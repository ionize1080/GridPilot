const {app,BrowserWindow,ipcMain,session,dialog,clipboard,ClipboardItem}=require('electron');
const path=require('path');
const fs=require('fs');
const os=require('os');
const readline=require('readline');
const net=require('net');
const {pathToFileURL}=require('url');
const profile=fs.mkdtempSync(path.join(os.tmpdir(),'GridPilot-ui-'));
app.setPath('userData',profile);
app.commandLine.appendSwitch('disable-background-networking');
app.commandLine.appendSwitch('disable-component-update');
app.commandLine.appendSwitch('disable-sync');
app.disableHardwareAcceleration(); // predictable performance on old Office PCs and remote desktops
app.enableSandbox();
let win,sequence=0;
const pending=new Map();
const pipeName=process.argv.find(v=>v.startsWith('--gridpilot-pipe='))?.slice('--gridpilot-pipe='.length);
if(!pipeName || !/^GridPilot-[a-f0-9]{32}$/.test(pipeName)){app.quit();throw new Error('GridPilot must be opened from Excel.');}
const connection=net.createConnection('\\\\.\\pipe\\'+pipeName);
const input=readline.createInterface({input:connection});
input.on('line',line=>{try{const message=JSON.parse(line);if(message.type==='progress'){win?.webContents.send('progress',message);return;}const slot=pending.get(message.id);if(slot){pending.delete(message.id);message.error?slot.reject(new Error(message.error)):slot.resolve(message.data);}}catch{}});
input.on('close',()=>app.quit());
connection.on('error',()=>app.quit());
function hostRequest(message){return new Promise((resolve,reject)=>{const id=++sequence;pending.set(id,{resolve,reject});connection.write(JSON.stringify({...message,id})+'\n');});}
const {DatabaseManager}=require('./database-main.cjs');
const databases=new DatabaseManager(path.join(app.getPath('appData'),'GridPilot','Databases'),hostRequest,p=>win?.webContents.send('progress',p));
const file=pathToFileURL(path.join(__dirname,'index.html')).href;
ipcMain.handle('request',async(event,message)=>{
  if(event.sender!==win?.webContents || event.senderFrame.url!==file)throw new Error('Invalid sender');
  if(!message || !['init','preview','export','cancel','saveSettings','navigationSettings','previewPage','copyMask','dbList','dbChoose','dbSample','dbCreate','dbInfo','dbScan','dbQuery','dbOpenSource','legal'].includes(message.action))throw new Error('Invalid action');
  if(JSON.stringify(message).length>32*1024*1024)throw new Error('请求过大');
  if(message.action==='legal'){const file=path.join(__dirname,'legal.bin');if(!fs.existsSync(file))return {titles:[],text:'许可信息将在正式发布包中提供。',offset:0,length:0};const entries=JSON.parse(require('node:zlib').gunzipSync(fs.readFileSync(file)));const selected=entries[Math.max(0,Math.min(entries.length-1,Number(message.index)||0))];let offset=Math.max(0,Number(message.offset)||0);if(message.search)offset=Math.max(0,selected.text.toLowerCase().indexOf(String(message.search).toLowerCase(),offset));return {titles:entries.map(e=>e.title),text:selected.text.slice(offset,offset+40000),offset,length:selected.text.length};}
  if(message.action.startsWith('db'))return databases.handle(message,dialog,win);
  if(message.action==='cancel')databases.cancel();
  if(message.action==='copyMask'){const result=await hostRequest(message);await clipboard.write([new ClipboardItem({'text/plain':result.text,'text/html':result.html})]);return {copied:true,message:`已复制 ${result.rows} 行 × ${result.columns} 列脱敏结果。`};}
  if(message.action==='export'){const choice=await dialog.showSaveDialog(win,{title:'保存脱敏副本',defaultPath:'GridPilot-masked.xlsx',filters:[{name:'Excel 工作簿',extensions:['xlsx']}],properties:['showOverwriteConfirmation']});if(choice.canceled)return {canceled:true,message:'已取消保存，预览仍保留。'};message={...message,path:choice.filePath};}
  return hostRequest(message);
});
ipcMain.on('close',event=>{if(event.sender===win?.webContents)win.close();});
app.whenReady().then(()=>{
  session.defaultSession.setPermissionRequestHandler((web,permission,callback)=>callback(false));
  session.defaultSession.webRequest.onBeforeRequest((details,callback)=>callback({cancel:!details.url.startsWith(pathToFileURL(__dirname+path.sep).href)}));
  win=new BrowserWindow({width:1180,height:820,minWidth:760,minHeight:540,show:false,backgroundColor:'#f5f7fa',title:'GridPilot',autoHideMenuBar:true,webPreferences:{preload:path.join(__dirname,'preload.js'),sandbox:true,contextIsolation:true,nodeIntegration:false,spellcheck:false,devTools:false}});
  win.setMenu(null);
  win.webContents.setWindowOpenHandler(()=>({action:'deny'}));
  win.webContents.on('will-navigate',event=>event.preventDefault());
  win.webContents.on('will-attach-webview',event=>event.preventDefault());
  win.once('ready-to-show',()=>win.show());
  win.loadFile(path.join(__dirname,'index.html'));
});
app.on('window-all-closed',()=>app.quit());
app.on('quit',()=>{databases.cancel();try{fs.rmSync(profile,{recursive:true,force:true});}catch{}});
