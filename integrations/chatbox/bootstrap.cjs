// GridPilot connection layer for the SHA-256 pinned official Chatbox 1.20.3.
const {app, ipcMain, BrowserWindow} = require('electron');
const path = require('node:path'), net = require('node:net');
const readline = require('node:readline'), crypto = require('node:crypto');
const {tools, handleTool} = require('./tools.cjs');
app.setPath('userData', process.env.GRIDPILOT_CHATBOX_PROFILE || path.join(app.getPath('appData'),'GridPilot','Chatbox-1.20.3'));
Object.defineProperty(app,'isPackaged',{get:()=>true});
Object.defineProperty(process,'resourcesPath',{value:path.join(__dirname,'resources')});
app.setAsDefaultProtocolClient = () => false;
const pipe = process.argv.find(v=>v.startsWith('--gridpilot-pipe='))?.split('=')[1];
if (!/^GridPilot-[a-f0-9]{32}$/.test(pipe || '')) throw Error('请从GridPilot打开Chatbox。');
const socket = net.createConnection('\\\\.\\pipe\\'+pipe), pending = new Map();
let sequence=0, primary,displayMode='window';
function applyMode(mode){
  displayMode=mode;if(!primary)return;
  const docked=mode==='sidebar';primary.setHasShadow(!docked);primary.setSkipTaskbar(docked);primary.setMovable(!docked);primary.setMaximizable(!docked);primary.setMinimizable(!docked);
  primary.webContents.send('gridpilot:mode',mode);
}
const token=crypto.randomBytes(32).toString('hex');
function host(message) {
  return new Promise((resolve,reject)=>{
    const id=++sequence, timer=setTimeout(()=>{pending.delete(id);reject(Error('Excel响应超时。请退出单元格编辑或关闭Excel对话框后重试。'));},600000);
    pending.set(id,{resolve,reject,timer});
    socket.write(JSON.stringify({...message,id})+'\n');
  });
}
readline.createInterface({input:socket}).on('line',line=>{
  try {const m=JSON.parse(line);if(m.type==='progress'){primary?.webContents.send('gridpilot:progress',m);return;}if(m.type==='mode'){applyMode(m.mode);return;}if(m.type==='activate'){primary?.show();primary?.focus();return;}const p=pending.get(m.id);if(!p)return;clearTimeout(p.timer);pending.delete(m.id);m.error?p.reject(Error(m.error)):p.resolve(m.data);}catch{}
});
function disconnect(){for(const p of pending.values()){clearTimeout(p.timer);p.reject(Error('Excel已关闭。'));}pending.clear();app.exit(0);}
socket.on('close',disconnect);socket.on('error',disconnect);
function publish(data){primary?.webContents.send('gridpilot:result',data);}
const allowed=new Set(['aiContext','aiCapture','aiRead','aiPropose','aiApply','aiUndo','aiNavigate','aiApplyBatch','aiMode','aiApproval','aiCheckpoint','aiRestore','aiOpenRecovery','cancel']);
ipcMain.handle('gridpilot:request',async(event,message)=>{
  if(!primary || event.sender!==primary.webContents || event.senderFrame!==primary.webContents.mainFrame ||
     !event.senderFrame.url.startsWith('file:') || !allowed.has(message?.action) || JSON.stringify(message).length>8*1024*1024)
    throw Error('无效的Excel请求。');
  const result=await host(message);
  if(message.action==='aiMode'){applyMode(result.mode);}
  if(message.action==='aiPropose')publish(result);
  return result;
});
// Private IPC only: the native Chatbox stdio child forwards approved tools here.
const bridgePipe = '\\\\.\\pipe\\GridPilot-MCP-' + crypto.randomBytes(16).toString('hex');
const clients = new Set();
const server=net.createServer(client=>{
  client.setEncoding('utf8');clients.add(client);client.on('close',()=>clients.delete(client));client.on('error',()=>client.destroy());
  let buffer='';
  client.on('data',chunk=>{
    buffer+=chunk.toString('utf8');if(buffer.length>8*1024*1024){client.destroy();return;}
    let at;while((at=buffer.indexOf('\n'))>=0){
      const line=buffer.slice(0,at);buffer=buffer.slice(at+1);
      (async()=>{let m;try{
        m=JSON.parse(line);if(m.token!==token)throw Error('Invalid bridge session');
        const result=await handleTool(m.params,host,publish);
        client.write(JSON.stringify({id:m.id,result})+'\n');
      }catch(e){client.write(JSON.stringify({id:m?.id,error:e.message})+'\n');}})();
    }
  });
});
server.listen(bridgePipe);
function withBridge(value){
  value=value||{};
  const servers=value.mcp?.servers||[], previous=servers.find(s=>s.id==='gridpilot-excel');
  return {...value,autoUpdate:false,mcp:{...value.mcp,enabledBuiltinServers:value.mcp?.enabledBuiltinServers||[],
    servers:[...servers.filter(s=>s.id!=='gridpilot-excel'),{id:'gridpilot-excel',name:'GridPilot Excel',enabled:previous?.enabled??true,
      transport:{type:'stdio',command:process.execPath,args:[path.join(__dirname,'stdio.cjs')],env:{ELECTRON_RUN_AS_NODE:'1',GRIDPILOT_MCP_PIPE:bridgePipe,GRIDPILOT_MCP_TOKEN:token}}}]}};
}
// Decorate native storage reads; leave every original setting and setting page
// in Chatbox. Other native/custom MCP servers retain their original entries.
const nativeHandle=ipcMain.handle.bind(ipcMain);
ipcMain.handle=(channel,listener)=>{
  if(channel==='install-update')return nativeHandle(channel,()=>{throw Error('GridPilot固定使用Chatbox 1.20.3，请通过GridPilot发布页升级。');});
  if(channel==='getSettings')return nativeHandle(channel,async(...args)=>withBridge(await listener(...args)));
  if(channel==='getStoreValue')return nativeHandle(channel,async(event,key,...args)=>{
    const value=await listener(event,key,...args);return key==='settings'?withBridge(value):value;
  });
  return nativeHandle(channel,listener);
};
app.on('browser-window-created',(_event,window)=>{
  if(primary)return;primary=window;
  window.setBackgroundColor('#ffffff');window.setMinimumSize(320,240);
  window.webContents.on('will-navigate',(event,url)=>{if(!url.startsWith('file:'))event.preventDefault();});
  window.once('ready-to-show',async()=>{
    try{
      const h=window.getNativeWindowHandle();
      const result=await host({action:'aiReady',handle:(h.length===8?h.readBigUInt64LE():BigInt(h.readUInt32LE())).toString()});
      if(result.mode==='sidebar'){applyMode('sidebar');}
      else if(window.getBounds().width<640)window.setSize(1000,800);
    }catch(error){require('electron').dialog.showErrorBox('GridPilot AI',error.message);app.exit(1);}
  });
  window.on('maximize',()=>{if(displayMode==='sidebar')window.unmaximize();});
  window.on('minimize',()=>{if(displayMode==='sidebar')window.restore();});
  window.on('closed',()=>app.exit(0));
});
// Register native ready handlers during module loading, before Electron emits ready.
require('./native/dist/main/main.js');
app.on('quit',()=>{for(const client of clients)client.destroy();server.close();socket.destroy();});
