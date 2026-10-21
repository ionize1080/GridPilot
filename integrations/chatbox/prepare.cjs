const fs=require('node:fs'),path=require('node:path'),crypto=require('node:crypto'),{execFileSync}=require('node:child_process');
const asar=require('@electron/asar');
const root=path.resolve(__dirname,'../..'),temp=path.join(root,'.temp');
const installer=path.join(temp,'Chatbox-1.20.3-Setup.exe');
const hash=file=>crypto.createHash('sha256').update(fs.readFileSync(file)).digest('hex');
const expected='874171c6c7358ff3915f01412eb036154a5ded7a76c4663fb0f592d2c5c23b13';
if(hash(installer)!==expected)throw Error('Chatbox 1.20.3 official SHA-256 mismatch');
const extracted=path.join(temp,'chatbox-official-installer');
if(!fs.existsSync(path.join(extracted,'resources/app.asar'))){
  execFileSync(require('7zip-bin').path7za,['x',installer,'-o'+extracted,'-y'],{stdio:'pipe',windowsHide:true});
}
const archive=path.join(extracted,'resources/app.asar');
if(hash(archive)!=='e43647af8168b731caa9afaed269a4700f8686c257fbc4dd6a254da42729fa01')throw Error('Unexpected official app.asar');
const out=path.join(temp,'Chatbox'),native=path.join(out,'native');
fs.mkdirSync(native,{recursive:true});asar.extractAll(archive,native);
const packageData=JSON.parse(fs.readFileSync(path.join(native,'package.json'),'utf8'));
if(packageData.version!=='1.20.3')throw Error('Unexpected native version');
fs.mkdirSync(path.join(out,'resources'),{recursive:true});
fs.cpSync(path.join(extracted,'resources/assets'),path.join(out,'resources/assets'),{recursive:true});
const patches=[];
function patch(relative,before,after){
  const file=path.join(native,relative),text=fs.readFileSync(file,'utf8');
  if(text.split(before).length!==2)throw Error('Pinned binary patch anchor changed: '+relative);
  fs.writeFileSync(file,text.replace(before,after));
  patches.push({file:relative,originalSha256:crypto.createHash('sha256').update(text).digest('hex'),patchedSha256:hash(file),before,after});
}
patch('dist/main/main.js',
  'await createWindow(),ensureTray(),new AppUpdater(()=>mainWindow?.webContents.send("update-downloaded",{}))',
  'await createWindow(),void 0');
patch('dist/renderer/js/InputBox.DH3jG147.js',
  'constructedMessage:E.message,needGenerating:s,',
  'constructedMessage:await window.gridpilotCaptureMessage(E.message),needGenerating:s,');
patch('dist/renderer/js/index.TlLbGKvw.js',
  'maxContextMessageCount:nu.maxContextMessageCount});const yu=[];',
  'maxContextMessageCount:nu.maxContextMessageCount});ku=window.gridpilotExcel.expandMessages(ku,eu);const yu=[];');
patch('dist/renderer/js/InputBox.DH3jG147.js',
  'rr.knowledgeBase&&!M&&o.jsx(Bo,{currentKnowledgeBaseId:de?.id,onSelect:nn,',
  'o.jsx("span",{id:"gridpilot-toolbar-slot"}),rr.knowledgeBase&&!M&&o.jsx(Bo,{currentKnowledgeBaseId:de?.id,onSelect:nn,');
patch('dist/renderer/js/index.TlLbGKvw.js','function getToolName(eu){return{',
  'function getToolName(eu){const names={excel_read:"读取单元格",excel_find:"查找数据",excel_statistics:"统计数据",excel_inspect:"检查工作簿",excel_propose_formula:"写入公式",excel_propose_values:"写入内容",excel_propose_formulas:"写入批量公式",excel_propose_format:"调整格式",excel_propose_chart:"生成图表",excel_propose_pivot:"生成透视表"};if(eu.startsWith("mcp__gridpilot_excel__")){const label=names[eu.slice("mcp__gridpilot_excel__".length)];if(label)return "Excel · "+label;}return{');
const renderer=path.join(native,'dist/renderer/js/index.TlLbGKvw.js');
const rendererBefore=hash(renderer);
fs.appendFileSync(renderer,'\nObject.defineProperty(window.File.prototype,"path",{configurable:true,get(){return window.gridpilotExcel.filePath(this);}});\nwindow.addEventListener("gridpilot:sidebar",()=>uiStore.getState().setShowSidebar(false));\n');
patches.push({file:'dist/renderer/js/index.TlLbGKvw.js',originalSha256:rendererBefore,patchedSha256:hash(renderer),append:'Restore File.path through Electron webUtils for native attachment parsing. Collapse the native conversation drawer once when entering the Excel sidebar; its original toggle stays available.'});
const readinessBefore=hash(renderer);
fs.appendFileSync(renderer,`
window.gridpilotToolsStatus=()=>({state:mcpController.getServer("gridpilot-excel")?.status?.state,count:Object.keys(mcpController.getAvailableTools()).filter(n=>n.includes("excel_")).length});
window.gridpilotCaptureMessage=async message=>{
  if(window.gridpilotExcel.wantsContext()){
    const started=Date.now();
    while(!window.gridpilotToolsStatus().count){
      const server=mcpController.getServer("gridpilot-excel");
      if(server?.status?.error || Date.now()-started>15000)throw Error("Excel工具尚未连接，请在工具菜单中启用 GridPilot Excel 后重试。");
      await new Promise(resolve=>setTimeout(resolve,100));
    }
  }
  return window.gridpilotExcel.captureMessage(message);
};
`);
patches.push({file:'dist/renderer/js/index.TlLbGKvw.js',originalSha256:readinessBefore,patchedSha256:hash(renderer),append:'Wait for native GridPilot MCP schema discovery before the first model request; never synthesize tool calls from text.'});
const preload=path.join(native,'dist/preload/index.js'),original=hash(preload);
fs.appendFileSync(preload,'\n'+fs.readFileSync(path.join(__dirname,'preload.cjs'),'utf8'));
patches.push({file:'dist/preload/index.js',originalSha256:original,patchedSha256:hash(preload),append:'integrations/chatbox/preload.cjs'});
for(const name of ['bootstrap.cjs','tools.cjs','stdio.cjs','LICENSE-GPL-3.0.txt'])fs.copyFileSync(path.join(__dirname,name),path.join(out,name));
fs.cpSync(path.join(__dirname,'node_modules'),path.join(out,'node_modules'),{recursive:true});
fs.cpSync(path.join(extracted,'resources/app.asar.unpacked'),path.join(out,'resources/app.asar.unpacked'),{recursive:true});
fs.writeFileSync(path.join(out,'package.json'),JSON.stringify({name:packageData.name,productName:packageData.productName,version:'1.20.3',private:true,main:'bootstrap.cjs'}));
fs.writeFileSync(path.join(out,'provenance.json'),JSON.stringify({version:'1.20.3',upstream:'https://github.com/chatboxai/chatbox/releases/tag/v1.20.3',installerSha256:expected,archiveSha256:hash(archive),patches},null,2));
console.log('Prepared official Chatbox 1.20.3 with auditable integration patches:',out);
