// Production native Chatbox UI, isolated profile, synthetic workbook and local MCP.
const {_electron:electron}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const fs=require('node:fs'),path=require('node:path'),net=require('node:net'),crypto=require('node:crypto'),assert=require('node:assert/strict');
const http=require('node:http');
const {Client}=require('../integrations/chatbox/node_modules/@modelcontextprotocol/sdk/dist/cjs/client/index.js');
const {StdioClientTransport}=require('../integrations/chatbox/node_modules/@modelcontextprotocol/sdk/dist/cjs/client/stdio.js');
const root=path.resolve(__dirname,'..'),profile=fs.mkdtempSync(path.join(root,'.temp/ai-ui-'));
const pipe='GridPilot-'+crypto.randomBytes(16).toString('hex'),snapshotId='a'.repeat(32);
let application,socket,mcpClient,applied=0,captured=0,approval={automatic:false,available:false};
const modelRequests=[],modelHeaders=[];
const modelServer=http.createServer((req,res)=>{
  let text='';req.on('data',d=>text+=d);req.on('end',()=>{
    if(req.method==='GET'){res.setHeader('Content-Type','application/json');res.end(JSON.stringify({object:'list',data:[{id:'gpt-4o',object:'model'}]}));return;}
    const body=JSON.parse(text);modelRequests.push(body);modelHeaders.push(req.headers);
    const name=body.tools?.find(t=>t.function.name.endsWith('excel_statistics'))?.function.name;
    const hasResult=body.messages.some(m=>m.role==='tool');
    res.writeHead(200,{'Content-Type':'text/event-stream'});
    const chunk=(delta,finish_reason=null)=>res.write('data: '+JSON.stringify({id:'synthetic',object:'chat.completion.chunk',created:1,model:'gpt-4o',choices:[{index:0,delta,finish_reason}]})+'\n\n');
    chunk({role:'assistant'});
    if(name&&!hasResult){chunk({tool_calls:[{index:0,id:'stats1',type:'function',function:{name,arguments:JSON.stringify({snapshotId,column:2})}}]});chunk({},'tool_calls');}
    else {chunk({content:'完整选区统计：合计为 60，来源 销售!B2:B4。'});chunk({},'stop');}
    res.end('data: [DONE]\n\n');
  });
});
const capture={snapshotId,book:'合成销售.xlsx',scope:'selection',capturedAt:new Date().toISOString(),excelVersion:'16.0',ranges:[{sheet:'销售',address:'A1:C4'}],scannedCells:12,nonEmptyCells:10,sample:[{sheet:'销售',address:'B2',value:10,formula:null}],statistics:{sum:60,numericCount:3}};
function respond(m){
  if(m.action==='aiApproval'){approval={automatic:!!m.enabled,approvalScope:m.scope,available:false};return approval;}
  if(m.action==='aiCheckpoint')return approval;
  if(m.action==='aiConversation'){if(approval.approvalScope==='session')approval.automatic=false;return approval;}
  if(m.action==='aiDiscard')return {message:'已清空'};
  if(m.action==='aiReady')return {mode:'window',version:'1.20.3'};
  if(m.action==='aiContext')return {available:true,book:'合成销售.xlsx',sheet:'销售',selection:'A1:C4'};
  if(m.action==='aiCapture'){captured++;return {...capture,scope:m.scope};}
  if(m.action==='aiPropose')return {proposalId:'test-proposal',snapshotId,book:'合成销售.xlsx',sheet:'销售',address:m.address,formula:m.formula,before:null};
  if(m.action==='aiApplyBatch'){applied++;return {message:'已应用整批修改',appliedIds:m.proposalIds};}
  if(m.action==='aiMode')return {mode:m.mode};
  if(m.action==='aiUndo')return {message:'已撤销AI公式'};
  if(m.action==='aiNavigate')return {located:true};
  if(m.action==='aiRead')return m.operation==='statistics'?{snapshotId,statistics:{sum:60,numericCount:3,average:20,min:10,max:30}}:{snapshotId,total:1,cells:[{sheet:'销售',address:'B2',value:10}]};
  throw Error(m.action);
}
const server=net.createServer(s=>{
  socket=s;let text='';s.on('data',data=>{text+=data;let at;while((at=text.indexOf('\n'))>=0){
    const m=JSON.parse(text.slice(0,at));text=text.slice(at+1);
    try{const reply=JSON.stringify({id:m.id,data:respond(m)})+'\n';if(m.action==='aiApplyBatch')setTimeout(()=>s.write(reply),900);else s.write(reply);}catch(e){s.write(JSON.stringify({id:m.id,error:e.message})+'\n');}
  }});
});
(async()=>{
  await new Promise(resolve=>server.listen('\\\\.\\pipe\\'+pipe,resolve));
  await new Promise(resolve=>modelServer.listen(0,'127.0.0.1',resolve));
  fs.writeFileSync(path.join(profile,'config.json'),JSON.stringify({settings:{language:'zh-Hans',autoUpdate:false,providers:{openai:{apiKey:'synthetic-test-only',apiHost:'http://127.0.0.1:'+modelServer.address().port,models:[{modelId:'gpt-4o',capabilities:['tool_use'],contextWindow:128000}]}}},lastShownAboutDialogVersion:'1.20.3'}));
  try{
    const env={...process.env,GRIDPILOT_CHATBOX_PROFILE:profile};delete env.ELECTRON_RUN_AS_NODE;delete env.NODE_OPTIONS;
    // With executablePath Playwright does not add its ready-event loader itself.
    // Hold Electron ready until CDP is attached, so the native preload cannot race it.
    const loader=path.resolve(path.dirname(require.resolve(process.env.PLAYWRIGHT_MODULE||'playwright')),'../playwright-core/lib/server/electron/loader.js');
    application=await electron.launch({executablePath:path.join(root,'.temp/Browser/electron.exe'),args:['-r',loader,path.join(root,'.temp/Chatbox'),'--gridpilot-pipe='+pipe],env});
    const page=await application.firstWindow();page.setDefaultTimeout(30000);
    const errors=[];page.on('pageerror',e=>errors.push(e.message));
    page.on('console',m=>{if(m.type()==='error')console.log('Renderer error:',m.text().slice(0,1000));});
    await page.waitForLoadState('domcontentloaded');
    await page.waitForFunction(()=>!!window.gridpilotExcel).catch(async e=>{
      console.log('Window diagnostics',await application.evaluate(({BrowserWindow,app})=>({packaged:app.isPackaged,windows:BrowserWindow.getAllWindows().map(w=>({url:w.webContents.getURL(),preload:w.webContents.getLastWebPreferences().preload,contextIsolation:w.webContents.getLastWebPreferences().contextIsolation,sandbox:w.webContents.getLastWebPreferences().sandbox}))})));throw e;
    });
    console.log('Preload:',await page.evaluate(()=>({native:!!window.electronAPI,excel:!!window.gridpilotExcel})));
    await page.getByText('Just chat',{exact:true}).click();
    await page.locator('#gridpilot-excel-tools').waitFor({timeout:30000}).catch(async e=>{
      console.log('UI diagnostic:',await page.locator('body').innerText());console.log('errors',errors);
      await page.screenshot({path:path.join(root,'.temp/ai-ui-failure.png')});throw e;
    });
    console.log('Native Chatbox ready',await page.locator('body').innerText().then(s=>s.slice(0,1200)));
    assert.equal(await page.locator('#gridpilot-excel').isVisible(),false,'idle integration uses no extra panel');
    assert(await page.locator('#gridpilot-toolbar-slot #gridpilot-excel-tools').count());
    assert(await page.getByRole('switch',{name:'自动批准'}).isVisible(),'approval is visible without opening More');
    assert.equal(await page.locator('#gridpilot-native-web-search').isVisible(),false,'network action moved out of toolbar');
    await page.getByRole('switch',{name:'自动批准'}).click();
    await page.waitForFunction(()=>!document.querySelector('.gp-menu').open);
    assert(approval.automatic&&approval.approvalScope==='workbook','default opt-in scope is current workbook');
    assert.equal(await page.getByRole('switch',{name:'自动批准'}).evaluate(e=>getComputedStyle(e).color),'rgb(220, 69, 75)','auto approval icon is red');
    fs.mkdirSync(path.join(root,'docs/v2.1'),{recursive:true});
    await page.screenshot({path:path.join(root,'docs/v2.1/auto-approval.png')});
    const webState=await page.locator('#gridpilot-native-web-search').getAttribute('data-enabled');
    await page.getByLabel('Excel工具',{exact:true}).click();await page.getByRole('button',{name:/联网搜索/}).click();
    await page.waitForFunction(previous=>document.getElementById('gridpilot-native-web-search').dataset.enabled!==previous,webState);
    await page.getByLabel('Excel工具',{exact:true}).click();await page.getByRole('button',{name:/联网搜索/}).click();
    await page.waitForFunction(previous=>document.getElementById('gridpilot-native-web-search').dataset.enabled===previous,webState);
    await page.getByLabel('Excel工具',{exact:true}).click();await page.locator('#message-input').focus();assert.equal(await page.locator('.gp-menu').evaluate(e=>e.open),false,'focus input closes menu');
    await page.getByLabel('Excel工具',{exact:true}).click();await page.keyboard.press('Escape');assert.equal(await page.locator('.gp-menu').evaluate(e=>e.open),false,'Escape closes menu');
    await page.getByLabel('Excel工具',{exact:true}).click();await page.locator('body').click({position:{x:15,y:250}});assert.equal(await page.locator('.gp-menu').evaluate(e=>e.open),false,'outside pointer closes menu');
    await page.getByLabel('Excel工具',{exact:true}).click();await page.getByRole('switch').click();await page.waitForFunction(()=>!document.querySelector('.gp-menu').open);
    console.log('PASS approval popup dismisses on action, outside click, input focus and Escape');
    await page.screenshot({path:path.join(root,'.temp/ai-ui-inspect.png')});
    await page.getByLabel('Excel工具',{exact:true}).click();

    await page.locator('#gridpilot-excel-tools button').filter({hasText:'解释公式'}).click();
    assert.match(await page.locator('#message-input').inputValue(),/解释选中/);
    await page.getByLabel('Excel上下文范围').selectOption('workbook');
    const message=await page.evaluate(()=>window.gridpilotExcel.captureMessage({id:'test-context',role:'user',contentParts:[{type:'text',text:'统计数据'}]}));
    assert.match(message.contentParts[1].text,/📎 Excel/);assert(!message.contentParts[1].text.includes('"sample"'));assert(!message.contentParts[1].text.includes(snapshotId));assert.equal(captured,1);
    const expanded=await page.evaluate(m=>window.gridpilotExcel.expandMessages([m]),message);
    assert.match(expanded[0].contentParts[2].text,/"scope":"workbook"/);
    await page.getByLabel('Excel上下文范围').selectOption('none');
    const plain=await page.evaluate(()=>window.gridpilotExcel.captureMessage({role:'user',contentParts:[{type:'text',text:'一般聊天'}]}));
    assert.equal(plain.contentParts.length,1);assert.equal(captured,1);
    const config=await page.evaluate(()=>window.electronAPI.invoke('getStoreValue','settings'));
    const bridge=config.mcp.servers.find(s=>s.id==='gridpilot-excel');assert(bridge);assert.equal(config.autoUpdate,false);
    assert.equal(bridge.transport.type,'stdio');assert(!JSON.stringify(bridge).includes('http://'));
    mcpClient=new Client({name:'GridPilot-test',version:'1.0.0'});
    await mcpClient.connect(new StdioClientTransport({...bridge.transport,stderr:'pipe'}));
    const rpc=async(method,params={})=>({result:method==='tools/list'?await mcpClient.listTools():await mcpClient.callTool(params)});
    const list=await rpc('tools/list');assert.equal(list.result.tools.length,13);assert(!list.result.tools.some(t=>/apply|undo/.test(t.name)));
    const proposed=await rpc('tools/call',{name:'excel_propose_formula',arguments:{snapshotId,sheet:'销售',address:'D2',formula:'=SUM(B2:B4)'}});
    assert(!proposed.result.isError);assert.equal(applied,0);
    await page.locator('#gridpilot-excel').getByText('=SUM(B2:B4)',{exact:true}).waitFor({state:'attached'});
    assert(await page.getByRole('button',{name:'一键应用',exact:true}).count(),'Apply is accessible outside native modals');
    console.log('Proposal DOM:',await page.locator('#gridpilot-excel').evaluate(e=>({text:e.textContent,details:[...e.querySelectorAll('details')].map(d=>d.open)})));
    await page.screenshot({path:path.join(root,'.temp/ai-proposal-inspect.png')});
    await page.locator('#gridpilot-excel button').filter({hasText:'一键应用'}).click();assert.equal(applied,1);
    const working=page.getByRole('button',{name:'正在应用并保存快照…',exact:true});assert(await working.isDisabled());
    const colors=await working.evaluate(e=>({color:getComputedStyle(e).color,background:getComputedStyle(e).backgroundColor,opacity:getComputedStyle(e).opacity}));assert.equal(colors.color,'rgb(255, 255, 255)');assert.equal(colors.background,'rgb(73, 110, 100)');assert.equal(colors.opacity,'1');
    fs.mkdirSync(path.join(root,'docs/v2.1'),{recursive:true});await page.locator('#gridpilot-excel').screenshot({path:path.join(root,'docs/v2.1/applying.png')});
    await page.waitForFunction(()=>!document.querySelector('#gridpilot-excel .gp-primary'));console.log('PASS clear, high-contrast applying state and finished batch');
    await page.getByLabel('Excel工具',{exact:true}).click();await page.locator('#gridpilot-excel-tools button').filter({hasText:'撤销上次应用'}).click();await page.getByLabel('Excel工具',{exact:true}).click();
    await page.getByLabel('Excel上下文范围').selectOption('selection');
    await page.getByText('选择模型',{exact:true}).click();
    console.log('Model menu:',(await page.locator('body').innerText()).slice(-1800));
    await page.getByText('GPT-4o',{exact:true}).last().click();
    await page.locator('#message-input').fill('统计当前选区B列的合计，使用Excel统计工具。');
    await page.locator('#message-input').press('Enter');
    await page.getByText('完整选区统计：合计为 60，来源 销售!B2:B4。',{exact:false}).first().waitFor();
    assert(modelRequests.some(r=>JSON.stringify(r.messages).includes(snapshotId)));
    assert(modelRequests.some(r=>r.messages.some(m=>m.role==='tool')),'native model tool loop returned full statistics');
    assert(modelHeaders.every(h=>h.authorization==='Bearer synthetic-test-only'));assert(modelHeaders.every(h=>!/gridpilot/i.test(h['user-agent']||'')));
    assert.equal(await application.evaluate(({app})=>app.getVersion()),'1.20.3');
    console.log('Native provider header names',Object.keys(modelHeaders[0]).sort());
    // Test the native parser plus the same File.path compatibility used by pick/drop.
    for(const ext of ['txt','pdf','docx']){
      const fixture=path.join(root,'tests/fixtures/ai/中文附件.'+ext);
      const parsed=await page.evaluate(async filePath=>JSON.parse(await window.electronAPI.invoke('parseFileLocally',JSON.stringify({filePath}))),fixture);
      assert(parsed.isSupported,ext+' native parser supports attachment');assert.match(parsed.text,/GridPilot/);
    }
    const upload=page.locator('#input-box-2024-02-22 input[type=file]').last();
    const pdf=path.join(root,'tests/fixtures/ai/中文附件.pdf');
    await page.evaluate(()=>document.addEventListener('change',e=>{if(e.target.files?.length)window.__fixturePath=e.target.files[0].path;},{capture:true,once:true}));
    await upload.setInputFiles(pdf);
    const actualPath=await page.evaluate(()=>window.__fixturePath);assert.equal(actualPath,pdf,'native File.path bridge preserves picker path');
    await page.getByText('中文附件.pdf',{exact:false}).first().waitFor();
    await page.getByLabel('Excel上下文范围').selectOption('none');
    await page.locator('#message-input').fill('读取已添加的 PDF 附件。');
    // Native parsing completes before its send control enables.
    await page.locator('#message-input').press('Enter');
    await page.waitForFunction(()=>!document.querySelector('#message-input').value,{timeout:30000});
    await page.waitForTimeout(300);
    assert(modelRequests.some(r=>JSON.stringify(r.messages).includes('GridPilot PDF attachment verification')),'native model receives parsed PDF text');
    console.log('PASS native TXT/PDF/DOCX parsing, Unicode picker path, PDF message through Chatbox provider');
    fs.mkdirSync(path.join(root,'docs/v2.1'),{recursive:true});
    await page.screenshot({animations:'disabled',path:path.join(root,'docs/v2.1/chatbox-window.png')});
    await application.evaluate(({BrowserWindow})=>{const w=BrowserWindow.getAllWindows()[0];w.setSize(480,850);w.webContents.send('gridpilot:mode','sidebar');});
    await page.waitForFunction(()=>document.getElementById('gridpilot-excel')?.dataset.mode==='sidebar');
    await page.screenshot({animations:'disabled',path:path.join(root,'docs/v2.1/chatbox-narrow.png')});
    const overflow=await page.locator('#gridpilot-excel').evaluate(e=>e.scrollWidth>e.clientWidth+2);assert.equal(overflow,false);
    await application.evaluate(({BrowserWindow})=>BrowserWindow.getAllWindows()[0].setSize(1100,850));
    await application.evaluate(({BrowserWindow})=>BrowserWindow.getAllWindows()[0].webContents.send('navigate-to','/settings/provider'));
    await page.waitForURL(/settings/);
    await page.getByText('API 密钥',{exact:true}).waitFor();
    await page.waitForTimeout(400); // Native modal transition, before the documentation screenshot.
    console.log('Native settings:',(await page.locator('body').innerText()).slice(-1500));
    await page.screenshot({animations:'disabled',path:path.join(root,'docs/v2.1/chatbox-settings.png')});
    assert.deepEqual(errors,[]);
    console.log('PASS official native UI, Chinese draft, context opt-out, full-scope MCP, stdio transport, proposal/apply/undo, narrow layout');
  }finally{if(mcpClient)await mcpClient.close();if(application)await application.close();socket?.destroy();server.close();modelServer.close();}
})().catch(e=>{console.error(e);process.exitCode=1;});
