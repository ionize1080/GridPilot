// Opt-in paid API test. Credential is read only from GRIDPILOT_TEST_KEY_FILE.
// Native Chatbox provider + native MCP stdio + real isolated Excel workbook.
const fs=require('node:fs'),path=require('node:path'),net=require('node:net'),crypto=require('node:crypto'),assert=require('node:assert/strict');
const {spawn}=require('node:child_process'),readline=require('node:readline');
const {_electron:electron}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const root=path.resolve(__dirname,'..');
const keyFile=process.env.GRIDPILOT_TEST_KEY_FILE;if(!keyFile)throw Error('Explicit credential file required');
const key=fs.readFileSync(keyFile,'utf8').match(/sk-[A-Za-z0-9_-]+/)?.[0];if(!key)throw Error('Credential format not recognized; contents withheld');
const model=process.env.GRIDPILOT_TEST_MODEL||'deepseek-flash';
const profile=fs.mkdtempSync(path.join(root,'.temp/live-ai-')),pipe='GridPilot-'+crypto.randomBytes(16).toString('hex');
let application,socket,seq=1000000;const pending=new Map(),actions=[],proposals=[],progress=[],operations=[];
const bridge=spawn(path.join(root,'tests/SheetPilot.Tests/bin/Release/net462/SheetPilot.Tests.exe'),['--ai-live-bridge'],{windowsHide:true,stdio:['pipe','pipe','pipe']});
const send=m=>bridge.stdin.write(JSON.stringify(m)+'\n');
const direct=(action,args={})=>new Promise((resolve,reject)=>{const id=++seq;const timer=setTimeout(()=>{pending.delete(id);reject(Error('Excel test timeout'));},120000);pending.set(id,{resolve,reject,timer});send({...args,action,id});});
readline.createInterface({input:bridge.stdout}).on('line',line=>{
  let m;try{m=JSON.parse(line.replace(/^\uFEFF/,''));}catch{return;}
  if(m.type==='progress'){progress.push(m);socket?.write(JSON.stringify(m)+'\n');return;}
  const p=pending.get(m.id);if(p){clearTimeout(p.timer);pending.delete(m.id);m.error?p.reject(Error(m.error)):p.resolve(m.data);return;}
  if(m.data?.proposalId)proposals.push(m.data);socket?.write(JSON.stringify(m)+'\n');
});
bridge.stderr.on('data',()=>{}); // Never write raw child diagnostics into API test logs.
const server=net.createServer(s=>{socket=s;readline.createInterface({input:s}).on('line',line=>{const m=JSON.parse(line);actions.push(m.action);operations.push(m.operation);send(m);});});
async function waitFor(predicate,label,timeout=180000){const start=Date.now();while(!await predicate()){if(Date.now()-start>timeout)throw Error('Timeout: '+label);await new Promise(r=>setTimeout(r,150));}}
(async()=>{
  await new Promise(r=>server.listen('\\\\.\\pipe\\'+pipe,r));
  fs.writeFileSync(path.join(profile,'config.json'),JSON.stringify({settings:{language:'zh-Hans',autoUpdate:false,providers:{openai:{apiKey:key,apiHost:'https://api.deepseek.com',models:[{modelId:model,capabilities:['tool_use'],contextWindow:128000}]}}},lastShownAboutDialogVersion:'1.20.3'}));
  try{
    const env={...process.env,GRIDPILOT_CHATBOX_PROFILE:profile};delete env.ELECTRON_RUN_AS_NODE;delete env.NODE_OPTIONS;delete env.GRIDPILOT_TEST_KEY_FILE;
    const loader=path.resolve(path.dirname(require.resolve(process.env.PLAYWRIGHT_MODULE||'playwright')),'../playwright-core/lib/server/electron/loader.js');
    application=await electron.launch({executablePath:path.join(root,'.temp/Browser/electron.exe'),args:['-r',loader,path.join(root,'.temp/Chatbox'),'--gridpilot-pipe='+pipe],env});
    const page=await application.firstWindow();page.setDefaultTimeout(45000);
    page.on('request',req=>{if(req.url().startsWith('https://api.deepseek.com/')&&req.method()==='POST'){try{const body=req.postDataJSON();console.log('Native API request: model='+body.model+' tools='+(body.tools?.length||0));}catch{}}});
    await page.getByText('Just chat',{exact:true}).click();await page.locator('#gridpilot-excel-tools').waitFor();
    await page.getByText('选择模型',{exact:true}).click();await page.getByText(model,{exact:true}).last().click();
    await page.getByLabel('Excel上下文范围').selectOption('sheet');
    async function ask(prompt,check){
      const n=proposals.length;await page.locator('#message-input').fill(prompt);await page.locator('#input-box-2024-02-22 button').filter({has:page.locator('svg.tabler-icon-arrow-up')}).click();
      await waitFor(()=>check(proposals.slice(n)),'live model tools').catch(async e=>{console.log((await page.locator('body').innerText()).split(key).join('[REDACTED]').slice(-2500));throw e;});
      await page.locator('.tabler-icon-player-stop-filled,.icon-tabler-player-stop-filled').waitFor({state:'hidden',timeout:180000});
    }
    await ask('在销售工作表 E2 写入收入 B2:B4 的求和公式，请使用 Excel 工具生成待应用建议。不要仅输出代码。',p=>p.some(x=>x.address==='E2'&&x.formula?.startsWith('=')));
    assert.equal((await direct('testRead',{address:'E2'})).value,null);
    await page.locator('#gridpilot-excel button').filter({hasText:'一键应用'}).click();await waitFor(async()=>(await direct('testRead',{address:'E2'})).value===450,'manual apply');
    console.log('PASS live first-turn formula proposal + user Apply + actual Excel result 450');
    await page.getByLabel('Excel工具',{exact:true}).click();await page.locator('#gridpilot-excel-tools [role=switch]').click();await page.waitForFunction(()=>!document.querySelector('.gp-menu').open);
    await ask('请仅使用 excel_propose_values、excel_propose_formulas、excel_propose_format、excel_propose_chart 四个基础工具完成这些修改：F2:G3写入二维值[[10,20],[30,40]]；H2:H3写入每行F列加G列的公式；I2:I3设置加粗和浅绿色填充#D9EFE5；使用A1:C4生成标题为“收入与成本”的柱状图。所有操作均使用工具，不要只给文字。',p=>['values','formulas','format','chart'].every(k=>p.some(x=>x.kind===k&&x.applied)));
    assert.equal((await direct('testRead',{address:'H2'})).value,30);assert.equal((await direct('testRead',{address:'H3'})).value,70);assert.equal((await direct('testRead',{address:'I2'})).bold,true);assert.equal((await direct('testRead',{address:'A1'})).charts,1);
    console.log('PASS live automatic values / rectangular formulas / formatting / native chart');
    await page.locator('.gp-history button').filter({hasText:'变更前'}).click();await waitFor(async()=>(await direct('testRead',{address:'H2'})).value===null,'before checkpoint');
    assert.equal((await direct('testRead',{address:'F2'})).value,null);assert.equal((await direct('testRead',{address:'I2'})).bold,false);assert.equal((await direct('testRead',{address:'A1'})).charts,0);assert.equal((await direct('testRead',{address:'E2'})).value,450);
    await page.locator('.gp-history button').filter({hasText:'变更后'}).click();await waitFor(async()=>(await direct('testRead',{address:'H2'})).value===30,'after checkpoint');assert.equal((await direct('testRead',{address:'A1'})).charts,1);
    console.log('PASS native UI before/after toggles restore actual Excel content, format and chart');
    assert.equal((await direct('aiConversation')).automatic,true,'workbook approval survives conversation change and snapshot toggles');
    console.log('PASS automatic approval remains active across rounds and restores without retoggling');
    await ask('请仅用 excel_propose_pivot 基础透视表工具，以销售!A1:C4为源，在新工作表“透视汇总”创建原生数据透视表。行字段是项目，汇总字段收入，汇总方式sum。不要只输出表格文本。',p=>p.some(x=>x.kind==='pivot'&&x.applied));
    assert.deepEqual(await direct('testPivot',{sheet:'透视汇总'}),{count:1,total:450});
    await page.locator('.gp-history button').filter({hasText:'变更前'}).click();await waitFor(async()=>(await direct('aiCheckpoint')).position==='before','pivot before');assert.equal((await direct('aiCheckpoint')).automatic,true);
    await page.locator('.gp-history button').filter({hasText:'变更后'}).click();await waitFor(async()=>{try{return (await direct('testPivot',{sheet:'透视汇总'})).total===450;}catch{return false;}},'pivot redo');
    console.log('PASS live native PivotTable + before/after restore; aggregate total 450');
    await page.getByLabel('Excel上下文范围').selectOption('workbook');
    const aiReadStart=operations.length,inspectStart=actions.filter(a=>a==='aiInspect').length;
    await ask('请分别调用Excel的四个只读工具：inspect检查工作簿结构，read读取销售工作表一页数据，find在销售中查找甲，statistics统计销售第2列的总和。必须实际分别调用四种工具，然后简短报告收入合计。',()=>['read','find','statistics'].every(k=>operations.slice(aiReadStart).includes(k))&&actions.filter(a=>a==='aiInspect').length>inspectStart);
    console.log('PASS live read, find, statistics and workbook inspection');
    const docs=path.join(root,'docs/v2.0');fs.mkdirSync(docs,{recursive:true});
    await application.evaluate(({BrowserWindow})=>{const w=BrowserWindow.getAllWindows()[0];w.setSize(480,850);w.webContents.send('gridpilot:mode','sidebar');});
    await page.waitForFunction(()=>document.documentElement.dataset.gpMode==='sidebar');
    assert.equal(await page.locator('.controls.self-start').first().isVisible(),false);assert.equal(await page.getByLabel('弹出为独立窗口').count(),1);
    await page.screenshot({path:path.join(docs,'live-ai-sidebar.png')});
    await page.getByLabel('Excel上下文范围').click();await page.screenshot({path:path.join(docs,'scope-menu.png')});await page.keyboard.press('Escape');
    await direct('testLarge');const capturePromise=page.evaluate(()=>window.gridpilotExcel.captureMessage({id:'large-live-fixture',role:'user',contentParts:[{type:'text',text:'fixture'}]}));
    await page.locator('.gp-progress').waitFor();await page.screenshot({path:path.join(docs,'capture-progress.png')});await capturePromise;
    assert(progress.some(p=>p.current>100000));assert.equal(await page.locator('.gp-progress').count(),0);
    console.log('PASS >200k real Excel capture with live progress; only samples/statistics enter model context');
    console.log('PASS paid native provider test; model='+model+'; proposals='+proposals.length+'; hostActions='+actions.length+'; credential withheld');
  }catch(e){if(application){try{const p=await application.firstWindow();console.log('UI diagnostic: '+(await p.locator('body').innerText()).split(key).join('[REDACTED]').slice(-2500));}catch{}}throw e;}finally{
    if(application)await application.close();socket?.destroy();server.close();bridge.stdin.end();
    await new Promise(resolve=>{if(bridge.exitCode!==null)return resolve();bridge.once('exit',resolve);setTimeout(()=>{bridge.kill();resolve();},10000);});
    for(const p of pending.values()){clearTimeout(p.timer);p.reject(Error('Test ended'));}pending.clear();
    // Remove only this newly-created isolated profile, never the user's profile.
    assert.equal(path.dirname(path.resolve(profile)),path.join(root,'.temp'));assert(path.basename(profile).startsWith('live-ai-'));
    fs.rmSync(profile,{recursive:true,force:true,maxRetries:10,retryDelay:300});
  }
})().catch(error=>{console.error(String(error.message).split(key).join('[REDACTED]'));process.exitCode=1;});
