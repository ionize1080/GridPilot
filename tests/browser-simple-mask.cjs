const {_electron:electron}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const assert=require('node:assert/strict'),fs=require('node:fs'),path=require('node:path'),os=require('node:os'),net=require('node:net'),crypto=require('node:crypto');
const root=path.resolve(__dirname,'..'),temp=fs.mkdtempSync(path.join(os.tmpdir(),'GridPilot-simple-ui-'));
async function main(){
 let application,socket,lastMessage,applied=false;const received=[];
 const pipe='GridPilot-'+crypto.randomBytes(16).toString('hex');
 const state={mode:'simple',book:'合成测试.xlsx',sheet:'客户清单',count:205,blocks:[{Address:'$A$1:$C$205',Row:1,Column:1,Rows:205,Columns:3}],mappingVersion:'GP-SIMPLE-1 / Unihan-17.0.0',letters:'岚砚穹榕汐棠霁珀笙岑翎澜槐璟沐筠霜屿檀洵苒曜栖珩蔚桐',digits:'7391058264'};
 state.headers={Message:'识别表头：第 4 行（2 个字段）。请核对，可手动修正。',Bands:[{Reason:'字段与下方数据判断',Fields:[{Label:'客户名'},{Label:'余额'}]}]};
 const server=net.createServer(s=>{socket=s;let buffer='';s.on('data',chunk=>{buffer+=chunk.toString();let end;while((end=buffer.indexOf('\n'))>=0){const m=JSON.parse(buffer.slice(0,end));buffer=buffer.slice(end+1);lastMessage=m;received.push(m);let data,error;
  if(m.action==='init'||m.action==='simpleCapture')data=state;
  else if(m.action==='simplePreview'||m.action==='simplePage'){
   if(m.options.Mode==='regex'&&m.options.Pattern==='[')error='正则表达式无效，请检查括号与转义。';
   else {if(m.action==='simplePreview')applied=false;const offset=m.offset||0,total=m.search?1:205;data={headers:state.headers,summary:'读取 205 格 · 变更 203 格 · 跳过 2 格。复制和导出中的跳过项留空。',changed:203,skipped:2,total,offset,rows:Array.from({length:Math.min(100,total-offset)},(_,i)=>({address:'$A$'+(offset+i+1),before:i===0?'新加坡油气有限公司':i===1?'P123456':i===2?'<script>alert(1)</script>':'示例'+i,after:i===0?'XJPYQYXGS':i===1?'56234P1':'结果'+i,status:applied?'已写入':i===3?'长度不足首尾各 2 个字符':'可处理'}))};}
  }else if(m.action==='simpleCopy')data={text:'XJPYQYXGS\t56234P1\r\n',html:'<table><tr><td>XJPYQYXGS</td><td>56234P1</td></tr></table>',rows:1,columns:2};
  else if(m.action==='simpleApply'){assert(m.backupConfirmed&&m.confirmed);applied=true;data={applied:true,message:'已写入 203 格；写入跳过 0 格；规则跳过 2 格。'};}
  else if(m.action==='simpleRecovery'){assert(m.sensitiveConfirmed&&m.confirmed);data={saved:m.path,message:'已导出含原文的还原记录'};}
  else if(m.action==='simpleExport')data={saved:m.path,message:'已导出脱敏结果'};
  else data={message:'已取消'};
  s.write(JSON.stringify({id:m.id,data,error})+'\n');
 }});});
 await new Promise(r=>server.listen('\\\\.\\pipe\\'+pipe,r));
 const bootstrap=path.join(temp,'boot.cjs'),appDir=process.env.GRIDPILOT_TEST_APP||path.join(root,'App');fs.writeFileSync(bootstrap,`require('electron').app.setPath('appData',${JSON.stringify(temp)});require(${JSON.stringify(path.join(appDir,'main.js'))});`);
 const env={...process.env};delete env.ELECTRON_RUN_AS_NODE;delete env.NODE_OPTIONS;
 try{
  application=await electron.launch({executablePath:process.env.GRIDPILOT_TEST_BROWSER||path.join(root,'.temp/Browser/electron.exe'),args:[bootstrap,'--gridpilot-pipe='+pipe],env});const page=await application.firstWindow(),errors=[];page.on('pageerror',e=>errors.push(e.message));
  await page.locator('#simplePreview').waitFor();assert(await page.locator('#simpleCopy').isDisabled());assert(await page.locator('#simpleApply').isDisabled());
  assert.equal(await page.locator('#simpleHeaderMode').inputValue(),'auto');assert(!(await page.locator('#simpleMaskHeaders').isChecked()));assert(!(await page.locator('#simplePreservePreamble').isChecked()));assert((await page.locator('#simpleHeaderInfo').innerText()).includes('第 4 行'));
  await page.locator('#simplePreview').click();await page.locator('#simpleRows tr').first().waitFor();assert.equal(await page.locator('#simpleRows tr').count(),100);assert(await page.locator('#simpleCopy').isDisabled());assert((await page.locator('#simpleRows').innerText()).includes('<script>alert(1)</script>'));
  await page.locator('#simpleNext').click();await page.locator('#simplePageInfo').filter({hasText:'101–200 / 205'}).waitFor();await page.locator('#simpleNext').click();await page.locator('#simplePageInfo').filter({hasText:'201–205 / 205'}).waitFor();assert.equal(await page.locator('#simpleRows tr').count(),5);
  await page.locator('#simpleSearch').fill('example');await page.locator('#simpleFind').click();await page.locator('#simplePageInfo').filter({hasText:'1–1 / 1'}).waitFor();
  await page.locator('#simpleConfirmed').check();await page.locator('#simpleCopy').click();await page.locator('#status').filter({hasText:'已复制'}).waitFor();assert.equal(await application.evaluate(async({clipboard})=>await clipboard.readText()),'XJPYQYXGS\t56234P1\r\n');
  assert(await page.locator('#simpleApply').isDisabled());await page.locator('#simpleBackup').check();await page.locator('#simpleApply').click();await page.locator('#status').filter({hasText:'已写入 203'}).waitFor();assert(await page.locator('#simpleApply').isDisabled());
  assert(await page.locator('#simpleRecovery').isDisabled());await page.locator('#simpleSensitive').check();await application.evaluate(({dialog},temp)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath:temp+'/recovery.json'});},temp);await page.locator('#simpleRecovery').click();await page.locator('#status').filter({hasText:'已导出含原文'}).waitFor();
  await application.evaluate(({dialog})=>{dialog.showSaveDialog=async()=>({canceled:true});});await page.locator('#simpleExport').click();await page.locator('#status').filter({hasText:'已取消保存'}).waitFor();
  await page.locator('#simplePrefix').fill('客户-');assert(await page.locator('#simpleCopy').isDisabled());assert.equal(await page.locator('#simpleRows tr').count(),0);
  await page.locator('#simpleMode').selectOption('swap');await page.getByRole('button',{name:'首尾各 3 位',exact:true}).click();assert.equal(await page.locator('#simpleEdge').inputValue(),'3');await page.locator('#simplePreview').click();await page.locator('#simpleRows tr').first().waitFor();assert.equal(lastMessage.options.EdgeCount,3);
  await page.locator('#simpleMode').selectOption('number');await page.locator('#simpleOperator').selectOption('/');await page.locator('#simpleOperand').fill('2.5');await page.locator('#simplePreview').click();await page.locator('#simpleRows tr').first().waitFor();assert.equal(lastMessage.options.Operand,'2.5');
  await page.locator('#simpleMode').selectOption('regex');await page.locator('#simplePattern').fill('[');await page.locator('#simplePreview').click();await page.locator('#status.error').waitFor();assert(await page.locator('#simpleCopy').isDisabled());await page.getByRole('button',{name:'首尾各两位模板',exact:true}).click();await page.locator('#simplePreview').click();await page.locator('#simpleRows tr').first().waitFor();assert.equal(lastMessage.options.Replacement,'$3$2$1');
  await page.locator('#simpleMode').selectOption('initials');await page.locator('#simplePrefix').fill('');await page.locator('#simplePreview').click();await page.locator('#simpleRows tr').first().waitFor();
  await page.locator('#simpleMaskHeaders').check();assert(await page.locator('#simpleCopy').isDisabled());await page.locator('#simpleHeaderMode').selectOption('manual');await page.locator('#simpleHeaderRows').fill('4-5');await page.locator('#simplePreservePreamble').check();await page.locator('#simplePreview').click();await page.locator('#simpleRows tr').first().waitFor();assert.equal(lastMessage.options.HeaderMode,'manual');assert.equal(lastMessage.options.HeaderRows,'4-5');assert(lastMessage.options.MaskHeaders&&lastMessage.options.PreservePreamble);
  await page.locator('#simpleHeaderMode').selectOption('auto');await page.locator('#simpleMaskHeaders').uncheck();await page.locator('#simplePreservePreamble').uncheck();await page.locator('#simplePreview').click();await page.locator('#simpleRows tr').first().waitFor();
  await page.setViewportSize({width:800,height:620});assert(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth));await page.setViewportSize({width:1180,height:900});fs.mkdirSync(path.join(root,'docs/v1.3'),{recursive:true});await page.screenshot({path:path.join(root,'docs/v1.3/simple-mask.png'),fullPage:true});
  await page.locator('#simpleCapture').click();await page.locator('#simpleSummary').filter({hasText:'选区已更新'}).waitFor();assert(await page.locator('#simpleCopy').isDisabled());assert.deepEqual(errors,[]);
  console.log('PASS simple masking UI: preview gates, pagination/search, escaped source, native clipboard, backup gate, one-shot apply, recovery gate, canceled save, rule invalidation, quick buttons, arithmetic, invalid regex recovery, narrow layout, rescan; '+received.length+' host requests');
 }finally{if(application)await application.close();socket?.destroy();await new Promise(r=>server.close(r));fs.rmSync(temp,{recursive:true,force:true,maxRetries:5,retryDelay:200});}
}
main().catch(e=>{console.error(e);process.exitCode=1;});
