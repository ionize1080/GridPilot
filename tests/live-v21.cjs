// Explicit paid-model integration against a new synthetic Excel instance.
// The credential is read into memory, never placed in child args/env/logs.
const fs=require('node:fs'),path=require('node:path'),assert=require('node:assert/strict');
const {spawn}=require('node:child_process'),readline=require('node:readline');
const {tools,handleTool}=require('../integrations/chatbox/tools.cjs');
const root=path.resolve(__dirname,'..');
const key=fs.readFileSync(process.env.GRIDPILOT_TEST_KEY_FILE,'utf8').match(/sk-[A-Za-z0-9_-]+/)?.[0];
if(!key)throw Error('Credential format unavailable; contents withheld');
const model=process.env.GRIDPILOT_TEST_MODEL||'deepseek-flash';
const env={...process.env};delete env.GRIDPILOT_TEST_KEY_FILE;
const child=spawn(path.join(root,'tests/SheetPilot.Tests/bin/Release/net462/SheetPilot.Tests.exe'),['--ai-live-bridge'],{stdio:['pipe','pipe','ignore'],windowsHide:true,env});
let sequence=0;const pending=new Map();
readline.createInterface({input:child.stdout}).on('line',line=>{try{const msg=JSON.parse(line.replace(/^\uFEFF/,'')),p=pending.get(msg.id);if(p){clearTimeout(p.timer);pending.delete(msg.id);msg.error?p.reject(Error(msg.error)):p.resolve(msg.data);}}catch{}});
function host(request){return new Promise((resolve,reject)=>{const id=++sequence,timer=setTimeout(()=>reject(Error('Excel timeout')),120000);pending.set(id,{resolve,reject,timer});child.stdin.write(JSON.stringify({...request,id})+'\n');});}
let usage=0,currentScope="auto";
async function task(prompt,expected,validate){
  const context=await host({action:'aiCapture',scope:currentScope}),seen=[];
  const messages=[{role:'system',content:'You are testing Excel tools on synthetic data. Follow context.workflow. Review the ENTIRE directory. FIRST call excel_context with the directory snapshotId and selected bookId; use returned child snapshotId for EVERY data/edit tool, never reuse directory ID for them. Use exact local statistics for totals, not samples; include hidden sheets when relevant. Read excel_help before unfamiliar operations. Use capability.list before modern operations. Do not claim success unless applied=true. VBA must never be run. Respond briefly. Context: '+JSON.stringify(context)},{role:'user',content:prompt}];
  for(let round=0;round<16;round++){
    const response=await fetch('https://api.deepseek.com/chat/completions',{method:'POST',headers:{Authorization:'Bearer '+key,'Content-Type':'application/json'},body:JSON.stringify({model,messages,tools:tools.map(t=>({type:'function',function:{name:t.name,description:t.description,parameters:t.inputSchema}})),temperature:0,max_tokens:1600}),signal:AbortSignal.timeout(90000)});
    if(!response.ok)throw Error('Provider status '+response.status);const body=await response.json();usage+=body.usage?.total_tokens||0;
    const message=body.choices?.[0]?.message;if(!message)throw Error('No provider message');messages.push(message);
    if(!message.tool_calls?.length){if(!expected.length){await validate(seen,message.content);console.log('PASS paid model cross-workbook analysis');return;}break;}
    for(const call of message.tool_calls){let args;try{args=JSON.parse(call.function.arguments);}catch{throw Error('Invalid model tool arguments');}
      const result=await handleTool({name:call.function.name,arguments:args},host,()=>{});
      messages.push({role:'tool',tool_call_id:call.id,content:JSON.stringify(result)});if(result.isError)console.log('Tool error '+String(result.content?.[0]?.text).split(key).join('[REDACTED]'));
      if(['excel_operation','excel_context'].includes(call.function.name)&&!result.isError){const data=JSON.parse(result.content[0].text);seen.push({operation:call.function.name==='excel_context'?'context.'+(args.mode||'capture'):args.operation,args:args.options,data});console.log('Tool '+seen.at(-1).operation+' applied='+data.applied);}
    }
    if(expected.length&&expected.every(op=>seen.some(v=>v.operation===op))){await validate(seen);console.log('PASS paid model '+expected.join(', '));return;}
  }
  throw Error('Model did not complete expected operations: '+expected.join(', '));
}
(async()=>{
  try{
    const books=await host({action:'testSetupV21'});
    await task('核对当前打开的所有工作簿：汇总每一本销售表的收入，另外加上隐藏的调整表金额，最后读取 RevenuePolicy 模块中的 RevenueFactor 系数，并计算调整后的最终金额。使用工具核验所有数据，说明来源，不要写入。',[],async(seen,answer)=>{
      const captures=seen.filter(v=>v.operation==='context.capture');assert(captures.length>=2,'both workbooks captured');
      assert(seen.some(v=>v.operation==='vba.read'),'VBA read');
      assert(/729/.test(answer),'(450+200+25)*1.08 = 729');
      assert(captures.some(v=>v.data.ranges?.some(r=>r.sheet==='调整')),'hidden adjustment captured');
    });
    currentScope='all';
    await task('仅核对所有打开工作簿的销售表收入：分别报告每一本合计及总收入。必须覆盖目录中的每本工作簿，不含调整表，不要修改。',[],async(seen,answer)=>{assert(new Set(seen.filter(v=>v.operation==='context.capture').map(v=>v.data.book)).size===2);assert(/650/.test(answer));});
    await host({action:'testActivate'});await host({action:'aiApproval',enabled:true,scope:'workbook'});currentScope='auto';
    await task('在含 RevenuePolicy 的工作簿的销售表 G2 插入动态数组公式 =SEQUENCE(3)，使用 range.formula2，先核验版本能力。',['range.formula2'],async seen=>{assert(seen.find(v=>v.operation==='range.formula2').data.applied);assert.equal((await host({action:'testRead',address:'G4'})).value,3);});
    await task('在含 RevenuePolicy 的工作簿，以销售!A1:B4 创建名为 ModernLive 的瀑布图。使用 chart.create，先核验当前 Excel 支持。',['chart.create'],async seen=>{assert(seen.find(v=>v.operation==='chart.create').data.applied);assert.equal((await host({action:'testRead',address:'A1'})).charts,1);});
    await task('在含 RevenuePolicy 的工作簿中创建 Power Query 定义 LiveQuery：let Source = #table({"A"}, {{1}}) in Source。仅保存定义，不刷新；然后读取定义核验。',['query.set'],async seen=>{const set=seen.find(v=>v.operation==='query.set').data;assert(set.applied);const verified=await host({action:'aiExtended',snapshotId:set.snapshotId,operation:'query.list',options:{}});assert(verified.items.some(q=>q.name==='LiveQuery'));});
    await task('在含 RevenuePolicy 的工作簿的销售!A1 添加线程批注“实际 API 验收”，然后必须调用 comment.reply 添加回复“通过”，两步都要完成。',['comment.add','comment.reply'],async seen=>{assert(seen.filter(v=>['comment.add','comment.reply'].includes(v.operation)).every(v=>v.data.applied));});
    console.log('PASS live v2.1 integration; model='+model+'; totalTokens='+usage+'; credential withheld');
  }finally{child.stdin.end();for(const p of pending.values())clearTimeout(p.timer);pending.clear();}
})().catch(e=>{console.error(String(e.message).split(key).join('[REDACTED]'));process.exitCode=1;child.stdin.end();});
