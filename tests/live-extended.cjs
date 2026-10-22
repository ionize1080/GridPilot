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
let usage=0;
async function task(prompt,expected,validate){
  const context=await host({action:'aiCapture',scope:'workbook'}),seen=[];
  const messages=[{role:'system',content:'You are testing Excel tools on synthetic data. Execute the user request via excel_help then excel_operation. Read exact help first. Never use old excel_propose tools in this test. Do not claim success unless applied=true. VBA must never be run. Respond briefly. Context: '+JSON.stringify(context)},{role:'user',content:prompt}];
  for(let round=0;round<9;round++){
    const response=await fetch('https://api.deepseek.com/chat/completions',{method:'POST',headers:{Authorization:'Bearer '+key,'Content-Type':'application/json'},body:JSON.stringify({model,messages,tools:tools.filter(t=>['excel_help','excel_operation'].includes(t.name)).map(t=>({type:'function',function:{name:t.name,description:t.description,parameters:t.inputSchema}})),temperature:0,max_tokens:1600}),signal:AbortSignal.timeout(90000)});
    if(!response.ok)throw Error('Provider status '+response.status);const body=await response.json();usage+=body.usage?.total_tokens||0;
    const message=body.choices?.[0]?.message;if(!message)throw Error('No provider message');messages.push(message);
    if(!message.tool_calls?.length)break;
    for(const call of message.tool_calls){let args;try{args=JSON.parse(call.function.arguments);}catch{throw Error('Invalid model tool arguments');}
      const result=await handleTool({name:call.function.name,arguments:args},host,()=>{});
      messages.push({role:'tool',tool_call_id:call.id,content:JSON.stringify(result)});if(result.isError)console.log('Tool error '+String(result.content?.[0]?.text).split(key).join('[REDACTED]')); 
      if(call.function.name==='excel_operation'&&!result.isError){const data=JSON.parse(result.content[0].text);seen.push({operation:args.operation,data});console.log('Tool '+args.operation+' applied='+data.applied);}
    }
    if(expected.every(op=>seen.some(v=>v.operation===op))){await validate(seen);console.log('PASS paid model '+expected.join(', '));return;}
  }
  throw Error('Model did not complete expected operations: '+expected.join(', '));
}
(async()=>{
  try{
    await host({action:'aiApproval',enabled:true,scope:'workbook'});
    await task('创建名为“汇总”的工作表。',['sheet.create'],async seen=>assert(seen[0].data.applied));
    await task('以销售!A1:C4为源，在汇总!A3创建名为TestPivot的原生透视表，项目作行，收入求和，数据字段名称Total。',['pivot.create'],async()=>assert.equal((await host({action:'testPivot',sheet:'汇总'})).total,450));
    await task('将汇总的TestPivot按项目名称降序排列，并且筛选项目仅保留甲和乙。',['pivot.sort','pivot.filter'],async()=>assert.equal((await host({action:'testPivot',sheet:'汇总'})).total,300));
    await task('在销售表用A1:C4创建名为LiveChart的带数据标记折线图，然后将主值轴最大值改为300。',['chart.create','chart.edit'],async()=>assert.equal((await host({action:'testRead',address:'A1'})).charts,1));
    await task('必须查询并使用 plugin.rename 操作，将汇总表加前缀“测试_”，不要改变销售表的名称。',['plugin.rename'],async seen=>assert(seen.at(-1).data.applied));
    await task('将销售!A1:C4按收入降序排序，有表头，使用自定义多级排序接口。',['range.sort'],async()=>assert.equal((await host({action:'testRead',address:'B2'})).value,200));
    await task('仅准备 VBA 写入建议，在RiskDemo标准模块中写入 Sub Demo() 换行 Shell "cmd /c echo test" 换行 End Sub。不要运行。即使自动批准开启也必须等待人工批准。',['vba.write'],async seen=>{assert.equal(seen[0].data.applied,false);assert.equal(seen[0].data.requiresManualApproval,true);});
    console.log('PASS live extended integration; model='+model+'; totalTokens='+usage+'; credential withheld');
  }finally{child.stdin.end();for(const p of pending.values())clearTimeout(p.timer);pending.clear();}
})().catch(e=>{console.error(String(e.message).split(key).join('[REDACTED]'));process.exitCode=1;child.stdin.end();});
