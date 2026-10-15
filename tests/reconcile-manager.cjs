const assert=require('node:assert/strict'),fs=require('node:fs'),os=require('node:os'),path=require('node:path');
const {ReconcileManager}=require('../App/reconcile-main.cjs');
const root=fs.mkdtempSync(path.join(os.tmpdir(),'GridPilot-reconcile-manager-')),file=path.join(root,'source.xlsx');fs.writeFileSync(file,'synthetic');
const config={schema:1,name:'test',previousDate:'2026-08-31',currentDate:'2026-09-30',tolerance:'0.01',displayUnit:'1',sources:['p','c'].map(id=>({id,path:file,sheet:id,name:id,headerRow:1,startRow:2,endRow:0,headers:['id','amount']})),rules:[{id:'r',name:'test',kind:'balance',bindings:['previous','current'].map((role,i)=>({role,source:i?'c':'p',key:['id'],group:['id'],amount:'amount',unit:'1',sign:1,filters:[],joins:[]}))}]};
let fail=false,hold=false,releaseRead;
const host=async m=>{if(hold)await new Promise(r=>releaseRead=r);if(fail&&m.sheet==='c')throw Error('synthetic read error');return {sheet:m.sheet,rows:[{row:1,values:['id','amount']},{row:2,values:['=untrusted',m.sheet==='c'?'5':'10']}]};};
const manager=new ReconcileManager(host,()=>{});manager.allowed.add(file);
(async()=>{
  let r=await manager.handle({action:'reconcileRun',config});assert.equal(r.results[0].status,'unverified');assert.equal(r.manifest.length,2);assert(r.config.rules);console.log('PASS manager worker and audit snapshot');
  fail=true;r=await manager.handle({action:'reconcileRun',config});assert.equal(r.results[0].status,'error');assert(r.manifest.find(m=>m.error));fail=false;console.log('PASS failed source isolated and reported');
  hold=true;const running=manager.handle({action:'reconcileRun',config});await new Promise(resolve=>{const timer=setInterval(()=>{if(releaseRead){clearInterval(timer);resolve();}},5);});manager.cancel();releaseRead();await assert.rejects(running,/取消/);assert.equal(manager.last,null);assert.equal(manager.active,false);hold=false;console.log('PASS cancellation clears stale result and allows retry');
  r=await manager.handle({action:'reconcileRun',config});assert.equal(r.results[0].status,'unverified');
  const csvFile=path.join(root,'audit.csv');await manager.handle({action:'reconcileExport',kind:'csv'},{showSaveDialog:async()=>({filePath:csvFile})});const csv=fs.readFileSync(csvFile,'utf8');assert(csv.includes('"-5.000000"'));assert(!csv.includes("'-5.000000"));assert(csv.includes("'=untrusted"));console.log('PASS CSV keeps negative amounts numeric and neutralizes text formulas');
  await assert.rejects(manager.handle({action:'reconcileOpenSource',ref:{file:'not-allowed.xlsx',sheet:'p',row:1}}),/来源/);
  const bad=structuredClone(config);bad.sources[1].id='p';await assert.rejects(manager.handle({action:'reconcileRun',config:bad}),/重复/);assert.equal(manager.last,null);console.log('PASS invalid plan and unrelated source rejected');
})().catch(e=>{console.error(e);process.exitCode=1;}).finally(()=>{manager.cancel();assert.equal(path.dirname(path.resolve(root)),path.resolve(os.tmpdir()));assert(path.basename(root).startsWith('GridPilot-reconcile-manager-'));fs.rmSync(root,{recursive:true,force:true});});
