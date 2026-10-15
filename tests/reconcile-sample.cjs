// Read-only acceptance. Raw data and local paths are never written into release documentation.
const fs=require('node:fs'),path=require('node:path'),cp=require('node:child_process'),crypto=require('node:crypto'),assert=require('node:assert/strict');
const {suggest}=require('../App/reconcile-preset.cjs'),{table,reconcile,money}=require('../App/reconcile-core.cjs');
const file=process.argv[2];if(!file)throw Error('Pass the local sanitized sample path.');
const digest=()=>crypto.createHash('sha256').update(fs.readFileSync(file)).digest('hex'),before=digest(),reports={};
for(const [role,sheet] of Object.entries({previous:'上期余额',current:'当期余额',issue:'发放表202609',recovery:'回收表202609'})){
  const raw=cp.execFileSync(path.resolve('tests/SheetPilot.Tests/bin/Release/net462/SheetPilot.Tests.exe'),['--reconcile-json',file,sheet],{encoding:'utf8',maxBuffer:64*1024*1024});reports[role]=JSON.parse(raw.replace(/^\uFEFF/,''));
}
const c=suggest(reports,file);assert.equal(c.sources.length,4);assert.equal(c.currentDate,'2026-09-30');assert.equal(c.previousDate,'2026-08-31');
const tables={};for(const b of c.rules[0].bindings){const s=c.sources.find(s=>s.id===b.source);tables[s.id]=table(reports[b.role],s);}
assert.equal(reconcile(c,tables).results[0].status,'error');c.rules[0].reviewed=true;
const result=reconcile(c,tables);assert.notEqual(result.results[0].status,'error');
for(const b of c.rules[0].bindings){
  const t=tables[b.source],seen=new Set();let sum=0n,invalid=0,duplicates=0;
  for(const row of t.rows){const key=JSON.stringify(b.key.map(k=>String(row[k]??'').trim()));if(b.key.some(k=>!String(row[k]??'').trim())){invalid++;continue;}if(seen.has(key)){duplicates++;continue;}seen.add(key);try{sum+=money(row[b.amount]);}catch{invalid++;}}
  assert.equal(result.results[0].tot[b.role],sum.toString());console.log(`${b.role}: rows=${t.rows.length}, duplicateKeys=${duplicates}, invalidRows=${invalid}, independent total verified`);
}
assert.equal(digest(),before);const summary={status:result.results[0].status,rows:Object.values(tables).reduce((n,t)=>n+t.rows.length,0),groups:result.results[0].details.length,issues:result.issues.length,sourceUnchanged:true};
console.log(JSON.stringify(summary));
fs.mkdirSync('.temp/reconcile',{recursive:true});fs.writeFileSync('.temp/reconcile/sample-config.json',JSON.stringify(c,null,2));fs.writeFileSync('.temp/reconcile/sample-summary.json',JSON.stringify(summary,null,2));
