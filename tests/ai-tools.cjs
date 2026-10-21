const assert=require('node:assert/strict');
const {tools,handleTool}=require('../integrations/chatbox/tools.cjs');
(async()=>{
  assert.equal(tools.length,9);assert(!tools.some(t=>/apply|undo/.test(t.name)));
  let calls=0,last,published;
  const host=async request=>{calls++;last=request;return {total:2,cells:[],proposalId:'proposal'};};
  const publish=value=>published=value;
  const snapshotId='a'.repeat(32);
  let result=await handleTool({name:'aiApply',arguments:{snapshotId}},host,publish);
  assert(result.isError);assert.equal(calls,0);
  result=await handleTool({name:'excel_read',arguments:{snapshotId:'another book'}},host,publish);
  assert(result.isError);assert.equal(calls,0);
  await handleTool({name:'excel_read',arguments:{snapshotId,action:'aiApply',operation:'delete',offset:200}},host,publish);
  assert.equal(last.action,'aiRead');assert.equal(last.operation,'read');assert.equal(last.snapshotId,snapshotId);assert.equal(last.offset,200);
  await handleTool({name:'excel_propose_formula',arguments:{snapshotId,sheet:'Sales',address:'D2',formula:'=SUM(B2:B5)',action:'aiApply'}},host,publish);
  assert.equal(last.action,'aiPropose');assert.equal(published.proposalId,'proposal');
  result=await handleTool({name:'excel_statistics',arguments:{snapshotId}},async()=>{throw Error('expired');},publish);
  assert(result.isError);assert.equal(result.content[0].text,'expired');
  for(const [name,args,action] of [
    ['excel_propose_values',{snapshotId,sheet:'Sales',address:'A1:B2',data:[[1,'=literal'],[null,true]]},'aiPropose'],
    ['excel_propose_formulas',{snapshotId,sheet:'Sales',address:'A1',data:[['=1+1']]},'aiPropose'],
    ['excel_propose_format',{snapshotId,sheet:'Sales',address:'A1:B2',style:{bold:true}},'aiPropose'],
    ['excel_propose_chart',{snapshotId,sheet:'Sales',address:'A1:B2',chartType:'column'},'aiPropose'],
    ['excel_inspect',{snapshotId},'aiInspect']]) { const r=await handleTool({name,arguments:args},host,publish);assert(!r.isError);assert.equal(last.action,action); }
  console.log('PASS Excel MCP contract: no mutation tools, snapshot validation, allowlisted parameters, proposal-only writes, errors');
})().catch(error=>{console.error(error);process.exitCode=1;});
