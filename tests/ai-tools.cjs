const assert=require('node:assert/strict');
const {tools,handleTool,permitted}=require('../integrations/chatbox/tools.cjs');
(async()=>{
  assert.equal(tools.length,13);assert(!tools.some(t=>/apply|undo/.test(t.name)));
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
  result=await handleTool({name:'excel_help',arguments:{operation:'pivot.field'}},host,publish);assert(!result.isError);assert(JSON.parse(result.content[0].text).options.includes('orientation'));
  const before=calls;result=await handleTool({name:'excel_operation',arguments:{snapshotId,operation:'aiApproval',options:{enabled:true}}},host,publish);assert(result.isError);assert.equal(calls,before);
  await handleTool({name:'excel_operation',arguments:{snapshotId,operation:'vba.write',options:{module:'Test',code:'Shell "cmd"'},action:'aiApply'}},host,publish);assert.equal(last.action,'aiExtended');assert.equal(last.operation,'vba.write');
  assert(!permitted('excel_propose_values',{}, {GRIDPILOT_MCP_READ_ONLY:'1'}));assert(!permitted('excel_operation',{operation:'sheet.delete'},{GRIDPILOT_MCP_READ_ONLY:'1'}));assert(permitted('excel_operation',{operation:'vba.read'},{GRIDPILOT_MCP_READ_ONLY:'1'}));assert(!permitted('excel_statistics',{}, {GRIDPILOT_MCP_TOOLS:'excel_read,excel_help'}));
  let binding;
  result=await handleTool({name:'excel_operation',arguments:{snapshotId,operation:'context.capture',options:{bookId:'b'.repeat(32)}}},async r=>{assert.equal(r.action,'aiExtended');return {snapshotId:'c'.repeat(32)};},r=>binding=r);
  assert.equal(binding.parentSnapshotId,snapshotId);assert.equal(binding.snapshotId,'c'.repeat(32));assert(binding.contextBinding);
  result=await handleTool({name:'excel_help',arguments:{snapshotId,topic:'range'}},async r=>{assert.equal(r.operation,'capability.list');return {enabledOperations:['range.calculate']};},()=>{});
  const discovered=JSON.parse(result.content[0].text);assert(!discovered.operations.includes('range.formula2'));assert(discovered.operations.includes('range.calculate'));
  assert(permitted('excel_operation',{operation:'context.capture'},{GRIDPILOT_MCP_READ_ONLY:'1'}));assert(permitted('excel_operation',{operation:'query.list'},{GRIDPILOT_MCP_READ_ONLY:'1'}));assert(!permitted('excel_operation',{operation:'query.set'},{GRIDPILOT_MCP_READ_ONLY:'1'}));
  console.log('PASS Excel MCP contract: compact discovery, no approval/restore tools, validated routing, read-only/filter policies and errors');
})().catch(error=>{console.error(error);process.exitCode=1;});
