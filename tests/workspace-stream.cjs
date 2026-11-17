const fs=require('node:fs'),vm=require('node:vm'),assert=require('node:assert/strict'),path=require('node:path');
const root=path.resolve(__dirname,'..'),sdk=fs.readFileSync(path.join(root,'dist/Chatbox/native/dist/renderer/js/vendor-ai.CMAD2Bsd.js'),'utf8');
const helper=fs.readFileSync(path.join(root,'integrations/chatbox/workspace-stream.js'),'utf8');
function loop(prefix){const start=sdk.indexOf(prefix);assert(start>=0,prefix);let depth=0,quote='',escape=false;for(let i=start;i<sdk.length;i++){const c=sdk[i];if(quote){if(escape){escape=false;continue;}if(c==='\\'){escape=true;continue;}if(c===quote)quote='';continue;}if('"\'`'.includes(c)){quote=c;continue;}if(c==='{')depth++;if(c==='}'&&--depth===0)return sdk.slice(start,i+1);}throw Error('Unclosed SDK loop');}
for(const [prefix,container,storage,controller] of [
 ['for(let ue of U.tool_calls)','U','c','b'],['for(let W of E.tool_calls)','E','s','_'],['for(let H of Z.tool_calls)','Z','i','_'],['for(let E of A.tool_calls)','A','i','h'],
]){
 const context=vm.createContext({events:[],state:[],Error,Map,WeakMap});
 vm.runInContext(helper+`;const lo=Error,Rs=Error,Ve=()=>"generated",Nt=Ve;const Ho=value=>{try{JSON.parse(value);return true;}catch{return false;}};const Bm=Ho;function consume(calls){var x,O,R,w,C,F,D,k,V,M,Z,J,H,z,A,E,y,b,S,I;var ${container}={tool_calls:calls},${storage}=state,${controller}={enqueue:e=>events.push(e)};${loop(prefix)}}`,context);
 const feed=call=>context.consume([call]);
 feed({index:0,id:'call-1',type:'function'});assert.equal(context.events.length,0);
 feed({index:1,id:'call-2',type:'function',function:{name:'workspace_list',arguments:'{}'}});
 feed({index:0,function:{name:'workspace_',arguments:''}});assert.equal(context.events.filter(e=>e.toolCallId==='call-1').length,0);
 feed({index:0,function:{name:'read',arguments:'{"file_id":'}});
 feed({index:0,function:{arguments:'"abc"}'}});
 const calls=context.events.filter(e=>e.type==='tool-call');assert.equal(calls.length,2);assert.equal(calls[1].toolName,'workspace_read');assert.equal(calls[1].input,'{"file_id":"abc"}');
 vm.runInContext('gridpilotFinishToolDeltas(state)',context);
 const broken=[];context.broken=broken;vm.runInContext('gridpilotToolDelta({id:"incomplete",index:0},broken,0)',context);assert.throws(()=>vm.runInContext('gridpilotFinishToolDeltas(broken)',context),/不完整/);
 console.log('PASS shipped streaming loop',prefix,'split names, interleaved calls, exact arguments, unfinished stream rejection');
}
