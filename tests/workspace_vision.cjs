// Exercise both shipped SDK serializers and the multimodal lifting path.
const fs=require('node:fs'),path=require('node:path'),vm=require('node:vm'),assert=require('node:assert/strict');
const sdk=fs.readFileSync(path.resolve('dist/Chatbox/native/dist/renderer/js/vendor-ai.CMAD2Bsd.js'),'utf8');
const c=vm.createContext({URL,TextDecoder,Ke:Error,Ut:x=>x});
function part(start,end){const a=sdk.indexOf(start),b=sdk.indexOf(end,a+start.length);assert(a>=0&&b>a,start);return sdk.slice(a,b)}
vm.runInContext(part('function ur(e)','function fp(')+part('function _D(e)','function Ds(')+fs.readFileSync('integrations/chatbox/workspace-vision.js','utf8'),c);
const data=Buffer.from('test pixels').toString('base64');
c.messages=[{role:'assistant',content:[{type:'tool-call',toolCallId:'a',toolName:'view',input:{}},{type:'tool-call',toolCallId:'b',toolName:'read',input:{}}]},{role:'tool',content:[{type:'tool-result',toolCallId:'a',toolName:'view',output:{type:'content',value:[{type:'text',text:'page 1'},{type:'image-data',mediaType:'image/jpeg',data}]}},{type:'tool-result',toolCallId:'b',toolName:'read',output:{type:'json',value:{text:'revenue'}}}]}];
for(const expression of ['FO(messages)','yD({prompt:messages}).messages']){
 const result=vm.runInContext(expression,c);assert.deepEqual(Array.from(result,x=>x.role),['assistant','tool','tool','user']);
 assert(!result[1].content.includes(data));assert.equal(result[3].content[1].image_url.url,'data:image/jpeg;base64,'+data);
}
console.log('PASS production OpenAI + compatible serializers: real image blocks, no base64 text, paired parallel tool results');
if(process.argv.includes('--transform')){let input='';process.stdin.on('data',s=>input+=s);process.stdin.on('end',()=>{c.input=JSON.parse(input);process.stdout.write(JSON.stringify(vm.runInContext('gridpilotLiftToolImages(input)',c)));});}
