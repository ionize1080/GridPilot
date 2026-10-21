// Official MIT MCP SDK. stdout is reserved exclusively for MCP protocol messages.
const {Server}=require('@modelcontextprotocol/sdk/server/index.js');
const {StdioServerTransport}=require('@modelcontextprotocol/sdk/server/stdio.js');
const {ListToolsRequestSchema,CallToolRequestSchema}=require('@modelcontextprotocol/sdk/types.js');
const net=require('node:net'),readline=require('node:readline');
const {tools}=require('./tools.cjs');
const pipe=process.env.GRIDPILOT_MCP_PIPE,token=process.env.GRIDPILOT_MCP_TOKEN;
if(!/^\\\\\.\\pipe\\GridPilot-MCP-[a-f0-9]{32}$/.test(pipe||'')||!token)process.exit(1);
const socket=net.createConnection(pipe),pending=new Map();let sequence=0;
function call(params){return new Promise((resolve,reject)=>{
  const id=++sequence,timer=setTimeout(()=>{pending.delete(id);reject(Error('Excel操作超时，请重试。'));},600000);
  pending.set(id,{resolve,reject,timer});socket.write(JSON.stringify({id,token,params})+'\n');
});}
readline.createInterface({input:socket}).on('line',line=>{
  try{const m=JSON.parse(line),p=pending.get(m.id);if(!p)return;pending.delete(m.id);clearTimeout(p.timer);m.error?p.reject(Error(m.error)):p.resolve(m.result);}catch{}
});
socket.on('error',()=>process.exit(1));socket.on('close',()=>process.exit(0));
const server=new Server({name:'GridPilot Excel',version:'1.8.0'},{capabilities:{tools:{}}});
server.setRequestHandler(ListToolsRequestSchema,async()=>({tools}));
server.setRequestHandler(CallToolRequestSchema,async req=>call(req.params));
server.connect(new StdioServerTransport()).catch(()=>process.exit(1));
process.stdin.on('end',()=>{socket.destroy();process.exit(0);});
