const core=require('./database-core.cjs');
let seq=0;const pending=new Map();
process.on('message',m=>{if(m.type==='start'){run(m);return;}const p=pending.get(m.id);if(p){pending.delete(m.id);m.error?p.reject(Error(m.error)):p.resolve(m.data);}});
const send=m=>{if(process.connected)process.send(m);};
const read=(file,sheet)=>new Promise((resolve,reject)=>{const id=++seq;pending.set(id,{resolve,reject});send({type:'read',id,file,sheet});});
async function run(m){try{const {action,filename,config,sql,offset,pageSize}=m;
  const data=action==='create'?core.initialize(filename,config):action==='scan'?await core.scan(filename,read,p=>send({type:'progress',...p})):action==='info'?core.info(filename):core.query(filename,sql,offset,pageSize);
  send({type:'result',data:data??{created:true}});
}catch(e){send({type:'error',error:e.message});}}
process.on('disconnect',()=>process.exit(0));
