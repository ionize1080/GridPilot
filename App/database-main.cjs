const fs=require('node:fs'),path=require('node:path'),crypto=require('node:crypto');
const {fork}=require('node:child_process');
const {filesIn,info}=require('./database-core.cjs');
class DatabaseManager{
  constructor(root,host,progress){this.root=root;this.host=host;this.progress=progress;this.active=new Set();this.scanning=false;fs.mkdirSync(root,{recursive:true});}
  filename(id){if(!/^[a-f0-9-]{36}$/.test(id))throw Error('数据库标识无效。');const file=path.join(this.root,id+'.sqlite');if(!fs.existsSync(file))throw Error('数据库不存在。');return file;}
  job(action,filename,extra={}){return new Promise((resolve,reject)=>{
    const env={...process.env,ELECTRON_RUN_AS_NODE:'1'};delete env.NODE_OPTIONS;
    const worker=fork(path.join(__dirname,'database-worker.cjs'),[],{env,execArgv:['--max-old-space-size=512'],windowsHide:true,stdio:['ignore','ignore','ignore','ipc'],serialization:'advanced'});
    worker.send({type:'start',action,filename,...extra});
    let settled=false;const finish=async(err,data)=>{if(settled)return;settled=true;clearTimeout(timer);this.active.delete(cancel);await new Promise(done=>{if(worker.exitCode!==null||worker.signalCode!==null)return done();worker.once('exit',done);worker.kill();});err?reject(err):resolve(data);};
    const cancel=()=>finish(Error('已取消，未提交的扫描已回滚。'));this.active.add(cancel);
    const timer=setTimeout(()=>finish(Error(action==='query'?'查询超过 15 秒，请缩小范围。':'扫描超过 15 分钟，请拆分文件夹后重试。')),action==='query'?15000:900000);
    worker.on('error',e=>finish(e));worker.on('exit',code=>{if(!settled)finish(Error('数据库任务中断（'+code+'）。上次已提交数据仍保留。'));});
    worker.on('message',async m=>{if(m.type==='result')finish(null,m.data);else if(m.type==='error')finish(Error(m.error));else if(m.type==='progress')this.progress(m);else if(m.type==='read'){
      try{const data=await this.host({action:'dbReadFile',path:m.file,sheet:m.sheet,maxRows:0});if(!settled)worker.send({id:m.id,data});}
      catch(e){if(!settled)worker.send({id:m.id,error:e.message});}
    }});
  });}
  cancel(){for(const stop of [...this.active])stop();}
  async handle(m,dialog,win){
    if(m.action==='dbList')return {root:this.root,databases:fs.readdirSync(this.root).filter(n=>/^[a-f0-9-]{36}\.sqlite$/.test(n)).flatMap(n=>{try{const d=info(path.join(this.root,n));return [{id:n.slice(0,-7),...d}];}catch{return [];}})};
    if(m.action==='dbChoose'){
      const choice=await dialog.showOpenDialog(win,{title:'选择同一类报表所在文件夹',properties:['openDirectory']});if(choice.canceled)return {canceled:true};
      const folder=choice.filePaths[0];this.chosenFolder=folder;const files=filesIn(folder,!!m.recursive);if(!files.length)throw Error('所选文件夹没有 .xlsx / .xlsm 报表。');
      this.sample=files[0];return {folder,sample:path.basename(this.sample),...await this.host({action:'dbReadFile',path:this.sample,maxRows:100})};
    }
    if(m.action==='dbSample'){if(!this.sample)throw Error('请先选择文件夹。');return this.host({action:'dbReadFile',path:this.sample,sheet:m.sheet,maxRows:100});}
    if(m.action==='dbCreate'){
      if(!this.chosenFolder||m.config.folder!==this.chosenFolder)throw Error('请先通过文件夹选择按钮选择来源。');
      const id=crypto.randomUUID(),filename=path.join(this.root,id+'.sqlite');
      try{await this.job('create',filename,{config:m.config});return {id,...info(filename)};}catch(e){for(const suffix of ['', '-wal','-shm'])try{fs.unlinkSync(filename+suffix);}catch{}throw e;}
    }
    const filename=this.filename(m.id);
    if(m.action==='dbInfo')return this.job('info',filename);
    if(m.action==='dbScan'){
      if(this.scanning)throw Error('已有扫描正在运行，请等待完成。');this.scanning=true;
      try{return await this.job('scan',filename);}finally{this.scanning=false;}
    }
    if(m.action==='dbQuery')return this.job('query',filename,{sql:m.sql,offset:m.offset,pageSize:m.pageSize});
    if(m.action==='dbOpenSource'){
      const d=info(filename),file=d.files.find(f=>f.path===m.file);if(!file)throw Error('来源文件不属于当前数据库。');
      if(!fs.existsSync(file.path))throw Error('源文件已移走，请重新扫描。');
      return this.host({action:'dbOpenSource',path:file.path,sheet:m.sheet||d.config.sheet,row:Number(m.row)||d.config.headerRow+1});
    }
    throw Error('未知数据库操作。');
  }
}
module.exports={DatabaseManager};
