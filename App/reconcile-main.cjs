'use strict';
const fs=require('node:fs'),path=require('node:path'),crypto=require('node:crypto');
const {Worker}=require('node:worker_threads');
const {table,format,validateConfig}=require('./reconcile-core.cjs');
class ReconcileManager{
  constructor(host,progress){this.host=host;this.progress=progress;this.allowed=new Set();this.active=false;this.cancelled=false;this.last=null;}
  cancel(){this.cancelled=true;this.worker?.terminate();this.reject?.(Error('已取消本次核对，可调整后重新运行。'));}
  async read(source,maxRows=0){if(!this.allowed.has(source.path))throw Error('请通过选择文件添加来源');return this.host({action:'reconcileRead',path:source.path,sheet:source.sheet,maxRows});}
  async handle(m,dialog,win){
    if(m.action==='reconcileSuggest'){
      const choice=await dialog.showOpenDialog(win,{title:'识别余额底稿中的两期及发放回收表',filters:[{name:'Excel 底稿',extensions:['xlsx','xlsm']}],properties:['openFile']});if(choice.canceled)return {canceled:true};
      const file=choice.filePaths[0];this.allowed.add(file);const first=await this.read({path:file},10),reports={};
      for(const [role,pattern] of Object.entries({previous:/^上期余额$/,current:/^当期余额$/,issue:/^发放表\d*$/,recovery:/^回收表\d*$/})){const sheet=first.sheets.find(s=>pattern.test(s));if(sheet)reports[role]=await this.read({path:file,sheet},30);}
      const config=require('./reconcile-preset.cjs').suggest(reports,file);if(!config.sources.length)throw Error('没有识别到支持的表名和字段，请使用“添加 Excel 来源”手动配置。');return {config};
    }
    if(m.action==='reconcileChoose'){
      const choice=await dialog.showOpenDialog(win,{title:'添加报表来源',filters:[{name:'Excel 报表',extensions:['xlsx','xlsm']}],properties:['openFile']});if(choice.canceled)return {canceled:true};
      const file=choice.filePaths[0];this.allowed.add(file);const report=await this.read({path:file},40);return {path:file,...report};
    }
    if(m.action==='reconcileSample')return this.read(m.source,100);
    if(m.action==='reconcileLoad'){
      const choice=await dialog.showOpenDialog(win,{title:'打开核对方案',filters:[{name:'GridPilot 核对方案',extensions:['json']}],properties:['openFile']});if(choice.canceled)return {canceled:true};
      const file=choice.filePaths[0];if(fs.statSync(file).size>4*1024*1024)throw Error('方案超过 4 MB');const c=validateConfig(JSON.parse(fs.readFileSync(file,'utf8')));for(const s of c.sources){if(!path.isAbsolute(s.path)||! /\.(xlsx|xlsm)$/i.test(s.path))throw Error('来源路径无效');this.allowed.add(s.path);}return {config:c};
    }
    if(m.action==='reconcileSave'){
      const choice=await dialog.showSaveDialog(win,{title:'保存可复用核对方案',defaultPath:'监管核对方案.json',filters:[{name:'核对方案',extensions:['json']}],properties:['showOverwriteConfirmation']});if(choice.canceled)return {canceled:true};
      if(this.allowed.has(choice.filePath))throw Error('不能覆盖来源文件');fs.writeFileSync(choice.filePath,JSON.stringify(m.config,null,2),'utf8');return {message:'方案已保存，包含字段映射和规则，不包含源数据。'};
    }
    if(m.action==='reconcileRun'){
      if(this.active)throw Error('核对正在运行');this.active=true;this.cancelled=false;this.last=null;
      try{
        const c=validateConfig(m.config);
        const tables=Object.create(null),manifest=[];
        const needed=new Set(c.rules.filter(r=>!r.skip).flatMap(r=>[r.report?.source,...(r.bindings||[]).flatMap(b=>[b.source,...(b.joins||[]).map(j=>j.source)])]).filter(Boolean));
        for(const s of c.sources.filter(s=>needed.has(s.id))){
          if(this.cancelled)throw Error('已取消');this.progress({phase:'读取 '+s.name,percent:0});
          try{const before=await digest(s.path),report=await this.read(s),after=await digest(s.path);if(before!==after)throw Error('读取期间来源文件发生变化，请保存后重试');tables[s.id]={...table(report,s),date1904:report.date1904};manifest.push({id:s.id,path:s.path,sheet:s.sheet,sha256:after,rows:tables[s.id].rows.length});}catch(e){tables[s.id]={error:s.name+'：'+e.message};manifest.push({id:s.id,error:e.message});}
        }
        if(this.cancelled)throw Error('已取消');this.progress({phase:'执行指标规则与逐笔核对',percent:65});
        const result=await new Promise((resolve,reject)=>{this.reject=reject;const w=new Worker(path.join(__dirname,'reconcile-worker.cjs'),{workerData:{config:c,tables},resourceLimits:{maxOldGenerationSizeMb:768}});this.worker=w;w.once('message',r=>r.error?reject(Error(r.error)):resolve(r));w.once('error',reject);w.once('exit',code=>{if(code)reject(Error('核对任务中断，请缩小数据范围'));});});
        result.manifest=manifest;result.config=structuredClone(c);this.last=result;return result;
      }finally{this.worker=null;this.reject=null;this.active=false;}
    }
    if(m.action==='reconcileOpenSource'){
      const ref=m.ref;if(!this.last||!this.last.manifest.some(s=>s.path===ref.file&&s.sheet===ref.sheet)||!Number.isInteger(ref.row)||ref.row<1)throw Error('来源不属于本次核对');return this.host({action:'dbOpenSource',path:ref.file,sheet:ref.sheet,row:ref.row});
    }
    if(m.action==='reconcileExport'){
      if(!this.last)throw Error('请先运行核对');const kind=['json','csv','md'].includes(m.kind)?m.kind:'md';
      const choice=await dialog.showSaveDialog(win,{title:'导出核对底稿与草稿',defaultPath:'监管变动核对.'+kind,filters:[{name:'核对结果',extensions:[kind]}],properties:['showOverwriteConfirmation']});if(choice.canceled)return {canceled:true};if(this.allowed.has(choice.filePath))throw Error('不能覆盖来源文件');
      const r=this.last;let content;
      if(kind==='json')content=JSON.stringify({...r,reviewNotes:m.notes||{}},null,2);
      else if(kind==='md')content=`# 监管变动报告（待复核草稿）\n\n期间：${r.period.previous} → ${r.period.current}\n\n生成时间：${r.createdAt}\n方案指纹：${r.configHash}\n\n`+r.results.map(x=>`## ${x.name}\n\n状态：${x.status}\n\n${x.narrative||x.reason}\n\n复核补充：${m.notes?.[x.id]||'未填写'}\n`).join('\n')+'\n## 异常清单\n\n'+r.issues.map(i=>`- ${i.rule} / ${i.ref?.sheet||''} / ${i.ref?.row||''}：${i.message}`).join('\n');
      else {const rows=[['行类型','指标','状态','维度','上期（元）','当期（元）','变动（元）','发放（元）','回收（元）','调整（元）','未解释（元）','来源文件','工作表','行号','取数角色','取数金额（元）']];for(const x of r.results){if(!x.details.length)rows.push(['指标状态',x.name,x.status,x.reason]);for(const d of x.details){rows.push(['维度汇总',x.name,x.status,d.labels.join(' / '),...['previous','current','delta','issue','recovery','adjustment','residual'].map(k=>d[k]==null?'':format(d[k],'1',6))]);for(const ref of d.refs)rows.push(['来源明细',x.name,x.status,d.labels.join(' / '),'','','','','','','',ref.file,ref.sheet,ref.row,ref.role,format(ref.amount,'1',6)]);}}content='\uFEFF'+rows.map(row=>row.map((v,col)=>{let cell=String(v??'');if(!([4,5,6,7,8,9,10,13,15].includes(col)&&/^-?\d+(\.\d+)?$/.test(cell)))cell=cell.replace(/^[=+@\-\t\r]/,"'$&");return '"'+cell.replaceAll('"','""')+'"';}).join(',')).join('\r\n');}
      fs.writeFileSync(choice.filePath,content,'utf8');return {message:'已导出 '+path.basename(choice.filePath)};
    }
    throw Error('未知核对操作');
  }
}
async function digest(file){const h=crypto.createHash('sha256');for await(const chunk of fs.createReadStream(file))h.update(chunk);return h.digest('hex');}
module.exports={ReconcileManager};
