'use strict';
const {DatabaseSync,constants:C}=require('node:sqlite');
const fs=require('node:fs'),path=require('node:path'),crypto=require('node:crypto');
const hash=v=>crypto.createHash('sha256').update(v).digest('hex');
const quote=s=>'"'+String(s).replaceAll('"','""')+'"';
const literal=s=>"'"+String(s).replaceAll("'","''")+"'";
const text=v=>v==null?'':String(v).trim();
function dateValue(value,date1904=false){
  let s=text(value);if(!s)return '';
  if(/^\d{6}$/.test(s))s=s.slice(0,4)+'-'+s.slice(4)+'-01';
  else if(/^\d{8}$/.test(s))s=s.slice(0,4)+'-'+s.slice(4,6)+'-'+s.slice(6);
  else if(/^\d{1,5}(\.\d+)?$/.test(s)){const n=Number(s);if(!date1904&&n>=60&&n<61)throw Error('Excel 虚构的 1900-02-29 不能作为业务日期。');return new Date(Date.UTC(date1904?1904:1899,date1904?0:11,date1904?1:30)+(n+(!date1904&&n<60?1:0))*86400000).toISOString().slice(0,10);}
  else s=s.replace(/[年/._]/g,'-').replace('月','-').replace('日','');
  const m=s.match(/^(\d{4})-(\d{1,2})(?:-(\d{1,2}))?$/);if(!m)throw Error('日期格式无法识别：请使用 YYYY-MM-DD、YYYYMM 或 Excel 日期。');
  const y=Number(m[1]),mo=Number(m[2]),d=Number(m[3]||1),dt=new Date(Date.UTC(y,mo-1,d));
  if(y<1900||y>9999||dt.getUTCFullYear()!==y||dt.getUTCMonth()!==mo-1||dt.getUTCDate()!==d)throw Error('报表含无效日期。');
  return dt.toISOString().slice(0,10);
}
function minor(value,scale=2,excelNumeric=false){
  let s=text(value);if(!s)return null;
  if(s.includes(',')){if(!/^[+-]?\d{1,3}(,\d{3})*(\.\d+)?$/.test(s))throw Error('金额千位分隔格式无效。');s=s.replaceAll(',','');}
  const scientific=s.match(/^([+-]?)(\d+)(?:\.(\d*))?[eE]([+-]?\d+)$/);
  if(scientific){const shift=Number(scientific[4]);if(Math.abs(shift)>30)throw Error('金额指数超出范围。');const digits=scientific[2]+(scientific[3]||''),point=scientific[2].length+shift;s=scientific[1]+(point<=0?'0.'+'0'.repeat(-point)+digits:point>=digits.length?digits+'0'.repeat(point-digits.length):digits.slice(0,point)+'.'+digits.slice(point));}
  const m=s.match(/^([+-]?)(\d+)(?:\.(\d+))?$/);if(!m)throw Error('金额不是有效数字。');
  let fraction=m[3]||'';if(fraction.length>scale&&/[^0]/.test(fraction.slice(scale))){const n=Number(s),rounded=n.toFixed(scale);if(excelNumeric&&Number.isFinite(n)&&Math.abs(n-Number(rounded))<=2*Number.EPSILON*Math.abs(n))return minor(rounded,scale,false);throw Error('金额小数位超出设定，请提高金额精度。');}
  const valueInt=BigInt(m[2]+fraction.slice(0,scale).padEnd(scale,'0'))*(m[1]==='-'?-1n:1n);
  if(valueInt>9007199254740991n||valueInt< -9007199254740991n)throw Error('金额超出精确查询范围。');return Number(valueInt);
}
function validateConfig(c){
  if(!c||!text(c.name)||!path.isAbsolute(c.folder)||!['snapshot','cumulative'].includes(c.mode))throw Error('请填写名称、文件夹和报表类型。');
  if(!Number.isInteger(c.headerRow)||c.headerRow<1||c.headerRow>100)throw Error('表头行须为 1–100。');
  if(!text(c.sheet)||!Array.isArray(c.headers)||!c.headers.length||c.headers.length>256||new Set(c.headers.map(h=>String(h).toLowerCase())).size!==c.headers.length||c.headers.some(h=>!text(h)||h.startsWith('__')))throw Error('表头不能为空、重复或以 __ 开头，请在 Excel 中修正。');
  for(const k of ['loanColumn','dateColumn','amountColumn',...(c.keys||[]).map((_,i)=>'key'+i)]){const v=k.startsWith('key')?c.keys[Number(k.slice(3))]:c[k];if(v&&!c.headers.includes(v))throw Error('所选字段不在表头中。');}
  if(!c.loanColumn)throw Error('请选择贷款号字段（允许记录中的贷款号为空）。');
  if(c.mode==='snapshot'&&!c.dateColumn&&!c.dateFromFilename)throw Error('时点表必须选择日期字段或从文件名提取日期。');
  if(!Number.isInteger(c.amountScale)||c.amountScale<0||c.amountScale>6)throw Error('金额精度为 0–6 位。');
  c.keys=[...new Set(c.keys||[])];return c;
}
function connect(filename){const db=new DatabaseSync(filename,{timeout:1500});db.exec('PRAGMA foreign_keys=ON; PRAGMA journal_mode=WAL; PRAGMA synchronous=FULL;');return db;}
function initialize(filename,c){
  validateConfig(c);const db=connect(filename);try{db.exec(`
    CREATE TABLE config(value TEXT NOT NULL);
    CREATE TABLE files(path TEXT PRIMARY KEY, hash TEXT NOT NULL, modified REAL NOT NULL, row_count INTEGER NOT NULL);
    CREATE TABLE source_rows(id INTEGER PRIMARY KEY, file TEXT NOT NULL REFERENCES files(path) ON DELETE CASCADE, sheet TEXT NOT NULL, row_number INTEGER NOT NULL, record_key TEXT NOT NULL, payload_hash TEXT NOT NULL, payload TEXT NOT NULL, loan TEXT, report_date TEXT, amount_minor INTEGER, quality TEXT NOT NULL);
    CREATE INDEX source_key ON source_rows(record_key); CREATE INDEX source_file ON source_rows(file);
    CREATE TABLE records(record_key TEXT PRIMARY KEY, source_id INTEGER NOT NULL REFERENCES source_rows(id), source_count INTEGER NOT NULL, conflict INTEGER NOT NULL);
    CREATE TABLE scans(id INTEGER PRIMARY KEY, finished_at TEXT NOT NULL, files INTEGER NOT NULL, changed INTEGER NOT NULL, removed INTEGER NOT NULL, records INTEGER NOT NULL, conflicts INTEGER NOT NULL);
  `);db.prepare('INSERT INTO config VALUES (?)').run(JSON.stringify(c));
  const columns=c.headers.map((h,i)=>`json_extract(s.payload,'$[${i}]') AS ${quote(h)}`).join(',');
  const meta=`s.record_key AS __key,s.file AS __file,s.sheet AS __sheet,s.row_number AS __row,s.loan AS __loan,s.report_date AS __date,s.amount_minor AS __amount_minor,s.quality AS __quality`;
  db.exec(`CREATE VIEW data AS SELECT ${columns},${meta},r.source_count AS __sources,r.conflict AS __conflict FROM records r JOIN source_rows s ON s.id=r.source_id;
    CREATE VIEW provenance AS SELECT ${columns},${meta},s.id AS __source_id FROM source_rows s;
    CREATE VIEW issues AS SELECT * FROM data WHERE __quality<>'ok' OR __conflict<>0;
    CREATE VIEW latest AS SELECT * FROM data WHERE __date=(SELECT MAX(__date) FROM data);`);
  }finally{db.close();}
}
function filesIn(folder,recursive){
  const result=[];
  function visit(dir){for(const e of fs.readdirSync(dir,{withFileTypes:true})){if(e.isSymbolicLink()||e.name.startsWith('~$'))continue;const p=path.join(dir,e.name);if(e.isDirectory()&&recursive)visit(p);else if(e.isFile()&&/\.(xlsx|xlsm)$/i.test(e.name))result.push(p);if(result.length>10000)throw Error('文件数量超过 10000，请缩小目录范围。');}}
  visit(folder);return result.sort();
}
async function fileHash(file){const h=crypto.createHash('sha256');for await(const chunk of fs.createReadStream(file))h.update(chunk);return h.digest('hex');}
async function scan(filename,read,progress=()=>{}){
  const db=connect(filename);let transaction=false;
  try{
    const c=JSON.parse(db.prepare('SELECT value FROM config').get().value),files=filesIn(c.folder,c.recursive);
    // An empty folder can be a disconnected share or an accidental selection. Never silently clear it.
    if(!files.length)throw Error('文件夹中没有报表；为防止误清空，保留上次数据。');
    db.exec('BEGIN IMMEDIATE');transaction=true;
    const old=new Map(db.prepare('SELECT * FROM files').all().map(f=>[f.path,f]));
    const observed=new Map();let changed=0;
    db.exec('DELETE FROM records');
    const insert=db.prepare('INSERT INTO source_rows(file,sheet,row_number,record_key,payload_hash,payload,loan,report_date,amount_minor,quality) VALUES(?,?,?,?,?,?,?,?,?,?)');
    for(let i=0;i<files.length;i++){
      const file=files[i];progress({phase:`扫描 ${i+1}/${files.length} · ${path.basename(file)}`,percent:Math.floor(i/files.length*90),current:i,total:files.length});
      const digest=await fileHash(file);observed.set(file,digest);
      if(old.get(file)?.hash===digest)continue;
      const report=await read(file,c.sheet);if(await fileHash(file)!==digest)throw Error('扫描期间文件发生变化，请保存后重新扫描。');
      const header=report.rows.find(r=>r.row===c.headerRow)?.values.map(text)||[];while(header.at(-1)==='')header.pop();
      if(JSON.stringify(header)!==JSON.stringify(c.headers))throw Error(path.basename(file)+'：表头与建库时不一致，请在源表统一列名和顺序。');
      const rows=report.rows.filter(r=>r.row>c.headerRow);
      db.prepare('DELETE FROM files WHERE path=?').run(file);
      db.prepare('INSERT INTO files VALUES(?,?,?,?)').run(file,digest,fs.statSync(file).mtimeMs,rows.length);
      const occurrences=new Map();
      for(const row of rows){
        const values=c.headers.map((_,j)=>row.values[j]??null),get=col=>col?values[c.headers.indexOf(col)]:null;
        if(row.values.slice(c.headers.length).some(v=>text(v)))throw Error(path.basename(file)+'：数据超出已配置表头。');
        const loan=text(get(c.loanColumn));
        let reportDate='';try{reportDate=c.dateColumn?dateValue(get(c.dateColumn),report.date1904):c.dateFromFilename?dateValue(path.basename(file).match(/(?:19|20)\d{2}[-_.]?\d{2}(?:[-_.]?\d{2})?/)?.[0]):'';}catch(e){throw Error(`${path.basename(file)} 第 ${row.row} 行：${e.message}`);}
        if(c.mode==='snapshot'&&!reportDate)throw Error(path.basename(file)+'：时点日期为空，无法归档快照。');
        let amount;try{amount=minor(get(c.amountColumn),c.amountScale,Array.isArray(row.numeric)&&row.numeric.includes(c.headers.indexOf(c.amountColumn)));}catch(e){throw Error(`${path.basename(file)} 第 ${row.row} 行：${e.message}`);}
        const payload=JSON.stringify(values),payloadHash=hash(payload);
        let quality='ok',key;
        const extra=c.keys.map(get).map(text);
        if(!loan||extra.some(v=>!v)){quality=!loan?'missing_loan':'missing_key';key=hash(JSON.stringify(['unresolved',digest,c.sheet,row.row]));}
        else if(c.mode==='snapshot')key=hash(JSON.stringify(['snapshot',reportDate,loan,...extra]));
        else if(c.keys.length)key=hash(JSON.stringify(['event',loan,...extra]));
        else {quality='weak_key';const n=(occurrences.get(payloadHash)||0)+1;occurrences.set(payloadHash,n);key=hash(JSON.stringify(['content',payloadHash,n]));}
        insert.run(file,report.sheet,row.row,key,payloadHash,payload,loan||null,reportDate||null,amount,quality);
      }
      changed++;
    }
    let removed=0;for(const file of old.keys())if(!observed.has(file)){db.prepare('DELETE FROM files WHERE path=?').run(file);removed++;}
    // All contributing rows remain available. Conflicting payloads are never silently discarded.
    db.exec(`INSERT INTO records SELECT record_key,id,n,conflict FROM (
      SELECT s.record_key,s.id,COUNT(*) OVER(PARTITION BY s.record_key) AS n,
      ROW_NUMBER() OVER(PARTITION BY s.record_key ORDER BY f.modified DESC,s.file,s.row_number) AS rank,
      CASE WHEN (SELECT COUNT(DISTINCT payload_hash) FROM source_rows p WHERE p.record_key=s.record_key)>1
       OR EXISTS(SELECT 1 FROM source_rows p WHERE p.record_key=s.record_key GROUP BY p.file HAVING COUNT(*)>1) THEN 1 ELSE 0 END AS conflict
      FROM source_rows s JOIN files f ON f.path=s.file) WHERE rank=1;`);
    // Verify all files, including unchanged ones, and membership again before committing deletions.
    if(JSON.stringify(filesIn(c.folder,c.recursive))!==JSON.stringify(files))throw Error('扫描期间目录内容变化，请重试。');
    for(const [file,digest] of observed)if(await fileHash(file)!==digest)throw Error('扫描期间报表发生变化，请保存后重试。');
    const records=db.prepare('SELECT COUNT(*) AS n FROM records').get().n,conflicts=db.prepare('SELECT COUNT(*) AS n FROM records WHERE conflict=1').get().n;
    db.prepare('INSERT INTO scans(finished_at,files,changed,removed,records,conflicts) VALUES(?,?,?,?,?,?)').run(new Date().toISOString(),files.length,changed,removed,records,conflicts);
    db.exec('COMMIT');transaction=false;return {files:files.length,changed,removed,records,conflicts,message:`扫描完成：${files.length} 个文件，${changed} 个更新，${removed} 个移除，${records} 条记录，${conflicts} 个冲突。`};
  }finally{if(transaction)db.exec('ROLLBACK');db.close();}
}
function info(filename){const db=new DatabaseSync(filename,{readOnly:true});try{return {config:JSON.parse(db.prepare('SELECT value FROM config').get().value),lastScan:db.prepare('SELECT * FROM scans ORDER BY id DESC LIMIT 1').get(),files:db.prepare('SELECT * FROM files ORDER BY path').all(),issues:db.prepare('SELECT COUNT(*) AS n FROM issues').get().n};}finally{db.close();}}
function query(filename,sql,offset=0,pageSize=100){
  if(typeof sql!=='string'||sql.length>50000)throw Error('SQL 最长 50000 字符。');
  sql=sql.trim().replace(/;\s*$/,'');if(!sql)throw Error('请输入 SELECT 查询。');
  offset=Math.max(0,Math.min(19900,Math.floor(Number(offset)||0)));pageSize=Math.max(1,Math.min(200,Math.floor(Number(pageSize)||100),20000-offset));
  const db=new DatabaseSync(filename,{readOnly:true,allowExtension:false,timeout:1000});
  try{
    db.exec('PRAGMA query_only=ON; PRAGMA trusted_schema=OFF;');
    const allowed=new Set([C.SQLITE_SELECT,C.SQLITE_READ,C.SQLITE_RECURSIVE]);
    const functions=new Set('abs avg cast coalesce count date datetime dense_rank first_value group_concat hex ifnull iif instr json_extract json_array_length julianday lag last_value lead length like lower ltrim max min nullif rank replace round row_number rtrim strftime substr substring sum time total trim typeof unicode upper'.split(' '));
    db.setAuthorizer((code,a,b)=>allowed.has(code)||(code===C.SQLITE_FUNCTION&&functions.has(String(b).toLowerCase()))?C.SQLITE_OK:C.SQLITE_DENY);
    const stmt=db.prepare(`SELECT * FROM (${sql}) AS result LIMIT ${pageSize+1} OFFSET ${offset}`);stmt.setReadBigInts(true);stmt.setReturnArrays(true);
    const columns=stmt.columns().map(c=>c.name),raw=stmt.all();const more=raw.length>pageSize;
    const rows=raw.slice(0,pageSize).map(r=>r.map(v=>typeof v==='bigint'?(v<=9007199254740991n&&v>= -9007199254740991n?Number(v):v.toString()):v));
    if(JSON.stringify(rows).length>16*1024*1024)throw Error('查询页内容过大，请减少列或每页条数。');
    return {columns,rows,offset,pageSize,more:more&&offset+pageSize<20000,capped:more&&offset+pageSize>=20000};
  }finally{db.close();}
}
module.exports={initialize,scan,info,query,validateConfig,dateValue,minor,filesIn,quote,literal,hash};
