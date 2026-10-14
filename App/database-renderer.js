'use strict';
window.showDatabase=async function(){
  document.body.classList.add('database');
  let library=[],active=null,sample=null,offset=0,pageSize=100,result=null,currentSql='',filters=[],selectedView='data';
  $('eyebrow').textContent='数据中心 / 本地报表库';$('title').textContent='把每次报表，变成可查询的记录';
  $('subtitle').textContent='Excel 是数据源：在源表增删改并保存，再点击扫描更新。';
  document.querySelector('.local p').textContent='报表只读扫描 · 数据本地保存';
  async function call(m){try{return await request(m);}catch(e){status(e.message,true);return null;}}
  function label(title,control,hint){const wrap=el('label',undefined,'db-field');wrap.append(el('span',title),control);if(hint)wrap.append(el('small',hint));return wrap;}
  function table(columns,rows,open){const wrap=el('div',undefined,'db-table');const t=el('table'),head=el('thead'),hr=el('tr'),body=el('tbody');if(open)hr.append(el('th','来源'));columns.forEach(c=>hr.append(el('th',c)));head.append(hr);for(const r of rows){const tr=el('tr');if(open){const cell=el('td');const file=r[columns.indexOf('__file')];if(file){cell.append(button('打开 Excel',()=>call({action:'dbOpenSource',id:active.id,file,sheet:r[columns.indexOf('__sheet')],row:r[columns.indexOf('__row')]}),'quiet'));const key=r[columns.indexOf('__key')];if(key)cell.append(button('所有来源',()=>{selectedView='provenance';$('dbSql').value=`SELECT * FROM provenance WHERE __key = ${lit(key)} ORDER BY __file, __row`;execute(0);},'quiet'));}tr.append(cell);}for(const v of r){const td=el('td',v==null?'':String(v));td.title=v==null?'':String(v);tr.append(td);}body.append(tr);}t.append(head,body);wrap.append(t);return wrap;}
  const lit=v=>"'"+String(v).replaceAll("'","''")+"'",qid=v=>'"'+String(v).replaceAll('"','""')+'"';
  async function load(id){const data=await call({action:'dbList'});if(!data)return;library=data.databases;active=library.find(d=>d.id===(id||active?.id))||library[0]||null;renderNav();active?dashboard():empty();}
  function renderNav(){nav([['new','＋ 创建数据库',wizard],...library.map(d=>[d.id,d.config.name,()=>{active=d;selectedView='data';filters=[];renderNav();dashboard();}])],active?.id);}
  function empty(){$('content').replaceChildren();const c=el('section',undefined,'card db-empty');c.append(el('h2','从一个报表文件夹开始'),el('p','余额表保存不同日期的快照；发放表、回收表归并重复导出，并保留每条记录的来源。'),button('选择文件夹并创建',wizard,'primary'));$('content').append(c);$('actions').replaceChildren();status('支持 .xlsx / .xlsm；不运行宏、不刷新外部连接，读取最近保存的值。');}
  function wizard(){
    active=null;renderNav();sample=null;$('title').textContent='创建新的报表数据库';$('subtitle').textContent='选择来源并确认报表结构，建立独立的数据空间。';const content=$('content');content.replaceChildren();$('actions').replaceChildren();
    const c=el('section',undefined,'card');c.append(el('h2','创建报表数据库'),el('p','每个库对应一种报表结构。先选择文件夹，再确认字段和记录标识。','help'));content.append(c);
    const name=input('','text',()=>{});name.id='dbName';name.placeholder='例如：贷款余额月报';
    const recursive=input(false,'checkbox',()=>{});recursive.id='dbRecursive';const folderText=el('p','尚未选择文件夹','db-path'),mapping=el('div');
    let config={name:'',folder:'',mode:'snapshot',headerRow:2,keys:[],amountScale:2};
    const choose=button('选择报表文件夹',async()=>{choose.disabled=true;try{const d=await call({action:'dbChoose',recursive:recursive.checked});if(!d||d.canceled)return;sample=d;config.folder=d.folder;config.sheet=d.sheet;folderText.textContent=d.folder+' · 样本：'+d.sample;const detected=d.rows.find(r=>r.values.some(v=>/^(贷款号|借据号|贷款编号)$/.test(String(v))));config.headerRow=detected?.row||1;if(!name.value)name.value=d.sample.replace(/\.[^.]+$/,'');renderMapping();}finally{choose.disabled=false;}});
    c.append(label('数据库名称',name),label('同时扫描子文件夹',recursive),choose,folderText,mapping);
    function renderMapping(){
      mapping.replaceChildren();config.headers=(sample.rows.find(r=>r.row===config.headerRow)?.values||[]).map(v=>v==null?'':String(v).trim());while(config.headers.at(-1)==='')config.headers.pop();
      const fields=[['','不使用'],...config.headers.filter(Boolean).map(h=>[h,h])];
      const sheet=select(sample.sheets.map(s=>[s,s]),config.sheet,async v=>{const d=await call({action:'dbSample',sheet:v});if(d){sample={...sample,...d};config.sheet=v;renderMapping();}});sheet.id='dbSheet';
      const header=input(config.headerRow,'number',v=>{config.headerRow=v;renderMapping();});header.min=1;header.max=100;header.id='dbHeaderRow';
      const mode=select([['snapshot','时点型 · 余额 / 存量'],['cumulative','累计型 · 发放 / 回收记录']],config.mode,v=>{config.mode=v;renderMapping();});mode.id='dbMode';
      const grid=el('div',undefined,'db-fields');grid.append(label('工作表',sheet),label('表头行',header),label('报表类型',mode));mapping.append(grid);
      config.loanColumn=config.headers.find(h=>/^(贷款号|借据号|贷款编号)$/.test(h))||config.headers[0];config.dateColumn=config.headers.find(h=>/^(月份|报表日期|统计日期|业务日期|发放日期|回收日期|日期)$/.test(h))||'';config.amountColumn=config.headers.find(h=>/^(余额|金额|发放金额|回收金额|贷款余额)$/.test(h))||'';
      for(const [key,title,hint] of [['loanColumn','贷款号字段','作为文本保存，保留前导零。'],['dateColumn',config.mode==='snapshot'?'快照日期字段':'业务日期字段','月份自动归到月初；支持 Excel 日期。'],['amountColumn','金额字段','同时保存精确整数金额 __amount_minor。']]){const control=select(fields,config[key],v=>config[key]=v);control.id='db-'+key;grid.append(label(title,control,hint));}
      const fromFile=input(!!config.dateFromFilename,'checkbox',v=>config.dateFromFilename=v);grid.append(label('日期为空配置时，从文件名提取日期',fromFile,'文件名须含 YYYYMM、YYYYMMDD 或 YYYY-MM-DD；以日期字段优先。'));
      const scale=input(config.amountScale,'number',v=>config.amountScale=v);scale.min=0;scale.max=6;grid.append(label('金额小数位',scale,'默认 2：__amount_minor 以分为单位。'));
      const keys=el('div',undefined,'db-keys');keys.append(el('h3','补充唯一键（可多选）'),el('p',config.mode==='snapshot'?'同一时点每个贷款有多行时，加选币种、机构等区分字段。':'优先选择业务流水号；没有流水号时，可选择日期、金额和业务类型。同一贷款不能直接视为同一笔业务。','help'));
      config.keys=config.keys.filter(k=>config.headers.includes(k));for(const h of config.headers.filter(Boolean)){const check=input(config.keys.includes(h),'checkbox',v=>{config.keys=v?[...config.keys,h]:config.keys.filter(k=>k!==h);});keys.append(label(h,check));}mapping.append(keys);
      mapping.append(el('p','空贷款号或不完整唯一键：保留并标记。累计型未选补充键时按完整行内容去重，同一文件内相同多行保留次数，并标记为弱标识。冲突行保留所有来源，需在 Excel 核对。','large-alert'));
      mapping.append(el('h3','样本前 6 行（按所选表头）'),table(config.headers,sample.rows.filter(r=>r.row>config.headerRow).slice(0,6).map(r=>config.headers.map((_,i)=>r.values[i]))));
      $('actions').replaceChildren(button('创建数据库',async()=>{config.name=name.value.trim();config.recursive=recursive.checked;const created=await run({action:'dbCreate',config});if(created){await load(created.id);status('数据库已创建。点击“扫描更新”开始导入。');}},'primary'));
    }
    status('报表结构变化会阻止整次更新；库文件保存在当前 Windows 用户的 GridPilot 数据目录。');
  }
  function dashboard(){
    const c=active.config;offset=0;result=null;currentSql='';$('title').textContent=c.name;$('subtitle').textContent=`${c.mode==='snapshot'?'时点快照':'累计记录'} · ${c.folder}`;
    const content=$('content');content.replaceChildren();
    const s=active.lastScan,stats=el('div',undefined,'stats');for(const [v,k] of [[s?.records||0,'归并后记录'],[active.files.length,'来源文件'],[active.issues||0,'需核对记录'],[s?new Date(s.finished_at).toLocaleString():'尚未扫描','最近成功更新']]){const tile=el('div',undefined,'stat');tile.append(el('b',v),el('span',k));stats.append(tile);}content.append(stats);
    const card=el('section',undefined,'card');const tabs=el('div',undefined,'toolbar');
    for(const [v,t] of [['data','全部记录'],...(c.mode==='snapshot'?[['latest','最新快照']]:[]),['issues','待核对'],['provenance','所有来源'],['files','文件清单'],['scans','扫描历史']])tabs.append(button(t,()=>{selectedView=v;filters=[];dashboard();},selectedView===v?'primary':'quiet'));card.append(tabs);
    const help=el('p','筛选生成 SQL 后可自由修改；仅允许读取。结果最多浏览 20,000 条，每页 100 / 200 条。','help');card.append(help);
    const filtersBox=el('div',undefined,'db-filters');card.append(filtersBox);
    const fields=['files','scans'].includes(selectedView)?(selectedView==='files'?['path','hash','modified','row_count']:['id','finished_at','files','changed','removed','records','conflicts']):[...c.headers,'__loan','__date','__amount_minor','__quality',...(selectedView==='provenance'?[]:['__conflict']),'__file'];
    function renderFilters(){filtersBox.replaceChildren();for(const f of filters){const line=el('div',undefined,'toolbar');line.append(select(fields.map(h=>[h,h]),f.field,v=>f.field=v),select([['contains','包含'],['=','等于'],['<>','不等于'],['>','大于（数字）'],['<','小于（数字）'],['>=','大于等于（数字）'],['<=','小于等于（数字）'],['empty','为空']],f.op,v=>f.op=v),input(f.value,'text',v=>f.value=v),button('移除',()=>{filters=filters.filter(x=>x!==f);renderFilters();},'quiet'));filtersBox.append(line);}}
    const toolbar=el('div',undefined,'toolbar');toolbar.append(button('＋ 添加条件',()=>{filters.push({field:fields[0],op:'contains',value:''});renderFilters();}),button('生成 SQL',()=>generate()));card.append(toolbar);renderFilters();
    const sql=el('textarea');sql.id='dbSql';sql.spellcheck=false;sql.rows=4;sql.setAttribute('aria-label','SQL 查询');card.append(sql);const queryActions=el('div',undefined,'toolbar');queryActions.append(button('执行查询',()=>execute(0),'primary'));const sizes=select([[100,'每页 100 条'],[200,'每页 200 条']],pageSize,v=>{pageSize=Number(v);execute(0);});sizes.setAttribute('aria-label','查询每页条数');queryActions.append(sizes);card.append(queryActions);
    const results=el('div');results.id='dbResults';card.append(results);const paging=el('div',undefined,'toolbar');paging.id='dbPaging';card.append(paging);content.append(card);
    const guide=el('details',undefined,'card');guide.append(el('summary','数据规则与 SQL 字段说明'),el('p',`金额精度为 ${c.amountScale} 位，精确汇总请用 SUM(__amount_minor)。__quality 标记 missing_loan / missing_key / weak_key；__conflict=1 表示同键数据不一致或同一文件中键重复。冲突展示文件修改时间较新的代表行，所有来源仍可查询。`),el('p','data：归并记录；latest：日期最大的快照；issues：待核对；provenance：全部源行；files：文件与哈希；scans：成功扫描记录。月份快照不可跨月直接求和。'),el('p','删除源文件会在下一次成功扫描时移除其贡献；其他文件仍有相同记录时会保留。目录为空、表头变化、金额或日期无效、文件变动等错误会使整次更新回滚。'),el('p','读取已保存的缓存值；请先在 Excel 完成计算并保存。累计数据未配置业务键时，修订行可能无法识别为同一业务，应回到源表核对。'));content.append(guide);
    const scanButton=button('扫描更新',async()=>{const id=active.id;const r=await run({action:'dbScan',id});if(r){await load(id);status(r.message);}},'primary');const cancel=button('取消任务',()=>request({action:'cancel'}));cancel.id='cancel';$('actions').replaceChildren(cancel,scanButton);
    function generate(){const conditions=filters.map(f=>{const field=qid(f.field);if(f.op==='empty')return `(${field} IS NULL OR ${field} = '')`;if(f.op==='contains')return `instr(lower(COALESCE(${field}, '')), lower(${lit(f.value)})) > 0`;if(['>','<','>=','<='].includes(f.op)){if(!/^[+-]?(?:\d+\.?\d*|\.\d+)$/.test(f.value))throw Error('数字比较需要有效数字。');return `CAST(${field} AS REAL) ${f.op} ${Number(f.value)}`;}return `${field} ${f.op} ${lit(f.value)}`;});sql.value=`SELECT * FROM ${selectedView}${conditions.length?'\nWHERE '+conditions.join('\n  AND '):''}\nORDER BY ${selectedView==='files'?'path':selectedView==='scans'?'id DESC':selectedView==='provenance'?'__file, __row':'__key'}`;}
    toolbar.lastChild.onclick=()=>{try{generate();status('SQL 已生成，可修改后执行。');}catch(e){status(e.message,true);}};
    generate();status(active.lastScan?'已载入上次成功数据。修改源表后请扫描更新。':'已创建空库，点击扫描更新。');if(active.lastScan)execute(0);
  }
  async function execute(next){if(busy)return;const sql=$('dbSql').value,id=active.id;const d=await run({action:'dbQuery',id,sql,offset:next,pageSize});if(!d||active?.id!==id)return;result=d;offset=d.offset;currentSql=sql;$('dbResults').replaceChildren(table(d.columns,d.rows,true));const prev=button('上一页',()=>{$('dbSql').value=currentSql;execute(Math.max(0,offset-pageSize));});prev.disabled=offset===0;const nextButton=button('下一页',()=>{$('dbSql').value=currentSql;execute(offset+pageSize);});nextButton.disabled=!d.more;$('dbPaging').replaceChildren(prev,el('span',d.rows.length?`第 ${offset+1}–${offset+d.rows.length} 条${d.capped?' · 已到 20,000 条上限，请缩小筛选范围':''}`:'没有匹配记录'),nextButton);status('查询完成。更改数据请打开源 Excel，保存后重新扫描。');}
  await load();
};
