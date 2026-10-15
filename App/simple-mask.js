/* Selection-only masking. No source data is persisted by the renderer. */
window.showSimpleMask=function(data){
  const get=id=>document.getElementById(id),make=(tag,text,cls)=>{const e=document.createElement(tag);if(text!==undefined)e.textContent=text;if(cls)e.className=cls;return e;};
  let snapshot=data,working=false,ready=false,applied=false,offset=0,total=0;
  const content=get('content');content.replaceChildren();get('eyebrow').textContent='数据工具 / 选区处理';get('title').textContent='简易脱敏';get('subtitle').textContent='固定规则 · 逐格容错 · 预览后选择输出方式';
  get('nav').replaceChildren(make('button','简易脱敏','active'));document.querySelector('.local p').textContent='默认只生成预览；写回原表前请备份。';
  const card=make('section',undefined,'card');content.append(card);
  const source=make('p',undefined,'help');card.append(source);
  function sourceInfo(){source.textContent=`${snapshot.book} / ${snapshot.sheet} · ${snapshot.count.toLocaleString()} 个有内容单元格 · ${snapshot.blocks.length} 个区域`;}
  const field=(label,id,type,value)=>{const wrap=make('label',label,'simple-field'),input=make('input');input.id=id;input.type=type;input.value=value;wrap.append(input);return wrap;};
  const modeLabel=make('label','处理方式','simple-field'),mode=make('select');mode.id='simpleMode';[['initials','拼音首字母 / 固定字符替换'],['swap','交换首尾字符'],['reverse','字符倒序'],['regex','自定义正则替换'],['number','纯数字加减乘除']].forEach(([value,label])=>{const o=make('option',label);o.value=value;mode.append(o);});modeLabel.append(mode);card.append(modeLabel);
  const hints=make('p',undefined,'help');card.append(hints);
  const specific=make('div');card.append(specific);
  const affixes=make('div',undefined,'simple-fields');affixes.append(field('固定前缀（可选）','simplePrefix','text',''),field('固定后缀（可选）','simpleSuffix','text',''));card.append(affixes);get('simplePrefix').maxLength=get('simpleSuffix').maxLength=256;
  const details=make('details');details.append(make('summary','查看固定映射和还原说明'));const map=make('p',undefined,'help');map.textContent=`规则版本：${snapshot.mappingVersion}。中文取固定普通话读音的首字母（多音字不按上下文判断）；英文大写 A–Z：${[...snapshot.letters].map((c,i)=>String.fromCharCode(65+i)+'→'+c).join('，')}。小写加「〔小〕」标记。数字 0–9：${snapshot.digits}。空格、标点、符号及未收录字符统一转「〔符XXXXXX〕」，编码为 Unicode 码点加 0x5A17 后对 0x110000 取余。纯空白单元格跳过。`;details.append(map,make('p','首字母、任意正则与数字舍入可能丢失信息。这是便于分享的确定性替换，不是加密。需要将来还原时，单独导出含原文、位置、类型、规则和处理状态的还原记录；不要与脱敏报表一起分享。','help'));card.append(details);
  const previewCard=make('section',undefined,'card');previewCard.append(make('h2','结果预览'));const summary=make('p','先选择规则，再生成预览。','help');summary.id='simpleSummary';previewCard.append(summary);
  const search=make('input');search.id='simpleSearch';search.type='search';search.placeholder='搜索位置、原值、结果或跳过原因';search.setAttribute('aria-label','搜索简易脱敏预览');
  const toolbar=make('div',undefined,'toolbar');toolbar.append(search);previewCard.append(toolbar);
  function btn(text,action,id){const b=make('button',text);if(id)b.id=id;b.onclick=action;return b;}
  const find=btn('搜索',()=>page(0),'simpleFind');toolbar.append(find);
  const tableWrap=make('div',undefined,'table-wrap simple-table'),table=make('table'),head=make('thead'),tr=make('tr');['位置','原值','脱敏结果','状态 / 跳过原因'].forEach(t=>tr.append(make('th',t)));head.append(tr);const body=make('tbody');body.id='simpleRows';table.append(head,body);tableWrap.append(table);previewCard.append(tableWrap);
  const paging=make('div',undefined,'toolbar'),pageInfo=make('span');pageInfo.id='simplePageInfo';const prev=btn('上一页',()=>page(offset-100),'simplePrev'),next=btn('下一页',()=>page(offset+100),'simpleNext');paging.append(prev,pageInfo,next);previewCard.append(paging);
  const check=(label,id)=>{const l=make('label',undefined,'confirm'),c=make('input');c.type='checkbox';c.id=id;l.append(c,document.createTextNode(label));c.onchange=sync;return l;};
  previewCard.append(check('我已检查预览、跳过项和输出范围','simpleConfirmed'));
  const output=make('div',undefined,'toolbar'),block=make('select');block.id='simpleBlock';block.setAttribute('aria-label','复制区域');
  const copy=btn('复制结果',()=>act('simpleCopy'),'simpleCopy'),exportButton=btn('导出 XLSX',()=>act('simpleExport'),'simpleExport');output.append(block,copy,exportButton);previewCard.append(output,make('p','复制按所选区域排列，导出每个区域为一张结果表；区域边缘空白裁去，内部空格保留。跳过项在输出中留空，绝不混入未脱敏原文。公式、合并格、错误值及受保护格跳过。输出为文本值，保留长编号及精度；不复制格式。','help'));
  previewCard.append(check('我已备份原数据，了解写回会覆盖选区中的可处理内容','simpleBackup'));
  const apply=btn('在原表上脱敏',()=>act('simpleApply'),'simpleApply');previewCard.append(apply,make('p','只修改预览中的源工作表；预览后被编辑的单元格会跳过。写回不自动保存，可用 GridPilot 撤销；取消时已完成的写入保留。','help'));
  previewCard.append(check('还原记录含原文，我会单独安全保管','simpleSensitive'));
  const recovery=btn('导出还原记录 JSON',()=>act('simpleRecovery'),'simpleRecovery');previewCard.append(recovery);content.append(previewCard);
  const generate=btn('生成预览',async()=>{invalidate();const r=await call({action:'simplePreview',options:options()});if(r){ready=true;applied=false;search.value='';get('simpleConfirmed').checked=false;display(r);sync();}},'simplePreview');generate.className='primary';
  const capture=btn('重新读取当前选区',async()=>{const r=await call({action:'simpleCapture'});invalidate();if(r){snapshot=r;sourceInfo();fillBlocks();summary.textContent='选区已更新，请重新生成预览。';}},'simpleCapture');
  const cancel=btn('取消处理',()=>window.gridpilot.request({action:'cancel'}),'simpleCancel');get('actions').replaceChildren(capture,cancel,generate);
  function fillBlocks(){block.replaceChildren(...snapshot.blocks.map((b,i)=>{const o=make('option',`${i+1}. ${b.Address}（${b.Rows} 行 × ${b.Columns} 列）`);o.value=i;return o;}));}
  function options(){return {Mode:mode.value,Prefix:get('simplePrefix').value,Suffix:get('simpleSuffix').value,EdgeCount:Number(get('simpleEdge')?.value||2),Pattern:get('simplePattern')?.value||'^(.{2})([\\s\\S]*)(.{2})$',Replacement:get('simpleReplacement')?.value??'$3$2$1',Operator:get('simpleOperator')?.value||'+',Operand:get('simpleOperand')?.value||'100'};}
  function invalidate(){ready=false;applied=false;get('simpleConfirmed').checked=false;get('simpleBackup').checked=false;get('simpleSensitive').checked=false;body.replaceChildren();summary.textContent='规则或选区已改变，请生成预览。';offset=0;total=0;sync();}
  function renderMode(){specific.replaceChildren();
    if(mode.value==='initials')hints.textContent='例如：新加坡油气有限公司 → XJPYQYXGS。英文、数字、标点和特殊字符均使用固定替换。';
    if(mode.value==='swap'){hints.textContent='交换首尾各 N 个字符，中间不变。例如 N=2：P123456 → 56234P1。长度不足的单元格跳过。';specific.append(field('首尾各几个字符','simpleEdge','number','2'));get('simpleEdge').min=1;get('simpleEdge').max=1000;const quick=make('div',undefined,'toolbar');[1,2,3,4].forEach(n=>quick.append(btn('首尾各 '+n+' 位',()=>{get('simpleEdge').value=n;invalidate();})));specific.append(quick);}
    if(mode.value==='reverse')hints.textContent='从后向前排列字符，例如 P123456 → 654321P；中英文、数字和符号均参与。';
    if(mode.value==='regex'){hints.textContent='使用 .NET 正则：$1、$2、${name} 引用捕获组，$$ 表示美元符号。未匹配的单元格跳过。每格超时 50 ms；累计 10 次超时后跳过剩余项。';specific.append(field('正则查找','simplePattern','text','^(.{2})([\\s\\S]*)(.{2})$'),field('替换内容（允许留空）','simpleReplacement','text','$3$2$1'));specific.append(btn('首尾各两位模板',()=>{get('simplePattern').value='^(.{2})([\\s\\S]*)(.{2})$';get('simpleReplacement').value='$3$2$1';invalidate();}));}
    if(mode.value==='number'){hints.textContent='支持数值及数字文本、正负号、小数和科学计数法；小数点使用 .，不接受千分位、百分号或日期文字。运算采用十进制，超出小数精度时可能舍入；结果以文本保存，前导零会消失。超过 28 位有效数字跳过。';const l=make('label','运算','simple-field'),s=make('select');s.id='simpleOperator';[['+','加'],['-','减'],['*','乘'],['/','除']].forEach(([v,t])=>{const o=make('option',t);o.value=v;s.append(o);});l.append(s);specific.append(l,field('运算数','simpleOperand','text','100'));}
    invalidate();
  }
  function sync(){generate.disabled=working||snapshot.count===0;capture.disabled=working;cancel.disabled=!working;prev.disabled=working||!ready||offset===0;next.disabled=working||!ready||offset+100>=total;find.disabled=working||!ready;const allowed=ready&&!working&&get('simpleConfirmed').checked;copy.disabled=!allowed||!snapshot.blocks.length;exportButton.disabled=!allowed;apply.disabled=!allowed||applied||!get('simpleBackup').checked;recovery.disabled=!allowed||!get('simpleSensitive').checked;}
  async function call(message){if(working)return;working=true;const controls=[...card.querySelectorAll('input,select,button'),search,block,...previewCard.querySelectorAll('input[type=checkbox]')];controls.forEach(c=>c.disabled=true);sync();get('status').classList.remove('error');get('status').textContent='正在处理，可取消…';try{return await window.gridpilot.request(message);}catch(e){get('status').textContent=e.message.replace(/^Error invoking remote method 'request': Error: /,'');get('status').classList.add('error');}finally{working=false;controls.forEach(c=>c.disabled=false);get('progressWrap').hidden=true;sync();}}
  function display(r){offset=r.offset;total=r.total;summary.textContent=r.summary;body.replaceChildren(...r.rows.map(row=>{const tr=make('tr');[row.address,row.before,row.after??'（留空）',row.status].forEach(v=>tr.append(make('td',v)));return tr;}));pageInfo.textContent=total?`${offset+1}–${Math.min(offset+100,total)} / ${total}`:'0 条';get('status').textContent=r.summary;}
  async function page(n){if(!ready)return;const r=await call({action:'simplePage',options:options(),search:search.value,offset:n});if(r)display(r);sync();}
  async function act(action){const r=await call({action,options:options(),confirmed:get('simpleConfirmed').checked,backupConfirmed:get('simpleBackup').checked,sensitiveConfirmed:get('simpleSensitive').checked,block:Number(block.value)});if(r){if(r.applied){applied=true;await page(offset);}get('status').textContent=r.message;sync();}}
  card.addEventListener('input',invalidate);card.addEventListener('change',event=>{if(event.target===mode)renderMode();else invalidate();});search.onkeydown=e=>{if(e.key==='Enter')page(0);};sourceInfo();fillBlocks();renderMode();get('status').textContent='选区已读取，设置规则后生成预览。';
};
