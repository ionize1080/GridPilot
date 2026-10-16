/* Sorting state is isolated from the other tool workspaces. Excel owns comparisons. */
window.showSort = function(initial) {
  const $=id=>document.getElementById(id),make=(tag,text,cls)=>{const n=document.createElement(tag);if(text!==undefined)n.textContent=text;if(cls)n.className=cls;return n;};
  let data=initial,options=structuredClone(initial.options),page=null,valid=false,busy=false,applied=false,canUndo=false,offset=0,size=100,search='';
  const say=(text,error=false)=>{$('status').textContent=text;$('status').classList.toggle('error',error);};
  const btn=(label,fn,cls,title)=>{const b=make('button',label,cls);b.type='button';b.onclick=fn;if(title){b.title=title;b.setAttribute('aria-label',title);}return b;};
  const select=(pairs,value,fn,label)=>{const s=make('select');s.setAttribute('aria-label',label);for(const [v,t]of pairs){const o=make('option',t);o.value=v;s.append(o);}s.value=value;s.onchange=()=>fn(s.value);return s;};
  const check=(label,value,fn)=>{const l=make('label',undefined,'sort-check'),i=make('input');i.type='checkbox';i.checked=value;i.onchange=()=>fn(i.checked);l.append(i,document.createTextNode(label));return l;};
  function dirty(){valid=false;page=null;offset=0;renderPreview();sync();say('规则已修改，点击“预览排序”检查新的顺序。');}
  function sync(){document.querySelectorAll('.sort-workspace button,.sort-workspace input,.sort-workspace select,.sort-workspace textarea,#actions button').forEach(n=>n.disabled=busy||n.dataset.disabled==='true');if($('sortApply'))$('sortApply').disabled=busy||!valid||applied;if($('sortUndo'))$('sortUndo').disabled=busy||!canUndo;}
  async function request(action,extra={}){
    if(busy)return;busy=true;sync();say(action==='sortPreview'?'正在生成完整排序计划…':'正在处理…');
    try{return await window.gridpilot.request({action,options,...extra});}catch(e){say(e.message.replace(/^Error invoking remote method 'request': Error: /,''),true);return null;}
    finally{busy=false;sync();}
  }
  async function capture(action){const result=await request(action);if(!result)return;data=result;options=structuredClone(result.options);valid=applied=canUndo=false;page=null;offset=0;search='';render();}
  async function preview(){const result=await request('sortPreview',{offset:0,size,search});if(!result){valid=false;sync();return;}page=result;valid=true;offset=result.offset;renderPreview();sync();$('sortPreview').parentElement.scrollIntoView({block:'start',behavior:'smooth'});say(`已预览 ${result.previewed.toLocaleString()} / ${result.total.toLocaleString()} 行，${options.Levels.length} 个层级。${result.records<result.total?'跨行合并按完整记录移动。':''}`);}
  async function turn(next){const result=await request('sortPage',{offset:next,size,search});if(result){page=result;offset=result.offset;renderPreview();say('预览已更新。');}}
  function levelRow(level,i){
    const row=make('div',undefined,'sort-level'),top=make('div',undefined,'sort-level-top');row.dataset.level=i;
    top.append(make('span',String(i+1).padStart(2,'0'),'sort-priority'));
    top.append(select(data.headers.map((h,c)=>[c,h]),level.Column,v=>{level.Column=Number(v);dirty();},`第 ${i+1} 级排序列`));
    top.append(select([['values','Excel 原生 · A–Z / 数值 / 日期'],['natural','自然顺序 · 项目2 < 项目10'],['custom','自定义序列 / 中文数字 / 正则']],level.Mode,v=>{level.Mode=v;dirty();renderLevels();},`第 ${i+1} 级排序方式`));
    const directions=make('div',undefined,'sort-directions');
    for(const [desc,label] of [[false,'↑ 升序'],[true,'↓ 降序']]){const b=btn(label,()=>{level.Descending=desc;dirty();renderLevels();},level.Descending===desc?'selected':'');b.setAttribute('aria-pressed',String(level.Descending===desc));b.setAttribute('aria-label',`第 ${i+1} 级${desc?'降序':'升序'}`);directions.append(b);}top.append(directions);
    const move=(delta)=>{const j=i+delta;[options.Levels[i],options.Levels[j]]=[options.Levels[j],options.Levels[i]];dirty();renderLevels();};
    const up=btn('↑',()=>move(-1),'sort-icon','提高优先级'),down=btn('↓',()=>move(1),'sort-icon','降低优先级'),remove=btn('×',()=>{options.Levels.splice(i,1);dirty();renderLevels();},'sort-icon','删除此层级');
    up.dataset.disabled=String(i===0);down.dataset.disabled=String(i===options.Levels.length-1);remove.dataset.disabled=String(options.Levels.length===1);top.append(up,down,remove);row.append(top);
    if(level.Mode==='values')row.append(check('数字文本按数值排序（如“2”和“10”）',level.TextAsNumbers,v=>{level.TextAsNumbers=v;dirty();}));
    if(level.Mode==='custom'){
      const disclosure=make('details',undefined,'sort-rule-details'),summary=make('summary','编辑规则 · '+(level.Rules||'未设置').split('\n').join(' / '));disclosure.append(summary);
      const detail=make('div',undefined,'sort-rule-grid');
      for(const [key,label,placeholder]of [['Rules','匹配规则 · 每行一条，按行优先','(国际)[中文数字]\n(业务)[中文数字]'],['Sequences','自定义序列 · 名称=顺序','部门=国际部,业务部,综合部']]){const l=make('label',label),t=make('textarea');t.value=level[key]??'';t.placeholder=placeholder;t.rows=2;t.maxLength=8192;t.setAttribute('aria-label',`第 ${i+1} 级${label}`);t.oninput=()=>{level[key]=t.value;dirty();};l.append(t);detail.append(l);}
      const presets=select([['','插入规则示例…'],['chinese','中文数字：一、二、十'],['size','大、中、小、微'],['combined','组合：国际 / 业务 + 中文数字'],...data.sequences.map((s,n)=>['excel'+n,s.name])],'',v=>{
        if(!v)return;if(v==='chinese'){level.Rules='.*?[中文数字]';level.Sequences='';}
        else if(v==='size'){level.Rules='[大中小微]';level.Sequences='';}
        else if(v==='combined'){level.Rules='(国际)[中文数字]\n(业务)[中文数字]';level.Sequences='';}
        else{const seq=data.sequences[Number(v.slice(5))];level.Rules='[Excel序列]$';level.Sequences='Excel序列='+seq.values.join(',');}
        dirty();renderLevels();
      },`第 ${i+1} 级规则示例`);detail.append(presets,make('small','先按规则行分组，再按标记升降序；未匹配项在最后，相同键继续比较下一层。'));disclosure.append(detail);row.append(disclosure);
    }
    return row;
  }
  function renderLevels(){
    $('sortLevels').replaceChildren(...options.Levels.map(levelRow));$('sortCount').textContent=`${options.Levels.length} / ${data.kind==='sheets'?1:20} 级`;
    $('sortAdd').dataset.disabled=String(options.Levels.length>=20||data.kind==='sheets'||!data.headers?.length);sync();
  }
  function renderPreview(){
    const area=$('sortPreview');if(!area)return;area.replaceChildren();
    if(!page){area.append(make('div',applied?'已应用排序。重新读取后可以继续调整。':'先设置排序优先级，再预览所有记录的排序结果。','sort-empty'));return;}
    const bar=make('div',undefined,'sort-preview-bar');bar.append(make('strong',`${page.records.toLocaleString()} 条记录 · ${page.total.toLocaleString()} 行`));
    const searchInput=make('input');searchInput.type='search';searchInput.value=search;searchInput.placeholder='搜索预览内容';searchInput.setAttribute('aria-label','搜索排序预览');
    const find=()=>{search=searchInput.value;turn(0);};searchInput.onkeydown=e=>{if(e.key==='Enter'){e.preventDefault();find();}};bar.append(searchInput,btn('搜索',find));area.append(bar);
    if(page.total>page.previewed)area.append(make('p',`排序处理全部 ${page.total.toLocaleString()} 行；分页预览显示排序后的前 20,000 行，搜索限于这些行。`,'sort-notice'));
    const wrap=make('div',undefined,'sort-table-wrap'),table=make('table'),head=make('thead'),tr=make('tr');
    ['新位置','原始行',...page.headers].forEach(x=>tr.append(make('th',x||'（空表头）')));head.append(tr);table.append(head);const body=make('tbody');
    for(const row of page.rows){const r=make('tr');for(const val of [row.position,row.sourceRow,...row.values]){const cell=make('td',String(val??''));cell.title=String(val??'');r.append(cell);}body.append(r);}table.append(body);wrap.append(table);area.append(wrap);
    const pager=make('div',undefined,'sort-pager');const prev=btn('← 上一页',()=>turn(Math.max(0,offset-size))),next=btn('下一页 →',()=>turn(offset+size));prev.dataset.disabled=String(offset===0);next.dataset.disabled=String(offset+size>=page.filtered);
    const jump=make('input');jump.type='number';jump.min=1;jump.max=Math.max(1,Math.ceil(page.filtered/size));jump.value=Math.floor(offset/size)+1;jump.setAttribute('aria-label','跳转到预览页');jump.onchange=()=>turn((Math.max(1,Math.min(Number(jump.max),Number(jump.value)||1))-1)*size);
    pager.append(prev,make('span',page.filtered?`${offset+1}–${Math.min(offset+size,page.filtered)} / ${page.filtered.toLocaleString()} 行`:'没有匹配记录'),jump,make('span',`/ ${jump.max} 页`),next,select([[50,'50 条 / 页'],[100,'100 条 / 页'],[200,'200 条 / 页']],size,v=>{size=Number(v);turn(0);},'每页条数'));area.append(pager);sync();
  }
  function render(){
    document.body.classList.add('sorting');$('title').textContent=data.kind==='sheets'?'工作表排序':'多级联动排序';$('eyebrow').textContent='排序工作区';$('subtitle').textContent=data.description||'在 Excel 选中数据区域，或插入一个示例工作簿。';
    $('nav').replaceChildren(make('div','先部门，再客户，最后项目。每层可独立选择升降序。','sort-guide'));
    document.querySelector('.local').textContent='本地处理 · 预览后应用 · 可恢复原序';
    const content=$('content');content.classList.add('sort-workspace');content.replaceChildren();
    const toolbar=make('div',undefined,'sort-toolbar');toolbar.append(btn('重新读取选区',()=>capture('sortCapture')));
    if(data.kind!=='sheets'){toolbar.append(btn('插入示例工作簿',()=>capture('sortSample')));toolbar.append(btn('Excel 原生选项…',async()=>{const r=await request('sortNative');if(r){valid=applied=canUndo=false;page=null;renderPreview();sync();say(r.message);}},'quiet','打开 Excel 原生排序，支持颜色、图标和按行排序'));
      const h=check('首行为表头',options.HasHeader,v=>{options.HasHeader=v;dirty();});h.querySelector('input').dataset.disabled=String(data.isTable);toolbar.append(h);}
    toolbar.append(check('区分大小写',options.MatchCase,v=>{options.MatchCase=v;dirty();}));
    if(data.kind!=='sheets')toolbar.append(select([[false,'中文按拼音'],[true,'中文按笔画']],options.StrokeOrder,v=>{options.StrokeOrder=v==='true';dirty();},'中文原生排序方法'));
    content.append(toolbar);if(data.error)content.append(make('p',data.error,'sort-notice'));
    if(data.merges)content.append(make('p',`${data.merges} 处合并：横向合并随行移动；跨行合并及相交块作为一条完整记录，以首行排序。`,'sort-notice'));
    const card=make('section',undefined,'sort-card'),heading=make('div',undefined,'sort-heading');heading.append(make('h2','排序优先级'));const count=make('span');count.id='sortCount';heading.append(count);
    const example=btn('使用部门 → 客户 → 项目示例',()=>{if(!data.headers||data.headers.length<3){say('此示例需要至少三列，可先插入示例工作簿。',true);return;}options.Levels=[{Column:0,Mode:'custom',Rules:'(国际)[中文数字]\n(业务)[中文数字]',Sequences:'',Descending:false},{Column:1,Mode:'values',Descending:false},{Column:2,Mode:'custom',Rules:'.*?[中文数字]',Sequences:'',Descending:true}];dirty();renderLevels();},'quiet');
    if(data.kind!=='sheets')heading.append(example);card.append(heading);const levels=make('div');levels.id='sortLevels';card.append(levels);
    const add=btn('＋ 添加层级',()=>{const used=new Set(options.Levels.map(l=>l.Column));options.Levels.push({Column:data.headers.findIndex((_,c)=>!used.has(c))<0?0:data.headers.findIndex((_,c)=>!used.has(c)),Mode:'values',Descending:false,Rules:'[中文数字]',Sequences:''});dirty();renderLevels();});add.id='sortAdd';card.append(add);content.append(card);
    const previewSection=make('section',undefined,'sort-card');previewSection.append(make('h2','排序结果预览'));const target=make('div');target.id='sortPreview';previewSection.append(target);content.append(previewSection);
    const undo=btn('恢复原序',async()=>{const r=await request('sortUndo');if(r){canUndo=valid=applied=false;page=null;renderPreview();sync();say(r.message+' 请重新读取后继续。');}});undo.id='sortUndo';
    const see=btn('预览排序',preview);see.id='sortBuild';see.dataset.disabled=String(!data.headers?.length);
    const apply=btn('应用排序',async()=>{const r=await request('sortApply');if(r){applied=true;canUndo=r.canUndo;sync();say(r.message);}},'primary');apply.id='sortApply';$('actions').replaceChildren(undo,see,apply);
    if(!data.headers?.length)options.Levels=[];renderLevels();renderPreview();sync();say(data.error||'按优先级从上向下比较，相同键保持原序。完整预览最多 20,000 行。',!!data.error);
  }
  render();
};
