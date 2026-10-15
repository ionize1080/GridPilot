'use strict';
// All monetary arithmetic uses integer millionths of a yuan. No eval or floating point summation.
const crypto=require('node:crypto');
const text=v=>v==null?'':String(v).trim();
const bad=v=>v&&typeof v==='object'&&v.error;
function money(v,unit='1'){
  if(bad(v)||!text(v))throw Error('金额为空或公式错误');
  let s=text(v);if(s.includes(',')){if(!/^[+-]?\d{1,3}(,\d{3})*(\.\d+)?$/.test(s))throw Error('金额分隔符无效');s=s.replaceAll(',','');}
  const e=s.match(/^([+-]?)(\d+)(?:\.(\d*))?[eE]([+-]?\d+)$/);
  if(e){const shift=Number(e[4]);if(Math.abs(shift)>24)throw Error('金额指数过大');const d=e[2]+(e[3]||''),p=e[2].length+shift;s=e[1]+(p<=0?'0.'+'0'.repeat(-p)+d:p>=d.length?d+'0'.repeat(p-d.length):d.slice(0,p)+'.'+d.slice(p));}
  const m=s.match(/^([+-]?)(\d+)(?:\.(\d+))?$/);if(!m||s.length>80)throw Error('金额不是有效数字');
  if(!['1','10000','100000000'].includes(String(unit)))throw Error('金额单位无效');
  const fraction=m[3]||'',digits=BigInt(m[2]+fraction),denom=10n**BigInt(fraction.length),numerator=digits*BigInt(unit)*1000000n;
  return ((numerator+denom/2n)/denom)*(m[1]==='-'?-1n:1n);
}
function format(v,unit='10000',places=2){const n=BigInt(v),d=BigInt(unit)*1000000n,p=10n**BigInt(places),x=((n<0n?-n:n)*p+d/2n)/d;return (n<0n?'-':'')+(x/p).toString()+(places?'.'+(x%p).toString().padStart(places,'0'):'');}
const abs=n=>n<0n?-n:n;
function validateConfig(c){
  if(!c||c.schema!==1||typeof c.name!=='string'||!Array.isArray(c.sources)||!Array.isArray(c.rules)||c.sources.length>80||c.rules.length>500)throw Error('核对方案格式无效或超出规模限制');
  if(!['1','10000','100000000'].includes(c.displayUnit)||money(c.tolerance)<0n)throw Error('方案单位或容差无效');
  for(const list of [c.sources,c.rules])if(new Set(list.map(x=>x.id)).size!==list.length||list.some(x=>typeof x.id!=='string'||!x.id))throw Error('来源或指标标识重复 / 无效');
  for(const s of c.sources)if(typeof s.path!=='string'||typeof s.sheet!=='string'||typeof s.name!=='string'||!Array.isArray(s.headers)||s.headers.some(h=>typeof h!=='string'))throw Error('来源配置格式无效');
  for(const r of c.rules){if(typeof r.name!=='string'||!Array.isArray(r.bindings))throw Error('指标配置格式无效');for(const b of r.bindings){if(!Array.isArray(b.key)||!Array.isArray(b.group)||!Array.isArray(b.filters)||!Array.isArray(b.joins))throw Error('取数项配置格式无效');}}
  return c;
}
function field(row,name){if(!Object.hasOwn(row,name))throw Error('字段不存在：'+name);const v=row[name];if(bad(v))throw Error(name+'：'+v.error);return v;}
function matches(row,filters=[]){
  // Validate every predicate, even if an earlier condition fails.
  const results=filters.map(f=>{const v=text(field(row,f.field)),a=text(f.value);switch(f.op){case 'eq':return v===a;case 'ne':return v!==a;case 'in':return a.split('|').map(text).includes(v);case 'prefix':return v.startsWith(a);case 'contains':return v.includes(a);case 'empty':return !v;case 'gt':return money(v)>money(a);case 'lt':return money(v)<money(a);default:throw Error('未知筛选条件');}});return results.every(Boolean);
}
function table(report,source){
  if(!Number.isInteger(source.headerRow)||source.headerRow<1||source.headerRow>100||!Number.isInteger(source.startRow)||source.startRow<=source.headerRow||!Number.isInteger(source.endRow||0)||(source.endRow&&source.endRow<source.startRow))throw Error('请检查表头行和数据起止行');
  const header=report.rows.find(r=>r.row===source.headerRow);if(!header)throw Error('找不到表头行');
  if(header.values.some(bad))throw Error('表头包含公式错误');const headers=header.values.map((v,i)=>text(v)||'列'+(i+1));if(new Set(headers).size!==headers.length||headers.includes('__ref'))throw Error('表头重复或使用保留名称，请选择正确表头行或修正来源');
  return {headers,rows:report.rows.filter(r=>r.row>=(source.startRow||source.headerRow+1)&&(!source.endRow||r.row<=source.endRow)).map(r=>{const out=Object.create(null);headers.forEach((h,i)=>out[h]=r.values[i]??'');out.__ref={source:source.id,file:source.path,sheet:source.sheet,row:r.row};return out;})};
}
function validatePeriod(c){for(const k of ['previousDate','currentDate']){if(!/^\d{4}-\d{2}-\d{2}$/.test(c[k]||'')||!Number.isFinite(Date.parse(c[k]))||new Date(c[k]).toISOString().slice(0,10)!==c[k])throw Error('请设置有效的两期日期');}if(c.previousDate>=c.currentDate)throw Error('当期日期必须晚于上期');}
function reconcile(config,tables){
  validatePeriod(config);if(!Array.isArray(config.rules)||!config.rules.length)throw Error('请至少添加一个指标规则');
  const issues=[],results=[],tol=money(config.tolerance||'0.01'),unit=config.displayUnit||'10000';if(tol<0n)throw Error('容差不能为负');format(0n,unit);
  const issue=(rule,ref,message)=>issues.push({rule,ref,message});
  for(const rule of config.rules){
    if(rule.skip){results.push({id:rule.id,name:rule.name,status:'skipped',reason:rule.skipReason||'用户跳过',details:[]});continue;}
    const begin=issues.length;
    try{
      if(rule.requiresReview&&!rule.reviewed)throw Error('自动识别的规则尚未复核，请在指标规则中确认映射和统计范围');
      if(!['balance','financial'].includes(rule.kind))throw Error('请选择余额或财务指标');
      if(!rule.bindings?.length)throw Error('请添加数据取数项');
      const groups=new Map(),tot={previous:0n,current:0n,issue:0n,recovery:0n,adjustment:0n},roles=new Set();let movesIn=0n,movesOut=0n;
      const snapshots={previous:new Map(),current:new Map()};
      for(const b of rule.bindings){
        if(!Object.hasOwn(tot,b.role))throw Error('取数角色无效');roles.add(b.role);
        const t=tables[b.source];if(!t||t.error)throw Error(t?.error||'来源尚未读取');
        if(!t.rows.length&&['previous','current'].includes(b.role))throw Error('时点来源没有数据行，不能视为零余额');
        if(!b.amount||!b.key?.length)throw Error('请指定金额及业务组合键');
        const available=[...t.headers,...(b.joins||[]).flatMap(j=>Object.keys(j.fields||{}))];
        for(const f of [b.amount,...b.key,...(b.group||[]),...(b.filters||[]).map(f=>f.field),...(b.date?[b.date]:[])])if(!available.includes(f))throw Error('来源缺少字段：'+f);
        if(b.sign!==1&&b.sign!==-1)throw Error('计算方向无效');
        const seen=new Set();let joins=(b.joins||[]).map(j=>{const d=tables[j.source];if(!d||d.error)throw Error('专项表不可用');if(!j.left?.length||j.left.length!==j.right?.length)throw Error('专项表两侧关联键数量必须一致且非空');for(const f of [...j.right,...Object.values(j.fields||{})])if(!d.headers.includes(f))throw Error('专项表字段不存在：'+f);const index=new Map();for(const r of d.rows){let parts;try{parts=j.right.map(f=>text(field(r,f)));}catch{continue;}if(parts.some(v=>!v))continue;const k=JSON.stringify(parts);if(index.has(k))index.set(k,null);else index.set(k,r);}return {...j,index};});
        for(const original of t.rows){
          try{
            const row={...original};
            for(const j of joins){const parts=j.left.map(f=>text(field(row,f)));if(parts.some(x=>!x))throw Error('专项表关联键为空');const hit=j.index.get(JSON.stringify(parts));if(!hit)throw Error(hit===null?'专项表关联键重复（禁止多对多放大）':'专项表未匹配');for(const [alias,f] of Object.entries(j.fields||{})){if(Object.hasOwn(row,alias))throw Error('专项字段别名与原字段重名');row[alias]=field(hit,f);}}
            const keys=b.key.map(f=>text(field(row,f)));if(keys.some(x=>!x))throw Error('业务键为空');const key=JSON.stringify(keys);
            // Snapshot keys must identify a loan/account, while flows need transaction IDs. No arbitrary deduplication.
            if(seen.has(key))throw Error('业务键重复，请补充组合键或拆分取数项');seen.add(key);
            if(b.date){const {dateValue}=require('./database-core.cjs');const date=dateValue(field(row,b.date),t.date1904);if(['previous','current'].includes(b.role)){if(date!==config[b.role==='previous'?'previousDate':'currentDate'])throw Error('时点日期与核对期间不一致');}else if(date<=config.previousDate||date>config.currentDate)continue;}
            const amount=money(field(row,b.amount),b.unit||'1')*(b.sign===-1?-1n:1n),included=matches(row,b.filters);
            const labels=(b.group?.length?b.group:b.key).map(f=>text(field(row,f))||'（空）'),groupKey=JSON.stringify(labels);
            if(rule.kind==='balance'&&['previous','current'].includes(b.role)){if(snapshots[b.role].has(key))throw Error('同一时点的取数项重叠，请避免重复取数');snapshots[b.role].set(key,{amount,included,groupKey,labels,ref:row.__ref});}
            if(!included)continue;
            if(!groups.has(groupKey))groups.set(groupKey,{key:groupKey,labels,previous:0n,current:0n,issue:0n,recovery:0n,adjustment:0n,refs:[]});
            const g=groups.get(groupKey);g[b.role]+=amount;tot[b.role]+=amount;g.refs.push({...row.__ref,role:b.role,amount:amount.toString(),businessKey:keys.join(' / ')});
          }catch(e){issue(rule.id,original.__ref,e.message);}
        }
      }
      if(!roles.has('previous')||!roles.has('current'))throw Error('必须配置当期和上期取数');
      for(const [k,now] of snapshots.current){const old=snapshots.previous.get(k);if(old&&!old.included&&now.included)movesIn+=old.amount;if(old&&old.included&&!now.included)movesOut+=old.amount;}
      const delta=tot.current-tot.previous,flow=tot.issue-tot.recovery+tot.adjustment,flowComplete=rule.kind==='balance'&&roles.has('issue')&&roles.has('recovery');
      let rp=null,rc=null,rd=null,reportError=null,reportRefs=[];
      if(rule.report){try{const r=rule.report,t=tables[r.source];if(!t||t.error)throw Error('监管报表不可用');const rows=t.rows.filter(row=>text(field(row,r.codeField))===text(r.code));if(rows.length!==1)throw Error('监管指标代码未匹配或重复');const row=rows[0];rp=money(field(row,r.previous),r.unit||'1');rc=money(field(row,r.current),r.unit||'1');rd=r.delta?money(field(row,r.delta),r.unit||'1'):rc-rp;if(abs(rc-rp-rd)>tol)throw Error('监管报表当期减上期不等于差值');reportRefs=[row.__ref];}catch(e){reportError=e.message;issue(rule.id,null,e.message);}}
      const gap=rd===null?null:delta-rd,previousGap=rp===null?null:tot.previous-rp,currentGap=rc===null?null:tot.current-rc,residual=flowComplete?delta-flow:null;
      const status=issues.length>begin?'partial':rd===null?'unverified':abs(gap)>tol||abs(previousGap)>tol||abs(currentGap)>tol||(residual!==null&&abs(residual)>tol)?'difference':'matched';
      const details=[...groups.values()].map(g=>({...g,delta:g.current-g.previous,residual:flowComplete?g.current-g.previous-g.issue+g.recovery-g.adjustment:null})).sort((a,b)=>abs(a.delta)>abs(b.delta)?-1:abs(a.delta)<abs(b.delta)?1:a.key.localeCompare(b.key));
      const top=details.filter(d=>d.delta!==0n).slice(0,Math.max(1,Math.min(20,Number(rule.topN)||5))),rest=delta-top.reduce((n,d)=>n+d.delta,0n);
      const unitName={'1':'元','10000':'万元','100000000':'亿元'}[unit];
      const fmt=n=>format(n,unit),signed=n=>(n>=0n?'增加':'减少')+fmt(abs(n));
      let narrative=`${rule.name}：${config.currentDate}为${fmt(tot.current)}${unitName}，较${config.previousDate}${signed(delta)}${unitName}。`;
      if(top.length)narrative+='主要变动：'+top.map(d=>d.labels.join(' / ')+signed(d.delta)+unitName).join('；')+`；其他净变动${fmt(rest)}${unitName}。`;
      if(flowComplete)narrative+=`期间发放${fmt(tot.issue)}、回收${fmt(tot.recovery)}、已登记调整${fmt(tot.adjustment)}${unitName}，未解释差额${fmt(residual)}${unitName}。`;
      if(rule.template){
        if(typeof rule.template!=='string'||rule.template.length>5000)throw Error('报告模板最多 5000 字');
        const tokens={'指标':rule.name,'当期日':config.currentDate,'上期日':config.previousDate,'当期':fmt(tot.current),'上期':fmt(tot.previous),'净变动':fmt(delta),'单位':unitName,'主要变动':top.map(d=>d.labels.join(' / ')+signed(d.delta)+unitName).join('；'),'其他净变动':fmt(rest),'发放':roles.has('issue')?fmt(tot.issue):'未配置','回收':roles.has('recovery')?fmt(tot.recovery):'未配置','调整':fmt(tot.adjustment),'未解释差额':flowComplete?fmt(residual):'未核对'};
        narrative=rule.template.replace(/\{([^{}]+)\}/g,(_,key)=>{if(!Object.hasOwn(tokens,key))throw Error('未知报告占位符：'+key);return tokens[key];})+'\n';
      }
      if(movesIn||movesOut)narrative+=`按两期各自分类，存量迁入${fmt(movesIn)}、迁出${fmt(movesOut)}${unitName}（以上期余额计，已包含在余额变动中，不重复相加）。`;
      narrative+=status==='matched'?'两期金额及差额在容差内一致。':status==='partial'?'存在被跳过的异常记录，本草稿不完整，需核实后定稿。':status==='unverified'?'尚未绑定监管金额，待核对。':`监管差额偏差${fmt(gap)}${unitName}，需核查未解释事项。`;
      if(rule.kind==='balance'&&!flowComplete)narrative+='未配置完整发放和回收，仅完成余额与监管金额核对。';
      results.push({id:rule.id,name:rule.name,status,kind:rule.kind,tot,delta,flow:flowComplete?flow:null,residual,report:{previous:rp,current:rc,delta:rd,refs:reportRefs,error:reportError},gap,previousGap,currentGap,movesIn,movesOut,details,narrative,rest,issueCount:issues.length-begin});
    }catch(e){issue(rule.id,null,e.message);results.push({id:rule.id,name:rule.name,status:'error',reason:e.message,details:[]});}
  }
  const payload={version:1,createdAt:new Date().toISOString(),configHash:crypto.createHash('sha256').update(JSON.stringify(config)).digest('hex'),period:{previous:config.previousDate,current:config.currentDate},unit,results,issues};
  return JSON.parse(JSON.stringify(payload,(_,v)=>typeof v==='bigint'?v.toString():v));
}
module.exports={money,format,table,matches,reconcile,validatePeriod,validateConfig};
