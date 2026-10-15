// Suggestions are explicit drafts. Never infer regulatory applicability from a sheet name alone.
const crypto=require('node:crypto');
function suggest(reports,file){
  const config={schema:1,name:'余额变动核对 · 识别草稿',previousDate:'',currentDate:'',tolerance:'0.01',displayUnit:'10000',sources:[],rules:[]};
  const bindings=[];
  for(const [role,report] of Object.entries(reports)){
    const h=report.rows.find(r=>r.values.includes('贷款号')&&r.values.some(v=>['余额折人民币','实际发放金额（折人民币）','实际回收金额(折人民币)'].includes(v)));if(!h)continue;
    const headers=h.values.map((v,i)=>String(v??'').trim()||'列'+(i+1));
    const s={id:crypto.randomUUID(),name:report.sheet,path:file,sheet:report.sheet,sheets:report.sheets,headerRow:h.row,startRow:h.row+1,endRow:0,headers};config.sources.push(s);
    const amount=['余额折人民币','实际发放金额（折人民币）','实际回收金额(折人民币)'].find(v=>headers.includes(v));
    const key=['贷款号','币种代码'];if(['issue','recovery'].includes(role))key.push(...['发放号','交易日期',...(role==='recovery'?['回收类型代码']:[])].filter(v=>headers.includes(v)));
    const group=[headers.includes('客户名')?'客户名':'客户名称',headers.includes('项目名')?'项目名':'项目名称'];
    bindings.push({role,source:s.id,key:key.filter(v=>headers.includes(v)),group,amount,unit:'1',sign:1,date:['issue','recovery'].includes(role)?'交易日期':'',filters:[],joins:[]});
    if(['previous','current'].includes(role)){const first=report.rows.find(r=>r.row>h.row),month=String(first?.values[headers.indexOf('月份')]||'');if(/^\d{6}$/.test(month)){const y=Number(month.slice(0,4)),m=Number(month.slice(4));if(m>=1&&m<=12)config[role+'Date']=new Date(Date.UTC(y,m,0)).toISOString().slice(0,10);}}
  }
  config.rules.push({id:crypto.randomUUID(),name:'余额与发放回收 · 全表总额（请复核统计范围）',kind:'balance',topN:5,reviewed:false,requiresReview:true,bindings});
  return config;
}
module.exports={suggest};
