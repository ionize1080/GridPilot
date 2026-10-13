// Development-only UI test. Launches the production Electron application with synthetic pipe data.
const path=require('path'),fs=require('fs'),assert=require('assert');
const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const root=path.resolve(__dirname,'..');
const options={SchemaVersion:1,IconTheme:0,MaskPreserveFormatting:true,FocusRipple:true,RippleDuration:280,RippleWidth:36,RippleStrength:24,FocusColor:'#42AD91',FocusOpacity:20,FocusRow:true,FocusColumn:true,ClearActiveCell:true,FocusAnimation:false,FocusInterval:33,PreserveVisualAlignment:true,VisibleCellsOnly:false,PreserveLeadingZeros:true,ProtectPrecision:true,TrimNumericText:true,AllowPercent:false,AllowThousands:true,NumberAsTextHint:true,ConversionLimit:50000,DecimalCulture:0,Copy:{Scope:0,CellFormats:true,NumberFormats:true,RowHeights:true,ColumnWidths:true,MergedCells:true,ConditionalFormatting:0,FormulaFallback:0,CellLimit:1000000},Rename:{Find:'',Replacement:'',IncludeHidden:true,IncludeCharts:true,Numbering:false,Start:1,Step:1,Digits:2}};
async function main(mode){
 const env={...process.env};delete env.ELECTRON_RUN_AS_NODE;delete env.NODE_OPTIONS;const start=Date.now();
 const pipeName='GridPilot-'+require('crypto').randomBytes(16).toString('hex');let connection;
 const server=require('net').createServer(socket=>{connection=socket;let buffer='';socket.on('data',chunk=>{buffer+=chunk.toString();let i;while((i=buffer.indexOf('\n'))>=0){const line=buffer.slice(0,i);buffer=buffer.slice(i+1);try{respond(JSON.parse(line));}catch{}}});});
 await new Promise(resolve=>server.listen('\\\\.\\pipe\\'+pipeName,resolve));
 const child=require('child_process').spawn(path.join(root,'.temp/Browser/electron.exe'),['--remote-debugging-port=0',path.join(root,'App'),'--gridpilot-pipe='+pipeName],{env,stdio:['ignore','pipe','pipe']});
 const endpoint=new Promise((resolve,reject)=>{child.stderr.on('data',data=>{const m=data.toString().match(/DevTools listening on (ws:\/\/[^\s]+)/);if(m)resolve(m[1]);});child.on('exit',code=>reject(new Error('Early exit '+code)));});
 const fixture={mode,settings:options,sheets:[{id:0,name:'客户总表'},{id:1,name:'贷款明细'}],cells:12648,rules:[{Sheet:0,SheetName:'客户总表',Column:1,Header:'客户名称',StartRow:5,Kind:1,Domain:'客户',Factor:1.7,DayShift:37,Reason:'根据合并表头建议'}, {Sheet:0,SheetName:'客户总表',Column:2,Header:'信用代码',StartRow:5,Kind:3,Domain:'信用代码',Factor:1.7,DayShift:37},{Sheet:0,SheetName:'客户总表',Column:3,Header:'本期贷款余额',StartRow:5,Kind:0,Domain:'金额',Factor:1.7,DayShift:37},{Sheet:1,SheetName:'贷款明细',Column:1,Header:'借款人名称',StartRow:3,Kind:1,Domain:'客户',Factor:1.7,DayShift:37,Reason:'与客户总表有重复值'}]};
 const respond=m=>{let data=m.action==='init'?fixture:m.action==='preview'?{summary:'已验证 328 个常量替换、120 条公式。',rows:[{location:'S1!A5',before:'示例客户甲',after:'青禾有限公司'},{location:'S2!A3',before:'示例客户甲',after:'青禾有限公司'}]}:m.action==='saveSettings'?{saved:true,message:'已保存',settings:m.settings}:{message:'已取消保存'};connection.write(JSON.stringify({id:m.id,data})+'\n');};
 const browser=await chromium.connectOverCDP(await endpoint);const context=browser.contexts()[0];const page=context.pages()[0]||await context.waitForEvent('page');const errors=[];page.on('pageerror',e=>errors.push(e.message));await page.locator('#title').filter({hasText:mode==='mask'?'让数据可分享':'按你的习惯'}).waitFor();
 console.log(mode+' ready_ms='+String(Date.now()-start));
 fs.mkdirSync(path.join(root,'docs/v0.9'),{recursive:true});
 if(mode==='mask'){
  assert.strictEqual(await page.locator('tbody tr').count(),4);await page.screenshot({path:path.join(root,'docs/v0.9/masking.png')});
  await page.locator('#search').fill('信用');assert.strictEqual(await page.locator('tbody tr').count(),1);await page.locator('#search').fill('');
  await page.getByRole('button',{name:'分析与预览',exact:true}).click();await page.locator('dialog[open]').waitFor();assert.strictEqual(await page.locator('#previewRows tr').count(),2);assert(await page.locator('#export').isDisabled());await page.locator('#confirmed').check();assert(!(await page.locator('#export').isDisabled()));await page.screenshot({path:path.join(root,'docs/v0.9/preview.png')});await page.locator('#dismiss').click();
 }else{
  await page.getByRole('button',{name:'晴蓝'}).click();await page.getByRole('button',{name:'保存设置',exact:true}).click();await page.locator('#status').filter({hasText:'已保存'}).waitFor();assert.strictEqual(options.IconTheme,0);await page.screenshot({path:path.join(root,'docs/v0.9/settings.png')});
  await page.getByRole('button',{name:'聚焦与涟漪',exact:true}).click();await page.locator('#setting-RippleDuration').fill('320');await page.screenshot({path:path.join(root,'docs/v0.9/ripple.png')});
  for(const name of ['数据脱敏','原样复制','内容转换','批量重命名','导航布局'])await page.getByRole('button',{name,exact:true}).click();
 }
 await page.setViewportSize({width:800,height:620});assert(await page.evaluate(()=>document.documentElement.scrollWidth<=window.innerWidth));assert.deepStrictEqual(errors,[]);await page.locator('#close').click();await new Promise(resolve=>child.once('exit',resolve));await browser.close();server.close();console.log('PASS browser '+mode);
}
(async()=>{await main('mask');await main('settings');})().catch(e=>{console.error(e);process.exitCode=1;});
