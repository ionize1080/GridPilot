const {_electron:electron}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const assert=require('node:assert/strict'),fs=require('node:fs'),path=require('node:path'),os=require('node:os'),net=require('node:net'),crypto=require('node:crypto');
const root=path.resolve(__dirname,'..'),temporary=fs.mkdtempSync(path.join(os.tmpdir(),'GridPilot-browser-v11-'));
const appDirectory=process.env.GRIDPILOT_TEST_APP||path.join(root,'App');
const browserExecutable=process.env.GRIDPILOT_TEST_BROWSER||path.join(root,'.temp/Browser/electron.exe');
const headers=['月份','贷款号','余额','流水号','备注'];
const report=rows=>({sheet:'Data',sheets:['Data'],rows:[{row:1,values:['合成测试报表']},{row:2,values:headers},...rows.map((values,i)=>({row:i+3,values}))]});
async function test(mode){
  let application,socket,lastMessage;const folder=path.join(temporary,'reports');fs.mkdirSync(folder,{recursive:true});
  const rows=Array.from({length:20101},(_,i)=>['202402','LN'+String(i).padStart(6,'0'),'100.01','T'+i,'合成记录']);
  fs.writeFileSync(path.join(folder,'report.xlsx'),JSON.stringify(rows));
  const pipe='GridPilot-'+crypto.randomBytes(16).toString('hex');
  const send=m=>socket.write(JSON.stringify(m)+'\n');
  const server=net.createServer(s=>{socket=s;let buffer='';s.on('data',chunk=>{buffer+=chunk.toString();let end;while((end=buffer.indexOf('\n'))>=0){const m=JSON.parse(buffer.slice(0,end));buffer=buffer.slice(end+1);lastMessage=m;
    let data;
    if(m.action==='init')data=mode==='database'?{mode}:{mode,settings:{MaskPreserveFormatting:false,MaskReplaceMentions:false},cells:20101,sheets:[{id:0,name:'合成客户清单',rows:20101,columns:5,nonEmptyRows:20101,nonEmptyColumns:5}],rules:[{Sheet:0,SheetName:'合成客户清单',Column:1,Header:'客户名',Domain:'customer',StartRow:2,Kind:1,Prefix:'',Suffix:'',Digits:6,ShortName:false,Factor:1.7,DayShift:37}]};
    else if(m.action==='dbReadFile')data=report(JSON.parse(fs.readFileSync(m.path,'utf8')).slice(0,m.maxRows||undefined));
    else if(m.action==='preview'||m.action==='previewPage'){const offset=m.offset||0,n=m.search?1:20000;data={summary:'合成预览',offset,total:n,capped:!m.search,rows:Array.from({length:Math.min(m.pageSize||100,n-offset)},(_,i)=>({location:'S1!A'+(offset+i+1),before:'原客户'+(offset+i),after:'替代客户'+(offset+i)}))};}
    else if(m.action==='copyMask'){assert(m.confirmed);data={text:'替代客户\t100.01\r\n',html:'<table><tr><td>替代客户</td><td>100.01</td></tr></table>',rows:1,columns:2};}
    else if(m.action==='dbOpenSource')data={message:'已定位源文件'};
    else data={message:'已取消'};
    send({id:m.id,data});
  } });});
  await new Promise(resolve=>server.listen('\\\\.\\pipe\\'+pipe,resolve));
  const bootstrap=path.join(temporary,mode+'.cjs');fs.writeFileSync(bootstrap,`require('electron').app.setPath('appData',${JSON.stringify(temporary)});require(${JSON.stringify(path.join(appDirectory,'main.js'))});`);
  const env={...process.env};delete env.NODE_OPTIONS;delete env.ELECTRON_RUN_AS_NODE;
  try{
    application=await electron.launch({executablePath:browserExecutable,args:[bootstrap,'--gridpilot-pipe='+pipe],env});
    const page=await application.firstWindow(),errors=[];page.on('pageerror',e=>errors.push(e.message));
    if(mode==='database'){
      await page.getByRole('button',{name:'选择文件夹并创建',exact:true}).click();
      await application.evaluate(({dialog},folder)=>{dialog.showOpenDialog=async()=>({canceled:false,filePaths:[folder]});},folder);
      await page.getByRole('button',{name:'选择报表文件夹',exact:true}).click();await page.locator('#dbName').fill('合成余额月报');
      await page.getByRole('button',{name:'创建数据库',exact:true}).click();await page.getByRole('button',{name:'扫描更新',exact:true}).click();
      await page.locator('#dbResults tbody tr').first().waitFor();assert.equal(await page.locator('#dbResults tbody tr').count(),100);
      await page.getByRole('button',{name:'下一页',exact:true}).click();await page.locator('#dbPaging').filter({hasText:'第 101–200 条'}).waitFor();
      await page.locator('#dbSql').fill('SELECT COUNT(*) AS n, SUM(__amount_minor) AS exact_amount FROM data');await page.getByRole('button',{name:'执行查询',exact:true}).click();await page.locator('#dbResults').filter({hasText:'201030101'}).waitFor();
      assert((await page.locator('#dbResults').innerText()).includes('20101'));
      await page.locator('#dbSql').fill('DELETE FROM source_rows');await page.getByRole('button',{name:'执行查询',exact:true}).click();await page.locator('#status.error').waitFor();
      await page.getByRole('button',{name:'全部记录',exact:true}).click();await page.locator('#dbResults tbody tr').first().waitFor();
      await page.getByRole('button',{name:'添加条件',exact:false}).click();await page.locator('.db-filters select').first().selectOption('贷款号');await page.locator('.db-filters input').fill('LN020100');await page.getByRole('button',{name:'生成 SQL',exact:true}).click();await page.getByRole('button',{name:'执行查询',exact:true}).click();await page.locator('#dbPaging').filter({hasText:'第 1–1 条'}).waitFor();
      assert((await page.locator('#dbSql').inputValue()).includes('LN020100'));
      await page.getByRole('button',{name:'打开 Excel',exact:true}).click();assert.equal(lastMessage.action,'dbOpenSource');assert.equal(lastMessage.row,20103);
      await page.locator('#dbResults').getByRole('button',{name:'所有来源',exact:true}).click();await page.locator('#dbResults th').filter({hasText:'__source_id'}).waitFor();
      await page.getByRole('button',{name:'全部记录',exact:true}).click();await page.locator('#dbResults tbody tr').first().waitFor();
      // Documentation screenshot replaces only the temporary user's directory with a synthetic display path.
      await page.evaluate(folder=>{const walk=document.createTreeWalker(document.body,NodeFilter.SHOW_TEXT);let node;while(node=walk.nextNode())if(node.textContent.includes(folder))node.textContent=node.textContent.replaceAll(folder,'D:/报表/余额样本');},folder);
      await page.screenshot({path:path.join(root,'docs/v1.1/database.png')});
      await page.setViewportSize({width:800,height:620});assert(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth));
      console.log('PASS database production UI: folder wizard, scan 20101 rows, paging, visual SQL, exact sum, read-only guard, source row, provenance, narrow layout');
    }else if(mode==='settings'){
      await page.getByRole('button',{name:'关于与许可',exact:true}).click();await page.locator('.legal-text').filter({hasText:'Permission'}).waitFor();
      const result=await page.evaluate(async()=>{const all=await window.gridpilot.request({action:'legal',index:0});const index=all.titles.findIndex(t=>t.includes('LICENSES.chromium'));if(index<0)throw new Error('Chromium license missing');return window.gridpilot.request({action:'legal',index});});assert(result.titles.length>=7);assert(result.text.length>0);assert(result.length>40000);console.log('PASS clean package bundled licenses visible in settings, including Chromium');
    }else{
      await page.getByRole('button',{name:'分析与预览',exact:true}).click();await page.locator('dialog[open]').waitFor();assert.equal(await page.locator('#previewRows tr').count(),100);assert(await page.locator('#copyMask').isDisabled());
      await page.locator('#previewNext').click();await page.locator('#previewPageInfo').filter({hasText:'101–200'}).waitFor();
      await page.locator('#previewSearch').fill('原客户20101');await page.locator('#previewSearchButton').click();await page.locator('#previewPageInfo').filter({hasText:'1–1 / 1'}).waitFor();
      await page.locator('#confirmed').check();await page.locator('#copyMask').click();await page.locator('#exportProgress').filter({hasText:'已复制'}).waitFor();
      const clipboard=await application.evaluate(async({clipboard})=>{const items=await clipboard.read();return {text:await clipboard.readText(),html:await(await items[0].getType('text/html')).text()};});assert.equal(clipboard.text,'替代客户\t100.01\r\n');assert(clipboard.html.includes('替代客户'));assert(await page.locator('dialog').evaluate(d=>d.open));
      await page.screenshot({path:path.join(root,'docs/v1.1/mask-preview.png')});
      console.log('PASS masking production UI: 20k count, pagination, search, confirmation gate, system HTML/text clipboard');
    }
    assert.deepEqual(errors,[]);
  }finally{if(application)await application.close();socket?.destroy();await new Promise(resolve=>server.close(resolve));}
}
(async()=>{await test('database');await test('mask');if(process.env.GRIDPILOT_TEST_APP)await test('settings');})().catch(e=>{console.error(e);process.exitCode=1;}).finally(()=>{fs.rmSync(temporary,{recursive:true,force:true,maxRetries:5,retryDelay:200});});
