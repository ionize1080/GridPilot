// Production Electron + named-pipe UI contract. Core/COM order is independently
// verified by --multi-sort --excel; this suite checks controls and pagination.
const { _electron: electron }=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const assert=require('node:assert/strict'),fs=require('node:fs'),path=require('node:path'),net=require('node:net'),crypto=require('node:crypto');
const root=path.resolve(__dirname,'..'),pipe='GridPilot-'+crypto.randomBytes(16).toString('hex');
let socket,last,application,previewCalls=0,applied=0;
const initial={mode:'sort',kind:'rows',description:'合成测试 · A1:E20001 · 20,000 条记录',headers:['所属部门','客户名称','项目名称','金额','日期'],rowCount:20001,isTable:false,merges:0,sequences:[{name:'Excel · 一月、二月、三月…',values:['一月','二月','三月']}],options:{HasHeader:true,MatchCase:false,StrokeOrder:false,Levels:[{Column:0,Mode:'values',Descending:false,Rules:'[中文数字]',Sequences:''}]}};
function respond(m){last=m;let data;
  if(['init','sortCapture','sortSample'].includes(m.action))data=initial;
  else if(m.action==='sortPreview'||m.action==='sortPage'){
    if(m.action==='sortPreview')previewCalls++;const count=m.search?1:20000,offset=Math.min(m.offset||0,count-1),size=m.size||100;
    data={total:20000,records:20000,previewed:20000,filtered:count,offset,size,headers:initial.headers,rows:Array.from({length:Math.min(size,count-offset)},(_,j)=>({position:offset+j+1,sourceRow:20001-offset-j,values:['国际一部','青禾','项目'+(offset+j+1),'123.45','2026-10-16']}))};
  }else if(m.action==='sortApply'){applied++;data={message:'已排序 20,000 行',canUndo:true};}
  else data={message:'已恢复原序'};
  socket.write(JSON.stringify({id:m.id,data})+'\n');
}
(async()=>{
  const server=net.createServer(s=>{socket=s;let buffer='';s.on('data',chunk=>{buffer+=chunk;let end;while((end=buffer.indexOf('\n'))>=0){const m=JSON.parse(buffer.slice(0,end));buffer=buffer.slice(end+1);respond(m);}});});
  await new Promise(resolve=>server.listen('\\\\.\\pipe\\'+pipe,resolve));
  const env={...process.env};delete env.NODE_OPTIONS;delete env.ELECTRON_RUN_AS_NODE;
  try{
    application=await electron.launch({executablePath:path.join(root,'.temp/Browser/electron.exe'),args:[path.join(root,'App'),'--gridpilot-pipe='+pipe],env});
    const page=await application.firstWindow(),errors=[];page.on('pageerror',e=>errors.push(e.message));
    await page.locator('#sortBuild').waitFor();assert(await page.locator('#sortApply').isDisabled());
    await page.getByRole('button',{name:'使用部门 → 客户 → 项目示例',exact:true}).click();assert.equal(await page.locator('.sort-level').count(),3);
    await page.locator('#sortBuild').click();await page.locator('#sortApply:enabled').waitFor();assert.equal(last.options.Levels.length,3);assert.equal(last.options.Levels[0].Descending,false);assert.equal(last.options.Levels[2].Descending,true);
    assert.equal(await page.locator('.sort-table-wrap tbody tr').count(),100);
    await page.getByLabel('跳转到预览页').fill('200');await page.getByLabel('跳转到预览页').press('Tab');await page.getByText('19901–20000 / 20,000 行',{exact:true}).waitFor();assert.equal(await page.locator('.sort-table-wrap tbody tr').last().locator('td').first().textContent(),'20000');assert.equal(previewCalls,1);
    await page.getByLabel('搜索排序预览').fill('青禾');await page.getByRole('button',{name:'搜索',exact:true}).click();await page.getByText('1–1 / 1 行',{exact:true}).waitFor();
    await page.locator('#sortApply').click();await page.locator('#sortUndo:enabled').waitFor();assert.equal(applied,1);assert(await page.locator('#sortApply').isDisabled());
    await page.locator('#sortUndo').click();await page.locator('#sortUndo:disabled').waitFor();
    await page.getByRole('button',{name:'重新读取选区',exact:true}).click();await page.getByText('1 / 20 级',{exact:true}).waitFor();
    for(let i=1;i<20;i++)await page.locator('#sortAdd').click();assert.equal(await page.locator('.sort-level').count(),20);assert(await page.locator('#sortAdd').isDisabled());
    await page.locator('#sortBuild').click();await page.locator('#sortApply:enabled').waitFor();assert.equal(last.options.Levels.length,20);
    await page.getByRole('button',{name:'使用部门 → 客户 → 项目示例',exact:true}).click();assert(await page.locator('#sortApply').isDisabled());
    await page.getByLabel('第 2 级降序',{exact:true}).click();await page.locator('#sortBuild').click();await page.locator('#sortApply:enabled').waitFor();assert(last.options.Levels[1].Descending);
    await page.getByLabel('第 2 级升序',{exact:true}).click();await page.locator('#sortBuild').click();await page.locator('#sortApply:enabled').waitFor();
    fs.mkdirSync(path.join(root,'docs/v1.6'),{recursive:true});await page.screenshot({path:path.join(root,'docs/v1.6/sort.png')});
    await application.evaluate(({BrowserWindow})=>BrowserWindow.getAllWindows()[0].setSize(780,650));
    assert(await page.locator('#sortApply').isVisible());await page.screenshot({path:path.join(root,'docs/v1.6/sort-narrow.png')});
    assert.deepEqual(errors,[]);console.log('PASS production sorting UI: 20 levels, directions, 20k final page, search, invalidation, apply, undo, narrow layout');
  }finally{if(application)await application.close();socket?.destroy();server.close();}
})().catch(e=>{console.error(e);process.exitCode=1;});
