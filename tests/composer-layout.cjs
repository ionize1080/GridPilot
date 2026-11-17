// Exercise the production native composer, not a simplified workspace-only fixture.
const assert=require('node:assert/strict'),fs=require('node:fs'),path=require('node:path');
module.exports=async function(page,application,root){
  const out=path.join(root,'.temp/composer-layout');fs.mkdirSync(out,{recursive:true});
  assert.equal(await page.locator('.disclaimer-safe-area').count(),0,'Removed footer must not reserve space');
  await page.locator('#message-input').fill('布局测试草稿：保留输入，不发送。');
  for(const [width,zoom] of [[320,1],[360,1],[440,1],[480,1],[660,1.5],[880,2],[1100,1]]){
    await application.evaluate(({BrowserWindow},arg)=>{const w=BrowserWindow.getAllWindows()[0];w.setSize(arg.width,920);w.webContents.setZoomFactor(arg.zoom);w.webContents.send('gridpilot:mode','sidebar');},{width,zoom});
    await page.waitForTimeout(250);
    const state=await page.evaluate(()=>{
      const rect=e=>{const r=e.getBoundingClientRect();return {x:r.x,y:r.y,right:r.right,bottom:r.bottom,width:r.width,height:r.height}};
      const slot=document.querySelector('#gridpilot-toolbar-slot'),excel=document.querySelector('#gridpilot-excel-tools'),native=document.querySelector('.gridpilot-native-toolbar');
      const workspace=slot.querySelector('gridpilot-workspace').shadowRoot.querySelector('#trigger');
      const controls=[...native.querySelectorAll('button')].filter(e=>e.getClientRects().length&&getComputedStyle(e).visibility!=='hidden').map(e=>({label:e.textContent||e.getAttribute('aria-label'),...rect(e)}));
      return {viewport:innerWidth,slot:rect(slot),excel:rect(excel),workspace:rect(workspace),native:rect(native),controls,overflow:document.documentElement.scrollWidth>innerWidth+1,draft:document.querySelector('#message-input').value};
    });
    fs.writeFileSync(path.join(out,`${width}-${zoom}.json`),JSON.stringify(state,null,2));
    await page.screenshot({path:path.join(out,`${width}-${zoom}.png`)});
    assert(!state.overflow,JSON.stringify(state));
    assert(state.slot.bottom<=state.native.y+1,'Native model controls must have their own row');
    assert(state.workspace.right+4<=state.excel.x||state.workspace.bottom<=state.excel.y,'Workspace and Excel controls must have a visible gap');
    assert(state.workspace.width>=70&&state.workspace.height>=30,'Readable workspace target');
    for(const c of state.controls){assert(c.x>=0&&c.right<=state.viewport+1,'Native control outside viewport: '+JSON.stringify(c));}
    for(let i=0;i<state.controls.length;i++)for(let j=i+1;j<state.controls.length;j++){
      const a=state.controls[i],b=state.controls[j];
      assert(Math.min(a.right,b.right)-Math.max(a.x,b.x)<=1||Math.min(a.bottom,b.bottom)-Math.max(a.y,b.y)<=1,'Overlapping native controls');
    }
    assert.equal(state.draft,'布局测试草稿：保留输入，不发送。');
    await page.getByLabel('Excel工具',{exact:true}).click();
    const popup=await page.locator('.gp-popup').boundingBox();
    assert(popup&&popup.x>=0&&popup.x+popup.width<=state.viewport+1,'Excel menu must remain inside sidebar');
    await page.getByLabel('Excel工具',{exact:true}).click();
    await page.locator('gridpilot-workspace').locator('#trigger').click();
    await page.locator('gridpilot-workspace').locator('#panel').waitFor({state:'visible'});
    const panel=await page.locator('gridpilot-workspace').locator('#panel').boundingBox();
    assert(panel&&panel.x>=0&&panel.x+panel.width<=state.viewport+1,'Workspace panel must fit');
    await page.locator('gridpilot-workspace').locator('#trigger').click();
    await page.locator('gridpilot-workspace').locator('#panel').waitFor({state:'hidden'});
    const before=await page.locator('#message-input').boundingBox();
    await page.locator('#message-input').fill('第一行\n第二行\n第三行\n第四行');
    const expanded=await page.locator('#message-input').boundingBox();
    assert(expanded.height>before.height,'Compact input must still grow for multiline drafts');
    await page.locator('#message-input').fill('布局测试草稿：保留输入，不发送。');
    console.log(`PASS native composer ${width}px / ${zoom*100}%: separate rows, no overlap, reachable menus, draft preserved`);
  }
};
