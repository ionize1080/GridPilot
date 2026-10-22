// Real packed XLL + native Chatbox: test-only CDP instrumentation in an isolated staging folder.
const fs=require('node:fs'),path=require('node:path'),assert=require('node:assert/strict'),{spawn}=require('node:child_process');
const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const root=path.resolve(__dirname,'..'),stage=fs.mkdtempSync(path.join(root,'.temp/switch-')),chat=path.join(stage,'Chatbox');fs.mkdirSync(chat);
fs.mkdirSync(path.join(stage,'ui-profile'));fs.writeFileSync(path.join(stage,'ui-profile/config.json'),JSON.stringify({settings:{language:'zh-Hans',autoUpdate:false,providers:{openai:{apiKey:'synthetic-test-only',apiHost:'http://127.0.0.1:9'}}},lastShownAboutDialogVersion:'1.20.3'}));
const packageRoot=process.env.GRIDPILOT_TEST_PACKAGE;
const runtime=packageRoot||path.join(root,'dist'),chatSource=packageRoot?path.join(packageRoot,'Chatbox'):path.join(root,'.temp/Chatbox');
const link=(target,dest)=>fs.symlinkSync(target,dest,'junction');
for(const folder of ['Browser','App','Resources'])link(path.join(runtime,folder),path.join(stage,folder));
for(const folder of ['native','node_modules','resources'])link(path.join(chatSource,folder),path.join(chat,folder));
for(const file of fs.readdirSync(chatSource).filter(n=>/\.cjs$|^package.json$/.test(n)))fs.copyFileSync(path.join(chatSource,file),path.join(chat,file));
const bootstrap=path.join(chat,'bootstrap.cjs');fs.writeFileSync(bootstrap,fs.readFileSync(bootstrap,'utf8').replace("const {app, ipcMain, BrowserWindow} = require('electron');","const {app, ipcMain, BrowserWindow} = require('electron'); app.commandLine.appendSwitch('remote-debugging-port','19331'); app.on('browser-window-created',(_e,w)=>{w.webContents.on('preload-error',(_e,p,e)=>require('node:fs').appendFileSync(require('node:path').join(__dirname,'preload-test.log'),p+' '+e.stack));});"));
const output=packageRoot||path.join(root,'src/SheetPilot.Excel/bin/Release/net462');for(const f of fs.readdirSync(output).filter(n=>n.endsWith('.dll')))fs.copyFileSync(path.join(output,f),path.join(stage,f));
fs.copyFileSync(path.join(output,packageRoot?'GridPilot-x64.xll':'publish/SheetPilot-AddIn64-packed.xll'),path.join(stage,'GridPilot-x64.xll'));
const child=spawn(path.join(root,'tests/SheetPilot.Tests/bin/Release/net462/SheetPilot.Tests.exe'),['--ai-ui',path.join(stage,'GridPilot-x64.xll')],{windowsHide:true});child.stdout.on('data',d=>process.stdout.write(d));child.stderr.on('data',d=>process.stderr.write(d));
const pause=ms=>new Promise(r=>setTimeout(r,ms));let browser,sequence=0;
async function state(action='state'){fs.writeFileSync(path.join(stage,'ui-command.json'),JSON.stringify({sequence:++sequence,action}));for(let i=0;i<100;i++){await pause(100);try{const s=JSON.parse(fs.readFileSync(path.join(stage,'ui-state.json'),'utf8'));if(s.sequence===sequence)return s;}catch{}}throw Error('Native state timeout');}
(async()=>{try{
  for(let i=0;i<100;i++){try{browser=await chromium.connectOverCDP('http://127.0.0.1:19331');break;}catch{await pause(300);}}assert(browser);
  const page=browser.contexts()[0].pages()[0];page.on('pageerror',e=>console.log('PAGE ERROR',e.message));await page.waitForFunction(()=>!!window.gridpilotExcel);await page.waitForTimeout(1000);if(await page.getByText('Skip guide',{exact:true}).isVisible())await page.getByText('Skip guide',{exact:true}).click();await page.getByText('Just chat',{exact:true}).click();await page.locator('.gp-window-mode').waitFor();await page.locator('#message-input').fill('保留中文草稿：侧栏与窗口切换。');
  assert.equal(await page.getByLabel('Excel上下文范围').inputValue(),'auto');
  const initial=await state();
  for(const decision of ['no','yes']){
    const helper=spawn(path.join(root,'tests/SheetPilot.Tests/bin/Release/net462/SheetPilot.Tests.exe'),['--ai-confirmation',String(initial.excelPid),decision],{windowsHide:true});helper.stdout.on('data',d=>process.stdout.write(d));helper.stderr.on('data',d=>process.stderr.write(d));
    const done=new Promise(resolve=>helper.once('exit',resolve));
    await page.getByRole('switch',{name:'自动批准'}).click();assert.equal(await done,0);
    await page.waitForFunction(()=>!document.querySelector('.gp-approval').disabled);
    assert.equal(await page.getByRole('switch',{name:'自动批准'}).getAttribute('aria-checked'),String(decision==='yes'));
  }
  await page.getByRole('switch',{name:'自动批准'}).click();await page.waitForFunction(()=>document.querySelector('.gp-approval').getAttribute('aria-checked')==='false');
  console.log('PASS native approval confirmation cancel / accept / disable and default auto mode');
  for(let i=0;i<8;i++){
    if(i===0||i===4){await page.locator('.controls.self-start button').nth(1).click();await pause(300);}
    if(i===2)await state('newBook');if(i===6)await state('originalBook');
    await page.locator('.gp-window-mode').click();const sidebar=i%2===0;
    await page.waitForFunction(sidebar=>document.documentElement.dataset.gpMode===(sidebar?'sidebar':'window'),sidebar);
    await page.waitForFunction(()=>!document.querySelector('.gp-window-mode').disabled);await pause(250);
    const actual=await state();assert(actual.present&&actual.visible,JSON.stringify(actual));assert.equal(actual.docked,sidebar,JSON.stringify(actual));if(!sidebar)assert(actual.unowned);
    assert.equal(await page.locator('#message-input').inputValue(),'保留中文草稿：侧栏与窗口切换。');console.log('PASS actual HWND '+(sidebar?'docked':'floating')+' transition '+(i+1)+' and retained draft');
  }
}catch(e){if(browser){const p=browser.contexts()[0].pages()[0];console.log('PRELOAD DIAGNOSTIC',fs.existsSync(path.join(chat,'preload-test.log'))?fs.readFileSync(path.join(chat,'preload-test.log'),'utf8'):'no preload error');console.log('PAGE',p.url(),await p.evaluate(()=>({native:!!window.electronAPI,excel:!!window.gridpilotExcel})));console.log('UI diagnostic',await p.locator('body').innerText());await p.screenshot({path:path.join(root,'.temp/switch-failure.png')});}throw e;}finally{if(browser)await browser.close();fs.writeFileSync(path.join(stage,'stop-ui-test'),'done');await new Promise(resolve=>{if(child.exitCode!==null)return resolve();child.once('exit',resolve);});}})().catch(e=>{console.error(e);process.exitCode=1;});
