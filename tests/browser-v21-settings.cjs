// Runs the production Electron main/preload/renderer against a synthetic private pipe.
const { _electron: electron } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const assert = require('assert');
const path = require('path');
const fs = require('fs');
const net = require('net');
const crypto = require('crypto');
const root = path.resolve(__dirname, '..');
async function test(mode) {
  let socket, lastRequest, application;
  const name = 'GridPilot-' + crypto.randomBytes(16).toString('hex');
  const settings = { IconTheme: 0, FocusRipple: true, RippleDuration: 280, RippleWidth: 36, RippleStrength: 24, FocusColor: '#42AD91', FocusOpacity: 20, FocusRow: true, FocusColumn: true, ClearActiveCell: true, FocusAnimation: false, FocusInterval: 33, MaskReplaceMentions: true, MaskPreserveFormatting: true, Copy: {}, Rename: {} };
  const rules = [0, 1].map(Sheet => ({ Sheet, SheetName: 'Sheet' + Sheet, Column: 1, Header: '客户名称', Domain: '客户', StartRow: 2, Kind: 1, Prefix: '', Suffix: '有限公司', Digits: 6, ShortName: false, Factor: 1.7, DayShift: 37 }));
  const send = data => socket.write(JSON.stringify(data) + '\n');
  const server = net.createServer(client => {
    socket = client; let buffer = '';
    client.on('data', data => {
      buffer += data.toString(); let end;
      while ((end = buffer.indexOf('\n')) >= 0) {
        const m = JSON.parse(buffer.slice(0, end)); buffer = buffer.slice(end + 1); lastRequest = m;
        if (m.action === 'init') send({ id: m.id, data: { mode, settings, rules, cells: 2500001, sheets: [0, 1].map(id => ({ id, name: 'Sheet' + id, nonEmptyRows: 50000, nonEmptyColumns: 50 })) } });
        else if (m.action === 'saveSettings') send({ id: m.id, data: { saved: true, message: '已保存', settings: m.settings } });
        else if (m.action === 'preview') {
          assert(m.rules.every(r => r.Prefix === '【测试】' && r.Suffix === '项目' && r.Digits === 3));
          send({ type: 'progress', phase: '验证公式关系', percent: 65, current: 1600000, total: 2500001 });
          setTimeout(() => send({ id: m.id, data: { summary: '完整扫描通过', rows: [{ location: 'S1!A2', before: '甲', after: '【测试】青禾123项目' }] } }), 200);
        } else send({ id: m.id, data: { message: '已取消' } });
      }
    });
  });
  await new Promise(resolve => server.listen('\\\\.\\pipe\\' + name, resolve));
  const env = { ...process.env }; delete env.ELECTRON_RUN_AS_NODE; delete env.NODE_OPTIONS;
  try {
    application = await electron.launch({ executablePath: path.join(root, '.temp/Browser/electron.exe'), args: [path.join(root, 'App'), '--gridpilot-pipe=' + name], env });
    const page = await application.firstWindow(); const errors = []; page.on('pageerror', e => errors.push(e.message));
    await page.getByRole('button', { name: mode === 'mask' ? '分析与预览' : '保存设置', exact: true }).waitFor();
    await page.getByRole('button',{name:'AI 助手',exact:true}).click();
    assert.equal(await page.getByLabel('快照预算 MB').inputValue(),'500');
    await page.getByLabel('快照预算 MB').fill('50');await page.getByRole('button',{name:'保存设置',exact:true}).click();
    await page.locator('#status').filter({hasText:'已保存'}).waitFor();assert.equal(lastRequest.settings.SnapshotBudgetMB,100);
    await page.getByLabel('快照预算 MB').fill('750');
    for(const label of ['解释公式','生成内容','分析数据','手动公式草稿'])await page.getByLabel(label,{exact:true}).uncheck();
    await page.getByRole('button',{name:'保存设置',exact:true}).click();await page.locator('#status').filter({hasText:'已保存'}).waitFor();
    assert.equal(lastRequest.settings.SnapshotBudgetMB,750);for(const key of ['AiShowExplain','AiShowGenerate','AiShowAnalyze','AiShowManualFormula'])assert.equal(lastRequest.settings[key],false);
    fs.mkdirSync(path.join(root,'docs/v2.1'),{recursive:true});await page.screenshot({path:path.join(root,'docs/v2.1/settings.png')});
    await page.setViewportSize({width:800,height:620});assert(await page.evaluate(()=>document.documentElement.scrollWidth<=window.innerWidth));assert.deepEqual(errors,[]);
  } finally {
    if (application) await application.close();
    socket?.destroy(); await new Promise(resolve => server.close(resolve));
  }
}
(async () => { await test('settings'); console.log('PASS v2.1 settings: budget defaults/minimum/custom value, four shortcuts and responsive layout'); })().catch(error => { console.error(error); process.exitCode = 1; });
