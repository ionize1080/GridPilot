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
    if (mode === 'settings') {
      await page.getByRole('button',{name:'AI 助手',exact:true}).click();
      assert.equal(await page.getByLabel('AI 助手工作方式').inputValue(),'0');
      await page.getByLabel('AI 助手工作方式').selectOption('1');
      await page.getByRole('button',{name:'保存设置',exact:true}).click();
      await page.locator('#status').filter({hasText:'已保存'}).waitFor();
      assert.equal(lastRequest.settings.AiDisplayMode,1);
      fs.mkdirSync(path.join(root,'docs/v1.7'),{recursive:true});
      await page.screenshot({path:path.join(root,'docs/v1.7/gridpilot-settings.png')});
      console.log('PASS AI sidebar default and standalone preference save');
      await page.getByRole('button',{name:'功能开关',exact:true}).click();
      assert.equal(await page.locator('[id^="feature-"]').count(),12);
      const before=await page.locator('[id^="feature-"]').evaluateAll(xs=>xs.map(x=>x.id));
      await page.locator('#feature-Sort').check();await page.locator('#feature-Mask').check();await page.locator('#sort-context-menu').uncheck();
      await page.getByRole('button',{name:'保存设置',exact:true}).click();await page.locator('#status').filter({hasText:'已保存'}).waitFor();
      assert.deepEqual(lastRequest.settings.EnabledFeatures,['GridPilot_Sort','GridPilot_Mask']);assert.equal(lastRequest.settings.SortContextMenu,false);
      await page.locator('#feature-Mask').uncheck();await page.getByRole('button',{name:'保存设置',exact:true}).click();await page.locator('#status').filter({hasText:'已保存'}).waitFor();
      assert.deepEqual(lastRequest.settings.EnabledFeatures,['GridPilot_Sort']);assert.deepEqual(await page.locator('[id^="feature-"]').evaluateAll(xs=>xs.map(x=>x.id)),before);
      fs.mkdirSync(path.join(root,'docs/v1.5'),{recursive:true});await page.locator('#content').evaluate(e=>e.scrollTop=0);await page.screenshot({path:path.join(root,'docs/v1.5/settings.png')});
      console.log('PASS optional feature switches, disable/re-enable persistence, fixed order and independent context menu');
      await page.getByRole('button', { name: '聚焦与涟漪', exact: true }).click();
      for (const value of ['320', '460']) {
        await page.locator('#setting-RippleDuration').fill(value);
        await page.getByRole('button', { name: '保存设置', exact: true }).click();
        await page.locator('#status').filter({ hasText: '已保存' }).waitFor();
        assert.equal(lastRequest.settings.RippleDuration, Number(value));
        assert.equal(await page.locator('#setting-RippleDuration').inputValue(), value);
      }
      console.log('PASS F07 repeated settings saves retain latest edits');
    } else {
      await page.getByRole('textbox', { name: '客户名称 前缀', exact: true }).first().fill('【测试】');
      await page.getByRole('textbox', { name: '客户名称 后缀', exact: true }).first().fill('项目');
      await page.getByRole('spinbutton', { name: '客户名称 随机位数', exact: true }).first().fill('3');
      await page.getByRole('button', { name: '分析与预览', exact: true }).click();
      await page.locator('#progressMeta').filter({ hasText: '65%' }).waitFor();
      await page.locator('dialog[open]').waitFor();
      assert.equal(await page.getByRole('textbox', { name: '客户名称 后缀', exact: true }).last().inputValue(), '项目');
      assert.equal(await page.locator('#previewRows tr').count(), 1);
      await page.locator('#confirmed').check();
      await application.evaluate(({ dialog }) => { dialog.showSaveDialog = async () => ({ canceled: true }); });
      await page.locator('#export').click();
      await page.locator('#exportProgress').filter({ hasText: '已取消保存' }).waitFor();
      assert(await page.locator('dialog').evaluate(d => d.open)); assert(!(await page.locator('#export').isDisabled()));
      console.log('PASS T03 domain affix synchronization, progress frames and canceled save retain preview');
    }
    fs.mkdirSync(path.join(root, 'docs/v1.0'), { recursive: true });
    if(mode==='mask')await page.locator('#dismiss').click();
    if(mode==='mask'){await page.locator('[data-parameter=ShortName]').first().check();assert(await page.locator('[data-parameter=ShortName]').nth(1).isChecked());await page.locator('[data-parameter=ShortName]').first().uncheck();assert(!await page.locator('[data-parameter=ShortName]').nth(1).isChecked());}
    await page.screenshot({ path: path.join(root, 'docs/v1.0', mode + '.png') });
    await page.setViewportSize({ width: 800, height: 620 });
    assert(await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth));
    assert.deepEqual(errors, []);
  } finally {
    if (application) await application.close();
    socket?.destroy(); await new Promise(resolve => server.close(resolve));
  }
}
(async () => { await test('settings'); await test('mask'); console.log('PASS v1 production browser integration'); })().catch(error => { console.error(error); process.exitCode = 1; });
