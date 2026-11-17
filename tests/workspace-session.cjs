// Run the shipped new-chat submit handler without a mounted workspace component.
// React replaces the home input before the router's onResolved event can adopt it.
const fs = require('node:fs'), path = require('node:path'), os = require('node:os');
const vm = require('node:vm'), assert = require('node:assert/strict');
const { WorkspaceStore } = require('../integrations/chatbox/workspace-main.cjs');
const root = path.resolve(__dirname, '..');
const home = fs.mkdtempSync(path.join(os.tmpdir(), 'gridpilot-session-'));
const store = new WorkspaceStore(path.join(home, 'Workspaces'), path.join(root, 'dist'));
const file = path.join(home, 'reference.txt');
fs.writeFileSync(file, 'Workspace session regression');
const source = fs.readFileSync(path.join(root, 'dist/Chatbox/native/dist/renderer/js/index.BVHwD0Ys.js'), 'utf8');
const start = source.indexOf('async({constructedMessage:');
const end = source.indexOf(',[a,x,w.knowledgeBase,A,o,p,_])', start);
assert(start >= 0 && end > start, 'Pinned Chatbox new-chat submit handler exists');
const handler = source.slice(start, end);
let sequence = 0;
async function submit(workspace) {
  const sessionId = 'created-' + ++sequence, events = [];
  const message = { id: 'message-' + sequence, role: 'user', contentParts: [{ type: 'text', text: 'Read my reference' }], gridpilotWorkspace: { workspace_id: workspace.id } };
  const scope = {
    a: { name: 'Workspace regression', messages: [], settings: {} }, w: {}, o: {},
    Ie: async () => ({ id: sessionId }),
    window: { gridpilotFiles: { adopt: async (...args) => {
      // Exercise an asynchronous IPC handoff, not a synchronous mock.
      await new Promise(resolve => setTimeout(resolve, 5));
      store.adopt(...args); events.push('bound');
    } } },
    ze: id => {
      assert.equal(store.get(id).id, workspace.id, 'Workspace must be bound before the new chat mounts');
      events.push('navigated');
    },
    Ne: (id, payload) => {
      assert.equal(id, sessionId);
      assert.equal(store.get(id).id, payload.newUserMsg.gridpilotWorkspace.workspace_id, 'UI and MCP use the same workspace');
      events.push('submitted');
    },
  };
  await vm.runInNewContext('(' + handler + ')', scope)({ constructedMessage: message });
  assert.deepEqual(events, ['bound', 'navigated', 'submitted']);
  return sessionId;
}
(async () => {
  store.add('new', [file]);
  store.collection('new', 'create', { name: 'Reference folder' });
  const draft = store.get('new'), sessionId = await submit(draft);
  const reopened = new WorkspaceStore(store.folder, store.root);
  assert.deepEqual(reopened.get(sessionId), draft, 'Files and virtual folders survive reopening');
  assert.notEqual(store.get('new').id, draft.id, 'The next new chat gets its own workspace');
  store.add('new', [file]);
  const nextDraft = store.get('new');
  store.adopt('new', sessionId, draft.id);
  assert.deepEqual(store.get('new'), nextDraft, 'Repeating a completed handoff must not consume the next draft');
  assert.throws(() => store.adopt('new', 'stale-send', draft.id), /工作区已变化/);
  assert.deepEqual(store.get('new'), nextDraft);
  store.add('existing', [file]);
  const existing = store.get('existing');
  assert.throws(() => store.adopt('new', 'existing', nextDraft.id), /已有/);
  assert.deepEqual(store.get('existing'), existing, 'An existing chat must not be overwritten');
  await submit(nextDraft);
  const empty = store.get('new'), emptySession = await submit(empty);
  const results = path.join(store.folder, 'results', empty.id);
  fs.mkdirSync(results, { recursive: true });
  fs.writeFileSync(path.join(results, 'working-paper.xlsx'), 'test result');
  assert.equal(store.outputCount(store.get(emptySession)), 1, 'First-prompt working papers remain in the visible workspace');
  assert.equal(fs.readFileSync(file, 'utf8'), 'Workspace session regression');
  console.log('PASS first-send workspace binding before remount, UI/MCP identity, reopen, draft isolation, retry, stale-send rejection and empty-workspace results');
})().catch(error => { console.error(error); process.exitCode = 1; }).finally(() => {
  assert.equal(path.dirname(path.resolve(home)), path.resolve(os.tmpdir()));
  assert(path.basename(home).startsWith('gridpilot-session-'));
  fs.rmSync(home, { recursive: true, force: true });
});
