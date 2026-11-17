'use strict';
const fs = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');
const { spawn } = require('node:child_process');
const extensions = ["avif", "bat", "bmp", "c", "cfg", "cpp", "cs", "css", "csv", "doc", "docm", "docx", "dot", "dotm", "dotx", "eml", "epub", "fodp", "fods", "fodt", "gif", "h", "hpp", "htm", "html", "ico", "ics", "ini", "ipynb", "java", "jfif", "jpeg", "jpg", "js", "json", "jsonl", "jsx", "log", "markdown", "md", "mht", "mhtml", "ndjson", "odp", "ods", "odt", "otp", "ots", "ott", "pbm", "pdf", "pgm", "png", "pot", "potm", "potx", "ppm", "pps", "ppsm", "ppsx", "ppt", "pptm", "pptx", "ps1", "py", "r", "rst", "rtf", "scss", "sh", "sql", "srt", "svg", "tex", "tif", "tiff", "toml", "ts", "tsv", "tsx", "txt", "vcf", "vtt", "webp", "xhtml", "xls", "xlsb", "xlsm", "xlsx", "xlt", "xltm", "xltx", "xml", "yaml", "yml"];
const id = () => crypto.randomBytes(16).toString('hex');
function atomic(file, value) {
  fs.mkdirSync(path.dirname(file), { recursive: true });
  const temporary = file + '.' + id() + '.tmp';
  try { fs.writeFileSync(temporary, JSON.stringify(value), { encoding: 'utf8', mode: 0o600 }); fs.renameSync(temporary, file); }
  finally { if (fs.existsSync(temporary)) fs.unlinkSync(temporary); }
}
function read(file, fallback) { return fs.existsSync(file) ? JSON.parse(fs.readFileSync(file, 'utf8')) : fallback; }
function preferences(folder) {
  const { maxFileMB = 100, searchDepth = 3 } = read(path.join(folder, 'preferences.json'), {});
  if (!Number.isInteger(maxFileMB) || maxFileMB < 1 || maxFileMB > 1024) throw Error('文件大小上限须为 1–1024 MB 的整数');
  if (!Number.isInteger(searchDepth) || searchDepth < 1 || searchDepth > 12) throw Error('递归深度须为 1–12 层');
  return { maxFileMB, searchDepth };
}
function savePreferences(folder, value) {
  value = { ...preferences(folder), ...value };
  if (!Number.isInteger(value?.maxFileMB) || value.maxFileMB < 1 || value.maxFileMB > 1024) throw Error('文件大小上限须为 1–1024 MB 的整数');
  if (!Number.isInteger(value.searchDepth) || value.searchDepth < 1 || value.searchDepth > 12) throw Error('递归深度须为 1–12 层');
  atomic(path.join(folder, 'preferences.json'), { maxFileMB: value.maxFileMB, searchDepth: value.searchDepth });
  return preferences(folder);
}
class WorkspaceStore {
  constructor(folder, root) { this.folder = folder; this.root = root; fs.mkdirSync(folder, { recursive: true }); }
  get(session) {
    if (typeof session !== 'string' || !session.length || session.length > 160) throw Error('无效聊天');
    const indexFile = path.join(this.folder, 'sessions.json'), index = read(indexFile, {});
    const key = crypto.createHash('sha256').update(session).digest('hex');
    if (!index[key]) {
      index[key] = id();
      atomic(path.join(this.folder, index[key] + '.json'), { id: index[key], sources: [] });
      atomic(indexFile, index);
    }
    const data = read(path.join(this.folder, index[key] + '.json'), { id: index[key], sources: [] });
    data.collections ||= [];
    return data;
  }
  save(data) { atomic(path.join(this.folder, data.id + '.json'), data); }
  clear(session) {
    const previous = this.get(session);
    if (!previous.sources.length && !previous.collections.length && !this.outputCount(previous)) return { cleared: true, workspace_id: previous.id };
    // Revoke the old grant before publishing the new one. Keep every disk file,
    // including generated papers; archive the manifest for local traceability.
    atomic(path.join(this.folder, previous.id + '.retired.json'), previous);
    atomic(path.join(this.folder, previous.id + '.json'), { id: previous.id, revoked: true, sources: [], collections: [] });
    const next = { id: id(), sources: [], collections: [], undoWorkspaceId: previous.id };
    this.save(next);
    const indexFile = path.join(this.folder, 'sessions.json'), index = read(indexFile, {});
    index[crypto.createHash('sha256').update(session).digest('hex')] = next.id;
    atomic(indexFile, index);
    return { cleared: true, workspace_id: next.id };
  }
  undoClear(session) {
    const current = this.get(session), previousId = current.undoWorkspaceId;
    if (!/^[a-f0-9]{32}$/.test(previousId || '')) throw Error('没有可撤销的清空操作');
    const previous = read(path.join(this.folder, previousId + '.retired.json'), null);
    if (!previous || previous.id !== previousId) throw Error('清空记录不可用');
    // Retain references added since clearing as well as new working papers.
    previous.collections = [...(previous.collections || []), ...current.collections];
    previous.sources.push(...current.sources);
    const results = path.join(this.folder, 'results', current.id);
    if (fs.existsSync(results)) for (const name of fs.readdirSync(results)) {
      if (name.endsWith('.xlsx')) previous.sources.push({ id: id(), path: this.resultPath(session, 'results:' + Buffer.from(name).toString('base64url')), kind: 'file', parent: '', order: Date.now() });
    }
    delete previous.undoWorkspaceId;
    this.save(previous);
    const indexFile = path.join(this.folder, 'sessions.json'), index = read(indexFile, {});
    index[crypto.createHash('sha256').update(session).digest('hex')] = previous.id;
    atomic(indexFile, index);
    atomic(path.join(this.folder, current.id + '.json'), { id: current.id, revoked: true, sources: [], collections: [] });
    return { restored: true, workspace_id: previous.id };
  }
  outputCount(data) { const folder = path.join(this.folder, 'results', data.id); return fs.existsSync(folder) && fs.readdirSync(folder).length ? 1 : 0; }
  resultPath(session, identifier) {
    const data = this.get(session);
    if (typeof identifier !== 'string' || !identifier.startsWith('results:')) throw Error('只能管理工作区生成的底稿');
    const name = Buffer.from(identifier.slice(8), 'base64url').toString('utf8');
    if (!name || name !== path.basename(name) || /[\\/:]/.test(name) || !name.endsWith('.xlsx')) throw Error('无效底稿');
    const folder = path.resolve(this.folder, 'results', data.id), target = path.resolve(folder, name);
    if (path.dirname(target) !== folder || fs.lstatSync(target).isSymbolicLink() || fs.realpathSync.native(folder).toLowerCase() !== folder.toLowerCase()) throw Error('底稿路径发生变化');
    return target;
  }
  adopt(from, to, workspaceId) {
    if (from === to) return;
    const next = this.get(to);
    if (workspaceId && next.id === workspaceId) return;
    const previous = this.get(from);
    if (workspaceId && previous.id !== workspaceId) throw Error('参考工作区已变化，请重新发送消息');
    if (next.sources.length || next.collections.length || this.outputCount(next)) throw Error('目标聊天已有工作区资料，无法继承新聊天的资料');
    // Keep even an empty captured workspace: MCP may create working papers in it.
    const indexFile = path.join(this.folder, 'sessions.json'), index = read(indexFile, {});
    const hash = x => crypto.createHash('sha256').update(x).digest('hex');
    index[hash(to)] = previous.id;
    delete index[hash(from)];
    atomic(indexFile, index);
  }
  parent(data, value) {
    const identifier = value?.startsWith('virtual:') ? value.slice(8) : value || '';
    if (identifier && !data.collections.some(c => c.id === identifier)) throw Error('虚拟文件夹已变化，请刷新');
    return identifier;
  }
  collection(session, action, args) {
    if (!['create','rename','delete'].includes(action)) throw Error('无效文件夹操作');
    const data = this.get(session), parent = this.parent(data, args.parent), identifier = String(args.id || '').replace(/^virtual:/, '');
    const current = data.collections.find(c => c.id === identifier);
    if (action !== 'create' && !current) throw Error('虚拟文件夹不存在');
    const name = String(args.name || '').trim();
    if (action !== 'delete' && (!name || name.length > 80 || /[\\/\x00-\x1f]/.test(name))) throw Error('文件夹名称须为 1–80 个字符，不能包含路径分隔符');
    if (action !== 'delete' && data.collections.some(c => c.id !== identifier && (c.parent || '') === (current ? current.parent || '' : parent) && c.name.toLowerCase() === name.toLowerCase())) throw Error('此位置已有同名文件夹');
    if (action === 'create') {
      if (data.collections.length >= 200) throw Error('最多创建 200 个虚拟文件夹');
      let depth = 1, ancestor = parent;
      while (ancestor) { depth++; ancestor = data.collections.find(c => c.id === ancestor)?.parent || ''; }
      if (depth > 12) throw Error('虚拟文件夹最多嵌套 12 层');
      data.collections.push({ id: id(), name, parent, order: Date.now() });
    } else if (action === 'rename') current.name = name;
    else {
      const removed = new Set([identifier]);
      for (let i = 0; i < 12; i++) for (const c of data.collections) if (removed.has(c.parent)) removed.add(c.id);
      data.collections = data.collections.filter(c => !removed.has(c.id));
      data.sources = data.sources.filter(s => !removed.has(s.parent));
    }
    this.save(data); return { saved: true };
  }
  move(session, identifiers, target, before = '', after = '') {
    const data = this.get(session), parent = this.parent(data, target);
    if (!Array.isArray(identifiers) || !identifiers.length || identifiers.length > 100) throw Error('每次可整理 1–100 个条目');
    const selected = [...new Set(identifiers)];
    if (selected.some(s => typeof s !== 'string')) throw Error('无效条目');
    if ((before && selected.includes(before)) || (after && selected.includes(after))) return { moved: 0 };
    const entries = [];
    for (const identifier of selected) {
      if (identifier.startsWith('virtual:')) {
        const c = data.collections.find(c => 'virtual:' + c.id === identifier);
        if (!c) throw Error('文件夹已被移除');
        let ancestor = parent, depth = 1;
        while (ancestor) { if (ancestor === c.id) throw Error('不能将文件夹放入自身或其子文件夹'); depth++; ancestor = data.collections.find(x => x.id === ancestor)?.parent || ''; }
        const subtreeDepth = root => 1 + Math.max(0, ...data.collections.filter(x => x.parent === root).map(x => subtreeDepth(x.id)));
        if (depth + subtreeDepth(c.id) - 1 > 12) throw Error('虚拟文件夹最多嵌套 12 层');
        entries.push(c);
      } else {
        const [sourceId, encoded = ''] = identifier.split(':');
        let source = data.sources.find(s => s.id === sourceId);
        if (encoded || sourceId === 'results') {
          let full;
          if (sourceId === 'results') full = this.resultPath(session, identifier);
          else {
            if (!source || source.kind !== 'folder') throw Error('无效文件引用');
            const rel = Buffer.from(encoded, 'base64url').toString('utf8');
            if (/[\\:\x00]/.test(rel) || rel.startsWith('/') || rel.split('/').some(p => p === '..' || p.startsWith('.'))) throw Error('无效文件引用');
            full = path.resolve(source.path, rel); const relative = path.relative(source.path, fs.realpathSync.native(full));
            if (relative.startsWith('..') || path.isAbsolute(relative)) throw Error('文件已移出授权文件夹');
            for (let check = full; check !== path.dirname(check); check = path.dirname(check)) if (fs.lstatSync(check).isSymbolicLink()) throw Error('不支持链接');
          }
          const stat = fs.statSync(full);
          source = data.sources.find(s => s.path.toLowerCase() === full.toLowerCase() && (s.parent || '') === parent);
          if (!source) {
            source = { id: id(), path: full, kind: stat.isDirectory() ? 'folder' : 'file', parent, order: Date.now() };
            if (data.sources.length >= 500) throw Error('最多添加 500 个来源');
            data.sources.push(source); // A virtual reference; the disk item is not moved.
          }
        }
        if (!source) throw Error('文件引用已被移除');
        entries.push(source);
      }
    }
    // If both a folder and a descendant are selected, retain the subtree intact.
    const groupIds = new Set(entries.filter(e => 'name' in e).map(e => e.id));
    const moving = entries.filter(e => { let p = e.parent; while (p) { if (groupIds.has(p)) return false; p = data.collections.find(c => c.id === p)?.parent; } return true; });
    for (const e of moving) e.parent = parent;
    const siblings = [...data.collections, ...data.sources].filter(e => (e.parent || '') === parent && !moving.includes(e)).sort((a, b) => (a.order || 0) - (b.order || 0));
    const names = new Set(siblings.filter(e => 'name' in e).map(e => e.name.toLowerCase()));
    for (const e of moving.filter(e => 'name' in e)) { const name = e.name.toLowerCase(); if (names.has(name)) throw Error('目标位置已有同名文件夹'); names.add(name); }
    const publicId = e => 'name' in e ? 'virtual:' + e.id : e.id + ':';
    const anchor = before || after;
    let position = anchor ? siblings.findIndex(e => publicId(e) === anchor) : siblings.length;
    if (after && position >= 0) position++;
    if (position < 0) position = siblings.length;
    siblings.splice(position, 0, ...moving); siblings.forEach((e, i) => e.order = i + 1);
    this.save(data); return { moved: moving.length };
  }
  add(session, paths, destination = '') {
    if (!Array.isArray(paths) || paths.length > 100) throw Error('每次最多添加 100 个文件');
    const data = this.get(session), errors = [], added = [];
    const parent = this.parent(data, destination);
    for (const input of paths) {
      try {
        if (typeof input !== 'string' || !path.isAbsolute(input) || input.startsWith('\\\\')) throw Error('请选择本机文件');
        let check = path.resolve(input);
        while (true) {
          if (fs.lstatSync(check).isSymbolicLink()) throw Error('请直接选择原文件，不支持符号链接');
          const parent = path.dirname(check); if (parent === check) break; check = parent;
        }
        const full = fs.realpathSync.native(input), stat = fs.statSync(full);
        if (!stat.isDirectory() && !stat.isFile()) throw Error('不是普通文件');
        if (stat.isFile() && !extensions.includes(path.extname(full).slice(1).toLowerCase())) throw Error('不支持此文件格式');
        if (stat.isFile() && stat.size > preferences(this.folder).maxFileMB * 1024 * 1024) throw Error('超过文件大小上限，可在首选项修改');
        let source = data.sources.find(s => s.path.toLowerCase() === full.toLowerCase() && (s.parent || '') === parent);
        if (!source) {
          if (data.sources.length >= 500) throw Error('工作区最多添加 500 个来源，请使用文件夹');
          source = { id: id(), path: full, kind: stat.isDirectory() ? 'folder' : 'file', parent, order: Date.now() };
          data.sources.push(source);
        }
        added.push({ id: source.id + ':', name: path.basename(full), kind: source.kind });
      } catch (error) { errors.push({ name: path.basename(String(input)), message: error.message }); }
    }
    this.save(data);
    return { workspace_id: data.id, added, errors };
  }
  remove(session, identifier) {
    const data = this.get(session);
    data.sources = data.sources.filter(s => s.id + ':' !== identifier);
    this.save(data);
    return { removed: true };
  }
  worker(operation, args) {
    return new Promise((resolve, reject) => {
      const child = spawn(path.join(this.root, 'Python/python.exe'), [path.join(this.root, 'bridge/workspace_files.py')], {
        windowsHide: true, env: { ...process.env, GRIDPILOT_WORKSPACES: this.folder, PYTHONIOENCODING: 'utf-8' }, stdio: ['pipe', 'pipe', 'ignore'],
      });
      const chunks = []; let bytes = 0, settled = false;
      const done = (error, value) => { if (settled) return; settled = true; clearTimeout(timer); error ? reject(error) : resolve(value); };
      const timer = setTimeout(() => { child.kill(); done(Error('文件处理超时，请缩小读取范围')); }, 45000);
      child.stdout.on('data', chunk => { bytes += chunk.length; if (bytes > 8 * 1024 * 1024) { child.kill(); done(Error('结果过大')); } else chunks.push(chunk); });
      child.on('error', error => done(error));
      child.on('close', () => { try { const data = JSON.parse(Buffer.concat(chunks).toString('utf8')); done(data.error ? Error(data.error) : null, data.result); } catch { done(Error('文件处理组件未就绪，请重新解压完整发布包')); } });
      child.stdin.on('error', () => {});
      child.stdin.end(JSON.stringify({ operation, arguments: args }));
    });
  }
  async summary(session, directory_id = '', cursor = 0) {
    const data = this.get(session);
    const result = await this.worker('list', { workspace_id: data.id, directory_id, cursor, limit: 40 });
    return { ...result, collections: data.collections, sourceCount: data.sources.length + this.outputCount(data), canUndoClear: !!data.undoWorkspaceId };
  }
}
function install({ ipcMain, dialog, primary, folder, root, openWorkbook }) {
  const store = new WorkspaceStore(folder, root);
  ipcMain.handle('gridpilot:workspace', async (event, action, args = {}) => {
    const win = primary();
    if (!win || event.sender !== win.webContents || event.senderFrame !== win.webContents.mainFrame || !event.senderFrame.url.startsWith('file:')) throw Error('Invalid sender');
    const session = args.session;
    if (action === 'preferences') return args.value ? savePreferences(folder, args.value) : preferences(folder);
    if (action === 'adopt') { store.adopt(args.from, session, args.workspace_id); return {}; }
    if (action === 'summary') return store.summary(session, args.directory_id, args.cursor);
    if (action === 'context') { const data = store.get(session); return { workspace_id: data.id, sourceCount: data.sources.length + store.outputCount(data) }; }
    if (action === 'clear') return store.clear(session);
    if (action === 'undoClear') return store.undoClear(session);
    if (action === 'open') {
      const source = store.resultPath(session, args.id);
      if (!openWorkbook) throw Error('Excel 连接不可用，请从 GridPilot 打开助手');
      return openWorkbook(source);
    }
    if (action === 'export') {
      const source = store.resultPath(session, args.id);
      const selected = await dialog.showSaveDialog(win, { title: '导出分析底稿', defaultPath: path.basename(source), filters: [{ name: 'Excel 工作簿', extensions: ['xlsx'] }] });
      if (selected.canceled) return { cancelled: true };
      await fs.promises.copyFile(source, selected.filePath);
      return { exported: true };
    }
    if (action === 'delete') {
      const target = store.resultPath(session, args.id);
      const answer = await dialog.showMessageBox(win, { type: 'question', title: '删除分析底稿', message: `删除“${path.basename(target)}”？`, detail: '原始参考文件不会删除。需要留存时，可先导出底稿。', buttons: ['保留', '删除底稿'], defaultId: 0, cancelId: 0, noLink: true });
      if (answer.response !== 1) return { cancelled: true };
      const checked = store.resultPath(session, args.id);
      await fs.promises.unlink(checked);
      return { deleted: true };
    }
    if (action === 'remove') return store.remove(session, args.id);
    if (action === 'collection') {
      if (args.operation === 'delete') {
        const response = await dialog.showMessageBox(win, { type: 'question', title: '删除虚拟文件夹', message: '移除此虚拟文件夹及其中的参考引用？', detail: '本机原文件和生成的分析底稿都会保留。', buttons: ['保留', '移除文件夹'], defaultId: 0, cancelId: 0, noLink: true });
        if (response.response !== 1) return { cancelled: true };
      }
      return store.collection(session, args.operation, args);
    }
    if (action === 'move') return store.move(session, args.ids, args.parent, args.before, args.after);
    if (action === 'choose') {
      const selection = await dialog.showOpenDialog(win, { title: args.folder ? '选择工作区文件夹' : '添加参考文件',
        properties: args.folder ? ['openDirectory'] : ['openFile', 'multiSelections'],
        ...(args.folder ? {} : { filters: [{ name: '文档与图片', extensions }] }) });
      return selection.canceled ? { cancelled: true } : store.add(session, selection.filePaths, args.parent);
    }
    if (action === 'add') return store.add(session, args.paths, args.parent);
    throw Error('Invalid action');
  });
  return store;
}
module.exports = { install, WorkspaceStore, preferences, savePreferences, extensions };
