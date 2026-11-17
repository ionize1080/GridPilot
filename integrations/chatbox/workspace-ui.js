// Native Web Component + Popover API: isolated styling, no extra framework/runtime.
(() => {
  const api = window.gridpilotFiles;
  const session = () => window.gridpilotWorkspaceSession?.() || 'new';
  const icons = {
    folder: '<path d="M3 7a2 2 0 0 1 2-2h5l2 2h7a2 2 0 0 1 2 2v10a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/>',
    plus: '<path d="M12 5v14M5 12h14"/>', close: '<path d="m6 6 12 12M18 6 6 18"/>',
    export: '<path d="M12 3v12m-4-4 4 4 4-4M5 16v5h14v-5"/>', trash: '<path d="M4 7h16M9 7V4h6v3M6 7l1 14h10l1-14M10 11v6M14 11v6"/>',
    edit: '<path d="m15 4 5 5-11 11H4v-5zM12 7l5 5"/>', move: '<path d="M3 7h6l2 2h10v11H3zM12 3v12m-3-3 3 3 3-3"/>',
    file: '<path d="M14 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9zM14 3v6h6M8 13h8M8 17h5"/>',
    image: '<rect x="3" y="3" width="18" height="18" rx="3"/><circle cx="8" cy="8" r="1.5"/><path d="m3 17 6-6 4 4 3-3 5 5"/>',
    back: '<path d="m14 6-6 6 6 6"/>', refresh: '<path d="M20 7v5h-5M4 17v-5h5M5.5 8A7 7 0 0 1 18 6l2 6M4 12l2 6a7 7 0 0 0 12.5-2"/>',
    settings: '<path d="M12 8a4 4 0 1 0 0 8 4 4 0 0 0 0-8M9 3h6l1 3 3 1 2 5-2 5-3 1-1 3H9l-1-3-3-1-2-5 2-5 3-1z"/>',
  };
  const icon = name => `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${icons[name]}</svg>`;
  const css = `
    :host{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","Microsoft YaHei UI",sans-serif;font-size:13px;color:#20232a;--surface:rgba(250,251,253,.97);--line:#22304816;--muted:#687184;--hover:#1c65db0c;--blue:#1764db;--field:#fff;display:inline-flex;align-items:center;-webkit-app-region:no-drag;color-scheme:light}
    *{box-sizing:border-box}button,input{font:inherit}button{color:inherit;cursor:pointer;border:0;background:transparent;display:inline-flex;align-items:center;justify-content:center;gap:6px;border-radius:9px;min-height:32px;padding:6px 9px}button:hover{background:var(--hover)}button:focus-visible,input:focus-visible{outline:3px solid #3989ef88;outline-offset:2px}button:disabled{opacity:.48;cursor:wait}svg{height:18px;width:18px;flex:none}#trigger{font-size:12px;gap:5px;padding:5px 8px;color:var(--muted);border:1px solid var(--line);border-radius:10px;height:30px;min-height:30px;white-space:nowrap}#trigger[aria-expanded=true]{background:#1764db12;color:var(--blue)}#badge:empty{display:none}#badge{font-size:10px;border-radius:5px;background:#1764db12;padding:1px 4px;color:var(--blue);font-variant-numeric:tabular-nums}
    #panel{inset:auto;margin:0;position:fixed;width:min(372px,calc(100vw - 20px));max-height:min(530px,calc(100vh - 72px));padding:0;border:1px solid var(--line);border-radius:19px;background:var(--surface);color:inherit;box-shadow:0 22px 70px #18253b28,0 3px 12px #18253b0b;backdrop-filter:blur(24px) saturate(140%);overflow:hidden;opacity:1}#panel:popover-open{display:flex;flex-direction:column}#panel::backdrop{background:transparent}header{display:flex;align-items:center;padding:17px 15px 10px;gap:10px}.brand{width:36px;height:36px;display:grid;place-items:center;background:linear-gradient(145deg,#edf5ff,#dceaff);color:var(--blue);border-radius:11px}.heading{flex:1;min-width:0}h2{font-size:15px;line-height:1.3;font-weight:650;letter-spacing:-.3px;margin:0}.subtitle{color:var(--muted);font-size:11px;margin-top:4px}.icon{width:32px;padding:7px;flex:none}nav{display:flex;gap:6px;padding:0 16px 12px}nav button{border:1px solid var(--line);background:var(--field);flex:1;font-size:12px}.primary{color:var(--blue)}#body{margin:0 10px;border-radius:12px;overflow:auto;min-height:140px;max-height:285px;overscroll-behavior:contain;scrollbar-width:thin}#empty{text-align:center;padding:25px 12px 28px;border:1px dashed var(--line);border-radius:12px;margin:2px 4px 8px}#empty .large{width:42px;height:42px;border-radius:13px;background:#1764db09;color:#729acc;margin:0 auto 12px;display:grid;place-items:center}#empty b{font-size:13px;font-weight:550}p{color:var(--muted);font-size:11px;line-height:1.7;margin:7px 0 0}#crumb{display:flex;align-items:center;gap:4px;padding:0 4px 6px;color:var(--muted);font-size:11px;min-width:0}#crumb span{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.row{display:flex;align-items:center;gap:9px;padding:8px 6px;border-radius:10px;min-width:0}.row:hover{background:var(--hover)}.fileicon{height:33px;width:29px;display:grid;place-items:center;border-radius:7px;background:#1764db09;color:#6b88b6;flex:none}.row[data-kind=folder] .fileicon{color:var(--blue)}.detail{min-width:0;flex:1;display:block;text-align:left;padding:1px 0;border-radius:3px}.name{font-size:12px;display:block;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.meta{font-size:10px;display:block;color:var(--muted);padding-top:4px}.remove{color:#87909f;opacity:0}.row:hover .remove,.remove:focus-visible{opacity:1}.warning{color:#b77b1b!important}#more[hidden]{display:none}#more{display:flex;margin:4px auto;color:var(--blue);font-size:11px}footer{padding:11px 14px;display:flex;align-items:center;gap:7px;color:var(--muted);font-size:10px;border-top:1px solid var(--line);margin-top:10px}.dot{height:5px;width:5px;border-radius:50%;background:#5b9b83}footer .hint{flex:1}#status{padding:0 17px 10px;color:var(--muted);font-size:11px;line-height:1.5;overflow-wrap:anywhere}#status:empty{display:none}#prefs{padding:8px 18px 18px}#prefs[hidden],#body[hidden]{display:none}label{display:flex;align-items:center;gap:8px;justify-content:space-between;margin:12px 0}input{width:85px;border:1px solid var(--line);background:var(--field);color:inherit;border-radius:8px;padding:7px}#save{background:var(--blue);color:white;width:100%;margin-top:12px}#panel.dragging #body{outline:2px solid var(--blue);background:#1764db0c}.support{font-size:10px;color:var(--muted);line-height:1.7}#status.error{color:#c0483d}
    :host([dark]){--surface:rgba(34,37,44,.97);--line:#ffffff19;--muted:#a5aebb;--hover:#8ab4f815;--blue:#8bb9ff;--field:#2c3039;color:#edf0f5;color-scheme:dark}:host([dark]) .brand{background:#32507a44}:host([dark]) #save{background:#417fd7;color:white}
    @media(prefers-reduced-motion:no-preference){#panel{animation:pp-appear .15s ease-out}@keyframes pp-appear{from{opacity:0;transform:translateY(5px)}to{opacity:1;transform:translateY(0)}}}
    @media(prefers-contrast:more){:host{--line:currentColor;--surface:Canvas;--muted:CanvasText}#panel{backdrop-filter:none}}
  `;
  // Semantic, appearance-specific colors; restrained translucent chrome and opaque reading surfaces.
  const theme = `
    :host{--surface:rgba(249,250,249,.98);--line:#dce3df;--muted:#586960;--hover:#e9f1ec;--blue:#196849;--field:#fff;color:#1d2b24;font-family:"Segoe UI Variable","Segoe UI","Microsoft YaHei UI",sans-serif;letter-spacing:.1px}
    #trigger{height:30px;border-radius:9px;background:var(--field);box-shadow:0 1px 2px #172b1e06;color:var(--blue);border-color:var(--line)}#trigger[aria-expanded=true]{background:var(--hover);color:var(--blue)}#badge{background:var(--hover);color:var(--blue)}
    #panel{width:min(408px,calc(100vw - 20px));max-height:min(620px,calc(100vh - 64px));border-radius:22px;border:1px solid var(--line);box-shadow:0 20px 60px #142a2026,0 4px 12px #142a2010,inset 0 1px #ffffffaa;backdrop-filter:blur(28px) saturate(115%)}
    header{padding:20px 18px 16px;gap:12px}.brand{width:40px;height:40px;border-radius:13px;background:linear-gradient(145deg,#e6f2eb,#d6e9df);color:#196849;box-shadow:inset 0 1px #fff9}.brand svg{width:22px;height:22px}h2{font-size:16px;font-weight:650;letter-spacing:.2px}.subtitle{margin-top:5px;font-size:11px;color:var(--muted)}
    nav{padding:0 18px 12px;gap:7px}nav button{min-height:36px;border-radius:10px;font-size:12px;box-shadow:0 1px 2px #18281d06}nav button.primary{color:#fff;background:#196849;border-color:#196849}nav button.primary:hover{background:#14563c}nav svg{width:16px;height:16px}
    .library-tools{display:flex;align-items:center;gap:5px;margin:0 18px 8px;min-height:28px}.library-tools>span{flex:1;color:var(--muted);font-size:11px}.library-tools button{width:auto;min-height:28px;font-size:11px;padding:4px 7px;gap:4px;color:var(--muted)}.library-tools svg{width:14px;height:14px}#undo-clear{color:var(--blue)}
    #body{margin:0 12px;border:1px solid var(--line);background:var(--field);min-height:170px;max-height:320px;border-radius:14px;padding:6px}.row{padding:9px 7px;gap:9px;min-height:54px}.row+.row{border-top:1px solid color-mix(in srgb,var(--line) 55%,transparent)}.row:hover{background:var(--hover)}.row[aria-selected=true]{background:var(--hover);box-shadow:inset 3px 0 var(--blue)}.row .name{font-size:12px;line-height:1.5}.row .meta{color:var(--muted);font-size:10px}.fileicon{color:var(--muted);background:var(--hover);border-radius:8px;width:31px;height:35px}.row[data-kind=folder] .fileicon{color:var(--blue)}.actions button{color:var(--muted)}.remove{color:var(--muted)}
    #empty{border:0;margin:0;padding:30px 12px;background:var(--field)}#empty .large{background:var(--hover);color:var(--blue);width:48px;height:48px;border-radius:15px;margin-bottom:14px}#empty b{font-size:14px;color:inherit}#empty p{max-width:240px;margin:9px auto 0;font-size:11px;color:var(--muted)}
    footer{margin-top:12px;padding:10px 17px;border-top:0;color:var(--muted);font-size:10px}.dot{background:#377c57}#status{padding:10px 19px 0;color:var(--blue);font-size:11px}#status:empty{padding:0}#status.error{color:#a52f32}.warning{color:#875712!important}#prefs{padding:8px 20px 20px}#prefs input{border-radius:9px}#save{border-radius:10px;min-height:36px}button:focus-visible,input:focus-visible,select:focus-visible{outline:2px solid var(--blue);outline-offset:3px}
    :host([dark]){--surface:rgba(28,34,31,.98);--line:#414d45;--muted:#b4c1b8;--hover:#303e35;--blue:#8ed4ad;--field:#222b25;color:#edf3ef}:host([dark]) .brand{background:linear-gradient(145deg,#304c3c,#253e30);color:#8ed4ad;box-shadow:inset 0 1px #ffffff0a}:host([dark]) #panel{box-shadow:0 20px 60px #0006,inset 0 1px #ffffff08}:host([dark]) nav button.primary{color:#14291c;background:#8ed4ad;border-color:#8ed4ad}:host([dark]) #save{background:#8ed4ad;color:#14291c}:host([dark]) #status.error{color:#ffaaa9}:host([dark]) .warning{color:#efc483!important}
    @media(max-width:360px){nav button svg{display:none}header{padding:16px 13px 12px;gap:8px}nav{padding:0 12px 10px;gap:5px}nav button{font-size:11px;padding:5px}.library-tools{margin:0 13px 6px}#body{margin:0 8px}.actions button{width:24px}h2{font-size:15px}.brand{width:34px;height:34px}}
    @media(prefers-reduced-transparency:reduce){#panel{backdrop-filter:none;background:var(--field)}}
    @media(prefers-reduced-motion:reduce){#panel{animation:none}*{transition:none!important}}
    @media(prefers-contrast:more){:host,:host([dark]){--line:CanvasText;--surface:Canvas;--field:Canvas;--muted:CanvasText;--blue:LinkText;color:CanvasText}#panel{backdrop-filter:none}.row+.row{border-top-color:CanvasText}}
  `;
  class Workspace extends HTMLElement {
    constructor() {
      super(); this.attachShadow({ mode: 'open' }); this.stack = []; this.currentSession = session(); this.request = 0;
      this.shadowRoot.innerHTML = `<style>${css}</style><button id="trigger" type="button" aria-label="打开参考工作区" aria-expanded="false" aria-controls="panel">${icon('folder')}<span>工作区</span><span id="badge"></span></button><section id="panel" popover="auto" role="dialog" aria-labelledby="title"><header><div class="brand">${icon('folder')}</div><div class="heading"><h2 id="title">参考工作区</h2><div class="subtitle">文件留在本机，AI 按需读取</div></div><button class="icon" id="refresh" aria-label="刷新文件列表" title="刷新">${icon('refresh')}</button><button class="icon" id="close" aria-label="收起工作区" title="收起">${icon('close')}</button></header><nav><button id="add" class="primary">${icon('plus')}添加文件</button><button id="folder">${icon('folder')}选择文件夹</button></nav><div id="body"><div id="crumb"></div><div id="list"></div><button id="more" hidden>显示更多</button></div><div id="prefs" hidden><b>工作区首选项</b><label for="limit">单文件大小上限 <span><input id="limit" type="number" min="1" max="1024" step="1" value="100" aria-label="单文件大小上限，MB"> MB</span></label><p>默认 100 MB，适用于新添加和已选文件。大文件仅在需要时分段读取。</p><button id="save">保存首选项</button><p class="support">PDF · Word · Excel · PowerPoint<br>OpenDocument · 邮件 · 文本与代码 · 图片<br>图片与扫描页需要支持视觉的模型。</p></div><div id="status" role="status" aria-live="polite"></div><footer><span class="dot"></span><span class="hint">原文保留 · 按需读取</span><button id="settings" class="icon" aria-label="工作区首选项" title="首选项">${icon('settings')}</button></footer></section>`;
      // Native invoker association prevents light-dismiss from closing the panel
      // on pointerdown and then reopening it in a subsequent click handler.
      this.$('trigger').popoverTargetElement = this.$('panel');
      this.selection = new Set();
      this.shadowRoot.querySelector('style').textContent += `nav button{padding:6px 7px;white-space:nowrap}.row[aria-selected=true]{background:#1764db15}.row[data-drop=inside],#crumb [data-drop=inside]{outline:2px solid var(--blue);outline-offset:-2px}.row[data-drop=before]{box-shadow:0 -2px 0 var(--blue)}.row[data-drop=after]{box-shadow:0 2px 0 var(--blue)}.actions{display:flex;gap:0}.actions button{width:26px;padding:5px}.actions svg{width:15px;height:15px}#organize{padding:8px 17px 15px}#organize[hidden]{display:none}#organize input,#organize select{width:100%;padding:9px;border:1px solid var(--line);border-radius:9px;background:var(--field);color:inherit;margin:10px 0}#organize .form-actions{display:flex;justify-content:flex-end;gap:5px}#organize .primary{background:var(--blue);color:var(--field)}#crumb{flex-wrap:wrap}#crumb button{min-height:26px;max-width:150px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:11px;padding:4px 6px}#body[data-drop=inside]{outline:2px dashed var(--blue);outline-offset:-2px}`;
      const create = document.createElement('button'); create.id = 'new-folder'; create.innerHTML = icon('folder') + '新建文件夹'; create.title = '新建虚拟文件夹'; create.onclick = () => this.organize('create'); this.shadowRoot.querySelector('nav').append(create);
      const depthLabel=document.createElement('label');depthLabel.innerHTML='递归查找深度 <span><input id="search-depth" type="number" min="1" max="12" value="3" aria-label="递归查找深度"> 层</span>';this.$('prefs').querySelector('label').after(depthLabel);
      const organize = document.createElement('div'); organize.id = 'organize'; organize.hidden = true;
      organize.innerHTML = '<b id="organize-title"></b><input id="folder-name" maxlength="80" aria-label="虚拟文件夹名称" placeholder="例如：市场资料"><select id="move-target" aria-label="目标虚拟文件夹"></select><p>整理工作区引用，保留本机文件位置。</p><div class="form-actions"><button id="organize-cancel">取消</button><button id="organize-save" class="primary">完成</button></div>';
      this.$('body').before(organize); this.$('organize-cancel').onclick = () => this.endOrganize();
      this.$('organize-save').onclick = () => this.run(async () => {
        if (this.organizing === 'move') await api.move(session(), this.organizeIds, this.$('move-target').value);
        else await api.collection(session(), { operation:this.organizing, id:this.organizeIds[0], name:this.$('folder-name').value, parent:this.virtualParent() });
        if(this.stack.at(-1)&&!this.stack.at(-1).virtual)this.goVirtual(this.virtualParent());
        this.endOrganize(); this.selection.clear(); await this.refresh(); this.status('已整理');
      });
      this.$('folder-name').onkeydown = e => { if (e.key === 'Enter') this.$('organize-save').click(); if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); this.endOrganize(); } };
      this.$('body').addEventListener('dragover', e => { if (!e.target.closest('.row')) this.dragOver(e, this.$('body'), this.virtualParent(), 'inside'); });
      this.$('body').addEventListener('drop', e => { if (!e.target.closest('.row')) this.drop(e, this.virtualParent()); });
      this.$('close').onclick = () => this.$('panel').hidePopover();
      this.$('panel').addEventListener('toggle', e => { const open = e.newState === 'open'; this.$('trigger').setAttribute('aria-expanded', String(open)); if (open) { this.place(); this.refresh(); } });
      this.$('add').onclick = () => this.choose(false); this.$('folder').onclick = () => this.choose(true);
      this.$('refresh').onclick = () => this.refresh();
      const clear = document.createElement('button'); clear.id = 'clear'; clear.className = 'icon';
      clear.title = '清空工作区（不删除源文件）'; clear.setAttribute('aria-label', clear.title); clear.innerHTML = icon('trash');
      const management=document.createElement('div');management.className='library-tools';management.innerHTML='<span>本次对话的资料</span>';this.$('body').before(management);management.append(clear);clear.innerHTML += '<span>清空</span>';
      const undo = document.createElement('button'); undo.id = 'undo-clear'; undo.className = 'icon'; undo.hidden = true;
      undo.title = '撤销清空'; undo.setAttribute('aria-label', undo.title); undo.innerHTML = icon('back'); clear.before(undo); undo.innerHTML += '<span>撤销清空</span>';
      this.shadowRoot.querySelector('style').textContent += '#undo-clear[hidden]{display:none}';
      undo.onclick = () => this.run(async () => {
        await api.undoClear(session()); this.stack = []; this.selection.clear(); this.endOrganize();
        this.$('prefs').hidden = true; await this.refresh(); this.status('已恢复原工作区，新添加的资料也已保留');
        window.dispatchEvent(new Event('gridpilot:files'));
      });
      clear.onclick = () => this.run(async () => {
        await api.clear(session()); this.stack = []; this.selection.clear(); this.endOrganize();
        this.$('prefs').hidden = true; await this.refresh();
        this.status('工作区已清空，源文件和已有底稿文件均保留');
        window.dispatchEvent(new Event('gridpilot:files'));
      });
      this.$('more').onclick = () => this.refresh(this.next, true);
      this.$('settings').onclick = async () => {
        const show = this.$('prefs').hidden; this.$('prefs').hidden = !show; this.$('body').hidden = show;
        if (show) { this.endOrganize();this.$('body').hidden=true;const value = await api.preferences(); this.$('limit').value = value.maxFileMB;this.$('search-depth').value=value.searchDepth; this.$('limit').focus(); }
      };
      this.$('save').onclick = () => this.run(async () => { await api.preferences({ maxFileMB: Number(this.$('limit').value), searchDepth:Number(this.$('search-depth').value) }); this.status('已保存'); this.$('prefs').hidden = true; this.$('body').hidden = false; await this.refresh(); });
      this.$('panel').addEventListener('dragover', e => { if (e.dataTransfer.types.includes('Files')) { e.preventDefault(); e.stopPropagation(); this.$('panel').classList.add('dragging'); } });
      this.$('panel').addEventListener('dragleave', () => this.$('panel').classList.remove('dragging'));
      this.$('panel').addEventListener('drop', e => this.drop(e, this.virtualParent()));
      this.$('panel').addEventListener('dragend', () => this.clearDrop());
      this.shadowRoot.querySelector('style').textContent += theme;
      this.resize = () => { if (this.$('panel').matches(':popover-open')) this.place(); };
      this.changed = () => this.onSession();
    }
    $(id) { return this.shadowRoot.getElementById(id); }
    virtualParent() { return [...this.stack].reverse().find(e => e.virtual)?.id || ''; }
    goVirtual(parent) { const stack=[];let id=parent.replace(/^virtual:/,'');while(id){const c=this.collections.find(x=>x.id===id);if(!c)break;stack.unshift({...c,id:'virtual:'+c.id,virtual:true,kind:'folder'});id=c.parent;}this.stack=stack; }
    endOrganize() { this.$('organize').hidden = true; this.$('body').hidden = false; this.organizing = ''; }
    organize(operation, entry) {
      this.organizing = operation; this.organizeIds = entry ? [entry.id] : [];
      this.$('organize').hidden = false; this.$('body').hidden = true; this.$('prefs').hidden = true;
      this.$('organize-title').textContent = {create:'新建虚拟文件夹',rename:'重命名虚拟文件夹',move:'移动到'}[operation];
      this.$('folder-name').hidden = operation === 'move'; this.$('move-target').hidden = operation !== 'move';
      if (operation === 'move') {
        const select = this.$('move-target'); select.replaceChildren(new Option('所有资料',''));
        for (const folder of this.collections || []) { let labels=[folder.name],parent=folder.parent; while(parent){const p=this.collections.find(c=>c.id===parent);if(!p)break;labels.unshift(p.name);parent=p.parent;} select.add(new Option(labels.join(' / '),'virtual:'+folder.id)); }
        select.focus();
      } else { this.$('folder-name').value = entry?.name || ''; this.$('folder-name').focus(); this.$('folder-name').select(); }
    }
    clearDrop() { for (const element of this.shadowRoot.querySelectorAll('[data-drop]')) delete element.dataset.drop; this.$('panel').classList.remove('dragging'); }
    dragOver(e, element, parent, placement='inside') {
      if (!e.dataTransfer.types.includes('application/x-gridpilot-entry') && !e.dataTransfer.types.includes('Files')) return;
      e.preventDefault(); e.stopPropagation(); this.clearDrop(); element.dataset.drop = placement; e.dataTransfer.dropEffect = e.dataTransfer.types.includes('Files') ? 'copy' : 'move';
    }
    drop(e, parent, before='', after='') {
      e.preventDefault(); e.stopPropagation(); this.clearDrop();
      const internal = e.dataTransfer.getData('application/x-gridpilot-entry');
      this.run(async () => {
        if (internal) { const value=JSON.parse(internal); if(value.session!==session())throw Error('请在同一聊天内整理资料'); await api.move(session(),value.ids,parent,before,after);this.selection.clear();if(this.stack.at(-1)&&!this.stack.at(-1).virtual)this.goVirtual(parent);await this.refresh();this.status('已整理，原文件位置保持不变'); }
        else if(e.dataTransfer.files.length) await this.added(await api.add(session(),Array.from(e.dataTransfer.files),parent),true);
      });
    }
    connectedCallback() { window.addEventListener('resize', this.resize); window.addEventListener('gridpilot:session', this.changed); window.addEventListener('gridpilot:files', this.changed); this.refresh(); }
    disconnectedCallback() { window.removeEventListener('resize', this.resize); window.removeEventListener('gridpilot:session', this.changed); window.removeEventListener('gridpilot:files', this.changed); }
    place() {
      const r = this.getBoundingClientRect(), p = this.$('panel');
      p.style.left = `${Math.max(10, Math.min(r.left, innerWidth - Math.min(372, innerWidth - 20) - 10))}px`;
      p.style.bottom = `${Math.max(10, innerHeight - r.top + 9)}px`;
      p.style.maxHeight = `${Math.max(160, Math.min(530, r.top - 22))}px`;
      const bg = getComputedStyle(document.body).backgroundColor.match(/[\d.]+/g);
      this.toggleAttribute('dark', document.documentElement.dataset.mantineColorScheme === 'dark' || document.documentElement.classList.contains('dark') || !!(bg && bg.length >= 3 && Number(bg[0]) + Number(bg[1]) + Number(bg[2]) < 240 && (bg.length < 4 || Number(bg[3]) > 0)));
    }
    status(text, error = false) { this.$('status').textContent = text; this.$('status').classList.toggle('error', error); }
    async run(fn) { try { this.status('正在处理…'); await fn(); } catch (e) { this.status(e.message, true); } }
    async onSession() {
      const next = session(), previous = this.currentSession; this.currentSession = next;
      if (previous !== next) { this.stack = []; this.selection.clear(); this.endOrganize(); ++this.request; }
      await this.refresh();
    }
    async choose(folder) { await this.run(async () => this.added(await api.choose(session(), folder, this.virtualParent()),true)); }
    async added(result, keepLocation=false) { this.status(result.errors?.map(e => `${e.name}：${e.message}`).join('；') || (result.cancelled ? '' : '已添加 · AI 可按需读取'), !!result.errors?.length); if(!keepLocation)this.stack = [];else if(this.stack.at(-1)&&!this.stack.at(-1).virtual)this.goVirtual(this.virtualParent()); await this.refresh(); }
    async refresh(cursor = 0, append = false) {
      const serial = ++this.request, selected = session(), directory = this.stack.at(-1)?.id || '';
      try {
        const value = await api.summary(selected, directory, cursor);
        if (serial !== this.request || selected !== session()) return;
        this.$('badge').textContent = value.sourceCount ? String(value.sourceCount) : '';
        this.$('undo-clear').hidden = !value.canUndoClear; this.collections = value.collections;
        this.$('trigger').title = `参考工作区 · ${value.sourceCount} 个来源 · 单文件上限 ${value.maxFileMB} MB`;
        const crumb = this.$('crumb'); crumb.replaceChildren();
        const breadcrumbs=[{id:'',name:'所有资料',virtual:true},...this.stack];
        breadcrumbs.forEach((entry,index)=>{const button=document.createElement('button');button.textContent=entry.name;button.title=entry.virtual?'可拖到此处':entry.name;button.onclick=()=>{this.stack=this.stack.slice(0,index);this.refresh();};if(entry.virtual){button.ondragover=e=>this.dragOver(e,button,entry.id);button.ondrop=e=>this.drop(e,entry.id);}crumb.append(button);});
        const list = this.$('list'); if (!append) list.replaceChildren();
        if (!value.entries.length && !cursor) { const empty = document.createElement('div'); empty.id = 'empty'; empty.innerHTML = `<div class="large">${icon('folder')}</div><b>${directory ? '这个文件夹里没有可见文件' : '把参考资料放在这里'}</b><p>${directory ? '可返回上一级继续浏览' : '拖入文件，或选择本机文件夹<br>无需整理上传，按问题读取所需内容'}</p>`; list.append(empty); }
        for (const entry of value.entries) {
          const row = document.createElement('div'); row.className = 'row'; row.dataset.kind = entry.kind;
          row.dataset.id=entry.id;row.tabIndex=0;row.setAttribute('aria-selected',String(this.selection.has(entry.id)));
          const movable = !entry.generated || entry.kind === 'file';
          row.draggable=movable;
          const selectEntry=e=>{if(e.ctrlKey||e.metaKey){e.preventDefault();e.stopPropagation();this.selection.has(entry.id)?this.selection.delete(entry.id):this.selection.add(entry.id);row.setAttribute('aria-selected',String(this.selection.has(entry.id)));return true;}return false;};
          row.onclick=e=>selectEntry(e);
          row.onkeydown=e=>{if(e.target!==row)return;if(e.key==='F2'&&entry.virtual){e.preventDefault();this.organize('rename',entry);}else if(e.key===' '){e.preventDefault();this.selection.has(entry.id)?this.selection.delete(entry.id):this.selection.add(entry.id);row.setAttribute('aria-selected',String(this.selection.has(entry.id)));}else if(e.key==='Escape'){this.selection.clear();this.refresh();}};
          row.ondragstart=e=>{if(!movable){e.preventDefault();return;}const ids=this.selection.has(entry.id)?[...this.selection]:[entry.id];e.dataTransfer.setData('application/x-gridpilot-entry',JSON.stringify({session:selected,ids}));e.dataTransfer.effectAllowed='copyMove';};
          row.ondragover=e=>{const box=row.getBoundingClientRect(),ratio=(e.clientY-box.top)/box.height;const place=entry.virtual&&ratio>.23&&ratio<.77?'inside':ratio<.5?'before':'after';if(entry.generated&&!entry.virtual)return;this.dragOver(e,row,place==='inside'?entry.id:this.virtualParent(),place);};
          row.ondrop=e=>{const placement=row.dataset.drop;if(placement)this.drop(e,placement==='inside'?entry.id:this.virtualParent(),placement==='before'?entry.id:'',placement==='after'?entry.id:'');};
          const mark = document.createElement('span'); mark.className = 'fileicon'; mark.innerHTML = icon(entry.kind === 'folder' ? 'folder' : entry.vision && entry.extension !== '.pdf' ? 'image' : 'file');
          const detail = document.createElement(entry.kind === 'folder' ? 'button' : 'div'); detail.className = 'detail'; detail.title = entry.name;
          const name = document.createElement('span'); name.className = 'name'; name.textContent = entry.name;
          const meta = document.createElement('span'); meta.className = 'meta';
          const issue = entry.unavailable ? '文件已移动或不可用' : entry.overLimit ? '超过大小上限' : !entry.supported ? '此格式暂不支持' : '';
          meta.textContent = issue || (entry.virtual ? '虚拟文件夹 · 拖入整理' : entry.kind === 'folder' ? '本机文件夹 · 点击浏览' : `${entry.size < 1048576 ? (entry.size / 1024).toFixed(1) + ' KB' : (entry.size / 1048576).toFixed(1) + ' MB'} · ${entry.vision ? '支持视觉读取' : '按需读取'}`);
          if (issue) meta.classList.add('warning');
          detail.append(name, meta); if (entry.kind === 'folder') detail.onclick = e => { if(selectEntry(e))return;e.stopPropagation();this.stack.push(entry); this.selection.clear();this.refresh(); };
          row.append(mark, detail);
          if(movable){const actions=document.createElement('span');actions.className='actions';for(const [action,label,glyph] of (entry.virtual?[['rename','重命名','edit'],['move','移动到','move'],['delete','删除虚拟文件夹','trash']]:[['move','整理到虚拟文件夹','move']])){const button=document.createElement('button');button.className='icon';button.title=label;button.setAttribute('aria-label',`${label} ${entry.name}`);button.innerHTML=icon(glyph);button.onclick=e=>{e.stopPropagation();if(action==='delete')this.run(async()=>{const r=await api.collection(selected,{operation:'delete',id:entry.id});if(!r.cancelled){this.selection.clear();await this.refresh();this.status('虚拟文件夹已移除，原文件保留');}});else this.organize(action,entry);};actions.append(button);}row.append(actions);}
          if (entry.generated && entry.kind === 'file') {
            for (const [action, label, glyph] of [['open','在 Excel 中打开','file'],['export','导出底稿','export'],['delete','删除底稿','trash']]) {
              const control = document.createElement('button'); control.className = 'icon'; control.innerHTML = icon(glyph); control.title = label; control.setAttribute('aria-label', `${label} ${entry.name}`);
              control.onclick = () => this.run(async () => { const r = await api[action](selected, entry.id); this.status(r.cancelled ? '' : action === 'open' ? '底稿已在 Excel 中打开' : action === 'export' ? '底稿已导出' : '底稿已删除'); await this.refresh(); }); row.append(control);
            }
          } else if ((!directory || directory.startsWith('virtual:')) && !entry.generated && !entry.virtual) { const remove = document.createElement('button'); remove.className = 'remove icon'; remove.innerHTML = icon('close'); remove.title = '移出工作区（不删除原文件）'; remove.setAttribute('aria-label', `移除 ${entry.name}`); remove.onclick = () => this.run(async () => { await api.remove(selected, entry.id); await this.refresh(); this.status('已移除，原文件保留'); }); row.append(remove); }
          list.append(row);
        }
        this.next = value.next_cursor; this.$('more').hidden = this.next === null;
        if(value.directory_limit_reached)this.status('此目录已达到 10 万项探测上限，请直接选择更具体的子文件夹。');
      } catch (e) { if (serial === this.request) this.status(e.message, true); }
    }
  }
  customElements.define('gridpilot-workspace', Workspace);
  function mount() { const slot = document.getElementById('gridpilot-toolbar-slot'); if (slot && !slot.querySelector('gridpilot-workspace')) slot.prepend(document.createElement('gridpilot-workspace')); }
  new MutationObserver(mount).observe(document.body, { childList: true, subtree: true }); mount();
})();
