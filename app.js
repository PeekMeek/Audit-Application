/* Additive Applications — notebook photo → review → Excel + Teams */
(function () {
  'use strict';
  const F = window.AAFields;
  const $ = (s) => document.querySelector(s);
  const el = (tag, props, ...kids) => {
    const n = document.createElement(tag);
    for (const k in (props || {})) {
      if (k === 'class') n.className = props[k];
      else if (k === 'text') n.textContent = props[k];
      else if (k.startsWith('on')) n.addEventListener(k.slice(2), props[k]);
      else if (props[k] !== undefined && props[k] !== null && props[k] !== false) n.setAttribute(k, props[k] === true ? '' : props[k]);
    }
    kids.flat().forEach(c => c != null && n.append(c));
    return n;
  };

  const DEFAULT_MODELS = { anthropic: 'claude-sonnet-5-5', openai: '', azure: '' };
  const NOTEBOOK_MAX = 1568, PART_MAX = 1600;

  /* ================= settings ================= */
  const SETTINGS_KEY = 'aa.settings.v1';
  let settings = loadSettings();
  function loadSettings() {
    try { return Object.assign({ name: '', owner: '', folder: '', provider: 'anthropic', apiKey: '', model: '', azureEndpoint: '', azureDeployment: '', webhook: '', thumb: false }, JSON.parse(localStorage.getItem(SETTINGS_KEY) || '{}')); }
    catch (e) { return { name: '', owner: '', provider: 'anthropic', apiKey: '', model: '', azureEndpoint: '', azureDeployment: '', webhook: '', thumb: false }; }
  }
  function saveSettingsToStore() {
    try { localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings)); return true; }
    catch (e) { toast('Settings could not be saved on this device.', true); return false; }
  }

  /* ================= storage (IndexedDB) ================= */
  const DB = {
    db: null,
    open() {
      if (this.db) return Promise.resolve(this.db);
      return new Promise((res, rej) => {
        const r = indexedDB.open('additive-applications', 1);
        r.onupgradeneeded = () => r.result.createObjectStore('entries', { keyPath: 'id' });
        r.onsuccess = () => { this.db = r.result; res(this.db); };
        r.onerror = () => rej(r.error);
      });
    },
    async tx(mode, fn) {
      const db = await this.open();
      return new Promise((res, rej) => {
        const t = db.transaction('entries', mode), store = t.objectStore('entries');
        const req = fn(store);
        t.oncomplete = () => res(req && req.result);
        t.onerror = () => rej(t.error);
      });
    },
    all() { return this.tx('readonly', s => s.getAll()); },
    put(e) { return this.tx('readwrite', s => s.put(e)); },
    del(id) { return this.tx('readwrite', s => s.delete(id)); }
  };

  /* ================= ui helpers ================= */
  let toastTimer;
  function toast(msg, isError) {
    const t = $('#toast');
    t.textContent = msg; t.classList.toggle('error', !!isError); t.hidden = false;
    clearTimeout(toastTimer); toastTimer = setTimeout(() => { t.hidden = true; }, isError ? 6000 : 2800);
  }
  const screens = ['settings', 'home', 'capture', 'review'];
  let current = null;
  function show(name, title, opts) {
    opts = opts || {};
    screens.forEach(s => { $('#screen-' + s).hidden = s !== name; });
    $('#title').textContent = title;
    $('#btn-back').hidden = name === 'home' || (name === 'settings' && !settings.name);
    $('#btn-settings').hidden = name === 'settings';
    if (name !== 'home' && current === 'home' && !opts.noPush) history.pushState({ s: name }, '');
    current = name;
    window.scrollTo(0, 0);
  }
  function goBack() {
    if (current === 'review' && draft && draft.isNew && draft.dirty && !confirm('Discard this entry? It hasn\'t been saved.')) return false;
    if (current === 'review' && draft && draft.isNew) { show('capture', 'New application', { noPush: true }); return true; }
    goHome(); return true;
  }
  function goHome() {
    draft = null; show('home', 'Additive Applications', { noPush: true }); renderHome();
    if (history.state && history.state.s) history.back(); // drop our pushed state; popstate on home is ignored
  }
  window.addEventListener('popstate', () => {
    if (current === 'home') return;
    goBack();
    if (current !== 'home') history.pushState({ s: current }, ''); // stayed in the app (cancelled, or review → capture)
  });
  $('#btn-back').addEventListener('click', () => { if (current !== 'home') goBack(); });
  $('#btn-settings').addEventListener('click', openSettings);

  $('#viewer').addEventListener('click', () => { $('#viewer').hidden = true; });
  function view(src) { $('#viewer-img').src = src; $('#viewer').hidden = false; }

  function fmtDate(iso) {
    if (!iso) return '';
    const d = new Date(iso);
    return d.toLocaleDateString(undefined, { month: 'short', day: 'numeric' }) + ' ' + d.toLocaleTimeString(undefined, { hour: 'numeric', minute: '2-digit' });
  }
  const urgColor = n => `hsl(${F.urgencyHue(n)} 70% 42%)`;
  const firstName = (n) => String(n || '').trim().split(/\s+/)[0] || '';
  const uid = () => (crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(36) + Math.random().toString(36).slice(2));

  /* ================= images ================= */
  function imageFrom(src) {
    return new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = () => rej(new Error('That file couldn\'t be opened as an image.')); i.src = src; });
  }
  async function toJpeg(src, maxDim, quality) {
    const img = await imageFrom(src);
    const k = Math.min(1, maxDim / Math.max(img.naturalWidth, img.naturalHeight));
    const w = Math.max(1, Math.round(img.naturalWidth * k)), h = Math.max(1, Math.round(img.naturalHeight * k));
    const c = document.createElement('canvas'); c.width = w; c.height = h;
    const ctx = c.getContext('2d'); ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, w, h); ctx.drawImage(img, 0, 0, w, h);
    return { dataUrl: c.toDataURL('image/jpeg', quality), w, h };
  }
  async function fileToPhoto(file, maxDim) {
    const url = URL.createObjectURL(file);
    try { return await toJpeg(url, maxDim, 0.85); } finally { URL.revokeObjectURL(url); }
  }
  // Turn a photo clockwise by 90, 180 or 270 degrees.
  async function rotatePhoto(photo, deg) {
    deg = ((deg % 360) + 360) % 360;
    if (!deg) return photo;
    const img = await imageFrom(photo.dataUrl);
    const w = img.naturalWidth, h = img.naturalHeight, swap = deg === 90 || deg === 270;
    const c = document.createElement('canvas'); c.width = swap ? h : w; c.height = swap ? w : h;
    const ctx = c.getContext('2d');
    ctx.translate(c.width / 2, c.height / 2); ctx.rotate(deg * Math.PI / 180); ctx.drawImage(img, -w / 2, -h / 2);
    return { dataUrl: c.toDataURL('image/jpeg', 0.88), w: c.width, h: c.height };
  }
  function thumb(photo, onRemove, onRotate) {
    return el('div', { class: 'thumb' },
      el('img', { src: photo.dataUrl, alt: '', onclick: () => view(photo.dataUrl) }),
      onRemove ? el('button', { class: 'x', 'aria-label': 'Remove photo', onclick: onRemove, text: '✕' }) : null,
      onRotate ? el('button', { class: 'rot', 'aria-label': 'Rotate photo', title: 'Rotate', onclick: onRotate, text: '⟳' }) : null);
  }

  /* ================= settings screen ================= */
  function openSettings() {
    $('#set-name').value = settings.name;
    $('#set-owner').value = settings.owner;
    $('#set-folder').value = settings.folder || '';
    $('#set-provider').value = settings.provider;
    $('#set-key').value = settings.apiKey;
    $('#set-model').value = settings.model || DEFAULT_MODELS[settings.provider] || '';
    $('#set-azure-endpoint').value = settings.azureEndpoint;
    $('#set-azure-deployment').value = settings.azureDeployment;
    $('#set-webhook').value = settings.webhook;
    $('#set-thumb').checked = !!settings.thumb;
    $('#setup-intro').textContent = settings.name ? 'Everything here stays on this device.' : 'Welcome! Set this phone up once. Everything here stays on this device.';
    syncProviderFields();
    show('settings', settings.name ? 'Settings' : 'Set up');
  }
  function syncProviderFields() {
    const p = $('#set-provider').value;
    $('#ai-fields').hidden = p === 'none';
    $('#azure-fields').hidden = p !== 'azure';
    $('#lbl-model').hidden = p === 'azure';
    $('#set-model').placeholder = p === 'openai' ? 'e.g. a current GPT vision model' : 'claude-sonnet-5-5';
  }
  $('#set-provider').addEventListener('change', () => {
    const p = $('#set-provider').value;
    if (!$('#set-model').value || Object.values(DEFAULT_MODELS).includes($('#set-model').value)) $('#set-model').value = DEFAULT_MODELS[p] || '';
    syncProviderFields();
  });
  function readSettingsForm() {
    return Object.assign({}, settings, {
      name: $('#set-name').value.trim(),
      owner: $('#set-owner').value.trim(),
      folder: $('#set-folder').value.trim() || firstName($('#set-name').value),
      provider: $('#set-provider').value,
      apiKey: $('#set-key').value.trim(),
      model: $('#set-model').value.trim(),
      azureEndpoint: $('#set-azure-endpoint').value.trim(),
      azureDeployment: $('#set-azure-deployment').value.trim(),
      webhook: $('#set-webhook').value.trim(),
      thumb: $('#set-thumb').checked
    });
  }
  $('#btn-save-settings').addEventListener('click', () => {
    const s = readSettingsForm();
    if (!s.name) { toast('Enter your name. It starts every Teams message.', true); $('#set-name').focus(); return; }
    if (s.webhook && !/^https:\/\//i.test(s.webhook)) { toast('The webhook URL should start with https://', true); return; }
    settings = s;
    if (saveSettingsToStore()) { toast('Saved'); goHome(); }
  });
  $('#btn-test-teams').addEventListener('click', async (ev) => {
    const s = readSettingsForm();
    if (!s.webhook) { toast('Paste the webhook URL first.', true); return; }
    ev.target.disabled = true;
    try { await sendToTeams(s.webhook, F.buildTestPayload(s.name)); toast('Test sent. Check the Teams chat in a few seconds.'); }
    catch (e) { toast(e.message, true); }
    finally { ev.target.disabled = false; }
  });

  /* ================= home ================= */
  let entries = [];
  async function renderHome() {
    try { entries = (await DB.all()).sort((a, b) => (b.createdAt || '').localeCompare(a.createdAt || '')); }
    catch (e) { entries = []; toast('Saved entries couldn\'t be loaded on this device.', true); }
    const list = $('#entry-list'); list.replaceChildren();
    entries.forEach(e => {
      const d = F.derive(e);
      const chip = el('span', { class: 'chip', text: d.urgency || '–' });
      chip.style.background = d.urgency ? urgColor(d.urgency) : 'var(--muted)';
      const pic = e.partPhoto ? el('img', { class: 'pic', src: e.partPhoto.dataUrl, alt: '' }) : el('div', { class: 'pic none', text: '▧' });
      const sub = [d.department, fmtDate(e.createdAt)].filter(Boolean).join(' · ');
      list.append(el('div', { class: 'entry', role: 'button', tabindex: '0', onclick: () => openReview(e, false) },
        pic,
        el('div', { class: 'meta' }, el('div', { class: 'name', text: d.appName || '(no name)' }),
          el('div', { class: 'sub' }, e.postedAt ? el('span', { class: 'posted', text: '✓ Posted · ' }) : null, sub)),
        chip));
    });
    $('#empty-state').hidden = entries.length > 0;
    $('#export-bar').hidden = entries.length === 0;
    $('#btn-export').textContent = `Export workbook (${entries.length} application${entries.length === 1 ? '' : 's'})`;
  }
  $('#btn-new').addEventListener('click', openCapture);
  $('#btn-export').addEventListener('click', async () => {
    const all = entries.slice().sort((a, b) => (a.createdAt || '').localeCompare(b.createdAt || ''));
    try {
      const blob = await AAXlsx.buildWorkbook(all);
      await deliverFile(blob, `Additive Applications ${new Date().toISOString().slice(0, 10)}.xlsx`);
    } catch (e) { toast('The workbook couldn\'t be created: ' + e.message, true); }
  });

  async function deliverFile(blob, filename) {
    const file = new File([blob], filename, { type: blob.type });
    const mobile = window.matchMedia('(pointer: coarse)').matches;
    if (mobile && navigator.canShare && navigator.canShare({ files: [file] })) {
      try { await navigator.share({ files: [file], title: filename }); return; }
      catch (e) { if (e.name === 'AbortError') return; }
    }
    const a = el('a', { href: URL.createObjectURL(blob), download: filename });
    document.body.append(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 5000);
  }

  /* ================= capture ================= */
  let cap = { notebook: [], part: null };
  function renderChecklist() {
    const fill = (ol, list) => { ol.replaceChildren(...list.map(f => el('li', { text: f.label }))); };
    fill($('#list-required'), F.REQUIRED);
    fill($('#list-supporting'), F.SUPPORTING);
  }
  function openCapture() {
    if (!settings.name) { openSettings(); return; }
    cap = { notebook: [], part: null };
    renderCapture();
    $('#checklist').open = true;
    $('#read-status').hidden = true;
    show('capture', 'New application');
  }
  function renderCapture() {
    $('#thumbs-notebook').replaceChildren(...cap.notebook.map((p, i) => thumb(p,
      () => { cap.notebook.splice(i, 1); renderCapture(); },
      async () => { cap.notebook[i] = await rotatePhoto(cap.notebook[i], 90); renderCapture(); })));
    $('#thumbs-part').replaceChildren(...(cap.part ? [thumb(cap.part,
      () => { cap.part = null; renderCapture(); },
      async () => { cap.part = await rotatePhoto(cap.part, 90); renderCapture(); })] : []));
    const ai = settings.provider !== 'none';
    $('#btn-read').hidden = !ai;
    $('#btn-read').disabled = cap.notebook.length === 0;
    $('#btn-manual').textContent = ai ? 'Type it in instead' : 'Continue';
    $('#btn-manual').className = ai ? 'btn link' : 'btn primary big';
  }
  async function addPhotos(input, kind) {
    const files = Array.from(input.files || []);
    input.value = '';
    for (const f of files) {
      try {
        const p = await fileToPhoto(f, kind === 'part' ? PART_MAX : NOTEBOOK_MAX);
        if (kind === 'part') {
          cap.part = p;
          autoStraightenPart(() => cap.part, (np) => {
            const old = cap.part; cap.part = np;
            if (current === 'capture') renderCapture();
            // If the review screen already opened with this photo, update it there too.
            if (draft && draft.isNew && draft.entry.partPhoto && draft.entry.partPhoto.dataUrl === old.dataUrl) {
              draft.entry.partPhoto = np; renderFields(collectFields());
            }
          });
        } else cap.notebook.push(p);
      } catch (e) { toast(e.message, true); }
    }
    if (kind === 'notebook' && cap.notebook.length) $('#checklist').open = false;
    renderCapture();
  }
  $('#in-notebook-cam').addEventListener('change', e => addPhotos(e.target, 'notebook'));
  $('#in-notebook-pick').addEventListener('change', e => addPhotos(e.target, 'notebook'));
  $('#in-part-cam').addEventListener('change', e => addPhotos(e.target, 'part'));
  $('#in-part-pick').addEventListener('change', e => addPhotos(e.target, 'part'));

  $('#btn-manual').addEventListener('click', () => {
    openReview(newEntry(Object.fromEntries(F.KEYS.map(k => [k, ''])), ''), true, { warnings: [], uncertain: [] });
  });
  $('#btn-read').addEventListener('click', async () => {
    if (!settings.apiKey) { toast('Add your AI service API key in Settings first.', true); return; }
    const btn = $('#btn-read');
    btn.disabled = true; $('#read-status').hidden = false;
    $('#read-status-text').textContent = 'Reading your handwriting…';
    try {
      let result = await readWithAI(cap.notebook);
      // If a page was sideways or upside down, straighten it and read once more for best accuracy.
      const turns = cap.notebook.map((_, i) => result.rotateToFix[i] || 0);
      if (turns.some(Boolean)) {
        $('#read-status-text').textContent = 'The page was turned. Straightening it and reading again…';
        cap.notebook = await Promise.all(cap.notebook.map((p, i) => rotatePhoto(p, turns[i])));
        renderCapture();
        try { result = await readWithAI(cap.notebook); }
        catch (e) { /* keep the first reading if the second one fails */ }
      }
      openReview(newEntry(result.fields, result.extraText), true, result);
    } catch (e) {
      toast(e.message, true);
    } finally {
      btn.disabled = cap.notebook.length === 0; $('#read-status').hidden = true;
    }
  });
  function newEntry(fields, otherNotes) {
    const now = new Date().toISOString();
    return {
      id: uid(), createdAt: now, updatedAt: now, postedAt: null,
      originator: settings.name, workstreamOwner: settings.owner, notesFolder: settings.folder || firstName(settings.name),
      fields: Object.assign(Object.fromEntries(F.KEYS.map(k => [k, ''])), fields),
      otherNotes: otherNotes || '', notebookPhotos: cap.notebook.slice(), partPhoto: cap.part
    };
  }

  /* ================= AI ================= */
  async function readWithAI(photos) {
    return F.parseAIResponse(await askAI(photos, F.buildPrompt(), 3000));
  }
  async function askAI(photos, prompt, maxTokens) {
    const s = settings;
    let res, text;
    try {
      if (s.provider === 'anthropic') {
        res = await fetch('https://api.anthropic.com/v1/messages', {
          method: 'POST',
          headers: { 'content-type': 'application/json', 'x-api-key': s.apiKey, 'anthropic-version': '2023-06-01', 'anthropic-dangerous-direct-browser-access': 'true' },
          body: JSON.stringify({
            model: s.model || DEFAULT_MODELS.anthropic, max_tokens: maxTokens || 3000,
            messages: [{ role: 'user', content: photos.map(p => ({ type: 'image', source: { type: 'base64', media_type: 'image/jpeg', data: p.dataUrl.split(',')[1] } })).concat([{ type: 'text', text: prompt }]) }]
          })
        });
      } else {
        const body = {
          messages: [{ role: 'user', content: [{ type: 'text', text: prompt }].concat(photos.map(p => ({ type: 'image_url', image_url: { url: p.dataUrl, detail: 'high' } }))) }],
          response_format: { type: 'json_object' }
        };
        let url, headers = { 'content-type': 'application/json' };
        if (s.provider === 'azure') {
          if (!s.azureEndpoint || !s.azureDeployment) throw new Error('Add the Azure endpoint and deployment name in Settings.');
          url = `${s.azureEndpoint.replace(/\/+$/, '')}/openai/deployments/${encodeURIComponent(s.azureDeployment)}/chat/completions?api-version=2024-10-21`;
          headers['api-key'] = s.apiKey;
        } else {
          if (!s.model) throw new Error('Enter an OpenAI model name in Settings.');
          url = 'https://api.openai.com/v1/chat/completions';
          headers.Authorization = 'Bearer ' + s.apiKey;
          body.model = s.model;
        }
        res = await fetch(url, { method: 'POST', headers, body: JSON.stringify(body) });
      }
    } catch (e) {
      if (/Settings/.test(e.message)) throw e;
      throw new Error('Couldn\'t reach the AI service. Check your connection and try again.');
    }
    let j = null;
    try { j = await res.json(); } catch (e) { /* not json */ }
    if (!res.ok) {
      const msg = j && (j.error && (j.error.message || j.error.type));
      if (res.status === 401 || res.status === 403) throw new Error('The AI service rejected the API key. Check it in Settings.');
      throw new Error(`The AI service returned an error (${res.status})${msg ? ': ' + msg : '.'}`);
    }
    if (s.provider === 'anthropic') text = ((j && j.content) || []).filter(c => c.type === 'text').map(c => c.text).join('');
    else text = j && j.choices && j.choices[0] && j.choices[0].message && j.choices[0].message.content;
    return text;
  }

  /* Part photos: ask the AI which way is up and straighten automatically (small, low-cost check). */
  const aiReady = () => settings.provider !== 'none' && !!settings.apiKey;
  async function partPhotoTurn(photo) {
    const small = await toJpeg(photo.dataUrl, 768, 0.7);
    const text = await askAI([small], F.buildPartOrientationPrompt(), 200);
    return F.parseOrientation(text);
  }
  // Runs in the background after a part photo is added. `get`/`set` read and replace the photo
  // wherever it lives (capture screen or review screen), only if it hasn't changed meanwhile.
  async function autoStraightenPart(get, set) {
    const original = get();
    if (!original || !aiReady()) return;
    try {
      const deg = await partPhotoTurn(original);
      if (!deg || get() !== original) return;
      set(await rotatePhoto(original, deg));
      toast('Part photo straightened. Tap ⟳ if it\'s still not right.');
    } catch (e) { /* leave the photo as taken; the ⟳ button is still there */ }
  }

  /* ================= Teams ================= */
  async function sendToTeams(url, payload) {
    let res;
    try { res = await fetch(url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) }); }
    catch (e) { throw new Error('Couldn\'t reach Teams. Check your connection. If it keeps failing, your company may be blocking posts from outside apps.'); }
    if (!res.ok) throw new Error(`Teams didn't accept the message (error ${res.status}). Check the webhook URL in Settings.`);
  }
  async function postEntry(entry) {
    if (!settings.webhook) throw new Error('Add the Teams webhook URL in Settings to post.');
    // Teams rejects messages over ~28 KB, so keep the whole card under a safe budget.
    const BUDGET = 22000;
    let payload = F.buildTeamsPayload(entry);
    const textSize = JSON.stringify(payload).length;
    if (settings.thumb && entry.partPhoto && textSize < BUDGET - 3000) {
      const room = BUDGET - textSize;
      const tries = [[320, 0.6], [260, 0.55], [220, 0.5], [180, 0.45], [140, 0.4]];
      for (const [dim, q] of tries) {
        const t = (await toJpeg(entry.partPhoto.dataUrl, dim, q)).dataUrl;
        if (t.length <= room) { payload = F.buildTeamsPayload(entry, { thumbnailDataUrl: t }); break; }
      }
    }
    // Entry data for the workflow's SharePoint Excel steps (ignored by the Teams card step).
    let photo = null;
    // The workflow accepts at most 256 KB per request, so keep the Excel photo well under that.
    if (entry.partPhoto) {
      const tries = [[1200, 0.8], [1000, 0.75], [850, 0.7], [700, 0.65], [560, 0.6], [440, 0.55]];
      for (const [dim, q] of tries) {
        const p = await toJpeg(entry.partPhoto.dataUrl, dim, q);
        const b64 = p.dataUrl.split(',')[1];
        if (b64.length <= 150000 || dim === 440) { photo = { base64: b64, w: p.w, h: p.h }; break; }
      }
    }
    payload.entry = F.buildEntryData(entry, photo);
    await sendToTeams(settings.webhook, payload);
  }

  /* ================= review ================= */
  let draft = null; // { entry, isNew, flagged:Set, dirty }
  function openReview(entry, isNew, result) {
    result = result || { warnings: [], uncertain: [] };
    const v = F.validate(entry.fields);
    const warnings = [];
    (result.warnings || []).concat(v.warnings).forEach(w => { if (!warnings.includes(w)) warnings.push(w); });
    draft = { entry: JSON.parse(JSON.stringify(entry)), isNew, flagged: new Set((result.uncertain || []).concat(v.flagged)), dirty: false, warnings, raw: {} };
    if (!isNew) { draft.flagged = new Set(); draft.warnings = []; }
    ['urgency', 'confidence'].forEach(k => { draft.raw[k] = entry.fields[k] || ''; });
    $('#rv-originator').value = entry.originator || '';
    $('#rv-owner').value = entry.workstreamOwner || '';
    $('#rv-folder').value = entry.notesFolder || settings.folder || firstName(settings.name);
    $('#rv-other').value = entry.otherNotes || '';
    $('#review-pages').replaceChildren(...(entry.notebookPhotos || []).map(p => thumb(p)));
    $('#page-preview').hidden = !(entry.notebookPhotos || []).length;
    $('#page-preview').open = false;
    renderWarnings();
    renderFields(entry.fields);
    $('#btn-delete').hidden = isNew;
    $('#btn-save-post').textContent = entry.postedAt ? 'Save & post to Teams again' : 'Save & post to Teams';
    $('#posted-note').hidden = !entry.postedAt;
    if (entry.postedAt) $('#posted-note').textContent = 'Posted to Teams ' + fmtDate(entry.postedAt);
    show('review', isNew ? 'Check & save' : (F.clean(entry.fields.appName) || 'Application'), { noPush: current !== 'home' });
  }
  function renderWarnings() {
    const w = $('#warnings');
    if (!draft.warnings.length) { w.hidden = true; return; }
    w.replaceChildren(el('b', { text: 'Please check' }), el('ul', null, draft.warnings.map(t => el('li', { text: t }))));
    w.hidden = false;
  }
  function renderFields(fields) {
    [['required', F.REQUIRED, '#rv-required'], ['supporting', F.SUPPORTING, '#rv-supporting']].forEach(([sec, list, sel]) => {
      $(sel).replaceChildren(...list.map((f, i) => fieldRow(f, i, fields[f.key] || '')));
    });
  }
  function fieldRow(f, i, value) {
    const row = el('div', { class: 'field' + (draft.flagged.has(f.key) ? ' flag' : ''), 'data-key': f.key });
    const unflag = () => { row.classList.remove('flag'); draft.flagged.delete(f.key); draft.dirty = true; };
    row.append(el('div', { class: 'field-head' },
      el('span', { class: 'num', text: F.fieldNumber(f.key) }),
      el('span', { class: 'lbl', text: f.label }),
      el('button', { class: 'shift', type: 'button', title: 'A line was skipped above: move this and everything below down one', 'aria-label': 'Shift down from here', text: '↓', onclick: () => doShift(f.key, 'down') }),
      el('button', { class: 'shift', type: 'button', title: 'This is an extra line: remove it and move everything below up one', 'aria-label': 'Shift up into here', text: '↑', onclick: () => doShift(f.key, 'up') })));
    if (f.type === 'scale') {
      row.append(slider(f, value, unflag));
    } else if (f.type === 'photo') {
      const box = el('div', { class: 'part-photo' });
      const renderPhoto = () => {
        const p = draft.entry.partPhoto;
        box.replaceChildren(
          p ? thumb(p, () => { draft.entry.partPhoto = null; draft.dirty = true; renderPhoto(); },
            async () => { draft.entry.partPhoto = await rotatePhoto(draft.entry.partPhoto, 90); draft.dirty = true; renderPhoto(); })
            : el('span', { class: 'hint', text: 'No part photo' }),
          el('label', { class: 'btn secondary file-btn' }, p ? 'Replace' : '📷 Add part photo',
            el('input', { type: 'file', accept: 'image/*', hidden: true, onchange: async (ev) => {
              const file = ev.target.files && ev.target.files[0]; if (!file) return;
              try {
                draft.entry.partPhoto = await fileToPhoto(file, PART_MAX); draft.dirty = true; renderPhoto();
                const d = draft;
                autoStraightenPart(() => d.entry.partPhoto, (np) => { d.entry.partPhoto = np; if (draft === d) renderPhoto(); });
              } catch (e) { toast(e.message, true); }
            } })));
      };
      renderPhoto();
      row.append(box);
      row.append(el('input', { type: 'hidden', class: 'val', value }));
      if (value) row.append(el('div', { class: 'read-as', text: `On the page: ${value}` }));
    } else {
      const ta = el('textarea', { class: 'val', rows: '1' });
      ta.value = value;
      const grow = () => { ta.style.height = 'auto'; ta.style.height = (ta.scrollHeight + 2) + 'px'; };
      ta.addEventListener('input', () => { grow(); unflag(); });
      requestAnimationFrame(grow);
      row.append(ta);
    }
    return row;
  }
  function slider(f, raw, onChange) {
    const isU = f.key === 'urgency';
    const n = F.scaleValue(raw);
    const wrap = el('div', { class: 'slider-wrap' + (n ? '' : ' unset') });
    const val = el('span', { class: 'slider-val' });
    const marks = el('span', { class: 'slider-marks' });
    const input = el('input', { type: 'range', min: '1', max: '10', step: '1', class: (isU ? 'urgency' : 'confidence') + ' val-range', 'aria-label': f.label });
    input.value = String(n || 5);
    const hidden = el('input', { type: 'hidden', class: 'val', value: n ? String(n) : '' });
    const paint = (v, set) => {
      if (set) {
        val.textContent = v + ' / 10';
        if (isU) { val.style.color = urgColor(v); input.style.setProperty('--thumb', urgColor(v)); }
        else { marks.textContent = F.confidenceMarks(v); input.style.setProperty('--thumb', 'var(--accent-text)'); }
      } else { val.textContent = 'Not set: slide to set'; val.style.color = ''; marks.textContent = ''; }
    };
    paint(n || 5, !!n);
    const set = () => { const v = parseInt(input.value, 10); hidden.value = String(v); wrap.classList.remove('unset'); paint(v, true); onChange(); };
    input.addEventListener('input', set);
    input.addEventListener('change', set);
    wrap.append(el('div', { class: 'slider-top' }, val, marks), input,
      el('div', { class: 'slider-ends' }, el('span', { text: isU ? '1 · low' : '1  !' }), el('span', { text: isU ? 'high · 10' : '!?!?!?!?  10' })), hidden);
    const frag = el('div', null, wrap);
    const r = (draft.raw[f.key] || '').trim();
    if (r && !F.scaleValue(r)) frag.append(el('div', { class: 'read-as', text: `Read from the page as: "${r}"` }));
    return frag;
  }
  function collectFields() {
    const out = {};
    document.querySelectorAll('#screen-review .field').forEach(row => {
      const v = row.querySelector('.val');
      out[row.dataset.key] = v ? v.value.trim() : '';
    });
    return out;
  }
  function doShift(key, dir) {
    const fields = collectFields();
    let next;
    if (dir === 'down') next = F.shiftDown(fields, key);
    else {
      const r = F.shiftUp(fields, key); next = r.fields;
      if (r.removed) { const o = $('#rv-other'); o.value = [o.value.trim(), r.removed].filter(Boolean).join('\n'); }
    }
    // scales: keep any non-number text visible so it can be moved again
    ['urgency', 'confidence'].forEach(k => { draft.raw[k] = next[k] || ''; });
    const v = F.validate(next);
    draft.flagged = new Set(v.flagged);
    draft.dirty = true;
    renderFields(next);
    toast(dir === 'down' ? 'Moved values down one slot' : 'Moved values up one slot');
  }

  function entryFromForm() {
    const e = draft.entry;
    e.fields = collectFields();
    e.originator = $('#rv-originator').value.trim();
    e.workstreamOwner = $('#rv-owner').value.trim();
    e.notesFolder = $('#rv-folder').value.trim();
    e.otherNotes = $('#rv-other').value.trim();
    e.updatedAt = new Date().toISOString();
    return e;
  }
  async function save(post) {
    const e = entryFromForm();
    if (!F.clean(e.fields.appName)) { toast('Add an Application Name before saving.', true); return; }
    const missing = ['urgency', 'confidence'].filter(k => !F.scaleValue(e.fields[k])).map(k => F.byKey[k].label.replace(' (1-10)', ''));
    if (missing.length && !confirm(`${missing.join(' and ')} ${missing.length > 1 ? 'aren\'t' : 'isn\'t'} set. Save anyway?`)) return;
    if (post && e.postedAt && !confirm('This was already posted to Teams. Post it again?')) return;
    const buttons = ['#btn-save-post', '#btn-save'].map(s => $(s));
    buttons.forEach(b => { b.disabled = true; });
    try {
      await DB.put(e);
      draft.isNew = false; draft.dirty = false;
      if (post) {
        try { await postEntry(e); e.postedAt = new Date().toISOString(); await DB.put(e); toast('Saved and posted to Teams'); }
        catch (err) { toast('Saved, but not posted: ' + err.message, true); openReview(e, false); return; }
      } else toast('Saved');
      goHome();
    } catch (err) {
      toast('Couldn\'t save on this device: ' + err.message, true);
    } finally { buttons.forEach(b => { b.disabled = false; }); }
  }
  $('#btn-save-post').addEventListener('click', () => save(true));
  $('#btn-save').addEventListener('click', () => save(false));
  $('#btn-export-one').addEventListener('click', async () => {
    const e = entryFromForm();
    try { const blob = await AAXlsx.buildWorkbook([e], { summary: false }); await deliverFile(blob, `${(F.clean(e.fields.appName) || 'Additive Application').replace(/[\\/:*?"<>|]/g, ' ').slice(0, 60)}.xlsx`); }
    catch (err) { toast('The form couldn\'t be created: ' + err.message, true); }
  });
  $('#btn-delete').addEventListener('click', async () => {
    if (!confirm('Delete this entry from this phone? Anything already posted to Teams stays there.')) return;
    try { await DB.del(draft.entry.id); toast('Deleted'); goHome(); } catch (e) { toast('Couldn\'t delete: ' + e.message, true); }
  });
  ['#rv-originator', '#rv-owner', '#rv-folder', '#rv-other'].forEach(s => $(s).addEventListener('input', () => { if (draft) draft.dirty = true; }));

  /* ================= start ================= */
  renderChecklist();
  if (!settings.name) openSettings(); else goHome();
  if ('serviceWorker' in navigator && location.protocol !== 'file:') {
    navigator.serviceWorker.register('sw.js').catch(() => { /* offline support unavailable */ });
  }
})();
