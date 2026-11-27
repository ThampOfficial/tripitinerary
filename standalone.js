/* ทริปด้วยกัน — ตัวเชื่อมสำหรับเวอร์ชัน GitHub Pages
 * ทำให้แอปตัวเดียวกับใน Claude ใช้ Google Apps Script (Google Sheet + Drive + Maps) แทนบริการของ Claude
 * ฐานข้อมูล: เก็บเป็นเอกสารใน Sheet อัปเดตจากเพื่อนทุก ~8 วินาที
 */
(function () {
  'use strict';
  window.TRIP_STANDALONE = true;

  const store = {
    get(k) { try { return localStorage.getItem(k); } catch (e) { return null; } },
    set(k, v) { try { localStorage.setItem(k, v); } catch (e) {} },
    del(k) { try { localStorage.removeItem(k); } catch (e) {} }
  };
  const API = () => (window.TRIP_API || store.get('tripApi') || '').trim();
  const API_OK = u => /^https:\/\/script\.google\.com\/.+\/exec$/.test(u || '');
  // invite links carry the backend URL, so friends who open one are connected without any setup
  try {
    const u = new URL(location.href), a = u.searchParams.get('api');
    if (a && API_OK(a)) { store.set('tripApi', a); u.searchParams.delete('api'); history.replaceState(null, '', u.pathname + u.search + u.hash); }
  } catch (e) {}
  window.TRIP_INVITE_URL = code => location.origin + location.pathname + '?api=' + encodeURIComponent(API()) + '#join-' + code;
  const rid = () => Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-4);

  async function call(action, payload) {
    let res;
    try {
      res = await fetch(API(), {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify(Object.assign({ action }, payload || {}))
      });
    } catch (e) { throw { code: 'unavailable', message: 'เชื่อมต่อ Google Sheet ไม่ได้' }; }
    let j;
    try { j = await res.json(); } catch (e) { throw { code: 'unavailable', message: 'Apps Script ตอบกลับไม่ถูกต้อง ตรวจว่า deploy เป็น Web app ให้ Anyone ใช้ได้' }; }
    if (j.error) throw { code: j.error === 'not found' ? 'invalid_argument' : 'unavailable', message: j.error };
    return j.result;
  }

  /* ---------- db: same surface the app uses in Claude ---------- */
  const meta = { fromCache: false, hasPendingWrites: false };
  const snapDoc = (path, d) => ({ id: path.split('/').pop(), exists: !!d, data: () => d || undefined, metadata: meta });
  const snapCol = list => {
    const docs = list.map(x => ({ id: x.id, exists: true, data: () => x.data, metadata: meta }));
    return { docs, size: docs.length, empty: !docs.length, docChanges: () => [], metadata: meta };
  };
  const subs = new Map();
  let sid = 0, timer = null, pulling = false, again = false;

  async function pull() {
    if (!subs.size) return;
    if (pulling) { again = true; return; }
    pulling = true;
    const list = [...subs.entries()];
    try {
      const res = await call('pull', { targets: list.map(([, s]) => s.target) });
      res.forEach((r, i) => {
        const [id, s] = list[i];
        if (!subs.has(id)) return;
        const j = JSON.stringify(r);
        if (j === s.last) return;
        s.last = j;
        try { s.next(s.target.t === 'doc' ? snapDoc(s.target.path, r.exists ? r.data : null) : snapCol(r.docs)); } catch (e) { console.error(e); }
      });
    } catch (e) { console.warn('sync', e); }
    pulling = false;
    if (again) { again = false; pull(); }
  }
  function schedule(ms) {
    clearTimeout(timer);
    timer = setTimeout(async () => { await pull(); schedule(document.hidden ? 30000 : 8000); }, ms);
  }
  document.addEventListener('visibilitychange', () => { if (!document.hidden) schedule(100); });
  function subscribe(target, next) {
    const id = ++sid;
    subs.set(id, { target, next, last: null });
    schedule(30);
    return () => subs.delete(id);
  }
  async function write(op, path, data) {
    await call('write', { ops: [{ op, path, data: data === undefined ? null : JSON.parse(JSON.stringify(data)) }] });
    schedule(20);
  }
  function docRef(path) {
    return {
      id: path.split('/').pop(), path,
      get: async () => { const [r] = await call('pull', { targets: [{ t: 'doc', path }] }); return snapDoc(path, r.exists ? r.data : null); },
      set: d => write('set', path, d),
      update: d => write('update', path, d),
      delete: () => write('delete', path),
      onSnapshot: next => subscribe({ t: 'doc', path }, next),
      collection: c => query(path + '/' + c, [], 0)
    };
  }
  function query(path, where, lim) {
    return {
      path,
      where: (f, op, v) => query(path, where.concat([[f, op, v]]), lim),
      orderBy() { return this; },
      limit: n => query(path, where, n),
      get: async () => { const [r] = await call('pull', { targets: [{ t: 'col', path, where, limit: lim }] }); return snapCol(r.docs); },
      onSnapshot: next => subscribe({ t: 'col', path, where, limit: lim }, next),
      doc: id => docRef(path + '/' + (id || rid())),
      add: async d => { const r = docRef(path + '/' + rid()); await r.set(d); return r; }
    };
  }
  const db = { doc: docRef, collection: p => query(p, [], 0) };

  /* ---------- who is using this device ---------- */
  let uid = store.get('tripUid');
  if (!uid) { uid = 'u_' + rid() + rid(); store.set('tripUid', uid); }
  const user = {
    id: async () => uid,
    me: async () => ({ id: uid, name: store.get('tripMyName') || '', avatarUrl: '', color: '#0E7A6B', email: null, isOwner: true, canEdit: true }),
    isOwner: async () => true,
    canEdit: async () => true,
    can: async () => true,
    profiles: async ids => Object.fromEntries([].concat(ids).map(i => [i, { id: i, name: '', avatarUrl: '', color: '#0E7A6B', email: null, isMe: i === uid, guest: false }])),
    search: async () => []
  };

  /* ---------- files go to Google Drive ---------- */
  const readB64 = blob => new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(String(r.result).split(',')[1]); r.onerror = rej; r.readAsDataURL(blob); });
  const loadImg = file => new Promise((res, rej) => { const u = URL.createObjectURL(file); const im = new Image(); im.onload = () => { res(im); URL.revokeObjectURL(u); }; im.onerror = rej; im.src = u; });
  async function shrink(file) {
    if (!/^image\/(jpeg|png|webp)$/.test(file.type) || file.size < 1.5e6) return file;
    try {
      const im = await loadImg(file), sc = Math.min(1, 2000 / Math.max(im.naturalWidth, im.naturalHeight));
      const c = document.createElement('canvas'); c.width = Math.round(im.naturalWidth * sc); c.height = Math.round(im.naturalHeight * sc);
      c.getContext('2d').drawImage(im, 0, 0, c.width, c.height);
      const b = await new Promise(r => c.toBlob(r, 'image/jpeg', 0.85));
      return b || file;
    } catch (e) { return file; }
  }
  const assets = {
    async upload(file, opt) {
      const small = await shrink(file);
      if (small.size > 25 * 1024 * 1024) throw { code: 'too_large' };
      const type = small === file ? ((opt && opt.type) || file.type || 'application/octet-stream') : 'image/jpeg';
      const r = await call('upload', { name: file.name || 'file', type, b64: await readB64(small) });
      return { id: r.id, url: window.TRIP_BLOB_URL(r.id), sizeBytes: r.sizeBytes, contentType: r.contentType };
    },
    async delete(id) { return call('deleteFile', { id }); },
    async list() { return { assets: [], usage: {} }; }
  };
  window.TRIP_BLOB_URL = id => 'https://drive.google.com/thumbnail?id=' + encodeURIComponent(id) + '&sz=w1600';
  window.TRIP_FILE_URL = id => 'https://drive.google.com/file/d/' + encodeURIComponent(id) + '/view';
  window.TRIP_FETCH_DATAURL = async id => { const r = await call('file', { id }); return 'data:' + r.type + ';base64,' + r.b64; };

  /* ---------- Google Maps through Apps Script ---------- */
  window.TRIP_ROUTE = (pairs, mode) => call('route', { pairs, mode });
  window.TRIP_RESOLVE = url => call('resolve', { url });

  /* ---------- saving a generated file (PDF) ---------- */
  const downloads = {
    async save({ filename, data }) {
      const blob = data instanceof Blob ? data : new Blob([data]);
      const u = URL.createObjectURL(blob), a = document.createElement('a');
      a.href = u; a.download = filename; document.body.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(u), 60000);
      return { status: 'saved' };
    }
  };

  /* ---------- first run: ask for the Apps Script URL ---------- */
  function setupScreen(msg) {
    const v = document.getElementById('view');
    if (!v) return;
    v.innerHTML = `<div class="card pad stack" style="margin-top:8px">
      <div><span class="eyebrow">ตั้งค่าครั้งแรก</span><h2 style="font-size:22px">เชื่อมกับ Google Sheet</h2></div>
      <p style="margin:0"><b>ถ้าเพื่อนชวนมา</b> ให้เปิดจากลิงก์เชิญที่เพื่อนส่งให้ จะเข้าได้เลยโดยไม่ต้องตั้งค่า</p>
      <p style="margin:0">ถ้าคุณเป็นเจ้าของ วาง Web app URL ของ Apps Script (ลงท้ายด้วย <span class="num">/exec</span>) ที่ได้จากขั้นตอน Deploy ใน README</p>
      <div class="field"><label for="apiUrl">Web app URL</label><input class="inp num" id="apiUrl" placeholder="https://script.google.com/macros/s/…/exec" value="${API().replace(/"/g, '&quot;')}"></div>
      <div class="hint err" id="apiMsg">${msg || ''}</div>
      <div class="row" style="justify-content:flex-end"><button class="btn pri" id="apiSave">เชื่อมต่อ</button></div></div>`;
    document.getElementById('apiSave').onclick = async () => {
      const u = document.getElementById('apiUrl').value.trim(), m = document.getElementById('apiMsg');
      if (!API_OK(u)) { m.textContent = 'URL ต้องขึ้นต้นด้วย https://script.google.com/ และลงท้ายด้วย /exec'; return; }
      m.className = 'hint'; m.textContent = 'กำลังทดสอบ…';
      try {
        const r = await fetch(u); const j = await r.json();
        if (!j.ok) throw 0;
        store.set('tripApi', u); location.reload();
      } catch (e) { m.className = 'hint err'; m.textContent = 'เชื่อมต่อไม่ได้ ตรวจว่า deploy เป็น Web app และตั้ง Who has access เป็น Anyone'; }
    };
  }
  window.TRIP_RESET_API = () => { store.del('tripApi'); location.reload(); };

  const caps = { db, user, assets, downloads };
  if (!API()) {
    window.claude = { use: () => new Promise(() => {}) };
    document.addEventListener('DOMContentLoaded', () => setupScreen(''));
  } else {
    window.claude = { use: async name => caps[name] || null };
  }
})();
