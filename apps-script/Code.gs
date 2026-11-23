/**
 * ทริปด้วยกัน — backend สำหรับเวอร์ชัน GitHub Pages
 * เก็บข้อมูลใน Google Sheet นี้ รูป/ไฟล์ใน Google Drive และใช้ Google Maps คำนวณระยะทาง/เวลา
 *
 * ติดตั้งครั้งเดียว
 * 1) สร้าง Google Sheet ใหม่ (ว่าง ๆ) > ส่วนขยาย > Apps Script
 * 2) ลบโค้ดเดิม แล้ววางไฟล์นี้ทั้งหมดลงใน Code.gs > กดบันทึก
 * 3) เลือกฟังก์ชัน setup > Run > อนุญาตสิทธิ์ทั้งหมด
 *    (ถ้าขึ้น "Google hasn't verified this app" ให้กด Advanced > Go to ... เพราะเป็นสคริปต์ของคุณเอง)
 * 4) Deploy > New deployment > เลือกชนิด Web app
 *      Execute as: Me
 *      Who has access: Anyone
 *    กด Deploy แล้วคัดลอก Web app URL (ลงท้ายด้วย /exec)
 * 5) เปิดหน้าแอปบน GitHub Pages แล้ววาง URL นี้ในหน้าตั้งค่าครั้งแรก
 *    (หรือใส่ใน config.js ของ repo เพื่อให้ทุกคนใช้ได้ทันที)
 *
 * ถ้าแก้โค้ดนี้ภายหลัง: Deploy > Manage deployments > Edit > Version: New version > Deploy
 */

const APP = 'ทริปด้วยกัน';
const SHEET = 'docs';
const MAX_CELL = 49000;

/* ================= entry points ================= */
function doGet() {
  return out_({ ok: true, app: 'tripitinerary' });
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (x) { return out_({ error: 'bad request' }); }
  try { return out_({ ok: true, result: handle_(req) }); }
  catch (x) { return out_({ error: String((x && x.message) || x) }); }
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** รันครั้งแรก: สร้างชีต + โฟลเดอร์ + ขอสิทธิ์ Maps/Drive/UrlFetch */
function setup() {
  sheet_();
  rootFolder_();
  Maps.newGeocoder().geocode('Bangkok');
  UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true });
  Logger.log('ติดตั้งเรียบร้อย ไปที่ Deploy > New deployment > Web app ได้เลย');
}

function handle_(req) {
  switch (req.action) {
    case 'pull': return pull_(req.targets || []);
    case 'write': return withLock_(() => write_(req.ops || []));
    case 'upload': return upload_(req);
    case 'deleteFile': return deleteFile_(req.id);
    case 'file': return readFile_(req.id);
    case 'route': return route_(req.pairs || [], req.mode || 'driving');
    case 'resolve': return resolve_(req.url);
    default: throw new Error('unknown action');
  }
}

/* ================= document store (one row per document) ================= */
function sheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.getRange(1, 1, 1, 4).setValues([['path', 'collection', 'json', 'updatedAt']]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A:C').setNumberFormat('@');
    ss.getSheets().forEach(s => { if (s.getName() !== SHEET && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s); });
  }
  return sh;
}

function load_() {
  const sh = sheet_(), n = sh.getLastRow(), map = {};
  if (n < 2) return map;
  sh.getRange(2, 1, n - 1, 3).getValues().forEach((r, i) => {
    if (!r[0]) return;
    let data = null;
    try { data = JSON.parse(r[2]); } catch (x) { data = null; }
    map[String(r[0])] = { row: i + 2, col: String(r[1]), data: data };
  });
  return map;
}

const parentOf_ = p => p.split('/').slice(0, -1).join('/');
const idOf_ = p => p.split('/').pop();

function match_(data, where) {
  return (where || []).every(w => {
    const v = data ? data[w[0]] : undefined, op = w[1], x = w[2];
    if (op === '==') return v === x;
    if (op === '!=') return v !== x;
    if (op === 'array-contains') return Array.isArray(v) && v.indexOf(x) >= 0;
    if (op === 'in') return Array.isArray(x) && x.indexOf(v) >= 0;
    return true;
  });
}

function pull_(targets) {
  const map = load_();
  return targets.slice(0, 40).map(t => {
    if (t.t === 'doc') {
      const d = map[t.path];
      return { exists: !!d, data: d ? d.data : null };
    }
    let docs = Object.keys(map).filter(p => map[p].col === t.path && match_(map[p].data, t.where))
      .sort().map(p => ({ id: idOf_(p), data: map[p].data }));
    if (t.limit) docs = docs.slice(0, t.limit);
    return { docs: docs };
  });
}

function merge_(a, b) {
  const out = Object.assign({}, a || {});
  Object.keys(b || {}).forEach(k => {
    const v = b[k];
    if (v && typeof v === 'object' && !Array.isArray(v) && out[k] && typeof out[k] === 'object' && !Array.isArray(out[k])) out[k] = merge_(out[k], v);
    else out[k] = v;
  });
  return out;
}

function write_(ops) {
  const sh = sheet_(), map = load_(), now = Date.now();
  const del = [];
  ops.forEach(o => {
    const path = String(o.path || '');
    if (!/^[A-Za-z0-9_\-.~:@+]+(\/[A-Za-z0-9_\-.~:@+]+)*$/.test(path) || path.split('/').length % 2) throw new Error('bad path');
    const cur = map[path];
    if (o.op === 'delete') { if (cur) del.push(cur.row); delete map[path]; return; }
    let data;
    if (o.op === 'update') { if (!cur) throw new Error('not found'); data = merge_(cur.data, o.data); }
    else data = o.data || {};
    const json = JSON.stringify(data);
    if (json.length > MAX_CELL) throw new Error('ข้อมูลรายการนี้ใหญ่เกินไป');
    const row = [path, parentOf_(path), json, now];
    if (cur) { sh.getRange(cur.row, 1, 1, 4).setValues([row]); cur.data = data; }
    else { sh.appendRow(row); map[path] = { row: sh.getLastRow(), col: parentOf_(path), data: data }; }
  });
  del.sort((a, b) => b - a).forEach(r => sh.deleteRow(r));
  return true;
}

function withLock_(fn) {
  const l = LockService.getScriptLock();
  l.waitLock(25000);
  try { return fn(); } finally { SpreadsheetApp.flush(); l.releaseLock(); }
}

/* ================= files (Google Drive) ================= */
function rootFolder_() {
  const props = PropertiesService.getScriptProperties(), id = props.getProperty('ROOT_FOLDER');
  if (id) { try { const f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (x) {} }
  const f = DriveApp.createFolder(APP + ' - ไฟล์แนบ');
  props.setProperty('ROOT_FOLDER', f.getId());
  return f;
}

function upload_(req) {
  const type = String(req.type || 'application/octet-stream');
  const blob = Utilities.newBlob(Utilities.base64Decode(req.b64), type, String(req.name || 'file').slice(0, 120));
  const f = rootFolder_().createFile(blob);
  try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (x) {}
  return { id: f.getId(), contentType: type, sizeBytes: f.getSize() };
}

function inRoot_(file) {
  const rootId = rootFolder_().getId(), it = file.getParents();
  while (it.hasNext()) if (it.next().getId() === rootId) return true;
  return false;
}

function deleteFile_(id) {
  try { const f = DriveApp.getFileById(id); if (inRoot_(f)) f.setTrashed(true); } catch (x) {}
  return { deleted: true };
}

function readFile_(id) {
  const f = DriveApp.getFileById(id);
  if (!inRoot_(f)) throw new Error('not found');
  const b = f.getBlob();
  return { type: b.getContentType(), b64: Utilities.base64Encode(b.getBytes()) };
}

/* ================= Google Maps ================= */
const MODES_ = {
  driving: Maps.DirectionFinder.Mode.DRIVING,
  walking: Maps.DirectionFinder.Mode.WALKING,
  transit: Maps.DirectionFinder.Mode.TRANSIT
};

function point_(df, q, origin) {
  const m = String(q).match(/^\s*(-?\d+(?:\.\d+)?)\s*,\s*(-?\d+(?:\.\d+)?)\s*$/);
  if (origin) return m ? df.setOrigin(+m[1], +m[2]) : df.setOrigin(String(q));
  return m ? df.setDestination(+m[1], +m[2]) : df.setDestination(String(q));
}

/** ระยะทาง (เมตร) และเวลา (วินาที) ของแต่ละคู่จุด: pairs = [[from, to], ...] เป็นพิกัด "lat,lng" หรือชื่อสถานที่ */
function route_(pairs, mode) {
  const cache = CacheService.getScriptCache();
  return pairs.slice(0, 25).map(p => {
    const key = 'L_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, mode + '|' + p[0] + '|' + p[1]));
    const hit = cache.get(key);
    if (hit) return JSON.parse(hit);
    let out;
    try {
      let df = Maps.newDirectionFinder().setMode(MODES_[mode] || MODES_.driving).setLanguage('th');
      df = point_(df, p[0], true);
      df = point_(df, p[1], false);
      const d = df.getDirections();
      if (d.routes && d.routes.length) { const L = d.routes[0].legs[0]; out = { ok: true, m: L.distance.value, s: L.duration.value }; }
      else out = { ok: false };
    } catch (x) { out = { ok: false }; }
    try { cache.put(key, JSON.stringify(out), 21600); } catch (x) {}
    return out;
  });
}

function parseLoc_(s) {
  s = String(s || '');
  const m = s.match(/!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)/) ||
    s.match(/[?&](?:q|query|ll|destination|daddr|center)=(-?\d+\.\d+)(?:,|%2C)\s*\+?(-?\d+\.\d+)/i) ||
    s.match(/@(-?\d+\.\d+),(-?\d+\.\d+)/);
  return m ? { lat: +m[1], lng: +m[2] } : null;
}

/** อ่านพิกัดจากลิงก์ Google Maps รวมถึงลิงก์ย่อ maps.app.goo.gl */
function resolve_(url) {
  let u = String(url || '').trim(), name = '';
  if (!/^https?:\/\//i.test(u)) return null;
  for (let i = 0; i < 6; i++) {
    const pm = u.match(/\/place\/([^/@?]+)/);
    if (pm && !name) { try { name = decodeURIComponent(pm[1].replace(/\+/g, ' ')); } catch (x) {} }
    const g = parseLoc_(u);
    if (g) { g.name = name; return g; }
    let r;
    try { r = UrlFetchApp.fetch(u, { followRedirects: false, muteHttpExceptions: true }); } catch (x) { break; }
    const h = r.getHeaders(), loc = h.Location || h.location;
    if (loc) { u = Array.isArray(loc) ? loc[0] : loc; continue; }
    const body = r.getContentText().slice(0, 300000);
    const g2 = parseLoc_((body.match(/https:\/\/www\.google\.[^"'\s]*\/maps[^"'\s]*/) || [''])[0]);
    if (g2) { g2.name = name; return g2; }
    break;
  }
  const q = name || ((u.match(/[?&]q=([^&]+)/) || [])[1] ? decodeURIComponent(u.match(/[?&]q=([^&]+)/)[1].replace(/\+/g, ' ')) : '');
  if (q) {
    const res = Maps.newGeocoder().setLanguage('th').geocode(q);
    if (res.results && res.results.length) {
      const l = res.results[0].geometry.location;
      return { lat: l.lat, lng: l.lng, name: name || q };
    }
  }
  return null;
}
