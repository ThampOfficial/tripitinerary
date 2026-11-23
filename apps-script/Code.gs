/**
 * ทริปด้วยกัน — Google Apps Script backend
 * ------------------------------------------------------------
 * วิธีติดตั้ง (ทำครั้งเดียว)
 * 1) สร้าง Google Sheet ใหม่ (ว่าง ๆ) > ส่วนขยาย > Apps Script
 * 2) วางไฟล์นี้ใน Code.gs และสร้างไฟล์ HTML ชื่อ Index แล้ววางโค้ด Index.html
 * 3) เลือกฟังก์ชัน setup แล้วกด Run > อนุญาตสิทธิ์ (Sheets, Drive, Maps, เชื่อมต่อภายนอก)
 * 4) Deploy > New deployment > Web app
 *      Execute as: Me   |   Who has access: Anyone
 * 5) เปิด Web app URL ที่ได้ แล้วส่งลิงก์ให้เพื่อน
 * ------------------------------------------------------------
 */

const APP_NAME = 'ทริปด้วยกัน';

const TABLES = {
  Trips:    ['id','code','name','start','end','mode','fundHolder','people','placeCats','expCats','dayModes','createdBy','createdAt','updatedAt'],
  Places:   ['id','tripId','name','cat','day','time','stay','order','lat','lng','url','address','note','files','createdAt','updatedAt'],
  Expenses: ['id','tripId','title','amount','type','cat','date','paidBy','split','among','custom','note','files','createdAt','updatedAt'],
  Deposits: ['id','tripId','personId','amount','date','note','createdAt','updatedAt'],
  Docs:     ['id','tripId','title','note','files','createdAt','updatedAt']
};
const CHILD = ['Places','Expenses','Deposits','Docs'];
const JSON_COLS = {people:1, placeCats:1, expCats:1, dayModes:1, files:1, among:1, custom:1};
const OBJ_COLS  = {custom:1, dayModes:1};
const NUM_COLS  = {amount:1, order:1, stay:1, lat:1, lng:1, createdAt:1, updatedAt:1};
const TRIP_PATCH_KEYS = ['name','start','end','mode','fundHolder','people','placeCats','expCats','dayModes'];

/* ================= Web app ================= */
function doGet(e) {
  const t = HtmlService.createTemplateFromFile('Index');
  t.join = (e && e.parameter && e.parameter.join) || '';
  t.appUrl = ScriptApp.getService().getUrl();
  return t.evaluate()
    .setTitle(APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** รันครั้งแรกครั้งเดียว: สร้างชีตทั้งหมด + ขอสิทธิ์ */
function setup() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(TABLES).forEach(n => sheet_(n));
  ss.getSheets().forEach(sh => {
    if (!TABLES[sh.getName()] && sh.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });
  rootFolder_();
  Maps.newGeocoder().geocode('Bangkok');
  UrlFetchApp.fetch('https://www.google.com', {muteHttpExceptions: true});
  Logger.log('ติดตั้งเรียบร้อย ไปที่ Deploy > New deployment > Web app ได้เลย');
}

/* ================= Sheet helpers ================= */
function sheet_(name) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  const H = TABLES[name];
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, H.length).setValues([H]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange(1, 1, sh.getMaxRows(), H.length).setNumberFormat('@');
  }
  return sh;
}
function fmtDate_(d, p) { return Utilities.formatDate(d, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), p); }
function cell_(h, v) {
  if (v instanceof Date) return h === 'time' ? fmtDate_(v, 'HH:mm') : fmtDate_(v, 'yyyy-MM-dd');
  if (JSON_COLS[h]) {
    const empty = OBJ_COLS[h] ? {} : [];
    if (v === '' || v == null) return empty;
    try { return JSON.parse(v); } catch (e) { return empty; }
  }
  if (NUM_COLS[h]) {
    if (v === '' || v == null) return (h === 'lat' || h === 'lng') ? null : 0;
    const n = Number(v);
    return isNaN(n) ? null : n;
  }
  return v == null ? '' : String(v);
}
function rows_(name) {
  const sh = sheet_(name), n = sh.getLastRow(), H = TABLES[name];
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, H.length).getValues()
    .map((r, i) => { const o = {_row: i + 2}; H.forEach((h, j) => o[h] = cell_(h, r[j])); return o; })
    .filter(o => o.id);
}
function toRow_(name, o) {
  return TABLES[name].map(h => {
    const v = o[h];
    if (v == null) return '';
    if (JSON_COLS[h]) return JSON.stringify(v);
    return NUM_COLS[h] ? v : String(v);
  });
}
function findRow_(name, id) {
  const sh = sheet_(name);
  if (sh.getLastRow() < 2) return 0;
  const hit = sh.getRange(2, 1, sh.getLastRow() - 1, 1).createTextFinder(id).matchEntireCell(true).findNext();
  return hit ? hit.getRow() : 0;
}
function getById_(name, id) {
  const r = findRow_(name, id);
  if (!r) return null;
  const H = TABLES[name], vals = sheet_(name).getRange(r, 1, 1, H.length).getValues()[0];
  const o = {_row: r}; H.forEach((h, j) => o[h] = cell_(h, vals[j]));
  return o;
}
/** merge-upsert: ฟิลด์ที่ไม่ได้ส่งมาจะคงค่าเดิม */
function upsert_(name, obj) {
  const sh = sheet_(name), old = getById_(name, obj.id);
  const merged = Object.assign({}, old || {}, obj);
  delete merged._row;
  const row = toRow_(name, merged);
  if (old) sh.getRange(old._row, 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);
  return merged;
}
function remove_(name, id) { const r = findRow_(name, id); if (r) sheet_(name).deleteRow(r); }
function clean_(o) { if (!o) return o; const c = Object.assign({}, o); delete c._row; return c; }
function uid_(p) { return (p || '') + Utilities.getUuid().replace(/-/g, '').slice(0, 16); }
function J(x) { return JSON.stringify(x); }
function withLock_(fn) {
  const l = LockService.getScriptLock();
  l.waitLock(25000);
  try { return fn(); } finally { l.releaseLock(); SpreadsheetApp.flush(); }
}
function requireTrip_(id) {
  const t = getById_('Trips', String(id || ''));
  if (!t) throw new Error('ไม่พบทริปนี้');
  return t;
}
function tripData_(id) {
  const trip = getById_('Trips', id);
  if (!trip) return null;
  const out = {trip: clean_(trip)};
  CHILD.forEach(n => out[n.toLowerCase()] = rows_(n).filter(r => r.tripId === id).map(clean_));
  return out;
}
function newCode_() {
  const a = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789', used = {};
  rows_('Trips').forEach(t => used[t.code] = 1);
  for (let k = 0; k < 50; k++) {
    let s = ''; for (let i = 0; i < 6; i++) s += a[Math.floor(Math.random() * a.length)];
    if (!used[s]) return s;
  }
  throw new Error('สร้างรหัสไม่สำเร็จ');
}

/* ================= Trip API ================= */
function api_listTrips(idsJson) {
  const ids = {}; JSON.parse(idsJson || '[]').forEach(i => ids[i] = 1);
  const counts = {};
  rows_('Places').forEach(p => { if (ids[p.tripId]) counts[p.tripId] = (counts[p.tripId] || 0) + 1; });
  return J(rows_('Trips').filter(t => ids[t.id]).map(t => {
    const c = clean_(t); c.placeCount = counts[t.id] || 0; return c;
  }));
}
function api_getTrip(id) { return J(tripData_(String(id))); }

function api_createTrip(json) {
  const o = JSON.parse(json);
  return withLock_(() => {
    const now = Date.now();
    const t = {
      id: uid_('t'), code: newCode_(), name: o.name || 'ทริปใหม่', start: o.start || '', end: o.end || o.start || '',
      mode: 'driving', fundHolder: '', people: o.people || [], placeCats: o.placeCats || [], expCats: o.expCats || [],
      dayModes: {}, createdBy: o.createdBy || '', createdAt: now, updatedAt: now
    };
    upsert_('Trips', t);
    return J(tripData_(t.id));
  });
}
function api_findByCode(code) {
  code = String(code || '').trim().toUpperCase();
  const t = rows_('Trips').find(x => String(x.code).toUpperCase() === code);
  return J(t ? {id: t.id, name: t.name, start: t.start, end: t.end, people: t.people} : null);
}
function api_join(tripId, personId, newName) {
  return withLock_(() => {
    const t = requireTrip_(tripId), ps = t.people || [];
    let p = personId && ps.find(x => x.id === personId);
    if (p) p.joined = true;
    else { p = {id: uid_('p'), name: String(newName || 'เพื่อน').slice(0, 40), joined: true}; ps.push(p); }
    upsert_('Trips', {id: t.id, people: ps, updatedAt: Date.now()});
    return J({personId: p.id, data: tripData_(t.id)});
  });
}
function api_updateTrip(tripId, patchJson) {
  const patch = JSON.parse(patchJson), clean = {id: tripId, updatedAt: Date.now()};
  TRIP_PATCH_KEYS.forEach(k => { if (k in patch) clean[k] = patch[k]; });
  return withLock_(() => {
    requireTrip_(tripId);
    if (patch.regenCode) clean.code = newCode_();
    upsert_('Trips', clean);
    return J(tripData_(tripId));
  });
}
function api_deleteTrip(tripId) {
  return withLock_(() => {
    requireTrip_(tripId);
    CHILD.forEach(n => {
      const sh = sheet_(n);
      rows_(n).filter(r => r.tripId === tripId).sort((a, b) => b._row - a._row).forEach(r => sh.deleteRow(r._row));
    });
    remove_('Trips', tripId);
    try { const f = tripFolder_(tripId, false); if (f) f.setTrashed(true); } catch (e) {}
    return J(true);
  });
}

/* ================= Item API (Places / Expenses / Deposits / Docs) ================= */
function api_save(table, json) {
  if (CHILD.indexOf(table) < 0) throw new Error('ตารางไม่ถูกต้อง');
  const o = JSON.parse(json);
  return withLock_(() => {
    requireTrip_(o.tripId);
    const now = Date.now();
    if (!o.id) { o.id = uid_(table.slice(0, 2).toLowerCase()); o.createdAt = now; }
    const old = getById_(table, o.id);
    if (old && old.tripId !== o.tripId) throw new Error('รายการไม่อยู่ในทริปนี้');
    o.updatedAt = now;
    if (old && 'files' in o) trashRemoved_(old.files, o.files);
    upsert_(table, o);
    return J(tripData_(o.tripId));
  });
}
function api_remove(table, tripId, id) {
  if (CHILD.indexOf(table) < 0) throw new Error('ตารางไม่ถูกต้อง');
  return withLock_(() => {
    const old = getById_(table, id);
    if (old && old.tripId === tripId) { trashRemoved_(old.files, []); remove_(table, id); }
    return J(tripData_(tripId));
  });
}
function api_saveOrders(tripId, json) {
  const list = JSON.parse(json);
  return withLock_(() => {
    const sh = sheet_('Places'), col = TABLES.Places.indexOf('order') + 1;
    const byId = {}; rows_('Places').forEach(r => { if (r.tripId === tripId) byId[r.id] = r._row; });
    list.forEach(x => { if (byId[x.id]) sh.getRange(byId[x.id], col).setValue(x.order); });
    return J(tripData_(tripId));
  });
}

/* ================= Files (Google Drive) ================= */
function rootFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('ROOT_FOLDER');
  if (id) { try { const f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) {} }
  const f = DriveApp.createFolder(APP_NAME + ' - ไฟล์แนบ');
  props.setProperty('ROOT_FOLDER', f.getId());
  return f;
}
function tripFolder_(tripId, create) {
  const root = rootFolder_(), it = root.getFoldersByName(tripId);
  if (it.hasNext()) return it.next();
  return create === false ? null : root.createFolder(tripId);
}
function api_upload(tripId, name, mime, b64) {
  requireTrip_(tripId);
  const blob = Utilities.newBlob(Utilities.base64Decode(b64), mime, name);
  const f = tripFolder_(tripId).createFile(blob);
  try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  return J({id: f.getId(), name: name, type: mime});
}
function trashRemoved_(before, after) {
  const keep = {}; (after || []).forEach(f => keep[f.id] = 1);
  (before || []).forEach(f => { if (!keep[f.id]) { try { DriveApp.getFileById(f.id).setTrashed(true); } catch (e) {} } });
}

/* ================= Maps ================= */
const MODE_ = {
  driving: Maps.DirectionFinder.Mode.DRIVING,
  walking: Maps.DirectionFinder.Mode.WALKING,
  transit: Maps.DirectionFinder.Mode.TRANSIT,
  bicycling: Maps.DirectionFinder.Mode.BICYCLING
};
const LABELS_ = '123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ';

function parseLoc_(s) {
  s = String(s || '');
  const m = s.match(/!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)/) ||
            s.match(/[?&](?:q|query|ll|destination|daddr|center)=(-?\d+\.\d+)(?:,|%2C)\s*\+?(-?\d+\.\d+)/i) ||
            s.match(/@(-?\d+\.\d+),(-?\d+\.\d+)/) ||
            s.match(/^\s*(-?\d{1,2}\.\d+)\s*,\s*(-?\d{1,3}\.\d+)\s*$/);
  if (!m) return null;
  const lat = +m[1], lng = +m[2];
  return (Math.abs(lat) <= 90 && Math.abs(lng) <= 180) ? {lat: lat, lng: lng} : null;
}
function placeName_(u) {
  const m = String(u).match(/\/place\/([^/@?]+)/);
  if (!m) return '';
  try { return decodeURIComponent(m[1].replace(/\+/g, ' ')); } catch (e) { return ''; }
}
/** แปลงลิงก์ Google Maps (รวมลิงก์ย่อ maps.app.goo.gl) เป็นพิกัด */
function api_resolveLink(url) {
  let u = String(url || '').trim(), name = '';
  for (let i = 0; i < 5; i++) {
    name = name || placeName_(u);
    const g = parseLoc_(u);
    if (g) { g.name = name; return J(g); }
    if (!/^https?:\/\//i.test(u)) break;
    let r;
    try { r = UrlFetchApp.fetch(u, {followRedirects: false, muteHttpExceptions: true}); } catch (e) { break; }
    const h = r.getHeaders(), loc = h.Location || h.location;
    if (loc) { u = Array.isArray(loc) ? loc[0] : loc; continue; }
    const body = r.getContentText().slice(0, 200000);
    const g2 = parseLoc_(body.match(/https:\/\/www\.google\.[^"'\s]*\/maps[^"'\s]*/) ? body.match(/https:\/\/www\.google\.[^"'\s]*\/maps[^"'\s]*/)[0] : '') ||
               (body.match(/center=(-?\d+\.\d+)%2C(-?\d+\.\d+)/) ? {lat: +RegExp.$1, lng: +RegExp.$2} : null);
    if (g2) { g2.name = name; return J(g2); }
    break;
  }
  const q = name || (u.match(/[?&]q=([^&]+)/) ? decodeURIComponent(RegExp.$1.replace(/\+/g, ' ')) : '');
  if (q) { const list = geocode_(q); if (list.length) { list[0].name = name || list[0].name; return J(list[0]); } }
  return J(null);
}
function geocode_(q) {
  const key = 'g_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, q));
  const c = CacheService.getScriptCache(), hit = c.get(key);
  if (hit) return JSON.parse(hit);
  const res = Maps.newGeocoder().setLanguage('th').setRegion('th').geocode(q);
  const out = (res.results || []).slice(0, 6).map(r => ({
    name: (r.address_components && r.address_components[0] && r.address_components[0].long_name) || q,
    address: r.formatted_address,
    lat: r.geometry.location.lat, lng: r.geometry.location.lng
  }));
  try { c.put(key, JSON.stringify(out), 21600); } catch (e) {}
  return out;
}
function api_geocode(q) { return J(geocode_(String(q || '').slice(0, 200))); }

function leg_(a, b, mode) {
  const key = ['L', mode, a.lat.toFixed(5), a.lng.toFixed(5), b.lat.toFixed(5), b.lng.toFixed(5)].join('_');
  const c = CacheService.getScriptCache(), hit = c.get(key);
  if (hit) return JSON.parse(hit);
  let out;
  try {
    const df = Maps.newDirectionFinder().setOrigin(a.lat, a.lng).setDestination(b.lat, b.lng)
      .setMode(MODE_[mode] || MODE_.driving).setLanguage('th').setRegion('th');
    if (mode === 'transit') df.setDepart(new Date());
    const d = df.getDirections();
    if (d.routes && d.routes.length) {
      const L = d.routes[0].legs[0];
      out = {ok: true, m: L.distance.value, s: L.duration.value, poly: d.routes[0].overview_polyline.points};
    } else out = {ok: false, status: d.status || 'NO_ROUTE'};
  } catch (e) { out = {ok: false, status: String(e.message || e)}; }
  try { c.put(key, JSON.stringify(out), 21600); } catch (e) {}
  return out;
}
/** ระยะทาง/เวลาแต่ละช่วง + รูปแผนที่เส้นทางทั้งวัน */
function api_route(ptsJson, mode, withMap) {
  const pts = JSON.parse(ptsJson).slice(0, 30);
  const legs = [], polys = [];
  for (let i = 0; i < pts.length - 1; i++) {
    const L = leg_(pts[i], pts[i + 1], mode);
    legs.push({ok: L.ok, m: L.m || 0, s: L.s || 0, status: L.status || ''});
    if (L.ok && L.poly) polys.push(L.poly);
  }
  let map = '';
  if (withMap && pts.length) { try { map = staticMap_(pts, polys); } catch (e) { map = ''; } }
  return J({legs: legs, map: map});
}
function staticMap_(pts, polys) {
  const key = 'M_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify([pts, polys])));
  const c = CacheService.getScriptCache(), hit = c.get(key);
  if (hit) return hit;
  const m = Maps.newStaticMap().setSize(640, 420).setLanguage('th').setFormat(Maps.StaticMap.Format.JPG);
  if (pts.length === 1) m.setCenter(pts[0].lat, pts[0].lng).setZoom(14);
  const budget = Math.max(20, Math.floor(260 / Math.max(1, polys.length)));
  polys.forEach(p => {
    const flat = Maps.decodePolyline(p), n = flat.length / 2, step = Math.max(1, Math.ceil(n / budget)), keep = [];
    for (let i = 0; i < n; i += step) keep.push(flat[2 * i], flat[2 * i + 1]);
    keep.push(flat[flat.length - 2], flat[flat.length - 1]);
    try { m.setPathStyle(5, '0x0E7A6BDD', null); } catch (e) {}
    m.addPath(Maps.encodePolyline(keep));
  });
  pts.forEach((p, i) => {
    m.setMarkerStyle(Maps.StaticMap.MarkerSize.MID, i === 0 ? '0x13302A' : '0xE9A23B', LABELS_[i] || '');
    m.addMarker(p.lat, p.lng);
  });
  const url = 'data:image/jpeg;base64,' + Utilities.base64Encode(m.getBlob().getBytes());
  if (url.length < 99000) { try { c.put(key, url, 21600); } catch (e) {} }
  return url;
}
/** จัดลำดับจุดกลางให้เส้นทางสั้นที่สุด (จุดแรก/จุดสุดท้ายคงที่) */
function api_optimize(ptsJson, mode) {
  const pts = JSON.parse(ptsJson);
  if (pts.length < 4) return J(null);
  const df = Maps.newDirectionFinder()
    .setOrigin(pts[0].lat, pts[0].lng).setDestination(pts[pts.length - 1].lat, pts[pts.length - 1].lng)
    .setMode(mode === 'walking' ? MODE_.walking : MODE_.driving).setOptimizeWaypoints(true);
  pts.slice(1, -1).slice(0, 23).forEach(p => df.addWaypoint(p.lat, p.lng));
  const d = df.getDirections();
  if (!d.routes || !d.routes.length) return J(null);
  return J(d.routes[0].waypoint_order);
}
