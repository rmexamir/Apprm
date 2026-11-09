// نسخه‌ی ۱۳ — API گیت‌هاب (آمار ابزارها، شمارش معکوس، کاربران، روبیکا)
// ===== Code.gs — بک‌اند با Google Sheets + Drive =====
const FH = ['id','title','grade','subject','type','year','fileId','size','downloads','created','exam','pin','link','sub'];
const P = () => PropertiesService.getScriptProperties();

// ===== API: صفحه‌ی گیت‌هاب با fetch به این وب‌اپ وصل می‌شود =====
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function doGet() {
  return json_({ ok: true, service: 'rah-mowafaghiat API' });
}
// هر درخواست: {fn: 'نام تابع', a: [آرگومان‌ها]} با بدنه‌ی text/plain (بدون preflight)
function doPost(e) {
  try {
    const b = JSON.parse(e.postData.contents || '{}');
    const f = API[b.fn];
    if (!f) throw new Error('درخواست نامعتبر');
    const data = f.apply(null, Array.isArray(b.a) ? b.a : []);
    return json_({ ok: true, data: data === undefined ? null : data });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

// یک‌بار اجرا کنید: شیت‌ها، پوشه و رمز اولیه را می‌سازد
function setup() {
  sheet_('Files', FH); catsSheet_();
  if (!P().getProperty('ADMIN_PASS')) P().setProperty('ADMIN_PASS', '1234');
  folder_();
}

function sheet_(n, h) {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(n);
  if (!s) { s = ss.insertSheet(n); s.appendRow(h); }
  if (n === 'Files' && s.getLastColumn() < h.length) s.getRange(1, 1, 1, h.length).setValues([h]);
  return s;
}
function folder_() {
  let id = P().getProperty('FOLDER');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder('فایل‌های سایت');
  P().setProperty('FOLDER', f.getId());
  return f;
}
// نرمال‌سازی رمز: ارقام فارسی/عربی به لاتین و حذف فاصله‌ی اضافی
function normPass_(s) {
  return String(s == null ? '' : s)
    .replace(/[۰-۹]/g, d => '۰۱۲۳۴۵۶۷۸۹'.indexOf(d))
    .replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d))
    .trim();
}
function auth_(p) {
  const stored = normPass_(P().getProperty('ADMIN_PASS'));
  if (!stored || normPass_(p) !== stored) { Utilities.sleep(1500); throw new Error('رمز اشتباه است'); }
}
function rows_() {
  const v = sheet_('Files', FH).getDataRange().getValues();
  v.shift();
  return v.filter(r => r[0]).map(r => Object.fromEntries(FH.map((k, i) => [k, r[i]])));
}
function rowIndex_(id) {
  const ids = sheet_('Files', FH).getRange(1, 1, sheet_('Files', FH).getLastRow(), 1).getValues();
  for (let i = 1; i < ids.length; i++) if (ids[i][0] === id) return i + 1;
  throw new Error('فایل پیدا نشد');
}

// ----- عمومی -----
function getData() {
  return { files: rows_().reverse(), cats: catsOut_(), notice: P().getProperty('NOTICE') || '', countdown: { title: P().getProperty('CD_TITLE') || '', date: P().getProperty('CD_DATE') || '' }, events: events_(), countdowns: cdList_() };
}
function events_() {
  try { const a = JSON.parse(P().getProperty('EVENTS') || '[]'); return Array.isArray(a) ? a : []; } catch (e) { return []; }
}
function hit(id) {
  const s = sheet_('Files', FH), r = rowIndex_(id), c = FH.indexOf('downloads') + 1;
  s.getRange(r, c).setValue((Number(s.getRange(r, c).getValue()) || 0) + 1);
}

// ----- مدیر -----
function checkPass(p) { auth_(p); return true; }
const CATH = ['kind', 'name', 'order', 'parent'];
function catsSheet_() {
  const s = sheet_('Cats', CATH);
  if (s.getLastColumn() < CATH.length) s.getRange(1, 1, 1, CATH.length).setValues([CATH]);
  return s;
}
function addCat(p, kind, name, parent) {
  auth_(p);
  const s = catsSheet_(), v = s.getDataRange().getValues(), par = String(parent || '');
  let mx = 0;
  for (let i = 1; i < v.length; i++) if (v[i][0] === kind && String(v[i][3] || '') === par) mx = Math.max(mx, Number(v[i][2]) || 0);
  s.appendRow([kind, String(name).trim(), mx + 1, par]);
}
function delCat(p, kind, name) {
  auth_(p);
  const s = catsSheet_(), v = s.getDataRange().getValues();
  for (let i = v.length - 1; i >= 1; i--) {
    const kill = (v[i][0] === kind && v[i][1] === name) || (kind === 'type' && v[i][0] === 'sub' && String(v[i][3]) === String(name));
    if (kill) s.deleteRow(i + 1);
  }
}
function setCatOrder(p, kind, names, parent) {
  auth_(p);
  const s = catsSheet_(), v = s.getDataRange().getValues(), par = String(parent || '');
  (names || []).forEach((n, idx) => {
    for (let i = 1; i < v.length; i++) {
      if (v[i][0] === kind && String(v[i][1]) === String(n) && String(v[i][3] || '') === par) { s.getRange(i + 1, 3).setValue(idx + 1); break; }
    }
  });
}
function catsOut_() {
  const cats = { grade: [], subject: [], type: [], exam: [] }, subs = [];
  const v = catsSheet_().getDataRange().getValues(); v.shift();
  v.filter(r => r[0]).map((r, i) => ({ k: String(r[0]), n: String(r[1]), o: Number(r[2]) || 0, p: String(r[3] || ''), i: i }))
    .sort((a, b) => a.o - b.o || a.i - b.i)
    .forEach(r => { if (r.k === 'sub') subs.push({ n: r.n, p: r.p }); else if (cats[r.k]) cats[r.k].push(r.n); });
  cats.subs = subs;
  return cats;
}
function addFile(p, m, b64, name, mime) {
  auth_(p);
  const blob = Utilities.newBlob(Utilities.base64Decode(b64), mime || 'application/octet-stream', name);
  const f = folder_().createFile(blob);
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  sheet_('Files', FH).appendRow([Utilities.getUuid(), m.title, m.grade, m.subject, m.type, m.year,
    f.getId(), f.getSize(), 0, new Date().toISOString(), m.exam || '', '', m.link || '', m.sub || '']);
}
function updateFile(p, id, m) {
  auth_(p);
  const s = sheet_('Files', FH), r = rowIndex_(id);
  ['title', 'grade', 'subject', 'type', 'year', 'exam', 'link', 'sub'].forEach(k =>
    s.getRange(r, FH.indexOf(k) + 1).setValue(m[k]));
}
function deleteFile(p, id) {
  auth_(p);
  const s = sheet_('Files', FH), r = rowIndex_(id);
  try { DriveApp.getFileById(s.getRange(r, 7).getValue()).setTrashed(true); } catch (e) {}
  s.deleteRow(r);
}

function seedDefaults(p, g, s, t, e) {
  auth_(p);
  const sh = catsSheet_();
  const have = new Set(sh.getDataRange().getValues().map(r => r[0] + '|' + r[1]));
  const add = [];
  [['grade', g], ['subject', s], ['type', t], ['exam', e || []]].forEach(x => x[1].forEach(n => { if (!have.has(x[0] + '|' + n)) add.push([x[0], n]); }));
  if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, 2).setValues(add);
}

function changePass(p, np) {
  auth_(p);
  if (!np || String(np).length < 6) throw new Error('رمز باید حداقل ۶ کاراکتر باشد');
  P().setProperty('ADMIN_PASS', normPass_(np));
}
function setNotice(p, text) { auth_(p); P().setProperty('NOTICE', String(text || '')); }

function togglePin(p, id) {
  auth_(p);
  const s = sheet_('Files', FH), r = rowIndex_(id), col = FH.indexOf('pin') + 1;
  s.getRange(r, col).setValue(s.getRange(r, col).getValue() ? '' : 1);
}
function setCountdown(p, title, date) {
  auth_(p); P().setProperty('CD_TITLE', String(title || '')); P().setProperty('CD_DATE', String(date || ''));
}
// مهلت‌ها و رویدادهای مهم (نمایش در صفحه اصلی)
function setEvents(p, json) {
  auth_(p);
  const a = JSON.parse(json);
  if (!Array.isArray(a) || a.length > 30) throw new Error('فهرست مهلت‌ها نامعتبر است');
  P().setProperty('EVENTS', JSON.stringify(a.map(x => ({ t: String(x.t || '').slice(0, 80), d: String(x.d || '').slice(0, 10) }))));
}
const RH = ['id', 'text', 'contact', 'date'];
function submitRequest(text, contact) {
  const t = String(text || '').trim().slice(0, 500);
  if (t.length < 3) throw new Error('متن درخواست را بنویسید');
  const ch = CacheService.getScriptCache(), n = Number(ch.get('rq')) || 0;
  if (n >= 30) throw new Error('تعداد درخواست‌ها زیاد است؛ کمی بعد دوباره تلاش کنید');
  ch.put('rq', String(n + 1), 3600);
  sheet_('Requests', RH).appendRow([Utilities.getUuid(), t, String(contact || '').trim().slice(0, 100), new Date().toISOString()]);
}
function getRequests(p) {
  auth_(p);
  const v = sheet_('Requests', RH).getDataRange().getValues(); v.shift();
  return v.filter(r => r[0]).map(r => ({ id: r[0], text: r[1], contact: r[2], date: r[3] })).reverse();
}
function delRequest(p, id) {
  auth_(p);
  const s = sheet_('Requests', RH), v = s.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (v[i][0] === id) { s.deleteRow(i + 1); return; }
}

// بارگذاری فایل‌های بزرگ (تا ۱۵۰ مگابایت) به‌صورت مستقیم از مرورگر به Drive
function startUpload(p, name, mime, size, origin) {
  auth_(p);
  const r = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable', {
    method: 'post', muteHttpExceptions: true, contentType: 'application/json; charset=UTF-8',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken(), 'X-Upload-Content-Type': mime,
      'X-Upload-Content-Length': String(size), Origin: origin },
    payload: JSON.stringify({ name: name, parents: [folder_().getId()] })
  });
  if (r.getResponseCode() !== 200) throw new Error('شروع بارگذاری ناموفق بود (' + r.getResponseCode() + ')');
  const h = r.getAllHeaders();
  return { url: h.Location || h.location };
}
function finishUpload(p, m, id, name) {
  auth_(p);
  let f = null;
  if (id) f = DriveApp.getFileById(id);
  else {
    const it = folder_().getFilesByName(name);
    while (it.hasNext()) { const x = it.next(); if (!f || x.getDateCreated() > f.getDateCreated()) f = x; }
  }
  if (!f) throw new Error('فایل بارگذاری‌شده پیدا نشد');
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  sheet_('Files', FH).appendRow([Utilities.getUuid(), m.title, m.grade, m.subject, m.type, m.year,
    f.getId(), f.getSize(), 0, new Date().toISOString(), m.exam || '', '', m.link || '', m.sub || '']);
}

// فقط این توابع از بیرون قابل فراخوانی هستند (setup عمداً نیست)
const API = {
  getData, hit, checkPass, addCat, delCat, addFile, updateFile, deleteFile,
  seedDefaults, changePass, setNotice, togglePin, setCountdown, setEvents,
  submitRequest, getRequests, delRequest, startUpload, finishUpload,
  trackClick, getClickStats, getToolStats, setCountdowns, registerUser, loginUser, getMe, logoutUser,
  adminCreateUser, listUsers, deleteUser, resetUserPass,
  adminUploadAsset, adminSaveExam, adminListExams, adminSetExam, adminDeleteExam, setCatOrder,
  listExams, startExam, submitExam, myResults, examAnalysis, raffle
};

// ----- شمارش کلیک‌ها (شیت Clicks: تاریخ به‌صورت yyyymmdd، کلید، تعداد) -----
const CLK = ['date', 'key', 'count'];
function tehran_(ts, f) { return Utilities.formatDate(new Date(ts), 'Asia/Tehran', f); }

// عمومی: هر کلیک یک بار صدا زده می‌شود
function trackClick(key) {
  const k = String(key || '').replace(/[^a-zA-Z0-9_-]/g, '').slice(0, 30);
  if (!k) return;
  const d = Number(tehran_(Date.now(), 'yyyyMMdd'));
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return; }
  try {
    const s = sheet_('Clicks', CLK), v = s.getDataRange().getValues();
    for (let i = 1; i < v.length; i++) {
      if (Number(v[i][0]) === d && String(v[i][1]) === k) {
        s.getRange(i + 1, 3).setValue((Number(v[i][2]) || 0) + 1);
        return;
      }
    }
    s.appendRow([d, k, 1]);
  } finally { lock.releaseLock(); }
}

// فقط مدیر: آمار روزانه، هفتگی و ماهانه‌ی کلیک روی روبیکا
function getClickStats(p) {
  auth_(p);
  const s = sheet_('Clicks', CLK), v = s.getDataRange().getValues(); v.shift();
  const per = {}; let total = 0;
  v.forEach(r => {
    if (String(r[1]) === 'rubika') {
      const n = Number(r[2]) || 0, k = String(r[0]);
      per[k] = (per[k] || 0) + n; total += n;
    }
  });
  const now = Date.now();
  const day = i => new Date(now - i * 864e5);
  const dayCount = i => per[tehran_(day(i), 'yyyyMMdd')] || 0;

  let week = 0, month = 0;
  for (let i = 0; i < 30; i++) { if (i < 7) week += dayCount(i); month += dayCount(i); }

  const daily = [];
  for (let i = 13; i >= 0; i--) daily.push({ d: tehran_(day(i), 'yyyy-MM-dd'), n: dayCount(i) });

  const weekly = [];
  for (let j = 7; j >= 0; j--) {
    let n = 0;
    for (let i = j * 7; i < j * 7 + 7; i++) n += dayCount(i);
    weekly.push({ d: tehran_(day(j * 7 + 6), 'yyyy-MM-dd'), n: n });
  }

  const ny = Number(tehran_(now, 'yyyy')), nm = Number(tehran_(now, 'M'));
  const monthly = [];
  for (let j = 5; j >= 0; j--) {
    let mm = nm - j, yy = ny;
    while (mm <= 0) { mm += 12; yy--; }
    const ym = yy * 100 + mm;
    let n = 0;
    Object.keys(per).forEach(k => { if (Math.floor(Number(k) / 100) === ym) n += per[k]; });
    monthly.push({ m: yy + '-' + String(mm).padStart(2, '0'), n: n });
  }

  return { today: dayCount(0), week: week, month: month, total: total, daily: daily, weekly: weekly, monthly: monthly };
}

// ----- شمارش معکوس‌ها (فهرست چند‌تایی) -----
function cdList_() {
  try { const a = JSON.parse(P().getProperty('COUNTDOWNS') || '[]'); return Array.isArray(a) ? a : []; } catch (e) { return []; }
}
function setCountdowns(p, json) {
  auth_(p);
  const a = JSON.parse(json);
  if (!Array.isArray(a) || a.length > 20) throw new Error('فهرست شمارش‌ها نامعتبر است');
  P().setProperty('COUNTDOWNS', JSON.stringify(a.map(x => ({ t: String(x.t || '').slice(0, 80), d: String(x.d || '').slice(0, 10) }))));
}

// ----- کاربران و ثبت‌نام -----
const UH = ['id', 'username', 'name', 'grade', 'salt', 'hash', 'token', 'created'];
function usersSheet_() { return sheet_('Users', UH); }
function hashPw_(salt, pw) {
  const d = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(salt) + '|' + String(pw), Utilities.Charset.UTF_8);
  return Utilities.base64Encode(d);
}
function normUser_(u) { return String(u || '').trim().toLowerCase(); }
function validUsername_(u) { return /^[\w.\u0600-\u06FF-]{3,30}$/.test(u); }
function findUserRow_(v, pred) { for (let i = 1; i < v.length; i++) if (pred(v[i])) return i; return -1; }
function newToken_() { return Utilities.getUuid() + Utilities.getUuid(); }

function registerUser(username, password, name, grade) {
  const u = normUser_(username), pw = String(password || '');
  if (!validUsername_(u)) throw new Error('نام کاربری باید ۳ تا ۳۰ نویسه (حرف، عدد، _ . -) باشد');
  if (pw.length < 6) throw new Error('رمز باید حداقل ۶ کاراکتر باشد');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const s = usersSheet_(), v = s.getDataRange().getValues();
    if (findUserRow_(v, r => normUser_(r[1]) === u) !== -1) throw new Error('این نام کاربری قبلاً ثبت شده است');
    const salt = Utilities.getUuid(), token = newToken_();
    const nm = String(name || '').trim().slice(0, 60), gr = String(grade || '').slice(0, 40);
    s.appendRow([Utilities.getUuid(), u, nm, gr, salt, hashPw_(salt, pw), token, new Date().toISOString()]);
    return { token: token, user: { username: u, name: nm, grade: gr } };
  } finally { lock.releaseLock(); }
}
function loginUser(username, password) {
  const u = normUser_(username), pw = String(password || '');
  const s = usersSheet_(), v = s.getDataRange().getValues();
  const i = findUserRow_(v, r => normUser_(r[1]) === u);
  if (i === -1 || hashPw_(v[i][4], pw) !== String(v[i][5])) { Utilities.sleep(1500); throw new Error('نام کاربری یا رمز اشتباه است'); }
  const token = newToken_();
  s.getRange(i + 1, 7).setValue(token);
  return { token: token, user: { username: v[i][1], name: v[i][2], grade: v[i][3] } };
}
function getMe(token) {
  if (!token) return null;
  const v = usersSheet_().getDataRange().getValues();
  const i = findUserRow_(v, r => r[6] !== '' && String(r[6]) === String(token));
  return i === -1 ? null : { username: v[i][1], name: v[i][2], grade: v[i][3] };
}
function logoutUser(token) {
  if (!token) return;
  const s = usersSheet_(), v = s.getDataRange().getValues();
  const i = findUserRow_(v, r => r[6] !== '' && String(r[6]) === String(token));
  if (i !== -1) s.getRange(i + 1, 7).setValue('');
}
function adminCreateUser(p, username, password, name, grade) {
  auth_(p);
  return registerUser(username, password, name, grade).user;
}
function listUsers(p) {
  auth_(p);
  const v = usersSheet_().getDataRange().getValues(); v.shift();
  return v.filter(r => r[0]).map(r => ({ id: r[0], username: r[1], name: r[2], grade: r[3], created: r[7] })).reverse();
}
function deleteUser(p, id) {
  auth_(p);
  const s = usersSheet_(), v = s.getDataRange().getValues();
  const i = findUserRow_(v, r => String(r[0]) === String(id));
  if (i !== -1) s.deleteRow(i + 1);
}
function resetUserPass(p, id, np) {
  auth_(p);
  if (String(np || '').length < 6) throw new Error('رمز باید حداقل ۶ کاراکتر باشد');
  const s = usersSheet_(), v = s.getDataRange().getValues();
  const i = findUserRow_(v, r => String(r[0]) === String(id));
  if (i === -1) throw new Error('کاربر پیدا نشد');
  const salt = Utilities.getUuid();
  s.getRange(i + 1, 5, 1, 4).setValues([[salt, hashPw_(salt, np), '', new Date().toISOString()]]);
}

// فقط مدیر: آمار استفاده از ابزارها (کلیدهایی که با tool_ شروع می‌شوند)
function getToolStats(p) {
  auth_(p);
  const s = sheet_('Clicks', CLK), v = s.getDataRange().getValues(); v.shift();
  const now = Date.now(), idx = {};
  for (let i = 0; i < 30; i++) idx[Number(tehran_(new Date(now - i * 864e5), 'yyyyMMdd'))] = i;
  const out = {};
  v.forEach(r => {
    const k = String(r[1]);
    if (k.indexOf('tool_') !== 0) return;
    const n = Number(r[2]) || 0, d = Number(r[0]);
    const o = out[k] || (out[k] = { key: k, today: 0, week: 0, month: 0, total: 0 });
    o.total += n;
    if (idx.hasOwnProperty(d)) {
      const i = idx[d];
      if (i === 0) o.today += n;
      if (i < 7) o.week += n;
      o.month += n;
    }
  });
  return Object.keys(out).map(k => out[k]).sort((a, b) => b.total - a.total);
}

// ===================== آزمون آنلاین =====================
const EXH = ['id', 'title', 'track', 'durationMin', 'startAt', 'endAt', 'published', 'showAnswers', 'questions', 'createdAt'];
const ATH = ['id', 'examId', 'username', 'name', 'startedAt', 'submittedAt', 'answers', 'score', 'total', 'percent', 'durationSec'];
const RFH = ['id', 'examId', 'at', 'count', 'winners'];
const EX_GRACE = 20000; // میلی‌ثانیه، فرصت کوتاه برای ارسال پاسخ‌ها در لحظه‌ی پایان

function uidE_() { return Utilities.getUuid(); }
function safeObj_(s) { try { const o = JSON.parse(s); return (o && typeof o === 'object') ? o : {}; } catch (e) { return {}; } }
function safeArr_(s) { try { const a = JSON.parse(s); return Array.isArray(a) ? a : []; } catch (e) { return []; } }
function examFolder_() {
  let id = P().getProperty('EXAM_FOLDER');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder('فایل‌های آزمون');
  P().setProperty('EXAM_FOLDER', f.getId());
  return f;
}
// آپلود تصویر یا PDF سؤال (فقط مدیر). فایل با لینک عمومی خوانده می‌شود.
function adminUploadAsset(p, name, mime, b64) {
  auth_(p);
  if (!/^(image\/(png|jpeg|webp)|application\/pdf)$/.test(String(mime))) throw new Error('فقط تصویر (PNG/JPEG/WEBP) یا PDF مجاز است');
  const bytes = Utilities.base64Decode(String(b64));
  if (bytes.length > 8 * 1024 * 1024) throw new Error('حجم فایل بیشتر از ۸ مگابایت است');
  const f = examFolder_().createFile(Utilities.newBlob(bytes, String(mime), String(name || 'asset').slice(0, 80)));
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { id: f.getId() };
}
function examRows_() {
  const v = sheet_('Exams', EXH).getDataRange().getValues(); v.shift();
  return v.filter(r => r[0]).map(r => ({
    id: String(r[0]), title: String(r[1]), track: String(r[2]), durationMin: Number(r[3]),
    startAt: r[4] ? String(r[4]) : '', endAt: r[5] ? String(r[5]) : '',
    published: r[6] === true || String(r[6]).toUpperCase() === 'TRUE',
    showAnswers: r[7] === true || String(r[7]).toUpperCase() === 'TRUE',
    data: safeObj_(r[8]), createdAt: r[9]
  }));
}
function attempts_() {
  const v = sheet_('Attempts', ATH).getDataRange().getValues(); v.shift();
  return v.filter(r => r[0]).map(r => ({
    id: String(r[0]), examId: String(r[1]), username: String(r[2]), name: String(r[3]),
    startedAt: Number(r[4]) || 0, submittedAt: Number(r[5]) || 0, answers: String(r[6] || ''),
    score: Number(r[7]) || 0, total: Number(r[8]) || 0, percent: Number(r[9]) || 0, durationSec: Number(r[10]) || 0
  }));
}
function adminSaveExam(p, json) {
  auth_(p);
  const e = JSON.parse(json);
  const qs = e.questions || [];
  if (!e.title || !qs.length) throw new Error('عنوان و حداقل یک سؤال لازم است');
  if (qs.length > 100) throw new Error('حداکثر ۱۰۰ سؤال در هر آزمون');
  const clean = qs.map((q, i) => {
    const opts = (q.options || []).map(o => String(o).slice(0, 300)).filter(Boolean);
    if (opts.length < 2 || opts.length > 6) throw new Error('سؤال ' + (i + 1) + ': بین ۲ تا ۶ گزینه لازم است');
    const ans = Number(q.answer);
    if (!(ans >= 0 && ans < opts.length)) throw new Error('سؤال ' + (i + 1) + ': پاسخ صحیح را مشخص کن');
    return { text: String(q.text || '').slice(0, 2000), img: String(q.img || ''), pdf: String(q.pdf || ''), options: opts, answer: ans };
  });
  const dur = Math.max(1, Math.min(600, Number(e.durationMin) || 30));
  const id = e.id || uidE_();
  const row = [id, String(e.title).slice(0, 120), String(e.track || '').slice(0, 60), dur,
    e.startAt || '', e.endAt || '', e.published === true, e.showAnswers === true,
    JSON.stringify({ pdf: String(e.pdf || ''), questions: clean }), new Date().toISOString()];
  const s = sheet_('Exams', EXH), v = s.getDataRange().getValues();
  let r = -1;
  for (let i = 1; i < v.length; i++) if (String(v[i][0]) === id) { r = i + 1; break; }
  if (r === -1) s.appendRow(row); else s.getRange(r, 1, 1, row.length).setValues([row]);
  return id;
}
function adminListExams(p) {
  auth_(p);
  return examRows_().map(e => ({ id: e.id, title: e.title, track: e.track, durationMin: e.durationMin, startAt: e.startAt, endAt: e.endAt,
    published: e.published, showAnswers: e.showAnswers, count: (e.data.questions || []).length, data: e.data }));
}
function adminSetExam(p, id, published, showAnswers) {
  auth_(p);
  const s = sheet_('Exams', EXH), v = s.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]) === String(id)) {
    s.getRange(i + 1, 7, 1, 2).setValues([[published === true || published === 'true', showAnswers === true || showAnswers === 'true']]);
    return true;
  }
  throw new Error('آزمون پیدا نشد');
}
function adminDeleteExam(p, id) {
  auth_(p);
  const s = sheet_('Exams', EXH), v = s.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]) === String(id)) { s.deleteRow(i + 1); return true; }
  throw new Error('آزمون پیدا نشد');
}
function examClosedReason_(e) {
  const now = Date.now();
  if (e.startAt && now < Date.parse(e.startAt)) return 'هنوز شروع نشده';
  if (e.endAt && now > Date.parse(e.endAt)) return 'مهلت شرکت تمام شده';
  return '';
}
function listExams(token) {
  const u = getMe(token), att = u ? attempts_() : [];
  return examRows_().filter(e => e.published).map(e => {
    const mine = u ? att.find(a => a.examId === e.id && a.username === u.username) : null;
    return { id: e.id, title: e.title, track: e.track, durationMin: e.durationMin, count: (e.data.questions || []).length,
      closed: examClosedReason_(e), state: mine ? (mine.submittedAt ? 'done' : 'running') : 'new',
      score: mine && mine.submittedAt ? mine.percent : null };
  }).reverse();
}
function userOfE_(token) {
  const u = getMe(token);
  if (!u) throw new Error('ابتدا وارد حساب کاربری شو');
  return u;
}
function startExam(token, id) {
  const u = userOfE_(token);
  const e = examRows_().find(x => x.id === String(id) && x.published);
  if (!e) throw new Error('آزمون پیدا نشد');
  const closed = examClosedReason_(e);
  if (closed) throw new Error(closed);
  const s = sheet_('Attempts', ATH), v = s.getDataRange().getValues();
  let r = -1;
  for (let i = 1; i < v.length; i++) if (String(v[i][1]) === e.id && String(v[i][2]) === u.username) { r = i + 1; break; }
  const now = Date.now();
  let started = now;
  if (r === -1) {
    s.appendRow([uidE_(), e.id, u.username, u.name, now, '', '', 0, 0, 0, 0]);
  } else {
    if (v[r - 1][5]) throw new Error('شما قبلاً در این آزمون شرکت کرده‌ای');
    started = Number(v[r - 1][4]);
  }
  const endsAt = started + e.durationMin * 60000;
  if (now > endsAt + EX_GRACE) { finalizeExam_(s, r, e, [], true); throw new Error('زمان آزمون تمام شده است'); }
  return { title: e.title, endsAt: endsAt, serverNow: now, pdf: e.data.pdf || '',
    questions: (e.data.questions || []).map(q => ({ text: q.text, img: q.img, pdf: q.pdf, options: q.options })) };
}
function submitExam(token, id, answers) {
  const u = userOfE_(token);
  const e = examRows_().find(x => x.id === String(id));
  if (!e) throw new Error('آزمون پیدا نشد');
  const s = sheet_('Attempts', ATH), v = s.getDataRange().getValues();
  let r = -1;
  for (let i = 1; i < v.length; i++) if (String(v[i][1]) === e.id && String(v[i][2]) === u.username) { r = i + 1; break; }
  if (r === -1) throw new Error('ابتدا آزمون را شروع کن');
  if (v[r - 1][5]) throw new Error('این آزمون قبلاً ثبت شده است');
  const started = Number(v[r - 1][4]), endsAt = started + e.durationMin * 60000;
  return finalizeExam_(s, r, e, Array.isArray(answers) ? answers : [], Date.now() > endsAt + EX_GRACE);
}
function finalizeExam_(s, r, e, ans, late) {
  const qs = e.data.questions || [];
  const now = Date.now();
  let score = 0;
  const correct = qs.map((q, i) => {
    const a = ans[i];
    if (!late && typeof a === 'number' && Number.isInteger(a) && a === q.answer) score++;
    return q.answer;
  });
  const total = qs.length, percent = total ? Math.round(score / total * 1000) / 10 : 0;
  const started = Number(s.getRange(r, 5).getValue());
  const dur = Math.min(Math.round((now - started) / 1000), e.durationMin * 60 + EX_GRACE / 1000);
  s.getRange(r, 6, 1, 6).setValues([[now, JSON.stringify(ans.slice(0, total)), score, total, percent, dur]]);
  return { score: score, total: total, percent: percent, late: !!late, correct: e.showAnswers ? correct : null };
}
function myResults(token) {
  const u = userOfE_(token);
  const exams = examRows_(), all = attempts_().filter(a => a.submittedAt);
  return all.filter(a => a.username === u.username).map(a => {
    const e = exams.find(x => x.id === a.examId) || { title: 'آزمون حذف‌شده', showAnswers: false, data: { questions: [] } };
    const peers = all.filter(x => x.examId === a.examId).sort((x, y) => y.percent - x.percent || x.durationSec - y.durationSec);
    const ans = safeArr_(a.answers);
    return { examId: a.examId, title: e.title, score: a.score, total: a.total, percent: a.percent, durationSec: a.durationSec,
      rank: peers.findIndex(x => x.username === u.username) + 1, participants: peers.length,
      top: peers.slice(0, 3).map(x => ({ name: x.name, percent: x.percent })),
      detail: e.showAnswers ? (e.data.questions || []).map((q, i) => ({ text: q.text, yours: ans[i], correct: q.answer, ok: ans[i] === q.answer })) : null
    };
  }).reverse();
}
function examAnalysis(p, id) {
  auth_(p);
  const e = examRows_().find(x => x.id === String(id));
  if (!e) throw new Error('آزمون پیدا نشد');
  const qs = e.data.questions || [];
  const atts = attempts_().filter(a => a.examId === e.id && a.submittedAt);
  const n = atts.length;
  const pcts = atts.map(a => a.percent).sort((a, b) => a - b);
  const r1 = x => Math.round(x * 10) / 10;
  const avg = n ? pcts.reduce((x, y) => x + y, 0) / n : 0;
  const median = n ? (n % 2 ? pcts[(n - 1) / 2] : (pcts[n / 2 - 1] + pcts[n / 2]) / 2) : 0;
  const perQ = qs.map((q, i) => {
    let ok = 0;
    atts.forEach(a => { const x = safeArr_(a.answers)[i]; if (typeof x === 'number' && x === q.answer) ok++; });
    return { i: i + 1, text: String(q.text || '').slice(0, 60), correct: ok, rate: n ? Math.round(ok / n * 100) : 0 };
  });
  const board = atts.slice().sort((a, b) => b.percent - a.percent || a.durationSec - b.durationSec)
    .map((a, k) => ({ rank: k + 1, name: a.name, username: a.username, percent: a.percent, score: a.score, total: a.total, durationSec: a.durationSec }));
  const rv = sheet_('Raffles', RFH).getDataRange().getValues(); rv.shift();
  const raffles = rv.filter(r => String(r[1]) === e.id).map(r => ({ at: String(r[2]), count: r[3], winners: safeArr_(r[4]) })).reverse();
  return { title: e.title, count: n, avg: r1(avg), median: r1(median), max: n ? pcts[n - 1] : 0, min: n ? pcts[0] : 0,
    avgDuration: n ? Math.round(atts.reduce((x, a) => x + a.durationSec, 0) / n) : 0, perQ: perQ, board: board, raffles: raffles };
}
// قرعه‌کشی: از شرکت‌کنندگانی که آزمون را ثبت کرده‌اند، با حداقل درصد دلخواه
function raffle(p, examId, count, minPercent) {
  auth_(p);
  const e = examRows_().find(x => x.id === String(examId));
  if (!e) throw new Error('آزمون پیدا نشد');
  const minP = Number(minPercent) || 0;
  const pool = attempts_().filter(a => a.examId === e.id && a.submittedAt && a.percent >= minP)
    .map(a => ({ username: a.username, name: a.name }));
  if (!pool.length) throw new Error('شرکت‌کننده‌ای برای قرعه‌کشی پیدا نشد');
  const k = Math.max(1, Math.min(Number(count) || 1, pool.length));
  for (let i = pool.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); const t = pool[i]; pool[i] = pool[j]; pool[j] = t; }
  const winners = pool.slice(0, k);
  sheet_('Raffles', RFH).appendRow([uidE_(), e.id, new Date().toISOString(), k, JSON.stringify(winners)]);
  return winners;
}
