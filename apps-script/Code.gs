/**
 * QUẢN LÝ YÊU CẦU SỬA CHỮA / BẢO DƯỠNG / HIỆU CHUẨN / CHẾ TẠO  (v2)
 * Trung tâm Bảo dưỡng Nội trường Hà Nội
 * Google Apps Script Web App + Google Sheets.
 * Chạy với quyền chủ dự án (USER_DEPLOYING), truy cập ẩn danh: ai có link đều XEM được (khách).
 * Admin/Editor đăng nhập bằng mã 6 số gửi qua email (danh sách trong sheet USERS).
 * Không cần chia sẻ Sheet/Drive cho ai. Mọi hàm công khai nhận tham số đầu tiên `tok` (mã phiên, rỗng = khách);
 * phân quyền được kiểm tra ở PHÍA SERVER trong mọi hàm.
 */

var SHEETS = {
  DATA_MAIN: 'DATA_MAIN', WORKFLOW: 'WORKFLOW', EXTENSIONS: 'EXTENSIONS', COMMENTS: 'COMMENTS',
  CHANGE_LOG: 'CHANGE_LOG', ATTACHMENTS: 'ATTACHMENTS', EMAIL_LOG: 'EMAIL_LOG', MASTER_DATA: 'MASTER_DATA',
  SETTINGS: 'SETTINGS', FORMS: 'FORMS', FORM_RULES: 'FORM_RULES', USERS: 'USERS', SESSIONS: 'SESSIONS', UNIT_CONTACTS: 'UNIT_CONTACTS'
};
var HEADERS = {
  DATA_MAIN: ['request_id','ngay_tn','nguon','nguoi_gui','sdt','dv_gui','noi_dung','loai','dv_tn','bieu_mau','ma_bravo','thiet_bi','nhom','deadline_goc','deadline','tt','ly_do_pending','ngay_xong','ghi_chu','legacy','ma_cu','pl_cu','ra_soat','ghi_chu_rs','tao_boi','tao_luc','cn_boi','cn_luc','uu_tien'],
  WORKFLOW: ['request_id','s1_ngay','s1_boi','s2_ngay','s2_boi','s3_ngay','s3_boi','s4_ngay','s4_boi'],
  EXTENSIONS: ['ext_id','request_id','deadline_cu','deadline_moi','ly_do','yeu_cau_boi','yeu_cau_luc','trang_thai','cap_duyet_can','duyet_boi','duyet_luc','ghi_chu_duyet'],
  COMMENTS: ['cm_id','request_id','thoi_gian','nguoi','noi_dung'],
  CHANGE_LOG: ['log_id','thoi_gian','nguoi','request_id','hanh_dong','truong','gia_tri_cu','gia_tri_moi','prev_hash','hash'],
  ATTACHMENTS: ['att_id','request_id','ten_file','drive_id','url','mime','thoi_gian','nguoi'],
  EMAIL_LOG: ['key','thoi_gian','den','loai','request_id','tieu_de','ket_qua'],
  MASTER_DATA: ['nhom_dm','gia_tri'], SETTINGS: ['key','value','mo_ta'], FORMS: ['ma','ten','ghi_chu'],
  FORM_RULES: ['loai','dv_tn','bieu_mau'],
  USERS: ['email','ten','vai_tro','cap_duyet','trang_thai','nhan_mail','ngay_them'],
  SESSIONS: ['th','email','tao_luc','het_han'],
  UNIT_CONTACTS: ['don_vi','emails','ghi_chu']
};
var PRIORITIES = ['Khẩn','Cao','Bình thường','Thấp'];
var SCHEMA_VER = '3';
var DATE_COLS = ['ngay_tn','deadline_goc','deadline','ngay_xong','deadline_cu','deadline_moi','s1_ngay','s2_ngay','s3_ngay','s4_ngay'];
var STATUSES = ['Open','Inprogress','Pending','Closed'];
var STEPS = [
  {k:'s1', ten:'Xác nhận thông tin'},
  {k:'s2', ten:'Cập nhật Bravo (ghi mã hồ sơ)'},
  {k:'s3', ten:'In tài liệu và trình ký'},
  {k:'s4', ten:'Cập nhật Vaeco', chiPV:true}
];
var DV_PV = 'Trung tâm Phục vụ Bảo dưỡng';

/* ============ Web app / API ============ */
var API_ = {
  deleteRequest: deleteRequest, deleteUser: deleteUser, getContacts: getContacts, saveContacts: saveContacts, notifyUnits: notifyUnits, sendReportNow: sendReportNow, getInit: getInit, getBootstrap: getBootstrap, listRequests: listRequests, getRequest: getRequest, getLog: getLog, createRequest: createRequest, updateRequest: updateRequest,
  setStep: setStep, changeStatus: changeStatus, requestExtension: requestExtension, decideExtension: decideExtension,
  listPendingExtensions: listPendingExtensions, addComment: addComment, uploadAttachment: uploadAttachment, getFile: getFile,
  listUsers: listUsers, saveUser: saveUser, saveSettings: saveSettings, resolveReview: resolveReview, verifyLog: verifyLog,
  exportReport: exportReport, importLegacy: importLegacy, requestLogin: requestLogin, verifyLogin: verifyLogin, logout: logout
};
/** API cho website tĩnh: POST {fn, args:[tok, ...]} (Content-Type text/plain để tránh preflight CORS). Chỉ các hàm trong API_ được gọi. */
function doPost(e) {
  SS_ = null; MEMO_ = {};
  try { ensureSchema_(); } catch (x) {}
  var out;
  try {
    var req = JSON.parse(e.postData.contents), fn = String(req.fn || '');
    if (!API_.hasOwnProperty(fn)) throw new Error('Hàm không hợp lệ');
    out = { ok: true, data: API_[fn].apply(null, req.args || []) };
  } catch (err) { out = { ok: false, error: String(err && err.message ? err.message : err) }; }
  return ContentService.createTextOutput(JSON.stringify(out === undefined ? null : out)).setMimeType(ContentService.MimeType.JSON);
}
function doGet() {
  SS_ = null; MEMO_ = {};
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Quản lý yêu cầu sửa chữa - bảo dưỡng')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ============ Hạ tầng ============ */
/* Bộ nhớ đệm trong MỘT lần thực thi: mở file Sheets một lần, mỗi sheet đọc một lần (xóa khi có ghi). */
var SS_ = null, MEMO_ = {};
function memoClear_(n) { if (n) delete MEMO_[n]; else MEMO_ = {}; }
function ss_() {
  if (SS_) return SS_;
  var id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  if (!id) throw new Error('Chưa chạy setup(). Xem HUONG_DAN.md.');
  SS_ = SpreadsheetApp.openById(id);
  return SS_;
}
function sh_(name) {
  var s = ss_().getSheetByName(name);
  if (!s) throw new Error('Thiếu sheet ' + name + '. Chạy lại setup().');
  return s;
}
function tz_() { return getSetting_('TZ') || 'Asia/Ho_Chi_Minh'; }
function today_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'); }
function now_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm:ss'); }
function norm_(v, col) {
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), (col && /_luc$|thoi_gian/.test(col)) ? 'yyyy-MM-dd HH:mm:ss' : 'yyyy-MM-dd');
  return v === null || v === undefined ? '' : v;
}
function readAll_(name) {
  if (MEMO_[name]) return MEMO_[name];
  var s = sh_(name), v = s.getDataRange().getValues();
  if (v.length < 1) return [];
  var h = v[0], out = [];
  for (var i = 1; i < v.length; i++) {
    var o = {}, empty = true;
    for (var j = 0; j < h.length; j++) { o[h[j]] = norm_(v[i][j], h[j]); if (v[i][j] !== '') empty = false; }
    if (!empty) { o._row = i + 1; out.push(o); }
  }
  MEMO_[name] = out;
  return out;
}
function appendRows_(name, objs) {
  if (!objs.length) return;
  memoClear_(name);
  var s = sh_(name), h = HEADERS[name];
  var rows = objs.map(function (o) { return h.map(function (k) { return o[k] === undefined ? '' : o[k]; }); });
  var start = s.getLastRow() + 1;
  var rng = s.getRange(start, 1, rows.length, h.length);
  rng.setNumberFormat('@');
  rng.setValues(rows);
}
function updateRow_(name, row, obj) {
  memoClear_(name);
  var s = sh_(name), h = HEADERS[name];
  var cur = s.getRange(row, 1, 1, h.length).getValues()[0];
  var rows = [h.map(function (k, i) { return obj.hasOwnProperty(k) ? obj[k] : cur[i]; })];
  var rng = s.getRange(row, 1, 1, h.length);
  rng.setNumberFormat('@'); rng.setValues(rows);
}
function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function getSettings_() {
  var o = {};
  readAll_(SHEETS.SETTINGS).forEach(function (r) { o[r.key] = r.value; });
  return o;
}
function getSetting_(k) {
  var cache = CacheService.getScriptCache(), c = cache.get('set_' + k);
  if (c !== null) return c;
  var v = getSettings_()[k]; v = v === undefined ? '' : String(v);
  return v;
}
function setSetting_(k, val, mota) {
  var s = sh_(SHEETS.SETTINGS), rows = readAll_(SHEETS.SETTINGS);
  for (var i = 0; i < rows.length; i++) if (rows[i].key === k) {
    updateRow_(SHEETS.SETTINGS, rows[i]._row, { value: val }); return;
  }
  appendRows_(SHEETS.SETTINGS, [{ key: k, value: val, mo_ta: mota || '' }]);
}
function num_(k, d) { var n = Number(getSetting_(k)); return isNaN(n) || getSetting_(k) === '' ? d : n; }

/* ============ Người dùng, đăng nhập & quyền ============ */
function findUser_(email) {
  var us = readAll_(SHEETS.USERS);
  for (var i = 0; i < us.length; i++)
    if (String(us[i].email).toLowerCase() === email && us[i].trang_thai !== 'Khóa') return us[i];
  return null;
}
function userObj_(u) { return { email: String(u.email).toLowerCase(), ten: u.ten || u.email, vai_tro: u.vai_tro, cap_duyet: u.cap_duyet || '', guest: false }; }
function authUser_(tok) {
  if (!tok) return null;
  var h = sha_('tok|' + tok), cache = CacheService.getScriptCache(), em = cache.get('ses_' + h);
  if (!em) {
    var now = now_(), s = readAll_(SHEETS.SESSIONS).filter(function (r) { return r.th === h && String(r.het_han) > now; })[0];
    if (!s) return null;
    em = String(s.email).toLowerCase(); cache.put('ses_' + h, em, 300);
  }
  var uc = cache.get('usr_' + em);
  if (uc) { try { return JSON.parse(uc); } catch (e) {} }
  var u = findUser_(em), uo = u ? userObj_(u) : null;
  if (uo) cache.put('usr_' + em, JSON.stringify(uo), 120);
  return uo;
}
function me_(tok) {
  var u = authUser_(tok);
  if (u) return u;
  if (tok) throw new Error('AUTH: Phiên đăng nhập đã hết hạn hoặc tài khoản bị khóa. Hãy đăng nhập lại.');
  if (getSetting_('KHACH_XEM') === '0') throw new Error('AUTH: Cần đăng nhập để xem.');
  return { email: '', ten: 'Khách', vai_tro: 'Viewer', cap_duyet: '', guest: true };
}
function maskEm_(me, s) { return me.guest ? String(s === undefined || s === null ? '' : s).replace(/\s*<.*$/, '').split('@')[0] : s; }
function normEmail_(e) {
  e = String(e || '').trim().toLowerCase();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e)) throw new Error('Email không hợp lệ');
  return e;
}
/** Bước 1: gửi mã 6 số (hiệu lực 10 phút) tới email nếu email có trong USERS. Trả lời giống nhau dù email có hay không. */
function requestLogin(tok, email) {
  email = normEmail_(email);
  var generic = { ok: true, msg: 'Nếu email có trong danh sách người dùng, mã xác nhận 6 số đã được gửi (hiệu lực 10 phút). Hãy kiểm tra cả thư rác.' };
  var c = CacheService.getScriptCache(), rk = 'rl_' + sha_(email), n = Number(c.get(rk) || 0);
  if (n >= 5) throw new Error('Đã yêu cầu quá nhiều lần, hãy thử lại sau 1 giờ.');
  c.put(rk, String(n + 1), 3600);
  var u = findUser_(email); if (!u) return generic;
  var code = ('000000' + (parseInt(Utilities.getUuid().replace(/-/g, '').slice(0, 8), 16) % 1000000)).slice(-6);
  c.put('otp_' + sha_(email), sha_('otp|' + email + '|' + code), 600); c.remove('att_' + sha_(email));
  MailApp.sendEmail({ to: email, subject: '[Quản lý yêu cầu] Mã đăng nhập: ' + code, name: 'Quản lý yêu cầu SC-BD',
    htmlBody: '<p>Mã đăng nhập của bạn: <b style="font-size:22px;letter-spacing:3px">' + code + '</b></p><p>Mã có hiệu lực 10 phút. Nếu không phải bạn yêu cầu, hãy bỏ qua thư này.</p>' });
  return generic;
}
/** Bước 2: đổi mã lấy mã phiên (token). */
function verifyLogin(tok, email, code) {
  email = normEmail_(email); code = String(code || '').trim();
  var c = CacheService.getScriptCache(), ak = 'att_' + sha_(email), a = Number(c.get(ak) || 0);
  if (a >= 5) throw new Error('Nhập sai quá 5 lần, hãy yêu cầu mã mới.');
  var want = c.get('otp_' + sha_(email)); c.put(ak, String(a + 1), 600);
  if (!want || want !== sha_('otp|' + email + '|' + code)) throw new Error('Mã không đúng hoặc đã hết hạn.');
  c.remove('otp_' + sha_(email));
  var u = findUser_(email); if (!u) throw new Error('Tài khoản không còn hiệu lực.');
  var token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
  var days = num_('SESSION_DAYS', 14);
  var exp = Utilities.formatDate(new Date(new Date().getTime() + days * 86400000), tz_(), 'yyyy-MM-dd HH:mm:ss');
  var me = userObj_(u);
  withLock_(function () {
    appendRows_(SHEETS.SESSIONS, [{ th: sha_('tok|' + token), email: email, tao_luc: now_(), het_han: exp }]);
    log_(me, [{ hanh_dong: 'Đăng nhập', truong: 'phiên', cu: '', moi: 'hết hạn ' + exp }]);
  });
  return { tok: token };
}
function logout(tok) {
  if (!tok) return { ok: true };
  var h = sha_('tok|' + tok);
  CacheService.getScriptCache().remove('ses_' + h);
  withLock_(function () {
    var s = readAll_(SHEETS.SESSIONS).filter(function (r) { return r.th === h; })[0];
    if (s) updateRow_(SHEETS.SESSIONS, s._row, { het_han: '1970-01-01 00:00:00' });
  });
  return { ok: true };
}
function can_(me, action) {
  var r = me.vai_tro;
  if (action === 'read') return r === 'Admin' || r === 'Editor' || r === 'Viewer';
  if (action === 'write') return r === 'Admin' || r === 'Editor';
  if (action === 'admin') return r === 'Admin';
  if (action === 'approve') return (r === 'Admin' || r === 'Editor') && (me.cap_duyet === 'Tổ' || me.cap_duyet === 'Phòng');
  return false;
}
function need_(me, action) {
  if (!can_(me, action)) throw new Error('Bạn không có quyền thực hiện thao tác này (' + me.vai_tro + ').');
}

/* ============ Nhật ký thay đổi (chỉ ghi thêm, có chuỗi băm) ============ */
function sha_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + (b < 0 ? b + 256 : b).toString(16)).slice(-2); }).join('');
}
function logHash_(prev, o) {
  return sha_([prev, o.log_id, o.thoi_gian, o.nguoi, o.request_id, o.hanh_dong, o.truong, o.gia_tri_cu, o.gia_tri_moi].join('|'));
}
/** entries: [{request_id, hanh_dong, truong, cu, moi}] - gọi bên trong withLock_ */
function log_(me, entries) {
  if (!entries.length) return;
  var s = sh_(SHEETS.CHANGE_LOG), last = s.getLastRow(), prev = '', seq = 0;
  if (last > 1) { var lr = s.getRange(last, 1, 1, 10).getValues()[0]; prev = String(lr[9]); seq = Number(String(lr[0]).replace(/\D/g, '')) || (last - 1); }
  var t = now_(), rows = [];
  entries.forEach(function (e) {
    seq++;
    var o = { log_id: 'LG-' + ('000000' + seq).slice(-6), thoi_gian: t, nguoi: me.email, request_id: e.request_id || '',
      hanh_dong: e.hanh_dong, truong: e.truong || '', gia_tri_cu: String(e.cu === undefined ? '' : e.cu), gia_tri_moi: String(e.moi === undefined ? '' : e.moi) };
    o.prev_hash = prev; o.hash = logHash_(prev, o); prev = o.hash; rows.push(o);
  });
  appendRows_(SHEETS.CHANGE_LOG, rows);
}
function verifyLog(tok) {
  var me = me_(tok); need_(me, 'admin');
  var rows = readAll_(SHEETS.CHANGE_LOG), prev = '', bad = [];
  rows.forEach(function (r) {
    if (String(r.prev_hash) !== prev || logHash_(prev, r) !== r.hash) bad.push(r.log_id);
    prev = r.hash;
  });
  return { tong: rows.length, loi: bad.slice(0, 20), ok: bad.length === 0 };
}

/* ============ Dữ liệu ============ */
function danhMuc_() {
  var m = {};
  readAll_(SHEETS.MASTER_DATA).forEach(function (r) { (m[r.nhom_dm] = m[r.nhom_dm] || []).push(r.gia_tri); });
  return m;
}
function rules_() { return readAll_(SHEETS.FORM_RULES).map(function (r) { return { loai: r.loai, dv_tn: r.dv_tn, bieu_mau: String(r.bieu_mau) }; }); }
function ruleFor_(loai, dvtn) {
  var rs = rules_();
  for (var i = 0; i < rs.length; i++) if (rs[i].loai === loai && rs[i].dv_tn === dvtn) return rs[i];
  return null;
}
function getBootstrap(tok) {
  try { ensureSchema_(); } catch (x) {}
  var me, cfg = getSettings_();
  try { me = me_(tok); } catch (e) {
    if (/AUTH:/.test(e.message)) return { needLogin: true, org: cfg.ORG_NAME || '', msg: e.message.replace(/^.*AUTH:\s*/, '') };
    throw e;
  }
  var pub = {};
  ['ORG_NAME','SAP_HAN_NGAY','PENDING_DAI_NGAY','KHONG_CAP_NHAT_NGAY','LEGACY_EDITABLE','EXT_PHONG_NGAY','PAGE_SIZE','KHACH_XEM','SESSION_DAYS','NOTIFY_AUTO','REPORT_WEEKLY','REPORT_MONTHLY','APP_URL'].forEach(function (k) { pub[k] = cfg[k]; });
  var cch = CacheService.getScriptCache(), st = null;
  try { var raw = cch.get('bs_static'); if (raw) st = JSON.parse(raw); } catch (e) { st = null; }
  if (!st) {
    var forms = {}; readAll_(SHEETS.FORMS).forEach(function (f) { forms[String(f.ma)] = f.ten; });
    st = { master: danhMuc_(), rules: rules_(), forms: forms };
    try { cch.put('bs_static', JSON.stringify(st), 180); } catch (e) {}
  }
  return { me: me, today: today_(), settings: pub, master: st.master, rules: st.rules, forms: st.forms,
    steps: STEPS, users: can_(me, 'admin') ? listUsers_() : null };
}
/** Một lần gọi duy nhất cho lần mở trang: cấu hình + danh sách (nhanh hơn gọi 2 lần nối tiếp). */
function getInit(tok) {
  var b = getBootstrap(tok);
  return { boot: b, list: b.needLogin ? null : listRequests(tok) };
}
function reqOut_(r, me) { var o = {}; HEADERS.DATA_MAIN.forEach(function (k) { o[k] = r[k]; }); if (PRIORITIES.indexOf(o.uu_tien) < 0) o.uu_tien = 'Bình thường'; if (me && me.guest) { o.tao_boi = maskEm_(me, o.tao_boi); o.cn_boi = maskEm_(me, o.cn_boi); } return o; }
function listRequests(tok) {
  var me = me_(tok); need_(me, 'read');
  var ex = {}; readAll_(SHEETS.EXTENSIONS).forEach(function (e) { (ex[e.request_id] = ex[e.request_id] || []).push(e); });
  return readAll_(SHEETS.DATA_MAIN).map(function (r) {
    var o = reqOut_(r, me);
    o.wf = {};
    var es = ex[r.request_id] || [];
    o.ext_cho = es.filter(function (e) { return e.trang_thai === 'Chờ duyệt'; }).length;
    var last = es.length ? es[es.length - 1] : null; o.ext_last = last ? last.trang_thai : '';
    o.ext_n = es.filter(function (e) { return e.trang_thai === 'Đã duyệt'; }).length;
    return o;
  });
}
function findReq_(id) {
  var rs = readAll_(SHEETS.DATA_MAIN);
  for (var i = 0; i < rs.length; i++) if (rs[i].request_id === id) return rs[i];
  throw new Error('Không tìm thấy yêu cầu ' + id);
}
function getRequest(tok, id) {
  var me = me_(tok); need_(me, 'read');
  var r = findReq_(id);
  function by(name) { return readAll_(name).filter(function (x) { return x.request_id === id; }); }
  var exts = by(SHEETS.EXTENSIONS), cms = by(SHEETS.COMMENTS), atts = by(SHEETS.ATTACHMENTS);
  if (me.guest) {   // khách: che email
    exts.forEach(function (e) { e.yeu_cau_boi = maskEm_(me, e.yeu_cau_boi); e.duyet_boi = maskEm_(me, e.duyet_boi); });
    cms.forEach(function (x) { x.nguoi = maskEm_(me, x.nguoi); });
    atts.forEach(function (x) { x.nguoi = maskEm_(me, x.nguoi); x.url = ''; });
  }
  return { req: reqOut_(r, me), wf: {}, exts: exts, comments: cms, atts: atts, log: [] };
}
/** Nhật ký thay đổi của một yêu cầu: chỉ tải khi người dùng mở mục này (sheet nhật ký lớn, đọc chậm). Khách không xem được. */
function getLog(tok, id) {
  var me = me_(tok); need_(me, 'read');
  if (me.guest) return [];
  return readAll_(SHEETS.CHANGE_LOG).filter(function (x) { return x.request_id === id; }).slice(-200);
}

/* ============ Kiểm tra dữ liệu ============ */
var ISO = /^\d{4}-\d{2}-\d{2}$/;
function needStr_(v, msg) { if (!v || !String(v).trim()) throw new Error(msg); return String(v).trim(); }
function checkDate_(v, msg) { if (!ISO.test(String(v || ''))) throw new Error(msg + ' (định dạng yyyy-mm-dd)'); return String(v); }
function legacyGuard_(r) {
  /* Dữ liệu cũ cũng được Editor cập nhật như yêu cầu mới. */
}
function addDays_(iso, n) { var d = new Date(iso + 'T00:00:00Z'); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); }
function diffDays_(a, b) { return Math.round((new Date(a + 'T00:00:00Z') - new Date(b + 'T00:00:00Z')) / 86400000); }

/* ============ Tạo / sửa yêu cầu ============ */
function createRequest(tok, p) {
  var me = me_(tok); need_(me, 'write');
  var dm = danhMuc_();
  var loai = needStr_(p.loai, 'Chọn loại yêu cầu'), dvtn = needStr_(p.dv_tn, 'Chọn đơn vị tiếp nhận');
  var rule = ruleFor_(loai, dvtn);
  if (!rule) throw new Error('Đơn vị "' + dvtn + '" không tiếp nhận loại "' + loai + '".');
  var o = {
    ngay_tn: checkDate_(p.ngay_tn || today_(), 'Ngày tiếp nhận'),
    nguon: String(p.nguon || '').trim(), nguoi_gui: needStr_(p.nguoi_gui, 'Nhập người gửi yêu cầu'),
    sdt: String(p.sdt || '').trim(), dv_gui: needStr_(p.dv_gui, 'Chọn đơn vị gửi'),
    noi_dung: needStr_(p.noi_dung, 'Nhập nội dung yêu cầu'), loai: loai, dv_tn: dvtn, bieu_mau: rule.bieu_mau,
    ma_bravo: '', thiet_bi: '', nhom: String(p.nhom || '').trim(),
    deadline_goc: checkDate_(p.deadline, 'Hạn hoàn thành dự kiến'), tt: 'Open', ly_do_pending: '', ngay_xong: '',
    ghi_chu: String(p.ghi_chu || '').trim(), legacy: 0, ma_cu: '', pl_cu: '', ra_soat: 0, ghi_chu_rs: '',
    tao_boi: me.email, tao_luc: now_(), cn_boi: me.email, cn_luc: now_(),
    uu_tien: PRIORITIES.indexOf(p.uu_tien) >= 0 ? p.uu_tien : 'Bình thường'
  };
  if (dm.dv_gui && dm.dv_gui.indexOf(o.dv_gui) < 0) throw new Error('Đơn vị gửi không hợp lệ.');
  if (o.deadline_goc < o.ngay_tn) throw new Error('Hạn hoàn thành không được trước ngày tiếp nhận.');
  o.deadline = o.deadline_goc;
  return withLock_(function () {
    var yr = o.ngay_tn.slice(0, 4), key = 'SEQ_' + yr, seq = (Number(getSettings_()[key]) || 0) + 1;
    setSetting_(key, seq, 'Bộ đếm mã yêu cầu năm ' + yr);
    o.request_id = 'YC-' + yr + '-' + ('00000' + seq).slice(-5);
    appendRows_(SHEETS.DATA_MAIN, [o]);
    appendRows_(SHEETS.WORKFLOW, [{ request_id: o.request_id }]);
    log_(me, [{ request_id: o.request_id, hanh_dong: 'Tạo yêu cầu', truong: 'tt', cu: '', moi: 'Open' }]);
    try { notifyAuto_(o, 'new', me); } catch (x) { }
    return { request_id: o.request_id, bieu_mau: o.bieu_mau };
  });
}

var EDITABLE = ['ngay_tn','nguoi_gui','sdt','dv_gui','loai','dv_tn','noi_dung','nhom','ghi_chu','ma_bravo','uu_tien'];
function updateRequest(tok, id, patch) {
  var me = me_(tok); need_(me, 'write');
  return withLock_(function () {
    var r = findReq_(id); legacyGuard_(r);
    var upd = {}, ents = [];
    EDITABLE.forEach(function (k) {
      if (patch.hasOwnProperty(k) && String(patch[k]).trim() !== String(r[k])) {
        if (k === 'noi_dung') needStr_(patch[k], 'Nội dung không được để trống');
        if (k === 'ngay_tn') checkDate_(String(patch[k]).trim(), 'Ngày tiếp nhận');
        if (k === 'uu_tien' && PRIORITIES.indexOf(String(patch[k])) < 0) throw new Error('Mức ưu tiên không hợp lệ');
        upd[k] = String(patch[k]).trim(); ents.push({ request_id: id, hanh_dong: 'Sửa thông tin', truong: k, cu: r[k], moi: upd[k] });
      }
    });
    if (!ents.length) return { ok: true, doi: 0 };
    if (upd.hasOwnProperty('loai') || upd.hasOwnProperty('dv_tn')) {
      var nl = upd.hasOwnProperty('loai') ? upd.loai : r.loai, nd = upd.hasOwnProperty('dv_tn') ? upd.dv_tn : r.dv_tn;
      if (nl && nd) { var ru = ruleFor_(nl, nd); if (!ru) throw new Error('Đơn vị "' + nd + '" không tiếp nhận loại "' + nl + '".'); upd.bieu_mau = ru.bieu_mau; }
    }
    upd.cn_boi = me.email; upd.cn_luc = now_();
    updateRow_(SHEETS.DATA_MAIN, r._row, upd); log_(me, ents);
    return { ok: true, doi: ents.length };
  });
}

/* ============ Xóa yêu cầu (Admin/Editor; bắt buộc lý do; lưu bản chụp vào nhật ký) ============ */
function deleteRequest(tok, id, ly_do) {
  var me = me_(tok); need_(me, 'write');
  needStr_(ly_do, 'Vui lòng nhập lý do xóa (nhập sai, yêu cầu bị hủy...)');
  var res = withLock_(function () {
    var r = findReq_(id);
    var snap = {}; ['request_id', 'ngay_tn', 'loai', 'dv_gui', 'dv_tn', 'noi_dung', 'ma_bravo', 'deadline', 'tt', 'uu_tien'].forEach(function (k) { snap[k] = r[k]; });
    log_(me, [{ request_id: id, hanh_dong: 'Xóa yêu cầu', truong: 'ly_do: ' + String(ly_do).trim(), cu: JSON.stringify(snap), moi: '(đã xóa)' }]);
    readAll_(SHEETS.EXTENSIONS).forEach(function (e) { if (e.request_id === id && e.trang_thai === 'Chờ duyệt') updateRow_(SHEETS.EXTENSIONS, e._row, { trang_thai: 'Đã hủy' }); });
    var w = readAll_(SHEETS.WORKFLOW).filter(function (x) { return x.request_id === id; })[0];
    if (w) { sh_(SHEETS.WORKFLOW).deleteRow(w._row); memoClear_(SHEETS.WORKFLOW); }
    sh_(SHEETS.DATA_MAIN).deleteRow(r._row); memoClear_();
    try { CacheService.getScriptCache().remove('bs_static'); } catch (x) {}
    return { ok: true };
  });
  return res;
}

/* ============ Quy trình 4 bước (chỉ điền thông tin để quản lý) ============ */
function setStep(tok, id, key, done, ngay) {
  var me = me_(tok); need_(me, 'write');
  if (!STEPS.some(function (s) { return s.k === key; })) throw new Error('Bước không hợp lệ');
  return withLock_(function () {
    var r = findReq_(id); legacyGuard_(r);
    var st = STEPS.filter(function (s) { return s.k === key; })[0];
    if (st.chiPV && r.dv_tn !== DV_PV) throw new Error('Bước "Cập nhật Vaeco" chỉ áp dụng cho luồng Trung tâm Phục vụ Bảo dưỡng.');
    var rows = readAll_(SHEETS.WORKFLOW), w = rows.filter(function (x) { return x.request_id === id; })[0];
    var d = done ? checkDate_(ngay || today_(), 'Ngày hoàn thành bước') : '';
    var upd = {}; upd[key + '_ngay'] = d; upd[key + '_boi'] = done ? me.email : '';
    if (w) updateRow_(SHEETS.WORKFLOW, w._row, upd); else { upd.request_id = id; appendRows_(SHEETS.WORKFLOW, [upd]); }
    log_(me, [{ request_id: id, hanh_dong: 'Quy trình', truong: st.ten, cu: w ? w[key + '_ngay'] : '', moi: d || '(bỏ đánh dấu)' }]);
    updateRow_(SHEETS.DATA_MAIN, r._row, { cn_boi: me.email, cn_luc: now_() });
    return { ok: true };
  });
}
function missingSteps_(r, w) {
  return STEPS.filter(function (s) { return (!s.chiPV || r.dv_tn === DV_PV) && !(w && w[s.k + '_ngay']); }).map(function (s) { return s.ten; });
}

/* ============ Trạng thái ============ */
/**
 * p: {tt, ly_do, deadline_moi, ngay_xong, xacNhanThieuBuoc}
 * Pending: bắt buộc lý do + ngày xin gia hạn -> tạo đề nghị gia hạn (Chờ duyệt).
 */
function changeStatus(tok, id, p) {
  var me = me_(tok); need_(me, 'write');
  if (STATUSES.indexOf(p.tt) < 0) throw new Error('Trạng thái không hợp lệ');
  return withLock_(function () {
    var r = findReq_(id); legacyGuard_(r);
    if (r.tt === p.tt) throw new Error('Yêu cầu đã ở trạng thái ' + p.tt);
    var upd = { tt: p.tt, cn_boi: me.email, cn_luc: now_() }, ents = [{ request_id: id, hanh_dong: 'Đổi trạng thái', truong: 'tt', cu: r.tt, moi: p.tt }];
    var w = readAll_(SHEETS.WORKFLOW).filter(function (x) { return x.request_id === id; })[0];
    if (p.tt === 'Pending') {
      var ly = needStr_(p.ly_do, 'Pending bắt buộc nhập lý do');
      var nm = checkDate_(p.deadline_moi, 'Ngày xin gia hạn đến');
      if (nm <= r.deadline) throw new Error('Ngày xin gia hạn phải sau hạn hiện tại (' + r.deadline + ').');
      upd.ly_do_pending = ly; ents.push({ request_id: id, hanh_dong: 'Lý do Pending', truong: 'ly_do_pending', cu: r.ly_do_pending, moi: ly });
      createExt_(me, r, nm, ly, ents);
    } else if (p.tt === 'Closed') {
      upd.ngay_xong = checkDate_(p.ngay_xong || today_(), 'Ngày hoàn thành');
      if (upd.ngay_xong < r.ngay_tn) throw new Error('Ngày hoàn thành không được trước ngày tiếp nhận.');
      if (me.vai_tro === 'Admin') ents.push({ request_id: id, hanh_dong: 'Admin đóng yêu cầu', truong: 'tt', cu: '', moi: me.email });
    } else {
      upd.ngay_xong = '';
      if (p.tt === 'Inprogress' || p.tt === 'Open') upd.ly_do_pending = '';
    }
    updateRow_(SHEETS.DATA_MAIN, r._row, upd); log_(me, ents);
    if (p.tt === 'Closed' || p.tt === 'Pending') { try { var nr = {}; for (var kk in r) nr[kk] = r[kk]; for (var k2 in upd) nr[k2] = upd[k2]; notifyAuto_(nr, p.tt === 'Closed' ? 'closed' : 'pending', me, { nm: nm }); } catch (x) { } }
    return { ok: true };
  });
}

/* ============ Gia hạn có duyệt ============ */
function capCan_(r, nm) {
  var n = num_('EXT_PHONG_NGAY', 0);
  return (n > 0 && diffDays_(nm, r.deadline) > n) ? 'Phòng' : 'Tổ';
}
function createExt_(me, r, nm, ly, ents) {
  var exts = readAll_(SHEETS.EXTENSIONS);
  exts.forEach(function (e) { if (e.request_id === r.request_id && e.trang_thai === 'Chờ duyệt') throw new Error('Yêu cầu đã có đề nghị gia hạn đang chờ duyệt.'); });
  var e = { ext_id: 'GH-' + ('00000' + (exts.length + 1)).slice(-5), request_id: r.request_id, deadline_cu: r.deadline, deadline_moi: nm,
    ly_do: ly, yeu_cau_boi: me.email, yeu_cau_luc: now_(), trang_thai: 'Chờ duyệt', cap_duyet_can: capCan_(r, nm),
    duyet_boi: '', duyet_luc: '', ghi_chu_duyet: '' };
  appendRows_(SHEETS.EXTENSIONS, [e]);
  ents.push({ request_id: r.request_id, hanh_dong: 'Xin gia hạn', truong: 'deadline', cu: r.deadline, moi: nm + ' (cần ' + e.cap_duyet_can + ' duyệt; ' + ly + ')' });
  try { mailExtRequest_(e, r); } catch (x) { }
  return e;
}
function requestExtension(tok, id, nm, ly) {
  var me = me_(tok); need_(me, 'write');
  return withLock_(function () {
    var r = findReq_(id); legacyGuard_(r);
    if (r.tt === 'Closed') throw new Error('Yêu cầu đã đóng.');
    needStr_(ly, 'Nhập lý do xin gia hạn'); checkDate_(nm, 'Ngày xin gia hạn đến');
    if (nm <= r.deadline) throw new Error('Ngày xin gia hạn phải sau hạn hiện tại (' + r.deadline + ').');
    var ents = []; var e = createExt_(me, r, nm, ly, ents);
    updateRow_(SHEETS.DATA_MAIN, r._row, { cn_boi: me.email, cn_luc: now_() });
    log_(me, ents); return { ok: true, ext_id: e.ext_id, cap_duyet_can: e.cap_duyet_can };
  });
}
function decideExtension(tok, extId, approve, note) {
  var me = me_(tok); need_(me, 'approve');
  return withLock_(function () {
    var e = readAll_(SHEETS.EXTENSIONS).filter(function (x) { return x.ext_id === extId; })[0];
    if (!e) throw new Error('Không tìm thấy đề nghị gia hạn');
    if (e.trang_thai !== 'Chờ duyệt') throw new Error('Đề nghị này đã được xử lý.');
    if (String(e.yeu_cau_boi).toLowerCase() === me.email) throw new Error('Không được tự duyệt đề nghị do chính mình lập.');
    if (e.cap_duyet_can === 'Phòng' && me.cap_duyet !== 'Phòng') throw new Error('Đề nghị này cần cán bộ cấp Phòng duyệt.');
    if (!approve) needStr_(note, 'Nhập lý do từ chối');
    var r = findReq_(e.request_id), ents = [];
    updateRow_(SHEETS.EXTENSIONS, e._row, { trang_thai: approve ? 'Đã duyệt' : 'Từ chối', duyet_boi: me.email, duyet_luc: now_(), ghi_chu_duyet: note || '' });
    ents.push({ request_id: r.request_id, hanh_dong: approve ? 'Duyệt gia hạn' : 'Từ chối gia hạn', truong: 'deadline', cu: r.deadline, moi: approve ? e.deadline_moi : r.deadline });
    if (approve) updateRow_(SHEETS.DATA_MAIN, r._row, { deadline: e.deadline_moi, cn_boi: me.email, cn_luc: now_() });
    log_(me, ents);
    try { mailExtDecision_(e, r, approve, note, me); } catch (x) { }
    if (approve) { try { notifyAuto_(r, 'extension', me, { cu: e.deadline_cu, moi: e.deadline_moi }); } catch (x) { } }
    return { ok: true };
  });
}
function listPendingExtensions(tok) {
  var me = me_(tok); need_(me, 'read');
  var reqs = {}; readAll_(SHEETS.DATA_MAIN).forEach(function (r) { reqs[r.request_id] = r; });
  return readAll_(SHEETS.EXTENSIONS).filter(function (e) { return e.trang_thai === 'Chờ duyệt'; }).map(function (e) {
    e.noi_dung = reqs[e.request_id] ? reqs[e.request_id].noi_dung : ''; e.the_duyet = can_(me, 'approve') &&
      String(e.yeu_cau_boi).toLowerCase() !== me.email && (e.cap_duyet_can !== 'Phòng' || me.cap_duyet === 'Phòng');
    e.yeu_cau_boi = maskEm_(me, e.yeu_cau_boi); return e;
  });
}

/* ============ Trao đổi, đính kèm ============ */
function addComment(tok, id, text) {
  var me = me_(tok); need_(me, 'write'); needStr_(text, 'Nhập nội dung trao đổi');
  return withLock_(function () {
    findReq_(id);
    var n = sh_(SHEETS.COMMENTS).getLastRow();
    appendRows_(SHEETS.COMMENTS, [{ cm_id: 'CM-' + ('00000' + n).slice(-5), request_id: id, thoi_gian: now_(), nguoi: me.ten + ' <' + me.email + '>', noi_dung: String(text).trim() }]);
    log_(me, [{ request_id: id, hanh_dong: 'Trao đổi', truong: 'comment', cu: '', moi: String(text).trim().slice(0, 120) }]);
    return { ok: true };
  });
}
function folder_() {
  var id = getSetting_('DRIVE_FOLDER_ID');
  if (!id) throw new Error('Chưa cấu hình DRIVE_FOLDER_ID (chạy setup()).');
  return DriveApp.getFolderById(id);
}
/** f: {name, mime, b64} - ảnh đã được thu nhỏ ở trình duyệt */
function uploadAttachment(tok, id, f) {
  var me = me_(tok); need_(me, 'write');
  if (!f || !f.b64) throw new Error('Thiếu tệp');
  if (f.b64.length > 9 * 1024 * 1024) throw new Error('Tệp quá lớn (giới hạn khoảng 6 MB).');
  var allowed = /^(image\/|application\/pdf|application\/vnd\.|application\/msword|text\/plain)/;
  if (!allowed.test(f.mime || '')) throw new Error('Chỉ nhận ảnh, PDF, Word/Excel.');
  findReq_(id);
  var blob = Utilities.newBlob(Utilities.base64Decode(f.b64), f.mime, id + '_' + String(f.name).replace(/[\\\/:*?"<>|]/g, '_'));
  var file = folder_().createFile(blob);
  return withLock_(function () {
    var n = sh_(SHEETS.ATTACHMENTS).getLastRow();
    appendRows_(SHEETS.ATTACHMENTS, [{ att_id: 'FI-' + ('00000' + n).slice(-5), request_id: id, ten_file: f.name, drive_id: file.getId(), url: file.getUrl(), mime: f.mime, thoi_gian: now_(), nguoi: me.email }]);
    log_(me, [{ request_id: id, hanh_dong: 'Đính kèm', truong: 'file', cu: '', moi: f.name }]);
    return { ok: true, url: file.getUrl() };
  });
}
/** Trả tệp đính kèm (base64) để trình duyệt tải về; khách cũng xem được vì tệp nằm trong Drive của chủ dự án. */
function getFile(tok, driveId) {
  var me = me_(tok); need_(me, 'read');
  var a = readAll_(SHEETS.ATTACHMENTS).filter(function (x) { return x.drive_id === driveId; })[0];
  if (!a) throw new Error('Tệp không thuộc hệ thống');
  var b = DriveApp.getFileById(driveId).getBlob(), bytes = b.getBytes();
  if (bytes.length > 7 * 1024 * 1024) throw new Error('Tệp quá lớn để tải qua ứng dụng.');
  return { name: a.ten_file, mime: b.getContentType(), b64: Utilities.base64Encode(bytes) };
}

/* ============ Quản trị ============ */
function listUsers_() { return readAll_(SHEETS.USERS).map(function (u) { return { email: u.email, ten: u.ten, vai_tro: u.vai_tro, cap_duyet: u.cap_duyet, trang_thai: u.trang_thai, nhan_mail: u.nhan_mail }; }); }
function listUsers(tok) { var me = me_(tok); need_(me, 'admin'); return listUsers_(); }
/** Xóa người dùng khỏi danh sách phân quyền (Admin). Không tự xóa mình; luôn còn ít nhất một Admin. */
function deleteUser(tok, email) {
  var me = me_(tok); need_(me, 'admin');
  email = String(email || '').toLowerCase().trim();
  if (email === String(me.email).toLowerCase()) throw new Error('Không thể tự xóa chính mình.');
  return withLock_(function () {
    var rows = readAll_(SHEETS.USERS), ex = rows.filter(function (x) { return String(x.email).toLowerCase() === email; })[0];
    if (!ex) throw new Error('Không tìm thấy người dùng');
    var admins = rows.filter(function (x) { return x.vai_tro === 'Admin' && x.trang_thai !== 'Khóa' && String(x.email).toLowerCase() !== email; });
    if (ex.vai_tro === 'Admin' && !admins.length) throw new Error('Phải còn ít nhất một Admin.');
    log_(me, [{ hanh_dong: 'Xóa người dùng', truong: email, cu: ex.vai_tro + '/' + ex.cap_duyet + '/' + ex.trang_thai, moi: '(đã xóa)' }]);
    sh_(SHEETS.USERS).deleteRow(ex._row); memoClear_();
    try { CacheService.getScriptCache().remove('usr_' + email); } catch (x) {}
    return listUsers_();
  });
}
function saveUser(tok, u) {
  var me = me_(tok); need_(me, 'admin');
  var email = needStr_(u.email, 'Nhập email').toLowerCase();
  CacheService.getScriptCache().remove('usr_' + email);
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('Email không hợp lệ');
  if (['Admin', 'Editor', 'Viewer'].indexOf(u.vai_tro) < 0) throw new Error('Vai trò không hợp lệ');
  if (['', 'Tổ', 'Phòng'].indexOf(u.cap_duyet || '') < 0) throw new Error('Cấp duyệt không hợp lệ');
  if (u.vai_tro === 'Viewer' && u.cap_duyet) throw new Error('Viewer không thể có quyền duyệt.');
  return withLock_(function () {
    var rows = readAll_(SHEETS.USERS), ex = rows.filter(function (x) { return String(x.email).toLowerCase() === email; })[0];
    var admins = rows.filter(function (x) { return x.vai_tro === 'Admin' && x.trang_thai !== 'Khóa' && String(x.email).toLowerCase() !== email; });
    if (!admins.length && (u.vai_tro !== 'Admin' || u.trang_thai === 'Khóa')) throw new Error('Phải còn ít nhất một Admin.');
    var o = { email: email, ten: u.ten || '', vai_tro: u.vai_tro, cap_duyet: u.cap_duyet || '', trang_thai: u.trang_thai === 'Khóa' ? 'Khóa' : 'Hoạt động', nhan_mail: u.nhan_mail ? 1 : 0 };
    if (ex) { updateRow_(SHEETS.USERS, ex._row, o); log_(me, [{ hanh_dong: 'Sửa người dùng', truong: email, cu: ex.vai_tro + '/' + ex.cap_duyet + '/' + ex.trang_thai, moi: o.vai_tro + '/' + o.cap_duyet + '/' + o.trang_thai }]); }
    else { o.ngay_them = today_(); appendRows_(SHEETS.USERS, [o]); log_(me, [{ hanh_dong: 'Thêm người dùng', truong: email, cu: '', moi: o.vai_tro + '/' + o.cap_duyet }]); }
    return listUsers_();
  });
}
function saveSettings(tok, obj) {
  var me = me_(tok); need_(me, 'admin');
  var okKeys = ['SAP_HAN_NGAY', 'PENDING_DAI_NGAY', 'KHONG_CAP_NHAT_NGAY', 'LEGACY_EDITABLE', 'EXT_PHONG_NGAY', 'PAGE_SIZE', 'ORG_NAME', 'KHACH_XEM', 'SESSION_DAYS', 'NOTIFY_AUTO', 'REPORT_WEEKLY', 'REPORT_MONTHLY', 'APP_URL'];
  return withLock_(function () {
    var cur = getSettings_(), ents = [];
    okKeys.forEach(function (k) {
      if (obj.hasOwnProperty(k) && String(obj[k]) !== String(cur[k])) {
        if (k !== 'ORG_NAME' && k !== 'APP_URL' && (isNaN(Number(obj[k])) || Number(obj[k]) < 0)) throw new Error(k + ' phải là số không âm');
        setSetting_(k, obj[k]); ents.push({ hanh_dong: 'Đổi cấu hình', truong: k, cu: cur[k], moi: obj[k] });
      }
    });
    log_(me, ents); return { ok: true };
  });
}
/** Admin: bỏ cờ "cần rà soát" sau khi đã kiểm tra */
function resolveReview(tok, id, note) {
  var me = me_(tok); need_(me, 'write');
  return withLock_(function () {
    var r = findReq_(id);
    updateRow_(SHEETS.DATA_MAIN, r._row, { ra_soat: 0, ghi_chu_rs: (r.ghi_chu_rs ? r.ghi_chu_rs + ' | ' : '') + 'Đã rà soát ' + today_() + (note ? ': ' + note : ''), cn_boi: me.email, cn_luc: now_() });
    log_(me, [{ request_id: id, hanh_dong: 'Rà soát dữ liệu cũ', truong: 'ra_soat', cu: 1, moi: 0 }]);
    return { ok: true };
  });
}

/* ============ Chỉ số (dùng cho nhắc việc & báo cáo) ============ */
function metrics_(r, t, cfg) {
  var sap = cfg.sap, open = r.tt !== 'Closed';
  var d = r.deadline;
  return {
    quaHan: open && d && d < t,
    sapHan: open && d && d >= t && diffDays_(d, t) <= sap,
    pendingDai: r.tt === 'Pending' && r.cn_luc && diffDays_(t, String(r.cn_luc).slice(0, 10)) > cfg.pend,
    khongCapNhat: open && r.cn_luc && diffDays_(t, String(r.cn_luc).slice(0, 10)) > cfg.kcn
  };
}
function cfg_() { return { sap: num_('SAP_HAN_NGAY', 3), pend: num_('PENDING_DAI_NGAY', 7), kcn: num_('KHONG_CAP_NHAT_NGAY', 7) }; }

/* ============ Email nhắc việc ============ */
function mailLogged_(key, to, loai, reqId, subject, html, attachments) {
  var log = readAll_(SHEETS.EMAIL_LOG);
  if (log.some(function (l) { return l.key === key; })) return false;
  var res = 'OK';
  try { var mo = { to: to, subject: subject, htmlBody: html, name: 'Quản lý yêu cầu SC-BD' }; if (attachments && attachments.length) mo.attachments = attachments; MailApp.sendEmail(mo); } catch (e) { res = 'LỖI: ' + e.message; }
  appendRows_(SHEETS.EMAIL_LOG, [{ key: key, thoi_gian: now_(), den: to, loai: loai, request_id: reqId, tieu_de: subject, ket_qua: res }]);
  return res === 'OK';
}
function approverEmails_(cap) {
  return readAll_(SHEETS.USERS).filter(function (u) {
    return u.trang_thai !== 'Khóa' && (u.vai_tro === 'Admin' || u.vai_tro === 'Editor') && (u.cap_duyet === 'Phòng' || (cap !== 'Phòng' && u.cap_duyet === 'Tổ'));
  }).map(function (u) { return u.email; });
}
function mailExtRequest_(e, r) {
  var to = approverEmails_(e.cap_duyet_can).filter(function (x) { return String(x).toLowerCase() !== String(e.yeu_cau_boi).toLowerCase(); });
  if (!to.length) return;
  mailLogged_('EXT|' + e.ext_id, to.join(','), 'Đề nghị gia hạn', r.request_id, '[Gia hạn] ' + r.request_id + ' chờ ' + e.cap_duyet_can + ' duyệt',
    '<p>' + esc_(e.yeu_cau_boi) + ' đề nghị gia hạn <b>' + esc_(r.request_id) + '</b> từ ' + e.deadline_cu + ' đến <b>' + e.deadline_moi + '</b>.</p><p>Lý do: ' + esc_(e.ly_do) + '</p><p>Nội dung: ' + esc_(r.noi_dung) + '</p><p>Vui lòng mở ứng dụng để duyệt.</p>');
}
function mailExtDecision_(e, r, ok, note, me) {
  mailLogged_('EXTD|' + e.ext_id, e.yeu_cau_boi, 'Kết quả gia hạn', r.request_id, '[Gia hạn] ' + r.request_id + (ok ? ' đã được duyệt' : ' bị từ chối'),
    '<p>' + esc_(me.email) + (ok ? ' đã duyệt' : ' đã từ chối') + ' gia hạn <b>' + esc_(r.request_id) + '</b> đến ' + e.deadline_moi + '.</p>' + (note ? '<p>Ghi chú: ' + esc_(note) + '</p>' : ''));
}
function esc_(s) { return String(s === undefined || s === null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }

/** Trigger hằng ngày: 1 thư tổng hợp/người/ngày (khóa chống gửi trùng bằng EMAIL_LOG). Bỏ qua dữ liệu cũ. */
function dailyDigest() {
  var t = today_(), c = cfg_();
  var reqs = readAll_(SHEETS.DATA_MAIN).filter(function (r) { return String(r.legacy) !== '1' && r.tt !== 'Closed'; });
  var groups = { 'Ưu tiên khẩn': [], 'Quá hạn': [], 'Sắp đến hạn': [], 'Pending kéo dài': [], 'Chưa cập nhật': [] };
  reqs.forEach(function (r) {
    var m = metrics_(r, t, c);
    if (r.uu_tien === 'Khẩn') groups['Ưu tiên khẩn'].push(r);
    if (m.quaHan) groups['Quá hạn'].push(r); else if (m.sapHan) groups['Sắp đến hạn'].push(r);
    if (m.pendingDai) groups['Pending kéo dài'].push(r);
    if (m.khongCapNhat && !m.quaHan) groups['Chưa cập nhật'].push(r);
  });
  var extWait = readAll_(SHEETS.EXTENSIONS).filter(function (e) { return e.trang_thai === 'Chờ duyệt'; });
  var users = readAll_(SHEETS.USERS).filter(function (u) { return u.trang_thai !== 'Khóa' && String(u.nhan_mail) === '1'; });
  var sent = 0;
  users.forEach(function (u) {
    var html = '', total = 0;
    Object.keys(groups).forEach(function (g) {
      if (!groups[g].length) return; total += groups[g].length;
      html += '<h3>' + g + ' (' + groups[g].length + ')</h3><ul>' + groups[g].slice(0, 40).map(function (r) {
        return '<li><b>' + esc_(r.request_id) + '</b> – ' + esc_(String(r.noi_dung).slice(0, 90)) + ' – hạn ' + r.deadline + ' – ' + esc_(r.dv_gui) + '</li>'; }).join('') + '</ul>';
    });
    if ((u.cap_duyet === 'Tổ' || u.cap_duyet === 'Phòng') && extWait.length) {
      total += extWait.length; html += '<h3>Đề nghị gia hạn chờ duyệt (' + extWait.length + ')</h3><ul>' + extWait.map(function (e) { return '<li>' + esc_(e.request_id) + ' → ' + e.deadline_moi + ' (' + esc_(e.ly_do) + ')</li>'; }).join('') + '</ul>';
    }
    if (!total) return;
    if (mailLogged_('DIGEST|' + t + '|' + u.email, u.email, 'Tổng hợp ngày', '', '[Nhắc việc] ' + t + ': ' + total + ' mục cần chú ý', html)) sent++;
  });
  return sent;
}

/* ============ Báo cáo tuần / tháng (Word / Excel) ============ */
function reportData_(kind, ref) {
  var from, to;
  if (kind === 'week') { var d = new Date(ref + 'T00:00:00Z'), wd = (d.getUTCDay() + 6) % 7; from = addDays_(ref, -wd); to = addDays_(from, 6); }
  else { from = ref.slice(0, 7) + '-01'; var y = +ref.slice(0, 4), m = +ref.slice(5, 7); to = new Date(Date.UTC(y, m, 0)).toISOString().slice(0, 10); }
  var t = today_(), c = cfg_();
  var all = readAll_(SHEETS.DATA_MAIN);
  var live = all.filter(function (r) { return String(r.legacy) !== '1'; });
  var moi = live.filter(function (r) { return r.ngay_tn >= from && r.ngay_tn <= to; });
  var xong = all.filter(function (r) { return r.tt === 'Closed' && r.ngay_xong && r.ngay_xong >= from && r.ngay_xong <= to; });
  var ton = all.filter(function (r) { return r.tt !== 'Closed'; });
  var qua = ton.filter(function (r) { return metrics_(r, t, c).quaHan; });
  var tk = {}; moi.forEach(function (r) { var k = r.loai + ' / ' + r.dv_tn; tk[k] = (tk[k] || 0) + 1; });
  return { from: from, to: to, moi: moi, xong: xong, ton: ton, qua: qua, tk: tk, soLegacyTon: ton.filter(function (r) { return String(r.legacy) === '1'; }).length };
}
function rowsOf_(list) {
  return list.map(function (r, i) { return [i + 1, r.request_id, r.ngay_tn, r.loai, r.dv_gui, r.dv_tn, r.noi_dung, r.ma_bravo, r.deadline, r.tt, r.ngay_xong]; });
}
var RH = ['STT', 'Mã YC', 'Ngày TN', 'Loại', 'Đơn vị gửi', 'Đơn vị TN', 'Nội dung', 'Mã Bravo', 'Hạn', 'Trạng thái', 'Ngày xong'];
function exportReport(tok, kind, fmt, ref) {
  var me = me_(tok); need_(me, 'read');
  if (me.guest) throw new Error('Hãy đăng nhập (Admin/Editor) để xuất báo cáo.');
  if (['week', 'month'].indexOf(kind) < 0 || ['xlsx', 'docx'].indexOf(fmt) < 0) throw new Error('Tham số báo cáo không hợp lệ');
  ref = checkDate_(ref || today_(), 'Ngày tham chiếu');
  var d = reportData_(kind, ref), org = getSetting_('ORG_NAME') || 'Trung tâm Bảo dưỡng Nội trường Hà Nội';
  var title = (kind === 'week' ? 'BÁO CÁO TUẦN' : 'BÁO CÁO THÁNG') + ' – ' + d.from + ' đến ' + d.to;
  var fileId, name = (kind === 'week' ? 'BaoCao_Tuan_' : 'BaoCao_Thang_') + d.from + '.' + fmt;
  if (fmt === 'xlsx') {
    var wb = SpreadsheetApp.create('tmp_' + name); fileId = wb.getId();
    var sections = [['Tiếp nhận mới', d.moi], ['Hoàn thành trong kỳ', d.xong], ['Quá hạn chưa đóng', d.qua], ['Tồn đọng (chưa đóng)', d.ton]];
    var first = wb.getSheets()[0]; first.setName('Tổng quan');
    var sum = [[org], [title], [''], ['Chỉ tiêu', 'Số lượng'], ['Tiếp nhận mới (không gồm dữ liệu cũ)', d.moi.length], ['Hoàn thành trong kỳ', d.xong.length], ['Tồn đọng (gồm ' + d.soLegacyTon + ' việc từ dữ liệu cũ)', d.ton.length], ['Quá hạn chưa đóng', d.qua.length], [''], ['Tiếp nhận mới theo loại / đơn vị tiếp nhận', '']];
    Object.keys(d.tk).forEach(function (k) { sum.push([k, d.tk[k]]); });
    sum = sum.map(function (r) { return r.length < 2 ? r.concat(['']) : r; });
    first.getRange(1, 1, sum.length, 2).setValues(sum); first.getRange(1, 1).setFontWeight('bold'); first.setColumnWidth(1, 380);
    sections.forEach(function (s) {
      var sheet = wb.insertSheet(s[0]), rows = [RH].concat(rowsOf_(s[1]));
      sheet.getRange(1, 1, rows.length, RH.length).setValues(rows);
      sheet.getRange(1, 1, 1, RH.length).setFontWeight('bold').setBackground('#1F4E78').setFontColor('#ffffff');
      sheet.setFrozenRows(1); sheet.setColumnWidth(7, 380);
    });
    SpreadsheetApp.flush();
  } else {
    var doc = DocumentApp.create('tmp_' + name); fileId = doc.getId();
    var b = doc.getBody();
    b.appendParagraph(org).setBold(true).setAlignment(DocumentApp.HorizontalAlignment.LEFT);
    b.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.HEADING1).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    b.appendParagraph('Tiếp nhận mới: ' + d.moi.length + ' | Hoàn thành: ' + d.xong.length + ' | Tồn đọng: ' + d.ton.length + ' (trong đó ' + d.soLegacyTon + ' từ dữ liệu cũ) | Quá hạn: ' + d.qua.length);
    var keys = Object.keys(d.tk);
    if (keys.length) { b.appendParagraph('Tiếp nhận mới theo loại / đơn vị tiếp nhận').setBold(true); keys.forEach(function (k) { b.appendListItem(k + ': ' + d.tk[k]); }); }
    [['1. Tiếp nhận mới', d.moi], ['2. Hoàn thành trong kỳ', d.xong], ['3. Quá hạn chưa đóng', d.qua]].forEach(function (s) {
      b.appendParagraph(s[0] + ' (' + s[1].length + ')').setHeading(DocumentApp.ParagraphHeading.HEADING2);
      if (!s[1].length) { b.appendParagraph('Không có.'); return; }
      var rows = [['STT', 'Mã YC', 'Nội dung', 'Đơn vị gửi', 'Hạn', 'Trạng thái']].concat(s[1].slice(0, 200).map(function (r, i) { return [String(i + 1), r.request_id, r.noi_dung, r.dv_gui, r.deadline, r.tt]; }));
      b.appendTable(rows).setFontSize(9);
    });
    b.appendParagraph('Ngày lập: ' + today_() + ' – Người lập: ' + me.ten);
    doc.saveAndClose();
  }
  var mime = fmt === 'xlsx' ? 'xlsx' : 'docx';
  var url = fmt === 'xlsx' ? 'https://docs.google.com/spreadsheets/d/' + fileId + '/export?format=xlsx' : 'https://docs.google.com/document/d/' + fileId + '/export?format=docx';
  var resp = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  DriveApp.getFileById(fileId).setTrashed(true);
  if (resp.getResponseCode() !== 200) throw new Error('Không xuất được tệp (' + resp.getResponseCode() + ')');
  return { name: name, b64: Utilities.base64Encode(resp.getContent()) };
}


/* ============ Cập nhật lại dữ liệu cũ từ 2 file Excel đã làm sạch ============ */
var MON_ = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
var STD_DV_ = { PKT: 'Phòng KT', SCCT: 'Đội SCCT', PVBD: 'Đội PVBD' };
function iso_(y, m, d) {
  if (y < 100) y += 2000;
  var dt = new Date(Date.UTC(y, m - 1, d));
  if (dt.getUTCFullYear() !== y || dt.getUTCMonth() !== m - 1 || dt.getUTCDate() !== d) return null;
  return dt.toISOString().slice(0, 10);
}
/** Trả về {v:'yyyy-MM-dd'|'', err:true|false} */
function parseDateAny_(v) {
  if (v === '' || v === null || v === undefined) return { v: '', err: false };
  if (v instanceof Date) return { v: Utilities.formatDate(v, tz_(), 'yyyy-MM-dd'), err: false };
  var s = String(v).trim(), m, r;
  if ((m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/))) r = iso_(+m[1], +m[2], +m[3]);
  else if ((m = s.match(/^(\d{1,2})[-\/ ]*([A-Za-z]{3})[a-z]*[-\/ ]*(\d{2,4})$/)) && MON_[m[2].toLowerCase()]) r = iso_(+m[3], MON_[m[2].toLowerCase()], +m[1]);
  else if ((m = s.match(/^(\d{1,2})[-\/ ]*thg\s*(\d{1,2})[-\/ ]*(\d{2,4})$/i))) r = iso_(+m[3], +m[2], +m[1]);
  else if ((m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/))) r = iso_(+m[3], +m[2], +m[1]);
  return r ? { v: r, err: false } : { v: '', err: true };
}
function doiOf_(s) { var m = String(s || '').trim().toUpperCase().match(/^ĐỘI\s*(\d)$/); return m ? 'Đội ' + m[1] : null; }
function padStt_(n) { return ('0000' + n).slice(-4); }
function legacyRecords_(errors, msgs) {
  var recs = [], pdn = [['VP', 'Văn phòng'], ['PV', DV_PV]];
  pdn.forEach(function (p) {
    var name = 'NHAP_CU_' + p[0], s = ss_().getSheetByName(name);
    if (!s) { s = ss_().insertSheet(name); msgs.push('Đã tạo sheet ' + name + ' (đang trống).'); return; }
    var v = s.getDataRange().getValues(), hr = -1;
    for (var i = 0; i < v.length && i < 30; i++) if (String(v[i][0]).trim().toUpperCase() === 'STT') { hr = i; break; }
    if (hr < 0) { if (v.length > 1 || (v[0] && v[0][0] !== '')) errors.push(name + ': không tìm thấy dòng tiêu đề (ô A có chữ "STT").'); else msgs.push('Sheet ' + name + ' trống, bỏ qua.'); return; }
    var seen = {}, n = 0;
    for (var k = hr + 1; k < v.length; k++) {
      var r = v[k]; if (!String(r[4] === undefined ? '' : r[4]).trim()) continue;
      n++;
      var stt = Number(r[0]), ma = p[0] + (stt > 0 ? padStt_(stt) : 'R' + padStt_(n));
      seen[ma] = (seen[ma] || 0) + 1; if (seen[ma] > 1) ma += '-' + seen[ma];
      var yc = String(r[7] || '').trim(), pl = String(r[2] || '').trim(), note = [], dv;
      if (STD_DV_[yc]) dv = STD_DV_[yc];
      else if (yc === 'Doi BD') { dv = doiOf_(pl); if (!dv) { dv = STD_DV_[pl] || pl; note.push('Doi BD nhưng cột Phân loại đơn vị không phải Đội 1-6'); } }
      else { dv = yc; note.push('đơn vị lạ'); }
      var cp = doiOf_(pl) || STD_DV_[pl];
      if (yc !== 'Doi BD' && cp && cp !== dv) note.push('Đvi yc=' + yc + ' khác Phân loại đơn vị=' + pl);
      var d1 = parseDateAny_(r[3]), d2 = parseDateAny_(r[10]);
      if (d1.err) { note.push('Không đọc được ngày làm yêu cầu: ' + r[3]); errors.push(name + ' dòng ' + (k + 1) + ': ngày làm yêu cầu "' + r[3] + '"'); }
      if (d2.err) { note.push('Không đọc được ngày dự kiến hoàn thành: ' + r[10]); errors.push(name + ' dòng ' + (k + 1) + ': hạn "' + r[10] + '"'); }
      if (d1.v && d2.v && d2.v < d1.v) note.push('Hạn (' + d2.v + ') trước ngày làm yêu cầu (' + d1.v + ')');
      var tt = String(r[6] || '').trim(); if (tt === 'Close') tt = 'Closed';
      if (STATUSES.indexOf(tt) < 0) { note.push('Tình trạng lạ: ' + tt); errors.push(name + ' dòng ' + (k + 1) + ': tình trạng "' + tt + '"'); tt = 'Inprogress'; }
      var gui = r[12] === undefined || r[12] === '' ? '' : (r[12] instanceof Date ? Utilities.formatDate(r[12], tz_(), 'dd/MM/yyyy') : String(r[12]).trim());
      recs.push({ request_id: 'L-' + ma, ma_cu: ma, ngay_tn: d1.v, nguoi_gui: r[9] ? String(r[9]).trim() : '', dv_gui: dv,
        noi_dung: String(r[4]).trim(), dv_tn: p[1], ma_bravo: r[8] ? String(r[8]).trim() : '', nhom: r[5] ? String(r[5]).trim() : '',
        deadline_goc: d2.v, deadline: d2.v, tt: tt, ghi_chu: gui ? 'Ngày gửi YC (file cũ): ' + gui : '',
        pl_cu: String(r[1] || '').trim() + ' | ' + pl + ' | Đvi yc: ' + yc, _note: note });
    }
  });
  var bc = {}; recs.forEach(function (x) { if (x.ma_bravo) bc[x.ma_bravo] = (bc[x.ma_bravo] || 0) + 1; });
  recs.forEach(function (x) {
    if (x.ma_bravo && bc[x.ma_bravo] > 1) x._note.push('Mã Bravo trùng (' + bc[x.ma_bravo] + ' dòng)');
    x.ra_soat = x._note.length ? 1 : 0; x.ghi_chu_rs = x._note.join('; '); delete x._note;
  });
  return recs;
}
/**
 * Admin: dán dữ liệu 2 file Excel (giữ nguyên bố cục, kể cả dòng tiêu đề có chữ STT) vào sheet NHAP_CU_VP và NHAP_CU_PV,
 * rồi chạy với apply=false để xem trước, apply=true để áp dụng.
 * Ghép theo mã cũ (VP0021...): dòng đã có -> cập nhật nếu khác; dòng mới -> thêm; dòng không còn trong file -> giữ lại và gắn cờ rà soát.
 * Trao đổi, đính kèm, nhật ký của yêu cầu được giữ nguyên.
 */
function importLegacy(tok, apply) {
  var me = me_(tok); need_(me, 'admin');
  var errors = [], msgs = [], recs = legacyRecords_(errors, msgs);
  var run = function () {
    var s = sh_(SHEETS.DATA_MAIN), v = s.getDataRange().getValues(), H = HEADERS.DATA_MAIN, ix = {};
    H.forEach(function (h, i) { ix[h] = i; });
    var byMa = {}, rows = [];
    for (var i = 1; i < v.length; i++) {
      var row = v[i].map(function (c, j) { return norm_(c, H[j]); });
      while (row.length < H.length) row.push('');
      rows.push(row);
      if (String(row[ix.legacy]) === '1') byMa[row[ix.ma_cu]] = rows.length - 1;
    }
    var res = { moi: 0, capNhat: 0, giuNguyen: 0, khongCon: 0, loi: errors, thongBao: msgs, mau: [], tongFile: recs.length };
    var F = ['ngay_tn', 'nguoi_gui', 'dv_gui', 'noi_dung', 'dv_tn', 'ma_bravo', 'nhom', 'deadline_goc', 'deadline', 'tt', 'ghi_chu', 'pl_cu', 'ra_soat', 'ghi_chu_rs'];
    var ents = [], inFile = {}, t = now_();
    recs.forEach(function (x) {
      inFile[x.ma_cu] = 1;
      var at = byMa[x.ma_cu];
      if (at === undefined) {
        res.moi++;
        if (apply) {
          var nr = H.map(function (h) { return x.hasOwnProperty(h) ? x[h] : ''; });
          nr[ix.legacy] = 1; nr[ix.tao_boi] = '(nhập từ file cũ)'; nr[ix.tao_luc] = t; rows.push(nr);
          ents.push({ request_id: x.request_id, hanh_dong: 'Nhập dữ liệu cũ', truong: 'tạo mới', cu: '', moi: x.ma_cu });
        }
        if (res.mau.length < 8) res.mau.push(x.request_id + ': thêm mới');
        return;
      }
      var ex = rows[at], upd = {}, ch = [];
      F.forEach(function (k) { upd[k] = x[k]; });
      if (ex[ix.ngay_xong]) { upd.tt = ex[ix.tt]; upd.deadline = ex[ix.deadline]; }            // đã đóng trong app: giữ nguyên
      else if (ex[ix.deadline] !== ex[ix.deadline_goc]) upd.deadline = ex[ix.deadline];       // đã gia hạn được duyệt: giữ hạn mới
      if (String(ex[ix.ra_soat]) === '0' && x.ghi_chu_rs && String(ex[ix.ghi_chu_rs]).indexOf(x.ghi_chu_rs) === 0) { upd.ra_soat = 0; upd.ghi_chu_rs = ex[ix.ghi_chu_rs]; } // đã rà soát, không đổi
      F.forEach(function (k) { if (String(ex[ix[k]]) !== String(upd[k])) ch.push(k); });
      if (!ch.length) { res.giuNguyen++; return; }
      res.capNhat++;
      if (res.mau.length < 8) res.mau.push(x.request_id + ': ' + ch.join(', '));
      if (apply) {
        F.forEach(function (k) { ex[ix[k]] = upd[k]; }); ex[ix.cn_boi] = me.email; ex[ix.cn_luc] = t;
        ents.push({ request_id: x.request_id, hanh_dong: 'Cập nhật từ file cũ', truong: ch.join(','), cu: '', moi: ch.length + ' trường' });
      }
    });
    rows.forEach(function (row) {
      if (String(row[ix.legacy]) !== '1' || inFile[row[ix.ma_cu]]) return;
      res.khongCon++;
      if (apply && String(row[ix.ghi_chu_rs]).indexOf('Không còn trong file') < 0) {
        row[ix.ra_soat] = 1; row[ix.ghi_chu_rs] = (row[ix.ghi_chu_rs] ? row[ix.ghi_chu_rs] + '; ' : '') + 'Không còn trong file cập nhật ' + today_();
        ents.push({ request_id: row[ix.request_id], hanh_dong: 'Cập nhật từ file cũ', truong: 'không còn trong file', cu: '', moi: 'gắn cờ rà soát' });
      }
    });
    if (apply) {
      if (errors.length) throw new Error('Còn ' + errors.length + ' lỗi dữ liệu trong file, hãy sửa trước khi áp dụng (bấm Kiểm tra để xem).');
      if (!recs.length) throw new Error('Không có dữ liệu trong NHAP_CU_VP / NHAP_CU_PV.');
      var rg = s.getRange(2, 1, rows.length, H.length); rg.setNumberFormat('@'); rg.setValues(rows); memoClear_();
      log_(me, ents.concat([{ hanh_dong: 'Nhập lại dữ liệu cũ', truong: 'tổng kết', cu: '', moi: 'mới ' + res.moi + ', cập nhật ' + res.capNhat + ', giữ nguyên ' + res.giuNguyen + ', không còn ' + res.khongCon }]));
    }
    res.daApDung = !!apply;
    return res;
  };
  return apply ? withLock_(run) : run();
}

/* ============ Thông báo tới đơn vị liên quan ============ */
function splitEmails_(s) {
  return String(s || '').split(/[,;\s]+/).map(function (x) { return x.trim().toLowerCase(); })
    .filter(function (x) { return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(x); });
}
function unitEmails_(name) {
  var row = readAll_(SHEETS.UNIT_CONTACTS).filter(function (x) { return x.don_vi === name; })[0];
  return row ? splitEmails_(row.emails) : [];
}
function unique_(a) { var s = {}; return a.filter(function (x) { if (s[x]) return false; s[x] = 1; return true; }); }
function appUrl_() { return String(getSetting_('APP_URL') || '').trim(); }
/** Danh bạ đơn vị: Admin xem/sửa email nhận thông báo của từng đơn vị. */
function getContacts(tok) {
  var me = me_(tok); need_(me, 'admin');
  var dm = danhMuc_(), have = {}, out = [];
  readAll_(SHEETS.UNIT_CONTACTS).forEach(function (x) { have[x.don_vi] = x; out.push({ don_vi: x.don_vi, emails: x.emails || '', ghi_chu: x.ghi_chu || '' }); });
  unique_((dm.dv_gui || []).concat(dm.dv_tn || [])).forEach(function (u) { if (u && !have[u]) out.push({ don_vi: u, emails: '', ghi_chu: '' }); });
  return out;
}
function saveContacts(tok, rows) {
  var me = me_(tok); need_(me, 'admin');
  if (!Array.isArray(rows)) throw new Error('Dữ liệu không hợp lệ');
  return withLock_(function () {
    var cur = {}; readAll_(SHEETS.UNIT_CONTACTS).forEach(function (x) { cur[x.don_vi] = x; });
    var ents = [], add = [];
    rows.forEach(function (r) {
      var u = needStr_(r.don_vi, 'Thiếu tên đơn vị'), raw = String(r.emails || '');
      var good = splitEmails_(raw), bad = raw.split(/[,;\s]+/).filter(function (x) { return x && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(x); });
      if (bad.length) throw new Error('Email không hợp lệ ở "' + u + '": ' + bad.join(', '));
      var val = unique_(good).join(', '), ex = cur[u];
      if (ex) { if (String(ex.emails || '') !== val || String(ex.ghi_chu || '') !== String(r.ghi_chu || '')) { updateRow_(SHEETS.UNIT_CONTACTS, ex._row, { emails: val, ghi_chu: String(r.ghi_chu || '') }); ents.push({ hanh_dong: 'Cập nhật liên hệ đơn vị', truong: u, cu: ex.emails, moi: val }); } }
      else if (val || r.ghi_chu) { add.push({ don_vi: u, emails: val, ghi_chu: String(r.ghi_chu || '') }); ents.push({ hanh_dong: 'Thêm liên hệ đơn vị', truong: u, cu: '', moi: val }); }
    });
    if (add.length) appendRows_(SHEETS.UNIT_CONTACTS, add);
    log_(me, ents);
    return { ok: true, doi: ents.length };
  });
}
function notifHtml_(r, heading, msg, who) {
  var url = appUrl_();
  return '<div style="font-family:Segoe UI,Arial,sans-serif;max-width:640px"><h2 style="color:#0a4cc2;margin:0 0 8px">' + esc_(heading) + '</h2>' +
    (msg ? '<p style="white-space:pre-wrap;font-size:15px">' + esc_(msg) + '</p>' : '') +
    '<table style="border-collapse:collapse;width:100%;font-size:14px">' + [
      ['Mã yêu cầu', r.request_id], ['Nội dung', r.noi_dung], ['Loại', r.loai], ['Đơn vị gửi', r.dv_gui], ['Đơn vị tiếp nhận', r.dv_tn],
      ['Ưu tiên', r.uu_tien || 'Bình thường'], ['Hạn hoàn thành', r.deadline], ['Trạng thái', r.tt]
    ].map(function (x) { return '<tr><td style="padding:6px 10px;border:1px solid #dde5f2;background:#f3f7fd;width:150px"><b>' + x[0] + '</b></td><td style="padding:6px 10px;border:1px solid #dde5f2">' + esc_(x[1]) + '</td></tr>'; }).join('') + '</table>' +
    (url ? '<p><a href="' + esc_(url) + '" style="background:#0a66e8;color:#fff;padding:9px 18px;border-radius:20px;text-decoration:none;font-weight:600">Mở hệ thống</a></p>' : '') +
    '<p style="color:#6b7a99;font-size:12px">Gửi bởi ' + esc_(who) + ' qua hệ thống Quản lý yêu cầu SC-BD. Vui lòng không trả lời thư này.</p></div>';
}
function sendMailNow_(to, subject, html, reqId, loai) {
  var res = 'OK';
  try {
    if (typeof MailApp.getRemainingDailyQuota === 'function' && MailApp.getRemainingDailyQuota() < to.length) throw new Error('Hết hạn mức gửi email trong ngày của Google');
    MailApp.sendEmail({ to: to.join(','), subject: subject, htmlBody: html, name: 'Quản lý yêu cầu SC-BD' });
  } catch (e) { res = 'LỖI: ' + e.message; }
  appendRows_(SHEETS.EMAIL_LOG, [{ key: 'NOTI|' + reqId + '|' + now_() + '|' + Math.floor(Math.random() * 1e6), thoi_gian: now_(), den: to.join(','), loai: loai, request_id: reqId, tieu_de: subject, ket_qua: res }]);
  return res;
}
/** Editor/Admin gửi thông báo thủ công tới đơn vị gửi / đơn vị tiếp nhận / email bổ sung. */
function notifyUnits(tok, id, p) {
  var me = me_(tok); need_(me, 'write'); p = p || {};
  var cache = CacheService.getScriptCache(), ck = 'noti_' + me.email, n = Number(cache.get(ck) || 0);
  if (n >= 20) throw new Error('Bạn đã gửi quá 20 thông báo trong 1 giờ. Hãy thử lại sau.');
  var r = findReq_(id), units = [], to = [], missing = [];
  if (p.gui) units.push(r.dv_gui); if (p.tn) units.push(r.dv_tn);
  unique_(units.filter(Boolean)).forEach(function (u) { var em = unitEmails_(u); if (em.length) to = to.concat(em); else missing.push(u); });
  var extra = splitEmails_(p.extra); to = unique_(to.concat(extra));
  if (!to.length) throw new Error(missing.length ? 'Chưa có email liên hệ cho: ' + missing.join(', ') + '. Admin cập nhật tại Quản trị → Liên hệ đơn vị.' : 'Chưa chọn người nhận.');
  var msg = String(p.msg || '').trim(), ro = reqOut_(r, me);
  var subject = '[Thông báo] ' + r.request_id + ' – ' + String(r.noi_dung).slice(0, 70);
  var res = sendMailNow_(to, subject, notifHtml_(ro, 'Thông báo về yêu cầu ' + r.request_id, msg, me.ten || me.email), r.request_id, 'Thông báo thủ công');
  if (res !== 'OK') throw new Error('Không gửi được: ' + res.replace(/^LỖI:\s*/, ''));
  cache.put(ck, String(n + 1), 3600);
  return withLock_(function () {
    var who = unique_(units.filter(Boolean)).concat(extra).join(', ');
    appendRows_(SHEETS.COMMENTS, [{ cm_id: 'CM-' + new Date().getTime(), request_id: id, thoi_gian: now_(), nguoi: me.email, noi_dung: '📣 Đã gửi thông báo tới: ' + who + (msg ? ' — ' + msg : '') }]);
    log_(me, [{ request_id: id, hanh_dong: 'Gửi thông báo', truong: 'thong_bao', cu: '', moi: to.length + ' người nhận (' + who + ')' }]);
    return { ok: true, soNguoi: to.length, thieu: missing };
  });
}
/** Thông báo tự động (bật/tắt bằng NOTIFY_AUTO). Không bao giờ làm hỏng thao tác chính; bỏ qua dữ liệu cũ. */
function notifyAuto_(r, ev, me, x) {
  if (getSetting_('NOTIFY_AUTO') === '0' || String(r.legacy) === '1') return;
  x = x || {};
  var unit = ev === 'new' ? r.dv_tn : r.dv_gui, to = unitEmails_(unit);
  to = to.filter(function (e) { return e !== String(me.email).toLowerCase(); });
  if (!to.length) return;
  var head, msg;
  if (ev === 'new') { head = 'Yêu cầu mới tại ' + r.dv_tn; msg = 'Đã tiếp nhận yêu cầu mới từ ' + r.dv_gui + '. Hạn hoàn thành dự kiến: ' + r.deadline + '.'; }
  else if (ev === 'closed') { head = 'Yêu cầu đã hoàn thành'; msg = 'Yêu cầu của đơn vị đã được đóng (hoàn thành ngày ' + (r.ngay_xong || today_()) + ').'; }
  else if (ev === 'pending') { head = 'Yêu cầu tạm dừng (Pending)'; msg = 'Lý do: ' + (r.ly_do_pending || '') + (x.nm ? '\nĐề nghị gia hạn đến: ' + x.nm : ''); }
  else { head = 'Hạn hoàn thành được điều chỉnh'; msg = 'Hạn hoàn thành đã được duyệt gia hạn từ ' + x.cu + ' sang ' + x.moi + '.'; }
  sendMailNow_(to, '[' + head + '] ' + r.request_id, notifHtml_(reqOut_(r, me), head, msg, me.ten || me.email), r.request_id, 'Tự động: ' + ev);
}

/* ============ Báo cáo định kỳ cho cán bộ Phòng / Tổ ============ */
function reportRecipients_() {
  return readAll_(SHEETS.USERS).filter(function (u) {
    return u.trang_thai !== 'Khóa' && String(u.nhan_mail) === '1' && (u.vai_tro === 'Admin' || u.cap_duyet === 'Tổ' || u.cap_duyet === 'Phòng');
  }).map(function (u) { return String(u.email).toLowerCase(); });
}
function prioRank_(p) { var i = PRIORITIES.indexOf(p); return i < 0 ? 2 : i; }
function csvCell_(v) { v = String(v === undefined || v === null ? '' : v); return /[",\n;]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; }
function buildPeriodReport_(kind, ref) {
  var d = reportData_(kind, ref), t = today_(), cf = cfg_();
  var ton = d.ton.map(function (r) { var m = metrics_(r, t, cf); return { r: r, m: m }; });
  var qua = ton.filter(function (x) { return x.m.quaHan; }).map(function (x) { return x.r; }).sort(function (a, b) { return prioRank_(a.uu_tien) - prioRank_(b.uu_tien) || (a.deadline < b.deadline ? -1 : 1); });
  var sap = ton.filter(function (x) { return x.m.sapHan; }).map(function (x) { return x.r; }).sort(function (a, b) { return a.deadline < b.deadline ? -1 : 1; });
  var khan = d.ton.filter(function (r) { return r.uu_tien === 'Khẩn'; });
  var pend = ton.filter(function (x) { return x.m.pendingDai; }).map(function (x) { return x.r; });
  var extWait = readAll_(SHEETS.EXTENSIONS).filter(function (e) { return e.trang_thai === 'Chờ duyệt'; });
  var byUnit = {};
  ton.forEach(function (x) { var u = x.r.dv_tn || '(trống)'; var o = byUnit[u] = byUnit[u] || { mo: 0, qua: 0 }; o.mo++; if (x.m.quaHan) o.qua++; });
  var th = 'style="padding:6px 8px;border:1px solid #dde5f2;background:#e8f0ff;text-align:left;font-size:12px"', td = 'style="padding:6px 8px;border:1px solid #dde5f2;font-size:13px;vertical-align:top"';
  function box(n, l, col) { return '<td style="padding:12px 16px;background:' + col + ';color:#fff;border-radius:10px;text-align:center;width:16%"><div style="font-size:26px;font-weight:800">' + n + '</div><div style="font-size:12px">' + l + '</div></td><td style="width:6px"></td>'; }
  function list(title, arr, col) {
    if (!arr.length) return '';
    return '<h3 style="color:' + col + ';margin:18px 0 6px">' + title + ' (' + arr.length + ')</h3><table style="border-collapse:collapse;width:100%"><tr><th ' + th + '>Mã</th><th ' + th + '>Nội dung</th><th ' + th + '>Đơn vị gửi</th><th ' + th + '>ĐV tiếp nhận</th><th ' + th + '>Hạn</th><th ' + th + '>Ưu tiên</th></tr>' +
      arr.slice(0, 15).map(function (r) { return '<tr><td ' + td + '><b>' + esc_(r.request_id) + '</b></td><td ' + td + '>' + esc_(String(r.noi_dung).slice(0, 110)) + '</td><td ' + td + '>' + esc_(r.dv_gui) + '</td><td ' + td + '>' + esc_(r.dv_tn) + '</td><td ' + td + '>' + esc_(r.deadline) + '</td><td ' + td + '>' + esc_(r.uu_tien || 'Bình thường') + '</td></tr>'; }).join('') +
      '</table>' + (arr.length > 15 ? '<p style="color:#6b7a99;font-size:12px">… và ' + (arr.length - 15) + ' mục khác (xem tệp CSV đính kèm).</p>' : '');
  }
  var title = (kind === 'week' ? 'Báo cáo tuần ' : 'Báo cáo tháng ') + d.from + ' → ' + d.to, url = appUrl_();
  var html = '<div style="font-family:Segoe UI,Arial,sans-serif;max-width:760px"><div style="background:linear-gradient(120deg,#06214f,#0a4cc2 60%,#14a0ff);color:#fff;padding:18px 22px;border-radius:14px"><div style="font-size:12px;opacity:.85">' + esc_(getSetting_('ORG_NAME') || '') + '</div><div style="font-size:22px;font-weight:800">' + esc_(title) + '</div></div>' +
    '<table style="border-collapse:separate;width:100%;margin:14px 0"><tr>' + box(d.ton.length, 'Đang mở', '#0a56d6') + box(qua.length, 'Quá hạn', '#d71a2b') + box(sap.length, 'Sắp đến hạn', '#ee6a00') + box(khan.length, 'Ưu tiên khẩn', '#b4137a') + box(d.moi.length, 'Tiếp nhận mới', '#0a8fb0') + box(d.xong.length, 'Hoàn thành', '#0b8f4d') + '</tr></table>' +
    (extWait.length ? '<p style="background:#fff4e0;padding:10px 14px;border-radius:10px">⏳ Có <b>' + extWait.length + '</b> đề nghị gia hạn đang chờ duyệt.</p>' : '') +
    list('Quá hạn', qua, '#d71a2b') + list('Ưu tiên khẩn đang mở', khan, '#b4137a') + list('Sắp đến hạn', sap, '#ee6a00') + list('Pending kéo dài', pend, '#b9770e') +
    '<h3 style="margin:18px 0 6px">Tồn theo đơn vị tiếp nhận</h3><table style="border-collapse:collapse"><tr><th ' + th + '>Đơn vị</th><th ' + th + '>Đang mở</th><th ' + th + '>Quá hạn</th></tr>' +
    Object.keys(byUnit).map(function (u) { return '<tr><td ' + td + '>' + esc_(u) + '</td><td ' + td + '>' + byUnit[u].mo + '</td><td ' + td + '>' + byUnit[u].qua + '</td></tr>'; }).join('') + '</table>' +
    '<p style="color:#6b7a99;font-size:12px;margin-top:16px">Số tiếp nhận mới không gồm dữ liệu nhập từ file cũ; phần đang mở/quá hạn gồm cả ' + d.soLegacyTon + ' việc từ dữ liệu cũ.</p>' +
    (url ? '<p><a href="' + esc_(url) + '" style="background:#0a66e8;color:#fff;padding:9px 18px;border-radius:20px;text-decoration:none;font-weight:600">Mở hệ thống</a></p>' : '') + '</div>';
  var head = ['Mã', 'Ngày TN', 'Loại', 'Đơn vị gửi', 'Đơn vị TN', 'Nội dung', 'Mã Bravo', 'Hạn', 'Trạng thái', 'Ưu tiên'];
  var csv = '\ufeff' + [head].concat(d.ton.slice().sort(function (a, b) { return prioRank_(a.uu_tien) - prioRank_(b.uu_tien) || (a.deadline < b.deadline ? -1 : 1); }).map(function (r) {
    return [r.request_id, r.ngay_tn, r.loai, r.dv_gui, r.dv_tn, r.noi_dung, r.ma_bravo, r.deadline, r.tt, r.uu_tien || 'Bình thường'];
  })).map(function (row) { return row.map(csvCell_).join(','); }).join('\r\n');
  return { title: title, html: html, csv: csv, from: d.from, so: { mo: d.ton.length, qua: qua.length, khan: khan.length } };
}
function sendPeriodReport_(kind, ref, only) {
  var rep = buildPeriodReport_(kind, ref), to = only ? [only] : reportRecipients_(), sent = 0, stamp = only ? '|' + now_() : '';
  to.forEach(function (em) {
    var att = [Utilities.newBlob(rep.csv, 'text/csv', 'ton_dong_' + today_() + '.csv')];
    if (mailLogged_('RPT|' + kind + '|' + rep.from + '|' + em + stamp, em, 'Báo cáo ' + (kind === 'week' ? 'tuần' : 'tháng'), '', '[Báo cáo] ' + rep.title, rep.html, att)) sent++;
  });
  return { sent: sent, title: rep.title, nguoiNhan: to.length };
}
/** Trigger: sáng thứ Hai gửi báo cáo tuần trước. */
function reportWeekly() { if (getSetting_('REPORT_WEEKLY') === '0') return 0; return sendPeriodReport_('week', addDays_(today_(), -1)).sent; }
/** Trigger: sáng ngày 1 gửi báo cáo tháng trước. */
function reportMonthly() { if (getSetting_('REPORT_MONTHLY') === '0') return 0; return sendPeriodReport_('month', addDays_(today_(), -1)).sent; }
/** Admin: gửi ngay để thử (mode 'me' = chỉ cho mình) hoặc gửi cho toàn bộ cán bộ Phòng/Tổ ('all', kỳ hiện tại). */
function sendReportNow(tok, kind, mode) {
  var me = me_(tok); need_(me, 'admin');
  if (['week', 'month'].indexOf(kind) < 0) throw new Error('Loại báo cáo không hợp lệ');
  var r = sendPeriodReport_(kind, today_(), mode === 'all' ? null : me.email);
  log_(me, [{ hanh_dong: 'Gửi báo cáo ngay', truong: kind, cu: '', moi: mode === 'all' ? 'tất cả (' + r.nguoiNhan + ')' : me.email }]);
  return r;
}

/* ============ Nâng cấp cấu trúc dữ liệu (tự chạy 1 lần cho mỗi phiên bản) ============ */
function ensureSchema_() {
  var sp = PropertiesService.getScriptProperties();
  if (!sp.getProperty('SS_ID') || sp.getProperty('SCHEMA') === SCHEMA_VER) return;
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    if (sp.getProperty('SCHEMA') === SCHEMA_VER) return;
    migrate_();
    sp.setProperty('SCHEMA', SCHEMA_VER);
  } finally { lock.releaseLock(); }
}
/** Tạo sheet còn thiếu, thêm cột còn thiếu vào cuối (không đụng dữ liệu cũ), gieo danh bạ đơn vị, cài trigger báo cáo. */
function migrate_() {
  var ss = ss_();
  Object.keys(HEADERS).forEach(function (n) {
    var s = ss.getSheetByName(n), h = HEADERS[n];
    if (!s) { s = ss.insertSheet(n); s.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold'); s.setFrozenRows(1); return; }
    var last = Math.max(1, s.getLastColumn()), cur = s.getRange(1, 1, 1, last).getValues()[0].map(String);
    var add = h.filter(function (k) { return cur.indexOf(k) < 0; });
    if (add.length) s.getRange(1, cur.filter(function (x) { return x !== ''; }).length + 1, 1, add.length).setValues([add]).setFontWeight('bold');
  });
  memoClear_();
  if (!readAll_(SHEETS.UNIT_CONTACTS).length) {
    var dm = danhMuc_(), seen = {}, rows = [];
    (dm.dv_gui || []).concat(dm.dv_tn || []).forEach(function (u) { if (u && !seen[u]) { seen[u] = 1; rows.push({ don_vi: u, emails: '', ghi_chu: '' }); } });
    if (rows.length) appendRows_(SHEETS.UNIT_CONTACTS, rows);
  }
  try { installTriggers(); } catch (e) {}
}

/* ============ Cài đặt (Admin chạy 1 lần trong trình soạn thảo) ============ */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Hãy tạo dự án Apps Script từ chính file Google Sheet (Tiện ích mở rộng > Apps Script).');
  PropertiesService.getScriptProperties().setProperty('SS_ID', ss.getId());
  SS_ = null; memoClear_();
  migrate_(); PropertiesService.getScriptProperties().setProperty('SCHEMA', SCHEMA_VER);
  var folderId = getSetting_('DRIVE_FOLDER_ID');
  if (!folderId) { var f = DriveApp.createFolder('Dinh_kem_Yeu_cau_SC_BD'); setSetting_('DRIVE_FOLDER_ID', f.getId(), 'Thư mục Drive chứa đính kèm'); }
  var me = Session.getEffectiveUser().getEmail().toLowerCase();
  if (!readAll_(SHEETS.USERS).length) appendRows_(SHEETS.USERS, [{ email: me, ten: 'Admin', vai_tro: 'Admin', cap_duyet: 'Tổ', trang_thai: 'Hoạt động', nhan_mail: 1, ngay_them: today_() }]);
  Logger.log('Xong. Admin đầu tiên: ' + me + '. Hãy Deploy > New deployment > Web app.');
}
function installTriggers() {
  var mine = ['dailyDigest', 'reportWeekly', 'reportMonthly'];
  ScriptApp.getProjectTriggers().forEach(function (t) { if (mine.indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('dailyDigest').timeBased().everyDays(1).atHour(7).create();
  ScriptApp.newTrigger('reportWeekly').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).create();
  ScriptApp.newTrigger('reportMonthly').timeBased().onMonthDay(1).atHour(7).create();
}
