// إعداد سجل اللوندري الرقمي — يُشغَّل مرة واحدة
const DEFS = [
  { name: 'المستخدمون',
    headers: ['المعرف', 'الاسم', 'الدور', 'المجموعة', 'الرمز المشفّر', 'فعّال'],
    rows: [] },
  { name: 'الجهات',
    headers: ['المعرف', 'الاسم', 'الرمز في السجل', 'النوع', 'ضمن العنايات', 'الترتيب', 'فعّال'],
    rows: [
      ['R01', 'الطوابق (1–4)', '', 'طوابق', 'لا', 1, 'نعم'],
      ['R02', 'الطوارئ', 'G', 'طوارئ', 'لا', 2, 'نعم'],
      ['R03', 'العناية المركزة ICU', '', 'عناية', 'نعم', 3, 'نعم'],
      ['R04', 'العناية المتوسطة HDU', '', 'عناية', 'نعم', 4, 'نعم'],
      ['R05', 'العناية القلبية CCU', '', 'عناية', 'نعم', 5, 'نعم'],
      ['R06', 'غسيل الكلى (الديلزة)', 'B', 'غسيل كلى', 'لا', 6, 'نعم'],
      ['R07', 'نوم الموظفين', '', 'موظفون', 'لا', 7, 'نعم'],
      ['R08', 'التمريض', '', 'تمريض', 'لا', 8, 'نعم'],
      ['R09', 'العمليات - طابق 2', '', 'عمليات', 'لا', 9, 'نعم'],
      ['R10', 'العمليات - طابق 3', '', 'عمليات', 'لا', 10, 'نعم']
    ] },
  { name: 'الأصناف',
    headers: ['المعرف', 'الاسم', 'الوحدة', 'يدخل في الخزين', 'فعّال'],
    rows: [
      ['I01', 'بطانية', 'قطعة', 'نعم', 'نعم'],
      ['I02', 'وسادة', 'قطعة', 'نعم', 'نعم'],
      ['I03', 'شرشف', 'قطعة', 'نعم', 'نعم'],
      ['I04', 'بدلة عمليات', 'قطعة', 'لا', 'نعم']
    ] },
  { name: 'الوجبات',
    headers: ['معرف الوجبة', 'التاريخ', 'النوع', 'المجموعة', 'المدوّن', 'الحالة', 'وقت الفتح', 'وقت القفل', 'وقت الإرسال'],
    rows: [] },
  { name: 'الحركات',
    headers: ['المعرف', 'معرف الوجبة', 'الوقت', 'النوع', 'الجهة', 'الصنف', 'العدد', 'اسم موظف الجهة', 'المدخل', 'ملاحظة'],
    rows: [] },
  { name: 'الغسلات',
    headers: ['المعرف', 'معرف الوجبة', 'الوقت', 'الغسالة', 'المدخل'],
    rows: [] },
  { name: 'الجرد',
    headers: ['المعرف', 'معرف الوجبة', 'النوع', 'الصنف', 'العدد الفعلي', 'المتوقع', 'الفرق', 'المدخل', 'الوقت'],
    rows: [] },
  { name: 'التصحيحات',
    headers: ['المعرف', 'الوقت', 'الورقة', 'معرف السجل', 'الحقل', 'القيمة السابقة', 'القيمة الجديدة', 'السبب', 'من عدّل', 'من اعتمد'],
    rows: [] },
  { name: 'الإعدادات',
    headers: ['المفتاح', 'القيمة', 'الوصف'],
    rows: [
      ['بداية الوجبة الصباحية', '07:00', 'الصباحية من 7:00 إلى 19:00'],
      ['بداية الوجبة الليلية', '19:00', 'الليلية تُنسب ليوم بدايتها'],
      ['دقائق الإرسال بعد نهاية الوجبة', '5', 'موعد إرسال التقرير وقفل الوجبة، وهي أيضاً مهلة الإدخال المتأخر'],
      ['دقائق التذكير قبل نهاية الوجبة', '45', 'تذكير البوت بما لم يُسجَّل (0 = إيقاف)'],
      ['ساعات بلا إدخال قبل التذكير', '3', 'تذكير البوت عند توقف الإدخال (0 = إيقاف)'],
      ['بطانيات لكل غسلة', '6', 'سعة الغسالة الواحدة'],
      ['شراشف لكل غسلة (تقدير)', '60', 'لتقدير غسلات الشراشف في اللوحة فقط'],
      ['عدد الغسالات', '3', ''],
      ['المنطقة الزمنية', 'Asia/Baghdad', '']
    ] }
];

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const names = DEFS.map(d => d.name);

  DEFS.forEach(def => {
    let sh = ss.getSheetByName(def.name) || ss.insertSheet(def.name);
    sh.setRightToLeft(true);
    if (sh.getLastRow() > 0) return; // لا نمسّ ورقة فيها بيانات
    const cols = def.headers.length;
    sh.getRange(1, 1, 1, cols).setValues([def.headers])
      .setFontWeight('bold').setBackground('#1f3a5f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.getRange(1, 1, sh.getMaxRows(), cols).setNumberFormat('@'); // نص كما هو
    if (def.rows.length) sh.getRange(2, 1, def.rows.length, cols).setValues(def.rows);
    sh.autoResizeColumns(1, cols);
  });

  // حذف الورقة الافتراضية الفارغة
  ss.getSheets().forEach(sh => {
    if (!names.includes(sh.getName()) && sh.getLastRow() === 0 && ss.getSheets().length > 1) {
      ss.deleteSheet(sh);
    }
  });

  ss.toast('تم إنشاء أوراق السجل بنجاح', 'سجل اللوندري', 10);
  Logger.log('تم: ' + names.join('، '));
}

// ===== القائمة وإضافة المستخدمين =====
function onOpen() {
  SpreadsheetApp.getUi().createMenu('سجل اللوندري')
    .addItem('إضافة مستخدم', 'addUser')
    .addItem('ربط مجموعة Telegram', 'linkTelegram')
    .addItem('إرسال تجريبي لتقرير الوجبة الحالية', 'sendTestReport')
    .addItem('تفعيل الإرسال التلقائي', 'enableAutoSend')
    .addToUi();
}

// تحويل الأرقام العربية (٠١٢٣) إلى (0123)
function toLatin_(s) {
  return String(s).replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d));
}

function hashPin_(id, pin) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, id + ':' + pin, Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

function addUser() {
  const ui = SpreadsheetApp.getUi();
  const ask = q => {
    const r = ui.prompt('إضافة مستخدم', q, ui.ButtonSet.OK_CANCEL);
    if (r.getSelectedButton() !== ui.Button.OK) throw new Error('أُلغيت العملية');
    return toLatin_(r.getResponseText().trim());
  };
  try {
    const name = ask('اسم المستخدم الكامل:');
    if (!name) throw new Error('الاسم فارغ');

    const roles = { '1': 'مدوّن', '2': 'مسؤول', '3': 'مدير النظام' };
    const role = roles[ask('الدور: 1 = مدوّن ، 2 = مسؤول ، 3 = مدير النظام')];
    if (!role) throw new Error('رقم الدور غير صحيح');

    let group = '';
    if (role === 'مدوّن') {
      const groups = { '1': 'صباحية أ', '2': 'صباحية ب', '3': 'ليلية أ', '4': 'ليلية ب' };
      group = groups[ask('المجموعة: 1 = صباحية أ ، 2 = صباحية ب ، 3 = ليلية أ ، 4 = ليلية ب')];
      if (!group) throw new Error('رقم المجموعة غير صحيح');
    }

    const pin = ask('الرمز السري (4 أرقام):');
    if (!/^\d{4}$/.test(pin)) throw new Error('الرمز يجب أن يكون 4 أرقام بالضبط');

    const sh = SpreadsheetApp.getActive().getSheetByName('المستخدمون');
    const names = sh.getRange(1, 2, sh.getLastRow(), 1).getValues().flat();
    if (names.includes(name)) throw new Error('هذا الاسم موجود مسبقاً');

    const id = 'U' + Date.now().toString(36).toUpperCase();
    sh.appendRow([id, name, role, group, hashPin_(id, pin), 'نعم']);
    ui.alert('تمت إضافة: ' + name + ' — ' + role + (group ? ' — ' + group : ''));
  } catch (e) {
    ui.alert(e.message);
  }
}

// ===== صفحة الهاتف: الدخول والوجبة الحالية =====
const TZ = 'Asia/Baghdad';
const DAYS = { 1: 'الاثنين', 2: 'الثلاثاء', 3: 'الأربعاء', 4: 'الخميس', 5: 'الجمعة', 6: 'السبت', 7: 'الأحد' };

// ?page=dashboard يفتح لوحة المتابعة للمسؤول، وبدونه صفحة الإدخال
// ALLOWALL: يسمح بعرض الصفحة داخل تطبيق الهاتف المنشور على GitHub Pages
function doGet(e) {
  const dash = e && e.parameter && e.parameter.page === 'dashboard';
  return HtmlService.createHtmlOutputFromFile(dash ? 'Dashboard' : 'Index')
    .setTitle(dash ? 'لوحة متابعة اللوندري' : 'سجل اللوندري')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function sheet_(name) {
  return SpreadsheetApp.getActive().getSheetByName(name);
}

// الصباحية 7:00–19:00، والليلية تُنسب ليوم بدايتها
function currentShift_(at) {
  const now = at || new Date();
  const hour = Number(Utilities.formatDate(now, TZ, 'H'));
  const night = hour >= 19 || hour < 7;
  const base = (night && hour < 7) ? new Date(now.getTime() - 24 * 3600 * 1000) : now;
  return {
    id: Utilities.formatDate(base, TZ, 'yyyy-MM-dd') + (night ? '-N' : '-D'),
    date: Utilities.formatDate(base, TZ, 'yyyy-MM-dd'),
    display: Utilities.formatDate(base, TZ, 'dd/MM/yyyy'),
    day: DAYS[Number(Utilities.formatDate(base, TZ, 'u'))],
    type: night ? 'ليلية' : 'صباحية'
  };
}

function listUsers() {
  const rows = sheet_('المستخدمون').getDataRange().getValues().slice(1);
  return rows.filter(r => r[5] === 'نعم').map(r => r[1]);
}

function login(name, pin) {
  name = String(name).trim();
  pin = toLatin_(String(pin).trim());
  const cache = CacheService.getScriptCache();
  const failKey = 'fail_' + name;
  const fails = Number(cache.get(failKey) || 0);
  if (fails >= 5) throw new Error('محاولات خاطئة كثيرة. حاول بعد 15 دقيقة.');

  const row = sheet_('المستخدمون').getDataRange().getValues().slice(1)
    .find(r => r[1] === name && r[5] === 'نعم');
  if (!row || hashPin_(row[0], pin) !== row[4]) {
    cache.put(failKey, String(fails + 1), 900);
    throw new Error('الاسم أو الرمز غير صحيح');
  }
  cache.remove(failKey);

  const shift = currentShift_();
  if (row[2] === 'مدوّن') ensureShift_(shift, row);

  const token = Utilities.getUuid();
  cache.put('s_' + token, JSON.stringify({ id: row[0], name: row[1], role: row[2] }), 21600);
  return { token: token, name: row[1], role: row[2], group: row[3], shift: shift,
    late: false, grace: graceShift_() };
}

// ===== مهلة الإدخال المتأخر =====
// خلال دقائق الإرسال بعد نهاية الوجبة (وقبل إرسال تقريرها) يمكن إكمال إدخالها،
// فيُنسب الإدخال إلى الوجبة الأصلية لا إلى الوجبة الجديدة
function sinceShiftStart_(now) {
  const h = Number(Utilities.formatDate(now, TZ, 'H'));
  const m = Number(Utilities.formatDate(now, TZ, 'm'));
  return ((h - 7 + 24) % 12) * 60 + m; // دقائق منذ بداية الوجبة الحالية (7:00 أو 19:00)
}

function graceShift_() {
  const now = new Date();
  const delay = Number(setting_('دقائق الإرسال بعد نهاية الوجبة', 5));
  if (sinceShiftStart_(now) >= delay) return null;
  const prev = currentShift_(new Date(now.getTime() - 12 * 3600 * 1000));
  const row = sheet_('الوجبات').getDataRange().getValues().slice(1).find(r => r[0] === prev.id);
  if (row && row[8] !== '') return null; // أُرسل تقريرها
  const startH = prev.type === 'ليلية' ? 7 : 19; // تنتهي السابقة حين تبدأ الحالية
  prev.until = ('0' + (startH + Math.floor(delay / 60))).slice(-2) + ':' + ('0' + (delay % 60)).slice(-2);
  return prev;
}

// الوجبة التي يُسجَّل فيها إدخال هذا المستخدم الآن
function shiftOf_(user) {
  if (!user.late) return currentShift_();
  const g = graceShift_();
  if (!g) throw new Error('انتهت مهلة الإدخال المتأخر للوجبة السابقة وأُرسل تقريرها');
  return g;
}

function setLateMode(token, on) {
  const s = session_(token);
  if (on && !graceShift_()) throw new Error('لا توجد مهلة إدخال متأخر الآن');
  s.late = !!on;
  CacheService.getScriptCache().put('s_' + token, JSON.stringify(s), 21600);
  return { late: s.late, shift: shiftOf_(s), grace: graceShift_() };
}

// تستدعيه الصفحة كل دقيقة لتعرف إن تغيّرت الوجبة أو بدأت المهلة أو انتهت
function getMode(token) {
  const s = session_(token);
  const g = graceShift_();
  if (s.late && !g) return setLateMode(token, false);
  return { late: !!s.late, shift: shiftOf_(s), grace: g };
}

// إنشاء صف الوجبة عند أول دخول فيها
function ensureShift_(shift, user) {
  const sh = sheet_('الوجبات');
  const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().flat();
  if (ids.includes(shift.id)) return;
  const now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
  sh.appendRow([shift.id, shift.date, shift.type, user[3], user[1], 'مفتوحة', now, '', '']);
}

// ===== تسجيل الحركات =====
function session_(token) {
  const s = CacheService.getScriptCache().get('s_' + token);
  if (!s) throw new Error('انتهت الجلسة. سجّل الدخول من جديد.');
  return JSON.parse(s);
}

function userRow_(id) {
  return sheet_('المستخدمون').getDataRange().getValues().slice(1).find(r => r[0] === id);
}

// الأصناف المسموحة لكل نوع جهة
function itemsFor_(type, night) {
  if (type === 'عمليات') return [{ item: 'بدلة عمليات', out: true, in: true }];
  if (type === 'تمريض') return night ? [] : [{ item: 'شرشف', out: true, in: false }];
  return ['بطانية', 'وسادة', 'شرشف'].map(i => ({ item: i, out: true, in: true }));
}

function getDepartments(token) {
  const night = shiftOf_(session_(token)).type === 'ليلية';
  return sheet_('الجهات').getDataRange().getValues().slice(1)
    .filter(r => r[6] === 'نعم')
    .sort((a, b) => Number(a[5]) - Number(b[5]))
    .map(r => ({ id: r[0], name: r[1], code: r[2], items: itemsFor_(r[3], night) }))
    .filter(d => d.items.length);
}

function saveMovements(token, deptId, entries, staffName, clientId) {
  const user = session_(token);
  if (user.role !== 'مدوّن') throw new Error('الإدخال للمدوّنين فقط');
  const cache = CacheService.getScriptCache();
  if (clientId && cache.get('c_' + clientId)) {
    return { saved: 0, duplicate: true, summary: getShiftSummary(token) };
  }

  const shift = shiftOf_(user);
  const dept = sheet_('الجهات').getDataRange().getValues().slice(1)
    .find(r => r[0] === deptId && r[6] === 'نعم');
  if (!dept) throw new Error('الجهة غير موجودة');
  const allowed = itemsFor_(dept[3], shift.type === 'ليلية');

  const now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
  const rows = [];
  (entries || []).forEach((e, i) => {
    const qty = Number(toLatin_(e.qty));
    if (!qty) return;
    if (!Number.isInteger(qty) || qty < 0 || qty > 2000) throw new Error('عدد غير صحيح: ' + e.qty);
    const rule = allowed.find(a => a.item === e.item);
    if (!rule || (e.type === 'صادر' && !rule.out) || (e.type === 'وارد' && !rule.in)) {
      throw new Error('حركة غير مسموحة لهذه الجهة: ' + e.item + ' ' + e.type);
    }
    rows.push(['M' + Date.now().toString(36).toUpperCase() + i, shift.id, now, e.type,
      dept[1], e.item, qty, String(staffName || '').trim(), user.name, '']);
  });
  if (!rows.length) throw new Error('لم تُدخل أي عدد');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    ensureShift_(shift, userRow_(user.id));
    const sh = sheet_('الحركات');
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  } finally {
    lock.releaseLock();
  }
  if (clientId) cache.put('c_' + clientId, '1', 21600);
  return { saved: rows.length, summary: getShiftSummary(token) };
}

// مجاميع الوجبة الحالية لكل جهة وصنف
function getShiftSummary(token) {
  const shiftId = shiftOf_(session_(token)).id;
  const totals = {};
  sheet_('الحركات').getDataRange().getValues().slice(1)
    .filter(r => r[1] === shiftId)
    .forEach(r => {
      const k = r[4] + '|' + r[5];
      totals[k] = totals[k] || { dept: r[4], item: r[5], out: 0, in: 0 };
      totals[k][r[3] === 'صادر' ? 'out' : 'in'] += Number(r[6]) || 0;
    });
  return Object.values(totals);
}

// ===== الغسلات =====
function getWashCounts(token) {
  const id = shiftOf_(session_(token)).id;
  const c = { 1: 0, 2: 0, 3: 0 };
  sheet_('الغسلات').getDataRange().getValues().slice(1)
    .filter(r => r[1] === id)
    .forEach(r => { const w = Number(r[3]); c[w] = (c[w] || 0) + 1; });
  return c;
}

function addWash(token, washer, clientId) {
  const user = session_(token);
  if (user.role !== 'مدوّن') throw new Error('الإدخال للمدوّنين فقط');
  washer = Number(washer);
  if ([1, 2, 3].indexOf(washer) === -1) throw new Error('رقم الغسالة غير صحيح');

  const cache = CacheService.getScriptCache();
  const done = clientId && cache.get('c_' + clientId);
  if (done) return { id: done, counts: getWashCounts(token) };

  const shift = shiftOf_(user);
  const id = 'W' + Date.now().toString(36).toUpperCase();
  const now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    ensureShift_(shift, userRow_(user.id));
    sheet_('الغسلات').appendRow([id, shift.id, now, washer, user.name]);
  } finally {
    lock.releaseLock();
  }
  if (clientId) cache.put('c_' + clientId, id, 21600);
  cache.put('w_' + id, user.id, 60); // مهلة التراجع: دقيقة
  return { id: id, counts: getWashCounts(token) };
}

function undoWash(token, washId) {
  const user = session_(token);
  const cache = CacheService.getScriptCache();
  if (cache.get('w_' + washId) !== user.id) throw new Error('انتهت مهلة التراجع (دقيقة واحدة)');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_('الغسلات');
    const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().flat();
    const i = ids.indexOf(washId);
    if (i > 0) sh.deleteRow(i + 1);
  } finally {
    lock.releaseLock();
  }
  cache.remove('w_' + washId);
  return getWashCounts(token);
}
// ===== الجرد والتسليم =====
const INV_FIELDS = [
  { key: 'بطانيات جاهزة', type: 'نهاية', label: 'البطانيات الجاهزة في المخزن', hint: 'تُسلَّم للوجبة التالية' },
  { key: 'وسائد جاهزة', type: 'نهاية', label: 'الوسائد الجاهزة', hint: '', nightOnly: true },
  { key: 'خزن وسائد ليلي', type: 'نهاية', label: 'خزن الوسائد الليلي', hint: '' },
  { key: 'شراشف المخزن الداخلي', type: 'نهاية', label: 'الشراشف في المخزن الداخلي', hint: 'اختياري' },
  { key: 'بطانيات متسخة مستلمة', type: 'بداية', label: 'البطانيات المتسخة المستلمة من الوجبة السابقة', hint: 'اختياري: تُعدّ عند غسلها' }
];

function invFields_(night) {
  return INV_FIELDS.filter(f => night || !f.nightOnly);
}

function getInventory(token) {
  const shift = shiftOf_(session_(token));
  const rows = sheet_('الجرد').getDataRange().getValues().slice(1).filter(r => r[1] === shift.id);
  return invFields_(shift.type === 'ليلية').map(f => {
    const r = rows.find(x => x[3] === f.key);
    return { key: f.key, label: f.label, hint: f.hint, value: r ? String(r[4]) : '' };
  });
}

function saveInventory(token, values) {
  const user = session_(token);
  if (user.role !== 'مدوّن') throw new Error('الإدخال للمدوّنين فقط');
  const shift = shiftOf_(user);
  const now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
  const rows = [];
  invFields_(shift.type === 'ليلية').forEach((f, i) => {
    const v = toLatin_(values[f.key] == null ? '' : values[f.key]).trim();
    if (v === '') return;
    const q = Number(v);
    if (!Number.isInteger(q) || q < 0 || q > 5000) throw new Error('عدد غير صحيح في: ' + f.label);
    rows.push(['J' + Date.now().toString(36).toUpperCase() + i, shift.id, f.type, f.key, q, '', '', user.name, now]);
  });
  if (!rows.length) throw new Error('لم تكتب أي عدد');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    ensureShift_(shift, userRow_(user.id));
    const sh = sheet_('الجرد');
    const data = sh.getDataRange().getValues();
    const keys = rows.map(r => r[3]);
    for (let r = data.length - 1; r >= 1; r--) {   // آخر رقم هو المعتمد
      if (data[r][1] === shift.id && keys.indexOf(data[r][3]) !== -1) sh.deleteRow(r + 1);
    }
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  } finally {
    lock.releaseLock();
  }
  return getInventory(token);
}

// ===== التقرير =====
function dayName_(date) {
  return DAYS[Number(Utilities.formatDate(new Date(date + 'T12:00:00+03:00'), TZ, 'u'))];
}

function buildReport_(shiftId) {
  const shiftRow = sheet_('الوجبات').getDataRange().getValues().slice(1).find(r => r[0] === shiftId);
  if (!shiftRow) throw new Error('لا توجد بيانات لهذه الوجبة بعد');
  let date = shiftRow[1];
  if (date instanceof Date) date = Utilities.formatDate(date, TZ, 'yyyy-MM-dd');
  date = String(date);
  const night = shiftRow[2] === 'ليلية';
  const recorder = shiftRow[4];

  const types = {};
  sheet_('الجهات').getDataRange().getValues().slice(1).forEach(r => { types[r[1]] = r[3]; });

  const mv = sheet_('الحركات').getDataRange().getValues().slice(1).filter(r => r[1] === shiftId);
  const sum = fn => mv.filter(fn).reduce((a, r) => a + (Number(r[6]) || 0), 0);
  const icuIn = sum(r => types[r[4]] === 'عناية' && r[5] === 'شرشف' && r[3] === 'وارد');
  const icuOut = sum(r => types[r[4]] === 'عناية' && r[5] === 'شرشف' && r[3] === 'صادر');
  const sheetsIn = sum(r => types[r[4]] !== 'عناية' && r[5] === 'شرشف' && r[3] === 'وارد');
  const nursingOut = sum(r => types[r[4]] === 'تمريض' && r[5] === 'شرشف' && r[3] === 'صادر');

  const orRows = mv.filter(r => types[r[4]] === 'عمليات' && r[5] === 'بدلة عمليات');
  const orReady = [...new Set(orRows
    .filter(r => r[3] === 'صادر' && Number(r[6]) > 0)
    .map(r => String(r[4]).replace('العمليات - ', '')))];

  const washes = sheet_('الغسلات').getDataRange().getValues().slice(1).filter(r => r[1] === shiftId).length;

  const inv = {};
  sheet_('الجرد').getDataRange().getValues().slice(1)
    .filter(r => r[1] === shiftId).forEach(r => { inv[r[3]] = r[4]; });
  const v = k => (inv[k] === undefined || inv[k] === '') ? 'لم يُدخل' : inv[k];

  const p = date.split('-');
  const lines = [
    'التقرير اليومي لوحدة اللوندري',
    (night ? 'الشفت الليلي' : 'الشفت الصباحي') + ' — المدوّن: ' + recorder,
    '',
    'يوم ' + dayName_(date) + ' / ' + p[0] + '/' + p[1] + '/' + p[2],
    ''
  ];

  if (night) {
    lines.push(
      '1- عدد الغسلات ' + washes,
      '2- عدد البطانيات الجاهزة ' + v('بطانيات جاهزة'),
      '3- عدد الوسائد الجاهزة ' + v('وسائد جاهزة'),
      '4- شراشف العنايات',
      'وارد ' + icuIn,
      'صادر ' + icuOut,
      '5- الشراشف الواردة من الرقود',
      'الوارد ' + sheetsIn
    );
  } else {
    lines.push(
      '1- عدد البطانيات الجاهزة ' + v('بطانيات جاهزة'),
      '2- الشراشف الصادرة إلى التمريض ' + nursingOut,
      '3- الشراشف الواردة من الطوابق (عدا العنايات) ' + sheetsIn,
      '4- شراشف العنايات',
      'وارد ' + icuIn,
      'صادر ' + icuOut
    );
  }

  // بدلات العمليات: دائماً في الليلي، وفي الصباحي فقط إذا وردت
  if (night || orRows.length) {
    lines.push('', 'بدلات العمليات');
    if (orReady.length) {
      lines.push('تم تجهيز بدلات العمليات');
      orReady.forEach(f => lines.push(' ' + f));
    } else {
      lines.push('لم تُجهَّز بدلات في هذه الوجبة');
    }
  }

  lines.push('', 'خزن وسائد ليلي ' + v('خزن وسائد ليلي'));
  return lines.join('\n');
}

function previewReport(token) {
  return buildReport_(shiftOf_(session_(token)).id);
}

// ===== لوحة المتابعة =====
function getDashboard(token, from, to) {
  const user = session_(token);
  if (user.role !== 'مسؤول' && user.role !== 'مدير النظام') throw new Error('لوحة المتابعة للمسؤولين فقط');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(from) || !/^\d{4}-\d{2}-\d{2}$/.test(to) || from > to) throw new Error('الفترة غير صحيحة');
  const day = d => Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
  const dateOf = id => String(id).slice(0, 10);
  const inRange = id => dateOf(id) >= from && dateOf(id) <= to;
  const rows = name => sheet_(name).getDataRange().getValues().slice(1);
  const qty = r => Number(r[6]) || 0;
  const perWash = Number(setting_('بطانيات لكل غسلة', 6));
  const sheetsPerWash = Number(setting_('شراشف لكل غسلة (تقدير)', 60));

  const types = {}, order = {};
  rows('الجهات').filter(r => r[6] === 'نعم').forEach(r => { types[r[1]] = r[3]; order[r[1]] = Number(r[5]); });
  const mvAll = rows('الحركات'), wsAll = rows('الغسلات'), invAll = rows('الجرد'), shAll = rows('الوجبات');
  const mv = mvAll.filter(r => inRange(r[1])), ws = wsAll.filter(r => inRange(r[1]));
  const sum = (list, f) => list.filter(f).reduce((a, r) => a + qty(r), 0);
  const is = (item, type, dept) => r => r[5] === item && r[3] === type && (!dept || r[4] === dept);

  // المتراكم المتسخ (تقدير، منذ بداية السجل): البطانيات الواردة − غسلات البطانيات × سعة الغسلة
  const dirty = Math.max(0, sum(mvAll, is('بطانية', 'وارد')) - wsAll.length * perWash);

  // السلسلة اليومية
  const days = [], byDay = {};
  for (let d = new Date(from + 'T12:00:00+03:00'); day(d) <= to && days.length < 370; d = new Date(d.getTime() + 864e5)) {
    const o = { date: day(d), washD: 0, washN: 0, sheetsIn: 0, blanketsIn: 0 };
    days.push(o);
    byDay[o.date] = o;
  }
  ws.forEach(r => { const o = byDay[dateOf(r[1])]; if (o) o[String(r[1]).slice(-1) === 'N' ? 'washN' : 'washD']++; });
  mv.filter(r => r[3] === 'وارد').forEach(r => {
    const o = byDay[dateOf(r[1])];
    if (!o) return;
    if (r[5] === 'شرشف') o.sheetsIn += qty(r);
    if (r[5] === 'بطانية') o.blanketsIn += qty(r);
  });

  // الغسلات لكل غسالة ولكل مجموعة
  const washers = { 1: 0, 2: 0, 3: 0 };
  ws.forEach(r => { washers[Number(r[3])] = (washers[Number(r[3])] || 0) + 1; });
  const groupOf = {};
  shAll.forEach(r => { groupOf[r[0]] = r[3] || 'غير محدد'; });
  const groups = {};
  ws.forEach(r => { const g = groupOf[r[1]] || 'غير محدد'; groups[g] = (groups[g] || 0) + 1; });

  // الاستهلاك حسب الجهة والصنف
  const depts = {};
  mv.forEach(r => {
    const d = depts[r[4]] = depts[r[4]] || { name: r[4], order: order[r[4]] || 99, items: {} };
    const it = d.items[r[5]] = d.items[r[5]] || { out: 0, in: 0 };
    it[r[3] === 'صادر' ? 'out' : 'in'] += qty(r);
  });

  // رصيد العنايات وبدلات العمليات ضمن الفترة
  const icu = Object.keys(types).filter(n => types[n] === 'عناية').map(n => {
    const o = sum(mv, is('شرشف', 'صادر', n)), i = sum(mv, is('شرشف', 'وارد', n));
    return { name: n, out: o, in: i, balance: o - i };
  });
  const or = Object.keys(types).filter(n => types[n] === 'عمليات').map(n => ({
    name: n.replace('العمليات - ', ''),
    out: sum(mv, is('بدلة عمليات', 'صادر', n)), in: sum(mv, is('بدلة عمليات', 'وارد', n)),
    shifts: new Set(mv.filter(r => is('بدلة عمليات', 'صادر', n)(r) && qty(r) > 0).map(r => r[1])).size
  }));

  // المخزن: آخر قيمة لكل بند جرد نهاية الوجبة
  const invBy = {}; // shiftId -> key -> value
  invAll.forEach(r => { (invBy[r[1]] = invBy[r[1]] || {})[r[3]] = Number(r[4]); });
  const shiftIds = Object.keys(invBy).sort();
  const store = INV_FIELDS.filter(f => f.type === 'نهاية').map(f => {
    const series = shiftIds.filter(id => inRange(id) && invBy[id][f.key] !== undefined)
      .map(id => ({ shift: id, value: invBy[id][f.key] }));
    const lastId = shiftIds.filter(id => invBy[id][f.key] !== undefined).pop();
    return { key: f.key, label: f.label, last: lastId ? invBy[lastId][f.key] : null, lastShift: lastId || '', series: series };
  });

  // فروق الجرد: البطانيات الجاهزة المتوقعة = جاهز الوجبة السابقة + غسلات الوجبة × السعة − صادر البطانيات
  const prevId = id => id.slice(-1) === 'N' ? dateOf(id) + '-D'
    : day(new Date(new Date(dateOf(id) + 'T12:00:00+03:00').getTime() - 864e5)) + '-N';
  const diffs = shiftIds.filter(inRange).map(id => {
    const prev = invBy[prevId(id)], cur = invBy[id];
    if (!prev || prev['بطانيات جاهزة'] === undefined || cur['بطانيات جاهزة'] === undefined) return null;
    const washes = wsAll.filter(r => r[1] === id).length;
    const out = sum(mvAll, r => r[1] === id && is('بطانية', 'صادر')(r));
    const expected = prev['بطانيات جاهزة'] + washes * perWash - out;
    return { shift: id, recorder: (shAll.find(r => r[0] === id) || [])[4] || '', expected: expected,
      actual: cur['بطانيات جاهزة'], diff: cur['بطانيات جاهزة'] - expected };
  }).filter(Boolean);

  // انتظام التقارير لكل وجبة في الفترة
  const cur = currentShift_().id;
  const shRow = {};
  shAll.forEach(r => { shRow[r[0]] = r; });
  const first = Object.keys(shRow).sort()[0] || cur; // أول وجبة في السجل
  const reports = [];
  days.forEach(o => ['D', 'N'].forEach(t => {
    const id = o.date + '-' + t, r = shRow[id];
    let status;
    if (id > cur || id < first) status = 'future';
    else if (id === cur) status = 'open';
    else if (r && r[5] === 'فارغة') status = 'empty';
    else if (r && r[8] !== '') status = 'sent';
    else status = 'missing';
    reports.push({ shift: id, status: status, recorder: r ? r[4] : '', group: r ? r[3] : '', sentAt: r ? String(r[8]) : '' });
  }));
  const due = reports.filter(r => r.status === 'sent' || r.status === 'empty' || r.status === 'missing');

  const sheetsIn = sum(mv, r => r[5] === 'شرشف' && r[3] === 'وارد');
  return {
    from: from, to: to, updated: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'),
    perWash: perWash, dirty: dirty,
    kpi: {
      washes: ws.length,
      blanketsIn: sum(mv, is('بطانية', 'وارد')),
      blanketsOut: sum(mv, is('بطانية', 'صادر')),
      sheetsIn: sheetsIn,
      sheetsOut: sum(mv, r => r[5] === 'شرشف' && r[3] === 'صادر'),
      sheetWashesEst: Math.round(sheetsIn / sheetsPerWash),
      suitsOut: sum(mv, r => r[5] === 'بدلة عمليات' && r[3] === 'صادر'),
      reportsSent: due.filter(r => r.status === 'sent').length,
      reportsDue: due.length
    },
    days: days, washers: washers, groups: groups,
    depts: Object.values(depts).sort((a, b) => a.order - b.order),
    icu: icu, or: or, store: store, diffs: diffs, reports: reports
  };
}

// ===== Telegram =====
// الرمز السري للبوت في خصائص النص البرمجي: TELEGRAM_TOKEN
// معرّف المجموعة يُحفظ تلقائياً عند الربط: TELEGRAM_CHAT_ID
function tgCall_(method, payload) {
  const token = PropertiesService.getScriptProperties().getProperty('TELEGRAM_TOKEN');
  if (!token) throw new Error('لم يُضبط TELEGRAM_TOKEN في خصائص النص البرمجي');
  const res = UrlFetchApp.fetch('https://api.telegram.org/bot' + token.trim() + '/' + method, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload || {}),
    muteHttpExceptions: true
  });
  const data = JSON.parse(res.getContentText());
  if (!data.ok) throw new Error('Telegram: ' + data.description);
  return data.result;
}

function tgSend_(text) {
  const chatId = PropertiesService.getScriptProperties().getProperty('TELEGRAM_CHAT_ID');
  if (!chatId) throw new Error('لم تُربط مجموعة Telegram بعد');
  return tgCall_('sendMessage', { chat_id: chatId, text: text });
}

// من القائمة: يبحث عن آخر مجموعة أُضيف إليها البوت ويحفظها، ثم يرسل رسالة تجريبية
function linkTelegram() {
  const ui = SpreadsheetApp.getUi();
  try {
    const updates = tgCall_('getUpdates', {});
    let chatId = null, title = '';
    updates.forEach(u => {
      const m = u.message || u.my_chat_member;
      if (!m || !m.chat || (m.chat.type !== 'group' && m.chat.type !== 'supergroup')) return;
      chatId = (u.message && u.message.migrate_to_chat_id) || m.chat.id;
      title = m.chat.title;
    });
    if (!chatId) throw new Error('لم أجد المجموعة. أرسل /start داخل المجموعة ثم أعد المحاولة.');
    PropertiesService.getScriptProperties().setProperty('TELEGRAM_CHAT_ID', String(chatId));
    tgSend_('✅ تم ربط سجل اللوندري بهذه المجموعة. ستصل التقارير هنا تلقائياً بعد كل وجبة.');
    ui.alert('تم الربط بالمجموعة: ' + title + '\nتحقق من وصول رسالة تجريبية فيها.');
  } catch (e) {
    ui.alert(e.message);
  }
}

// ===== الإرسال التلقائي وقفل الوجبة =====
function setting_(key, def) {
  const row = sheet_('الإعدادات').getDataRange().getValues().slice(1).find(r => r[0] === key);
  return row && row[1] !== '' ? row[1] : def;
}

// من القائمة: يرسل تقرير الوجبة الحالية إلى المجموعة دون قفلها
function sendTestReport() {
  const ui = SpreadsheetApp.getUi();
  try {
    tgSend_('🧪 تقرير تجريبي (الوجبة لم تنتهِ بعد)\n\n' + buildReport_(currentShift_().id));
    ui.alert('أُرسل التقرير التجريبي إلى المجموعة');
  } catch (e) {
    ui.alert(e.message);
  }
}

// من القائمة: ينشئ مشغّلاً كل 5 دقائق (ويحذف أي مشغّل سابق حتى لا يتكرر)
function enableAutoSend() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'autoSend')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('autoSend').timeBased().everyMinutes(5).create();
  SpreadsheetApp.getUi().alert('تم تفعيل الإرسال التلقائي: يُرسل تقرير كل وجبة بعد نهايتها بـ '
    + setting_('دقائق الإرسال بعد نهاية الوجبة', 5) + ' دقائق.');
}

function shiftTitle_(s) {
  return 'الوجبة ال' + s.type + ' — يوم ' + s.day + ' ' + s.display;
}

// النص المخزَّن 'yyyy-MM-dd HH:mm:ss' بتوقيت بغداد
function toDate_(v) {
  return v instanceof Date ? v : new Date(String(v).replace(' ', 'T') + '+03:00');
}

// ما لم يُسجَّل بعد في الوجبة: الجهات بلا حركة، الغسلات، بنود الجرد
function missing_(shift) {
  const night = shift.type === 'ليلية';
  const done = {};
  sheet_('الحركات').getDataRange().getValues().slice(1)
    .filter(r => r[1] === shift.id).forEach(r => { done[r[4]] = true; });
  const depts = sheet_('الجهات').getDataRange().getValues().slice(1)
    .filter(r => r[6] === 'نعم' && itemsFor_(r[3], night).length && !done[r[1]])
    .sort((a, b) => Number(a[5]) - Number(b[5]))
    .map(r => r[1]);
  const washes = sheet_('الغسلات').getDataRange().getValues().slice(1).some(r => r[1] === shift.id);
  const inv = {};
  sheet_('الجرد').getDataRange().getValues().slice(1)
    .filter(r => r[1] === shift.id).forEach(r => { inv[r[3]] = true; });
  const invMissing = invFields_(night)
    .filter(f => f.type === 'نهاية' && f.hint !== 'اختياري' && !inv[f.key])
    .map(f => f.label);

  const lines = [];
  if (!washes) lines.push('• الغسلات: لم تُسجَّل أي غسلة');
  if (depts.length) lines.push('• جهات بلا حركة: ' + depts.join('، '));
  if (invMissing.length) lines.push('• الجرد والتسليم: ' + invMissing.join('، '));
  return lines;
}

// آخر وقت إدخال في الوجبة (حركة أو غسلة أو جرد)، أو بداية الوجبة إن لم يوجد
function lastEntry_(shift) {
  let last = new Date(shift.date + 'T' + (shift.type === 'ليلية' ? '19' : '07') + ':00:00+03:00');
  [['الحركات', 2], ['الغسلات', 2], ['الجرد', 8]].forEach(([name, col]) => {
    sheet_(name).getDataRange().getValues().slice(1).filter(r => r[1] === shift.id).forEach(r => {
      const t = toDate_(r[col]);
      if (!isNaN(t) && t > last) last = t;
    });
  });
  return last;
}

// تذكير قبل نهاية الوجبة بما لم يُسجَّل، وتذكير إذا مرّت ساعات دون أي إدخال
function remind_(now, since) {
  const cur = currentShift_(now);
  const cache = CacheService.getScriptCache();
  const before = Number(setting_('دقائق التذكير قبل نهاية الوجبة', 45));
  if (before > 0 && since >= 720 - before && !cache.get('remEnd_' + cur.id)) {
    cache.put('remEnd_' + cur.id, '1', 21600);
    const miss = missing_(cur);
    tgSend_('⏰ باقي ' + (720 - since) + ' دقيقة على نهاية ' + shiftTitle_(cur) + '\n\n'
      + (miss.length ? 'لم يُسجَّل بعد:\n' + miss.join('\n') : '✅ كل البنود الأساسية مسجّلة. تأكد من الجرد قبل التسليم.'));
    return;
  }
  const idleH = Number(setting_('ساعات بلا إدخال قبل التذكير', 3));
  if (idleH <= 0 || since >= 720 - before) return;
  const last = lastEntry_(cur);
  const key = 'remIdle_' + cur.id + '_' + last.getTime();
  if (now - last >= idleH * 3600 * 1000 && !cache.get(key)) {
    cache.put(key, '1', 21600);
    tgSend_('🔔 لم يُسجَّل أي إدخال في ' + shiftTitle_(cur) + ' منذ '
      + Math.floor((now - last) / 3600000) + ' ساعات (آخر إدخال '
      + Utilities.formatDate(last, TZ, 'HH:mm') + '). هل هناك حركات لم تُدخل؟');
  }
}

// يعمل كل 5 دقائق: التذكيرات، ثم إذا انتهت الوجبة السابقة ومضت دقائق الإرسال يقفلها ويرسل تقريرها مرة واحدة
function autoSend() {
  const now = new Date();
  const sinceStart = sinceShiftStart_(now);
  try { remind_(now, sinceStart); } catch (e) { console.error('التذكير: ' + e.message); }
  if (sinceStart < Number(setting_('دقائق الإرسال بعد نهاية الوجبة', 5))) return;

  const prev = currentShift_(new Date(now.getTime() - 12 * 3600 * 1000)); // الوجبة التي انتهت
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;
  try {
    const sh = sheet_('الوجبات');
    const data = sh.getDataRange().getValues();
    const i = data.findIndex((r, k) => k > 0 && r[0] === prev.id);
    if (i > 0 && data[i][8] !== '') return; // أُرسل سابقاً
    const stamp = Utilities.formatDate(now, TZ, 'yyyy-MM-dd HH:mm');

    const has = name => sheet_(name).getDataRange().getValues().slice(1).some(r => r[1] === prev.id);
    const empty = !has('الحركات') && !has('الغسلات') && !has('الجرد');
    if (empty) {
      tgSend_('⚠️ تنبيه: انتهت ' + shiftTitle_(prev) + ' دون أي إدخال في السجل'
        + (i > 0 && data[i][4] ? ' (المدوّن: ' + data[i][4] + ')' : '') + '.');
    } else {
      tgSend_(buildReport_(prev.id));
    }

    // الحالة، وقت القفل، وقت الإرسال
    if (i > 0) {
      sh.getRange(i + 1, 6).setValue(empty ? 'فارغة' : 'مقفلة');
      sh.getRange(i + 1, 8, 1, 2).setValues([[stamp, stamp]]);
    } else {
      sh.appendRow([prev.id, prev.date, prev.type, '', '', 'فارغة', '', stamp, stamp]);
    }
  } finally {
    lock.releaseLock();
  }
}