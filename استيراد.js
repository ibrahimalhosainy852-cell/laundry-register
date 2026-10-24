// ===== استيراد التقارير الورقية لشهر تشرين الأول 2026 + المدوّنين =====
// يُشغَّل مرة واحدة (يستدعيه autoSend إن لم يُشغَّل بعد، أو من القائمة). آمن للتكرار: لا يكرر ما أُدخل.
const IMPORT_TAG = 'من التقارير الورقية';
const IMPORT_KEY = 'IMPORT_OCT_2026';

// الاسم، المجموعة، الرمز (تسلسلي: 2222، 3333، ...)
const IMPORT_USERS = [
  ['علي صلاح', 'صباحية أ', '2222'],
  ['مصطفى طالب', 'صباحية ب', '3333'],
  ['عبد الله أحمد ناصر', 'ليلية أ', '4444'],
  ['إبراهيم جواد', 'ليلية ب', '5555'],
  ['أحمد صباح', 'صباحية ب', '6666'],
  ['قحطان علي كاظم', 'صباحية أ', '7777'],
  ['قاسم عيسى عطوان', 'صباحية أ', '8888']
];

// "موقف اللوندري" = الوجبة الصباحية:
// [التاريخ, شراشف مكوية, شراشف المخزن الداخلي, بطانيات المخزن الداخلي, وسائد المخزن الداخلي, وسائد جاهزة,
//  بطانيات جاهزة, شراشف الطوابق صادر, شراشف الطوابق وارد, العنايات صادر, العنايات وارد, الغسلات, بدلات (طوابق), خزن وسائد ليلي]
const IMPORT_DAY = [
  ['2026-10-01', 370, 440, 0, 0, 0, 50, 180, 92, 80, 78, 29, [2], 28],
  ['2026-10-02', 280, 400, 0, 0, 0, 55, 180, 98, 80, 42, 28, [3], 30],
  ['2026-10-03', 370, 420, 0, 0, 0, 40, 180, 89, 80, 38, 30, [2, 3], 10],
  ['2026-10-04', 300, 440, 0, 27, 0, 50, 180, 126, 0, 30, 24, [], 30],
  ['2026-10-05', 300, 460, 0, 0, 0, 40, 160, 79, 40, 40, 29, [], 22],
  ['2026-10-06', 360, 520, 0, 27, 0, 40, 180, 119, 50, 33, 29, [], 30],
  ['2026-10-07', 280, 440, 0, 10, 0, 45, 180, 95, 80, 38, 31, [], 27],
  ['2026-10-08', 320, 420, 0, 30, 0, 65, 190, 113, 80, 110, 25, [2], 30],
  ['2026-10-09', 290, 380, 0, 30, 0, 70, 180, 120, 80, 76, 24, [], 32],
  ['2026-10-10', 300, 380, 0, 15, 0, 70, 180, 91, 54, 40, 34, [], 35], // الصادر/الوارد كانا معكوسين في الأصل
  ['2026-10-11', 320, 360, 0, 46, 0, 175, 160, 95, 80, 63, 40, [2, 3], 30],
  ['2026-10-12', 380, 420, 0, 80, 0, 70, 180, 103, 60, 53, 25, [], 30],
  ['2026-10-13', 300, 380, 0, 15, 0, 40, 160, 105, 60, 40, 29, [3], 33],
  ['2026-10-14', 340, 380, 0, 30, 0, 70, 180, 95, 80, 96, 31, [2], 30],
  ['2026-10-15', 360, 440, 0, 20, 0, 60, 160, 80, 80, 65, 26, [3], 30],
  ['2026-10-16', 300, 480, 0, 41, 0, 80, 180, 107, 40, 76, 32, [3], 30],
  ['2026-10-17', 320, 420, 0, 0, 0, 40, 160, 85, 80, 52, 27, [3], 25],
  ['2026-10-18', 340, 400, 0, 35, 12, 105, 200, 117, 40, 45, 30, [3], 30],
  ['2026-10-19', 280, 360, 0, 30, 0, 60, 160, 89, 80, 50, 28, [3], 34],
  ['2026-10-20', 360, 380, 0, 33, 0, 90, 180, 88, 60, 41, 30, [3], 30],
  ['2026-10-21', 300, 360, 0, 30, 0, 75, 160, 110, 80, 58, 26, [3], 30],
  ['2026-10-22', 320, 360, 0, 33, 0, 80, 180, 135, 80, 84, 27, [], 30],
  ['2026-10-23', 300, 340, 0, 33, 0, 100, 160, 94, 80, 78, 23, [3], 33]
];

// "التقرير اليومي — الشفت الليلي" (بدلات طابق 2 و3 في كل ليلة):
// [التاريخ, الغسلات, بطانيات جاهزة, وسائد جاهزة, العنايات وارد, العنايات صادر, شراشف الرقود وارد, شراشف الرقود صادر, خزن وسائد ليلي]
const IMPORT_NIGHT = [
  ['2026-10-01', 14, 25, 0, 19, 40, 35, 0, 33],
  ['2026-10-02', 13, 56, 0, 55, 20, 45, 20, 31],
  ['2026-10-03', 13, 18, 0, 43, 40, 33, 0, 35],
  ['2026-10-04', 15, 60, 0, 18, 20, 57, 0, 45],
  ['2026-10-05', 14, 20, 0, 26, 40, 44, 0, 23],
  ['2026-10-06', 15, 50, 0, 20, 40, 42, 0, 36],
  ['2026-10-07', 15, 12, 0, 21, 40, 32, 0, 15],
  ['2026-10-08', 16, 70, 0, 22, 40, 60, 0, 42],
  ['2026-10-09', 8, 18, 0, 43, 40, 33, 0, 20],
  ['2026-10-11', 11, 50, 0, 22, 40, 38, 0, 36],
  ['2026-10-13', 18, 27, 0, 8, 40, 22, 0, 33],
  ['2026-10-14', 17, 70, 0, 10, 20, 50, 20, 25],
  ['2026-10-15', 16, 13, 0, 18, 40, 28, 0, 18],
  ['2026-10-16', 13, 60, 0, 31, 40, 50, 0, 39],
  ['2026-10-17', 19, 35, 0, 32, 40, 54, 0, 6],
  ['2026-10-18', 15, 90, 15, 30, 40, 65, 0, 73],
  ['2026-10-19', 17, 30, 0, 15, 20, 71, 0, 40],
  ['2026-10-20', 14, 80, 0, 35, 40, 78, 0, 41],
  ['2026-10-21', 13, 15, 0, 39, 40, 29, 0, 18],
  ['2026-10-23', 10, 10, 0, 13, 40, 25, 0, 20]
];

function importOctober() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const result = importOctober_();
    PropertiesService.getScriptProperties().setProperty(IMPORT_KEY, JSON.stringify(result));
    return result;
  } finally {
    lock.releaseLock();
  }
}

function importOctober_() {
  const res = { users: 0, shifts: 0, movements: 0, washes: 0, inventory: 0, skipped: [] };

  // المدوّنون
  const us = sheet_('المستخدمون');
  const names = us.getDataRange().getValues().slice(1).map(r => r[1]);
  IMPORT_USERS.forEach(([name, group, pin], i) => {
    if (names.includes(name)) return;
    const id = 'U' + (Date.now() + i).toString(36).toUpperCase();
    us.appendRow([id, name, 'مدوّن', group, hashPin_(id, pin), 'نعم']);
    res.users++;
  });

  // أسماء الجهات كما في ورقة الجهات
  const dept = {};
  sheet_('الجهات').getDataRange().getValues().slice(1).forEach(r => { dept[r[0]] = r[1]; });
  const FLOORS = dept.R01, ICU = dept.R03, NURSING = dept.R08, OR = { 2: dept.R09, 3: dept.R10 };

  // وجبات فيها إدخال حقيقي من النظام لا نلمسها
  const has = {};
  [['الحركات', 8], ['الغسلات', 4], ['الجرد', 7]].forEach(([n, col]) => {
    sheet_(n).getDataRange().getValues().slice(1).forEach(r => { if (r[col] !== IMPORT_TAG) has[r[1]] = true; });
  });
  const shSheet = sheet_('الوجبات');
  const shData = shSheet.getDataRange().getValues();
  const shRow = {};
  shData.forEach((r, i) => { if (i) shRow[r[0]] = i + 1; });

  const mv = [], ws = [], inv = [], newShifts = [];
  let seq = 0;
  const uid = p => p + Date.now().toString(36).toUpperCase() + (seq++).toString(36).toUpperCase();
  const canImport = shiftId => {
    if (has[shiftId]) { res.skipped.push(shiftId); return false; }
    const existing = shRow[shiftId];
    if (existing && shData[existing - 1][4] === IMPORT_TAG) return false; // مستورد سابقاً
    return true;
  };

  function addShift(date, night, washes, moves, invs) {
    const id = date + (night ? '-N' : '-D');
    if (!canImport(id)) return;
    const t = date + (night ? ' 20:00:00' : ' 08:00:00');
    moves.forEach(([type, d, item, qty, note]) => {
      if (qty > 0) mv.push([uid('M'), id, t, type, d, item, qty, '', IMPORT_TAG, note || IMPORT_TAG]);
    });
    for (let i = 0; i < washes; i++) ws.push([uid('W'), id, t, 0, IMPORT_TAG]);
    invs.forEach(([key, qty]) => inv.push([uid('J'), id, 'نهاية', key, qty, '', '', IMPORT_TAG, t]));
    const row = [id, date, night ? 'ليلية' : 'صباحية', '', IMPORT_TAG, 'مقفلة', t, t, t];
    if (shRow[id]) shSheet.getRange(shRow[id], 1, 1, row.length).setValues([row]); // وجبة فارغة أنشأها الإرسال التلقائي
    else newShifts.push(row);
    res.shifts++;
  }

  const ICU_NOTE = IMPORT_TAG + ' — مجموع العنايات الثلاث';
  const SUIT_NOTE = IMPORT_TAG + ' — تجهيز (العدد غير مسجّل)';
  IMPORT_DAY.forEach(r => {
    const [date, ironed, inSheets, inBlankets, inPillows, readyP, readyB, flOut, flIn, icuOut, icuIn, washes, ors, pillowStore] = r;
    addShift(date, false, washes, [
      ['صادر', NURSING, 'شرشف', flOut], ['وارد', FLOORS, 'شرشف', flIn],
      ['صادر', ICU, 'شرشف', icuOut, ICU_NOTE], ['وارد', ICU, 'شرشف', icuIn, ICU_NOTE]
    ].concat(ors.map(f => ['صادر', OR[f], 'بدلة عمليات', 1, SUIT_NOTE])), [
      ['شراشف مكوية', ironed], ['شراشف المخزن الداخلي', inSheets], ['بطانيات المخزن الداخلي', inBlankets],
      ['وسائد المخزن الداخلي', inPillows], ['وسائد جاهزة', readyP], ['بطانيات جاهزة', readyB], ['خزن وسائد ليلي', pillowStore]
    ]);
  });
  IMPORT_NIGHT.forEach(r => {
    const [date, washes, readyB, readyP, icuIn, icuOut, flIn, flOut, pillowStore] = r;
    addShift(date, true, washes, [
      ['وارد', FLOORS, 'شرشف', flIn], ['صادر', FLOORS, 'شرشف', flOut],
      ['وارد', ICU, 'شرشف', icuIn, ICU_NOTE], ['صادر', ICU, 'شرشف', icuOut, ICU_NOTE],
      ['صادر', OR[2], 'بدلة عمليات', 1, SUIT_NOTE], ['صادر', OR[3], 'بدلة عمليات', 1, SUIT_NOTE]
    ], [['بطانيات جاهزة', readyB], ['وسائد جاهزة', readyP], ['خزن وسائد ليلي', pillowStore]]);
  });

  const put = (name, rows) => {
    if (!rows.length) return;
    const sh = sheet_(name);
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  };
  put('الحركات', mv); put('الغسلات', ws); put('الجرد', inv); put('الوجبات', newShifts);
  res.movements = mv.length; res.washes = ws.length; res.inventory = inv.length;
  return res;
}

// مرة واحدة (بطلب المستخدم): حذف إدخالات التجارب في صباحية 23 أكتوبر ثم استيراد الموقف الورقي مكانها
const REPLACE_SHIFT = '2026-10-23-D';
function replaceTestShift() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  const deleted = {};
  try {
    [['الحركات', 8], ['الغسلات', 4], ['الجرد', 7]].forEach(([name, col]) => {
      const sh = sheet_(name);
      const data = sh.getDataRange().getValues();
      deleted[name] = 0;
      for (let r = data.length - 1; r >= 1; r--) {
        if (data[r][1] === REPLACE_SHIFT && data[r][col] !== IMPORT_TAG) { sh.deleteRow(r + 1); deleted[name]++; }
      }
    });
  } finally {
    lock.releaseLock();
  }
  const result = importOctober(); // يستورد الآن الوجبة لأنها لم تعد تحتوي إدخالاً من النظام
  result.deletedTest = deleted;
  PropertiesService.getScriptProperties().setProperty(IMPORT_KEY, JSON.stringify(result));
  PropertiesService.getScriptProperties().setProperty('REPLACE_23D_DONE', '1');
  return result;
}

function importOctoberMenu() {
  const r = importOctober();
  SpreadsheetApp.getUi().alert('الاستيراد: ' + r.users + ' مدوّن، ' + r.shifts + ' وجبة، ' + r.movements + ' حركة، '
    + r.washes + ' غسلة، ' + r.inventory + ' بند جرد.' + (r.skipped.length ? '\nتُركت (فيها إدخال من النظام): ' + r.skipped.join('، ') : ''));
}

// حالة الاستيراد (أعداد فقط) — للتحقق عن بُعد
function importStatus() {
  return JSON.parse(PropertiesService.getScriptProperties().getProperty(IMPORT_KEY) || 'null');
}
