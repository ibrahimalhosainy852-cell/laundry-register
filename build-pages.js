// يبني صفحات تطبيق الهاتف في docs/ من Index.html و Dashboard.html نفسيهما.
// خارج Apps Script لا يوجد google.script.run، فنضيف بديلاً يتصل بالخادم عبر doPost
// دون إرسال حسابات Google (credentials: 'omit')، فلا تتأثر الصفحة بتعدد الحسابات في المتصفح.
// التشغيل: node build-pages.js
const fs = require('fs');

// رابط عملية النشر الرسمية (لا يتغيّر عند نشر إصدار جديد)
const API = process.env.API_URL ||
  'https://script.google.com/macros/s/AKfycbyBHtn7_j7wnNjLbM7lLiKFARtroKlrAPymMngOi_X99iVTUM1sBWxqGdqT5goQH6vTTg/exec';

const shim = `<script>
  // بديل google.script.run: نفس طريقة الاستدعاء، والاتصال مباشر بالخادم
  (function () {
    var API = ${JSON.stringify(API)};
    function runner() {
      var ok = function () {}, fail = function () {};
      var p = new Proxy({}, { get: function (_, fn) {
        if (fn === 'withSuccessHandler') return function (f) { ok = f; return p; };
        if (fn === 'withFailureHandler') return function (f) { fail = f; return p; };
        return function () {
          var args = Array.prototype.slice.call(arguments);
          fetch(API, { method: 'POST', credentials: 'omit', redirect: 'follow',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify({ fn: fn, args: args }) })
            .then(function (r) { return r.json(); })
            .then(function (d) { setTimeout(function () { d.ok ? ok(d.result) : fail(new Error(d.error)); }); },
              function () { fail(new Error('تعذّر الاتصال بالخادم. تحقق من الإنترنت وحاول مجدداً.')); });
        };
      } });
      return p;
    }
    window.google = { script: { get run() { return runner(); } } };
  })();
</script>`;

const head = title => `<title>${title}</title>
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<meta name="theme-color" content="#1f3a5f">
<link rel="manifest" href="manifest.webmanifest">
<link rel="icon" href="icons/favicon-32.png" sizes="32x32">
<link rel="apple-touch-icon" href="icons/apple-touch-icon.png">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-title" content="سجل اللوندري">`;

const sw = `<script>if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(function () {});</script>`;

[['Index.html', 'docs/index.html', 'سجل اللوندري'],
 ['Dashboard.html', 'docs/dashboard.html', 'لوحة متابعة اللوندري']].forEach(([src, dest, title]) => {
  let h = fs.readFileSync(src, 'utf8');
  if (h.indexOf('<script>') === -1) throw new Error(src + ': لا يوجد <script>');
  h = h.replace('<meta charset="utf-8">', '<meta charset="utf-8">\n' + head(title))
       .replace('<script>', shim + '\n<script>')
       .replace('</body>', sw + '\n</body>');
  fs.writeFileSync(dest, '<!-- مولَّد من ' + src + ' بواسطة build-pages.js — لا تعدّله مباشرة -->\n' + h);
  console.log('built', dest);
});
