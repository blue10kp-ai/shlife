(function () {
  'use strict';
  // 시작 순서: 구글 로그인 → Drive에서 데이터 읽기 → 화면 코드(access.js, page-*.js, app.js) 불러오기
  // 데이터 사본(data\*.js)은 쓰지 않는다. 로그인 전에는 어떤 고객·수수료 정보도 화면에 없다.
  var main = document.getElementById('main'), nav = document.getElementById('nav');
  var esc = function (x) { return String(x == null ? '' : x).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  var SCRIPTS = ['access.js', 'page-dashboard.js', 'page-contract.js', 'page-team.js', 'page-customer.js', 'page-fee.js', 'page-settings.js', 'app.js'];

  function box(html) { main.innerHTML = '<div class="bootbox">' + html + '</div>'; }
  function loginView(msg) {
    box('<h1>신한FC</h1><p class="sub">구글 계정으로 로그인하면 Google Drive의 고객·수수료 데이터를 불러옵니다.</p>' +
      (msg ? '<p class="warn">' + esc(msg) + '</p>' : '') + '<button type="button" class="btn pri" id="bootLogin">구글 로그인</button>');
    document.getElementById('bootLogin').onclick = function () {
      var b = this; b.disabled = true; b.textContent = '로그인 창을 확인하세요…';
      DRIVE.login().then(start).catch(function (e) { loginView(e.message || String(e)); });
    };
  }
  function loadScript(src) {
    return new Promise(function (ok, no) {
      var el = document.createElement('script'); el.src = src + '?v=' + Date.now().toString(36); el.async = false;
      el.onload = ok; el.onerror = function () { no(new Error(src + ' 를 불러오지 못했습니다.')); };
      document.body.appendChild(el);
    });
  }
  function shells() {   // 아직 Drive에서 읽지 않는 자료의 빈 틀(화면이 깨지지 않게)
    if (!window.PERF) window.PERF = { fcs: [], currentFcs: [], current: [], years: {}, from: '', upto: '', made: '' };
    if (!window.CONTRACTS) window.CONTRACTS = { made: '', rows: [] };
  }
  async function start() {
    try {
      box('<h1>신한FC</h1><p class="sub" id="bootMsg">Drive에서 읽는 중…</p>');
      var errs = await DRIVE.loadAll(function (t) { var m = document.getElementById('bootMsg'); if (m) m.textContent = t; });
      shells();
      for (var i = 0; i < SCRIPTS.length; i++) await loadScript(SCRIPTS[i]);
      if (errs.length) alert('일부 자료를 읽지 못했습니다:\n' + errs.join('\n'));
    } catch (e) {
      var m = e && e.message || String(e);
      if (m === 'GOOGLE_AUTH') return loginView('로그인이 만료되었습니다. 다시 로그인해 주세요.');
      box('<h1>신한FC</h1><p class="warn">' + esc(m) + '</p><button type="button" class="btn" id="bootOut">다른 계정으로 로그인</button> <button type="button" class="btn" id="bootRetry">다시 시도</button>');
      document.getElementById('bootOut').onclick = function () { DRIVE.logout(); loginView(); };
      document.getElementById('bootRetry').onclick = start;
    }
  }
  if (nav) nav.innerHTML = '';
  var tries = 0;
  (function wait() {   // 구글 로그인 모듈(gsi)이 올 때까지 잠깐 기다림
    if (DRIVE.on()) return start();
    if (window.google && window.google.accounts) return loginView();
    if (++tries > 100) return loginView('구글 로그인 모듈을 불러오지 못했습니다. 인터넷 연결을 확인하세요.');
    setTimeout(wait, 100);
  })();
})();
