(function () {
  'use strict';
  // 시작 순서: 구글 로그인 → Drive에서 데이터 읽기 → 화면 코드(access.js, page-*.js, app.js) 불러오기
  // 데이터 사본(data\*.js)은 쓰지 않는다. 로그인 전에는 어떤 고객·수수료 정보도 화면에 없다.
  var main = document.getElementById('main'), nav = document.getElementById('nav');
  var esc = function (x) { return String(x == null ? '' : x).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  var SCRIPTS = ['access.js', 'page-dashboard.js', 'page-contract.js', 'page-team.js', 'page-customer.js', 'page-fee.js', 'admin-perm.js', 'page-settings.js', 'app.js'];

  function box(html) { main.innerHTML = '<div class="bootbox">' + html + '</div>'; }
  function loginView(msg) {
    box('<h1>신한FC</h1><p class="sub">구글 계정으로 로그인하면 Google Drive의 고객·수수료 데이터를 불러옵니다.</p>' +
      (msg ? '<p class="warn">' + esc(msg) + '</p>' : '') + '<button type="button" class="btn pri" id="bootLogin">구글 로그인</button>');
    document.getElementById('bootLogin').onclick = function () {
      var b = this; b.disabled = true; b.textContent = '로그인 창을 확인하세요…';
      DRIVE.login().then(start).catch(function (e) { loginView(e.message || String(e)); });
    };
  }
  // 화면 코드는 버전 번호(VER)로 캐시를 쓴다(예전에는 매번 시간값을 붙여 매번 새로 받았음). 화면 코드를 바꾸면 VER를 올린다.
  var VER = '2026-11-12s9';
  function preloadScripts() {   // 데이터를 받는 동안 화면 코드를 미리 받아 둔다(내용은 아직 실행하지 않음)
    SCRIPTS.forEach(function (src) { var l = document.createElement('link'); l.rel = 'preload'; l.as = 'script'; l.href = src + '?v=' + VER; document.head.appendChild(l); });
  }
  function loadScripts() {   // 한꺼번에 받고 순서대로 실행(async=false)
    return Promise.all(SCRIPTS.map(function (src) {
      return new Promise(function (ok, no) {
        var el = document.createElement('script'); el.src = src + '?v=' + VER; el.async = false;
        el.onload = ok; el.onerror = function () { no(new Error(src + ' 를 불러오지 못했습니다.')); };
        document.body.appendChild(el);
      });
    }));
  }
  // 권한 확인이 끝나면(= 접속 허용된 계정) 메뉴와 첫 화면의 틀을 먼저 그린다. 고객·수수료 같은 자료는 아직 없고, 자료가 오면 실제 화면으로 바뀐다.
  function skeleton(cfg) {
    var me = String(DRIVE.email() || '').toLowerCase(), u = (cfg.users || []).filter(function (x) { return String(x.email || '').toLowerCase() === me; })[0], role = u ? u.role : 'FC';
    var all = [['dashboard', '대시보드'], ['team', '팀원'], ['customer', '고객'], ['contract', '계약'], ['performance', '업적'], ['fee', '수수료'], ['activity', '활동'], ['report', '보고서'], ['settings', '관리자']];
    var ids = role === 'ESL' ? all.map(function (x) { return x[0]; }) : (role === 'FC' ? ['dashboard', 'customer', 'contract', 'performance', 'fee'] : ['dashboard', 'team', 'customer', 'contract', 'performance', 'fee']);
    var own = { customer: '내 고객', contract: '내 계약', performance: '내 업적', fee: '내 수수료' };
    var cur = ((location.hash || '#/dashboard').replace('#/', '') || 'dashboard').split('?')[0], curName = '';
    nav.innerHTML = all.filter(function (x) { return ids.indexOf(x[0]) >= 0; }).map(function (x) {
      var nm = role === 'FC' && own[x[0]] ? own[x[0]] : x[1]; if (x[0] === cur) curName = nm;
      return '<a class="' + (x[0] === cur ? 'on' : '') + '"><i></i>' + nm + '</a>';
    }).join('');
    main.innerHTML = '<h1>' + esc(curName || '대시보드') + '</h1><div class="skel"><div class="skb"></div><div class="skb"></div><div class="skb"></div></div><p class="sub" id="bootMsg2">자료를 불러오는 중…</p>';
  }
  function shells() {   // 아직 Drive에서 읽지 않는 자료의 빈 틀(화면이 깨지지 않게)
    if (!window.PERF) window.PERF = { fcs: [], currentFcs: [], current: [], years: {}, from: '', upto: '', made: '' };
    if (!window.CONTRACTS) window.CONTRACTS = { made: '', rows: [] };
  }
  async function start() {
    try {
      box('<h1>신한FC</h1><p class="sub" id="bootMsg">Drive에서 읽는 중…</p>');
      preloadScripts();
      var errs = await DRIVE.loadAll(function (t) { var m = document.getElementById('bootMsg') || document.getElementById('bootMsg2'); if (m) m.textContent = t; }, skeleton);
      shells();
      await loadScripts();
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
