(function () {
  'use strict';
  // 설정 > 업데이트 일정: 매일·매주·매월 해야 할 업데이트와 완료 날짜 기록
  // 저장: 이 브라우저(localStorage). 항목은 '항목 편집'에서 추가·수정·삭제
  var KEY = 'shfc.updates';
  var DEFAULT = [
    { id: 'd1', name: '따행 SOL 반영 확인', freq: 'daily' },
    { id: 'd2', name: '일정 캘린더 점검', freq: 'daily' },
    { id: 'w1', name: '계약고객 업데이트', freq: 'weekly', wd: 1 },
    { id: 'w2', name: '이관고객 업데이트', freq: 'weekly', wd: 1 },
    { id: 'm1', name: '따행 시트 새 대상년월', freq: 'monthly', day: 1 },
    { id: 'm2', name: '월말 이관 리스트 정리', freq: 'monthly', day: 28 }
  ];
  var FREQ = { daily: '매일', weekly: '매주', monthly: '매월' };
  var WDN = ['일', '월', '화', '수', '목', '금', '토'];
  var esc = function (s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  var p2 = function (n) { return n < 10 ? '0' + n : '' + n; };
  var ymd = function (d) { return d.getFullYear() + '-' + p2(d.getMonth() + 1) + '-' + p2(d.getDate()); };
  var addDays = function (d, n) { return new Date(d.getFullYear(), d.getMonth(), d.getDate() + n); };
  var mdL = function (s) { return parseInt(s.slice(5, 7), 10) + '월 ' + parseInt(s.slice(8, 10), 10) + '일'; };
  var today = function () { var n = new Date(); return new Date(n.getFullYear(), n.getMonth(), n.getDate()); };

  var store = (function () {
    try { var s = JSON.parse(localStorage.getItem(KEY) || 'null'); if (s && Array.isArray(s.items)) return { items: s.items, done: s.done || {} }; } catch (e) {}
    return { items: DEFAULT.map(function (x) { return Object.assign({}, x); }), done: {} };
  })();
  var state = { sub: '', filter: 'all', draft: null };
  var SUBS = [['common', '공통 설정(M과 공유)'], ['updates', '업데이트 일정']];
  var cs = { data: null, err: '', busy: false, msg: '' };   // 공통 설정(Drive)
  var pulled = false;
  function save() {
    Object.keys(store.done).forEach(function (k) { if (!store.items.some(function (x) { return x.id === k; })) delete store.done[k]; else store.done[k] = store.done[k].slice(-120); });
    var ok = true;
    try { localStorage.setItem(KEY, JSON.stringify(store)); } catch (e) { ok = false; }
    pushUpdates();
    return ok;
  }
  // 업데이트 일정은 Drive 설정 파일의 desktopUpdates 항목에도 저장 (연결되어 있을 때)
  function pushUpdates() { if (window.DRIVE && DRIVE.on()) DRIVE.writeSettings({ desktopUpdates: store }).catch(function () { /* 연결 실패 시 이 브라우저 저장만 */ }); }
  function commonHtml() {
    var on = window.DRIVE && DRIVE.on(), d = cs.data;
    var h = '<h1>설정</h1><p class="sub">신한FC M의 관리자 › 설정과 같은 파일(shlife-설정.json)을 읽고 씁니다. 여기서 저장하면 M에도 바로 반영됩니다.</p>';
    h += '<div class="tabs">' + SUBS.map(function (x) { return '<button class="' + (state.sub === x[0] ? 'on' : '') + '" data-sub="' + x[0] + '">' + x[1] + '</button>'; }).join('') + '</div>';
    if (!on) return h + '<section><p class="mut">왼쪽 아래 Drive 연결(구글 로그인)을 하면 공통 설정을 불러옵니다.</p></section>';
    if (!d) return h + '<section><p class="mut">' + (cs.err ? esc(cs.err) : '불러오는 중…') + '</p>' + (cs.err ? '<button class="btn" id="csReload">다시 불러오기</button>' : '') + '</section>';
    var g = d.google || {}, cals = (g.calendars || []).map(function (c) { return esc(c.name); }).join(', ') || '(없음)';
    var ids = function (a) { return (a || []).map(function (c) { return esc(c.name || c.id); }).join(', ') || '(없음)'; };
    h += '<section><h2>공통 설정</h2><div class="fg">' +
      '<div class="fr"><span>회사 주소</span><span><input id="csAddr" type="text" value="' + esc(d.companyAddress || '') + '" placeholder="예: 광주광역시 서구 상무중앙로 98" autocomplete="off" style="width:100%"></span></div>' +
      '<div class="fr"><span>카카오 REST API 키</span><span><input id="csKakao" type="password" value="' + esc(d.kakaoRestKey || '') + '" autocomplete="off" spellcheck="false" style="width:100%"></span></div>' +
      '</div><div class="ue-act">' + (window.SETTINGS_ADMIN() ? '<button type="button" class="btn pri" id="csSave">설정 저장</button>' : '') + '<button type="button" class="btn" id="csReload">다시 불러오기</button> <span class="mut">' + esc(cs.msg) + '</span></div></section>';
    h += '<section><h2>구글 캘린더 (보기 전용)</h2><div class="fg"><div class="fr"><span>구글 계정</span><span>' + esc(g.account || '') + '</span></div><div class="fr"><span>선택한 캘린더</span><span>' + cals + '</span></div><div class="fr"><span>영업일정</span><span>' + ids(g.sales) + '</span></div><div class="fr"><span>정보 캘린더</span><span>' + ids(g.info) + '</span></div></div><p class="mut">캘린더 선택은 M의 관리자 › 설정에서 바꿉니다.</p></section>';
    return h;
  }
  function loadCommon() {
    cs.err = ''; cs.busy = true;
    return DRIVE.readSettings().then(function (d) {
      cs.data = d; cs.busy = false;
      if (d.desktopUpdates && Array.isArray(d.desktopUpdates.items)) { store = { items: d.desktopUpdates.items, done: d.desktopUpdates.done || {} }; try { localStorage.setItem(KEY, JSON.stringify(store)); } catch (e) { /* 무시 */ } }
      else DRIVE.writeSettings({ desktopUpdates: store }).catch(function () { /* 무시 */ });
      window.APP.rerender();
    }).catch(function (e) { cs.busy = false; cs.err = '불러오지 못했습니다: ' + (e && e.message || e); window.APP.rerender(); });
  }

  function period(it, t) {
    if (it.freq === 'daily') return [ymd(t), ymd(t)];
    if (it.freq === 'weekly') { var mon = addDays(t, -((t.getDay() + 6) % 7)); return [ymd(mon), ymd(addDays(mon, 6))]; }
    return [ymd(new Date(t.getFullYear(), t.getMonth(), 1)), ymd(new Date(t.getFullYear(), t.getMonth() + 1, 0))];
  }
  function due(it, t) {
    if (it.freq === 'daily') return t;
    if (it.freq === 'weekly') { var mon = addDays(t, -((t.getDay() + 6) % 7)); return addDays(mon, (((+it.wd || 0) % 7) + 6) % 7); }
    var last = new Date(t.getFullYear(), t.getMonth() + 1, 0).getDate();
    return new Date(t.getFullYear(), t.getMonth(), Math.min(Math.max(+it.day || 1, 1), last));
  }
  function status(it, t) {
    var pr = period(it, t), all = (store.done[it.id] || []).slice().sort();
    var inP = all.filter(function (s) { return s >= pr[0] && s <= pr[1]; }), lastAll = all[all.length - 1];
    var diff = Math.round((due(it, t) - t) / 864e5);
    var sched = it.freq === 'daily' ? '매일' : it.freq === 'weekly' ? '매주 ' + WDN[+it.wd || 0] + '요일' : '매월 ' + (+it.day || 1) + '일';
    var last = lastAll ? '마지막 ' + mdL(lastAll) : '기록 없음';
    if (inP.length) return { k: 'ok', pill: '완료', sub: sched + ' · ' + mdL(inP[inP.length - 1]) + ' 완료' };
    if (diff === 0) return { k: 'due', pill: '오늘', sub: sched + ' · ' + last };
    if (diff < 0) return { k: 'late', pill: (-diff) + '일 지연', sub: sched + ' · ' + last };
    return { k: 'wait', pill: diff + '일 남음', sub: sched + ' · ' + last };
  }

  function render() {
    var menu = '<div class="tabs">' + SUBS.map(function (x) { return '<button class="' + (state.sub === x[0] ? 'on' : '') + '" data-sub="' + x[0] + '">' + x[1] + '</button>'; }).join('') + '</div>';
    if (state.sub === 'common') return commonHtml();
    if (state.sub !== 'updates') return '<h1>설정</h1><p class="sub">세부 메뉴를 선택하세요</p>' + menu;
    var t = today(), ts = ymd(t), h = '<h1>설정</h1><p class="sub">업데이트 일정 · 체크하면 오늘 날짜로 완료가 기록됩니다 (이 브라우저에 저장)</p>' + menu;
    if (state.draft) {
      h += '<section><h2>항목 편집</h2><div class="ue-list">' + state.draft.map(function (it, i) {
        return '<div class="ue-row" data-i="' + i + '"><input type="text" data-f="name" value="' + esc(it.name) + '" placeholder="항목 이름" maxlength="40">' +
          '<select data-f="freq">' + Object.keys(FREQ).map(function (k) { return '<option value="' + k + '"' + (it.freq === k ? ' selected' : '') + '>' + FREQ[k] + '</option>'; }).join('') + '</select>' +
          (it.freq === 'weekly' ? '<select data-f="wd">' + [1, 2, 3, 4, 5, 6, 0].map(function (w) { return '<option value="' + w + '"' + ((+it.wd || 0) === w ? ' selected' : '') + '>' + WDN[w] + '요일</option>'; }).join('') + '</select>' : '') +
          (it.freq === 'monthly' ? '<select data-f="day">' + Array.from({ length: 31 }, function (_, k) { return k + 1; }).map(function (n) { return '<option value="' + n + '"' + ((+it.day || 1) === n ? ' selected' : '') + '>' + n + '일</option>'; }).join('') + '</select>' : '') +
          '<button type="button" class="btn" data-udel="' + i + '">삭제</button></div>';
      }).join('') + '</div><div class="ue-act"><button type="button" class="btn" id="uAdd">항목 추가</button><button type="button" class="btn" id="uCancel">취소</button><button type="button" class="btn pri" id="uSave">저장</button></div></section>';
      return h;
    }
    var mon = addDays(t, -((t.getDay() + 6) % 7)), daily = store.items.filter(function (x) { return x.freq === 'daily'; });
    var week = [0, 1, 2, 3, 4, 5, 6].map(function (i) {
      var d = addDays(mon, i), s = ymd(d), dot = '';
      if (daily.length) { var all = daily.every(function (x) { return (store.done[x.id] || []).indexOf(s) >= 0; }); dot = '<i class="ud ' + (all ? 'g' : s < ts ? 'r' : 'o') + '"></i>'; }
      return '<div class="' + (s === ts ? 't' : '') + '">' + WDN[d.getDay()] + '<br><small>' + d.getDate() + '</small>' + dot + '</div>';
    }).join('');
    var groups = [['daily', '오늘 (매일)'], ['weekly', '이번 주 (매주)'], ['monthly', '이번 달 (매월)']].filter(function (g) { return state.filter === 'all' || state.filter === g[0]; });
    var tabs = [['all', '전체'], ['daily', '매일'], ['weekly', '매주'], ['monthly', '매월']].map(function (x) { return '<button class="' + (state.filter === x[0] ? 'on' : '') + '" data-ufilter="' + x[0] + '">' + x[1] + '</button>'; }).join('');
    h += '<section><h2>업데이트 일정 <small class="mut">' + (t.getMonth() + 1) + '월 ' + t.getDate() + '일 (' + WDN[t.getDay()] + ') 기준</small></h2><div class="tabs">' + tabs + '</div><div class="uweek">' + week + '</div>' +
      groups.map(function (g) {
        var its = store.items.filter(function (x) { return x.freq === g[0]; });
        return '<h3 class="uh">' + g[1] + '</h3>' + (its.length ? its.map(function (it) {
          var st = status(it, t);
          return '<div class="urow ' + (st.k === 'ok' ? 'ok' : '') + '"><button type="button" class="uchk" data-utoggle="' + esc(it.id) + '" aria-pressed="' + (st.k === 'ok') + '" aria-label="' + esc(it.name) + ' 완료 표시">' + (st.k === 'ok' ? '✓' : '') + '</button><div class="utx"><b>' + esc(it.name) + '</b><small>' + esc(st.sub) + '</small></div><span class="upill p-' + st.k + '">' + esc(st.pill) + '</span></div>';
        }).join('') : '<p class="mut">항목이 없습니다.</p>');
      }).join('') + '<div class="ue-act"><button type="button" class="btn" id="uEdit">항목 편집</button></div></section>';
    return h;
  }

  function bind() {
    var rr = function () { window.APP.rerender(); };
    document.querySelectorAll('[data-sub]').forEach(function (b) { b.onclick = function () { var k = b.getAttribute('data-sub'); state.sub = state.sub === k ? '' : k; state.draft = null; rr(); }; });
    var g0 = function (id) { return document.getElementById(id); };
    if (window.DRIVE && DRIVE.on() && !pulled && !cs.busy) { pulled = true; loadCommon(); }
    if (g0('csReload')) g0('csReload').onclick = function () { cs.msg = ''; cs.data = null; loadCommon(); };
    if (g0('csSave')) g0('csSave').onclick = function () {
      var b = g0('csSave'); b.disabled = true; cs.msg = '저장 중…';
      DRIVE.writeSettings({ companyAddress: g0('csAddr').value.trim(), kakaoRestKey: g0('csKakao').value.trim() }).then(function (n) { cs.data = n; cs.msg = '저장했습니다.'; rr(); })
        .catch(function (e) { cs.msg = '저장하지 못했습니다: ' + (e && e.message || e); rr(); });
    };
    var q = function (sel, fn) { document.querySelectorAll(sel).forEach(fn); };
    q('[data-ufilter]', function (b) { b.onclick = function () { state.filter = b.getAttribute('data-ufilter'); rr(); }; });
    q('[data-utoggle]', function (b) {
      b.onclick = function () {
        var id = b.getAttribute('data-utoggle'), it = store.items.filter(function (x) { return x.id === id; })[0]; if (!it) return;
        var t = today(), pr = period(it, t), ok = status(it, t).k === 'ok';
        var arr = (store.done[id] || []).filter(function (s) { return !(s >= pr[0] && s <= pr[1]); });
        if (!ok) arr.push(ymd(t));
        arr.sort(); store.done[id] = arr; save(); rr();
      };
    });
    var g = function (id) { return document.getElementById(id); };
    if (g('uEdit')) g('uEdit').onclick = function () { state.draft = store.items.map(function (x) { return Object.assign({}, x); }); rr(); };
    if (g('uCancel')) g('uCancel').onclick = function () { state.draft = null; rr(); };
    if (g('uAdd')) g('uAdd').onclick = function () { state.draft.push({ id: 'i' + Date.now().toString(36), name: '', freq: 'weekly', wd: 1 }); rr(); var ins = document.querySelectorAll('[data-f="name"]'); if (ins.length) ins[ins.length - 1].focus(); };
    q('[data-udel]', function (b) { b.onclick = function () { state.draft.splice(+b.getAttribute('data-udel'), 1); rr(); }; });
    q('.ue-row [data-f]', function (el) {
      var apply = function () {
        var it = state.draft[+el.closest('.ue-row').getAttribute('data-i')], f = el.getAttribute('data-f');
        if (f === 'name') it.name = el.value;
        else if (f === 'freq') { it.freq = el.value; if (it.freq === 'weekly' && it.wd == null) it.wd = 1; if (it.freq === 'monthly' && !it.day) it.day = 1; }
        else it[f] = +el.value;
      };
      if (el.getAttribute('data-f') === 'name') el.oninput = apply; else el.onchange = function () { apply(); if (el.getAttribute('data-f') === 'freq') rr(); };
    });
    if (g('uSave')) g('uSave').onclick = function () {
      store.items = state.draft.map(function (x) { return Object.assign({}, x, { name: String(x.name || '').trim() }); }).filter(function (x) { return x.name; });
      state.draft = null;
      if (!save()) alert('이 브라우저에 저장하지 못했습니다.');
      rr();
    };
  }
  var ADMINS = ['blue10kp@gmail.com'];   // M의 config.admins와 같은 관리자 계정(ESL)
  window.SETTINGS_ADMIN = function () { return !!(window.DRIVE && DRIVE.on() && ADMINS.indexOf(String(DRIVE.email()).toLowerCase()) >= 0); };
  // 로그인하면 바로 공통 설정(관리자 설정)을 불러옴 — 설정 화면을 열지 않아도
  if (window.DRIVE) DRIVE.onChange(function () {
    pulled = false; cs.data = null; cs.err = '';
    if (DRIVE.on()) { pulled = true; setTimeout(function () { if (DRIVE.on() && !cs.busy) loadCommon(); }, 300); }
  });
  if (window.DRIVE && DRIVE.on()) { pulled = true; setTimeout(function () { if (window.APP) loadCommon(); }, 500); }   // 이미 연결된 상태로 열었을 때
  window.PAGES = window.PAGES || {};
  window.PAGES.settings = { render: render, bind: bind };
})();
