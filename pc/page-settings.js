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
  var SUBS = [['perm', '권한·공유 현황'], ['common', '공통 설정(M과 공유)'], ['updates', '업데이트 일정']];
  var pageTitle = function () { return window.SETTINGS_ADMIN && window.SETTINGS_ADMIN() ? '관리자' : '설정'; };
  var pageHead = function (sub) { return '<h1>' + pageTitle() + '</h1><p class="sub">' + sub + '</p><div class="tabs">' + SUBS.map(function (x) { return '<button class="' + (state.sub === x[0] ? 'on' : '') + '" data-sub="' + x[0] + '">' + x[1] + '</button>'; }).join('') + '</div>'; };
  var cs = { data: null, err: '', busy: false, msg: '', cals: null, calMsg: '', calBusy: false, calTried: false };   // 공통 설정(Drive)
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

  // ===== 권한·공유 현황 (읽기 전용): 권한 파일(shlife-권한.json)과 실제 Drive 공유가 같은지 보여 준다. 수정은 M의 관리자 화면에서 한다. =====
  var pm = { data: null, err: '', busy: false };
  var roleKo = { owner: '소유자', writer: '편집', commenter: '댓글', reader: '보기' };
  function permHtml() {
    var h = pageHead('권한·공유 현황은 읽기 전용입니다. 수정은 신한FC M의 관리자 화면에서 합니다(Drive 공유도 함께 바뀌기 때문).');
    if (!(window.DRIVE && DRIVE.on())) return h + '<section><p class="mut">Drive에 로그인하면 불러옵니다.</p></section>';
    var d = pm.data;
    if (!d) return h + '<section><p class="mut">' + (pm.err ? esc(pm.err) : '불러오는 중…') + '</p>' + (pm.err ? '<button class="btn" id="pmReload">다시 불러오기</button>' : '') + '</section>';
    var emails = {};
    Object.keys(d.perms.users).forEach(function (e) { emails[e.toLowerCase()] = 1; });
    Object.keys(d.perms.names).forEach(function (e) { emails[e.toLowerCase()] = 1; });
    Object.keys(d.shares).forEach(function (k) { d.shares[k].forEach(function (s) { if (s.role !== 'owner') emails[s.email] = 1; }); });
    var nameOf = function (e) { var n = d.perms.names[e] || d.perms.names[Object.keys(d.perms.names).filter(function (k) { return k.toLowerCase() === e; })[0]]; if (n) return n; var u = (window.ACCESS_CFG.users || []).filter(function (x) { return String(x.email || '').toLowerCase() === e; })[0]; return u ? u.name : ''; };
    var permList = function (e) { var k = Object.keys(d.perms.users).filter(function (x) { return x.toLowerCase() === e; })[0]; return k ? d.perms.users[k] : []; };
    var fcNames = Object.keys(d.perms.fcFolders);
    var rows = Object.keys(emails).sort().map(function (e) {
      var wants = permList(e).map(function (x) { var m = String(x).match(/^(.*):r$/); return { fc: m ? m[1] : x, ro: !!m }; });
      var has = []; fcNames.forEach(function (fc) { (d.shares[fc] || []).forEach(function (s) { if (s.email === e && s.role !== 'owner') has.push({ fc: fc, role: s.role }); }); });
      var wantTxt = wants.map(function (w) { return esc(w.fc) + (w.ro ? '(보기)' : ''); }).join(', ') || '-';
      var hasTxt = has.map(function (x) { return esc(x.fc) + '(' + (roleKo[x.role] || x.role) + ')'; }).join(', ') || '-';
      var bad = [];
      wants.forEach(function (w) { var s = has.filter(function (x) { return x.fc === w.fc; })[0]; if (!s) bad.push(w.fc + ' 공유 없음'); else if (w.ro && s.role !== 'reader') bad.push(w.fc + ' 보기 전용인데 ' + (roleKo[s.role] || s.role)); else if (!w.ro && s.role !== 'writer') bad.push(w.fc + ' 편집인데 ' + (roleKo[s.role] || s.role)); });
      has.forEach(function (x) { if (!wants.some(function (w) { return w.fc === x.fc; })) bad.push(x.fc + ' 공유만 있고 권한 목록에 없음'); });
      var app = (d.appShares.filter(function (s) { return s.email === e; })[0] || {}).role;
      return '<tr><td class="l">' + esc(nameOf(e) || '(이름 없음)') + '</td><td class="l">' + esc(e) + '</td><td class="l">' + wantTxt + '</td><td class="l">' + hasTxt + '</td><td>' + (app ? roleKo[app] || app : '-') + '</td><td class="l">' + (bad.length ? '<span class="warnTxt">⚠ ' + esc(bad.join(' · ')) + '</span>' : '일치') + '</td></tr>';
    }).join('');
    h += '<section><h2>아이디별 접근 <small class="mut">' + rows.split('<tr>').length - 1 + '개 계정</small></h2><div class="wrap"><table><tr><th class="l">이름</th><th class="l">구글 아이디</th><th class="l">권한 목록(shlife-권한.json)</th><th class="l">실제 Drive 폴더 공유</th><th>설정 폴더</th><th class="l">점검</th></tr>' + rows + '</table></div>' +
      '<p class="mut">고객 루트·수수료·업적·계약 폴더는 소유자 전용입니다(점검: ' + (d.rootShared ? '<b>공유가 있음</b>' : '공유 없음') + ').</p><div class="ue-act"><button class="btn" id="pmReload">다시 불러오기</button></div></section>';
    return h;
  }
  function loadPerm() {
    pm.err = ''; pm.busy = true;
    return Promise.all([DRIVE.readPerms(), DRIVE.folderShares(DRIVE.appDirId)]).then(function (r) {
      var perms = r[0], appShares = r[1], names = Object.keys(perms.fcFolders);
      return Promise.all(names.map(function (n) { return DRIVE.folderShares(perms.fcFolders[n]).catch(function () { return []; }); })).then(function (list) {
        var shares = {}; names.forEach(function (n, i) { shares[n] = list[i]; });
        return DRIVE.baseShares().then(function (root) { pm.data = { perms: perms, shares: shares, appShares: appShares, rootShared: root }; pm.busy = false; window.APP.rerender(); });
      });
    }).catch(function (e) { pm.busy = false; pm.err = '불러오지 못했습니다: ' + (e && e.message || e); window.APP.rerender(); });
  }
  function commonHtml() {
    var on = window.DRIVE && DRIVE.on(), d = cs.data;
    var h = '<h1>' + pageTitle() + '</h1><p class="sub">신한FC M의 관리자 › 설정과 같은 파일(shlife-설정.json)을 읽고 씁니다. 여기서 저장하면 M에도 바로 반영됩니다.</p>';
    h += '<div class="tabs">' + SUBS.map(function (x) { return '<button class="' + (state.sub === x[0] ? 'on' : '') + '" data-sub="' + x[0] + '">' + x[1] + '</button>'; }).join('') + '</div>';
    if (!on) return h + '<section><p class="mut">왼쪽 아래 Drive 연결(구글 로그인)을 하면 공통 설정을 불러옵니다.</p></section>';
    if (!d) return h + '<section><p class="mut">' + (cs.err ? esc(cs.err) : '불러오는 중…') + '</p>' + (cs.err ? '<button class="btn" id="csReload">다시 불러오기</button>' : '') + '</section>';
    var g = d.google || {}, cals = (g.calendars || []).map(function (c) { return esc(c.name); }).join(', ') || '(없음)';
    var ids = function (a) { return (a || []).map(function (c) { return esc(c.name || c.id); }).join(', ') || '(없음)'; };
    h += '<section><h2>공통 설정</h2><div class="fg">' +
      '<div class="fr"><span>회사 주소</span><span><input id="csAddr" type="text" value="' + esc(d.companyAddress || '') + '" placeholder="예: 광주광역시 서구 상무중앙로 98" autocomplete="off" style="width:100%"></span></div>' +
      '<div class="fr"><span>카카오 REST API 키</span><span><input id="csKakao" type="password" value="' + esc(d.kakaoRestKey || '') + '" autocomplete="off" spellcheck="false" style="width:100%"></span></div>' +
      '</div><div class="ue-act">' + (window.SETTINGS_ADMIN() ? '<button type="button" class="btn pri" id="csSave">설정 저장</button>' : '') + '<button type="button" class="btn" id="csReload">다시 불러오기</button> <span class="mut">' + esc(cs.msg) + '</span></div></section>';
    h += '<section><h2>구글 캘린더</h2><div class="fg"><div class="fr"><span>구글 계정</span><span>' + esc(g.account || '') + '</span></div><div class="fr"><span>선택한 캘린더</span><span>' + cals + '</span></div><div class="fr"><span>영업일정</span><span>' + ids(g.sales) + '</span></div><div class="fr"><span>정보 캘린더</span><span>' + ids(g.info) + '</span></div></div>';
    if (window.SETTINGS_ADMIN()) {
      h += '<div class="ue-act"><button type="button" class="btn" id="calLoad">' + (cs.cals ? '목록 다시 불러오기' : (cs.calBusy ? '불러오는 중…' : '캘린더 바꾸기')) + '</button> <span class="mut">' + esc(cs.calMsg) + '</span></div>';
      if (cs.cals) {
        var saved = (g.calendars || []).map(function (c) { return c.id; }), sIds = (g.sales || []).map(function (c) { return c.id; }), iIds = (g.info || []).map(function (c) { return c.id; });
        h += '<div class="calpick"><div class="cp-head"><span>캘린더</span><span>보이기</span><span>영업일정</span><span>정보</span></div>' + cs.cals.map(function (c, i) {
          var on = saved.indexOf(c.id) >= 0;
          return '<div class="cp-row"><span class="cp-name"><i class="cp-dot" style="background:' + esc(c.color || '#999') + '"></i>' + esc(c.name) + '</span>' +
            '<label><input type="checkbox" data-pcal="' + i + '"' + (on ? ' checked' : '') + '></label>' +
            '<label><input type="checkbox" data-scal="' + esc(c.id) + '"' + (on && sIds.indexOf(c.id) >= 0 ? ' checked' : '') + '></label>' +
            '<label><input type="checkbox" data-ical="' + esc(c.id) + '"' + (on && (iIds.indexOf(c.id) >= 0 || sIds.indexOf(c.id) >= 0) ? ' checked' : '') + '></label></div>';
        }).join('') + '</div><div class="ue-act"><button type="button" class="btn pri" id="calSave">캘린더 설정 저장</button></div>';
      }
    } else h += '<p class="mut">캘린더 선택은 관리자 계정으로 로그인했을 때만 바꿀 수 있습니다.</p>';
    return h + '</section>';
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
    if (state.sub === 'perm') return permHtml();
    if (state.sub === 'common') return commonHtml();
    if (state.sub !== 'updates') return '<h1>' + pageTitle() + '</h1><p class="sub">세부 메뉴를 선택하세요</p>' + menu;
    var t = today(), ts = ymd(t), h = '<h1>' + pageTitle() + '</h1><p class="sub">업데이트 일정 · 체크하면 오늘 날짜로 완료가 기록됩니다 (이 브라우저에 저장)</p>' + menu;
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
    if (state.sub === 'perm' && window.DRIVE && DRIVE.on() && !pm.data && !pm.busy && !pm.err) loadPerm();
    if (g0('pmReload')) g0('pmReload').onclick = function () { pm.data = null; pm.err = ''; loadPerm(); };
    if (window.DRIVE && DRIVE.on() && !pulled && !cs.busy) { pulled = true; loadCommon(); }
    if (g0('csReload')) g0('csReload').onclick = function () { cs.msg = ''; cs.data = null; loadCommon(); };
    if (g0('csSave')) g0('csSave').onclick = function () {
      var b = g0('csSave'); b.disabled = true; cs.msg = '저장 중…';
      DRIVE.writeSettings({ companyAddress: g0('csAddr').value.trim(), kakaoRestKey: g0('csKakao').value.trim() }).then(function (n) { cs.data = n; cs.msg = '저장했습니다.'; rr(); })
        .catch(function (e) { cs.msg = '저장하지 못했습니다: ' + (e && e.message || e); rr(); });
    };
    if (!cs.cals && !cs.calBusy && cs.data && window.SETTINGS_ADMIN() && DRIVE.calReady && DRIVE.calReady() && !cs.calTried) {
      cs.calTried = true; cs.calBusy = true;
      DRIVE.calendarList().then(function (l) { cs.cals = l; cs.calBusy = false; rr(); }).catch(function (e) { cs.calBusy = false; cs.calMsg = e.message || String(e); rr(); });
    }
    if (g0('calLoad')) g0('calLoad').onclick = function () {
      cs.calMsg = '구글 창에서 캘린더 읽기 권한을 허용해 주세요…';
      DRIVE.calendarList().then(function (l) { cs.cals = l; cs.calMsg = ''; rr(); }).catch(function (e) { cs.calMsg = e.message || String(e); rr(); });
    };
    document.querySelectorAll('.calpick .cp-row').forEach(function (row) {   // M과 같은 규칙: 보이기가 꺼지면 영업일정·정보도 꺼짐, 영업일정을 켜면 정보도 켜짐, 정보를 끄면 영업일정도 꺼짐
      var show = row.querySelector('[data-pcal]'), sales = row.querySelector('[data-scal]'), info = row.querySelector('[data-ical]');
      var sync = function () { sales.disabled = info.disabled = !show.checked; if (!show.checked) sales.checked = info.checked = false; row.classList.toggle('off', !show.checked); };
      show.onchange = sync; sales.onchange = function () { if (sales.checked) info.checked = true; }; info.onchange = function () { if (!info.checked) sales.checked = false; };
      sync();
    });
    if (g0('calSave')) g0('calSave').onclick = function () {
      var list = [].slice.call(document.querySelectorAll('.calpick [data-pcal]')).filter(function (x) { return x.checked; }).map(function (x) { return cs.cals[+x.getAttribute('data-pcal')]; });
      var ck = function (attr) { return [].slice.call(document.querySelectorAll('.calpick [' + attr + ']')).filter(function (x) { return x.checked; }).map(function (x) { return x.getAttribute(attr); }); };
      var sIds = ck('data-scal'), iIds = ck('data-ical'), primary = cs.cals.filter(function (c) { return c.primary; })[0];
      var pick = function (ids) { return list.filter(function (c) { return ids.indexOf(c.id) >= 0; }).map(function (c) { return { id: c.id, name: c.name }; }); };
      var g1 = Object.assign({}, (cs.data && cs.data.google) || {}, { account: primary ? primary.id : (((cs.data || {}).google || {}).account || ''), calendars: list.map(function (c) { return { id: c.id, name: c.name, color: c.color }; }), sales: pick(sIds), info: pick(iIds) });
      var b = g0('calSave'); b.disabled = true; cs.calMsg = '저장 중…';
      DRIVE.writeSettings({ google: g1 }).then(function (n) { cs.data = n; cs.cals = null; cs.calMsg = '캘린더 설정을 저장했습니다.'; rr(); })
        .catch(function (e) { cs.calMsg = '저장하지 못했습니다: ' + (e && e.message || e); rr(); });
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
    pulled = false; cs.data = null; cs.err = ''; cs.cals = null; cs.calTried = false;
    if (DRIVE.on()) { pulled = true; setTimeout(function () { if (DRIVE.on() && !cs.busy) loadCommon(); }, 300); }
  });
  if (window.DRIVE && DRIVE.on()) { pulled = true; setTimeout(function () { if (window.APP) loadCommon(); }, 500); }   // 이미 연결된 상태로 열었을 때
  window.PAGES = window.PAGES || {};
  window.PAGES.settings = { render: render, bind: bind };
})();
