// 접근 범위 적용: 근무 중인 FC가 접속한다는 전제.
// 열람 = 본인 / GSL·SSL은 소속 팀원 / ESL 전부. 수정은 ESL만.
// ESL·SSL·GSL은 화면 위의 "본인 | 팀"으로 보기 범위를 나눠 본다. 기본은 본인.
// 이 PC(로컬) 버전은 ESL 본인 PC 전용이라 로그인 없이 ESL로 열린다. ESL은 "FC 화면으로 보기"로 각 FC의 화면을 미리 볼 수 있다.
// 주의: 아래 걸러내기는 화면용이다. 웹으로 열 때의 실제 차단은 Drive의 FC 폴더별 공유로 한다.
(function () {
  'use strict';
  var CFG = window.ACCESS_CFG, P = window.PERF, C = window.CONTRACTS, K = window.CUSTOMERS;
  var users = CFG.users || [];
  // 로그인한 구글 계정에 맞는 사용자(사용자권한.json의 email). 이메일이 없으면 ESL로 본다(로컬 시험용)
  var loginEmail = String(window.DRIVE && DRIVE.email && DRIVE.email() || '').toLowerCase();
  var real = (loginEmail && users.filter(function (u) { return String(u.email || '').toLowerCase() === loginEmail; })[0]) || users.filter(function (u) { return u.role === 'ESL'; })[0] || users[0];
  var h = location.hash || '';
  var m = h.match(/[?&]as=(\w+)/), mm = h.match(/[?&]mode=(self|team)/);
  var as = '', modeStored = '';
  try {
    if (m) { localStorage.setItem('shfc.viewAs', m[1] === 'me' ? '' : m[1]); }
    if (mm) { localStorage.setItem('shfc.mode', mm[1]); }
    as = localStorage.getItem('shfc.viewAs') || '';
    modeStored = localStorage.getItem('shfc.mode') || '';
  } catch (e) { as = m && m[1] !== 'me' ? m[1] : ''; modeStored = mm ? mm[1] : ''; }
  var me = real;
  if (as && real.role !== 'ESL') as = '';   // 다른 사람 화면으로 보기는 ESL만
  if (as) { var t = users.filter(function (u) { return u.no === as; })[0]; if (t) me = t; }
  var viewing = me !== real;

  function subs(no) {
    var out = [];
    users.forEach(function (u) { if (u.boss === no) { out.push(u.no); out = out.concat(subs(u.no)); } });
    return out;
  }
  var roleAll = me.role === 'ESL';
  var roleScope = roleAll ? null : [me.no].concat(me.role === 'FC' ? [] : subs(me.no));
  var canTeam = me.role !== 'FC';                       // 팀을 가진 직급만 본인/팀 분류가 있다
  // 임시(시험용): ESL의 보기 범위는 '최수진 | 팀' 두 가지. 최수진 = 최수진 로그인 화면(FC, 본인), 팀 = 팀 전체. ESL 본인 화면은 항상 팀 전체.
  var mode = canTeam ? (modeStored === 'team' || (!viewing && me.role === 'ESL') ? 'team' : 'self') : 'self';
  var dataAll = roleAll && mode === 'team';
  var dataScope = mode === 'team' ? roleScope : [me.no];
  var inScope = function (no) { return dataAll || (dataScope || []).indexOf(no) >= 0; };
  var teamSize = roleAll ? users.length : roleScope.length;
  var scopeNames = users.filter(function (u) { return inScope(u.no); }).map(function (u) { return u.name; });

  // 이달 이관 목록 등은 보기 범위(본인|팀)와 상관없이 내 팀(ESL은 전체, GSL·SSL은 소속, FC는 본인) 고객을 모두 볼 수 있게 따로 둔다
  var teamNos = roleAll ? users.map(function (u) { return u.no; }) : (roleScope || []);
  var teamNames = users.filter(function (u) { return teamNos.indexOf(u.no) >= 0; }).map(function (u) { return u.name; });
  K.teamCustomers = K.customers.filter(function (c) { return teamNames.indexOf(c['담당FC']) >= 0; });
  if (!dataAll) {
    P.fcs = P.fcs.filter(function (f) { return inScope(f.no); });
    P.currentFcs = (P.currentFcs || []).filter(function (f) { return inScope(f.no); });
    P.current = P.current.filter(inScope);
    Object.keys(P.years).forEach(function (y) { Object.keys(P.years[y]).forEach(function (no) { if (!inScope(no)) delete P.years[y][no]; }); });
    C.rows = C.rows.filter(function (r) { return inScope(r.no); });
    K.customers = K.customers.filter(function (c) { return scopeNames.indexOf(c['담당FC']) >= 0; });
    var ids = {}; K.teamCustomers.forEach(function (c) { ids[c['고객ID']] = 1; });   // 계약·이력·따행·파일은 팀 고객까지 남긴다(이달 이관 목록·상세용)
    ['contracts', 'history', 'haeng'].forEach(function (k) { Object.keys(K[k]).forEach(function (id) { if (!ids[id]) delete K[k][id]; }); });
    if (window.CUSTFILES) { Object.keys(window.CUSTFILES.byId).forEach(function (id) { if (!ids[id]) delete window.CUSTFILES.byId[id]; }); }
    if (window.FEES) {
      // 수수료명세서는 보기 범위(본인|팀)와 상관없이 내 팀(ESL은 전체, GSL·SSL은 소속, FC는 본인)을 모두 고를 수 있게 따로 둔다
      window.FEES.teamRows = window.FEES.rows.filter(function (r) { return roleAll || (roleScope || []).indexOf(r.no) >= 0; });
      window.FEES.rows = window.FEES.rows.filter(function (r) { return inScope(r.no); });
    }
    Object.keys(K.files).forEach(function (n) { if (teamNames.indexOf(n) < 0) delete K.files[n]; });
  }

  var menus = roleAll ? ['dashboard', 'team', 'customer', 'contract', 'performance', 'fee', 'activity', 'report', 'settings']
    : (me.role === 'FC' ? ['dashboard', 'customer', 'contract', 'performance', 'fee'] : ['dashboard', 'team', 'customer', 'contract', 'performance', 'fee']);
  // FC 선택 드롭다운 정렬: 직급(ESL > SSL > GSL > FC) 먼저, 같은 직급은 사번순. 명단에 없는 FC는 FC로 본다.
  var RANKS = { ESL: 0, SSL: 1, GSL: 2, FC: 3 };
  var fcSortKey = function (no, name, role) {
    var u = users.filter(function (x) { return (no && x.no === no) || (!no && name && x.name === name); })[0];
    var r = RANKS[role || (u && u.role) || 'FC']; if (r == null) r = 3;
    var n = parseInt(no || (u && u.no), 10); return [r, isNaN(n) ? 1e12 : n];
  };
  var fcCompare = function (a, b) { var x = fcSortKey(a.no, a.name, a.role), y = fcSortKey(b.no, b.name, b.role); return x[0] - y[0] || x[1] - y[1] || String(a.name).localeCompare(String(b.name), 'ko'); };
  window.ACCESS = {
    fcCompare: fcCompare,
    user: me, real: real, viewing: viewing, role: me.role, all: roleAll, dataAll: dataAll, scope: dataScope,
    canEdit: me.role === 'ESL' && !viewing, menus: menus, own: me.role === 'FC', inScope: inScope, users: users,
    canTeam: canTeam, mode: mode, teamSize: teamSize,
    viewAs: function (no) { try { localStorage.setItem('shfc.viewAs', no || ''); localStorage.setItem('shfc.mode', 'self'); } catch (e) {} location.hash = (location.hash || '#/dashboard').replace(/[?&](as|mode)=\w+/g, ''); location.reload(); },
    setMode: function (md) { try { localStorage.setItem('shfc.mode', md); } catch (e) {} location.hash = (location.hash || '#/dashboard').replace(/[?&]mode=\w+/g, ''); location.reload(); }
  };
})();
