(function () {
  'use strict';
  var D = window.PERF;
  // 현재 재직 FC를 사번순으로 먼저, 퇴직 FC는 괄호 한 줄(참조)로 뒤에. 색은 FC마다 원래대로.
  var FCL = (function () {
    var nk = function (n) { return parseInt(n, 10) || 1e12; };
    var cur = (D.currentFcs || []).map(function (f) { return { no: f.no, name: f.name }; }).sort(function (a, b) { return nk(a.no) - nk(b.no); });
    var curNos = cur.map(function (f) { return f.no; });
    var ret = D.fcs.filter(function (f) { return curNos.indexOf(f.no) < 0; }).sort(function (a, b) { return nk(a.no) - nk(b.no); });
    var orig = D.fcs.map(function (f) { return f.no; });
    var zero = function () { return { n: [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], cmip: [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], canp: [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0] }; };
    Object.keys(D.years).forEach(function (y) {
      var yd = D.years[y];
      cur.forEach(function (f) { if (!yd[f.no]) yd[f.no] = zero(); });
      if (ret.length) {
        var g = zero();
        ret.forEach(function (f) { ['n', 'cmip', 'canp'].forEach(function (k) { (yd[f.no] ? yd[f.no][k] : g[k]).forEach(function (v, i) { g[k][i] += v; }); }); });
        yd.RET = g;
      }
    });
    var list = cur.map(function (f) { var k = orig.indexOf(f.no); return { no: f.no, name: f.name, ci: k >= 0 ? k : orig.length }; });
    if (ret.length) list.push({ no: 'RET', name: '(' + ret.map(function (f) { return f.name; }).join(', ') + ')', group: true, ci: -1 });
    return list;
  })();
  var CIX = FCL.map(function (f) { return f.ci; });
  var MENUS = [
    { id: 'dashboard', name: '대시보드' },
    { id: 'team', name: '팀원' },
    { id: 'customer', name: '고객' },
    { id: 'contract', name: '계약' },
    { id: 'performance', name: '업적' },
    { id: 'fee', name: '수수료' },
    { id: 'activity', name: '활동', soon: true },
    { id: 'report', name: '보고서', soon: true },
    { id: 'settings', name: '설정' }
  ];
  var AC = window.ACCESS;
  var LABEL = { customer: '고객', contract: '계약', performance: '업적', fee: '수수료' };
  if (window.SETTINGS_ADMIN && window.SETTINGS_ADMIN()) LABEL.settings = '관리자';
  if (AC.own) { LABEL = { customer: '내 고객', contract: '내 계약', performance: '내 업적', fee: '내 수수료' }; }
  MENUS = MENUS.filter(function (m) { return AC.menus.indexOf(m.id) >= 0; }).map(function (m) { return Object.assign({}, m, { name: LABEL[m.id] || m.name }); });
  var COLORS = ['#2B5FD9', '#1E7F74', '#B5651D', '#7C4DBB', '#C0392B', '#0E8FB3', '#8A9A1B', '#D6336C', '#5B6B91', '#E0A800', '#3C8D40', '#9B59B6', '#16A085', '#E67E22', '#34495E', '#C2185B', '#7F8C8D', '#2E86C1', '#A04000', '#117A65'];
  var METRICS = [
    { k: 'n', t: '월별 건수 (기여율 반영)', u: '건', f: function (v) { return v.toLocaleString('ko-KR', { minimumFractionDigits: v % 1 ? 1 : 0, maximumFractionDigits: 1 }); }, ax: function (v) { return Math.round(v) + '건'; } },
    { k: 'cmip', t: '월별 CMIP (월환산보험료, 원, 기여율 반영)', u: '원', f: function (v) { return Math.round(v).toLocaleString('ko-KR'); }, ax: function (v) { return Math.round(v / 10000).toLocaleString('ko-KR') + '만'; } },
    { k: 'canp', t: '월별 CANP (원, 기여율 반영)', u: '원', f: function (v) { return Math.round(v).toLocaleString('ko-KR'); }, ax: function (v) { return Math.round(v / 10000).toLocaleString('ko-KR') + '만'; } }
  ];
  var sum = function (a) { return a.reduce(function (x, y) { return x + y; }, 0); };
  var col = function (i) { var c = CIX[i]; return c < 0 ? '#C3CAD8' : COLORS[c % COLORS.length]; };
  var fl = function (Y) {
    return FCL.map(function (f, j) { return Object.assign({ j: j }, f); }).filter(function (f) {
      return f.group || isCur(f.no) || ['n', 'cmip', 'canp'].some(function (k) { return sum(D.years[Y][f.no][k]) > 0; });
    });
  };
  var years = Object.keys(D.years).sort();
  var perfTab = years[0];
  var cmpYears = years.filter(function (y) { return y < D.upto.slice(0, 4); }).slice(0, 2);
  var isCur = function (no) { return D.current.indexOf(no) >= 0; };

  // 데이터 갱신 시각: 읽은 FC별 data.json의 made(UTC). 여러 FC면 가장 오래된 것 기준(마우스를 올리면 FC별)
  function refreshDStamp() {
    var el = document.getElementById('dstamp'); if (!el) return;
    var st = window.DATASTAMP || {}, ks = Object.keys(st).filter(function (k) { return st[k]; });
    var fmt = function (m) { var d = new Date(m + (/Z|[+-]\d\d:?\d\d$/.test(m) ? '' : 'Z')); if (isNaN(d)) return ''; var p = function (n) { return (n < 10 ? '0' : '') + n; }; return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate()) + ' ' + p(d.getHours()) + ':' + p(d.getMinutes()); };
    if (!ks.length) { el.textContent = ''; el.title = ''; return; }
    ks.sort(function (a, b) { return st[a] < st[b] ? -1 : 1; });
    el.textContent = 'Drive 데이터 ' + fmt(st[ks[0]]) + (ks.length > 1 ? ' (가장 오래된 FC 기준)' : '');
    el.title = ks.map(function (k) { return k + ' ' + fmt(st[k]); }).join(' · ');
  }
  window.refreshDStamp = refreshDStamp;
  function userBox() {
    var foot = document.querySelector('.foot'); if (!foot) return;
    var h = '<div class="ub"><b>' + AC.user.name + '</b> · ' + AC.role + (AC.canEdit ? '' : ' · 열람 전용') + '</div>';
    if (AC.real.role === 'ESL') {
      h += '<select id="asSel" class="asSel"><option value="">내 화면</option>' + AC.users.filter(function (u) { return u.role !== 'ESL'; }).sort(AC.fcCompare).map(function (u) { return '<option value="' + u.no + '"' + (AC.user.no === u.no ? ' selected' : '') + '>' + u.name + '</option>'; }).join('') + '</select>';
    }
    var dv = window.DRIVE;
    if (dv) h += '<div class="drv">' + (dv.on() ? '<span class="drvok">● Google Drive 연결됨</span><small>' + (dv.email() || '') + '</small><button type="button" id="drvOut" class="asSel">연결 해제</button>' : '<button type="button" id="drvIn" class="asSel">Google Drive 연결 (로그인)</button><small>연결하면 고객 정보 수정이 엑셀에 바로 저장됩니다</small>') + '</div>';
    h += '<div style="margin-top:8px">데탑 v3 · 2026-11-12<br><span id="stamp"></span><br><span id="dstamp"></span></div>';
    foot.innerHTML = h;
    refreshDStamp();
    var sel = document.getElementById('asSel'); if (sel) sel.onchange = function () { AC.viewAs(sel.value); };
    var di = document.getElementById('drvIn'); if (di) di.onclick = function () { di.disabled = true; di.textContent = '로그인 창을 확인하세요…'; DRIVE.login().then(function () { location.reload(); }).catch(function (e) { alert(e.message || e); userBox(); }); };
    var dout = document.getElementById('drvOut'); if (dout) dout.onclick = function () { DRIVE.logout(); };
  }
  // 구글 Drive 연결 뒤: 엑셀의 최신 값으로 고객 데이터를 다시 읽는다
  function driveSync() {
    userBox();
    var st = document.getElementById('stamp'); if (st) st.textContent = 'Drive에서 읽는 중…';
    return DRIVE.refresh().then(function (r) {
      if (window.PAGES && PAGES.customer && PAGES.customer.reloadExtra) PAGES.customer.reloadExtra();
      route();
      if (r.errors && r.errors.length) alert('일부 FC 파일을 읽지 못했습니다:\n' + r.errors.join('\n'));
    }).catch(function (e) { alert('Drive에서 읽지 못했습니다: ' + (e.message || e)); route(); });
  }
  function nav(active) {
    document.getElementById('nav').innerHTML = MENUS.map(function (m) {
      return '<a href="#/' + m.id + '" class="' + (m.id === active ? 'on ' : '') + (m.soon ? 'soon' : '') + '"><i></i>' + m.name + (m.soon ? '<small>준비 중</small>' : '') + '</a>';
    }).join('');
  }

  function stacked(m, Y) {
    var W = 820, H = 300, L = 64, R = 10, T = 14, B = 36;
    var tot = Array.from({ length: 12 }, function (_, i) { return sum(FCL.map(function (f) { return D.years[Y][f.no][m.k][i]; })); });
    var mx = Math.max.apply(null, [1].concat(tot)) * 1.08, bw = (W - L - R) / 12;
    var s = '<svg viewBox="0 0 ' + W + ' ' + H + '" width="100%">';
    for (var g = 0; g <= 4; g++) {
      var y = T + (H - T - B) * (1 - g / 4);
      s += '<line x1="' + L + '" x2="' + (W - R) + '" y1="' + y + '" y2="' + y + '" stroke="#E1E4EC"/><text x="' + (L - 8) + '" y="' + (y + 4) + '" text-anchor="end" font-size="12" fill="#5A667E">' + m.ax(mx * g / 4) + '</text>';
    }
    for (var i = 0; i < 12; i++) {
      var y0 = H - B, x = L + i * bw + bw * 0.15, w = bw * 0.7;
      FCL.forEach(function (f, j) {
        var v = D.years[Y][f.no][m.k][i]; if (!v) return;
        var h = (H - T - B) * v / mx; y0 -= h;
        s += '<rect x="' + x + '" y="' + y0 + '" width="' + w + '" height="' + h + '" fill="' + col(j) + '"><title>' + (i + 1) + '월 ' + f.name + ': ' + m.f(v) + m.u + '</title></rect>';
      });
      s += '<text x="' + (x + w / 2) + '" y="' + (H - B + 18) + '" text-anchor="middle" font-size="12" fill="#5A667E">' + (i + 1) + '월</text>';
    }
    return s + '</svg>';
  }

  function shareBars(m, Y) {
    var F = fl(Y), tots = F.map(function (f) { return sum(D.years[Y][f.no][m.k]); }), all = sum(tots) || 1;
    var rowH = 24, H = F.length * rowH + 8, W = 470, L = 150;
    var maxp = Math.max.apply(null, tots.map(function (t) { return t / all; }).concat([0.01]));
    var s = '<svg viewBox="0 0 ' + W + ' ' + H + '" width="100%">';
    F.forEach(function (f, ix) {
      var p = tots[ix] / all, y = 4 + ix * rowH, w = Math.max((W - L - 70) * p / maxp, 0);
      s += '<text x="' + (L - 8) + '" y="' + (y + 16) + '" text-anchor="end" font-size="13" fill="#14213D">' + (f.group ? '(퇴직 FC)' : f.name) + '</text><rect x="' + L + '" y="' + (y + 3) + '" width="' + w + '" height="16" rx="3" fill="' + col(f.j) + '"/><text x="' + (L + w + 6) + '" y="' + (y + 16) + '" font-size="13" fill="#5A667E">' + (p * 100).toFixed(1) + '%</text>';
    });
    return s + '</svg>';
  }

  function table(m, Y) {
    var F = fl(Y), tots = F.map(function (f) { return sum(D.years[Y][f.no][m.k]); }), all = sum(tots) || 1;
    var h = '<div class="wrap"><table><tr><th>FC</th><th>사번</th>' + Array.from({ length: 12 }, function (_, i) { return '<th>' + (i + 1) + '월</th>'; }).join('') + '<th>연 합계</th><th>점유율</th></tr>';
    F.forEach(function (f, ix) {
      h += '<tr class="' + (isCur(f.no) ? 'cur' : '') + '"><td class="' + (f.group ? 'wrapc' : '') + '"><span class="sw" style="background:' + col(f.j) + '"></span>' + f.name + '</td><td>' + (f.group ? '-' : f.no) + '</td>' +
        D.years[Y][f.no][m.k].map(function (v) { return '<td>' + (v ? m.f(v) : '-') + '</td>'; }).join('') + '<td><b>' + m.f(tots[ix]) + '</b></td><td>' + (tots[ix] / all * 100).toFixed(1) + '%</td></tr>';
    });
    var col12 = Array.from({ length: 12 }, function (_, i) { return sum(FCL.map(function (f) { return D.years[Y][f.no][m.k][i]; })); });
    h += '<tr class="tot"><td>합계</td><td></td>' + col12.map(function (v) { return '<td>' + (v ? m.f(v) : '-') + '</td>'; }).join('') + '<td>' + m.f(all) + '</td><td>100%</td></tr></table></div>';
    return h;
  }

  function legend(Y) {
    var F = Y ? fl(Y) : FCL.map(function (f, j) { return Object.assign({ j: j }, f); });
    return '<div class="legend">' + F.map(function (f) { return '<span><i class="sw" style="background:' + col(f.j) + '"></i>' + f.name + '</span>'; }).join('') + '</div>';
  }

  function yearView(Y) {
    var T = function (k) { return sum(FCL.map(function (f) { return sum(D.years[Y][f.no][k]); })); };
    var part = Y === D.upto.slice(0, 4) ? ' (1~' + parseInt(D.upto.slice(5), 10) + '월)' : '';
    var h = '<div class="kpis"><div class="k"><div class="l">' + Y + '년' + part + ' 건수 (기여율 반영)</div><div class="v">' + METRICS[0].f(T('n')) + '<small>건</small></div></div>' +
      '<div class="k"><div class="l">' + Y + '년' + part + ' CMIP (월환산보험료)</div><div class="v">' + Math.round(T('cmip') / 10000).toLocaleString('ko-KR') + '<small>만원</small></div></div>' +
      '<div class="k"><div class="l">' + Y + '년' + part + ' CANP</div><div class="v">' + Math.round(T('canp') / 10000).toLocaleString('ko-KR') + '<small>만원</small></div></div></div>';
    METRICS.forEach(function (m) {
      h += '<section><h2>' + Y + '년 ' + m.t + '</h2>' + table(m, Y) + '<div class="two"><div><b style="font-size:15px">월별 FC 누적 막대</b>' + stacked(m, Y) + '</div><div><b style="font-size:15px">FC별 점유율 (연간)</b>' + shareBars(m, Y) + '</div></div>' + legend(Y) + '</section>';
    });
    return h;
  }

  function cmpView() {
    var A = cmpYears[0], B2 = cmpYears[1], h = '';
    METRICS.forEach(function (m) {
      var a = FCL.map(function (f) { return sum(D.years[A][f.no][m.k]); });
      var b = FCL.map(function (f) { return sum(D.years[B2][f.no][m.k]); });
      var SA = sum(a) || 1, SB = sum(b) || 1;
      h += '<section><h2>' + m.t.replace('월별 ', '') + ' — ' + A + ' vs ' + B2 + '</h2><div class="wrap"><table><tr><th>FC</th><th>사번</th><th>' + A + '년</th><th>' + A + ' 점유율</th><th>' + B2 + '년</th><th>' + B2 + ' 점유율</th><th>증감</th><th>증감률</th></tr>';
      FCL.forEach(function (f, j) {
        if (!a[j] && !b[j]) return;
        var diff = b[j] - a[j], pct = a[j] ? diff / a[j] * 100 : null;
        h += '<tr class="' + (isCur(f.no) ? 'cur' : '') + '"><td class="' + (f.group ? 'wrapc' : '') + '"><span class="sw" style="background:' + col(j) + '"></span>' + f.name + '</td><td>' + (f.group ? '-' : f.no) + '</td><td>' + m.f(a[j]) + '</td><td>' + (a[j] / SA * 100).toFixed(1) + '%</td><td>' + m.f(b[j]) + '</td><td>' + (b[j] / SB * 100).toFixed(1) + '%</td><td class="' + (diff >= 0 ? 'up' : 'dn') + '">' + (diff >= 0 ? '+' : '') + m.f(diff) + '</td><td class="' + (diff >= 0 ? 'up' : 'dn') + '">' + (pct === null ? '-' : (pct >= 0 ? '+' : '') + pct.toFixed(0) + '%') + '</td></tr>';
      });
      var dd = SB - SA;
      h += '<tr class="tot"><td>합계</td><td></td><td>' + m.f(SA) + '</td><td>100%</td><td>' + m.f(SB) + '</td><td>100%</td><td>' + (dd >= 0 ? '+' : '') + m.f(dd) + '</td><td>' + ((SB - SA) / SA * 100 >= 0 ? '+' : '') + ((SB - SA) / SA * 100).toFixed(0) + '%</td></tr></table></div>' + legend() + '</section>';
    });
    return h;
  }

  function performance() {
    var tabs = years.map(function (y) {
      var part = y === D.upto.slice(0, 4) ? ' (1~' + parseInt(D.upto.slice(5), 10) + '월)' : '';
      return '<button class="' + (perfTab === y ? 'on' : '') + '" data-tab="' + y + '">' + y + '년' + part + '</button>';
    });
    if (cmpYears.length === 2) tabs.push('<button class="' + (perfTab === 'cmp' ? 'on' : '') + '" data-tab="cmp">' + cmpYears[0] + ' vs ' + cmpYears[1] + '</button>');
    return '<h1>업적</h1><p class="sub">청약계약List조회 기준 · 성립 계약만 · 건수·CMIP·CANP 모두 기여율 반영 · 현재 FC는 사번순, 퇴직 FC는 뒤에 괄호로 참조 · 수집 범위 ' + D.from + ' ~ ' + D.upto + ' · 만든 날 ' + D.made +
      '</p><div class="tabs">' + tabs.join('') + '</div>' +
      (perfTab === 'cmp' ? cmpView() : yearView(perfTab));
  }

  function soon(name) {
    return '<h1>' + name + '</h1><div class="soonbox">이 메뉴는 아직 만들지 않았습니다.<br>업적 메뉴부터 만들고 있으며, 이어서 하나씩 추가합니다.</div>';
  }


  // 표 정렬 규칙: 맨 윗줄(머리글)은 항상 가운데, 데이터는 가운데, 금액이 들어간 열만 오른쪽
  var MONEY = /보험료|CMIP|CANP|월납|금액|환산|원\)|만원/;
  function alignTables() {
    document.querySelectorAll('#main table').forEach(function (t) {
      var hr = null; for (var i = 0; i < t.rows.length; i++) { if (t.rows[i].querySelector('th')) { hr = t.rows[i]; break; } }
      if (!hr) return;
      var sec = t.closest('section'), h2 = sec && sec.querySelector('h2'), secMoney = !!(h2 && /CMIP|CANP|원\)/.test(h2.textContent || ''));
      var money = Array.prototype.map.call(hr.cells, function (c) { var tx = (c.textContent || '').trim(); return MONEY.test(tx) || (secMoney && /^(\d+월|연 합계)$/.test(tx)); });
      var left = Array.prototype.map.call(hr.cells, function (c) { return /상품명|주소|지역/.test(c.textContent || ''); });
      Array.prototype.forEach.call(t.rows, function (r) {
        var full = r.cells.length === money.length;
        Array.prototype.forEach.call(r.cells, function (c, k) {
          c.classList.remove('r', 'c');
          c.classList.remove('lft');
          var isTh = !!r.querySelector('th');
          c.classList.add(isTh || !full ? 'c' : (money[k] ? 'r' : (left[k] ? 'lft' : 'c')));
        });
      });
    });
  }
  function route() {
    var startId = 'dashboard'; try { startId = localStorage.getItem('shfc.start') || 'dashboard'; } catch (e) {}   // 설정 > 시작 화면 (주소에 화면 표시가 없을 때만)
    var id = ((location.hash || '#/' + startId).replace('#/', '') || 'dashboard').split('?')[0];   // 첫 화면은 대시보드(기본)
    if (id === 'team' && AC.canTeam && AC.mode === 'self') { AC.setMode('team'); return; }   // 팀원 메뉴는 팀 보기로 열린다
    var m = MENUS.filter(function (x) { return x.id === id; })[0] || MENUS[0];
    nav(m.id);
    var P = window.PAGES && window.PAGES[m.id];
    document.getElementById('main').innerHTML = m.id === 'performance' ? performance() : (P ? P.render() : soon(m.name));
    if (P && P.bind) P.bind();
    document.getElementById('stamp').textContent = '데이터 ' + D.upto + ' 까지';
    document.querySelectorAll('[data-tab]').forEach(function (b) {
      b.onclick = function () { perfTab = b.getAttribute('data-tab'); route(); };
    });
    document.title = '신한FC · ' + m.name;
    userBox();
    document.getElementById('stamp').textContent = '데이터 ' + D.upto + ' 까지';
    if (AC.viewing) {
      document.getElementById('main').insertAdjacentHTML('afterbegin', '<div class="warn viewas">' + AC.user.name + '(' + AC.role + ') 화면으로 미리 보는 중입니다. 이 사람이 접속하면 보게 될 범위만 표시됩니다. <button class="btn" id="asBack">내 화면으로</button></div>');
      var b = document.getElementById('asBack'); if (b) b.onclick = function () { AC.viewAs(''); };
    }
    if (AC.real.role === 'ESL') {   // 임시(시험용): 보기 범위 = 최수진(최수진 로그인 화면) | 팀(팀 전체)
      var TNO = '52028131', tu = AC.users.filter(function (u) { return u.no === TNO; })[0], tn = tu ? tu.name : '최수진';
      document.getElementById('main').insertAdjacentHTML('afterbegin', '<div class="modebar"><span class="mut">보기 범위</span><div class="tabs mini"><button class="' + (AC.user.no === TNO ? 'on' : '') + '" data-asfc="' + TNO + '">' + tn + '</button><button class="' + (AC.user.no === TNO ? '' : 'on') + '" data-asteam="1">팀</button></div></div>');
      document.querySelectorAll('[data-asfc]').forEach(function (b) { b.onclick = function () { if (AC.user.no !== b.getAttribute('data-asfc')) AC.viewAs(b.getAttribute('data-asfc')); }; });
      document.querySelectorAll('[data-asteam]').forEach(function (b) { b.onclick = function () { if (AC.user.no === TNO) { try { localStorage.setItem('shfc.viewAs', ''); localStorage.setItem('shfc.mode', 'team'); } catch (e) {} location.hash = (location.hash || '#/dashboard').replace(/[?&](as|mode)=w+/g, ''); location.reload(); } }; });
    } else if (AC.canTeam) {
      var lab = AC.role === 'ESL' ? '팀 전체 (' + AC.teamSize + '명)' : '팀 (소속 포함 ' + AC.teamSize + '명)';
      document.getElementById('main').insertAdjacentHTML('afterbegin', '<div class="modebar"><span class="mut">보기 범위</span><div class="tabs mini"><button class="' + (AC.mode === 'self' ? 'on' : '') + '" data-mode="self">본인 (' + AC.user.name + ')</button><button class="' + (AC.mode === 'team' ? 'on' : '') + '" data-mode="team">' + lab + '</button></div></div>');
      document.querySelectorAll('[data-mode]').forEach(function (b) { b.onclick = function () { if (b.getAttribute('data-mode') !== AC.mode) AC.setMode(b.getAttribute('data-mode')); }; });
    }
    alignTables();
    if (/dbg=1/.test(location.hash)) { document.body.setAttribute('data-sw', document.documentElement.scrollWidth + '/' + window.innerWidth); }
  }
  window.PAGES = window.PAGES || {};
  window.APP = { rerender: route, align: alignTables };
  window.addEventListener('hashchange', route);
  if (window.DRIVE) DRIVE.onChange(function () { if (!DRIVE.on()) { location.reload(); return; } userBox(); });
  route();
})();
