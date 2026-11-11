(function () {
  'use strict';
  var D = window.PERF, C = window.CONTRACTS;
  var sum = function (a) { return a.reduce(function (x, y) { return x + y; }, 0); };
  var fmtN = function (v) { return v.toLocaleString('ko-KR', { minimumFractionDigits: v % 1 ? 1 : 0, maximumFractionDigits: 1 }); };
  var fmtW = function (v) { return Math.round(v).toLocaleString('ko-KR'); };
  var man = function (v) { return Math.round(v / 10000).toLocaleString('ko-KR'); };
  var esc = function (s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  var METR = [
    { k: 'n', t: '건수', f: fmtN, ax: function (v) { return Math.round(v * 10) / 10 + '건'; } },
    { k: 'cmip', t: 'CMIP', f: fmtW, ax: function (v) { return man(v) + '만'; } },
    { k: 'canp', t: 'CANP', f: fmtW, ax: function (v) { return man(v) + '만'; } }
  ];
  // 명단 (사번은 SOL 기준). 교육생·후보자는 아직 사번이 없다.
  var ROSTER = [
    { no: '00026637', name: '김현성', role: 'ESL' },
    { no: '00044090', name: '임지섭', role: 'FC' },
    { no: '52028131', name: '최수진', role: 'FC' },
    { no: '52038578', name: '박상선', role: 'FC' },
    { no: '52055143', name: '김경미', role: 'FC' },
    { no: '52087261', name: '오상만', role: 'FC' },
    { no: '', name: '이현미', role: '교육생' },
    { no: '', name: '김주형', role: '후보자' }
  ];
  // 접근 범위: ESL은 전체(교육생·후보자·퇴직 FC 참조 포함), GSL·SSL은 소속, FC는 본인만(팀원 메뉴 없음)
  var AC = window.ACCESS;
  ROSTER = ROSTER.filter(function (r) { return AC.dataAll || (r.no && AC.inScope(r.no)); });
  var MONTHS = [];
  Object.keys(D.years).sort().forEach(function (y) { for (var m = 1; m <= 12; m++) { var ym = y + '-' + (m < 10 ? '0' : '') + m; if (ym >= D.from && ym <= D.upto) MONTHS.push(ym); } });
  var known = {}; ROSTER.forEach(function (r) { known[r.no] = 1; });
  var people = ROSTER.slice();
  // 현재 재직자(명단) 뒤에 퇴직 FC를 괄호 한 줄(참조)로 붙인다
  var RET = D.fcs.filter(function (f) { return !known[f.no]; }).sort(function (a, b) { return (parseInt(a.no, 10) || 1e12) - (parseInt(b.no, 10) || 1e12); });
  var RETNOS = RET.map(function (f) { return f.no; });
  if (RET.length && AC.dataAll) people.push({ no: 'RET', name: '(' + RET.map(function (f) { return f.name; }).join(', ') + ')', role: '퇴직 FC', group: true });
  var state = { sel: '52028131', metric: 'n', others: false, span: 'ytd' };

  function mval(no, ym, k) { if (no === 'RET') return sum(RETNOS.map(function (n) { return mval(n, ym, k); })); var a = D.years[ym.slice(0, 4)]; return a && a[no] && a[no].n && a[no][k] ? a[no][k][parseInt(ym.slice(5), 10) - 1] : 0; }
  function series(no, k) { return MONTHS.map(function (ym) { return mval(no, ym, k); }); }
  function stats(p) {
    var n = series(p.no, 'n'), cm = series(p.no, 'cmip'), ca = series(p.no, 'canp');
    var first = '', last = '';
    MONTHS.forEach(function (ym, i) { if (n[i] > 0) { if (!first) first = ym; last = ym; } });
    var l12 = n.slice(-12), l3 = n.slice(-3);
    var ytdMonths = parseInt(D.upto.slice(5), 10), cy = D.upto.slice(0, 4);
    var ytd = p.no === 'RET' ? sum(RETNOS.map(function (n) { return D.years[cy] && D.years[cy][n] ? sum(D.years[cy][n].n.slice(0, ytdMonths)) : 0; })) : (D.years[cy] && D.years[cy][p.no] ? sum(D.years[cy][p.no].n.slice(0, ytdMonths)) : 0);
    return { n: sum(n), cmip: sum(cm), canp: sum(ca), first: first, last: last, l12: sum(l12), l12c: sum(cm.slice(-12)), l12p: sum(ca.slice(-12)), active: sum(l3) > 0, ytd: ytd };
  }
  // 기간 기준 합계: 올해(1월~데이터 마지막 달) 또는 최근 12개월
  function spanStats(p, span) {
    var cy = D.upto.slice(0, 4), idx = [];
    MONTHS.forEach(function (ym, i) { if (span === 'ytd' ? ym.slice(0, 4) === cy : i >= MONTHS.length - 12) idx.push(i); });
    var pick = function (k) { var a = series(p.no, k); return sum(idx.map(function (i) { return a[i]; })); };
    var m1 = idx.length ? MONTHS[idx[0]] : '', m2 = idx.length ? MONTHS[idx[idx.length - 1]] : '';
    return { n: pick('n'), cmip: pick('cmip'), canp: pick('canp'), label: span === 'ytd' ? cy + '년 (1~' + parseInt(D.upto.slice(5), 10) + '월)' : '최근 12개월 (' + m1.slice(2).replace('-', '.') + '~' + m2.slice(2).replace('-', '.') + ')' };
  }
  function chart(p, m) {
    var vals = series(p.no, m.k), W = 900, H = 260, L = 64, R = 10, T = 14, B = 40;
    var mx = Math.max.apply(null, [1].concat(vals)) * 1.1, bw = (W - L - R) / vals.length;
    var s = '<svg viewBox="0 0 ' + W + ' ' + H + '" width="100%">';
    for (var g = 0; g <= 4; g++) { var y = T + (H - T - B) * (1 - g / 4); s += '<line x1="' + L + '" x2="' + (W - R) + '" y1="' + y + '" y2="' + y + '" stroke="#E1E4EC"/><text x="' + (L - 8) + '" y="' + (y + 4) + '" text-anchor="end" font-size="12" fill="#5A667E">' + m.ax(mx * g / 4) + '</text>'; }
    vals.forEach(function (v, i) {
      var x = L + i * bw + bw * 0.15, w = bw * 0.7, h = (H - T - B) * v / mx;
      if (v) s += '<rect x="' + x + '" y="' + (H - B - h) + '" width="' + w + '" height="' + h + '" rx="2" fill="#2B5FD9"><title>' + MONTHS[i] + ': ' + m.f(v) + '</title></rect>';
      if (i % 3 === 0) s += '<text x="' + (x + w / 2) + '" y="' + (H - B + 16) + '" text-anchor="middle" font-size="11" fill="#5A667E">' + MONTHS[i].slice(2).replace('-', '.') + '</text>';
    });
    return s + '</svg>';
  }
  function detail(p) {
    var st = stats(p), sp = spanStats(p, state.span), m = METR.filter(function (x) { return x.k === state.metric; })[0];
    var tabs = METR.map(function (x) { return '<button class="' + (state.metric === x.k ? 'on' : '') + '" data-metric="' + x.k + '">' + x.t + '</button>'; }).join('');
    var yrs = Object.keys(D.years).sort().map(function (y) {
      var a = p.no === 'RET' ? (function () { var g = { n: [], cmip: [], canp: [] }; RETNOS.forEach(function (n) { var z = D.years[y][n]; if (!z) return; ['n', 'cmip', 'canp'].forEach(function (k) { z[k].forEach(function (v, i) { g[k][i] = (g[k][i] || 0) + v; }); }); }); return g.n.length ? g : null; })() : D.years[y][p.no]; var part = y === D.upto.slice(0, 4) ? ' (1~' + parseInt(D.upto.slice(5), 10) + '월)' : '';
      return '<tr><td>' + y + '년' + part + '</td><td>' + (a ? fmtN(sum(a.n)) : '-') + '</td><td>' + (a ? fmtW(sum(a.cmip)) : '-') + '</td><td>' + (a ? fmtW(sum(a.canp)) : '-') + '</td></tr>';
    }).join('');
    var recent = C.rows.filter(function (r) { return (p.no === 'RET' ? RETNOS.indexOf(r.no) >= 0 : r.no === p.no) && r.st === '성립'; }).slice(0, 8).map(function (r) {
      return '<tr><td>' + esc(r.ymd) + '</td><td class="l">' + esc(r.goods) + '</td><td>' + r.ctrt + '%</td></tr>';
    }).join('');
    var nodata = !p.no || !st.n;
    return '<section class="detail"><h2>' + esc(p.name) + ' <small class="mut">' + esc(p.role) + (p.no ? ' · ' + p.no : '') + '</small></h2>' +
      (nodata ? '<p class="mut">이 기간(' + D.from + ' ~ ' + D.upto + ')에 성립 계약 실적이 없습니다.' + (p.no ? '' : ' 사번이 등록되면 실적이 연결됩니다.') + '</p>' :
        '<div class="spanbar"><span class="mut">집계 기준</span><div class="tabs mini"><button class="' + (state.span === 'ytd' ? 'on' : '') + '" data-span="ytd">올해</button><button class="' + (state.span === 'l12' ? 'on' : '') + '" data-span="l12">최근 12개월</button></div><span class="mut">' + sp.label + '</span></div><div class="kpis4"><div class="k"><div class="l">건수 (기여율 반영)</div><div class="v">' + fmtN(sp.n) + '<small>건</small></div></div><div class="k"><div class="l">CMIP</div><div class="v">' + man(sp.cmip) + '<small>만원</small></div></div><div class="k"><div class="l">CANP</div><div class="v">' + man(sp.canp) + '<small>만원</small></div></div><div class="k"><div class="l">첫 성립 ~ 최근 성립</div><div class="v" style="font-size:22px">' + st.first + ' ~ ' + st.last + '</div></div></div>' +
        '<h3>월별 추이 <span class="tabs mini">' + tabs + '</span></h3>' + chart(p, m) +
        '<div class="two2"><div><h3>연도별</h3><div class="wrap"><table><tr><th>연도</th><th>건수</th><th>CMIP(원)</th><th>CANP(원)</th></tr>' + yrs + '</table></div></div>' +
        '<div><h3>최근 성립 계약</h3><div class="wrap"><table class="ct"><tr><th>계약일</th><th class="l">상품명</th><th>기여율</th></tr>' + recent + '</table></div></div></div>') + '</section>';
  }
  function render() {
    var list = people;
    var rows = list.map(function (p) {
      var st = p.no ? stats(p) : null;
      var badge = p.role === 'ESL' ? 't3' : p.role === 'FC' ? 't3' : p.role === '퇴직 FC' ? 'tg' : 't2';
      var status = !st ? '<span class="mut">-</span>' : st.active ? '<span class="up">활동</span>' : '<span class="mut">최근 3개월 실적 없음</span>';
      return '<tr class="pick-row ' + (state.sel === (p.no || p.name) ? 'selrow' : '') + '" data-sel="' + esc(p.no || p.name) + '"><td class="l' + (p.group ? ' wrapc' : '') + '"><b>' + esc(p.name) + '</b></td><td><span class="tag ' + badge + '">' + esc(p.role) + '</span></td><td>' + (p.group ? '-' : esc(p.no || '-')) + '</td><td>' + (st ? fmtN(st.l12) : '-') + '</td><td>' + (st ? fmtW(st.l12c) : '-') + '</td><td>' + (st ? fmtW(st.l12p) : '-') + '</td><td>' + (st ? fmtN(st.ytd) : '-') + '</td></tr>';
    }).join('');
    var sel = people.filter(function (p) { return (p.no || p.name) === state.sel; })[0] || people[2];
    return '<div class="topbar"><div><h1>팀원</h1><p class="sub">현재 팀원 ' + ROSTER.length + '명 · 사번순 · 실적은 성립 계약, 기여율 반영 (' + D.from + ' ~ ' + D.upto + ')</p></div>' +
      '</div>' +
      '<section><div class="wrap"><table class="ct tm"><tr><th class="l">이름</th><th>구분</th><th>사번</th><th>최근 12개월<br>건수</th><th>최근 12개월<br>CMIP(원)</th><th>최근 12개월<br>CANP(원)</th><th>올해 누적<br>건수</th></tr>' + rows + '</table></div><p class="note">행을 누르면 아래에 상세가 열립니다. 면담·교육·자격 정보는 이후 데이터를 연결하면 이 화면에 추가합니다.</p></section>' + detail(sel);
  }
  function bind() {
    var rr = function () { window.APP.rerender(); };
    document.querySelectorAll('[data-sel]').forEach(function (tr) { tr.onclick = function () { state.sel = tr.getAttribute('data-sel'); rr(); window.scrollTo(0, document.body.scrollHeight); }; });
    document.querySelectorAll('[data-span]').forEach(function (b) { b.onclick = function () { state.span = b.getAttribute('data-span'); rr(); }; });
    document.querySelectorAll('[data-metric]').forEach(function (b) { b.onclick = function () { state.metric = b.getAttribute('data-metric'); rr(); }; });
  }
  window.PAGES = window.PAGES || {};
  window.PAGES.team = { render: render, bind: bind };
})();
