(function () {
  'use strict';
  // 수수료: 본인 것 / 팀(소속) 것. 수정 없음(조회 전용). 접근 범위는 access.js가 FEES를 이미 걸러 둔다.
  var F = window.FEES, AC = window.ACCESS;
  var sum = function (a) { return a.reduce(function (x, y) { return x + y; }, 0); };
  var won = function (v) { return v == null ? '-' : Math.round(v).toLocaleString('ko-KR'); };
  var man = function (v) { return Math.round(v / 10000).toLocaleString('ko-KR'); };
  var esc = function (s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  var METR = [
    { k: 'net', t: '실 지급금', get: function (r) { return r.tot['실지급금']; } },
    { k: 'gross', t: '지급계', get: function (r) { return r.tot['지급계']; } },
    { k: 'ded', t: '공제계', get: function (r) { return r.tot['공제계']; } },
    { k: 'sub', t: '수수료 소계', get: function (r) { return r.pay['수수료 소계']; } },
    { k: 'rec', t: '모집수수료', get: function (r) { return r.pay['모집수수료']; } }
  ];
  var nk = function (n) { return parseInt(n, 10) || 1e12; };
  var rows = F ? F.rows : [];
  var fcs = []; var seen = {};
  rows.forEach(function (r) { if (!seen[r.no]) { seen[r.no] = 1; fcs.push({ no: r.no, name: r.name }); } });
  fcs.sort(function (a, b) { return nk(a.no) - nk(b.no); });
  var years = []; rows.forEach(function (r) { var y = r.ym.slice(0, 4); if (years.indexOf(y) < 0) years.push(y); }); years.sort();
  var idx = {}; rows.forEach(function (r) { idx[r.no + '|' + r.ym] = r; });
  // 수수료명세서용: 보기 범위와 상관없이 내 팀 전체(access.js가 teamRows로 둠)
  var srows = (F && F.teamRows) || rows;
  var sidx = {}; srows.forEach(function (r) { sidx[r.no + '|' + r.ym] = r; });
  var state = { year: years[years.length - 1], metric: 'net', fc: '', ym: '' };
  var metric = function () { return METR.filter(function (m) { return m.k === state.metric; })[0]; };
  var cell = function (no, ym) { var r = idx[no + '|' + ym]; return r ? (metric().get(r) || 0) : 0; };
  var lastYm = rows.length ? rows[rows.length - 1].ym : '';
  if (!state.ym) state.ym = lastYm;

  // 수수료명세서 FC 선택: 선택한 수수료지급년월에 명세서가 있는 FC만(퇴사한 FC 포함), 그 달 신분 기준 ESL > SSL > GSL > FC, 같은 직급은 사번순
  var RANK = { ESL: 0, SSL: 1, GSL: 2, FC: 3 };
  function roleOf(r) { var t = String((r.info || {})['신분'] || ''); if (/E[ ]*SL/i.test(t)) return 'ESL'; if (/SSL/i.test(t)) return 'SSL'; if (/GSL/i.test(t)) return 'GSL'; return 'FC'; }
  function leftOf(r) { var i = r.info || {}, st = String(i['상태'] || ''); return /해촉|퇴/.test(st) || !!(i['해촉일'] && (!i['재위촉일'] || i['재위촉일'] < i['해촉일'])); }
  function fcOptionsAt(ym, allowed) {
    return srows.filter(function (r) { return r.ym === ym && (!allowed || allowed.indexOf(r.no) >= 0); }).map(function (r) { return { no: r.no, name: r.name, role: roleOf(r), left: leftOf(r) }; })
      .sort(function (a, b) { return RANK[a.role] - RANK[b.role] || nk(a.no) - nk(b.no); });
  }
  function monthsOf(y) { var a = []; for (var m = 1; m <= 12; m++) a.push(y + '-' + (m < 10 ? '0' : '') + m); return a; }
  function seriesOf(no, y) { return monthsOf(y).map(function (ym) { return cell(no, ym); }); }
  function chart(vals, labels) {
    var W = 900, H = 230, L = 64, R = 10, T = 12, B = 32, mx = Math.max.apply(null, [1].concat(vals)) * 1.1, bw = (W - L - R) / vals.length;
    var s = '<svg viewBox="0 0 ' + W + ' ' + H + '" width="100%">';
    for (var g = 0; g <= 4; g++) { var y = T + (H - T - B) * (1 - g / 4); s += '<line x1="' + L + '" x2="' + (W - R) + '" y1="' + y + '" y2="' + y + '" stroke="#E1E4EC"/><text x="' + (L - 8) + '" y="' + (y + 4) + '" text-anchor="end" font-size="12" fill="#5A667E">' + man(mx * g / 4) + '만</text>'; }
    vals.forEach(function (v, i) {
      var x = L + i * bw + bw * 0.15, w = bw * 0.7, h = (H - T - B) * v / mx;
      if (v) s += '<rect x="' + x + '" y="' + (H - B - h) + '" width="' + w + '" height="' + h + '" rx="2" fill="#2B5FD9"><title>' + labels[i] + ': ' + won(v) + '원</title></rect>';
      s += '<text x="' + (x + w / 2) + '" y="' + (H - B + 16) + '" text-anchor="middle" font-size="12" fill="#5A667E">' + labels[i] + '</text>';
    });
    return s + '</svg>';
  }

  // SOL 수수료명세서 화면과 같은 모양: 요약 3줄×3칸 → 수수료지급내역(항목 순서·0 포함, 수수료 소계) → 선지급 항목(선지급액 소계) → 공제내역 → 지급계·공제계·실 지급금
  // 인트라넷(SOL FC) 수수료명세서 화면과 같은 구성(2026-11-11 화면 확인):
  //   기본정보 표(본부·지점·팀·성명·상태 / 위촉일·해촉일·재위촉일·JOB TYPE·신분) → 요약 3줄×3칸 → [수수료지급내역 | 공제내역] 두 칸(항목 순서·0 포함) → 지급계·공제계·실 지급금
  function statement(no, ym) {
    var r = sidx[no + '|' + ym];
    if (!r) return '<p class="mut">이 달의 명세서가 없습니다.</p>';
    var inf = r.info || {};
    // "중앙본부 2110053" 처럼 이름+코드가 한 칸이면 두 줄로 나눈다
    var two = function (s) { var m = String(s || '').match(/^(.*?)\s*(\d{5,})$/); return m ? esc(m[1]) + '<br>' + esc(m[2]) : esc(s || ''); };
    var one = function (k) { return esc(inf[k] || ''); };
    var infoT = '<table class="sol-info"><tr>' +
      '<th>본부</th><td>' + two(inf['본부']) + '</td><th>지점</th><td>' + two(inf['지점']) + '</td><th>팀</th><td>' + two(inf['팀']) + '</td><th>성명</th><td>' + (inf['성명'] ? two(inf['성명']) : esc(r.name) + '<br>' + esc(no)) + '</td><th>상태</th><td>' + one('상태') + '</td></tr><tr>' +
      '<th>위촉일</th><td>' + one('위촉일') + '</td><th>해촉일</th><td>' + one('해촉일') + '</td><th>재위촉일</th><td>' + one('재위촉일') + '</td><th>JOB TYPE</th><td>' + one('JOB TYPE') + '</td><th>신분</th><td>' + one('신분') + '</td></tr></table>';
    var num = function (v) { return v == null || v === '' ? '' : String(v); };
    var isNeg = function (v) { return /^-/.test(String(v)); };
    var sm = r.sum || [], sumRows = '';
    for (var i = 0; i < sm.length; i += 3) {
      sumRows += '<tr>' + sm.slice(i, i + 3).map(function (p) { return '<th>' + esc(p[0]) + '</th><td class="num' + (isNeg(p[1]) ? ' neg' : '') + (/^\d{4}-\d{2}-\d{2}$/.test(String(p[1])) ? ' dt' : '') + '">' + esc(num(p[1])) + '</td>'; }).join('') + '</tr>';
    }
    var rowsOf = function (items) {
      return items.map(function (p) {
        var sub = /소계$/.test(p[0]);
        return '<tr class="' + (sub ? 'sub' : '') + '"><th>' + esc(p[0]) + '</th><td class="num' + (p[1] < 0 ? ' neg' : '') + '">' + won(p[1]) + '</td></tr>';
      }).join('');
    };
    // 공제내역이 지급내역보다 짧으면 SOL 화면처럼 빈 칸(라벨 음영)으로 높이를 맞춘다
    var fill = function (n) { var s = ''; for (var k = 0; k < n; k++) s += '<tr class="fill"><th></th><td></td></tr>'; return s; };
    var t = r.tot || {};
    return (r.chk ? '<div class="warn">이 달 값은 확인이 필요합니다 (' + esc(r.chk) + ')</div>' : '') +
      infoT +
      '<div class="sol-bar"><span class="sol-hint">조회 전용입니다. 금액 항목은 SOL FC 수수료명세서의 항목 순서 그대로이며 0원 항목도 포함합니다.</span><button class="sol-btn" id="feePrint">인쇄</button></div>' +
      '<table class="sol-sum">' + sumRows + '</table>' +
      '<div class="sol-cols"><div><div class="sol-h">수수료지급내역</div><table class="sol-t">' + rowsOf(r.a || []) + rowsOf(r.b || []) + '</table></div>' +
      '<div><div class="sol-h">공제내역</div><table class="sol-t">' + rowsOf(r.d || []) + fill((r.a || []).length + (r.b || []).length - (r.d || []).length) + '</table></div></div>' +
      '<div class="sol-tot"><span>지급계 <b>' + won(t['지급계']) + '</b></span><i>|</i><span>공제계 <b>' + won(t['공제계']) + '</b></span><span class="net">실 지급금 <b>' + won(t['실지급금']) + '</b></span></div>';
  }

  function render() {
    if (!rows.length) return '<h1>수수료</h1><div class="soonbox">표시할 수수료 데이터가 없습니다.</div>';
    var m = metric(), Y = state.year, ms = monthsOf(Y);
    var team = AC.canTeam && AC.mode === 'team';
    var list = team ? fcs : fcs.filter(function (f) { return f.no === AC.user.no; });
    if (!list.length) list = fcs;
    var data = list.map(function (f) { var v = seriesOf(f.no, Y); return { f: f, v: v, total: sum(v), n: v.filter(function (x) { return x !== 0; }).length }; });
    var grand = ms.map(function (_, i) { return sum(data.map(function (d) { return d.v[i]; })); });
    var gTotal = sum(grand), cnt = ms.map(function (_, i) { return data.filter(function (d) { return d.v[i] !== 0; }).length; });
    var avgPerFc = ms.map(function (_, i) { return cnt[i] ? grand[i] / cnt[i] : 0; });
    var monthsPaid = grand.filter(function (x) { return x !== 0; }).length || 1;
    var tabs = years.map(function (y) { return '<button class="' + (Y === y ? 'on' : '') + '" data-fy="' + y + '">' + y + '년</button>'; }).join('');
    var mt = METR.map(function (x) { return '<button class="' + (state.metric === x.k ? 'on' : '') + '" data-fm="' + x.k + '">' + x.t + '</button>'; }).join('');
    var card = function (t, v, d) { return '<div class="k"><div class="l">' + t + '</div><div class="v">' + v + '</div><div class="d">' + d + '</div></div>'; };
    var cards = card(Y + '년 ' + m.t + (team ? ' (팀 합계)' : ''), man(gTotal) + '<small>만원</small>', '지급이 있는 ' + monthsPaid + '개월');
    cards += card('월평균', man(gTotal / monthsPaid) + '<small>만원</small>', team ? '팀 합계 기준' : '본인');
    if (team) {
      var nzMonths = avgPerFc.filter(function (x) { return x !== 0; });
      cards += card('FC 1인당 월평균', man(nzMonths.length ? sum(nzMonths) / nzMonths.length : 0) + '<small>만원</small>', '그 달 지급이 있던 FC 기준');
    } else {
      var lastNz = ms.map(function (ym, i) { return [ym, data[0].v[i]]; }).filter(function (x) { return x[1] !== 0; }).pop();
      cards += card('가장 최근 지급 달', lastNz ? man(lastNz[1]) + '<small>만원</small>' : '-', lastNz ? lastNz[0] : '');
    }
    var best = data.slice().sort(function (a, b) { return b.total - a.total; })[0];
    cards += card(team ? '연 합계 1위' : '월 최대', team ? esc(best.f.name) : man(Math.max.apply(null, data[0].v)) + '<small>만원</small>', team ? man(best.total) + '만원' : '');
    var table = '<div class="wrap"><table class="cmpt"><tr><th>FC</th>' + ms.map(function (_, i) { return '<th>' + (i + 1) + '월</th>'; }).join('') + '<th>연 합계</th><th>월평균</th></tr>' +
      data.map(function (d) { return '<tr><td>' + esc(d.f.name) + '</td>' + d.v.map(function (v) { return '<td>' + (v ? man(v) : '-') + '</td>'; }).join('') + '<td><b>' + man(d.total) + '</b></td><td>' + (d.n ? man(d.total / d.n) : '-') + '</td></tr>'; }).join('') +
      (team ? '<tr class="tot"><td>팀 합계</td>' + grand.map(function (v) { return '<td>' + (v ? man(v) : '-') + '</td>'; }).join('') + '<td>' + man(gTotal) + '</td><td>' + man(gTotal / monthsPaid) + '</td></tr>' +
        '<tr class="tot"><td>FC 1인당 평균</td>' + avgPerFc.map(function (v) { return '<td>' + (v ? man(v) : '-') + '</td>'; }).join('') + '<td></td><td></td></tr>' : '') + '</table></div>';
    var allYms = []; srows.forEach(function (r) { if (allYms.indexOf(r.ym) < 0) allYms.push(r.ym); }); allYms.sort().reverse();
    if (allYms.indexOf(state.ym) < 0) state.ym = allYms[0];
    var opts = fcOptionsAt(state.ym, null);
    if (!opts.length) opts = fcOptionsAt(state.ym, null);
    var stNo = state.fc && opts.some(function (f) { return f.no === state.fc; }) ? state.fc : (opts.some(function (f) { return f.no === AC.user.no; }) ? AC.user.no : (opts[0] ? opts[0].no : list[0].no));
    var fcSel = '<select id="fcSel">' + opts.map(function (f) { return '<option value="' + f.no + '"' + (f.no === stNo ? ' selected' : '') + '>' + esc(f.name) + '</option>'; }).join('') + '</select>';
    var ymSel = '<select id="ymSel">' + allYms.map(function (ym) { return '<option value="' + ym + '"' + (ym === state.ym ? ' selected' : '') + '>' + ym + '</option>'; }).join('') + '</select>';
    var chartVals = team ? grand : data[0].v;
    return '<h1>수수료</h1><p class="sub">수수료명세서 기준 · ' + (team ? '팀 전체(소속 ' + list.length + '명)' : '본인') + ' · 금액은 원 단위(표는 만원) · 조회 전용입니다. 개인 지급 정보라 본인과 상위 관리자만 볼 수 있습니다.</p>' +
      '<div class="tabs">' + tabs + '</div>' +
      '<div class="kpis5">' + cards + '</div>' +
      '<section><h2>월별 ' + m.t + ' (만원) <span class="tabs mini">' + mt + '</span></h2>' + table + '<h3 class="h3" style="margin-top:14px">월별 추이 ' + (team ? '(팀 합계)' : '') + '</h3>' + chart(chartVals, ms.map(function (_, i) { return (i + 1) + '월'; })) + '</section>' +
      '<section class="sol"><div class="sol-title">수수료명세서</div><div class="sol-filter"><div class="r"><label>FC</label>' + fcSel + '</div><div class="r"><label>수수료지급년월</label>' + ymSel + '<button class="sol-btn go" id="feeGo">조회</button></div></div>' + statement(stNo, state.ym) + '</section>';
  }
  function bind() {
    var rr = function () { window.APP.rerender(); };
    document.querySelectorAll('[data-fy]').forEach(function (b) { b.onclick = function () { state.year = b.getAttribute('data-fy'); rr(); }; });
    document.querySelectorAll('[data-fm]').forEach(function (b) { b.onclick = function () { state.metric = b.getAttribute('data-fm'); rr(); }; });
    var f = document.getElementById('fcSel'); if (f) f.onchange = function () { state.fc = f.value; rr(); };
    var y = document.getElementById('ymSel'); if (y) y.onchange = function () { state.ym = y.value; rr(); };
    var go = document.getElementById('feeGo'); if (go) go.onclick = function () { rr(); };
    var pr = document.getElementById('feePrint'); if (pr) pr.onclick = function () { window.print(); };
  }
  window.PAGES = window.PAGES || {};
  window.PAGES.fee = { render: render, bind: bind };
})();
