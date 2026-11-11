(function () {
  'use strict';
  var C = window.CONTRACTS;
  var sum = function (a) { return a.reduce(function (x, y) { return x + y; }, 0); };
  var fmtN = function (v) { return v.toLocaleString('ko-KR', { minimumFractionDigits: v % 1 ? 1 : 0, maximumFractionDigits: 1 }); };
  var fmtW = function (v) { return Math.round(v).toLocaleString('ko-KR'); };
  var PAGE = 1e9;   // 쪽 나눔 없이 전체 표시
  var state = { q: '', fc: '', st: '', from: '', to: '', only: true, page: 0 };
  var FCS = []; var seen = {};
  C.rows.forEach(function (r) { if (!seen[r.no]) { seen[r.no] = 1; FCS.push({ no: r.no, name: r.fc }); } });
  var CURNOS = (window.PERF && window.PERF.current) || [];
  ((window.PERF && window.PERF.currentFcs) || []).forEach(function (f) { if (!FCS.some(function (x) { return x.no === f.no; })) FCS.push({ no: f.no, name: f.name }); });
  var isCurFc = function (no) { return CURNOS.indexOf(no) >= 0; };
  var fcLabel = function (no, name) { return isCurFc(no) ? name : '(' + name + ')'; };
  // 현재 재직 FC 먼저(사번순), 퇴직 FC는 뒤에 괄호로(사번순)
  FCS.sort(function (a, b) { return window.ACCESS.fcCompare(a, b); });   // 직급 먼저, 같은 직급은 사번순
  var STS = []; var s2 = {}; C.rows.forEach(function (r) { if (!s2[r.st]) { s2[r.st] = 1; STS.push(r.st); } });
  var YMS = []; var s3 = {}; C.rows.forEach(function (r) { if (!s3[r.ym]) { s3[r.ym] = 1; YMS.push(r.ym); } }); YMS.sort();
  var esc = function (s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };

  function filtered() {
    var q = state.q.trim().toLowerCase();
    return C.rows.filter(function (r) {
      if (state.only && r.st !== '성립') return false;
      if (state.fc && r.no !== state.fc) return false;
      if (state.st && r.st !== state.st) return false;
      if (state.from && r.ym < state.from) return false;
      if (state.to && r.ym > state.to) return false;
      if (q && (r.owner + ' ' + r.insured + ' ' + r.goods + ' ' + r.ppnu + ' ' + r.fc).toLowerCase().indexOf(q) < 0) return false;
      return true;
    });
  }
  function stClass(st) { return st === '성립' ? 't3' : (st.indexOf('철회') >= 0 || st.indexOf('거절') >= 0 ? 't1' : 't2'); }
  function render() {
    var rows = filtered();
    var ok = rows.filter(function (r) { return r.st === '성립'; });
    var n = sum(ok.map(function (r) { return r.ctrt / 100; })), cm = sum(ok.map(function (r) { return r.cmip * r.ctrt / 100; })), ca = sum(ok.map(function (r) { return r.canp * r.ctrt / 100; }));
    var pages = Math.max(1, Math.ceil(rows.length / PAGE)); if (state.page >= pages) state.page = pages - 1;
    var view = rows.slice(state.page * PAGE, state.page * PAGE + PAGE);
    var opt = function (arr, cur, lab) { return arr.map(function (x) { var v = x.v !== undefined ? x.v : x, t = x.t !== undefined ? x.t : x; return '<option value="' + esc(v) + '"' + (String(cur) === String(v) ? ' selected' : '') + '>' + esc(t) + '</option>'; }).join(''); };
    var body = view.map(function (r) {
      return '<tr class="' + (r.st === '성립' ? '' : 'gray') + '"><td>' + esc(r.ymd) + '</td><td>' + esc(fcLabel(r.no, r.fc)) + '</td><td><span class="tag ' + stClass(r.st) + '">' + esc(r.st) + '</span></td><td class="l">' + esc(r.goods) + '</td><td>' + esc(r.owner) + '</td><td>' + esc(r.insured) + '</td><td>' + fmtW(r.sum) + '</td><td>' + fmtW(r.cmip) + '</td><td>' + fmtW(r.canp) + '</td><td>' + r.ctrt + '%</td></tr>';
    }).join('');
    return '<h1>계약</h1><p class="sub">청약계약List조회 원본 · ' + C.rows.length + '건 (' + YMS[0] + ' ~ ' + YMS[YMS.length - 1] + ') · 금액은 SOL 화면 값(기여율 미반영), 합계 카드는 기여율 반영</p>' +
      '<div class="filters"><input id="q" type="search" placeholder="계약자·피보험자·상품명·청약번호 검색" value="' + esc(state.q) + '">' +
      '<select id="fc"><option value="">팀</option>' + opt(FCS.map(function (f) { return { v: f.no, t: f.name }; }), state.fc) + '</select>' +
      '<select id="st"><option value="">청약상태 전체</option>' + opt(STS, state.st) + '</select>' +
      '<select id="from"><option value="">시작월</option>' + opt(YMS, state.from) + '</select><select id="to"><option value="">끝월</option>' + opt(YMS, state.to) + '</select>' +
      '<label class="chk"><input id="only" type="checkbox"' + (state.only ? ' checked' : '') + '> 성립만</label><button id="csv" class="btn">CSV 내보내기</button></div>' +
      '<div id="ctBody"><div class="kpis4"><div class="k"><div class="l">조회된 계약</div><div class="v">' + rows.length.toLocaleString('ko-KR') + '<small>건</small></div></div>' +
      '<div class="k"><div class="l">성립 건수 (기여율 반영)</div><div class="v">' + fmtN(n) + '<small>건</small></div></div>' +
      '<div class="k"><div class="l">CMIP 합계 (원, 기여율 반영)</div><div class="v">' + fmtW(cm) + '</div></div>' +
      '<div class="k"><div class="l">CANP 합계 (원, 기여율 반영)</div><div class="v">' + fmtW(ca) + '</div></div></div>' +
      '<section><div class="wrap"><table class="ct hc"><tr><th>계약일</th><th>FC</th><th>청약상태</th><th class="l">상품명</th><th>계약자</th><th>피보험자</th><th>보험료</th><th>월납화보험료</th><th>CANP</th><th>기여율</th></tr>' + (body || '<tr><td colspan="10" class="mutc">조건에 맞는 계약이 없습니다.</td></tr>') + '</table></div>' +
      '<div class="pager"><span>전체 ' + rows.length.toLocaleString('ko-KR') + '건 표시</span></div></section></div>';
  }
  function partial() {
    var tmp = document.createElement('div'); tmp.innerHTML = render();
    var a = document.getElementById('ctBody'), b = tmp.querySelector('#ctBody'); if (a && b) a.innerHTML = b.innerHTML;
    if (window.APP && APP.align) APP.align();
  }
  function bind() {
    var g = function (id) { return document.getElementById(id); };
    var rr = function () { window.APP.rerender(); };
    var q = g('q'); if (q) { q.oninput = function () { state.q = q.value; state.page = 0; partial(); }; }   // 입력칸은 그대로 두고 아래만 바꿔 한글 조합이 깨지지 않게
    ['fc', 'st', 'from', 'to'].forEach(function (id) { var el = g(id); el.onchange = function () { state[id] = el.value; state.page = 0; rr(); }; });
    g('only').onchange = function () { state.only = g('only').checked; state.page = 0; rr(); };
    
    
    g('csv').onclick = function () {
      var rows = filtered();
      var head = ['계약일', 'FC', '청약상태', '상품명', '계약자', '피보험자', '보험료', '월납화보험료', 'CANP', '기여율(%)'];
      var csv = [head].concat(rows.map(function (r) { return [r.ymd, r.fc, r.st, r.goods, r.owner, r.insured, r.sum, r.cmip, r.canp, r.ctrt]; })).map(function (a) { return a.map(function (x) { return '"' + String(x == null ? '' : x).replace(/"/g, '""') + '"'; }).join(','); }).join('\r\n');
      var blob = new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8' });
      var a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = '계약목록.csv'; a.click();
    };
  }
  window.PAGES = window.PAGES || {};
  window.PAGES.contract = { render: render, bind: bind };
})();
