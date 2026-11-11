(function () {
  'use strict';
  var D = window.PERF, C = window.CONTRACTS;
  // 정렬 규칙: 현재 근무 FC를 사번순, 그 뒤에 퇴직 FC를 괄호로 사번순
  var noKey = function (n) { return parseInt(n, 10) || 1e12; };
  var isCurNo = function (n) { return D.current.indexOf(n) >= 0; };
  var CURNAMES = {}; (D.currentFcs || []).forEach(function (f) { CURNAMES[f.no] = f.name; });
  var ALLFC = D.fcs.slice();
  D.current.forEach(function (n) { if (!ALLFC.some(function (f) { return f.no === n; })) ALLFC.push({ no: n, name: CURNAMES[n] || n }); });
  var SORTED = ALLFC.sort(function (a, b) {
    var ca = isCurNo(a.no), cb = isCurNo(b.no);
    return ca !== cb ? (ca ? -1 : 1) : noKey(a.no) - noKey(b.no);
  });
  var RET = SORTED.filter(function (f) { return !isCurNo(f.no); });
  var RETNOS = RET.map(function (f) { return f.no; });
  // 색은 FC마다 원래대로(사번순 목록 기준 위치), 퇴직 묶음만 회색
  var ci = function (no) { var k = D.fcs.map(function (f) { return f.no; }).indexOf(no); return k >= 0 ? k : D.fcs.length; };
  var FCS = SORTED.filter(function (f) { return isCurNo(f.no); }).map(function (f) { return { no: f.no, name: f.name, ci: ci(f.no) }; });
  if (RET.length) FCS.push({ no: 'RET', name: '(' + RET.map(function (f) { return f.name; }).join(', ') + ')', group: true, ci: -1 });
  var COLORS = ['#2B5FD9', '#1E7F74', '#B5651D', '#7C4DBB', '#C0392B', '#0E8FB3', '#8A9A1B', '#D6336C', '#5B6B91', '#E0A800', '#3C8D40', '#9B59B6', '#16A085', '#E67E22', '#34495E', '#C2185B', '#7F8C8D', '#2E86C1', '#A04000', '#117A65'];
  var col = function (i) { return i < 0 ? '#C3CAD8' : COLORS[i % COLORS.length]; };
  var sum = function (a) { return a.reduce(function (x, y) { return x + y; }, 0); };
  var fmtN = function (v) { return v.toLocaleString('ko-KR', { minimumFractionDigits: v % 1 ? 1 : 0, maximumFractionDigits: 1 }); };
  var fmtW = function (v) { return Math.round(v).toLocaleString('ko-KR'); };
  var man = function (v) { return Math.round(v / 10000).toLocaleString('ko-KR'); };
  var METR = [
    { k: 'n', t: '건수', u: '건', f: fmtN, ax: function (v) { return (v < 10 ? Math.round(v * 10) / 10 : Math.round(v)) + '건'; } },
    { k: 'cmip', t: 'CMIP', u: '원', f: fmtW, ax: function (v) { return man(v) + '만'; } },
    { k: 'canp', t: 'CANP', u: '원', f: fmtW, ax: function (v) { return man(v) + '만'; } }
  ];
  // all months in order
  var MONTHS = [];
  Object.keys(D.years).sort().forEach(function (y) { for (var m = 1; m <= 12; m++) { var ym = y + '-' + (m < 10 ? '0' : '') + m; if (ym >= D.from && ym <= D.upto) MONTHS.push(ym); } });
  var state = { base: D.upto, metric: 'n' };
  function val(no, ym, k) { if (no === 'RET') return sum(RETNOS.map(function (n) { return val(n, ym, k); })); var a = D.years[ym.slice(0, 4)]; return a && a[no] ? a[no][k][parseInt(ym.slice(5), 10) - 1] : 0; }
  function tot(ym, k) { return sum(FCS.map(function (f) { return val(f.no, ym, k); })); }
  function prevYm(ym, back) { var y = parseInt(ym.slice(0, 4), 10), m = parseInt(ym.slice(5), 10) - back; while (m < 1) { m += 12; y--; } return y + '-' + (m < 10 ? '0' : '') + m; }
  function diffHtml(cur, prev, f) {
    var d = cur - prev; if (!prev && !cur) return '<span class="mut">변화 없음</span>';
    var pct = prev ? ' (' + (d >= 0 ? '+' : '') + (d / prev * 100).toFixed(0) + '%)' : '';
    return '<span class="' + (d >= 0 ? 'up' : 'dn') + '">' + (d >= 0 ? '▲ ' : '▼ ') + f(Math.abs(d)) + pct + '</span>';
  }
  function ytd(no, y, upToMonth, k) { if (no === 'RET') return sum(RETNOS.map(function (n) { return ytd(n, y, upToMonth, k); })); var a = D.years[y]; if (!a || !a[no]) return 0; return sum(a[no][k].slice(0, upToMonth)); }

  function chart(metric) {
    var m = METR.filter(function (x) { return x.k === metric; })[0];
    var list = []; for (var i = 11; i >= 0; i--) list.push(prevYm(state.base, i));
    list = list.filter(function (ym) { return ym >= D.from; });
    var W = 820, H = 300, L = 64, R = 10, T = 14, B = 36;
    var totals = list.map(function (ym) { return tot(ym, metric); });
    var mx = Math.max.apply(null, [1].concat(totals)) * 1.08, bw = (W - L - R) / list.length;
    var s = '<svg viewBox="0 0 ' + W + ' ' + H + '" width="100%">';
    for (var g = 0; g <= 4; g++) { var y = T + (H - T - B) * (1 - g / 4); s += '<line x1="' + L + '" x2="' + (W - R) + '" y1="' + y + '" y2="' + y + '" stroke="#E1E4EC"/><text x="' + (L - 8) + '" y="' + (y + 4) + '" text-anchor="end" font-size="12" fill="#5A667E">' + m.ax(mx * g / 4) + '</text>'; }
    list.forEach(function (ym, i) {
      var y0 = H - B, x = L + i * bw + bw * 0.15, w = bw * 0.7;
      FCS.forEach(function (f, j) { var v = val(f.no, ym, metric); if (!v) return; var h = (H - T - B) * v / mx; y0 -= h; s += '<rect x="' + x + '" y="' + y0 + '" width="' + w + '" height="' + h + '" fill="' + col(FCS[j].ci) + '"><title>' + ym + ' ' + f.name + ': ' + m.f(v) + m.u + '</title></rect>'; });
      s += '<text x="' + (x + w / 2) + '" y="' + (H - B + 18) + '" text-anchor="middle" font-size="12" fill="#5A667E">' + ym.slice(2).replace('-', '.') + '</text>';
    });
    return s + '</svg>';
  }


  function todo() {
    var K = window.CUSTOMERS; if (!K) return '<p class="mut">고객 데이터가 없습니다.</p>';
    var today = realToday(), ym = today.slice(0, 7);
    var plus = function (n) { var t = new Date(today + 'T00:00:00'); t.setDate(t.getDate() + n); return t.getFullYear() + '-' + p2(t.getMonth() + 1) + '-' + p2(t.getDate()); };
    var in30 = plus(30);
    var h = [].concat.apply([], Object.keys(K.haeng).map(function (k) { return K.haeng[k]; })).filter(function (x) { return String(x.ym).slice(0, 7) === ym; });
    var tg = h.filter(function (x) { return x.target !== '아니오' && x.target !== 'N'; });
    var done = tg.filter(function (x) { return x.done; }).length;
    var exp = K.customers.filter(function (c) { var d = c['가입설계동의 만료일']; return d && d >= today && d <= in30; }).length;
    var expired = K.customers.filter(function (c) { var d = c['가입설계동의 만료일']; return d && d < today && c['현재단계'] !== '피보험자'; }).length;
    var cons = [].concat.apply([], Object.keys(K.contracts).map(function (k) { return K.contracts[k]; }));
    var lapsed = cons.filter(function (x) { return x.st === '실효'; }).length;
    var renew = cons.filter(function (x) { return x.renew && x.renew >= today && x.renew <= in30; }).length;
    var line = function (t, v, c) { return '<div class="kv"><span>' + t + '</span><b class="' + (c || '') + '">' + v + '</b></div>'; };
    return line(ym + ' 따행 완료', done + ' / ' + tg.length + '명') + line('가입설계동의 만료 30일 내', exp + '명', exp ? 'dn' : '') + line('가입설계동의 이미 만료', expired + '명') + line('실효 계약', lapsed + '건', lapsed ? 'dn' : '') + line('갱신 도래 30일 내', renew + '건') +
      '<p class="note">수집이 끝난 FC(' + Object.keys(K.files).join('·') + ') 기준이고, 나머지 FC는 고객 데이터를 모은 뒤 포함됩니다. 따행 대상은 캠페인대상여부가 아닌 건을 뺀 수입니다.</p>';
  }

  // ===== 상단: 이번달 / 이전달 / 최근 3개월 / 최근 1년 계약 통계 =====
  var PERIODS = [{ k: 'this', t: '이번달' }, { k: 'prev', t: '지난달' }];
  state.period = ((location.hash.match(/period=(this|prev)/) || [])[1]) || 'this'; state.pmetric = 'n';
  var p2 = function (n) { return n < 10 ? '0' + n : '' + n; };
  function realToday() { var d = new Date(); return d.getFullYear() + '-' + p2(d.getMonth() + 1) + '-' + p2(d.getDate()); }
  var TESTYM = (function () {
    var y = realToday().slice(0, 4), cnt = {};
    C.rows.forEach(function (r) { if (r.ym.slice(0, 4) === y) cnt[r.ym] = (cnt[r.ym] || 0) + 1; });
    var best = '', mx = 0; Object.keys(cnt).forEach(function (k) { if (cnt[k] > mx) { mx = cnt[k]; best = k; } });
    return best || realToday().slice(0, 7);
  })();
  state.test = true;
  function localToday() { return state.test ? TESTYM + '-' + p2(lastDay(TESTYM)) : realToday(); }
  function lastDay(ym) { return new Date(+ym.slice(0, 4), +ym.slice(5), 0).getDate(); }
  function mAdd(ym, n) { var y = +ym.slice(0, 4), m = +ym.slice(5) + n; while (m < 1) { m += 12; y--; } while (m > 12) { m -= 12; y++; } return y + '-' + p2(m); }
  function range(k) {
    var td = localToday(), cm = td.slice(0, 7), day = +td.slice(8), pm = mAdd(cm, -1);
    if (k === 'this') { var cd = pm + '-' + p2(Math.min(day, lastDay(pm))); return { from: cm + '-01', to: td, cf: pm + '-01', ct: cd, unit: 'day', cmp: '지난달 같은 기간 (' + pm + '-01 ~ ' + cd + ')' }; }
    if (k === 'prev') { var pp = mAdd(cm, -2); return { from: pm + '-01', to: pm + '-' + lastDay(pm), cf: pp + '-01', ct: pp + '-' + lastDay(pp), unit: 'day', cmp: '그 전달 (' + pp + ')' }; }
    if (k === 'q3') { var a = mAdd(cm, -3), b = mAdd(cm, -6); return { from: a + '-01', to: pm + '-' + lastDay(pm), cf: b + '-01', ct: mAdd(cm, -4) + '-' + lastDay(mAdd(cm, -4)), unit: 'month', cmp: '직전 3개월 (' + b + ' ~ ' + mAdd(cm, -4) + ')' }; }
    var a1 = mAdd(cm, -12), b1 = mAdd(cm, -24);
    return { from: a1 + '-01', to: pm + '-' + lastDay(pm), cf: b1 + '-01', ct: mAdd(cm, -13) + '-' + lastDay(mAdd(cm, -13)), unit: 'month', cmp: '직전 1년 (' + b1 + ' ~ ' + mAdd(cm, -13) + ')' };
  }
  var fcKey = function (no) { return isCurNo(no) ? no : 'RET'; };
  function category(g) {
    g = g || '';
    if (/변액/.test(g)) return '변액'; if (/연금/.test(g)) return '연금'; if (/종신/.test(g)) return '종신';
    if (/건강|암|질병|케어|간병|치아|실속|통합|위너스|치매/.test(g)) return '건강·보장'; return '기타';
  }
  function agg(from, to) {
    var rows = C.rows.filter(function (r) { return r.ymd >= from && r.ymd <= to; });
    var ok = rows.filter(function (r) { return r.st === '성립'; });
    var o = { rows: rows, ok: ok, total: rows.length, byFc: {}, n: 0, cmip: 0, canp: 0, bad: 0, mix: {} };
    ok.forEach(function (r) {
      var w = r.ctrt / 100, k = fcKey(r.no), b = o.byFc[k] || (o.byFc[k] = { n: 0, cmip: 0, canp: 0 });
      b.n += w; b.cmip += r.cmip * w; b.canp += r.canp * w; o.n += w; o.cmip += r.cmip * w; o.canp += r.canp * w;
      var c = category(r.goods); o.mix[c] = (o.mix[c] || 0) + w;
    });
    rows.forEach(function (r) { if (r.st.indexOf('철회') >= 0 || r.st.indexOf('거절') >= 0) o.bad++; });
    return o;
  }
  function stackedPeriod(R, metric) {
    var m = METR.filter(function (x) { return x.k === metric; })[0];
    var keys = [], rows = C.rows.filter(function (r) { return r.st === '성립' && r.ymd >= R.from && r.ymd <= R.to; });
    if (R.unit === 'day') { var last = +R.to.slice(8); for (var d = 1; d <= last; d++) keys.push(R.from.slice(0, 8) + p2(d)); }
    else { var cur = R.from.slice(0, 7); while (cur <= R.to.slice(0, 7)) { keys.push(cur); cur = mAdd(cur, 1); } }
    var idx = {}; keys.forEach(function (k, i) { idx[k] = i; });
    var fi = {}; FCS.forEach(function (f, j) { fi[f.no] = j; });
    var mat = keys.map(function () { return FCS.map(function () { return 0; }); });
    rows.forEach(function (r) {
      var k = R.unit === 'day' ? r.ymd : r.ym; if (idx[k] === undefined) return; var w = r.ctrt / 100;
      var v = metric === 'n' ? w : (metric === 'cmip' ? r.cmip * w : r.canp * w); mat[idx[k]][fi[fcKey(r.no)]] += v;
    });
    var tot = mat.map(function (a) { return sum(a); }), mx = Math.max.apply(null, [1].concat(tot)) * 1.1;
    var W = 1100, H = 260, L = 64, Rr = 10, T = 12, B = 34, bw = (W - L - Rr) / keys.length;
    var s = '<svg viewBox="0 0 ' + W + ' ' + H + '" width="100%">';
    for (var g = 0; g <= 4; g++) { var yy = T + (H - T - B) * (1 - g / 4); s += '<line x1="' + L + '" x2="' + (W - Rr) + '" y1="' + yy + '" y2="' + yy + '" stroke="#E1E4EC"/><text x="' + (L - 8) + '" y="' + (yy + 4) + '" text-anchor="end" font-size="12" fill="#5A667E">' + m.ax(mx * g / 4) + '</text>'; }
    keys.forEach(function (k, i) {
      var y0 = H - B, x = L + i * bw + bw * 0.15, w = bw * 0.7;
      FCS.forEach(function (f, j) { var v = mat[i][j]; if (!v) return; var h = (H - T - B) * v / mx; y0 -= h; s += '<rect x="' + x + '" y="' + y0 + '" width="' + w + '" height="' + h + '" fill="' + col(f.ci) + '"><title>' + k + ' ' + f.name + ': ' + m.f(v) + m.u + '</title></rect>'; });
      var lab = R.unit === 'day' ? String(+k.slice(8)) : k.slice(2).replace('-', '.');
      if (R.unit === 'month' || +k.slice(8) % 2 === 1 || keys.length <= 12) s += '<text x="' + (x + w / 2) + '" y="' + (H - B + 16) + '" text-anchor="middle" font-size="11" fill="#5A667E">' + lab + '</text>';
    });
    return s + '</svg>';
  }

  // 상태별 청약 리스트 — 이번달/이전달 탭: 청약상태(계약상태 분류)별로 묶어 SOL 열 그대로 보여 준다
  var ST_ORDER = ['성립', '신계약마감', '청약후미입금', '신계약청약', '청약철회접수', '청약철회지급', '거절접수', '거절지급', '모니터링거절접수', '모니터링거절지급'];
  function stClass(k) { return k === '성립' ? 't3' : (k.indexOf('철회') >= 0 || k.indexOf('거절') >= 0 ? 't1' : 't2'); }
  function listPanel(R) {
    if (state.period !== 'this' && state.period !== 'prev') return '';
    var rows = C.rows.filter(function (r) { return r.ymd >= R.from && r.ymd <= R.to; }).sort(function (a, b) { return a.ymd < b.ymd ? 1 : a.ymd > b.ymd ? -1 : (a.ppnu < b.ppnu ? 1 : -1); });
    var groups = {}; rows.forEach(function (r) { (groups[r.st] = groups[r.st] || []).push(r); });
    var keys = Object.keys(groups).sort(function (a, b) { var ia = ST_ORDER.indexOf(a), ib = ST_ORDER.indexOf(b); return (ia < 0 ? 99 : ia) - (ib < 0 ? 99 : ib); });
    var nm = function (r) { return isCurNo(r.no) ? r.fc : '(' + r.fc + ')'; };
    var esc2 = function (x) { return String(x == null ? '' : x).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
    var head = '<tr><th>계약일자</th><th>심사상태</th><th>청약번호</th><th>증권번호</th><th>FC명</th><th class="l">상품명</th><th>계약자</th><th>주피보험자</th><th>보험료</th><th>환산보험료</th><th>기여율</th><th>건수</th><th>월납화보험료</th><th>CANP</th><th>승낙일자</th><th>철회일자</th></tr>';
    var out = '<h3 class="h3" style="margin-top:20px">청약 상태별 리스트 (SOL 청약계약List조회, 총 ' + rows.length + '건)</h3><p class="note" style="margin:0 0 8px">건수·월납화보험료·CANP는 기여율을 곱한 값이고 보험료·환산보험료는 SOL 화면 값 그대로입니다. 성립만 위 통계에 반영됩니다. 진단종류는 아직 수집하지 않았습니다.</p>';
    keys.forEach(function (k) {
      var g = groups[k], n = sum(g.map(function (r) { return r.ctrt / 100; })), cm = sum(g.map(function (r) { return r.cmip * r.ctrt / 100; })), ca = sum(g.map(function (r) { return r.canp * r.ctrt / 100; }));
      var body = g.map(function (r) {
        return '<tr><td>' + esc2(r.ymd) + '</td><td>' + esc2(r.ispa || '-') + '</td><td>' + esc2(r.ppnu) + '</td><td>' + esc2(r.pol || '-') + '</td><td>' + esc2(nm(r)) + '</td><td class="l">' + esc2(r.goods) + '</td><td>' + esc2(r.owner) + '</td><td>' + esc2(r.insured) + '</td><td>' + fmtW(r.sum) + '</td><td>' + (r.conv ? fmtW(r.conv) : '-') + '</td><td>' + r.ctrt + '%</td><td><b>' + fmtN(r.ctrt / 100) + '</b></td><td>' + fmtW(r.cmip * r.ctrt / 100) + '</td><td>' + fmtW(r.canp * r.ctrt / 100) + '</td><td>' + esc2(r.ok || '-') + '</td><td>' + esc2(r.wd || '-') + '</td></tr>';
      }).join('');
      out += '<details class="stg' + (k === '성립' ? '' : ' nok') + '" open><summary><span class="tag ' + stClass(k) + '">' + esc2(k) + '</span> <b>' + g.length + '건</b> <span class="mut">· 건수 ' + fmtN(n) + ' · 월납화보험료 ' + man(cm) + '만원 · CANP ' + man(ca) + '만원</span></summary><div class="wrap"><table class="ct">' + head + body + '</table></div></details>';
    });
    return out + (rows.length ? '' : '<p class="mut">이 기간에 청약이 없습니다.</p>');
  }
  function topPanel() {
    var R = range(state.period), A = agg(R.from, R.to), B = agg(R.cf, R.ct);
    var td = localToday(), notYet = !state.test && state.period === 'this' && td.slice(0, 7) > D.upto;
    var day = R.unit === 'day';
    var tabs = PERIODS.map(function (x) { return '<button class="' + (state.period === x.k ? 'on' : '') + '" data-period="' + x.k + '">' + x.t + '</button>'; }).join('');
    var mt = METR.map(function (m) { return '<button class="' + (state.pmetric === m.k ? 'on' : '') + '" data-pmetric="' + m.k + '">' + m.t + '</button>'; }).join('');
    var rate = A.total ? Math.round(A.ok.length / A.total * 100) : 0, rateB = B.total ? Math.round(B.ok.length / B.total * 100) : 0;
    var card = function (t, v, d) { return '<div class="k"><div class="l">' + t + '</div><div class="v">' + v + '</div><div class="d">' + d + '</div></div>'; };
    var avg = A.n ? A.cmip / A.n : 0, avgB = B.n ? B.cmip / B.n : 0;
    var avgC = A.n ? A.canp / A.n : 0, avgCB = B.n ? B.canp / B.n : 0;
    // 평균 건수 = 성립 건수 ÷ 그 달(기간 끝) 근무 중이던 팀원 수(ESL 포함, 실적과 무관). 근무 시작 = 위촉일
    var hc = function (dateStr) { return (D.currentFcs || []).filter(function (f) { return !f.joined || f.joined <= dateStr; }).length || 1; };
    var nFc = hc(R.to), nFcB = hc(R.ct), per = A.n / nFc, perB = B.n / nFcB;
    var f1 = function (v) { return v.toLocaleString('ko-KR', { minimumFractionDigits: 1, maximumFractionDigits: 1 }); };
    var cards = card('성립 건수', fmtN(A.n) + '<small>건</small>', '비교 ' + diffHtml(A.n, B.n, fmtN)) +
      card('CMIP', man(A.cmip) + '<small>만원</small>', '비교 ' + diffHtml(A.cmip, B.cmip, man)) +
      card('CANP', man(A.canp) + '<small>만원</small>', '비교 ' + diffHtml(A.canp, B.canp, man)) +
      card('평균 건수', f1(per) + '<small>건</small>', '비교 ' + diffHtml(per, perB, f1)) +
      card('CMIP / 건', man(avg) + '<small>만원</small>', '비교 ' + diffHtml(avg, avgB, man)) +
      card('CANP / 건', man(avgC) + '<small>만원</small>', '비교 ' + diffHtml(avgC, avgCB, man));
    var rowsH = FCS.map(function (f) {
      var a = A.byFc[f.no] || { n: 0, cmip: 0, canp: 0 };
      return '<tr class="' + (isCurNo(f.no) ? 'cur' : '') + '"><td class="' + (f.group ? 'wrapc' : '') + '"><span class="sw" style="background:' + col(f.ci) + '"></span>' + f.name + '</td><td><b>' + (a.n ? fmtN(a.n) : '-') + '</b></td><td>' + (a.cmip ? man(a.cmip) : '-') + '</td><td>' + (a.canp ? man(a.canp) : '-') + '</td></tr>';
    }).join('');
    var mixKeys = ['종신', '건강·보장', '변액', '연금', '기타'], mixMax = Math.max.apply(null, mixKeys.map(function (k) { return A.mix[k] || 0; }).concat([0.01]));
    var mixH = mixKeys.map(function (k) { var v = A.mix[k] || 0; return '<div class="mixr"><span>' + k + '</span><i style="width:' + Math.max(v / mixMax * 100, v ? 3 : 0) + '%"></i><b>' + fmtN(v) + '건 (' + (A.n ? Math.round(v / A.n * 100) : 0) + '%)</b></div>'; }).join('');
    var legendH = '<div class="legend">' + FCS.map(function (f) { return '<span><i class="sw" style="background:' + col(f.ci) + '"></i>' + f.name + '</span>'; }).join('') + '</div>';
    var stc = {}; A.rows.forEach(function (r) { stc[r.st] = (stc[r.st] || 0) + 1; });
    var stKeys = Object.keys(stc).sort(function (a, b) { return stc[b] - stc[a]; }), stMax = Math.max.apply(null, stKeys.map(function (k) { return stc[k]; }).concat([1]));
    var stH = stKeys.map(function (k) { return '<div class="mixr st"><span>' + k + '</span><i class="' + (k === '성립' ? '' : (k.indexOf('철회') >= 0 || k.indexOf('거절') >= 0 ? 'bad' : 'mid')) + '" style="width:' + Math.max(stc[k] / stMax * 100, 4) + '%"></i><b>' + stc[k] + '건</b></div>'; }).join('') || '<p class="mut">청약이 없습니다.</p>';
    var distH = '<h3 class="h3">팀 업적 분포 <small class="mut">(CMIP·CANP 만원)</small></h3><div class="wrap"><table class="cmpt"><tr><th>FC</th><th>건수</th><th>CMIP</th><th>CANP</th></tr>' + rowsH + '</table></div>';
    var mixBox = '<h3 class="h3" style="margin-top:14px">상품 구성 <small class="mut">(상품명으로 추정)</small></h3>' + mixH;
    var grid = day ? ''
      : '<div class="two3b"><div class="colL"><h3 class="h3">월별 추이 <span class="tabs mini">' + mt + '</span></h3>' + stackedPeriod(R, state.pmetric) + legendH + '<h3 class="h3" style="margin-top:14px">청약 상태 현황</h3>' + stH + '</div></div>';
    var testBar = '<label class="chk testbar"><input id="testChk" type="checkbox"' + (state.test ? ' checked' : '') + '> 테스트 기준월 사용: ' + TESTYM + ' (' + new Date().getFullYear() + '년 중 청약이 가장 많은 달)을 "이번달"로 봅니다</label>';
    return '<section class="topsec"><div class="tophead"><h2>계약 통계</h2><div class="tabs mini big">' + tabs + '</div></div>' + testBar +
      '<p class="sub" style="margin:4px 0 10px">' + PERIODS.filter(function (x) { return x.k === state.period; })[0].t + ' <b>' + R.from + ' ~ ' + R.to + '</b> · 비교 기간: ' + R.cmp + ' · 성립 계약만, 기여율 반영</p>' +
      (notYet ? '<div class="warn">이번달(' + td.slice(0, 7) + ') 계약은 아직 수집하지 않았습니다. 마지막 수집은 ' + D.upto + '까지이며, 수집하면 이 화면에 반영됩니다.</div>' : '') +
      '<div class="kpis5">' + cards + '</div>' + grid + listPanel(R) + '</section>';
  }
  function render() {
    var base = state.base, y = base.slice(0, 4), mo = parseInt(base.slice(5), 10), py = String(parseInt(y, 10) - 1);
    var prev = prevYm(base, 1);
    var partial = base === D.upto;
    var cards = METR.map(function (m) {
      return '<div class="k"><div class="l">' + base + ' ' + m.t + (m.k === 'n' ? '' : '') + '</div><div class="v">' + (m.k === 'n' ? fmtN(tot(base, m.k)) + '<small>건</small>' : man(tot(base, m.k)) + '<small>만원</small>') + '</div><div class="d">전월 대비 ' + diffHtml(tot(base, m.k), prev >= D.from ? tot(prev, m.k) : 0, m.k === 'n' ? fmtN : man) + '</div></div>';
    }).join('');
    var ytdN = sum(FCS.map(function (f) { return ytd(f.no, y, mo, 'n'); })), ytdP = sum(FCS.map(function (f) { return ytd(f.no, py, mo, 'n'); }));
    cards += '<div class="k"><div class="l">' + y + '년 누적 건수 (1~' + mo + '월)</div><div class="v">' + fmtN(ytdN) + '<small>건</small></div><div class="d">전년 동기간 ' + diffHtml(ytdN, ytdP, fmtN) + '</div></div>';

    // FC table (사번순)
    var rows = FCS.map(function (f, j) { return { f: f, j: j, n: ytd(f.no, y, mo, 'n'), cmip: ytd(f.no, y, mo, 'cmip'), canp: ytd(f.no, y, mo, 'canp'), pn: ytd(f.no, py, mo, 'n'), cur: val(f.no, base, 'n') }; })
      .filter(function (r) { return r.n || r.pn || r.cur || isCurNo(r.f.no); });
    var tbl = '<div class="wrap"><table class="cmpt"><tr><th>FC</th><th>' + base + ' 건수</th><th>누적 건수</th><th>누적 CMIP(만원)</th><th>누적 CANP(만원)</th></tr>' +
      rows.map(function (r) { return '<tr class="' + (D.current.indexOf(r.f.no) >= 0 ? 'cur' : '') + '"><td class="' + (r.f.group ? 'wrapc' : '') + '"><span class="sw" style="background:' + col(FCS[r.j].ci) + '"></span>' + r.f.name + '</td><td>' + (r.cur ? fmtN(r.cur) : '-') + '</td><td><b>' + (r.n ? fmtN(r.n) : '-') + '</b></td><td>' + (r.cmip ? man(r.cmip) : '-') + '</td><td>' + (r.canp ? man(r.canp) : '-') + '</td></tr>'; }).join('') + '</table></div>';

    // contract status in base month
    var st = {}; C.rows.filter(function (r) { return r.ym === base; }).forEach(function (r) { st[r.st] = (st[r.st] || 0) + 1; });
    var order = ['성립', '신계약마감', '청약후미입금', '청약철회지급', '청약철회접수', '거절지급', '거절접수', '모니터링거절지급', '모니터링거절접수', '신계약청약'];
    var stHtml = Object.keys(st).sort(function (a, b) { var ia = order.indexOf(a), ib = order.indexOf(b); return (ia < 0 ? 99 : ia) - (ia < 0 && ib < 0 ? 0 : ib < 0 ? 99 : ib); }).map(function (k) {
      return '<div class="al"><span class="tag ' + (k === '성립' ? 't3' : (k.indexOf('철회') >= 0 || k.indexOf('거절') >= 0 ? 't1' : 't2')) + '">' + k + '</span><b>' + st[k] + '건</b></div>';
    }).join('');

    var opts = MONTHS.slice().reverse().map(function (ym) { return '<option value="' + ym + '"' + (ym === base ? ' selected' : '') + '>' + ym + '</option>'; }).join('');
    var tabs = METR.map(function (m) { return '<button class="' + (state.metric === m.k ? 'on' : '') + '" data-metric="' + m.k + '">' + m.t + '</button>'; }).join('');
    var legend = '<div class="legend">' + FCS.map(function (f, j) { return '<span><i class="sw" style="background:' + col(FCS[j].ci) + '"></i>' + f.name + '</span>'; }).join('') + '</div>';

    return '<h1>대시보드</h1><p class="sub">성립 계약 · 기여율 반영 · 수집 ' + D.from + ' ~ ' + D.upto + '</p>' + topPanel() +
      '<div class="topbar" style="margin-top:26px"><div><h2 class="secH">월별 분석</h2></div><label class="pick">기준월 <select id="baseSel">' + opts + '</select></label></div>' +
      (partial ? '<div class="warn">최근 달(' + base + ')은 마감이 끝나기 전이라 이후 수집 때 바뀔 수 있습니다.</div>' : '') +
      '<div class="kpis4">' + cards + '</div>' +
      '<section><h2>최근 12개월 추이 <span class="tabs mini">' + tabs + '</span></h2>' + chart(state.metric) + legend + '</section>' +
      '<div class="two2"><section><h2>' + y + '년 누적 · FC별 (사번순)</h2>' + tbl + '</section>' +
      '<section><h2>' + base + ' 청약 현황</h2>' + (stHtml || '<p class="mut">데이터 없음</p>') + '<p class="note">성립이 아닌 계약은 실적에 반영되지 않습니다.</p>' +
      '<h2 style="margin-top:18px">할 일 · 주의</h2>' + todo() + '</section></div>';
  }
  function bind() {
    var sel = document.getElementById('baseSel');
    if (sel) sel.onchange = function () { state.base = sel.value; window.APP.rerender(); };
    var tc = document.getElementById('testChk'); if (tc) tc.onchange = function () { state.test = tc.checked; window.APP.rerender(); };
    document.querySelectorAll('[data-period]').forEach(function (b) { b.onclick = function () { state.period = b.getAttribute('data-period'); window.APP.rerender(); }; });
    document.querySelectorAll('[data-pmetric]').forEach(function (b) { b.onclick = function () { state.pmetric = b.getAttribute('data-pmetric'); window.APP.rerender(); }; });
    document.querySelectorAll('[data-metric]').forEach(function (b) { b.onclick = function () { state.metric = b.getAttribute('data-metric'); window.APP.rerender(); }; });
  }
  window.PAGES = window.PAGES || {};
  window.PAGES.dashboard = { render: render, bind: bind };
})();
