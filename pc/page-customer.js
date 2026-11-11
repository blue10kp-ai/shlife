(function () {
  'use strict';
  var K = window.CUSTOMERS;
  // 고객추가정보 엑셀(신한FC M이 쓰는 '기본'·'상담기록' 시트)에서 읽은 data/extra.js: 통신사·차량번호·자동차보험갱신일·미팅장소·메모, 상담기록
  //   고객 찾기: 담당FC + 이름 (이름 뒤 '(YYMMDD)'는 같은 이름 구분용이라 생년월일과 맞을 때만)
  var X = window.CUSTEXTRA || { basic: [], consult: [] };
  var xKey = function (fc, name, birth) {
    var m = String(name || '').match(/^(.*?)\((\d{6})\)$/);
    if (m) return (birth && String(birth).replace(/\D/g, '').slice(2, 8) === m[2]) ? fc + '|' + m[1] : null;
    return fc + '|' + name;
  };
  var xBasic = {}, xCons = {};
  var xPick = function (arr, c) { return (arr || []).filter(function (e) { return !e.sfx || String(c['생년월일'] || '').replace(/\D/g, '').slice(2, 8) === e.sfx; }); };
  function indexExtra() {
    X = window.CUSTEXTRA || X; xBasic = {}; xCons = {};
    X.basic.forEach(function (b) { var k = b.fc + '|' + String(b.name).replace(/\(\d{6}\)$/, ''), sfx = (String(b.name).match(/\((\d{6})\)$/) || [])[1]; (xBasic[k] = xBasic[k] || []).push({ b: b, sfx: sfx }); });
    X.consult.forEach(function (r) { var k = r.fc + '|' + String(r.name).replace(/\(\d{6}\)$/, ''), sfx = (String(r.name).match(/\((\d{6})\)$/) || [])[1]; (xCons[k] = xCons[k] || []).push({ r: r, sfx: sfx }); });
  }
  function applyExtraBasic() {   // 고객 시트에 값이 없을 때만 고객추가정보 기본 시트 값으로 채움
    K.customers.forEach(function (c) {
      xPick(xBasic[c['담당FC'] + '|' + c['이름']], c).forEach(function (e) {
        var b = e.b;
        if (b.carrier && !c['통신사']) c['통신사'] = b.carrier; if (b.car && !c['차량번호']) c['차량번호'] = b.car; if (b.renew && !c['자동차보험갱신일']) c['자동차보험갱신일'] = b.renew;
        if (b.meeting && !c['미팅장소']) c['미팅장소'] = b.meeting; if (b.memo && !c['메모']) c['메모'] = b.memo;
      });
    });
  }
  indexExtra(); applyExtraBasic();
  // 상담기록: 고객추가정보 엑셀(extra.js)의 기록 + 이 브라우저에서 입력·수정·삭제한 내용 { 고객ID: { add: [{id, at, text}], del: [키], mod: {키: 바꾼 내용} } }
  var CONS_KEY = 'shfc.consultEdits';
  function readCons() { try { return JSON.parse(localStorage.getItem(CONS_KEY) || '{}'); } catch (e) { return {}; } }
  function writeCons(o) { try { localStorage.setItem(CONS_KEY, JSON.stringify(o)); } catch (e) { alert('이 브라우저에 저장하지 못했습니다.'); } }
  function consOv(id) { var o = readCons(); return { all: o, ov: o[id] || { add: [], del: [], mod: {} } }; }
  function nowStamp() { var d = new Date(), p = function (n) { return n < 10 ? '0' + n : '' + n; }; return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate()) + ' ' + p(d.getHours()) + ':' + p(d.getMinutes()); }
  function consultsOf(c) {
    var ov = consOv(c['고객ID']).ov;
    var base = xPick(xCons[c['담당FC'] + '|' + c['이름']], c).map(function (e) { return e.r; }).map(function (r) { var key = 'x|' + r.at + '|' + r.text; return { key: key, at: r.at, text: (ov.mod || {})[key] != null ? ov.mod[key] : r.text, voice: r.voice, src: 'x' }; })
      .filter(function (r) { return (ov.del || []).indexOf(r.key) < 0; });
    var add = (ov.add || []).map(function (a) { return { key: 'a|' + a.id, at: a.at, text: a.text, voice: '', src: 'a' }; });
    return base.concat(add).sort(function (a, b) { return String(b.at).localeCompare(String(a.at)); });
  }
  function consChange(c, fn) { var r = consOv(c['고객ID']); r.ov.add = r.ov.add || []; r.ov.del = r.ov.del || []; r.ov.mod = r.ov.mod || {}; fn(r.ov); r.all[c['고객ID']] = Object.assign({ name: c['이름'], fc: c['담당FC'] }, r.ov); writeCons(r.all); }
  // 고객 정보 수정: 이 브라우저(localStorage)에 저장해 화면에 반영하고, 수정 내역을 파일로 내려받을 수 있다
  //   { 고객ID: { name, fc, at, o: {항목: 처음 값}, v: {항목: 바꾼 값} } }
  var EDIT_KEY = 'shfc.custEdits';
  function readEdits() { try { return JSON.parse(localStorage.getItem(EDIT_KEY) || '{}'); } catch (e) { return {}; } }
  (function () {
    var ed = readEdits();
    K.customers.forEach(function (c) { var e = ed[c['고객ID']]; if (e && e.v) Object.keys(e.v).forEach(function (k) { c[k] = e.v[k]; }); });
  })();
  // 계약 단위 유입경로(소개)·소개자: Drive 미연결이면 이 브라우저에 보관 { 증권번호: { route, intro } }
  var CE_KEY = 'shfc.contractEdits';
  function readCE() { try { return JSON.parse(localStorage.getItem(CE_KEY) || '{}'); } catch (e) { return {}; } }
  (function () { var ce = readCE(); Object.keys(K.contracts).forEach(function (id) { K.contracts[id].forEach(function (x) { var e = ce[x.pol]; if (e) { x.route = e.route; x.intro = e.intro; } }); }); })();
  var insIdx = {};   // 피보험자 고객ID → 그 사람이 피보험자인 계약 (소개 계약은 계약자·피보험자 모두 소개로 표시)
  function indexIns() { insIdx = {}; Object.keys(K.contracts).forEach(function (id) { K.contracts[id].forEach(function (x) { [x.insured, x.insured2].forEach(function (i) { if (i && i !== id) (insIdx[i] = insIdx[i] || []).push(x); }); }); }); }
  indexIns();
  function allContractsOf(c) { return (K.contracts[c['고객ID']] || []).concat(insIdx[c['고객ID']] || []); }
  function routeText(c) {
    var r = c['최초유입경로'] || '', parts = [], sc = allContractsOf(c).filter(function (x) { return x.route === '소개'; });
    if (sc.length) { parts.push('소개'); var n = sc.map(function (x) { return x.intro; }).filter(Boolean)[0]; if (n) parts.push(n); }
    else {
      if (r && r !== '이관') parts.push(r);
      if (r === 'DB' && c['유입DB명']) parts.push(c['유입DB명']);
      if ((r === '소개' || r === '지인') && c['소개자명']) parts.push(c['소개자명']);
    }
    if (c['유입일']) parts.push(c['유입일']);
    return parts.join(' · ');
  }
  // 자동차보험갱신일: 'MMDD'(또는 날짜) → '3월6일'
  function mdParts(v) { var d = String(v || '').replace(/\D/g, ''); if (d.length >= 8) d = d.slice(4, 8); return d.length === 4 ? [parseInt(d.slice(0, 2), 10), parseInt(d.slice(2), 10)] : null; }
  function mdText(v) { var p = mdParts(v); return p ? p[0] + '월' + p[1] + '일' : String(v || ''); }
  var PAGE = 1e9;   // 쪽 나눔 없이 전체 표시
  var esc = function (s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  var won = function (v) { return v == null || v === '' ? '-' : Math.round(v).toLocaleString('ko-KR'); };
  var visible = function (c) { return c['현재단계'] !== '가망'; };   // 가망은 회사 분류라 고객 목록·집계에서 제외 (데이터는 그대로)
  var custAll = function () { return K.customers.filter(visible); };
  var custTeam = function () { return (K.teamCustomers || K.customers).filter(visible); };
  // 고객 분류 칩: 전체 / 본인(계약) / 이관(이관관리) / DB / 가망 | 이달이관 / 따행
  var CHIP_STAGES = [['계약', '본인'], ['이관관리', '이관'], ['DB', 'DB']];   // 가망은 회사 분류라 칩에서 뺌 (진행이 비슷한 개념)
  var STAGES = ['DB', '가망', '진행', '계약', '이관관리', '피보험자'];
  // FC 드롭다운 목록: 이달 이관은 내 팀 전체, 그 외는 지금 보기 범위의 고객 FC (직급 먼저, 같은 직급은 사번순)
  var fcCache = {};
  function fcList() {
    var key = state.transfer || state.progress ? 'team' : 'view';
    if (!fcCache[key]) { var a = []; (state.transfer || state.progress ? custTeam() : custAll()).forEach(function (c) { if (a.indexOf(c['담당FC']) < 0) a.push(c['담당FC']); }); a.sort(function (x, y) { return window.ACCESS.fcCompare({ name: x }, { name: y }); }); fcCache[key] = a; }
    return fcCache[key];
  }
  var SRC = []; custAll().forEach(function (c) { var s = c['최초유입경로'] || '(없음)'; if (SRC.indexOf(s) < 0) SRC.push(s); });
  var state = { q: '', fc: '', stage: '', src: '', page: 0, sel: null, haeng: false, hs: '', transfer: false, progress: false, edit: false, cedit: '' };
  var todayStr = (function () { var d = new Date(), p = function (n) { return n < 10 ? '0' + n : '' + n; }; return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate()); })();
  (function () { var hm = (location.hash || '').match(/sel=(C\d+)/); if (hm) state.sel = hm[1]; })();
  var ymNow = todayStr.slice(0, 7);

  function addDays(d, n) { var t = new Date(d + 'T00:00:00'); t.setDate(t.getDate() + n); return t.toISOString().slice(0, 10); }
  function contractsOf(c) { return K.contracts[c['고객ID']] || []; }
  function active(x) { return x.st === '유지' || x.st === '완납' || x.st === '납입면제' || x.st === '연금지급'; }
  function filtered() {
    var q = state.q.trim().toLowerCase();
    return (state.transfer || state.progress ? custTeam() : custAll()).filter(function (c) {
      if (state.fc && c['담당FC'] !== state.fc) return false;
      if (state.stage && c['현재단계'] !== state.stage) return false;
      if (state.src && (c['최초유입경로'] || '(없음)') !== state.src) return false;
      if (state.transfer && !transferOf(c).length) return false;
      if (state.progress && c['현재단계'] !== '진행') return false;   // 진행중인 고객
      if (state.haeng) { var hs = hStatus(c); if (!hs || (state.hs && hs !== state.hs)) return false; }
      if (q && (c['이름'] + ' ' + c['연락처'] + ' ' + c['주소'] + ' ' + c['직업'] + ' ' + c['고객ID']).toLowerCase().indexOf(q) < 0) return false;
      return true;
    }).sort(function (a, b) { return String(a['이름']).localeCompare(String(b['이름']), 'ko'); });
  }
  // 이달 이관 버튼: 이관일이 이번 달인 계약이 있는 고객 (계약 단위로 표시)
  function transferOf(c) { return contractsOf(c).filter(function (x) { return String(x.trans || '').slice(0, 7) === ymNow; }); }
  // 표 칸: 연락처 / 생일(위아래 두 줄), 직업 · 지역 · 주소는 각각 따로 한 줄
  var cellContact = function (c) { return esc(c['연락처'] || '-') + '<br><span class="sm">' + esc(c['생년월일'] || '-') + '</span>'; };
  var cellPerson = function (c) { return '<td class="l job" title="' + esc(c['직업'] || '') + '">' + esc(c['직업'] || '-') + '</td><td class="l">' + esc(c['지역'] || '-') + '</td><td class="l">' + esc(c['주소'] || '-') + '</td>'; };
  function transferBody(view) {
    var n = 0, h = '';
    view.forEach(function (c) {
      transferOf(c).sort(function (a, b) { return String(b.trans).localeCompare(String(a.trans)); }).forEach(function (x) {
        n++;
        h += '<tr class="pick-row ' + (state.sel === c['고객ID'] ? 'selrow' : '') + (active(x) ? '' : ' gray') + '" data-cid="' + esc(c['고객ID']) + '"><td>' + n + '</td><td class="l nm"><b class="' + gcls(c) + '">' + esc(c['이름']) + '</b></td><td>' + cellContact(c) + '</td>' + cellPerson(c) + '<td>' + esc(x.rfc || '-') + '</td><td class="l">' + esc(x.goods) + '</td><td>' + won(x.sum) + '</td><td>' + esc(x.st) + '</td><td>' + esc(x.trans) + '</td></tr>';
      });
    });
    return h;
  }
  // 따행(따뜻한 행복) 버튼: 이달 따행 대상 고객만. 하위 구분 = 방문 완료 / 약속 / 미완료
  function hStatus(c) {
    var h = haengNow(c); if (!h || h.target === '아니오' || h.target === 'N') return '';
    return h.done ? 'done' : (h.plan ? 'plan' : 'todo');
  }
  function haengNow(c) { var h = (K.haeng[c['고객ID']] || []).filter(function (x) { return String(x.ym).slice(0, 7) === ymNow; })[0]; return h; }
  // 고객 성향에 따른 이름 색: Good 검정(기본) · Bad 주황 · Shit 빨강
  function gcls(c) { var g = String(c['성향'] || '').trim().toLowerCase(); return g === 'bad' ? 'g-bad' : g === 'shit' ? 'g-shit' : 'g-good'; }
  function rowH(f, html) { return html ? '<div class="fr"><span>' + f + '</span><b>' + html + '</b></div>' : ''; }
  function row(f, v) { return v ? '<div class="fr"><span>' + f + '</span><b>' + esc(v) + '</b></div>' : ''; }


  // ===== 보장분석: 설계동의 확인 → 만들어진 파일 내려받기 / 없으면 요청 =====
  var FILES = (window.CUSTFILES && window.CUSTFILES.byId) || {};
  var REQ_KEY = 'shfc.covanReq';
  function reqList() { try { return JSON.parse(localStorage.getItem(REQ_KEY) || '[]'); } catch (e) { return []; } }
  function consentOf(c) {
    var d = c['가입설계동의 만료일'];
    if (!d) return { ok: false, t: '가입설계동의 정보가 없습니다' };
    return d >= todayStr ? { ok: true, t: '가입설계동의 유효 · 만료 ' + d } : { ok: false, t: '가입설계동의가 만료되었습니다 (' + d + ')' };
  }
  function kb(n) { return n >= 1048576 ? (n / 1048576).toFixed(1) + 'MB' : Math.max(1, Math.round(n / 1024)) + 'KB'; }
  function covanSection(c) {
    var cs = consentOf(c), all = FILES[c['고객ID']] || [];
    var cov = all.filter(function (f) { return f.k === '스마트보장분석' || f.k === '손해담보'; }).sort(function (a, b) { return (a.k === '스마트보장분석' ? 0 : 1) - (b.k === '스마트보장분석' ? 0 : 1); });
    var req = reqList().filter(function (r) { return r.id === c['고객ID']; })[0];
    var h = '<div class="cv ' + (cs.ok ? 'ok' : 'no') + '">' + esc(cs.t) + '</div>';
    if (cov.length) {
      h += cov.map(function (f) { return '<div class="dlr"><a class="dl" href="' + esc(f.link || '') + '" target="_blank" rel="noopener">' + esc(f.k) + ' 다운로드</a></div>'; }).join('');
    } else if (cs.ok) {
      /* 파일이 없으면 안내 문구 없이 동의 상태 줄만 표시 */
    } else {
      h += '<p class="mut">가입설계동의가 있어야 보장분석을 진행할 수 있습니다.</p>';
    }
    var rest = all.filter(function (f) { return f.k !== '상세보장' && f.k !== '고객카드' && f.k !== '손해담보' && !(f.k === '스마트보장분석' && /\.pdf$/i.test(f.n)) && /\.(pdf|xlsx|xlsm|xls)$/i.test(f.n); });   // 분석 보고서(PDF·엑셀)만
    if (rest.length) {
      h += '<details class="files"><summary>보고서 ' + rest.length + '개</summary>' + rest.map(function (f) { return '<div class="dlr"><a class="dl" href="' + esc(f.link || '') + '" target="_blank" rel="noopener">' + esc(f.n) + '</a></div>'; }).join('') + '</details>';
    }
    return h;
  }

  // 통신사: 휴대전화 번호 앞 드롭다운 (알뜰폰은 쓰는 망별로 구분)
  var CARRIERS = ['SKT', 'KT', 'LG', 'SKT알뜰', 'KT알뜰', 'LG알뜰'];
  var rowE = function (f, html) { return '<div class="fr fe"><span>' + f + '</span><b>' + html + '</b></div>'; };
  var inp = function (c, k, ph, type) { return '<input data-e="' + k + '" type="' + (type || 'text') + '" value="' + esc(c[k] || '') + '"' + (ph ? ' placeholder="' + esc(ph) + '"' : '') + ' autocomplete="off">'; };
  function carrierSel(c) {
    var cur = c['통신사'] || '', list = CARRIERS.indexOf(cur) < 0 && cur ? [cur].concat(CARRIERS) : CARRIERS;
    return '<select data-e="통신사" class="csel"><option value="">통신사</option>' + list.map(function (o) { return '<option value="' + esc(o) + '"' + (o === cur ? ' selected' : '') + '>' + esc(o) + '</option>'; }).join('') + '</select>';
  }
  function mdSel(c) {
    var p = mdParts(c['자동차보험갱신일']) || [0, 0];
    return '<span class="efr"><select data-emm><option value="">월</option>' + Array.from({ length: 12 }, function (_, i) { return '<option value="' + (i + 1) + '"' + (p[0] === i + 1 ? ' selected' : '') + '>' + (i + 1) + '월</option>'; }).join('') + '</select><select data-edd><option value="">일</option>' + Array.from({ length: 31 }, function (_, i) { return '<option value="' + (i + 1) + '"' + (p[1] === i + 1 ? ' selected' : '') + '>' + (i + 1) + '일</option>'; }).join('') + '</select></span>';
  }
  // 유입경로: DB·이관은 자동 값, 지인·소개는 직접 고름(소개자 입력)
  function routeSel(c) {
    var cur = c['최초유입경로'] || '', list = ['', '지인', '소개'];
    if (cur && list.indexOf(cur) < 0) list.splice(1, 0, cur);
    return '<span class="efr"><select data-e="최초유입경로" class="csel">' + list.map(function (o) { return '<option value="' + esc(o) + '"' + (o === cur ? ' selected' : '') + '>' + (o || '(없음)') + '</option>'; }).join('') + '</select>' + inp(c, '소개자명', '소개자') + '</span>';
  }
  function gradeDots(c) {
    var g = String(c['성향'] || 'Good').toLowerCase();
    return '<span class="gdots">' + ['Good', 'Bad', 'Shit'].map(function (o) { return '<button type="button" data-g="' + o + '" class="gdot g-' + o.toLowerCase() + (g === o.toLowerCase() ? ' on' : '') + '" title="' + o + '"><i></i><em>' + o + '</em></button>'; }).join('') + '</span><input type="hidden" data-e="성향" value="' + esc(c['성향'] || 'Good') + '">';
  }
  function gradeTog(c) {
    var g = String(c['성향'] || 'Good').toLowerCase();
    return '<span class="gtog">' + ['Good', 'Bad', 'Shit'].map(function (o) { return '<button type="button" data-g="' + o + '" class="g-' + o.toLowerCase() + (g === o.toLowerCase() ? ' on' : '') + '">' + o + '</button>'; }).join('') + '</span><input type="hidden" data-e="성향" value="' + esc(c['성향'] || 'Good') + '">';
  }
  function editsBox() {
    var n = Object.keys(readEdits()).length, m = Object.keys(readCons()).length;
    return n || m ? '<p class="note">수정한 고객 ' + n + '명 · 상담기록을 바꾼 고객 ' + m + '명 · <button type="button" class="link" id="dlEdits">수정 내역 내려받기</button></p>' : '';
  }
  // 계약 상품명: 같은 계약의 상세보장 PDF(계약일·보험료로 찾음)가 있으면 눌러서 연다
  function pdfHref(c, x) {
    var files = ((window.CUSTFILES && window.CUSTFILES.byId || {})[c['고객ID']] || []).filter(function (f) { return f.k === '상세보장' && /\.pdf$/i.test(f.n); });
    var d = String(x.start || '').replace(/-/g, ''), fee = won(x.sum) + '원';
    var hit = files.filter(function (f) { return d && f.n.indexOf(d) === 0; });
    if (hit.length > 1) { var h2 = hit.filter(function (f) { return f.n.indexOf(fee) >= 0; }); if (h2.length) hit = h2; }
    if (!hit.length) hit = files.filter(function (f) { return f.n.indexOf(fee) >= 0 && String(x.goods || '').length > 3 && f.n.indexOf(String(x.goods).slice(0, 6)) >= 0; });
    return hit.length ? (hit[0].link || '') : '';
  }
  function drawer(c) {
    if (!c) return '';
    var allH = (K.history[c['고객ID']] || []).slice().sort(function (a, b) { return String(b.at).localeCompare(String(a.at)); }), cs = contractsOf(c),
      cons = consultsOf(c), hg = (K.haeng[c['고객ID']] || []).filter(function (x) { return String(x.ym).slice(0, 7) === ymNow; });
    var ct = cs.map(function (x) { var href = pdfHref(c, x); return '<tr class="' + (active(x) ? '' : 'gray') + (href ? ' pdfrow' : '') + '"' + (href ? ' data-pdf="' + esc(href) + '" title="눌러서 상품 PDF 열기"' : '') + '><td class="l"><span class="dot"></span>' + esc(x.goods) + '</td><td>' + esc(x.st) + '</td><td>' + won(x.sum) + '</td><td>' + esc(x.start) + '</td><td>' + esc(x.recruit) + '</td><td class="cr">' + (state.edit ? '<select data-cr="' + esc(x.pol) + '" data-crfc="' + esc(x.fc) + '"><option value=""></option><option value="소개"' + (x.route === '소개' ? ' selected' : '') + '>소개</option></select><input data-ci="' + esc(x.pol) + '" value="' + esc(x.intro || '') + '" placeholder="소개자">' : esc(x.route ? x.route + (x.intro ? '(' + x.intro + ')' : '') : '')) + '</td></tr>'; }).join('');
    var ht = hg.map(function (h) { return '<tr><td>' + esc(String(h.ym).slice(0, 7)) + '</td><td>' + esc(h.cat) + '</td><td>' + esc(h.contact) + '</td><td>' + esc(h.plan || '-') + '</td><td>' + esc(h.done || '-') + '</td><td>' + (h.survey ? '발송' : '-') + (h.answer ? '/응답' : '') + '</td></tr>'; }).join('');
    var evs = allH.map(function (h) { return { at: h.at, type: h.type, res: h.res, memo: h.memo }; })
      .concat(cons.map(function (r) { return { at: r.at, type: '상담기록 입력', res: '', memo: '', plain: true }; }))
      .concat(readLog().filter(function (e) { return e.id === c['고객ID']; }).map(function (e) { return { at: e.at, type: e.type === '상담기록 삭제' ? '상담기록 삭제' : e.type, res: '', memo: '', plain: true }; }))
      .sort(function (x, y) { return String(y.at).localeCompare(String(x.at)); }).slice(0, 20);
    var hist = evs.map(function (h) { return '<div class="tl"><b>' + esc(h.plain ? String(h.at).slice(0, 10) : h.at) + '</b> ' + esc(h.type) + (h.res ? ' · ' + esc(h.res) : '') + (h.memo ? '<div class="mut">' + esc(h.memo) + '</div>' : '') + '</div>'; }).join('');
    var E = state.edit, canE = window.ACCESS.canEdit;
    var memoBox = '<h3>메모 <span class="mut sv" id="memoSv"></span></h3><textarea id="memoBox" class="memobox" rows="3" placeholder="메모를 입력하면 바로 저장됩니다">' + esc(c['메모'] || '') + '</textarea>';
    var consHtml = '<h3>상담기록 ' + cons.length + '건</h3><div class="cadd"><textarea id="cNew" rows="1" placeholder="상담 내용을 입력하세요"></textarea><button type="button" class="btn pri" id="cAdd">입력</button></div>' +
      (cons.length ? cons.slice(0, 20).map(function (h) {
        return '<div class="tl cons"><span class="cdate">' + esc(String(h.at).slice(0, 10)) + '</span><span class="ctext">' + esc(h.text) + (h.voice ? ' <span class="mut">· 음성기록</span>' : '') + '</span><button type="button" class="cx" data-cdel="' + esc(h.key) + '" aria-label="삭제" title="삭제">✕</button></div>';
      }).join('') + (cons.length > 20 ? '<p class="mut">최근 20건만 표시 (전체 ' + cons.length + '건)</p>' : '') : '<p class="mut">상담기록이 없습니다.</p>');
    var rowC = function (f, v) { return row(f, v); };
    var route = c['최초유입경로'] === '이관' ? '' : (c['최초유입경로'] || '');
    var birth = (c['생년월일'] || '') + (c['성별'] ? (c['생년월일'] ? ' (' + c['성별'] + ')' : c['성별']) : '');
    var basic = '<h3>기본 정보</h3><div class="fg">' + (E ? rowE('유입경로', routeSel(c)) + (c['유입DB명'] ? row('유입DB', c['유입DB명']) : '') : rowC('유입', routeText(c))) + rowC('생년월일', birth) +
      (E ? rowE('직업', inp(c, '직업')) : row('직업', [c['직업'], c['직장명']].filter(Boolean).join(' · '))) +
      row('소개자', c['소개자명']) + row('관심·니즈', c['관심상품·니즈']) + '</div>';
    var phoneView = c['연락처'] ? (c['통신사'] ? esc(c['통신사']) + ' ' : '') + esc(c['연락처']) : '';
    var pick = function (k) { return '<button type="button" class="pickbtn" data-pick="' + k + '" title="이 값을 미팅장소로 복사">선택</button>'; };
    var contact = '<h3>연락처 · 주소</h3><div class="fg">' + (E
      ? rowE('휴대전화', '<span class="efr">' + carrierSel(c) + inp(c, '연락처', '', 'tel') + '</span>') + rowE('자택', '<span class="efr">' + inp(c, '자택전화', '', 'tel') + pick('자택전화') + '</span>') + rowE('직장', '<span class="efr">' + inp(c, '직장전화', '', 'tel') + pick('직장전화') + '</span>') + rowE('이메일', inp(c, '이메일', '', 'email')) + rowE('미팅장소', inp(c, '미팅장소')) + rowE('주소', inp(c, '주소')) + rowE('계좌', inp(c, '계좌', '예: 신한 110-123-456789')) + rowE('차량번호', inp(c, '차량번호', '예: 123가4321')) + rowE('자동차보험갱신일', mdSel(c)) + '<input type="hidden" data-e="자동차보험갱신일" value="' + esc(c['자동차보험갱신일'] || '') + '">'
      : rowH('휴대전화', phoneView) + row('자택', c['자택전화']) + row('직장', c['직장전화']) + row('이메일', c['이메일']) + row('미팅장소', c['미팅장소']) + row('주소', c['주소']) + row('계좌', c['계좌']) + rowH('차량', [c['차량번호'] ? esc(c['차량번호']) : '', c['자동차보험갱신일'] ? esc(mdText(c['자동차보험갱신일'])) : ''].filter(Boolean).join(' | '))) + '</div>';
    return '<div class="drawer"><div class="dh"><div><div class="nmrow"><h2 class="' + gcls(c) + '">' + esc(c['이름']) + '</h2>' + (E ? gradeDots(c) : '') + '</div><span class="tag t3">' + esc(c['현재단계'] === '이관관리' ? '이관' : c['현재단계']) + '</span> <span class="mut">' + esc(c['담당FC']) + ' · ' + esc(c['고객ID']) + '</span></div><div class="dhb">' +
      (canE && !E ? '<button class="btn" id="editD">수정</button>' : '') + (E ? '<button class="btn pri" id="eSave">확인</button><button class="btn" id="eCancel">취소</button>' : '') + '<button class="btn" id="closeD">닫기</button></div></div>' +
      covanSection(c) + basic + contact +
      memoBox + consHtml +
      '<h3>계약 ' + cs.length + '건</h3>' + (cs.length ? '<div class="wrap"><table class="ct"><tr><th class="l">상품명</th><th>상태</th><th>보험료</th><th>계약일</th><th>모집유형</th><th>유입</th></tr>' + ct + '</table></div>' : '<p class="mut">등록된 계약이 없습니다.</p>') +
      '<h3>접촉 이력 (최근 20건)</h3>' + (hist || '<p class="mut">이력이 없습니다.</p>') + '</div>';
  }

  // 따행 표: 순번·이름·생일·연락처·직업·주소·모집FC + 방문완료·설문발송·설문응답·쿠폰발송(했으면 체크 표시만)
  var DONE_CONTACT = ['방문완료', '설문발송', '설문응답', '쿠폰발송'];
  function recFc(c) {
    var seen = [], list = contractsOf(c).filter(active).map(function (x) { return x.rfc; }).filter(function (n) { if (!n || seen.indexOf(n) >= 0) return false; seen.push(n); return true; });
    return (list.length ? list : [c['원모집FC']].filter(Boolean));   // 여러 명이면 한 줄씩
  }
  var ck = function (on, t) { return '<td class="ckc"><span class="ck ' + (on ? 'on' : '') + '" title="' + t + (on ? ' 함' : ' 안 함') + '">' + (on ? '✓' : '') + '</span></td>'; };
  function haengBody(view) {
    return view.map(function (c, vi) {
      var h = haengNow(c) || {};
      return '<tr class="pick-row ' + (state.sel === c['고객ID'] ? 'selrow' : '') + '" data-cid="' + esc(c['고객ID']) + '"><td>' + (state.page * PAGE + vi + 1) + '</td><td class="l nm"><b class="' + gcls(c) + '">' + esc(c['이름']) + '</b></td><td>' + cellContact(c) + '</td>' + cellPerson(c) + '<td class="mfc">' + (recFc(c).map(esc).join('<br>') || '-') + '</td>' +
        ck(!!(h.done || DONE_CONTACT.indexOf(h.contact) >= 0), '방문완료') + ck(!!h.survey, '설문발송') + ck(!!h.answer, '설문응답') + ck(!!h.coupon, '쿠폰발송') + '</tr>';
    }).join('');
  }
  function render() {
    var FCS = fcList();
    if (state.fc && FCS.indexOf(state.fc) < 0) state.fc = '';   // 범위가 바뀌어 없는 FC는 해제
    var rows = filtered(), pages = Math.max(1, Math.ceil(rows.length / PAGE)); if (state.page >= pages) state.page = pages - 1;
    var view = rows.slice(state.page * PAGE, state.page * PAGE + PAGE);
    var cnt = {}; custAll().forEach(function (c) { if (!state.fc || c['담당FC'] === state.fc) cnt[c['현재단계']] = (cnt[c['현재단계']] || 0) + 1; });
    var chips = '<button class="chip ' + (state.stage ? '' : 'on') + '" data-stage="">전체 ' + Object.keys(cnt).reduce(function (a, k) { return a + cnt[k]; }, 0) + '</button>' + CHIP_STAGES.map(function (p) { return '<button class="chip ' + (state.stage === p[0] ? 'on' : '') + '" data-stage="' + p[0] + '">' + p[1] + ' ' + (cnt[p[0]] || 0) + '</button>'; }).join('');
    var hc = { all: 0, done: 0, plan: 0, todo: 0 };
    custAll().forEach(function (c) { if (state.fc && c['담당FC'] !== state.fc) return; var x = hStatus(c); if (x) { hc.all++; hc[x]++; } });
    chips += '<span class="csep" aria-hidden="true">|</span>';
    chips += '<button class="chip haeng ' + (state.transfer ? 'on' : '') + '" data-transfer="1">이달이관</button>';
    chips += '<button class="chip haeng ' + (state.haeng ? 'on' : '') + '" data-haeng="1">따행</button>';
    chips += '<button class="chip haeng ' + (state.progress ? 'on' : '') + '" data-progress="1">진행</button>';
    if (state.haeng) {
      chips += '<span class="hsub">' + [['', '전체', hc.all], ['todo', '미완료', hc.todo], ['plan', '방문 약속', hc.plan], ['done', '방문 완료', hc.done]].map(function (x) { return '<button class="chip small ' + (state.hs === x[0] ? 'on' : '') + '" data-hs="' + x[0] + '">' + x[1] + ' ' + x[2] + '</button>'; }).join('') + '</span>';
    }
    var opt = function (arr, cur) { return arr.map(function (x) { return '<option value="' + esc(x) + '"' + (cur === x ? ' selected' : '') + '>' + esc(x) + '</option>'; }).join(''); };
    var body = view.map(function (c, vi) {
      var cs = contractsOf(c), act = cs.filter(active).length, h = haengNow(c);
      return '<tr class="pick-row ' + (state.sel === c['고객ID'] ? 'selrow' : '') + '" data-cid="' + esc(c['고객ID']) + '"><td>' + (state.page * PAGE + vi + 1) + '</td><td class="l nm"><b class="' + gcls(c) + '">' + esc(c['이름']) + '</b></td><td>' + esc(c['담당FC']) + '</td><td>' + esc(c['연락처'] || '-') + '</td><td>' + esc(c['최초유입경로'] || '-') + '</td><td class="l job" title="' + esc(c['직업'] || '') + '">' + esc(c['직업'] || '-') + '</td><td class="l">' + esc(c['지역'] || '-') + '</td><td>' + (cs.length ? act + '/' + cs.length : '-') + '</td></tr>';
    }).join('');
    if (state.haeng) body = haengBody(view);
    if (state.transfer) body = transferBody(view);
    var sel = state.sel ? (K.teamCustomers || K.customers).filter(function (c) { return c['고객ID'] === state.sel; })[0] : null;
    return '<h1>고객</h1>' +
      '<div class="chips" id="custChips">' + chips + '</div>' +
      '<div class="filters"><input id="q" type="search" autocomplete="off" placeholder="이름·연락처·주소·직업·고객ID 검색" value="' + esc(state.q) + '"><select id="fc"><option value="">팀</option>' + opt(FCS, state.fc) + '</select><select id="src"><option value="">유입경로 전체</option>' + opt(SRC, state.src) + '</select><span class="mut" id="custCount">' + rows.length.toLocaleString('ko-KR') + '명 조회</span></div>' +
      '<div id="custBody">' +
      '<section><div class="wrap tall"><table class="ct' + (state.haeng ? ' hg' : '') + (state.transfer ? ' hg tr' : '') + '">' + (state.transfer ? '<tr><th>순번</th><th class="l">이름</th><th>연락처<br>생일</th><th class="l">직업</th><th class="l">지역</th><th class="l">주소</th><th>모집FC</th><th class="l">상품명</th><th>보험료</th><th>상태</th><th>이관일</th></tr>' : state.haeng ? '<tr><th>순번</th><th class="l">이름</th><th>연락처<br>생일</th><th class="l">직업</th><th class="l">지역</th><th class="l">주소</th><th>모집FC</th><th>방문<br>완료</th><th>설문<br>발송</th><th>설문<br>응답</th><th>쿠폰<br>발송</th></tr>' : '<tr><th>순번</th><th class="l">이름</th><th>담당FC</th><th>연락처</th><th>유입</th><th class="l">직업</th><th class="l">지역</th><th>유지계약</th></tr>') + (body || '<tr><td colspan="' + (state.haeng || state.transfer ? 11 : 8) + '" class="mutc">조건에 맞는 고객이 없습니다.</td></tr>') + '</table></div>' +
      '<div class="pager"><span>전체 ' + rows.length.toLocaleString('ko-KR') + '명 표시</span></div></section>' + drawer(sel) + '</div>';
  }
  // 고객 항목 값을 저장 (이 브라우저에 반영 + 수정 내역 기록)
  // 알림 한 줄(저장 중/저장됨/실패)
  function dToast(msg, bad) {
    var t = document.getElementById('dtoast'); if (!t) { t = document.createElement('div'); t.id = 'dtoast'; document.body.appendChild(t); }
    t.className = bad ? 'bad' : ''; t.textContent = msg; t.style.display = 'block'; clearTimeout(t._h); t._h = setTimeout(function () { t.style.display = 'none'; }, bad ? 7000 : 2500);
  }
  // 메모·고객 정보 수정 내역 (접촉 이력에 표시). 계좌번호는 값 없이 '수정'만 남김
  var LOG_KEY = 'shfc.editLog';
  function readLog() { try { return JSON.parse(localStorage.getItem(LOG_KEY) || '[]'); } catch (e) { return []; } }
  function logChanges(c, diff, oldv) {
    var keys = Object.keys(diff); if (!keys.length) return;
    var text, type;
    if (keys.length === 1 && keys[0] === '메모') { type = '메모 수정'; text = diff['메모'] || '(비움)'; }
    else { type = '정보 수정'; text = keys.map(function (k) { return k === '계좌' ? '계좌 수정' : k + ': ' + (oldv[k] || '-') + ' → ' + (diff[k] || '-'); }).join(' / '); }
    var l = readLog(); l.push({ id: c['고객ID'], at: nowStamp(), type: type, text: text });
    try { localStorage.setItem(LOG_KEY, JSON.stringify(l.slice(-500))); } catch (e) { /* 무시 */ }
  }
  function applyVals(c, vals) {
    var diff0 = {}, old0 = {}; Object.keys(vals).forEach(function (k) { if ((c[k] || '') !== vals[k]) { diff0[k] = vals[k]; old0[k] = c[k] || ''; } });
    if (!Object.keys(diff0).length) return 0;
    logChanges(c, diff0, old0);
    if (window.DRIVE && DRIVE.on()) {   // 구글 Drive 연결됨: 엑셀에 바로 저장 (화면은 바로 바꿈, 실패하면 되돌림)
      var ch = {}, old = {}; Object.keys(vals).forEach(function (k) { if ((c[k] || '') !== vals[k]) { ch[k] = vals[k]; old[k] = c[k] || ''; } });
      var n0 = Object.keys(ch).length; if (!n0) return 0;
      Object.keys(ch).forEach(function (k) { c[k] = ch[k]; });
      dToast('엑셀에 저장하는 중…');
      DRIVE.updateCustomer(c['담당FC'], c['고객ID'], ch).then(function (name) { dToast('엑셀에 저장했습니다 (' + name + ')'); })
        .catch(function (e) { Object.keys(old).forEach(function (k) { c[k] = old[k]; }); dToast('엑셀 저장 실패: ' + (e.message || e), true); window.APP.rerender(); });
      return n0;
    }
    var ed = readEdits(), e = ed[c['고객ID']] || { name: c['이름'], fc: c['담당FC'], o: {}, v: {} }, n = 0;
    Object.keys(vals).forEach(function (k) {
      if ((c[k] || '') === vals[k]) return;
      if (!(k in e.o)) e.o[k] = c[k] || '';
      e.v[k] = vals[k]; c[k] = vals[k]; n++;
    });
    if (n) { e.at = nowStamp(); ed[c['고객ID']] = e; try { localStorage.setItem(EDIT_KEY, JSON.stringify(ed)); } catch (er) { alert('이 브라우저에 저장하지 못했습니다.'); } }
    return n;
  }
  // 상담기록 Drive 저장: 화면(메모리)은 바로 바꾸고 엑셀에 저장, 실패하면 되돌림
  function driveCons(c, op) {
    var list = X.consult, name = c['이름'], fc = c['담당FC'], undo;
    if (op.add) { var e = { at: op.add.at, name: name, fc: fc, text: op.add.text, voice: '' }; list.push(e); undo = function () { list.splice(list.indexOf(e), 1); }; }
    else {
      var src = op.edit || op.del, e2 = list.filter(function (r) { return r.fc === fc && String(r.name).replace(/\(\d{6}\)$/, '') === name && r.at === src.at && r.text === src.text; })[0];
      if (!e2) { dToast('이 상담기록을 찾지 못했습니다.', true); return; }
      if (op.edit) { e2.text = op.edit.newText; undo = function () { e2.text = src.text; }; } else { var i = list.indexOf(e2); list.splice(i, 1); undo = function () { list.splice(i, 0, e2); }; }
    }
    indexExtra(); dToast('엑셀에 저장하는 중…'); window.APP.rerender();
    DRIVE.consultOp(fc, name, op).then(function (n) { dToast('상담기록을 엑셀에 저장했습니다 (' + n + ')'); })
      .catch(function (er) { undo(); indexExtra(); dToast('엑셀 저장 실패: ' + (er.message || er), true); window.APP.rerender(); });
  }
  var curCust = function () { return (K.teamCustomers || K.customers).filter(function (x) { return x['고객ID'] === state.sel; })[0]; };
  function saveEdit() {
    var c = curCust(); if (!c) return;
    var vals = {};
    document.querySelectorAll('.drawer [data-e]').forEach(function (el) { vals[el.getAttribute('data-e')] = el.value.trim(); });
    var mm = document.querySelector('.drawer [data-emm]'), dd = document.querySelector('.drawer [data-edd]');
    if (mm && dd) { var m = +mm.value, d = +dd.value; vals['자동차보험갱신일'] = m && d ? (m < 10 ? '0' : '') + m + (d < 10 ? '0' : '') + d : ''; }
    applyVals(c, vals);
    document.querySelectorAll('.drawer [data-cr]').forEach(function (sel) {   // 계약 단위 유입경로(소개)·소개자
      var pol = sel.getAttribute('data-cr'), x = (K.contracts[c['고객ID']] || []).filter(function (y) { return y.pol === pol; })[0]; if (!x) return;
      var ii = document.querySelector('.drawer [data-ci="' + pol + '"]'), route = sel.value, intro = ii ? ii.value.trim() : '';
      if ((x.route || '') === route && (x.intro || '') === intro) return;
      var oldR = x.route || '', oldI = x.intro || ''; x.route = route; x.intro = intro; indexIns();
      var lg = readLog(); lg.push({ id: c['고객ID'], at: nowStamp(), type: '계약 유입 수정', text: '' }); try { localStorage.setItem(LOG_KEY, JSON.stringify(lg.slice(-500))); } catch (e) { /* 무시 */ }
      if (window.DRIVE && DRIVE.on()) {
        dToast('엑셀에 저장하는 중…');
        DRIVE.updateContract(x.fc || c['담당FC'], pol, { '유입경로': route, '소개자': intro }).then(function (n) { dToast('엑셀에 저장했습니다 (' + n + ')'); })
          .catch(function (e) { x.route = oldR; x.intro = oldI; indexIns(); dToast('엑셀 저장 실패: ' + (e.message || e), true); window.APP.rerender(); });
      } else { var ce = readCE(); ce[pol] = { route: route, intro: intro }; try { localStorage.setItem(CE_KEY, JSON.stringify(ce)); } catch (e) { alert('이 브라우저에 저장하지 못했습니다.'); } }
    });
    state.edit = false; window.APP.rerender();
  }
  function partial() {
    var tmp = document.createElement('div'); tmp.innerHTML = render();
    ['custChips', 'custCount', 'custBody'].forEach(function (id) { var a = document.getElementById(id), b = tmp.querySelector('#' + id); if (a && b) a.innerHTML = b.innerHTML; });
    bind();
    if (window.APP && APP.align) APP.align();
  }
  var filesAsked = {};
  function ensureFiles(cid) {   // 고객 폴더의 파일 목록은 상세를 열 때 Drive에서 읽는다
    if (!cid || filesAsked[cid] || !window.DRIVE || !DRIVE.customerFiles || !DRIVE.on()) return;
    filesAsked[cid] = 1;
    DRIVE.customerFiles(cid).then(function () { if (state.sel === cid) window.APP.rerender(); }).catch(function () { /* 파일 목록만 못 보임 */ });
  }
  function bind() {
    ensureFiles(state.sel);
    var g = function (id) { return document.getElementById(id); }, rr = function () { window.APP.rerender(); };
    // 검색어를 칠 때는 입력칸을 다시 만들지 않고 목록·칩·건수만 바꾼다 (한글 조합 중에 글자가 깨지지 않게)
    var q = g('q'); q.oninput = function () { state.q = q.value; state.page = 0; partial(); };
    ['fc', 'src'].forEach(function (id) { g(id).onchange = function () { state[id] = g(id).value; state.page = 0; rr(); }; });
    document.querySelectorAll('[data-transfer]').forEach(function (b) { b.onclick = function () { state.transfer = !state.transfer; if (state.transfer) { state.haeng = false; state.progress = false; state.hs = ''; } state.page = 0; rr(); }; });
    document.querySelectorAll('[data-progress]').forEach(function (b) { b.onclick = function () { state.progress = !state.progress; if (state.progress) { state.transfer = false; state.haeng = false; state.hs = ''; state.stage = ''; } state.page = 0; rr(); }; });
    document.querySelectorAll('[data-haeng]').forEach(function (b) { b.onclick = function () { state.haeng = !state.haeng; if (state.haeng) { state.transfer = false; state.progress = false; } if (!state.haeng) state.hs = ''; state.page = 0; rr(); }; });
    document.querySelectorAll('[data-hs]').forEach(function (b) { b.onclick = function () { state.hs = b.getAttribute('data-hs'); state.page = 0; rr(); }; });
    document.querySelectorAll('[data-stage]').forEach(function (b) { b.onclick = function () { state.stage = b.getAttribute('data-stage'); state.progress = false; state.page = 0; rr(); }; });
    document.querySelectorAll('[data-cid]').forEach(function (tr) { tr.onclick = function () { state.sel = tr.getAttribute('data-cid'); state.edit = false; state.cedit = ''; rr(); }; });
    var ed = g('editD'); if (ed) ed.onclick = function () { state.edit = true; rr(); };
    var ec = g('eCancel'); if (ec) ec.onclick = function () { state.edit = false; rr(); };
    var es = g('eSave'); if (es) es.onclick = saveEdit;
    document.querySelectorAll('.gdots [data-g]').forEach(function (b) { b.onclick = function () { var box = b.closest('.gdots'); box.querySelectorAll('[data-g]').forEach(function (x) { x.classList.remove('on'); }); b.classList.add('on'); var h = document.querySelector('.drawer input[data-e="성향"]'); if (h) h.value = b.getAttribute('data-g'); var nm = document.querySelector('.drawer h2'); if (nm) nm.className = 'g-' + b.getAttribute('data-g').toLowerCase(); }; });
    var mb = g('memoBox');   // 메모는 수정 모드가 아니어도 항상 수정 가능, 입력칸을 벗어나면 바로 저장
    if (mb) mb.onblur = function () { var c = curCust(); if (!c) return; var n = applyVals(c, { '메모': mb.value.trim() }); var sv = g('memoSv'); if (sv) { sv.textContent = n ? '저장됨' : ''; setTimeout(function () { if (g('memoSv')) g('memoSv').textContent = ''; }, 1500); } };
    var ca = g('cAdd'); if (ca) ca.onclick = function () {
      var c = curCust(), t = (g('cNew').value || '').trim(); if (!c || !t) { g('cNew').focus(); return; }
      if (window.DRIVE && DRIVE.on()) return driveCons(c, { add: { at: nowStamp(), text: t } });
      consChange(c, function (ov) { ov.add.push({ id: 'a' + Date.now().toString(36), at: nowStamp(), text: t }); }); rr();
    };
    document.querySelectorAll('[data-pick]').forEach(function (b) { b.onclick = function () { var v = document.querySelector('.drawer [data-e="' + b.getAttribute('data-pick') + '"]'), m = document.querySelector('.drawer [data-e="미팅장소"]'); if (v && m) m.value = v.value; }; });
    document.querySelectorAll('tr[data-pdf]').forEach(function (tr) { tr.onclick = function () { window.open(tr.getAttribute('data-pdf'), '_blank', 'noopener'); }; });
    var cn = g('cNew'); if (cn) cn.oninput = function () { cn.style.height = 'auto'; cn.style.height = cn.scrollHeight + 'px'; };
    if (cn) cn.onkeydown = function (ev) { if (ev.key === 'Enter' && (ev.ctrlKey || ev.metaKey)) { ev.preventDefault(); ca.click(); } };
    document.querySelectorAll('[data-cedit]').forEach(function (b) { b.onclick = function () { state.cedit = b.getAttribute('data-cedit'); rr(); }; });
    document.querySelectorAll('[data-ccancel]').forEach(function (b) { b.onclick = function () { state.cedit = ''; rr(); }; });
    document.querySelectorAll('[data-csave]').forEach(function (b) { b.onclick = function () {
      var c = curCust(), k = b.getAttribute('data-csave'), t = (g('cEditBox').value || '').trim(); if (!c) return;
      if (k.indexOf('x|') === 0 && window.DRIVE && DRIVE.on()) { var p = k.split('|'); state.cedit = ''; return driveCons(c, { edit: { at: p[1], text: p.slice(2).join('|'), newText: t } }); }
      consChange(c, function (ov) { if (k.indexOf('a|') === 0) { ov.add.forEach(function (a) { if ('a|' + a.id === k) a.text = t; }); } else ov.mod[k] = t; });
      state.cedit = ''; rr();
    }; });
    document.querySelectorAll('[data-cdel]').forEach(function (b) { b.onclick = function () {
      var c = curCust(), k = b.getAttribute('data-cdel'); if (!c) return;
      var dl0 = consultsOf(c).filter(function (r) { return r.key === k; })[0]; if (dl0) { var lg = readLog(); lg.push({ id: c['고객ID'], at: nowStamp(), type: '상담기록 삭제', text: dl0.text }); try { localStorage.setItem(LOG_KEY, JSON.stringify(lg.slice(-500))); } catch (e) { /* 무시 */ } }
      if (k.indexOf('x|') === 0 && window.DRIVE && DRIVE.on()) { var p = k.split('|'); return driveCons(c, { del: { at: p[1], text: p.slice(2).join('|') } }); }
      consChange(c, function (ov) { if (k.indexOf('a|') === 0) ov.add = ov.add.filter(function (a) { return 'a|' + a.id !== k; }); else { if (ov.del.indexOf(k) < 0) ov.del.push(k); delete ov.mod[k]; } });
      rr();
    }; });
    var dl = g('dlEdits'); if (dl) dl.onclick = function () { var ed2 = readEdits(), cons2 = readCons(), out = Object.keys(ed2).map(function (id) { var e = ed2[id], ch = {}; Object.keys(e.v || {}).forEach(function (k) { ch[k] = { 전: (e.o || {})[k] || '', 후: e.v[k] }; }); return { 고객ID: id, 이름: e.name, 담당FC: e.fc, 수정일: e.at, 변경: ch }; }); var a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([JSON.stringify({ 고객정보수정: out, 상담기록변경: cons2 }, null, 1)], { type: 'application/json' })); a.download = '고객수정내역.json'; a.click(); };
    
    var cd = g('closeD'); if (cd) cd.onclick = function () { state.sel = null; state.edit = false; state.cedit = ''; rr(); };
    var rq = g('reqCovan');
    if (rq) rq.onclick = function () {
      var c = K.customers.filter(function (x) { return x['고객ID'] === state.sel; })[0]; if (!c) return;
      var list = reqList(); list.push({ id: c['고객ID'], name: c['이름'], fc: c['담당FC'], at: todayStr });
      try { localStorage.setItem(REQ_KEY, JSON.stringify(list)); } catch (e) {}
      var blob = new Blob([JSON.stringify(list, null, 1)], { type: 'application/json' });
      var a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = '보장분석요청.json'; a.click();
      rr();
    };
  }
  window.PAGES = window.PAGES || {};
  window.PAGES.customer = { render: render, bind: bind, reloadExtra: function () { indexExtra(); applyExtraBasic(); indexIns(); } };
})();
