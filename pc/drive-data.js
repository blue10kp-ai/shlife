(function () {
  'use strict';
  // Drive 직접 읽기: 정적 사본(data\*.js) 없이 화면이 쓰는 데이터를 Drive의 원본 파일에서 만든다.
  //   ACCESS_CFG  ← Claude\_양식\사용자권한.json
  //   CUSTOMERS   ← Claude\고객\<FC>\날짜 - FC - 리스트.xlsx 의 고객·계약·이력·따행 시트
  //   CUSTEXTRA   ← Claude\고객\날짜 - 고객추가정보.xlsx 의 기본·상담기록 시트
  //   FEES        ← Claude\수수료\수수료명세서.json
  //   PERF        ← Claude\업적\performance.json   (없으면 빈 화면; 데이터수집 세션이 만들어 줌)
  //   CONTRACTS   ← Claude\계약\contracts.json     (없으면 빈 화면)
  //   CUSTFILES   ← 고객 상세를 열 때 그 고객 폴더만 읽는다 (DRIVE.customerFiles)
  // DRIVE(drive.js)가 먼저 불러와져 있어야 한다.
  var D = window.DRIVE, I = D._i;
  var FOLDER = 'application/vnd.google-apps.folder';
  var CUST_KEEP = ['고객ID', '이름', '담당FC', '현재단계', '최초유입경로', '유입일', '원모집FC', '고객분류', '고객등급', '생년월일', '실제생일', '성별', '연락처',
    '자택전화', '직장전화', '이메일', '주소', '지역', '직업', '직장명', '미팅장소', '성향', '메모', '관심상품·니즈', '소개자명',
    '가입설계동의 만료일', '마케팅동의 만료일', '개인정보 수집·이용 동의일', '계약안내 수신', '상령일', '최근 보장분석', '최근접촉일', '다음접촉일',
    '유입DB명', '소개고객ID', '통신사', '차량번호', '자동차보험갱신일', '계약현황', '신한라이프 유지계약', '최근문자일시', '최근통화일시'];   // 계좌·후번호는 화면용 목록에 넣지 않는다
  var p2 = function (n) { return n < 10 ? '0' + n : '' + n; };
  var s = function (v) { return v == null ? '' : String(v).trim(); };

  // 셀 → 값: 날짜 서식이면 YYYY-MM-DD, 숫자는 숫자, 나머지는 글자
  function cellVal(c) {
    if (!c || c.v == null || c.v === '') return '';
    if (c.t === 'd' && c.v instanceof Date) return c.v.getUTCFullYear() + '-' + p2(c.v.getUTCMonth() + 1) + '-' + p2(c.v.getUTCDate());
    if (c.t === 'n') {
      if (c.z && XLSX.SSF.is_date(c.z)) { var d = XLSX.SSF.parse_date_code(c.v); if (d) return d.y + '-' + p2(d.m) + '-' + p2(d.d); }
      return c.v;
    }
    return c.v;
  }
  function rowsOf(ws) {   // 첫 줄이 머리글인 시트 → [{머리글: 값}]
    if (!ws || !ws['!ref']) return [];
    var rg = XLSX.utils.decode_range(ws['!ref']), hdr = [], out = [];
    for (var c = rg.s.c; c <= rg.e.c; c++) { var h = ws[XLSX.utils.encode_cell({ r: rg.s.r, c: c })]; hdr[c] = h ? s(h.v) : ''; }
    for (var r = rg.s.r + 1; r <= rg.e.r; r++) {
      var o = {}, any = false;
      for (var c2 = rg.s.c; c2 <= rg.e.c; c2++) {
        if (!hdr[c2]) continue;
        var v = cellVal(ws[XLSX.utils.encode_cell({ r: r, c: c2 })]);
        if (v !== '') any = true;
        o[hdr[c2]] = v;
      }
      if (any) out.push(o);
    }
    return out;
  }
  var nz = function (v) { return v === '' ? null : v; };
  var d10 = function (v) { return typeof v === 'string' && /^\d{4}-\d\d-\d\d \d\d:/.test(v) ? v.slice(0, 10) : v; };   // 엑셀 때와 같게 날짜만

  // JSON 원본(FC별 data.json, 고객추가정보.json). 주소 끝에 ?json=test 를 붙이면 시험 폴더(_변환시험6-11-11)에서 읽는다.
  var TEST = /[?&]json=test/.test(location.search);
  function sheetObjs(doc, name) {   // {columns, rows:[[…]]} → [{열: 값}]
    var sh = doc && doc.sheets && doc.sheets[name]; if (!sh) return [];
    return sh.rows.map(function (r) { var o = {}; sh.columns.forEach(function (c, i) { if (c) o[c] = r[i] == null ? '' : r[i]; }); return o; });
  }
  async function testRoot() {
    if (I.cache.testRoot) return I.cache.testRoot;
    var a = await child(await claudeId(), '_변환시험', true), b = a && await child(a.id, '고객-시험', true);
    if (!b) throw new Error("시험 복사본 폴더 '_변환시험/고객-시험' 을 찾지 못했습니다.");
    return (I.cache.testRoot = b.id);
  }
  async function jsonOf(fcName, fcFolderId) {   // 없으면 null (엑셀로 대신 읽음)
    var parent = fcFolderId;
    if (TEST) { var tf = await child(await testRoot(), fcName, true); if (!tf) return null; parent = tf.id; }
    var f = await child(parent, 'data.json', false);
    return f ? I.gd(I.API + '/files/' + encodeURIComponent(f.id) + '?alt=media&supportsAllDrives=true') : null;
  }
  async function extraJson() {
    var parent = TEST ? await testRoot() : await I.baseId();
    var f = await child(parent, '고객추가정보.json', false);
    return f ? I.gd(I.API + '/files/' + encodeURIComponent(f.id) + '?alt=media&supportsAllDrives=true') : null;
  }

  async function claudeId() {
    if (I.cache.claude) return I.cache.claude;
    var b = await I.baseId();
    var f = await I.gd(I.API + '/files/' + encodeURIComponent(b) + '?fields=parents&supportsAllDrives=true');
    if (!f.parents || !f.parents[0]) throw new Error("'Claude' 폴더를 찾지 못했습니다.");
    return (I.cache.claude = f.parents[0]);
  }
  async function child(parent, name, folderOnly) {
    var a = await I.qlist("'" + I.qv(parent) + "' in parents and name='" + I.qv(name) + "' and trashed=false" + (folderOnly ? " and mimeType='" + FOLDER + "'" : ''));
    return a[0] || null;
  }
  async function readJson(path) {   // Claude 폴더 아래 경로 ['수수료','수수료명세서.json']
    var cur = await claudeId();
    for (var i = 0; i < path.length; i++) {
      var f = await child(cur, path[i], i < path.length - 1);
      if (!f) throw new Error("Drive에서 '" + path.join('\\') + "' 을(를) 찾지 못했습니다. (접근 권한이 없을 수도 있습니다)");
      cur = f.id;
    }
    return I.gd(I.API + '/files/' + encodeURIComponent(cur) + '?alt=media&supportsAllDrives=true');
  }

  async function loadCustomers(progress) {
    var fcs = await I.fcFolders(), out = { customers: [], contracts: {}, history: {}, haeng: {}, files: {} }, errs = [], folders = {};
    var results = await Promise.all(fcs.slice().sort(function (a, b) { return a.name < b.name ? -1 : 1; }).map(async function (fo) {
      try {
        var js = await jsonOf(fo.name, fo.id), f = null, wb = null, doc = null;
        if (js) { doc = js; f = { name: js.source + ' (JSON)' }; }
        else { f = await I.listFile(fo.name); wb = await I.getWB(f); }
        // 고객 폴더 지도: 고객ID → 폴더 ID (파일 목록은 상세를 열 때 읽는다)
        var map = {};
        try {
          var cf = await child(fo.id, '_고객', true);
          if (cf) (await I.qlist("'" + I.qv(cf.id) + "' in parents and mimeType='" + FOLDER + "' and trashed=false")).forEach(function (x) { var m = x.name.match(/ - (C\d{5,})\s*$/); if (m) map[m[1]] = x.id; });
        } catch (e) { /* 파일 목록만 못 보임 */ }
        return { fc: fo.name, file: f.name, wb: wb, doc: doc, map: map };
      } catch (e) { if (!/리스트\.xlsx가 없습니다/.test(e.message || '')) errs.push(fo.name + ': ' + e.message); return null; }   // 리스트가 없는 폴더(따행리스트·이관리스트 같은 내보내기 폴더)는 건너뜀
    }));
    results.filter(Boolean).forEach(function (res) {
      var wb = res.wb, rowsOf2 = function (name) { return res.doc ? sheetObjs(res.doc, name) : rowsOf(wb.Sheets[name]); };
      out.files[res.fc] = res.file;
      Object.assign(folders, res.map);
      rowsOf2('고객').forEach(function (r) {
        var c = {};
        CUST_KEEP.forEach(function (k) { c[k] = r[k] == null ? '' : d10(r[k]); });
        c.sol = s(r['SOL고객번호']);
        out.customers.push(c);
      });
      rowsOf2('계약').forEach(function (r) {
        var cid = r['계약자ID'];
        (out.contracts[cid] = out.contracts[cid] || []).push({
          pol: s(r['증권번호']), goods: nz(r['상품명']), type: nz(r['상품구분']), sum: nz(r['합계보험료']),
          start: d10(r['최초계약일']), end: d10(r['만기일']), renew: d10(r['갱신일']), st: nz(r['계약상태']),
          recruit: nz(r['모집유형']), rfc: nz(r['모집FC']), fc: nz(r['담당FC']), insured: nz(r['주피보험자ID']), insured2: nz(r['종피보험자ID']), trans: d10(r['이관일자']),
          route: nz(r['유입경로']), intro: nz(r['소개자'])
        });
      });
      rowsOf2('이력').forEach(function (r) {
        var cid = r['고객ID'];
        (out.history[cid] = out.history[cid] || []).push({ at: d10(r['일시']), type: nz(r['유형']), res: nz(r['결과']), memo: nz(r['내용']), next: d10(r['다음접촉일']) });
      });
      rowsOf2('따행').forEach(function (r) {
        var cid = r['고객ID'];
        (out.haeng[cid] = out.haeng[cid] || []).push({
          ym: nz(r['대상년월']), cat: nz(r['따행분류']), contact: nz(r['접촉현황']), target: nz(r['캠페인대상여부']),
          from: d10(r['활동기한시작일']), to: d10(r['활동기한종료일']), plan: d10(r['방문약속일']), done: d10(r['방문완료일']),
          survey: d10(r['설문발송일']), answer: d10(r['설문응답일']), coupon: d10(r['쿠폰발송일']), gift: nz(r['선물']), handed: nz(r['전달여부'])
        });
      });
    });
    var t = new Date(); out.made = t.getFullYear() + '-' + p2(t.getMonth() + 1) + '-' + p2(t.getDate());
    window.CUSTOMERS = out;
    I.cache.folders = folders;
    window.CUSTFILES = { made: out.made, byId: {} };
    return errs;
  }

  async function loadExtra() {
    var ej = await extraJson(), ef = ej ? null : await I.extraFile();
    var X = window.CUSTEXTRA = { file: ej ? ej.source + ' (JSON)' : (ef ? ef.name : ''), basic: [], consult: [] };
    if (!ej && !ef) return;
    var wb = ej ? null : await I.getWB(ef), objs = function (n) { return ej ? sheetObjs(ej, n) : rowsOf(wb.Sheets[n]); };
    X.basic = objs('기본').map(function (r) {
      return { fc: s(r['담당 FC']), name: s(r['고객명']), meeting: s(r['미팅장소']), memo: s(r['메모']), carrier: s(r['통신사']), car: s(r['차량번호']), renew: s(r['자동차보험갱신일']) };
    });
    X.consult = objs('상담기록').filter(function (r) { return s(r['내용']); }).map(function (r) {
      var at = r['저장날짜시간'] != null && r['저장날짜시간'] !== '' ? r['저장날짜시간'] : r['저장날짜'];
      return { at: s(at).slice(0, 16), name: s(r['고객명']), fc: s(r['담당 FC']), text: s(r['내용']), voice: s(r['음성파일']) };
    });
  }

  async function loadFees() {
    var d = await readJson(['수수료', '수수료명세서.json']), rows = [];
    d.forEach(function (r) {
      var pay = {};
      (r.a || []).concat(r.b || []).forEach(function (kv) { if (typeof kv[1] === 'number') pay[kv[0]] = kv[1]; });
      rows.push({
        no: r.no, name: r.fc, ym: r.ym,
        tot: { '지급계': (r.tot || {})['지급계'], '공제계': (r.tot || {})['공제계'], '실지급금': (r.tot || {})['실지급금'] },
        pay: pay, info: r.info || {}, sum: r.summary, a: r.a, b: r.b, d: r.d, chk: r.check || ''
      });
    });
    rows.sort(function (x, y) { return x.ym < y.ym ? -1 : x.ym > y.ym ? 1 : x.no < y.no ? -1 : x.no > y.no ? 1 : 0; });
    var t = new Date();
    window.FEES = { made: t.getFullYear() + '-' + p2(t.getMonth() + 1) + '-' + p2(t.getDate()), rows: rows };
  }

  // 고객 한 명의 폴더 파일 (스마트보장분석·손해담보·상세보장·고객카드). 상세를 열 때 호출
  var kind = function (n) {
    if (n.indexOf('스마트보장분석') >= 0) return '스마트보장분석';
    if (n.indexOf('손해담보') >= 0) return '손해담보';
    if (/\.docx$/i.test(n)) return '고객카드';
    if (/\.pdf$/i.test(n)) return '상세보장';
    return '기타';
  };
  async function customerFiles(cid) {
    var C = window.CUSTFILES; if (!C) return [];
    if (C.byId[cid]) return C.byId[cid];
    var fid = (I.cache.folders || {})[cid]; if (!fid) return (C.byId[cid] = []);
    var out = [];
    async function walk(id, depth) {
      var kids = await I.qlist("'" + I.qv(id) + "' in parents and trashed=false");
      for (var i = 0; i < kids.length; i++) {
        var k = kids[i];
        if (k.mimeType === FOLDER) { if (depth < 2) await walk(k.id, depth + 1); continue; }
        if (k.name.indexOf('~$') === 0 || k.name.charAt(0) === '.') continue;
        out.push({ n: k.name, id: k.id, link: 'https://drive.google.com/file/d/' + k.id + '/view', k: kind(k.name), m: (k.modifiedTime || '').slice(0, 10) });
      }
    }
    await walk(fid, 0);
    out.sort(function (a, b) { return a.k < b.k ? -1 : a.k > b.k ? 1 : a.n < b.n ? -1 : 1; });
    return (C.byId[cid] = out);
  }

  // 한꺼번에 읽기. progress(문구)로 진행 상황을 알린다. 못 읽은 것은 errors에 모은다.
  async function loadAll(progress) {
    var say = progress || function () {}, errors = [];
    say('권한 정보 읽는 중…');
    window.ACCESS_CFG = await readJson(['_양식', '사용자권한.json']);
    var me = String(D.email() || '').toLowerCase();
    var ok = (window.ACCESS_CFG.users || []).some(function (u) { return String(u.email || '').toLowerCase() === me; });
    if (!ok) throw new Error('이 구글 계정(' + me + ')은 신한FC 접속 권한 목록에 없습니다. 관리자에게 요청하세요.');
    say('고객·계약·이력 읽는 중…');
    var steps = [
      loadCustomers().then(function (e) { errors = errors.concat(e); }),
      loadExtra().catch(function (e) { errors.push('고객추가정보: ' + e.message); }),
      readJson(['업적', 'performance.json']).then(function (d) { window.PERF = d; }).catch(function (e) { console.warn('업적 자료를 읽지 못했습니다(화면은 비어 있음):', e.message); }),
      readJson(['계약', 'contracts.json']).then(function (d) { window.CONTRACTS = d; }).catch(function (e) { console.warn('계약 자료를 읽지 못했습니다(화면은 비어 있음):', e.message); }),
      loadFees().catch(function (e) { errors.push('수수료: ' + e.message); window.FEES = { made: '', rows: [] }; })
    ];
    await Promise.all(steps);
    return errors;
  }

  // 접속 명단 저장(_양식\사용자권한.json): 저장 직전에 최신을 읽고 fn으로 바꾼 뒤 올림(다른 항목 보존)
  function updateRoster(fn) {
    return I.serialize(async function () {
      var cur = await claudeId(), f = await child(cur, '_양식', true);
      f = f && await child(f.id, '사용자권한.json', false);
      if (!f) throw new Error("'_양식\\사용자권한.json' 을(를) 찾지 못했습니다.");
      var data = (await I.gd(I.API + '/files/' + encodeURIComponent(f.id) + '?alt=media&supportsAllDrives=true')) || {};
      data.users = data.users || [];
      var next = fn(data) || data;
      await I.gd('https://www.googleapis.com/upload/drive/v3/files/' + encodeURIComponent(f.id) + '?uploadType=media&supportsAllDrives=true', { method: 'PATCH', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(next, null, 2) });
      return next;
    });
  }

  D.updateRoster = updateRoster;
  D.loadAll = loadAll;
  D.customerFiles = customerFiles;
  D.readJson = readJson;
})();
