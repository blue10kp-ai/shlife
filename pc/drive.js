(function () {
  'use strict';
  // 구글 로그인 + Google Drive 엑셀 읽기·쓰기 (신한FC M과 같은 파일·같은 구조)
  //   고객 정보: 고객\<FC>\날짜 - FC - 리스트.xlsx 의 '고객' 시트 (고객ID 줄)
  //   상담기록: 고객\날짜 - 고객추가정보.xlsx 의 '상담기록' 시트
  // 로그인은 사용자가 직접 한다. 토큰은 이 탭(sessionStorage)에만 보관한다.
  var CID = '482206908404-5sdbvetcrc8j7ctrhe6aa23ccmd3d05d.apps.googleusercontent.com';
  var SCOPE = 'https://www.googleapis.com/auth/drive openid email';
  var API = 'https://www.googleapis.com/drive/v3', UP = 'https://www.googleapis.com/upload/drive/v3';
  var KEY = 'shfcd.gtoken', MAIL = 'shfcd.gmail';
  var FOLDER = 'application/vnd.google-apps.folder', XMIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  var listeners = [], cache = {}, queue = Promise.resolve();
  var EDITABLE = ['최초유입경로', '소개자명', '유입DB명', '연락처', '자택전화', '직장전화', '이메일', '주소', '직업', '성향', '통신사', '차량번호', '자동차보험갱신일', '계좌', '미팅장소', '메모'];

  function stored() { try { var t = JSON.parse(sessionStorage.getItem(KEY) || 'null'); return t && t.exp > Date.now() + 60000 ? t : null; } catch (e) { return null; } }
  function token() { var t = stored(); return t ? t.token : ''; }
  function on() { return !!token(); }
  function email() { try { return sessionStorage.getItem(MAIL) || ''; } catch (e) { return ''; } }
  function emit() { listeners.forEach(function (f) { try { f(); } catch (e) { /* 무시 */ } }); }
  function onChange(f) { listeners.push(f); }

  function login() {
    return new Promise(function (resolve, reject) {
      var g = window.google && window.google.accounts && window.google.accounts.oauth2;
      if (!g) return reject(new Error('구글 로그인 모듈을 불러오지 못했습니다. 인터넷 연결을 확인하세요.'));
      var c = g.initTokenClient({
        client_id: CID, scope: SCOPE,
        callback: function (r) {
          if (r.error) return reject(new Error('구글 연결 실패: ' + r.error));
          if (!g.hasGrantedAllScopes(r, 'https://www.googleapis.com/auth/drive')) return reject(new Error('Google Drive 권한을 허용해야 합니다. 다시 로그인하면서 모든 항목에 체크해 주세요.'));
          try { sessionStorage.setItem(KEY, JSON.stringify({ token: r.access_token, exp: Date.now() + (r.expires_in || 3600) * 1000 })); } catch (e) { /* 무시 */ }
          cache = {};
          fetch('https://www.googleapis.com/oauth2/v3/userinfo', { headers: { Authorization: 'Bearer ' + r.access_token } }).then(function (x) { return x.json(); })
            .then(function (u) { try { sessionStorage.setItem(MAIL, u.email || ''); } catch (e) { /* 무시 */ } }).catch(function () { /* 무시 */ })
            .then(function () { emit(); resolve(); });
        },
        error_callback: function (e) { reject(new Error(e && e.type === 'popup_closed' ? '구글 로그인 창이 닫혔습니다.' : '구글 로그인 창을 열 수 없습니다. 팝업 차단을 확인하세요.')); }
      });
      c.requestAccessToken({ prompt: '' });
    });
  }
  function logout() {
    var t = token();
    try { sessionStorage.removeItem(KEY); sessionStorage.removeItem(MAIL); } catch (e) { /* 무시 */ }
    cache = {};
    try { if (t && window.google) window.google.accounts.oauth2.revoke(t, function () {}); } catch (e) { /* 무시 */ }
    emit();
  }

  async function gd(url, opt, buf) {
    var t = token();
    if (!t) throw new Error('GOOGLE_AUTH');
    var o = opt || {};
    var res = await fetch(url, Object.assign({}, o, { headers: Object.assign({ Authorization: 'Bearer ' + t }, o.headers || {}), cache: 'no-store' }));
    if (res.status === 401) { try { sessionStorage.removeItem(KEY); } catch (e) { /* 무시 */ } emit(); throw new Error('GOOGLE_AUTH'); }
    if (!res.ok) throw new Error('Drive ' + res.status + ': ' + (await res.text()).slice(0, 200));
    if (res.status === 204) return null;
    return buf ? res.arrayBuffer() : res.json();
  }
  var qv = function (s) { return String(s).replace(/\\/g, '\\\\').replace(/'/g, "\\'"); };
  async function qlist(q) {
    var out = [], page = '';
    do {
      var d = await gd(API + '/files?pageSize=1000&supportsAllDrives=true&includeItemsFromAllDrives=true&fields=' + encodeURIComponent('nextPageToken,files(id,name,mimeType,modifiedTime)') + '&q=' + encodeURIComponent(q) + (page ? '&pageToken=' + encodeURIComponent(page) : ''));
      out = out.concat(d.files || []); page = d.nextPageToken || '';
    } while (page);
    return out;
  }
  async function baseId() {
    if (cache.base) return cache.base;
    var a = await qlist("name='Claude' and 'root' in parents and mimeType='" + FOLDER + "' and trashed=false");
    if (!a[0]) throw new Error("내 드라이브에 'Claude' 폴더가 없습니다. 고객 폴더 주인 계정으로 로그인해 주세요.");
    var b = await qlist("name='고객' and '" + a[0].id + "' in parents and mimeType='" + FOLDER + "' and trashed=false");
    if (!b[0]) throw new Error("'Claude/고객' 폴더를 찾지 못했습니다.");
    return (cache.base = b[0].id);
  }
  async function fcFolders() {
    if (cache.fcs) return cache.fcs;
    var kids = await qlist("'" + qv(await baseId()) + "' in parents and mimeType='" + FOLDER + "' and trashed=false");
    return (cache.fcs = kids.filter(function (k) { return k.name.charAt(0) !== '_'; }));
  }
  async function listFile(fc) {   // FC 폴더의 가장 최근(날짜 접두사) '리스트.xlsx'
    var key = 'list|' + fc;
    if (cache[key]) return cache[key];
    var fo = (await fcFolders()).filter(function (f) { return f.name === fc; })[0];
    if (!fo) throw new Error("고객 폴더에 '" + fc + "' 폴더가 없습니다.");
    var files = (await qlist("'" + fo.id + "' in parents and trashed=false")).filter(function (f) { return /리스트\.xlsx$/.test(f.name) && f.name.indexOf('(구)') < 0 && f.name.indexOf('~$') !== 0; }).sort(function (a, b) { return a.name.localeCompare(b.name); });
    if (!files.length) throw new Error("'" + fc + "' 폴더에 리스트.xlsx가 없습니다.");
    return (cache[key] = files[files.length - 1]);
  }
  async function extraFile() {
    if (cache.extra !== undefined) return cache.extra;
    var files = (await qlist("'" + qv(await baseId()) + "' in parents and trashed=false")).filter(function (f) { return /\.xlsx$/i.test(f.name) && f.name.replace(/\s+/g, '').indexOf('고객추가정보') >= 0; }).sort(function (a, b) { return a.name.localeCompare(b.name); });
    return (cache.extra = files[files.length - 1] || null);
  }
  async function getWB(f) { return XLSX.read(await gd(API + '/files/' + encodeURIComponent(f.id) + '?alt=media&supportsAllDrives=true', null, true), { type: 'array', cellNF: true }); }
  async function putWB(f, wb) {
    var out = XLSX.write(wb, { type: 'array', bookType: 'xlsx', compression: true });
    await gd(UP + '/files/' + encodeURIComponent(f.id) + '?uploadType=media&supportsAllDrives=true', { method: 'PATCH', headers: { 'Content-Type': XMIME }, body: out });
  }
  function info(ws, key) {   // 머리글 줄(첫 10줄 안에서 key가 있는 줄)
    var rg = XLSX.utils.decode_range(ws['!ref'] || 'A1'), enc = function (r, c) { return XLSX.utils.encode_cell({ r: r, c: c }); };
    var txt = function (r, c) { var x = ws[enc(r, c)]; return x ? String(x.w != null ? x.w : x.v).trim() : ''; };
    var hr = -1;
    for (var r = rg.s.r; r <= Math.min(rg.e.r, rg.s.r + 10) && hr < 0; r++) for (var c = rg.s.c; c <= rg.e.c; c++) if (txt(r, c) === key) { hr = r; break; }
    if (hr < 0) throw new Error("시트에서 '" + key + "' 열을 찾지 못했습니다.");
    var colOf = function (n) { for (var c = rg.s.c; c <= rg.e.c; c++) if (txt(hr, c) === n) return c; return -1; };
    return { ws: ws, rg: rg, enc: enc, txt: txt, hr: hr, colOf: colOf };
  }
  function serial(d) { return (Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes()) - Date.UTC(1899, 11, 30)) / 86400000; }
  function dateText(cell) {   // 셀(일련번호 또는 글자) → 'YYYY-MM-DD HH:MM'
    if (!cell) return '';
    if (cell.t === 'n' && cell.v > 1) { var d = new Date(Math.round(cell.v * 1440) * 60000 + Date.UTC(1899, 11, 30)); var p = function (n) { return n < 10 ? '0' + n : '' + n; }; return d.getUTCFullYear() + '-' + p(d.getUTCMonth() + 1) + '-' + p(d.getUTCDate()) + ' ' + p(d.getUTCHours()) + ':' + p(d.getUTCMinutes()); }
    return String(cell.w != null ? cell.w : cell.v).trim().slice(0, 16);
  }
  function serialize(fn) { var run = queue.then(fn, fn); queue = run.catch(function () { /* 다음 작업은 계속 */ }); return run; }

  // 고객 항목 저장: values = { 열이름: 값 } (빈 값은 칸을 비움)
  function updateCustomer(fc, id, values) { return updateRow(fc, '고객', '고객ID', id, values); }
  function updateContract(fc, pol, values) { return updateRow(fc, '계약', '증권번호', pol, values); }
  function updateRow(fc, sheet, keyCol, key, values) {
    return serialize(async function () {
      var f = await listFile(fc), wb = await getWB(f), ws = wb.Sheets[sheet];
      if (!ws) throw new Error("'" + f.name + "'에 '" + sheet + "' 시트가 없습니다.");
      var I = info(ws, keyCol), idc = I.colOf(keyCol), row = -1;
      for (var r = I.hr + 1; r <= I.rg.e.r; r++) if (I.txt(r, idc) === String(key)) { row = r; break; }
      if (row < 0) throw new Error(sheet + ' 시트에서 ' + key + ' 줄을 찾지 못했습니다.');
      Object.keys(values).forEach(function (k) {
        var c = I.colOf(k);
        if (c < 0) { c = I.rg.e.c + 1; I.rg.e.c = c; ws[I.enc(I.hr, c)] = { t: 's', v: k }; }
        var v = values[k];
        if (v === '' || v == null) delete ws[I.enc(row, c)]; else ws[I.enc(row, c)] = { t: 's', v: String(v) };
      });
      ws['!ref'] = XLSX.utils.encode_range(I.rg);
      await putWB(f, wb);
      return f.name;
    });
  }

  // 상담기록: op = { add: {at, text} } | { edit: {at, text, newText} } | { del: {at, text} }   (at = 'YYYY-MM-DD HH:MM')
  function consultOp(fc, name, op) {
    return serialize(async function () {
      var f = await extraFile();
      if (!f) throw new Error("고객 폴더에 '고객추가정보' 엑셀이 없습니다.");
      var wb = await getWB(f), ws = wb.Sheets['상담기록'];
      if (!ws) { ws = XLSX.utils.aoa_to_sheet([['저장날짜시간', '고객명', '담당 FC', '내용', '음성파일']]); XLSX.utils.book_append_sheet(wb, ws, '상담기록'); }
      var I = info(ws, '고객명'), dc = Math.max(I.colOf('저장날짜시간'), I.colOf('저장날짜')), nc = I.colOf('고객명'), fcc = I.colOf('담당 FC'), tc = I.colOf('내용');
      if (dc < 0 || tc < 0) throw new Error("상담기록 시트에 '저장날짜시간'·'내용' 열이 필요합니다.");
      var same = function (r, o) {
        var t = ws[I.enc(r, tc)], n = ws[I.enc(r, nc)], fcell = fcc >= 0 ? ws[I.enc(r, fcc)] : null;
        return t && n && String(t.v) === o.text && String(n.v).trim() === name && (!fcell || !String(fcell.v).trim() || String(fcell.v).trim() === fc) && dateText(ws[I.enc(r, dc)]) === o.at;
      };
      var find = function (o) { for (var r = I.hr + 1; r <= I.rg.e.r; r++) if (same(r, o)) return r; return -1; };
      if (op.add) {
        var r0 = I.rg.e.r + 1, m = op.add.at.match(/^(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2})/);
        ws[I.enc(r0, nc)] = { t: 's', v: name }; if (fcc >= 0) ws[I.enc(r0, fcc)] = { t: 's', v: fc }; ws[I.enc(r0, tc)] = { t: 's', v: op.add.text };
        ws[I.enc(r0, dc)] = { t: 'n', v: serial(new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5])), z: 'yyyy-mm-dd hh:mm' };
        I.rg.e.r = r0;
      } else if (op.edit) {
        var re = find(op.edit); if (re < 0) throw new Error('엑셀에서 이 상담기록을 찾지 못했습니다. 새로고침 후 다시 시도하세요.');
        ws[I.enc(re, tc)] = { t: 's', v: op.edit.newText };
      } else if (op.del) {
        var rd = find(op.del); if (rd < 0) throw new Error('엑셀에서 이 상담기록을 찾지 못했습니다. 새로고침 후 다시 시도하세요.');
        for (var r = rd; r < I.rg.e.r; r++) for (var c = I.rg.s.c; c <= I.rg.e.c; c++) { var nx = ws[I.enc(r + 1, c)]; if (nx) ws[I.enc(r, c)] = nx; else delete ws[I.enc(r, c)]; }
        for (var c2 = I.rg.s.c; c2 <= I.rg.e.c; c2++) delete ws[I.enc(I.rg.e.r, c2)];
        I.rg.e.r -= 1;
      }
      ws['!ref'] = XLSX.utils.encode_range(I.rg);
      await putWB(f, wb);
      return f.name;
    });
  }

  // Drive의 최신 값으로 화면 데이터를 갱신 (고객 시트의 수정 가능 항목 + 고객추가정보 기본·상담기록)
  async function refresh() {
    var K = window.CUSTOMERS, X = window.CUSTEXTRA = window.CUSTEXTRA || { basic: [], consult: [] };
    var byId = {};
    K.customers.forEach(function (c) { byId[c['고객ID']] = c; });
    var polIdx = {}; Object.keys(K.contracts || {}).forEach(function (id) { K.contracts[id].forEach(function (x) { polIdx[x.pol] = x; }); });
    var fcs = await fcFolders(), n = 0, errs = [];
    for (var i = 0; i < fcs.length; i++) {
      try {
        var f = await listFile(fcs[i].name), wb = await getWB(f), ws = wb.Sheets['고객'];
        if (!ws) continue;
        var cws = wb.Sheets['계약'];
        if (cws) XLSX.utils.sheet_to_json(cws, { defval: '', raw: false }).forEach(function (r) { var x = polIdx[String(r['증권번호'])]; if (x) { x.route = String(r['유입경로'] || '').trim(); x.intro = String(r['소개자'] || '').trim(); } });
        XLSX.utils.sheet_to_json(ws, { defval: '', raw: false }).forEach(function (r) {
          var c = byId[r['고객ID']]; if (!c) return;
          EDITABLE.forEach(function (k) { if (k in r) c[k] = String(r[k]).trim(); });
          n++;
        });
      } catch (e) { errs.push(fcs[i].name + ': ' + e.message); }
    }
    var ef = await extraFile();
    if (ef) {
      var wb2 = await getWB(ef), s = function (v) { return v == null ? '' : String(v).trim(); };
      var bs = wb2.Sheets['기본'] ? XLSX.utils.sheet_to_json(wb2.Sheets['기본'], { defval: '', raw: false }) : [];
      X.basic = bs.map(function (r) { return { fc: s(r['담당 FC']), name: s(r['고객명']), meeting: s(r['미팅장소']), memo: s(r['메모']), carrier: s(r['통신사']), car: s(r['차량번호']), renew: s(r['자동차보험갱신일']) }; });
      var cs = wb2.Sheets['상담기록'];
      if (cs) {
        var I = info(cs, '고객명'), dc = Math.max(I.colOf('저장날짜시간'), I.colOf('저장날짜')), nc = I.colOf('고객명'), fcc = I.colOf('담당 FC'), tc = I.colOf('내용'), vc = I.colOf('음성파일'), out = [];
        for (var r = I.hr + 1; r <= I.rg.e.r; r++) {
          var tx = cs[I.enc(r, tc)]; if (!tx || !String(tx.v).trim()) continue;
          out.push({ at: dateText(cs[I.enc(r, dc)]), name: s((cs[I.enc(r, nc)] || {}).v), fc: fcc >= 0 ? s((cs[I.enc(r, fcc)] || {}).v) : '', text: String(tx.v).trim(), voice: vc >= 0 ? s((cs[I.enc(r, vc)] || {}).v) : '' });
        }
        X.consult = out;
      }
    }
    return { customers: n, errors: errs };
  }

  // 공통 설정: 신한FC M과 같은 앱 설정 파일(shlife-설정.json). 읽을 때마다 최신을 읽고, 쓸 때는 바꾼 항목만 합쳐 저장 (M이 쓰는 다른 항목은 그대로 둠)
  var APPDIR = '1ERsiC6gUGh0LljPSKxRx04Z5dZUGJGC0', SETF = 'shlife-설정.json';
  async function settingsFile() {
    var a = await qlist("'" + APPDIR + "' in parents and name='" + SETF + "' and trashed=false");
    return a[0] || null;
  }
  async function readSettings() {
    var f = await settingsFile();
    if (!f) return {};
    return (await gd(API + '/files/' + encodeURIComponent(f.id) + '?alt=media&supportsAllDrives=true')) || {};
  }
  function writeSettings(patch) {
    return serialize(async function () {
      var f = await settingsFile();
      if (!f) throw new Error("앱 설정 폴더에서 '" + SETF + "' 파일을 찾지 못했습니다.");
      var cur = (await gd(API + '/files/' + encodeURIComponent(f.id) + '?alt=media&supportsAllDrives=true')) || {};
      var next = Object.assign({}, cur, patch);
      await gd(UP + '/files/' + encodeURIComponent(f.id) + '?uploadType=media&supportsAllDrives=true', { method: 'PATCH', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(next, null, 2) });
      return next;
    });
  }

  window.DRIVE = { on: on, email: email, login: login, logout: logout, onChange: onChange, updateCustomer: updateCustomer, updateContract: updateContract, consultOp: consultOp, refresh: refresh, readSettings: readSettings, writeSettings: writeSettings,
    _i: { get cache() { return cache; }, API: API, gd: gd, qlist: qlist, qv: qv, baseId: baseId, fcFolders: fcFolders, listFile: listFile, extraFile: extraFile, getWB: getWB } };
})();
