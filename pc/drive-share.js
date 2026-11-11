(function () {
  'use strict';
  // Drive 폴더 공유(권한) 읽기·변경 + 접근 권한 파일(shlife-권한.json)·접속 명단(사용자권한.json) 저장.
  // 신한FC m의 관리자 화면(grantFolders·unshareFolders·shareMatches)과 같은 규칙을 데스크톱에 옮긴 것이다.
  //   고객\<FC> 폴더: 체크한 FC = 보기(reader), 수정 체크 = 편집(writer) / 앱 설정 폴더: 보기 / 빠진 FC 폴더는 공유 해제
  //   고객 루트·상위 폴더(Claude 등)에 직접 공유돼 있으면 해제(그대로 두면 FC 제한이 Drive에서 안 걸리기 때문)
  // 이 파일의 함수는 호출하는 쪽(관리자 화면)이 확인창을 거친 뒤에만 부른다.
  var D = window.DRIVE, I = D._i;
  var APPDIR = D.appDirId;
  var fcOf = function (s) { return String(s).replace(/:r$/, ''); };
  var roOf = function (s) { return /:r$/.test(String(s)); };
  var sleep = function (ms) { return new Promise(function (r) { setTimeout(r, ms); }); };
  var permUrl = function (id, pid) { return I.API + '/files/' + encodeURIComponent(id) + '/permissions' + (pid ? '/' + encodeURIComponent(pid) : '') + '?supportsAllDrives=true'; };

  async function permList(id, directOnly) {   // 이 폴더에 '직접' 준 사용자 권한만(상위에서 물려받은 것 제외)
    var d = await I.gd(I.API + '/files/' + encodeURIComponent(id) + '/permissions?supportsAllDrives=true&fields=' + encodeURIComponent('permissions(id,emailAddress,role,type,permissionDetails(inherited))'));
    return (d.permissions || []).filter(function (p) { return p.type === 'user' && p.emailAddress; })
      .filter(function (p) { return !directOnly || !(p.permissionDetails || []).length || (p.permissionDetails || []).some(function (x) { return !x.inherited; }); })
      .map(function (p) { return { id: p.id, email: p.emailAddress.toLowerCase(), role: p.role }; });
  }
  async function ancestors(baseId) {   // 고객 폴더의 상위 폴더(Claude 등)와 거기에 직접 준 공유
    var out = [], jobs = [];
    var getF = function (id) { return I.gd(I.API + '/files/' + encodeURIComponent(id) + '?fields=id,name,parents&supportsAllDrives=true').catch(function () { return null; }); };
    var cur = await getF(baseId);
    for (var i = 0; i < 5; i++) {
      var pid = cur && (cur.parents || [])[0]; if (!pid) break;
      var pf = await getF(pid); if (!pf) break;
      var item = { id: pid, name: pf.name, perms: [] }; out.push(item);
      jobs.push(permList(pid, true).catch(function () { return []; }).then((function (it) { return function (p) { it.perms = p; }; })(item)));
      if (!(pf.parents || []).length) break;
      cur = pf;
    }
    await Promise.all(jobs);
    return out;
  }
  async function loadShares() {
    var base = await I.baseId(), fcs = await I.fcFolders(), fcIds = {}, st = { fc: {} };
    fcs.forEach(function (f) { fcIds[f.name] = f.id; });
    await Promise.all([
      permList(base, true).then(function (p) { st.cust = p; }),
      permList(APPDIR, true).then(function (p) { st.app = p; }),
      ancestors(base).catch(function () { return []; }).then(function (a) { st.anc = a; }),
      Promise.all(fcs.map(function (f) { return permList(f.id, true).catch(function () { return []; }).then(function (p) { st.fc[f.name] = p; }); }))
    ]);
    st.baseId = base; st.fcIds = fcIds;
    return st;
  }
  var ancShares = function (st, em) {
    var out = []; (st.anc || []).forEach(function (a) { a.perms.filter(function (p) { return p.email === em && p.role !== 'owner'; }).forEach(function (p) { out.push({ folderId: a.id, folder: a.name, perm: p }); }); });
    return out;
  };
  function shareState(st, em) {   // 'ok' | 'part' | 'none'
    var anyFc = Object.keys(st.fc).some(function (n) { return st.fc[n].some(function (p) { return p.email === em; }); });
    var app = st.app.find(function (p) { return p.email === em; }), cust = st.cust.find(function (p) { return p.email === em; });
    return (anyFc && app) ? 'ok' : (anyFc || app || cust || ancShares(st, em).length) ? 'part' : 'none';
  }
  function shareMatches(st, em, list) {   // 권한 목록이 Drive 공유와 정확히 맞는가
    if (!st.app.find(function (p) { return p.email === em; })) return false;
    if (ancShares(st, em).length) return false;
    if (st.cust.find(function (p) { return p.email === em && p.role !== 'owner'; })) return false;
    return Object.keys(st.fc).every(function (name) {
      var mine = (st.fc[name] || []).find(function (p) { return p.email === em; });
      var want = list.find(function (x) { return fcOf(x) === name; });
      if (!want) return !mine;
      return !!mine && (roOf(want) ? mine.role === 'reader' : (mine.role === 'writer' || mine.role === 'owner'));
    });
  }

  // 공유 맞추기: 상위 공유 해제 → 고객 루트 직접 공유 해제 → 체크한 FC 폴더 → 빠진 FC 폴더 해제 → 설정 폴더 보기. 실패하면 지운 공유를 되돌린다.
  async function retry(fn, tries) {
    var last;
    for (var k = 0; k < (tries || 3); k++) { try { return await fn(); } catch (e) { last = e; if (!/Drive (5\d\d|429)|Failed to fetch|NetworkError/.test(String(e.message))) throw e; await sleep(2500); } }
    throw last;
  }
  async function grant(st, em, list) {
    list = (list || []).filter(function (g) { return g !== '*'; });
    if (!list.length) throw new Error('FC를 하나 이상 지정하세요.');
    var ids = st.fcIds;
    list.forEach(function (g) { if (!ids[fcOf(g)]) throw new Error('FC 폴더를 찾을 수 없습니다: ' + fcOf(g)); });
    var hadAny = st.cust.some(function (p) { return p.email === em; }) || st.app.some(function (p) { return p.email === em; }) || Object.keys(st.fc).some(function (n) { return st.fc[n].some(function (p) { return p.email === em; }); });
    var add = function (id, role) { return I.gd(permUrl(id) + '&sendNotificationEmail=' + (hadAny ? 'false' : 'true'), { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ type: 'user', role: role, emailAddress: em }) }); };
    var patch = function (id, pid, role) { return I.gd(permUrl(id, pid), { method: 'PATCH', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ role: role }) }); };
    var del = function (id, pid) { return I.gd(permUrl(id, pid), { method: 'DELETE' }); };
    var undo = [], touched = { cust: false, app: false, fc: [], anc: [] };
    var remove = async function (id, perm, label) { await retry(function () { return del(id, perm.id); }); undo.push({ label: label, fn: function () { return add(id, perm.role === 'owner' ? 'writer' : perm.role); } }); };
    var mineIn = async function (id) { return (await permList(id, true)).find(function (p) { return p.email === em; }); };
    var roleOk = function (m, role) { return !!m && (m.role === role || (role === 'writer' && m.role === 'owner')); };
    var waitGone = async function (id) { for (var k = 0; k < 8; k++) { if (!(await mineIn(id))) return; await sleep(800); } throw new Error('공유 해제가 반영되지 않았습니다: ' + id); };
    var ensureDirect = async function (id, role) {
      for (var a = 0; a < 3; a++) {
        var mine = await mineIn(id);
        if (roleOk(mine, role)) return;
        if (mine && mine.role !== 'owner') await retry(function () { return patch(id, mine.id, role); }); else await retry(function () { return add(id, role); });
        for (var k = 0; k < 8; k++) { await sleep(800); if (roleOk(await mineIn(id), role)) return; }
      }
      throw new Error('공유가 반영되지 않았습니다: ' + id);
    };
    try {
      var ancs = ancShares(st, em);
      for (var i = 0; i < ancs.length; i++) { await remove(ancs[i].folderId, ancs[i].perm, '상위 폴더 ' + ancs[i].folder); touched.anc.push(ancs[i].folderId); }
      var c = st.cust.find(function (p) { return p.email === em; });
      if (c && c.role !== 'owner') { await remove(st.baseId, c, '고객 루트'); touched.cust = true; }
      for (var j = 0; j < ancs.length; j++) await waitGone(ancs[j].folderId);
      if (touched.cust) await waitGone(st.baseId);
      for (var w = 0; w < list.length; w++) { var n = fcOf(list[w]); touched.fc.push(n); await ensureDirect(ids[n], roOf(list[w]) ? 'reader' : 'writer'); }
      var extra = Object.keys(st.fc).filter(function (fc) { return st.fc[fc].some(function (p) { return p.email === em && p.role !== 'owner'; }) && !list.some(function (g) { return fcOf(g) === fc; }); });
      for (var x = 0; x < extra.length; x++) { touched.fc.push(extra[x]); await remove(ids[extra[x]], st.fc[extra[x]].find(function (p) { return p.email === em; }), 'FC ' + extra[x]); }
      await ensureDirect(APPDIR, 'reader'); touched.app = true;
    } catch (e) {
      var failed = [];
      for (var u = undo.length - 1; u >= 0; u--) { try { await undo[u].fn(); } catch (e2) { failed.push(undo[u].label); } }
      if (failed.length) e.message += ' (되돌리지 못한 공유: ' + failed.join(', ') + ')';
      throw e;
    }
    return touched;
  }
  async function unshare(st, em) {
    var del = function (id, p) { return I.gd(permUrl(id, p.id), { method: 'DELETE' }); };
    var touched = { cust: false, app: false, fc: [], anc: [] };
    var a = ancShares(st, em); for (var i = 0; i < a.length; i++) { await del(a[i].folderId, a[i].perm); touched.anc.push(a[i].folderId); }
    var cs = st.cust.filter(function (x) { return x.email === em && x.role !== 'owner'; }); for (var k = 0; k < cs.length; k++) { await del(st.baseId, cs[k]); touched.cust = true; }
    var ap = st.app.filter(function (x) { return x.email === em && x.role !== 'owner'; }); for (var m = 0; m < ap.length; m++) { await del(APPDIR, ap[m]); touched.app = true; }
    var names = Object.keys(st.fc);
    for (var n = 0; n < names.length; n++) { var ps = st.fc[names[n]].filter(function (x) { return x.email === em && x.role !== 'owner'; }); for (var q = 0; q < ps.length; q++) { await del(st.fcIds[names[n]], ps[q]); touched.fc.push(names[n]); } }
    return touched;
  }
  async function refreshShares(st, t) {   // 바뀐 폴더의 공유만 다시 읽음
    var jobs = [];
    if (t.cust) jobs.push(permList(st.baseId, true).then(function (p) { st.cust = p; }));
    if (t.app) jobs.push(permList(APPDIR, true).then(function (p) { st.app = p; }));
    t.fc.forEach(function (n) { if (st.fcIds[n]) jobs.push(permList(st.fcIds[n], true).then(function (p) { st.fc[n] = p; })); });
    t.anc.forEach(function (id) { var a = (st.anc || []).find(function (x) { return x.id === id; }); if (a) jobs.push(permList(id, true).then(function (p) { a.perms = p; })); });
    await Promise.all(jobs);
  }

  // 접근 권한 파일(shlife-권한.json) 저장: 저장 직전에 최신을 읽고 fn으로 바꾼 뒤 올림(다른 항목 보존)
  function updatePerms(fn) {
    return I.serialize(async function () {
      var a = await I.qlist("'" + APPDIR + "' in parents and name='shlife-권한.json' and trashed=false");
      if (!a[0]) throw new Error("앱 설정 폴더에서 'shlife-권한.json' 파일을 찾지 못했습니다.");
      var cur = (await I.gd(I.API + '/files/' + encodeURIComponent(a[0].id) + '?alt=media&supportsAllDrives=true')) || {};
      cur.users = cur.users || {}; cur.names = cur.names || {};
      var next = fn(cur) || cur;
      await I.gd('https://www.googleapis.com/upload/drive/v3/files/' + encodeURIComponent(a[0].id) + '?uploadType=media&supportsAllDrives=true', { method: 'PATCH', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(next, null, 2) });
      return next;
    });
  }

  D.share = { loadShares: loadShares, shareMatches: shareMatches, shareState: shareState, ancShares: ancShares, grant: grant, unshare: unshare, refresh: refreshShares, updatePerms: updatePerms, fcOf: fcOf, roOf: roOf };
})();
