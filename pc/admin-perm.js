(function () {
  'use strict';
  // 관리자 > 권한·공유 현황: 신한FC m의 '사용 FC 현황'과 같은 방식으로 계정별 볼 FC를 지정하고 Drive 공유를 자동으로 맞춘다.
  //   계정마다 FC별 [없음 | 보기 | 편집], 이름, 직급·사번·상위자(데스크톱 접속 명단)를 고친 뒤 '변경 저장' → 무엇을 누구에게 공유/해제하는지 확인창 → 확인하면
  //   ① shlife-권한.json 저장 ② Drive 폴더 공유 부여·해제 ③ 사용자권한.json(접속 명단) 저장.
  var D = window.DRIVE, S = D && D.share;
  var ADMINS = ['blue10kp@gmail.com'];   // 엠 config.admins 와 같은 관리자(전체 FC, 이 표에서 수정하지 않음)
  var esc = function (x) { return String(x == null ? '' : x).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  var ap = { data: null, err: '', busy: false, extra: [], msg: '', gone: {}, onChange: function () {} };
  var ROLES = ['FC', 'GSL', 'SSL', 'ESL'];
  var cmpKo = function (a, b) { return String(a).localeCompare(String(b), 'ko'); };

  function load() {
    ap.err = ''; ap.busy = true;
    return Promise.all([D.readPerms(), S.loadShares(), D.readJson(['_양식', '사용자권한.json'])]).then(function (r) {
      ap.data = { perms: r[0], st: r[1], roster: (r[2] && r[2].users) || [], fcs: Object.keys(r[1].fcIds).sort(cmpKo) };
      ap.busy = false; ap.gone = {}; ap.onChange();
    }).catch(function (e) { ap.busy = false; ap.err = '불러오지 못했습니다: ' + (e && e.message || e); ap.onChange(); });
  }
  var lc = function (s) { return String(s || '').toLowerCase(); };
  function rosterOf(d, em) { return d.roster.filter(function (u) { return lc(u.email) === em; })[0] || null; }
  function listOf(d, em) { var k = Object.keys(d.perms.users).filter(function (x) { return lc(x) === em; })[0]; return k ? d.perms.users[k] : []; }
  function nameOf(d, em) { var k = Object.keys(d.perms.names).filter(function (x) { return lc(x) === em; })[0]; return (k && d.perms.names[k]) || (rosterOf(d, em) || {}).name || ''; }
  function emails(d) {
    var set = {};
    Object.keys(d.perms.users).forEach(function (e) { set[lc(e)] = 1; });
    Object.keys(d.perms.names).forEach(function (e) { set[lc(e)] = 1; });
    d.roster.forEach(function (u) { if (u.email) set[lc(u.email)] = 1; });
    ap.extra.forEach(function (e) { set[e] = 1; });
    Object.keys(d.st.fc).forEach(function (n) { d.st.fc[n].forEach(function (p) { if (p.role !== 'owner') set[p.email] = 1; }); });
    d.st.app.forEach(function (p) { if (p.role !== 'owner') set[p.email] = 1; });
    ADMINS.forEach(function (a) { delete set[a]; });
    return Object.keys(set).filter(function (e) { return !ap.gone[e]; }).sort();
  }

  function html() {
    if (!(D && D.on())) return '<section><p class="mut">Drive에 로그인하면 불러옵니다.</p></section>';
    var d = ap.data;
    if (!d) return '<section><p class="mut">' + (ap.err ? esc(ap.err) : '불러오는 중…') + '</p>' + (ap.err ? '<button class="btn" id="apReload">다시 불러오기</button>' : '') + '</section>';
    var rosterNames = d.roster.filter(function (u) { return u.no; });
    var rows = emails(d).map(function (em) {
      var r = rosterOf(d, em), list = listOf(d, em), cur = {};
      list.forEach(function (x) { cur[S.fcOf(x)] = S.roOf(x) ? 'r' : 'w'; });
      var cells = d.fcs.map(function (fc) {
        var v = cur[fc] || '';
        return '<td><select data-fc="' + esc(fc) + '"><option value=""' + (v === '' ? ' selected' : '') + '>-</option><option value="r"' + (v === 'r' ? ' selected' : '') + '>보기</option><option value="w"' + (v === 'w' ? ' selected' : '') + '>편집</option></select></td>';
      }).join('');
      var warn = [];
      if (list.length && !S.shareMatches(d.st, em, list)) warn.push('Drive 공유가 목록과 다름');
      S.ancShares(d.st, em).forEach(function (a) { warn.push("상위 폴더 '" + a.folder + "' 공유(" + a.perm.role + ")"); });
      if (!list.length && S.shareState(d.st, em) !== 'none') warn.push('공유가 남아 있음');
      if (list.length && !r) warn.push('접속 명단에 없음(데스크톱 사용 불가)');
      var bossSel = '<select data-boss><option value="">-</option>' + rosterNames.map(function (u) { return '<option value="' + esc(u.no) + '"' + (r && r.boss === u.no ? ' selected' : '') + '>' + esc(u.name) + '</option>'; }).join('') + '</select>';
      return '<tr data-em="' + esc(em) + '"><td class="l"><input type="text" data-name value="' + esc(nameOf(d, em)) + '" maxlength="20" size="7" placeholder="이름"></td><td class="l">' + esc(em) + '</td>' +
        '<td><select data-role>' + ROLES.map(function (x) { return '<option' + ((r ? r.role : 'FC') === x ? ' selected' : '') + '>' + x + '</option>'; }).join('') + '</select></td>' +
        '<td><input type="text" data-no value="' + esc(r ? r.no : '') + '" maxlength="8" size="9" placeholder="사번"></td><td>' + bossSel + '</td>' + cells +
        '<td class="l">' + (warn.length ? '<span class="warnTxt">⚠ ' + esc(warn.join(' · ')) + '</span>' : '일치') + '</td><td><button type="button" class="btn" data-del="' + esc(em) + '">삭제</button></td></tr>';
    }).join('');
    var head = '<tr><th class="l">이름</th><th class="l">구글 아이디</th><th>직급</th><th>사번</th><th>상위자</th>' + d.fcs.map(function (fc) { return '<th>' + esc(fc) + '</th>'; }).join('') + '<th class="l">점검</th><th></th></tr>';
    return '<section><h2>사용 FC 현황 <small class="mut">관리자(' + esc(ADMINS.join(', ')) + ')는 전체 FC라 표에 없습니다</small></h2>' +
      '<p class="mut">FC별로 [보기] 또는 [편집]을 고르고 \'변경 저장\'을 누르면, 저장 전에 공유·해제할 내용을 확인창으로 보여 주고, 확인하면 Google Drive 폴더 공유가 함께 바뀝니다. 직급·사번·상위자는 데스크톱 접속 명단(사용자권한.json)에 저장됩니다.</p>' +
      '<div class="wrap"><table>' + head + rows + '</table></div>' + (rows ? '' : '<p class="mut">아직 사용자가 없습니다. 아래에서 추가하세요.</p>') +
      '<div class="ue-act"><button type="button" class="btn pri" id="apSave">변경 저장</button><button type="button" class="btn" id="apReload">되돌리기(다시 불러오기)</button> <span class="mut" id="apMsg">' + esc(ap.msg) + '</span></div></section>' +
      '<section><h2>사용자 추가</h2><div class="ue-row"><input id="apNewName" type="text" placeholder="이름(FC명)" maxlength="20" autocomplete="off" style="max-width:140px"><input id="apNewEm" type="email" placeholder="구글 계정 이메일" autocomplete="off"><button type="button" class="btn" id="apAdd">추가</button></div>' +
      '<p class="mut">추가하면 위 표에 줄이 생깁니다. 볼 FC를 고르고 \'변경 저장\'을 누르세요. 데스크톱을 쓰게 하려면 직급과 사번도 적어야 합니다.</p></section>';
  }

  function collect() {   // 화면의 현재 값 → { 이메일: {name, role, no, boss, list} }
    var out = {};
    document.querySelectorAll('#main tr[data-em]').forEach(function (tr) {
      var list = [];
      tr.querySelectorAll('select[data-fc]').forEach(function (s) { if (s.value) list.push(s.getAttribute('data-fc') + (s.value === 'r' ? ':r' : '')); });
      out[tr.getAttribute('data-em')] = { name: tr.querySelector('[data-name]').value.trim(), role: tr.querySelector('[data-role]').value, no: tr.querySelector('[data-no]').value.trim(), boss: tr.querySelector('[data-boss]').value, list: list };
    });
    return out;
  }
  var sameSet = function (a, b) { return a.length === b.length && a.every(function (x) { return b.indexOf(x) >= 0; }); };
  var roLabel = function (x) { return S.fcOf(x) + (S.roOf(x) ? '(보기)' : '(편집)'); };

  // 바뀐 내용을 계산하고 확인창에 보여 줄 문장을 만든다 (아무것도 바꾸지 않음)
  function plan() {
    var d = ap.data, want = collect(), lines = [], ops = { perms: [], share: [], unshare: [], roster: [], rosterDel: [] }, warns = [];
    Object.keys(ap.gone).forEach(function (em) { want[em] = { name: '', role: 'FC', no: '', boss: '', list: [], del: true }; });
    Object.keys(want).forEach(function (em) {
      var w = want[em], cur = listOf(d, em), r = rosterOf(d, em);
      var listChanged = !sameSet(w.list, cur), nameChanged = w.name !== nameOf(d, em);
      if (w.list.length && (listChanged || nameChanged)) ops.perms.push([em, w.list, w.name]);
      if (!w.list.length && (cur.length || Object.keys(d.perms.names).some(function (k) { return lc(k) === em; }))) ops.perms.push([em, [], '']);
      if (w.list.length && !S.shareMatches(d.st, em, w.list)) ops.share.push([em, w.list]);
      if (!w.list.length && S.shareState(d.st, em) !== 'none') ops.unshare.push(em);
      if (w.del) { if (r) ops.rosterDel.push(em); return; }
      var rc = !r ? (w.no && w.name) : (r.role !== w.role || String(r.no || '') !== w.no || String(r.boss || '') !== w.boss || (r.name || '') !== w.name);
      if (rc) {
        if (!r && !(w.no && w.name)) { if (w.list.length) warns.push(em + ': 사번·이름이 없어 접속 명단에는 추가하지 않습니다(데스크톱 사용 불가).'); }
        else ops.roster.push([em, { email: em, no: w.no, name: w.name, role: w.role, boss: w.boss }, r ? 'edit' : 'add']);
      }
    });
    ops.share.forEach(function (x) { lines.push('· 공유: ' + x[0] + ' ← ' + x[1].map(roLabel).join(', ') + ' + 앱 설정 폴더(보기). 처음 공유하는 계정에는 구글 알림 메일이 갑니다.'); });
    ops.unshare.forEach(function (em) { lines.push('· 공유 해제: ' + em + ' (FC 폴더·설정 폴더 공유를 모두 해제)'); });
    ops.share.concat(ops.unshare.map(function (e) { return [e]; })).forEach(function (x) {
      S.ancShares(d.st, x[0]).forEach(function (a) { lines.push("· 함께 해제: " + x[0] + " ← '" + a.folder + "' 폴더에 직접 준 공유(" + a.perm.role + ") — 남겨 두면 FC 제한이 Drive에서 안 걸립니다"); });
      var c = d.st.cust.find(function (p) { return p.email === x[0] && p.role !== 'owner'; });
      if (c) lines.push('· 함께 해제: ' + x[0] + " ← 고객 폴더 전체 공유(" + c.role + ")");
    });
    ops.perms.forEach(function (x) { lines.push('· 권한 파일(shlife-권한.json) ' + (x[1].length ? (listOf(d, x[0]).length ? '변경' : '추가') : '삭제') + ': ' + x[0] + (x[1].length ? ' [' + x[1].map(roLabel).join(', ') + ']' : '')); });
    ops.roster.forEach(function (x) { lines.push('· 접속 명단(사용자권한.json) ' + (x[2] === 'add' ? '추가' : '변경') + ': ' + x[1].name + '(' + x[0] + ') ' + x[1].role + ' 사번 ' + x[1].no); });
    ops.rosterDel.forEach(function (em) { lines.push('· 접속 명단(사용자권한.json)에서 삭제: ' + em); });
    warns.forEach(function (x) { lines.push('· 참고: ' + x); });
    return { lines: lines, ops: ops, changed: !!(ops.perms.length || ops.share.length || ops.unshare.length || ops.roster.length || ops.rosterDel.length) };
  }

  async function apply(p) {   // 확인창을 통과한 뒤에만 호출
    var d = ap.data, o = p.ops, failed = [];
    if (o.perms.length) {
      await S.updatePerms(function (perms) {
        o.perms.forEach(function (x) {
          var em = x[0], list = x[1], nm = x[2];
          Object.keys(perms.users).forEach(function (k) { if (lc(k) === em) delete perms.users[k]; });
          Object.keys(perms.names).forEach(function (k) { if (lc(k) === em && !list.length) delete perms.names[k]; });
          if (list.length) { perms.users[em] = list; if (nm) perms.names[em] = nm; else delete perms.names[em]; }
        });
        perms.fcFolders = d.st.fcIds;   // 사용자 앱이 FC 폴더를 ID로 바로 열 수 있게(엠과 같음)
      });
    }
    var tAll = { cust: false, app: false, fc: [], anc: [] }, merge = function (t) { tAll.cust = tAll.cust || t.cust; tAll.app = tAll.app || t.app; tAll.fc = tAll.fc.concat(t.fc); tAll.anc = tAll.anc.concat(t.anc); };
    for (var i = 0; i < o.share.length; i++) { try { merge(await S.grant(d.st, o.share[i][0], o.share[i][1])); } catch (e) { console.error(e); failed.push(o.share[i][0] + ': ' + (e.message || e)); } }
    for (var j = 0; j < o.unshare.length; j++) { try { merge(await S.unshare(d.st, o.unshare[j])); } catch (e) { console.error(e); failed.push(o.unshare[j] + ': ' + (e.message || e)); } }
    if (o.roster.length || o.rosterDel.length) {
      await D.updateRoster(function (data) {
        o.roster.forEach(function (x) {
          var idx = -1; data.users.forEach(function (u, k) { if (lc(u.email) === x[0]) idx = k; });
          if (idx >= 0) data.users[idx] = Object.assign({}, data.users[idx], x[1]); else data.users.push(x[1]);
        });
        o.rosterDel.forEach(function (em) { data.users = data.users.filter(function (u) { return lc(u.email) !== em; }); });
      });
    }
    return failed;
  }

  function bind(rerender) {
    ap.onChange = rerender;
    var g = function (id) { return document.getElementById(id); };
    if (D && D.on() && !ap.data && !ap.busy && !ap.err) load();
    if (g('apReload')) g('apReload').onclick = function () { ap.data = null; ap.extra = []; ap.msg = ''; load(); };
    document.querySelectorAll('#main button[data-del]').forEach(function (b) { b.onclick = function () { var em = b.getAttribute('data-del'); ap.gone[em] = 1; ap.msg = '삭제할 계정으로 표시했습니다. \'변경 저장\'을 누르면 공유 해제와 함께 반영됩니다.'; rerender(); }; });
    if (g('apAdd')) g('apAdd').onclick = function () {
      var em = lc(g('apNewEm').value.trim());
      if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(em)) { ap.msg = '이메일 형식을 확인해 주세요.'; return rerender(); }
      if (ADMINS.indexOf(em) >= 0) { ap.msg = '관리자는 이미 모든 폴더를 볼 수 있습니다.'; return rerender(); }
      if (!ap.data.perms.names[em]) ap.data.perms.names[em] = g('apNewName').value.trim();
      if (ap.extra.indexOf(em) < 0) ap.extra.push(em);
      delete ap.gone[em]; ap.msg = em + ' 줄을 추가했습니다. 볼 FC와 직급·사번을 입력하고 저장하세요.'; rerender();
    };
    if (g('apSave')) g('apSave').onclick = async function () {
      var p = plan();
      if (!p.changed) { ap.msg = '바뀐 내용이 없습니다.'; return rerender(); }
      if (!confirm('저장하면 다음이 실제로 바뀝니다.\n\n' + p.lines.join('\n') + '\n\n진행할까요?')) return;
      var b = g('apSave'); b.disabled = true; ap.msg = '저장 중… (Drive 공유는 몇 초 걸릴 수 있습니다)'; var m = g('apMsg'); if (m) m.textContent = ap.msg;
      try {
        var failed = await apply(p);
        ap.msg = failed.length ? '저장했지만 공유하지 못한 계정이 있습니다: ' + failed.join(' / ') : '저장했습니다. Drive 공유도 함께 바꿨습니다.';
        ap.data = null; ap.extra = []; ap.gone = {}; await load();
      } catch (e) { ap.msg = '저장하지 못했습니다: ' + (e && e.message || e); ap.data = null; await load(); }
    };
  }
  window.ADMINPERM = { html: html, bind: bind, plan: plan, collect: collect, state: ap };
})();
