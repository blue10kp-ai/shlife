/* 캘린더 일정 ↔ 고객 연결 공통 모듈 (규칙: app\_공통\캘린더고객연결.md, 2026-11-24)
 * 웹앱은 이 파일을 앱 소스 폴더에 calcust.js로 복사해 불러온다. 원본은 app\_공통\캘린더고객연결.js 하나.
 * 이 파일에는 고객·사번·토큰 값을 넣지 않는다. 토큰은 공통 출입증(localStorage blue10kp.gauth) 또는 CalCust.config({token})으로 받는다.
 * 쓰는 법(요약):
 *   CalCust.linkOf(event.extendedProperties.private) → {value, custId, ignore, old, how, fc, custFc, kind, at, manual}
 *   CalCust.timesOf(event) → {날짜, 시작, 끝, 종일}      CalCust.kindOf(제목)
 *   CalCust.fcOfCreator(event.creator.email, names) → 일정 담당 FC(= 일정을 만든 계정, 권한 파일 names로 이름). 못 바꾸면 ""
 *   CalCust.fcOfTitle(제목, prefixes) → 붙임말 FC(보기용 참고만, 담당 판정에 쓰지 않음)
 *   await CalCust.read({calId, calName, timeMin, timeMax}) → [{일정ID, calId, eventId, 캘린더, 제목, updated, link, ...timesOf}]
 *   await CalCust.find({calId, custId}) → 그 고객ID로 연결된 일정(기간 없이)
 *   await CalCust.write({calId, eventId, custId | ignore:true | clear:true, how, fc, custFc, kind, priv}) → 바뀌었으면 구글 응답 일정, 같으면 null
 *   CalCust.rowOf({...}) → 기록 파일 한 줄      await CalCust.recordPut(folderId, rows, account) → 그 FC 폴더 일정매칭\YYYY-MM - 계정.json에 합쳐 씀
 */
(function (root) {
  "use strict";
  const GCAL = "https://www.googleapis.com/calendar/v3", GDRV = "https://www.googleapis.com/drive/v3", GUP = "https://www.googleapis.com/upload/drive/v3";
  const K = { cust: "shlifeCust", fc: "shlifeFc", custFc: "shlifeCustFc", kind: "shlifeKind", how: "shlifeHow", at: "shlifeAt" };
  const IGNORE = "__ignore__", REC_DIR = "일정매칭";
  const KINDS = [["AP", "AP"], ["PC", "PC"], ["CS", "CS"], ["청약", "청약"], ["계약", "계약"], ["통화", "통화"], ["전화", "통화"], ["상담", "상담"]];
  let getToken = () => { try { const a = JSON.parse(localStorage.getItem("blue10kp.gauth") || "null"); return a && a.exp > Date.now() + 60000 ? a.token : ""; } catch (e) { return ""; } };
  function config(o) { if (o && typeof o.token === "function") getToken = o.token; }

  async function api(url, opt) {
    opt = opt || {};
    const t = getToken();
    if (!t) throw new Error("GOOGLE_AUTH");
    const r = await fetch(url, { method: opt.method || "GET", body: opt.body, headers: { Authorization: "Bearer " + t, ...(opt.headers || {}) } });
    if (r.status === 401) throw new Error("GOOGLE_AUTH");
    if (!r.ok) { const e = new Error("Google " + r.status + ": " + (await r.text()).slice(0, 160)); e.status = r.status; throw e; }
    if (opt.text) return r.text();
    return r.status === 204 ? null : r.json();
  }
  const pad = (n) => String(n).padStart(2, "0");
  const ymd = (d) => d.getUTCFullYear() + "-" + pad(d.getUTCMonth() + 1) + "-" + pad(d.getUTCDate());
  function kst(d) { const k = new Date(+new Date(d) + 9 * 3600e3); return ymd(k) + "T" + pad(k.getUTCHours()) + ":" + pad(k.getUTCMinutes()) + ":" + pad(k.getUTCSeconds()) + "+09:00"; }
  const nowKst = () => kst(new Date());

  // 활동 종류: 제목에서 처음 나온 낱말 (영문은 낱말 단위, 대소문자 무시)
  function kindOf(title) {
    const t = String(title || "");
    let best = null;
    KINDS.forEach(([w, v]) => {
      let i;
      if (/^[A-Z]+$/.test(w)) { const m = new RegExp("(^|[^A-Za-z])" + w + "(?![A-Za-z])", "i").exec(t); i = m ? m.index + m[1].length : -1; }
      else i = t.indexOf(w);
      if (i >= 0 && (!best || i < best[0])) best = [i, v];
    });
    return best ? best[1] : "";
  }
  // 일정 담당 FC: 제목 앞 붙임말 'X)'를 prefixes {FC이름: 붙임말}과 맞춤. 없거나 둘 이상 맞으면 "" (짐작하지 않음)
  function fcOfTitle(title, prefixes) {
    const m = /^\s*\[?([^\s)\]]{1,3})[)\]]/.exec(String(title || ""));
    if (!m) return "";
    const hit = Object.entries(prefixes || {}).filter(([, p]) => String(p || "").replace(/[)\]\s]/g, "") === m[1]).map(([fc]) => fc);
    return hit.length === 1 ? hit[0] : "";
  }
  // 일정 담당 FC = 일정을 만든 계정(creator). 권한 파일 names {이메일: 이름}에 없으면 "" (외부 계정·캘린더 자체 = 담당 미정)
  function fcOfCreator(email, names) {
    const em = String(email || "").trim().toLowerCase();
    if (!em) return "";
    const hit = Object.entries(names || {}).find(([k]) => String(k).trim().toLowerCase() === em);
    return hit ? String(hit[1] || "").trim() : "";
  }
  // 숨은 칸 읽기. shlifeHow가 없는 옛 값은 '선택'(무시 값이면 '무시')으로 봄
  function linkOf(priv) {
    priv = priv || {};
    const v = String(priv[K.cust] || "");
    const how = String(priv[K.how] || "") || (v === IGNORE ? "무시" : v ? "선택" : "");
    const l = { value: v, custId: /^id:/.test(v) ? v.slice(3) : "", ignore: v === IGNORE, old: !!v && v !== IGNORE && !/^id:/.test(v), how, fc: String(priv[K.fc] || ""), custFc: String(priv[K.custFc] || ""), kind: String(priv[K.kind] || ""), at: String(priv[K.at] || "") };
    l.manual = how === "선택" || how === "무시";
    return l;
  }
  // 날짜·시각 (규칙 '일정ID·날짜 적는 법')
  function timesOf(ev) {
    const s = ev.start || {}, e = ev.end || {};
    if (s.date) {
      const last = new Date(Date.parse((e.date || s.date) + "T00:00:00Z") - 864e5);
      const endIncl = e.date && e.date > s.date ? ymd(last) : s.date;
      return { 날짜: s.date, 시작: s.date, 끝: endIncl, 종일: "Y" };
    }
    const st = kst(s.dateTime), en = e.dateTime ? kst(e.dateTime) : st;
    return { 날짜: st.slice(0, 10), 시작: st, 끝: en, 종일: "N" };
  }
  const calPath = (calId) => GCAL + "/calendars/" + encodeURIComponent(calId) + "/events";
  async function listAll(url) {
    let out = [], pt = "";
    do { const d = await api(url + (pt ? "&pageToken=" + encodeURIComponent(pt) : "")); out = out.concat(d.items || []); pt = d.nextPageToken || ""; } while (pt);
    return out.filter((e) => e.status !== "cancelled");
  }
  const shape = (calId, calName, ev) => ({ 일정ID: calId + "|" + ev.id, calId, eventId: ev.id, 캘린더: calName || "", 제목: String(ev.summary || ""), 만든계정: String((ev.creator || {}).email || ""), updated: ev.updated || "", link: linkOf((ev.extendedProperties || {}).private), ...timesOf(ev) });
  async function read(o) {
    const q = "?singleEvents=true&orderBy=startTime&maxResults=250&timeMin=" + encodeURIComponent(new Date(o.timeMin).toISOString()) + "&timeMax=" + encodeURIComponent(new Date(o.timeMax).toISOString());
    return (await listAll(calPath(o.calId) + q)).map((ev) => shape(o.calId, o.calName, ev));
  }
  async function find(o) {
    const q = "?singleEvents=true&maxResults=250&privateExtendedProperty=" + encodeURIComponent(K.cust + "=id:" + o.custId);
    return (await listAll(calPath(o.calId) + q)).map((ev) => shape(o.calId, o.calName, ev));
  }
  // 연결 쓰기: custId(연결) / ignore(무시) / clear(해제). priv = 지금 숨은 칸(같으면 쓰지 않음). 사람이 고른 것은 자동이 덮지 않음
  function nextPriv(o) {
    const how = o.how || (o.ignore ? "무시" : o.clear ? "해제" : "선택");
    const v = o.ignore ? IGNORE : o.clear ? "" : o.custId ? "id:" + o.custId : "";
    return { [K.cust]: v, [K.how]: how, [K.fc]: o.fc || "", [K.custFc]: o.custId ? o.custFc || "" : "", [K.kind]: o.kind || "" };
  }
  function same(priv, next) { priv = priv || {}; return Object.keys(next).every((k) => String(priv[k] || "") === String(next[k] || "")); }
  async function write(o) {
    const cur = linkOf(o.priv);
    if (o.how === "자동" && cur.manual) return null; // 사람이 고른 것(선택·무시)은 자동이 덮지 않음
    const next = nextPriv(o);
    if (o.priv && same(o.priv, next)) return null;
    next[K.at] = nowKst();
    return api(calPath(o.calId) + "/" + encodeURIComponent(o.eventId), { method: "PATCH", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ extendedProperties: { private: next } }) });
  }
  // 기록 파일 한 줄 (ev = 구글 일정 또는 {start,end,summary,updated})
  function rowOf(o) {
    const ev = o.ev || {};
    return { 일정ID: o.calId + "|" + (o.eventId || ev.id), 캘린더: o.calName || "", ...timesOf(ev), 제목: String(o.title != null ? o.title : ev.summary || ""), FC: o.fc || "", 만든계정: o.creator != null ? o.creator : String((ev.creator || {}).email || ""), 붙임말FC: o.tagFc || "", 고객ID: o.custId || "", 고객FC: o.custId ? o.custFc || "" : "", 종류: o.kind || "", 방식: o.how || "", 일정수정시각: o.updated || ev.updated || "", 적은시각: nowKst(), 적은계정: o.account || "" };
  }
  // ---- 드라이브 기록 파일 ----
  const qv = (s) => String(s).replace(/\\/g, "\\\\").replace(/'/g, "\\'");
  async function child(parentId, name, folder) {
    const q = `'${qv(parentId)}' in parents and name='${qv(name)}' and trashed=false` + (folder ? " and mimeType='application/vnd.google-apps.folder'" : "");
    const d = await api(GDRV + "/files?q=" + encodeURIComponent(q) + "&fields=files(id,name)&supportsAllDrives=true&includeItemsFromAllDrives=true&pageSize=10");
    return (d.files || [])[0] || null;
  }
  async function ensureDir(parentId) {
    const f = await child(parentId, REC_DIR, true);
    if (f) return f.id;
    const d = await api(GDRV + "/files?supportsAllDrives=true&fields=id", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ name: REC_DIR, mimeType: "application/vnd.google-apps.folder", parents: [parentId] }) });
    return d.id;
  }
  async function putJson(dirId, name, fileId, obj) {
    const body = JSON.stringify(obj, null, 1);
    if (fileId) return api(GUP + "/files/" + encodeURIComponent(fileId) + "?uploadType=media&supportsAllDrives=true", { method: "PATCH", headers: { "Content-Type": "application/json" }, body });
    const b = "calcust" + Date.now();
    const mp = `--${b}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify({ name, parents: [dirId], mimeType: "application/json" })}\r\n--${b}\r\nContent-Type: application/json\r\n\r\n${body}\r\n--${b}--`;
    return api(GUP + "/files?uploadType=multipart&supportsAllDrives=true&fields=id", { method: "POST", headers: { "Content-Type": "multipart/related; boundary=" + b }, body: mp });
  }
  const chains = {}; // 같은 파일 쓰기는 차례로(이 페이지 안)
  async function recordPut(folderId, rows, account) {
    rows = (rows || []).filter((r) => r && r.일정ID && r.날짜);
    if (!folderId || !rows.length) return 0;
    const who = String(account || "").split("@")[0] || "unknown";
    const byMonth = {};
    rows.forEach((r) => (byMonth[r.날짜.slice(0, 7)] = byMonth[r.날짜.slice(0, 7)] || []).push(r));
    let n = 0;
    for (const [ym, list] of Object.entries(byMonth)) {
      const key = folderId + "|" + ym + "|" + who;
      const run = async () => {
        const dirId = await ensureDir(folderId), name = ym + " - " + who + ".json";
        const f = await child(dirId, name, false); // 쓰기 직전에 다시 읽어 합침
        let doc = null;
        if (f) { try { doc = JSON.parse(await api(GDRV + "/files/" + encodeURIComponent(f.id) + "?alt=media&supportsAllDrives=true", { text: true })); } catch (e) { doc = null; } }
        doc = doc && typeof doc === "object" && doc.rows ? doc : { v: 1, 계정: account || "", 달: ym, rows: {} };
        let changed = 0;
        list.forEach((r) => { const o = doc.rows[r.일정ID]; if (o && String(o.적은시각 || "") > String(r.적은시각 || "")) return; const keep = ["날짜", "시작", "끝", "종일", "제목", "FC", "만든계정", "붙임말FC", "고객ID", "고객FC", "종류", "방식", "일정수정시각"]; if (o && keep.every((k) => String(o[k] || "") === String(r[k] || ""))) return; doc.rows[r.일정ID] = r; changed++; });
        if (!changed) return 0;
        await putJson(dirId, name, f && f.id, doc);
        return changed;
      };
      const p = (chains[key] || Promise.resolve()).catch(() => {}).then(run);
      chains[key] = p;
      n += await p;
    }
    return n;
  }
  root.CalCust = { version: 2, keys: K, IGNORE, config, kindOf, fcOfCreator, fcOfTitle, linkOf, timesOf, read, find, nextPriv, same, write, rowOf, recordPut, kst };
})(typeof window !== "undefined" ? window : globalThis);
