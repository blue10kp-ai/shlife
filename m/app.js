(function () {
  "use strict";
  const CFG = window.APP_CONFIG;
  // Files.Read.All: 내 파일 + 다른 사람이 공유해 준 폴더까지 읽기 (읽기 전용)
  const SCOPES = ["Files.Read.All", "User.Read"];
  const GRAPH = "https://graph.microsoft.com/v1.0";
  const $ = (id) => document.getElementById(id);

  // ---------- 로그인 (MSAL) ----------
  const redirectUri = location.origin + location.pathname.replace(/index\.html$/, "");
  const msal = !window.msal ? null : new window.msal.PublicClientApplication({
    auth: {
      clientId: CFG.clientId,
      authority: "https://login.microsoftonline.com/consumers",
      redirectUri
    },
    cache: { cacheLocation: "localStorage" }
  });

  async function getToken() {
    const account = msal.getAllAccounts()[0];
    if (!account) throw new Error("NOT_SIGNED_IN");
    try {
      const r = await msal.acquireTokenSilent({ scopes: SCOPES, account });
      return r.accessToken;
    } catch (e) {
      await msal.acquireTokenRedirect({ scopes: SCOPES, account });
      throw e;
    }
  }

  async function graph(path, asBuffer) {
    const token = await getToken();
    const res = await fetch(path.startsWith("http") ? path : GRAPH + path, {
      headers: { Authorization: "Bearer " + token },
      cache: "no-store" // OneDrive 파일이 바뀌면 바로 반영 (브라우저 캐시 사용 안 함)
    });
    if (!res.ok) {
      const t = await res.text();
      throw new Error("Graph " + res.status + ": " + t.slice(0, 200));
    }
    return asBuffer ? res.arrayBuffer() : res.json();
  }

  async function listChildren(pathOrId) {
    let url = pathOrId + (pathOrId.includes("?") ? "&" : "?") +
      "$top=999&$select=id,name,folder,file,webUrl,size,lastModifiedDateTime";
    const out = [];
    while (url) {
      const data = await graph(url);
      out.push(...data.value);
      url = data["@odata.nextLink"] || null;
    }
    return out;
  }

  // ---------- 관리자 ----------
  // 관리자: 공유 권한을 바꾸려면 쓰기 권한이 필요 (관리자만 요청)
  const ADMIN_SCOPES = ["Files.ReadWrite.All", "User.Read"];
  const account = () => GD ? (gAccount() ? { username: gAccount() } : undefined) : msal.getAllAccounts()[0];
  const isAdmin = () => !!account() &&
    (CFG.admins || []).map((a) => a.toLowerCase()).includes(String(account().username).toLowerCase());

  async function gx(method, path, body, contentType) {
    const acc = account();
    let token;
    try {
      token = (await msal.acquireTokenSilent({ scopes: ADMIN_SCOPES, account: acc })).accessToken;
    } catch (e) {
      sessionStorage.setItem("openAdmin", "1");
      await msal.acquireTokenRedirect({ scopes: ADMIN_SCOPES, account: acc });
      throw e;
    }
    const raw = body instanceof ArrayBuffer || ArrayBuffer.isView(body);
    const res = await fetch(GRAPH + path, {
      method,
      headers: { Authorization: "Bearer " + token, "Content-Type": contentType || "application/json" },
      body: body == null ? undefined : raw ? body : JSON.stringify(body),
      cache: "no-store"
    });
    if (!res.ok) throw new Error("Graph " + res.status + ": " + (await res.text()).slice(0, 200));
    return res.status === 204 ? null : res.json();
  }

  // ---------- 저장소: OneDrive 또는 Google Drive ----------
  //   기본은 config.storage. 주소 끝에 ?drive=google / ?drive=onedrive 를 붙이면 이 기기에서 그 저장소로 바뀜(시험·전환용)
  const DRIVE = (() => {
    const q = new URLSearchParams(location.search).get("drive");
    // (시험 때 쓴 예전 키 shlife.drive 는 무시: 전환 후 모든 기기가 config 기본값을 따름)
    if (q) { try { localStorage.setItem("shlife.drive2", q); } catch (e) { /* 이번만 적용 */ } }
    let v = q;
    if (!v) { try { v = localStorage.getItem("shlife.drive2"); } catch (e) { v = null; } }
    return (v || CFG.storage || "onedrive") === "google" ? "google" : "onedrive";
  })();
  const GD = DRIVE === "google";
  const DRIVE_NAME = GD ? "Google Drive" : "OneDrive";
  const GDCFG = CFG.gdrive || {};
  const GDAPI = "https://www.googleapis.com/drive/v3";
  const GDUP = "https://www.googleapis.com/upload/drive/v3";
  const GD_FIELDS = "id,name,mimeType,size,modifiedTime,webViewLink";
  const GD_FOLDER = "application/vnd.google-apps.folder";
  const GACC_KEY = "shlife.gaccount";
  function gAccount() { try { return localStorage.getItem(GACC_KEY) || ""; } catch (e) { return ""; } }
  // Google Drive 파일 정보를 OneDrive(Graph)와 같은 모양으로
  const gdNorm = (f) => ({
    id: f.id, name: f.name, size: Number(f.size || 0), lastModifiedDateTime: f.modifiedTime, webUrl: f.webViewLink,
    folder: f.mimeType === GD_FOLDER ? {} : undefined, file: f.mimeType === GD_FOLDER ? undefined : { mimeType: f.mimeType }
  });
  async function gdFetch(url, opt, asBuffer) {
    const t = storedGToken() || gTokenMem;
    if (!t) throw new Error("GOOGLE_AUTH");
    const o = opt || {};
    const res = await fetch(url, { ...o, headers: { Authorization: "Bearer " + t, ...(o.headers || {}) }, cache: "no-store" });
    if (res.status === 401) { clearGToken(t); throw new Error("GOOGLE_AUTH"); }
    if (!res.ok) throw new Error("Drive " + res.status + ": " + (await res.text()).slice(0, 200));
    if (res.status === 204) return null;
    return asBuffer ? res.arrayBuffer() : res.json();
  }
  const qv = (s) => String(s).replace(/\\/g, "\\\\").replace(/'/g, "\\'"); // 검색식 안 문자열
  async function gdQuery(q) {
    const out = [];
    let page = "";
    do {
      const d = await gdFetch(GDAPI + "/files?pageSize=1000&supportsAllDrives=true&includeItemsFromAllDrives=true&fields=" +
        encodeURIComponent("nextPageToken,files(" + GD_FIELDS + ")") + "&q=" + encodeURIComponent(q) + (page ? "&pageToken=" + encodeURIComponent(page) : ""));
      out.push(...(d.files || []).map(gdNorm));
      page = d.nextPageToken || "";
    } while (page);
    return out;
  }
  async function gdFindIn(dirId, name) {
    return (await gdQuery("'" + qv(dirId) + "' in parents and name='" + qv(name) + "' and trashed=false"))[0] || null;
  }

  // 폴더 안 목록
  async function stList0(id) {
    return GD ? gdQuery("'" + qv(id) + "' in parents and trashed=false") : listChildren(itemPath(id) + "/children");
  }
  // 파일 내용 (ArrayBuffer)
  async function stGet0(id) {
    return GD ? gdFetch(GDAPI + "/files/" + encodeURIComponent(id) + "?alt=media&supportsAllDrives=true", null, true)
      : graph(itemPath(id) + "/content", true);
  }
  // 파일 내용 바꾸기
  async function stPut0(id, data, mime) {
    if (!GD) return gx("PUT", itemPath(id) + "/content", data, mime);
    return gdNorm(await gdFetch(GDUP + "/files/" + encodeURIComponent(id) + "?uploadType=media&supportsAllDrives=true&fields=" + GD_FIELDS,
      { method: "PATCH", headers: { "Content-Type": mime }, body: data }));
  }

  // ---------- JSON 저장소 층 (2026-11-11) ----------
  // FC 폴더의 data.json, 고객 폴더의 고객추가정보.json 이 있으면 엑셀 대신 그것을 원본으로 쓴다.
  //   앱의 나머지 코드는 그대로 엑셀(시트) 모양으로 읽고 쓰므로, 여기서 JSON ↔ 엑셀 모양을 바꿔 준다.
  //   목록에서는 JSON을 "99999999 - …리스트.xlsx" 라는 가상 파일로 보여 주어 항상 가장 최신으로 고르게 한다(같은 폴더의 옛 엑셀은 무시됨).
  //   저장 때 파일 버전을 확인해, 읽은 뒤 다른 곳에서 바뀌었으면 덮어쓰지 않고 멈춘다.
  const JSON_FILES = { "data.json": " - 리스트.xlsx", "고객추가정보.json": " - 고객추가정보.xlsx" };
  const jsonInfo = new Map(); // fileId → { name, meta, ver }
  const isoText = (d) => d.y + "-" + String(d.m).padStart(2, "0") + "-" + String(d.d).padStart(2, "0");
  function jsonToBook(doc) {
    const wb = XLSX.utils.book_new();
    Object.keys(doc.sheets || {}).forEach((name) => {
      const sh = doc.sheets[name];
      const ws = XLSX.utils.aoa_to_sheet([sh.columns].concat(sh.rows.map((r) => r.map((v) => (v == null ? "" : v)))));
      XLSX.utils.book_append_sheet(wb, ws, name);
    });
    return wb;
  }
  // ===== 과도기 읽기(2026-11-17 데이터구조 확정안 1단계) =====
  // 새 칸이 있으면 새 칸, 없으면 옛 칸. 읽을 때만 쓰는 변환이라 data.json 원본은 바꾸지 않는다.
  //   새 칸 → 앱 코드가 아직 쓰는 옛 칸 이름으로 채워 주어(옛 칸이 비었거나 없을 때) 나머지 코드는 그대로 돈다.
  //   주민번호 → 생년월일·성별·후번호 / 고객분류(직접모집·타인모집·가망) + 유입경로 + 계약진행 표 → 현재단계·진행·진행제외·영업단계
  //   유입경로(신규유입·이관유입·DB유입) → 내부 값(계약·이관·DB) / 계약분류 → 모집유형 / 이관상태·따행상태 → 이관보관·따행보관(날짜)
  const NEWCLS = { "직접모집": "계약", "타인모집": "이관관리", "가망": "가망" };
  const ROUTE_IN = { "신규유입": "계약", "이관유입": "이관", "DB유입": "DB" }; // 새 값 → 내부 값(옛 코드가 보는 값)
  const ROUTE_OUT = { "계약": "신규유입", "지인": "신규유입", "소개": "신규유입", "신규": "신규유입", "신규유입": "신규유입", "이관": "이관유입", "이관유입": "이관유입", "DB": "DB유입", "DB유입": "DB유입" }; // 화면에 보이는 말
  const routeShow = (v) => { const t = String(v == null ? "" : v).trim(); return ROUTE_OUT[t] || t; };
  const STAGE_NEW_OLD = { "고객등록": "", "전화": "통화", "AP": "AP", "PC": "PC", "계약": "계약" }; // 계약진행 단계 → 영업단계 칩
  function parseJumin(v) { // 주민번호(13자리 / 앞6+뒤1 / 생년월일 8자리 / 뒤1) → { birth: YYYY-MM-DD, sex: 남·여, back: 뒤7자리 }
    const d = String(v == null ? "" : v).replace(/\D/g, ""), o = { birth: "", sex: "", back: "" };
    const cen = (g) => ("1256".includes(g) ? "19" : "3478".includes(g) ? "20" : "90".includes(g) ? "18" : "");
    if (d.length >= 13 || d.length === 7) { const g = d[6]; o.sex = +g % 2 === 1 ? "남" : "여"; if (d.length >= 13) o.back = d.slice(6, 13); const c = cen(g); if (c) o.birth = c + d.slice(0, 2) + "-" + d.slice(2, 4) + "-" + d.slice(4, 6); }
    else if (d.length === 8) o.birth = d.slice(0, 4) + "-" + d.slice(4, 6) + "-" + d.slice(6, 8);
    else if (d.length === 1) o.sex = +d % 2 === 1 ? "남" : "여";
    return o;
  }
  function compatDoc(doc) {
    if (!doc || !doc.sheets) return doc;
    const SS = {};
    Object.keys(doc.sheets).forEach((n) => { const s0 = doc.sheets[n]; SS[n] = { columns: (s0.columns || []).slice(), rows: (s0.rows || []).map((r) => r.slice()) }; });
    const txt = (v) => (v == null ? "" : String(v).trim());
    const day = (v) => txt(v).slice(0, 10);
    const ix = (s) => { const m = {}; s.columns.forEach((c, i) => { if (c && !(c in m)) m[c] = i; }); return m; };
    const ensure = (s, c) => { let i = s.columns.indexOf(c); if (i < 0) { s.columns.push(c); i = s.columns.length - 1; } return i; };
    const setv = (s, r, i, v) => { while (r.length < s.columns.length) r.push(null); r[i] = v === "" ? null : v; };
    const hist = SS["이력"], histRows = hist ? (() => { const m = ix(hist); return hist.rows.map((r) => ({ at: txt(r[m["일시"]]), cust: txt(r[m["고객ID"]]), ct: txt(r[m["계약ID"]]), type: txt(r[m["유형"]]), prog: txt(r[m["진행ID"]]) })); })() : [];
    const archAt = (pred) => { let best = ""; histRows.forEach((h) => { if (h.type === "이동(보관)" && pred(h) && h.at > best) best = h.at; }); return day(best); };
    // 계약진행 표 → 고객별 가장 최근 줄
    const progBy = new Map();
    if (SS["계약진행"]) { const s = SS["계약진행"], m = ix(s); s.rows.forEach((r) => { const id = txt(r[m["고객ID"]]); if (!id) return; const o = { 단계: txt(r[m["단계"]]), 상태: txt(r[m["상태"]]), 종료일: day(r[m["종료일"]]), 단계일: day(r[m["단계일"]]), 시작일: day(r[m["시작일"]]) }; const p = progBy.get(id); if (!p || o.시작일 >= p.시작일) progBy.set(id, o); }); }
    // 고객 시트
    const C = SS["고객"];
    if (C) {
      const m = ix(C), jm = m["주민번호"], v2 = !C.columns.includes("현재단계"); // v2 = 현재단계 칸이 없는 새 구조(옛 파일의 고객분류 '가망'과 헷갈리지 않게)
      const iBirth = ensure(C, "생년월일"), iSex = ensure(C, "성별"), iBack = ensure(C, "후번호"), iStage = ensure(C, "현재단계"), iRoute = ensure(C, "최초유입경로"), iProg = ensure(C, "진행"), iEx = ensure(C, "진행제외"), iBiz = ensure(C, "영업단계");
      const iCls = m["고객분류"], iIdc = m["고객ID"], iSr = ensure(C, "상령일");
      const td = new Date(), tdKey = td.getFullYear() * 10000 + (td.getMonth() + 1) * 100 + td.getDate();
      C.rows.forEach((r) => {
        if (jm != null) { // 주민번호 → 생년월일·성별·후번호. 형식이 틀렸거나(있을 수 없는 날짜) 옛 칸 값과 어긋나면 새 칸을 믿지 않고 옛 칸을 그대로 둠(2026-11-17: 잘못 채워진 주민번호 50건 사고)
          const p = parseJumin(r[jm]), ob = day(r[iBirth]), os = txt(r[iSex]);
          const dateOk = (v) => { const mm = /^(\d{4})-(\d\d)-(\d\d)$/.exec(v); if (!mm) return false; const dt = new Date(+mm[1], +mm[2] - 1, +mm[3]); return dt.getFullYear() === +mm[1] && dt.getMonth() === +mm[2] - 1 && dt.getDate() === +mm[3] && +mm[1] >= 1900 && +mm[1] <= td.getFullYear(); };
          const ok = (p.birth ? dateOk(p.birth) && (!ob || ob === p.birth) : true) && (p.sex ? (!os || os === p.sex) : true) && (p.birth || p.sex);
          if (ok) { if (p.birth) setv(C, r, iBirth, p.birth); if (p.sex) setv(C, r, iSex, p.sex); if (p.back) setv(C, r, iBack, p.back); } }
        { const bd = /^(\d{4})-(\d\d)-(\d\d)$/.exec(txt(r[iBirth])); // 상령일 = 생년월일 + 6개월의 월·일 중 오늘 이후 가장 가까운 날(칸 없이 계산, 확정안 3-2)
          if (bd && !txt(r[iSr])) { const bm = +bd[2] - 1 + 6, dim = (y, mo) => new Date(y, mo + 1, 0).getDate(); let y = td.getFullYear(); const mk = (yy) => { const mo = bm % 12, yy2 = yy + Math.floor(bm / 12); return [yy2, mo, Math.min(+bd[3], dim(yy2, mo))]; }; let a = mk(y); if (a[0] * 10000 + (a[1] + 1) * 100 + a[2] < tdKey) a = mk(y + 1); setv(C, r, iSr, a[0] + "-" + String(a[1] + 1).padStart(2, "0") + "-" + String(a[2]).padStart(2, "0")); } }
        let rt = txt(r[iRoute]); if (ROUTE_IN[rt]) { rt = ROUTE_IN[rt]; setv(C, r, iRoute, rt); }
        const cls0 = iCls != null ? txt(r[iCls]) : "", insOnly = /\(피보험자만\)\s*$/.test(cls0), cls = cls0.replace(/\s*\(피보험자만\)\s*$/, ""); // '타인모집(피보험자만)'처럼 뒤에 붙은 표시는 떼고 읽음
        if (NEWCLS[cls] && (v2 || cls !== "가망")) { // 새 고객분류 값이 있으면 새 칸 기준
          let base = insOnly ? "피보험자" : NEWCLS[cls];
          if (base === "가망" && rt === "DB") base = "DB";
          const pr = progBy.get(txt(r[iIdc]));
          if (pr) {
            if (pr.상태 === "진행") { if (base === "가망" || base === "DB") base = "진행"; else setv(C, r, iProg, "Y"); }
            else if ((pr.상태 === "보관" || pr.상태 === "삭제" || pr.상태 === "완료") && base === "가망") { base = "진행"; setv(C, r, iEx, pr.종료일 || pr.단계일 || "보관"); }
            if (pr.단계 && pr.단계 !== "고객등록" && STAGE_NEW_OLD[pr.단계] != null) setv(C, r, iBiz, STAGE_NEW_OLD[pr.단계]);
          }
          if (cls === "직접모집" && !insOnly && !txt(r[iBiz])) setv(C, r, iBiz, "계약"); // 영업단계 칩의 '계약': 옛 영업단계 칸이 없어져도 고객분류가 직접모집이면 계약으로 읽음(2026-11-17 사용자 "추천대로")
          setv(C, r, iStage, base);
        }
      });
    }
    // 계약현황(옛 칸이 없을 때): 이 고객이 계약자인 유효 계약(유지·완납·납입면제·연금지급·신계약승낙)이 있으면 '유효'
    if (C && SS["계약"] && C.columns.indexOf("계약현황") < 0) {
      const T0 = SS["계약"], m0 = ix(T0), iOwn = m0["계약자ID"], iSt0 = m0["계약상태"] != null ? m0["계약상태"] : m0["상태"], valid = new Set();
      if (iOwn != null && iSt0 != null) T0.rows.forEach((r) => { if (["유지", "완납", "납입면제", "연금지급", "신계약승낙"].includes(txt(r[iSt0]))) valid.add(txt(r[iOwn])); });
      const iIdc2 = C.columns.indexOf("고객ID"), iCur = ensure(C, "계약현황");
      if (iIdc2 >= 0) C.rows.forEach((r) => { if (valid.has(txt(r[iIdc2]))) setv(C, r, iCur, "유효"); });
    }
    // 계약 시트: 납입완료 예정 = 최초계약일 + 납입기간(년)(칸 없이 계산) / 계약분류 → 모집유형 / 이관상태 → 이관보관(날짜)
    const T = SS["계약"];
    if (T) {
      { const m0 = ix(T), iFirst = m0["최초계약일"], iPay = m0["납입기간"];
        if (iFirst != null && iPay != null) { const iDone = ensure(T, "납입완료 예정"); T.rows.forEach((r) => { if (txt(r[iDone])) return; const d0 = /^(\d{4})-(\d\d)-\d\d/.exec(txt(r[iFirst])), n0 = /^(\d+)\s*\(?\s*년/.exec(txt(r[iPay])); if (d0 && n0) setv(T, r, iDone, (+d0[1] + +n0[1]) + "-" + d0[2]); }); } }
      const m = ix(T), iKind = m["계약분류"], iSt = m["이관상태"];
      if (iKind != null || iSt != null) {
        const iType = ensure(T, "모집유형"), iArch = ensure(T, "이관보관"), iId = m["계약ID"], iIn = m["입력일시"];
        T.rows.forEach((r) => {
          if (iKind != null) setv(T, r, iType, txt(r[iKind]) || "타인모집(미정)"); // 계약분류가 빈칸(정해지지 않은 계약)이면 직접모집으로 세지 않음 → '타인모집…'으로 읽어 이관도 직접모집도 아닌 칸에 둠
          if (iSt != null) { const st = txt(r[iSt]); if (st === "보관" || st === "삭제") { const id = txt(r[iId]); setv(T, r, iArch, archAt((h) => h.ct === id) || (iIn != null ? day(r[iIn]) : "") || "보관"); } else if (st === "목표" || st === "") setv(T, r, iArch, ""); }
        });
      }
    }
    // 따행 시트: 따행상태 → 따행보관(날짜)
    const D = SS["따행"];
    if (D) {
      const m = ix(D), iSt = m["따행상태"];
      if (iSt != null) {
        const iArch = ensure(D, "따행보관"), iCust = m["고객ID"], iIn = m["입력일시"];
        D.rows.forEach((r) => { const st = txt(r[iSt]); if (st === "보관" || st === "삭제") setv(D, r, iArch, archAt((h) => h.cust === txt(r[iCust]) && !h.ct) || (iIn != null ? day(r[iIn]) : "") || "보관"); else setv(D, r, iArch, ""); });
      }
    }
    return { ...doc, sheets: SS };
  }

  function bookToDoc(wb, meta) {
    const sheets = {};
    wb.SheetNames.forEach((name) => {
      if (name.includes("(구)")) return;
      const ws = wb.Sheets[name]; if (!ws || !ws["!ref"]) { sheets[name] = { columns: [], rows: [] }; return; }
      const rg = XLSX.utils.decode_range(ws["!ref"]), val = (r, c) => {
        const x = ws[XLSX.utils.encode_cell({ r, c })];
        if (!x || x.v == null || x.v === "") return null;
        if (x.t === "n" && x.z && XLSX.SSF.is_date(x.z)) {
          const d = XLSX.SSF.parse_date_code(x.v);
          if (d) return isoText(d) + (d.H || d.M || Math.round(d.S) ? " " + String(d.H).padStart(2, "0") + ":" + String(d.M).padStart(2, "0") + ":" + String(Math.round(d.S) % 60).padStart(2, "0") : "");
        }
        if (x.t === "d" && x.v instanceof Date) return x.v.toISOString().slice(0, 10);
        return x.v;
      };
      const columns = [];
      for (let c = rg.s.c; c <= rg.e.c; c++) { const h = val(rg.s.r, c); columns.push(h == null ? "" : String(h).trim()); }
      while (columns.length && columns[columns.length - 1] === "") columns.pop();
      const rows = [];
      for (let r = rg.s.r + 1; r <= rg.e.r; r++) {
        const row = columns.map((_, i) => val(r, rg.s.c + i));
        if (row.some((v) => v != null)) rows.push(row);
      }
      sheets[name] = { columns, rows };
    });
    return { schema: 1, ...meta, made: new Date().toISOString().slice(0, 19), sheets };
  }
  // ===== 위치(ID) 기억(2026-11-12): 한 번 찾은 폴더·파일 ID를 이 기기에 저장해 두고 다음부터는 찾기 없이 쓴다. 틀리면(404) 기억을 지우고 다시 찾는다 =====
  const idcKey = () => "shlife.idc.v1." + (typeof me === "function" ? me() : "");
  let idcUsed = false;
  const idcAll = () => { try { return JSON.parse(localStorage.getItem(idcKey()) || "{}") || {}; } catch (e) { return {}; } };
  const idcGet = (k) => idcAll()[k];
  const idcSet = (k, v) => { try { const o = idcAll(); if (v == null) delete o[k]; else o[k] = v; localStorage.setItem(idcKey(), JSON.stringify(o)); } catch (e) { /* 저장 못 해도 동작은 같음 */ } };
  const idcClear = () => { try { localStorage.removeItem(idcKey()); } catch (e) { /* 무시 */ } idcUsed = false; };
  async function stFindPathCached(path) {
    const c = GD ? idcGet("path:" + path) : null;
    if (c) { idcUsed = true; return c; }
    const r = await stFindPath(path);
    if (r && GD) idcSet("path:" + path, r);
    return r;
  }
  // JSON 읽기: 같은 파일을 여러 곳에서 읽어도 한 번만 받고(쓰기 때까지), 읽기에는 버전 확인과 엑셀 변환 왕복을 하지 않는다
  const jsonDocP = new Map();
  function readJsonDoc(id) {
    if (!jsonDocP.has(id)) jsonDocP.set(id, stGet0(id).then(async (buf) => {
      const inf = jsonInfo.get(id), name = inf && inf.name;
      let doc = null, why = "";
      try { doc = parseDoc(buf); } catch (e) { why = "JSON\uC73C\uB85C \uC77D\uD788\uC9C0 \uC54A\uC74C"; }
      if (doc && name && JSON_FILES[name]) why = jsonDocBad(name, doc);
      if (why && name && JSON_FILES[name]) { // v509: \uAE68\uC84C\uC73C\uBA74 \uC790\uB3D9 \uBCF5\uAD6C \uC2DC\uB3C4, \uBABB \uD558\uBA74 \uC54C\uB9BC(\uC608\uC804\uCC98\uB7FC)
        const fixed = await fixBad(id, why);
        if (fixed) { doc = fixed; rowsWatch(id, name, doc); }
        else if (!doc) throw new Error(`\uB370\uC774\uD130 \uD30C\uC77C(${name})\uC744 \uC77D\uC9C0 \uBABB\uD588\uC2B5\uB2C8\uB2E4(${why}).`);
      } else if (doc && name) rowsWatch(id, name, doc);
      if (!doc) throw new Error("\uB370\uC774\uD130 \uD30C\uC77C\uC744 \uC77D\uC9C0 \uBABB\uD588\uC2B5\uB2C8\uB2E4.");
      return doc;
    }).catch((e) => { jsonDocP.delete(id); throw e; }));
    return jsonDocP.get(id);
  }
  // v508(2026-11-20 최수진 data.json 사고): 깨진 구조의 JSON은 저장하지 않고, 읽을 때 알림
  //   사고 경위: 앱을 연 직후(파일 목록이 오기 전) 전화·문자 기록 반영이 data.json을 엑셀로 읽음 → JSON 글자가 'Sheet1' 한 줄로 바뀜 → 그사이 목록이 와서 JSON 저장 경로로 덮어씀
  const looksJson = (buf) => { const b = new Uint8Array(buf instanceof ArrayBuffer ? buf : (buf && buf.buffer) || []).subarray(0, 8); let i = 0; if (b[0] === 0xef && b[1] === 0xbb && b[2] === 0xbf) i = 3; while (i < b.length && (b[i] === 0x20 || b[i] === 0x0a || b[i] === 0x0d || b[i] === 0x09)) i++; return b[i] === 0x7b; };
  function jsonDocBad(name, doc) { // 문제가 있으면 이유 글, 없으면 ""
    const sh = doc && doc.sheets;
    if (!sh || typeof sh !== "object") return "시트가 없음";
    const names = Object.keys(sh);
    if (!names.length) return "시트가 없음";
    if (names.some((n) => /^Sheet\d*$/i.test(n))) return "엑셀 기본 시트(Sheet1)가 들어 있음";
    if (name === "data.json" && !sh["고객"]) return "고객 시트가 없음";
    if (name === "data.json" && !doc.fc) return "FC 이름(fc)이 없음";
    if (name === "고객추가정보.json" && !sh["기본"]) return "기본 시트가 없음";
    for (const n of names) { const cols = (sh[n] && sh[n].columns) || []; if (cols.some((c) => typeof c === "string" && (c.length > 80 || /["{}]/.test(c)))) return n + " 시트의 칸 이름이 이상함"; }
    return "";
  }
  const jsonRows = (doc) => Object.values((doc && doc.sheets) || {}).reduce((a, x) => a + ((x && x.rows && x.rows.length) || 0), 0);
  const jsonWarned = new Set();
  function jsonWarnBad(name, doc) {
    const why = name && JSON_FILES[name] ? jsonDocBad(name, doc) : "";
    if (!why) return;
    console.error("데이터 파일 구조 이상", name, why);
    const k = name + why; if (jsonWarned.has(k)) return; jsonWarned.add(k);
    setTimeout(() => { try { toast(`데이터 파일(${name}) 하나가 깨져 있어 그 고객을 불러오지 못했습니다(${why}). 관리자에게 알려 주세요.`); } catch (e) { /* 무시 */ } }, 500);
  }
  const jsonNot = new Set(); // 이름을 확인해 보니 데이터 파일이 아닌 ID(설정 JSON 등)
  async function jsonAdopt(id) { // 목록을 거치지 않은 파일 ID가 data.json·고객추가정보.json이면 JSON 파일로 등록
    if (!GD || jsonNot.has(id)) return false;
    const m = await gdFetch(GDAPI + "/files/" + encodeURIComponent(id) + "?fields=name,parents,modifiedTime&supportsAllDrives=true");
    if (!m || !JSON_FILES[m.name]) { jsonNot.add(id); return false; }
    if (!jsonInfo.get(id)) jsonInfo.set(id, { name: m.name, meta: null, ver: null, mod: m.modifiedTime || "", parent: (m.parents || [])[0] || "" });
    return true;
  }
  // ===== v509(2026-11-21) 자동 백업·자동 복구 — 기준은 app\_공통\데이터파일_백업복구.md(데이터수집 PC 쪽과 같은 기준) =====
  //   ① 저장 직전: 바로 앞 내용을 같은 폴더의 _앱백업\<파일>.YYYYMMDD-HHMMSS.저장전 으로 서버 복사(계정별 최근 10개)
  //   ② 깨진 파일(jsonDocBad)을 읽으면: .반영중.lock·10분 안 된 .복구중.lock이 없을 때만 .복구중.lock을 걸고
  //      깨진 내용을 _앱백업\…깨진본 으로 남긴 뒤 Drive 버전 기록(없으면 저장전 사본)에서 마지막 정상본으로 되돌림
  //   ③ 되돌린 뒤 이 기기에서 보낸 전화·문자 기록(최근 14일)을 다시 반영(이력에 이미 있으면 건너뜀)
  //   줄 수만 크게 준 것은 되돌리지 않고 알리기만
  const BK_DIR = "_앱백업", BK_KEEP = 10, LOCK_FIX = ".복구중.lock", LOCK_SYNC = ".반영중.lock", LOCK_TTL = 10 * 60 * 1000;
  const LOCK_COLLECT = ".수집중.lock", COLLECT_TTL = 30 * 60 * 1000; // v550(2026-11-22 사용자 "이 안대로"): PC가 SOL 수집·반영하는 동안. 10분마다 갱신, 30분 넘게 갱신 없으면 무시
  const tsName = (d) => { d = d || new Date(); const p = (n) => String(n).padStart(2, "0"); return d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate()) + "-" + p(d.getHours()) + p(d.getMinutes()) + p(d.getSeconds()); };
  const bkDirP = new Map(); // 부모 폴더 ID → _앱백업 폴더 ID(Promise)
  function bkDir(parent) {
    if (!bkDirP.has(parent)) bkDirP.set(parent, (async () => {
      const ex = await gdFindIn(parent, BK_DIR);
      if (ex) return ex.id;
      const f = await gdFetch(GDAPI + "/files?supportsAllDrives=true&fields=id", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ name: BK_DIR, mimeType: GD_FOLDER, parents: [parent] }) });
      return f.id;
    })().catch((e) => { bkDirP.delete(parent); throw e; }));
    return bkDirP.get(parent);
  }
  async function bkCopy(id, info, kind) { // 지금 Drive에 있는 내용을 _앱백업으로 서버 복사(기기로 받지 않음) → 새 파일 ID
    const dir = await bkDir(info.parent);
    const f = await gdFetch(GDAPI + "/files/" + encodeURIComponent(id) + "/copy?supportsAllDrives=true&fields=id", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ name: info.name + "." + tsName() + "." + kind, parents: [dir], appProperties: { shlifeBackup: kind, shlifeOf: id } }) });
    return f.id;
  }
  async function bkPrune(info) { // 내 계정이 만든 저장전 사본 중 최근 BK_KEEP개만 남기고 휴지통으로
    const dir = await bkDir(info.parent);
    const d = await gdFetch(GDAPI + "/files?pageSize=200&supportsAllDrives=true&includeItemsFromAllDrives=true&fields=files(id,name)&q=" + encodeURIComponent("'" + qv(dir) + "' in parents and 'me' in owners and trashed=false and name contains '" + qv(info.name) + ".'"));
    const mine = (d.files || []).filter((f) => f.name.startsWith(info.name + ".") && f.name.endsWith(".저장전")).sort((a, b) => (a.name < b.name ? 1 : -1));
    for (const f of mine.slice(BK_KEEP)) await stDel(f.id).catch(() => {});
  }
  async function lockList(parent) { // 같은 폴더의 잠금 파일(.반영중.lock·.복구중.lock)
    const d = await gdFetch(GDAPI + "/files?pageSize=50&supportsAllDrives=true&includeItemsFromAllDrives=true&fields=files(id,name,createdTime,modifiedTime)&q=" + encodeURIComponent("'" + qv(parent) + "' in parents and trashed=false and (name contains '.반영중' or name contains '.복구중' or name contains '.수집중')"));
    const now = Date.now(); // 동기화가 '.복구중 (1).lock'처럼 이름을 바꿔도 같은 잠금으로 봄(데이터수집과 합의)
    const kind = (n) => (/^\.반영중.*\.lock$/.test(n) ? LOCK_SYNC : /^\.복구중.*\.lock$/.test(n) ? LOCK_FIX : /^\.수집중.*\.lock$/.test(n) ? LOCK_COLLECT : "");
    return (d.files || []).filter((f) => kind(f.name)).map((f) => { const k = kind(f.name); return { ...f, name: k, fresh: k === LOCK_COLLECT ? now - new Date(f.modifiedTime || f.createdTime).getTime() < COLLECT_TTL : now - new Date(f.createdTime || f.modifiedTime).getTime() < LOCK_TTL }; });
  }
  const lockBlocks = (ls) => ls.find((f) => f.name === LOCK_SYNC) || ls.find((f) => f.name === LOCK_COLLECT && f.fresh) || ls.find((f) => f.name === LOCK_FIX && f.fresh) || null;
  async function lockMake(parent, info, id) {
    const b = "shlife" + Date.now();
    const meta = { name: LOCK_FIX, parents: [parent], mimeType: "application/json" };
    const body = new Blob(["--" + b + "\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n" + JSON.stringify(meta) + "\r\n--" + b + "\r\nContent-Type: application/json\r\n\r\n" +
      JSON.stringify({ by: "앱", account: me(), file: info.name, fileId: id, at: new Date().toISOString() }) + "\r\n--" + b + "--"]);
    return gdFetch(GDUP + "/files?uploadType=multipart&supportsAllDrives=true&fields=id,createdTime", { method: "POST", headers: { "Content-Type": "multipart/related; boundary=" + b }, body });
  }
  async function lockDrop(lid) {
    try { await gdFetch(GDAPI + "/files/" + encodeURIComponent(lid) + "?supportsAllDrives=true", { method: "DELETE" }); }
    catch (e) { await stDel(lid).catch(() => {}); }
  }
  const parseDoc = (buf) => JSON.parse((typeof buf === "string" ? buf : new TextDecoder().decode(buf)).replace(/^﻿/, ""));
  async function findGood(id, info) { // 마지막 정상본 { text, at, src } — 버전 기록 먼저, 안 되면 _앱백업 저장전 사본
    let revErr = "";
    try {
      const d = await gdFetch(GDAPI + "/files/" + encodeURIComponent(id) + "/revisions?pageSize=1000&fields=revisions(id,modifiedTime)");
      const rv = (d.revisions || []).slice(0, -1).reverse().slice(0, 20); // 맨 끝(지금 깨진 것) 빼고 최근부터
      for (const r of rv) {
        try {
          const buf = await gdFetch(GDAPI + "/files/" + encodeURIComponent(id) + "/revisions/" + encodeURIComponent(r.id) + "?alt=media", null, true);
          const doc = parseDoc(buf);
          if (!jsonDocBad(info.name, doc)) return { buf, doc, at: r.modifiedTime, src: "버전 기록" };
        } catch (e) { /* 다음 버전 */ }
      }
    } catch (e) { revErr = String(e.message || e).slice(0, 60); console.warn("버전 기록을 읽지 못함", revErr); }
    try {
      const dir = await bkDir(info.parent);
      const d = await gdFetch(GDAPI + "/files?pageSize=200&supportsAllDrives=true&includeItemsFromAllDrives=true&fields=files(id,name,createdTime)&q=" + encodeURIComponent("'" + qv(dir) + "' in parents and trashed=false and name contains '" + qv(info.name) + ".'"));
      const bks = (d.files || []).filter((f) => f.name.startsWith(info.name + ".") && f.name.endsWith(".저장전")).sort((a, b) => (a.name < b.name ? 1 : -1)).slice(0, 10);
      for (const f of bks) {
        try { const buf = await stGet0(f.id); const doc = parseDoc(buf); if (!jsonDocBad(info.name, doc)) return { buf, doc, at: f.createdTime, src: "저장전 사본" }; } catch (e) { /* 다음 사본 */ }
      }
    } catch (e) { /* 사본 폴더 없음 */ }
    return { fail: revErr ? "버전 기록을 읽을 권한이 없고 정상 사본도 없음" : "정상본을 찾지 못함" };
  }
  const fixTried = new Set();
  async function jsonRestore(id) { // → { doc, at, src } 되돌림 / { doc, already } 누가 이미 고침 / { skip|fail: 이유 }
    const info = jsonInfo.get(id);
    if (!GD || !info || !info.parent) return { fail: "파일 위치를 모름" };
    if (fixTried.has(id)) return { skip: "이번에 이미 시도함" };
    fixTried.add(id);
    const blk = lockBlocks(await lockList(info.parent));
    if (blk) return { skip: blk.name === LOCK_SYNC ? "데이터 반영 중" : "다른 곳에서 복구 중" };
    const lock = await lockMake(info.parent, info, id);
    try {
      const fixes = (await lockList(info.parent)).filter((f) => f.name === LOCK_FIX && f.fresh).sort((a, b) => (a.createdTime < b.createdTime ? -1 : a.createdTime > b.createdTime ? 1 : a.id < b.id ? -1 : 1));
      if (fixes.length && fixes[0].id !== lock.id) return { skip: "다른 곳에서 복구 중" };
      const headRev = () => gdFetch(GDAPI + "/files/" + encodeURIComponent(id) + "?fields=headRevisionId&supportsAllDrives=true").then((x) => x.headRevisionId); // 내용이 바뀌었는지(복사·메타데이터로는 안 바뀜)
      const [v0, buf0] = await Promise.all([headRev(), stGet0(id)]);
      let cur = null; try { cur = parseDoc(buf0); } catch (e) { /* JSON도 아님 */ }
      if (cur && !jsonDocBad(info.name, cur)) return { doc: cur, already: true };
      await bkCopy(id, info, "깨진본"); // 깨진 내용은 지우지 않고 남김(못 남기면 되돌리지 않음)
      const good = await findGood(id, info);
      if (!good.doc) return good;
      if ((await headRev()) !== v0) return { skip: "그사이 다른 곳에서 바뀜" };
      await gdFetch(GDUP + "/files/" + encodeURIComponent(id) + "?uploadType=media&supportsAllDrives=true&fields=id", { method: "PATCH", headers: { "Content-Type": "application/json" }, body: new Blob([good.buf], { type: "application/json" }) });
      await gdFetch(GDAPI + "/files/" + encodeURIComponent(id) + "?supportsAllDrives=true&fields=id", { method: "PATCH", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ appProperties: { shlifeRestoredAt: new Date().toISOString(), shlifeRestoredBy: "앱", shlifeRestoredFrom: good.src + " " + (good.at || "") } }) }).catch(() => {});
      info.ver = null; info.rows = undefined;
      return { doc: good.doc, at: good.at, src: good.src };
    } finally { await lockDrop(lock.id); }
  }
  // 되돌린 뒤: 이 기기에서 보낸 전화·문자 기록을 다시 반영(writeContacts는 이력에 같은 기록이 있으면 건너뜀)
  const sentKey = () => "shlife.contactSent." + me();
  const readSent = () => { try { return JSON.parse(localStorage.getItem(sentKey()) || "[]"); } catch (e) { return []; } };
  function addSent(items) {
    try {
      const cut = Date.now() - 14 * 864e5;
      const all = readSent().concat(items.map((i) => ({ ...i, sentAt: Date.now() }))).filter((i) => i.sentAt > cut).slice(-500);
      localStorage.setItem(sentKey(), JSON.stringify(all));
    } catch (e) { /* 기기 저장 불가 — 다시 반영만 못 함 */ }
  }
  function reapplySent(id, sinceIso) {
    const since = sinceIso ? new Date(sinceIso).getTime() - 60000 : 0;
    const items = readSent().filter((i) => i.fileId === id && (i.sentAt || 0) >= since);
    if (!items.length) return;
    setTimeout(async () => {
      try { await writeContacts(id, items); toast(`이 기기에서 보낸 전화·문자 기록 ${items.length}건을 다시 반영했습니다.`); }
      catch (e) { console.error("전화·문자 기록 다시 반영 실패", e); }
    }, 2500);
  }
  function afterRestore(id, r) {
    const info = jsonInfo.get(id) || {};
    if (r.already) return;
    setTimeout(() => { try { toast(`데이터 파일(${info.name})이 깨져 있어 ${stampText(r.at) || "이전"} 정상본으로 자동으로 되돌렸습니다(${r.src}). 깨진 내용은 ${BK_DIR} 폴더에 남겼습니다.`); } catch (e) { /* 무시 */ } }, 500);
    reapplySent(id, r.at);
  }
  async function fixBad(id, why) { // 깨진 파일 → 되돌린 정상 doc 또는 null(못 되돌림 — 이유는 알림에)
    let r;
    try { r = await jsonRestore(id); } catch (e) { r = { fail: String(e.message || e).slice(0, 80) }; }
    if (r && r.doc) { afterRestore(id, r); return r.doc; }
    console.error("자동 복구 못 함", why, r);
    const info = jsonInfo.get(id) || {};
    const k = id + "fix"; if (!jsonWarned.has(k)) { jsonWarned.add(k); setTimeout(() => { try { toast(`데이터 파일(${info.name}) 하나가 깨져 있어 그 고객을 불러오지 못했습니다(${why}). 자동 복구도 못 했습니다(${(r && (r.skip || r.fail)) || "알 수 없음"}). 관리자에게 알려 주세요.`); } catch (e) { /* 무시 */ } }, 500); }
    return null;
  }
  // 줄 수 경고: 이 기기가 지난번에 읽은 줄 수보다 절반 아래로 줄었으면 알리기만(되돌리지 않음)
  function rowsWatch(id, name, doc) {
    try {
      const k = "shlife.jsonRows." + me(), o = JSON.parse(localStorage.getItem(k) || "{}") || {}, n = jsonRows(doc), was = o[id] || 0;
      if (was > 20 && n < was * 0.5 && !jsonWarned.has(id + "rows")) { jsonWarned.add(id + "rows"); setTimeout(() => { try { toast(`데이터 파일(${name})의 줄이 ${was}줄에서 ${n}줄로 크게 줄었습니다. 이상하면 관리자에게 알려 주세요.`); } catch (e) { /* 무시 */ } }, 800); }
      o[id] = n; localStorage.setItem(k, JSON.stringify(o));
    } catch (e) { /* 무시 */ }
  }
  if (/^(127\.0\.0\.1|localhost)$/.test(location.hostname)) window.__shlifeT = { jsonAdopt, readJsonDoc: (id) => readJsonDoc(id), jsonGetDoc: (id) => jsonGetDoc(id), jsonPutDoc: (id, d) => jsonPutDoc(id, d), jsonInfo, addSent, readSent }; // 로컬 시험용
  // 데이터 갱신 시각: 읽은 FC별 data.json의 Drive 수정시각(목록 조회 때 이미 받은 값). 여러 FC면 가장 오래된 것 기준(자세히는 길게 눌러/마우스를 올려 FC별로 표시)
  const stampText = (iso) => { const d = new Date(iso); return isNaN(d) ? "" : d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0") + " " + String(d.getHours()).padStart(2, "0") + ":" + String(d.getMinutes()).padStart(2, "0"); };
  function renderDataStamp() {
    const es = [...jsonInfo.values()].filter((x) => x.name === "data.json" && x.mod).sort((a, b) => (a.mod < b.mod ? -1 : 1));
    const txt = es.length ? "데이터 " + stampText(es[0].mod) : "";
    const tip = es.map((x) => ((x.meta && x.meta.fc) || "FC") + " " + stampText(x.mod)).join(" · ");
    document.querySelectorAll(".data-stamp").forEach((el) => { el.textContent = txt; el.title = tip; });
  }
  async function stList(id) {
    const items = await stList0(id);
    if (!GD) return items;
    const virt = [];
    items.forEach((i) => {
      const suffix = i.file && JSON_FILES[i.name];
      if (!suffix) return;
      const old = jsonInfo.get(i.id); // v508: 다시 목록을 받아도 저장 중인 버전·줄 수 기억은 지우지 않음
      jsonInfo.set(i.id, { name: i.name, meta: old ? old.meta : null, ver: old ? old.ver : null, rows: old ? old.rows : undefined, mod: i.lastModifiedDateTime || "", parent: id });
      virt.push({ ...i, name: "99999999" + (i.name === "data.json" ? " - JSON" : "") + suffix });
    });
    return virt.length ? virt.concat(items.filter((i) => !(i.file && JSON_FILES[i.name]))) : items;
  }
  // JSON 파일 저장 속도(2026-11-15): 버전 확인·내용 받기, 버전 확인·잠금 확인을 동시에 하고, 올린 응답에서 새 버전을 받아 마지막 확인 요청을 없앰
  const jsonVer = (id) => gdFetch(GDAPI + "/files/" + encodeURIComponent(id) + "?fields=version&supportsAllDrives=true");
  async function jsonGetDoc(id) { // 저장용 최신 JSON(버전 기억)
    const info = jsonInfo.get(id);
    const [m, buf] = await Promise.all([jsonVer(id), stGet0(id)]);
    let doc = null, bad = "";
    try { doc = parseDoc(buf); bad = jsonDocBad(info.name, doc); } catch (e) { bad = "JSON으로 읽히지 않음"; }
    if (bad) { // v509: 깨졌으면 자동 복구 후 다시 읽기(한 번만)
      if (!fixTried.has(id) && (await fixBad(id, bad))) { jsonDocP.delete(id); return jsonGetDoc(id); }
      throw new Error(`데이터 파일(${info.name})이 깨져 있어 저장하지 않았습니다(${bad}). 관리자에게 알려 주세요.`);
    }
    info.ver = m.version; info.meta = { fc: doc.fc, source: doc.source }; info.rows = jsonRows(doc);
    return doc;
  }
  async function jsonPutDoc(id, doc) {
    const info = jsonInfo.get(id);
    const bad = jsonDocBad(info.name, doc);
    if (bad) throw new Error(`데이터 파일(${info.name})을 잘못된 모양으로 저장하려 해서 멈췄습니다(${bad}). 새로고침한 뒤 다시 해 주세요.`);
    const nRows = jsonRows(doc);
    if (info.rows > 20 && nRows < info.rows * 0.5) throw new Error(`데이터 줄이 ${info.rows}줄에서 ${nRows}줄로 크게 줄어 저장하지 않았습니다. 새로고침한 뒤 다시 해 주세요.`);
    const [m, locks] = await Promise.all([jsonVer(id), info.parent ? lockList(info.parent) : []]);
    if (info.ver != null && String(m.version) !== String(info.ver)) throw new Error("다른 곳에서 먼저 수정되었습니다. 새로고침한 뒤 다시 시도해 주세요.");
    // 데이터 반영(수집) 중에는 FC 폴더에 .반영중.lock, 복구 중에는 .복구중.lock(10분 안 된 것)이 있다 → 저장 직전에만 확인해 막는다(읽기는 그대로)
    const blk = lockBlocks(locks.filter((f) => f.name === LOCK_FIX || info.name === "data.json"));
    if (blk) { const er = new Error(blk.name === LOCK_SYNC ? "데이터 반영 중이라 저장하지 못했습니다. 잠시 후 다시 시도해 주세요." : blk.name === LOCK_COLLECT ? "데이터 업데이트 중이라 지금은 수정할 수 없습니다. 업데이트가 끝난 뒤 다시 해 주세요." : "데이터 파일을 복구하는 중이라 저장하지 못했습니다. 잠시 후 다시 시도해 주세요."); er.lockKind = blk.name; if (info.parent) syncMark(info.parent, blk.name === LOCK_COLLECT); throw er; }
    // v509: 저장 직전 사본(바로 앞 내용, 서버 복사). 사본을 못 만들어도 저장은 계속(버전 기록이 남음)
    if (info.parent) { try { await bkCopy(id, info, "저장전"); bkPrune(info).catch(() => {}); } catch (e) { console.warn("저장 전 사본을 만들지 못함", e); } }
    const raw = await gdFetch(GDUP + "/files/" + encodeURIComponent(id) + "?uploadType=media&supportsAllDrives=true&fields=" + GD_FIELDS + ",version",
      { method: "PATCH", headers: { "Content-Type": "application/json" }, body: new Blob([JSON.stringify(doc)], { type: "application/json" }) });
    const out = gdNorm(raw);
    jsonDocP.delete(id);
    if (out && out.lastModifiedDateTime) { info.mod = out.lastModifiedDateTime; renderDataStamp(); } // 저장하면 데이터 시각도 바로 갱신
    info.ver = raw && raw.version != null ? raw.version : (await jsonVer(id)).version; info.rows = nRows;
    return out;
  }
  async function stGet(id) {
    if (!jsonInfo.get(id)) { const buf = await stGet0(id); if (!looksJson(buf) || !(await jsonAdopt(id))) return buf; }
    return XLSX.write(jsonToBook(await jsonGetDoc(id)), { type: "array", bookType: "xlsx" });
  }
  async function stPut(id, data, mime) {
    const info = jsonInfo.get(id);
    if (!info) return stPut0(id, data, mime);
    return jsonPutDoc(id, bookToDoc(XLSX.read(data, { type: "array", cellNF: true }), info.meta || {}));
  }
  // 저장용 통합문서 읽기/쓰기: JSON이면 엑셀 파일로 바꿨다 다시 읽는 왕복 없이 시트 모양으로 바로
  async function stGetBook(id) {
    if (jsonInfo.get(id)) return jsonToBook(await jsonGetDoc(id));
    const buf = await stGet0(id);
    if (looksJson(buf)) { // v508: JSON 글자를 엑셀로 읽지 않음
      if (await jsonAdopt(id)) return jsonToBook(await jsonGetDoc(id));
      throw new Error("데이터 파일을 엑셀로 읽을 수 없습니다. 새로고침한 뒤 다시 해 주세요.");
    }
    return XLSX.read(buf, { type: "array", cellNF: true });
  }
  async function stPutBook(id, wb) {
    let info = jsonInfo.get(id);
    if (!info && (await jsonAdopt(id))) info = jsonInfo.get(id);
    if (info && !info.rows && info.ver == null) throw new Error("데이터 파일을 다시 읽은 뒤 저장해야 합니다. 새로고침한 뒤 다시 해 주세요.");
    if (info) return jsonPutDoc(id, bookToDoc(wb, info.meta || {}));
    return stPut0(id, XLSX.write(wb, { type: "array", bookType: "xlsx", compression: true }), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  }
  // 폴더(dir = {id, driveId?}) 안에 이름으로 저장 (같은 이름이 있으면 그 파일을 고침)
  async function stPutIn(dir, name, data, mime) {
    if (!GD) return gx("PUT", "/drives/" + (dir.driveId || base.driveId) + "/items/" + dir.id + ":/" + encodeURIComponent(name) + ":/content", data, mime);
    const ex = await gdFindIn(dir.id, name);
    if (ex) return stPut(ex.id, data, mime);
    const b = "shlife" + Date.now();
    const body = new Blob([
      "--" + b + "\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n" + JSON.stringify({ name, parents: [dir.id] }) +
      "\r\n--" + b + "\r\nContent-Type: " + mime + "\r\n\r\n", data, "\r\n--" + b + "--"]);
    return gdNorm(await gdFetch(GDUP + "/files?uploadType=multipart&supportsAllDrives=true&fields=" + GD_FIELDS,
      { method: "POST", headers: { "Content-Type": "multipart/related; boundary=" + b }, body }));
  }
  // 삭제 (휴지통으로)
  async function stDel(id) {
    if (!GD) return gx("DELETE", itemPath(id));
    return gdFetch(GDAPI + "/files/" + encodeURIComponent(id) + "?supportsAllDrives=true&fields=id",
      { method: "PATCH", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ trashed: true }) });
  }
  // 내 드라이브 기준 폴더 경로("Claude/고객") → { id, driveId } / 없으면 null
  async function stFindPath(path) {
    if (!GD) {
      try {
        const it = await graph("/me/drive/root:/" + path.split("/").map(encodeURIComponent).join("/") + "?$select=id,parentReference");
        return { id: it.id, driveId: it.parentReference.driveId };
      } catch (e) {
        if (String(e.message).includes("404")) return null;
        throw e;
      }
    }
    let id = "root";
    for (const part of path.split("/").filter(Boolean)) {
      const r = await gdQuery("'" + qv(id) + "' in parents and name='" + qv(part) + "' and mimeType='" + GD_FOLDER + "' and trashed=false");
      if (!r.length) return null;
      id = r[0].id;
    }
    return { id };
  }
  // 폴더 안 JSON 파일 읽기 / 쓰기 (없으면 null)
  async function stReadJsonIn(dir, name) {
    if (!dir || !dir.id) return null;
    if (!GD) {
      try { return await graph("/drives/" + dir.driveId + "/items/" + dir.id + ":/" + encodeURIComponent(name) + ":/content?x=" + Date.now()); } catch (e) {
        if (String(e.message).includes("404")) return null;
        throw e;
      }
    }
    const fk = "file:" + dir.id + "/" + name, cid = idcGet(fk);
    if (cid) { try { return JSON.parse(new TextDecoder().decode(await stGet(cid)).replace(/^﻿/, "")); } catch (e) { idcSet(fk, null); } } // 기억한 위치가 틀리면 찾기로
    const f = await gdFindIn(dir.id, name);
    if (!f) return null;
    idcSet(fk, f.id);
    return JSON.parse(new TextDecoder().decode(await stGet(f.id)).replace(/^﻿/, ""));
  }
  async function stWriteJsonIn(dir, name, obj) {
    return stPutIn(dir, name, new TextEncoder().encode(JSON.stringify(obj, null, 2)), "application/json");
  }

  // ---------- 다크 모드 선택 (v408, 2026-11-17 사용자 지시) ----------
  // 설정 > 화면의 [다크 모드] 체크. 고르지 않으면 폰 설정(시스템)을 따르고, 고르면 이 기기에서 그 모양으로 고정. 스타일의 '다크' 미디어 규칙을 켜고 끄는 방식이라 모든 화면에 함께 적용
  const THEME_KEY = "shlife.theme";
  // v428(2026-11-18 말만해요 규칙): 말만해요 안(embed)에서는 주소의 &theme=dark|light 를 그대로 따르고(앱 자체 설정보다 우선, 저장하지 않음), 말만해요 설정에서 바꾸면 postMessage(blue10kp-theme)로 알려 줌
  let embedTheme = (() => { try { const q = new URLSearchParams(location.search), t = q.get("theme"); return q.get("embed") === "1" && (t === "dark" || t === "light") ? t : ""; } catch (e) { return ""; } })();
  const themePref = () => embedTheme || (() => { try { return localStorage.getItem(THEME_KEY) || "dark"; } catch (e) { return "dark"; } })(); // v414(2026-11-18 사용자 지시): 고른 적 없으면 다크 모드가 기본
  const darkMQ = (() => { try { return matchMedia("(prefers-color-scheme: dark)"); } catch (e) { return null; } })();
  const themeIsDark = () => { const p = themePref(); return p ? p === "dark" : !!(darkMQ && darkMQ.matches); };
  const themeOrig = window.__themeOrig || new WeakMap(); // index.html 머리에서 미리 적용하며 적어 둔 원래 값을 이어 씀
  function themeApply() {
    const p = themePref();
    if (p) document.documentElement.dataset.theme = p; else delete document.documentElement.dataset.theme;
    const walk = (rules) => {
      for (const r of rules) {
        if (r.media && r.cssRules) {
          if (!themeOrig.has(r)) themeOrig.set(r, r.media.mediaText);
          const o = themeOrig.get(r);
          if (/prefers-color-scheme:\s*dark/.test(o)) { try { r.media.mediaText = !p ? o : p === "dark" ? (o.replace(/\(\s*prefers-color-scheme:\s*dark\s*\)/g, "(min-width: 0px)")) : "not all"; } catch (e) { /* 무시 */ } }
          else if (/prefers-color-scheme:\s*light/.test(o)) { try { r.media.mediaText = !p ? o : p === "light" ? o.replace(/\(\s*prefers-color-scheme:\s*light\s*\)/g, "(min-width: 0px)") : "not all"; } catch (e) { /* 무시 */ } }
          walk(r.cssRules);
        }
      }
    };
    for (const sh of document.styleSheets) { let rules; try { rules = sh.cssRules; } catch (e) { continue; } if (rules) walk(rules); }
    const tc = document.querySelectorAll('meta[name="theme-color"]'); // 폰 상단 색도 맞춤
    tc.forEach((m) => { if (!m.dataset.media) m.dataset.media = m.getAttribute("media") || ""; const md = m.dataset.media; m.setAttribute("media", !p ? md : (p === "dark" ? /dark/.test(md) : /light/.test(md)) ? "all" : "not all"); });
  }
  themeApply();
  window.addEventListener("load", themeApply);
  document.addEventListener("change", (ev) => {
    if (!ev.target || ev.target.id !== "darkMode") return;
    const v = ev.target.checked ? "dark" : "light";
    if (embedTheme) { embedTheme = v; themeApply(); toast("말만해요 안에서는 말만해요 설정을 따릅니다(이번에만 바뀜)."); return; } // embed 중에는 저장하지 않음
    try { localStorage.setItem(THEME_KEY, v); } catch (e) { /* 이번에만 적용 */ }
    themeApply();
    toast(ev.target.checked ? "다크 모드로 바꿨습니다." : "라이트 모드로 바꿨습니다.");
    // v427(2026-11-18 프팀장 전달 사용자 지시): 기기 밖에서도 같게 — 내 설정 파일(Drive)에 저장하고, localStorage는 화면 깜빡임을 막는 빠른 캐시
    (async () => { try { const pr = { ...(await loadMyPrefs()) }; pr.theme = v; await saveMyPrefs(pr); } catch (e) { console.warn("theme save", e && e.message); } })();
  });
  // ---------- 화면 전환 ----------
  // 오른손 모드: 설정에서 켜면(이 폰에만 저장) 하단 메뉴·탭·선택 버튼을 오른쪽으로 모음 (모양은 design.css의 .rh)
  const RH_KEY = "shlife.rightHand";
  const rhOn = () => { try { return localStorage.getItem(RH_KEY) === "1"; } catch (e) { return false; } };
  const rhApply = () => document.documentElement.classList.toggle("rh", rhOn());
  rhApply();
  const rhOptHtml = () => `<div class="fld rh-fld">
        <label class="rh-opt"><span>오른손 모드</span><input type="checkbox" class="cb" id="rhMode" ${rhOn() ? "checked" : ""}></label>
      </div>`;
  // ---- 앱 설치(홈 화면에 추가): 설치할 수 있을 때 설정 화면에 [앱 설치] 버튼, 이미 설치됐으면 카드 숨김 ----
  let installEvt = null;
  const isStandalone = () => { try { return matchMedia("(display-mode: standalone)").matches || navigator.standalone === true; } catch (e) { return false; } };
  const installCardHtml = () => isStandalone() ? "" : `<div class="card" id="installCard"><h3 class="sec-title">앱 설치</h3><p class="hint">홈 화면에 앱처럼 추가합니다. 안드로이드 크롬은 아래 [앱 설치] 또는 브라우저 메뉴의 ‘앱 설치’, 아이폰 사파리는 공유 → ‘홈 화면에 추가’를 누르세요.</p><button type="button" class="secondary" id="installBtn"${installEvt ? "" : " hidden"}>앱 설치</button></div>`;
  window.addEventListener("beforeinstallprompt", (e) => { e.preventDefault(); installEvt = e; const b = document.getElementById("installBtn"); if (b) b.hidden = false; });
  window.addEventListener("appinstalled", () => { installEvt = null; const c = document.getElementById("installCard"); if (c) c.remove(); });
  document.addEventListener("click", async (ev) => {
    if (!ev.target.closest || !ev.target.closest("#installBtn") || !installEvt) return;
    const e = installEvt; installEvt = null; ev.target.closest("#installBtn").hidden = true;
    try { e.prompt(); await e.userChoice; } catch (er) { /* 무시 */ }
  });
  const rhCardHtml = () => { // v337(2026-11-15 사용자 지시): 화면 카드 3칸, 바꾸면 바로 저장
    const cur = readStartPref();
    return `<div class="card" id="startCard"><h3 class="sec-title">화면</h3><div class="disp3">
      <label class="d3"><span>시작하면 AI</span><span class="d3c"><input type="checkbox" class="cb" id="startAi" ${cur.ai ? "checked" : ""}> 바로 열기</span></label>
      <label class="d3"${START_FIXED ? " hidden" : ""}><span>시작 화면</span><select id="startPage" aria-label="시작 화면">${START_PAGES.map(([k, l]) => `<option value="${k}" ${cur.page === k ? "selected" : ""}>${l}</option>`).join("")}</select></label>
      <label class="d3"><span>오른손 모드</span><span class="d3c"><input type="checkbox" class="cb" id="rhMode" ${rhOn() ? "checked" : ""}> 켜기</span></label>
    </div>
    <label class="wake-opt"><input type="checkbox" class="cb" id="darkMode" ${themeIsDark() ? "checked" : ""}> <span><b>다크 모드</b> <small>기본은 다크, 끄면 밝은 화면(내 설정 파일에 저장되어 다른 기기에서도 같음)</small></span></label>
    <label class="wake-opt"><input type="checkbox" class="cb" id="wakeMode" ${wakeOn() ? "checked" : ""}> <span><b>“몰리야”로 AI 부르기</b> <small>시험용 · 앱 화면이 열려 있는 동안만 마이크로 듣습니다(안드로이드 크롬). 이 기기에만 저장</small></span></label></div>`;
  };
  document.addEventListener("change", (ev) => {
    if (!ev.target || ev.target.id !== "rhMode") return;
    try { localStorage.setItem(RH_KEY, ev.target.checked ? "1" : "0"); } catch (e) { /* 저장 못 해도 이번엔 적용 */ }
    rhApply();
    if (!isAdmin()) { const on = ev.target.checked; readPerms().then((perms) => readStartShared(perms, me(), perms.fcFolders).then((v) => writeStartShared(perms, me(), { ...(v || readStartPref()), rh: on }, perms.fcFolders))).catch(() => {}); } // 관리자 조직도에 보이게 (v333)
    toast(ev.target.checked ? "오른손 모드를 켰습니다." : "오른손 모드를 껐습니다.");
  });
  const views = ["loginView", "scheduleView", "searchView", "customerView", "salesView", "adminView", "introView", "dashView", "teamView"];
  const APP_NAME = "신한FC m";
  const TITLES = { scheduleView: "일정", searchView: "고객", customerView: "고객 정보", salesView: "세일즈", adminView: "설정", introView: "소개찾기", dashView: "대시보드", teamView: "팀관리" };
  const SECTION = { scheduleView: "schedule", searchView: "customer", customerView: "customer", salesView: "sales", adminView: "admin", dashView: "dash", introView: "customer", teamView: "team" };
  // 말만해요 안에 모듈(iframe)로 떠 있는 모드 (?embed=1)
  const EMBED = (() => { try { return new URLSearchParams(location.search).get("embed") === "1" && window.parent !== window; } catch (e) { return false; } })();
  let notifyWhereNow = () => {};
  if (EMBED) document.documentElement.classList.add("embed");
  function show(view, keepScroll) {
    notifyWhereNow();
    views.forEach((v) => ($(v).hidden = v !== view));
    document.documentElement.dataset.view = view; // 화면별 모양 조정용 (예: 일정 화면은 위 막대 밑 선 없음)
    // 고객 정보 화면: "‹ 고객 정보" 전체가 뒤로 가기 버튼
    $("backBtn").hidden = view !== "customerView";
    $("title").hidden = view === "customerView";
    $("logoutBtn").hidden = view === "loginView" || EMBED; // 모듈로 떠 있을 땐 로그아웃은 말만해요에서
    $("reloadBtn").hidden = view === "loginView";
    $("topSearch").hidden = view === "loginView"; // 상단 검색창은 로그인 후 항상
    $("title").textContent = TITLES[view] || APP_NAME;
    $("mainNav").hidden = view === "loginView";
    try { const ss = $("salesSub"); if (ss) { ss.hidden = view !== "salesView"; document.documentElement.style.setProperty("--subh", ss.hidden ? "0px" : Math.round(ss.getBoundingClientRect().height) + "px"); ss.querySelectorAll("[data-ssub]").forEach((b) => b.classList.toggle("on", b.dataset.ssub === (SALES.sub || "check"))); } } catch (e) { /* 세일즈 서브메뉴(이달목표·자료) 표시 */ }
    // 관리자: '관리자' 메뉴 / 일반 사용자: '설정' 메뉴 (내 캘린더 고르기)
    $("navAdmin").hidden = !account();
    { const nt = $("navTeam"); if (nt) nt.hidden = !isAdmin(); } // v442: 팀관리 메뉴는 관리자만
    $("navAdminLabel").textContent = "설정";
    if (view === "searchView") $("pageSearch").hidden = $("results").children.length > 0; // 보이는 고객 목록이 없으면 화면 안 검색창
    if (view !== "salesView" && view !== "customerView" && view !== "loginView") salesClearState(); // 세일즈 화면을 떠나면 이어보기 상태 삭제
    if (view !== "scheduleView" && view !== "customerView" && view !== "loginView") schedClearState(); // 일정 화면을 떠나도 마찬가지
    $("mainNav").querySelectorAll("button[data-section]").forEach((b) => b.classList.toggle("on", b.dataset.section === SECTION[view]));
    if (view === "loginView" && typeof aiClose === "function") aiClose(true); // 로그인 화면에서는 AI 도우미 없음
    document.documentElement.style.setProperty("--tbh", Math.round(topbarH()) + "px"); // 떠 있는 줄(일정 날짜·고객 수정/닫기)의 위치 기준
    { const nv = $("mainNav"); if (nv && !nv.hidden) document.documentElement.style.setProperty("--navh", Math.round(nv.getBoundingClientRect().height) + "px"); } // 떠 있는 저장 줄은 아래 메뉴 바로 위
    fsRefresh();
    if (!keepScroll) window.scrollTo(0, 0);
  }

  function toast(msg) {
    const el = $("toast");
    el.textContent = msg;
    el.hidden = false;
    clearTimeout(toast._t);
    toast._t = setTimeout(() => (el.hidden = true), 3500);
  }

  function esc(s) {
    return String(s ?? "").replace(/[&<>"']/g, (c) =>
      ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  }

  // ---------- 검색 ----------
  // 고객 폴더 위치 찾기
  //   소유자(관리자): 내 OneDrive의 Claude/고객
  //   다른 사용자: 관리자가 OneDrive에서 공유해 준 같은 '고객' 폴더 (config.owner)
  let base = null; // { driveId, id, owner }
  async function resolveBase() {
    if (base) return base;
    const mine = await stFindPathCached(CFG.baseFolder);
    if (mine) return (base = { driveId: mine.driveId, id: mine.id, owner: true });
    if (GD) {
      if (!GDCFG.baseId) throw new Error("NO_SHARE");
      return (base = { id: GDCFG.baseId, owner: false });
    }
    return (base = { driveId: CFG.owner.driveId, id: CFG.owner.baseId, owner: false });
  }
  const itemPath = (id) => "/drives/" + base.driveId + "/items/" + id;

  // ---------- 앱 자체 권한 (고객 폴더 안 권한 파일) ----------
  // { "users": { "someone@outlook.com": ["김현성"], "other@hotmail.com": ["*"] } }   "*" = 전체 FC
  const PERM_FILE = CFG.permFile || "shlife-권한.json";

  // ---------- 앱 저장 폴더: Claude/app/신한FC/설정 (설정·권한 파일) ----------
  //   소유자(관리자): 내 OneDrive 경로 / 다른 사용자: 관리자가 공유해 준 같은 폴더 (config.owner.appDirId)
  let appDir = null;
  let appDirP = null;
  function resolveAppDir() { if (appDir) return Promise.resolve(appDir); return appDirP || (appDirP = resolveAppDirRaw().finally(() => { appDirP = null; })); }
  async function resolveAppDirRaw() {
    if (appDir) return appDir;
    await resolveBase();
    // Google Drive: 공유받은 사용자는 관리자의 설정 폴더(ID)를 씀 (자기 드라이브에 같은 이름 폴더가 있어도)
    if (GD && !base.owner) return (appDir = { id: GDCFG.appDirId || "" });
    const mine = await stFindPathCached(CFG.appFolder);
    if (mine) return (appDir = mine);
    if (GD) return (appDir = { id: GDCFG.appDirId || "" });
    return (appDir = { driveId: CFG.owner.driveId, id: CFG.owner.appDirId });
  }

  // ---------- 관리자 설정 (앱 저장 폴더의 설정 파일, 모든 사용자가 같은 설정 사용) ----------
  //   { companyAddress, kakaoRestKey, google: { account, calendars: [] } }
  const SETTINGS_FILE = CFG.settingsFile || "shlife-설정.json";
  let appSettings = { companyAddress: "", google: { account: "", calendars: [] } };
  async function loadSettings() {
    try {
      await resolveAppDir();
      const s = await stReadJsonIn(appDir, SETTINGS_FILE);
      if (s) appSettings = { ...appSettings, ...s, google: { ...appSettings.google, ...(s.google || {}) } };
    } catch (e) { /* 설정 파일을 못 읽으면 기본값 */ }
    return appSettings;
  }
  async function saveSettings(next) {
    await resolveAppDir();
    await stWriteJsonIn(appDir, SETTINGS_FILE, next);
    appSettings = next;
  }
  // 세일즈 화면이 매번 권한 파일을 다시 읽지 않도록 'FC 이름(names)'만 세션 캐시. 권한 파일을 읽거나 쓸 때마다 갱신 (관리자 화면은 열 때 최신을 읽음). 접근 FC 판단(myAllowedFcs)은 캐시를 쓰지 않음
  let permNamesCache = null, permPrefixCache = null; // permPrefixCache: { FC이름: 일정 제목 앞 글자 }
  let salesExtraAt = 0;
  // 세일즈 화면을 열 때(1분에 한 번) 권한 파일을 다시 읽어, 관리자가 방금 바꾼 '추가 FC'도 앱을 다시 열지 않고 반영
  async function refreshSalesExtra() { if (Date.now() - salesExtraAt < 60000) return; salesExtraAt = Date.now(); try { await readPerms(); } catch (e) { /* 이전 값 유지 */ } }
  let salesPermsAll = null; // v325: 관리자 세일즈 화면의 사용자 선택(각 사용자의 FC + '세일즈 페이지 접근 허용')
  let salesExtra = []; // 세일즈(이달이관·따행) 화면의 FC 선택창에 추가로 보이는 FC (관리자 화면에서 계정마다 지정)
  // ---------- FC 이름 ↔ 사번 (v359) ----------
  // 기준표: Claude/_양식/사용자권한.json (users[{no,name,...}]) — 관리자만 읽을 수 있음(일반 사용자는 권한 파일의 fcNos 보조칸을 씀)
  // 못 읽거나 이름이 안 맞으면 빈 값을 돌려주고, 화면·권한 판단은 지금처럼 이름으로 동작(사번은 어디에도 새로 보이지 않음)
  let fcNoMap = {};
  async function loadFcTable() {
    if (!isAdmin()) return null;
    try {
      const dir = await stFindPathCached("Claude/_양식");
      const t = dir && await stReadJsonIn(dir, "사용자권한.json");
      const out = {};
      ((t && t.users) || []).forEach((u) => { const n = String((u && u.name) || "").trim(), no = String((u && u.no) || "").trim(); if (n && no) out[n] = no; });
      return Object.keys(out).length ? out : null;
    } catch (e) { return null; }
  }
  const fcNoOf = (name) => fcNoMap[String(name || "").trim()] || "";
  const fcNameOfNo = (no) => { const k = Object.keys(fcNoMap).find((n) => fcNoMap[n] === String(no || "").trim()); return k || ""; };
  async function fillFcNos(perms) { // 기준표에서 채움(권한 파일에 있는 FC 이름과 맞는 것만). 실패하면 기존 값 그대로
    try {
      const t = await loadFcTable(); if (!t) return;
      const known = new Set([...Object.keys(perms.fcFolders || {}), ...Object.values(perms.names || {}), ...Object.values(perms.users || {}).flat().map((x) => String(x).replace(/:r$/, ""))]);
      const next = { ...(perms.fcNos || {}) };
      Object.keys(t).forEach((n) => { if (known.has(n)) next[n] = t[n]; });
      perms.fcNos = next;
    } catch (e) { /* 무시 */ }
  }
  async function writePerms(perms) {
    await resolveAppDir();
    if (isAdmin()) await fillFcNos(perms);
    await stWriteJsonIn(appDir, PERM_FILE, perms);
    fcNoMap = { ...(perms.fcNos || {}) };
    permNamesCache = { ...(perms.names || {}) };
    permPrefixCache = { ...(perms.prefixes || {}) };
  }
  const me = () => String((account() || {}).username || "").toLowerCase();

  async function readPerms() {
    await resolveAppDir();
    const p = await stReadJsonIn(appDir, PERM_FILE);
    const out = { users: (p && p.users) || {}, fcFolders: (p && p.fcFolders) || {}, names: (p && p.names) || {}, prefixes: (p && p.prefixes && typeof p.prefixes === "object") ? p.prefixes : {}, sensitive: Array.isArray(p && p.sensitive) ? p.sensitive : [], salesFcs: (p && p.salesFcs && typeof p.salesFcs === "object") ? p.salesFcs : {}, starts: (p && p.starts && typeof p.starts === "object") ? p.starts : {}, org: (p && p.org && typeof p.org === "object") ? p.org : {}, orgApplied: (p && p.orgApplied && typeof p.orgApplied === "object") ? p.orgApplied : {}, orgSalesApplied: (p && p.orgSalesApplied && typeof p.orgSalesApplied === "object") ? p.orgSalesApplied : {}, orgVer: (p && p.orgVer) || 1, titles: (p && p.titles && typeof p.titles === "object") ? p.titles : {}, archived: (p && p.archived && typeof p.archived === "object") ? p.archived : {}, fcNos: (p && p.fcNos && typeof p.fcNos === "object") ? p.fcNos : {} }; // fcNos(v359): { FC이름: 사번 } 보조칸(이름 키는 그대로, 일반 사용자가 기준표를 못 읽을 때 사번을 알 수 있게) // archived(v343): { 이메일: {...} } 보관한(삭제한) 사용자 — 데이터는 지우지 않음 // titles(v339): { 이메일: "ESL"|"SSL"|"GSL"|"FC" } 직책 // orgVer 2(v335): 보기·수정 = 그 사용자가 자기 FC를 위 사용자들에게 주는 권한 // orgSalesApplied(v330): { 이메일: [FC] } 조직도가 세일즈 페이지 접근 허용에 더한 FC // orgApplied(v327): { 이메일: [{fc, prev}] } 조직도가 더한 권한(조직도에서 빼면 이것만 되돌림, 직접 준 권한은 그대로) // org(v326): { 이메일: { parent, view, edit } } 조직도 // salesFcs: { 이메일: [세일즈 화면에서 추가로 볼 수 있는 FC 이름] } / sensitive: 민감정보 접근을 허용한 이메일 목록 (다시 저장해도 지워지지 않게 읽어 둠)
    if (p && p.login && typeof p.login === "object") out.login = p.login; // v513: 로그인 승인/거부 토글(없으면 옛 파일 = 예전 방식)
    permNamesCache = { ...out.names };
    permPrefixCache = { ...out.prefixes };
    fcNoMap = { ...out.fcNos };
    try { const k = "shlife.startAdmin." + me(), v = out.starts[me()]; if (v) localStorage.setItem(k, JSON.stringify(v)); else localStorage.removeItem(k); } catch (e) { /* 무시 */ } // 관리자가 정한 시작 화면 (다음 실행부터)
    salesExtra = Array.isArray(out.salesFcs[me()]) ? out.salesFcs[me()].slice() : [];
    salesExtraAt = Date.now();
    salesPermsAll = out;
    return out;
  }
  // 권한 항목: "최수진"(보기+수정) / "최수진:r"(보기만) / "*"(전체 FC, 보기+수정)
  // names: { 이메일: "FC 이름" } — 관리자 화면에서 아이디와 이름을 짝지음(화면에는 이름을 먼저 보여 줌)
  const grantFc = (s) => String(s).replace(/:r$/, "");
  const grantRo = (s) => /:r$/.test(String(s));
  const canSens = () => isAdmin() || !!base.owner || sensOk; // 관리자는 항상, 그 밖의 계정은 권한 파일의 민감정보 허용 목록에 있을 때만
  let sensOk = false; // 민감정보(주민번호·계좌) 볼 수 있는 계정: 관리자 또는 관리자 화면에서 '민감정보 접근 허용'을 체크한 계정
  let roFcs = new Set(); // 내가 '보기만' 권한인 FC (화면에서 수정 버튼 숨김 — 실제 막는 것은 Drive 공유 권한)
  let fcFolderIds = {}; // FC 이름 → Drive 폴더 ID (관리자가 변경 저장할 때 권한 파일에 기록)
  const isRO = (fc) => !!fc && roFcs.has(fc);

  // 지금 로그인한 아이디가 볼 수 있는 FC: "*" 또는 FC 이름 목록
  async function myAllowedFcs() {
    if (isAdmin() || base.owner) { sensOk = true; return "*"; }
    const perms = await readPerms();
    sensOk = (perms.sensitive || []).includes(me());
    fcFolderIds = perms.fcFolders || {};
    const list = (perms.users[me()] || []).filter((g) => g !== "*"); // 전체("*")는 관리자 전용 — 일반 계정은 FC 이름 목록만 인정
    roFcs = new Set(list.filter(grantRo).map(grantFc));
    return list.map(grantFc);
  }

  // 고객 폴더 구조 읽기
  //   고객/김현성/_계약/<고객폴더>, 고객/김현성/계약고객리스트.xlsx   ← FC별 폴더 (현재 구조)
  //   고객/_계약/<고객폴더>, 고객/계약고객리스트.xlsx               ← FC 폴더 없이 (예전 구조)
  // "_"로 시작하는 폴더 = 분류(계약/이관/진행), 그 안의 폴더 = 고객 폴더
  let scanCache = null, scanFilesCache = null; // Promise<{ folders, files, fcs }>
  // scanFiles: 처음 화면에 필요한 '리스트 파일' 위치만(빠름). scanBase: 고객 폴더 목록(사진·파일용)까지(고객을 열 때만)
  function scanFiles() { return scanFilesCache || (scanFilesCache = scanRun(false).catch((e) => { scanFilesCache = null; throw e; })); }
  let scanAt = 0; // 고객 폴더 목록을 마지막으로 읽기 시작한 때
  function scanBase() { return scanCache || (scanAt = Date.now(), scanCache = scanRun(true).catch((e) => { scanCache = null; throw e; })); }
  async function scanRun(wantFolders) {
    try { return await scanOnce(wantFolders); } catch (e) {
      if (idcUsed && /40[34]/.test(String(e.message))) { idcClear(); base = null; appDir = null; return scanOnce(wantFolders); } // 기억한 위치가 틀리면 지우고 다시 찾기
      throw e;
    }
  }
  function scanOnce(wantFolders) {
    return (async () => {
      await resolveBase();
      const allowed = await myAllowedFcs();
      const ok = (name) => allowed === "*" || allowed.includes(name);
      const folders = [], files = [], done = new Set(); // done: 폴더 ID로 이미 읽은 FC (루트 목록에서 다시 읽지 않음)
      async function scan(id, fc) {
        const items = await stList(id);
        // FC 폴더 바깥(고객 폴더 바로 아래) 파일·분류폴더는 '전체' 권한만
        const rootOk = !!fc || allowed === "*";
        if (rootOk) items.filter((i) => i.file && i.name !== PERM_FILE).forEach((f) => files.push({ ...f, fc }));
        await Promise.all(items.filter((i) => i.folder).map(async (d) => {
          if (d.name.startsWith("_")) {
            if (!wantFolders || !rootOk) return; // 고객 폴더 목록은 필요할 때만
            // 고객 폴더가 들어 있는 분류 폴더만 읽음 (_고객 = 새 규칙, _계약·_이관·_진행·_DB = 옛 구조). _결과·_전환중복 같은 정리용 폴더는 고객으로 보지 않음
            if (CFG.customerFolderGroups && !CFG.customerFolderGroups.includes(d.name.slice(1))) return;
            (await stList(d.id))
              .filter((i) => i.folder)
              .forEach((f) => folders.push({ ...f, group: d.name.slice(1), fc }));
          } else if (!fc && ok(d.name) && !done.has(d.name)) {
            await scan(d.id, d.name); // 허용된 FC 폴더만
          } else if (fc) {
            if (wantFolders) folders.push({ ...d, group: "", fc }); // 새 구조: FC 폴더 바로 아래가 고객 폴더 (이름 (YYMMDD) - 고객ID)
          }
        }));
      }
      try {
        if (allowed !== "*" && allowed.length) {
          // FC 폴더 단위로 공유받은 사용자: 폴더 ID로 바로 열기 (고객 루트 폴더 권한이 없어도 됨). Drive 권한이 없는 FC는 건너뜀
          let opened = 0;
          const rest = allowed.filter((f) => !fcFolderIds[f]);
          await Promise.all(allowed.filter((f) => fcFolderIds[f]).map(async (f) => { // FC 폴더 목록은 동시에
            try { await scan(fcFolderIds[f], f); done.add(f); opened++; }
            catch (e) { if (!/40[34]/.test(String(e.message))) throw e; console.warn("FC 폴더 접근 없음, 건너뜀:", f); }
          }));
          if (rest.length) { // 폴더 ID를 모르는 FC는 예전 방식(고객 루트 목록)으로
            try { await scan(base.id, ""); opened++; } catch (e) { if (!/40[34]/.test(String(e.message))) throw e; }
          }
          if (!opened) throw new Error("NO_SHARE");
        } else await scan(base.id, "");
      } catch (e) {
        if (!base.owner && /40[34]/.test(String(e.message))) throw new Error("NO_SHARE");
        throw e;
      }
      const fcs = [];
      folders.forEach((f) => f.fc && !fcs.includes(f.fc) && fcs.push(f.fc));
      files.forEach((f) => f.fc && !fcs.includes(f.fc) && fcs.push(f.fc));
      return { folders, files, fcs: fcs.sort((a, b) => a.localeCompare(b, "ko")), all: allowed === "*" };
    })();
  }
  const loadCustomerFolders = async () => (await scanBase()).folders;

  // ---------- 고객 리스트 엑셀 (계약고객리스트 / 이관고객리스트) ----------
  const bookCache = {};
  // 저장한 통합문서를 읽기용 캐시에 넣을 때: JSON 파일이면 과도기 읽기 변환을 거친다
  const cacheBook = (fileId, wb) => { bookCache[fileId] = Promise.resolve(jsonInfo.get(fileId) ? jsonToBook(compatDoc(bookToDoc(wb, {}))) : wb); };
  async function readBook(id) {
    if (!bookCache[id]) {
      bookCache[id] = jsonInfo.has(id)
        ? readJsonDoc(id).then((doc) => { const inf = jsonInfo.get(id); if (inf && !inf.meta) inf.meta = { fc: doc.fc, source: doc.source }; return jsonToBook(compatDoc(doc)); }) // JSON을 앱 구조(시트)로 바로 읽기 (v383: 새 칸·옛 칸 둘 다 읽는 과도기 변환)
        : stGet(id).then((buf) => XLSX.read(buf, { type: "array" }));
    }
    return bookCache[id];
  }

  // <schema> 새 데이터 구조(고객·계약·이력·관계 4시트) 이중 읽기 — '고객' 시트에 '고객ID' 열이 있는 파일만 새 구조로 읽고, 아니면 예전 시트(계약고객·이관고객…)를 그대로 읽음
  const CUST_SHEET = "고객", ID_COL = "고객ID";
  // 현재단계 → 예전 type(화면 태그·우선순위·검색 범위 버튼)
  const STAGE_TYPE = { "계약": "계약", "이관관리": "이관", "이관": "이관", "진행": "진행", "가망": "가망", "DB": "DB", "피보험자": "피보험자" };
  const sheetRows = (wb, name) => {
    const sn = wb.SheetNames.find((n) => n.trim() === name);
    return sn ? XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: "", raw: false, blankrows: true }) : null;
  };
  const headIndex = (rows, col) => (rows || []).findIndex((r) => r.some((c) => String(c).trim() === col));
  const isNewStruct = (wb) => { const rows = sheetRows(wb, CUST_SHEET); return !!rows && headIndex(rows, ID_COL) >= 0; };
  const stageType = (row) => { const st = String(row["현재단계"] || "").trim(); if (st === "진행" && String(row["진행제외"] || "").trim()) return "진행제외"; return STAGE_TYPE[st] || "가망"; }; // 진행제외 = 진행 목록에서 뺀 고객(날짜 기록)
  // 고객ID → 이름 (새 구조의 계약 시트가 이름 대신 ID를 가지고 있어서 이름으로 풀어 줌)
  function idNameMap(wb) {
    const rows = sheetRows(wb, CUST_SHEET) || [];
    const h = headIndex(rows, ID_COL), map = new Map();
    if (h < 0) return map;
    const head = rows[h].map((c) => String(c).trim()), ic = head.indexOf(ID_COL), nc = head.indexOf("이름");
    rows.slice(h + 1).forEach((r) => { const id = String(r[ic] || "").trim(), nm = String(r[nc] || "").trim(); if (id && nm) map.set(id, nm); });
    return map;
  }
  // 새 구조의 계약 줄을 예전 계약 탭이 쓰는 모양으로: 계약자·주피보험자 이름을 채움
  function contractRowCompat(row, ids) {
    const o = { ...row };
    if (!o["계약자"] && o["계약자ID"]) o["계약자"] = ids.get(String(o["계약자ID"]).trim()) || "";
    if (!o["주피보험자"] && o["주피보험자ID"]) o["주피보험자"] = ids.get(String(o["주피보험자ID"]).trim()) || "";
    return o;
  }
  // 모집유형 → 계약 구분: '타인모집'(이관) = 이관, '타인모집(비이관)' 등 그 밖의 타인모집… = 타인모집(이관 아님), 그 외 = 계약(직접모집)
  const contractLabel = (mt) => { const v = String(mt || "").trim(); return v === "타인모집" ? "이관" : /^(타인모집|보류)/.test(v) ? "타인모집" : "계약"; }; // v386: '보류(SOL 확인 필요)'는 직접모집으로 세지 않음
  const isNewContractSheet = (head) => head.includes("계약자ID") && !head.includes("계약자");
  // </schema>
  // 앱 버전 표시 (화면 맨 아래 작은 글씨) — 새 버전이 모든 기기에 반영됐는지 확인용
  const APP_VER = (((document.querySelector('script[src*="app.js"]') || {}).src || "").match(/[?&]v=(\d+)/) || [])[1] || "";
  if (APP_VER) { const el = document.createElement("div"); el.className = "app-ver"; el.textContent = "v" + APP_VER; document.body.appendChild(el); }

  const pick = (row, names) => { for (const n of names) if (row[n] != null && String(row[n]).trim()) return String(row[n]).trim(); return ""; };
  const NAME_COLS = ["이름", "계약자", "고객명"];
  const BIRTH_COLS = ["생년월일", "생일"];
  const PHONE_COLS = ["연락처", "휴대폰", "핸드폰", "전화번호"];
  // v576(2026-11-25 사용자 지시): 주민번호·전화번호는 저장할 때 숫자만, 보일 때·입력할 때는 '-'를 알아서 넣음
  const onlyDigits = (v) => String(v == null ? "" : v).replace(/\D/g, "");
  function fmtPhone(v) { // 숫자만 → 010-1234-5678 / 02-123-4567 / 031-123-4567 (입력 중 일부만 있어도 그 자리까지)
    const d = onlyDigits(v); if (!d) return String(v == null ? "" : v).trim();
    if (d.length > 11) return d;
    if (d.startsWith("02")) { const r = d.slice(2); return d.length <= 2 ? d : r.length <= 3 ? "02-" + r : r.length <= 7 ? "02-" + r.slice(0, 3) + "-" + r.slice(3) : "02-" + r.slice(0, 4) + "-" + r.slice(4); }
    if (d.length <= 3) return d;
    if (d.length <= 6) return d.slice(0, 3) + "-" + d.slice(3);
    if (d.length <= 10) return d.slice(0, 3) + "-" + d.slice(3, 6) + "-" + d.slice(6);
    return d.slice(0, 3) + "-" + d.slice(3, 7) + "-" + d.slice(7);
  }
  const PHONE_ONLY = /^[\d\s().+-]+$/;
  const fmtPhoneText = (v) => { const t = String(v == null ? "" : v).trim(); if (!t) return t; return PHONE_ONLY.test(t) ? fmtPhone(t) : t.replace(/\d[\d\s.-]{7,}\d/g, (m) => { const d = onlyDigits(m); return d.length >= 9 && d.length <= 11 && /^0/.test(d) ? fmtPhone(d) : m; }); };
  const phoneStore = (v) => { const t = String(v == null ? "" : v).trim(); return t && PHONE_ONLY.test(t) ? onlyDigits(t) : t; }; // 번호 하나면 숫자만, 글이 섞이면 그대로
  function fmtRrn(v, typing) { // 주민번호 숫자 → 800823-1234567 / 800823-1 / 19800823(생년월일) / -1(성별만)
    const t = String(v == null ? "" : v).trim(), d = onlyDigits(t); if (!d) return t;
    if (typing) return d.length <= 6 ? d : d.slice(0, 6) + "-" + d.slice(6, 13);
    if (d.length === 13 || d.length === 7) return d.slice(0, 6) + "-" + d.slice(6);
    if (d.length === 1) return "-" + d;
    return d;
  }
  document.addEventListener("input", (ev) => { // 입력하는 동안 '-' 자동(커서는 같은 숫자 뒤에 둠)
    const el = ev.target; if (!(el instanceof HTMLInputElement)) return;
    const k = el.dataset.edit || el.dataset.cdf || el.dataset.tmf || "";
    const f = k === "주민번호" ? (x) => fmtRrn(x, true) : PHONE_COLS.includes(k) || (el.type === "tel" && (el.dataset.edit || el.dataset.cdf || el.dataset.tmf)) ? (x) => (PHONE_ONLY.test(x) ? fmtPhone(x) : x) : null;
    if (!f || !el.value) return;
    const pos = el.selectionStart == null ? el.value.length : el.selectionStart, nBefore = onlyDigits(el.value.slice(0, pos)).length, nv = f(el.value);
    if (nv === el.value) return;
    el.value = nv; let i = 0, c = 0; while (i < nv.length && c < nBefore) { if (/\d/.test(nv[i])) c++; i++; }
    try { el.setSelectionRange(i, i); } catch (e) {}
  }, true);

  let listCache = null; // [{ type, name, birth, phone, fc, sheet, fields }]
  // 설정의 파일 조건(fileRegex 또는 fileKeyword)에 맞는 엑셀 중 FC 폴더마다 가장 최근 것 하나만
  //   (예: "20261108 - 최수진 - 계약.xlsx"와 나중의 "20261115 - 최수진 - 계약.xlsx"가 있으면 20261115만)
  function pickSourceFiles(files, spec) {
    const re = spec.fileRegex ? new RegExp(spec.fileRegex, "i") : null;
    // 고객 리스트는 FC 폴더 안의 파일만 (고객 폴더 바로 아래로 옮겨 둔 예전 파일은 읽지 않음)
    const hit = files.filter((f) => f.fc && /\.xlsx?$/i.test(f.name) && (re ? re.test(f.name) : spec.fileKeyword && f.name.includes(spec.fileKeyword)));
    const byFc = new Map();
    // 통합 '리스트' 파일이 있으면 그것을 우선, 그다음 최신 날짜 (예전 '계약'/'기타' 파일은 리스트가 없는 FC만)
    const rank = (f) => (/ - 리스트\.xlsx?$/i.test(f.name) ? "1" : "0") + f.name;
    hit.forEach((f) => { const k = f.fc || ""; if (!byFc.has(k) || rank(f) > rank(byFc.get(k))) byFc.set(k, f); });
    return [...byFc.values()];
  }
  // 읽을 시트: 설정에 sheet가 있으면 그 시트만, 없으면 전부
  const sheetsOf = (wb, spec) => (spec.sheet ? wb.SheetNames.filter((n) => n.trim() === spec.sheet) : wb.SheetNames);

  const knownFcs = new Set(); // 리스트 파일을 읽은 FC 이름(고객 줄이 없어도)
  async function loadLists() {
    if (listCache) return listCache;
    const { files } = await scanFiles();
    const folders = CFG.lists.some((l) => l.folderGroup) ? (await scanBase()).folders : []; // 폴더 기반 범위가 있을 때만 고객 폴더 목록을 읽음
    const sources = CFG.lists.filter((l) => l.fileRegex || l.fileKeyword).flatMap((l) =>
      pickSourceFiles(files, l).map((f) => ({ ...l, file: f })));
    const out = [], newDone = new Set();
    // 폴더 기반 범위 (예: 진행 = "_진행" 폴더 안 고객 폴더)
    CFG.lists.filter((l) => l.folderGroup).forEach((l) =>
      folders.filter((f) => f.group === l.folderGroup).forEach((f) => out.push({
        type: l.label, name: custPart(f.name), fc: f.fc, sheet: "", birth: "", phone: "",
        fields: [["고객 폴더", f.name], ["분류", l.folderGroup]].concat(f.fc ? [["담당 FC", f.fc]] : [])
      })));
    await Promise.all(sources.map(async (s) => {
      const wb = await readBook(s.file.id);
      // FC 이름: FC 폴더 이름 → 없으면 파일명 "날짜 - FC - ..." 가운데 부분
      const fc = s.file.fc || (s.file.name.split(" - ")[1] || "").trim();
      if (fc) knownFcs.add(fc); // v511: 고객이 0명인 FC(빈 data.json)도 FC 선택에 보이게
      // 새 구조 파일: '고객' 시트 하나만 (파일당 한 번), 예전 시트(계약고객 등)는 건너뜀
      const newMode = isNewStruct(wb);
      if (newMode) { if (newDone.has(s.file.id)) return; newDone.add(s.file.id); }
      (newMode ? [wb.SheetNames.find((n) => n.trim() === CUST_SHEET)] : sheetsOf(wb, s)).forEach((sheet) => {
        const ws = wb.Sheets[sheet];
        if (!ws || !ws["!ref"]) return;
        const rg = XLSX.utils.decode_range(ws["!ref"]);
        // blankrows: true → rows[i]는 엑셀의 (rg.s.r + i)번째 줄 (나중에 같은 줄에 저장하기 위해)
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: false, blankrows: true });
        // 제목 줄이 있을 수 있으니 '이름/계약자' 칸이 있는 줄을 머리글로 사용
        const h = rows.findIndex((r) => r.some((c) => NAME_COLS.includes(String(c).trim())));
        if (h < 0) return;
        const head = rows[h].map((c) => String(c).trim());
        rows.slice(h + 1).forEach((r, i) => {
          const row = {};
          head.forEach((k, j) => { if (k) row[k] = r[j]; });
          const name = nameShow(pick(row, NAME_COLS));
          if (!name) return;
          out.push({
            type: newMode ? stageType(row) : s.label, name, fc, sheet,
            birth: pick(row, BIRTH_COLS),
            phone: pick(row, PHONE_COLS),
            grade: String(row[GRADE_COL] ?? "").trim(),
            src: { fileId: s.file.id, sheet, row: rg.s.r + h + 1 + i },
            fields: head.map((k, j) => [k, String(r[j] ?? "").trim()]).filter(([k, v]) => k && v)
          });
        });
      });
    }));
    listCache = out; introCache = null;
    renderDataStamp();
    return out;
  }

  // ---------- 고객리스트 엑셀에 값 쓰기 (성향, 연락처 수정 + 우선) ----------
  const GRADE_COL = "성향", PRIORITY_COL = "우선";
  const GRADES = [
    { id: "Good", label: "Good", color: "#ffffff" },
    { id: "Bad", label: "Bad", color: "#f39c12" },
    { id: "Shit", label: "Shit", color: "#e53935" }
  ];
  const gradeOf = (c) => { const g = (c.entries || []).map((e) => e.grade).find(Boolean) || ""; return GRADES.find((x) => x.id.toLowerCase() === g.toLowerCase()) ? g : "Good"; }; // 성향이 비어 있으면 Good으로 봄(이름 굵게)
  const gradeClass = (g) => (g ? "grade-" + g.toLowerCase() : "");

  // entries의 원래 줄(파일·시트·행)에 values({열이름: 값})를 기록. 열이 없으면 머리글 끝에 새로 만듦
  async function writeEntryCells(entries, values, extra) { // extra(wb, fileId): 같은 파일을 올리기 전에 더 고칠 것(예: 이력 한 줄) — 한 번 받고 한 번 저장
    entries = ownEntries(entries); // 담당 FC 행에만 저장
    if (entries.some((e) => isRO(e.fc))) throw new Error("보기 전용 권한이라 수정할 수 없습니다.");
    const byFile = new Map();
    entries.filter((e) => e.src).forEach((e) => { if (!byFile.has(e.src.fileId)) byFile.set(e.src.fileId, []); byFile.get(e.src.fileId).push(e); });
    for (const [fileId, list] of byFile) {
      const wb = await stGetBook(fileId); // 최신 파일 (cellNF: 숫자·날짜 서식 유지)
      for (const e of list) {
        const ws = wb.Sheets[e.src.sheet];
        if (!ws) continue;
        const rg = XLSX.utils.decode_range(ws["!ref"]);
        const enc = (r, c) => XLSX.utils.encode_cell({ r, c });
        const txt = (r, c) => { const cell = ws[enc(r, c)]; return cell ? String(cell.w ?? cell.v).trim() : ""; };
        // 머리글 줄 찾기
        let hr = rg.s.r;
        for (let r = rg.s.r; r <= Math.min(rg.e.r, rg.s.r + 10); r++) {
          let found = false;
          for (let c = rg.s.c; c <= rg.e.c; c++) if (NAME_COLS.includes(txt(r, c))) found = true;
          if (found) { hr = r; break; }
        }
        const colOf = (name) => { for (let c = rg.s.c; c <= rg.e.c; c++) if (txt(hr, c) === name) return c; return -1; };
        const nameCol = NAME_COLS.map(colOf).find((c) => c >= 0);
        const birthCol = BIRTH_COLS.map(colOf).find((c) => c >= 0);
        // 저장해 둔 줄이 맞는지 확인, 아니면 이름(+생년월일)으로 다시 찾기
        let row = e.src.row;
        const same = (r) => norm(txt(r, nameCol)) === norm(e.name) && (birthCol == null || birthCol < 0 || !e.birth || txt(r, birthCol).replace(/\D/g, "") === String(e.birth).replace(/\D/g, ""));
        if (!same(row)) {
          row = -1;
          for (let r = hr + 1; r <= rg.e.r; r++) if (same(r)) { row = r; break; }
        }
        if (row < 0) continue;
        const pairs = [];
        for (const [k0, v0] of Object.entries(values)) pairs.push(...writeTargets(wb, (n) => colOf(n) >= 0, k0, v0, e));
        for (const [k, v] of pairs) {
          let c = colOf(k);
          if (c < 0) { // 새 열: 맨 끝이 '입력일시'면 그 앞에 만들고 입력일시는 한 칸 뒤로
            const sc = colOf(STAMP_COL);
            if (sc >= 0 && sc === rg.e.c) {
              for (let r = rg.s.r; r <= rg.e.r; r++) { const cell = ws[enc(r, sc)]; delete ws[enc(r, sc)]; if (cell) ws[enc(r, sc + 1)] = cell; }
              if (ws["!cols"] && ws["!cols"][sc]) { ws["!cols"][sc + 1] = ws["!cols"][sc]; ws["!cols"][sc] = { wch: 12 }; }
              c = sc;
            } else c = rg.e.c + 1;
            ws[enc(hr, c)] = { t: "s", v: k }; rg.e.c += 1;
          }
          ws[enc(row, c)] = typeof v === "number" ? { t: "n", v } : { t: "s", v: String(v) };
        }
        ws["!ref"] = XLSX.utils.encode_range(rg);
        e.src.row = row;
        // 화면에 쓰는 값도 갱신
        for (const [k, v] of Object.entries(values)) {
          const f = e.fields.find(([fk]) => fk === k);
          if (f) f[1] = String(v); else e.fields.push([k, String(v)]);
          if (k === "주민번호") { const pj = parseJumin(v), setf = (kk, vv) => { const ff = e.fields.find(([fk]) => fk === kk); if (ff) ff[1] = vv; else if (vv) e.fields.push([kk, vv]); }; if (pj.birth) { e.birth = pj.birth; setf("생년월일", pj.birth); } if (pj.sex) setf("성별", pj.sex); if (pj.back) setf("후번호", pj.back); }
          if (/^소개/.test(k)) introCache = null;
          if (k === GRADE_COL) e.grade = String(v);
          if (PHONE_COLS.includes(k)) e.phone = String(v);
        }
      }
      if (extra) extra(wb, fileId);
      await stPutBook(fileId, wb);
      cacheBook(fileId, wb); // 방금 저장한 내용이 최신 → 다시 받지 않음(목록을 다시 그릴 때 1초쯤 빨라짐)
    }
  }

  // ---------- 검색 ----------
  let mode = "all";
  const norm = (s) => String(s).replace(/\s+/g, "").toLowerCase();
  // v579(2026-11-25 사용자 지시): 영어 이름은 낱말마다 첫 글자만 대문자(OKAMOTO KEIKO → Okamoto Keiko). 한글이 섞인 이름은 그대로
  const nameShow = (v) => { const t = String(v == null ? "" : v); return /[A-Za-z]/.test(t) && !/[가-힣]/.test(t) ? t.replace(/[A-Za-z]+/g, (w) => w[0].toUpperCase() + w.slice(1).toLowerCase()) : t; };
  const custPart = (n) => n.split(" - ")[0]; // 폴더명 "고객명 - 담당자명"
  // 동명이인 표시를 뺀 이름: "이영숙(640927)" → "이영숙" (일정 제목·계약 매칭은 이 이름으로)
  const baseName = (n) => String(n || "").replace(/\(\d{6}\)\s*$/, "").trim();
  // ---- 새 구조: 고객ID·담당FC ----
  const progFlag = (r) => (r.type === "계약" || r.type === "이관") && /^y$/i.test(fieldOf(r, "진행")); // v375(2026-11-16): 계약·이관관리 고객도 고객 시트 '진행'=Y면 진행 목록에 함께 나옴(계약·이관 목록에도 그대로)
  const fieldOf = (e, k) => { const f = (e.fields || []).find(([x]) => x.trim() === k); return f ? String(f[1]).trim() : ""; };
  const idOf = (e) => fieldOf(e, "고객ID");
  // 해지 고객 표시(화면에서 계산, 데이터 수정 없음): 계약자 기준 계약이 1건 이상 있고 유효 계약(유지·완납·납입면제·연금지급·신계약승낙)이 0건이면 '해지', 실효 계약만 남았으면 '실효'. 계약이 없으면(가망·진행·피보험자) 표시 없음
  const VALID_ST = ["유지", "완납", "납입면제", "연금지급", "신계약승낙"];
  function churnOf(ids, cts) {
    const mine = cts.filter((ct) => ids.has(String(ct.row["계약자ID"] || "").trim()));
    if (!mine.length) return null;
    const st = mine.map((ct) => String(ct.row["계약상태"] || ct.row["상태"] || "").trim());
    if (st.some((s) => VALID_ST.includes(s))) return null;
    return st.includes("실효") ? "실효" : "해지";
  }
  const churnTagHtml = (k) => k ? `<span class="tag t-${k}">${k === "해지" ? "해지" : "실효"}</span>` : "";
  // 이 고객의 담당 FC: '담당FC' 칸이 자기 파일의 FC와 같은 행의 FC (없으면 첫 행의 FC)
  function primaryFc(entries) {
    const p = entries.find((e) => { const f = fieldOf(e, "담당FC") || fieldOf(e, "담당 FC"); return f && e.fc === f; });
    return ((p || entries[0] || {}).fc) || "";
  }
  // 저장은 담당 FC의 행에만 (다른 FC 파일에 같은 사람이 있어도 거기엔 쓰지 않음)
  function ownEntries(entries) {
    const fc = primaryFc(entries), own = entries.filter((e) => e.fc === fc);
    return own.length ? own : entries;
  }
  // 같은 사람 묶기: 고객ID가 같으면 같은 사람(한 FC 파일에만 보이는 행끼리만 묶임), ID가 없는 옛 행은 이름+생년월일(YYMMDD)
  function groupCustomers(rs) {
    const ymd = (b) => { const x = String(b || "").replace(/\D/g, ""); return /^(19|20)\d{6}/.test(x) ? x.slice(2, 8) : x.slice(0, 6); };
    const groups = [], byId = new Map(), byNb = new Map();
    rs.forEach((r) => {
      const id = idOf(r), nb = norm(baseName(r.name)) + "|" + ymd(r.birth);
      let g = id ? byId.get(id) : null;
      if (!g && !id) g = byNb.get(nb); // ID 없는 행: 같은 이름+생년월일 그룹(ID 그룹 포함)에 합류
      if (!g) { g = { name: r.name, birth: r.birth, entries: [] }; groups.push(g); if (id) byId.set(id, g); }
      if (id && !byId.has(id)) byId.set(id, g);
      if (!byNb.has(nb) || ymd(r.birth)) { if (!byNb.has(nb)) byNb.set(nb, g); }
      g.entries.push(r);
      if (!g.birth && r.birth) g.birth = r.birth;
    });
    return groups;
  }
  // 시트 우선순위: config.lists 순서 (계약 > 이관 > 따행 > DB > 추가 > 진행)
  const LIST_PRIORITY = [...CFG.lists.map((l) => l.label), "진행제외"];
  const prio = (type) => { const i = LIST_PRIORITY.indexOf(type); return i < 0 ? 99 : i; };
  // 미팅장소: 고객리스트 시트의 '미팅장소' 칸(우선순위 순) → 없으면 고객추가정보 엑셀
  function meetingOf(c, extra) {
    const es = [...(c.entries || [])].sort((a, b) => prio(a.type) - prio(b.type));
    for (const e of es) { const f = e.fields.find(([k]) => k.trim() === "미팅장소"); if (f && String(f[1]).trim()) return String(f[1]).trim(); }
    const x = extra || c.extra || {};
    return String((x.values || {})["미팅장소"] || "").trim();
  }

  // 주소 → 지역 약칭: "광주광역시 서구…" → 광주, "전라남도 영광군…" → 전남 영광
  const METRO = ["서울", "부산", "대구", "인천", "광주", "대전", "울산", "세종"];
  const PROVINCE = [[/^경기/, "경기"], [/^강원/, "강원"], [/^(충청북|충북)/, "충북"], [/^(충청남|충남)/, "충남"],
    [/^(전라북|전북)/, "전북"], [/^(전라남|전남)/, "전남"], [/^(경상북|경북)/, "경북"], [/^(경상남|경남)/, "경남"], [/^제주/, "제주"]];
  function regionFromAddress(addr) {
    const t = String(addr || "").trim().split(/\s+/);
    if (!t[0]) return "";
    const cityOf = (s) => (s || "").match(/^(\S+?)(특례시|시|군)$/);
    // 전남광주통합특별시: 구(서구·북구…)는 광주, 시·군(영광군…)은 전남 ○○
    if (/전남광주통합/.test(t[0])) { const c = cityOf(t[1]); return c ? "전남 " + c[1] : "광주"; }
    if (/^광주/.test(t[0])) return "광주";
    const metro = METRO.find((m) => t[0].startsWith(m));
    if (metro) return metro;
    const prov = PROVINCE.find(([re]) => re.test(t[0]));
    if (!prov) return "";
    // 띄어쓰기 없는 주소: "전남무안군삼향면" → 전남 무안
    const glued = t[0].replace(/^(전라남도|전라북도|경상남도|경상북도|충청남도|충청북도|경기도|강원특별자치도|강원도|전북특별자치도|제주특별자치도|전남|전북|경남|경북|충남|충북|경기|강원|제주)/, "");
    const g = glued && glued.match(/^(\S+?)(시|군)/);
    const c = g ? [null, g[1]] : cityOf(t[1]);
    const city = c ? c[1] : "";
    return city && city !== prov[1] ? prov[1] + " " + city : prov[1];
  }
  function regionOf(c) {
    for (const e of c.entries) {
      const a = e.fields.find(([k]) => k.trim() === "주소");
      if (a && a[1]) return regionFromAddress(a[1]);
    }
    return "";
  }

  // 미팅장소 (고객추가정보 '기본' 시트) — 검색용, 저장하면 다시 읽음
  let meetCache = null;
  function loadMeetingPlaces() {
    if (meetCache) return meetCache;
    meetCache = (async () => {
      const x = await readExtra();
      if (!x) return [];
      const ni = x.head.indexOf(EXTRA.nameCol), fi = x.head.indexOf(EXTRA.fcCol), mi = x.head.indexOf("미팅장소");
      if (ni < 0 || mi < 0) return [];
      return x.rows.slice(1).map((r) => ({ name: String(r[ni] ?? "").trim(), fc: fi < 0 ? "" : String(r[fi] ?? "").trim(), mp: String(r[mi] ?? "").trim() }))
        .filter((m) => m.name && m.mp);
    })();
    meetCache.catch(() => { meetCache = null; });
    return meetCache;
  }
  const digits = (v) => String(v || "").replace(/\D/g, "");
  // 전체검색(위쪽 검색) 확장: 고객 메모(고객추가정보 '메모')·상담기록 글에서도 찾음 (2분 동안 기억)
  let extraTextCache = null, extraTextAt = 0;
  function loadExtraText() {
    if (extraTextCache && Date.now() - extraTextAt < 120000) return extraTextCache;
    extraTextAt = Date.now();
    extraTextCache = (async () => {
      const x = await readExtra();
      if (!x) return { memos: [], consults: [] };
      const ni = x.head.indexOf(EXTRA.nameCol), fi = x.head.indexOf(EXTRA.fcCol), mi = x.head.indexOf("메모");
      const memos = ni < 0 || mi < 0 ? [] : x.rows.slice(1).map((r) => ({ name: String(r[ni] ?? "").trim(), fc: fi < 0 ? "" : String(r[fi] ?? "").trim(), text: String(r[mi] ?? "").trim() })).filter((m) => m.name && m.text);
      const consults = [];
      const ws = x.wb.Sheets[EXTRA.consult.sheet];
      if (ws && ws["!ref"]) {
        const { rg, ni: cn, fi: cf, ti } = consultSheetInfo(ws);
        if (cn >= 0 && ti >= 0) for (let r = rg.s.r + 1; r <= rg.e.r; r++) {
          const a = ws[XLSX.utils.encode_cell({ r, c: cn })], t = ws[XLSX.utils.encode_cell({ r, c: ti })], f = cf < 0 ? null : ws[XLSX.utils.encode_cell({ r, c: cf })];
          if (a && t && String(t.v).trim()) consults.push({ name: String(a.v).trim(), fc: f ? String(f.v).trim() : "", text: String(t.v).trim() });
        }
      }
      return { memos, consults };
    })();
    extraTextCache.catch(() => { extraTextCache = null; });
    return extraTextCache;
  }
  const snippetOf = (text, nq) => { const t = String(text).replace(/\s+/g, " "), i = norm(t).indexOf(nq); if (i < 0) return t.slice(0, 24); const a = Math.max(0, i - 8); return (a ? "…" : "") + t.slice(a, a + 26) + (t.length > a + 26 ? "…" : ""); };
  // 전체검색 확장: 오늘 기준 1년 전 ~ 1개월 뒤 일정(제목·설명·장소). 구글 캘린더 검색(q) 사용, 연결이 없거나 실패하면 조용히 건너뜀
  async function searchCalendarEvents(q) {
    try {
      const cals = await scheduleCalendars();
      const from = dayStart(new Date()); from.setFullYear(from.getFullYear() - 1);
      const to = dayStart(new Date()); to.setMonth(to.getMonth() + 1);
      const qs = `?singleEvents=true&orderBy=startTime&maxResults=100&q=${encodeURIComponent(q)}&timeMin=${encodeURIComponent(from.toISOString())}&timeMax=${encodeURIComponent(to.toISOString())}`;
      const lists = await Promise.all(cals.map(async (c) => {
        try {
          return ((await gcal("/calendars/" + encodeURIComponent(c.id) + "/events" + qs)).items || []).filter((e) => e.status !== "cancelled").map((e) => ({
            cal: calAlias(c.id, c.name || ""), title: String(e.summary || "(제목 없음)").trim(), location: String(e.location || "").trim(),
            allDay: !!e.start.date, start: e.start.dateTime ? new Date(e.start.dateTime) : ymdToDate(e.start.date)
          }));
        } catch (e) { return []; }
      }));
      return lists.flat().sort((a, b) => b.start - a.start).slice(0, 40); // 최근 일정이 위로
    } catch (e) { return []; }
  }
  let searchWide = true; // true = 위쪽 전체검색(고객·메모·상담기록·일정), false = 고객 화면 안 고객검색(고객만)
  // 전체검색 결과 팝업(아래에서 올라오는 창): 고객·메모·상담기록·일정 결과를 보여 주고, 누르면 해당 페이지(고객 정보·상담 탭·그날 일정)로 이동
  function closeSearchPopup() { const b = $("searchPopup"); if (b) b.hidden = true; document.documentElement.classList.remove("help-open"); }
  function openSearchPopup(q, custs, evs) {
    let box = $("searchPopup");
    if (!box) { box = document.createElement("div"); box.id = "searchPopup"; box.className = "help-modal"; document.body.appendChild(box); box.addEventListener("click", (e) => { if (e.target === box || e.target.closest("[data-spclose]")) closeSearchPopup(); }); }
    const WDK = ["일", "월", "화", "수", "목", "금", "토"];
    const custHtml = custs.map((c, i) => {
      const types = [...new Set(c.entries.map((e) => e.type))];
      const why = c.why && c.why.size ? [...c.why].map(([k, v]) => esc(k) + (v && k !== "연락처" ? " · " + esc(v) : "")).join(" / ") : "";
      return `<li><button type="button" data-spc="${i}"><span class="avatar sm">${esc(c.name[0])}</span><span class="r-name"><span><span class="cname ${gradeClass(gradeOf(c))}">${trophyOf(c)}${esc(c.name)}</span>${resultInfo(c) ? ` <span class="rinfo">${esc(resultInfo(c))}</span>` : ""}</span>${why ? `<span class="r-why">${why}</span>` : ""}</span>${types.map((t) => `<span class="tag t-${esc(t)}">${esc(typeLabel(t))}</span>`).join("")}<span class="chev">›</span></button></li>`;
    }).join("");
    const evHtml = evs.map((e, i) => `<li><button type="button" data-spe="${i}"><span class="r-name"><span class="cname">${esc(e.title)}</span><span class="r-sub">${e.start.getFullYear()}.${pad2(e.start.getMonth() + 1)}.${pad2(e.start.getDate())}(${WDK[e.start.getDay()]}) ${e.allDay ? "종일" : hm(e.start)} · ${esc(e.cal)}${e.location ? " · " + esc(e.location) : ""}</span></span><span class="chev">›</span></button></li>`).join("");
    box.innerHTML = `<div class="help-sheet" role="dialog" aria-label="전체검색 결과"><div class="help-top"><div class="ai-grab" aria-hidden="true"></div><div class="help-head"><b>전체검색 · ${esc(q)}</b><button type="button" class="secondary" data-spclose>닫기</button></div></div><div class="help-body sp-body">` +
      (custs.length ? `<h4>고객 ${custs.length}명</h4><ul class="list">${custHtml}</ul>` : "") +
      (evs.length ? `<h4>일정 ${evs.length}건 <span class="hint">(1년 전 ~ 1개월 뒤)</span></h4><ul class="list">${evHtml}</ul>` : "") +
      (!custs.length && !evs.length ? `<p class="hint pad">'${esc(q)}'에서 찾은 것이 없습니다.</p>` : "") + `</div></div>`;
    box.querySelectorAll("[data-spc]").forEach((b) => b.addEventListener("click", async () => {
      const c = custs[+b.dataset.spc]; closeSearchPopup();
      await openCustomer(c);
      if (c.why && c.why.has("상담기록") && !c.why.has("메모")) selectTab("consult"); // 상담기록에서 걸렸으면 상담 탭으로
    }));
    box.querySelectorAll("[data-spe]").forEach((b) => b.addEventListener("click", () => { const e = evs[+b.dataset.spe]; closeSearchPopup(); schedMode = "day"; schedDate = dayStart(e.start); openSchedule(); }));
    box.hidden = false; document.documentElement.classList.add("help-open");
  }
  document.addEventListener("keydown", (e) => { if (e.key === "Escape") closeSearchPopup(); });

  // 고객 범위 옵션: 전체 / 본인(계약) / 이관 / DB / 따행 / 가망. 본인부터는 다중 선택(고른 것 중 하나라도 해당하면 포함), 전체는 선택 해제.
  const optSel = new Set();
  // 성향 버튼: 누를 때마다 Good → Bad → Shit → Good. Good은 기본(걸러내지 않음), Bad·Shit은 그 성향의 고객만 보기
  let gradeFilter = "Good";
  function syncGradeOpt() {
    const b = $("gradeOpt"); if (!b) return;
    // v225: 글자 G·B·S 세 개, 지금 고른 성향 글자만 색(G 초록·B 주황·S 빨강), 나머지는 회색
    b.innerHTML = [["Good", "G"], ["Bad", "B"], ["Shit", "S"]].map(([id, ch]) => `<span class="gl gl-${id}${gradeFilter === id ? " on" : ""}">${ch}</span>`).join("");
    b.setAttribute("aria-label", "성향 " + gradeFilter);
  }
  syncGradeOpt();
  const OPT_LABEL = { contract: "직접모집", transfer: "이관고객", progress: "계약진행", db: "DB", ddahaeng: "따행", plan: "가망" }; // v383: 새 말(직접모집·이관고객·계약진행)
  function syncOpts() { document.querySelectorAll("#searchOpts [data-opt]").forEach((b) => b.classList.toggle("on", b.dataset.opt === "all" ? optSel.size === 0 : optSel.has(b.dataset.opt))); }
  const typeOfOpt = (id) => (CFG.lists.find((l) => l.id === id) || {}).label;
  // 고객 리스트의 FC 선택: 전체 / 각 FC. 기본은 내 FC(권한 파일의 내 이름), 고른 값은 이 폰에 기억. FC가 하나뿐이면 선택칸을 숨김
  const FC_ALL = "__all__";
  const listFcKey = () => "shlife.listFc." + me();
  async function renderFcOpts(records) {
    const box = $("fcOpts");
    if (!box) return FC_ALL;
    const fcs = [...new Set([...records.map((r) => r.fc), ...knownFcs].filter(Boolean))].sort((a, b) => { const o = CFG.fcOrder || [], ia = o.indexOf(a), ib = o.indexOf(b); return (ia < 0 ? 999 : ia) - (ib < 0 ? 999 : ib) || a.localeCompare(b, "ko"); });
    if (fcs.length <= 1) { box.hidden = true; box.innerHTML = ""; return FC_ALL; }
    const myFc = await myFcName();
    let saved = ""; try { saved = localStorage.getItem(listFcKey()) || ""; } catch (e) { /* 무시 */ }
    const sel = saved === FC_ALL || fcs.includes(saved) ? saved : fcs.includes(myFc) ? myFc : FC_ALL;
    const opts = [[FC_ALL, "전체"], ...fcs.map((f) => [f, f])];
    // FC가 5명 이하면 버튼 줄, 6명 이상이면 드롭박스(선택창)
    box.innerHTML = fcs.length <= 5
      ? opts.map(([v, l]) => `<button type="button" data-fco="${esc(v)}" class="${v === sel ? "on" : ""}">${esc(l)}</button>`).join("")
      : `<select id="fcoSel" class="fco-sel" aria-label="고객 FC 선택">${opts.map(([v, l]) => `<option value="${esc(v)}"${v === sel ? " selected" : ""}>${esc(l)}</option>`).join("")}</select>`;
    box.hidden = false;
    return sel;
  }
  const curQuery = () => ((($("q") && $("q").value) || ($("q2") && $("q2").value) || "")).trim();
  // 검색 결과에서 같은 이름을 구분하도록 이름 옆에 작게: 생년 · 남/여 · 지역
  function birthYear(b) { const raw = String(b || "").trim(), m4 = raw.match(/^((?:19|20)\d{2})(?:\D|$)/); if (m4) return m4[1]; const x = raw.replace(/\D/g, ""); if (/^(19|20)\d{2}($|\d{4})/.test(x) && (x.length === 4 || x.length === 8)) return x.slice(0, 4); if (x.length >= 6) { const yy = +x.slice(0, 2), cy = new Date().getFullYear() % 100; return (yy > cy ? 1900 + yy : 2000 + yy) + ""; } return ""; }
  const resultInfo = (c) => [birthYear(c.birth), genderOf(c), regionOf(c)].filter(Boolean).join(" · ");
  let progShowRemoved = false; // [진행] 목록: false = 진행 고객, true = 보관(진행에서 이동한) 고객
  let trShowRemoved = false; // [이관] 목록: false = 이관 고객, true = 보관(이관 계약을 모두 보관한 고객)
  let dbShowRemoved = false; // [DB] 목록: false = DB 고객, true = 보관(이동한 고객 + 계약 고객)
  const TYPE_SHOW = { "계약": "직접모집", "이관": "이관고객", "진행": "계약진행" }; // v383: 화면에 보이는 고객 구분 말(데이터 안의 값은 그대로)
  const typeLabel = (t) => (t === "진행제외" ? "계약진행 보관" : TYPE_SHOW[t] || t);
  let goTabPending = ""; // 말만해요 tab=analysis 등: 고객 정보를 열 때 이 탭으로
  let goKeepList = null; // 말만해요에서 넘어온 검색어: 한 명만 나와도 목록을 보여 줌
  async function search(q, force) {
    q = q.trim();
    const keepList = goKeepList !== null && goKeepList === q; goKeepList = null;
    const tabReq = goTabPending; goTabPending = "";
    const status = $("searchStatus");
    const list = $("results");
    list.innerHTML = "";
    { const eh = $("evHits"); if (eh) eh.remove(); const pb = $("progBar"); if (pb) pb.remove(); $("searchStatus").classList.remove("has-pbar"); }
    const wide = searchWide;
    syncOpts();
    $("pageSearch").hidden = !!q || !!force;
    if (!q && !force) { status.textContent = ""; return; }
    $("q2").value = q;
    notifyWhereNow();
    status.textContent = "찾는 중…";
    try {
      const records = await loadLists();
      const fcPick = await renderFcOpts(records); // 위쪽 FC 선택 (전체 또는 FC 이름)
      const want = CFG.lists.find((l) => l.id === mode);
      if (want && !records.some((r) => r.type === want.label)) {
        status.textContent = want.folderGroup
          ? `'_${want.folderGroup}' 폴더에 고객 폴더가 없습니다.`
          : want.sheet ? `'${want.sheet}' 시트에 아직 고객이 없습니다.`
          : `'${want.label}' 엑셀 파일이 없습니다. FC 폴더에 이름에 '${want.fileKeyword}'가 들어간 엑셀을 넣어 주세요.`;
        return;
      }
      // 고객 이름 · 연락처 · 주소 · 미팅장소로 검색
      const nq = norm(q), dq = digits(q);
      const evP = wide && nq.length >= 2 ? searchCalendarEvents(q) : Promise.resolve([]); // 일정 검색은 나머지와 동시에
      const meets = await loadMeetingPlaces().catch(() => []);
      const ex = wide && nq.length >= 2 ? await loadExtraText().catch(() => ({ memos: [], consults: [] })) : { memos: [], consults: [] };
      const exMap = (arr) => { const m = new Map(); arr.forEach((a) => { if (!norm(a.text).includes(nq)) return; const k = norm(a.name); if (!m.has(k)) m.set(k, []); m.get(k).push(a); }); return m; };
      const memoHit = exMap(ex.memos), consHit = exMap(ex.consults);
      const why = new Map(); // 기록 → { 이유: 글 조각 } (이름이 아닌 곳에서 걸린 것을 목록에 보여 주기 위함)
      const addWhy = (r, k, snip) => { if (!why.has(r)) why.set(r, new Map()); if (!why.get(r).has(k)) why.get(r).set(k, snip || ""); };
      const meetOf = (r) => meets.filter((m) => norm(m.name) === norm(r.name) && sameFc(m.fc, r.fc)).map((m) => m.mp)
        .concat(r.fields.filter(([k]) => k.trim() === "미팅장소").map(([, v]) => v)).join(" ");
      // 증권번호 검색: 숫자 5자리 이상이면 그 증권번호가 들어간 계약의 계약자·피보험자도 찾음
      let polIds = null, polNames = null;
      if (dq.length >= 5 && dq === q.replace(/[\s-]/g, "")) {
        const cts = (await loadContracts().catch(() => [])).filter((c) => digits(c.row["증권번호"]).includes(dq));
        if (cts.length) {
          polIds = new Set(cts.flatMap((c) => [c.row["계약자ID"], c.row["주피보험자ID"], c.row["종피보험자ID"]]).map((x) => String(x || "").trim()).filter(Boolean));
          polNames = cts.flatMap((c) => (c.names || []).map((n) => n + "|" + (c.srcFc || c.fc)));
        }
      }
      const optTypes = [...optSel]; // 고른 범위 (따행은 따행 시트 고객 + 고객 시트의 따행대상='예')
      // 이관 필터: 고객 단위 보관 상태(그 고객의 이관 계약이 모두 보관이면 보관)
      const trMap = new Map();
      if (optTypes.includes("transfer")) {
        const cts = await loadContracts().catch(() => []);
        const byHolder = new Map();
        trCtsOf({ has: () => true }, cts).forEach((c) => { const id = String(c.row["계약자ID"] || "").trim(); (byHolder.get(id) || byHolder.set(id, []).get(id)).push(c); });
        byHolder.forEach((list, id) => trMap.set(id, { list, st: trStateOf(list) }));
      }
      const trOff = (r) => { const t = trMap.get(idOf(r)); return !!t && t.st === "off"; };
      const hits = records.filter((r) => {
        if (fcPick !== FC_ALL && r.fc !== fcPick) return false;
        if (want && r.type !== want.label) return false;
        if (optTypes.length && !optTypes.some((id) => id === "ddahaeng" ? (r.type === typeOfOpt(id) || fieldOf(r, "따행대상") === "예") : id === "progress" ? (progShowRemoved ? r.type === "진행제외" : r.type === typeOfOpt(id) || progFlag(r)) : id === "db" ? (isDbRec(r) && dbOffRec(r) === dbShowRemoved) : id === "transfer" ? (r.type === typeOfOpt(id) && trOff(r) === trShowRemoved) : r.type === typeOfOpt(id))) return false; // DB = 현재단계가 DB인 고객 + 유입경로가 DB인 고객(이미 계약·진행으로 넘어간 고객 포함)
        if (!q) return true;
        let ok = false;
        if (polIds && (polIds.has(idOf(r)) || (!idOf(r) && polNames.includes(norm(baseName(r.name)) + "|" + r.fc)))) ok = true;
        if (norm(r.name).includes(nq)) ok = true;
        if (dq.length >= 3 && dq === q.replace(/[\s-]/g, "") && digits(r.phone).includes(dq)) { ok = true; if (wide) addWhy(r, "연락처", r.phone); }
        if (nq.length < 2) return ok;
        if (norm(entryAddr(r)).includes(nq)) { ok = true; if (wide) addWhy(r, "주소", snippetOf(entryAddr(r), nq)); }
        if (norm(meetOf(r)).includes(nq)) { ok = true; if (wide) addWhy(r, "미팅장소", snippetOf(meetOf(r), nq)); }
        if (!wide) return ok;
        // 전체검색: 메모(고객추가정보·고객 시트의 '메모' 열)와 상담기록 — 같은 이름 + 같은 FC
        const keys = [norm(r.name), norm(baseName(r.name))];
        const pick = (map) => keys.flatMap((k) => map.get(k) || []).filter((a) => sameFc(a.fc, r.fc))[0];
        const mh = pick(memoHit); if (mh) { ok = true; addWhy(r, "메모", snippetOf(mh.text, nq)); }
        const own = fieldOf(r, "메모"); if (own && norm(own).includes(nq)) { ok = true; addWhy(r, "메모", snippetOf(own, nq)); }
        const ch = pick(consHit); if (ch) { ok = true; addWhy(r, "상담기록", snippetOf(ch.text, nq)); }
        return ok;
      });
      // 같은 사람 묶기: 고객ID가 같으면 한 고객, ID가 없는 옛 행은 같은 이름 + 같은 생년월일
      const groups = groupCustomers(hits);
      // 고객 안의 시트는 우선순위 순 (계약 > 이관 > 따행 > DB > 가망)
      groups.forEach((c) => { c.entries.sort((a, b) => prio(a.type) - prio(b.type)); c.name = c.entries[0].name; });
      groups.forEach((c) => { c.why = new Map(); c.entries.forEach((e) => (why.get(e) || new Map()).forEach((v, k) => { if (!c.why.has(k)) c.why.set(k, v); })); });
      const custs = groups.filter((c) => gradeFilter === "Good" || gradeOf(c).toLowerCase() === gradeFilter.toLowerCase()).sort((a, b) => baseName(a.name).localeCompare(baseName(b.name), "ko") || a.name.localeCompare(b.name, "ko"));
      // 피보험자 고객: 계약의 주·종피보험자ID로 계약자를 찾아 한 줄로 보여 줌 (보이는 범위(records) 안의 계약자만, 표시 전용)
      const insLines = new Map(), churnMap = new Map();
      if (custs.length) {
        const cts = await loadContracts().catch(() => []);
        custs.forEach((c) => { const k = churnOf(new Set(c.entries.map(idOf).filter(Boolean)), cts); if (k) churnMap.set(c, k); });
        const byId = new Map(); records.forEach((r) => { const id = idOf(r); if (id && !byId.has(id)) byId.set(id, r); });
        custs.forEach((c) => {
          if (!c.entries.some((e) => e.type === "피보험자")) return;
          const mine = new Set(c.entries.map(idOf).filter(Boolean));
          const holders = new Map();
          cts.forEach((ct) => {
            if (![ct.row["주피보험자ID"], ct.row["종피보험자ID"]].some((x) => mine.has(String(x || "").trim()))) return;
            const hid = String(ct.row["계약자ID"] || "").trim();
            if (!hid || mine.has(hid) || !byId.has(hid)) return;
            if (!holders.has(hid)) holders.set(hid, { id: hid, rec: byId.get(hid), goods: [] });
            holders.get(hid).goods.push(String(ct.row["상품명"] || "").trim());
          });
          if (holders.size) insLines.set(c, [...holders.values()]);
        });
      }
      const insLineHtml = (c) => {
        const hs = insLines.get(c); if (!hs) return "";
        const h = hs[0]; // 계약자 이름과 연락처만 (계약자가 여럿이면 '외 N명'만 덧붙임)
        return `<div class="ins-line"><a role="button" tabindex="0" class="ins-link" data-ins="${esc(h.id)}">${esc(h.rec.name)}</a>${h.rec.phone ? ` <span class="num">${esc(h.rec.phone)}</span>` : ""}${hs.length > 1 ? `<span class="ins-more"> 외 ${hs.length - 1}명</span>` : ""}</div>`;
      };
      const openHolder = (id) => { // 계약자 상세로 이동
        const grp = groupCustomers(records.filter((r) => idOf(r) === id))[0]; if (!grp) return;
        grp.entries.sort((a, b) => prio(a.type) - prio(b.type)); grp.name = grp.entries[0].name;
        openCustomer(grp);
      };
      const label = (want ? want.label + " " : "") + (optTypes.length ? optTypes.map((id) => OPT_LABEL[id]).join("·") + " " : "") + (gradeFilter !== "Good" ? "성향 " + gradeFilter + " " : "");
      const evs = q ? await evP : [];
      status.textContent = (custs.length ? label + custs.length + "명 찾음" : (q ? "'" + q + "' " : "") + label + (evs.length ? "고객 없음" : "고객이 없습니다.")) + (evs.length ? " · 일정 " + evs.length + "건" : "");
      $("pageSearch").hidden = custs.length > 0 || evs.length > 0 || !!force; // 못 찾으면 화면 안 검색창으로 다시 검색
      if (wide && q && !force) { openSearchPopup(q, custs, evs); return; } // 위쪽 전체검색: 결과는 팝업으로
      if (custs.length === 1 && q && (!keepList || tabReq)) { await openCustomer(custs[0]); if (tabReq) selectTab(tabReq); return; }
      const progView = optTypes.includes("progress");
      const dbOn = optTypes.includes("db"); // DB 필터가 켜져 있음(다른 필터와 함께여도)
      const dbView = !progView && dbOn; // DB 보기 칩·정렬은 진행 필터가 없을 때
      const trOn = optTypes.includes("transfer");
      const trView = !progView && !dbOn && trOn; // 이관 보기 칩은 진행·DB 필터가 없을 때
      if (trView) {
        const on = new Set(), off = new Set();
        records.forEach((r) => { if (fcPick !== FC_ALL && r.fc !== fcPick) return; if (r.type !== "이관" || !trMap.has(idOf(r))) return; (trOff(r) ? off : on).add(idOf(r) || r.name); }); // 칩 건수 = 이달 이관 고객만
        status.classList.add("has-pbar");
        status.insertAdjacentHTML("beforeend", `<span id="progBar" class="prog-bar" role="group" aria-label="이달 이관 보기"><button type="button" data-tview="on" class="${trShowRemoved ? "" : "on"}">이달 ${on.size}</button><button type="button" data-tview="off" class="${trShowRemoved ? "on" : ""}">보관 ${off.size}</button></span>`);
      }
      const progDay = new Map(); // 고객 → 정렬·표시용 날짜(YYYY-MM-DD)
      if (progView) { // 상태 줄 오른쪽 작은 칩: [진행 N] [삭제 M] (기본은 진행 고객). 한 줄을 따로 차지하지 않음
        const on = new Set(), off = new Set();
        records.forEach((r) => { if (fcPick !== FC_ALL && r.fc !== fcPick) return; const k = idOf(r) || r.name; if (r.type === "진행" || progFlag(r)) on.add(k); else if (r.type === "진행제외") off.add(k); });
        status.classList.add("has-pbar");
        status.insertAdjacentHTML("beforeend", `<span id="progBar" class="prog-bar" role="group" aria-label="진행 보기"><button type="button" data-pview="on" class="${progShowRemoved ? "" : "on"}">계약진행 ${on.size}</button><button type="button" data-pview="off" class="${progShowRemoved ? "on" : ""}">보관 ${off.size}</button></span>`);
        // 정렬: 진행 = 진행이 된 날(이력의 '단계변경 … → 진행' 일시, 없으면 고객 줄 입력일시) 최신순 / 삭제 = 진행제외 날짜 최신순
        const d10 = (v) => String(v || "").replace(/\./g, "-").slice(0, 10);
        let hist = [];
        if (!progShowRemoved) { try { hist = await loadHistory(); } catch (e) { hist = []; } }
        const stageDay = new Map();
        hist.forEach((h) => { if (String(h["유형"] || "").trim() !== "단계변경" || !/→\s*진행/.test(String(h["내용"] || ""))) return; const id = String(h["고객ID"] || "").trim(), t = d10(h["일시"]); if (id && t && (!stageDay.has(id) || stageDay.get(id) < t)) stageDay.set(id, t); });
        custs.forEach((c) => {
          let t = "";
          if (progShowRemoved) t = d10(c.entries.map((e) => fieldOf(e, "진행제외")).find(Boolean));
          else { const ids = c.entries.map(idOf).filter(Boolean); t = ids.map((id) => stageDay.get(id)).filter(Boolean).sort().pop() || d10(c.entries.map((e) => fieldOf(e, "입력일시")).filter(Boolean).sort().pop()); }
          progDay.set(c, t);
        });
        custs.sort((x, y) => (progDay.get(y) || "").localeCompare(progDay.get(x) || "") || x.name.localeCompare(y.name, "ko"));
      }
      if (dbView) { // DB 목록: 상태 줄 오른쪽 [DB N] [삭제 M] 칩. 정렬 = 유입일 최신순(삭제 목록은 삭제한 날), 같은 날은 이름순
        const on = new Set(), off = new Set();
        records.forEach((r) => { if (fcPick !== FC_ALL && r.fc !== fcPick) return; if (!isDbRec(r)) return; (dbOffRec(r) ? off : on).add(idOf(r) || r.name); });
        status.classList.add("has-pbar");
        status.insertAdjacentHTML("beforeend", `<span id="progBar" class="prog-bar" role="group" aria-label="DB 보기"><button type="button" data-dview="on" class="${dbShowRemoved ? "" : "on"}">DB ${on.size}</button><button type="button" data-dview="off" class="${dbShowRemoved ? "on" : ""}">보관 ${off.size}</button></span>`);
        const d10 = (v) => String(v || "").replace(/\./g, "-").slice(0, 10);
        custs.forEach((c) => {
          const ex = c.entries.some((e) => dbState(e) === "보관") ? (c.entries.map((e) => fieldOf(e, "DB제외")).find(Boolean) || "보관") : "";
          const auto = c.entries.some(dbAutoRec);
          c.dbAuto = dbShowRemoved && auto && !ex;
          progDay.set(c, d10(dbShowRemoved && ex && /^\d{4}/.test(ex) ? ex : c.entries.map((e) => fieldOf(e, "유입일")).find(Boolean)));
        });
        custs.sort((x, y) => (progDay.get(y) || "").localeCompare(progDay.get(x) || "") || x.name.localeCompare(y.name, "ko"));
      }
      // 진행 보기: 계약한 고객이면 '계약' 태그, 영업단계(통화·AP·PC)도 함께 보여 줌
      list.innerHTML = custs.map((c, i) => {
        let types = [...new Set(c.entries.map((e) => e.type))], stage = "";
        if (progView) {
          const ids = new Set(c.entries.map(idOf).filter(Boolean));
          const all = ids.size ? records.filter((r) => ids.has(idOf(r))) : c.entries;
          types = [...new Set(all.map((e) => e.type))];
          if (!types.includes("계약") && all.some((e) => fieldOf(e, "계약현황") === "유효")) types.push("계약");
          stage = all.map((e) => fieldOf(e, "영업단계")).find(Boolean) || "";
        }
        const fcs = [...new Set(c.entries.map((e) => e.fc).filter(Boolean))];
        const trInfo = trOn ? c.entries.map((e) => trMap.get(idOf(e))).find(Boolean) : null;
        const trDay = trView && trShowRemoved && trInfo ? trInfo.list.map((k) => String(k.row["이관보관"] || "")).sort().pop() : "";
        const pd = progView || dbView ? progDay.get(c) : trDay ? String(trDay).replace(/\./g, "-").slice(0, 10) : "";
        const pdHtml = c.dbAuto && dbView ? `<span class="r-day">계약</span>` : trDay ? `<span class="r-day">보관 ${esc(pd.slice(5))}</span>` : pd ? `<span class="r-day">${dbView ? (dbShowRemoved ? "보관 " + esc(pd.slice(5)) : "유입 " + esc(pd.slice(2))) : (progShowRemoved ? "보관" : "진행") + " " + esc(pd.slice(5))}</span>` : ""; // FC 이름 뒤에 한 칸 띄우고 작게
        // 보관 버튼: 필터가 여럿이어도 이 고객이 보관 가능할 때만 (진행 필터 → 진행 고객, DB 필터 → DB 고객). 둘 다 해당하면 진행 우선
        const ro = isRO((c.entries[0] || {}).fc);
        const srcLbl = progView ? routeShow(c.entries.map((e) => fieldOf(e, "최초유입경로")).find(Boolean) || "") : ""; // v376(2026-11-16): 진행 목록의 '진행' 태그 대신 유입경로
        const canProg = progView && !ro && c.entries.some((e) => e.sheet === CUST_SHEET && e.type === (progShowRemoved ? "진행제외" : "진행"));
        const canDb = dbOn && !ro && !canProg && !c.dbAuto && c.entries.some((e) => e.sheet === CUST_SHEET && isDbRec(e) && (dbShowRemoved ? dbState(e) === "보관" && !dbAutoRec(e) : !dbOffRec(e)));
        const canTr = trOn && !ro && !canProg && !canDb && !!trInfo && c.entries.some((e) => e.type === "이관") && trInfo.st === (trShowRemoved ? "off" : "on");
        return `<li${progView || dbOn || trOn ? ' class="pv"' : ""}><button data-i="${i}"><span class="avatar sm">${esc(c.name[0])}</span>
          <span class="r-name"><span><span class="cname ${gradeClass(gradeOf(c))}">${esc(c.name)}</span>${resultInfo(c) ? ` <span class="rinfo">${esc(resultInfo(c))}</span>` : ""}</span>${fcs.length || pdHtml ? `<span class="r-sub">${esc(fcs.join(", "))}${pdHtml}</span>` : ""}${c.why && c.why.size ? `<span class="r-why">${[...c.why].map(([k, v]) => esc(k) + (v && k !== "연락처" ? " · " + esc(v) : "")).join(" / ")}</span>` : ""}</span>
          ${churnTagHtml(churnMap.get(c))}${types.map((t) => progView && t === "진행" ? (srcLbl ? `<span class="tag t-진행">${esc(srcLbl)}</span>` : "") : `<span class="tag t-${esc(t)}">${esc(typeLabel(t))}</span>`).join("")}${stage ? `<span class="tag stg stg-${esc(stage)}">${esc(stage)}</span>` : ""}<span class="chev">›</span></button>${canTr ? `<button type="button" class="prog-mini" data-trc="${i}" aria-label="${esc(c.name)} ${trShowRemoved ? "이관으로 원복" : "보관으로 이동"}">${trShowRemoved ? "원복" : "이동"}</button>` : ""}${canDb ? `<button type="button" class="prog-mini" data-dbc="${i}" aria-label="${esc(c.name)} ${dbShowRemoved ? "DB로 원복" : "보관으로 이동"}">${dbShowRemoved ? "원복" : "이동"}</button>` : ""}${canProg ? (progShowRemoved ? `<button type="button" class="prog-mini" data-prest="${i}" aria-label="${esc(c.name)} 진행으로 원복">원복</button>` : `<button type="button" class="prog-mini" data-pdel="${i}" aria-label="${esc(c.name)} 보관으로 이동">이동</button>`) : ""}${insLineHtml(c)}</li>`;
      }).join("");
      list.querySelectorAll("button[data-i]").forEach((b) =>
        b.addEventListener("click", async () => { await openCustomer(custs[+b.dataset.i]); if (tabReq) selectTab(tabReq); }));
      list.querySelectorAll("[data-prest]").forEach((b) => b.addEventListener("click", (ev) => { ev.stopPropagation(); progressRestore(custs[+b.dataset.prest], () => search(curQuery(), true)); }));
      const pbar = $("progBar"); if (pbar) pbar.addEventListener("click", (ev) => { const b = ev.target.closest("[data-pview]"), d = ev.target.closest("[data-dview]"), t = ev.target.closest("[data-tview]"); if (b) progShowRemoved = b.dataset.pview === "off"; else if (d) dbShowRemoved = d.dataset.dview === "off"; else if (t) trShowRemoved = t.dataset.tview === "off"; else return; search(curQuery(), true); });
      list.querySelectorAll("[data-trc]").forEach((b) => b.addEventListener("click", (ev) => { ev.stopPropagation(); const c = custs[+b.dataset.trc]; const t = c.entries.map((e) => trMap.get(idOf(e))).find(Boolean); if (t) trArchive(c.name, t.list, !trShowRemoved, () => search(curQuery(), true)); }));
      list.querySelectorAll("[data-dbc]").forEach((b) => b.addEventListener("click", (ev) => { ev.stopPropagation(); const c = custs[+b.dataset.dbc], ids = new Set(c.entries.map(idOf).filter(Boolean)); const all = ids.size ? records.filter((r) => ids.has(idOf(r))) : c.entries; (dbShowRemoved ? dbRestore : dbRemove)(c.name, all, () => search(curQuery(), true)); }));
      list.querySelectorAll("[data-pdel]").forEach((b) => b.addEventListener("click", (ev) => { ev.stopPropagation(); progressRemove(custs[+b.dataset.pdel], () => search(curQuery(), true)); }));
      if (evs.length) { // 일정에서 찾은 것: 날짜·제목 줄, 누르면 그날 일정 화면
        const WDK = ["일", "월", "화", "수", "목", "금", "토"];
        list.insertAdjacentHTML("afterend", `<div id="evHits" class="ev-hits"><h3 class="sec-title">일정 ${evs.length}건 <span class="hint">(1년 전 ~ 1개월 뒤)</span></h3><ul class="list">${evs.map((e, i) => `<li><button type="button" data-evgo="${i}"><span class="r-name"><span class="cname">${esc(e.title)}</span><span class="r-sub">${e.start.getFullYear()}.${pad2(e.start.getMonth() + 1)}.${pad2(e.start.getDate())}(${WDK[e.start.getDay()]}) ${e.allDay ? "종일" : hm(e.start)} · ${esc(e.cal)}${e.location ? " · " + esc(e.location) : ""}</span></span><span class="chev">›</span></button></li>`).join("")}</ul></div>`);
        $("evHits").querySelectorAll("[data-evgo]").forEach((b) => b.addEventListener("click", () => { const e = evs[+b.dataset.evgo]; schedMode = "day"; schedDate = dayStart(e.start); openSchedule(); }));
      }
      list.querySelectorAll("[data-ins]").forEach((a) => {
        const go = (ev) => { ev.stopPropagation(); openHolder(a.dataset.ins); };
        a.addEventListener("click", go);
        a.addEventListener("keydown", (ev) => { if (ev.key === "Enter" || ev.key === " ") { ev.preventDefault(); go(ev); } });
      });
    } catch (e) {
      status.textContent = "";
      handleError(e);
    }
  }

  // ---------- 계약 상세 (계약리스트 / 이관리스트) ----------
  const CONTRACT_NAME_COLS = ["계약자", "주피보험자", "피보", "피보험자"];
  let contractCache = null;
  function loadContracts() {
    if (contractCache) return contractCache;
    contractCache = (async () => {
      const { files } = await scanFiles();
      const sources = (CFG.contractLists || []).flatMap((l) =>
        pickSourceFiles(files, l).map((f) => ({ ...l, file: f })));
      const out = [], newCtDone = new Set();
      await Promise.all(sources.map(async (s) => {
        const wb = await readBook(s.file.id);
        const fc = s.file.fc || (s.file.name.split(" - ")[1] || "").trim();
        // 새 구조 파일: '계약' 시트 하나(직접모집·타인모집 함께)만 파일당 한 번, 예전 '이관' 시트는 건너뜀
        const newMode = isNewStruct(wb);
        if (newMode) { if (newCtDone.has(s.file.id) || s.sheet !== "계약") return; newCtDone.add(s.file.id); }
        const ids = newMode ? idNameMap(wb) : null;
        sheetsOf(wb, s).forEach((sheet) => {
          const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheet], { header: 1, defval: "", raw: false });
          const h = rows.findIndex((r) => r.some((c) => ["계약자", "계약자ID"].includes(String(c).trim())));
          if (h < 0) return;
          const head = rows[h].map((c) => String(c).trim());
          rows.slice(h + 1).forEach((r) => {
            const fields = head.map((k, i) => [k, String(r[i] ?? "").trim()]).filter(([k, v]) => k && v);
            if (!fields.length) return;
            const row = newMode ? contractRowCompat(Object.fromEntries(fields), ids) : Object.fromEntries(fields);
            if (newMode) { // 계약자·주피보험자 이름도 필드로 (계약 탭 표시용), 타인모집은 '이관'
              ["계약자", "주피보험자"].forEach((k) => { if (row[k] && !fields.some(([fk]) => fk === k)) fields.push([k, row[k]]); });
              s = { ...s, label: contractLabel(row["모집유형"]) };
            }
            const names = CONTRACT_NAME_COLS.map((c) => row[c]).filter(Boolean).map(norm);
            // 줄에 담당FC가 있으면 그 FC로 표시 (다른 FC 파일에 같이 들어 있는 계약)
            const rowFc = String(row["담당FC"] || row["담당 FC"] || "").trim() || fc;
            if (names.length) out.push({ type: s.label, fc: rowFc, srcFc: fc, sheet, row, fields, names, src: { fileId: s.file.id, sheet } });
          });
        });
      }));
      // 같은 증권번호가 여러 FC 파일에 있으면 한 번만 (담당 FC 파일의 것을 우선)
      const seen = new Map();
      out.forEach((c, i) => {
        const no = String(c.row["증권번호"] || "").trim();
        if (!no) return;
        const key = c.type + "|" + no;
        if (!seen.has(key)) seen.set(key, i);
      });
      return out.filter((c, i) => {
        const no = String(c.row["증권번호"] || "").trim();
        return !no || seen.get(c.type + "|" + no) === i;
      });
    })();
    contractCache.catch(() => { contractCache = null; });
    return contractCache;
  }

  // ---------- 고객 추가 정보 (고객 폴더 안 '고객추가정보' 엑셀) ----------
  //   '기본' 시트: 담당 FC + 고객명으로 줄을 찾음. 나머지 모든 열을 읽고/수정/저장
  //   '상담기록' 시트: 고객명 + 담당 FC로 기록을 찾고, 새 기록은 맨 아래 줄에 추가
  const EXTRA = Object.assign({
    fileKeyword: "고객추가정보", sheet: "기본", fcCol: "담당 FC", nameCol: "고객명",
    primary: ["미팅장소", "메모"],             // 항상 보이는 항목
    carrierCol: "통신사",                      // 연락처 앞에 표시
    afterAddress: ["차량번호", "자동차보험갱신일"], // 기본정보 주소 밑에 표시
    monthDayCols: ["자동차보험갱신일"],          // 월/일을 골라서 'MMDD'로 저장
    consult: { sheet: "상담기록", nameCol: "고객명", fcCol: "담당 FC", dateCol: ["저장날짜시간", "저장날짜"], textCol: "내용", voiceCol: "음성파일", voiceText: "음성기록" }
  }, CFG.extraInfo || {});
  const nospace = (s) => String(s || "").replace(/\s+/g, "");

  async function readExtra() {
    await resolveBase();
    let items = [];
    try { items = await stList(base.id); } catch (e) { if (!/40[34]/.test(String(e.message))) throw e; } // FC 폴더만 공유받은 사용자는 루트 목록이 없음
    const file = items.find((i) => i.file && /\.xlsx$/i.test(i.name) && nospace(i.name).includes(nospace(EXTRA.fileKeyword)));
    if (!file) return null;
    const wb = await stGetBook(file.id); // 서식 유지 (다시 저장할 때 #,##0·날짜 서식이 사라지지 않게)
    const sheetName = wb.SheetNames.includes(EXTRA.sheet) ? EXTRA.sheet : wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: false, blankrows: true });
    const head = (rows[0] || []).map((c) => String(c).trim());
    return { file, wb, ws, rows, head };
  }

  const STAMP_COL = "입력일시"; // 줄이 처음 들어간 시각 (엑셀 날짜·시간, yyyy-mm-dd hh:mm). 새 줄에만 넣고 고칠 때는 그대로
  const extraCols = (x) => x.head.filter((h) => h && h !== EXTRA.fcCol && h !== EXTRA.nameCol && h !== STAMP_COL);
  const sameFc = (rowFc, fc) => !fc || !String(rowFc || "").trim() || norm(rowFc) === norm(fc);

  // (담당 FC, 고객명)이 같은 줄 찾기. 엑셀에 FC가 비어 있으면 이름만으로 매칭
  function extraRowIndex(x, fc, name) {
    const ci = x.head.indexOf(EXTRA.fcCol), ni = x.head.indexOf(EXTRA.nameCol);
    if (ni < 0) return -1;
    for (let r = 1; r < x.rows.length; r++) {
      const row = x.rows[r];
      if (norm(row[ni] ?? "") === norm(name) && sameFc(ci < 0 ? "" : row[ci], fc)) return r;
    }
    return -1;
  }

  // 월일 값 해석: "0326", "3/26", "03-26", "3월 26일", "2026-03-26", 20260326 → [3, 26]
  function parseMonthDay(v) {
    const s = String(v ?? "").trim();
    if (!s) return null;
    let m = s.match(/(\d{1,2})\s*월\s*(\d{1,2})/);
    if (!m) m = s.match(/^\d{4}[-./](\d{1,2})[-./](\d{1,2})$/);
    if (!m) m = s.match(/^(\d{1,2})[-./](\d{1,2})$/);
    if (!m) { const d = s.replace(/\D/g, ""); if (d.length === 8) m = [0, d.slice(4, 6), d.slice(6, 8)]; else if (d.length === 4 || d.length === 3) m = [0, d.slice(0, d.length - 2), d.slice(-2)]; }
    if (!m) return null;
    const mm = +m[1], dd = +m[2];
    return mm >= 1 && mm <= 12 && dd >= 1 && dd <= 31 ? [mm, dd] : null;
  }
  const monthDayText = (v) => { const p = parseMonthDay(v); return p ? `${p[0]}월 ${p[1]}일` : String(v || ""); };
  const pad2 = (n) => String(n).padStart(2, "0");

  // 상담기록: 날짜(시간) 값 → { key: 정렬용 YYYYMMDDHHmm, text: "2026.11.08 15:30" }
  //   20261108(숫자) · 엑셀 날짜/시간 · "2026-11-08 15:30" 글자 모두 읽음
  function consultDate(cell) {
    if (!cell) return { key: 0, text: "" };
    const v = cell.v;
    let y, m, d, hh = null, mi = null;
    if (typeof v === "number" && v > 190000000000) { // 202611081530
      y = Math.floor(v / 1e8); m = Math.floor(v / 1e6) % 100; d = Math.floor(v / 1e4) % 100; hh = Math.floor(v / 100) % 100; mi = v % 100;
    } else if (typeof v === "number" && v > 19000000) { // 20261108
      y = Math.floor(v / 10000); m = Math.floor(v / 100) % 100; d = v % 100;
    } else if (typeof v === "number") { // 엑셀 날짜(시간) 일련번호
      const t = new Date(Date.UTC(1899, 11, 30) + Math.round(v * 86400) * 1000);
      y = t.getUTCFullYear(); m = t.getUTCMonth() + 1; d = t.getUTCDate();
      if (v % 1) { hh = t.getUTCHours(); mi = t.getUTCMinutes(); }
    } else {
      const s = String(v).match(/(\d{4})\D?(\d{1,2})\D?(\d{1,2})(?:\D+(\d{1,2})\D?(\d{2}))?/);
      if (!s) return { key: 0, text: String(v) };
      [y, m, d] = [+s[1], +s[2], +s[3]];
      if (s[4] != null) { hh = +s[4]; mi = +s[5]; }
    }
    const key = (y * 10000 + m * 100 + d) * 10000 + (hh == null ? 0 : hh * 100 + mi);
    return { key, text: `${y}.${pad2(m)}.${pad2(d)}` + (hh == null ? "" : ` ${pad2(hh)}:${pad2(mi)}`) };
  }

  // Date → 엑셀 날짜·시간 일련번호 (휴대폰 시간 그대로)
  const toExcelDateTime = (dt) =>
    (Date.UTC(dt.getFullYear(), dt.getMonth(), dt.getDate(), dt.getHours(), dt.getMinutes()) - Date.UTC(1899, 11, 30)) / 86400000;

  function consultSheetInfo(ws) {
    const C = EXTRA.consult;
    const rg = XLSX.utils.decode_range(ws["!ref"] || "A1:D1");
    const head = [];
    for (let c = rg.s.c; c <= rg.e.c; c++) { const cell = ws[XLSX.utils.encode_cell({ r: rg.s.r, c })]; head[c] = cell ? String(cell.v).trim() : ""; }
    // 열 이름 후보 여러 개 허용 (예: 날짜 열 = "저장날짜시간" 또는 "저장날짜")
    const col = (h) => { for (const x of [].concat(h)) { const i = head.indexOf(x); if (i >= 0) return i; } return -1; };
    return { rg, head, ni: col(C.nameCol), fi: col(C.fcCol), di: col(C.dateCol), ti: col(C.textCol), vi: col(C.voiceCol) };
  }

  function readConsults(wb, fc, name) {
    const ws = wb.Sheets[EXTRA.consult.sheet];
    if (!ws || !ws["!ref"]) return [];
    const { rg, ni, fi, di, ti, vi } = consultSheetInfo(ws);
    if (ni < 0 || ti < 0) return [];
    const at = (r, c) => c < 0 ? null : ws[XLSX.utils.encode_cell({ r, c })];
    const out = [];
    for (let r = rg.s.r + 1; r <= rg.e.r; r++) {
      const nm = at(r, ni), tx = at(r, ti);
      if (!nm || !tx || norm(nm.v) !== norm(name)) continue;
      if (!sameFc(at(r, fi) ? at(r, fi).v : "", fc)) continue;
      const d = consultDate(at(r, di));
      const voice = at(r, vi) ? String(at(r, vi).v).trim() : "";
      out.push({ row: r, date: d.text, key: d.key, text: String(tx.v), voice });
    }
    return out.sort((a, b) => b.key - a.key || a.row - b.row); // 최근 날짜·시간이 위로 (같으면 엑셀 순서)
  }

  // 새 구조: 메모·통신사·차량번호 등은 '고객' 시트의 같은 이름 열이 있으면 거기서 읽고 쓴다 (없는 값은 옛 고객추가정보 파일에서)
  const sheetExtraKeys = () => [...new Set([...EXTRA.primary, EXTRA.carrierCol, ...EXTRA.afterAddress])].filter(Boolean);
  const sheetEntriesOf = (cust) => (cust ? ownEntries(cust.entries).filter((e) => e.sheet === CUST_SHEET && e.src) : []);
  async function loadExtraFor(fc, name, cust) {
    const x = await readExtra();
    let base;
    if (!x) base = { missing: true, fields: [], values: {}, consults: [] };
    else {
      const r = extraRowIndex(x, fc, name);
      const fields = extraCols(x);
      const values = {};
      fields.forEach((k) => { const c = x.head.indexOf(k); values[k] = r > 0 ? String(x.rows[r][c] ?? "").trim() : ""; });
      base = { found: r > 0, fields, values, fileName: x.file.name, consults: readConsults(x.wb, fc, name) };
    }
    const own = sheetEntriesOf(cust);
    if (!own.length) return base;
    const keys = [...new Set([...sheetExtraKeys(), ...base.fields])];
    const values = { ...base.values };
    keys.forEach((k) => { const sv = own.map((e) => fieldOf(e, k)).find(Boolean); if (sv) values[k] = sv; else if (!(k in values)) values[k] = ""; });
    return { ...base, found: true, missing: false, fields: keys, values, sheetMode: true };
  }

  // 날짜처럼 보이는 값(2026-12-01, 26-12-01)은 엑셀 날짜로 저장 (월일 열은 'MMDD' 글자 그대로)
  function toCell(col, v) {
    if (v === "" || v == null) return null;
    if (EXTRA.monthDayCols.includes(col)) return { t: "s", v: String(v) };
    const m = /일$|날짜|date/i.test(col) && String(v).trim().match(/^(\d{4}|\d{2})[-./](\d{1,2})[-./](\d{1,2})$/);
    if (m) {
      const y = m[1].length === 2 ? 2000 + +m[1] : +m[1];
      const serial = (Date.UTC(y, +m[2] - 1, +m[3]) - Date.UTC(1899, 11, 30)) / 86400000;
      return { t: "n", v: serial, z: "yyyy-mm-dd" };
    }
    return { t: "s", v: String(v) };
  }

  async function uploadExtra(x) {
    await stPutBook(x.file.id, x.wb);
    delete bookCache[x.file.id];
  }

  async function saveExtraFor(fc, name, values, before) {
    const own = sheetEntriesOf(current && current.cust && current.cust.name === name ? current.cust : null);
    if (own.length) { // 새 구조: 바뀐 칸만 '고객' 시트(담당 FC 행)에 저장
      const changed = {};
      Object.keys(values).forEach((k) => { if ((values[k] ?? "") !== (before[k] ?? "")) changed[k] = values[k] ?? ""; });
      if (Object.keys(changed).length) await writeEntryCells(own, changed);
      return;
    }
    const x = await readExtra(); // 저장 직전 최신 파일을 다시 읽어 이 고객 줄의 바뀐 칸만 수정
    if (!x) throw new Error("NO_EXTRA_FILE");
    let r = extraRowIndex(x, fc, name);
    const isNew = r < 0;
    if (isNew) r = Math.max(x.rows.length, 1); // 없으면 맨 아래에 새 줄
    const setCell = (col, v) => {
      const c = x.head.indexOf(col);
      if (c < 0) return;
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = toCell(col, v);
      if (cell) x.ws[addr] = cell; else delete x.ws[addr];
    };
    if (isNew) {
      setCell(EXTRA.fcCol, fc); setCell(EXTRA.nameCol, name);
      const sc = x.head.indexOf(STAMP_COL);
      if (sc >= 0) x.ws[XLSX.utils.encode_cell({ r, c: sc })] = { t: "n", v: toExcelDateTime(new Date()), z: "yyyy-mm-dd hh:mm" };
    }
    Object.keys(values).forEach((k) => {
      if (isNew || (values[k] ?? "") !== (before[k] ?? "")) setCell(k, values[k] ?? "");
    });
    const range = XLSX.utils.decode_range(x.ws["!ref"] || "A1");
    range.e.r = Math.max(range.e.r, r);
    range.e.c = Math.max(range.e.c, x.head.length - 1);
    x.ws["!ref"] = XLSX.utils.encode_range(range);
    await uploadExtra(x);
  }

  // 상담기록 저장: ① ✕로 표시한 기록 삭제 ② 새 기록 추가(날짜·시간) ③ 시트 전체를 날짜·시간 최신순으로 정렬
  //   adds: [{ text, when: Date, voice: "파일명" }]
  async function saveConsults(fc, name, adds, dels) {
    if (dels && dels.length && syncLockedFc(fc)) throw new Error("데이터 업데이트 중이라 상담기록 수정·삭제는 끝난 뒤에 해 주세요."); // v550
    const C = EXTRA.consult;
    const x = await readExtra(); // 저장 직전 최신 파일을 다시 읽음
    if (!x) throw new Error("NO_EXTRA_FILE");
    let ws = x.wb.Sheets[C.sheet];
    if (!ws) {
      ws = XLSX.utils.aoa_to_sheet([[[].concat(C.dateCol)[0], C.nameCol, C.fcCol, C.textCol, C.voiceCol]]);
      XLSX.utils.book_append_sheet(x.wb, ws, C.sheet);
    }
    const enc = (r, c) => XLSX.utils.encode_cell({ r, c });
    let info = consultSheetInfo(ws);
    // 음성파일 열이 없으면 머리글에 추가
    if (adds.some((a) => a.voice) && info.vi < 0) {
      const c = info.rg.e.c + 1;
      ws[enc(info.rg.s.r, c)] = { t: "s", v: C.voiceCol };
      info.rg.e.c = c;
      ws["!ref"] = XLSX.utils.encode_range(info.rg);
      info = consultSheetInfo(ws);
    }
    const { rg, ni, fi, di, ti, vi } = info;
    const width = rg.e.c - rg.s.c + 1;

    // 본문 줄을 [셀...] 배열로 꺼냄
    let rows = [];
    for (let r = rg.s.r + 1; r <= rg.e.r; r++) {
      const cells = [];
      let any = false;
      for (let c = rg.s.c; c <= rg.e.c; c++) { const cell = ws[enc(r, c)]; cells.push(cell || null); if (cell) any = true; }
      if (any) rows.push(cells);
    }
    const cellOf = (row, c) => (c < 0 ? null : row[c - rg.s.c]);

    // ① 삭제: 같은 고객·FC·날짜·내용인 줄
    dels.forEach((d) => {
      const i = rows.findIndex((row) => {
        const nm = cellOf(row, ni), tx = cellOf(row, ti), f = cellOf(row, fi);
        return nm && tx && norm(nm.v) === norm(name) && String(tx.v) === d.text &&
          sameFc(f ? f.v : "", fc) && consultDate(cellOf(row, di)).key === d.key;
      });
      if (i >= 0) rows.splice(i, 1);
    });

    // ② 추가
    adds.forEach((a) => {
      const row = new Array(width).fill(null);
      const put = (c, cell) => { if (c >= 0) row[c - rg.s.c] = cell; };
      put(ni, { t: "s", v: name });
      put(fi, { t: "s", v: fc });
      put(di, { t: "n", v: toExcelDateTime(a.when || new Date()), z: "yyyy-mm-dd hh:mm" });
      put(ti, { t: "s", v: a.text });
      if (a.voice) put(vi, { t: "s", v: a.voice });
      rows.push(row);
    });

    // ③ 최신 날짜·시간이 위로 (같으면 기존 순서 유지)
    rows = rows.map((row, i) => ({ row, i, key: consultDate(cellOf(row, di)).key }))
      .sort((a, b) => b.key - a.key || a.i - b.i).map((o) => o.row);

    for (let r = rg.s.r + 1; r <= rg.e.r; r++) for (let c = rg.s.c; c <= rg.e.c; c++) delete ws[enc(r, c)];
    rows.forEach((row, i) => row.forEach((cell, j) => { if (cell) ws[enc(rg.s.r + 1 + i, rg.s.c + j)] = cell; }));
    rg.e.r = rg.s.r + rows.length;
    ws["!ref"] = XLSX.utils.encode_range(rg);
    await uploadExtra(x);
    return readConsults(x.wb, fc, name);
  }

  // PDF는 OneDrive 웹 대신 휴대폰의 PDF 보기 앱/브라우저로 열기 (임시 다운로드 주소 사용)
  //   PDF 내용을 받아 브라우저 PDF 보기로 엶 (안드로이드는 'PDF 앱으로 열기' 선택 창)
  async function openFileDirect(id) {
    const w = window.open("", "_blank"); // 팝업 차단을 피하려고 먼저 창을 엶
    if (w) w.document.write('<p style="font-family:sans-serif;padding:24px">PDF를 여는 중…</p>');
    try {
      const buf = await stGet(id);
      const url = URL.createObjectURL(new Blob([buf], { type: "application/pdf" }));
      if (w) w.location.href = url; else location.href = url;
      setTimeout(() => URL.revokeObjectURL(url), 5 * 60 * 1000);
    } catch (e) {
      if (w) w.close();
      handleError(e);
    }
  }

  // ---------- 고객 성향: 이름 색 + 다른 등급 동그라미 버튼 ----------
  function renderGrade(cust) {
    const g = gradeOf(cust) || "Good";
    $("custName").className = gradeClass(g) + " cname";
    const box = $("gradeBtns");
    box.hidden = !editMode;
    box.innerHTML = editMode ? '<span class="grade-title">성향</span>' + GRADES.map((x) =>
      `<button type="button" class="grade-opt gl-${x.id}${x.id === g ? " on" : ""}" data-grade="${x.id}" aria-label="성향 ${x.label}" aria-pressed="${x.id === g}"><span class="grade-dot">${x.id.charAt(0)}</span></button>`).join("") : "";
    box.querySelectorAll("[data-grade]").forEach((b) => b.addEventListener("click", () => { if (b.dataset.grade !== g) saveGrade(cust, b.dataset.grade); }));
  }

  async function saveGrade(cust, g) {
    const before = gradeOf(cust);
    // 화면 먼저 바꾸고 엑셀에 저장
    cust.entries.forEach((e) => (e.grade = g));
    renderGrade(cust);
    try {
      await writeEntryCells(cust.entries, { [GRADE_COL]: g });
      toast(`성향을 ${g}(으)로 저장했습니다.`);
    } catch (e) {
      cust.entries.forEach((x) => (x.grade = before));
      renderGrade(cust);
      saveError(e);
    }
  }

  // ---------- 고객 화면 ----------
  let current = null; // { cust, folders, files, phones, contracts, fc, extra }

  let returnView = "searchView"; // 고객 정보에서 뒤로 가면 돌아갈 화면 (일정 / 고객 검색)
  async function openCustomer(cust) {
    if (!$("salesView").hidden) SALES.cardAnchor = salesAnchor(); // 세일즈 화면에서 열면 돌아올 때 이 위치로
    returnView = !$("scheduleView").hidden ? "scheduleView" : !$("salesView").hidden ? "salesView" : !$("introView").hidden ? "introView" : "searchView";
    show("customerView");
    $("custName").innerHTML = trophyOf(cust) + esc(cust.name);
    editMode = false;
    renderGrade(cust);
    renderStageAct(cust);
    $("avatar").textContent = cust.name[0];
    $("avatar").style.backgroundImage = "";
    editMode = false;
    $("editBtn").textContent = "수정";
    $("editBtn").hidden = false;
    $("editBtn").classList.remove("on");
    $("editBtns").hidden = true;
    $("closeCust").hidden = false;
    $("photoMenu").hidden = true;
    $("avatar").classList.remove("editable");
    consultShowAll = false; // 상담기록은 처음엔 최근 10건
    $("avatar").classList.remove("has-photo");
    $("custSub").innerHTML = custSubHtml(cust);
    { const tel = custTelHtml(cust); $("custTel").innerHTML = tel; $("custTel").hidden = !tel; } // 명함 카드 안 전화 줄(표시만 추가, 기능 그대로)
    $("tabBody").innerHTML = '<p class="hint pad">고객 정보를 불러오는 중…</p>';
    $("tabs").innerHTML = "";
    history.pushState({ v: "customer" }, "");
    current = { cust, folders: [], files: [], phones: null, contracts: [], fc: "", extra: { values: {} }, profilePhoto: null, photoUrl: null };
    notifyWhereNow();
    showCachedPhoto(current); // 기기에 저장된 사진이 있으면 바로
    try {
      const names = new Set(cust.entries.map((e) => norm(e.name)).concat(norm(cust.name)));
      // 새 구조: 시트의 '폴더' 칸이 가리키는 폴더 우선, 없으면 예전처럼 이름으로
      const linked = cust.entries.map((e) => { const f = e.fields.find(([k]) => k.trim() === "폴더"); return f ? norm(f[1]) : ""; }).filter(Boolean);
      let allFolders = await loadCustomerFolders();
      { // v556: 폴더 목록은 앱을 연 뒤 한 번만 읽어 둠 → 그 뒤 새로 만든 고객 폴더(오늘 등록·보장분석)를 못 찾으면 목록을 한 번 새로 읽음
        const cid0 = cust.entries.map((e) => idOf(e)).filter(Boolean);
        const hit = allFolders.some((f) => linked.includes(norm(f.name)) || cid0.some((id) => String(f.name).endsWith(" - " + id)));
        if ((linked.length || cid0.length) && !hit && Date.now() - scanAt > 60000) { scanCache = null; allFolders = await loadCustomerFolders(); } // 폴더가 아예 없는 고객 때문에 자꾸 다시 읽지 않게 1분에 한 번만
      }
      const byLink =linked.length ? allFolders.filter((f) => linked.includes(norm(f.name))) : [];
      const cids = cust.entries.map((e) => idOf(e)).filter(Boolean); // v494: '폴더' 칸이 비어 있으면 폴더 이름 끝의 고객ID('이름 (생년월일) - C01249')로 찾음
      const byId = !byLink.length && cids.length ? allFolders.filter((f) => cids.some((id) => String(f.name).endsWith(" - " + id))) : [];
      const byName = byLink.length ? byLink : byId.length ? byId : allFolders.filter((f) => names.has(norm(custPart(f.name))));
      // 같은 이름이 여러 FC에 있으면 이 고객의 담당 FC 폴더만
      const fcs = cust.entries.map((e) => e.fc).filter(Boolean);
      const sameFc = byName.filter((f) => f.fc && fcs.includes(f.fc));
      // 새 규칙(_고객)의 폴더를 먼저, 그다음 FC 바로 아래, 옛 분류 폴더 순 (사진 올리기 등은 첫 폴더에 저장)
      const gRank = (f) => (f.group === "고객" ? 0 : f.group === "" ? 1 : 2);
      const folders = [...(sameFc.length ? sameFc : byName)].sort((a, b) => gRank(a) - gRank(b));
      let files = [];
      for (const f of folders) files = files.concat(await collectFiles(f.id, folders.length > 1 ? f.name + "/" : "", 0));
      current.folders = folders;
      current.files = files;
      current.profilePhoto = findProfilePhoto(files);
      if (current.profilePhoto) loadProfilePhoto(); // 머리 부분 동그라미에 사진 (기기 저장 사진 확인 → 필요할 때만 받음)
      else if (current.cachedPhoto) { // 사진이 지워짐 → 기기 저장도 지움
        photoCacheDel(photoKey(current.cust));
        current.cachedPhoto = null;
        if (current.photoUrl) URL.revokeObjectURL(current.photoUrl);
        current.photoUrl = null;
        const a = $("avatar"); a.style.backgroundImage = ""; a.classList.remove("has-photo"); a.textContent = current.cust.name[0];
      }
      current.fc = primaryFc(cust.entries) || fcs[0] || (folders[0] && folders[0].fc) || "";
      $("editBtn").hidden = isRO(current.fc); // 보기만 권한 FC는 수정 버튼 숨김
      $("editBtn").hidden = isRO(current.fc); // 보기 전용 FC는 수정 버튼 숨김
      const [contracts, extra] = await Promise.all([
        loadContracts().catch((e) => { console.error(e); return []; }),
        loadExtraFor(current.fc, cust.name, cust).catch((e) => { console.error(e); return { error: true, values: {} }; })
      ]);
      const phones = new Set(cust.entries.map((e) => String(e.phone || "").replace(/\D/g, "")).filter(Boolean));
      const custIds = new Set(cust.entries.map(idOf).filter(Boolean));
      current.contracts = contracts.filter((c) => {
        const rid = [c.row["계약자ID"], c.row["주피보험자ID"]].filter(Boolean);
        if (custIds.size && rid.length) return rid.some((x) => custIds.has(String(x).trim())); // 새 구조: 고객ID로 정확히 연결
        if (!c.names.includes(norm(baseName(cust.name)))) return false;
        if (fcs.length && !fcs.includes(c.fc) && !fcs.includes(c.srcFc)) return false;
        const cp = String(c.row["연락처"] || "").replace(/\D/g, "");
        return !cp || !phones.size || phones.has(cp); // 이관 계약은 계약자 연락처로 동명이인 구분
      });
      { // 정보 탭 '유지 계약'·'계약현황' 표시용: 이 고객이 피보험자(주·종)이고 계약자는 다른 사람인 유효 계약 (고객 시트 값은 계약자 기준이라 그대로 둠)
        const VALID = /^(유지|완납|납입면제|연금지급|신계약승낙)$/;
        const ins = contracts.filter((c) => {
          if (!VALID.test(String(c.row["계약상태"] || c.row["상태"] || "").trim())) return false;
          if (custIds.has(String(c.row["계약자ID"] || "").trim())) return false;
          return [c.row["주피보험자ID"], c.row["종피보험자ID"]].some((x) => custIds.has(String(x || "").trim()));
        });
        current.insCt = { n: ins.length, holders: [...new Set(ins.map((c) => String(c.row["계약자"] || "").trim()).filter(Boolean))] };
      }
      current.extra = extra;
      { // 해지 고객 표시(머리 카드): 계약을 불러온 뒤 판단
        const k = churnOf(custIds, contracts), tg = $("custChurn"), note = $("custNote");
        tg.className = k ? `tag t-${k}` : "tag"; tg.textContent = k === "해지" ? "해지" : k === "실효" ? "실효" : ""; tg.hidden = !k;
        const noInfo = !cust.entries.some((e) => fieldOf(e, "연락처") || entryAddr(e));
        note.textContent = k === "해지" && noInfo ? "해지 고객이라 정보 조회 불가" : ""; note.hidden = !note.textContent;
      }
      try { // 따행 시트(B): 이 고객의 대상년월별 기록, 최신 월부터
        const fw = await loadFollow(), ids = new Set(cust.entries.map(idOf).filter(Boolean));
        current.follows = fw.rows.filter((r) => ids.has(r["고객ID"])).sort((a, b) => ymOfFollow(b["대상년월"]).localeCompare(ymOfFollow(a["대상년월"])));
        current.followIdx = 0;
      } catch (e) { current.follows = []; }
    } catch (e) {
      handleError(e);
    }
    renderTabs();
    selectTab("basic");
  }

  document.addEventListener("change", (ev) => {
    if (ev.target && ev.target.id === "hideRef") {
      try { localStorage.setItem("shlife.hideRefContracts", ev.target.checked ? "1" : "0"); } catch (e) { /* 무시 */ }
      if (current) selectTab("contract");
    }
  });

  // 고객 폴더 안 파일 (하위 폴더 2단계까지)
  async function collectFiles(id, prefix, depth) {
    const items = await stList(id);
    let out = [];
    for (const it of items) {
      if (it.folder && depth < 2) {
        out = out.concat(await collectFiles(it.id, prefix + it.name + "/", depth + 1));
      } else if (it.file) {
        out.push({ ...it, path: prefix + it.name });
      }
    }
    return out;
  }

  function filesForTab(tab) {
    if (tab.id === "all") return current.files;
    const kws = tab.keywords.map((k) => k.toLowerCase());
    const ex = tab.exclude ? new RegExp(tab.exclude, "i") : null;
    const out = current.files.filter((f) => kws.some((k) => f.path.toLowerCase().includes(k)) && !(ex && ex.test(f.name)) && !DOCS_RE.test(f.path)); // 자료 폴더의 파일은 자료 탭에서만
    if (tab.id !== "analysis") return out;
    // 분석 탭: 워드(docx) 파일은 숨기고, ① AI/클로드 분석 보고서 ② 스마트보장분석 ③ 손해담보 ④ 그 밖의 순으로 (같은 종류는 파일 이름 앞 날짜가 최근인 것부터)
    const rank = (f) => { const n = f.name.replace(/\s/g, ""); return /AI.*분석|클로드|claude|보장분석보고서|심화분석/i.test(n) ? 0 : /스마트보장분석/.test(n) ? 1 : /손해담보/.test(n) ? 2 : 3; };
    // v377(데이터수집 전달, 사용자 지시): 분석 탭은 'old'(예전 '과거 - 날짜') 하위 폴더의 옛 파일을 숨기고 고객 폴더의 최신 파일만
    return out.filter((f) => kind(f.name) !== "word" && !/(^|\/)(old|과거 - [^/]*)\//i.test(f.path))
      .map((f, i) => [f, i]).sort((a, b) => rank(a[0]) - rank(b[0]) || String(b[0].name).localeCompare(String(a[0].name)) || a[1] - b[1]).map((x) => x[0]);
  }

  function renderTabs() {
    $("tabs").innerHTML = custTabs().map((t) => {
      const n = t.id === "docs" ? (docsCount() || "") : t.id === "contact" || t.id === "all" || t.id === "basic" || t.id === "claim" ? "" :
        t.id === "consult" ? ((current.extra && current.extra.consults) || []).length :
        filesForTab(t).length + (t.id === "contract" ? ownContracts().length : 0);
      const refN = t.id === "contract" ? refContracts().length : 0;
      return `<button data-tab="${t.id}">${esc(t.label)}${n !== "" ? `<small>${n}${refN ? ` (+참고 ${refN})` : ""}</small>` : ""}</button>`;
    }).join("");
    $("tabs").querySelectorAll("button").forEach((b) =>
      b.addEventListener("click", async () => {
        if (b.dataset.tab !== "basic" && !(await confirmLeaveEdit())) return; // 수정 중이면 저장·취소·이동 취소 확인
        selectTab(b.dataset.tab);
      }));
  }

  // 수정모드에서 다른 곳으로 이동하려 할 때: 저장하고 이동 / 수정 취소하고 이동 / 이동 취소
  function askLeave() {
    return new Promise((resolve) => {
      const box = document.createElement("div");
      box.className = "modal";
      box.innerHTML = `<div class="modal-box leave-box" role="dialog" aria-modal="true">
          <p class="lead"><b>수정 중인 내용이 있어요</b></p>
          <p class="hint">이동하기 전에 어떻게 할까요?</p>
          <div class="leave-btns">
            <button type="button" class="primary" data-a="save">저장하고 이동</button>
            <button type="button" class="secondary" data-a="discard">수정 취소하고 이동</button>
            <button type="button" class="secondary" data-a="stay">이동 취소</button>
          </div></div>`;
      const done = (a) => { box.remove(); resolve(a); };
      box.addEventListener("click", (ev) => { const b = ev.target.closest("[data-a]"); if (b) done(b.dataset.a); else if (ev.target === box) done("stay"); });
      document.body.appendChild(box);
    });
  }
  // 이동해도 되면 true (수정 중이 아니거나, 저장했거나, 수정을 취소한 경우)
  async function confirmLeaveEdit() {
    if (!editMode || $("customerView").hidden) return true;
    const a = await askLeave();
    if (a === "stay") return false;
    if (a === "discard") { setEditMode(false, true); return true; }
    return saveInfo({ stay: true });
  }

  function kind(name) {
    const ext = name.split(".").pop().toLowerCase();
    if (["xlsx", "xlsm", "xls", "csv"].includes(ext)) return "excel";
    if (ext === "pdf") return "pdf";
    if (["doc", "docx"].includes(ext)) return "word";
    return "other";
  }
  const KIND_LABEL = { excel: "엑셀", pdf: "PDF", word: "워드", other: "파일" };

  // 파일 표시 이름: 계약 PDF("날짜 - 고객 - FC - 상품명 - 보험료원.pdf")는 "상품명 - 보험료원"만
  function fileLabel(f) {
    const m = f.name.match(/^\d{8} - .+? - .+? - (.+) - ([\d,]+원)\.pdf$/i);
    if (m) return esc(productName(m[1]) + " - " + m[2]);
    return `${esc(f.path)}<div class="hint">${new Date(f.lastModifiedDateTime).toLocaleDateString("ko-KR")}</div>`;
  }

  // v466(2026-11-19 사용자 지시): 고객 정보 > [자료] 탭(청구 오른쪽). 고객 폴더 아래 '자료' 폴더의 파일을 썸네일로 보여 주고, 올리기·삭제(드라이브 휴지통)
  const DOCS_DIR = "자료", DOCS_RE = /(^|\/)자료\//, DOCS_MAX = 50 * 1024 * 1024;
  const custTabs = () => (CFG.tabs.some((t) => t.id === "docs") ? CFG.tabs : [...CFG.tabs, { id: "docs", label: "자료", keywords: [] }]);
  const docsCount = () => (current && current.docs ? current.docs.length : ((current && current.files) || []).filter((f) => DOCS_RE.test(f.path)).length);
  let docsUrls = [];
  const docsFree = () => { docsUrls.forEach((u) => { try { URL.revokeObjectURL(u); } catch (e) { /* 무시 */ } }); docsUrls = []; };
  const docsIsImg = (f) => /^image\//.test(f.mime || "") || /\.(jpe?g|png|gif|webp|heic|bmp)$/i.test(f.name);
  const docsExt = (f) => { const m = String(f.name).match(/\.([A-Za-z0-9]{1,5})$/); return m ? m[1].toUpperCase() : "파일"; };
  async function docsDir(make) { // 고객 폴더 아래 '자료' 폴더 ({id}) — 없으면 make일 때만 만듦
    const folder = current.folders[0]; if (!folder) return null;
    if (current.docsDir) return current.docsDir;
    const r = await gdQuery("'" + qv(folder.id) + "' in parents and name='" + qv(DOCS_DIR) + "' and mimeType='" + GD_FOLDER + "' and trashed=false");
    if (r.length) return (current.docsDir = { id: r[0].id });
    if (!make) return null;
    const mk = await gdFetch(GDAPI + "/files?fields=id&supportsAllDrives=true", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ name: DOCS_DIR, mimeType: GD_FOLDER, parents: [folder.id] }) });
    return (current.docsDir = { id: mk.id });
  }
  async function docsList(dir) {
    const out = []; let page = "";
    do {
      const d = await gdFetch(GDAPI + "/files?pageSize=200&supportsAllDrives=true&includeItemsFromAllDrives=true&orderBy=" + encodeURIComponent("modifiedTime desc") + "&fields=" + encodeURIComponent("nextPageToken,files(id,name,mimeType,size,modifiedTime,webViewLink,thumbnailLink)") + "&q=" + encodeURIComponent("'" + qv(dir.id) + "' in parents and mimeType!='" + GD_FOLDER + "' and trashed=false") + (page ? "&pageToken=" + encodeURIComponent(page) : ""));
      (d.files || []).forEach((f) => out.push({ id: f.id, name: f.name, mime: f.mimeType || "", size: Number(f.size || 0), mod: f.modifiedTime || "", webUrl: f.webViewLink || "", thumb: f.thumbnailLink || "" }));
      page = d.nextPageToken || "";
    } while (page);
    return out;
  }
  const docsOnTab = (cust) => current && current.cust === cust && !$("customerView").hidden && !!$("tabs").querySelector('button.on[data-tab="docs"]');
  async function renderDocs(note) {
    const body = $("tabBody"), cust = current.cust, ro = isRO(current.fc);
    { const ub0 = $("docsUpBtn"); if (ub0) ub0.hidden = ro || !GD || !current.folders[0]; }
    if (!GD) { body.innerHTML = '<div class="card"><p class="hint">자료는 구글 드라이브에서만 쓸 수 있습니다.</p></div>'; return; }
    if (!current.folders[0]) { body.innerHTML = '<div class="card"><p class="hint">이 고객의 폴더가 없어 자료를 올릴 수 없습니다.</p></div>'; return; }
    if (!current.docs) body.innerHTML = '<p class="hint pad">자료를 불러오는 중…</p>';
    try {
      const dir = await docsDir(false);
      let list = dir ? await docsList(dir) : [];
      if (!docsOnTab(cust)) return;
      { // v488: 방금 올린 파일이 드라이브 목록에 아직 안 잡혔으면 앞에 붙여 바로 보여 주고, 잠시 뒤 한 번 더 읽음
        const pend = (current.docsPend || []).filter((x) => x.id && !list.some((f) => f.id === x.id));
        current.docsPend = pend;
        if (pend.length) { list = pend.concat(list); if (!current.docsRetry) { current.docsRetry = 1; setTimeout(() => { if (docsOnTab(cust)) renderDocs(note); }, 3000); } } else current.docsRetry = 0;
      }
      { const ub = $("docsUpBtn"); if (ub) ub.hidden = ro; }
      current.docs = list;
      { const sm = $("tabs").querySelector('button[data-tab="docs"]'); if (sm) sm.innerHTML = "자료" + (list.length ? `<small>${list.length}</small>` : ""); }
      docsFree();
      if (current.docsSel) { const ids = new Set(list.map((f) => f.id)); [...current.docsSel].forEach((x) => { if (!ids.has(x)) current.docsSel.delete(x); }); } // v487: 여러 개 골라 삭제
      const selMode = !ro && !!current.docsSel;
      { const ub1 = $("docsUpBtn"); if (ub1) ub1.disabled = selMode; } // v489: 자료를 고르는 중에는 [업로드]를 못 누르게
      const sizeTx = (n) => (n >= 1048576 ? (n / 1048576).toFixed(1) + "MB" : n >= 1024 ? Math.round(n / 1024) + "KB" : n ? n + "B" : "");
      body.innerHTML = `<div class="card docs-card"><div class="docs-head"><span class="docs-cnt">${selMode ? `<b data-docs-seln>${current.docsSel.size}</b>개 선택` : `자료 ${list.length}개${note ? ` · ${esc(note)}` : ""}`}</span>${ro ? "" : selMode ? `<button type="button" class="ev-edit docs-delbtn" data-docs-seldel ${current.docsSel.size ? "" : "disabled"}>삭제</button><button type="button" class="ev-edit" data-docs-selcancel>취소</button>` : `${list.length ? `<button type="button" class="ev-edit" data-docs-selmode>선택</button>` : ""}<input type="file" id="docsFile" multiple hidden>`}</div>
        ${list.length ? `<ul class="docs-grid${selMode ? " selecting" : ""}">${list.map((f, i) => `<li class="docs-item${selMode && current.docsSel.has(f.id) ? " sel" : ""}" data-di="${i}"><button type="button" class="docs-thumb" data-docs-open="${i}" aria-label="${esc(f.name)} 열기">${f.thumb || docsIsImg(f) ? `<img alt="" loading="lazy" referrerpolicy="no-referrer" data-docs-img="${i}" ${f.thumb ? `src="${esc(f.thumb.replace(/=s\d+$/, "=s240"))}"` : ""}>` : ""}<span class="docs-ext"${f.thumb || docsIsImg(f) ? " hidden" : ""}>${esc(docsExt(f))}</span><span class="docs-chk" aria-hidden="true"></span></button>${ro || selMode ? "" : `<button type="button" class="docs-del" data-docs-del="${i}" aria-label="${esc(f.name)} 삭제">×</button>`}<div class="docs-nm" title="${esc(f.name)}">${esc(f.name)}</div><div class="docs-meta">${esc(String(f.mod).slice(2, 10))}${sizeTx(f.size) ? " · " + sizeTx(f.size) : ""}</div></li>`).join("")}</ul>` : `<p class="hint">올린 자료가 없습니다.${ro ? "" : " [업로드]로 사진·파일을 올리면 이 고객 폴더의 ‘자료’ 폴더에 저장됩니다."}</p>`}</div>`;
      body.querySelectorAll("img[data-docs-img]").forEach((img) => { // 썸네일 주소가 안 열리면: 사진은 원본을 받아 작게 줄여 보여 주고, 그 밖은 파일 종류 글자
        const f = list[+img.dataset.docsImg];
        const fail = async () => {
          img.onerror = null;
          const ext = img.parentNode.querySelector(".docs-ext");
          if (docsIsImg(f) && f.size <= 15 * 1024 * 1024) {
            try { const blob = new Blob([await stGet0(f.id)], { type: f.mime || "image/jpeg" }), small = await shrinkImage(blob, 320, 0.8).catch(() => blob), u = URL.createObjectURL(small); docsUrls.push(u); img.src = u; return; } catch (e) { /* 아래로 */ }
          }
          img.remove(); if (ext) ext.hidden = false;
        };
        if (!img.getAttribute("src")) fail(); else img.onerror = fail;
      });
    } catch (e) { if (docsOnTab(cust)) body.innerHTML = '<div class="card"><p class="hint">자료를 불러오지 못했습니다.</p></div>'; handleError(e); }
  }
  async function docsUpload(dir, name, file) { // 5MB까지는 한 번에, 더 크면 이어 올리기 방식
    const mime = file.type || "application/octet-stream";
    if (file.size <= 4.5 * 1024 * 1024) return stPutIn(dir, name, await file.arrayBuffer(), mime);
    const t = storedGToken() || gTokenMem; if (!t) throw new Error("GOOGLE_AUTH");
    const r1 = await fetch(GDUP + "/files?uploadType=resumable&supportsAllDrives=true", { method: "POST", headers: { Authorization: "Bearer " + t, "Content-Type": "application/json; charset=UTF-8", "X-Upload-Content-Type": mime }, body: JSON.stringify({ name, parents: [dir.id] }) });
    const loc = r1.ok ? r1.headers.get("Location") : "";
    if (!loc) throw new Error("Drive " + r1.status + ": 큰 파일 올리기를 시작하지 못했습니다");
    const r2 = await fetch(loc, { method: "PUT", headers: { "Content-Type": mime }, body: file });
    if (!r2.ok) throw new Error("Drive " + r2.status + ": " + (await r2.text()).slice(0, 200));
    return r2.json();
  }
  async function docsAddFiles(files, btn) {
    const cust = current.cust, big = files.filter((f) => f.size > DOCS_MAX), ok = files.filter((f) => f.size <= DOCS_MAX);
    if (!ok.length) { toast("50MB가 넘는 파일은 올릴 수 없습니다"); return; }
    if (btn) { btn.disabled = true; }
    let done = 0; const failed = [], upl = [];
    try {
      const dir = await docsDir(true);
      const have = new Set((await docsList(dir)).map((f) => f.name));
      const d = new Date(), ymd = `${d.getFullYear()}${pad2(d.getMonth() + 1)}${pad2(d.getDate())}`;
      for (const f of ok) {
        if (btn) btn.textContent = `올리는 중 ${done + 1}/${ok.length}`;
        let base = String(f.name || "파일").replace(/[\\/:*?"<>|]/g, "").trim() || "파일";
        if (!/^\d{8}\b/.test(base)) base = ymd + " - " + base; // 파일 이름 앞 날짜 규칙
        const m = base.match(/^(.*?)(\.[A-Za-z0-9]{1,5})?$/); let name = base, k = 2;
        while (have.has(name)) name = `${m[1]} (${k++})${m[2] || ""}`; // 같은 이름이 있으면 덮어쓰지 않고 (2), (3)…
        try { const it = await docsUpload(dir, name, f); if (it && it.id) upl.push({ id: it.id, name: it.name || name, mime: (it.file && it.file.mimeType) || it.mimeType || f.type || "", size: Number(it.size || f.size || 0), mod: it.lastModifiedDateTime || it.modifiedTime || new Date().toISOString(), webUrl: it.webUrl || it.webViewLink || "", thumb: "" }); have.add(name); done++; } catch (e) { if (String(e.message).includes("GOOGLE_AUTH")) throw e; console.error(e); failed.push(f.name); }
      }
    } catch (e) { handleError(e); }
    if (btn) { btn.disabled = false; btn.textContent = "업로드"; }
    if (current.cust === cust) { current.docs = null; current.docsPend = upl.reverse().concat(current.docsPend || []); current.docsRetry = 0; if (docsOnTab(cust)) await renderDocs([done ? `${done}개 올림` : "", failed.length ? `${failed.length}개 실패` : "", big.length ? `${big.length}개는 50MB 초과` : ""].filter(Boolean).join(", ")); }
  }
  function docsView(f) { // 사진은 화면 안에서 크게, PDF는 새 창, 그 밖은 드라이브에서
    if (!docsIsImg(f)) { if (/\.pdf$/i.test(f.name) || f.mime === "application/pdf") openFileDirect(f.id); else if (f.webUrl) window.open(f.webUrl, "_blank", "noopener"); return; }
    const box = document.createElement("div");
    box.className = "modal docs-view"; box.innerHTML = `<p class="hint">사진을 여는 중…</p>`;
    let url = "";
    const close = () => { box.remove(); if (url) URL.revokeObjectURL(url); document.removeEventListener("keydown", onKey); };
    const onKey = (e) => { if (e.key === "Escape") close(); };
    box.addEventListener("click", close); document.addEventListener("keydown", onKey);
    document.body.appendChild(box);
    stGet0(f.id).then((buf) => { if (!box.isConnected) return; url = URL.createObjectURL(new Blob([buf], { type: f.mime || "image/jpeg" })); box.innerHTML = `<img alt="${esc(f.name)}" src="${url}">`; }).catch((e) => { close(); handleError(e); });
  }
  document.addEventListener("click", async (ev) => {
    const t = ev.target && ev.target.closest && ev.target.closest("#docsUpBtn, #tabBody [data-docs-add], #tabBody [data-docs-open], #tabBody [data-docs-del], #tabBody [data-docs-selmode], #tabBody [data-docs-selcancel], #tabBody [data-docs-seldel]");
    if (!t || !current) return;
    if (t.hasAttribute("data-docs-selmode")) { current.docsSel = new Set(); renderDocs(); return; }
    if (t.hasAttribute("data-docs-selcancel")) { current.docsSel = null; renderDocs(); return; }
    if (t.hasAttribute("data-docs-seldel")) {
      const picks = (current.docs || []).filter((x) => current.docsSel && current.docsSel.has(x.id));
      if (!picks.length || isRO(current.fc) || !confirm(`고른 자료 ${picks.length}개를 삭제할까요?\n(구글 드라이브 휴지통으로 옮겨집니다)`)) return;
      t.disabled = true; let ok = 0, bad = 0;
      for (const x of picks) { t.textContent = `삭제 중 ${ok + bad + 1}/${picks.length}`; try { await stDel(x.id); ok++; current.files = current.files.filter((y) => y.id !== x.id); } catch (e) { if (String(e.message).includes("GOOGLE_AUTH")) { handleError(e); break; } console.error(e); bad++; } }
      current.docsSel = null; current.docs = null; await renderDocs([ok ? `${ok}개 삭제` : "", bad ? `${bad}개 실패` : ""].filter(Boolean).join(", "));
      return;
    }
    if (t.hasAttribute("data-docs-open") && current.docsSel) { // 선택 중에는 누르면 고름/풂
      const g = (current.docs || [])[+t.dataset.docsOpen]; if (!g) return;
      if (current.docsSel.has(g.id)) current.docsSel.delete(g.id); else current.docsSel.add(g.id);
      t.closest(".docs-item").classList.toggle("sel", current.docsSel.has(g.id));
      const body = $("tabBody"), n = body.querySelector("[data-docs-seln]"), db = body.querySelector("[data-docs-seldel]");
      if (n) n.textContent = current.docsSel.size; if (db) db.disabled = !current.docsSel.size;
      return;
    }
    if (t.hasAttribute("data-docs-add")) { const inp = $("docsFile"); if (inp) { inp.value = ""; inp.click(); } return; }
    const f = (current.docs || [])[+(t.dataset.docsOpen || t.dataset.docsDel)]; if (!f) return;
    if (t.hasAttribute("data-docs-open")) { docsView(f); return; }
    if (isRO(current.fc) || !confirm(`‘${f.name}’ 자료를 삭제할까요?\n(구글 드라이브 휴지통으로 옮겨집니다)`)) return;
    t.disabled = true;
    try { await stDel(f.id); current.files = current.files.filter((x) => x.id !== f.id); current.docs = null; await renderDocs("1개 삭제"); } catch (e) { t.disabled = false; handleError(e); }
  });
  document.addEventListener("change", (ev) => {
    const inp = ev.target; if (!inp || inp.id !== "docsFile" || !inp.files || !inp.files.length) return;
    docsAddFiles([...inp.files], $("docsUpBtn"));
  });
  // v494(2026-11-19 사용자 지시): 고객 정보 > [청구] 탭. 고객 폴더 > 보험금청구 > 보험금청구서 폴더에 보험사 청구서를 올리고, 파일 이름에서 보험사를 찾아 보험사마다 체크박스로 보여 줌
  const CLM_DIR = "보험금청구", CLM_SUB = "보험금청구서", CLM_RE = /(^|\/)보험금청구\//;
  const CLM_CO = [ // [표시 이름, 파일 이름에서 찾을 말(띄어쓰기 무시)] — 긴 말부터
    ["삼성생명", ["삼성생명"]], ["한화생명", ["한화생명"]], ["교보생명", ["교보생명", "교보"]], ["신한라이프", ["신한라이프", "신한생명", "오렌지라이프"]], ["흥국생명", ["흥국생명"]], ["동양생명", ["동양생명"]], ["미래에셋생명", ["미래에셋생명", "미래에셋"]],
    ["NH농협생명", ["농협생명", "NH생명"]], ["KB라이프", ["KB라이프", "KB생명", "푸르덴셜"]], ["DB생명", ["DB생명"]], ["ABL생명", ["ABL생명", "ABL"]], ["메트라이프", ["메트라이프", "MetLife"]], ["라이나생명", ["라이나생명", "라이나"]], ["AIA생명", ["AIA생명", "AIA"]],
    ["처브라이프", ["처브라이프"]], ["하나생명", ["하나생명"]], ["KDB생명", ["KDB생명"]], ["푸본현대생명", ["푸본현대생명", "푸본현대"]], ["iM라이프", ["iM라이프", "DGB생명"]], ["IBK연금보험", ["IBK연금"]], ["BNP파리바카디프생명", ["카디프"]],
    ["삼성화재", ["삼성화재"]], ["현대해상", ["현대해상"]], ["DB손해보험", ["DB손해보험", "DB손보", "동부화재"]], ["KB손해보험", ["KB손해보험", "KB손보"]], ["메리츠화재", ["메리츠화재", "메리츠"]], ["한화손해보험", ["한화손해보험", "한화손보"]],
    ["롯데손해보험", ["롯데손해보험", "롯데손보"]], ["흥국화재", ["흥국화재"]], ["MG손해보험", ["MG손해보험", "MG손보"]], ["NH농협손해보험", ["농협손해보험", "농협손보", "NH손보"]], ["하나손해보험", ["하나손해보험", "하나손보"]], ["신한EZ손해보험", ["신한EZ손해보험", "신한EZ손보", "신한EZ", "신한이지손해보험", "신한이지손보"]], ["AIG손해보험", ["AIG손해보험", "AIG"]],
    ["캐롯손해보험", ["캐롯"]], ["우체국보험", ["우체국"]]
  ];
  const clmNorm = (t) => String(t || "").replace(/\s+/g, "").toUpperCase();
  function clmCompany(name) { // 파일 이름 → 보험사 (못 찾으면 "")
    const n = clmNorm(name); let best = "", len = 0;
    CLM_CO.forEach(([co, keys]) => keys.forEach((k) => { const kk = clmNorm(k); if (kk.length > len && n.includes(kk)) { best = co; len = kk.length; } }));
    return best;
  }
  const clmOnTab = (cust) => current && current.cust === cust && !$("customerView").hidden && !!$("tabs").querySelector('button.on[data-tab="claim"]');
  async function clmFolder(parentId, name, make) {
    const r = await gdQuery("'" + qv(parentId) + "' in parents and name='" + qv(name) + "' and mimeType='" + GD_FOLDER + "' and trashed=false");
    if (r.length) return { id: r[0].id };
    if (!make) return null;
    const mk = await gdFetch(GDAPI + "/files?fields=id&supportsAllDrives=true", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ name, mimeType: GD_FOLDER, parents: [parentId] }) });
    return { id: mk.id };
  }
  async function clmDir(make) { // 고객 폴더 > 보험금청구 > 보험금청구서 ({id}) — 없으면 make일 때만 만듦
    if (current.clmDir) return current.clmDir;
    const folder = current.folders[0]; if (!folder) return null;
    const top = await clmFolder(folder.id, CLM_DIR, make); if (!top) return null;
    current.clmTop = top;
    const sub = await clmFolder(top.id, CLM_SUB, make); if (!sub) return null;
    return (current.clmDir = sub);
  }
  const HSP_SUB = "병원서류"; // v495(2026-11-19 사용자 지시): 보험금청구서 칸 아래 병원서류 칸 — 고객 폴더 > 보험금청구 > 병원서류
  async function hspDir(make) {
    if (current.hspDir) return current.hspDir;
    if (!current.clmTop && !(await clmDir(make))) return null;
    const d = await clmFolder(current.clmTop.id, HSP_SUB, make); if (!d) return null;
    return (current.hspDir = d);
  }
  const clmPrefix = () => { const d = new Date(); return [`${d.getFullYear()}${pad2(d.getMonth() + 1)}${pad2(d.getDate())}`, current.fc || myFcName() || "", baseName(current.cust && current.cust.name)].filter(Boolean).join(" - "); };
  const HSP_KIND = [[/신분증|주민등록|운전면허|면허증|여권|외국인등록/, "신분증"], [/영수|계산서/, "영수증"], [/세부|내역|산정/, "세부내역"], [/진단/, "진단서"], [/소견/, "소견서"], [/처방/, "처방전"], [/입.?퇴원|입원.?확인|퇴원/, "입퇴원확인서"], [/수술/, "수술확인서"], [/통원|진료.?확인/, "통원확인서"], [/검사|결과/, "검사결과"]];
  const hspKind = (n) => (HSP_KIND.find(([re]) => re.test(n)) || [0, "병원서류"])[1]; // v502: 원래 이름에서 서류 종류를 찾아 짧은 이름으로(못 찾으면 '병원서류')
  const clmYmd = () => { const d = new Date(); return `${d.getFullYear()}${pad2(d.getMonth() + 1)}${pad2(d.getDate())}`; };
  const clmStdName = (co, ext) => `${clmYmd()} - ${co || "청구서"}${ext || ""}`; // v501(2026-11-19 사용자 지시): 청구서 이름은 짧게 'YYYYMMDD - 보험사.pdf' (이미 고객 폴더 안이라 FC·고객 이름은 뺌)
  // v497(2026-11-19 사용자 지시): 보험금청구서 제작 — 보험사 청구서 PDF마다 병원서류를 뒤에 붙여 한 PDF로, 보험금청구 › 보험금청구 기록 폴더에 저장
  // 스킬 insurance-claim-docs의 규칙을 따름: 쪽 순서 = 청구서 → 진료비 계산서·영수증 → 세부내역 등 나머지, A4 세로, 사진은 흑백(회색조)으로. 손글씨 칸 채우기·잡티 지우기는 PC 스킬에서만
  const REC_SUB = "보험금청구 기록";
  async function recDir(make) {
    if (current.recDir) return current.recDir;
    if (!current.clmTop && !(await clmDir(make))) return null;
    const d = await clmFolder(current.clmTop.id, REC_SUB, make); if (!d) return null;
    return (current.recDir = d);
  }
  const clmTargets = () => { const list = current.clmList || [], sel = current.clmSel || new Set(); const ord = (f) => clmCompany(f.name) || "￿"; return (sel.size ? list.filter((f) => sel.has(clmCompany(f.name))) : list).slice().sort((a, b) => ord(a).localeCompare(ord(b), "ko")); }; // v504: 화면에 보이는 보험사 순서대로
  const clmBuildLabel = () => { const n = clmTargets().length, sel = current.clmSel || new Set(); return `보험금청구서 제작 시작 (${sel.size ? "체크한 " : ""}청구서 ${n}개 + 병원서류)`; };
  const hspOrder = (f) => (/영수|계산서/.test(f.name) ? 0 : /세부|내역|산정/.test(f.name) ? 1 : 2);
  const isPdfFile = (f) => /\.pdf$/i.test(f.name) || f.mime === "application/pdf";
  async function imgToA4(doc, buf, mime) { // 사진 한 장 → A4 세로 한 쪽
    const blob = new Blob([buf], { type: mime || "image/jpeg" });
    let bmp; try { bmp = await createImageBitmap(blob, { imageOrientation: "from-image" }); } catch (e) { bmp = await createImageBitmap(blob); }
    return srcToA4(doc, bmp, bmp.width, bmp.height);
  }
  async function tifPages(buf) { // 팩스 TIF(여러 쪽) → 쪽마다 캔버스. 가로·세로 해상도가 다르면(예: 204×98dpi) 세로 비율을 바로잡음
    await rptLoad("https://cdn.jsdelivr.net/npm/pako@1.0.11/dist/pako.min.js");
    await rptLoad("https://cdn.jsdelivr.net/npm/utif@3.1.0/UTIF.js");
    const ifds = UTIF.decode(buf).filter((d) => d.t256 && d.t257), out = [];
    for (const d of ifds) {
      UTIF.decodeImage(buf, d);
      const w = d.width, h = d.height, rgba = UTIF.toRGBA8(d);
      const c0 = document.createElement("canvas"); c0.width = w; c0.height = h;
      c0.getContext("2d").putImageData(new ImageData(new Uint8ClampedArray(rgba.buffer, rgba.byteOffset, w * h * 4), w, h), 0, 0);
      const rv = (t) => (Array.isArray(t) ? (t.length > 1 && t[1] ? t[0] / t[1] : t[0]) : t) || 0;
      const xr = rv(d.t282), yr = rv(d.t283), ratio = xr && yr ? xr / yr : 1;
      if (Math.abs(ratio - 1) > 0.05) { const c1 = document.createElement("canvas"); c1.width = w; c1.height = Math.round(h * ratio); const g1 = c1.getContext("2d"); g1.imageSmoothingQuality = "high"; g1.drawImage(c0, 0, 0, c1.width, c1.height); out.push(c1); } else out.push(c0);
    }
    return out;
  }
  // v503(2026-11-19 사용자 지시): 병원서류를 팩스 보내기 좋게 흑백(검정·흰색 두 가지)으로 정리 — 스킬 insurance-claim-docs의 fax_clean 규칙을 따름
  //   ① 주변 밝기(41×41 평균)보다 어두운 곳만 글자로 봐서 그림자·누런 바탕·회색 음영을 흰색으로 ② 아주 진한 점은 늘 글자 ③ 외톨이 잡티(3픽셀 이하 덩어리)만 지움 — 마침표·쉼표(200dpi에서 보통 9픽셀 넘음)는 남김
  function faxClean(a, W, H) {
    const n = W * H, gray = new Uint8Array(n);
    for (let i = 0, j = 0; i < n; i++, j += 4) gray[i] = (a[j] * 299 + a[j + 1] * 587 + a[j + 2] * 114) / 1000;
    const S = new Float64Array((W + 1) * (H + 1)); // 적분 영상(주변 평균을 빠르게)
    for (let y = 0; y < H; y++) { let row = 0; for (let x = 0; x < W; x++) { row += gray[y * W + x]; S[(y + 1) * (W + 1) + x + 1] = S[y * (W + 1) + x + 1] + row; } }
    const R = Math.max(8, Math.round(20 * W / 1654)), ink = new Uint8Array(n);
    for (let y = 0; y < H; y++) {
      const y0 = Math.max(0, y - R), y1 = Math.min(H, y + R + 1);
      for (let x = 0; x < W; x++) {
        const x0 = Math.max(0, x - R), x1 = Math.min(W, x + R + 1), cnt = (y1 - y0) * (x1 - x0);
        const mean = (S[y1 * (W + 1) + x1] - S[y0 * (W + 1) + x1] - S[y1 * (W + 1) + x0] + S[y0 * (W + 1) + x0]) / cnt;
        const v = gray[y * W + x];
        ink[y * W + x] = v < 70 || (v < 215 && v < mean * 0.86 - 4) ? 1 : 0;
      }
    }
    const seen = new Uint8Array(n), st = new Int32Array(n); // 외톨이 잡티 지우기(8방향으로 이어진 덩어리 크기)
    for (let i = 0; i < n; i++) {
      if (!ink[i] || seen[i]) continue;
      let sp = 0, size = 0; st[sp++] = i; seen[i] = 1; const comp = [];
      while (sp) { const c = st[--sp]; size++; if (size <= 4) comp.push(c); const cx = c % W, cy = (c - cx) / W;
        for (let dy = -1; dy <= 1; dy++) for (let dx = -1; dx <= 1; dx++) { const nx = cx + dx, ny = cy + dy; if (nx < 0 || ny < 0 || nx >= W || ny >= H) continue; const q = ny * W + nx; if (ink[q] && !seen[q]) { seen[q] = 1; st[sp++] = q; } } }
      if (size <= 3) comp.forEach((c) => { ink[c] = 0; });
    }
    return ink;
  }
  async function bitPage(doc, ink, W, H) { // v505(2026-11-19 사용자 지시 "용량 줄이기"): 흑백 쪽을 1비트(점 하나 = 1비트)로 압축해 넣음 — PNG(컬러용)보다 훨씬 작음
    await rptLoad("https://cdn.jsdelivr.net/npm/pako@1.0.11/dist/pako.min.js");
    const rb = (W + 7) >> 3, bits = new Uint8Array(rb * H).fill(255); // 1 = 흰색, 0 = 검정
    for (let y = 0; y < H; y++) { const o = y * W, r = y * rb; for (let x = 0; x < W; x++) if (ink[o + x]) bits[r + (x >> 3)] &= ~(0x80 >> (x & 7)); }
    const ref = doc.context.register(doc.context.stream(pako.deflate(bits), { Type: "XObject", Subtype: "Image", Width: W, Height: H, ColorSpace: "DeviceGray", BitsPerComponent: 1, Filter: "FlateDecode" }));
    const pg = doc.addPage([595.28, 841.89]), nm = pg.node.newXObject("Im", ref);
    pg.pushOperators(PDFLib.pushGraphicsState(), PDFLib.concatTransformationMatrix(595.28, 0, 0, 841.89, 0, 0), PDFLib.drawObject(nm), PDFLib.popGraphicsState());
  }
  async function pdfToA4(doc, buf) { // 스캔 PDF 병원서류 → 쪽마다 200dpi 그림으로 그려 팩스용 정리
    await rptLoad("https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.min.js");
    if (window.pdfjsLib && !pdfjsLib.GlobalWorkerOptions.workerSrc) pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js";
    const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(buf.slice(0)) }).promise;
    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i), vp = page.getViewport({ scale: clmDpi / 72 });
      const c = document.createElement("canvas"); c.width = Math.round(vp.width); c.height = Math.round(vp.height);
      const g = c.getContext("2d"); g.fillStyle = "#fff"; g.fillRect(0, 0, c.width, c.height);
      await page.render({ canvasContext: g, viewport: vp }).promise;
      await srcToA4(doc, c, c.width, c.height); c.width = c.height = 0; page.cleanup();
    }
    pdf.destroy();
  }
  let clmDpi = 300; // v506: 기본 300dpi(사용자 지시), 결과가 8MB를 넘으면 200 → 150dpi로 낮춰 다시 만듦
  async function srcToA4(doc, bmp, bw, bh) {
    if (clmGray && clmIdName) { const r = bw / bh; return clmStylePage(doc, bmp, bw, bh, (r > 0.66 && r < 0.76) || (r > 1.32 && r < 1.5)); } // A4 모양이면 쪽 그대로, 아니면 사진 // v552: 신분증은 원래 모양 그대로 // 그림 하나 → A4 세로 한 쪽(흑백). 가로로 긴 서류는 누운 채로 돌려 꽉 차게
    const W = Math.round(8.268 * clmDpi), H = Math.round(11.693 * clmDpi), rot = bw > bh, iw = rot ? bh : bw, ih = rot ? bw : bh;
    const k = Math.min(W / iw, H / ih), dw = iw * k, dh = ih * k;
    const cv = document.createElement("canvas"); cv.width = W; cv.height = H;
    const g = cv.getContext("2d"); g.fillStyle = "#fff"; g.fillRect(0, 0, W, H);
    g.save(); g.translate(W / 2, H / 2); if (rot) g.rotate(-Math.PI / 2);
    g.drawImage(bmp, -(rot ? dh : dw) / 2, -(rot ? dw : dh) / 2, rot ? dh : dw, rot ? dw : dh); g.restore();
    if (clmGray) { await grayPage(doc, cv); cv.width = cv.height = 0; return; } // v551: 회색조 JPEG
    const im = g.getImageData(0, 0, W, H);
    const ink = faxClean(im.data, W, H); cv.width = cv.height = 0; // 큰 캔버스는 바로 비움(폰 메모리)
    await bitPage(doc, ink, W, H);
  }
  async function clmAppend(doc, f, buf, clean) {
    clmIdName = ID_RE.test(String(f.name || "")); try { return await clmAppend0(doc, f, buf, clean); } finally { clmIdName = false; }
  }
  async function clmAppend0(doc, f, buf, clean) { // 파일 하나를 쪽으로 붙임. clean = 병원서류(팩스용 흑백 정리), 아니면 PDF는 쪽 그대로
    if (isPdfFile(f)) { if (clean) return pdfToA4(doc, buf); const src = await PDFLib.PDFDocument.load(buf, { ignoreEncryption: true }); (await doc.copyPages(src, src.getPageIndices())).forEach((pg) => doc.addPage(pg)); return; }
    if (/\.tiff?$/i.test(f.name) || f.mime === "image/tiff") { for (const c of await tifPages(buf)) await srcToA4(doc, c, c.width, c.height); return; } // 팩스 TIF
    if (docsIsImg(f) && !/\.heic$/i.test(f.name)) return imgToA4(doc, buf, f.mime);
    throw new Error("SKIP");
  }
  // v498(2026-11-19 사용자 지시): 청구서에 사고일시·사고경위·작성일(오늘)을 손글씨체로 적음
  // 좌표 = 스킬 insurance-claim-docs의 방식(배율 1.6 그림의 픽셀, 글자 왼쪽 아래 기준선). xs = 한 칸에 한 글자. 양식 판이 다르면 다시 잡아야 함
  const CLM_FORMS = {
    "메리츠화재": [
      { page: 1, f: "accY", xs: [148, 185, 221, 258], y: 666, size: 16 }, { page: 1, f: "accM", xs: [320, 355], y: 666, size: 16 }, { page: 1, f: "accD", xs: [416, 450], y: 666, size: 16 },
      { page: 1, f: "accH", xs: [512, 546], y: 666, size: 16 }, { page: 1, f: "accMin", xs: [608, 642], y: 666, size: 16 },
      { page: 1, f: "story", x: 124, y: 714, size: 11.5 },
      { page: 1, f: "todY", x: 236, y: 1201, size: 13 }, { page: 1, f: "todM", x: 308, y: 1201, size: 13 }, { page: 1, f: "todD", x: 355, y: 1201, size: 13 },
      { page: 3, f: "todY", x: 116, y: 1160, size: 13 }, { page: 3, f: "todM", x: 228, y: 1160, size: 13 }, { page: 3, f: "todD", x: 320, y: 1160, size: 13 },
      { page: 1, f: "chkIn", x: 142, y: 627, size: 13 }, { page: 1, f: "chkOut", x: 217, y: 627, size: 13 }, // v499: 세부유형 입원·통원 체크
      { page: 1, f: "hosp", x: 584, y: 752, size: 9.5 }, { page: 1, f: "dept", x: 770, y: 752, size: 11 } // 최초 내원 병원 / 진료과
    ],
    "신한EZ손해보험": [ // v502(2026-11-19): 김정인 고객 건 양식(1쪽 청구서, 4쪽 동의서 3/3)으로 잡음. 병원·과 글자는 양식에 인쇄돼 있어 끝의 '병원'·'과'는 빼고 적음
      { page: 1, f: "accY", x: 318, y: 514, size: 13 }, { page: 1, f: "accM", x: 428, y: 514, size: 13 }, { page: 1, f: "accD", x: 520, y: 514, size: 13 },
      { page: 1, f: "accH", x: 615, y: 514, size: 13 }, { page: 1, f: "accMin", x: 712, y: 514, size: 13 },
      { page: 1, f: "story", x: 252, y: 604, size: 12 },
      { page: 1, f: "hosp", x: 520, y: 662, size: 11, strip: /\s*병원$/ }, { page: 1, f: "dept", x: 746, y: 662, size: 11, strip: /과$/ },
      { page: 1, f: "chkIn", x: 746, y: 784, size: 12 }, { page: 1, f: "chkOut", x: 815, y: 784, size: 12 },
      { page: 1, f: "todY", x: 165, y: 1211, size: 13 }, { page: 1, f: "todM", x: 262, y: 1211, size: 13 }, { page: 1, f: "todD", x: 335, y: 1211, size: 13 },
      { page: 4, f: "todYY", x: 152, y: 1190, size: 13 }, { page: 4, f: "todM", x: 212, y: 1190, size: 13 }, { page: 4, f: "todD", x: 272, y: 1190, size: 13 } // 동의일: 20[26]년
    ],
    "KB손해보험": [ // 2026-11-19 김정인 고객건 양식(인보험용, 1쪽)으로 잡음. 사고일시에 시·분 칸 없음
      { page: 1, f: "accY", xs: [188, 217, 246, 274], y: 664, size: 16 }, { page: 1, f: "accM", xs: [324, 348], y: 664, size: 16 }, { page: 1, f: "accD", xs: [400, 424], y: 664, size: 16 },
      { page: 1, f: "story", x: 178, y: 714, size: 11 },
      { page: 1, f: "todY", xs: [188, 217, 246, 274], y: 1226, size: 16 }, { page: 1, f: "todM", xs: [324, 348], y: 1226, size: 16 }, { page: 1, f: "todD", xs: [400, 424], y: 1226, size: 16 },
      { page: 1, f: "hosp", x: 552, y: 694, size: 10 }, { page: 1, f: "dept", x: 690, y: 694, size: 11 } // v499: 병원명 / 진료과
    ]
  };
  let clmFontP = null;
  function clmFont() { // 손글씨체 '그여자체' = 앱 설정 폴더의 그여자체.ttf(드라이브). 없으면 기본 글꼴
    return clmFontP || (clmFontP = (async () => {
      try {
        const dir = await resolveAppDir(); if (!dir || !dir.id) return "sans-serif";
        const r = await gdQuery("'" + qv(dir.id) + "' in parents and name='그여자체.ttf' and trashed=false"); if (!r.length) return "sans-serif";
        const ff = new FontFace("ClmHand", await stGet0(r[0].id)); await ff.load(); document.fonts.add(ff); return "ClmHand";
      } catch (e) { console.warn("손글씨체를 불러오지 못함", e); return "sans-serif"; }
    })());
  }
  async function clmWrite(doc, page, text, xPx, yPx, size, fam) { // 글자 하나(또는 한 줄)를 그림으로 그려 그 자리에 붙임(한글 글꼴을 PDF에 넣지 않고도 됨)
    const K = 4, sz = size * (0.96 + Math.random() * 0.09), cv = document.createElement("canvas"), g = cv.getContext("2d");
    g.font = `${sz * K}px ${fam}`; const m = g.measureText(text), asc = Math.ceil(m.actualBoundingBoxAscent || sz * K * 0.8) + 4, des = Math.ceil(m.actualBoundingBoxDescent || sz * K * 0.2) + 4;
    cv.width = Math.max(4, Math.ceil(m.width) + 8); cv.height = asc + des;
    g.font = `${sz * K}px ${fam}`; g.fillStyle = "rgb(28,28,40)"; g.textBaseline = "alphabetic"; g.fillText(text, 2, asc);
    const png = await doc.embedPng(await new Promise((r) => cv.toBlob(async (b) => r(new Uint8Array(await b.arrayBuffer())), "image/png")));
    const H = page.getHeight(), x = xPx / 1.6 + (Math.random() - 0.5) * 1.2, base = H - yPx / 1.6 + (Math.random() - 0.5) * 1.2;
    page.drawImage(png, { x: x - 2 / K, y: base - des / K, width: cv.width / K, height: cv.height / K, rotate: PDFLib.degrees((Math.random() - 0.5) * 2.4) }); // 사람이 쓴 것처럼 크기·기울기를 살짝 흔듦
  }
  async function clmFill(src, co, v) { // 알고 있는 양식이면 칸에 적음. 적은 칸 수를 돌려줌
    const form = CLM_FORMS[co]; if (!form) return 0;
    const fam = await clmFont(), pages = src.getPages(); let n = 0;
    for (const e of form) {
      const t = String(v[e.f] || "").replace(e.strip || /$^/, "").trim(), pg = pages[e.page - 1]; if (!t || !pg) continue;
      if (e.xs) { for (let i = 0; i < e.xs.length && i < t.length; i++) await clmWrite(src, pg, t[i], e.xs[i], e.y, e.size, fam); } else await clmWrite(src, pg, t, e.x, e.y, e.size, fam);
      n++;
    }
    return n;
  }
  // v499(2026-11-19 사용자 지시 "Gemini로 읽기"): 병원서류(영수증·세부내역·진단서 등)를 Gemini에 보내 사고일·시각·경위·병원·진료과·입원/통원을 읽어 입력창에 미리 채움
  //   [보험금청구서 제작 시작]을 누를 때만 보냄. 키는 말만해요 설정 파일(aiGeminiConfig). 주민번호·계좌는 묻지도 적지도 않음
  // ---- v565(2026-11-24 진단 뒤 고침): Gemini 호출 하나로 — 성공한 모델을 기억해 먼저 쓰고, 없는 모델(404)은 이 세션에서 건너뜀, 요청마다 20초 제한, 전체 50초 안, flash 모델은 생각 단계 끔 ----
  //   onStep(상태, 모델, 차례, 경과ms): 'start' | 'retry' — 화면 진행 표시용. 돌려주는 값은 fetch 응답(없으면 null)
  const GEM_LAST = "shlife.gemModel", GEM_BAD = new Set();
  async function gemCall(cfg, req, onStep) {
    let last = ""; try { last = localStorage.getItem(GEM_LAST) || ""; } catch (e) { /* 무시 */ }
    const order = [...new Set([last, "gemini-3.5-flash", cfg.model, "gemini-flash-latest", "gemini-3.8-flash", "gemini-flash-lite-latest", "gemini-3.1-flash-lite"].filter((m) => m && !GEM_BAD.has(m)))];
    const t0 = Date.now(); let res = null, n = 0;
    for (const mdl of order) {
      if (n && Date.now() - t0 > 45000) break; // 전체 50초 안쪽에서 멈춤
      if (onStep) try { onStep(n ? "retry" : "start", mdl, n, Date.now() - t0); } catch (e) { /* 무시 */ }
      n++;
      for (const think of [true, false]) {
        const body = JSON.stringify(think && /flash/.test(mdl) ? { ...req, generationConfig: { ...(req.generationConfig || {}), thinkingConfig: { thinkingBudget: 0 } } } : req);
        const ac = new AbortController(), tm = setTimeout(() => ac.abort(), 20000);
        try { res = await fetch("https://generativelanguage.googleapis.com/v1beta/models/" + encodeURIComponent(mdl) + ":generateContent", { method: "POST", headers: { "Content-Type": "application/json", "x-goog-api-key": cfg.key }, body, signal: ac.signal }); }
        catch (e) { res = null; console.warn("Gemini", mdl, e && e.name === "AbortError" ? "20초 넘음" : e && e.message, "→ 다음 모델"); }
        finally { clearTimeout(tm); }
        if (res && res.status === 400 && think && /flash/.test(mdl)) { const tx = await res.clone().text().catch(() => ""); if (/thinking/i.test(tx)) continue; } // 생각 끄기를 모르는 모델이면 끄지 않고 한 번 더
        break;
      }
      if (!res) continue;
      if (res.status === 404) { GEM_BAD.add(mdl); continue; }
      if ([429, 500, 503].includes(res.status)) { console.warn("Gemini", mdl, res.status, "→ 다음 모델"); continue; }
      if (res.ok) { try { localStorage.setItem(GEM_LAST, mdl); } catch (e) { /* 무시 */ } }
      return res;
    }
    return res;
  }
  const b64 = (blob) => new Promise((r, j) => { const fr = new FileReader(); fr.onload = () => r(String(fr.result).split(",")[1] || ""); fr.onerror = j; fr.readAsDataURL(blob); });
  async function cvJpeg(src, w, h, max) { const k = Math.min(1, max / Math.max(w, h)), c = document.createElement("canvas"); c.width = Math.round(w * k); c.height = Math.round(h * k); const g = c.getContext("2d"); g.fillStyle = "#fff"; g.fillRect(0, 0, c.width, c.height); g.drawImage(src, 0, 0, c.width, c.height); return new Promise((r) => c.toBlob(r, "image/jpeg", 0.85)); }
  async function hspParts(hsp, bufs) { // Gemini에 보낼 조각: PDF는 그대로, 사진·TIF는 긴 변 1800px JPEG (최대 10장)
    const parts = []; let imgs = 0;
    for (let i = 0; i < hsp.length && imgs < 10; i++) {
      const f = hsp[i], buf = bufs[i];
      try {
        if (isPdfFile(f)) { if (buf.byteLength <= 15 * 1024 * 1024) { parts.push({ inline_data: { mime_type: "application/pdf", data: await b64(new Blob([buf])) } }); imgs++; } continue; }
        if (/\.tiff?$/i.test(f.name) || f.mime === "image/tiff") { for (const c of await tifPages(buf)) { if (imgs >= 10) break; parts.push({ inline_data: { mime_type: "image/jpeg", data: await b64(await cvJpeg(c, c.width, c.height, 1800)) } }); imgs++; } continue; }
        if (docsIsImg(f) && !/\.heic$/i.test(f.name)) { const bm = await createImageBitmap(new Blob([buf], { type: f.mime || "image/jpeg" })); parts.push({ inline_data: { mime_type: "image/jpeg", data: await b64(await cvJpeg(bm, bm.width, bm.height, 1800)) } }); imgs++; }
      } catch (e) { console.warn("Gemini에 못 보낸 병원서류", f.name, e); }
    }
    return parts;
  }
  async function clmReadHosp(hsp, bufs) { // → { accDate, accTime, story, hosp, dept, kind } / { noKey: true } / null(실패)
    const cfg = await aiGeminiConfig(); if (!cfg) return { noKey: true };
    const parts = await hspParts(hsp, bufs); if (!parts.length) return null;
    const prompt = "너는 보험금 청구서를 대신 적는 보험 설계사의 도우미다. 첨부한 병원 서류(진료비 영수증·계산서, 세부산정내역, 진단서, 처방전 등)를 읽고 아래 JSON 하나로만 답하라(설명 금지).\n" +
      '{"accDate":"YYYY-MM-DD","accTime":"HH:MM 또는 빈 문자열","story":"사고(내원) 경위 한 줄, 25자 이내","hosp":"병원 이름","dept":"진료과","kind":"통원|입원|빈 문자열"}\n' +
      "accDate = 이번 진료의 시작일(진료기간·입원일·내원일 중 가장 이른 날). accTime = 서류에 내원·접수 시각이 있을 때만. story = 서류의 처치·검사·시술·진료과를 바탕으로 '무릎 통증으로 정형외과 통원 치료'처럼 짧게. hosp = 영수증 머리·도장의 병원 이름 그대로. kind = 영수증의 외래/입원 구분.\n" +
      "서류에서 확인되지 않는 값은 지어내지 말고 빈 문자열로 둬라. 주민등록번호·계좌번호·환자 이름은 답에 넣지 마라.";
    try {
      const gemReq = { contents: [{ role: "user", parts: [{ text: prompt }, ...parts] }], generationConfig: { responseMimeType: "application/json", temperature: 0 } };
      let res = await gemCall(cfg, gemReq, null); // v565: 모델 기억·없는 모델 건너뛰기·20초 제한·생각 끄기
      if (!res || !res.ok) { console.warn("Gemini 병원서류 읽기 실패", res && res.status, (await res.text().catch(() => "")).slice(0, 200)); return null; }
      const d = await res.json(), c0 = (d.candidates || [])[0] || {}, tx = ((c0.content || {}).parts || []).filter((x) => !x.thought).map((x) => x.text || "").join("");
      if (!tx.trim()) { console.warn("Gemini 병원서류 읽기: 빈 답", cfg.model, c0.finishReason || "", JSON.stringify(d).slice(0, 300)); return null; }
      const j = JSON.parse(tx.replace(/^```json\s*|\s*```$/g, ""));
      const ok = (re, v) => (re.test(String(v || "")) ? String(v) : "");
      return { accDate: ok(/^\d{4}-\d{2}-\d{2}$/, j.accDate), accTime: ok(/^\d{2}:\d{2}$/, j.accTime), story: String(j.story || "").replace(/\d{6}-?\d{7}/g, "").slice(0, 40).trim(), hosp: String(j.hosp || "").slice(0, 30).trim(), dept: String(j.dept || "").slice(0, 15).trim(), kind: /입원/.test(j.kind) ? "입원" : /통원|외래/.test(j.kind) ? "통원" : "" };
    } catch (e) { console.warn("Gemini 병원서류 읽기 실패", e); return null; }
  }
  function clmAsk(targets, hspN, pre) { // 만들기 전에 사고일시·사고경위·작성일을 받음
    return new Promise((done) => {
      const old = document.getElementById("clmPop"); if (old) old.remove();
      const z = (x) => String(x).padStart(2, "0"), d = new Date(), today = `${d.getFullYear()}-${z(d.getMonth() + 1)}-${z(d.getDate())}`, rd = pre && !pre.noKey ? pre : null, last = rd ? { ...(current.clmVals || {}), ...Object.fromEntries(Object.entries(rd).filter(([, v]) => v)) } : current.clmVals || {};
      const aiNote = !pre ? "병원서류를 읽지 못해 빈칸으로 둡니다. 직접 넣어 주세요." : pre.noKey ? "Gemini 키가 없어 병원서류를 읽지 못했습니다(말만해요 설정). 직접 넣어 주세요." : "Gemini가 병원서류를 읽어 채운 값입니다. 맞는지 확인하고 고쳐 주세요.";
      const known = targets.map((t) => clmCompany(t.name)).filter((c) => CLM_FORMS[c]), unknown = targets.filter((t) => !CLM_FORMS[clmCompany(t.name)]).map((t) => clmCompany(t.name) || "보험사 모름");
      const box = document.createElement("div"); box.id = "clmPop"; box.className = "modal";
      box.innerHTML = `<div class="modal-box cd-conv-box" role="dialog" aria-modal="true" aria-label="보험금청구서 만들기"><h3 class="sec-title">보험금청구서 만들기</h3>
        <p class="lead">청구서 ${targets.length}개 · 병원서류 ${hspN}개를 붙입니다(영수증 → 세부내역 → 나머지).</p>
        <div class="clm-pop-row"><label class="fld"><span>사고일</span><input type="text" data-dp="date" data-cp="accDate" value="${esc(last.accDate || "")}"></label><label class="fld"><span>시각(있으면)</span><input type="time" data-cp="accTime" value="${esc(last.accTime || "")}"></label></div>
        <p class="hint clm-ai">${esc(aiNote)}</p>
        <label class="fld"><span>사고경위</span><input type="text" data-cp="story" maxlength="40" placeholder="예: 무릎 통증으로 정형외과 통원 치료" value="${esc(last.story || "")}"></label>
        <div class="clm-pop-row"><label class="fld"><span>병원</span><input type="text" data-cp="hosp" maxlength="30" value="${esc(last.hosp || "")}"></label><label class="fld clm-dept"><span>진료과</span><input type="text" data-cp="dept" maxlength="15" value="${esc(last.dept || "")}"></label><label class="fld clm-kind"><span>구분</span><select data-cp="kind"><option value=""${last.kind ? "" : " selected"}>-</option><option${last.kind === "통원" ? " selected" : ""}>통원</option><option${last.kind === "입원" ? " selected" : ""}>입원</option></select></label></div>
        <label class="fld"><span>작성일</span><input type="text" data-dp="date" data-cp="today" value="${esc(today)}"></label>
        <p class="hint">${known.length ? `손글씨체로 적는 양식: ${esc([...new Set(known)].join(", "))}.` : ""}${unknown.length ? ` 칸 위치를 모르는 양식(${esc([...new Set(unknown)].join(", "))})은 글자 없이 합칩니다.` : ""}</p>
        <div class="file-actions"><button type="button" class="secondary" data-cp-cancel>취소</button><button type="button" class="primary" data-cp-ok>만들기</button></div></div>`;
      document.body.appendChild(box);
      box.addEventListener("click", (e) => {
        if (e.target === box || e.target.closest("[data-cp-cancel]")) { box.remove(); done(null); return; }
        if (!e.target.closest("[data-cp-ok]")) return;
        const get = (k) => String(box.querySelector(`[data-cp="${k}"]`).value || "").trim();
        const v = { accDate: get("accDate"), accTime: get("accTime"), story: get("story"), hosp: get("hosp"), dept: get("dept"), kind: get("kind"), today: get("today") || today };
        if (!v.accDate && !confirm("사고일이 비어 있습니다. 그대로 만들까요?")) return;
        current.clmVals = { accDate: v.accDate, accTime: v.accTime, story: v.story, hosp: v.hosp, dept: v.dept, kind: v.kind };
        const [ay, am, ad] = v.accDate ? v.accDate.split("-") : ["", "", ""], [th, tm] = v.accTime ? v.accTime.split(":") : ["", ""], [ty, tmo, td] = v.today.split("-");
        box.remove(); done({ accY: ay, accM: am, accD: ad, accH: th, accMin: tm, story: v.story, hosp: v.hosp, dept: v.dept, chkIn: v.kind === "입원" ? "V" : "", chkOut: v.kind === "통원" ? "V" : "", todY: ty, todYY: String(ty || "").slice(2), todM: tmo, todD: td });
      });
    });
  }
  async function clmBuild() {
    const bset = (t) => { current.clmBusy = t; const b = $("tabBody") && $("tabBody").querySelector("[data-clm-build]"); if (b) { b.textContent = t || clmBuildLabel(); b.disabled = !!t; } }; // 진행 중에 화면을 다시 그려도 버튼 글자·잠김이 이어지게
    const cust = current.cust, targets = clmTargets(), hsp = (current.hspList || []).slice().sort((a, b) => hspOrder(a) - hspOrder(b) || a.name.localeCompare(b.name, "ko"));
    if (!targets.length || !hsp.length) return toast("보험사 청구서와 병원서류가 모두 있어야 만들 수 있습니다");
    if (current.clmBusy) return;
    const cos = targets.map((t) => clmCompany(t.name) || "보험사 모름");
    if (!confirm(`보험금청구서를 한 파일로 만들까요?\n\n· 순서: ${cos.join(" → ")} 청구서 → 병원서류 ${hsp.length}개(영수증 → 세부내역 → 나머지)\n· 병원서류는 팩스 보내기 좋게 흑백으로 정리합니다. 청구서 칸은 채우지 않습니다.\n· ‘보험금청구 › 보험금청구 기록’ 폴더에 저장합니다.`)) return;
    bset("준비 중…");
    const made = [], skipped = new Set(), failed = [];
    try {
      bset("병원서류 받는 중…");
      const hbuf = [];
      for (const f of hsp) hbuf.push(await stGet0(f.id));
      // v503(2026-11-19 사용자 지시): 병원서류를 읽어 청구서 칸을 채우지 않음(Gemini·손글씨 함수는 남겨 둠). 병원서류만 팩스용으로 정리해 각 청구서 뒤에 붙임
      bset("PDF 도구 불러오는 중…");
      await rptLoad("https://cdn.jsdelivr.net/npm/pdf-lib@1.17.1/dist/pdf-lib.min.js");
      // v506(2026-11-19 사용자 지시 "파일이 커질 때 대비"): 200dpi로 만들어 8MB가 넘으면 150dpi, 그래도 넘으면 100dpi로 다시 만듦(팩스로 보내기 좋은 크기)
      const LIMIT = 8 * 1024 * 1024, tbuf = [];
      for (let k = 0; k < targets.length; k++) { bset(`청구서 받는 중 ${k + 1}/${targets.length}`); try { tbuf.push(await stGet0(targets[k].id)); } catch (e) { if (String(e.message).includes("GOOGLE_AUTH")) throw e; tbuf.push(null); } }
      let bytes = null, used = [], dpiUsed = 300;
      for (const dpi of [300, 200, 150]) {
        clmDpi = dpi; dpiUsed = dpi; skipped.clear(); failed.length = 0; used = [];
        const tag = dpi === 300 ? "" : ` (${dpi}dpi로 다시)`;
        const hdoc = await PDFLib.PDFDocument.create();
        for (let i = 0; i < hsp.length; i++) { bset(`병원서류 정리 중 ${i + 1}/${hsp.length}${tag}`); await new Promise((r) => setTimeout(r, 30)); try { await clmAppend(hdoc, hsp[i], hbuf[i], true); } catch (e) { console.warn(hsp[i].name, e); skipped.add(hsp[i].name); } }
        // v504: 인쇄하기 좋게 한 파일로 — 보험사 청구서들(위 칸 순서) → 정리한 병원서류 한 벌. v505: 청구서도 팩스용 흑백으로
        const doc = await PDFLib.PDFDocument.create();
        for (let k = 0; k < targets.length; k++) {
          const t = targets[k]; bset(`청구서 붙이는 중 ${k + 1}/${targets.length}${tag}`);
          try { if (!tbuf[k]) throw new Error("받지 못함"); await clmAppend(doc, t, tbuf[k], true); used.push(clmCompany(t.name) || "보험사 모름"); }
          catch (e) { console.error(e); failed.push(t.name); }
        }
        if (!used.length) break;
        if (hdoc.getPageCount()) (await doc.copyPages(hdoc, hdoc.getPageIndices())).forEach((pg) => doc.addPage(pg));
        bset(`저장 준비 중…${tag}`);
        bytes = await doc.save();
        if (bytes.length <= LIMIT) break;
        console.warn("보험금청구서", dpi + "dpi", Math.round(bytes.length / 1048576) + "MB → 해상도 낮춰 다시");
      }
      clmDpi = 300;
      if (used.length && bytes) {
        const dir = await recDir(true), have = new Set((await docsList(dir)).map((f) => f.name));
        const base = `${clmYmd()} - 보험금청구서 - ${[...new Set(used)].join("·")}`;
        let name = base + ".pdf", n = 2;
        while (have.has(name)) name = `${base} (${n++}).pdf`;
        bset(`저장 중 (${(bytes.length / 1048576).toFixed(1)}MB)…`);
        const it = await docsUpload(dir, name, new File([bytes], name, { type: "application/pdf" }));
        if (it && it.id) made.push({ id: it.id, name: it.name || name, mime: "application/pdf", size: bytes.length, mod: new Date().toISOString(), webUrl: it.webViewLink || "", thumb: "" });
        current.clmSize = `${(bytes.length / 1048576).toFixed(1)}MB${dpiUsed !== 300 ? ` · 용량이 커서 ${dpiUsed}dpi로 낮춤` : ""}${bytes.length > LIMIT ? " · 그래도 8MB가 넘어 팩스로 보내기 어려울 수 있음" : ""}`;
      }
    } catch (e) { handleError(e); }
    clmDpi = 300; current.clmBusy = "";
    if (current.cust !== cust) return;
    current.recPend = made.reverse().concat(current.recPend || []);
    current.recNote = [made.length ? `1개 만듦(${current.clmSize || ""})` : "", failed.length ? `청구서 ${failed.length}개는 붙이지 못함` : "", skipped.size ? `병원서류 ${skipped.size}개는 붙이지 못함(PDF·사진·TIF만 가능)` : ""].filter(Boolean).join(", ");
    if (made.length) toast("보험금청구서를 한 파일로 만들었습니다");
    if (clmOnTab(cust)) await renderClaim();
  }
  async function clmDownload(f, btn) {
    if (btn) btn.disabled = true;
    try { const u = URL.createObjectURL(new Blob([await stGet0(f.id)], { type: "application/pdf" })), a = document.createElement("a"); a.href = u; a.download = (current && current.cust && !String(f.name).includes(baseName(current.cust.name)) ? baseName(current.cust.name) + " - " : "") + f.name; document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(u), 60000); }
    catch (e) { handleError(e); } finally { if (btn) btn.disabled = false; }
  }
  // ===== v551(2026-11-22 사용자 지시, 시안 확인 "이대로 코딩") 청구 탭 새 구조 =====
  //   맨 위 메뉴 4개(개인청구서 업로드 / 기본청구서 선택 / 병원서류 업로드 / 청구기록) → 누른 것의 칸이 아래에 열림(처음 = 개인청구서)
  //   기본청구서 = 앱 설정 폴더 > 보험금청구서식(안내 박스 없는 원본 + 위치.json). 1쪽 안내 박스를 담당 FC(이름·연락처)로 새로 그려 넣음(예전 공용 샘플 '김신한' 자리)
  //   보험사 순서: 자주 쓴 서식(내설정 clmUse 횟수 많은 순) → 추천 7곳 → 나머지 이름순, 묶음 사이 가로줄
  //   [보험금청구서 다운로드]: 고른 개인청구서 → 고른 기본청구서 → (있으면) 병원서류를 A4 회색조 300dpi JPEG 품질 70%로 끝에 → 한 PDF. 보험금청구 기록 폴더에 저장 + 기기에 내려받기 + 청구기록.json에 한 줄
  const CLM_PIN = ["신한라이프", "삼성화재", "메리츠화재", "현대해상", "한화손해보험", "DB손해보험", "KB손해보험"];
  const BF_DIR = "보험금청구서식", BF_POS = "위치.json", REC_FILE = "청구기록.json";
  let bfP = null; // 기본청구서 목록 { files: Map(보험사 → 파일), pos: {보험사: {x,y,w,alpha}}, ratio }
  function bfLoad() {
    return bfP || (bfP = (async () => {
      const dir = await resolveAppDir(); if (!dir || !dir.id) return { files: new Map(), pos: {}, ratio: 1168 / 401 };
      const sub = (await gdQuery("'" + qv(dir.id) + "' in parents and name='" + qv(BF_DIR) + "' and mimeType='" + GD_FOLDER + "' and trashed=false"))[0];
      if (!sub) return { files: new Map(), pos: {}, ratio: 1168 / 401 };
      const list = await gdQuery("'" + qv(sub.id) + "' in parents and trashed=false"), files = new Map();
      let pos = {}, ratio = 1168 / 401;
      for (const f of list) {
        if (/\.pdf$/i.test(f.name)) files.set(f.name.replace(/\.pdf$/i, ""), { id: f.id, name: f.name, mime: "application/pdf" });
        else if (f.name === BF_POS) { try { const j = parseDoc(await stGet0(f.id)); pos = j.서식 || {}; ratio = Number(j.박스비율) || ratio; } catch (e) { console.warn("기본청구서 위치 파일", e); } }
      }
      return { files, pos, ratio };
    })().catch((e) => { bfP = null; throw e; }));
  }
  const clmUse = () => ((myPrefs && myPrefs.clmUse) || {});
  function bfGroups(cos) { // [[묶음 이름, [보험사…]]…]
    const use = clmUse(), have = new Set(cos);
    const used = cos.filter((c) => use[c] > 0).sort((a, b) => use[b] - use[a] || a.localeCompare(b, "ko"));
    const pin = CLM_PIN.filter((c) => have.has(c) && !used.includes(c));
    const rest = cos.filter((c) => !used.includes(c) && !pin.includes(c)).sort((a, b) => a.localeCompare(b, "ko"));
    return [["자주 쓴 서식", used], ["추천 보험사", pin], ["나머지 (이름순)", rest]].filter(([, l]) => l.length);
  }
  async function clmFcDefault() { // 로그인한 FC의 이름·연락처: 내설정에 적어 둔 것 → (관리자) 팀원 파일 → 이름만
    const nm = await myFcName(), pr = (await loadMyPrefs().catch(() => null)) || {}, sv = pr.clmFc || {};
    if (sv.name && (!nm || sv.name === nm)) return { name: sv.name, phone: sv.phone || "" };
    let phone = "";
    try { if (isAdmin()) { const d = await stFindPath(TM_DIR), doc = d ? await stReadJsonIn(d, TM_FILE) : null; const p = ((doc && doc.사람) || []).find((x) => (nm && x.이름 === nm) || String(x.이메일 || "").toLowerCase() === me()); if (p) phone = String(p.연락처 || ""); } } catch (e) { /* 없으면 빈칸 */ }
    return { name: nm, phone };
  }
  const recDay = (f) => { const m = String(f.name || "").match(/^(\d{4})(\d{2})(\d{2})\b/); return m ? `${m[1]}-${m[2]}-${m[3]}` : String(f.mod || "").slice(0, 10); };
  async function recRead() { // 청구기록.json(보험금청구 폴더) + 예전에 만든 PDF(기록 파일에 없는 것)
    const top = current.clmTop; let doc = null;
    try { doc = top ? await stReadJsonIn(top, REC_FILE) : null; } catch (e) { console.warn("청구기록 읽기", e); }
    const items = Array.isArray(doc && doc.기록) ? doc.기록.slice() : [];
    const known = new Set(items.map((x) => x.pdfId).filter(Boolean));
    (current.recList || []).forEach((f) => { if (!known.has(f.id)) { const m = String(f.name).match(/보험금청구서 - (.+?)(?: \(\d+\))?\.pdf$/i); items.push({ date: recDay(f), acc: "", cos: m ? m[1].split("·") : [], pdfId: f.id, pdfName: f.name, old: true }); } });
    return items.sort((a, b) => String(b.date).localeCompare(String(a.date)) || String(b.at || "").localeCompare(String(a.at || "")));
  }
  const clmSelN = () => ((current.clmOwnSel || new Set()).size + (current.clmBaseSel || new Set()).size);
  function clmFootHtml(ro) {
    const own = (current.clmOwnSel || new Set()).size, bs = (current.clmBaseSel || new Set()).size, h = (current.hspList || []).length, n = own + bs;
    const busy = current.clmBusy || "";
    const sum = n ? `<b>${[own ? `개인 ${own}` : "", bs ? `기본 ${bs}` : ""].filter(Boolean).join(" · ")}</b> 청구서${h ? ` + 병원서류 ${h}개` : ""} → 한 PDF${h ? "<br>병원서류는 A4 회색조(300dpi)로 바꿔 끝에 붙입니다" : ""}` : "청구서를 하나 이상 고르면 다운로드할 수 있습니다";
    return `<div class="card cx-foot" id="clmFoot"><div class="sum">${sum}${current.recNote ? `<br>${esc(current.recNote)}` : ""}</div>${ro ? "" : `<div class="bar"><input type="text" data-clm-acc placeholder="사고 내용 (예: 무릎 염좌 통원)" value="${esc(current.clmAcc || "")}" maxlength="60"><button type="button" class="go" data-clm-build${n && !busy ? "" : " disabled"}>${esc(busy || "보험금청구서 다운로드")}</button></div>`}</div>`;
  }
  function clmFootUpdate() { const f = $("clmFoot"); if (f) f.outerHTML = clmFootHtml(isRO(current.fc)); const m = $("tabBody").querySelector(".cx-menu"); if (m) m.outerHTML = clmMenuHtml(); }
  function clmMenuHtml() {
    const tab = current.clmTab || "own", bs = (current.clmBaseSel || new Set()).size, os = (current.clmOwnSel || new Set()).size;
    const it = [["own", "개인청구서", "업로드", `${(current.clmList || []).length}개${os ? ` · ${os}선택` : ""}`], ["base", "기본청구서", "선택", bs ? `${bs}곳 선택` : "고르기"], ["hsp", "병원서류", "업로드", `${(current.hspList || []).length}개`], ["rec", "청구기록", "", `${(current.recItems || []).length}건`]];
    return `<div class="cx-menu" role="tablist">${it.map(([k, a, b, c]) => `<button type="button" role="tab" data-clm-tab="${k}" class="${k === tab ? "on" : ""}" aria-selected="${k === tab}">${a}<br>${b || "&nbsp;"}<small>${c}</small></button>`).join("")}</div>`;
  }
  async function renderClaim(note) {
    const body = $("tabBody"), cust = current.cust, ro = isRO(current.fc);
    if (!GD) { body.innerHTML = '<div class="card"><p class="hint">보험금 청구는 구글 드라이브에서만 쓸 수 있습니다.</p></div>'; return; }
    if (!current.folders[0]) { body.innerHTML = '<div class="card"><p class="hint">이 고객의 폴더가 없어 보험금 청구를 쓸 수 없습니다.</p></div>'; return; }
    if (!current.clmList) body.innerHTML = '<p class="hint pad">보험금 청구 자료를 확인하는 중…</p>';
    try {
      const dir = await clmDir(false);
      if (!clmOnTab(cust)) return;
      let list = dir ? await docsList(dir) : [];
      if (!clmOnTab(cust)) return;
      const pend = (current.clmPend || []).filter((x) => !list.some((f) => f.id === x.id)); current.clmPend = pend; list = pend.concat(list);
      if (pend.length && !current.clmRetry) { current.clmRetry = 1; setTimeout(() => { if (clmOnTab(cust)) renderClaim(note); }, 3000); } else if (!pend.length) current.clmRetry = 0;
      current.clmList = list;
      const hdir = dir ? await hspDir(false) : null; let hlist = hdir ? await docsList(hdir) : [];
      if (!clmOnTab(cust)) return;
      const hpend = (current.hspPend || []).filter((x) => !hlist.some((f) => f.id === x.id)); current.hspPend = hpend; hlist = hpend.concat(hlist);
      if (hpend.length && !current.clmRetry) { current.clmRetry = 1; setTimeout(() => { if (clmOnTab(cust)) renderClaim(note); }, 3000); }
      current.hspList = hlist;
      const rdir = dir ? await recDir(false) : null; let rlist = rdir ? await docsList(rdir) : [];
      if (!clmOnTab(cust)) return;
      const rpend = (current.recPend || []).filter((x) => !rlist.some((f) => f.id === x.id)); current.recPend = rpend; rlist = rpend.concat(rlist);
      current.recList = rlist;
      current.recItems = dir ? await recRead() : [];
      let bf = { files: new Map(), pos: {} };
      try { bf = await bfLoad(); } catch (e) { console.warn("기본청구서 목록", e); }
      await loadMyPrefs().catch(() => null);
      if (!current.clmFcSet) { const d = await clmFcDefault(); if (!current.clmFcSet) { current.clmFc = d; current.clmFcDef = d; current.clmFcSet = true; } }
      if (!clmOnTab(cust)) return;
      const tab = current.clmTab || (current.clmTab = "own");
      const ownSel = current.clmOwnSel || (current.clmOwnSel = new Set()), baseSel = current.clmBaseSel || (current.clmBaseSel = new Set());
      [...ownSel].forEach((id) => { if (!list.some((f) => f.id === id)) ownSel.delete(id); });
      const xBtn = (kind, i, f) => (ro ? "" : `<button type="button" class="x" data-${kind}-del="${i}" aria-label="${esc(f.name)} 삭제">×</button>`);
      const md = (f) => { const d = recDay(f); return d ? d.slice(5) : ""; };
      const fsz = (n) => (n >= 1048576 ? (n / 1048576).toFixed(1) + "MB" : n ? Math.max(1, Math.round(n / 1024)) + "KB" : "");
      let panel = "";
      if (tab === "own") {
        panel = `<div${ro ? "" : ' data-drop="clm"'}><div class="cx-head"><b>개인청구서 리스트</b><span id="clmOwnCnt">${list.length}개${ownSel.size ? ` · ${ownSel.size}개 선택` : ""}${note ? ` · ${esc(note)}` : ""}</span>${ro ? "" : '<button type="button" class="ev-edit" data-clm-add>업로드</button><input type="file" id="clmFile" accept=".pdf,image/*" multiple hidden>'}</div>
          ${list.length ? list.map((f, i) => `<div class="cx-row"><input type="checkbox" data-own-sel="${esc(f.id)}"${ownSel.has(f.id) ? " checked" : ""} aria-label="${esc(f.name)} 고르기"><button type="button" class="nm cx-link" data-clm-open="${i}" title="${esc(f.name)}">${esc(f.name)}</button><span class="mt">${esc(md(f))}</span>${xBtn("clm", i, f)}</div>`).join("") : `<p class="hint">이 고객에게 올린 개인청구서가 없습니다.${ro ? "" : " [업로드]로 PDF를 올리거나 여기에 끌어다 놓으세요."}</p>`}</div>`;
      } else if (tab === "base") {
        const fc = current.clmFc || {}, cos = [...bf.files.keys()];
        [...baseSel].forEach((c) => { if (!bf.files.has(c)) baseSel.delete(c); });
        const use = clmUse();
        panel = `<div class="cx-fc"><div class="lb">청구서 안내 박스의 담당 FC (공용 샘플 ‘김신한’ 대신 넣을 사람)</div>
          <div class="in"><label for="clmFcN">이름</label><input id="clmFcN" class="nmI" data-clmfc="name" value="${esc(fc.name || "")}" maxlength="10"${ro ? " disabled" : ""}><label for="clmFcP">연락처</label><input id="clmFcP" class="phI" data-clmfc="phone" type="tel" value="${esc(fc.phone || "")}" placeholder="010-0000-0000" maxlength="16"${ro ? " disabled" : ""}></div>
          <div class="pv" id="clmFcPv">${clmFcPv()}</div>
          <div class="rs"><span>기본값: 로그인한 FC</span>${ro ? "" : '<button type="button" class="cx-a" data-clmfc-reset>되돌리기</button>'}</div></div>
          <div class="cx-head"><b>보험사 기본청구서</b><span id="clmBaseCnt">${cos.length}곳${baseSel.size ? ` · ${baseSel.size}곳 선택` : ""}</span></div>
          ${cos.length ? `<div class="cx-grid">${bfGroups(cos).map(([g, l]) => `<div class="cx-gl">${g}</div>` + l.map((c) => `<label class="cx-row"><input type="checkbox" data-base-sel="${esc(c)}"${baseSel.has(c) ? " checked" : ""}><span class="nm">${esc(c)}</span>${use[c] ? `<span class="mt">${use[c]}회</span>` : ""}</label>`).join("")).join("")}</div>` : '<p class="hint">기본청구서 서식을 찾지 못했습니다(앱 설정 폴더 › 보험금청구서식).</p>'}`;
      } else if (tab === "hsp") {
        panel = `<div${ro ? "" : ' data-drop="hsp"'}><div class="cx-head"><b>병원서류</b><span>${hlist.length}개${current.hspNote ? ` · ${esc(current.hspNote)}` : ""}</span></div>
          ${hlist.map((f, i) => `<div class="cx-row"><button type="button" class="nm cx-link" data-hsp-open="${i}" title="${esc(f.name)}">${esc(f.name)}</button><span class="mt">${esc(fsz(f.size))}</span>${xBtn("hsp", i, f)}</div>`).join("")}
          ${ro ? (hlist.length ? "" : '<p class="hint">올린 병원서류가 없습니다.</p>') : hlist.length
            ? '<div class="cx-drop sm"><button type="button" class="ev-edit" data-hsp-add>추가 업로드</button><span>또는 파일을 여기에 끌어다 놓으세요 (PDF·사진·팩스 TIF)</span></div>'
            : '<div class="cx-drop"><button type="button" class="ev-edit" data-hsp-add>업로드</button><br>영수증·세부내역·진단서를 올리거나<br>파일을 여기에 끌어다 놓으세요 (PDF·사진·팩스 TIF)</div>'}
          ${ro ? "" : '<input type="file" id="hspFile" accept=".pdf,image/*,.tif,.tiff" multiple hidden>'}</div>`;
        current.hspNote = "";
      } else {
        const it = current.recItems || [];
        panel = `<div class="cx-head"><b>청구기록</b><span>${it.length}건</span></div>
          ${it.length ? it.map((r, i) => `<div class="cx-row cx-rec"><span class="dt">${esc(r.date || "")}</span><button type="button" class="nm cx-link" data-rec-open="${i}"><b>${esc(r.acc || (r.old ? "예전 기록" : "사고 내용 없음"))}</b><span class="co">${esc((r.cos || []).join(" · ") || "-")}</span></button>${ro ? "" : `<button type="button" class="x" data-rec-del="${i}" aria-label="청구기록 삭제">×</button>`}</div>`).join("") + '<p class="hint cx-note">줄을 누르면 그때 만든 PDF가 열립니다</p>' : '<p class="hint">아직 만든 보험금청구서가 없습니다. 아래 [보험금청구서 다운로드]로 만들면 여기에 남습니다.</p>'}`;
      }
      body.innerHTML = `${clmMenuHtml()}<div class="card cx-panel">${panel}</div>${clmFootHtml(ro)}`;
      current.recNote = "";
    } catch (e) { if (clmOnTab(cust)) body.innerHTML = '<div class="card"><p class="hint">보험금 청구 자료를 불러오지 못했습니다.</p></div>'; handleError(e); }
  }
  const clmFcLine = () => { const fc = current.clmFc || {}; return `${fc.name || "○○○"}FC(${fc.phone || "010-0000-0000"})`; };
  const clmFcPv = () => `서류 보완 및 기타문의 시<br><b>${esc(clmFcLine())}</b>에게<br>반드시 먼저 연락주세요 ^^`;
  async function clmBoxPng(doc, alpha) { // 안내 박스 그림(1168×401, 원래 박스와 같은 모양): 테두리·글씨는 진하게, 바탕은 alpha만큼 흰색
    const W = 1168, H = 401, cv = document.createElement("canvas"); cv.width = W; cv.height = H;
    const g = cv.getContext("2d"); g.fillStyle = `rgba(255,255,255,${alpha == null ? 1 : alpha})`; g.fillRect(0, 0, W, H);
    g.strokeStyle = "#000"; g.lineWidth = 12; g.strokeRect(14, 14, W - 28, H - 28);
    g.fillStyle = "#000"; g.textAlign = "center"; g.textBaseline = "middle";
    const fam = "'Malgun Gothic','Apple SD Gothic Neo','Noto Sans KR',sans-serif", line2 = clmFcLine() + "에게";
    let sz = 86; g.font = `${sz}px ${fam}`; while (sz > 40 && g.measureText(line2).width > 1060) { sz -= 2; g.font = `${sz}px ${fam}`; }
    g.font = `86px ${fam}`; g.fillText("서류 보완 및 기타문의 시", W / 2, 110);
    g.font = `${sz}px ${fam}`; g.fillText(line2, W / 2, 205);
    g.font = `86px ${fam}`; g.fillText("반드시 먼저 연락주세요 ^^", W / 2, 300);
    return doc.embedPng(await new Promise((r) => cv.toBlob(async (b) => r(new Uint8Array(await b.arrayBuffer())), "image/png")));
  }
  let clmGray = false; // v551: true면 그림 쪽을 흑백 정리 대신 회색조 JPEG(품질 70%)로
  async function grayPage(doc, cv) {
    const W = cv.width, H = cv.height, g = cv.getContext("2d"), im = g.getImageData(0, 0, W, H), a = im.data;
    for (let i = 0; i < a.length; i += 4) { const v = (a[i] * 299 + a[i + 1] * 587 + a[i + 2] * 114) / 1000; a[i] = a[i + 1] = a[i + 2] = v; }
    g.putImageData(im, 0, 0);
    const jpg = await doc.embedJpg(await new Promise((r) => cv.toBlob(async (b) => r(new Uint8Array(await b.arrayBuffer())), "image/jpeg", 0.7)));
    const pg = doc.addPage([595.28, 841.89]); pg.drawImage(jpg, { x: 0, y: 0, width: 595.28, height: 841.89 });
  }
  async function clmBuild2() {
    const ro = isRO(current.fc); if (ro || current.clmBusy) return;
    const cust = current.cust, ownSel = current.clmOwnSel || new Set(), baseSel = current.clmBaseSel || new Set();
    const own = (current.clmList || []).filter((f) => ownSel.has(f.id));
    let bf; try { bf = await bfLoad(); } catch (e) { return handleError(e); }
    const cos = [...baseSel].filter((c) => bf.files.has(c)), order = new Map(bfGroups([...bf.files.keys()]).flatMap(([, l]) => l).map((c, i) => [c, i]));
    cos.sort((a, b) => (order.get(a) ?? 999) - (order.get(b) ?? 999));
    if (!own.length && !cos.length) return toast("청구서를 하나 이상 골라 주세요");
    const hsp = (current.hspList || []).slice().sort((a, b) => hspOrder(a) - hspOrder(b) || a.name.localeCompare(b.name, "ko"));
    const fc = current.clmFc || {};
    if (cos.length && (!fc.name || !fc.phone)) { current.clmTab = "base"; await renderClaim(); return toast("기본청구서 안내 박스에 넣을 담당 FC 이름과 연락처를 넣어 주세요"); }
    const acc = String(current.clmAcc || "").trim();
    const bset = (t) => { current.clmBusy = t; const b = $("tabBody") && $("tabBody").querySelector("[data-clm-build]"); if (b) { b.textContent = t || "보험금청구서 다운로드"; b.disabled = !!t; } };
    bset("준비 중…");
    const skipped = [], failed = []; let made = null, bytes = null;
    try {
      bset("PDF 도구 불러오는 중…");
      await rptLoad("https://cdn.jsdelivr.net/npm/pdf-lib@1.17.1/dist/pdf-lib.min.js");
      const doc = await PDFLib.PDFDocument.create();
      clmGray = true; clmDpi = 300;
      for (let k = 0; k < own.length; k++) { // 개인청구서: PDF는 쪽 그대로, 사진은 A4 회색조
        bset(`개인청구서 붙이는 중 ${k + 1}/${own.length}`);
        try { await clmAppend(doc, own[k], await stGet0(own[k].id), false); } catch (e) { if (String(e.message).includes("GOOGLE_AUTH")) throw e; console.warn(own[k].name, e); failed.push(own[k].name); }
      }
      for (let k = 0; k < cos.length; k++) { // 기본청구서: 1쪽 안내 박스를 담당 FC로
        const co = cos[k], f = bf.files.get(co), p = bf.pos[co];
        bset(`${co} 청구서 받는 중 ${k + 1}/${cos.length}`);
        try {
          const src = await PDFLib.PDFDocument.load(await stGet0(f.id), { ignoreEncryption: true });
          if (p && p.w) { const pg = src.getPages()[0], h = p.w / (bf.ratio || 1168 / 401), png = await clmBoxPng(src, p.alpha); pg.drawImage(png, { x: p.x, y: pg.getHeight() - p.y - h, width: p.w, height: h }); }
          (await doc.copyPages(src, src.getPageIndices())).forEach((pg) => doc.addPage(pg));
        } catch (e) { if (String(e.message).includes("GOOGLE_AUTH")) throw e; console.warn(co, e); failed.push(co); }
      }
      if (!doc.getPageCount()) throw new Error("붙일 수 있는 청구서가 없습니다");
      for (let i = 0; i < hsp.length; i++) { // 병원서류: A4 회색조 300dpi JPEG 70%
        bset(`병원서류 바꾸는 중 ${i + 1}/${hsp.length}`); await new Promise((r) => setTimeout(r, 30));
        try { await clmAppend(doc, hsp[i], await stGet0(hsp[i].id), true); } catch (e) { if (String(e.message).includes("GOOGLE_AUTH")) throw e; console.warn(hsp[i].name, e); skipped.push(hsp[i].name); }
      }
      bset("PDF 만드는 중…");
      bytes = await doc.save();
      const label = [...new Set(cos.concat(own.map((f) => clmCompany(f.name) || "개인청구서")))];
      const fname = `${clmYmd()} - 보험금청구서 - ${label.join("·")}`.slice(0, 120) + ".pdf";
      { // 기기에 내려받기
        const u = URL.createObjectURL(new Blob([bytes], { type: "application/pdf" })), a = document.createElement("a");
        a.href = u; a.download = `${baseName(cust.name)} - ${fname}`; document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(u), 60000);
      }
      const dir = await recDir(true), have = new Set((await docsList(dir)).map((f) => f.name));
      const base = `${clmYmd()} - 보험금청구서 - ${label.join("·")}`.slice(0, 120);
      let name = base + ".pdf", n = 2; while (have.has(name)) name = `${base} (${n++}).pdf`;
      bset(`저장 중 (${(bytes.length / 1048576).toFixed(1)}MB)…`);
      const it = await docsUpload(dir, name, new File([bytes], name, { type: "application/pdf" }));
      if (it && it.id) made = { id: it.id, name: it.name || name, mime: "application/pdf", size: bytes.length, mod: new Date().toISOString(), webUrl: it.webViewLink || "", thumb: "" };
      if (made) { // 청구기록.json에 한 줄
        const top = current.clmTop, doc0 = (await stReadJsonIn(top, REC_FILE).catch(() => null)) || {}, rows = Array.isArray(doc0.기록) ? doc0.기록 : [];
        rows.push({ date: `${clmYmd().slice(0, 4)}-${clmYmd().slice(4, 6)}-${clmYmd().slice(6)}`, acc, cos: label, own: own.map((f) => f.name), hsp: hsp.filter((f) => !skipped.includes(f.name)).map((f) => f.name), fc: `${fc.name || ""} ${fc.phone || ""}`.trim(), pdfId: made.id, pdfName: made.name, at: new Date().toISOString(), by: me() });
        await stWriteJsonIn(top, REC_FILE, { 설명: "신한FC m 청구 탭 청구기록(다운로드할 때마다 한 줄)", 기록: rows });
      }
      try { // 자주 쓴 서식 횟수 + 내 담당 FC 정보(로그인한 FC 이름일 때만) 기억
        const pr = { ...((await loadMyPrefs()) || {}) }, use = { ...(pr.clmUse || {}) };
        cos.forEach((c) => { use[c] = (use[c] || 0) + 1; }); pr.clmUse = use;
        const nm = await myFcName(); if (fc.name && fc.phone && (!nm || fc.name === nm)) pr.clmFc = { name: fc.name, phone: fc.phone };
        await saveMyPrefs(pr);
      } catch (e) { console.warn("청구 횟수 저장", e); }
    } catch (e) { handleError(e); }
    clmGray = false; clmDpi = 300; current.clmBusy = "";
    if (current.cust !== cust) return;
    if (made) { current.recPend = [made].concat(current.recPend || []); current.clmAcc = ""; }
    current.recNote = [made ? `다운로드함 (${(bytes.length / 1048576).toFixed(1)}MB) · 청구기록에 남김` : "", failed.length ? `청구서 ${failed.length}개는 붙이지 못함` : "", skipped.length ? `병원서류 ${skipped.length}개는 붙이지 못함(PDF·사진·TIF만 가능)` : ""].filter(Boolean).join(", ");
    if (made) toast("보험금청구서를 한 파일로 만들어 내려받았습니다");
    if (clmOnTab(cust)) await renderClaim();
  }
  // v551(2026-11-22 사용자 지시): 개인청구서를 올리면 원본은 '보험금청구서 › 원본'에 두고, 보험금청구서 스타일로 바꾼 PDF를 이 고객 청구서로 저장
  //   스타일 = A4 세로, 이미지 개선(회색조·밝기 단계 맞춤·바깥 여백 잘라냄), 위쪽 왼쪽부터 놓기, 회색조 300dpi JPEG 품질 70%
  const CLM_ORIG = "원본";
  // v552(2026-11-22 사용자 지시 "신분증은 따로 정렬하지 말것"): 신분증 사본은 자르기·돌리기·키우기·위쪽 왼쪽으로 옮기기를 하지 않고 원래 놓인 모양 그대로(회색조·밝기 맞춤만)
  //   판별: 파일 이름(신분증·주민등록·운전면허·여권·외국인등록) 또는 쪽 안 글씨 덩어리가 카드 비율(가로:세로 약 1.4~1.8)이고 쪽의 40% 이하인 경우
  const ID_RE = /신분증|주민등록|운전면허|면허증|여권|외국인등록/;
  let clmIdName = false; // 지금 바꾸는 파일 이름이 신분증인지(병원서류 쪽 바꾸기에서도 씀)
  async function clmStylePage(doc, src, sw, sh, isPage) {
    const c0 = document.createElement("canvas"); c0.width = sw; c0.height = sh;
    const g0 = c0.getContext("2d"); g0.fillStyle = "#fff"; g0.fillRect(0, 0, sw, sh); g0.drawImage(src, 0, 0, sw, sh);
    const im = g0.getImageData(0, 0, sw, sh), a = im.data, n = sw * sh, hist = new Uint32Array(256);
    for (let i = 0; i < a.length; i += 4) { const v = (a[i] * 299 + a[i + 1] * 587 + a[i + 2] * 114) / 1000 | 0; a[i] = v; hist[v]++; }
    const pct = (q) => { let acc = 0; for (let v = 0; v < 256; v++) { acc += hist[v]; if (acc >= n * q) return v; } return 255; };
    const lo = Math.min(90, pct(0.01)), hi = Math.max(lo + 60, Math.min(250, pct(0.6)) - 8); // 바탕(대부분 밝은 곳)은 흰색으로, 가장 진한 1%는 검정으로
    let x0 = sw, y0 = sh, x1 = -1, y1 = -1;
    for (let y = 0; y < sh; y++) for (let x = 0; x < sw; x++) {
      const i = (y * sw + x) * 4; let v = Math.round((a[i] - lo) * 255 / (hi - lo)); v = v < 0 ? 0 : v > 255 ? 255 : v;
      a[i] = a[i + 1] = a[i + 2] = v; a[i + 3] = 255;
      if (v < 200 && x % 2 === 0 && y % 2 === 0) { if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y; }
    }
    g0.putImageData(im, 0, 0);
    if (x1 < 0) { x0 = 0; y0 = 0; x1 = sw - 1; y1 = sh - 1; } // 빈 쪽
    const pad = Math.round(Math.max(sw, sh) * 0.01); x0 = Math.max(0, x0 - pad); y0 = Math.max(0, y0 - pad); x1 = Math.min(sw - 1, x1 + pad); y1 = Math.min(sh - 1, y1 + pad);
    const cw = x1 - x0 + 1, ch = y1 - y0 + 1, ar = cw / ch;
    const W = 2480, H = 3508, m = 59;
    const cv = document.createElement("canvas"); cv.width = W; cv.height = H;
    const g = cv.getContext("2d"); g.fillStyle = "#fff"; g.fillRect(0, 0, W, H); g.imageSmoothingQuality = "high";
    if (clmIdName || (isPage && ((ar > 1.4 && ar < 1.8) || (ar > 0.55 && ar < 0.72)) && cw * ch < sw * sh * 0.4) || (!isPage && ((ar > 1.5 && ar < 1.7) || (ar > 0.59 && ar < 0.67)))) { // 사진은 카드 비율(A4 가로 1.41과 겹치지 않게 좁게) // 신분증: 원래 모양 그대로
      if (isPage) { const k0 = Math.min(W / sw, H / sh); g.drawImage(c0, 0, 0, sw, sh, 0, 0, Math.round(sw * k0), Math.round(sh * k0)); } // 쪽 전체를 그대로
      else { const k0 = Math.min(1, (W - 2 * m) / sw, (H - 2 * m) / sh, 1300 / Math.max(sw, sh)); g.drawImage(c0, 0, 0, sw, sh, m, m, Math.round(sw * k0), Math.round(sh * k0)); } // 사진: 돌리거나 자르지 않고 원래 크기 정도로
      c0.width = c0.height = 0; await grayPage(doc, cv); cv.width = cv.height = 0; return;
    }
    const rot = cw > ch * 1.15, iw = rot ? ch : cw, ih = rot ? cw : ch;
    const k = Math.min((W - 2 * m) / iw, (H - 2 * m) / ih), dw = Math.round(iw * k), dh = Math.round(ih * k);
    if (rot) { g.save(); g.translate(m, m + dh); g.rotate(-Math.PI / 2); g.drawImage(c0, x0, y0, cw, ch, 0, 0, dh, dw); g.restore(); } // 가로로 긴 쪽은 세워서
    else g.drawImage(c0, x0, y0, cw, ch, m, m, dw, dh); // 위쪽 왼쪽부터
    c0.width = c0.height = 0;
    await grayPage(doc, cv); cv.width = cv.height = 0;
  }
  async function clmStylePdf(f) { // 올린 파일 하나 → 스타일 PDF(Uint8Array). 못 바꾸는 종류면 null
    await rptLoad("https://cdn.jsdelivr.net/npm/pdf-lib@1.17.1/dist/pdf-lib.min.js");
    const buf = await f.arrayBuffer(), doc = await PDFLib.PDFDocument.create(), name = String(f.name || ""), type = f.type || "";
    clmIdName = ID_RE.test(name);
    if (/\.pdf$/i.test(name) || type === "application/pdf") {
      await rptLoad("https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.min.js");
      if (window.pdfjsLib && !pdfjsLib.GlobalWorkerOptions.workerSrc) pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js";
      const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(buf.slice(0)) }).promise;
      for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i), vp = page.getViewport({ scale: 300 / 72 }), c = document.createElement("canvas"); c.width = Math.round(vp.width); c.height = Math.round(vp.height);
        const g = c.getContext("2d"); g.fillStyle = "#fff"; g.fillRect(0, 0, c.width, c.height); await page.render({ canvasContext: g, viewport: vp }).promise;
        await clmStylePage(doc, c, c.width, c.height, true); c.width = c.height = 0; page.cleanup();
      }
      pdf.destroy();
    } else if (/\.tiff?$/i.test(name) || type === "image/tiff") { for (const c of await tifPages(buf)) await clmStylePage(doc, c, c.width, c.height); }
    else if (/^image\//.test(type) && !/heic|heif/i.test(type + name)) {
      const blob = new Blob([buf], { type }); let bmp; try { bmp = await createImageBitmap(blob, { imageOrientation: "from-image" }); } catch (e) { bmp = await createImageBitmap(blob); }
      await clmStylePage(doc, bmp, bmp.width, bmp.height);
    } else return null;
    clmIdName = false;
    return doc.getPageCount() ? doc.save() : null;
  }
  async function clmAddFiles(files, btn, hsp) { // hsp = 병원서류 칸
    const cust = current.cust, ok = files.filter((f) => f.size <= DOCS_MAX), big = files.length - ok.length;
    if (!ok.length) { toast("50MB가 넘는 파일은 올릴 수 없습니다"); return; }
    const lbl0 = btn ? btn.textContent : "";
    if (btn) btn.disabled = true;
    let done = 0; const failed = [], upl = [];
    try {
      const dir = hsp ? await hspDir(true) : await clmDir(true);
      const have = new Set((await docsList(dir)).map((f) => f.name));
      for (const f of ok) {
        if (btn) btn.textContent = `올리는 중 ${done + 1}/${ok.length}`;
        const orig = String(f.name || "파일").replace(/[\\/:*?"<>|]/g, "").trim() || "파일", em = orig.match(/(\.[A-Za-z0-9]{1,5})$/), ext = em ? em[1] : "", co = hsp ? "" : clmCompany(orig); // 병원서류는 원래 이름 앞에 날짜·FC·고객만
        let base = hsp ? `${clmYmd()} - ${hspKind(orig)}${ext}` : clmStdName(co, ext); // v502: 병원서류도 'YYYYMMDD - 영수증.jpg'처럼 짧게 // 청구서: 'YYYYMMDD - 보험사.pdf'(못 찾으면 'YYYYMMDD - 청구서.pdf'), 병원서류: 원래 이름 앞에 날짜·FC·고객
        const m = base.match(/^(.*?)(\.[A-Za-z0-9]{1,5})?$/); let name = base, k = 2;
        while (have.has(name)) name = `${m[1]} (${k++})${m[2] || ""}`;
        let upF = f;
        if (!hsp) { // v551: 원본 보관 + 스타일 PDF로 바꿔 저장
          try {
            if (btn) btn.textContent = `바꾸는 중 ${done + 1}/${ok.length}`;
            const pdf = await clmStylePdf(f);
            if (pdf) {
              try { const od = await clmFolder(dir.id, CLM_ORIG, true); await docsUpload(od, name, f); } // 원본 보관(실패해도 바꾼 PDF는 저장)
              catch (e) { if (String(e.message).includes("GOOGLE_AUTH")) throw e; console.warn("개인청구서 원본 보관 실패", f.name, e); }
              const pb = name.replace(/\.[A-Za-z0-9]{1,5}$/, ""); name = pb + ".pdf"; let k2 = 2; while (have.has(name)) name = `${pb} (${k2++}).pdf`;
              upF = new File([pdf], name, { type: "application/pdf" });
            }
          } catch (e) { if (String(e.message).includes("GOOGLE_AUTH")) throw e; console.warn("개인청구서 스타일 바꾸기 실패 — 원본 그대로 저장", f.name, e); }
          if (btn) btn.textContent = `올리는 중 ${done + 1}/${ok.length}`;
        }
        try { const it = await docsUpload(dir, name, upF); if (it && it.id) upl.push({ id: it.id, name: it.name || name, mime: it.mimeType || upF.type || "", size: upF.size, mod: new Date().toISOString(), webUrl: it.webViewLink || "", thumb: "" }); have.add(name); done++; } catch (e) { if (String(e.message).includes("GOOGLE_AUTH")) throw e; console.error(e); failed.push(f.name); }
      }
    } catch (e) { handleError(e); }
    if (btn) { btn.disabled = false; btn.textContent = lbl0 || "업로드"; }
    const msg = [done ? `${done}개 올림` : "", failed.length ? `${failed.length}개 실패` : "", big ? `${big}개는 50MB 초과` : ""].filter(Boolean).join(", ");
    if (current.cust === cust) { current.clmRetry = 0; if (hsp) { current.hspPend = upl.reverse().concat(current.hspPend || []); current.hspNote = msg; } else current.clmPend = upl.reverse().concat(current.clmPend || []); if (clmOnTab(cust)) await renderClaim(hsp ? "" : msg); }
  }
  document.addEventListener("click", async (ev) => {
    const t = ev.target && ev.target.closest && ev.target.closest("#tabBody [data-clm-start], #tabBody [data-clm-add], #tabBody [data-clm-open], #tabBody [data-hsp-add], #tabBody [data-hsp-open], #tabBody [data-clm-del], #tabBody [data-hsp-del], #tabBody [data-clm-build], #tabBody [data-rec-open], #tabBody [data-rec-dl], #tabBody [data-rec-del], #tabBody [data-clm-tab], #tabBody [data-clmfc-reset]");
    if (!t || !current) return;
    if (t.hasAttribute("data-clm-start")) {
      if (isRO(current.fc)) return;
      t.disabled = true; t.textContent = "폴더 만드는 중…";
      try { await clmDir(true); await hspDir(true); toast("‘보험금청구’ 안에 보험금청구서·병원서류 폴더를 만들었습니다"); await renderClaim(); } catch (e) { t.disabled = false; t.textContent = "보험금 청구 시작"; handleError(e); }
      return;
    }
    if (t.hasAttribute("data-clm-add") || t.hasAttribute("data-hsp-add")) { const inp = $(t.hasAttribute("data-hsp-add") ? "hspFile" : "clmFile"); if (inp) { inp.value = ""; inp.click(); } return; }
    if (t.hasAttribute("data-clm-tab")) { current.clmTab = t.dataset.clmTab; await renderClaim(); return; } // v551: 메뉴 4개
    if (t.hasAttribute("data-clmfc-reset")) { current.clmFc = { ...(current.clmFcDef || {}) }; await renderClaim(); return; }
    if (t.hasAttribute("data-clm-build")) { if (!isRO(current.fc)) clmBuild2(); return; }
    if (t.hasAttribute("data-rec-open") || t.hasAttribute("data-rec-dl")) { const r = (current.recItems || [])[+(t.dataset.recOpen || t.dataset.recDl)]; if (!r || !r.pdfId) return toast("이 기록의 PDF를 찾지 못했습니다"); const g = (current.recList || []).find((x) => x.id === r.pdfId) || { id: r.pdfId, name: r.pdfName || "보험금청구서.pdf", mime: "application/pdf" }; if (t.hasAttribute("data-rec-dl")) clmDownload(g, t); else docsView(g); return; }
    if (t.hasAttribute("data-rec-del")) { // v551: 청구기록 한 줄 삭제 = 기록 파일에서 빼고, 그때 만든 PDF는 드라이브 휴지통으로
      const r = (current.recItems || [])[+t.dataset.recDel];
      if (!r || isRO(current.fc) || !confirm(`${r.date || ""} ${r.acc || ""} 청구기록을 삭제할까요?\n(그때 만든 PDF는 구글 드라이브 휴지통으로 옮겨집니다)`)) return;
      t.disabled = true;
      try {
        if (r.pdfId) { await stDel(r.pdfId).catch((e) => { if (!String(e.message).includes("404")) throw e; }); current.recPend = (current.recPend || []).filter((x) => x.id !== r.pdfId); }
        if (!r.old && current.clmTop) { const d0 = (await stReadJsonIn(current.clmTop, REC_FILE)) || {}; const rows = (Array.isArray(d0.기록) ? d0.기록 : []).filter((x) => !(x.pdfId === r.pdfId && x.at === r.at)); await stWriteJsonIn(current.clmTop, REC_FILE, { ...d0, 기록: rows }); }
        current.recNote = "청구기록 1건 삭제"; await renderClaim();
      } catch (e) { t.disabled = false; handleError(e); }
      return;
    }
    if (t.hasAttribute("data-clm-del") || t.hasAttribute("data-hsp-del")) { // v496: 파일 삭제(드라이브 휴지통)
      const hsp = t.hasAttribute("data-hsp-del"), g = hsp ? (current.hspList || [])[+t.dataset.hspDel] : (current.clmList || [])[+t.dataset.clmDel];
      if (!g || isRO(current.fc) || !confirm(`‘${g.name}’ 파일을 삭제할까요?\n(구글 드라이브 휴지통으로 옮겨집니다)`)) return;
      t.disabled = true;
      try { await stDel(g.id); if (hsp) { current.hspPend = (current.hspPend || []).filter((x) => x.id !== g.id); current.hspNote = "1개 삭제"; } else current.clmPend = (current.clmPend || []).filter((x) => x.id !== g.id); current.files = (current.files || []).filter((x) => x.id !== g.id); await renderClaim(hsp ? "" : "1개 삭제"); } catch (e) { t.disabled = false; handleError(e); }
      return;
    }
    const f = t.hasAttribute("data-hsp-open") ? (current.hspList || [])[+t.dataset.hspOpen] : (current.clmList || [])[+t.dataset.clmOpen]; if (f) docsView(f);
  });
  // v496: 청구서·병원서류 칸에 파일을 끌어다 놓아도 올림
  const clmDropBox = (ev) => (ev.target && ev.target.closest ? ev.target.closest("#tabBody [data-drop]") : null);
  const hasFiles = (ev) => !!(ev.dataTransfer && [...(ev.dataTransfer.types || [])].includes("Files"));
  document.addEventListener("dragover", (ev) => { const b = clmDropBox(ev); if (!b || !hasFiles(ev)) return; ev.preventDefault(); ev.dataTransfer.dropEffect = "copy"; document.querySelectorAll("#tabBody [data-drop].drop-on").forEach((x) => { if (x !== b) x.classList.remove("drop-on"); }); b.classList.add("drop-on"); });
  document.addEventListener("dragleave", (ev) => { const b = clmDropBox(ev); if (b && !b.contains(ev.relatedTarget)) b.classList.remove("drop-on"); });
  document.addEventListener("drop", (ev) => {
    const b = clmDropBox(ev); if (!b || !hasFiles(ev)) return;
    ev.preventDefault(); b.classList.remove("drop-on");
    if (!current || isRO(current.fc)) return;
    const files = [...ev.dataTransfer.files]; if (!files.length) return;
    const hsp = b.dataset.drop === "hsp";
    clmAddFiles(files, b.querySelector(hsp ? "[data-hsp-add]" : "[data-clm-add]"), hsp);
  });
  document.addEventListener("change", async (ev) => {
    const t = ev.target; if (!t || !current) return;
    if (t.id === "clmFile" && t.files && t.files.length) { clmAddFiles([...t.files], $("tabBody").querySelector("[data-clm-add]")); return; }
    if (t.id === "hspFile" && t.files && t.files.length) { clmAddFiles([...t.files], $("tabBody").querySelector("[data-hsp-add]"), true); return; }
    if (t.matches && t.matches("#tabBody [data-own-sel], #tabBody [data-base-sel]")) { // v551: 개인·기본 청구서 고르기
      const own = t.hasAttribute("data-own-sel"), set = own ? (current.clmOwnSel || (current.clmOwnSel = new Set())) : (current.clmBaseSel || (current.clmBaseSel = new Set())), k = own ? t.dataset.ownSel : t.dataset.baseSel;
      t.checked ? set.add(k) : set.delete(k); clmFootUpdate();
      const hs = $("clmOwnCnt"); if (hs && own) hs.textContent = `${(current.clmList || []).length}개${set.size ? ` · ${set.size}개 선택` : ""}`;
      const hb = $("clmBaseCnt"); if (hb && !own) hb.textContent = hb.textContent.replace(/ · \d+곳 선택$/, "") + (set.size ? ` · ${set.size}곳 선택` : "");
      return;
    }
    if (t.matches && t.matches("#tabBody [data-clm-sel]")) { const sel = current.clmSel || (current.clmSel = new Set()); t.checked ? sel.add(t.dataset.clmSel) : sel.delete(t.dataset.clmSel); { const bb = $("tabBody").querySelector("[data-clm-build]"); if (bb && !bb.disabled) bb.textContent = clmBuildLabel(); } const c = $("tabBody").querySelector(".clm-card .docs-cnt"); if (c) c.textContent = c.textContent.replace(/ · 보험사 (\d+)곳( 중 \d+곳 선택)?/, (a, n) => ` · 보험사 ${n}곳${sel.size ? ` 중 ${sel.size}곳 선택` : ""}`); return; }
    if (t.matches && t.matches("#tabBody [data-clm-co]")) { // 보험사를 못 찾은 파일: 고르면 이름을 바꿔 보험사가 들어가게
      const f = (current.clmList || [])[+t.dataset.clmCo], co = t.value; if (!f || !co || isRO(current.fc)) return;
      const em = String(f.name).match(/(\.[A-Za-z0-9]{1,5})$/), name = clmStdName(co, em ? em[1] : "");
      t.disabled = true;
      try { await gdFetch(GDAPI + "/files/" + encodeURIComponent(f.id) + "?fields=id,name&supportsAllDrives=true", { method: "PATCH", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ name }) }); current.clmPend = (current.clmPend || []).filter((x) => x.id !== f.id); current.clmList = null; await renderClaim(`${co}(으)로 이름 바꿈`); } catch (e) { t.disabled = false; handleError(e); }
    }
  });
  document.addEventListener("input", (ev) => { // v551: 담당 FC·사고 내용 입력
    const t = ev.target; if (!t || !current || !t.matches) return;
    if (t.matches("#tabBody [data-clmfc]")) { current.clmFc = { ...(current.clmFc || {}), [t.dataset.clmfc]: t.value.trim() }; const pv = $("clmFcPv"); if (pv) pv.innerHTML = clmFcPv(); return; }
    if (t.matches("#tabBody [data-clm-acc]")) current.clmAcc = t.value;
  });
  async function selectTab(id) {
    $("tabs").querySelectorAll("button").forEach((b) =>
      b.classList.toggle("on", b.dataset.tab === id));
    const tab = CFG.tabs.find((t) => t.id === id);
    const body = $("tabBody");
    $("customerView").classList.toggle("on-docs", id === "docs"); // v488: 자료 탭 = 아래 줄 왼쪽 끝 [업로드], [수정]은 숨김
    $("customerView").classList.toggle("no-edit", id !== "basic"); // v489(2026-11-19 사용자 지시): [수정]은 고칠 수 있는 정보 탭에서만(계약·분석·상담·청구·자료에서는 숨김)

    if (id === "contact") return renderContact();
    if (id === "consult") { body.innerHTML = consultHtml(); bindConsult(); return; }
    if (id === "claim") return renderClaim(); // v494: 보험금 청구
    if (id === "docs") return renderDocs();

    const files = filesForTab(tab);
    const listHtml = id === "basic" ? basicHtml() : id === "contract" ? contractHtml() : id === "analysis" ? coverageSummaryHtml() : ""; // 보장 요약은 분석 탭 맨 위
    if (!files.length) {
      body.innerHTML = listHtml; // 파일이 없으면 안내 문구 없이 비워 둠
      if (id === "basic") bindExtra();
      return;
    }
    body.innerHTML = listHtml + files.map((f, i) => {
      const k = kind(f.name);
      // 따로 '열기' 버튼 없이 칸(줄) 전체를 누르면 열림: 엑셀 = 내용 펼치기/접기, PDF = PDF 열기, 그 밖 = 새 창
      const inner = `<span class="badge ${k}">${KIND_LABEL[k]}</span><div class="file-name">${fileLabel(f)}</div>${k === "excel" ? '<span class="file-hint">내용 보기</span>' : ""}<span class="chev">›</span>`;
      const row = k === "excel" ? `<div class="file-row tap" data-act="view" role="button" tabindex="0">${inner}</div>`
        : k === "pdf" ? `<div class="file-row tap" data-open="${esc(f.id)}" role="button" tabindex="0">${inner}</div>`
        : `<a class="file-row tap" href="${esc(f.webUrl)}" target="_blank" rel="noopener">${inner}</a>`;
      return `<div class="card file" data-i="${i}">${row}<div class="sheet" hidden></div></div>`;
    }).join("");
    body.querySelectorAll('[data-act="view"]').forEach((b) =>
      b.addEventListener("click", () => {
        const card = b.closest(".file");
        toggleSheet(card, files[+card.dataset.i], b.querySelector(".file-hint") || b);
      }));
    // 엑셀 내용은 기본으로 접어 둠 (줄을 누르면 펼침)
    if (id === "basic") bindExtra();
  }

  // 정보 탭
  //   보기: 연락처 · (이메일) · 주소 · 통신사/차량번호/자동차보험갱신일 · (그 밖의 고객리스트 항목) / 미팅장소 · 메모
  //   수정(맨 위 '수정' 버튼): 같은 항목을 한 화면에서 고치고 저장
  //     연락처·주소·이메일 → 고객리스트 엑셀 + '우선' 1 / 나머지 → 고객추가정보 엑셀
  // 금액 표시: 숫자면 천 단위 쉼표 + 원 (엑셀 서식이 General이어도)
  const wonText = (v) => {
    const t = String(v ?? "").trim().replace(/원\s*$/, "");
    const n = Number(t.replace(/,/g, ""));
    return (t && isFinite(n) ? n.toLocaleString("ko-KR") : t) + "원";
  };
  const BASIC_EXCLUDE = ["상품명 - 보험료", "상품명", "보험료", "이관월", "이관일", "미팅장소", "고객 폴더", "분류", "No", "순번", "메모", "모집FC", "담당FC", "담당 FC", "우편번호", "성별", "폴더", "성향", "우선", "직접수정", "주민번호", "입력일시", "고객ID", "고객 ID", "SOL고객번호", "SOL 고객번호", "현재단계", "현재 단계", "고객분류", "고객 분류", "고객등급", "고객 등급", "마케팅동의 만료일", "마케팅 동의 만료일", "계약안내 수신", "통신사", "보장요약", "생일구분", "생일 구분", "차량번호", "자동차보험갱신일"]; // 차량번호·갱신일은 아래 '차량' 줄 하나로 합쳐 보여 줌 // 정보 탭 보기에서 숨기는 항목
  const BIRTH_KEYS = ["생년월일", "생일"];
  const EMAIL_COLS = ["이메일", "e-mail", "Email"];
  let editMode = false;

  function genderMark(v) {
    const s = String(v || "").trim();
    if (/^(남|남자|M|male)$/i.test(s)) return "남";
    if (/^(여|여자|F|female)$/i.test(s)) return "여";
    return s;
  }
  const genderOf = (c) => { for (const e of c.entries) { const f = e.fields.find(([k]) => k.trim() === "성별"); if (f && f[1]) return genderMark(f[1]); } return ""; };

  // 연락처(휴대폰): 📞 Call · 💬 SMS / 이메일: ✉️ @ — 오른쪽 끝 정렬 버튼
  const LAND_COLS = /^(자택|집|직장|회사)\s*전화(번호)?$/;
  const MOBILE_RE = /01[016789][-.\s]?\d{3,4}[-.\s]?\d{4}/;
  // 주소 옆 위치 아이콘: 누르면 카카오맵에서 그 주소를 검색한 지도가 새 창으로 열림 (누를 때만 주소가 카카오로 전달됨)
  // v456(2026-11-19 사용자 지시): 주소가 나오는 곳은 모두 끝에 [지도보기] (카카오맵에서 그 주소 검색, 우편번호 [12345]는 떼고 보냄)
  const mapBtn = (v) => { const a = String(v || "").trim().replace(/^\[[^\]]*\]\s*/, ""); return a ? `<a class="map-btn" href="https://map.kakao.com/link/search/${encodeURIComponent(a)}" target="_blank" rel="noopener noreferrer" title="카카오맵에서 이 주소 보기" aria-label="지도보기">지도</a>` : ""; };
  // v467(2026-11-19 사용자 지시): 고객 정보의 계좌 오른쪽 [복사] — 계좌 내용을 클립보드에 복사
  async function copyText(t) {
    try { await navigator.clipboard.writeText(t); return true; } catch (e) { /* 아래 방식으로 */ }
    try { const ta = document.createElement("textarea"); ta.value = t; ta.setAttribute("readonly", ""); ta.style.cssText = "position:fixed;left:-9999px;top:0;font-size:16px"; document.body.appendChild(ta); ta.select(); ta.setSelectionRange(0, t.length); const ok = document.execCommand("copy"); ta.remove(); return ok; } catch (e) { return false; }
  }
  document.addEventListener("click", async (ev) => {
    const b = ev.target && ev.target.closest && ev.target.closest("[data-copy]");
    if (!b) return;
    ev.preventDefault();
    const ok = await copyText(b.dataset.copy || "");
    toast(ok ? (b.dataset.copyMsg || "계좌를 복사했습니다") : "복사하지 못했습니다. 길게 눌러 복사해 주세요");
    if (ok) { b.textContent = "복사됨"; setTimeout(() => { if (b.isConnected) b.textContent = "복사"; }, 1500); }
  });
  function addrWithMap(v) {
    const a = String(v || "").trim();
    if (!a) return "";
    return `<div class="val-row"><span class="val">${esc(a)}</span><span class="acts">${mapBtn(a)}</span></div>`;
  }
  function valueWithActions(k, v) {
    let btns = "";
    const m = PHONE_COLS.includes(k.trim()) && String(v).match(MOBILE_RE);
    if (m) {
      const num = m[0].replace(/\D/g, "");
      btns = `<a class="act-btn" href="tel:${num}" aria-label="전화 걸기"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M22 16.9v3a2 2 0 0 1-2.2 2 19.8 19.8 0 0 1-8.6-3.1 19.5 19.5 0 0 1-6-6 19.8 19.8 0 0 1-3.1-8.7A2 2 0 0 1 4.1 2h3a2 2 0 0 1 2 1.7c.1 1 .4 1.9.7 2.8a2 2 0 0 1-.5 2.1L8.1 9.9a16 16 0 0 0 6 6l1.3-1.3a2 2 0 0 1 2.1-.4c.9.3 1.8.6 2.8.7a2 2 0 0 1 1.7 2z"/></svg><span>Call</span></a>
        <a class="act-btn" href="sms:${num}" aria-label="문자 보내기"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/></svg><span>SMS</span></a>`;
    } else if (LAND_COLS.test(k.trim()) && /^[\d\s().+-]{8,16}$/.test(String(v).trim()) && String(v).replace(/\D/g, "").length >= 8) {
      // 자택·직장 전화(일반 전화): 문자는 못 받으니 전화 버튼만
      const num = String(v).replace(/[^\d+]/g, "");
      btns = `<a class="act-btn" href="tel:${num}" aria-label="전화 걸기"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M22 16.9v3a2 2 0 0 1-2.2 2 19.8 19.8 0 0 1-8.6-3.1 19.5 19.5 0 0 1-6-6 19.8 19.8 0 0 1-3.1-8.7A2 2 0 0 1 4.1 2h3a2 2 0 0 1 2 1.7c.1 1 .4 1.9.7 2.8a2 2 0 0 1-.5 2.1L8.1 9.9a16 16 0 0 0 6 6l1.3-1.3a2 2 0 0 1 2.1-.4c.9.3 1.8.6 2.8.7a2 2 0 0 1 1.7 2z"/></svg><span>Call</span></a>`;
    } else if (/이메일|e-?mail/i.test(k) && /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(String(v).trim())) {
      btns = `<a class="act-btn" href="mailto:${esc(String(v).trim())}" aria-label="메일 보내기"><svg viewBox="0 0 24 24" aria-hidden="true"><rect x="2" y="4" width="20" height="16" rx="2"/><path d="m22 7-10 6L2 7"/></svg><span>@</span></a>`;
    }
    return btns ? `<div class="val-row"><span class="val ${m ? "nowrap" : ""}">${esc(v)}</span><span class="acts">${btns}</span></div>`
      : `<div>${esc(v)}</div>`;
  }

  // 명함 카드 안 전화 줄: 고객 리스트의 휴대전화 번호와 전화·문자 버튼 (정보 탭의 연락처 줄과 같은 값)
  function carrierOf(c) {
    for (const e of c.entries) { const v = fieldOf(e, EXTRA.carrierCol); if (v) return String(v).trim(); }
    return String(((current && current.extra && current.extra.values) || {})[EXTRA.carrierCol] || "").trim();
  }
  function custTelHtml(c) {
    for (const e of c.entries) for (const [k, v] of e.fields) {
      if (PHONE_COLS.includes(k.trim()) && String(v).match(MOBILE_RE)) {
        const num = String(v).match(MOBILE_RE)[0];
        const d = num.replace(/\D/g, "");
        const carrier = carrierOf(c); // 통신사는 번호 왼쪽에 작게
        return `<div class="n">${carrier ? `<span class="carrier">${esc(carrier)}</span>` : ""}${esc(fmtPhone(num))}</div><span class="acts" data-cn="${esc(c.name)}" data-cfc="${esc(memoFcOf(c))}">` +
          `<a class="act-btn" href="tel:${d}" aria-label="전화 걸기"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M22 16.9v3a2 2 0 0 1-2.2 2 19.8 19.8 0 0 1-8.6-3.1 19.5 19.5 0 0 1-6-6 19.8 19.8 0 0 1-3.1-8.7A2 2 0 0 1 4.1 2h3a2 2 0 0 1 2 1.7c.1 1 .4 1.9.7 2.8a2 2 0 0 1-.5 2.1L8.1 9.9a16 16 0 0 0 6 6l1.3-1.3a2 2 0 0 1 2.1-.4c.9.3 1.8.6 2.8.7a2 2 0 0 1 1.7 2z"/></svg></a>` +
          `<a class="act-btn" href="sms:${d}" aria-label="문자 보내기"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/></svg></a></span>`;
      }
    }
    return "";
  }

  // 맨 위: 생년월일 (남) · [계약] · 담당: 최수진FC   (계약/이관 등은 색 태그)
  function custSubHtml(c) {
    const g = genderOf(c);
    const types = [...new Set(c.entries.map((e) => e.type))];
    const fcs = [...new Set(c.entries.map((e) => e.fc).filter(Boolean))];
    // 왼쪽: 생년월일 · 담당 / 오른쪽 끝: 분류(계약·이관 등) 한 줄
    const cal = c.entries.map((e) => fieldOf(e, "생일구분")).find(Boolean) || ""; // 양력/음력 → 생년월일 앞에 (양)·(음)
    const calMark = /음/.test(cal) ? "(음) " : /양/.test(cal) ? "(양) " : "";
    const left = [c.birth ? esc(calMark + c.birth + (g ? ` (${g})` : "")) : "", fcs.length ? esc("담당: " + fcs.map((f) => f + "FC").join(", ")) : ""].filter(Boolean).join(" · ");
    return `<span class="sub-l">${left}</span><span class="sub-tags">${types.map((t) => `<span class="tag t-${esc(t)}">${esc(typeLabel(t))}${t === "진행제외" ? " · " + esc(c.entries.map((e) => fieldOf(e, "진행제외")).find(Boolean) || "") : ""}</span>`).join(" ")}</span>`;
  }

  const mdText = (v) => { const p = parseMonthDay(v); return p ? `${p[0]}월 ${p[1]}일` : String(v || ""); };

  function basicHtml() {
    return editMode ? basicFormHtml() : basicViewHtml();
  }

  // 주민번호(v396, 2026-11-17): data.json의 '주민번호' 한 칸을 그대로 보임(13자리 YYMMDD-NNNNNNN / 앞6+뒤1 YYMMDD-N / 생년월일 8자리 / 성별만 -N). 민감정보 접근이 허용된 계정에게만 보임
  function ssnRow(c, backs, row) {
    const v = String(currentValue(c, ["주민번호"]) || "").trim();
    if (v) return row("주민번호", `<div>${esc(fmtRrn(v))}</div>`);
    const bd = String(c.birth || "").replace(/\D/g, ""), front = bd.length >= 8 ? bd.slice(2, 8) : bd.slice(0, 6), back = String((backs[0] || [])[1] || "").replace(/\D/g, ""); // 옛 칸만 있는 파일
    if (front.length === 6 && back) return row("주민번호", `<div>${esc(front + "-" + back)}</div>`);
    return "";
  }
  // ---------- 소개 (4단계 ③, 2026-11-17 확정안 3-4) ----------
  // 칸: 소개확인(빈칸=안 물어봄 / 아님 / 소개 / 모름), 소개고객ID(소개자가 등록 고객일 때), 소개자명(등록 안 된 분일 때 이름만). 유입경로는 바꾸지 않음
  const introOf = (c) => { for (const e of c.entries) { const id = fieldOf(e, "소개고객ID"), nm = fieldOf(e, "소개자명"), cf = fieldOf(e, "소개확인"); if (cf || id || nm) return { conf: cf || "소개", id, name: nm }; } return { conf: "", id: "", name: "" }; };
  const custIdOf = (c) => (c.entries.map((e) => idOf(e)).find(Boolean) || "");
  let introCache = null;
  function introCounts() { // 소개자 고객ID → 소개한 서로 다른 고객 수(새 칸 없이 소개고객ID를 세어 계산)
    if (introCache) return introCache;
    const m = new Map();
    (listCache || []).forEach((e) => { if (e.sheet !== CUST_SHEET) return; const id = fieldOf(e, "소개고객ID"), me = idOf(e); if (!id || !me) return; if (!m.has(id)) m.set(id, new Set()); m.get(id).add(me); });
    introCache = new Map([...m].map(([k, v]) => [k, v.size]));
    return introCache;
  }
  const TROPHY_CLS = (n) => (n >= 5 ? "gold" : n >= 3 ? "silver" : n >= 1 ? "bronze" : ""); // 1~2명 동, 3~4명 은, 5명 이상 금
  const trophyHtml = (n) => { const k = TROPHY_CLS(n); return k ? `<svg class="trophy ${k}" viewBox="0 0 16 16" width="13" height="13" role="img" aria-label="소개 ${n}명"><title>소개 ${n}명</title><path fill="currentColor" d="M4 1.5h8V5a4 4 0 0 1-3.2 3.9V11H11v1.5H5V11h2.2V8.9A4 4 0 0 1 4 5V1.5ZM2 2.5h1.5V5H2.7A.7.7 0 0 1 2 4.3V2.5Zm10.5 0H14v1.8a.7.7 0 0 1-.7.7h-.8V2.5ZM4.5 13h7v1.5h-7V13Z"/></svg>` : ""; };
  const trophyOf = (c) => trophyHtml(introCounts().get(custIdOf(c)) || 0);
  // 소개자 후보: 같은 FC 등록 고객 중 이름에 글자가 들어간 사람(자기 자신 제외)
  function introCands(q, selfId, fc) {
    q = String(q || "").trim(); if (!q) return [];
    const seen = new Set(), out = [];
    (listCache || []).forEach((e) => {
      if (e.sheet !== CUST_SHEET || (fc && e.fc !== fc)) return;
      const id = idOf(e); if (!id || id === selfId || seen.has(id) || !String(e.name).includes(q)) return;
      seen.add(id);
      const by = (fieldOf(e, "생년월일") || String(e.birth || "")).slice(2, 4);
      out.push({ id, name: e.name, yy: /^\d\d$/.test(by) ? by : "", route: routeShow(fieldOf(e, "최초유입경로")) });
    });
    out.sort((a, b) => (a.name === q ? -1 : 0) - (b.name === q ? -1 : 0) || a.name.localeCompare(b.name, "ko"));
    return out.slice(0, 8);
  }
  const introNameOf = (id) => { const e = (listCache || []).find((x) => x.sheet === CUST_SHEET && idOf(x) === id); return e ? e.name : ""; };
  function introFldHtml(c) { // 고객 정보 수정 화면의 '소개' 입력(소개확인 선택 + 소개자 찾기)
    const it = introOf(c), conf = it.conf || "";
    const cur = it.id ? introNameOf(it.id) || it.id : it.name;
    return `<div class="fld intro-fld" id="introFld" data-conf="${esc(conf)}" data-id="${esc(it.id)}"><span>소개</span>
      <div class="intro-conf" role="group" aria-label="소개 여부">${[["소개", "소개"], ["아님", "아님"], ["모름", "모름"]].map(([v, l]) => `<button type="button" data-intro-conf="${v}" class="${conf === v ? "on" : ""}" aria-pressed="${conf === v ? "true" : "false"}">${l}</button>`).join("")}</div>
      <div class="intro-src"${conf === "소개" ? "" : " hidden"}><input id="introName" value="${esc(cur)}" placeholder="소개자 이름 — 등록 고객이면 고르기" autocomplete="off" aria-label="소개자 이름"><div id="introSug" class="intro-sug"></div><div class="hint" id="introNote">${it.id ? "등록 고객과 연결됨" : it.name ? "이름만 저장됨" : ""}</div></div></div>`;
  }
  function introSuggest() {
    const f = $("introFld"); if (!f || !current || !current.cust) return;
    const c = current.cust, q = $("introName").value.trim(), box = $("introSug"), note = $("introNote");
    const own = c.entries.find((e) => e.fc) || c.entries[0], cands = introCands(q, custIdOf(c), own && own.fc);
    box.innerHTML = cands.map((k) => `<button type="button" data-intro-pick="${esc(k.id)}" class="${f.dataset.id === k.id ? "on" : ""}">${esc(k.name)}${k.yy ? ` <small>${k.yy}년생</small>` : ""}${k.route ? ` <small>${esc(k.route)}</small>` : ""}</button>`).join("") +
      (q ? `<button type="button" data-intro-pick="" class="name-only${f.dataset.id ? "" : " on"}">‘${esc(q)}’ 이름만 남기기</button>` : "");
    note.textContent = !q ? "" : f.dataset.id ? "등록 고객과 연결됨" : cands.length ? `등록 고객 ${cands.length}명 — 골라야 연결됩니다. 고르지 않으면 이름만 저장` : "등록 고객에 없음 → 이름만 저장";
  }
  document.addEventListener("click", (ev) => {
    const cf = ev.target.closest && ev.target.closest("[data-intro-conf]"), pk = ev.target.closest && ev.target.closest("[data-intro-pick]");
    const f = $("introFld"); if (!f) return;
    if (cf) { const v = cf.dataset.introConf, on = f.dataset.conf === v; f.dataset.conf = on ? "" : v; f.querySelectorAll("[data-intro-conf]").forEach((b) => { const o = b.dataset.introConf === f.dataset.conf; b.classList.toggle("on", o); b.setAttribute("aria-pressed", o ? "true" : "false"); }); f.querySelector(".intro-src").hidden = f.dataset.conf !== "소개"; if (f.dataset.conf === "소개") { $("introName").focus(); introSuggest(); } return; }
    if (pk) { const id = pk.dataset.introPick; f.dataset.id = id; if (id) $("introName").value = introNameOf(id) || $("introName").value; introSuggest(); }
  });
  document.addEventListener("input", (ev) => { if (ev.target && ev.target.id === "introName") { const f = $("introFld"); if (f) { f.dataset.id = ""; introSuggest(); } } });
  function introChange(c) { // 저장할 값: { 소개확인, 소개고객ID, 소개자명 } 또는 바뀐 게 없으면 null
    const f = $("introFld"); if (!f) return null;
    const cur = introOf(c), conf = f.dataset.conf, nm = ($("introName") && $("introName").value.trim()) || "", id = f.dataset.id || "";
    let nConf = conf, nId = "", nName = "";
    if (conf === "소개") { if (id) nId = id; else if (nm) nName = nm; else nConf = ""; }
    if (nConf === (cur.conf || "") && nId === (cur.id || "") && nName === (cur.name || "")) return null;
    return { 소개확인: nConf, 소개고객ID: nId, 소개자명: nName };
  }
  // ---------- 소개찾기 (4단계 ③, 고객 화면 GBS 옆 버튼) ----------
  // 대상 = 등록된 고객 전부(고객ID별 한 사람). 묶음(지역·성별·직업·직장·유입 시기·가족·유입경로)을 겹쳐 쓰고, 소개고객부터 / 소개자부터 / 랭킹 세 가지로 본다
  const IS = { dir: "cust", fc: "", keys: new Set(), all: false, sel: new Set(), who: null, sheet: null };
  const IS_KEYS = ["지역", "성별", "직업", "직장", "유입 시기", "가족", "유입경로"];
  function introPeople() {
    const m = new Map();
    (listCache || []).forEach((e) => { if (e.sheet !== CUST_SHEET) return; const id = idOf(e); if (!id) return; if (!m.has(id)) m.set(id, []); m.get(id).push(e); });
    return [...m].map(([id, ents]) => {
      const e0 = ents.find((x) => x.fc) || ents[0], it = introOf({ entries: ents }), addr = fieldOf(e0, "주소");
      return { id, ents, fc: e0.fc, name: e0.name, route: routeShow(fieldOf(e0, "최초유입경로")), sex: fieldOf(e0, "성별"), yy: String(fieldOf(e0, "생년월일") || e0.birth || "").slice(2, 4), job: fieldOf(e0, "직업"), work: fieldOf(e0, "직장명"),
        region: fieldOf(e0, "지역") || String(addr).replace(/^\[[^\]]*\]\s*/, "").split(/\s+/).slice(0, 2).join(" "), ym: String(fieldOf(e0, "유입일") || "").slice(0, 7), fam: String(addr).replace(/\s+/g, "").slice(0, 20), conf: it.conf, introId: it.id, introName: it.name };
    });
  }
  const isGroupKey = (p) => [...IS.keys].map((k) => ({ 지역: p.region, 성별: p.sex, 직업: p.job, 직장: p.work, "유입 시기": p.ym, 가족: p.fam, 유입경로: p.route }[k] || "(없음)")).join(" · ") || "전체";
  const isOwnNames = (people) => { const t = new Map(); people.forEach((p) => { if (p.introName && !p.introId) t.set(p.introName, (t.get(p.introName) || 0) + 1); }); return t; };
  async function openIntro() { await loadLists(); IS.sel.clear(); IS.who = null; IS.sheet = null; show("introView"); renderIntro(); }
  function isRowHtml(p, mode) {
    const st = p.introId ? `소개: ${esc(introNameOf(p.introId) || "등록 고객")}` : p.introName ? `소개: ${esc(p.introName)}` : p.conf === "아님" ? "소개 아님" : p.conf === "모름" ? "모름" : "";
    const lead = mode === "pick" ? `<input type="checkbox" class="is-cb" data-is-sel="${esc(p.id)}" ${IS.sel.has(p.id) ? "checked" : ""} aria-label="${esc(p.name)} 선택">` : "";
    return `<li class="is-row">${lead}<span class="is-nm"><a href="#" data-cust="${esc(p.id)}">${trophyHtml(introCounts().get(p.id) || 0)}${esc(p.name)}</a>${p.yy ? ` <small>${p.yy}년생</small>` : ""}${p.route ? ` <small class="rt">${esc(p.route)}</small>` : ""}${p.sex ? ` <small>${esc(p.sex)}</small>` : ""}</span>
      <span class="is-st">${st}</span>${mode === "cust" ? `${p.conf ? "" : `<button type="button" class="ev-edit is-none" data-is-none="${esc(p.id)}" title="소개자 없음으로 저장(다시 묻지 않음)">없음</button>`}<button type="button" class="ev-edit" data-is-find="${esc(p.id)}">소개자 찾기</button>` : ""}</li>`; // v567: [없음] = 소개 아님으로 저장 → 다시 묻지 않음
  }
  // 로그인 때 새 고객 확인(확정안 3-4 화면 1): 지난번 접속 뒤 입력된 내 FC 고객(유입경로 상관없음) 중 소개를 아직 안 물은 고객이 있으면 소개찾기 화면이 뜸. 처음 접속은 기준 시각만 기록(기존 고객은 소개찾기에서 점검)
  const introSeenKey = () => "shlife.introSeen." + (typeof me === "function" ? me() : "");
  const introStamp = () => { const d = new Date(), p2 = (n) => String(n).padStart(2, "0"); return d.getFullYear() + "-" + p2(d.getMonth() + 1) + "-" + p2(d.getDate()) + " " + p2(d.getHours()) + ":" + p2(d.getMinutes()) + ":" + p2(d.getSeconds()); };
  // 기준 시각은 Drive의 내 설정 파일(shlife-내설정.json의 introSeen)에 둠 → 기기를 바꿔도 같은 기준. localStorage는 빠른 표시용 사본
  async function readIntroSeen() {
    let loc = ""; try { loc = localStorage.getItem(introSeenKey()) || ""; } catch (e) { /* 무시 */ }
    let drv = ""; try { drv = String(((await loadMyPrefs()) || {}).introSeen || ""); } catch (e) { /* 무시 */ }
    const last = drv && drv > loc ? drv : loc || drv; // 둘 중 최근 것
    if (last && last !== drv) writeIntroSeen(last, true); // 기기에만 있던 값은 Drive로 올림
    return last;
  }
  async function writeIntroSeen(stamp, quiet) {
    try { localStorage.setItem(introSeenKey(), stamp); } catch (e) { /* 무시 */ }
    try { // 저장 직전 최신 내 설정 파일을 읽어(못 읽으면 저장하지 않음 — 다른 설정을 덮어쓰지 않게) introSeen만 더해 저장
      const d = await resolveMyDir(true); if (!d || !d.id) return;
      const cur = await stReadJsonIn(d, MY_FILE);
      await saveMyPrefs({ ...(cur || myPrefs || {}), introSeen: stamp });
    }
    catch (e) { if (!quiet) console.warn("새 고객 확인 기준 시각을 Drive에 저장하지 못했습니다(이 기기에는 저장됨)", e && e.message); }
  }
  async function checkNewCustomers() {
    try {
      if (EMBED || !account()) return;
      { let lt = ""; try { lt = localStorage.getItem(introLaterKey()) || ""; } catch (e) { /* 무시 */ } if (lt && lt === introLoginMark()) return; } // v567: [다음에]를 누른 로그인 동안은 묻지 않음 → 다음 구글 로그인 뒤에 다시
      await loadLists();
      const last = await readIntroSeen();
      if (!last) { await writeIntroSeen(introStamp()); return; }
      const myFc = await myFcName();
      const fresh = introPeople().filter((p) => !p.conf && p.fc === myFc && String(fieldOf(p.ents[0], "입력일시") || "") > last).map((p) => p.id);
      if (!fresh.length) { await writeIntroSeen(introStamp()); return; }
      IS.fresh = fresh; IS.sel.clear(); IS.who = null; IS.sheet = null; IS.dir = "cust";
      show("introView"); renderIntro();
    } catch (e) { console.warn("새 고객 확인 건너뜀", e && e.message); }
  }
  const introLaterKey = () => "shlife.introLater." + (typeof me === "function" ? me() : "");
  const introLoginMark = () => { const a = readGAuth(); return a ? String(gLoginAt(a)) : ""; }; // 지금 구글 로그인(출입증을 받은 시각)
  function introFreshDone(saveSeen) { if (saveSeen) writeIntroSeen(introStamp()); IS.fresh = null; IS.sheet = null; resetSearch(); }
  function renderIntroFresh(box) {
    const map = new Map(introPeople().map((p) => [p.id, p])), ps = IS.fresh.map((id) => map.get(id)).filter((p) => p && !p.conf);
    if (!ps.length) { introFreshDone(true); return; }
    box.innerHTML = `<div class="is-head"><b>새 고객 확인</b></div><p class="hint">새로 들어온 고객 ${ps.length}명입니다. 소개고객이 있으면 소개자를 알려 주세요. 소개자가 없으면 [없음]을 누르면 다시 묻지 않습니다.</p>
      <ul class="is-list">${ps.map((p) => isRowHtml(p, "cust")).join("")}</ul>
      <div class="is-fresh-act"><button type="button" class="primary" data-is-rest="1">나머지는 소개 없음</button><button type="button" class="secondary" data-is-later="1">다음에</button></div>` + (IS.sheet ? isSheetHtml() : "");
    if (IS.sheet) { const i = $("isQ"); if (i) i.focus(); }
  }
  function renderIntro() {
    const box = $("introView"); if (!box) return;
    if (IS.fresh) return renderIntroFresh(box);
    const all = introPeople(), fcs = [...new Set(all.map((p) => p.fc).filter(Boolean))];
    const people = IS.fc ? all.filter((p) => p.fc === IS.fc) : all;
    const done = people.filter((p) => p.conf).length, pct = people.length ? Math.round((done / people.length) * 100) : 0;
    const head = `<div class="is-head"><b>소개찾기</b>${fcs.length > 1 ? `<select id="isFc" aria-label="FC 선택"><option value="">전체 FC</option>${fcs.map((f) => `<option value="${esc(f)}"${f === IS.fc ? " selected" : ""}>${esc(f)}</option>`).join("")}</select>` : ""}</div>
      <div class="is-prog"><span>등록고객 ${people.length}명 중 ${done}명 확인</span><i style="--w:${pct}%"></i></div>
      <div class="seg is-dir"><button type="button" data-is-dir="cust" class="${IS.dir === "cust" ? "on" : ""}">소개고객부터</button><button type="button" data-is-dir="who" class="${IS.dir === "who" ? "on" : ""}">소개자부터</button><button type="button" data-is-dir="rank" class="${IS.dir === "rank" ? "on" : ""}">랭킹</button></div>`;
    let body = "";
    if (IS.dir === "rank") {
      const reg = [...introCounts()].filter(([id]) => people.some((p) => p.id === id) || all.some((p) => p.id === id)).map(([id, n]) => ({ name: introNameOf(id) || id, n, id, reg: true }));
      const unreg = [...isOwnNames(people)].map(([name, n]) => ({ name, n, reg: false }));
      const rows = [...reg, ...unreg].sort((a, b) => b.n - a.n || a.name.localeCompare(b.name, "ko"));
      body = rows.length ? `<ol class="is-rank">${rows.map((r, i) => `<li><span class="no">${i + 1}</span><span class="is-nm">${r.reg ? trophyHtml(r.n) : ""}${r.reg ? `<a href="#" data-cust="${esc(r.id)}">${esc(r.name)}</a>` : esc(r.name)}${r.reg ? "" : ` <small class="rt">미등록</small>`}</span><b>${r.n}명</b>${r.reg ? "" : `<button type="button" class="ev-edit" data-is-link="${esc(r.name)}">고객으로 연결</button>`}</li>`).join("")}</ol>` : `<p class="hint pad">아직 소개 기록이 없습니다.</p>`;
    } else {
      const keyBar = `<div class="is-keys" role="group" aria-label="묶음">${IS_KEYS.map((k) => `<button type="button" data-is-key="${k}" class="${IS.keys.has(k) ? "on" : ""}" aria-pressed="${IS.keys.has(k) ? "true" : "false"}">${k}</button>`).join("")}</div>`;
      let list = people.filter((p) => IS.all || !p.conf);
      if (IS.dir === "who" && IS.who) list = list.filter((p) => p.id !== IS.who.id);
      if (IS.keys.has("가족")) { const fc = new Map(); people.forEach((p) => p.fam && fc.set(p.fam, (fc.get(p.fam) || 0) + 1)); list = list.filter((p) => p.fam && fc.get(p.fam) > 1); }
      const groups = new Map(); list.forEach((p) => { const k = isGroupKey(p); if (!groups.has(k)) groups.set(k, []); groups.get(k).push(p); });
      const gs = [...groups].sort((a, b) => b[1].length - a[1].length || a[0].localeCompare(b[0], "ko"));
      const who = IS.dir === "who" ? `<div class="is-who">${IS.who ? `소개자: <b>${esc(IS.who.name)}</b>${IS.who.id ? "" : ' <small class="rt">미등록</small>'} <button type="button" class="ev-edit" data-is-who="1">바꾸기</button>` : `<button type="button" class="ev-edit ev-done" data-is-who="1">소개자 고르기</button><span class="hint"> 먼저 소개자를 고른 뒤, 그분이 소개한 고객을 체크합니다.</span>`}</div>` : "";
      const act = IS.dir === "who" && IS.who && IS.sel.size ? `<div class="is-act"><button type="button" class="primary" data-is-go="1">${IS.sel.size}명을 ${esc(IS.who.name)}님 소개로 연결</button></div>` : "";
      body = who + keyBar + `<label class="is-all"><input type="checkbox" id="isAll" ${IS.all ? "checked" : ""}> 확인한 고객도 보기</label>` +
        (IS.dir === "who" && !IS.who ? "" : gs.length ? gs.map(([k, ps]) => `<div class="is-grp"><div class="is-gt"><b>${esc(k)}</b><span>${ps.length}명</span></div><ul class="is-list">${ps.map((p) => isRowHtml(p, IS.dir === "who" ? "pick" : "cust")).join("")}</ul></div>`).join("") : `<p class="hint pad">${IS.all ? "고객이 없습니다." : "아직 확인하지 않은 고객이 없습니다."}</p>`) + act;
    }
    box.innerHTML = head + body + (IS.sheet ? isSheetHtml() : "");
    if (IS.sheet) { const i = $("isQ"); if (i) i.focus(); }
  }
  // 소개자 고르기 창: 등록 고객(이름 검색) · 이미 적힌 이름 · 이름만 · 아님/모름
  function isSheetHtml() {
    const sh = IS.sheet, q = (sh.q || "").trim(), people = introPeople();
    const cands = q ? people.filter((p) => p.id !== sh.selfId && p.name.includes(q)).sort((a, b) => (a.name === q ? -1 : 0) - (b.name === q ? -1 : 0) || a.name.localeCompare(b.name, "ko")).slice(0, 8) : [];
    const names = q ? [...isOwnNames(people).keys()].filter((n) => n.includes(q)).slice(0, 5) : [];
    return `<div class="is-sheet" role="dialog" aria-label="소개자 고르기"><div class="is-sheet-in"><div class="is-sheet-h"><b>${esc(sh.title)}</b><button type="button" data-is-close="1" aria-label="닫기">닫기</button></div>
      <input id="isQ" value="${esc(sh.q || "")}" placeholder="소개자 이름" autocomplete="off">
      <div class="intro-sug">${cands.map((k) => `<button type="button" data-is-pick-id="${esc(k.id)}">${esc(k.name)}${k.yy ? ` <small>${k.yy}년생</small>` : ""}${k.route ? ` <small>${esc(k.route)}</small>` : ""}</button>`).join("")}${names.map((n) => `<button type="button" data-is-pick-name="${esc(n)}" class="name-only">이미 적힌 이름 ‘${esc(n)}’</button>`).join("")}${q && !names.includes(q) && !sh.noName ? `<button type="button" data-is-pick-name="${esc(q)}" class="name-only">‘${esc(q)}’ 이름만 남기기</button>` : ""}</div>
      ${sh.noName ? "" : `<div class="is-sheet-f"><button type="button" data-is-conf="아님">소개 아님</button><button type="button" data-is-conf="모름">모름</button></div>`}</div></div>`;
  }
  async function isApply(ids, vals, msg) {
    const map = new Map(introPeople().map((p) => [p.id, p])), ents = ids.flatMap((id) => (map.get(id) ? map.get(id).ents : []));
    if (!ents.length) return;
    try { await writeEntryCells(ents, vals); toast(msg || "저장했습니다."); }
    catch (e) { handleError(e); return; }
    IS.sheet = null; IS.sel.clear(); renderIntro();
  }
  const isVals = (kind, v) => (kind === "id" ? { 소개확인: "소개", 소개고객ID: v, 소개자명: "" } : kind === "name" ? { 소개확인: "소개", 소개고객ID: "", 소개자명: v } : { 소개확인: v, 소개고객ID: "", 소개자명: "" });
  document.addEventListener("click", async (ev) => {
    const t = ev.target; if (!t.closest || !$("introView") || $("introView").hidden) return;
    const g = (sel) => t.closest(sel);
    let b;
    if ((b = g("[data-is-dir]"))) { IS.dir = b.dataset.isDir; IS.sel.clear(); IS.sheet = null; return renderIntro(); }
    if ((b = g("[data-is-key]"))) { const k = b.dataset.isKey; IS.keys.has(k) ? IS.keys.delete(k) : IS.keys.add(k); return renderIntro(); }
    if ((b = g("[data-is-none]"))) { const p = introPeople().find((x) => x.id === b.dataset.isNone); b.disabled = true; await isApply([b.dataset.isNone], isVals("conf", "아님"), `${p ? p.name + "님은 " : ""}소개자 없음으로 저장했습니다. 다시 묻지 않습니다.`); return; }
    if ((b = g("[data-is-find]"))) { const p = introPeople().find((x) => x.id === b.dataset.isFind); if (p) { IS.sheet = { title: `${p.name}님의 소개자`, q: "", selfId: p.id, ids: [p.id] }; renderIntro(); } return; }
    if ((b = g("[data-is-who]"))) { IS.sheet = { title: "소개자 고르기", q: "", selfId: "", who: true }; return renderIntro(); }
    if ((b = g("[data-is-link]"))) { IS.sheet = { title: `‘${b.dataset.isLink}’ 님을 등록 고객으로 연결`, q: b.dataset.isLink, selfId: "", link: b.dataset.isLink, noName: true }; return renderIntro(); }
    if (g("[data-is-close]")) { IS.sheet = null; return renderIntro(); }
    if (g("[data-is-later]")) { try { localStorage.setItem(introLaterKey(), introLoginMark()); } catch (e) { /* 무시 */ } toast("다음 구글 로그인 뒤에 다시 물어봅니다."); return introFreshDone(false); } // v567
    if (g("[data-is-rest]")) { const map = new Map(introPeople().map((p) => [p.id, p])), ids = (IS.fresh || []).filter((id) => map.get(id) && !map.get(id).conf); await isApply(ids, isVals("conf", "아님"), `${ids.length}명을 소개 없음으로 저장했습니다.`); return; }
    if ((b = g("[data-is-sel]"))) { const id = b.dataset.isSel; b.checked ? IS.sel.add(id) : IS.sel.delete(id); return renderIntro(); }
    if (g("[data-is-go]")) { if (IS.who) return isApply([...IS.sel], IS.who.id ? isVals("id", IS.who.id) : isVals("name", IS.who.name), `${IS.sel.size}명을 연결했습니다.`); return; }
    const pickId = g("[data-is-pick-id]"), pickName = g("[data-is-pick-name]"), conf = g("[data-is-conf]");
    if (pickId || pickName || conf) {
      const sh = IS.sheet; if (!sh) return;
      if (sh.who) { const id = pickId && pickId.dataset.isPickId; IS.who = id ? { id, name: introNameOf(id) || id } : { id: "", name: (pickName || {}).dataset ? pickName.dataset.isPickName : "" }; IS.sheet = null; if (!IS.who.name) IS.who = null; return renderIntro(); }
      if (sh.link) { const id = pickId && pickId.dataset.isPickId; if (!id) return; const ids = introPeople().filter((p) => p.introName === sh.link && !p.introId).map((p) => p.id); return isApply(ids, isVals("id", id), `${ids.length}명을 등록 고객으로 연결했습니다.`); }
      if (pickId) return isApply(sh.ids, isVals("id", pickId.dataset.isPickId), "소개자를 저장했습니다.");
      if (pickName) return isApply(sh.ids, isVals("name", pickName.dataset.isPickName), "소개자 이름을 저장했습니다.");
      if (conf) return isApply(sh.ids, isVals("conf", conf.dataset.isConf), "저장했습니다.");
    }
  });
  document.addEventListener("input", (ev) => { // v493: 입력칸은 그대로 두고 아래 목록만 바꿈(화면 전체를 다시 그리면 한글 조합이 'ㅈㅓㅇ'처럼 깨짐)
    const t = ev.target; if (!(t && t.id === "isQ" && IS.sheet)) return;
    IS.sheet.q = t.value;
    const sug = t.parentNode && t.parentNode.querySelector(".intro-sug"); if (!sug) return;
    const tmp = document.createElement("div"); tmp.innerHTML = isSheetHtml();
    const nsug = tmp.querySelector(".intro-sug"); if (nsug) sug.innerHTML = nsug.innerHTML;
  });
  document.addEventListener("change", (ev) => { const t = ev.target; if (!t || !$("introView") || $("introView").hidden) return; if (t.id === "isFc") { IS.fc = t.value; renderIntro(); } else if (t.id === "isAll") { IS.all = t.checked; renderIntro(); } });
  document.addEventListener("click", (ev) => {
    const b = ev.target.closest && ev.target.closest("#introBtn"); if (b) return openIntro();
    const cu = ev.target.closest && ev.target.closest("#introView [data-cust]");
    if (cu) { ev.preventDefault(); const es = (listCache || []).filter((e) => idOf(e) === cu.dataset.cust); if (es.length) { const g = groupCustomers(es); g.forEach((x) => { x.entries.sort((a, b2) => prio(a.type) - prio(b2.type)); x.name = x.entries[0].name; }); openCustomer(g[0]); } }
  });
  // ---------- 대시보드 (4단계 ⑤ A안, 2026-11-17 사용자 선택) ----------
  // 위: v516 이번달 영업실적(건수·CMIP·CANP) + 최근 수수료 · 계약진행 인원 — 업적·수수료는 민감정보 허용 계정이 그 파일을 읽을 수 있을 때만
  // 가운데: 이관·따행·DB·계약진행 목표의 단계별 막대 / 아래: 소개 랭킹 상위 3명. 보는 범위는 FC 선택(내 권한으로 보이는 FC만, 기본은 본인)
  const DSH = { fc: "" };
  async function readClaudeJson(dirPath, name) { try { const d = await stFindPathCached(dirPath); if (!d) return null; return await stReadJsonIn(d, name); } catch (e) { return null; } }
  const wonShort = (v) => { const n = Math.round(+v || 0); return n >= 10000 ? (Math.round(n / 1000) / 10).toLocaleString("ko-KR") + "만원" : n.toLocaleString("ko-KR") + "원"; };
  function dashBar(label, total, segs) { // segs: [[이름, 수, 색 클래스]]
    const sum = segs.reduce((a, x) => a + x[1], 0) || 1;
    return `<div class="dsh-goal"><div class="dsh-gh"><b>${esc(label)}</b><span>${total}건</span></div>
      <div class="dsh-bar" role="img" aria-label="${esc(label)} ${segs.map((x) => x[0] + " " + x[1]).join(", ")}">${segs.filter((x) => x[1]).map((x) => `<i class="${x[2]}" style="flex:${x[1]}"></i>`).join("") || `<i class="none" style="flex:1"></i>`}</div>
      <div class="dsh-lg">${segs.map((x) => `<span class="${x[2]}">${esc(x[0])} <b>${x[1]}</b></span>`).join("")}</div></div>`;
  }
  // v442(2026-11-19 사용자 지시): 관리자에게만 보이는 '팀관리' 메뉴(세일즈 오른쪽). 내용은 아직 정해지지 않아 빈 화면으로 시작
  document.addEventListener("click", (ev) => { if (ev.target.closest && ev.target.closest("[data-team-org]")) openOrgFromTeam(); });
  function openTeam() {
    if (!isAdmin()) return;
    show("teamView");
    $("teamView").innerHTML = `<div class="card"><div class="team-head"><h3 class="sec-title">팀관리</h3><div class="team-btns"><button type="button" class="ev-edit" data-team-org>조직도</button><button type="button" class="ev-edit" data-team-info>팀원정보</button><button type="button" class="ev-edit" data-team-cand>후보자정보</button><button type="button" class="ev-edit" data-team-intro>도입현황</button></div></div></div>`;
  }
  // v465(2026-11-19 사용자 지시): 팀관리 > [팀원정보]. 관리자만. 관리자 드라이브의 Claude/팀 폴더에 있는 팀원 파일을 그때그때 읽어 보여 주고(기기에 내용을 남기지 않음), 기본 항목을 고쳐 같은 파일에 저장
  //   위 드롭다운 = 팀원(직급순 → 사번순), 처음에는 로그인한 본인이 선택됨
  const TM_DIR = "Claude/팀", TM_FILE = "팀원(초안).json";
  const TM_RANK = { ESL: 0, SL: 1, FC: 2 };
  // v468(2026-11-19 사용자 지시): 이름·사번·직급은 위 드롭다운에 있어 줄에서 뺌, 단계·신분 삭제, 맨 위 주민번호, 연락처 앞에 통신사, 메모 밑 구분선 뒤 사이트 계정(비번)
  //   줄 = [라벨, [[파일의 칸 이름, 입력 종류, 입력칸 안내]…]]. 종류 secret = 가려서 보여 주고 [보기]·[복사]
  const TM_ROWS = [["주민번호", [["주민번호", "secret", "000000-0000000"]]], ["연락처", [["통신사", "text", "통신사"], ["연락처", "tel", "010-0000-0000"]]], ["이메일", [["이메일", "email", ""]]], ["주소", [["주소", "text", ""]]], ["도입자", [["도입자", "text", ""]]], ["도입일", [["도입일", "date", ""]]], ["소개자", [["소개자", "text", "도입자에게 소개한 사람"]]], ["입사일", [["입사일", "date", ""]]], ["메모", [["메모", "area", ""]]]];
  const TM_LOCK = new Set(); // v492(2026-11-19 사용자 지시): 관리자만 쓰는 화면이라 주민번호·입사일도 수정 가능
  // v485(2026-11-19 사용자 지시): 도입자는 팀원 중에서 고르고 '도입자사번'으로 연결(화면에는 이름). 사번이 없는 사람(팀원이 아닌 사람)은 바로 아래 '소개자' 칸에 글자로 적음
  const introName = (p, list) => { const t = (list || []).find((x) => String(x.사번) === String(p.도입자사번 || "")); return t ? String(t.이름 || "") : String(p.도입자 || ""); };
  function introSel(attr, p, list) {
    const cur = String(p.도입자사번 || ""), inTeam = !!cur && (list || []).some((x) => String(x.사번) === cur), ext = inTeam ? "" : String(p.도입자 || "");
    const os = (list || []).filter((x) => !p.사번 || String(x.사번) !== String(p.사번)).slice().sort((a, b) => String(a.이름 || "").localeCompare(String(b.이름 || ""), "ko"));
    return `<select data-${attr}="도입자사번" aria-label="도입자"><option value=""${inTeam || ext ? "" : " selected"}>없음</option>${ext ? `<option value="__keep__" selected>${esc(ext)} (팀 밖)</option>` : ""}${os.map((x) => `<option value="${esc(String(x.사번))}"${inTeam && String(x.사번) === cur ? " selected" : ""}>${esc(x.이름 || "")}</option>`).join("")}</select>`;
  }
  function introFix(vals, list) { // 고른 사번 → 도입자 이름도 함께. '팀 밖' 그대로면 건드리지 않음
    if (!("도입자사번" in vals)) return vals;
    if (vals.도입자사번 === "__keep__") { delete vals.도입자사번; return vals; }
    const t = (list || []).find((x) => String(x.사번) === vals.도입자사번);
    vals.도입자 = t ? String(t.이름 || "") : "";
    return vals;
  }
  const TM_ACCTS = [["Sol FC", [["SolFC비번", "text", "비번"]]], ["신한이지", [["신한이지아이디", "text", "아이디"], ["신한이지비번", "text", "비번"]]], ["지식인", [["지식인비번", "text", "비번"]]], ["보험연수원", [["보험연수원아이디", "text", "아이디"], ["보험연수원비번", "text", "비번"]]]];
  let tmDoc = null, tmDir = null, tmSel = "", tmEdit = false;
  const tmRank = (p) => { const r = TM_RANK[String(p.직급 || "").trim().toUpperCase()]; return r == null ? 9 : r; };
  const tmPeople = () => ((tmDoc && tmDoc.사람) || []).slice().sort((a, b) => tmRank(a) - tmRank(b) || String(a.사번 || "").localeCompare(String(b.사번 || "")) || String(a.이름 || "").localeCompare(String(b.이름 || ""), "ko"));
  // ===== v578(2026-11-25 사용자 지시 "도입현황") 팀관리 > [도입현황] =====
  //   도입 1건 = 팀원 파일에서 도입자(도입자사번·도입자)가 있는 팀원 1명. 날짜 = 도입일(없으면 입사일). 퇴직해도 남음
  const IB = { y: 0, recent: false, ppl: [] };
  const ibDay = (p) => String(p.도입일 || p.입사일 || "").replace(/\./g, "-").slice(0, 10);
  async function openIntroBoard() {
    if (!isAdmin()) return;
    show("teamView");
    const box = $("teamView");
    box.innerHTML = `<p class="hint pad">도입현황을 불러오는 중…</p>`;
    try {
      const dir = await stFindPath(TM_DIR), doc = dir ? await stReadJsonIn(dir, TM_FILE) : null;
      IB.ppl = (doc && Array.isArray(doc.사람) && doc.사람) || [];
      if (!IB.y) IB.y = new Date().getFullYear();
      renderIntroBoard();
    } catch (e) { box.innerHTML = ""; handleError(e); }
  }
  function renderIntroBoard() {
    const box = $("teamView"), ppl = IB.ppl, now = new Date(), z = (x) => String(x).padStart(2, "0");
    const recs = ppl.filter((p) => (p.도입자사번 || p.도입자) && /^\d{4}-\d{2}/.test(ibDay(p))).map((p) => {
      const t = ppl.find((x) => String(x.사번) === String(p.도입자사번 || ""));
      return { by: t ? String(t.이름 || "") : String(p.도입자 || "") + (p.도입자사번 ? "" : " (팀 밖)"), who: String(p.이름 || ""), day: ibDay(p), out: !!String(p.퇴직일 || "").trim() };
    });
    const ymOf = (d) => d.slice(0, 7);
    let months;
    if (IB.recent) { months = [0, 1, 2].map((k) => { const d = new Date(now.getFullYear(), now.getMonth() - k, 1); return d.getFullYear() + "-" + z(d.getMonth() + 1); }); }
    else { const last = IB.y === now.getFullYear() ? now.getMonth() + 1 : 12; months = []; for (let m = last; m >= 1; m--) months.push(IB.y + "-" + z(m)); }
    const mset = new Set(months), sel = recs.filter((r) => mset.has(ymOf(r.day)));
    const cnt = new Map(); sel.forEach((r) => cnt.set(r.by, (cnt.get(r.by) || 0) + 1));
    const rank = [...cnt].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0], "ko")), mx = rank.length ? rank[0][1] : 1;
    const years = [...new Set([now.getFullYear(), ...recs.map((r) => +r.day.slice(0, 4))])].sort((a, b) => b - a);
    const ymL = (m) => m.slice(0, 4) + "년 " + +m.slice(5) + "월", period = IB.recent ? `최근 3개월 (${ymL(months[2])}~${ymL(months[0])})` : `${IB.y}년 1월~${IB.y === now.getFullYear() ? now.getMonth() + 1 : 12}월`;
    // 연도 보기에서는 도입이 있는 달과, 올해는 이번 달까지 모두(없는 달은 '도입 없음'), 지난해는 도입 있는 달만
    const showM = IB.recent || IB.y === now.getFullYear() ? months : months.filter((m) => sel.some((r) => ymOf(r.day) === m));
    box.innerHTML = `<div class="card ib"><h3 class="sec-title">도입현황</h3>
      <div class="tm-bar ib-bar"><button type="button" class="ev-edit org-back-btn" data-ib-back>‹ 팀관리</button><select data-ib-year aria-label="연도"${IB.recent ? " disabled" : ""}>${years.map((y) => `<option value="${y}"${y === IB.y ? " selected" : ""}>${y}년</option>`).join("")}</select><button type="button" class="ib-chip${IB.recent ? " on" : ""}" data-ib-recent aria-pressed="${IB.recent}">최근 3개월</button></div>
      <div class="ib-sum"><b>${sel.length}</b><span>명 도입 · ${esc(period)}</span></div>
      ${rank.length ? `<div class="ib-rank">${rank.map(([n, c]) => `<span>${esc(n)}</span><span class="bar"><i style="width:${Math.max(6, Math.round((c / mx) * 100))}%"></i></span><span class="n">${c}명</span>`).join("")}</div>` : `<p class="hint">이 기간에 도입이 없습니다.</p>`}
      <div class="ib-ms">${showM.map((m) => { const rs = sel.filter((r) => ymOf(r.day) === m).sort((a, b) => a.day.localeCompare(b.day)); return `<div class="ib-m"><h4>${IB.recent ? ymL(m) : +m.slice(5) + "월"} <span>${rs.length}명</span></h4>${rs.length ? rs.map((r) => `<div class="ib-r">${esc(r.by)} <span class="ar">→</span> ${esc(r.who)}${r.out ? ' <span class="out">퇴직</span>' : ""}<small>${esc(r.day.slice(5))} 도입</small></div>`).join("") : '<div class="ib-none">도입 없음</div>'}</div>`; }).join("")}</div>
      <p class="hint ib-foot">도입 = 팀원정보에서 도입자가 있는 팀원. 날짜는 도입일(없으면 입사일)입니다. 팀원정보에서 도입자·도입일을 고치면 여기도 바뀝니다.</p></div>`;
  }
  document.addEventListener("click", (ev) => {
    const t = ev.target && ev.target.closest && ev.target.closest("[data-team-intro],[data-ib-back],[data-ib-recent]");
    if (!t) return;
    if (t.hasAttribute("data-team-intro")) { IB.recent = false; IB.y = new Date().getFullYear(); openIntroBoard(); return; }
    if (t.hasAttribute("data-ib-back")) { openTeam(); window.scrollTo(0, 0); return; }
    if (t.hasAttribute("data-ib-recent")) { IB.recent = !IB.recent; renderIntroBoard(); }
  });
  document.addEventListener("change", (ev) => { const t = ev.target; if (t && t.matches && t.matches("[data-ib-year]")) { IB.y = +t.value; renderIntroBoard(); } });
  async function openTeamInfo() {
    if (!isAdmin()) return;
    show("teamView");
    const box = $("teamView");
    box.innerHTML = `<p class="hint pad">팀원 정보를 불러오는 중…</p>`;
    try {
      tmEdit = false;
      tmDir = await stFindPath(TM_DIR);
      tmDoc = tmDir ? await stReadJsonIn(tmDir, TM_FILE) : null;
      if (!tmDoc || !Array.isArray(tmDoc.사람) || !tmDoc.사람.length) { box.innerHTML = `<div class="card"><h3 class="sec-title">팀원정보</h3><div class="tm-bar"><button type="button" class="ev-edit org-back-btn" data-tm-back>‹ 팀관리</button></div><p class="hint">팀원 파일을 찾지 못했습니다(내 드라이브 > Claude > 팀 > ${esc(TM_FILE)}).</p></div>`; return; }
      const ppl = tmPeople(), em = me(), nm = await myFcName();
      const mine = ppl.find((x) => em && String(x.이메일 || "").toLowerCase() === em) || ppl.find((x) => nm && x.이름 === nm) || ppl[0];
      if (!ppl.some((x) => String(x.사번) === tmSel)) tmSel = String(mine.사번);
      renderTeamInfo(); window.scrollTo(0, 0);
    } catch (e) { box.innerHTML = ""; handleError(e); openTeam(); }
  }
  function renderTeamInfo() {
    const box = $("teamView"), ppl = tmPeople(), p = ppl.find((x) => String(x.사번) === tmSel) || ppl[0];
    if (!box || !p) return;
    const opts = ppl.map((x) => `<option value="${esc(String(x.사번))}" ${String(x.사번) === String(p.사번) ? "selected" : ""}>${esc(x.이름)} · ${esc(x.직급 || "-")} · ${esc(String(x.사번))}</option>`).join("");
    const val = (k) => String(p[k] == null ? "" : p[k]);
    const view = ([k, t]) => {
      const v = val(k);
      if (k === "도입자") { const n = introName(p, ppl); return n ? esc(n) : ""; }
      if (!v) return "";
      if (t === "secret") return `<span class="tm-sec"><span class="tm-secv">${esc(fmtRrn(v))}</span><button type="button" class="map-btn copy-btn" data-copy="${esc(fmtRrn(v))}" data-copy-msg="복사했습니다">복사</button></span>`;
      if (k === "주소") return `${esc(v)} ${mapBtn(v)}`;
      if (t === "tel" || PHONE_COLS.includes(k)) return esc(fmtPhoneText(v));
      return esc(v).replace(/\n/g, "<br>");
    };
    const edit = ([k, t, ph]) => {
      const v = val(k);
      if (k === "도입자") return introSel("tmf", p, ppl);
      if (k === "통신사") { const base = ["SKT", "KT", "LG U+", "SKT 알뜰폰", "KT 알뜰폰", "LG U+ 알뜰폰"], os = v && !base.includes(v) ? [v, ...base] : base; return `<select data-tmf="통신사" class="tm-narrow" aria-label="통신사"><option value=""${v ? "" : " selected"}>통신사</option>${os.map((o) => `<option${o === v ? " selected" : ""}>${esc(o)}</option>`).join("")}</select>`; } /* v478: 통신사는 고르기(고객 정보 수정과 같은 목록) */
      if (t === "area") return `<textarea data-tmf="${k}" rows="1">${esc(v)}</textarea>`;
      return `<input data-tmf="${k}" type="${t === "secret" || t === "date" ? "text" : t}"${t === "date" ? ' data-dp="date"' : ""} value="${esc(k === "주민번호" ? fmtRrn(v) : t === "tel" ? fmtPhoneText(v) : v)}" placeholder="${esc(ph || "")}" aria-label="${esc(k)}"${t === "tel" ? ' inputmode="tel"' : ""} autocomplete="off" autocapitalize="off" autocorrect="off" spellcheck="false"${k === "통신사" ? ' class="tm-narrow"' : ""}>`;
    };
    const rowHtml = ([lab, cells]) => {
      const ed = tmEdit && !cells.some(([k]) => TM_LOCK.has(k)); // v471: 주민번호·입사일·도입자는 바뀌지 않는 값이라 수정 중에도 보기 그대로
      const parts = ed ? cells.map(edit) : cells.map(view).filter(Boolean);
      return `<dt>${lab}</dt><dd class="${cells.length > 1 ? "tm-multi" : ""}${tmEdit && !ed ? " tm-lock" : ""}">${parts.length ? parts.join(ed ? "" : '<span class="tm-gap"></span>') : '<span class="tm-none">-</span>'}</dd>`;
    };
    const acts = tmEdit ? `<button type="button" class="ev-edit" data-tm-cancel>취소</button><button type="button" class="ev-edit ev-done" data-tm-save>저장</button>` : `<button type="button" class="ev-edit" data-tm-edit>수정</button>`;
    box.innerHTML = `<div class="card tm-wrap"><h3 class="sec-title">팀원정보</h3>
      <div class="tm-bar"><button type="button" class="ev-edit org-back-btn" data-tm-back>‹ 팀관리</button><select data-tm-sel aria-label="팀원 고르기(직급순, 사번순)" ${tmEdit ? "disabled" : ""}>${opts}</select>${acts}</div>
      <dl class="tm-info${tmEdit ? " editing" : ""}">${TM_ROWS.map(rowHtml).join("")}</dl>
      <hr class="tm-hr">
      <dl class="tm-info tm-acct${tmEdit ? " editing" : ""}">${TM_ACCTS.map(rowHtml).join("")}</dl></div>`;
  }
  document.addEventListener("click", (ev) => { // 가린 값 [보기] ↔ [가리기]
    const b = ev.target && ev.target.closest && ev.target.closest("#teamView [data-tm-show]");
    if (!b) return;
    const sv = b.parentNode.querySelector(".tm-secv"), on = b.textContent === "보기";
    sv.textContent = on ? sv.dataset.v : "••••••"; b.textContent = on ? "가리기" : "보기";
  });
  async function saveTeamInfo(btn) {
    const box = $("teamView"), vals = {};
    box.querySelectorAll("[data-tmf]").forEach((el) => { vals[el.dataset.tmf] = String(el.value || "").trim(); }); introFix(vals, (tmDoc && tmDoc.사람) || []); numVals(vals);
    btn.disabled = true; btn.textContent = "저장 중";
    try {
      const fresh = await stReadJsonIn(tmDir, TM_FILE); // 저장 직전에 최신 파일을 다시 읽어 그 사람의 바뀐 칸만 고침
      const q = fresh && Array.isArray(fresh.사람) ? fresh.사람.find((x) => String(x.사번) === tmSel) : null;
      if (!q) throw new Error("팀원 파일에서 이 사람을 찾지 못했습니다");
      let n = 0;
      Object.keys(vals).forEach((k) => { if (String(q[k] == null ? "" : q[k]) !== vals[k]) { q[k] = vals[k]; n++; } });
      if (n) { const d = new Date(), z = (x) => String(x).padStart(2, "0"); q.수정일시 = `${d.getFullYear()}-${z(d.getMonth() + 1)}-${z(d.getDate())} ${z(d.getHours())}:${z(d.getMinutes())}`; await stWriteJsonIn(tmDir, TM_FILE, fresh); }
      tmDoc = fresh; tmEdit = false; renderTeamInfo(); toast(n ? "저장했습니다" : "바뀐 것이 없습니다");
    } catch (e) { btn.disabled = false; btn.textContent = "저장"; handleError(e); }
  }
  document.addEventListener("click", (ev) => {
    const t = ev.target && ev.target.closest && ev.target.closest("[data-team-info],[data-tm-back],[data-tm-edit],[data-tm-cancel],[data-tm-save]");
    if (!t) return;
    if (t.hasAttribute("data-team-info")) { openTeamInfo(); return; }
    if (t.hasAttribute("data-tm-back")) { if (tmEdit && !confirm("저장하지 않은 수정을 버릴까요?")) return; tmEdit = false; tmDoc = null; openTeam(); window.scrollTo(0, 0); return; }
    if (t.hasAttribute("data-tm-edit")) { tmEdit = true; renderTeamInfo(); return; }
    if (t.hasAttribute("data-tm-cancel")) { tmEdit = false; renderTeamInfo(); return; }
    saveTeamInfo(t);
  });
  // v481(2026-11-19 사용자 지시): 팀관리 > [후보자정보]. 관리자만. 후보자는 Claude/팀 폴더의 후보자 파일에 따로 두고(나중에 팀원정보로 옮길 예정이라 이름·주민번호·통신사·연락처·이메일·주소 칸 이름은 팀원 파일과 같게),
  //   기본 정보 + 짧은 질문 6개 + 상세 자기소개서 3개를 보여 주고 그 자리에서 고쳐 저장
  const CD_FILE = "후보자(초안).json";
  const CD_ROWS = [["이름", [["이름", "text", ""]]], ["사번", [["사번", "text", "나오면 입력"]]], ["입사일", [["입사일", "date", ""]]], ["도입자", [["도입자", "text", ""]]], ["소개자", [["소개자", "text", "도입자에게 소개한 사람"]]], ["영문이름", [["영문이름", "text", ""]]], ["주민번호", [["주민번호", "secret", "000000-0000000"]]], ["연락처", [["통신사", "text", "통신사"], ["연락처", "tel", "010-0000-0000"]]], ["이메일", [["이메일", "email", ""]]], ["주소", [["주소", "text", ""]]], ["고등학교", [["고등학교", "text", "학교명"], ["고등학교졸업", "year", "졸업 연도"]]], ["최종학력", [["최종학력", "pick:고등학교|대학교|대학원", ""]]], ["병역", [["병역", "pick:면제|육군|공군|해군", ""], ["병역시작", "month", "시작"], ["병역끝", "month", "끝"]]], ["결혼여부", [["결혼여부", "check", "기혼"], ["결혼기념일", "md", "기념일"]]], ["자녀여부", [["자녀여부", "text", ""]]], ["종교", [["종교", "text", ""]]], ["취미", [["취미", "text", ""]]], ["특기", [["특기", "text", ""]]]];
  const CD_QS = [["자기소개1", "1. 본인 자신을 30글자 이내로 표현해 보세요."], ["자기소개2", "2. 10년 후의 나의 모습은?"], ["자기소개3", "3. 신한라이프가 당신을 뽑아야 하는 이유는?"], ["자기소개4", "4. 본인의 좌우명은?"], ["자기소개5", "5. 본인의 슬럼프 극복 방법은 무엇인가요?"], ["자기소개6", "6. 현재의 꿈은 무엇인가요?"]];
  const CD_ESSAYS = [["자소서1", "1. 성장 배경과 자신의 가치관(인생관)"], ["자소서2", "2. 본인이 생각하는 본인의 장단점"], ["자소서3", "3. 지원 동기 및 향후 Vision과 꿈"]];
  let cdDoc = null, cdSel = "", cdEdit = false, cdTeam = [];
  const cdKept = () => ((cdDoc && cdDoc.사람) || []).filter((x) => !x.전환일시 && x.담아둔일시).sort((a, b) => String(b.담아둔일시).localeCompare(String(a.담아둔일시))); // v563: 담아둔 후보자(고르기 목록에서 빠짐)
  let cdKeepView = false;
  const cdPeople = () => ((cdDoc && cdDoc.사람) || []).filter((x) => !x.전환일시 && !x.담아둔일시).sort((a, b) => String(a.이름 || "").localeCompare(String(b.이름 || ""), "ko"));
  const cdHead = (mid, acts) => `<div class="cd-titlebar"><h3 class="sec-title">후보자정보</h3>${cdEdit ? "" : `<span class="cd-tr">${cdKept().length ? `<button type="button" class="cd-keepn" data-cd-keeplist>담아둔 ${cdKept().length}</button>` : ""}<button type="button" class="cd-addt" data-cd-add>추가</button></span>`}</div><div class="tm-bar"><button type="button" class="ev-edit org-back-btn" data-cd-back>‹ 팀관리</button>${mid || '<span class="tm-sp" aria-hidden="true"></span>'}${acts || ""}</div>`;
  async function openCandInfo() {
    if (!isAdmin()) return;
    show("teamView");
    const box = $("teamView");
    box.innerHTML = `<p class="hint pad">후보자 정보를 불러오는 중…</p>`;
    try {
      cdEdit = false; cdKeepView = false;
      tmDir = await stFindPath(TM_DIR);
      cdDoc = tmDir ? await stReadJsonIn(tmDir, CD_FILE) : null;
      { const tj = tmDir ? await stReadJsonIn(tmDir, TM_FILE).catch(() => null) : null; cdTeam = (tj && Array.isArray(tj.사람) && tj.사람) || []; }
      if (!cdDoc || !Array.isArray(cdDoc.사람) || !cdPeople().length) { box.innerHTML = `<div class="card">${cdHead("")}<p class="hint">팀원으로 전환하지 않은 후보자가 없습니다(내 드라이브 > Claude > 팀 > ${esc(CD_FILE)}).</p></div>`; return; }
      const ppl = cdPeople();
      if (!ppl.some((x) => String(x.id) === cdSel) && !cdKept().some((x) => String(x.id) === cdSel)) cdSel = String(ppl[0].id);
      renderCandInfo(); window.scrollTo(0, 0);
    } catch (e) { box.innerHTML = ""; handleError(e); openTeam(); }
  }
  // v560(2026-11-24 사용자 지시): 고등학교 = 학교명 + 졸업 연도, 최종학력 = 고등학교/대학교/대학원, 병역 = 면제/육군/공군/해군 + 시작·끝 년월, 결혼여부 = 체크 + 기념일(월·일)
  //   예전에 글로 적은 값은 수정 화면에서 알아볼 수 있는 만큼 새 칸으로 옮겨 보여 주고(예: 고등학교 칸의 '1988' → 졸업 연도, '결혼11월18일' → 기혼·11-18), 저장으로 바뀌면 원래 글은 '<칸>_이전'에 남김
  const cdPrefill = (p) => {
    const g = (k) => String(p[k] == null ? "" : p[k]).trim(), o = {};
    const hs = g("고등학교"); if (/^\d{4}$/.test(hs)) { o.고등학교 = ""; if (!g("고등학교졸업")) o.고등학교졸업 = hs; }
    const mr = g("결혼여부"); if (mr && mr !== "기혼") { o.결혼여부 = /미혼/.test(mr) ? "" : "기혼"; const m = /(\d{1,2})\s*월\s*(\d{1,2})\s*일/.exec(mr); if (m && !g("결혼기념일")) o.결혼기념일 = String(m[1]).padStart(2, "0") + "-" + String(m[2]).padStart(2, "0"); }
    return o;
  };
  const cdYm = (v) => String(v || "").replace(/^(\d{4})-(\d{2}).*/, "$1.$2");
  // ---------- v562(2026-11-24 사용자 "이대로 코딩"): 후보자 [추가](메모·이미지 → Gemini 분석 → 채운 칸 확인 → 완료), [수정]의 '자료로 채우기'(빈 칸만 / 기존 정보도 바꾸기), 맨 아래 '자료' 박스 ----------
  //   올린 이미지·PDF는 Claude/팀/후보자자료/<후보자ID - 이름> 폴더에 '날짜 - 추가|수정 - 원래이름'으로 저장, 후보자 기록의 '자료폴더'에 폴더 ID. 메모는 '추가메모'에 날짜와 함께 쌓음
  const CD_DOCS = "후보자자료", CD_CARRIERS = ["SKT", "KT", "LG U+", "SKT 알뜰폰", "KT 알뜰폰", "LG U+ 알뜰폰"];
  const cdSrcFiles = { add: [], edit: [] }; // 아직 올리지 않은 파일(이 화면에서만)
  let cdAdd = null; // { step, memo, vals, filled:Set }
  let cdFillMode = "empty";
  const cdYmd = () => { const d = new Date(); return d.getFullYear() + String(d.getMonth() + 1).padStart(2, "0") + String(d.getDate()).padStart(2, "0"); };
  function cdSrcHtml(id) {
    return `<div class="cd-src" data-cdsrc="${id}"><label class="cd-lab" for="cdMemo_${id}">메모</label><textarea class="cd-memo" id="cdMemo_${id}" rows="3" placeholder="이름·연락처·학력·병역 등 아는 것을 적어 주세요">${esc(id === "add" && cdAdd ? cdAdd.memo || "" : "")}</textarea>
      <span class="cd-lab">이미지 (이력서·자기소개서·명함 사진)</span>
      <div class="cd-drop" tabindex="0" role="button" data-cddrop="${id}">이미지를 여기로 <b>끌어다 놓거나 눌러서 고르기</b><small>여러 장 가능 · JPG·PNG·PDF</small><input type="file" accept="image/*,application/pdf" multiple hidden data-cdfile="${id}"></div>
      <div class="cd-thumbs" data-cdthumbs="${id}">${cdThumbsHtml(id)}</div></div>`;
  }
  function cdThumbsHtml(id) {
    return cdSrcFiles[id].map((f, i) => { if (!f._url && /^image\//.test(f.type) && !/heic/i.test(f.type + f.name)) f._url = URL.createObjectURL(f); return `<span class="cd-th" title="${esc(f.name)}">${f._url ? `<img src="${f._url}" alt="">` : `<b>${/pdf/i.test(f.type + f.name) ? "PDF" : "파일"}</b>`}<i>${esc(f.name.replace(/\.[^.]+$/, "").slice(0, 8))}</i><button type="button" data-cdthx="${id}|${i}" aria-label="${esc(f.name)} 빼기">×</button></span>`; }).join("");
  }
  function cdAddFiles(id, list) {
    const ok = [...(list || [])].filter((f) => /^image\//.test(f.type) || /pdf/i.test(f.type) || /\.(pdf|jpe?g|png|webp|heic)$/i.test(f.name));
    if (!ok.length) return toast("이미지나 PDF만 넣을 수 있습니다.");
    cdSrcFiles[id].push(...ok.slice(0, 10 - cdSrcFiles[id].length));
    const t = document.querySelector(`[data-cdthumbs="${id}"]`); if (t) t.innerHTML = cdThumbsHtml(id);
  }
  function cdClearFiles(id) { cdSrcFiles[id].forEach((f) => { if (f._url) URL.revokeObjectURL(f._url); }); cdSrcFiles[id] = []; }
  document.addEventListener("dragover", (ev) => { const d = ev.target.closest && ev.target.closest("[data-cddrop]"); if (!d) return; ev.preventDefault(); d.classList.add("over"); });
  document.addEventListener("dragleave", (ev) => { const d = ev.target.closest && ev.target.closest("[data-cddrop]"); if (d) d.classList.remove("over"); });
  document.addEventListener("drop", (ev) => { const d = ev.target.closest && ev.target.closest("[data-cddrop]"); if (!d) return; ev.preventDefault(); d.classList.remove("over"); cdAddFiles(d.dataset.cddrop, ev.dataTransfer && ev.dataTransfer.files); });
  document.addEventListener("change", (ev) => { const i = ev.target.closest && ev.target.closest("[data-cdfile]"); if (!i) return; cdAddFiles(i.dataset.cdfile, i.files); i.value = ""; });
  document.addEventListener("keydown", (ev) => { const d = ev.target.closest && ev.target.closest("[data-cddrop]"); if (d && (ev.key === "Enter" || ev.key === " ")) { ev.preventDefault(); d.querySelector("[data-cdfile]").click(); } });

  // 분석: 메모·이미지 → Gemini → 후보자 칸 값(형식 맞춘 것만)
  function cdNorm(k, v) {
    v = String(v == null ? "" : v).trim(); if (!v) return "";
    const one = (list) => (list.includes(v) ? v : "");
    if (k === "입사일") return /^\d{4}-\d{2}-\d{2}$/.test(v) ? v : "";
    if (k === "고등학교졸업") return /^\d{4}$/.test(v) ? v : "";
    if (k === "최종학력") return one(["고등학교", "대학교", "대학원"]);
    if (k === "병역") return one(["면제", "육군", "공군", "해군"]);
    if (k === "병역시작" || k === "병역끝") return /^\d{4}-\d{2}$/.test(v) ? v : "";
    if (k === "결혼여부") return v === "기혼" ? "기혼" : "";
    if (k === "결혼기념일") return /^\d{2}-\d{2}$/.test(v) ? v : "";
    if (k === "통신사") return one(CD_CARRIERS);
    if (k === "주민번호") return /^\d{6}-?\d{1,7}$/.test(v) ? onlyDigits(v) : "";
    if (k === "연락처") { const d = v.replace(/\D/g, ""); return /^01\d{8,9}$/.test(d) ? d.replace(/^(\d{3})(\d{3,4})(\d{4})$/, "$1-$2-$3") : ""; }
    if (k === "이메일") return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v) ? v : "";
    return v.slice(0, /^자/.test(k) ? 3000 : 200);
  }
  const cdAiKeys = () => [...CD_ROWS.flatMap(([, cells]) => cells.map(([k]) => k)).filter((k) => k !== "사번" && k !== "도입자"), ...CD_QS.map(([k]) => k), ...CD_ESSAYS.map(([k]) => k)];
  async function cdAnalyze(memo, files, onStep) {
    const cfg = await aiGeminiConfig(); if (!cfg) return { noKey: true };
    const parts = [];
    for (const f of files.slice(0, 10)) {
      try {
        const buf = await f.arrayBuffer();
        if (/pdf/i.test(f.type) || /\.pdf$/i.test(f.name)) { if (buf.byteLength <= 15 * 1024 * 1024) parts.push({ inline_data: { mime_type: "application/pdf", data: await b64(new Blob([buf])) } }); }
        else if (!/heic/i.test(f.type + f.name)) { const bm = await createImageBitmap(new Blob([buf], { type: f.type || "image/jpeg" })); parts.push({ inline_data: { mime_type: "image/jpeg", data: await b64(await cvJpeg(bm, bm.width, bm.height, 1800)) } }); }
      } catch (e) { console.warn("Gemini에 못 보낸 후보자 자료", f.name, e); }
    }
    if (!parts.length && !String(memo || "").trim()) return null;
    const keys = cdAiKeys();
    const prompt = "너는 보험회사 FC 후보자(입사 지원자)의 정보를 정리하는 도우미다. 아래 메모와 첨부한 이력서·자기소개서·명함 이미지에서 후보자 본인의 정보만 읽어 JSON 객체 하나로 답해라.\n" +
      "키(모두 문자열): " + keys.join(", ") + "\n" +
      "형식: 입사일 YYYY-MM-DD, 고등학교 = 고등학교 이름, 고등학교졸업 = 졸업 연도 YYYY, 최종학력 = 고등학교|대학교|대학원 중 하나, 병역 = 면제|육군|공군|해군 중 하나, 병역시작·병역끝 YYYY-MM, 결혼여부 = 기혼이면 \"기혼\" 아니면 빈 문자열, 결혼기념일 MM-DD, 통신사 = " + CD_CARRIERS.join("|") + " 중 하나, 주민번호 000000-0000000, 연락처 010-0000-0000, 영문이름 = 여권식 영문 이름. " +
      "자기소개1~6은 차례로 [" + CD_QS.map(([, q]) => q).join(" / ") + "]의 답, 자소서1~3은 [" + CD_ESSAYS.map(([, q]) => q).join(" / ") + "]에 해당하는 글(원문 그대로, 길면 그대로 둠).\n" +
      "자료에서 확인되지 않는 값은 지어내지 말고 빈 문자열로 둬라.\n\n메모:\n" + String(memo || "(없음)");
    try {
      const gemReq = { contents: [{ role: "user", parts: [{ text: prompt }, ...parts] }], generationConfig: { responseMimeType: "application/json", temperature: 0 } };
      let res = await gemCall(cfg, gemReq, onStep); // v565: 모델 기억·없는 모델 건너뛰기·20초 제한·생각 끄기
      if (!res || !res.ok) { console.warn("Gemini 후보자 분석 실패", res && res.status); return null; }
      const d = await res.json(), c0 = (d.candidates || [])[0] || {}, tx = ((c0.content || {}).parts || []).filter((x) => !x.thought).map((x) => x.text || "").join("");
      const j = JSON.parse(String(tx).replace(/^```json\s*|\s*```$/g, "") || "{}"), out = {};
      keys.forEach((k) => { const v = cdNorm(k, j[k]); if (v) out[k] = v; });
      return out;
    } catch (e) { console.warn("Gemini 후보자 분석 실패", e); return null; }
  }
  // 분석 결과를 화면의 칸에 넣기. mode = empty(빈 칸만) | replace(기존도 바꿈) → { filled, replaced }
  function cdApplyFill(box, res, mode) {
    let filled = 0, replaced = 0;
    box.querySelectorAll(".cd-old").forEach((x) => x.remove());
    Object.entries(res || {}).forEach(([k, v]) => {
      const el = box.querySelector(`[data-cdf="${k}"]`); if (!el) return;
      const cur = el.type === "checkbox" ? (el.checked ? "기혼" : "") : String(el.value || "").trim();
      if (cur === v || (cur && mode !== "replace")) return;
      if (el.type === "checkbox") el.checked = v === "기혼";
      else if (el.tagName === "SELECT") { if (![...el.options].some((o) => (o.value || o.textContent) === v)) el.insertAdjacentHTML("beforeend", `<option>${esc(v)}</option>`); el.value = v; }
      else el.value = v;
      const show = el._dpFace || (el.type === "checkbox" ? el.closest(".cd-chk") : el);
      if (show) show.classList.add("cd-ai");
      if (cur) { replaced++; const host = el.closest("dd") || el.parentElement; host.insertAdjacentHTML("beforeend", `<span class="cd-old">이전: ${esc(cur)}</span>`); } else filled++;
    });
    return { filled, replaced };
  }
  // 자료 폴더: Claude/팀/후보자자료/<ID - 이름>
  async function cdDocsDir(p, create) {
    if (!GD) return null;
    if (p.자료폴더) return { id: p.자료폴더 };
    let par = cdDocsParent || (await stFindPath(TM_DIR + "/" + CD_DOCS));
    if (par) cdDocsParent = par;
    if (!par) { if (!create) return null; const t = tmDir || (await stFindPath(TM_DIR)); par = { id: (await gdFetch(GDAPI + "/files?fields=id&supportsAllDrives=true", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ name: CD_DOCS, mimeType: GD_FOLDER, parents: [t.id] }) })).id }; cdDocsParent = par; }
    const nm = String(p.id) + " - " + String(p.이름 || "이름 없음").replace(/[\\/:*?"<>|]/g, "");
    const ex = await gdQuery("'" + qv(par.id) + "' in parents and name='" + qv(nm) + "' and mimeType='" + GD_FOLDER + "' and trashed=false");
    if (ex.length) return { id: ex[0].id };
    if (!create) return null;
    return { id: (await gdFetch(GDAPI + "/files?fields=id&supportsAllDrives=true", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ name: nm, mimeType: GD_FOLDER, parents: [par.id] }) })).id };
  }
  async function cdShrink(f) { // v565: 사진은 긴 변 2000px JPEG로 줄여 올림(PDF·HEIC 등은 원본 그대로)
    if (!/^image\/(jpeg|png|webp|bmp)$/i.test(f.type || "")) return { blob: f, name: f.name, type: f.type || "application/octet-stream" };
    try {
      const bm = await createImageBitmap(f);
      if (Math.max(bm.width, bm.height) <= 2000 && f.size <= 1.5 * 1024 * 1024) return { blob: f, name: f.name, type: f.type };
      const b = await cvJpeg(bm, bm.width, bm.height, 2000);
      return { blob: b, name: f.name.replace(/\.[^.]+$/, "") + ".jpg", type: "image/jpeg" };
    } catch (e) { return { blob: f, name: f.name, type: f.type || "application/octet-stream" }; }
  }
  let cdDocsParent = null; // v565: 후보자자료 폴더 ID 기억
  async function cdUpload(p, files, how) { // → 폴더 ID (여러 장 동시에)
    const dir = await cdDocsDir(p, true), day = cdYmd();
    await Promise.all(files.map(async (f) => { const x = await cdShrink(f); return stPutIn(dir, `${day} - ${how} - ${x.name.replace(/[\\/:*?"<>|]/g, "_")}`, x.blob, x.type); }));
    return dir.id;
  }
  // 자료 박스(보기 화면 맨 아래)
  let cdDocsList = [];
  async function cdLoadDocs(p) {
    const box = $("cdDocsList"), n = $("cdDocsN"); if (!box) return;
    try {
      const dir = await cdDocsDir(p, false);
      cdDocsList = dir ? (await stList(dir.id)).filter((f) => f.file).sort((a, b) => String(b.name).localeCompare(String(a.name), "ko", { numeric: true })) : [];
      if (String(p.id) !== cdSel || !$("cdDocsList")) return;
      n.textContent = cdDocsList.length + "개";
      const size = (b) => (b >= 1048576 ? (b / 1048576).toFixed(1) + "MB" : Math.max(1, Math.round(b / 1024)) + "KB");
      box.innerHTML = cdDocsList.length ? cdDocsList.map((f, i) => { const m = /^(\d{4})(\d{2})(\d{2}) - (추가|수정|자료) - (.+)$/.exec(f.name), ty = /pdf/i.test(f.name) ? "PDF" : /\.(jpe?g|png|gif|webp|heic|bmp)$/i.test(f.name) ? "IMG" : "FILE";
        return `<div class="cd-doc"><span class="cd-ty">${ty}</span><button type="button" class="cd-dn" data-cddoc="${i}"><b>${esc(m ? m[5] : f.name)}</b><small>${m ? m[2] + "-" + m[3] : esc(String(f.lastModifiedDateTime || "").slice(5, 10))} · ${size(f.size || 0)}${m ? " · " + (m[4] === "자료" ? "자료에서 올림" : m[4] + "할 때") : ""}</small></button><button type="button" class="cd-dd" data-cddocdel="${i}">삭제</button></div>`; }).join("") : `<p class="hint">첨부한 파일이 없습니다. [올리기]를 누르거나 파일을 끌어다 놓으세요.</p>`;
    } catch (e) { box.innerHTML = `<p class="hint">자료를 불러오지 못했습니다.</p>`; console.warn(e); }
  }
  // v566(2026-11-24 사용자 지시 "업로드 버튼 만들어줘 끌어놓기도"): 자료 박스에서 바로 올리기·끌어다 놓기 (사진은 줄여서, PDF 등은 원본, 여러 개 동시에)
  async function cdDocsUpload(list) {
    const files = [...(list || [])]; if (!files.length) return;
    const p = cdPeople().concat(cdKept()).find((x) => String(x.id) === cdSel); if (!p) return;
    const n = $("cdDocsN"), btn = document.querySelector("[data-cd-up]");
    if (btn) { btn.disabled = true; btn.textContent = "올리는 중…"; }
    if (n) n.textContent = `${files.length}개 올리는 중…`;
    try {
      const fid = await cdUpload(p, files, "자료");
      if (p.자료폴더 !== fid) { // 처음 올린 후보자면 후보자 기록에 폴더 ID를 남김
        const fresh = await stReadJsonIn(tmDir, CD_FILE), q = fresh && Array.isArray(fresh.사람) ? fresh.사람.find((x) => String(x.id) === String(p.id)) : null;
        if (q) { q.자료폴더 = fid; await stWriteJsonIn(tmDir, CD_FILE, fresh); cdDoc = fresh; }
      }
      toast(`${files.length}개를 올렸습니다`);
      cdLoadDocs(cdPeople().concat(cdKept()).find((x) => String(x.id) === String(p.id)) || p);
    } catch (e) { handleError(e); if (n) n.textContent = ""; }
    finally { const b = document.querySelector("[data-cd-up]"); if (b) { b.disabled = false; b.textContent = "올리기"; } }
  }
  document.addEventListener("click", (ev) => { const b = ev.target.closest && ev.target.closest("[data-cd-up]"); if (b) { const i = b.parentElement.querySelector("[data-cd-upfile]"); if (i) i.click(); } });
  document.addEventListener("change", (ev) => { const i = ev.target.closest && ev.target.closest("[data-cd-upfile]"); if (!i) return; cdDocsUpload(i.files); i.value = ""; });
  document.addEventListener("dragover", (ev) => { const d = ev.target.closest && ev.target.closest("[data-cd-docdrop]"); if (!d) return; ev.preventDefault(); d.classList.add("over"); });
  document.addEventListener("dragleave", (ev) => { const d = ev.target.closest && ev.target.closest("[data-cd-docdrop]"); if (d && !d.contains(ev.relatedTarget)) d.classList.remove("over"); });
  document.addEventListener("drop", (ev) => { const d = ev.target.closest && ev.target.closest("[data-cd-docdrop]"); if (!d) return; ev.preventDefault(); d.classList.remove("over"); cdDocsUpload(ev.dataTransfer && ev.dataTransfer.files); });
  // [추가] 창
  function cdAddOpen() { cdClearFiles("add"); cdAdd = { step: 1, memo: "", vals: {}, filled: new Set() }; cdAddRender(); }
  function cdAddClose() { const b = $("cdAddSheet"); if (b) b.remove(); if (window.DatePick) DatePick.close(); cdClearFiles("add"); cdAdd = null; }
  function cdAddRender(busy) {
    let box = $("cdAddSheet");
    if (!box) { box = document.createElement("div"); box.id = "cdAddSheet"; box.className = "modal ntf-shw"; document.body.appendChild(box); }
    if (cdAdd.step === 1) {
      box.innerHTML = `<div class="ntf-sh cd-add" role="dialog" aria-modal="true" aria-label="후보자 추가"><h3>후보자 추가<small class="cd-step">1/2 자료 넣기</small><button type="button" data-cdadd-x>취소</button></h3>
        ${cdSrcHtml("add")}
        <p class="hint cd-note">[분석해서 채우기]를 누르면 메모와 이미지를 구글 Gemini로 보내 후보자 정보 칸을 채웁니다. 올린 이미지는 Claude/팀/후보자자료 폴더에도 저장합니다.</p>
        <button type="button" class="ntf-save" data-cdadd-go ${busy ? "disabled" : ""}>${busy ? esc(busy) : "분석해서 채우기"}</button>
        <button type="button" class="cd-skip" data-cdadd-skip ${busy ? "disabled" : ""}>건너뛰고 직접 입력</button></div>`;
      return;
    }
    const v = cdAdd.vals, val = (k) => String(v[k] == null ? "" : v[k]), p0 = { ...v };
    const rows = CD_ROWS.map(([lab, cells]) => `<dt>${lab}</dt><dd class="${cells.length > 1 ? "tm-multi" : ""}">${cells.map((c) => cdCellEdit(c, val, p0)).join("")}</dd>`).join("");
    const qa = ([k, q]) => `<div class="cd-q">${esc(q)}</div><textarea class="cd-a${cdAdd.filled.has(k) ? " cd-ai" : ""}" data-cdf="${k}" rows="1" aria-label="${esc(q)}">${esc(val(k))}</textarea>`;
    const qn = [...CD_QS, ...CD_ESSAYS].filter(([k]) => cdAdd.filled.has(k)).length;
    box.innerHTML = `<div class="ntf-sh cd-add" role="dialog" aria-modal="true" aria-label="후보자 추가"><h3>후보자 추가<small class="cd-step">2/2 정보 확인</small><button type="button" data-cdadd-back>‹ 자료</button></h3>
      ${cdAdd.filled.size ? `<div class="cd-sum"><span class="cd-sw"></span>자동으로 채운 칸 ${cdAdd.filled.size}개 · 확인하고 빈 칸을 채워 주세요</div>` : `<div class="cd-sum">${cdAdd.note ? esc(cdAdd.note) : "빈 칸을 채워 주세요. 이름은 꼭 넣어야 합니다."}</div>`}
      <dl class="tm-info cd-info editing cd-add-form">${rows}</dl>
      <details class="cd-more"${qn ? " open" : ""}><summary>짧은 질문 6 · 상세 자기소개서 3${qn ? ` <span>${qn}개 채움</span>` : ""}</summary><h4 class="cd-h">짧은 질문</h4><div class="cd-qa editing">${CD_QS.map(qa).join("")}</div><h4 class="cd-h">상세 자기소개서</h4><div class="cd-qa editing">${CD_ESSAYS.map(qa).join("")}</div></details>
      <button type="button" class="ntf-save" data-cdadd-done>완료</button></div>`;
    box.querySelectorAll("[data-cdf]").forEach((el) => { if (cdAdd.filled.has(el.dataset.cdf)) { const s = el._dpFace || (el.type === "checkbox" ? el.closest(".cd-chk") : el); setTimeout(() => { const s2 = el._dpFace || s; if (s2) s2.classList.add("cd-ai"); }, 0); } });
  }
  const numVals = (vals) => { if (vals["주민번호"]) vals["주민번호"] = onlyDigits(vals["주민번호"]); PHONE_COLS.forEach((k) => { if (vals[k]) vals[k] = phoneStore(vals[k]); }); return vals; }; // v576: 저장은 숫자만
  const cdFormVals = (box) => { const vals = {}; box.querySelectorAll("[data-cdf]").forEach((el) => { vals[el.dataset.cdf] = el.type === "checkbox" ? (el.checked ? el.dataset.on || "예" : "") : String(el.value || "").trim(); }); introFix(vals, cdTeam); return numVals(vals); };
  // v565: 분석 진행 표시 — 버튼 글자에 단계·경과 초("읽는 중… 12초", 다른 모델로 바꾸면 "다른 모델로 다시 시도 중… 25초")
  function cdProgress(btn, n) {
    const t0 = Date.now(); let stage = "읽는 중";
    const draw = () => { const b = typeof btn === "function" ? btn() : btn; if (b) b.textContent = (stage === "읽는 중" ? `메모와 이미지 ${n}장을 읽는 중… ` : "다른 모델로 다시 시도 중… ") + Math.round((Date.now() - t0) / 1000) + "초"; };
    draw(); const tm = setInterval(draw, 1000);
    return { step: (st) => { stage = st === "retry" ? "다시" : "읽는 중"; draw(); }, stop: () => clearInterval(tm) };
  }
  async function cdAddGo(skip) {
    const memoEl = document.querySelector("#cdAddSheet .cd-memo"); cdAdd.memo = memoEl ? memoEl.value : cdAdd.memo;
    if (!skip) {
      if (!cdAdd.memo.trim() && !cdSrcFiles.add.length) return toast("메모나 이미지를 넣어 주세요.");
      cdAddRender(`메모와 이미지 ${cdSrcFiles.add.length}장을 읽는 중…`);
      const pg = cdProgress(() => document.querySelector("[data-cdadd-go]"), cdSrcFiles.add.length);
      const res = await cdAnalyze(cdAdd.memo, cdSrcFiles.add, (st) => pg.step(st)); pg.stop();
      if (!cdAdd) return;
      if (res && res.noKey) { cdAdd.note = "Gemini 키가 없어 자동으로 채우지 못했습니다(말만해요 설정에서 키를 넣으면 됩니다). 직접 채워 주세요."; }
      else if (!res) cdAdd.note = "자료를 읽지 못했습니다. 직접 채워 주세요.";
      else { cdAdd.vals = { ...cdAdd.vals, ...res }; cdAdd.filled = new Set(Object.keys(res)); cdAdd.note = ""; }
    }
    cdAdd.step = 2; cdAddRender();
  }
  async function cdAddDone(btn) {
    const box = $("cdAddSheet"), vals = cdFormVals(box);
    if (!vals.이름) return toast("이름을 넣어 주세요.");
    btn.disabled = true; btn.textContent = "저장 중";
    try {
      if (!tmDir) tmDir = await stFindPath(TM_DIR);
      const fresh = (await stReadJsonIn(tmDir, CD_FILE)) || { 사람: [] }; if (!Array.isArray(fresh.사람)) fresh.사람 = [];
      const day = cdYmd(), pre = "C" + day + "-", nn = fresh.사람.map((x) => String(x.id || "")).filter((x) => x.startsWith(pre)).map((x) => +x.slice(pre.length) || 0);
      const d = new Date(), z = (x) => String(x).padStart(2, "0"), stamp = `${d.getFullYear()}-${z(d.getMonth() + 1)}-${z(d.getDate())} ${z(d.getHours())}:${z(d.getMinutes())}`;
      const rec = { id: pre + String(Math.max(0, ...nn) + 1).padStart(2, "0"), ...Object.fromEntries(Object.entries(vals).filter(([, v]) => v !== "")), 등록일시: stamp };
      if (cdAdd.memo.trim()) rec.추가메모 = stamp.slice(0, 10) + " " + cdAdd.memo.trim();
      if (cdSrcFiles.add.length) rec.자료폴더 = await cdUpload(rec, cdSrcFiles.add, "추가");
      fresh.사람.push(rec);
      await stWriteJsonIn(tmDir, CD_FILE, fresh);
      cdDoc = fresh; cdSel = rec.id; cdEdit = false; cdAddClose(); renderCandInfo(); window.scrollTo(0, 0); toast("후보자를 추가했습니다");
    } catch (e) { btn.disabled = false; btn.textContent = "완료"; handleError(e); }
  }
  async function cdEditFill(btn) {
    const card = document.querySelector("#teamView .tm-wrap"), memoEl = card.querySelector('[data-cdsrc="edit"] .cd-memo'), memo = memoEl ? memoEl.value : "";
    if (!memo.trim() && !cdSrcFiles.edit.length) return toast("메모나 이미지를 넣어 주세요.");
    btn.disabled = true;
    const pg = cdProgress(btn, cdSrcFiles.edit.length);
    const res = await cdAnalyze(memo, cdSrcFiles.edit, (st) => pg.step(st)); pg.stop();
    btn.disabled = false; btn.textContent = "분석해서 채우기";
    const sum = $("cdSum");
    if (res && res.noKey) return toast("Gemini 키가 없어 분석하지 못했습니다(말만해요 설정에서 키를 넣어 주세요).");
    if (!res) return toast("자료를 읽지 못했습니다.");
    const r = cdApplyFill(card, res, cdFillMode);
    if (sum) { sum.hidden = false; sum.innerHTML = `<span class="cd-sw"></span>빈 칸 ${r.filled}개 채움${cdFillMode === "replace" ? ` · 기존 칸 ${r.replaced}개 바꿈` : ""} · [저장]을 눌러야 저장됩니다`; }
  }
  document.addEventListener("click", (ev) => {
    const t = ev.target.closest && ev.target.closest("[data-cd-add],[data-cdadd-x],[data-cdadd-go],[data-cdadd-skip],[data-cdadd-back],[data-cdadd-done],[data-cddrop],[data-cdthx],[data-cdmode],[data-cdfill],[data-cddoc],[data-cddocdel]");
    if (!t) { if (ev.target.id === "cdAddSheet" && confirm("후보자 추가를 그만둘까요?")) cdAddClose(); return; }
    const d = t.dataset;
    if (d.cdAdd != null) return cdAddOpen();
    if (d.cdaddX != null) { if (confirm("후보자 추가를 그만둘까요?")) cdAddClose(); return; }
    if (d.cdaddGo != null) return cdAddGo(false);
    if (d.cdaddSkip != null) return cdAddGo(true);
    if (d.cdaddBack != null) { const box = $("cdAddSheet"); cdAdd.vals = { ...cdAdd.vals, ...cdFormVals(box) }; cdAdd.step = 1; if (window.DatePick) DatePick.close(); return cdAddRender(); }
    if (d.cdaddDone != null) return cdAddDone(t);
    if (d.cdthx) { const [id, i] = d.cdthx.split("|"); const f = cdSrcFiles[id].splice(+i, 1)[0]; if (f && f._url) URL.revokeObjectURL(f._url); const box = document.querySelector(`[data-cdthumbs="${id}"]`); if (box) box.innerHTML = cdThumbsHtml(id); return; }
    if (d.cddrop) { if (!ev.target.closest("input")) t.querySelector("[data-cdfile]").click(); return; }
    if (d.cdmode) { cdFillMode = d.cdmode === "replace" ? "replace" : "empty"; t.parentElement.querySelectorAll("[data-cdmode]").forEach((b) => { b.classList.toggle("on", b === t); b.setAttribute("aria-checked", String(b === t)); }); return; }
    if (d.cdfill != null) return cdEditFill(t);
    if (d.cddoc) { const f = cdDocsList[+d.cddoc]; if (f) docsView({ ...f, mime: (f.file && f.file.mimeType) || "" }); return; }
    if (d.cddocdel) { const f = cdDocsList[+d.cddocdel], p = cdPeople().find((x) => String(x.id) === cdSel); if (!f || !confirm(`'${f.name}' 파일을 지울까요?\n\n드라이브 휴지통으로 가서 30일 안에 되살릴 수 있습니다.`)) return; t.disabled = true; stDel(f.id).then(() => { toast("지웠습니다(드라이브 휴지통)"); if (p) cdLoadDocs(p); }, (e) => { t.disabled = false; handleError(e); }); }
  });
  // ---------- v563(2026-11-24 사용자 지시): 담아두기 — 담아둔 후보자는 고르기 목록에서 빼고, [담아둔 N] 목록에서 검색·보기·꺼내기 ----------
  function renderCdKeep() {
    const box = $("teamView"), list = cdKept();
    const md = (s) => String(s || "").slice(5, 10);
    box.innerHTML = `<div class="card tm-wrap cd-keep"><div class="cd-titlebar"><h3 class="sec-title">담아둔 후보자</h3><span class="cd-tr"><button type="button" class="cd-keepn" data-cd-keepback>‹ 후보자정보</button></span></div>
      <input type="search" class="cd-kq" id="cdKq" placeholder="이름·연락처로 찾기" aria-label="담아둔 후보자 찾기" autocomplete="off">
      <div class="cd-klist">${list.map((x) => `<div class="cd-krow" data-k="${esc(String(x.이름 || "") + " " + String(x.연락처 || "").replace(/\D/g, "") + " " + String(x.연락처 || ""))}"><button type="button" class="cd-kn" data-cd-keepopen="${esc(String(x.id))}">${esc(x.이름 || "(이름 없음)")}</button><small>${esc(md(x.담아둔일시))} 담음</small><button type="button" class="cd-kout" data-cd-keepout="${esc(String(x.id))}">꺼내기</button></div>`).join("") || '<p class="hint">담아둔 후보자가 없습니다.</p>'}<p class="hint cd-knone" hidden>찾는 후보자가 없습니다.</p></div>
      <p class="hint cd-knote">담아둔 후보자는 위 고르기 목록에 나오지 않습니다. 이름을 누르면 정보를 보고, [꺼내기]를 누르면 다시 목록에 나옵니다.</p></div>`;
  }
  document.addEventListener("input", (ev) => {
    if (!ev.target || ev.target.id !== "cdKq") return;
    const q = ev.target.value.trim().toLowerCase(), qd = q.replace(/\D/g, "");
    let n = 0;
    document.querySelectorAll(".cd-krow").forEach((r) => { const k = r.dataset.k.toLowerCase(), hit = !q || k.includes(q) || (qd.length >= 3 && k.includes(qd)); r.hidden = !hit; if (hit) n++; });
    const none = document.querySelector(".cd-knone"); if (none) none.hidden = n > 0 || !q;
  });
  async function cdKeepSet(id, on) { // on = 담아두기, off = 꺼내기
    const fresh = await stReadJsonIn(tmDir, CD_FILE);
    const q = fresh && Array.isArray(fresh.사람) ? fresh.사람.find((x) => String(x.id) === String(id)) : null;
    if (!q) throw new Error("후보자 파일에서 이 사람을 찾지 못했습니다");
    if (on) { const d = new Date(), z = (x) => String(x).padStart(2, "0"); q.담아둔일시 = `${d.getFullYear()}-${z(d.getMonth() + 1)}-${z(d.getDate())} ${z(d.getHours())}:${z(d.getMinutes())}`; }
    else delete q.담아둔일시;
    await stWriteJsonIn(tmDir, CD_FILE, fresh);
    cdDoc = fresh;
    return q;
  }
  document.addEventListener("click", async (ev) => {
    const t = ev.target.closest && ev.target.closest("[data-cd-keep],[data-cd-keepout],[data-cd-keeplist],[data-cd-keepback],[data-cd-keepopen]");
    if (!t) return;
    const d = t.dataset;
    if (d.cdKeeplist != null) { cdKeepView = true; renderCdKeep(); window.scrollTo(0, 0); return; }
    if (d.cdKeepback != null) { cdKeepView = false; if (!cdPeople().some((x) => String(x.id) === cdSel)) cdSel = String((cdPeople()[0] || {}).id || ""); if (cdPeople().length) renderCandInfo(); else openCandInfo(); window.scrollTo(0, 0); return; }
    if (d.cdKeepopen) { cdKeepView = false; cdSel = d.cdKeepopen; renderCandInfo(); window.scrollTo(0, 0); return; }
    t.disabled = true;
    try {
      if (d.cdKeep) {
        const q = await cdKeepSet(d.cdKeep, true), ppl = cdPeople();
        cdSel = String((ppl[0] || {}).id || "");
        if (ppl.length) renderCandInfo(); else openCandInfo();
        undoToast(`${q.이름 || "후보자"} 님을 담아 두었습니다(목록에서 빠짐)`, async () => { try { await cdKeepSet(q.id, false); cdSel = String(q.id); cdKeepView = false; renderCandInfo(); toast("되돌렸습니다."); } catch (e) { handleError(e); } });
      } else if (d.cdKeepout) {
        const q = await cdKeepSet(d.cdKeepout, false);
        toast(`${q.이름 || "후보자"} 님을 꺼냈습니다(목록에 다시 나옴)`);
        if (cdKeepView) renderCdKeep(); else { cdSel = String(q.id); renderCandInfo(); }
      }
    } catch (e) { t.disabled = false; handleError(e); }
  });
  function cdCellEdit([k, t, ph], val, p) { // 후보자 칸 하나(수정·추가 창 공통)
      const v = val(k);
      if (k === "도입자") return introSel("cdf", p, cdTeam);
      if (t === "year") return `<input data-cdf="${k}" type="text" data-dp="year" class="tm-narrow cd-year" value="${esc(v)}" placeholder="${esc(ph || "")}" aria-label="${esc(k)}">`;
      if (t === "month") return `<input data-cdf="${k}" type="text" data-dp="month" data-dp-range="cd-${k.replace(/(시작|끝)$/, "")}"${/끝$/.test(k) ? " data-dp-end" : ""} class="cd-month" value="${esc(v)}" aria-label="${esc("병역 " + (ph || k))}">`;
      if (/^pick:/.test(t)) { const os = t.slice(5).split("|"), old = v && !os.includes(v) ? v : ""; return `<select data-cdf="${k}" aria-label="${esc(k)}"><option value="">선택</option>${os.map((x) => `<option ${v === x ? "selected" : ""}>${esc(x)}</option>`).join("")}${old ? `<option value="${esc(old)}" selected>이전: ${esc(old)}</option>` : ""}</select>`; }
      if (t === "check") return `<label class="cd-chk"><input type="checkbox" class="cb" data-cdf="${k}" data-on="${esc(ph || "예")}" ${v ? "checked" : ""}> ${esc(ph || "예")}</label>`;
      if (t === "md") return `<input data-cdf="${k}" type="text" data-dp="md" class="cd-mdin" value="${esc(v)}" placeholder="${esc((ph || "") + " 월·일")}" aria-label="${esc(ph || k)}">`; // v561: 공통 날짜 달력(월일)
      if (k === "통신사") { const base = ["SKT", "KT", "LG U+", "SKT 알뜰폰", "KT 알뜰폰", "LG U+ 알뜰폰"], os = v && !base.includes(v) ? [v, ...base] : base; return `<select data-cdf="통신사" class="tm-narrow" aria-label="통신사"><option value=""${v ? "" : " selected"}>통신사</option>${os.map((o) => `<option${o === v ? " selected" : ""}>${esc(o)}</option>`).join("")}</select>`; }
      return `<input data-cdf="${k}" type="${t === "secret" || t === "date" ? "text" : t}"${t === "date" ? ' data-dp="date"' : ""} value="${esc(k === "주민번호" ? fmtRrn(v) : t === "tel" ? fmtPhoneText(v) : v)}" placeholder="${esc(ph || "")}" aria-label="${esc(k)}"${t === "tel" ? ' inputmode="tel"' : ""} autocomplete="off" autocapitalize="off" autocorrect="off" spellcheck="false">`;
    }
  function renderCandInfo() {
    if (cdKeepView) return renderCdKeep();
    const box = $("teamView"), ppl = cdPeople(), kp = cdKept().find((x) => String(x.id) === cdSel), p = kp || ppl.find((x) => String(x.id) === cdSel) || ppl[0];
    if (!box || !p) return;
    const pre = cdEdit ? cdPrefill(p) : {};
    const val = (k) => (k in pre ? pre[k] : String(p[k] == null ? "" : p[k]));
    const view = ([k, t]) => {
      const v = val(k);
      if (k === "도입자") { const n = introName(p, cdTeam); return n ? esc(n) : ""; }
      if (!v) return "";
      if (t === "secret") return `<span class="tm-sec"><span class="tm-secv">${esc(fmtRrn(v))}</span></span>`;
      if (t === "tel" || PHONE_COLS.includes(k)) return esc(fmtPhoneText(v)); // v561(사용자 지시): 후보자정보 주민번호 복사 버튼 없앰
      if (k === "주소") return `${esc(v)} ${mapBtn(v)}`;
      if (k === "고등학교졸업") return /^\d{4}$/.test(v) ? esc(v) + "년 졸업" : esc(v);
      if (k === "병역시작") return esc(cdYm(v)) + " ~";
      if (k === "병역끝") return (val("병역시작") ? "" : "~ ") + esc(cdYm(v));
      if (k === "결혼기념일") { const m = /^(\d{2})-(\d{2})$/.exec(v); return "기념일 " + (m ? +m[1] + "월 " + +m[2] + "일" : esc(v)); }
      return esc(v);
    };
    const edit = (c) => cdCellEdit(c, val, p);
    const rowHtml = ([lab, cells]) => {
      const parts = cdEdit ? cells.map(edit) : cells.map(view).filter(Boolean);
      return `<dt>${lab}</dt><dd class="${cells.length > 1 ? "tm-multi" : ""}">${parts.length ? parts.join(cdEdit ? "" : '<span class="tm-gap"></span>') : '<span class="tm-none">-</span>'}</dd>`;
    };
    const cvMiss = ["이름", "사번", "입사일"].filter((k) => !val(k).trim()); // v490(2026-11-19 사용자 지시): 필수 항목(이름·사번·입사일)이 다 들어가야 [팀원 전환]이 눌림
    const qa = ([k, q]) => `<div class="cd-q">${esc(q)}</div>${cdEdit ? `<textarea class="cd-a" data-cdf="${k}" rows="1" aria-label="${esc(q)}">${esc(val(k))}</textarea>` : `<div class="cd-a">${val(k) ? esc(val(k)) : '<span class="tm-none">-</span>'}</div>`}`;
    const opts = (kp ? `<option value="${esc(String(kp.id))}" selected>(담아둠) ${esc(kp.이름 || "")}</option>` : "") + ppl.map((x) => `<option value="${esc(String(x.id))}" ${String(x.id) === String(p.id) ? "selected" : ""}>${esc(x.이름 || "(이름 없음)")}</option>`).join("");
    const acts = cdEdit ? `<button type="button" class="ev-edit" data-cd-cancel>취소</button><button type="button" class="ev-edit ev-done" data-cd-save>저장</button>` : `<button type="button" class="ev-edit" data-cd-edit>수정</button>${kp ? `<button type="button" class="ev-edit" data-cd-keepout="${esc(String(p.id))}">꺼내기</button>` : `<button type="button" class="ev-edit" data-cd-keep="${esc(String(p.id))}">담아두기</button>`}`; // v563: [추가]는 제목 줄로, [수정] 오른쪽 [담아두기]/[꺼내기]
    box.innerHTML = `<div class="card tm-wrap">${cdHead(`<select data-cd-sel aria-label="후보자 고르기" ${cdEdit ? "disabled" : ""}>${opts}</select>`, acts)}
      ${cdEdit ? `<div class="cd-fill"><div class="cd-fill-h"><b>자료로 채우기</b><span class="hint"> · 메모나 이미지를 넣고 [분석]</span></div>${cdSrcHtml("edit")}<div class="cd-seg" role="radiogroup" aria-label="채우는 방법"><button type="button" role="radio" data-cdmode="empty" class="${cdFillMode === "empty" ? "on" : ""}" aria-checked="${cdFillMode === "empty"}">빈 칸만 채우기<small>있는 정보는 그대로</small></button><button type="button" role="radio" data-cdmode="replace" class="${cdFillMode === "replace" ? "on" : ""}" aria-checked="${cdFillMode === "replace"}">기존 정보도 바꾸기<small>읽은 값으로 덮어씀</small></button></div><button type="button" class="cd-fill-go" data-cdfill>분석해서 채우기</button></div><div class="cd-sum" id="cdSum" hidden></div>` : ""}
      <dl class="tm-info cd-info${cdEdit ? " editing" : ""}">${CD_ROWS.map(rowHtml).join("")}</dl>
      <hr class="tm-hr"><h4 class="cd-h">짧은 질문</h4><div class="cd-qa${cdEdit ? " editing" : ""}">${CD_QS.map(qa).join("")}</div>
      <hr class="tm-hr"><h4 class="cd-h">상세 자기소개서</h4><div class="cd-qa${cdEdit ? " editing" : ""}">${CD_ESSAYS.map(qa).join("")}</div>
      ${cdEdit ? "" : `<hr class="tm-hr"><div class="cd-conv"><span class="hint">${cvMiss.length ? "필수 항목을 먼저 넣어 주세요: " + cvMiss.join(" · ") : "사번 " + esc(val("사번")) + " · 입사일 " + esc(val("입사일")) + " · 팀원정보에 등록합니다"}</span><button type="button" class="ev-edit ev-done" data-cd-conv${cvMiss.length ? " disabled" : ""}>팀원 전환</button></div>`}</div>${cdEdit ? "" : `<div class="card cd-docs" data-cd-docdrop><div class="cd-docs-h"><h3 class="sec-title">자료</h3><span class="hint" id="cdDocsN"></span><button type="button" class="cd-up" data-cd-up>올리기</button><input type="file" multiple hidden data-cd-upfile></div><div id="cdDocsList"><p class="hint">불러오는 중…</p></div><p class="hint cd-uphint">파일을 이 상자에 끌어다 놓아도 올라갑니다.</p></div>`}`;
    if (!cdEdit) cdLoadDocs(p); // v562: 맨 아래 '자료' 박스
  }
  // v484(2026-11-19 사용자 지시): [팀원 전환] — 사번이 있어야 하고, 확인 화면(입사일·직급·도입자) 뒤 팀원 파일에 등록. 단계는 후보자 → 팀원 둘뿐(교육생 단계 없음).
  //   후보자 기록(자기소개 등)은 모두 팀원 파일로 옮기고, 후보자 파일에는 지우지 않고 '전환일시'만 적어 목록에서 빠지게 함
  function candConvertAsk(p) {
    const old = document.getElementById("cdConvPop"); if (old) old.remove();
    const d = new Date(), z = (x) => String(x).padStart(2, "0"), today = `${d.getFullYear()}-${z(d.getMonth() + 1)}-${z(d.getDate())}`;
    const box = document.createElement("div");
    box.id = "cdConvPop"; box.className = "modal";
    box.innerHTML = `<div class="modal-box cd-conv-box" role="dialog" aria-modal="true" aria-label="팀원 전환"><h3 class="sec-title">팀원 전환</h3>
      <p class="lead"><b>${esc(p.이름 || "")}</b> · 사번 ${esc(String(p.사번 || ""))}</p>
      <label class="fld"><span>입사일</span><input type="text" data-dp="date" data-cv="입사일" value="${esc(String(p.입사일 || today))}"></label>
      <label class="fld"><span>직급</span><select data-cv="직급"><option selected>FC</option><option>SL</option><option>ESL</option></select></label>
      <label class="fld"><span>도입자</span>${introSel("cv", p, cdTeam)}</label>
      <p class="hint">팀원정보에 등록하고 후보자 목록에서는 빠집니다. 후보자 기록은 지우지 않고 남겨 둡니다.</p>
      <div class="file-actions"><button type="button" class="secondary" data-cv-cancel>취소</button><button type="button" class="primary" data-cv-ok>전환</button></div></div>`;
    document.body.appendChild(box);
    box.addEventListener("click", async (e) => {
      if (e.target === box || e.target.closest("[data-cv-cancel]")) { box.remove(); return; }
      const ok = e.target.closest("[data-cv-ok]"); if (!ok) return;
      const get = (k) => String(box.querySelector(`[data-cv="${k}"]`).value || "").trim();
      ok.disabled = true; ok.textContent = "전환 중";
      try { await candConvert(String(p.id), introFix({ 입사일: get("입사일"), 직급: get("직급"), 도입자사번: get("도입자사번") }, cdTeam)); box.remove(); }
      catch (err) { ok.disabled = false; ok.textContent = "전환"; if (/^CONV:/.test(String(err.message))) toast(String(err.message).slice(5)); else handleError(err); }
    });
  }
  async function candConvert(id, extra) {
    const cands = await stReadJsonIn(tmDir, CD_FILE), team = await stReadJsonIn(tmDir, TM_FILE); // 저장 직전에 두 파일을 다시 읽음
    const c = cands && Array.isArray(cands.사람) ? cands.사람.find((x) => String(x.id) === id) : null;
    if (!c) throw new Error("CONV:후보자 파일에서 이 사람을 찾지 못했습니다");
    if (!team || !Array.isArray(team.사람)) throw new Error("CONV:팀원 파일을 읽지 못했습니다");
    const no = String(c.사번 || "").trim();
    if (!String(c.이름 || "").trim()) throw new Error("CONV:이름을 먼저 넣어 주세요");
    if (!/^\d{8}$/.test(no)) throw new Error("CONV:사번은 숫자 8자리여야 합니다");
    if (!extra.입사일) throw new Error("CONV:입사일을 넣어 주세요");
    const d = new Date(), z = (x) => String(x).padStart(2, "0"), now = `${d.getFullYear()}-${z(d.getMonth() + 1)}-${z(d.getDate())} ${z(d.getHours())}:${z(d.getMinutes())}`;
    const dup = team.사람.find((x) => String(x.사번) === no);
    if (dup && String(dup.후보자id || "") !== id) throw new Error("CONV:팀원정보에 같은 사번이 이미 있습니다");
    if (!dup) { // 팀원 파일에 먼저 등록(이미 등록됐는데 후보자 표시만 못 한 경우는 건너뜀)
      const np = { 사번: no, 이름: c.이름 || "", 단계: "팀원", 단계근거: "후보자정보에서 팀원 전환", 입사일: extra.입사일, 퇴직일: "", 직급: extra.직급 || "FC", 도입자: "도입자사번" in extra ? extra.도입자 : (c.도입자 || ""), 도입자사번: "도입자사번" in extra ? extra.도입자사번 : (c.도입자사번 || "") };
      Object.keys(c).forEach((k) => { if (!["id", "사번", "이름", "입사일", "도입자", "도입자사번", "등록일시", "수정일시", "전환일시"].includes(k)) np[k] = c[k]; });
      if (np.메모 == null) np.메모 = "";
      np.후보자id = id; np.전환일시 = now; np.수정일시 = now;
      if (np.도입자사번 || np.도입자) np.도입일 = extra.입사일; // v578(2026-11-25 사용자 지시): 도입 기록 = 도입일(기본 입사일)
      team.사람.push(np);
      await stWriteJsonIn(tmDir, TM_FILE, team);
    }
    let permNote = "";
    if (isAdmin()) { // v513: 전환과 함께 관리자 > FC 사용 목록에 올림(이메일은 나중에, 로그인은 '거부'로 시작)
      try {
        const pm = await readPerms(), nm = String(c.이름 || "").trim();
        if (!Object.keys(pm.users || {}).some((e) => (pm.names || {})[e] === nm)) {
          const key = nm + "@email.pending";
          pm.names = pm.names || {}; pm.titles = pm.titles || {};
          loginInit(pm);
          pm.users[key] = []; pm.names[key] = nm; pm.titles[key] = ORG_TITLES.includes(extra.직급) ? extra.직급 : "FC"; pm.login[key] = "거부";
          await writePerms(pm);
          if (adminState) adminState.perms = pm;
          permNote = " · 관리자 FC 사용 목록에 올렸습니다(로그인 거부)";
        } else permNote = " · FC 사용 목록에는 이미 있습니다";
      } catch (e) { console.error("FC 사용 목록 등록 실패", e); permNote = " · FC 사용 목록 등록은 못 했습니다(설정 > 사용자 추가로 넣어 주세요)"; }
    }
    c.전환일시 = now; if ("도입자사번" in extra) { c.도입자사번 = extra.도입자사번; c.도입자 = extra.도입자; }
    await stWriteJsonIn(tmDir, CD_FILE, cands);
    cdDoc = null; cdEdit = false; tmSel = no;
    toast(`${c.이름 || ""} 님을 팀원으로 전환했습니다${permNote}`);
    await openTeamInfo();
  }
  async function saveCandInfo(btn) {
    const box = $("teamView"), vals = {};
    box.querySelectorAll("[data-cdf]").forEach((el) => { vals[el.dataset.cdf] = el.type === "checkbox" ? (el.checked ? el.dataset.on || "예" : "") : String(el.value || "").trim(); }); introFix(vals, cdTeam); numVals(vals);
    box.querySelectorAll("[data-cdmd]").forEach((el) => { const m = el.querySelector('[data-md="m"]').value, d = el.querySelector('[data-md="d"]').value; vals[el.dataset.cdmd] = m && d ? m + "-" + d : ""; }); // v560: 결혼기념일 MM-DD (월·일 둘 다 골라야)
    btn.disabled = true; btn.textContent = "저장 중";
    try {
      const fresh = await stReadJsonIn(tmDir, CD_FILE);
      const q = fresh && Array.isArray(fresh.사람) ? fresh.사람.find((x) => String(x.id) === cdSel) : null;
      if (!q) throw new Error("후보자 파일에서 이 사람을 찾지 못했습니다");
      let n = 0;
      { // v562: '자료로 채우기'에 넣은 이미지·메모도 함께 저장
        const memo = String((box.querySelector('[data-cdsrc="edit"] .cd-memo') || {}).value || "").trim();
        if (cdSrcFiles.edit.length) { const fid = await cdUpload(q, cdSrcFiles.edit, "수정"); if (q.자료폴더 !== fid) q.자료폴더 = fid; cdClearFiles("edit"); n++; }
        if (memo) { const d0 = new Date(), z0 = (x) => String(x).padStart(2, "0"); q.추가메모 = (q.추가메모 ? q.추가메모 + String.fromCharCode(10) : "") + `${d0.getFullYear()}-${z0(d0.getMonth() + 1)}-${z0(d0.getDate())} ${memo}`; n++; }
      }
      Object.keys(vals).forEach((k) => {
        const old = String(q[k] == null ? "" : q[k]);
        if (old === vals[k]) return;
        if (old && ["고등학교", "최종학력", "병역", "결혼여부"].includes(k) && !(k + "_이전" in q) && !/^(고등학교|대학교|대학원|면제|육군|공군|해군|기혼)$/.test(old)) q[k + "_이전"] = old; // v560: 예전 글 값은 지우지 않고 남김
        q[k] = vals[k]; n++;
      });
      if (n) { const d = new Date(), z = (x) => String(x).padStart(2, "0"); q.수정일시 = `${d.getFullYear()}-${z(d.getMonth() + 1)}-${z(d.getDate())} ${z(d.getHours())}:${z(d.getMinutes())}`; await stWriteJsonIn(tmDir, CD_FILE, fresh); }
      cdDoc = fresh; cdEdit = false; renderCandInfo(); toast(n ? "저장했습니다" : "바뀐 것이 없습니다");
    } catch (e) { btn.disabled = false; btn.textContent = "저장"; handleError(e); }
  }
  document.addEventListener("click", (ev) => {
    const t = ev.target && ev.target.closest && ev.target.closest("[data-team-cand],[data-cd-back],[data-cd-edit],[data-cd-cancel],[data-cd-save],[data-cd-conv]");
    if (!t) return;
    if (t.hasAttribute("data-cd-conv")) { const cp = cdPeople().find((x) => String(x.id) === cdSel); if (cp) candConvertAsk(cp); return; }
    if (t.hasAttribute("data-team-cand")) { openCandInfo(); return; }
    if (t.hasAttribute("data-cd-back")) { if (cdEdit && !confirm("저장하지 않은 수정을 버릴까요?")) return; cdEdit = false; cdDoc = null; openTeam(); window.scrollTo(0, 0); return; }
    if (t.hasAttribute("data-cd-edit")) { cdEdit = true; cdClearFiles("edit"); renderCandInfo(); return; }
    if (t.hasAttribute("data-cd-cancel")) { cdEdit = false; renderCandInfo(); return; }
    saveCandInfo(t);
  });
  document.addEventListener("change", (ev) => {
    const t = ev.target && ev.target.closest && ev.target.closest("[data-cd-sel]");
    if (!t) return;
    cdSel = t.value; renderCandInfo(); window.scrollTo(0, 0);
  });
  document.addEventListener("change", (ev) => {
    const t = ev.target && ev.target.closest && ev.target.closest("[data-tm-sel]");
    if (!t) return;
    tmSel = t.value; renderTeamInfo();
  });
  async function openDash() { show("dashView"); renderDash(); }
  async function renderDash() {
    const box = $("dashView"); if (!box) return;
    box.innerHTML = `<p class="hint pad">불러오는 중…</p>`;
    try {
      const data = await loadSalesData(), myFc = await myFcName(), ord = CFG.fcOrder || [];
      // v516(2026-11-21 사용자 지시): 보는 범위 선택은 ESL만(맨 위 '팀' → 팀 합계). 관리자는 처음에 팀. 그 밖의 계정은 본인 FC만
      const pm = salesPermsAll || (await readPerms().catch(() => null));
      const esl = isAdmin() || (!!pm && titleOf(pm, me()) === "ESL");
      const [perf, fee, ctj] = canSens() ? await Promise.all([readClaudeJson("Claude/업적", "performance.json"), readClaudeJson("Claude/수수료", "수수료명세서.json"), readClaudeJson("Claude/계약", "contracts.json")]) : [null, null, null];
      const curFcs = ((perf && perf.currentFcs) || []).map((x) => x.name).filter(Boolean);
      const byOrd = (a, b) => ((ord.indexOf(a) + 1 || 999) - (ord.indexOf(b) + 1 || 999)) || a.localeCompare(b, "ko");
      const fcs = [...new Set([...data.lists.map((e) => e.fc), ...knownFcs, ...(esl ? curFcs : [])].filter(Boolean))].sort(byOrd);
      const TEAM = "__team";
      if (!esl) DSH.fc = fcs.includes(myFc) ? myFc : fcs[0] || "";
      else if (!DSH.fc || (DSH.fc !== TEAM && !fcs.includes(DSH.fc))) DSH.fc = isAdmin() || !fcs.includes(myFc) ? TEAM : myFc;
      const team = DSH.fc === TEAM, fc = DSH.fc, ym = ymNow(), inF = (f) => (team ? fcs.includes(f) : f === fc);
      const tr = transferRows(data, ym, false).filter((r) => inF(r.srcFc) && !r.archDay);
      const dd = ddaRows(data, ym).filter((r) => inF(r.srcFc) && !r.archDay && r.dda && r.dda.target !== false);
      const dbr = valueupRows(data).filter((r) => inF(r.srcFc) && !r.dbOff);
      const pg = (team ? fcs.flatMap((f) => progRows(data, f)) : progRows(data, fc)).filter((r) => r.status === "진행");
      const stg = (rows) => { const c = { 통화: 0, AP: 0, PC: 0, 계약: 0, 없음: 0 }; rows.forEach((r) => { c[c[r.stage] !== undefined && r.stage ? r.stage : "없음"]++; }); return c; };
      const ts = stg(tr), bs = stg(dbr);
      const dLast = (r) => (r.dda.coupon ? "쿠폰발송" : r.dda.resp ? "설문응답" : r.dda.send ? "설문발송" : r.dda.done ? "방문완료" : "미방문");
      const dc = { 방문완료: 0, 설문발송: 0, 설문응답: 0, 쿠폰발송: 0, 미방문: 0 }; dd.forEach((r) => dc[dLast(r)]++);
      const pc = { 고객등록: 0, 전화: 0, AP: 0, PC: 0, 계약: 0 }; pg.forEach((r) => { if (pc[r.stage] !== undefined) pc[r.stage]++; });
      // v525(2026-11-21 사용자 지시): 계약진행 카드 = (최근 3개월에 시작한 계약진행 가운데 이번 달 계약한 사람 수) / (최근 3개월에 시작한 계약진행 사람 수). 앞 숫자는 초록
      const pgAll = team ? fcs.flatMap((f) => progRows(data, f)) : progRows(data, fc);
      const ym0 = String(ymNow()), d3 = new Date(+ym0.slice(0, 4), +ym0.slice(4, 6) - 3, 1), from3 = d3.getFullYear() + "-" + String(d3.getMonth() + 1).padStart(2, "0") + "-01", thisM = ym0.slice(0, 4) + "-" + ym0.slice(4, 6);
      const pgStart = (r) => r.start || r.stageDay || String((r.p && r.p["입력일시"]) || "").slice(0, 10); // 시작일이 비었으면(11-17 옮겨 넣은 줄) 단계일 → 입력일시
      const rec3 = pgAll.filter((r) => r.status !== "보관" && r.status !== "삭제" && dateKey(pgStart(r)) >= dateKey(from3)); // 보관(중단)·삭제는 빼고 진행·완료만
      // v526(2026-11-21 사용자 지시): 사람 수가 아니라 건수로. 이달 계약 = 계약진행 단계가 이번 달 '계약'이거나, 그 고객(같은 FC·같은 이름)이 이번 달 청약 목록에 철회·거절이 아닌 청약이 있는 건
      // v527(2026-11-21 사용자 지시): 청약 관련 건은 메인FC 한 사람에게만 — 같은 청약을 여러 FC가 나눴으면(청약번호 같음) 그 계약자의 담당FC(메인FC)인 FC 줄만. 청약철회도 계약한 뒤 일이라 계약으로 셈(거절만 뺌)
      const ctRows = ((ctj && ctj.rows) || []).filter((c) => String(c.ym) === thisM); // v528(2026-11-21 사용자 지시): 거절·철회도 계약한 뒤 일이라 모두 계약으로 셈
      const mainFcOfOwner = (nm, fcsIn) => { const ents = data.lists.filter((e) => e.sheet === CUST_SHEET && String(e.name || "").trim() === nm && fcsIn.includes(e.fc)); return ents.length ? primaryFc(ents) : ""; };
      const ctGrp = new Map(); ctRows.forEach((c) => { const k = c.ppnu || c.pol || c.fc + c.owner + c.ymd; (ctGrp.get(k) || ctGrp.set(k, []).get(k)).push(c); });
      const ctNow = new Set();
      ctGrp.forEach((g) => { const own = String(g[0].owner || "").trim(), fsIn = [...new Set(g.map((c) => c.fc))]; const main = fsIn.length > 1 ? mainFcOfOwner(own, fsIn) || g.slice().sort((x, y) => (+y.ctrt || 0) - (+x.ctrt || 0))[0].fc : fsIn[0]; ctNow.add(main + "|" + own); });
      const nmCnt = new Map(); data.lists.forEach((e) => { if (e.sheet !== CUST_SHEET) return; const k = e.fc + "|" + String(e.name || "").trim(); const id = idOf(e) || e.name; (nmCnt.get(k) || nmCnt.set(k, new Set()).get(k)).add(id); });
      const pgIsDone = (r) => { const k = r.fc + "|" + String(r.name || "").trim(); return (r.stage === "계약" && String(r.stageDay || r.end || "").slice(0, 7) === thisM) || (ctNow.has(k) && (nmCnt.get(k) || new Set()).size <= 1); }; // 이름 맞추기는 그 FC에 같은 이름 고객이 한 명일 때만(PC 따라가기와 같은 규칙)
      const pgRecent = rec3.length, pgDone = rec3.filter(pgIsDone).length;
      Object.keys(pc).forEach((k) => { pc[k] = 0; }); rec3.forEach((r) => { if (pc[r.stage] !== undefined) pc[r.stage]++; }); // v525: 이달목표의 계약진행 막대도 최근 3개월에 시작한 계약진행 기준(상태와 상관없이)
      // v524(2026-11-21 사용자 지시): 이달목표 이행 점수 = (이관 통화 이상 + 따행 쿠폰발송 + DB AP 이상) ÷ (이관 + 따행 + DB) × 100. 세 줄 막대 대신 점수 한 개 + 진도 막대 한 줄
      const trDone = ts.통화 + ts.AP + ts.PC + ts.계약, ddDone = dc.쿠폰발송, dbDone = bs.AP + bs.PC + bs.계약;
      const gTot = tr.length + dd.length + dbr.length, gDone = trDone + ddDone + dbDone, gScore = gTot ? Math.round((gDone / gTot) * 100) : 0;
      const goalHtml = `<div class="dsh-goal dsh-score"><div class="dsh-gh"><b>이행 점수</b><span class="sc"><b>${gScore}</b>점</span></div>
        <div class="dsh-bar" role="img" aria-label="이행 ${gDone} / 목표 ${gTot}"><i class="s4" style="flex:${gDone || 0}"></i><i class="s0" style="flex:${Math.max(gTot - gDone, gTot ? 0 : 1)}"></i></div>
        <div class="dsh-lg"><span class="s4">이행 <b>${gDone}</b> / ${gTot}건</span><span>이관(통화) <b>${trDone}</b>/${tr.length}</span><span>따행(쿠폰) <b>${ddDone}</b>/${dd.length}</span><span>DB(AP) <b>${dbDone}</b>/${dbr.length}</span></div></div>`;
      // 이번달 영업실적(업적 자료 performance.json: 연도·사번별 n=건수, cmip=월납환산보험료, canp) — 팀은 현재 팀원 합계. 민감정보 허용 계정만
      // v517(2026-11-21 사용자 지시): 이번달 영업실적 = 청약 목록(계약\contracts.json)에서 이번 달 청약 중 철회·거절을 뺀 것을 합산, 건수·CMIP·CANP 모두 기여율(ctrt %)을 곱해 더함
      //   그 아래 '지난달 청약상태' 표(청약상태·심사상태·진단종류·[FC명: 팀일 때만]·상품명·계약일자·계약자·주피보험자·CMIP·CANP). 민감정보 허용 계정만
      let perfHtml = "", feeHtml = "", lastHtml = "", curHtml = "", fcTblHtml = "", avgHtml = "";
      if (ctj && Array.isArray(ctj.rows)) {
        const ymd = String(ym).replace(/\D/g, ""), ymH = ymd.slice(0, 4) + "-" + ymd.slice(4, 6);
        const ymAdd = (k) => { const d = new Date(+ymd.slice(0, 4), +ymd.slice(4, 6) - 1 + k, 1); return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0"); };
        const ymP = ymAdd(-1);
        const mineCt = ctj.rows.filter((r) => inF(r.fc));
        const live = (r) => !/철회|거절/.test(String(r.st || ""));
        const rate = (r) => (r.ctrt == null || r.ctrt === "" ? 1 : (+r.ctrt || 0) / 100);
        const tot = (rows) => rows.filter(live).reduce((a, r) => { const k = rate(r); a.n += k; a.cmip += (+r.cmip || 0) * k; a.canp += (+r.canp || 0) * k; return a; }, { n: 0, cmip: 0, canp: 0 });
        const won = (v) => Math.round(+v || 0).toLocaleString("ko-KR") + "원";
        const nTxt = (v) => (Math.round(v * 10) / 10).toLocaleString("ko-KR") + "건";
        const made = String(ctj.made || ""), asOf = made.replace(/\D/g, "").slice(0, 6) === ymd ? `${+made.slice(5, 7)}.${+made.slice(8, 10)} 기준` : "";
        // 이달 실적: 이번 달 청약 중 철회·거절 제외, 건수·CMIP·CANP 모두 기여율 반영
        const c = tot(mineCt.filter((r) => String(r.ym) === ymH));
        perfHtml = `<div class="dsh-perf"><div class="dsh-ph"><b>이달 실적</b><small>${esc(ymLabel(ym))}${asOf ? " · " + esc(asOf) : ""}</small></div>
            <div class="dsh-p3"><div><span>건수</span><b>${nTxt(c.n)}</b></div><div><span>CMIP</span><b>${won(c.cmip)}</b></div><div><span>CANP</span><b>${won(c.canp)}</b></div></div></div>`;
        // 지난달 청약상태(금액 말고는 가운데 정렬, 기여율 표시 없음)
        const prev = mineCt.filter((r) => String(r.ym) === ymP).sort((a, b) => String(b.ymd).localeCompare(String(a.ymd)));
        const num = (v) => (+v || 0).toLocaleString("ko-KR");
        const stTable = (rows, title, ymX, none) => `<div class="dsh-perf dsh-last"><div class="dsh-ph"><b>${title}</b><small>${esc(ymLabel(ymX.replace("-", "")))} · ${rows.length}건</small></div>` + (rows.length ? `<div class="dsh-tw"><table class="dsh-tb"><thead><tr><th>청약상태</th><th>심사상태</th><th>진단종류</th>${team ? "<th>FC명</th>" : ""}<th>상품명</th><th>계약일자</th><th>계약자</th><th>주피보험자</th><th class="n">CMIP</th><th class="n">CANP</th></tr></thead><tbody>${rows.map((r) => `<tr class="${live(r) ? "" : "off"}"><td>${esc(r.st || "")}</td><td>${esc(r.ispa || "")}</td><td>${esc(r.diag || r["진단종류"] || "")}</td>${team ? `<td>${esc(r.fc || "")}</td>` : ""}<td class="gd">${esc(r.goods || "")}</td><td>${esc(String(r.ymd || "").slice(2).replace(/-/g, "."))}</td><td>${esc(r.owner || "")}</td><td>${esc(r.insured || "")}</td><td class="n">${num(r.cmip)}</td><td class="n">${num(r.canp)}</td></tr>`).join("")}</tbody></table></div>` : `<p class="hint">${none}</p>`) + `</div>`;
        lastHtml = stTable(prev, "지난달 청약상태", ymP, "지난달 청약이 없습니다.");
        // v521(2026-11-21 사용자 지시): 이달 실적 밑에 이달 청약상태 표(지난달 표와 같은 칸)
        curHtml = stTable(mineCt.filter((r) => String(r.ym) === ymH).sort((a, b) => String(b.ymd).localeCompare(String(a.ymd))), "이달 청약상태", ymH, "이달 청약이 아직 없습니다.");
        // 팀일 때: 지난달 FC별(철회·거절 제외, 기여율 반영)
        if (team) {
          const jn = new Map(((perf && perf.currentFcs) || []).map((x) => [x.name, String(x.joined || "").slice(0, 7)])); // v520: 그달에 아직 입사 전인 FC는 빼고 보여 줌
          const rowsF = fcs.filter((f) => !jn.get(f) || jn.get(f) <= ymP).map((f) => ({ f, t: tot(prev.filter((r) => r.fc === f)) }));
          const all = tot(prev);
          fcTblHtml = `<div class="dsh-perf dsh-last"><div class="dsh-ph"><b>지난달 FC별 실적</b><small>${esc(ymLabel(ymP.replace("-", "")))} · 철회·거절 제외 · 기여율 반영</small></div><div class="dsh-tw"><table class="dsh-tb dsh-fc"><thead><tr><th>FC명</th><th class="n">건수</th><th class="n">CMIP</th><th class="n">CANP</th></tr></thead><tbody>${rowsF.map(({ f, t }) => `<tr${t.n ? "" : ' class="off"'}><td>${esc(f)}</td><td class="n">${nTxt(t.n)}</td><td class="n">${won(t.cmip)}</td><td class="n">${won(t.canp)}</td></tr>`).join("")}<tr class="sum"><td>합계</td><td class="n">${nTxt(all.n)}</td><td class="n">${won(all.cmip)}</td><td class="n">${won(all.canp)}</td></tr></tbody></table></div></div>`;
        }
        // 최근 6개월 실적: 이번 달 전 6개월(다 끝난 달). FC는 입사한 달부터만 세어 그 달 수로 나눔, 팀은 팀 합계를 (팀원이 한 명이라도 있던 달 수)로 나눔. 수수료 = 실지급금
        const win = [-6, -5, -4, -3, -2, -1].map(ymAdd);
        const joined = new Map(((perf && perf.currentFcs) || []).map((x) => [x.name, String(x.joined || "").slice(0, 7)]));
        const who = team ? fcs : [fc];
        const activeIn = (f, m) => { const j = joined.get(f); return !j || j <= m; };
        const feeOf = (f, m) => (Array.isArray(fee) ? fee.filter((r) => r.fc === f && String(r.ym) === m && r.tot).reduce((a, r) => a + (+r.tot["실지급금"] || 0), 0) : 0);
        // v523(2026-11-21 사용자 지시): '최근 6개월 평균 실적' — FC마다 월평균(입사한 달부터만 세어 나눔)을 내고, 팀은 그 FC별 월평균의 평균(더한 값 아님). 입사 전 FC는 뺌
        const per = who.map((f) => { const ms = win.filter((m) => activeIn(f, m)); if (!ms.length) return null; const a = { f, k: ms.length, n: 0, cmip: 0, canp: 0, fee: 0 }; ms.forEach((m) => { const t = tot(ctj.rows.filter((r) => r.fc === f && String(r.ym) === m)); a.n += t.n; a.cmip += t.cmip; a.canp += t.canp; a.fee += feeOf(f, m); }); ["n", "cmip", "canp", "fee"].forEach((x) => { a[x] /= a.k; }); return a; }).filter(Boolean);
        const months = win.filter((m) => who.some((f) => activeIn(f, m)));
        if (per.length) {
          const avg = (x) => per.reduce((a, p) => a + p[x], 0) / per.length;
          const k = months.length, span = `${+months[0].slice(5)}~${+months[k - 1].slice(5)}월`;
          avgHtml = `<div class="dsh-perf"><div class="dsh-ph"><b>최근 6개월 평균 실적</b><small>${esc(span)} · 월평균${team ? ` · FC ${per.length}명 평균` : per[0].k < 6 ? ` · 입사 뒤 ${per[0].k}개월` : ""}</small></div>
            <div class="dsh-p3 dsh-p4"><div><span>건수</span><b>${nTxt(avg("n"))}</b></div><div><span>CMIP</span><b>${won(avg("cmip"))}</b></div><div><span>CANP</span><b>${won(avg("canp"))}</b></div><div><span>수수료</span><b>${Array.isArray(fee) ? won(avg("fee")) : "-"}</b></div></div></div>`;
          if (team) avgHtml += `<div class="dsh-perf dsh-last"><div class="dsh-ph"><b>FC별 최근 6개월 평균</b><small>${esc(span)} · 월평균</small></div><div class="dsh-tw"><table class="dsh-tb dsh-fc"><thead><tr><th>FC명</th><th class="n">건수</th><th class="n">CMIP</th><th class="n">CANP</th><th class="n">수수료</th></tr></thead><tbody>${per.map((a) => `<tr><td>${esc(a.f)}${a.k < 6 ? `<small class="mo">${a.k}개월</small>` : ""}</td><td class="n">${nTxt(a.n)}</td><td class="n">${won(a.cmip)}</td><td class="n">${won(a.canp)}</td><td class="n">${Array.isArray(fee) ? won(a.fee) : "-"}</td></tr>`).join("")}</tbody></table></div></div>`;
        } else avgHtml = `<div class="dsh-perf"><div class="dsh-ph"><b>최근 6개월 평균 실적</b><small>입사 전이라 지난 실적이 없습니다</small></div></div>`;
      }
      if (Array.isArray(fee)) {
        const mine = fee.filter((r) => inF(r.fc) && r.tot), last = mine.map((r) => String(r.ym)).sort().pop();
        if (last) feeHtml = `<div class="dsh-card"><span>최근 수수료</span><b>${wonShort(mine.filter((r) => String(r.ym) === last).reduce((a, r) => a + (+r.tot["실지급금"] || 0), 0))}</b><small>${esc(last)} · 실지급${team ? " · 팀 합계" : ""}</small></div>`;
      }
      // 소개 랭킹 상위 3명(고른 범위 고객 기준)
      const ppl = introPeople().filter((p) => inF(p.fc)), rk = new Map();
      ppl.forEach((p) => { const k = p.introId ? "id:" + p.introId : p.introName ? "nm:" + p.introName : ""; if (k) rk.set(k, (rk.get(k) || 0) + 1); });
      const top = [...rk].sort((a, b) => b[1] - a[1]).slice(0, 3).map(([k, n]) => ({ n, reg: k.startsWith("id:"), id: k.slice(3), name: k.startsWith("id:") ? introNameOf(k.slice(3)) || k.slice(3) : k.slice(3) }));
      box.innerHTML = (esl ? `<div class="dsh-head"><span></span><select id="dshFc" aria-label="보는 범위(팀·FC)">${[TEAM, ...fcs].map((f) => `<option value="${esc(f)}"${f === fc ? " selected" : ""}>${f === TEAM ? "팀" : esc(f)}</option>`).join("")}</select></div>` : "") + `
        ${perfHtml}${curHtml}${lastHtml}${fcTblHtml}${avgHtml}
        <div class="dsh-cards">${feeHtml}<div class="dsh-card"><span>계약진행</span><b><span class="ok">${pgDone}건</span> / ${pgRecent}건</b><small>최근 3개월 중 이달 계약</small></div></div>
        <h3 class="dsh-t">이달목표 · ${esc(ymLabel(ym))}</h3>
        ${goalHtml}
        ${dashBar("계약진행 · 최근 3개월", rec3.length, [["고객등록", pc.고객등록, "s0"], ["전화", pc.전화, "s1"], ["AP", pc.AP, "s2"], ["PC", pc.PC, "s3"], ["계약", pc.계약, "s4"]])}
        <h3 class="dsh-t">소개 랭킹</h3>
        ${top.length ? `<ol class="dsh-rank">${top.map((r, i) => `<li><span class="no">${i + 1}</span><span class="nm">${r.reg ? trophyHtml(introCounts().get(r.id) || r.n) : ""}${esc(r.name)}${r.reg ? "" : ' <small>미등록</small>'}</span><b>${r.n}명</b></li>`).join("")}</ol>` : `<p class="hint">아직 소개 기록이 없습니다. 고객 화면의 [소개찾기]에서 넣을 수 있습니다.</p>`}
        <p class="hint">목표는 보관하지 않은 이관·따행(목표)·신규 DB입니다.${canSens() && !perfHtml ? " 청약 목록을 읽지 못해 이달 실적·청약상태는 숨겼습니다." : ""} 이달 실적은 이번 달 청약 중 철회·거절을 뺀 것의 합계이고, 건수·CMIP·CANP 모두 기여율을 곱했습니다. 지난달 청약상태의 금액은 계약 금액 그대로입니다. 최근 6개월 평균 실적은 이번 달 전 6개월 월평균(입사 뒤 달만, 팀은 FC별 월평균의 평균)이고 수수료는 실지급금입니다.</p>`;
    } catch (e) { box.innerHTML = ""; handleError(e); }
  }
  document.addEventListener("change", (ev) => { if (ev.target && ev.target.id === "dshFc") { DSH.fc = ev.target.value; renderDash(); } });
  // 주민번호 입력 검사: 13자리 / 앞6+뒤1(7자리) / 생년월일 8자리 / 성별만(뒤 1자리). 틀리면 메시지, 맞으면 저장 형식(norm)과 생년월일·성별
  function jminCheck(raw) {
    const d = String(raw || "").replace(/\D/g, "");
    if (!d) return String(raw || "").trim() ? { ok: false, msg: "숫자로 적어 주세요 (예: 800823-1234567)" } : { ok: true, empty: true, norm: "" };
    const bad = { ok: false, msg: "주민번호 형식을 확인해 주세요 (예: 800823-1234567 · 800823-1 · 19800823 · 성별만은 -1)" };
    const dateOk = (y, m, dd) => { const t = new Date(y, m - 1, dd); return t.getFullYear() === y && t.getMonth() === m - 1 && t.getDate() === dd && y >= 1900 && y <= new Date().getFullYear(); };
    if (d.length === 13 || d.length === 7) {
      const g = d[6], cen = "1256".includes(g) ? 1900 : "3478".includes(g) ? 2000 : 0; if (!cen) return bad;
      const y = cen + +d.slice(0, 2), m = +d.slice(2, 4), dd = +d.slice(4, 6); if (!dateOk(y, m, dd)) return { ok: false, msg: "있을 수 없는 생년월일입니다. 앞 6자리를 확인해 주세요." };
      return { ok: true, norm: d, birth: y + "-" + String(m).padStart(2, "0") + "-" + String(dd).padStart(2, "0"), sex: +g % 2 === 1 ? "남" : "여" };
    }
    if (d.length === 8) { const y = +d.slice(0, 4), m = +d.slice(4, 6), dd = +d.slice(6, 8); if (!dateOk(y, m, dd)) return { ok: false, msg: "있을 수 없는 생년월일입니다. 8자리를 확인해 주세요." }; return { ok: true, norm: d, birth: d.slice(0, 4) + "-" + d.slice(4, 6) + "-" + d.slice(6, 8) }; }
    if (d.length === 1 && "12345678".includes(d)) return { ok: true, norm: d, sex: +d % 2 === 1 ? "남" : "여" };
    return bad;
  }
  function basicViewHtml() {
    const c = current.cust;
    const x = current.extra || { values: {}, fields: [] };
    const ev = x.values || {};
    { const tel = custTelHtml(c); $("custTel").innerHTML = tel; $("custTel").hidden = !tel; } // 통신사는 추가정보를 읽은 뒤에 알 수 있어 정보 탭을 그릴 때 위쪽 전화 줄도 다시 그림
    // 고객리스트 항목 (같은 항목은 한 번만)
    const merged = new Map();
    c.entries.forEach((e) => e.fields.forEach(([k, v]) => {
      const key = k.trim();
      if (BASIC_EXCLUDE.includes(key) || NAME_COLS.includes(key) || BIRTH_KEYS.includes(key) || /^(차량번호|자동차보험갱신일)$/.test(key.replace(/\s+/g, ""))) return; // 띄어쓰기가 다른 이름도 차량 줄에 합쳐 보여 주므로 따로 표시 안 함
      if (!merged.has(key)) merged.set(key, []);
      if (!merged.get(key).includes(v)) merged.get(key).push(v);
    }));
    const take = (cols) => { const out = []; cols.forEach((k) => { (merged.get(k) || []).forEach((v) => out.push([k, v])); merged.delete(k); }); return out; };
    const phones = take(PHONE_COLS), backs = take(["후번호"]), emails = take(EMAIL_COLS), addrs = take(["주소"]), accts = take(["계좌"]);
    if (addrs.length) merged.delete("지역"); // 주소가 있으면 지역 항목은 보여 주지 않음
    const row = (label, html) => `<dt>${esc(label)}</dt><dd>${html}</dd>`;
    // 통신사 / 차량번호 / 자동차보험갱신일 (+ 미팅장소·메모가 아닌 추가 항목) — 주소 밑 한 줄
    const extraKeys = (x.fields || []).filter((k) => !EXTRA.primary.includes(k) && k !== EXTRA.carrierCol && k !== CAR_NUM && k !== CAR_DATE && !(addrs.length && k === "지역"));
    const carRows = carRowsOf({ [CAR_NUM]: ev[CAR_NUM] || currentValue(c, [CAR_NUM]), [CAR_DATE]: ev[CAR_DATE] || currentValue(c, [CAR_DATE]) });
    const chips = extraKeys.filter((k) => (ev[k] || "").trim()).map((k) =>
      `<span class="chip" title="${esc(k)}">${esc(EXTRA.monthDayCols.includes(k) ? mdText(ev[k]) : ev[k])}</span>`).join("");
    const rest = [...merged];
    const introContracts = (current.contracts || []).filter((k) => k.row["유입경로"] === "소개"); // 소개 계약은 계약자·피보험자 모두 소개로 표시
    const introNames = [...new Set(introContracts.map((k) => k.row["소개자"]).filter(Boolean))].join(", ");
    const info = `<div class="card" id="infoCard"><dl class="kv">
        ${canSens() ? ssnRow(c, backs, row) : ""}
        ${canSens() && accts.length ? row("계좌", accts.map(([, v]) => `<div class="val-row"><span class="val">${esc(v)}</span><span class="acts"><button type="button" class="map-btn copy-btn" data-copy="${esc(v)}" aria-label="계좌 복사">복사</button></span></div>`).join("")) : ""}
        ${introContracts.length ? row("소개 계약", `<div>소개${introNames ? " · " + esc(introNames) : ""} <span class="hint">(${introContracts.length}건)</span></div>`) : ""}
        ${emails.map(([k, v]) => row("이메일", valueWithActions(k, v))).join("")}
        ${addrs.length ? `<dt>주소</dt><dd class="dd-left">${addrs.map(([, v]) => addrWithMap(v)).join("")}</dd>` : ""}
        ${carRows.length ? row("차량", carRows.map((r) => `<div>${esc(r.num || "-")}${r.d ? ` · 보험 갱신 ${esc(mdText(r.d))}` : ""}</div>`).join("")) : ""}
        ${chips ? row("추가", `<div class="chips">${chips}</div>`) : ""}
        ${rest.map(([k, vs]) => /가입\s*설계\s*동의\s*만료일/.test(k) ? row("설계 동의 만료", vs.map((v) => `<div>${esc(v)}</div>`).join("")) : k === "최초유입경로" ? row("유입경로", vs.map((v) => `<div>${esc(routeShow(v))}</div>`).join("")) : k === "신한라이프 유지계약" ? row("유지 계약", (current.insCt && current.insCt.n) ? `<div>본인 ${esc(/^\d+$/.test(String(vs[0]).trim()) ? String(vs[0]).trim() : (vs[0] || "0"))}건 · 피보험자 ${current.insCt.n}건${current.insCt.holders.length ? ` <span class="hint">(계약자 ${esc(current.insCt.holders.join(", "))})</span>` : ""}</div>` : vs.map((v) => `<div>${esc(/^\d+$/.test(String(v).trim()) ? String(v).trim() + "건" : v)}</div>`).join("")) : k === "계약현황" && current.insCt && current.insCt.n && vs.some((v) => /계약\s*없음/.test(String(v))) ? row("계약현황", `<div>본인 계약 없음 · 피보험자 계약 있음</div>`) : LAND_COLS.test(k.trim()) ? row(k, vs.map((v) => valueWithActions(k, v)).join("")) : row(k, vs.map((v) => `<div>${esc(v)}</div>`).join(""))).join("")}
      </dl></div>`;
    const prim = (x.fields || []).filter((k) => EXTRA.primary.includes(k));
    // 미팅장소·메모: 수정 모드가 아니어도 바로 입력하고 저장
    const memo = x.missing ? `<div class="card"><p class="hint">'${esc(EXTRA.fileKeyword)}' 엑셀 파일을 찾지 못했습니다.</p></div>`
      : x.error ? '<div class="card"><p class="hint">추가 정보 엑셀을 읽지 못했습니다.</p></div>'
      : prim.length ? `<div class="card" id="memoCard">${prim.map((k) => `<label class="fld"><span>${esc(k)}</span>
          ${/메모/.test(k) ? `<textarea data-pm="${esc(k)}" rows="3">${esc(ev[k] || "")}</textarea>`
            : `<input data-pm="${esc(k)}" value="${esc(k === "미팅장소" ? meetingOf(c, x) : ev[k] || "")}" autocomplete="off">`}</label>`).join("")}
          <div class="file-actions"><button type="button" class="primary" id="memoSave" disabled>저장</button></div></div>` : "";
    return info + memo + followSectionHtml();
  }
  // 정보 탭 '따행' 섹션: 따행 시트(B)의 대상년월별 기록 (월 선택 칩 + 항목). 수정 모드에서는 방문약속일·방문완료·선물·전달여부 입력.
  const FOLLOW_VIEW = [["따행분류", "따행분류"], ["접촉현황", "접촉현황"], ["캠페인대상여부", "캠페인대상"], ["활동기한", ""], ["방문약속일", "방문약속일"], ["방문완료일", "방문완료일"], ["설문발송일", "설문발송일"], ["설문응답일", "설문응답일"], ["쿠폰발송", "쿠폰발송"], ["선물", "선물"], ["전달여부", "전달여부"]];
  function followSectionHtml() {
    const fl = current.follows || [];
    if (!fl.length) return "";
    const idx = Math.min(current.followIdx || 0, fl.length - 1), f = fl[idx];
    const months = fl.length > 1 ? `<div class="fl-months">${fl.map((x, i) => `<button type="button" class="cls-chip${i === idx ? " on" : ""}" data-fidx="${i}">${esc(x["대상년월"])}</button>`).join("")}</div>` : "";
    // 내용이 없는 항목은 보여 주지 않음
    const period = f["활동기한시작일"] || f["활동기한종료일"] ? `${f["활동기한시작일"] || "-"} ~ ${f["활동기한종료일"] || "-"}` : "";
    const frows = [["따행분류", f["따행분류"]], ["접촉현황", f["접촉현황"]], ["캠페인대상", f["캠페인대상여부"]], ["활동기한", period], ["방문약속일", f["방문약속일"]], ["방문완료일", f["방문완료일"]], ["설문발송일", f["설문발송일"]], ["설문응답일", f["설문응답일"]], ["쿠폰발송", f["쿠폰발송일"]], ["선물", f["선물"]], ["전달여부", f["전달여부"]]]
      .filter(([, val]) => String(val == null ? "" : val).trim());
    return `<div class="card" id="followCard"><div class="fl-head"><b>따행</b> <span class="hint">${esc(f["대상년월"])}</span></div>${months}
      ${frows.length ? `<dl class="kv">${frows.map(([k, val]) => `<dt>${esc(k)}</dt><dd>${esc(val)}</dd>`).join("")}</dl>` : '<p class="hint">입력된 내용이 없습니다.</p>'}
      </div>`;
  }
  // 수정 모드의 따행 입력칸 (선택한 대상년월의 따행 행에 저장)
  const FOLLOW_EDIT = [["방문약속일", "date"], ["방문완료일", "text"], ["선물", "text"], ["전달여부", "sel"]];
  function followFormHtml() {
    const fl = current.follows || [];
    if (!fl.length) return "";
    const f = fl[Math.min(current.followIdx || 0, fl.length - 1)];
    return `<div class="fld-sec"><div class="fl-head"><b>따행</b> <span class="hint">${esc(f["대상년월"])} — 방문 정보</span></div>` + FOLLOW_EDIT.map(([k, ty]) => `<label class="fld"><span>${esc(k)}</span>${ty === "sel"
      ? `<select data-fl="${esc(k)}"><option value=""></option>${["완료", "미완료"].map((o) => `<option${f[k] === o ? " selected" : ""}>${o}</option>`).join("")}</select>`
      : `<input data-fl="${esc(k)}" type="${ty === "date" ? "date" : "text"}" value="${esc(f[k] || "")}" autocomplete="off"${k === "방문완료일" ? ' placeholder="날짜 또는 방문완료"' : ""}>`}</label>`).join("") + `</div>`;
  }

  function bindMemo() {
    const btn = $("memoSave");
    if (!btn) return;
    const inputs = [...$("tabBody").querySelectorAll("[data-pm]")];
    const orig = { ...((current.extra && current.extra.values) || {}) };
    const shown = { ...orig, 미팅장소: meetingOf(current.cust, current.extra) };
    const dirty = () => inputs.some((i) => i.value.trim() !== (shown[i.dataset.pm] || "").trim());
    inputs.forEach((i) => i.addEventListener("input", () => { btn.disabled = !dirty(); }));
    btn.addEventListener("click", async () => {
      btn.disabled = true; btn.textContent = "저장 중…";
      const v = { ...orig };
      inputs.forEach((i) => { v[i.dataset.pm] = i.value.trim(); });
      try {
        // 미팅장소는 이 고객이 있는 고객리스트 시트마다 '미팅장소' 칸에 저장 (칸이 없으면 만듦)
        if ("미팅장소" in v && v["미팅장소"] !== (shown["미팅장소"] || "")) {
          await writeEntryCells(current.cust.entries, { 미팅장소: v["미팅장소"] });
        }
        await saveExtraFor(current.fc, current.cust.name, v, orig);
        current.extra = { ...current.extra, found: true, values: v };
        meetCache = null;
        toast("저장했습니다.");
        selectTab("basic");
      } catch (e) {
        btn.disabled = false; btn.textContent = "저장";
        saveError(e);
      }
    });
  }

  // 정보 수정: 연락처·후번호·주소·계좌·이메일 → 고객리스트 엑셀 + '우선' 1
  //   pair: 같은 줄에 반반 (연락처 | 후번호)
  const EDIT_FIELDS = [
    { label: "연락처", cols: PHONE_COLS, type: "tel" },
    { label: "주민번호", cols: ["주민번호"], type: "text", paired: true }, // v396: 후번호 대신 주민번호 한 칸(수정 칸은 ssnFld가 따로 그림)
    { label: "주소", cols: ["주소"], type: "text", multi: true },
    { label: "계좌", cols: ["계좌"], type: "text", ph: "예: 신한 110-123-456789" },
    { label: "이메일", cols: EMAIL_COLS, type: "email" },
    { label: "유입경로", col: "최초유입경로", cols: ["최초유입경로"], type: "text", ro: true }, // v383(2026-11-17): 읽기만 하는 표시(신규유입·이관유입·DB유입, 처음 한 번 정해져 안 바뀜). 소개는 유입경로가 아니라 소개자 칸
    { label: "소개자", col: "소개자명", cols: ["소개자명"], type: "text", ph: "소개자 이름", paired: true } // v398: 소개자 입력은 introFldHtml(소개확인·소개고객ID·소개자명)이 맡음
  ];
  function currentValue(c, cols) {
    for (const e of c.entries) for (const k of cols) { const f = e.fields.find(([fk]) => fk.trim() === k); if (f && f[1]) return k === "최초유입경로" ? routeShow(f[1]) : f[1]; }
    return "";
  }

  // ---- 차량 여러 대: 차량번호·자동차보험갱신일 칸에 구분자 '|'로 이어서 저장 (두 칸의 순서가 짝). 한 대만 있는 예전 값도 그대로 읽힘 ----
  const CAR_SEP = "|", CAR_NUM = "차량번호", CAR_DATE = "자동차보험갱신일";
  const splitCars = (v) => String(v || "").split(CAR_SEP).map((x) => x.trim());
  function carRowsOf(ev) {
    const nums = splitCars(ev[CAR_NUM]), ds = splitCars(ev[CAR_DATE]), out = [];
    for (let i = 0; i < Math.max(nums.length, ds.length); i++) if (nums[i] || ds[i]) out.push({ num: nums[i] || "", d: ds[i] || "" });
    return out;
  }
  function carRowHtml(r) {
    const p = parseMonthDay(r.d) || [0, 0];
    return `<div class="car-row"><input class="car-num" value="${esc(r.num)}" placeholder="차량번호" autocomplete="off">
      <input class="car-md" type="text" data-dp="md" data-dp-format="MMDD" value="${p[0] ? String(p[0]).padStart(2, "0") + String(p[1]).padStart(2, "0") : ""}" placeholder="갱신 월·일" aria-label="갱신 월일">
      <button type="button" class="car-del" aria-label="이 차량 삭제">삭제</button></div>`;
  }
  function carsFormHtml(ev) {
    const rows = carRowsOf(ev);
    if (!rows.length) rows.push({ num: "", d: "" });
    return `<div class="fld cars" id="carsBox"><span>차량번호 · 자동차보험 갱신일</span>
      <div class="car-list">${rows.map(carRowHtml).join("")}</div>
      <button type="button" class="car-add">차량 추가</button>
      <input type="hidden" data-ex="${CAR_NUM}" value="${esc(ev[CAR_NUM] || "")}"><input type="hidden" data-ex="${CAR_DATE}" value="${esc(ev[CAR_DATE] || "")}"></div>`;
  }
  function bindCars() {
    const box = $("carsBox");
    if (!box) return;
    const list = box.querySelector(".car-list"), hn = box.querySelector(`[data-ex="${CAR_NUM}"]`), hd = box.querySelector(`[data-ex="${CAR_DATE}"]`);
    const sync = () => {
      const rows = [...list.querySelectorAll(".car-row")].map((r) => {
        return { num: r.querySelector(".car-num").value.trim().replace(/\|/g, ""), d: r.querySelector(".car-md").value }; // v561: 달력(월일) 'MMDD'
      }).filter((r) => r.num || r.d);
      hn.value = rows.some((r) => r.num) ? rows.map((r) => r.num).join(CAR_SEP) : "";
      hd.value = rows.some((r) => r.d) ? rows.map((r) => r.d).join(CAR_SEP) : "";
    };
    box.addEventListener("input", sync);
    box.addEventListener("change", sync);
    box.addEventListener("click", (ev) => {
      if (ev.target.closest(".car-add")) { list.insertAdjacentHTML("beforeend", carRowHtml({ num: "", d: "" })); return; }
      const del = ev.target.closest(".car-del");
      if (!del) return;
      const row = del.closest(".car-row");
      if (list.querySelectorAll(".car-row").length > 1) row.remove();
      else { row.querySelector(".car-num").value = ""; row.querySelector(".car-md").value = ""; }
      sync();
    });
  }
  function basicFormHtml() {
    const c = current.cust;
    const x = current.extra || { values: {}, fields: [] };
    const ev = x.values || {};
    const canList = c.entries.some((e) => e.src);
    const selOpts = (f) => { const cur = currentValue(c, f.cols); const list = cur && !f.select.includes(cur) ? [cur, ...f.select] : f.select; return `<option value=""></option>` + list.map((o) => `<option value="${esc(o)}"${o === cur ? " selected" : ""}>${esc(o)}</option>`).join(""); };
    const fld = (f) => f.ro ? `<label class="fld"><span>${esc(f.label)}</span><input class="ro" readonly value="${esc(currentValue(c, f.cols))}" title="처음 들어온 길이라 바꿀 수 없습니다"></label>` : `<label class="fld"><span>${esc(f.label)}</span>
        ${f.select ? `<select data-edit="${esc(f.label)}">${selOpts(f)}</select>` : f.multi ? `<textarea data-edit="${esc(f.label)}" rows="2">${esc(currentValue(c, f.cols))}</textarea>`
          : `<input data-edit="${esc(f.label)}" value="${esc(f.label === "연락처" ? fmtPhoneText(currentValue(c, f.cols)) : currentValue(c, f.cols))}" type="${f.type}" autocomplete="off"${f.list ? ` list="${f.list}"` : ""}${f.ph ? ` placeholder="${esc(f.ph)}"` : ""}>`}</label>`;
    const field = (k) => {
      const v = esc(ev[k] || "");
      let input;
      if (EXTRA.monthDayCols.includes(k)) {
        const p = parseMonthDay(ev[k]) || [0, 0];
        input = `<input type="text" data-dp="md" data-dp-format="MMDD" data-ex="${esc(k)}" value="${p[0] ? pad2(p[0]) + pad2(p[1]) : v}" aria-label="${esc(k)}">`; // v561: 공통 날짜 달력(월일), 저장은 예전처럼 'MMDD'
      } else if (/메모/.test(k)) input = `<textarea data-ex="${esc(k)}" rows="4">${v}</textarea>`;
      else if (k === EXTRA.carrierCol) { // 통신사는 목록에서 고르기만 함 (예전에 다른 글자로 저장된 값은 그대로 보이게 목록에 추가)
        const cur = String(ev[k] || "").trim(), base = ["SKT", "KT", "LG U+", "SKT 알뜰폰", "KT 알뜰폰", "LG U+ 알뜰폰"], opts = cur && !base.includes(cur) ? [cur, ...base] : base;
        input = `<select data-ex="${esc(k)}"><option value=""></option>${opts.map((o) => `<option value="${esc(o)}"${o === cur ? " selected" : ""}>${esc(o)}</option>`).join("")}</select>`;
      }
      else input = `<input data-ex="${esc(k)}" value="${v}" autocomplete="off">`;
      return `<label class="fld ${EXTRA.monthDayCols.includes(k) ? "fld-md" : ""}"><span>${esc(k)}</span>${input}</label>`;
    };
    const fields = x.fields || [];
    // 통신사는 연락처 왼쪽 같은 줄에, 주민번호(앞 6자리-뒤 7자리)는 연락처 아래 줄에 (주민번호는 민감정보 접근 허용 계정만)
    //   앞 6자리는 생년월일에서 가져온 값을 보여 주기만 함 (생년월일은 고객 구분·사진 키로 쓰여 여기서 고치지 않음), 뒤 7자리는 '후번호' 칸에 저장
    const carrierInRow = canList && fields.includes(EXTRA.carrierCol) && !(x.missing || x.error);
    const ssnFld = () => { // v396: 주민번호 한 칸(앞 6자리 + 뒤 7자리를 이어서, 또는 앞6+뒤1 / 생년월일 8자리 / 성별만 -1). 성별·생년월일이 바뀌면 저장할 때 한 번 더 묻는다
      const cur = String(currentValue(c, ["주민번호"]) || "").trim();
      return `<label class="fld ssn-fld"><span>주민번호</span><input data-edit="주민번호" value="${esc(fmtRrn(cur))}" type="text" inputmode="numeric" autocomplete="off" placeholder="예: 800823-1234567 · 19800823 · 800823-1" aria-label="주민번호"></label>`;
    };
    const listFields = canList ? EDIT_FIELDS.filter((f) => !f.paired && (canSens() || f.label !== "계좌")).map((f) => f.label === "연락처"
      ? (carrierInRow ? `<div class="fld-2 carrier-phone">${field(EXTRA.carrierCol)}${fld(f)}</div>` : fld(f)) + (canSens() ? ssnFld() : "")
      : fld(f)).join("") + introFldHtml(c) : "";
    const hasAddr = !!currentValue(c, ["주소"]); // 주소가 있으면 지역 항목은 수정 화면에서도 숨김
    const more = fields.filter((k) => !EXTRA.primary.includes(k) && k !== CAR_NUM && k !== CAR_DATE && !(carrierInRow && k === EXTRA.carrierCol) && !(hasAddr && k === "지역")); // 미팅장소·메모는 정보 화면에서 바로 입력 (차량은 여러 대 입력 칸으로 따로)
    const hasCars = fields.includes(CAR_NUM) || fields.includes(CAR_DATE);
    return `<div class="card" id="infoForm">
        ${listFields}
        ${x.missing || x.error ? `<p class="hint">추가 정보 엑셀을 읽지 못해 통신사·차량번호·갱신일은 수정할 수 없습니다.</p>` : more.map(field).join("") + (hasCars ? carsFormHtml(ev) : "")}
        ${followFormHtml()}
        <datalist id="routeList"><option value="DB유입"><option value="이관유입"><option value="신규유입"></datalist>
        <datalist id="carrierList"><option value="SKT"><option value="KT"><option value="LG U+"><option value="SKT 알뜰폰"><option value="KT 알뜰폰"><option value="LG U+ 알뜰폰"></datalist>
        <div class="file-actions">
          <button type="button" class="secondary" id="infoCancel">취소</button>
          <button type="button" class="primary" id="infoSave">저장</button>
        </div>      </div>`;
  }

  function bindExtra() {
    if (!editMode) return bindMemo();
    const body = $("tabBody");
    // 월/일 선택 → 'MMDD'
    body.querySelectorAll(".mmdd").forEach((box) => {
      const mm = box.querySelector("[data-mm]"), dd = box.querySelector("[data-dd]"), hid = box.querySelector("[data-ex]");
      const upd = () => { hid.value = mm.value && dd.value ? pad2(mm.value) + pad2(dd.value) : (!mm.value && !dd.value ? "" : hid.value); };
      mm.addEventListener("change", upd);
      dd.addEventListener("change", upd);
    });
    bindCars();
    body.querySelectorAll(".ssn-rest").forEach((inp) => {
      const hid = inp.parentElement.querySelector("[data-d1]");
      inp.addEventListener("input", () => { inp.value = inp.value.replace(/\D/g, "").slice(0, 6); hid.value = inp.value ? hid.dataset.d1 + inp.value : ""; });
    });
    $("infoCancel").addEventListener("click", () => setEditMode(false));
    $("infoSave").addEventListener("click", () => saveInfo());
  }

  // 정보 수정 저장. stay: 저장 후 정보 화면을 다시 그리지 않고 수정모드만 끔(다른 곳으로 이동할 때). 성공하면 true
  async function saveInfo(opt) {
    const body = $("tabBody");
    const done = () => setEditMode(false, !!(opt && opt.stay));
    {
      const c = current.cust;
      const btn = $("infoSave");
      // 고객리스트: 바뀐 칸만
      const listValues = {};
      body.querySelectorAll("[data-edit]").forEach((i) => {
        const f = EDIT_FIELDS.find((x) => x.label === i.dataset.edit);
        const numF = f.label === "주민번호" || f.label === "연락처", nv = numF && f.label === "연락처" ? phoneStore(i.value) : i.value.trim(); // v576: 번호 칸은 숫자 기준으로 비교·저장
        if (numF ? onlyDigits(nv) === onlyDigits(currentValue(c, f.cols)) && /\D/.test(nv) === /[^\d\s().+-]/.test(currentValue(c, f.cols)) : nv === currentValue(c, f.cols).trim()) return;
        const col = c.entries.flatMap((e) => e.fields.map(([k]) => k.trim())).find((k) => f.cols.includes(k)) || f.col || f.label;
        listValues[col] = nv;
      });
      // 고객추가정보: 바뀐 칸이 있으면
      const inputs = [...body.querySelectorAll("[data-ex]")];
      const orig = { ...((current.extra && current.extra.values) || {}) };
      const v = { ...orig, ...Object.fromEntries(inputs.map((i) => [i.dataset.ex, i.value.trim()])) };
      const extraDirty = inputs.some((i) => i.value.trim() !== (orig[i.dataset.ex] || ""));
      // 따행 입력칸 (바뀐 칸만, 선택한 대상년월의 따행 행)
      const fl = current.follows || [], fcur = fl[Math.min(current.followIdx || 0, Math.max(fl.length - 1, 0))];
      const followChanges = {};
      if (fcur) body.querySelectorAll("[data-fl]").forEach((i) => { if (i.value.trim() !== (fcur[i.dataset.fl] || "")) followChanges[i.dataset.fl] = i.value.trim(); });
      { const ic = introChange(c); if (ic) Object.assign(listValues, ic); } // v398: 소개확인·소개고객ID·소개자명
      if ("주민번호" in listValues) { // v396: 형식 검사 + 성별·생년월일이 바뀌면 확인
        const jc = jminCheck(listValues["주민번호"]);
        if (!jc.ok) { toast(jc.msg); return false; }
        listValues["주민번호"] = jc.norm;
        { // v396: 정보가 줄어드는 입력(비우기·13자리→8자리·성별 잃기 등)은 저장 전에 확인
          const kind = (norm) => (/^\d{13}$/.test(norm) ? { t: "13자리", b: 1, x: 1, k: 1 } : /^\d{7}$/.test(norm) ? { t: "앞6+뒤1", b: 1, x: 1, k: 0 } : /^\d{8}$/.test(norm) ? { t: "생년월일 8자리", b: 1, x: 0, k: 0 } : /^\d$/.test(norm) ? { t: "성별만", b: 0, x: 1, k: 0 } : { t: "없음", b: 0, x: 0, k: 0 });
          const cj = jminCheck(currentValue(c, ["주민번호"])), ck = kind(cj.ok ? cj.norm : ""), nk = kind(jc.norm);
          if ((ck.b && !nk.b) || (ck.x && !nk.x) || (ck.k && !nk.k)) { if (!confirm(`주민번호 정보가 줄어듭니다(지금 ${ck.t} → ${nk.t}).
저장할까요?`)) return false; }
        }
        const curSex = genderOf(c), curBirth = String(c.birth || "").replace(/\./g, "-").slice(0, 10), chg = [];
        if (jc.sex && curSex && jc.sex !== curSex) chg.push(`성별이 ${curSex} → ${jc.sex}`);
        if (jc.birth && curBirth && /^\d{4}-\d\d-\d\d$/.test(curBirth) && jc.birth !== curBirth) chg.push(`생년월일이 ${curBirth} → ${jc.birth}`);
        if (chg.length && !confirm(`${chg.join(", ")}(으)로 바뀝니다.\n입력한 주민번호가 맞나요?`)) return false;
      }
      if (!Object.keys(listValues).length && !extraDirty && !Object.keys(followChanges).length) { done(); return true; }
      btn.disabled = true; btn.textContent = "저장 중…"; $("editSave").disabled = true; $("editSave").textContent = "저장 중…";
      try {
        if (Object.keys(listValues).length) {
          if (Object.keys(listValues).some((k) => !/^소개/.test(k))) { listValues[PRIORITY_COL] = 1; listValues["직접수정"] = 1; } // v385(2026-11-17): '우선'(옛 칸)과 '직접수정' 두 칸에 함께 씀 (v398: 소개 칸만 고칠 때는 표시하지 않음)
          await writeEntryCells(c.entries, listValues);
        }
        if (extraDirty) {
          await saveExtraFor(current.fc, c.name, v, orig);
          current.extra = { ...current.extra, found: true, values: v };
          meetCache = null;
        }
        if (Object.keys(followChanges).length) {
          const row = { custId: idOf(c.entries[0]) || fcur["고객ID"], fc: current.fc, follow: fcur, fileId: fcur._fileId };
          const before = {}; Object.keys(followChanges).forEach((k) => { before[k] = fcur[k] || ""; });
          Object.assign(fcur, followChanges);
          try { await writeFollowRow(row, followChanges, Object.keys(followChanges).map((k) => ({ col: k, value: followChanges[k], prev: before[k] }))); }
          catch (err) { Object.assign(fcur, before); throw err; }
        }
        current.phones = null;
        toast("저장했습니다.");
        done();
        return true;
      } catch (e) {
        btn.disabled = false; btn.textContent = "저장"; $("editSave").disabled = false; $("editSave").textContent = "확인";
        saveError(e);
        return false;
      }
    }
  }

  // 맨 위 '수정' 버튼: 성향 버튼 · 사진 바꾸기 · 정보 수정 화면
  function setEditMode(on, noRender) {
    editMode = on;
    $("editBtn").textContent = "수정";
    $("editBtn").hidden = on || !!(current && isRO(current.fc)); // 수정 중에는 위쪽에 저장·취소
    $("editBtns").hidden = !on;
    $("closeCust").hidden = on; // 일반 모드: 수정/닫기, 수정 모드: 확인/취소
    $("photoMenu").hidden = true;
    renderGrade(current.cust);
    $("avatar").classList.toggle("editable", on);
    if (!noRender) selectTab("basic"); // noRender: 곧바로 다른 탭으로 이동할 때
  }

  // ---------- 프로필 사진 ----------
  // 원형 자르기: 끌어서 위치 이동, 막대/두 손가락으로 크기 조절 → 정사각형 800px JPEG (원 안 영역)
  function cropToCircle(file) {
    return new Promise((resolve, reject) => {
      const modal = $("cropModal"), view = $("cropView"), img = $("cropImg"), zoom = $("cropZoom");
      const url = URL.createObjectURL(file);
      let V = 0, base = 1, z = 1, tx = 0, ty = 0, w = 0, h = 0;
      const pointers = new Map();
      let pinchStart = null;

      const s = () => base * z;
      const clamp = () => {
        tx = Math.min(0, Math.max(V - w * s(), tx));
        ty = Math.min(0, Math.max(V - h * s(), ty));
      };
      const draw = () => {
        clamp();
        img.style.width = w * s() + "px";
        img.style.transform = `translate(${tx}px, ${ty}px)`;
      };
      const setZoom = (nz) => {
        nz = Math.min(4, Math.max(1, nz));
        const cx = (V / 2 - tx) / s(), cy = (V / 2 - ty) / s(); // 가운데 기준으로 확대/축소
        z = nz;
        tx = V / 2 - cx * s(); ty = V / 2 - cy * s();
        zoom.value = z;
        draw();
      };
      const onDown = (e) => { view.setPointerCapture(e.pointerId); pointers.set(e.pointerId, { x: e.clientX, y: e.clientY }); pinchStart = null; };
      const onMove = (e) => {
        if (!pointers.has(e.pointerId)) return;
        const prev = pointers.get(e.pointerId);
        pointers.set(e.pointerId, { x: e.clientX, y: e.clientY });
        if (pointers.size === 1) { tx += e.clientX - prev.x; ty += e.clientY - prev.y; draw(); }
        else if (pointers.size === 2) {
          const [a, b] = [...pointers.values()];
          const d = Math.hypot(a.x - b.x, a.y - b.y);
          if (!pinchStart) pinchStart = { d, z };
          else setZoom(pinchStart.z * d / pinchStart.d);
        }
      };
      const onUp = (e) => { pointers.delete(e.pointerId); pinchStart = null; };
      const onZoom = () => setZoom(+zoom.value);
      const onWheel = (e) => { e.preventDefault(); setZoom(z * (e.deltaY < 0 ? 1.08 : 0.93)); };
      const cleanup = () => {
        modal.hidden = true;
        view.removeEventListener("pointerdown", onDown); view.removeEventListener("pointermove", onMove);
        view.removeEventListener("pointerup", onUp); view.removeEventListener("pointercancel", onUp);
        view.removeEventListener("wheel", onWheel); zoom.removeEventListener("input", onZoom);
        $("cropOk").onclick = $("cropCancel").onclick = null;
        URL.revokeObjectURL(url);
      };
      img.onload = () => {
        modal.hidden = false;
        V = view.clientWidth;
        w = img.naturalWidth; h = img.naturalHeight;
        base = V / Math.min(w, h); // 원을 꽉 채우는 크기
        z = 1; zoom.value = 1;
        tx = (V - w * base) / 2; ty = (V - h * base) / 2; // 가운데
        draw();
        view.addEventListener("pointerdown", onDown); view.addEventListener("pointermove", onMove);
        view.addEventListener("pointerup", onUp); view.addEventListener("pointercancel", onUp);
        view.addEventListener("wheel", onWheel, { passive: false }); zoom.addEventListener("input", onZoom);
        $("cropCancel").onclick = () => { cleanup(); resolve(null); };
        $("cropOk").onclick = () => {
          const OUT = 800, cv = document.createElement("canvas");
          cv.width = cv.height = OUT;
          cv.getContext("2d").drawImage(img, -tx / s(), -ty / s(), V / s(), V / s(), 0, 0, OUT, OUT);
          cleanup();
          cv.toBlob((b) => (b ? resolve(b) : reject(new Error("이미지 변환 실패"))), "image/jpeg", 0.88);
        };
      };
      img.onerror = () => { cleanup(); reject(new Error("이미지를 열 수 없음")); };
      img.src = url;
    });
  }

  // 사진 줄이기: 긴 쪽 maxSide 픽셀, JPEG
  function shrinkImage(file, maxSide, quality) {
    return new Promise((resolve, reject) => {
      const img = new Image();
      const url = URL.createObjectURL(file);
      img.onload = () => {
        const scale = Math.min(1, maxSide / Math.max(img.naturalWidth, img.naturalHeight));
        const cv = document.createElement("canvas");
        cv.width = Math.round(img.naturalWidth * scale);
        cv.height = Math.round(img.naturalHeight * scale);
        cv.getContext("2d").drawImage(img, 0, 0, cv.width, cv.height);
        URL.revokeObjectURL(url);
        cv.toBlob((b) => (b ? resolve(b) : reject(new Error("이미지 변환 실패"))), "image/jpeg", quality);
      };
      img.onerror = () => { URL.revokeObjectURL(url); reject(new Error("이미지를 열 수 없음")); };
      img.src = url;
    });
  }

  // 고객 폴더에 "날짜 - FC명 - 고객명 - 프로필사진.jpg"로 저장 (같은 날 다시 올리면 덮어씀)
  async function uploadProfilePhoto(blob) {
    const folder = current.folders[0];
    if (!folder) throw new Error("이 고객의 폴더를 찾을 수 없어 사진을 저장할 수 없습니다.");
    const d = new Date();
    const ymd = `${d.getFullYear()}${pad2(d.getMonth() + 1)}${pad2(d.getDate())}`;
    const fc = current.fc || folder.fc || "";
    const name = `${ymd} - ${fc} - ${current.cust.name} - 프로필사진.jpg`.replace(/[\\/:*?"<>|]/g, "");
    const item = await stPutIn(folder, name, await blob.arrayBuffer(), "image/jpeg");
    current.files = current.files.filter((f) => f.id !== item.id).concat({ ...item, path: item.name });
    current.profilePhoto = item;
    if (current.photoUrl) URL.revokeObjectURL(current.photoUrl);
    current.photoUrl = URL.createObjectURL(blob);
    current.photoFullUrl = current.photoUrl;
    setAvatarPhoto();
    // 기기 저장 (작은 사진)
    const small = await shrinkImage(blob, PHOTO_SIZE, 0.85).catch(() => blob);
    current.cachedPhoto = { id: item.id, mod: item.lastModifiedDateTime || "", blob: small };
    photoCachePut(photoKey(current.cust), current.cachedPhoto);
  }

  // ---------- 고객 사진 기기 저장: 작은 사진(320px)을 이 기기 브라우저(IndexedDB)에 보관 → 다음부터 바로 표시 ----------
  //   사진이 바뀌면(파일 ID·수정 시각이 다르면) 새로 받음. 로그아웃하면 모두 지움.
  const PHOTO_DB = "shlife-photos", PHOTO_SIZE = 320;
  function photoDb() {
    if (!photoDb.p) photoDb.p = new Promise((res, rej) => {
      const r = indexedDB.open(PHOTO_DB, 1);
      r.onupgradeneeded = () => r.result.createObjectStore("p");
      r.onsuccess = () => res(r.result);
      r.onerror = () => rej(r.error);
    });
    return photoDb.p;
  }
  async function photoCacheGet(key) {
    try {
      const db = await photoDb();
      return await new Promise((res) => { const q = db.transaction("p").objectStore("p").get(key); q.onsuccess = () => res(q.result || null); q.onerror = () => res(null); });
    } catch (e) { return null; }
  }
  async function photoCacheSet(key, val) {
    try {
      const db = await photoDb();
      await new Promise((res) => { const t = db.transaction("p", "readwrite"); const s = t.objectStore("p"); if (val) s.put(val, key); else s.delete(key); t.oncomplete = res; t.onerror = res; });
    } catch (e) { /* 저장 못 해도 화면엔 표시 */ }
  }
  const photoCachePut = (key, val) => photoCacheSet(key, val);
  const photoCacheDel = (key) => photoCacheSet(key, null);
  async function photoCacheClear() {
    try { if (photoDb.p) (await photoDb.p).close(); } catch (e) { /* 무시 */ }
    photoDb.p = null;
    await new Promise((res) => {
      try { const r = indexedDB.deleteDatabase(PHOTO_DB); r.onsuccess = r.onerror = r.onblocked = () => res(); } catch (e) { res(); }
      setTimeout(res, 1500);
    });
  }
  const photoKey = (cust) => me() + "|" + norm(baseName(cust.name)) + "|" + String(cust.birth || "").replace(/\D/g, "");
  // 고객 화면을 열자마자: 기기에 저장된 사진이 있으면 바로 표시 (폴더 목록을 읽기 전)
  async function showCachedPhoto(c) {
    const v = await photoCacheGet(photoKey(c.cust));
    if (!v || !v.blob || c !== current || c.photoUrl) return;
    c.cachedPhoto = v;
    c.photoUrl = URL.createObjectURL(v.blob);
    setAvatarPhoto();
    showPhotoThumb();
  }
  function showPhotoThumb() {
    const t = $("photoThumb");
    if (t && current.photoUrl && !t.querySelector("img")) t.innerHTML = `<img src="${current.photoUrl}" alt="프로필 사진">`;
  }

  // 프로필 사진 삭제 (이 고객 폴더의 '프로필사진' 이미지 모두 → OneDrive 휴지통)
  async function deleteProfilePhotos() {
    const list = current.files.filter((f) => /프로필사진/.test(f.name) && /\.(jpe?g|png|heic|webp)$/i.test(f.name));
    for (const f of list) await stDel(f.id);
    current.files = current.files.filter((f) => !list.includes(f));
    current.profilePhoto = null;
    if (current.photoUrl) URL.revokeObjectURL(current.photoUrl);
    current.photoUrl = null; current.photoFullUrl = null; current.cachedPhoto = null;
    photoCacheDel(photoKey(current.cust));
    const a = $("avatar");
    a.style.backgroundImage = ""; a.classList.remove("has-photo"); a.textContent = current.cust.name[0];
  }

  function findProfilePhoto(files) {
    return files.filter((f) => /프로필사진/.test(f.name) && /\.(jpe?g|png|heic|webp)$/i.test(f.name))
      .sort((a, b) => (b.name > a.name ? 1 : -1) || (new Date(b.lastModifiedDateTime) - new Date(a.lastModifiedDateTime)))[0] || null;
  }

  // 폴더 목록을 읽은 뒤: 기기 저장 사진이 최신이면 그대로, 아니면 Drive에서 받아 작게 줄여 기기에 저장
  async function loadProfilePhoto() {
    const c = current, p = c.profilePhoto;
    if (!p) return;
    const mod = p.lastModifiedDateTime || "";
    if (c.cachedPhoto && c.cachedPhoto.id === p.id && c.cachedPhoto.mod === mod) { showPhotoThumb(); return; }
    try {
      const full = new Blob([await stGet(p.id)], { type: "image/jpeg" });
      if (c !== current) return; // 다른 고객으로 넘어감
      c.photoFullUrl = URL.createObjectURL(full);
      const small = await shrinkImage(full, PHOTO_SIZE, 0.85).catch(() => full);
      c.cachedPhoto = { id: p.id, mod, blob: small };
      photoCachePut(photoKey(c.cust), c.cachedPhoto);
      if (c !== current) return;
      if (c.photoUrl) URL.revokeObjectURL(c.photoUrl);
      c.photoUrl = URL.createObjectURL(small);
      const t = $("photoThumb");
      if (t) t.innerHTML = "";
      setAvatarPhoto();
      showPhotoThumb();
    } catch (e) { console.error(e); }
  }

  function setAvatarPhoto() {
    const a = $("avatar");
    if (current.photoUrl) { a.style.backgroundImage = `url("${current.photoUrl}")`; a.classList.add("has-photo"); a.textContent = ""; }
  }

  function saveError(e) {
    const m = String(e && e.message);
    if (m === "NO_CUSTOMER_FOLDER") return toast("이 고객의 폴더가 없어 사진을 저장할 수 없습니다.");
    if (/423|locked/i.test(m)) return toast("엑셀 파일이 다른 곳에서 열려 있어 저장하지 못했습니다. 닫고 다시 시도하세요.");
    if (/403/.test(m)) return toast("이 아이디는 엑셀을 수정할 권한이 없습니다.");
    handleError(e);
  }

  // 상담: '상담기록' 시트에서 이 고객(고객명+담당 FC) 기록, 최근 날짜·시간이 위로
  //   처음엔 최근 10건, '더보기'로 전부 / 각 기록 끝 ✕ → 삭제 표시, 저장 버튼으로 한 번에 엑셀 반영
  let consultShowAll = false;
  function consultHtml() {
    const x = current.extra || {};
    if (x.missing) return `<p class="hint pad">'${esc(EXTRA.fileKeyword)}' 엑셀 파일을 찾지 못했습니다.</p>`;
    const list = x.consults || [];
    const N = CFG.consultPageSize || 10;
    const shown = consultShowAll ? list : list.slice(0, N);
    return `<div class="card">
        <h3 class="sec-title">상담기록</h3>
        ${list.length ? `<ul class="consult-list">${shown.map((c, i) => `<li data-ci="${i}">
            <span class="c-body"><b>${esc(c.date)}</b> ${esc(c.text).replace(/\n/g, "<br>")}</span>
            <button type="button" class="x-btn" data-delc="${i}" aria-label="이 상담기록 삭제">✕</button></li>`).join("")}</ul>
            ${list.length > N && !consultShowAll ? `<button type="button" class="more-btn center" id="consultMore">더보기 (${list.length - N}건 더)</button>` : ""}`
          : '<p class="hint">아직 상담기록이 없습니다.</p>'}
      </div>
      <div class="card">
        <h3 class="sec-title">새 상담기록</h3>
        <label class="fld"><textarea id="consultNew" rows="4" placeholder="상담 내용을 입력하세요"></textarea></label>
        <div class="file-actions"><button class="primary" id="consultSave" disabled>저장</button></div>
      </div>`;
  }

  function bindConsult() {
    const ta = $("consultNew"), btn = $("consultSave");
    if (!ta) return;
    const list = (current.extra && current.extra.consults) || [];
    const dels = new Set(); // 삭제 표시한 기록 번호
    const refresh = () => {
      btn.disabled = !ta.value.trim() && !dels.size;
      btn.textContent = dels.size ? `저장 (삭제 ${dels.size}건${ta.value.trim() ? " · 새 기록 1건" : ""})` : "저장";
    };
    const more = $("consultMore");
    if (more) more.addEventListener("click", () => { consultShowAll = true; selectTab("consult"); });
    $("tabBody").querySelectorAll("[data-delc]").forEach((x) => x.addEventListener("click", () => {
      const i = +x.dataset.delc;
      const li = x.closest("li");
      if (dels.has(i)) { dels.delete(i); li.classList.remove("to-delete"); x.textContent = "✕"; x.setAttribute("aria-label", "이 상담기록 삭제"); }
      else { dels.add(i); li.classList.add("to-delete"); x.textContent = "↺"; x.setAttribute("aria-label", "삭제 취소"); }
      refresh();
    }));
    ta.addEventListener("input", refresh);
    btn.addEventListener("click", async () => {
      if (dels.size && !confirm(`상담기록 ${dels.size}건을 엑셀에서 삭제할까요?`)) return;
      btn.disabled = true;
      btn.textContent = "저장 중…";
      try {
        const text = ta.value.trim();
        const newList = await saveConsults(current.fc, current.cust.name, text ? [{ text, when: new Date() }] : [],
          [...dels].map((i) => ({ key: list[i].key, text: list[i].text })));
        current.extra = { ...current.extra, consults: newList };
        toast(dels.size ? `상담기록을 저장했습니다 (삭제 ${dels.size}건).` : "상담기록을 엑셀에 저장했습니다.");
        renderTabs();
        selectTab("consult");
      } catch (e) {
        refresh();
        saveError(e);
      }
    });
  }

  // 계약현황: 계약리스트·이관리스트에서 이 고객(계약자/피보험자)의 계약
  //   표시 항목: [표시 이름, 엑셀 열 이름 후보] (이관리스트는 피보/계약일/보험료)
  const CONTRACT_SHOW = [
    ["증권번호", ["증권번호"]],
    ["계약자", ["계약자"]],
    ["주피보험자", ["주피보험자", "피보", "피보험자"]],
    ["최초계약일", ["최초계약일", "계약일"]],
    ["납입기간", ["납입기간"]],
    ["합계보험료", ["합계보험료", "보험료"]],
    ["이체일", ["이체일"]],
    ["유입경로", ["유입경로"]],
    ["소개자", ["소개자"]]
  ];

  // 고객 폴더의 계약 PDF: "날짜 - 고객 - FC - 상품명 - 금액원.pdf"
  const CONTRACT_PDF = /^(\d{8}) - .+? - .+? - (.+) - [\d,]+원(?: \(\d+\))?\.pdf$/i; // v510: 같은 이름이 겹칠 때 붙는 '(증권번호)'도 허용
  const normP = (s) => String(s || "").replace(/무배당|[\s()（）\[\]·,.\-_*]/g, "").toLowerCase();
  function findContractPdf(date, product, link) {
    if (link) { // v510: 계약 시트 '증권파일링크'(고객 폴더 기준 상대 경로)가 있으면 그 파일 이름으로 먼저 찾음
      const nm = String(link).split(/[\\/]/).pop();
      const hit = nm && current.files.find((f) => f.name === nm);
      if (hit) return hit;
    }
    const d = String(date || "").replace(/\D/g, "").slice(0, 8);
    const p = normP(product);
    let best = null, bestScore = 0;
    current.files.forEach((f) => {
      const m = f.name.match(CONTRACT_PDF);
      if (!m) return;
      const fp = normP(m[2]);
      const prodOk = p && fp && (fp.startsWith(p.slice(0, 8)) || p.startsWith(fp.slice(0, 8)));
      const score = (d && m[1] === d ? 2 : 0) + (prodOk ? 1 : 0);
      if (score > bestScore) { best = f; bestScore = score; }
    });
    return best;
  }
  const pdfButton = (f) => f
    ? `<div class="file-actions pdf-mini"><button class="secondary" data-open="${esc(f.id)}">계약 열기</button></div>`
    : "";

  // 상품명에서 '무배당'이 들어간 괄호 삭제: "신한통합건강보험 원(ONE)(무배당, 해약환급금 미지급형)" → "신한통합건강보험 원(ONE)"
  const productName = (s) => String(s || "").replace(/\s*[(（][^()（）]*무배당[^()（）]*[)）]/g, "").replace(/\s{2,}/g, " ").trim();

  // 담당 FC가 지금 보는 고객의 담당과 다른 계약 = 참고용(다른 FC 담당)
  const isRefContract = (c) => !!(c.fc && current.fc && c.fc !== current.fc);
  const ownContracts = () => current.contracts.filter((c) => !isRefContract(c));
  const refContracts = () => current.contracts.filter(isRefContract);
  const hideRefOn = () => { try { return localStorage.getItem("shlife.hideRefContracts") === "1"; } catch (e) { return false; } };
  // 분석 탭 맨 위 '보장 요약': 고객 시트의 '보장요약' 칸(JSON 글자)을 읽기만 함 (고객 등록·업데이트 때 미리 만들어 둔 값, 여기서 계산하지 않음)
  //   새 형식: {"출처":"전체"|"신한","기준":"2026-10-05","사망":"1억원","3대진단":{"암":"3,000만원","뇌":"2,000만원","심":"2,000만원"},"수술":"30~100만원","입원":"일 3만원","간병":"일 10만원","실손":{…}}
  //   값 "-" = 가입 없음, 키가 없음 = 원문에서 읽지 못함. 옛 형식({"기준","실손","담보":{암진단비…}})도 읽음(출처 없으면 스마트보장분석=전체로 봄)
  function coverageSummaryHtml() {
    const raw = (current.cust.entries.map((e) => fieldOf(e, "보장요약")).find(Boolean)) || "";
    if (!raw) return "";
    let j = null;
    try { j = JSON.parse(raw); } catch (e) { return ""; }
    if (!j || typeof j !== "object") return "";
    const dam = j["담보"] && typeof j["담보"] === "object" ? j["담보"] : {};
    const pick = (o, keys) => { for (const k of keys) if (o && o[k] != null && String(o[k]).trim() !== "") return String(o[k]).trim(); };
    const three = j["3대진단"] && typeof j["3대진단"] === "object" ? j["3대진단"] : {};
    const vals = {
      사망: pick(j, ["사망"]),
      암: pick(three, ["암"]) ?? pick(dam, ["암진단비", "암 진단비", "암"]),
      뇌: pick(three, ["뇌"]) ?? pick(dam, ["뇌진단비", "뇌 진단비", "뇌"]),
      심: pick(three, ["심"]) ?? pick(dam, ["심진단비", "심 진단비", "심", "심장진단비", "심장 진단비"]),
      수술: pick(j, ["수술"]) ?? pick(dam, ["수술비", "수술"]),
      입원: pick(j, ["입원"]) ?? pick(dam, ["입원비", "입원"]),
      간병: pick(j, ["간병"])
    };
    const sil = j["실손"];
    let silHtml = "";
    if (sil && typeof sil === "object" && (sil["세대"] || sil["상품"])) {
      const meta = [sil["상품"], sil["가입일"] ? sil["가입일"] + " 가입" : ""].filter(Boolean).join(" · ");
      const ben = Array.isArray(sil["혜택"]) ? sil["혜택"].filter(Boolean) : [];
      silHtml = `<div class="cov-sil"><b>실손${sil["세대"] ? " " + esc(sil["세대"]) : ""}</b>${meta ? ` <span class="hint">${esc(meta)}</span>` : ""}${ben.length ? `<ul>${ben.map((b2) => `<li>${esc(b2)}</li>`).join("")}</ul>` : ""}</div>`;
    } else if (sil === "없음") {
      silHtml = `<div class="cov-sil"><b>실손</b> <span class="hint">없음</span></div>`;
    }
    if (Object.values(vals).every((v) => v === undefined) && !silHtml) return "";
    const src = j["출처"] === "신한" ? "신한" : "전체";
    const note = src === "신한" ? "신한라이프에 가입된 보장 요약입니다." : "전체 보험사에 가입된 보장 요약입니다.";
    const fmt = (v) => v === undefined ? `<span class="cov-na">확인 안 됨</span>` : v === "-" ? `<span class="cov-none">미가입</span>` : esc(v);
    const row = (label, html) => `<div class="cov-row"><span class="cov-k">${label}</span><span class="cov-v">${html}</span></div>`;
    const tri = [["암", "암"], ["뇌", "뇌혈관"], ["심", "심혈관"]].map(([k, lab]) => `<span class="cov-tri"><i>${lab}</i>${fmt(vals[k])}</span>`).join("");
    // 설명란: 정보를 모은 날(기준, 없으면 작성일)과 '그 뒤로 달라졌을 수 있으니 다시 분석'을 안내. 6개월이 넘었으면 눈에 띄게
    const when = String(j["기준"] || j["작성일"] || "").trim();
    let foot = "";
    if (when) {
      const dt = new Date(when.replace(/\./g, "-")), old = !isNaN(+dt) && (Date.now() - +dt) / 86400000 > 183;
      foot = `<p class="cov-foot${old ? " old" : ""}">${old ? "⚠ " : ""}${esc(when)} ${src === "신한" ? "신한라이프 상세보장내용" : "스마트보장분석"}으로 모은 정보입니다. 그 뒤로 보장이 달라졌을 수 있으니, 정확한 내용은 다시 분석해 확인하세요.</p>`;
    }
    return `<div class="card cov"><h3 class="sec-title">보장 요약</h3><p class="cov-note">${note}</p>${j["코멘트"] ? `<p class="cov-comment">${esc(j["코멘트"])}</p>` : ""}<div class="cov-rows">${row("사망", fmt(vals.사망))}<div class="cov-row cov-3"><span class="cov-k">3대진단 <small>(암·뇌혈관·심혈관)</small></span><span class="cov-tris">${tri}</span></div>${row("수술", fmt(vals.수술))}${row("입원", fmt(vals.입원))}${row("간병", fmt(vals.간병))}</div>${silHtml}${foot}</div>`;
  }
  function contractHtml() {
    const all = current.contracts;
    const list = hideRefOn() ? ownContracts() : all;
    const refN = refContracts().length;
    const toggle = refN ? `<label class="ref-toggle"><input type="checkbox" id="hideRef" ${hideRefOn() ? "checked" : ""}> 참고 계약 숨기기 <span class="hint">(담당이 다른 FC ${refN}건)</span></label>` : "";
    return toggle + contractListHtml(list);
  }
  function contractListHtml(list) {
    if (!list.length) {
      const lines = [...new Set(current.cust.entries.flatMap((e) =>
        e.fields.filter(([k]) => k.includes("상품명")).flatMap(([, v]) => v.split(/\n/)).map((s) => s.trim()).filter(Boolean)))];
      return lines.length ? lines.map((l) => {
          const product = l.replace(/\s*-\s*[\d,]+원\s*$/, "");
          return `<div class="card"><div class="contract-title">${esc(productName(l))}</div>${pdfButton(findContractPdf("", product))}</div>`;
        }).join("") + '<p class="hint pad">계약리스트·이관리스트에서 이 고객의 계약을 찾지 못해 고객리스트의 상품 정보를 표시합니다.</p>'
        : '<p class="hint pad">계약리스트·이관리스트에서 이 고객의 계약을 찾지 못했습니다.</p>';
    }
    return list.map((c) => {
      const status = c.row["계약상태"] || c.row["상태"] || "";
      const val = (cols) => { for (const k of cols) if (c.row[k]) return c.row[k]; return ""; };
      const shown = CONTRACT_SHOW.map(([label, cols]) => [label, val(cols)]).filter(([, v]) => v)
        .map(([k, v]) => [k, k === "합계보험료" ? wonText(v) : v]); // 79644 / 79,644 → 79,644원
      const pdf = findContractPdf(val(["최초계약일", "계약일"]), c.row["상품명"], c.row["증권파일링크"]);
      const ref = isRefContract(c);
      return `<div class="card${ref ? " ref-card" : ""}">
        <div class="entry-head"><span class="tag t-${esc(c.type)}">${esc(typeLabel(c.type))}</span>
          ${status ? `<span class="tag">${esc(status)}</span>` : ""}
          ${ref ? `<span class="tag ref-tag">참고</span>` : ""}
          ${c.row["이관일자"] ? `<span class="hint">${esc("이관 " + c.row["이관일자"])}</span>` : ""}
          ${c.fc ? `<span class="hint rec-fc${ref ? " rec-other" : ""}">${esc("모집 FC: " + c.fc)}</span>` : ""}</div>
        <div class="contract-title">${esc(productName(c.row["상품명"]) || "(상품명 없음)")}</div>
        <dl class="kv">${shown.map(([k, v]) => `<dt>${esc(k)}</dt><dd>${esc(v)}</dd>`).join("")}</dl>
        ${pdfButton(pdf)}
      </div>`;
    }).join("");
  }

  const readWorkbook = (f) => readBook(f.id);

  async function toggleSheet(card, f, btn) {
    const box = card.querySelector(".sheet");
    if (!box.hidden) { box.hidden = true; btn.textContent = "내용 보기"; return; }
    box.hidden = false;
    btn.textContent = "접기";
    box.innerHTML = '<p class="hint">읽는 중…</p>';
    try {
      const wb = await readWorkbook(f);
      box.innerHTML = wb.SheetNames.map((name) => {
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: "", raw: false })
          .filter((r) => r.some((c) => String(c).trim() !== ""));
        if (!rows.length) return "";
        return `<h4>${esc(name)}</h4>${renderRows(rows)}`;
      }).join("") || '<p class="hint">내용이 없습니다.</p>';
    } catch (e) {
      box.innerHTML = '<p class="hint">엑셀을 읽지 못했습니다. "열기"로 확인하세요.</p>';
      console.error(e);
    }
  }

  // 2열(항목/값) 형태면 카드 목록, 아니면 가로 스크롤 표
  function renderRows(rows) {
    const twoCol = rows.every((r) => r.filter((c) => String(c).trim() !== "").length <= 2);
    if (twoCol) {
      return `<dl class="kv">${rows.map((r) => {
        const cells = r.filter((c) => String(c).trim() !== "");
        return `<dt>${esc(cells[0])}</dt><dd>${esc(cells[1] ?? "")}</dd>`;
      }).join("")}</dl>`;
    }
    const width = Math.max(...rows.map((r) => r.length));
    return `<div class="table-wrap"><table>${rows.map((r, i) =>
      `<tr>${Array.from({ length: width }, (_, j) =>
        i === 0 ? `<th>${esc(r[j])}</th>` : `<td>${esc(r[j])}</td>`).join("")}</tr>`
    ).join("")}</table></div>`;
  }

  // 엑셀 파일에서 휴대폰 번호 찾기
  async function renderContact() {
    const body = $("tabBody");
    body.innerHTML = '<p class="hint pad">엑셀에서 연락처를 찾는 중…</p>';
    if (!current.phones) {
      const phones = new Set();
      const re = /01[016789][-.\s]?\d{3,4}[-.\s]?\d{4}/g;
      current.cust.entries.forEach((e) => (e.phone.match(re) || []).forEach((p) => phones.add(p.replace(/\D/g, ""))));
      for (const f of current.files.filter((x) => kind(x.name) === "excel")) {
        try {
          const wb = await readWorkbook(f);
          wb.SheetNames.forEach((n) => {
            const text = XLSX.utils.sheet_to_csv(wb.Sheets[n]);
            (text.match(re) || []).forEach((p) => phones.add(p.replace(/\D/g, "")));
          });
        } catch (e) { console.error(e); }
      }
      current.phones = [...phones];
    }
    if (!current.phones.length) {
      body.innerHTML = '<p class="hint pad">엑셀 파일에서 휴대폰 번호를 찾지 못했습니다.</p>';
      return;
    }
    body.innerHTML = current.phones.map((p) => {
      const pretty = p.replace(/(\d{3})(\d{3,4})(\d{4})/, "$1-$2-$3");
      const carrier = ((current.extra && current.extra.values) || {})[EXTRA.carrierCol] || "";
      return `<div class="card">
        <div class="phone">${carrier ? `<span class="carrier">${esc(carrier)}</span> ` : ""}${pretty}</div>
        <div class="file-actions">
          <a class="primary" href="tel:${p}">전화</a>
          <a class="secondary" href="sms:${p}">문자</a>
        </div>
      </div>`;
    }).join("");
  }

  // ---------- 관리자: 설정 ----------
  function settingsCardHtml() {
    const s = appSettings;
    return `<div class="card" id="settingsCard">
        <h3 class="sec-title">회사 주소·카카오 키</h3>
        <div class="fld"><span>회사 주소 (일정 이동시간 계산의 출발·도착 기준)</span>
          <div class="key-row"><input id="setCompany" value="${esc(s.companyAddress || "")}" placeholder="예: 광주광역시 서구 상무중앙로 98" autocomplete="off" readonly><button type="button" class="secondary" id="companyEdit">수정</button></div></div>
        <div class="fld"><span>카카오 REST API 키 (이동 거리·시간 계산용)</span>
          <div class="key-row"><input id="setKakao" value="${esc(s.kakaoRestKey || "")}" placeholder="developers.kakao.com › 앱 › 플랫폼 키" autocomplete="off" spellcheck="false" readonly><button type="button" class="secondary" id="kakaoEdit">수정</button></div></div>
        <div class="file-actions"><button type="button" class="primary" id="setSave" disabled>설정 저장</button></div>
      </div>`;
  }
  function bindSettings() {
    const btn = $("setSave");
    if (!btn) return;
    const read = () => {
      const g = readCalendarChoices();
      const { dataSaver: _old, ...rest } = appSettings; // 예전 데이터 절약 설정은 저장할 때 함께 지움
      return { ...rest, companyAddress: $("setCompany").value.trim(),
        kakaoRestKey: $("setKakao").value.trim(),
        google: g ? { ...(appSettings.google || {}), ...g } : appSettings.google };
    };
    const calKey = (g) => JSON.stringify([((g || {}).calendars || []).map((c) => c.id).sort(), salesArr((g || {}).sales).map((c) => c.id).sort(), salesArr((g || {}).info).map((c) => c.id).sort()]);
    const dirty = () => { const v = read(); return v.companyAddress !== (appSettings.companyAddress || "") || v.kakaoRestKey !== (appSettings.kakaoRestKey || "") || calKey(v.google) !== calKey(appSettings.google); };
    const gbtn = $("gCalSave"); // 관리자 본인 카드(이메일 아래)의 구글 캘린더 [저장]: 같은 설정 파일에 저장
    const sync = () => { const d = dirty(); btn.disabled = !d; if (gbtn) gbtn.disabled = !d; };
    ["input", "change"].forEach((ev) => { $("settingsCard").addEventListener(ev, sync); if ($("adminCalCard")) $("adminCalCard").addEventListener(ev, sync); });
    // 카카오 키는 실수로 바꾸지 않게 잠가 두고, '수정'을 눌러야 고칠 수 있음 (저장하면 다시 잠금)
    //   회사 주소도 같은 방식
    const lockPair = [["setKakao", "kakaoEdit"], ["setCompany", "companyEdit"]].map(([iid, bid]) => {
      const inp = $(iid), b = $(bid);
      const lock = (on) => { inp.readOnly = on; b.textContent = on ? "수정" : "잠금"; if (!on) inp.focus(); };
      b.addEventListener("click", () => lock(!inp.readOnly));
      return lock;
    });
    const kakaoLock = (on) => lockPair.forEach((f) => f(on));
    const doSave = async () => {
      btn.disabled = true; btn.textContent = "저장 중…"; if (gbtn) { gbtn.disabled = true; gbtn.textContent = "저장 중…"; }
      try {
        await saveSettings(read());
        toast("설정을 저장했습니다.");
        kakaoLock(true);
      } catch (e) { handleError(e); }
      btn.textContent = "설정 저장"; if (gbtn) gbtn.textContent = "저장";
      sync();
    };
    btn.addEventListener("click", doSave);
    if (gbtn) gbtn.addEventListener("click", doSave);
    if ($("gFold")) $("gFold").addEventListener("toggle", () => { if ($("gFold").open && !$("gCalBox").dataset.loaded) { $("gCalBox").dataset.loaded = "1"; loadCalendarChoices(); } }); // 접혀 있다가 열 때 캘린더 목록을 불러옴
  }

  // ---------- 관리자 페이지: 아이디별 FC 폴더 권한 (앱 자체 권한 파일) ----------
  const emailOf = (p) => String(
    (p.invitation && p.invitation.email) ||
    (p.grantedToV2 && p.grantedToV2.user && p.grantedToV2.user.email) ||
    (p.grantedTo && p.grantedTo.user && p.grantedTo.user.email) ||
    (p.grantedToIdentitiesV2 && p.grantedToIdentitiesV2[0] && p.grantedToIdentitiesV2[0].user && p.grantedToIdentitiesV2[0].user.email) || ""
  ).toLowerCase();
  const nameOf = (p) => String(
    (p.grantedToV2 && p.grantedToV2.user && p.grantedToV2.user.displayName) ||
    (p.grantedTo && p.grantedTo.user && p.grantedTo.user.displayName) || "");

  let adminState = null; // { fcs: [names], perms: { users }, shares: [{email,name}] }
  const extraUsers = new Set();
  const permOpen = new Set(); // 관리자 화면에서 펼쳐 둔 사용자 카드 (평소엔 접힘, 다시 그려도 유지)

  async function loadAdminState() {
    await resolveBase();
    // v95: 권한 파일 읽기와 공유 조회(loadShares)를 동시에 시작 (서로 필요 없는 일이라 직렬로 기다릴 이유가 없음). 결과는 이전과 같음
    const itemsP = stList(base.id);
    const permsP = readPerms();
    const idsOf = (items) => Object.fromEntries(items.filter((i) => i.folder && !i.name.startsWith("_")).map((f) => [f.name, f.id]));
    const sharesP = GD && base.owner ? itemsP.then((items) => loadShares(idsOf(items))).catch((e) => { console.error(e); return null; }) : null;
    const items = await itemsP;
    const fcFolders = items.filter((i) => i.folder && !i.name.startsWith("_"));
    const fcs = fcFolders.map((f) => f.name).sort((a, b) => a.localeCompare(b, "ko"));
    const fcIds = Object.fromEntries(fcFolders.map((f) => [f.name, f.id]));
    const perms = await permsP;
    const shares = sharesP ? await sharesP : null;
    adminState = { fcs, perms, shares, fcIds };
  }

  // ---------- Google Drive 공유 (관리자: 고객 폴더 = 편집, 설정 폴더 = 보기) ----------
  // directOnly: 상위 폴더에서 물려받은 권한(고객 폴더 전체 공유 등)은 제외하고 이 폴더에 직접 준 권한만
  async function gdPermList(fileId, directOnly) {
    const d = await gdFetch(GDAPI + "/files/" + encodeURIComponent(fileId) + "/permissions?supportsAllDrives=true&fields=" + encodeURIComponent("permissions(id,emailAddress,role,type,permissionDetails(inherited))"));
    return (d.permissions || []).filter((p) => p.type === "user" && p.emailAddress)
      .filter((p) => !directOnly || !(p.permissionDetails || []).length || (p.permissionDetails || []).some((x) => !x.inherited))
      .map((p) => ({ id: p.id, email: p.emailAddress.toLowerCase(), role: p.role }));
  }
  // 고객 폴더의 상위 폴더들(예: Claude 폴더)과 거기에 직접 준 공유 — 상위에 공유돼 있으면 하위 폴더 권한을 바꿀 수 없다
  async function loadAncestors() {
    const out = [], permPs = [];
    const getF = (id) => gdFetch(GDAPI + "/files/" + encodeURIComponent(id) + "?fields=id,name,parents&supportsAllDrives=true").catch(() => null);
    let cur = await getF(base.id); // v95: 이전에는 같은 폴더를 두 번 조회했음
    for (let i = 0; i < 5; i++) {
      const pid = cur && (cur.parents || [])[0];
      if (!pid) break;
      const pf = await getF(pid);
      if (!pf) break;
      const item = { id: pid, name: pf.name, perms: [] };
      out.push(item);
      permPs.push(gdPermList(pid, true).catch(() => []).then((p) => { item.perms = p; })); // 다음 폴더 조회와 동시에
      if (!(pf.parents || []).length) break;
      cur = pf;
    }
    await Promise.all(permPs);
    return out;
  }
  async function loadShares(fcIds) {
    await resolveAppDir();
    const ids = fcIds || (adminState && adminState.fcIds) || {};
    const fc = {}; // FC 폴더별 공유
    const [cust, app, anc] = await Promise.all([gdPermList(base.id, true), appDir.id ? gdPermList(appDir.id, true) : [], loadAncestors().catch(() => []),
      Promise.all(Object.entries(ids).map(async ([name, id]) => { fc[name] = await gdPermList(id, true).catch(() => []); }))]);
    return { cust, app, fc, anc };
  }
  // 이 계정에게 직접 준 상위 폴더 공유 (소유자 제외)
  const ancShares = (em) => ((adminState && adminState.shares && adminState.shares.anc) || [])
    .flatMap((a) => a.perms.filter((p) => p.email === em && p.role !== "owner").map((p) => ({ folderId: a.id, folder: a.name, perm: p })));
  // 공유 상태: "ok" (고객 편집 + 설정 보기 이상) / "part" / "none"
  function shareState(em) {
    const s = adminState && adminState.shares;
    if (!s) return null;
    const c = s.cust.find((p) => p.email === em);
    const a = s.app.find((p) => p.email === em);
    const cOk = c && (c.role === "writer" || c.role === "owner");
    const anyFc = Object.values(s.fc || {}).some((l) => l.some((p) => p.email === em));
    if (cOk && a) return "ok";
    return c || a || anyFc || ancShares(em).length ? "part" : "none";
  }
  // 이 계정의 FC 권한 목록(list)이 Drive 공유와 정확히 맞는지: 전체("*")는 고객 폴더 편집, 그 외는 FC 폴더별 편집/보기 + 설정 폴더 보기
  function shareMatches(em, list) {
    const s = adminState && adminState.shares;
    if (!s) return true;
    if (!s.app.find((p) => p.email === em)) return false;
    if (ancShares(em).length) return false; // 상위 폴더(예: Claude)가 이 계정에 공유돼 있으면 FC 제한이 Drive에서 안 걸림
    if (list.includes("*")) { const c = s.cust.find((p) => p.email === em); return !!c && (c.role === "writer" || c.role === "owner"); }
    if (s.cust.find((p) => p.email === em && p.role !== "owner")) return false; // 예전 방식(고객 폴더 전체 공유)이면 FC 단위로 바꿔야 함
    return Object.keys(s.fc || {}).every((name) => {
      const mine = (s.fc[name] || []).find((p) => p.email === em);
      const want = list.find((x) => grantFc(x) === name);
      if (!want) return !mine;
      return !!mine && (grantRo(want) ? mine.role === "reader" : mine.role === "writer" || mine.role === "owner");
    });
  }
  async function shareFolders(em) {
    if (!confirm(`${em} 계정에 Google Drive 폴더를 공유할까요?\n· 고객 폴더: 편집 (메모·미팅장소·사진 저장)\n· 설정 폴더: 보기\n공유 알림 메일이 그 계정으로 갑니다.`)) return;
    try {
      const tt = await grantFolders(em, adminState.perms.users[em]);
      await refreshSharesFor(tt);
      toast(`${em} 계정에 폴더를 공유했습니다.`);
      const pending = collectPending();
      renderAdmin();
      applyPending(pending);
    } catch (e) {
      handleError(/Drive 400/.test(e.message) ? new Error("공유하지 못했습니다. 구글 계정 이메일이 맞는지 확인해 주세요.") : e);
    }
  }
  // 공유 맞추기(FC 단위): 상위 폴더(Claude 등) 직접 공유 해제 → 고객 루트 직접 공유 해제 → 체크한 FC 폴더 편집/보기 → 빠진 FC 폴더 해제 → 설정 폴더 보기.
  // 중간에 실패하면 지운 공유를 모두 복구하고, 만든 뒤에는 실제로 직접 권한이 생겼는지 다시 확인(Drive 반영 지연 대비 재시도).
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  async function retry(fn, tries) {
    let last;
    for (let k = 0; k < (tries || 3); k++) {
      try { return await fn(); } catch (e) { last = e; if (!/Drive (5\d\d|429)|Failed to fetch|NetworkError/.test(String(e.message))) throw e; await sleep(2500); }
    }
    throw last;
  }
  async function grantFolders(em, list) {
    const st = adminState.shares;
    list = (list || []).filter((g) => g !== "*");
    if (!list.length) throw new Error("FC를 하나 이상 체크하세요. (전체(*)는 지원하지 않습니다)");
    const permUrl = (id, pid) => GDAPI + "/files/" + encodeURIComponent(id) + "/permissions" + (pid ? "/" + encodeURIComponent(pid) : "") + "?supportsAllDrives=true";
    const hadAny = st.cust.some((p) => p.email === em) || st.app.some((p) => p.email === em) || Object.values(st.fc || {}).some((l) => l.some((p) => p.email === em));
    const add = (id, role) => gdFetch(permUrl(id) + "&sendNotificationEmail=" + (hadAny ? "false" : "true"),
      { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ type: "user", role, emailAddress: em }) });
    const patch = (id, pid, role) => gdFetch(permUrl(id, pid), { method: "PATCH", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ role }) });
    const del = (id, pid) => gdFetch(permUrl(id, pid), { method: "DELETE" });
    const ids = adminState.fcIds || {};
    for (const g of list) if (!ids[grantFc(g)]) throw new Error("FC 폴더를 찾을 수 없습니다: " + grantFc(g));
    // v96: 단계 순서(상위 해제 → 고객 루트 해제 → FC 공유 → 빠진 FC 해제 → 설정 폴더 보기)는 그대로, 같은 단계 안에서만 병렬(동시 3개).
    //       고정 2.5초 대기 대신 0.8초 간격 확인. 해제한 공유는 실제로 사라진 것을 확인한 뒤에만 FC 공유 진행. 실패하면 지운 공유를 모두 되돌림
    const undo = []; // { label, fn } — 지운 공유를 다시 주는 함수
    const remove = async (id, perm, label) => { await retry(() => del(id, perm.id)); undo.push({ label, fn: () => add(id, perm.role === "owner" ? "writer" : perm.role) }); };
    const mineIn = async (id) => (await gdPermList(id, true)).find((p) => p.email === em);
    const roleOk = (m, role) => !!m && (m.role === role || (role === "writer" && m.role === "owner"));
    const waitGone = async (id) => { // 이 폴더의 '직접' 공유가 실제로 사라질 때까지 (0.8초 간격, 최대 약 6.4초)
      for (let k = 0; k < 8; k++) { if (!(await mineIn(id))) return; await sleep(800); }
      throw new Error("공유 해제가 반영되지 않았습니다: " + id);
    };
    // 이 폴더에 이 계정의 '직접' 권한이 role대로 생길 때까지 (요청 후 0.8초 간격 확인, 최대 약 6.4초, 안 되면 처음부터 다시 최대 3번)
    const ensureDirect = async (id, role) => {
      for (let a = 0; a < 3; a++) {
        const mine = await mineIn(id);
        if (roleOk(mine, role)) return;
        if (mine && mine.role !== "owner") await retry(() => patch(id, mine.id, role)); else await retry(() => add(id, role));
        for (let k = 0; k < 8; k++) { await sleep(800); if (roleOk(await mineIn(id), role)) return; }
      }
      throw new Error("공유가 반영되지 않았습니다: " + id);
    };
    const failIfAny = (rs) => { const bad = rs.find((r) => r && !r.ok); if (bad) throw bad.e; };
    const touched = { cust: false, app: false, fc: [], anc: [] }; // 바뀐(바뀌었을 수 있는) 폴더 — 저장 후 이 폴더의 공유 목록만 다시 읽음
    try {
      const ancs = ancShares(em);
      for (const a of ancs) { await remove(a.folderId, a.perm, "상위 폴더 " + a.folder); touched.anc.push(a.folderId); }      // 1 상위 폴더 공유 해제 (예: Claude)
      const c = st.cust.find((p) => p.email === em);
      if (c && c.role !== "owner") { await remove(base.id, c, "고객 루트"); touched.cust = true; }                         // 2 고객 루트 직접 공유 해제
      for (const a of ancs) await waitGone(a.folderId);                                                                    //   해제가 실제로 반영된 뒤에만 다음 단계
      if (touched.cust) await waitGone(base.id);
      const want = list.map((g) => [grantFc(g), grantRo(g) ? "reader" : "writer"]);                                        // 3 FC 폴더 공유 (동시 3개)
      want.forEach(([n]) => touched.fc.push(n));
      failIfAny(await runLimited(want.map(([n, role]) => () => ensureDirect(ids[n], role)), 3));
      const extra = Object.entries(st.fc || {}).filter(([fc, ps]) => ps.some((p) => p.email === em && p.role !== "owner") && !list.some((g) => grantFc(g) === fc));
      extra.forEach(([fc]) => touched.fc.push(fc));                                                                         // 4 목록에서 빠진 FC 폴더 해제 (동시 3개)
      failIfAny(await runLimited(extra.map(([fc, ps]) => () => remove(ids[fc], ps.find((p) => p.email === em), "FC " + fc)), 3));
      if (appDir.id) { await ensureDirect(appDir.id, "reader"); touched.app = true; }                                       // 5 설정 폴더 보기 (상위 공유가 사라져도 직접 권한 유지)
    } catch (e) {
      const failed = [];
      for (const u of undo.reverse()) { try { await u.fn(); } catch (e2) { console.error("복구 실패", e2); failed.push(u.label); } }
      if (failed.length) e.message += " (되돌리지 못한 공유: " + failed.join(", ") + ")";
      throw e;
    }
    return touched;
  }
  // 동시에 최대 limit개씩 실행하고, 진행 중인 것이 모두 끝난 뒤(실패 포함) 결과 배열 [{ok, v|e}] 반환 (실패 뒤에는 새로 시작하지 않아 시작 안 한 칸은 undefined)
  async function runLimited(tasks, limit) {
    const res = new Array(tasks.length);
    let i = 0;
    await Promise.all(Array.from({ length: Math.min(limit, tasks.length) }, async () => {
      while (i < tasks.length && !res.some((r) => r && !r.ok)) { const k = i++; try { res[k] = { ok: true, v: await tasks[k]() }; } catch (e) { res[k] = { ok: false, e }; } } // 하나라도 실패하면 새 작업은 시작하지 않고 진행 중인 것만 끝까지 기다림
    }));
    return res;
  }
  // 저장 후: 바뀐 폴더의 공유 목록만 다시 읽어 adminState.shares 갱신 (전체 재조회 대신)
  async function refreshSharesFor(tl) {
    const s = adminState && adminState.shares;
    if (!s) return;
    const ids = adminState.fcIds || {};
    const jobs = [];
    if (tl.cust) jobs.push(gdPermList(base.id, true).then((p) => { s.cust = p; }));
    if (tl.app && appDir.id) jobs.push(gdPermList(appDir.id, true).then((p) => { s.app = p; }));
    new Set(tl.fc).forEach((n) => { if (ids[n]) jobs.push(gdPermList(ids[n], true).then((p) => { s.fc[n] = p; })); });
    new Set(tl.anc).forEach((id) => { const a = (s.anc || []).find((x) => x.id === id); if (a) jobs.push(gdPermList(id, true).then((p) => { a.perms = p; })); });
    await Promise.all(jobs);
  }
  // v513: 로그인 승인/거부 토글. 승인 = 그 계정의 FC 폴더(본인 FC + 체크한 FC)와 설정 폴더 공유, 거부 = 로그인 막기(+ 원하면 공유 해제)
  async function shareAppDirReader(em) {
    const st = adminState.shares;
    if (!appDir || !appDir.id || (st && st.app.some((p) => p.email === em))) return;
    await gdFetch(GDAPI + "/files/" + encodeURIComponent(appDir.id) + "/permissions?supportsAllDrives=true&sendNotificationEmail=false", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ type: "user", role: "reader", emailAddress: em }) });
    if (st) st.app = await gdPermList(appDir.id, true);
  }
  async function toggleLogin(em, btn) {
    const p0 = adminState.perms || {}, on = loginOn(p0, em), nm = (p0.names || {})[em] || em.split("@")[0];
    if (!on && isPend(em)) return toast(`${nm} 님의 구글 계정 이메일이 아직 없습니다. 조직도에서 그 사람을 눌러 이메일을 먼저 넣어 주세요.`);
    const ids = adminState.fcIds || {}, own = (adminState.fcs || []).includes((p0.names || {})[em]) ? p0.names[em] : "";
    const list0 = ((p0.users || {})[em] || []).filter((g) => g !== "*");
    const list = [...list0.filter((g) => grantFc(g) !== own), ...(own ? [own] : [])].filter((g) => ids[grantFc(g)]);
    let unshare = false;
    if (!on) {
      if (!confirm(`${nm} (${em}) 로그인을 승인합니다.\n` + (GD ? (list.length ? `Google Drive 공유: FC 폴더(${list.map(grantFc).join(", ")}) 편집, 설정 폴더 보기` : "볼 FC 폴더가 아직 없어 설정 폴더만 공유합니다.") : ""))) return;
    } else {
      if (!confirm(`${nm} (${em}) 로그인을 거부합니다.\n다음에 앱을 열 때부터 로그인 화면에서 멈춥니다.`)) return;
      unshare = GD && shareState(em) && shareState(em) !== "none" && confirm("Google Drive 공유도 해제할까요?\n\n확인 = 공유 해제 (드라이브에서도 고객 파일을 볼 수 없음)\n취소 = 로그인만 막기 (공유가 남아 있어 드라이브에서 파일을 직접 열 수 있음)");
    }
    btn.disabled = true; btn.textContent = "저장 중";
    try {
      const pm = await readPerms();
      loginInit(pm)[em] = on ? "거부" : "승인";
      await writePerms(pm);
      adminState.perms = pm;
      let note = "";
      if (GD && adminState.shares) {
        try {
          if (!on) { if (list.length) await refreshSharesFor(await grantFolders(em, list)); await shareAppDirReader(em); }
          else if (unshare) await unshareFolders(em);
        } catch (e) { console.error(e); note = ` 다만 Drive 공유는 바꾸지 못했습니다(${String(e.message || e).slice(0, 60)}).`; }
      }
      toast(`${nm} 로그인을 ${on ? "거부" : "승인"}했습니다.${on ? (unshare ? " Drive 공유도 해제했습니다." : "") : ""}${note}`);
      renderAdmin();
    } catch (e) { btn.disabled = false; btn.textContent = on ? "승인" : "거부"; handleError(e); }
  }
  async function unshareFolders(em) {
    const st = adminState && adminState.shares;
    if (!st) return;
    const del = (fileId, p) => gdFetch(GDAPI + "/files/" + encodeURIComponent(fileId) + "/permissions/" + encodeURIComponent(p.id) + "?supportsAllDrives=true", { method: "DELETE" });
    const touched = { cust: false, app: false, fc: [], anc: [] };
    for (const a of ancShares(em)) { await del(a.folderId, a.perm); touched.anc.push(a.folderId); }
    for (const p of st.cust.filter((x) => x.email === em && x.role !== "owner")) { await del(base.id, p); touched.cust = true; }
    for (const p of st.app.filter((x) => x.email === em && x.role !== "owner")) { await del(appDir.id, p); touched.app = true; }
    for (const [fc, ps] of Object.entries(st.fc || {})) for (const p of ps.filter((x) => x.email === em && x.role !== "owner")) { await del((adminState.fcIds || {})[fc], p); touched.fc.push(fc); }
    await refreshSharesFor(touched); // v96: 전체 재조회 대신 바뀐 폴더만
    return touched;
  }
  // 표에서 아직 저장 안 한 체크 상태 (다시 그릴 때 유지)
  let collectPending = () => [];
  let applyPending = () => {};

  // ---------- 설정 (일반 사용자): 구글 캘린더(보여질 캘린더·영업일정 캘린더) + 데이터 절약 모드 ----------
  async function openMySettings() {
    show("adminView");
    const body = $("adminBody");
    body.innerHTML = featCardHtml() + `<div class="card" id="myCalCard"><details class="fold">
        <summary class="sec-title">구글 캘린더</summary>
        <div id="myCalBox"><p class="hint">캘린더 목록을 불러오는 중…</p></div>
        <div class="file-actions"><button type="button" class="primary" id="myCalSave" disabled>저장</button></div></details></div>` + startCardHtml() + smsCardHtml() + rhCardHtml() + installCardHtml();
    bindStartCard();
    bindSmsCard();
    let cals;
    try { [cals] = await Promise.all([calendarList(), loadMyPrefs(), loadSettings()]); } catch (e) {
      $("myCalBox").innerHTML = "";
      return handleError(e);
    }
    const prefs = myPrefs || {};
    const saved = (prefs.calendars || []).map((c) => c.id);
    const want = defaultCalNames();
    const isOn = (c) => (saved.length ? saved.includes(c.id) : want.includes(c.name));
    const box = $("myCalBox");
    box.innerHTML = cals.length ? calPickerHtml(cals, isOn) : '<p class="hint">볼 수 있는 캘린더가 없습니다.</p>';
    const sIds = salesArr(prefs.sales).length ? salesArr(prefs.sales).map((c) => c.id) : defaultSalesIds(cals);
    if (cals.length) bindCalPicker(box, cals, sIds, Array.isArray(prefs.info) ? prefs.info.map((c) => c.id) : sIds);
    const read = () => (cals.length ? readCalPicker(box, cals) : { calendars: [], sales: [], info: [] });
    const key = (v) => JSON.stringify([(v.calendars || []).map((c) => c.id).sort(), salesArr(v.sales).map((c) => c.id).sort(), salesArr(v.info).map((c) => c.id).sort()]);
    let base0 = saved.length ? key({ calendars: prefs.calendars, sales: prefs.sales, info: prefs.info }) : "";
    const btn = $("myCalSave");
    const dirty = () => key(read()) !== base0;
    btn.disabled = !dirty();
    body.addEventListener("change", () => { btn.disabled = !dirty(); });
    btn.addEventListener("click", async () => {
      const v = read();
      if (!v.calendars.length) return toast("보여질 캘린더를 하나 이상 고르세요.");
      btn.disabled = true; btn.textContent = "저장 중…";
      try {
        const { dataSaver: _old, ...keep } = myPrefs || {};
        await saveMyPrefs({ ...keep, calendars: v.calendars, sales: v.sales, info: v.info });
        base0 = key(v);
        toast("저장했습니다. 일정 화면에 반영됩니다.");
      } catch (e) { handleError(e); }
      btn.textContent = "저장";
      btn.disabled = !dirty();
    });
  }

  async function openAdmin() {
    if (!isAdmin()) return openMySettings();
    show("adminView");
    $("adminBody").innerHTML = '<p class="hint pad">권한 정보를 불러오는 중…</p>';
    try {
      await Promise.all([loadAdminState(), loadSettings()]);
      renderAdmin();
    } catch (e) {
      $("adminBody").innerHTML = "";
      handleError(e);
    }
  }

  // ---------- v326 조직도: 관리자(ESL) 밑으로 하부 조직, 사용자마다 보기·수정(그 사람 아래 조직의 고객) ----------
  const orgOpen = new Set(); // v369: 펼친 조직도 카드(이메일). 처음엔 모두 접힘(이름만)
  let orgDraft = null, orgSaved = "", orgStartCache = null; // orgStartCache(v333): { 이메일: {page, ai, rh} } 시작 화면·AI·오른손
  function orgLoadStarts() {
    const P = adminState.perms || {}, ids = adminState.fcIds, out = {};
    return Promise.all(orgUsers().map(async (em) => {
      let v = null; try { v = await readStartShared(P, em, ids); } catch (e) { v = null; }
      if (!v && (P.starts || {})[em]) { const st = P.starts[em]; v = { page: st.page, ai: !!st.ai, rh: null }; }
      out[em] = v;
    })).then(() => { orgStartCache = out; if (orgDraft && document.querySelector(".org-wrap")) renderOrg(); });
  }
  const orgSetHtml = (v, self) => { // v338: 오른쪽 표시란 정리 — 연한 바탕 상자, 켜진 것만 색, 꺼짐·기록 없음은 흐리게
    const ld = !orgStartCache && !self;
    const pg = v ? (START_PAGES.find((x) => x[0] === v.page) || [, "고객"])[1] : "고객";
    const row = (k, val, st) => `<span class="${st}"><i>${k}</i><b>${val}</b></span>`;
    const onoff = (x) => (x === true ? ["켬", "on"] : x === false ? ["끔", "off"] : ["–", "off"]);
    const ai = onoff(v ? !!v.ai : false), rh = onoff(v ? v.rh : null);
    return `<div class="org-set"${self ? ' title="이 기기의 설정"' : ""}>${START_FIXED ? "" : row("시작", ld ? "…" : pg, "")}${row("AI", ld ? "…" : ai[0], ld ? "off" : ai[1])}${row("오른손", ld ? "…" : rh[0], ld ? "off" : rh[1])}</div>`;
  };
  // v370(2026-11-16 사용자 지시): 이메일 없이 이름만 넣어 등록한 사용자는 임시 키 '<이름>@email.pending'. 이메일을 채워 저장하면 모든 기록의 키를 바꿈
  const isPend = (em) => /@email\.pending$/.test(String(em || ""));
  // v513(2026-11-21 사용자 규칙): 신한FC 로그인은 관리자 페이지의 승인/거부 토글로만 정함(이메일·사번이 있고 없고로 정하지 않음)
  //   권한 파일에 login 칸이 아예 없으면(토글 도입 전 파일) 예전 방식: FC 사용 목록(users)에 있으면 들어감 — 배포 직후 기존 계정이 잠기지 않게
  const loginOn = (perms, em) => {
    em = String(em || "").toLowerCase();
    if ((CFG.admins || []).map((a) => a.toLowerCase()).includes(em)) return true;
    if (!perms) return false;
    if (!perms.login || typeof perms.login !== "object") return !!(perms.users && Object.prototype.hasOwnProperty.call(perms.users, em));
    return perms.login[em] === "승인";
  };
  function loginInit(perms) { // 처음 토글을 쓸 때: 옛 방식 그대로의 값으로 login 칸을 만듦(다른 계정이 잠기지 않게)
    if (perms.login && typeof perms.login === "object") return perms.login;
    perms.login = {};
    Object.keys(perms.users || {}).forEach((em) => { perms.login[em] = isPend(em) ? "거부" : "승인"; });
    return perms.login;
  }
  function rekeyOrg(o, a, b) { if (!o || !o[a]) return o; o[b] = o[a]; delete o[a]; Object.keys(o).forEach((k) => { if (o[k] && o[k].parent === a) o[k].parent = b; }); return o; }
  function rekeyPerms(perms, a, b) {
    ["users", "names", "titles", "salesFcs", "starts", "orgApplied", "orgSalesApplied", "archived", "login"].forEach((k) => { const m = perms[k]; if (m && Object.prototype.hasOwnProperty.call(m, a)) { m[b] = m[a]; delete m[a]; } });
    if (Array.isArray(perms.sensitive)) perms.sensitive = perms.sensitive.map((x) => (x === a ? b : x));
    if (perms.org) rekeyOrg(perms.org, a, b);
  }
  const orgRoot = () => String((CFG.admins || [])[0] || "").toLowerCase();
  const orgUsers = () => { const adm = (CFG.admins || []).map((a) => a.toLowerCase()); return Object.keys((adminState.perms || {}).users || {}).map((e) => e.toLowerCase()).filter((e) => e && !adm.includes(e)); };
  const orgName = (em) => ((adminState.perms || {}).names || {})[em] || em.split("@")[0];
  function orgNorm(org) {
    const users = orgUsers(), root = orgRoot(), out = {};
    users.forEach((em) => { const o = (org || {})[em] || {}; const par = String(o.parent || "").toLowerCase(); out[em] = { parent: par && par !== em && (par === root || users.includes(par)) ? par : root, view: !!o.view, edit: !!o.view && !!o.edit }; });
    users.forEach((em) => { const seen = new Set([em]); let q = out[em].parent; while (q && q !== root) { if (seen.has(q)) { out[em].parent = root; break; } seen.add(q); q = (out[q] || {}).parent; } }); // 돌고 도는 연결은 관리자 밑으로
    return out;
  }
  const orgKids = (d, em) => Object.keys(d).filter((x) => d[x].parent === em).sort((a, b) => { const o = CFG.fcOrder || [], ia = o.indexOf(orgName(a)), ib = o.indexOf(orgName(b)); return (ia < 0 ? 999 : ia) - (ib < 0 ? 999 : ib) || orgName(a).localeCompare(orgName(b), "ko"); });
  const orgDesc = (d, em) => { const r = []; const walk = (x) => orgKids(d, x).forEach((k) => { r.push(k); walk(k); }); walk(em); return r; };
  const ORG_TITLES = ["ESL", "SSL", "GSL", "FC"];
  let orgTitles = {}, orgTitlesSaved = "";
  const titleOf = (perms, em) => { const t = ((perms || {}).titles || {})[em]; return ORG_TITLES.includes(t) ? t : (CFG.admins || []).map((a) => a.toLowerCase()).includes(em) ? "ESL" : "FC"; };
  const orgDirty = () => JSON.stringify(orgDraft) !== orgSaved || JSON.stringify(orgTitles) !== orgTitlesSaved;
  // v335(2026-11-15 사용자 지시): 사용자 X의 보기 = X의 FC를 X 위의 사용자들(관리자 제외)이 봄, 수정 = 수정 권한까지 위임
  function orgGets(d, em, fcOf) {
    const m = new Map();
    orgDesc(d, em).forEach((x) => { const o = d[x] || {}, fc = fcOf(x); if (!o.view || !fc) return; m.set(fc, m.get(fc) || !!o.edit); });
    return [...m].map(([fc, edit]) => ({ fc, edit }));
  }
  // 예전 방식(위 사용자에 체크 = 아래 조직을 봄)으로 저장된 조직도를 새 방식으로 바꿔 읽음 → 권한 결과는 같게
  function orgFromV1(d) {
    const nv = {}; Object.keys(d).forEach((x) => { nv[x] = { view: false, edit: false }; });
    Object.keys(d).forEach((em) => { const o = d[em]; if (!o.view) return; orgDesc(d, em).forEach((x) => { nv[x].view = true; if (o.edit) nv[x].edit = true; }); });
    Object.keys(d).forEach((x) => { d[x].view = nv[x].view; d[x].edit = nv[x].edit; });
    return d;
  }
  function renderOrg() {
    const d = orgDraft, root = orgRoot(), fcs = adminState.fcs || [];
    const fcOf = (em) => { const n = ((adminState.perms || {}).names || {})[em]; return n && fcs.includes(n) ? n : ""; };
    const node = (em) => {
      const isRoot = em === root, o = d[em] || {}, kids = orgKids(d, em), desc = isRoot ? [] : orgDesc(d, em);
      const subFcs = desc.map(fcOf).filter(Boolean);
      const gets = isRoot ? [] : orgGets(d, em, fcOf);
      const note = ""; // v373: ESL(관리자) 카드의 "전체 FC" 문구 삭제(2026-11-16 사용자 지시) // v339: '받은 권한' 줄 삭제(사용자 지시)
      const tSel = `<select class="org-title" data-org-t aria-label="${esc(orgName(em))} 직책">${ORG_TITLES.map((t) => `<option ${orgTitles[em] === t ? "selected" : ""}>${t}</option>`).join("")}</select>`;
      return `<li class="org-node${isRoot ? " root" : ""}${orgOpen.has(em) ? " open" : ""}" data-org-em="${esc(em)}"><div class="org-card">${isRoot ? "" : `<span class="org-grip" data-org-grip role="button" aria-label="${esc(orgName(em))} 끌어서 옮기기"><svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor" aria-hidden="true"><circle cx="9" cy="6" r="1.6"/><circle cx="15" cy="6" r="1.6"/><circle cx="9" cy="12" r="1.6"/><circle cx="15" cy="12" r="1.6"/><circle cx="9" cy="18" r="1.6"/><circle cx="15" cy="18" r="1.6"/></svg></span>`}<div class="org-main"><div class="org-nm"><button type="button" class="org-nmbtn" data-org-tg aria-expanded="${orgOpen.has(em) ? "true" : "false"}" aria-label="${esc(orgName(em))} 펼치기·접기">${esc(orgName(em))}</button> ${tSel}<button type="button" class="org-info" data-org-prof aria-label="${esc(orgName(em))} 사용자 정보">정보</button>${isRoot ? ' <span class="hint org-dev">(이 기기)</span>' : ""}</div><div class="org-mail hint">${isPend(em) ? "이메일 미입력" : esc(em)}</div>${isRoot ? "" : `<div class="org-cbs"><label><input type="checkbox" class="cb" data-org-v ${o.view ? "checked" : ""}> 보기</label><label><input type="checkbox" class="cb" data-org-e ${o.edit ? "checked" : ""}> 수정</label><i class="org-sep" aria-hidden="true"></i><label title="이 사용자 본인만(아래 조직에는 적용 안 됨)"><input type="checkbox" class="cb" data-org-s ${o.sens ? "checked" : ""}> 민감정보</label></div>`}${note ? `<div class="org-note hint">${note}</div>` : ""}</div>${isRoot ? orgSetHtml({ ...readStartPref(), rh: rhOn() }, true) : orgSetHtml((orgStartCache || {})[em], false)}</div>${kids.length ? `<ul class="org-kids">${kids.map(node).join("")}</ul>` : ""}</li>`;
    };
    const plan = orgPlan(), off = plan.changes.length + Object.keys(plan.sales).length; // 지금 권한이 조직도와 다른 계정 수
    $("adminBody").innerHTML = `<div class="card org-wrap"><div class="org-top"><span class="org-top-sp" aria-hidden="true"></span><h3 class="sec-title">조직도</h3><span class="org-top-sp r"><button type="button" class="ev-edit ev-done" data-org-save hidden ${orgDirty() || off ? "" : "disabled"}>저장</button></span></div><div class="org-back"><button type="button" class="ev-edit org-back-btn" data-org-back>${orgFromTeam ? "‹ 팀관리" : "‹ 관리자"}</button></div>
      <ul class="org-tree">${node(root)}</ul>${orgUsers().length ? "" : '<p class="hint">관리자 화면에서 사용자를 먼저 추가하세요.</p>'}</div>`;
  }
  let orgFromTeam = false; // v446: 조직도는 팀관리에서 열고, 뒤로 가면 팀관리로
  async function openOrgFromTeam() {
    orgFromTeam = true;
    show("adminView");
    $("mainNav").querySelectorAll("button[data-section]").forEach((b) => b.classList.toggle("on", b.dataset.section === "team"));
    $("adminBody").innerHTML = '<p class="hint pad">권한 정보를 불러오는 중…</p>';
    try { await Promise.all([loadAdminState(), loadSettings()]); orgOpen.clear(); openOrg(); }
    catch (e) { orgFromTeam = false; $("adminBody").innerHTML = ""; handleError(e); openTeam(); }
  }
  // v485(2026-11-19 사용자 지시): 조직도는 '도입자' 기준. 팀원 파일의 도입자사번을 따라 그 사람 밑에 두고, 도입자가 팀원(이 앱 사용자)에 없으면 관리자 바로 밑.
  //   여기서는 화면의 초안만 바꾸고, 권한이 걸려 있으므로 실제 적용은 지금처럼 아래 [저장]을 눌러 확인창을 거친 뒤에만 함
  async function orgApplyIntro() {
    const draft = orgDraft; if (!draft || !isAdmin()) return;
    let team = null;
    try { const dir = await stFindPath(TM_DIR); team = dir ? await stReadJsonIn(dir, TM_FILE) : null; } catch (e) { return; }
    if (orgDraft !== draft || !team || !Array.isArray(team.사람)) return;
    const root = orgRoot(), names = (adminState.perms || {}).names || {}, adm = (CFG.admins || []).map((a) => a.toLowerCase());
    const emOf = (q) => { const e = String(q.이메일 || "").toLowerCase(); if (e && adm.includes(e)) return root; if (e && draft[e]) return e; return Object.keys(draft).find((x) => names[x] && names[x] === q.이름) || ""; };
    const byNo = new Map(team.사람.map((q) => [String(q.사번), q]));
    let n = 0;
    team.사람.forEach((q) => {
      const em = emOf(q); if (!em || em === root || !draft[em]) return;
      const ip = byNo.get(String(q.도입자사번 || "")), pe = ip ? emOf(ip) : "";
      const par = pe && pe !== em && (pe === root || draft[pe]) ? pe : root;
      if (draft[em].parent !== par) { draft[em].parent = par; n++; }
    });
    Object.keys(draft).forEach((em) => { const seen = new Set([em]); let q = draft[em].parent; while (q && q !== root) { if (seen.has(q)) { draft[em].parent = root; break; } seen.add(q); q = (draft[q] || {}).parent; } }); // 돌고 도는 연결은 관리자 밑으로
    if (n) { renderOrg(); toast(`도입자 기준으로 ${n}명의 자리를 맞췄습니다. 아래 [저장]을 눌러야 적용됩니다`); }
  }
  function openOrg() {
    const d0 = orgNorm((adminState.perms || {}).org), d = (adminState.perms || {}).orgVer >= 2 ? d0 : orgFromV1(d0), sens = ((adminState.perms || {}).sensitive || []).map((e) => String(e).toLowerCase());
    Object.keys(d).forEach((em) => { d[em].sens = sens.includes(em); }); // v329: 민감정보 접근 = 권한 파일 sensitive (계정 카드 체크와 같은 값)
    orgSaved = JSON.stringify(d); orgDraft = JSON.parse(orgSaved);
    orgTitles = {}; [orgRoot(), ...Object.keys(d)].forEach((em) => { orgTitles[em] = titleOf(adminState.perms, em); }); orgTitlesSaved = JSON.stringify(orgTitles);
    renderOrg(); window.scrollTo(0, 0);
    orgApplyIntro();
    orgStartCache = null; orgLoadStarts();
  }
  document.addEventListener("click", async (ev) => {
    const t = ev.target && ev.target.closest && ev.target.closest("[data-org],[data-org-back],[data-org-save]");
    if (!t || !adminState) return;
    if (t.hasAttribute("data-org")) { orgOpen.clear(); openOrg(); return; } // v374: [조직도]를 열 때마다 모두 접힌 상태
    if (t.hasAttribute("data-org-back")) { if (orgDirty() && !confirm("저장하지 않은 조직도 변경을 버릴까요?")) return; orgDraft = null; if (orgFromTeam) { orgFromTeam = false; openTeam(); } else renderAdmin(); window.scrollTo(0, 0); return; }
    await saveOrg(t);
  });
  // v327: 조직도대로 '고객 페이지 접근 허용'을 맞춤. 보기 → 아래 조직 FC 보기만(FC:r), 수정 → 보기+수정(FC).
  //   조직도가 더한 권한만 orgApplied에 적어 두고, 조직도에서 빼면 그것만 되돌림(직접 체크한 권한은 그대로). Drive 공유는 savePerms가 함께 바꿈(확인창)
  function orgPlan() {
    const P = adminState.perms || {}, d = orgDraft, fcs = adminState.fcs || [], prev = P.orgApplied || {};
    const fcOf = (em) => { const n = (P.names || {})[em]; return n && fcs.includes(n) ? n : ""; };
    const changes = [], lines = [], applied = {}, sales = {}, salesApplied = {}, salesLines = [], prevS = P.orgSalesApplied || {};
    const lbl = (g) => grantFc(g) + (/:r$/.test(g) ? "(보기)" : "(수정)");
    orgUsers().forEach((em) => {
      const cur = ((P.users || {})[em] || []).slice(), own = fcOf(em);
      // v332(2026-11-15 사용자 지시 '정리'): 조직도가 유일한 기준. 고객 = 본인 FC(보기+수정) + 조직도 권한, 계정 카드에서 예전에 직접 준 권한은 뺌
      const gets = orgGets(d, em, fcOf).filter((g) => g.fc !== own), want = gets.map((g) => g.fc); // v335: 아래 사용자가 보기·수정을 켜서 나에게 준 FC
      let list = [...(own ? [own] : []), ...gets.map((g) => (g.edit ? g.fc : g.fc + ":r"))];
      if (!list.length) list = cur.slice(); // 본인 FC도 조직도 권한도 없으면 계정이 지워지지 않게 그대로 둠
      if (want.length) applied[em] = want.map((fc) => ({ fc, prev: null }));
      // 세일즈 페이지 접근 허용 = 조직도 아래 조직 FC (보기·수정 어느 쪽이든)
      const sCur = ((P.salesFcs || {})[em] || []).slice(), sList = want.slice();
      if (want.length) salesApplied[em] = want.slice();
      const sKey = (l) => [...l].sort().join("|");
      if (sKey(sList) !== sKey(sCur)) { sales[em] = sList; salesLines.push("· " + orgName(em) + " 영업: " + [...sList.filter((f) => !sCur.includes(f)).map((f) => "+" + f), ...sCur.filter((f) => !sList.includes(f)).map((f) => "−" + f)].join(" ")); }
      const key = (l) => [...l].sort().join("|");
      if (key(list) !== key(cur)) {
        const add = list.filter((g) => !cur.includes(g)), del = cur.filter((g) => !list.includes(g) && !add.some((x) => grantFc(x) === grantFc(g)));
        lines.push("· " + orgName(em) + ": " + [...add.map((g) => "+" + lbl(g)), ...del.map((g) => "−" + lbl(g))].join(" "));
        changes.push([em, list, undefined]);
      }
    });
    return { changes, lines, applied, sales, salesApplied, salesLines };
  }
  // 민감정보: 조직도 체크대로 권한 파일 sensitive를 맞춤(조직도 밖 계정은 그대로). org에는 sens를 저장하지 않음
  function orgSensApply(perms, draft) {
    const set = new Set((perms.sensitive || []).map((e) => String(e).toLowerCase()));
    Object.keys(draft).forEach((em) => { if (draft[em].sens) set.add(em); else set.delete(em); });
    perms.sensitive = [...set];
    perms.orgVer = 2;
    const org = {}; Object.keys(draft).forEach((em) => { const { sens, ...rest } = draft[em]; org[em] = rest; }); perms.org = org;
  }
  async function saveOrg(btn) {
    const { changes, lines, applied, sales, salesApplied, salesLines } = orgPlan(), draft = orgDraft;
    const salesApply = (perms) => { perms.salesFcs = perms.salesFcs || {}; Object.entries(sales).forEach(([em, l]) => { if (l.length) perms.salesFcs[em] = l; else delete perms.salesFcs[em]; }); perms.orgSalesApplied = salesApplied; };
    const was = new Set(((adminState.perms || {}).sensitive || []).map((e) => String(e).toLowerCase()));
    const sAdd = Object.keys(draft).filter((em) => draft[em].sens && !was.has(em)), sDel = Object.keys(draft).filter((em) => !draft[em].sens && was.has(em));
    const sLine = sAdd.length || sDel.length ? ["· 민감정보 접근(주민번호·계좌): " + [...sAdd.map((em) => "+" + orgName(em)), ...sDel.map((em) => "−" + orgName(em))].join(" ")] : [];
    const titles = { ...orgTitles };
    const done = () => { orgSaved = JSON.stringify(draft); orgTitlesSaved = JSON.stringify(titles); adminState.perms = { ...adminState.perms, titles }; toast("조직도를 저장했어요"); renderOrg(); };
    if (!changes.length) { // 고객 접근 권한은 그대로: 조직도(+민감정보)만 기록
      if ((sLine.length || salesLines.length) && !confirm(["조직도를 저장합니다.", ...(salesLines.length ? ["세일즈 페이지 접근 허용을 바꿉니다."] : []), ...salesLines, ...sLine].join("\n"))) return;
      btn.disabled = true; btn.textContent = "저장 중…";
      try { const perms = await readPerms(); orgSensApply(perms, draft); salesApply(perms); perms.orgApplied = applied; perms.titles = titles; await writePerms(perms); adminState.perms = { ...adminState.perms, org: perms.org, orgApplied: applied, sensitive: perms.sensitive, salesFcs: perms.salesFcs, orgSalesApplied: perms.orgSalesApplied }; done(); }
      catch (e) { handleError(e); renderOrg(); }
      return;
    }
    await savePerms(changes, { btn, pre: ["조직도대로 고객 페이지 접근 권한을 바꿉니다.", ...lines, ...(salesLines.length ? ["세일즈 페이지 접근 허용도 바꿉니다."] : []), ...salesLines, ...sLine], mutate: (perms) => { orgSensApply(perms, draft); salesApply(perms); perms.orgApplied = applied; perms.titles = titles; }, after: done });
  }
  document.addEventListener("change", (ev) => {
    const ts = ev.target && ev.target.closest && ev.target.closest("[data-org-t]");
    if (ts && orgDraft) { orgTitles[ts.closest(".org-node").dataset.orgEm] = ts.value; renderOrg(); return; }
    const cb = ev.target && ev.target.closest && ev.target.closest("[data-org-v],[data-org-e],[data-org-s]");
    if (!cb || !orgDraft) return;
    const em = cb.closest(".org-node").dataset.orgEm, o = orgDraft[em]; if (!o) return;
    if (cb.hasAttribute("data-org-s")) o.sens = cb.checked;
    else if (cb.hasAttribute("data-org-v")) { o.view = cb.checked; if (!o.view) o.edit = false; } else { o.edit = cb.checked; if (o.edit) o.view = true; } // 수정은 보기를 포함
    renderOrg();
  });
  // 끌어서 옮기기 (손가락·마우스 모두: pointer 이벤트)
  document.addEventListener("pointerdown", (ev) => {
    const g = ev.target && ev.target.closest && ev.target.closest("[data-org-grip]");
    if (!g || !orgDraft) return;
    ev.preventDefault();
    const li = g.closest(".org-node"), em = li.dataset.orgEm, bad = new Set([em, ...orgDesc(orgDraft, em)]);
    const card = li.querySelector(".org-card"), r = card.getBoundingClientRect(), dx = ev.clientX - r.left, dy = ev.clientY - r.top;
    const ghost = card.cloneNode(true); ghost.classList.add("org-ghost"); ghost.style.width = r.width + "px"; document.body.appendChild(ghost);
    li.classList.add("dragging");
    let tgt = null;
    const clear = () => document.querySelectorAll(".org-node.drop").forEach((x) => x.classList.remove("drop"));
    const move = (e) => {
      ghost.style.transform = `translate(${e.clientX - dx}px, ${e.clientY - dy}px)`;
      const el = document.elementFromPoint(e.clientX, e.clientY), n = el && el.closest && el.closest(".org-node");
      const t = n && !bad.has(n.dataset.orgEm) ? n.dataset.orgEm : null;
      if (t !== tgt) { clear(); tgt = t; if (t) n.classList.add("drop"); }
      const H = window.innerHeight; if (e.clientY < 80) window.scrollBy(0, -14); else if (e.clientY > H - 110) window.scrollBy(0, 14);
    };
    const up = () => {
      document.removeEventListener("pointermove", move); document.removeEventListener("pointerup", up); document.removeEventListener("pointercancel", up);
      ghost.remove(); li.classList.remove("dragging"); clear();
      if (tgt && orgDraft[em] && orgDraft[em].parent !== tgt) { orgDraft[em].parent = tgt; renderOrg(); }
    };
    document.addEventListener("pointermove", move); document.addEventListener("pointerup", up); document.addEventListener("pointercancel", up);
    move(ev);
  });

  // ---------- v340 사용자 정보 (조직도에서 이름을 누르면): 개인 정보는 관리자만 보는 '관리자전용' 폴더에 저장(사용자에게 공유되는 설정 폴더가 아님) ----------
  const PROF_DIR = "관리자전용", PROF_FILE = "shlife-사용자정보.json";
  const PROF_FIELDS = [["empNo", "사번", "text", "예: 1234567"], ["carrier", "통신사", "sel", ["", "SKT", "KT", "LG U+", "알뜰폰"]], ["phone", "연락처", "tel", "010-0000-0000"], ["addr", "주소", "text", ""], ["joined", "입사일", "date", ""], ["memo", "메모", "area", ""]];
  let profDirC = null, profEm = null, profAll = null;
  async function profFolder(create) {
    if (profDirC) return profDirC;
    const par = await stFindPathCached(CFG.appFolder.split("/").slice(0, -1).join("/"));
    if (!par || !par.id) throw new Error("신한FC 앱 폴더를 찾을 수 없습니다.");
    const r = await gdQuery("'" + qv(par.id) + "' in parents and name='" + qv(PROF_DIR) + "' and mimeType='" + GD_FOLDER + "' and trashed=false");
    if (r.length) return (profDirC = { id: r[0].id });
    if (!create) return null;
    const mk = await gdFetch(GDAPI + "/files?fields=id&supportsAllDrives=true", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ name: PROF_DIR, mimeType: GD_FOLDER, parents: [par.id] }) });
    return (profDirC = { id: mk.id });
  }
  async function readProfiles() { const d = await profFolder(false); if (!d) return {}; const v = await stReadJsonIn(d, PROF_FILE); return v && typeof v === "object" ? v : {}; }
  async function writeProfiles(all) { const d = await profFolder(true); await stWriteJsonIn(d, PROF_FILE, all); }
  // v342(2026-11-15 사용자 지시): 사용자 정보는 팝업. 조직도에서 이름을 누르거나 관리자 화면 계정 목록에서 사용자를 누르면 열림
  let profFrom = "org", profStart = null;
  let profNew = null; // v343: 새 사용자 추가 중이면 { name }
  let profEdit = false; // v349: 보기 모드(기본) / 수정 모드([수정] 버튼)
  async function openOrgProfile(em, from, newOpt, edit) {
    profEm = em; profFrom = from || "org"; profNew = newOpt || null; profEdit = !!newOpt || !!edit;
    try { history.pushState({ v: "orgprof" }, ""); } catch (e) { /* 무시 */ }
    let m = $("profModal");
    if (!m) { m = document.createElement("div"); m.id = "profModal"; m.className = "pf-modal"; m.innerHTML = '<div class="pf-sheet card" id="profSheet" role="dialog" aria-modal="true" aria-label="사용자 정보"></div>'; document.body.appendChild(m); m.addEventListener("click", (ev) => { if (ev.target === m) tryCloseProfile(); }); }
    m.hidden = false; document.documentElement.classList.add("pf-open");
    $("profSheet").innerHTML = '<p class="hint">사용자 정보를 불러오는 중…</p>';
    const P = adminState.perms || {}, isRoot = em === orgRoot();
    const [all, st] = await Promise.all([readProfiles().catch((e) => { console.error(e); toast("사용자 정보 파일을 읽지 못했습니다: " + e.message); return {}; }),
      isRoot ? Promise.resolve({ ...readStartPref(), rh: rhOn() }) : readStartShared(P, em, adminState.fcIds).catch(() => null)]);
    profAll = all;
    profStart = st || (P.starts || {})[em] || null;
    if (profEm === em) { renderOrgProfile(); profApplyMode(); }
  }
  function closeOrgProfile() {
    const m = $("profModal"); if (m) { m.hidden = true; $("profSheet").innerHTML = ""; }
    document.documentElement.classList.remove("pf-open");
    const from = profFrom; profEm = null; profNew = null;
    if (from === "org" && orgDraft && document.querySelector(".org-wrap")) renderOrg(); else if (from === "admin" && !$("adminView").hidden) renderAdmin();
  }
  function tryCloseProfile() {
    const b = document.querySelector("[data-prof-save]");
    if (b && !b.disabled && !confirm(profNew ? "사용자 추가를 취소할까요?" : "저장하지 않은 변경을 버릴까요?")) return;
    try { history.back(); } catch (e) { closeOrgProfile(); }
  }
  function renderOrgProfile() {
    const em = profEm, P = adminState.perms || {}, root = orgRoot(), isRoot = em === root, pr = (profAll || {})[em] || {}, isNew = !!profNew;
    const d = orgDraft || (P.orgVer >= 2 ? orgNorm(P.org) : orgFromV1(orgNorm(P.org))), o = d[em] || {}, nm = isNew ? profNew.name : (P.names || {})[em] || "", title = (!isNew && orgTitles && orgTitles[em]) || titleOf(P, em);
    const fld = ([k, label, type, ex]) => {
      const v = pr[k] || "";
      const inp = type === "sel" ? `<select data-pf="${k}">${ex.map((x) => `<option value="${esc(x)}" ${v === x ? "selected" : ""}>${x || "선택"}</option>`).join("")}</select>`
        : type === "area" ? `<textarea data-pf="${k}" rows="3" maxlength="1000">${esc(v)}</textarea>`
        : `<input type="${type}" data-pf="${k}" value="${esc(v)}" placeholder="${esc(ex || "")}" autocomplete="off">`;
      return `<label class="pf-row"><span>${label}</span>${inp}</label>`;
    };
    const par = isRoot || isNew ? root : o.parent || root, kids = isNew ? [] : orgKids(d, em).map(orgName);
    const lst = (l) => (l.length ? l.map(esc).join(", ") : '<span class="hint">없음</span>');
    const st = profStart || { page: "customer", ai: false, rh: null };
    const onoff = (x) => (x === true ? "켬" : x === false ? "끔" : '<span class="hint">기록 없음</span>');
    const cust = isRoot ? ["전체 FC"] : ((P.users || {})[em] || []).map((g) => grantFc(g) + (/:r$/.test(g) ? "(보기)" : "(수정)"));
    const own = (P.names || {})[em], sales = isRoot ? ["전체(사용자 선택)"] : [...new Set([own, ...((P.salesFcs || {})[em] || [])])].filter((f) => f && (adminState.fcs || []).includes(f));
    const kv = (k, v) => `<div class="pf-kv"><span>${k}</span><b>${v}</b></div>`;
    $("profSheet").innerHTML = `<div class="pf-top"><h3 class="sec-title">${isNew ? "사용자 추가" : profEdit ? "사용자 정보 수정" : "사용자 정보"}</h3>${profEdit ? `<button type="button" class="ev-edit ev-done" data-prof-save ${isNew ? "" : "disabled"}>저장</button><button type="button" class="ev-edit" data-prof-cancel>취소</button>` : `<button type="button" class="ev-edit" data-prof-edit>수정</button><button type="button" class="ev-edit" data-prof-back>취소</button>`}</div>
      <div class="pf-head"><b>${esc(nm || orgName(em))}</b> <span class="tag rank-tag">${esc(title)}</span>${isRoot ? ' <span class="hint">(이 기기)</span>' : ""}</div>
      <div class="pf-form">
        <label class="pf-row"><span>이름</span><input type="text" data-pf-name value="${esc(nm)}" maxlength="20" placeholder="이름(FC명)"></label>
        ${fld(PROF_FIELDS[0])}
        ${isPend(em) && !isRoot ? `<label class="pf-row"><span>이메일</span><input type="email" data-pf-email value="" placeholder="구글 계정 이메일 (예: name@gmail.com)" autocomplete="off"></label>` : `<label class="pf-row"><span>이메일</span><input type="email" value="${esc(em)}" readonly class="ro" title="로그인 계정이라 여기서 바꿀 수 없습니다"></label>`}
        ${PROF_FIELDS.slice(1, 5).map(fld).join("")}
        <label class="pf-row"><span>직책</span><select data-pf-title>${ORG_TITLES.map((t) => `<option ${title === t ? "selected" : ""}>${t}</option>`).join("")}</select></label>
        <label class="pf-row"><span>일정 구분</span><input type="text" data-pf-pfx value="${esc(nm ? (P.prefixes || {})[nm] || "" : "")}" maxlength="10" placeholder="예: 최" autocomplete="off"></label>
        <div class="pf-row"${START_FIXED ? " hidden" : ""}><span>시작 화면</span><span class="pf-start"><select data-pf-start aria-label="시작 화면">${START_PAGES.map(([k, l]) => `<option value="${k}" ${st.page === k ? "selected" : ""}>${l}</option>`).join("")}</select><label class="pf-ai"><input type="checkbox" class="cb" data-pf-ai ${st.ai ? "checked" : ""}> AI 바로 열기</label></span></div>
        ${fld(PROF_FIELDS[5])}
      </div>
      ${isNew ? `<p class="hint pf-note">[저장]하면 사용자로 등록되고 조직도에서 관리자 바로 밑에 들어갑니다. 이름·이메일·사번이 모두 있고 이름이 FC 폴더 이름과 같을 때만 본인 FC(보기+수정)와 폴더 공유가 함께 됩니다.</p>` : `<h4 class="pf-sub">신한FC 설정</h4>
      <div class="pf-kvs">
        ${kv("조직 위치", isRoot ? "맨 위(관리자)" : esc(orgName(par)) + " 밑")}
        ${kv("아래 사용자", lst(kids))}
        ${isRoot ? "" : kv("위로 주는 권한", o.view ? (o.edit ? "보기·수정" : "보기") : '<span class="hint">없음</span>')}
        ${kv("민감정보", isRoot ? "전체(관리자)" : o.sens || (P.sensitive || []).includes(em) ? "허용" : '<span class="hint">없음</span>')}
        ${kv("고객 페이지 접근", lst(cust))}
        ${kv("세일즈 페이지 접근", lst(sales))}
        ${kv("오른손 모드", onoff(st.rh))}
        ${pr.at ? kv("정보 저장", esc(String(pr.at).slice(0, 16).replace("T", " "))) : ""}
      </div>`}
      <p class="hint pf-note">권한·민감정보·위치는 조직도에서 바꿉니다. 오른손 모드는 그 사용자 폰의 설정입니다. 개인 정보는 관리자만 보는 '관리자전용' 폴더에 저장됩니다.</p>`;
  }
  function profApplyMode() { const sh = $("profSheet"); if (!sh) return; sh.classList.toggle("pf-view", !profEdit); sh.querySelectorAll(".pf-form input, .pf-form select, .pf-form textarea").forEach((x) => { if (!x.classList.contains("ro")) x.disabled = !profEdit; }); }
  async function saveOrgProfile(btn) {
    const em = profEm, box = $("profSheet"), P = adminState.perms || {}, isRoot = em === orgRoot();
    const val = (sel) => { const x = box.querySelector(sel); return x ? String(x.value || "").trim() : ""; };
    const prof = {}; PROF_FIELDS.forEach(([k]) => { prof[k] = val(`[data-pf="${k}"]`); });
    const newName = val("[data-pf-name]"), newTitle = val("[data-pf-title]") || titleOf(P, em), oldName = (P.names || {})[em] || "";
    const newPfx = fixPrefix(val("[data-pf-pfx]")), oldPfx = (P.prefixes || {})[oldName] || "";
    const aiBox = box.querySelector("[data-pf-ai]"), newStart = { page: val("[data-pf-start]") || "customer", ai: !!(aiBox && aiBox.checked) };
    const st0 = profStart || { page: "customer", ai: false, rh: null };
    if (!newName) return toast("이름을 적어 주세요.");
    const newEm = isPend(em) ? val("[data-pf-email]").toLowerCase() : "";
    if (newEm) { // v370: 이메일 미입력 사용자에 이메일을 채움 → 등록 키를 이메일로 바꾸고 공유 후 나머지 정보 저장
      if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(newEm)) return toast("이메일 형식을 확인해 주세요.");
      if ((CFG.admins || []).map((a) => a.toLowerCase()).includes(newEm)) return toast("관리자는 이미 모든 폴더를 볼 수 있습니다.");
      if (((P.users || {})[newEm]) || ((P.archived || {})[newEm])) return toast("이미 있는(또는 보관된) 사용자입니다.");
      if (!confirm(`${newName} 사용자의 이메일을 ${newEm}(으)로 정합니다.` + ((adminState.fcs || []).includes(newName) ? (prof.empNo ? ` 이름·이메일·사번이 모두 있어 본인 FC '${newName}' 폴더도 공유합니다.` : " 사번이 아직 없어 FC 폴더 공유는 사번을 넣고 저장할 때 합니다.") : ""))) return;
      btn.disabled = true; btn.textContent = "저장 중…";
      try {
        const perms = await readPerms(); rekeyPerms(perms, em, newEm); await writePerms(perms); adminState.perms = perms;
        const all0 = await readProfiles(); if (all0[em]) { all0[newEm] = all0[em]; delete all0[em]; await writeProfiles(all0); }
        if (orgDraft) { rekeyOrg(orgDraft, em, newEm); try { orgSaved = JSON.stringify(rekeyOrg(JSON.parse(orgSaved), em, newEm)); } catch (e) { /* 무시 */ } }
        if (orgTitles && orgTitles[em]) { orgTitles[newEm] = orgTitles[em]; delete orgTitles[em]; try { const t = JSON.parse(orgTitlesSaved || "{}"); if (t[em]) { t[newEm] = t[em]; delete t[em]; orgTitlesSaved = JSON.stringify(t); } } catch (e) { /* 무시 */ } }
        if (orgOpen.has(em)) { orgOpen.delete(em); orgOpen.add(newEm); }
        orgStartCache = null; profEm = newEm;
      } catch (e) { btn.disabled = false; btn.textContent = "저장"; handleError(e); return; }
      return saveOrgProfile(btn);
    }
    btn.disabled = true; btn.textContent = "저장 중…";
    const after = () => { if (orgTitles) { orgTitles[em] = newTitle; try { const t = JSON.parse(orgTitlesSaved || "{}"); t[em] = newTitle; orgTitlesSaved = JSON.stringify(t); } catch (e) { /* 무시 */ } } profStart = { ...st0, ...newStart }; toast("사용자 정보를 저장했습니다."); profEdit = false; renderOrgProfile(); profApplyMode(); };
    try {
      const all = await readProfiles(); all[em] = { ...prof, at: new Date().toISOString(), by: me() }; await writeProfiles(all); profAll = all;
      const nameCh = newName !== oldName, titleCh = newTitle !== titleOf(P, em), pfxCh = newPfx !== oldPfx || (nameCh && !!oldPfx);
      const startCh = newStart.page !== st0.page || newStart.ai !== !!st0.ai;
      let startToPerms = !!profNew; // 새 사용자는 FC 폴더 연결 전이라 시작 설정을 권한 파일에
      if (startCh && !profNew) {
        if (isRoot) writeStartPref(newStart); // 관리자 본인 = 이 기기
        else { let ok = false; try { ok = await writeStartShared(await readPerms(), em, { ...newStart, rh: st0.rh }, adminState.fcIds); } catch (e) { console.error(e); } startToPerms = !ok; orgStartCache = null; }
      }
      const mutate = (perms) => {
        perms.titles = { ...(perms.titles || {}), [em]: newTitle };
        if (pfxCh) { const nx = { ...(perms.prefixes || {}) }; if (oldName && nameCh) delete nx[oldName]; if (newPfx) nx[newName] = newPfx; else delete nx[newName]; perms.prefixes = nx; }
        if (startCh && !isRoot) { const nx = { ...(perms.starts || {}) }; if (startToPerms) nx[em] = newStart; else delete nx[em]; perms.starts = nx; }
      };
      if (profNew) { // v343: 새 사용자 등록 (확인창 → 권한 파일 + 본인 FC 폴더 공유)
        const own = (adminState.fcs || []).includes(newName) && prof.empNo ? [newName] : [];
        await savePerms([[em, own, newName]], { btn, pre: [`사용자 ${newName} (${em})를 추가합니다.` + (own.length ? ` 이름·이메일·사번이 모두 있어 본인 FC '${newName}' 보기+수정과 폴더 공유를 함께 합니다.` : " (이름·이메일·사번이 모두 채워지고 같은 이름의 FC 폴더가 있어야 폴더가 공유됩니다 — 지금은 권한 없이 등록, 조직도에서 받은 권한이 생기면 열림)")],
          mutate: (perms) => { mutate(perms); if (perms.archived) delete perms.archived[em]; }, after: () => { profNew = null; orgStartCache = null; after(); } });
        if (profNew) { btn.disabled = false; btn.textContent = "저장"; } // 확인창에서 취소
        return;
      }
      const ownFc2 = (adminState.fcs || []).includes(newName) ? newName : "";
      if (!isRoot && !isPend(em) && ownFc2 && prof.empNo && !((P.users || {})[em] || []).some((g) => grantFc(g) === ownFc2)) { // v371(2026-11-16 사용자 지시): 이름·이메일·사번이 모두 채워지면 그때 본인 FC 폴더 공유
        await savePerms([[em, [...((P.users || {})[em] || []), ownFc2], newName]], { btn, pre: [`이름·이메일·사번이 모두 채워졌습니다. 본인 FC '${ownFc2}' 폴더를 공유합니다.`], mutate, after });
        return;
      }
      if (nameCh && !isRoot) { // 이름(FC명)을 바꾸면 본인 FC 연결·공유가 달라질 수 있어 계정 저장과 같은 길로
        await savePerms([[em, ((P.users || {})[em] || []).slice(), newName]], { btn, pre: [`이름을 '${oldName}' → '${newName}'(으)로 바꿉니다. 이름이 FC 폴더 이름과 같아야 본인 FC가 연결됩니다.`], mutate, after });
        return;
      }
      if (nameCh || titleCh || pfxCh || (startCh && !isRoot)) {
        const perms = await readPerms(); perms.names = perms.names || {};
        if (nameCh) perms.names[em] = newName;
        mutate(perms);
        await writePerms(perms); adminState.perms = { ...adminState.perms, names: perms.names, titles: perms.titles, prefixes: perms.prefixes, starts: perms.starts };
        permPrefixCache = { ...(perms.prefixes || {}) };
      }
      after();
    } catch (e) { btn.disabled = false; btn.textContent = "저장"; handleError(e); }
  }
  document.addEventListener("click", (ev) => {
    const t = ev.target && ev.target.closest && ev.target.closest("[data-org-tg],[data-org-prof],[data-prof-back],[data-prof-save],[data-prof-edit],[data-prof-cancel]");
    if (!t || !adminState) return;
    if (t.hasAttribute("data-prof-cancel")) { // v352: 수정 취소 → 바꾼 내용 버리고 보기 모드(새 사용자 추가면 팝업 닫기)
      if (profNew) { tryCloseProfile(); return; }
      const b = document.querySelector("[data-prof-save]");
      if (b && !b.disabled && !confirm("수정한 내용을 버릴까요?")) return;
      profEdit = false; renderOrgProfile(); profApplyMode(); return;
    }
    if (t.hasAttribute("data-prof-edit")) { profEdit = true; renderOrgProfile(); profApplyMode(); const f = $("profSheet").querySelector("[data-pf-name]"); if (f) f.focus(); return; }
    if (t.hasAttribute("data-org-tg")) { const li = t.closest(".org-node"), em = li.dataset.orgEm, on = !orgOpen.has(em); on ? orgOpen.add(em) : orgOpen.delete(em); li.classList.toggle("open", on); t.setAttribute("aria-expanded", on ? "true" : "false"); return; }
    if (t.hasAttribute("data-org-prof")) { openOrgProfile(t.closest(".org-node").dataset.orgEm, "org"); return; }
    if (t.hasAttribute("data-prof-back")) { tryCloseProfile(); return; }
    saveOrgProfile(t);
  });
  const profDirty = (ev) => { if (profEm && ev.target && ev.target.closest && ev.target.closest(".pf-form")) { const b = document.querySelector("[data-prof-save]"); if (b) b.disabled = false; } };
  document.addEventListener("input", profDirty);
  document.addEventListener("change", profDirty);
  window.addEventListener("popstate", () => { if (profEm) closeOrgProfile(); }); // 폰의 뒤로가기 → 팝업 닫기

  function renderAdmin() {
    const { fcs, perms } = adminState;
    // v372(2026-11-16 사용자 지시): 사용자 순서 = 1순위 직급(ESL > SSL > GSL > FC), 2순위 사번(작은 것이 위, 사번 모르면 뒤), 그다음 이름
    const noKey = (em) => { const n = fcNoOf((perms.names || {})[em]); return n && /^\d+$/.test(n) ? Number(n) : Infinity; };
    const users = [...new Set([...Object.keys(perms.users), ...extraUsers])].sort((x, y) => ORG_TITLES.indexOf(titleOf(perms, x)) - ORG_TITLES.indexOf(titleOf(perms, y)) || (noKey(x) === noKey(y) ? 0 : noKey(x) < noKey(y) ? -1 : 1) || String((perms.names || {})[x] || x).localeCompare(String((perms.names || {})[y] || y), "ko"));
    const listOf = (em) => perms.users[em] || [];
    const isAll = (em) => listOf(em).includes("*");
    const nameOfEm = (em) => (perms.names || {})[em] || "";
    const ownFc = (em) => (nameOfEm(em) && fcs.includes(nameOfEm(em)) ? nameOfEm(em) : ""); // 본인 FC: 항상 보기+수정 (바꿀 수 없음)
    const has = (em, fc) => isAll(em) || fc === ownFc(em) || listOf(em).some((g) => grantFc(g) === fc);
    const canEdit = (em, fc) => isAll(em) || fc === ownFc(em) || listOf(em).includes(fc); // 'FC' = 수정 가능, 'FC:r' = 보기만
    const effList = (em) => (ownFc(em) && !listOf(em).some((g) => grantFc(g) === ownFc(em) && g === ownFc(em)) ? [...listOf(em).filter((g) => grantFc(g) !== ownFc(em)), ownFc(em)] : listOf(em));
    // FC 순서: 본인 → 직급(ESL > SSL > GSL > FC) → 사번순 (config.fcOrder), 명단에 없는 FC는 맨 뒤 이름순
    const fcRank = (fc) => { const o = CFG.fcOrder || [], i = o.indexOf(fc); return i < 0 ? 999 : i; };
    const fcsFor = (em) => [...fcs].sort((a, b) => (a === ownFc(em) ? -1 : b === ownFc(em) ? 1 : 0) || fcRank(a) - fcRank(b) || a.localeCompare(b, "ko"));
    // 세일즈 페이지(이달목표) FC 선택창에 나올 FC: 본인 FC + 체크한 FC (관리자 계정도 같은 기준)
    const salesHtml = (em) => `<div class="perm-sales"><p class="hint perm-lbl">• 세일즈 페이지 접근 허용</p><div class="perm-fcs">${fcsFor(em).map((fc) => { const own = !!nameOfEm(em) && fc === nameOfEm(em); return `<span class="chip"><label><input type="checkbox" class="cb" data-sf="${esc(em)}|${esc(fc)}" ${own ? 'data-own="1" checked disabled' : ((perms.salesFcs || {})[em] || []).includes(fc) ? "checked" : ""} aria-label="${esc(em)} 영업 ${esc(fc)}"> ${esc(fc)}${own ? ' <small class="hint">(본인)</small>' : ""}</label></span>`; }).join("")}</div></div>`;
    const ACC = GD ? "구글 계정" : "아이디";
    const pfxHtml = (em) => { const nm = nameOfEm(em); return `<label class="pfx-l"><span>일정 구분</span><input type="text" class="perm-name" data-pfx="${esc(em)}" maxlength="10" value="${esc(nm ? (perms.prefixes || {})[nm] || "" : "")}" placeholder="예: 최" autocomplete="off" ${nm ? "" : "disabled"} aria-label="${esc(em)} 일정 구분"><button type="button" class="secondary" data-pfx-save="${esc(em)}" disabled>저장</button></label>${nm ? "" : '<p class="hint">이름(FC명)을 먼저 입력·저장하면 정할 수 있습니다.</p>'}`; };

    // 순서: 사용 FC 현황 → 사용자 추가 → 설정  (Drive 공유는 '변경 저장' 때 자동 처리, 표에는 표시 안 함)
    const personal = featCardHtml() + `<div class="card" id="adminCalCard"><details class="fold" id="gFold"><summary class="sec-title">구글 캘린더</summary>
        <div id="gCalBox"><p class="hint">캘린더 목록을 불러오는 중…</p></div>
        <div class="file-actions"><button type="button" class="primary" id="gCalSave" disabled>저장</button></div></details></div>` + startCardHtml() + smsCardHtml() + rhCardHtml() + installCardHtml();
    $("adminBody").innerHTML = personal + `<h2 class="admin-sec">관리자</h2>
      <div class="card ntf-entry"><div><h3 class="sec-title">전체 알림</h3><p class="hint">FC 로그인 직후 보여 줄 알림(날짜·요일·시스템)을 정합니다.</p></div><button type="button" class="secondary" data-ntf-open>열기 ›</button></div>
      <div class="card">
        <h3 class="sec-title org-h">FC 사용 설정</h3>
        <p class="hint org-note-top">고객·세일즈 페이지 접근과 민감정보 권한은 팀관리 메뉴의 [조직도]에서 정합니다.</p>
        ${(() => {
          const sh = adminState.shares, adm = (CFG.admins || []).map((a) => a.toLowerCase());
          const hits = ((sh && sh.anc) || []).flatMap((a) => a.perms.filter((p) => p.role !== "owner" && !adm.includes(p.email)).map((p) => esc((perms.names || {})[p.email] || p.email) + " ← '" + esc(a.name) + "' 폴더(" + p.role + ")"));
          return hits.length ? `<div class="perm-warn"><b>⚠ 상위 폴더 공유 감지</b><br>${hits.join("<br>")}<br><span class="hint">이 계정은 FC 체크와 상관없이 상위 폴더 전체를 Drive에서 열 수 있습니다. 해당 계정 카드의 [저장]을 누르면 상위 공유를 해제하고 FC 폴더만 공유합니다.</span></div>` : "";
        })()}
        <div class="perm-list">
          ${(CFG.admins || []).map((a) => `<div class="perm-admin"><span class="perm-em">${esc(a)}</span> <span class="tag">관리자</span>
            <span class="adm-name"><input type="text" class="perm-name" data-admin-name="${esc(a.toLowerCase())}" placeholder="내 FC 이름(예: 김현성)" maxlength="20" value="${esc(nameOfEm(a.toLowerCase()))}" aria-label="관리자 FC 이름"><span class="tag rank-tag">${titleOf(perms, a.toLowerCase())}</span><button type="button" class="secondary" data-admin-save="${esc(a.toLowerCase())}" disabled>저장</button></span></div><div class="perm-pfx perm-pfx-adm">${pfxHtml(a.toLowerCase())}</div>`).join("")}
          ${users.map((em) => `<div class="perm-row pf-click" data-em="${esc(em)}">
            <div class="perm-head"><button type="button" class="perm-fold" data-pfold aria-expanded="${permOpen.has(em)}" aria-label="${esc(nameOfEm(em) || em)} 펼치기·접기"><svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M9 6l6 6-6 6"/></svg></button>
              <span class="perm-id"><span class="perm-nmrow"><b class="perm-nm-txt">${esc(nameOfEm(em) || em.split("@")[0])}</b><input type="text" class="perm-name" data-name placeholder="이름(FC명) 입력" value="${esc(nameOfEm(em))}" maxlength="20" aria-label="${esc(em)} 이름"><span class="tag rank-tag">${titleOf(perms, em)}</span></span><span class="perm-mail">${esc(em)}</span></span>
              <span class="perm-btns"><button type="button" class="lg-tg${loginOn(perms, em) ? " on" : ""}" data-login="${esc(em)}" aria-pressed="${loginOn(perms, em)}" title="신한FC 로그인 승인/거부">${loginOn(perms, em) ? "승인" : "거부"}</button><button type="button" class="secondary" data-rowedit="${esc(em)}">수정</button><button type="button" class="del-btn" data-del="${esc(em)}" aria-label="${esc(em)} 삭제">삭제</button><button type="button" class="secondary" data-rowcancel="${esc(em)}" hidden>취소</button><button type="button" class="primary row-save" data-rowsave="${esc(em)}" hidden>저장</button></span>
            </div>
            <div class="perm-body"><div class="perm-pfx">${pfxHtml(em)}</div>${startRowHtml(perms, em)}
            <div class="perm-moved" hidden>
            <div class="perm-sens"><label class="chip"><input type="checkbox" class="cb" data-sens="${esc(em)}" ${(perms.sensitive || []).includes(em) ? "checked" : ""} aria-label="${esc(em)} 민감정보 접근"> 민감정보 접근 허용</label></div>
            ${salesHtml(em)}
            <p class="hint perm-lbl">• 고객 페이지 접근 허용</p>
            <div class="perm-fcs fc-lines">
              <label class="chip chip-all" hidden title="전체(*)는 폐지 — 해제하고 FC를 체크하세요"><input type="checkbox" class="cb" data-all ${isAll(em) ? "checked" : ""} aria-label="${esc(em)} 전체"> 전체(폐지 예정)</label>
              <div class="fc-line fc-head"><span class="fc-nm">FC</span><span class="fc-c">보기</span><span class="fc-c">수정</span></div>
              ${fcsFor(em).map((fc) => { const own = fc === ownFc(em); return `<div class="fc-line${own ? " own" : ""}"><span class="fc-nm">${esc(fc)}${own ? ' <small class="hint">(본인)</small>' : ""}</span><label class="fc-c"><input type="checkbox" class="cb" data-fc="${esc(fc)}" ${has(em, fc) ? "checked" : ""} ${own ? 'data-own="1" disabled' : ""} aria-label="${esc(em)} ${esc(fc)} 보기"></label><label class="fc-c"><input type="checkbox" class="cb ro" data-edit="${esc(fc)}" ${has(em, fc) && canEdit(em, fc) ? "checked" : ""} ${own || isAll(em) ? 'data-own="1" disabled' : ""} aria-label="${esc(em)} ${esc(fc)} 수정"></label></div>`; }).join("")}
            </div></div></div>
          </div>`).join("")}
        </div>
        ${users.length ? "" : '<p class="hint">아직 권한을 준 사용자가 없습니다. 아래에서 이메일을 추가하세요.</p>'}
        ${(() => { const ar = Object.entries(perms.archived || {}); return ar.length ? `<details class="fold arch-users"><summary>보관된 사용자 ${ar.length}</summary>${ar.map(([em, a]) => `<div class="arch-row"><span><b>${esc(a.name || em.split("@")[0])}</b> <small class="hint">${esc(em)} · 보관 ${esc(String(a.at || "").slice(0, 10))}</small></span><button type="button" class="ev-edit" data-unarch="${esc(em)}">원복</button></div>`).join("")}</details>` : ""; })()}
        <div class="add-user"><h4 class="sub-title">사용자 추가</h4>
        <form id="addUserForm" class="search">
          <input id="newName" type="text" placeholder="이름(FC명)" autocomplete="off" maxlength="20" style="max-width:130px">
          <input id="newUser" type="email" placeholder="${GD ? "구글 계정 이메일 (비우면 나중에 입력)" : "Microsoft 계정 이메일"}" autocomplete="off">
          <button class="primary" type="submit">추가</button>
        </form>
        </div>
        <div class="file-actions">
          <button class="secondary" id="resetPerms" disabled>되돌리기</button>
          <button class="primary" id="savePerms" disabled>변경 저장</button>
        </div>
      </div>` + supportCardHtml() + smsTeamCardHtml() + settingsCardHtml(); // 오른손 모드는 '설정' 카드 안에 있음

    $("addUserForm").addEventListener("submit", async (ev) => {
      ev.preventDefault();
      const em0 = $("newUser").value.trim().toLowerCase();
      if (!em0) { // v370(2026-11-16 사용자 지시): 이름만 넣으면 바로 사용자로 등록하고, 조직도에서 그 사용자를 펼친 수정 상태로 열어 나머지 정보를 넣게 함
        const nm0 = $("newName").value.trim();
        if (!nm0) return toast("이름(FC명)을 적어 주세요.");
        const key = nm0 + "@email.pending";
        if ((perms.users || {})[key]) return toast("이메일 없이 등록된 같은 이름이 이미 있습니다. 조직도에서 그 사용자를 펼쳐 주세요.");
        const btn0 = ev.submitter || $("addUserForm").querySelector("button[type=submit]"); if (btn0) btn0.disabled = true;
        try {
          const pm = await readPerms(); pm.users = pm.users || {}; pm.names = pm.names || {}; pm.titles = pm.titles || {};
          pm.users[key] = []; pm.names[key] = nm0; pm.titles[key] = "FC";
          await writePerms(pm); adminState.perms = pm;
          orgOpen.add(key); openOrg(); toast(nm0 + " 사용자를 등록했습니다. 이메일 등 정보를 채워 주세요.");
          openOrgProfile(key, "org", null, true);
        } catch (e) { handleError(e); } finally { if (btn0) btn0.disabled = false; }
        return;
      }
      const em = em0;
      if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(em)) return toast("이메일 형식을 확인해 주세요.");
      if ((CFG.admins || []).map((a) => a.toLowerCase()).includes(em)) return toast("관리자는 이미 모든 폴더를 볼 수 있습니다.");
      const nm = $("newName").value.trim();
      if (!nm) return toast("이름(FC명)을 적어 주세요.");
      if ((perms.users || {})[em]) return toast("이미 있는 사용자입니다.");
      // v336(2026-11-15 사용자 지시): 추가하면 바로 저장 → 조직도에 자동으로 나옴. 이름이 FC 폴더와 같으면 본인 FC(보기+수정)·폴더 공유, 아니면 권한 없이 등록(조직도에서 받은 권한으로 열림)
      if ((perms.archived || {})[em]) return toast("보관된 사용자입니다. 아래 '보관된 사용자'에서 [원복]해 주세요.");
      openOrgProfile(em, "admin", { name: nm }); // v343(2026-11-15 사용자 지시): 추가 = 사용자 정보 팝업(수정 모드)에서 채우고 [저장]하면 등록
    });

    const rows = [...$("adminBody").querySelectorAll(".perm-row")];
    const same = (a, b) => a.length === b.length && a.every((x) => b.includes(x));
    const sel = (row) => {
      const all = row.querySelector("input[data-all]");
      const l = all.checked ? ["*"] : [...row.querySelectorAll("input[data-fc]")].filter((b) => b.checked).map((b) => {
        const ed = row.querySelector(`input[data-edit="${CSS.escape(b.dataset.fc)}"]`);
        return b.dataset.fc + (ed && ed.checked ? "" : ":r"); // 수정 체크 없으면 보기만
      });
      const nm = (row.querySelector("[data-name]") || {}).value; // v331: FC 체크 칸이 조직도로 옮겨져, 새 사용자는 이름(FC명) = 본인 FC로 시작
      if (!l.length && nm && fcs.includes(nm.trim())) l.push(nm.trim());
      return l;
    };
    const nameVal = (row) => row.querySelector("[data-name]").value.trim();
    // 체크가 바뀐 줄 + Drive 공유가 체크 상태와 다른 줄(예: 예전 방식 공유)도 '변경 저장' 대상으로 표시
    const needsShare = (r) => GD && adminState.shares && sel(r).length > 0 && !shareMatches(r.dataset.em, sel(r));
    const isDiff = (r) => !same(sel(r), effList(r.dataset.em)) || needsShare(r) || nameVal(r) !== nameOfEm(r.dataset.em);
    const changed = () => rows.filter(isDiff);
    const refresh = () => {
      rows.forEach((r) => { r.classList.toggle("changed", isDiff(r)); });
      const n = changed().length;
      $("savePerms").disabled = !n;
      $("resetPerms").disabled = !n;
      $("savePerms").textContent = n ? `변경 저장 (${n})` : "변경 저장";
    };
    rows.forEach((row) => {
      const all = row.querySelector("input[data-all]");
      const boxes = [...row.querySelectorAll("input[data-fc]")];
      const eds = [...row.querySelectorAll("input[data-edit]")];
      const syncEd = () => eds.forEach((e) => {
        const fb = boxes.find((b) => b.dataset.fc === e.dataset.edit);
        if (e.dataset.own) { e.checked = true; if (fb) fb.checked = true; return; } // 본인 FC는 항상 보기+수정
        if (all.checked) { e.checked = true; e.disabled = true; } // 전체 = 보기+수정
        else { e.disabled = false; if (!(fb && fb.checked)) e.checked = false; } // 수정 칸은 항상 누를 수 있음 (누르면 보기도 함께 켜짐)
      });
      const syncAll = () => { all.indeterminate = !all.checked && boxes.some((b) => b.checked); syncEd(); };
      all.addEventListener("change", () => { boxes.forEach((b) => (b.checked = all.checked)); syncAll(); refresh(); });
      boxes.forEach((b) => b.addEventListener("change", () => {
        if (!b.checked) { all.checked = false; const ed = eds.find((x) => x.dataset.edit === b.dataset.fc); if (ed) ed.checked = false; } // 보기를 끄면 수정도 꺼짐
        // (모든 FC를 체크해도 '전체'로 바뀌지 않음 — '전체'는 앞으로 생길 FC까지 열림)
        syncAll(); refresh();
      }));
      eds.forEach((e) => e.addEventListener("change", () => { // 수정을 켜면 보기도 같이 켜지고, 수정을 다시 꺼도 보기는 그대로 둠
        if (e.checked) { const fb = boxes.find((b) => b.dataset.fc === e.dataset.edit); if (fb && !fb.checked) fb.checked = true; }
        syncAll(); refresh();
      }));
      row.querySelector("[data-name]").addEventListener("input", refresh);
      syncAll();
    });
    refresh();
    $("resetPerms").addEventListener("click", () => renderAdmin());
    bindSettings();
    bindStartCard();
    bindStartRows();
    bindSmsTeamCard();
    bindSmsCard();
    bindPfxRows(); // 각 사용자 카드의 '일정 구분'(AI로 등록하는 일정 제목 앞 글자) 저장
    $("adminBody").querySelectorAll("[data-admin-name]").forEach((inp) => {
      const em = inp.dataset.adminName, btn = $("adminBody").querySelector(`[data-admin-save="${CSS.escape(em)}"]`), base = inp.value.trim();
      inp.addEventListener("input", () => { btn.disabled = inp.value.trim() === base; });
      btn.addEventListener("click", async () => { // 관리자 본인 FC 이름: 권한 파일 names에 저장 (세일즈 화면 FC 선택 기본값·문자 {FC명}에 사용)
        btn.disabled = true; btn.textContent = "저장 중…";
        try {
          const perms = await readPerms();
          perms.names = perms.names || {};
          if (inp.value.trim()) perms.names[em] = inp.value.trim(); else delete perms.names[em];
          await writePerms(perms);
          adminState.perms = perms;
          toast("저장했습니다.");
          renderAdmin();
        } catch (e) { btn.disabled = false; handleError(e); }
        btn.textContent = "저장";
      });
    });
    // 영업 추가 FC 체크: 누르는 즉시 권한 파일(salesFcs)에 저장 (FC 보기·편집 '변경 저장'과 따로)
    $("adminBody").querySelectorAll("[data-sf]:not([data-own])").forEach((cb) => cb.addEventListener("change", async () => {
      const em = cb.dataset.sf.split("|")[0];
      const group = [...$("adminBody").querySelectorAll("[data-sf]:not([data-own])")].filter((x) => x.dataset.sf.split("|")[0] === em);
      group.forEach((x) => (x.disabled = true));
      try {
        const perms = await readPerms(); // 저장 직전 최신 파일
        perms.salesFcs = perms.salesFcs || {};
        const list = group.filter((x) => x.checked).map((x) => x.dataset.sf.split("|").slice(1).join("|"));
        if (list.length) perms.salesFcs[em] = list; else delete perms.salesFcs[em];
        await writePerms(perms);
        adminState.perms = { ...adminState.perms, salesFcs: perms.salesFcs };
        toast(`${nameOfEm(em) || em}: 영업 추가 FC를 저장했습니다.`);
      } catch (e) { cb.checked = !cb.checked; handleError(e); }
      group.forEach((x) => (x.disabled = false));
    }));
    // 민감정보 접근 체크박스: 누르는 즉시 권한 파일(sensitive 목록)에 저장 — FC 보기·편집 '변경 저장'과는 따로 동작
    $("adminBody").querySelectorAll("[data-sens]").forEach((cb) => cb.addEventListener("change", async () => {
      const em = cb.dataset.sens, want = cb.checked;
      cb.disabled = true;
      try {
        const perms = await readPerms(); // 저장 직전 최신 파일
        const set = new Set(perms.sensitive || []);
        if (want) set.add(em); else set.delete(em);
        perms.sensitive = [...set].sort();
        await writePerms(perms);
        adminState.perms = { ...adminState.perms, sensitive: perms.sensitive };
        toast(`${nameOfEm(em) || em}: 민감정보 접근을 ${want ? "허용" : "해제"}했습니다.`);
      } catch (e) { cb.checked = !want; handleError(e); }
      cb.disabled = false;
    }));
    $("adminBody").querySelectorAll("[data-del]").forEach((b) => b.addEventListener("click", () => deleteUser(b.dataset.del)));
    $("adminBody").querySelectorAll("[data-login]").forEach((b) => b.addEventListener("click", (ev) => { ev.stopPropagation(); toggleLogin(b.dataset.login, b); }));
    $("adminBody").querySelectorAll("[data-share]").forEach((b) => b.addEventListener("click", () => shareFolders(b.dataset.share)));
    collectPending = collectRows;
    applyPending = applyRows;
    $("savePerms").addEventListener("click", () => savePerms(changed().map((r) => [r.dataset.em, sel(r), nameVal(r)])));
    // 사용자 카드는 평소 잠겨 있음: [수정][삭제] → 수정 모드에서는 [취소][저장] (저장·취소하면 다시 잠기고 버튼도 [수정][삭제]로 돌아옴)
    const rowLock = (r, lock) => {
      r.classList.toggle("locked", lock);
      r.querySelector("[data-name]").readOnly = lock;
      r.querySelectorAll(".perm-sens, .perm-sales, .perm-fcs").forEach((x) => { if (lock) x.setAttribute("inert", ""); else x.removeAttribute("inert"); });
      r.querySelector("[data-rowedit]").hidden = !lock; r.querySelector("[data-del]").hidden = !lock;
      r.querySelector("[data-rowcancel]").hidden = lock; r.querySelector("[data-rowsave]").hidden = lock;
    };
    const setOpen = (r, on) => { r.classList.toggle("open", on); const f = r.querySelector("[data-pfold]"); if (f) f.setAttribute("aria-expanded", String(on)); if (on) permOpen.add(r.dataset.em); else permOpen.delete(r.dataset.em); };
    rows.forEach((r) => {
      rowLock(r, true);
      r.querySelector(".perm-head").addEventListener("click", (ev) => { // v342: 머리 줄([삭제] 제외)을 누르면 사용자 정보 팝업 (접기·펼치기 없앰)
        return; // v353(2026-11-16 사용자 지시): 카드를 눌러도 아무 일 없음 — [수정]으로만 엶
      });
      r.querySelector("[data-rowedit]").addEventListener("click", () => openOrgProfile(r.dataset.em, "admin", null, true)); // v349: [수정] = 사용자 정보 수정 모드 팝업
      r.querySelector("[data-rowcancel]").addEventListener("click", () => { applyRows([[r.dataset.em, effList(r.dataset.em), nameOfEm(r.dataset.em)]]); rowLock(r, true); refresh(); }); // 바꾼 내용을 버리고 처음 값으로
      r.querySelector("[data-rowsave]").addEventListener("click", () => { if (!isDiff(r)) { rowLock(r, true); return; } savePerms([[r.dataset.em, sel(r), nameVal(r)]]); }); // 이 계정만 저장 (저장하면 화면을 다시 그려 잠김)
    });

    // 사용자 추가 시 저장 전 체크 상태 보존용
    function collectRows() { return rows.map((r) => [r.dataset.em, sel(r), nameVal(r)]); }
    function applyRows(pending) {
      pending.forEach(([em, list, nm]) => {
        const row = $("adminBody").querySelector(`.perm-row[data-em="${CSS.escape(em)}"]`);
        if (!row) return;
        const all = row.querySelector("input[data-all]");
        all.checked = list.includes("*");
        row.querySelectorAll("input[data-fc]").forEach((b) => (b.checked = all.checked || list.some((g) => grantFc(g) === b.dataset.fc)));
        row.querySelector("[data-name]").value = nm || "";
        // v96: '전체'가 꺼진 줄에서 all의 change를 보내면 방금 복원한 FC 체크를 모두 지워 버렸음(사용자 추가 시 다른 줄 체크가 풀리던 문제) → '전체'가 켜진 줄만 보냄
        if (all.checked) all.dispatchEvent(new Event("change"));
        else row.querySelectorAll("input[data-fc]").forEach((b) => b.dispatchEvent(new Event("change")));
        row.querySelectorAll("input[data-edit]").forEach((e) => { if (!all.checked) e.checked = list.includes(e.dataset.edit); e.dispatchEvent(new Event("change")); });
      });
    }
  }

  // v343(2026-11-15 사용자 지시): [삭제] = 보관. 사용을 막고(권한·공유 해제) 데이터는 지우지 않고 perms.archived에 옮겨 둠 → [원복] 가능. 사용자 정보 파일은 그대로
  async function deleteUser(em) {
    const shared = GD && shareState(em) && shareState(em) !== "none", nm0 = ((adminState.perms || {}).names || {})[em] || em;
    if (!confirm(`${nm0} (${em}) 계정을 보관할까요?\n보관하면 이 계정은 앱을 쓸 수 없지만, 사용자 정보와 권한 기록은 지우지 않고 남겨 두어 아래 '보관된 사용자'에서 [원복]할 수 있습니다.` +
      (shared ? "\nGoogle Drive 폴더 공유는 해제합니다." : ""))) return;
    try {
      const perms = await readPerms();
      perms.archived = perms.archived || {};
      perms.archived[em] = { users: (perms.users || {})[em] || [], name: (perms.names || {})[em] || "", title: (perms.titles || {})[em] || "", salesFcs: (perms.salesFcs || {})[em] || [], sensitive: (perms.sensitive || []).includes(em), org: (perms.org || {})[em] || null, starts: (perms.starts || {})[em] || null, at: new Date().toISOString(), by: me() };
      delete perms.users[em];
      if (perms.names) delete perms.names[em];
      perms.sensitive = (perms.sensitive || []).filter((x) => x !== em);
      ["salesFcs", "titles", "org", "starts", "orgApplied", "orgSalesApplied", "login"].forEach((k) => { if (perms[k]) delete perms[k][em]; }); // v513: 보관하면 로그인 승인도 지움(원복하면 거부로 시작)
      extraUsers.delete(em);
      await writePerms(perms);
      adminState.perms = perms;
      if (shared) await unshareFolders(em);
      toast(`${nm0} 계정을 보관했습니다. (데이터는 그대로)`);
      renderAdmin();
    } catch (e) {
      handleError(e);
    }
  }

  async function unarchiveUser(em) {
    const a = ((adminState.perms || {}).archived || {})[em]; if (!a) return;
    await savePerms([[em, (a.users || []).slice(), a.name || undefined]], { pre: [`${a.name || em} (${em}) 계정을 원복합니다. 보관할 때의 권한·직책·조직 위치로 돌아갑니다.`],
      mutate: (perms) => {
        if (a.title) perms.titles = { ...(perms.titles || {}), [em]: a.title };
        if ((a.salesFcs || []).length) perms.salesFcs = { ...(perms.salesFcs || {}), [em]: a.salesFcs };
        if (a.sensitive) perms.sensitive = [...new Set([...(perms.sensitive || []), em])];
        if (a.org) perms.org = { ...(perms.org || {}), [em]: a.org };
        if (a.starts) perms.starts = { ...(perms.starts || {}), [em]: a.starts };
        if (perms.archived) delete perms.archived[em];
      } });
  }
  document.addEventListener("click", (ev) => { const b = ev.target && ev.target.closest && ev.target.closest("[data-unarch]"); if (b && adminState) unarchiveUser(b.dataset.unarch); });
  async function savePerms(changes, opt) {
    opt = opt || {};
    if (!changes.length) return;
    if (changes.some(([, l]) => l.includes("*"))) return toast("전체(*)는 더 이상 지원하지 않습니다. FC를 하나씩 체크해 주세요.");
    // Google Drive: FC를 체크한 계정은 폴더 공유, 체크를 모두 뺀 계정은 공유 해제 (변경 저장 한 번으로)
    const auto = GD && adminState.shares;
    const toShare = auto ? changes.filter(([em, list]) => list.length && !shareMatches(em, list) && loginOn(adminState.perms, em)).map(([em]) => em) : []; // v513: 로그인 거부 계정은 공유하지 않음(승인할 때 공유)
    const toUnshare = auto ? changes.filter(([em, list]) => !list.length && shareState(em) !== "none").map(([em]) => em) : [];
    const listOfChange = (em) => (changes.find((c) => c[0] === em) || [em, []])[1];
    if (toShare.length || toUnshare.length || opt.pre) {
      const msg = [...(opt.pre || [])];
      if (toShare.length || toUnshare.length) msg.push("저장하면서 Google Drive 공유도 바꿉니다.");
      if (toShare.length) msg.push("· 공유(체크한 FC 폴더만 편집·보기만, 설정 폴더 보기): " + toShare.map((em) => em + " [" + listOfChange(em).join(", ") + "]").join(" / "));
      const ancMsg = [...toShare, ...toUnshare].flatMap((em) => ancShares(em).map((a) => em + " ← '" + a.folder + "' 폴더 공유(" + a.perm.role + ")"));
      if (ancMsg.length) msg.push("· 상위 폴더에 직접 공유된 것도 함께 해제합니다(이 공유가 있으면 FC 제한이 Drive에서 안 걸림): " + ancMsg.join(", "));
      if (toUnshare.length) msg.push("· 공유 해제: " + toUnshare.join(", "));
      if (!confirm(msg.join("\n"))) return;
    }
    const btn = opt.btn || $("savePerms") || {};
    btn.disabled = true;
    btn.textContent = "저장 중…";
    try {
      // 저장 직전에 최신 파일을 다시 읽고, 바뀐 아이디만 반영 (다른 설정 보존)
      const perms = await readPerms();
      perms.names = perms.names || {};
      changes.forEach(([em, list, nm]) => {
        if (list.length) { perms.users[em] = list; if (nm !== undefined) { if (nm) perms.names[em] = nm; else delete perms.names[em]; } }
        else { perms.users[em] = []; if (nm) perms.names[em] = nm; extraUsers.delete(em); } // v336: 권한 없이도 계정은 남김(조직도에 보이게). 계정 삭제는 [삭제]로
      });
      perms.fcFolders = adminState.fcIds || perms.fcFolders || {}; // 사용자 앱이 FC 폴더를 ID로 바로 열 수 있게
      if (opt.mutate) opt.mutate(perms);
      await writePerms(perms);
      adminState.perms = perms;
      const failed = [], tAll = { cust: false, app: false, fc: [], anc: [] };
      for (const em of toShare) { try { const tt = await grantFolders(em, listOfChange(em)); tAll.cust = tAll.cust || tt.cust; tAll.app = tAll.app || tt.app; tAll.fc.push(...tt.fc); tAll.anc.push(...tt.anc); } catch (e) { console.error(e); failed.push(em); } }
      for (const em of toUnshare) { try { await unshareFolders(em); } catch (e) { console.error(e); failed.push(em); } }
      if (toShare.length || toUnshare.length) { // 실패가 있으면 상태를 확실히 알 수 없으니 전체를 다시 읽고, 모두 성공이면 바뀐 폴더만 갱신
        if (failed.length) adminState.shares = await loadShares(); else await refreshSharesFor(tAll);
      }
      toast(failed.length ? `저장했습니다. 공유하지 못한 계정: ${failed.join(", ")} (구글 계정 이메일인지 확인)` : `저장했습니다 (${changes.length}개 계정${toShare.length ? ", 폴더 공유 " + toShare.length : ""}${toUnshare.length ? ", 공유 해제 " + toUnshare.length : ""}).`);
      if (opt.after) opt.after(); else renderAdmin();
    } catch (e) {
      btn.disabled = false;
      btn.textContent = opt.btn ? "저장" : "변경 저장";
      handleError(e);
    }
  }

  // ---------- 오류 처리 ----------
  function handleError(e) {
    console.error(e);
    const m = String(e && e.message);
    if (m === "NOT_SIGNED_IN") return show("loginView");
    if (GD && m === "GOOGLE_AUTH") { showGLogin(); return toast("구글 연결 시간이 지났습니다. 다시 연결해 주세요."); }
    if (GD && /Drive 403/.test(m) && /accessNotConfigured|has not been used|disabled/i.test(m)) return toast("Google Drive API가 켜져 있지 않습니다. 관리자에게 알려 주세요.");
    if (m === "NO_SHARE") {
      $("scopeHint").textContent = "아직 고객 폴더에 접근할 수 없습니다. 관리자에게 " + DRIVE_NAME + " '고객' 폴더 공유를 요청하세요.";
      return toast("고객 폴더 공유가 필요합니다. 관리자에게 요청하세요.");
    }
    if (m.includes("404")) return toast(DRIVE_NAME + "에서 '" + CFG.baseFolder + "' 폴더를 찾을 수 없습니다.");
    if (m.includes("403")) return toast("권한이 없습니다. 관리자에게 요청하세요.");
    toast("오류: " + m.slice(0, 120));
  }

  // ---------- 시작 ----------
  // Google Drive 모드: 구글 로그인 창(팝업) 하나로 Drive·캘린더 권한을 함께 받음 (토큰 1시간, 지나면 버튼 한 번)
  function showGLogin() {
    const em = gAccount() || (readGAuth() || {}).email || "";
    const sh = readGAuth();
    const needScope = !!(sh && sh.exp > Date.now() + 60000 && !gHasScopes(sh.scopes)); // 다른 앱에서 로그인했지만 이 앱 권한이 모자란 경우
    $("loginBtn").textContent = needScope ? "구글 연결 (권한 추가)" : em ? "Google 다시 연결 (" + em + ")" : "Google 계정으로 로그인";
    $("loginOtherBtn").hidden = false;
    if (EMBED) { // 모듈로 떠 있을 땐 여기서 로그인하지 않고 말만해요에서 다시 로그인 (공통 출입증으로 이어짐)
      $("lmsg").textContent = "말만해요에서 다시 로그인해 주세요.";
      $("loginBtn").hidden = true; $("loginOtherBtn").hidden = true;
    }
    show("loginView");
  }
  // other = true: 구글 계정 선택 창을 띄워 다른 계정으로 로그인
  async function googleLogin(other) {
    try {
      await requestGToken(other || am6Pending() ? "select_account" : undefined, other); // v557: 6시 로그아웃 뒤 첫 로그인은 계정 고르기 창
      try { localStorage.removeItem(AM6_KEY); } catch (e) { /* 무시 */ }
      const u = await gdFetch("https://www.googleapis.com/oauth2/v3/userinfo");
      const em = String(u.email || "").toLowerCase();
      if (!em) throw new Error("구글 계정 이메일을 읽지 못했습니다.");
      const changed = gAccount() && gAccount() !== em;
      try { localStorage.setItem(GACC_KEY, em); } catch (e) { /* 이번만 */ }
      const sh = readGAuth(); if (sh) writeGAuth({ ...sh, email: em }); // 공통 출입증에 확인한 계정 기록
      if (changed) { await photoCacheClear(); return location.reload(); } // 계정이 바뀌면 앱을 새로 시작 (이전 계정의 자료·설정 비우기)
      startApp();
    } catch (e) {
      if (e.message !== "GOOGLE_AUTH") toast(e.message); else handleError(e);
    }
  }
  $("loginBtn").addEventListener("click", () => (GD ? googleLogin() : msal.loginRedirect({ scopes: SCOPES })));
  $("loginOtherBtn").addEventListener("click", () => (GD ? googleLogin(true) : msal.loginRedirect({ scopes: SCOPES, prompt: "select_account" })));
  // ---- 관리자·설정 화면의 저장·취소 하나로: 화면 안의 모든 저장 버튼을 한꺼번에 눌러 주고, 아래 메뉴 바로 위에 떠 있음 ----
  //   저장 순서: 설정 → 문자 문구 → 내 캘린더 → (마지막) FC 권한. FC 권한 저장은 화면을 다시 그리므로 맨 마지막
  //   설정 화면의 어떤 카드든 저장할 내용이 있으면(그 카드의 저장 버튼이 켜져 있으면) 아래 [저장]이 켜지고, 누르면 켜진 카드를 모두 저장
  const FS_FIRST = ["setSave", "smsSave", "startSave", "smsTeamSave", "myCalSave"];
  const fsBtns = () => {
    const list = FS_FIRST.map((id) => $(id)).filter(Boolean);
    $("adminBody").querySelectorAll("[data-pfx-save], [data-sa-save], [data-org-save]").forEach((b) => list.push(b)); // v344: 조직도는 위 [저장] 없이 아래 [저장]으로
    const pp = $("savePerms"); if (pp) list.push(pp);
    return list;
  };
  const fsDirty = () => fsBtns().some((b) => !b.disabled);
  function fsRefresh() {
    const bar = $("floatSave");
    if (!bar) return;
    const on = !$("adminView").hidden;
    bar.hidden = !on;
    if (!on) return;
    const dirty = fsDirty(), busy = $("fsSave").dataset.busy === "1";
    $("fsSave").disabled = !dirty || busy;
  }
  new MutationObserver(fsRefresh).observe($("adminBody"), { subtree: true, childList: true, attributes: true, attributeFilter: ["disabled"] });
  // 세일즈 화면 위쪽 줄(이달이관·따행 선택, 월 이동, FC 선택)의 높이를 재 두어, 그 아래 '설문 발송' 줄이 바로 밑에 붙어 고정되게 함
  new MutationObserver(() => { const st = document.querySelector("#salesView .sales-top"); if (st) document.documentElement.style.setProperty("--salestop", Math.round(st.getBoundingClientRect().height) + "px"); }).observe($("salesView"), { subtree: true, childList: true });
  const fsWait = async (b) => { for (let i = 0; i < 160 && b.isConnected && /저장 중/.test(b.textContent); i++) await new Promise((r) => setTimeout(r, 150)); };
  // ---------- 설명서: 관리자·설정 화면의 설명을 한곳에 모음 (아래 '설명서' 버튼) ----------
  function helpHtml() {
    const iconBox = `<div class="help-icon"><img src="icon.svg?v=3" width="72" height="72" alt="신한FC m 앱 아이콘"><div><b>앱 아이콘 안내</b><p>고객의 보장을 지키는 <b>방패</b> 안에, 계약과 실적이 꾸준히 <b>오르는 화살표</b>를 담았습니다. 보험 설계사의 일(고객을 지키고, 함께 성장하는 것)을 한 그림으로 줄였습니다.</p><p>연한 하늘색 바탕과 파란 선은 신뢰감 있고 차분한 업무용 앱의 느낌을 주고, 굵은 윤곽선 방패는 작은 홈 화면 아이콘에서도 또렷하게 보이도록 단순하게 만들었습니다. 신한FC m(모바일)에서는 아래 메뉴 맨 왼쪽의 이 아이콘을 누르면 AI 도우미가 열리고, PC 화면 왼쪽 위에는 "신한FC"로 표시됩니다.</p></div></div>`;
    const sec = (t, items) => `<h4>${t}</h4><ul>${items.map((x) => `<li>${x}</li>`).join("")}</ul>`;
    const grp = (t) => `<h3 class="help-grp">${t}</h3>`; // 하단 메뉴 순서대로 묶음: AI · 일정 · (고객) · 세일즈 · 설정/관리자
    const s_team = sec("팀관리 (관리자)", ["<b>관리자</b>로 로그인하면 아래 메뉴 <b>세일즈 오른쪽</b>에 <b>팀관리</b>가 생깁니다. <b>[조직도]</b> · <b>[팀원정보]</b> · <b>[후보자정보]</b> 버튼이 있고(설정에 있던 조직도를 옮겼습니다), 내용은 정해지는 대로 채웁니다."]);
    const s_start = sec("시작 화면", ["<b>2026-11-21부터 처음 화면은 대시보드로 고정</b>되어 시작 화면 고르기는 숨겼습니다(아래 설명은 다시 살릴 때를 위한 참고). 앱을 쓰다 잠깐 나갔다 30분 안에 돌아오면 보던 화면(일정·영업)을 이어서 보여 주는 기능은 그대로입니다.", "<b>처음 보일 화면</b>: 설정 화면의 '화면' 카드 가운데 칸(시작 화면)에서 일정 · 고객(기본) · 세일즈 중 하나를 고르면 바로 저장되고, 다음에 앱을 열 때부터 적용됩니다. 같은 카드 왼쪽 칸은 ‘시작하면 AI 도우미 바로 열기’, 오른쪽 칸은 ‘오른손 모드’입니다. 저장한 값은 내 FC 폴더 안 '_앱설정' 폴더에도 기록되어 관리자가 확인·수정할 수 있고, 다른 폰에도 따라갑니다. 관리자는 설정 화면 아래 ‘관리자’의 각 사용자 카드(눌러서 펼침) 안 ‘시작 화면’에서 그 사용자의 값을 보고 바꿀 수 있습니다.", "<b>AI 도우미 바로 열기</b>: '화면' 카드 왼쪽 칸을 켜면 앱을 열자마자 AI 도우미 창이 올라옵니다.", "내 FC 폴더에 쓸 수 없는 계정은 이 폰에만 저장됩니다. 앱을 쓰다 잠깐 나갔다 30분 안에 돌아오면 보던 화면(일정·영업)을 이어서 보여 주는 기능이 먼저 적용될 수 있습니다."]);
    const s_ai = sec("AI 도우미", ["<b>2026-11-19부터 AI 버튼(아래 메뉴 맨 왼쪽 아이콘·떠 있는 버튼)은 우선 숨겼습니다.</b> 다시 보이게 할 때까지 아래 설명은 참고용입니다.", "<b>열기·닫기</b>: 아래 메뉴 맨 왼쪽의 앱 아이콘을 누르면 AI 도우미가 아래에서 올라옵니다. 아래로 끌어내리거나, 아이콘을 한 번 더 누르거나, 바깥 어두운 곳을 누르거나, 폰의 뒤로가기를 누르면 내려갑니다.", "<b>말하기·입력</b>: 큰 마이크 버튼을 눌러 말하면 입력칸에 글자로 들어가고, 키보드로 직접 써도 됩니다. 아래 예시 문구를 누르면 입력칸에 들어갑니다.", "<b>지금 알아듣는 말</b>: ① “홍길동 검색” → 고객 검색 ② “내일 오후 3시 홍길동 상담 일정 추가” → 구글 캘린더(1팀)에 1시간짜리 일정 등록 (“내 일정에 …”이라고 하면 내 FC 이름과 같은 이름의 내 캘린더에 등록, 이때는 앞 글자를 붙이지 않음) (시간을 말하지 않으면 종일 일정, 날짜를 말하지 않으면 오늘) ③ “오늘 일정 보여줘” → 일정 화면. 말한 뒤에는 \"이렇게 이해했습니다\" 확인 카드가 뜨고, [실행]을 눌러야 실제로 처리됩니다. 못 알아들으면 말하는 방법을 안내합니다.", "<b>일정 구분(제목 앞 글자)</b>: 관리자가 사용자마다 정해 두면(설정 화면 아래 ‘관리자’의 각 사용자 카드 안 ‘일정 구분’) 등록할 때 제목 맨 앞에 붙습니다. 예: 최)홍길동 상담. 칸에는 글자만 적어도 뒤에 ‘)’가 자동으로 붙고, ‘)’까지 적어도 하나만 붙습니다", "<b>이름 보호</b>: 위 말처럼 규칙으로 알아듣는 말은 이 폰 안에서만 처리하고 밖으로 보내지 않습니다. 규칙으로 못 알아들은 말만 Gemini(외부 AI)에 '무슨 뜻인지'를 묻는데, 그때 고객 이름은 [고객1] 같은 자리표시로, 주민번호·전화·계좌·이메일·주소 같은 숫자·주소는 [숫자]·[주소]로 가려서 보내고 답을 받은 뒤 앱 안에서 되돌립니다. 말만해요 설정에 Gemini 키가 없거나 연결이 안 되면 규칙만 쓰며, 확인 카드에 ‘AI가 해석’이라고 표시되고 [실행]을 눌러야 처리됩니다."]);
    const s_task = sec("구글 할 일 (일정 화면)", ["<b>[일정으로]</b>: 할 일 줄의 [일정으로]를 누르면 그 할 일을 캘린더 일정으로 옮기고 할 일은 지웁니다. 일정에서 [할 일로] 옮길 때 메모에 적힌 <b>예전 시간</b>(‘시간: 09:00~10:00’)이 있으면 그 시간으로, 없으면 <b>종일 일정</b>으로 만듭니다. 날짜는 마감일(없으면 보고 있는 날), 캘린더는 원래 있던 캘린더(없으면 기본 캘린더)이고, 장소·메모·작성자·고객 연결도 이어집니다.", "<b>할 일 제목에 등록된 고객 이름이 있으면</b> 그 할 일 줄 아래에 <b>일정과 똑같은 고객 카드</b>(이름·생년월일·연락처·주소·미팅장소·상담기록, 이름을 누르면 고객 정보)가 붙습니다. <b>종일 일정</b>도 캘린더와 상관없이 이름으로 고객을 찾습니다. 같은 이름이 여러 명이면 [선택]으로 한 명을 고르고(나중에 [다시 선택]), 관계없으면 [해당 없음]([되돌리기]로 다시 연결). 고른 값은 할 일 메모의 ‘고객연결: …’ 줄에 저장되어 모든 기기에서 같게 보입니다.", "<b>관리자</b>는 할 일을 추가·수정할 때 <b>사용자</b>를 고를 수 있습니다(기본 = 본인). 구글 할 일은 로그인한 계정의 것이라, 다른 사용자를 고르면 할 일 메모 맨 윗줄에 ‘작성: 이름’이 적히고 목록에는 ‘- 이름’으로 보입니다(본인으로 되돌리면 그 줄이 지워짐).", "<b>모든 일정</b> 카드(종일이 아니어도)에는 <b>[할 일로]</b> 버튼이 있습니다(시간이 있는 일정은 할 일 메모에 시간이 적힙니다). 누르면 그 날짜가 마감인 구글 할 일(기본 목록)을 만들고 캘린더 일정은 지웁니다. 장소·메모는 할 일의 메모로 옮겨집니다. 지운 일정은 화면을 새로 고치기 전까지 [원위치]로 되살릴 수 있습니다(만든 할 일은 그대로 남음).", "그날 마감인 구글 할 일이 일정 화면 맨 위에 나오고, 체크하면 완료로 바뀝니다.", "오늘(Today)을 볼 때는 아래에 작은 묶음 두 개가 더 나옵니다: '밀린 할 일'(마감이 지났는데 아직 안 한 일, 오래된 것부터, 앞에 마감일)과 '날짜 없음'(마감일이 없는 일). 체크하면 완료되어 묶음에서 빠지고, 6.5초 안에 [되돌리기]로 되살릴 수 있습니다. 제목 옆 숫자는 남은 할 일/전체입니다.", "할 일 줄 오른쪽 [수정]으로 제목·마감일을 고치고(마감일을 비우면 '날짜 없음'), [삭제]는 확인 뒤 지우며 6.5초 안에 [되돌리기]로 다시 만들 수 있습니다. 목록 이름은 할 일 목록이 둘 이상일 때만 보입니다.", "로그인할 때 할 일 권한도 함께 허용하면 따로 연결하지 않아도 됩니다. 연결이 풀리면(약 1시간 뒤) 앱이 먼저 조용히 다시 연결을 시도하고, 안 되면 나오는 [연결] 버튼을 한 번 누르세요."]);
    const s_nm = sec("일정 제목 → 고객 찾기 (2026-11-17)", ["<b>낱말로 찾기(v579)</b>: 일정 제목에 고객의 전체 이름이 없어도, 제목의 낱말이 이름 안에 들어 있으면 고객을 추천합니다(예: 'Keiko 상담' → Okamoto Keiko). 미팅·상담 같은 흔한 말은 빼고 찾습니다. 영어 이름은 낱말마다 첫 글자만 대문자로 보입니다.", "<b>고객 연결 기록(v559)</b>: 일정에서 고객이 한 명으로 찾아지거나 [선택]·[무시]·[다시 선택]을 누르면, 그 연결(고객ID·담당 FC·고객의 FC·종류)을 구글 일정의 숨은 칸과 고객 폴더의 <b>일정매칭</b> 기록에 남깁니다. 그래서 다른 앱·관리 세션도 \"이 일정 = 어느 고객\"을 바로 알 수 있습니다. 화면 모양은 그대로이고, 설명란에는 적지 않습니다. 캘린더 수정 권한이 없으면 기록 파일에만 남깁니다.", "일정 화면을 <b>좌우로 밀면</b> Today는 전날·다음날, Week는 전주·다음주, Month는 전월·다음월로 넘어갑니다.", "일정 제목에 <b>등록된 고객 이름</b>이 있으면(‘고객’ 글자가 없어도, 날짜와 상관없이) 그 고객 카드가 일정에 붙고, 카드에서 장소 ‘입력’을 할 수 있습니다. 같은 이름의 고객이 여럿이면 [선택]으로 고르고, 상관없는 고객이면 [무시]로 숨깁니다. 등록된 고객에 없는 이름이면 아무것도 붙이지 않습니다. (2026-11-18: 하루 동안 쓰던 ‘이름 + 고객’ 규칙은 풀었습니다.)"]);
    const s_cal = sec("구글 캘린더 (표)", ["<b>Show</b>: 일정 화면에 보일 캘린더를 고릅니다. 꺼 두면 일정·정보를 켤 수 없습니다.", "<b>일정</b>: 영업 일정으로 쓸 캘린더입니다. 고객 매칭을 하고, 이동 거리·예상 시간·회사 출발·도착은 그 일정 카드의 <b>비영업</b> 버튼을 눌러 영업일정으로 바꾼 일정에만 계산됩니다(예전에 1팀 캘린더에 만든 일정은 그대로 영업일정). 켜면 정보도 함께 켜집니다.", "<b>정보</b>: 이 캘린더의 일정에서 고객 정보를 찾아 붙입니다. 끄면 일정도 꺼집니다."]);
    const s_addev = sec("일정 추가 (일정 화면)", ["<b>날짜 고르기(v561)</b>: 앱의 모든 날짜 칸(일정·할 일·청구·알림 기간·입사일·병역·졸업 연도·결혼기념일·자동차보험 갱신일)은 칸을 누르면 바로 아래에 달력이 떠서 고릅니다(직접 입력 없음). 기간은 시작과 끝을 차례로 누르고, [지우기]로 비울 수 있습니다. PC에서는 Enter로 열고 화살표·Enter로 고릅니다.", "[선택]으로 고객과 묶은 일정·할 일에는 <b>[해당 없음]</b> 버튼이 보이지 않습니다. [다시 선택]으로 풀면 다시 보입니다.", "<b>일정과 고객이 묶여 있으면</b>(이름이 맞는 고객이 있어 [선택]을 눌러 한 명과 연결한 일정만, [다시 선택]이 보이는 상태) 카드 왼쪽에 <b>파란색 굵은 선</b>이 보입니다.", "<b>일정 카드의 [수정]·[삭제] 버튼은 평소에 접혀 있습니다.</b> ([할 일로]는 [영업/비영업] 왼쪽에 항상 보입니다.) 일정 칸(시간 줄이든 제목 줄이든)을 터치하면 제목 줄 아래 오른쪽 끝으로 펼쳐지고, 한 번 더 터치하면 접힙니다. [영업/비영업]은 예전처럼 윗줄 오른쪽에 항상 보입니다. <b>[수정]</b>에는 <b>종일</b> 체크가 있어, 종일 일정은 체크를 풀고 시작·종료 시간을 넣으면 시간 일정으로, 시간 일정은 체크하면 종일 일정으로 저장됩니다.", "<b>관리자</b>는 이미 있는 일정을 <b>[수정]</b>할 때도 맨 아래 <b>작성자</b>를 바꿀 수 있습니다(고르는 사람은 내가 쓸 수 있는(공유받은) 캘린더를 가진 사람만, 본인으로 되돌리면 작성자 표시가 지워짐). 할 일 [수정]의 <b>사용자</b>도 같은 사람들 중에서 고릅니다.", "<b>2026-11-18부터</b> 추가 칸 맨 위에서 <b>[일정]</b>과 <b>[할 일]</b>을 먼저 고릅니다. <b>일정</b>은 제목·캘린더·날짜·종일/시간·장소(관리자는 작성자), <b>할 일</b>은 제목·<b>마감일</b>·메모(관리자는 <b>사용자</b>도 고름)만 입력하고 [할 일 등록]을 누르면 구글 할 일(기본 목록)에 그 날짜가 마감인 할 일이 만들어집니다(처음 한 번은 할 일 연결 창이 뜰 수 있음).", "일정 화면 오른쪽 위의 <b>추가</b> 버튼을 누르면 제목·캘린더·날짜·시간(또는 종일)·장소를 넣어 새 일정을 구글 캘린더에 등록할 수 있습니다. 처음 열 때 캘린더는 1팀이 선택되어 있습니다.", "일정은 처음에 모두 <b>비영업일정</b>입니다. 카드의 <b>비영업</b> 버튼(수정 왼쪽)을 누르면 <b>영업일정</b>이 되어 이동 거리·시간이 나옵니다(영업일정 카드에는 ‘영업’ 버튼이 나오고, 누르면 다시 비영업일정으로 돌아갑니다. 버튼 글자는 지금 상태입니다). (일정 위쪽의 [초기화] 버튼은 2026-11-19에 없앴습니다.) 새로 만든 일정과 1팀 외 캘린더의 일정은 비영업일정으로 시작합니다."]);
    const s_sub = sec("세일즈 서브메뉴", ["세일즈 화면 맨 위의 <b>이달목표</b>(이관·따행·DB 목록)와 <b>보고서</b>를 눌러 바꿉니다. <b>보고서</b>에는 <b>조건 보고서</b>가 있습니다: 사용자(내 권한 범위) · 유형(따행/이관/DB, 여러 개) · <b>단계</b>(신규·계약·완료·지옥, 여러 개 — 계약은 DB의 계약 고객) · <b>기간</b>(년월 하나)을 고르면 건수가 바로 나오고, <b>진행</b>을 누르면 보고서가 화면에 나타납니다. 기간의 <b>[3개월]</b>·<b>[1년]</b>은 이번 달까지의 그 기간으로 <b>누르는 즉시 보고서</b>를 보여 줍니다(진행을 누르지 않아도 됨). 이관은 이관일, 따행은 대상년월이 기간 안인 고객입니다. 그 아래 <b>엑셀 다운</b> · <b>PDF 다운</b> · <b>PDF 열람</b> · <b>이메일(나에게)</b>(본문에 명단만, 파일 첨부는 안 됨) 버튼이 있습니다. DB는 월과 상관없이 지금 DB 고객 전체입니다. 이달목표 안에는 월 줄 오른쪽 드롭박스로 고르는 <b>따행</b> · <b>이관</b>(이달 이관) · <b>DB</b>(고객 시트의 유입경로가 ‘DB’인 고객, 유입일 최신순·같은 날은 이름순, 전화·문자 버튼 포함, 줄 오른쪽에 이관과 같은 영업단계 칩)가 있습니다. 이관·따행·DB 모두 줄의 작은 <b>이동</b>으로 보관에 넣고, ‘N명’ 줄 오른쪽 <b>보관 M</b> 칩에서 보관된 고객을 보며 <b>원복</b>으로 되돌립니다(자세한 것은 ‘보관 (이동 · 원복)’). 계약 고객이 된 DB 고객은 저절로 보관에 ‘계약’으로 보입니다. 이관·따행·DB 줄 왼쪽의 굵은 색 선은 지금 단계(이관·DB = 통화·AP·PC·계약, 따행 = 마지막으로 한 단계)의 색이고, 따행 칩은 방문완료 > 설문발송 > 설문응답 > 쿠폰발송 중 마지막으로 한 칩만 색이 들어옵니다(설문·쿠폰 필터는 그대로 값이 있는 고객을 모두 보여 줌). 월 옆 FC 선택창에는 본인 FC와, 관리자 화면 계정 카드의 ‘세일즈 페이지 접근 허용’에 체크한 FC만 나옵니다(하나뿐이면 선택창 없음). 관리자는 월 왼쪽 <b>사용자 선택</b>에서 아무 사용자나 고를 수 있고, 그러면 오른쪽 FC 선택창에 그 사용자의 본인 FC와 ‘세일즈 페이지 접근 허용’ FC만 나옵니다(그 사용자 화면과 같은 기준). 보고서는 세일즈에 쓰는 자료를 모아 둘 자리이며, 담을 내용은 정해지는 대로 채웁니다."]);
    const s_inflow = sec("DB 받은 FC와 진행하는 FC (유입FC, v514)", ["고객마다 <b>유입FC</b>(DB를 배분받은 FC) 칸이 있으면, 그 고객은 <b>유입FC의 이달목표 DB 목록</b>에 나옵니다. 담당FC(메인FC)가 다른 FC여도 DB 목록은 받은 FC 쪽입니다.", "DB 목록의 <b>보관 · 지옥 · 원복</b>(DB상태)은 유입FC 파일의 그 고객 줄에 저장됩니다. 영업단계 칩·전화·문자·상담기록은 담당FC 쪽에 남습니다(진행은 다른 FC가 할 수 있음).", "유입FC가 비어 있으면 예전처럼 담당FC 기준입니다."]);
    const s_od = OD_ON && sec("연체/실효 리스트 (2026-11-21)", ["세일즈 > 보고서의 <b>조건 보고서</b> 아래에 있습니다. 각 FC는 <b>본인 것만</b>, ESL(관리자·직급 ESL)은 <b>팀 전체</b>(FC를 고르거나 ‘전체’)를 봅니다. 자료는 각 FC 폴더 안에만 있어 드라이브에서 열어도 남의 것은 보이지 않습니다.", "<b>연체</b>: 가장 최근 확인일의 리스트를 ‘○월 ○일자 연체리스트’로 알려 주고 아래에 연체 계약 중 <b>최종납입 2~24회</b>만 보여 줍니다(회차를 모르면 넣고 ‘회차 모름’ 표시)(바로 전 확인일에 없던 계약은 <b>신규</b>, 실효리스트는 회차와 상관없이 모두). 제목 줄 오른쪽 끝에 <b>[PDF 다운]</b>. 연락처 번호를 누르면 한 번 더 물은 뒤 전화가 걸리고, 세일즈 화면 전화 버튼처럼 그 FC 이력에 통화 시도가 남습니다.", "<b>실효</b>: 그 아래 ‘○○○○년 ○월 실효리스트’와 실효 표(실효일이 그 달인 계약, 메인FC 기준, 지금 상태 = 실효 그대로·부활·해약), 제목 줄 오른쪽 끝에 <b>[PDF 다운]</b>. 위 요약 줄의 실효 숫자는 이 표의 건수입니다.", "‘전체’ PDF는 앞부분에 모든 FC를 이어 붙인 묶음(ESL이 볼 것), 뒷부분에 FC마다 새 쪽에서 시작하는 페이지(인쇄 뒤 FC별로 나눠 줄 것)가 들어 있습니다. FC 한 명을 고르면 데이터수집이 만든 PDF를 그대로 받습니다. 정식 확인일 기록이 생기기 전에는 시험 기록을 ‘시험 기록’으로 표시합니다."]);
    const s_stage = sec("영업단계 칩 (이관 · DB)", ["<b>다른 FC에서 진행 중이면(v515)</b>: 그 고객의 칩 값이 이 FC 데이터에 없으면, 같은 고객ID의 다른 FC 계약진행 중 가장 최근 단계가 칩에 보입니다(보이기만 하고 저장은 그대로). 이관받은 고객이 이관일 뒤 어느 FC에서든 새로 계약하면 PC 반영 때 그 이관은 자동으로 완료됩니다.", "이관·DB 목록의 고객 줄 오른쪽에 <b>통화 · AP · PC · 계약</b> 칩이 있고, 그 고객의 지금 영업단계인 칩만 색이 들어옵니다. 칩을 누르면 고객 시트의 ‘영업단계’ 칸에 누른 칩의 글자가 저장됩니다(칸이 없으면 새로 만듭니다). 이미 색이 켜진 칩을 다시 누르면 ‘고객의 영업상태를 제거하시겠습니까?’ 확인 뒤 빈칸이 됩니다."]);
    const s_nf = sec("전화·문자 필터 (이달이관 · 따행)", ["목록 위의 <b>전화</b>·<b>문자</b> 칸을 체크하면 그 달에 아직 전화(또는 문자) 시도 ✓가 없는 고객만 보입니다. 둘 다 체크하면 전화도 문자도 안 한 고객만 보입니다.", "체크를 끄면 전체가 다시 보입니다. 따행 줄은 ‘설문 ☐발송 ☐응답 | ☐쿠폰 | ☐전화 ☐문자’ 순서입니다. 발송·응답·쿠폰은 체크하면 그 날짜가 있는 고객만, 전화·문자는 반대로 아직 안 한 고객만 보여 주며, 함께 쓸 수 있습니다."]);
    const s_wake = sec("“몰리야”로 AI 부르기 (시험용)", ["오른쪽 아래에 떠 있는 <b>몰리(두더지) 선 그림 버튼</b>을 누르면 AI 창이 열립니다(아래 메뉴 맨 왼쪽 AI 버튼과 같음). 반투명이라 밑의 내용이 비쳐 보입니다. 누르면 AI 창이 열리면서 <b>바로 듣기(마이크)</b>가 켜져 말하면 됩니다. 버튼을 <b>꾹 눌러 끌면</b> 원하는 자리로 옮겨지고, 그 자리는 이 기기에 기억됩니다. 말만해요 안에서는 그 자리를 비워 두므로 나오지 않습니다.", "설정 > 화면의 <b>“몰리야”로 AI 부르기</b>를 켜면, 이 기기에서 앱 화면이 열려 있는 동안 마이크로 듣다가 <b>‘몰리야’</b>라고 부르면 AI 창이 뜹니다. ‘몰리야 내일 일정 알려줘’처럼 이어 말하면 그 말이 바로 AI에 들어가고, 이름만 부르면 AI 창의 마이크가 켜집니다.", "듣는 동안 화면 위에 작은 <b>‘몰리 듣는 중’</b> 표시가 나오고, 누르면 바로 꺼집니다. 화면이 꺼지거나 다른 앱으로 가면 듣지 않고, AI 창이나 다른 음성 입력을 쓰는 동안에도 멈춥니다. 안드로이드 크롬에서만 되며(아이폰 사파리는 거의 안 됨), 음성은 크롬이 구글 서버로 보내 글자로 바꿉니다. 상담 중에는 꺼 두세요. 설정은 이 기기에만 저장됩니다."]);
    const s_theme = sec("다크 모드", ["설정 > 화면의 <b>다크 모드</b>를 체크하면 어두운 화면, 끄면 밝은 화면으로 이 기기에서 고정됩니다. 한 번도 고르지 않았으면 <b>다크 모드가 기본</b>입니다(2026-11-18부터, 폰의 밝은·어두운 설정과 상관없음). 고른 값은 내 설정 파일(드라이브)에 저장되어 다른 기기에서도 같게 따르고, 이 기기에는 화면 깜빡임을 막으려고 같이 기억합니다. 로그인 화면도 같은 값을 따릅니다. 말만해요 안에서 열면 말만해요의 설정(다크/라이트)을 따르고, 말만해요에서 바꾸면 바로 바뀝니다(이 앱에는 저장하지 않음)."]);
    const s_dsh = sec("대시보드 (2026-11-21 바뀜)", ["아래 메뉴 <b>AI 오른쪽 [대시보드]</b>입니다. 보는 범위 선택(오른쪽 위)은 <b>ESL만</b> 보입니다: 맨 위 <b>팀</b>(팀 합계) → 김현성 · 임지섭 … 순서이고, 관리자는 처음에 <b>팀</b>을 봅니다. 그 밖의 계정은 본인 FC만 봅니다.", "<b>이달 실적</b>: 이번 달 <b>건수 · CMIP(월납환산보험료) · CANP</b> 세 칸입니다(금액은 원 단위). 청약 목록(계약 자료)에서 이번 달 청약 중 <b>철회·거절을 뺀 것</b>을 더하고, 세 값 모두 <b>기여율</b>을 곱해 더합니다. 팀은 팀원 합계입니다.", "<b>이달 청약상태</b>: 이달 실적 바로 아래에 이번 달 청약을 지난달 청약상태와 같은 칸으로 보여 줍니다.", "<b>지난달 청약상태</b>: 지난달 청약을 계약일자 최근 순으로 보여 줍니다(청약상태 · 심사상태 · 진단종류 · FC명(팀일 때만) · 상품명 · 계약일자 · 계약자 · 주피보험자 · CMIP · CANP). 표는 옆으로 밀어 봅니다. 철회·거절 줄은 흐리게, 금액은 계약 금액 그대로입니다.", "<b>지난달 FC별 실적</b>(팀일 때만): FC마다 지난달 철회·거절을 뺀 건수 · CMIP · CANP를 기여율을 곱해 보여 줍니다(예: 2건 중 1건이 기여율 50%면 1.5건). 맨 아래 합계.", "<b>최근 6개월 평균 실적</b>: 이번 달 전 6개월의 월평균 건수 · CMIP · CANP · 수수료(실지급금)입니다. 입사한 지 6개월이 안 된 FC는 입사한 달부터의 합계를 그 달 수로 나눕니다. 팀은 FC별 월평균의 평균입니다(더한 값이 아님). 팀을 볼 때는 그 아래에 <b>FC별 최근 6개월 평균</b> 표가 나옵니다(6개월이 안 된 FC는 이름 옆에 개월 수, 입사 전 FC는 빠짐). 지난달 FC별 실적에서도 그달에 입사 전인 FC는 빠집니다. 이 상자들은 민감정보 허용 계정만 보입니다.", "<b>최근 수수료</b>(가장 최근 달 실지급, 팀은 합계)와 <b>계약진행</b> 인원 카드가 그 아래에 있습니다.", "<b>이달목표 이행 점수</b>: (이관 중 통화 이상 + 따행 중 쿠폰발송 + DB 중 AP 이상) ÷ (이관 + 따행 + DB 목표 전체) × 100을 점수로 보여 주고(80%면 80점), 그 아래 한 줄 막대가 진도입니다. 계약진행 막대는 최근 3개월에 시작한 계약진행의 고객등록·전화·AP·PC·계약별 건수입니다. 위쪽 <b>계약진행 카드</b>는 ‘이달 계약 건 / 최근 3개월 계약진행 건’입니다. 최근 3개월에 시작한 계약진행 중 이번 달에 계약한 건(진행 단계가 이달 계약이거나, 그 고객의 이번 달 청약이 있음 — 청약철회·거절도 계약한 뒤 일이라 계약으로 셈. 여러 FC가 나눈 청약은 메인FC(계약자의 담당FC)에게만 셈)이 앞의 초록 숫자입니다. 목표는 보관하지 않은 이관·따행(목표)·신규 DB입니다. 맨 아래는 <b>소개 랭킹</b> 상위 3명입니다."]);
    const s_in = sec("소개 (2026-11-17)", ["<b>새 고객 확인의 [없음](v567)</b>: 새로 들어온 고객 줄의 [없음]을 누르면 그 고객은 소개자 없음(소개 아님)으로 저장되고, 다음부터 소개자를 묻지 않습니다. [나머지는 소개 없음]은 남은 고객을 한꺼번에 같은 방식으로 저장합니다. [다음에]는 이번 로그인 동안은 묻지 않고, 다음 구글 로그인 뒤에 다시 묻습니다.", "<b>고객 정보 > 수정</b>의 <b>소개</b>: <b>소개 · 아님 · 모름</b> 중 고르고, <b>소개</b>이면 소개자 이름을 적습니다. 같은 FC의 등록 고객 후보(이름·출생 두 자리·유입경로)가 나오고, 고르면 <b>등록 고객과 연결</b>(소개고객ID 저장), 고르지 않거나 후보에 없으면 <b>이름만 남기기</b>(소개자명 저장)입니다. 다시 누르면 풀리고 이름을 비우면 지워집니다. 소개만 고치면 ‘직접 수정’ 표시는 붙지 않습니다.", "<b>소개자 트로피</b>: 소개해 준 고객 이름 앞에 1~2명 동, 3~4명 은, 5명 이상 금 트로피가 붙습니다(소개고객ID를 세어 계산, 고객 정보 머리·고객 검색 결과·소개찾기에 나옵니다).", "<b>소개찾기</b>(고객 화면 GBS 옆 버튼): 등록 고객 전부를 대상으로 ‘등록고객 N명 중 M명 확인’을 보여 줍니다. <b>소개고객부터</b>는 고객마다 [소개자 찾기](창에서 이름 검색 → 고객 고르기 · 이미 적힌 이름 · 이름만 · 소개 아님 · 모름), <b>소개자부터</b>는 소개자를 먼저 고른 뒤 그분이 소개한 고객을 체크해 한 번에 연결, <b>랭킹</b>은 소개한 수 순서로 보여 주며 고객이 아닌 소개자는 ‘미등록’으로 묶고 [고객으로 연결]로 한 번에 고객ID에 이어 줍니다. 묶음 <b>지역·성별·직업·직장·유입 시기·가족(같은 주소)·유입경로</b>를 겹쳐 눌러 같은 사람끼리 묶어 볼 수 있고, ‘확인한 고객도 보기’를 켜면 이미 확인한 고객도 나옵니다.", "<b>새 고객 확인</b>: 지난번 접속 뒤 입력된 내 FC 고객이 있으면 로그인 때 소개찾기 화면이 먼저 뜹니다. 고객마다 [소개자 찾기], 모두 끝내려면 <b>나머지는 소개 없음</b>, 미루려면 <b>나중에</b>(다음 접속 때 다시 뜹니다). 처음 접속한 기기는 기준 시각만 기록하고 묻지 않습니다(기존 고객은 소개찾기에서 점검)."]);
    const s_smemo = sec("영업 상담기록 칸 (2026-11-18)", ["<b>전화·문자 시도 기록(2026-11-21)</b>: 고객에게 전화·문자를 누르면 상담기록에 ‘<b>○○에서 전화 시도</b>’ / ‘○○에서 문자 시도’가 저절로 남습니다(○○ = 누른 화면: 연체리스트·실효리스트·이달목표 이관/따행/DB·고객 목록·고객 정보·일정). 세일즈 화면과 연체/실효 리스트는 ‘정말로 거시겠습니까?’에 확인을 눌렀을 때만 남습니다. 전화·문자 앱이 열린 사이에는 이 폰에 적어 두었다가 앱으로 돌아오면 저장합니다.", "세일즈 화면(이달목표의 이관·따행·DB, 계약진행)의 고객 줄 맨 아래에 <b>가장 최근 상담기록 한 줄</b>이 나오고, 그 오른쪽에 <b>[수정]</b> · <b>[삭제]</b>가 있습니다. 수정하면 처음 날짜·시간은 그대로이고 내용만 바뀌며, 삭제는 확인을 거쳐 엑셀에서 지워집니다. 최근 1년 안의 상담기록이 2건 이상이면 오른쪽 끝에 <b>[더보기]</b>가 생겨 전체를 팝업으로 볼 수 있습니다(보기 전용).", "<b>일정 화면의 고객 카드</b>도 같은 상담기록 줄입니다: 카드 맨 아래 최근 상담기록 <b>최대 5줄(날짜와 함께)</b>과 [수정]·[삭제]·[더보기], 고객 정보 칸(주소 등 글자가 있는 곳)을 터치하면 위에 추가칸(마이크 포함)과 최근 6개월 상담기록이 펼쳐지고 한 번 더 터치하면 접힙니다.", "고객 정보가 있는 줄(버튼·링크가 아닌 곳)을 <b>한 번 터치</b>하면 아래로 상담기록 칸이 펼쳐집니다. <b>맨 위 줄은 상담기록 추가칸</b>([저장]을 누르면 일정 화면의 한 줄 메모처럼 상담기록에 저장되어 고객 정보 > 상담 탭에서도 보임)이고, 그 아래에 <b>최근 6개월 안의 상담기록이 날짜와 함께</b> 나옵니다(각 기록에 [수정]·[삭제], 6개월보다 오래된 것은 ‘6개월 이전 N건 더보기’로 팝업). <b>한 번 더 터치</b>하면 위로 올라가며 사라지고 최근 상담기록 한 줄만 남습니다. (예전 설정의 ‘영업 메모 칸’ 옵션은 없앴고 항상 사용합니다.)"]);
    const s_mat = sec("세일즈 > 자료 (v554, 2026-11-22)", ["세일즈 위 서브메뉴 맨 오른쪽 <b>자료</b>: 팀이 함께 보는 상품 안내장·설계안·교육자료·모의고사·서식 파일 목록입니다. 위쪽 칩(전체·분류별, 옆 숫자 = 파일 수)을 누르면 그 분류만, 검색 칸에 글자를 넣으면 이름으로 바로 걸러집니다.", "<b>줄을 누르면 바로 열립니다</b>(PDF는 새 창, 사진은 화면 안). <b>[공유]</b>를 누르면 파일을 받아 휴대폰 공유창(카톡·문자 등)이 바로 열립니다. 큰 파일이라 공유창이 바로 안 열리면 버튼이 <b>[보내기]</b>로 바뀌고, 한 번 더 누르면 열립니다.", "관리자는 <b>[올리기]</b>로 파일을 올립니다. 고른 분류 폴더에 저장되고, ‘전체’일 때는 ‘기타’로 들어갑니다. 파일은 드라이브 <b>app › 신한FC › 설정 › 자료</b>에 있고, 그 안에 폴더를 만들면 분류가 하나 생깁니다(폴더 이름 앞 숫자는 순서용으로 화면에선 빠짐)."]);
    const s_pg = sec("세일즈 > 계약진행 (2026-11-17)", ["<b>2026-11-18부터</b> 계약진행의 보기는 <b>진행 · 보관 · 완료</b> 세 가지입니다. 예전의 ‘삭제’는 ‘완료’로 합쳐, 전에 삭제로 둔 고객도 완료에 보입니다. 보관에서는 [진행으로]·[완료로], 완료에서는 [진행으로]·[보관으로]로 옮깁니다.", "세일즈 화면 위 서브메뉴가 <b>이달목표 · 계약진행 · 자료</b>입니다. <b>계약진행</b>은 설계를 시작한 고객 한 줄이 한 회차이고, 이름 옆에 유입경로(신규유입·이관유입·DB유입), 아래에 단계 5칸 <b>고객등록 · 전화 · AP · PC · 계약</b>(지금 단계만 색)과 시작일·단계일이 보입니다. 단계를 눌러 바꾸고(건너뛰어도 됩니다), <b>계약</b>을 누르면 상태가 <b>완료</b>가 됩니다.", "목록 위 칩 <b>진행 N · 보관 N · 삭제 N · 완료 N</b>(기본은 진행)으로 상태별 목록을 봅니다. <b>완료</b>는 단계가 ‘계약’이 되어 끝난 회차입니다(계약이 되면 진행 목록에서 완료로 옮겨지고, 다른 단계를 누르면 진행으로 되돌아옵니다). 줄 오른쪽 버튼은 진행·완료 줄은 <b>보관으로</b>, 보관 줄은 <b>진행으로 · 삭제로</b>, 삭제 줄은 <b>진행으로 · 보관으로</b>입니다. 삭제는 자료를 지우지 않고 삭제 칩으로 옮기는 것이며 확인창이 뜹니다. 바꿀 때마다 이력에 ‘계약진행’ 한 줄이 남고, 6초 안에 [되돌리기]가 나옵니다. 고객 화면의 <b>[진행] 칩은 없어졌고</b> 계약진행은 이곳 한 곳에서만 봅니다.", "<b>이달목표</b>의 이관·따행 목록은 이제 달 이동 없이 <b>목표</b> 상태 기준입니다. 이번 달 것이 위, 달이 지나도 보관·삭제하지 않은 것은 아래 ‘지난 목표’에 남습니다."]);
    const s_sup = sec("지원 기능 (관리자)", ["관리자 화면 ‘문자 문구 (팀 기본)’ 바로 위의 <b>지원 기능</b> 박스에 있는 <b>[고객 등록]</b>을 누르면 새 고객 등록 페이지가 새 탭으로 열립니다. 담당 FC·이름·연락처·유입경로(신규유입/DB유입)를 넣어 저장하며, 같은 연락처나 같은 신한라이프 고객번호가 이미 있으면 멈춥니다. 등록 직후에는 고객 목록이 새로 읽을 때까지 보이지 않을 수 있습니다."]);
    const s_sms = sec("문자 문구", ["설정 화면의 ‘문자 문구’ 카드는 접혀 있으니 제목을 눌러 펼칩니다. 세일즈 화면에서 [문자]를 누르면 이 문구가 채워진 문자 작성 화면이 열립니다. 내가 저장한 문구가 있으면 그것을 쓰고, 없으면 관리자가 정해 준 팀 기본 문구(관리자 화면의 ‘문자 문구 (팀 기본)’)가 미리 들어 있습니다. 고치거나 지우면 내 설정으로 저장됩니다. 지우면 문구 없이 작성 화면만 열립니다.", `쓸 수 있는 항목: <b>${SMS_VARS}</b> — {고객명}=계약자 이름, {FC명}=담당 FC 이름, {상품명}, {이관일}(예: 11월 9일, 따행 목록에서는 비어 있음)`, "문자는 작성 화면을 연 뒤 직접 '보내기'를 눌러야 발송되며, 앱에는 '시도'만 기록됩니다. 일부 폰·문자앱·카카오톡 안의 브라우저에서는 문구가 채워지지 않을 수 있고, 45자를 넘으면 장문 문자로 나가거나 잘릴 수 있습니다."]);
    const s_db = sec("밸류업DB 목록·보고서 (v570, 2026-11-24)", ["<b>이달목표 › DB</b>: 배분 건마다 한 줄 — 이름·성별/나이 / 연락처와 [전화][문자] / 유입회사·유입일·지역(v574) / 유입메모(SOL DB 메모, 있을 때만). 정렬은 최신 유입일시가 위, 같으면 이름 가나다순입니다(유입일시가 없는 줄은 아래에 배분일 최신순).", "<b>세일즈 › 보고서</b>에서 DB를 고르면 DB 표는 이름·성별/나이·연락처·유입회사·유입일·유입메모·지역 칸이고 가운데 정렬입니다(v573: 유입DB명 칸 없음, 유입메모는 왼쪽 정렬, 유입메모·지역은 잘리지 않고 칸 안에서 줄바꿈 — PDF는 넘치면 다음 쪽으로). <b>v575</b>: 보고서는 고른 단계마다 따로 만들고(신규 → 계약 → 완료 → 지옥 순), 제목에 단계를 붙입니다(예: 2026년 11월 밸류업DB 리스트 - 신규). PDF는 쪽마다 고르게 나누고 줄 높이를 늘려 아래쪽 빈 공간을 채웁니다. <b>PDF</b>는 A4 가로 한 쪽 폭, 쪽마다 12명, 맨 왼쪽에 손글씨 메모 칸과 성향·통화·미팅 칸이 있습니다. <b>엑셀</b>도 같은 칸 순서입니다.", "성별/나이는 주민번호 앞 7자리로 계산하고, 유입일은 회사 유입 날짜(FC 배분일과 다름)입니다.", "<b>DB 시트(v571)</b>: 밸류업DB는 배분 1건마다 'DB' 시트 한 줄(배분받은 FC 파일)에 있습니다. DB 시트에 줄이 있는 FC는 그 시트를 읽고, [보관으로]·[지옥으로]·[신규로]도 그 줄의 DB상태·완료일에 씁니다(이력에 DB 번호를 함께 적음). 아직 DB 시트가 없는 FC는 예전처럼 고객 시트 칸을 읽습니다.", "<b>진행FC 표시(v572)</b>: DB를 배분받은 FC와 계약진행을 맡은 FC가 다르면, 배분받은 FC의 DB 목록에만 나오고 그 줄 이름 옆에 <b>진행: FC이름</b>이 붙습니다. 진행하는 FC는 계약진행 목록에서 봅니다. DB 건수·3개월 자동완료는 배분받은 FC 기준으로 한 번만 셉니다."]);
    const s_ct = sec("설정 > 다양한 기능 > 주소록 관리 (v568, 2026-11-24)", ["설정 맨 위 <b>다양한 기능</b> 상자에 앞으로 여러 기능 버튼이 모입니다. 첫 번째가 <b>주소록 관리</b>입니다.", "<b>주소록 넣기</b>: 이번 달 이관 고객을 <b>내 구글 주소록</b>(로그인한 계정)에 '202611이관' 같은 라벨과 '신한FC' 라벨을 붙여 넣습니다. 이름·휴대전화·주소·메모 한 줄(상품·보험료·담당)을 넣고, 휴대전화가 같은 연락처가 이미 있으면 새로 만들지 않고 라벨과 빈 칸만 채웁니다. 번호가 없는 고객은 넣지 않습니다.", "<b>주소록 정리</b>: 앱이 넣은 연락처(신한FC 라벨)만 보여 주고, 이달목표에서 빠진 고객을 먼저 골라 둡니다. 고른 것만 지우며, 구글 주소록에서 <b>영구 삭제</b>되어 되돌릴 수 없습니다(고객 정보는 그대로).", "처음 누를 때 구글 창에서 주소록 권한을 한 번 허용합니다(이 앱에서만, 1시간마다 다시 연결). 관리자가 구글 콘솔에서 People API를 켜 두어야 합니다."]);
    const s_ntf = sec("중요 알림 · 오전 6시 로그아웃 (v557, 2026-11-24)", ["<b>중요 알림</b>: 로그인한 뒤 그날 처음 앱을 열면, 오늘에 해당하는 알림이 한 창에 모두 나옵니다. 버튼은 <b>[확인]</b> 하나뿐이고, 누르면 알림에 정해 둔 페이지로 이동합니다(시스템 알림이 먼저, 그다음 목록 위쪽 알림). 한 번 확인한 알림은 다시 나오지 않습니다(요일 알림은 그날만).", "<b>연체·실효 알림(v564)</b>: 내 FC의 연체리스트(나온 지 7일 안)나 실효리스트(이번 달·지난달)가 새로 나오면 로그인 직후 알림이 뜨고, [확인]을 누르면 세일즈 › 보고서의 연체/실효 리스트로 바로 갑니다. 한 번 확인하면 다시 뜨지 않고, 다음에 앱을 열면 첫 화면(대시보드)부터 보입니다. 연체나 실효가 없으면 PDF 다운 버튼은 숨깁니다.", "<b>시스템 알림</b>: 내 FC에 새 이관 계약이나 새 밸류업DB가 생기면 자동으로 알립니다. 처음 한 번은 지금 목록을 기억만 하고, 그 뒤로 새로 생긴 건수를 알려 줍니다.", "<b>관리자 › 전체 알림</b>(관리자만): 알림 추가·수정·삭제, 알림마다 켜기·끄기, 시스템 알림 전체 켜기·끄기(기본 켬). 알림마다 날짜(시작~끝) 또는 요일, 대상(전체·FC별), 확인하면 이동할 페이지를 정합니다. 날짜 알림에는 확인한 사람 수가 나옵니다.", "<b>오전 6시 로그아웃</b>: 매일 오전 6시가 지나면 모든 앱(신한FC·말만해요 등 같은 로그인을 쓰는 앱)이 로그아웃되고, 다음 로그인 때 구글 계정 고르기 창이 한 번 나옵니다. 구글 비밀번호를 다시 넣게 하지는 못하며, 폰 시계를 기준으로 합니다."]);
    const s_login = sec("로그인 (공통 로그인)", ["<b>로그인 승인/거부(v513)</b>: 신한FC 로그인은 관리자 화면 FC 사용 설정의 계정 카드 오른쪽 <b>[승인]/[거부]</b> 토글로만 정합니다(이메일·사번이 있다고 저절로 들어가지 않음). 거부인 계정으로 로그인하면 로그인 화면에 ‘아직 신한FC 로그인이 승인되지 않았습니다’가 뜨고 [다른 계정으로 로그인]이 보입니다. 승인으로 켜면 그 계정의 FC 폴더(본인 FC와 체크한 FC)와 설정 폴더를 Google Drive로 공유합니다. 거부로 끌 때는 공유도 해제할지 묻습니다 — 공유를 남기면 앱은 막혀도 드라이브에서 파일을 직접 열 수 있습니다. 이메일이 아직 없는 사람은 조직도에서 이메일을 먼저 넣어야 승인할 수 있습니다.", "<b>새 팀원</b>: 팀관리 > 후보자정보에서 [전환]하면 팀원정보에 옮겨지고, FC 사용 목록에 ‘이메일 미등록’으로 올라가며 로그인은 거부로 시작합니다. 팀원정보에 등록되어 데이터가 있는 FC는 계정이 없어도 관리자 화면(이달목표 사용자 선택·고객 FC 선택)에서 볼 수 있습니다.", "한 번 로그인하면 같은 브라우저(크롬·사파리 등)에서 다른 내 앱(차차차·항상일상·말만해요·메뉴메이커)도 다시 로그인하지 않고 바로 열립니다. 로그인은 약 1시간 유지되고, 끊기면 아무 앱에서 [다시 연결]을 한 번 누르면 나머지도 이어집니다.", "다른 앱에서만 로그인했고 이 앱에 필요한 권한이 모자라면 버튼이 ‘구글 연결 (권한 추가)’로 바뀝니다. 한 번 눌러 허용하세요.", "<b>로그아웃은 모든 앱에서</b> 함께 됩니다(오른쪽 위 전원 버튼). 아이폰은 홈 화면에 추가한 앱마다 저장 공간이 따로라 서로 나눠 쓰지 못합니다."]);
    const s_go = sec("말만해요 연결", ["말만해요에서 “김현성 고객 검색해줘”처럼 말하면 말만해요 화면 안에 이 앱의 고객 검색 목록이 나타납니다. 한 명만 나와도 목록이 먼저 보이고, 고객을 누르면 그 안에서 고객 정보가 열립니다. 이때는 앱 전체가 아니라 그 페이지만 보입니다(위쪽 바·아래 메뉴 없음). 목록→고객 정보, 정보·계약·분석·상담 탭 이동은 그대로이고, 뒤로가기는 말만해요의 왼쪽 위 버튼이 한 단계씩 되돌립니다.", "말만해요의 [앱으로 열기]를 누르면 보던 바로 그 화면(검색 목록 또는 그 고객 정보)으로 신한FC m이 전체 화면으로 열립니다. 말만해요 안에서는 로그아웃 버튼이 없고(로그아웃은 말만해요에서), 로그인이 풀렸으면 ‘말만해요에서 다시 로그인’ 안내가 나옵니다.", "“오늘 일정 보여줘”처럼 일정을 물으면 일정 화면(하루 보기)만 말만해요 안에 나타납니다. 날짜 이동·추가 버튼은 그대로 쓰고, [앱으로 열기]를 누르면 보던 날짜로 열립니다.", "검색한 이름은 주소의 # 뒤에만 실려 서버로 전송되지 않습니다."]);
    const s_map = sec("고객 주소 지도", ["고객 정보의 주소 오른쪽 <b>위치 아이콘</b>을 누르면 카카오맵에서 그 주소를 찾은 지도가 새 창으로 열립니다. 누를 때만 주소가 카카오맵에 전달됩니다."]);
    const s_inst = sec("앱 설치", ["안드로이드 크롬: 설정 화면의 [앱 설치](없으면 브라우저 메뉴의 ‘앱 설치’), 아이폰 사파리: 공유 → ‘홈 화면에 추가’. 설치하면 앱처럼 전체 화면으로 열립니다. 이미 설치했으면 설정에 설치 카드가 보이지 않습니다."]);
    const s_srch = sec("전체검색 · 고객검색", ["<b>위쪽 전체검색</b>: 고객 이름·연락처·주소·미팅장소·증권번호뿐 아니라 고객 메모, 상담기록, 오늘 기준 1년 전~1개월 뒤 일정(제목·장소)에서도 찾습니다. 결과는 아래에서 올라오는 창에 고객(어디서 찾았는지 표시)과 일정으로 나뉘어 나오고, 누르면 고객 정보(상담기록에서 찾았으면 상담 탭) 또는 그날 일정으로 이동합니다.", "<b>고객 화면의 고객검색</b>: 고객 이름·연락처·주소·미팅장소·증권번호로 고객만 찾는 기존 검색입니다. 한 명만 나오면 바로 고객 정보가 열립니다.", "검색칸이 비어 있을 때의 글자(전체검색·고객검색)는 오른쪽에 보입니다."]);
    const s_cov = sec("보장 요약 (고객 정보 > 분석 탭)", ["분석 탭 맨 위에 사망 · 3대진단(암·뇌·심) · 수술 · 입원 · 간병을 간단히 보여 줍니다. 스마트보장분석이 있으면 ‘전체 보험사에 가입된 보장 요약입니다.’, 신한라이프 상세보장내용만 있으면 ‘신한라이프에 가입된 보장 요약입니다.’라고 위에 적고, 맨 아래에 어느 문서의 몇 월 며칠 기준인지 나옵니다.", "‘미가입’은 가입한 보장이 없다는 뜻이고, ‘확인 안 됨’은 원본 문서에서 그 항목을 읽지 못했다는 뜻입니다. 이 값은 고객 등록·업데이트 때 미리 만들어 고객 시트의 ‘보장요약’ 칸에 저장해 둔 것을 읽기만 하므로 분석 탭을 열 때 계산하지 않습니다.", "앱을 켜 둔 사이에 새로 만든 고객 폴더(오늘 등록하거나 보장분석한 고객)는 그 고객을 열 때 폴더 목록을 다시 읽어 찾아옵니다(v556)."]);
    const s_cmp = sec("고객 구분 · 유입경로 말 (2026-11-17)", ["화면의 말이 <b>직접모집</b>(예전 ‘계약·본인’), <b>이관고객</b>(예전 ‘이관’), <b>계약진행</b>(예전 ‘진행’), <b>가망</b>으로 바뀌었습니다. 고객 정보의 <b>유입경로</b>는 <b>신규유입 · 이관유입 · DB유입</b> 세 가지이고 처음 한 번 정해지면 바뀌지 않아 <b>읽기만</b> 합니다(예전 ‘계약·지인·소개’ 값은 신규유입으로 보입니다). 소개는 유입경로가 아니라 <b>소개자</b> 칸으로 따로 둡니다.", "데이터가 새 칸 이름(주민번호·고객분류·이관상태·따행상태·계약진행 표 등)으로 바뀌는 중에는 앱이 <b>새 칸이 있으면 새 칸, 없으면 옛 칸</b>을 읽고, 저장은 옛 칸과 새 칸 <b>둘 다</b>에 합니다(이관보관·이관상태, 따행보관·따행상태, 진행제외·영업단계와 계약진행 표, DB제외·DB상태, 후번호와 주민번호). 따행 <b>방문완료</b>는 방문완료일만 보고, 따행 <b>대상</b>은 따행상태(목표, 빈칸이면 활동기한이 안 끝난 줄)로 셉니다. 상령일(생년월일+6개월)과 납입완료 예정(최초계약일+납입기간)은 칸이 없어도 계산해서 보여 줍니다."]);
    const s_prog = sec("보관 (이동 · 원복)", ["<b>[지옥으로]</b>를 누르면 옮긴 뒤 지옥문 그림이 1초 동안 화면 전체에 나옵니다.", "<b>지옥으로 보낸 계약·계약자</b>(이관·따행·DB)는 고객 <b>성향이 자동으로 Shit</b>이 됩니다. 되돌리기를 누르면 예전 성향으로 돌아가고, 다른 단계로 옮겨도 성향은 그대로 두니 필요하면 고객 정보에서 바꾸세요.", "<b>2026-11-18부터 이달목표의 이관·따행</b>은 위 칩이 <b>신규 · 완료 · 지옥</b>입니다(예전 ‘이관/따행 · 보관’). 신규 줄은 <b>[완료]</b>, 완료 줄은 <b>[신규로] [지옥으로]</b>, 지옥 줄은 <b>[신규로] [완료로]</b>로 옮깁니다. 예전에 보관한 고객은 완료에, 삭제한 고객은 지옥에 보입니다. 파일에는 지금처럼 목표 · 보관 · 삭제로 저장됩니다(이력에 남음). DB 탭도 같은 말로 <b>신규 · 계약 · 보관 · 지옥</b>입니다(예전 ‘삭제’ = 지옥, 보관 줄의 [삭제로] = [지옥으로]).", "세일즈 > 이달목표 목록 줄의 오른쪽 버튼 글자는 상태에 따라 달라집니다. 신규 목록 줄은 <b>보관으로</b>, 보관 목록 줄은 <b>신규로</b>(원래 목록으로)와 <b>지옥으로</b>(DB), 지옥 목록 줄은 <b>신규로</b>와 <b>보관으로</b>입니다. 아래 설명의 ‘이동·원복’은 각각 ‘보관으로·신규로’를 뜻합니다.", "고객을 지우지 않고 목록에서 <b>보관</b>으로 옮기는 기능입니다. 고객 화면의 [진행]·[DB] 목록, 세일즈 > 이달목표의 <b>이관 · 따행 · DB</b> 목록 줄 오른쪽 작은 <b>이동</b>을 누르면 ‘OO 고객을 보관으로 이동할까요? 고객 정보는 그대로예요.’라고 묻고, 확인하면 그 목록에서 바로 빠져 보관으로 갑니다(저장은 뒤에서 하며 왼쪽 아래에 작게 ‘저장 중…’, 저장이 안 되면 원래대로 돌아오고 알림). 6초 안에 나오는 [되돌리기]로 취소할 수 있습니다.", "목록 위 ‘N명’ 줄 오른쪽 작은 칩 <b>진행 N / 보관 M</b>(DB는 <b>DB N / 보관 M</b>, 이관은 <b>이관 N / 보관 M</b>, 따행은 <b>따행 N / 보관 M</b>)으로 보관된 고객만 볼 수 있고, 보관 목록 줄의 <b>원복</b>을 누르면 확인 없이 원래 목록으로 돌아갑니다. 줄에는 ‘보관 11-15’처럼 날짜가 작게 보입니다. DB DB 목록 위 칩은 <b>신규 / 계약 / 보관</b> 순서이고(신규 검정 · 계약 초록 · 보관 파랑), <b>보관</b>을 누르면 빨간 <b>지옥</b> 칩이 더해집니다. 계약 고객이 된 DB 고객은 고객 시트의 ‘DB상태’가 ‘계약’이 되고 계약 칩에서 봅니다. <b>지옥</b> 목록은 가장 최근에 지옥으로 보낸 고객이 맨 위에 나오며 줄의 <b>복구</b>를 누르면 보관으로 돌아갑니다. 보관 목록에는 원복 오른쪽에 <b>지옥으로</b> 버튼도 있습니다: 확인 뒤 누르면 자료는 지우지 않고 고객 시트의 ‘DB상태’ 칸(신규·보관·삭제 중 하나, 빈칸은 신규)에 ‘삭제’(=지옥)만 표시해, 보관 목록에서도 안 보이게 합니다(6초 안에 [되돌리기] 가능).", "기준: 진행 = 고객 단위(고객 시트 ‘진행제외’ 칸), DB = 고객 단위(‘DB제외’ 칸, 계약 고객이 된 DB 고객은 저절로 보관에 ‘계약’으로 보임), 이관 = 계약 단위(계약 시트 ‘이관보관’ 칸, 같은 고객의 다른 계약은 그대로), 따행 = 그 달 대상 단위(따행 시트 ‘따행보관’ 칸, 다음 달 대상에는 다시 나옴). 이동·원복할 때마다 이력에 ‘이동(보관)’·‘원복’ 한 줄이 남습니다. 진행 고객은 고객 정보 머리 카드 아래 <b>보관으로 이동</b>/<b>원복</b> 버튼으로도 할 수 있습니다. 고객 화면에서 필터를 여러 개 함께 켜도(예: DB + 가망) 보관할 수 있는 고객(진행 필터의 진행 고객, DB 필터의 DB 고객)에만 이동·원복 버튼이 보이고, 가망 같은 다른 고객에는 나오지 않습니다. 고객 화면 [이관] 목록에서는 <b>이달 N / 보관 M</b> 칩(이달 이관 고객 수)과 줄의 이동·원복으로 그 고객의 <b>이달 이관 계약만</b> 한꺼번에 보관하거나 원복합니다(이관일자가 이번 달인 계약 = 이달목표 > 이관에 나오는 계약, 지난 달 이관·담당이 다른 FC 계약은 그대로). 이달 이관 계약이 없는 이관 고객에는 버튼이 없습니다. 고객 정보 화면 탭 줄 오른쪽에도 보관할 수 있는 목록마다 <b>보관으로 이동</b>/<b>원복</b> 버튼이 나옵니다(여러 개면 ‘진행 보관’·‘DB 보관’·‘이관 보관’처럼 앞에 목록 이름). 그 FC를 수정할 수 있는 계정만 쓸 수 있습니다."]);
    const s_rh = sec("오른손 모드", ["'화면' 카드 오른쪽 칸입니다. 하단 메뉴, 탭, 선택 버튼을 오른쪽으로 모아 엄지로 누르기 쉽게 합니다. 이 폰에만 적용되고, 관리자 조직도에서 볼 수 있게 기록됩니다."]);
    const s_bk = sec("데이터 업데이트 중 잠금 (v550)", ["PC가 팀 데이터를 모으는 동안(그 FC 폴더에 업데이트 표시가 있는 동안) 화면 맨 위에 노란 띠 <b>데이터 업데이트 중</b>이 뜹니다.", "그동안은 <b>보기와 새 기록만</b> 됩니다. 전화·문자 시도는 따로 모아 두었다가 업데이트가 끝나면 자동으로 붙이고, 상담기록 추가는 바로 저장됩니다.", "고객 정보 수정·단계 변경·완료·상담기록 수정/삭제는 띠가 사라진 뒤 다시 해 주세요(1분마다 확인).", "업데이트 표시가 30분 넘게 새로 고쳐지지 않으면 끝난 것으로 보고 다시 열립니다."]) + sec("데이터 자동 백업·복구 (v509)", ["<b>자동 백업</b>: 앱이 고객 데이터 파일을 저장하기 직전에, 바로 앞 내용을 같은 폴더의 <b>_앱백업</b> 폴더에 사본으로 남깁니다. 계정마다 최근 10개만 두고 오래된 것은 휴지통으로 보냅니다. 사본은 원본과 같은 폴더 안에 있어 원본과 같은 사람만 볼 수 있습니다.", "<b>자동 복구</b>: 고객 데이터 파일이 깨져 있으면(모양이 망가짐) 앱이 Drive 버전 기록에서 가장 최근의 정상본을 찾아 자동으로 되돌리고 화면에 알립니다. 깨진 내용은 지우지 않고 _앱백업 폴더에 ‘깨진본’으로 남깁니다. 데이터 반영 중이거나 다른 곳에서 복구 중이면 손대지 않습니다.", "되돌린 뒤에는 이 기기에서 최근 14일 동안 보낸 <b>전화·문자 기록</b>을 다시 반영합니다(이미 있는 기록은 건너뜀). 다른 기기에서 한 기록이나 칩·정보 수정은 다시 반영되지 않으니, 알림이 뜨면 최근에 고친 내용을 한 번 확인해 주세요.", "줄 수만 크게 줄어든 경우는 되돌리지 않고 알리기만 합니다. 자동 복구를 못 했다는 알림이 뜨면 관리자에게 알려 주세요."]);
    const s_save = sec("저장", ["저장할 내용이 하나라도 있으면 아래 [저장]이 켜집니다. 누르면 바뀐 모든 카드(시작 화면·문자 문구·캘린더·사용자별 일정 구분과 시작 화면·FC 권한 등)를 한꺼번에 저장합니다. 각 카드 안의 [저장]을 눌러도 됩니다."]);
    const s_perm = sec("FC 사용 설정 (권한)", ["[추가]를 누르면 확인창 뒤 바로 저장되어 조직도에 나옵니다. 이름(FC명)을 FC 폴더 이름과 같게 적으면 본인 FC(보기+수정)로 시작하고, 같은 폴더가 없으면 권한 없이 등록되어 조직도에서 받은 권한으로 열립니다. 다른 FC의 고객·세일즈 페이지 접근과 민감정보 권한은 계정 카드가 아니라 <b>팀관리 메뉴의 [조직도]</b>에서 정합니다(2026-11-15부터)." + (GD ? " 저장하면 Google Drive 폴더 공유도 함께 바뀝니다." : ""), "<b>구글 캘린더</b>: 설정 화면 맨 위 ‘구글 캘린더’ 카드에서 관리자의 기본 캘린더와 영업·정보 캘린더를 고르고 그 카드의 [저장]을 누릅니다.", "<b>민감정보</b>: 조직도에서 체크한 사용자만 주민번호·계좌를 보고 고칠 수 있습니다.", "<b>세일즈 페이지 접근</b>: 세일즈(이달이관·따행) 선택창에 나올 FC로, 조직도의 보기·수정대로 정해집니다. 관리자는 세일즈 화면의 사용자 선택에서 사용자를 고르면 그 사용자의 FC를 봅니다. 관리자는 세일즈 화면의 사용자 선택으로 모든 사용자를 볼 수 있어 관리자 줄의 세일즈 페이지 접근 허용은 없앴습니다.", "관리자는 모든 FC를 볼 수 있어서 별도 체크가 필요 없습니다.", "<b>조직도</b>: 아래 메뉴 <b>팀관리</b>의 [조직도]를 누르면(2026-11-19에 설정에서 옮김) 관리자(ESL) 밑으로 조직을 짭니다. 이름 왼쪽 점 6개(⠿)를 끌어 다른 사람 위에 놓으면 그 사람 밑으로 옮겨지고, 이름 아래 <b>보기</b>를 켜면 그 사용자의 FC 고객을 위 사용자들(관리자 제외, 바로 위부터 맨 위까지)이 볼 수 있고, <b>수정</b>을 켜면 수정 권한까지 위임합니다(수정을 누르면 보기도 함께 켜짐). 카드 아래 ‘받은 권한’은 아래 사용자들에게서 받은 FC입니다. 카드 아래 작은 글씨에 그 결과 FC가 보이고, 아래 오른쪽 [저장]을 누르면 바뀌는 권한을 확인창으로 보여 준 뒤 ‘고객 페이지 접근 허용’(보기 = 보기만, 수정 = 보기+수정)·‘세일즈 페이지 접근 허용’(보기·수정 모두 아래 조직 FC 추가)과 Google Drive 폴더 공유를 함께 바꿉니다. 조직도가 유일한 기준이라, 저장하면 각 사용자의 권한이 ‘본인 FC + 조직도 권한’으로 맞춰지고 그 밖의 권한은 빠집니다(권한이 조직도와 다르면 조직도 위에 안내가 나오고 [저장]이 켜짐). 이름 아래 <b>민감정보</b>는 그 사용자 본인만의 주민번호·계좌 보기·고치기 허용이고(아래 조직에는 내려가지 않음), 계정 카드의 ‘민감정보 접근 허용’과 같은 값입니다. 이름 옆 드롭박스로 <b>직책</b>(ESL·SSL·GSL·FC)을 정하고 [저장]합니다. 직책은 관리자 화면의 계정 카드 이름 옆에도 보입니다. 카드 오른쪽 작은 상자는 그 사용자의 <b>시작 화면 · AI 바로 열기 · 오른손 모드</b>입니다(보기만, 시작 화면·AI는 계정 카드에서 바꿈). 오른손 모드는 사용자가 앱을 열거나 바꿀 때 기록되며, 아직 기록이 없으면 ‘–’로 보입니다. 관리자 줄은 지금 이 기기의 설정이라 이름 옆에 ‘(이 기기)’가 붙습니다.", "<b>사용자 추가</b>: 이름과 이메일을 넣고 [추가]를 누르면 사용자 정보 팝업이 수정 모드로 열리고, 나머지 정보를 채워 [저장]하면 확인창 뒤 등록됩니다(×로 닫으면 추가 취소).", "<b>삭제 = 보관</b>: 계정 카드의 [삭제]는 계정을 지우지 않고 보관합니다. 앱 사용과 Drive 공유는 막히지만 이름·권한·직책·조직 위치·사용자 정보는 남고, 계정 목록 아래 ‘보관된 사용자’에서 [원복]하면 그대로 돌아옵니다.", "<b>사용자 정보</b>(팝업): 사용자 추가에서 이메일을 비우고 이름만 넣어도 바로 등록되고 조직도에서 그 사용자가 펼쳐진 수정 화면으로 열립니다(이메일·사번은 나중에 채움). 폴더 공유는 이름·이메일·사번이 모두 채워졌을 때만 합니다. 조직도에서 이름을 눌러 펼친 뒤 [정보]를 누르거나(조직도는 처음에 이름만 보이고 이름을 누르면 펼쳐짐) 보기 모드로(칸이 잠겨 고칠 수 없음; 관리자 화면 계정 카드는 눌러도 열리지 않음), 계정 카드의 [수정](또는 팝업 위 [수정])을 누를 때만 수정 모드로(위에 [저장]·[취소], [취소]는 바꾼 내용을 버리고 보기 모드로) 그 사용자의 이름·사번·이메일·통신사·연락처·주소·입사일·직책·메모를 보고 고칠 수 있고, 아래 ‘신한FC 설정’에 조직 위치·권한·민감정보·고객/세일즈 페이지 접근·일정 구분·시작 화면·AI·오른손 모드가 모여 보입니다. 일정 구분·시작 화면·AI 바로 열기도 이 팝업에서 바꿉니다. 보기 모드의 [취소]나 바깥, 폰의 뒤로가기로 닫습니다. 개인 정보는 사용자에게 공유되지 않는 ‘관리자전용’ 폴더에 저장됩니다(이메일은 로그인 계정이라 바꿀 수 없음)."]);
    const s_add = sec("사용자 추가", ["이름과 구글 계정 이메일을 넣고 [추가]를 누르면 위 표에 줄이 생깁니다. 볼 FC를 체크하고 [저장]을 눌러야 적용됩니다."]);
    const s_set = sec("설정", ["<b>회사 주소</b>: 일정 이동시간 계산의 출발·도착 기준입니다.", "<b>카카오 REST API 키</b>: 이동 거리·시간 계산에 씁니다. 실수로 바뀌지 않게 잠겨 있고, [수정]을 눌러야 고칠 수 있습니다. 저장하면 다시 잠깁니다."]);
    const s_ib = sec("팀관리 > 도입현황 (관리자, v578)", ["팀관리의 <b>[도입현황]</b>에서 연도를 고르면 그해 도입 합계, 도입자별 인원, 월별 '도입자 → 피도입자'가 나옵니다. <b>[최근 3개월]</b>을 누르면 이번 달을 포함한 3개월만 봅니다.", "도입 1건 = 팀원정보에서 도입자가 있는 팀원 1명입니다. 날짜는 <b>도입일</b>(없으면 입사일)로 셉니다. [팀원 전환]을 하면 도입일이 입사일로 자동으로 들어가고, 팀원정보 도입자 밑 도입일 칸(달력)에서 고칠 수 있습니다. 퇴직해도 도입 기록은 남고 '퇴직'으로 표시합니다."]);
    const s_tm = sec("팀관리 > 팀원정보 (관리자, 2026-11-19)", ["팀관리의 <b>[팀원정보]</b>를 누르면 팀원 한 사람의 정보가 나옵니다. 위 <b>드롭다운</b>에서 팀원을 고르며, 순서는 <b>직급순 → 사번순</b>이고 처음에는 <b>로그인한 본인</b>이 골라져 있습니다.", "보이는 항목: <b>주민번호 · 연락처(앞에 통신사) · 이메일 · 주소([지도]) · 도입자 · 소개자 · 입사일 · 메모</b>, 구분선 아래에 <b>Sol FC(비번) · 신한이지(아이디 · 비번) · 지식인(비번) · 보험연수원(아이디 · 비번)</b>. 이름·사번·직급은 위 드롭다운에 나옵니다.", "주민번호와 비번은 가리지 않고 그대로 보여 주고(관리자만 보는 화면), 주민번호 옆에는 <b>[복사]</b> 버튼이 있습니다(구분선 아래 계정 항목에는 버튼 없이 글자만).", "<b>[수정]</b>을 누르면 보는 화면 그대로 그 자리에서 <b>주민번호 · 통신사 · 연락처 · 이메일 · 주소 · 입사일 · 메모와 구분선 아래 계정 항목</b>을 고칠 수 있고(<b>도입자</b>는 팀원 중에서 골라 사번으로 연결하고 화면에는 이름이 나옵니다. 우리 팀원이 아닌 사람(다른 팀에서 도입한 경우 포함)은 바로 아래 <b>소개자</b> 칸에 적습니다. 통신사는 SKT · KT · LG U+ · 각 알뜰폰 중에서 고름) <b>[저장]</b>하면 관리자 드라이브의 <b>Claude > 팀</b> 폴더 팀원 파일에 바로 저장됩니다. 팀원 정보는 관리자에게만 보이고 기기에 남기지 않습니다.", "<b>‹ 팀관리</b>로 돌아갑니다."]);
    const s_num = sec("주민번호·전화번호 '-' 자동 (v576)", ["주민번호와 전화번호는 <b>숫자만</b> 저장합니다. 화면에 보일 때는 '-'를 알아서 넣습니다(800823-1234567, 010-1234-5678, 02-123-4567). 수정 칸에서도 숫자만 치면 자리에 맞춰 '-'가 들어갑니다. 고객 정보·팀원정보·후보자정보·세일즈 목록·보고서에 모두 같습니다."]);
    const s_copy = sec("계좌 복사 (고객 정보 > 정보 탭, 2026-11-19)", ["고객 정보에 <b>계좌</b>가 있으면 오른쪽에 <b>[복사]</b> 버튼이 나옵니다. 누르면 계좌 내용이 그대로 클립보드에 복사되어 문자·카톡·은행 앱에 붙여 넣을 수 있습니다. 계좌는 민감정보 권한이 있는 계정에만 보입니다."]);
    const s_orgin = sec("조직도 = 도입자 기준 (관리자, 2026-11-19)", ["<b>조직도</b>를 열면 팀원정보의 <b>도입자</b>를 따라 자리를 맞춥니다. 도입자가 팀원(이 앱 사용자)이면 그 사람 밑에, <b>도입자가 팀원에 없으면 관리자 바로 밑</b>에 둡니다.", "자리가 바뀐 사람이 있으면 알림이 뜨고, <b>아래 [저장]을 눌러 확인창을 거쳐야</b> 실제로 적용됩니다(조직도에 따라 고객·세일즈 페이지 접근 권한이 달라지기 때문). 저장하지 않고 나가면 그대로입니다."]);
    const s_cd = sec("팀관리 > 후보자정보 (관리자, 2026-11-19)", ["<b>담아두기(v563)</b>: [수정] 오른쪽 [담아두기]를 누르면 그 후보자는 위 고르기 목록에서 빠집니다(바로 [되돌리기] 가능). 제목 줄의 [담아둔 N]을 누르면 담아둔 후보자 목록이 나오고, 이름·연락처로 찾고, 이름을 눌러 정보를 보고, [꺼내기]로 다시 목록에 넣습니다.", "<b>후보자 추가(v562, v563에서 [추가]는 제목 줄 오른쪽)</b>: [추가] → 먼저 메모와 이미지(이력서·자기소개서·명함 사진, 끌어다 놓거나 눌러서 고르기)를 넣고 [분석해서 채우기]를 누르면, 구글 Gemini가 읽어 후보자 정보 칸을 미리 채웁니다(색 칸). 확인하고 빈 칸을 채운 뒤 [완료]. [건너뛰고 직접 입력]도 됩니다.", "<b>수정의 자료로 채우기(v562)</b>: [수정]을 누르면 맨 위에 메모·이미지 칸이 나옵니다. [빈 칸만 채우기]는 있는 정보를 그대로 두고, [기존 정보도 바꾸기]는 읽은 값으로 덮어씁니다(바뀐 칸 아래에 이전 값). [저장]을 눌러야 저장됩니다.", "<b>자료 박스(v562, v566에서 올리기 더함)</b>: 상자 오른쪽 위 [올리기]를 누르거나 파일을 상자에 끌어다 놓으면 바로 올라갑니다(사진은 줄여서, PDF 등은 원본 그대로). 후보자정보 맨 아래에 올린 파일 목록이 나옵니다. 이름을 누르면 열리고, [삭제]하면 드라이브 휴지통으로 갑니다(30일 안에 되살림). 저장 위치: Claude/팀/후보자자료/후보자ID - 이름.", "<b>칸 고르기(v560)</b>: 수정할 때 고등학교는 학교명과 졸업 연도, 최종학력은 고등학교·대학교·대학원 중 고르기, 병역은 면제·육군·공군·해군 중 고르고 시작·끝 년월, 결혼여부는 체크와 결혼기념일(월·일)로 넣습니다. 예전에 글로 적은 값은 알아볼 수 있는 만큼 새 칸으로 옮겨 보여 주고, 저장하면 원래 글은 따로 남겨 둡니다.", "팀관리의 <b>[후보자정보]</b>(팀원정보 오른쪽)를 누르면 후보자 한 사람의 정보가 나옵니다. 위 한 줄은 <b>‹ 팀관리 · 후보자 드롭다운 · [수정]</b>입니다.", "보이는 항목: <b>이름 · 사번 · 입사일 · 도입자(팀원 중에서 고름) · 소개자(도입자에게 소개한 사람, 팀원이 아닌 사람) · 영문이름 · 주민번호 · 연락처(앞에 통신사) · 이메일 · 주소([지도]) · 고등학교(졸업 연도) · 최종학력 · 병역(입대 년월) · 결혼여부(기념일) · 자녀여부 · 종교 · 취미 · 특기</b>, 그 아래 <b>짧은 질문 6개</b>와 <b>상세 자기소개서 3개</b>(질문 아래에 답).", "<b>[수정]</b>을 누르면 보는 화면 그대로 그 자리에서 모든 항목을 고치고 <b>[저장]</b>하면 관리자 드라이브의 <b>Claude > 팀</b> 폴더 후보자 파일에 저장됩니다. 관리자에게만 보이고 기기에 남기지 않습니다. ", "<b>팀원 전환</b>: 정보가 다 들어가 있어도 자동으로 팀원이 되지는 않습니다. <b>필수 항목(이름 · 사번 · 입사일)</b>을 넣어 저장하면 맨 아래 <b>[팀원 전환]</b>이 눌리고(빠진 항목은 버튼 옆에 알려 줌), 그 버튼을 눌러야 등록이 시작됩니다. 누르면 확인 화면에서 <b>입사일(적어 둔 날짜) · 직급(기본 FC) · 도입자</b>를 확인하고 <b>[전환]</b>합니다. 사번이 숫자 8자리가 아니거나 팀원정보에 같은 사번이 있으면 전환하지 않습니다. 전환하면 후보자 기록(자기소개 포함)을 모두 팀원 파일로 옮겨 <b>팀원정보</b>에 나오고, 후보자 목록에는 <b>전환하지 않은 후보자만</b> 보입니다(후보자 기록은 지우지 않고 남겨 둠). 단계는 후보자 → 팀원 두 가지입니다(교육생 단계 없음). 앱 계정 등록·조직도·고객 폴더는 지금처럼 따로 합니다."]);
    const s_clm = sec("청구 탭 (고객 정보, v551 2026-11-22 바뀜)", ["맨 위 메뉴 4개: <b>개인청구서 업로드 · 기본청구서 선택 · 병원서류 업로드 · 청구기록</b>. 누르면 아래 칸이 그 내용으로 바뀝니다(처음은 개인청구서).", "<b>개인청구서</b>: 이 고객에게 올린 청구서가 보이고, 앞의 체크박스로 고릅니다. [업로드] 또는 끌어다 놓기로 올리면 원본은 보험금청구서 › 원본 폴더에 두고, <b>보험금청구서 스타일</b>(A4 세로, 회색조·밝기 맞춤, 바깥 여백을 잘라 위쪽 왼쪽부터, 300dpi JPEG 품질 70%)로 바꾼 PDF를 이 고객 청구서로 저장합니다. 신분증 사본(파일 이름이나 카드 모양으로 판단)은 자르기·돌리기·키우기·위치 옮기기를 하지 않고 원래 모양 그대로 둡니다(병원서류도 같음). ×는 삭제(드라이브 휴지통).", "<b>기본청구서</b>: 위쪽 담당 FC 칸(기본 = 로그인한 FC, 고칠 수 있음)이 청구서 1쪽 안내 박스에 들어갑니다(예전 공용 샘플 ‘김신한’ 자리). 보험사는 <b>자주 쓴 서식(많이 쓴 순, 횟수 표시) → 추천 7곳 → 나머지 이름순</b>으로, 묶음 사이에 가로줄이 있습니다.", "<b>병원서류</b>: [업로드] 또는 끌어다 놓기. 올린 파일마다 이름과 ×, 아래에 [추가 업로드].", "<b>청구기록</b>: 다운로드할 때마다 날짜 · 사고 내용 · 청구한 회사가 한 줄로 남습니다. 줄을 누르면 그때 만든 PDF, ×는 기록 삭제(PDF는 휴지통).", "맨 아래 <b>[보험금청구서 다운로드]</b>: 개인청구서나 기본청구서를 하나라도 고르면 켜집니다. 고른 청구서(개인 → 기본) 뒤에, 병원서류가 있으면 A4 회색조 300dpi(JPEG 품질 70%)로 바꿔 끝에 붙여 한 PDF로 내려받고, 고객 폴더 › 보험금청구 › 보험금청구 기록에도 저장합니다. 옆 칸에 사고 내용을 짧게 적으면 청구기록에 남습니다."]);
    const s_docs = sec("자료 탭 (고객 정보, 2026-11-19)", ["고객 정보의 탭 맨 오른쪽(<b>청구 오른쪽</b>)에 <b>자료</b> 탭이 있습니다. 그 고객 폴더 아래 <b>‘자료’ 폴더</b>에 든 파일을 <b>썸네일</b>(작은 그림)로 보여 줍니다. 사진은 누르면 화면에서 크게, PDF는 새 창, 그 밖의 파일은 구글 드라이브에서 열립니다.", "화면 아래 줄의 <b>[수정]은 정보 탭에서만</b> 나옵니다(계약 · 분석 · 상담 · 청구 · 자료 탭에서는 [닫기]만). 자료 탭에서는 그 줄 <b>왼쪽 끝에 [업로드]</b>가 나오고, [선택]으로 자료를 고르는 중에는 [업로드]가 눌리지 않습니다. <b>[업로드]</b>를 누르면 사진·파일을 여러 개 고를 수 있고, 올리면 바로 목록 맨 앞에 보입니다. 또 고객 폴더에 ‘자료’ 폴더가 없으면 그때 만듭니다. 파일 이름 앞에 오늘 날짜(YYYYMMDD)를 붙이고, 같은 이름이 있으면 덮어쓰지 않고 (2), (3)을 붙입니다. 한 파일 50MB까지 올릴 수 있습니다.", "<b>여러 개 삭제</b>: <b>[선택]</b>을 누르면 썸네일을 눌러 여러 개 고를 수 있고(고른 것은 파란 테두리와 체크), <b>[삭제]</b>로 한꺼번에 지웁니다. <b>[취소]</b>로 선택을 끝냅니다.", "썸네일 오른쪽 위 <b>×</b>를 누르면 확인 뒤 그 하나만 삭제합니다(구글 드라이브 <b>휴지통</b>으로 옮겨져 30일 안에는 되살릴 수 있습니다). 보기만 권한인 FC의 고객은 올리기·삭제 버튼이 나오지 않습니다."]);
    const parts = [grp("AI (아래 메뉴 맨 왼쪽)"), s_ai, s_wake, s_dsh, s_theme, s_team, s_tm, s_ib, s_cd, s_orgin, grp("검색"), s_srch, grp("고객 정보"), s_num, s_copy, s_clm, s_docs, s_cov, s_cmp, s_in, s_prog, grp("일정"), s_addev, s_nm, s_task, s_cal, grp("세일즈"), s_sub, ...(s_od ? [s_od] : []), s_pg, s_mat, s_smemo, s_stage, s_inflow, s_nf, s_sms, grp("설정"), s_start, s_rh, s_save, s_bk, s_db, s_ct, s_ntf, s_login, s_go, s_map, s_inst];
    if (isAdmin()) parts.push(grp("관리자 (설정 화면 아래쪽)"), s_perm, s_add, s_sup, s_set);
    return iconBox + parts.join("");
  }
  // 설명서 = AI 도우미처럼 아래 시트: 아래에서 올라오고, 끌어내리기·뒤로가기·닫기·바깥·Esc 모두 아래로 내려감(AI 시트와 같은 방식)
  var helpIsOpen = false, helpClosing = null, helpHist = false, helpExpectPop = false, helpMouse = false, helpDrag = null;
  function helpFinish() {
    if (!helpClosing) return;
    clearTimeout(helpClosing); helpClosing = null;
    const box = $("helpModal"); if (box) { box.hidden = true; box.classList.remove("closing"); }
    document.documentElement.classList.remove("help-open");
  }
  function openHelp() {
    helpFinish();
    if (helpIsOpen) return;
    let box = $("helpModal");
    if (!box) { box = document.createElement("div"); box.id = "helpModal"; box.className = "help-modal"; document.body.appendChild(box); box.addEventListener("click", (e) => { if (e.target === box || e.target.closest("[data-helpclose]")) helpClose(); }); }
    box.innerHTML = `<div class="help-sheet" role="dialog" aria-label="설명서"><div class="help-top"><div class="ai-grab" aria-hidden="true"></div><div class="help-head"><b>설명서</b><button type="button" class="secondary" data-helpclose>닫기</button></div></div><div class="help-body">${helpHtml()}</div></div>`;
    box.classList.remove("closing"); box.hidden = false; helpIsOpen = true;
    document.documentElement.classList.add("help-open");
    history.pushState({ v: "help" }, ""); helpHist = true; // 폰의 뒤로가기로 닫을 수 있게 기록 한 칸 쌓기
    helpBindDrag(box.querySelector(".help-sheet"));
  }
  function helpClose(fromPop) {
    if (!helpIsOpen) return;
    helpIsOpen = false;
    const box = $("helpModal"), sh = box.querySelector(".help-sheet");
    sh.style.transition = "transform 200ms var(--ease, ease-out)"; sh.style.transform = "translateY(100%)"; box.classList.add("closing");
    helpClosing = setTimeout(helpFinish, 220);
    if (helpHist) { helpHist = false; if (fromPop !== "pop") { helpExpectPop = true; history.back(); } }
  }
  window.addEventListener("popstate", () => { // AI 시트와 같은 방식: 우리가 되돌린 기록·뒤로가기는 설명서만 닫고 화면은 그대로
    if (helpExpectPop) { helpExpectPop = false; aiSkipOther = true; return; }
    if (helpIsOpen) { helpClose("pop"); aiSkipOther = true; }
  });
  document.addEventListener("keydown", (e) => { if (e.key === "Escape" && helpIsOpen) helpClose(); });
  // 아래로 끌어 닫기: 내용이 맨 위일 때만 끌림(스크롤 중이면 스크롤이 먼저). 80px 넘게 끌면 닫힘
  function helpBindDrag(sh) {
    const bd = sh.querySelector(".help-body");
    let y0 = null, dy = 0, drag = false;
    const begin = (y) => { y0 = y; dy = 0; drag = false; };
    const move = (y, ev) => {
      if (y0 === null) return;
      const d = y - y0;
      if (!drag) {
        if (d > 8 && bd.scrollTop <= 0) { drag = true; sh.style.transition = "none"; }
        else if (d < -8 || bd.scrollTop > 0) { y0 = null; return; }
        else return;
      }
      dy = Math.max(0, d); sh.style.transform = "translateY(" + dy + "px)";
      if (ev && ev.cancelable) ev.preventDefault();
    };
    const end = () => { if (y0 === null) return; const was = drag; y0 = null; drag = false; sh.style.transition = ""; if (!was) return; if (dy > 80) helpClose(); else sh.style.transform = ""; };
    sh.addEventListener("touchstart", (e) => begin(e.touches[0].clientY), { passive: true });
    sh.addEventListener("touchmove", (e) => move(e.touches[0].clientY, e), { passive: false });
    sh.addEventListener("touchend", end); sh.addEventListener("touchcancel", end);
    sh.addEventListener("pointerdown", (e) => { if (e.pointerType === "mouse") { helpMouse = true; begin(e.clientY); } });
    helpDrag = { move, end };
  }
  window.addEventListener("pointermove", (e) => { if (helpMouse && helpDrag) helpDrag.move(e.clientY, null); });
  window.addEventListener("pointerup", () => { if (helpMouse && helpDrag) { helpMouse = false; helpDrag.end(); } });
  $("fsHelp").addEventListener("click", openHelp);
  $("fsSave").addEventListener("click", async () => {
    const sb = $("fsSave"); sb.dataset.busy = "1"; sb.disabled = true; sb.textContent = "저장 중…";
    try {
      for (const b of fsBtns()) { if (b.isConnected && !b.disabled) { b.click(); await new Promise((r) => setTimeout(r, 60)); await fsWait(b); } }
    } finally { sb.dataset.busy = ""; sb.textContent = "저장"; fsRefresh(); }
  });
  // 새로고침 버튼(로그아웃 왼쪽): 이 화면을 다시 불러옴. 세일즈·일정 화면은 보던 자리로 돌아옴
  $("reloadBtn").addEventListener("click", async () => {
    if (typeof confirmLeaveEdit === "function" && !(await confirmLeaveEdit())) return; // 수정 중이면 저장·취소 확인
    location.reload();
  });
  $("logoutBtn").addEventListener("click", async () => {
    if (GD && !confirm("모든 앱에서 로그아웃됩니다. 로그아웃할까요?")) return;
    await photoCacheClear(); // 기기에 저장한 고객 사진 지우기
    if (!GD) return msal.logoutRedirect();
    clearGToken(); // 공통 출입증 삭제 = 모든 앱 로그아웃 (revoke는 안 함: 다른 앱의 동의가 사라지지 않게)
    try { localStorage.removeItem(GACC_KEY); } catch (e) { /* 무시 */ }
    salesClearState();
    location.reload();
  });
  // 상단 검색창: 어느 화면에서든 검색 → 고객 검색 결과 화면으로
  function runTopSearch(q, narrow) { // narrow=true: 고객 화면 안 검색창(고객검색), 아니면 위쪽 전체검색
    searchWide = !narrow;
    if (editMode && !$("customerView").hidden) { confirmLeaveEdit().then((ok) => { if (ok) runTopSearch(q, narrow); }); return; }
    $("q").value = q;
    $("q").blur();
    $("recentList").hidden = true;
    if (!q.trim()) return;
    addRecent(q.trim());
    if (!$("customerView").hidden) {
      // 고객 화면에서 검색: 뒤로 가기(검색 화면)가 끝난 뒤에 검색해야 결과가 1명일 때 바로 열린 고객 화면이 덮이지 않음
      window.addEventListener("popstate", () => { show("searchView"); search(q); }, { once: true });
      history.back();
      return;
    }
    if ($("searchView").hidden) show("searchView");
    search(q);
  }
  $("topSearch").addEventListener("submit", (ev) => { ev.preventDefault(); runTopSearch($("q").value); });

  // 최근 검색어: 검색창을 누르면 아래로 펼쳐짐 (이 휴대폰에만 저장)
  const RECENT_KEY = "shlife.recentSearches";
  const getRecent = () => { try { return JSON.parse(localStorage.getItem(RECENT_KEY)) || []; } catch (e) { return []; } };
  function addRecent(q) {
    const list = [q, ...getRecent().filter((x) => x !== q)].slice(0, 10);
    try { localStorage.setItem(RECENT_KEY, JSON.stringify(list)); } catch (e) { /* 저장 못 해도 검색은 계속 */ }
  }
  function showRecent() {
    const list = getRecent();
    const box = $("recentList");
    if (!list.length) { box.hidden = true; return; }
    box.innerHTML = `<li class="recent-head">최근 검색</li>` +
      list.map((q) => `<li><button type="button" data-q="${esc(q)}">🕘 ${esc(q)}</button></li>`).join("");
    box.hidden = false;
  }
  $("q").addEventListener("focus", showRecent);
  $("q").addEventListener("blur", () => setTimeout(() => { $("recentList").hidden = true; }, 200));
  // 버튼을 누르는 순간 입력창 blur가 먼저 일어나므로 pointerdown에서 처리
  $("recentList").addEventListener("pointerdown", (ev) => {
    const b = ev.target.closest("[data-q]");
    if (b) { ev.preventDefault(); runTopSearch(b.dataset.q); }
  });
  $("backBtn").addEventListener("click", async () => { if (await confirmLeaveEdit()) history.back(); });
  // 고객 사진을 누르면 크게 보기, 다시 누르면 닫기
  $("avatar").addEventListener("click", () => {
    if (!current) return;
    if (editMode) {
      if (!current.folders.length) return toast("이 고객의 폴더가 없어 사진을 저장할 수 없습니다.");
      if (current.profilePhoto) { $("photoMenu").hidden = !$("photoMenu").hidden; return; }
      return $("photoInput").click();
    }
    if (!current.photoUrl) return;
    // 크게 보기: 원본이 있으면 원본, 없으면 작은 사진을 먼저 보여 주고 원본을 받아 바꿈
    $("photoViewerImg").src = current.photoFullUrl || current.photoUrl;
    $("photoViewer").hidden = false;
    const c = current;
    if (!c.photoFullUrl && c.profilePhoto) {
      stGet(c.profilePhoto.id).then((buf) => {
        c.photoFullUrl = URL.createObjectURL(new Blob([buf], { type: "image/jpeg" }));
        if (c === current && !$("photoViewer").hidden) $("photoViewerImg").src = c.photoFullUrl;
      }).catch((e) => console.error(e));
    }
  });
  $("editBtn").addEventListener("click", () => { if (current) setEditMode(!editMode); });
  $("closeCust").addEventListener("click", async () => { if (await confirmLeaveEdit()) history.back(); });
  $("editCancel").addEventListener("click", () => { if (current) setEditMode(false); });
  $("editSave").addEventListener("click", () => { if (!current) return; if ($("infoSave")) saveInfo(); else setEditMode(false); });
  $("photoPick").addEventListener("click", () => { $("photoMenu").hidden = true; $("photoInput").click(); });
  $("photoMenuClose").addEventListener("click", () => { $("photoMenu").hidden = true; });
  $("photoDel").addEventListener("click", async () => {
    $("photoMenu").hidden = true;
    if (!confirm("이 고객의 프로필 사진을 삭제할까요? (" + DRIVE_NAME + " 휴지통으로 이동)")) return;
    try { await deleteProfilePhotos(); toast("사진을 삭제했습니다."); } catch (e) { saveError(e); }
  });
  $("photoInput").addEventListener("change", async () => {
    const file = $("photoInput").files && $("photoInput").files[0];
    $("photoInput").value = "";
    if (!file) return;
    try {
      const cropped = await cropToCircle(file); // 원 안에 맞추기 (취소하면 null)
      if (!cropped) return;
      toast("사진 저장 중…");
      await uploadProfilePhoto(cropped);
      toast("사진을 저장했습니다.");
    } catch (e) { saveError(e); }
  });
  $("photoViewer").addEventListener("click", () => { $("photoViewer").hidden = true; });
  // PDF 열기 버튼 (계약 열기, 보장분석 등): 휴대폰 PDF 앱/브라우저로
  $("tabBody").addEventListener("click", (ev) => {
    const b = ev.target.closest("[data-open]");
    if (b) { ev.preventDefault(); openFileDirect(b.dataset.open); }
  });
  // (검색 범위 버튼은 없앰: 항상 전체 검색, 결과는 계약 > 이관 > 따행 > DB > 추가 순 표시)
  // 고객 화면 안 검색창 (결과가 없을 때 보임) → 상단 검색과 같은 동작
  $("pageSearch").addEventListener("submit", (ev) => {
    ev.preventDefault();
    const q = $("q2").value.trim();
    if (!q) return;
    $("q2").blur();
    runTopSearch(q, true); // 고객 화면 안 검색 = 고객검색(고객만)
  });
  // ---------- 구글 캘린더 (일정) ----------
  const GCFG = CFG.google || {};
  const GCAL = "https://www.googleapis.com/calendar/v3";
  const GTOKEN_KEY = "shlife.gtoken";
  // v2 = 캘린더(읽기·수정) / v3 = Google Drive 모드: Drive + 캘린더 + 이메일
  const GTOKEN_V = GD ? 3 : 2;
  const GSCOPE = "https://www.googleapis.com/auth/calendar.readonly https://www.googleapis.com/auth/calendar.events" +
    " https://www.googleapis.com/auth/tasks" + // v221: 구글 할 일 권한도 로그인 때 한 번에 받음
    (GD ? " https://www.googleapis.com/auth/drive openid email" : "");
  // ---- 공통 로그인: 모든 앱이 localStorage `blue10kp.gauth` {v,token,exp,scopes,email}를 같이 씀 (app\_공통\공통로그인.md) ----
  const GAUTH_KEY = "blue10kp.gauth";
  const GS = "https://www.googleapis.com/auth/";
  // 이 앱이 꼭 필요한 권한(할 일 tasks는 없어도 앱이 돌아가서 필수가 아님). 상위 권한이 있으면 포함으로 봄
  const GNEED = [["calendar", "calendar.readonly"], ["calendar", "calendar.events"], ["drive"], ["userinfo.email", "email"]];
  function gHasScopes(list) {
    const s = new Set((list || []).map((x) => String(x).replace(GS, "")));
    const has = (k) => s.has(k) ||
      (k === "calendar.readonly" && s.has("calendar")) || (k === "calendar.events" && s.has("calendar")) ||
      (k === "tasks" && s.has("tasks")) || (k === "drive" && s.has("drive"));
    return GNEED.every((alts) => alts.some(has));
  }
  function readGAuth() {
    try {
      const a = JSON.parse(localStorage.getItem(GAUTH_KEY) || "null");
      return a && typeof a.token === "string" && a.token && a.exp > 0 ? a : null;
    } catch (e) { return null; }
  }
  // v557: login = 구글 창으로 출입증을 받은 시각(ms). 오전 6시 일괄 로그아웃 판단용(app\_공통\20261124 - 오전6시 일괄로그아웃 안.md). 모르는 칸은 지우지 않고 그대로 둠
  function writeGAuth(a) { try { const { v, token, exp, scopes, email, login, ...rest } = a; localStorage.setItem(GAUTH_KEY, JSON.stringify({ ...rest, v: 1, token, exp, scopes: scopes || [], email: email || "", login: login > 0 ? login : exp - 3600000 })); } catch (e) { /* 저장 못 해도 이번엔 사용 */ } }
  // ---- v557 오전 6시 일괄 로그아웃: 출입증을 받은 시각이 오늘 06:00(지금이 6시 전이면 어제 06:00)보다 앞이면 출입증을 지우고 로그인 화면으로 ----
  const AM6_KEY = "blue10kp.gauth.am6"; // 6시 로그아웃이 된 날(YYYY-MM-DD) → 다음 로그인은 계정 고르기 창
  const gLoginAt = (a) => (a && a.login > 0 ? a.login : a ? a.exp - 3600000 : 0); // login 칸이 없는 옛 출입증은 만료 1시간 전을 받은 시각으로 봄
  function am6Cut() { const d = new Date(), c = new Date(d.getFullYear(), d.getMonth(), d.getDate(), 6, 0, 0); if (d < c) c.setDate(c.getDate() - 1); return +c; }
  const am6Today = () => { const d = new Date(); return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0"); };
  function am6Pending() { try { return localStorage.getItem(AM6_KEY) === am6Today(); } catch (e) { return false; } }
  let am6Busy = false;
  async function am6Check() {
    if (!GD || am6Busy) return false;
    const a = readGAuth();
    if (!a || gLoginAt(a) >= am6Cut()) return false;
    am6Busy = true;
    try {
      const shown = appStarted && $("loginView").hidden;
      if (shown && typeof editMode !== "undefined" && editMode) { try { await confirmLeaveEdit(); } catch (e) { /* 저장 확인 못 해도 로그아웃 */ } }
      try { localStorage.removeItem(GAUTH_KEY); localStorage.setItem(AM6_KEY, am6Today()); } catch (e) { /* 무시 */ }
      gTokenMem = null;
      if (shown || !$("loginView").hidden) {
        appStarted = false; scanCache = null;
        showGLogin();
        const msg = $("lmsg"); if (msg && !EMBED) msg.textContent = "오전 6시가 지나 로그아웃되었습니다. 다시 로그인해 주세요.";
      }
      return true;
    } finally { am6Busy = false; }
  }
  // 예전 신한FC 전용 출입증(shlife.gtoken)을 한 번 공통 키로 옮기고 지움
  function migrateLegacyGToken() {
    try {
      const raw = localStorage.getItem(GTOKEN_KEY);
      if (raw == null) return;
      localStorage.removeItem(GTOKEN_KEY);
      if (readGAuth()) return; // 이미 공통 키가 있으면 그것을 씀
      const t = JSON.parse(raw);
      if (!t || !t.token || !(t.exp > Date.now() + 60000) || t.v !== GTOKEN_V || !gAccount()) return;
      const sc = [GS + "calendar.readonly", GS + "calendar.events", GS + "drive", "openid", "email"];
      if (t.tk) sc.push(GS + "tasks");
      writeGAuth({ token: t.token, exp: t.exp, scopes: sc, email: gAccount(), login: t.exp - 3600000 });
    } catch (e) { /* 무시 */ }
  }
  // 지금 쓸 수 있는 공통 출입증 (만료 1분 전까지, 필요한 권한이 다 있어야 함)
  function storedGToken() {
    const a = readGAuth();
    return a && a.exp > Date.now() + 60000 && gHasScopes(a.scopes) ? a.token : null;
  }
  // used = 방금 401을 받은 토큰. 공통 키의 토큰이 그것과 같을 때만 지움(다른 앱이 새로 받았으면 그대로 둠)
  function clearGToken(used) {
    const a = readGAuth();
    if (a && (!used || a.token === used)) { try { localStorage.removeItem(GAUTH_KEY); } catch (e) { /* 무시 */ } }
    gTokenMem = null;
  }
  // 구글 로그인 창(팝업) → 캘린더 읽기 권한 토큰 (1시간 유효, 버튼을 눌렀을 때만 호출)
  function requestGToken(prompt, noHint) {
    return new Promise((resolve, reject) => {
      const g = window.google && window.google.accounts && window.google.accounts.oauth2;
      if (!g) return reject(new Error("구글 로그인 모듈을 불러오지 못했습니다. 잠시 후 다시 시도하세요."));
      const client = g.initTokenClient({
        client_id: GCFG.clientId,
        scope: GSCOPE,
        include_granted_scopes: true,
        hint: noHint ? undefined : ((readGAuth() || {}).email || (GD ? gAccount() : appSettings.google && appSettings.google.account)) || undefined,
        callback: (r) => {
          if (r.error) return reject(new Error("구글 연결 실패: " + r.error));
          if (GD && !google.accounts.oauth2.hasGrantedAllScopes(r, "https://www.googleapis.com/auth/drive")) return reject(new Error("Google Drive 권한을 허용해야 합니다. 다시 로그인하면서 모든 항목에 체크해 주세요."));
          // 공통 출입증에 기록(받은 권한 그대로). 이메일은 userinfo로 확인한 뒤 googleLogin이 채움
          const prev = readGAuth();
          writeGAuth({ token: r.access_token, exp: Date.now() + (r.expires_in || 3600) * 1000, scopes: String(r.scope || "").split(/\s+/).filter(Boolean), email: !noHint && prev ? prev.email : "", login: Date.now() });
          gTokenMem = r.access_token;
          resolve(r.access_token);
        },
        error_callback: (e) => reject(new Error(e && e.type === "popup_closed" ? "구글 로그인 창이 닫혔습니다." : "구글 로그인 창을 열 수 없습니다. 팝업 차단을 확인하세요."))
      });
      client.requestAccessToken({ prompt: prompt == null ? "" : prompt });
    });
  }
  let gTokenMem = null;
  async function gcal(path) {
    const t = storedGToken() || gTokenMem;
    if (!t) throw new Error("GOOGLE_AUTH");
    const res = await fetch(GCAL + path, { headers: { Authorization: "Bearer " + t } });
    if (res.status === 401) { clearGToken(t); throw new Error("GOOGLE_AUTH"); }
    if (!res.ok) throw new Error("Calendar " + res.status + ": " + (await res.text()).slice(0, 160));
    return res.json();
  }
  // 일정 수정 (장소 입력, 일정 수정 모드)
  // 다른 캘린더로 옮기기 (원래 캘린더에서 이 일정을 만든 경우에만 가능)
  async function gcalMove(calId, eventId, destId) {
    const t = storedGToken() || gTokenMem;
    if (!t) throw new Error("GOOGLE_AUTH");
    const res = await fetch(GCAL + "/calendars/" + encodeURIComponent(calId) + "/events/" + encodeURIComponent(eventId) + "/move?destination=" + encodeURIComponent(destId), {
      method: "POST", headers: { Authorization: "Bearer " + t }
    });
    if (res.status === 401) { clearGToken(t); throw new Error("GOOGLE_AUTH"); }
    if (res.status === 403) throw new Error("이 일정은 다른 캘린더로 옮길 권한이 없습니다.");
    if (!res.ok) throw new Error("캘린더 변경 실패 (" + res.status + "): " + (await res.text()).slice(0, 120));
    return res.json();
  }
  let schedCals = []; // 지금 일정 화면에 보이는 캘린더 (일정 수정의 캘린더 선택용)
  async function gcalPatch(calId, eventId, body) {
    const t = storedGToken() || gTokenMem;
    if (!t) throw new Error("GOOGLE_AUTH");
    const res = await fetch(GCAL + "/calendars/" + encodeURIComponent(calId) + "/events/" + encodeURIComponent(eventId), {
      method: "PATCH", headers: { Authorization: "Bearer " + t, "Content-Type": "application/json" }, body: JSON.stringify(body)
    });
    if (res.status === 401) { clearGToken(t); throw new Error("GOOGLE_AUTH"); }
    if (res.status === 403) throw new Error("이 캘린더의 일정을 수정할 권한이 없습니다.");
    if (!res.ok) throw new Error("Calendar " + res.status + ": " + (await res.text()).slice(0, 160));
    return res.json();
  }

  // 일정 새로 만들기 (AI 도우미용)
  async function gcalInsert(calId, body) {
    const t = storedGToken() || gTokenMem;
    if (!t) throw new Error("GOOGLE_AUTH");
    const res = await fetch(GCAL + "/calendars/" + encodeURIComponent(calId) + "/events", {
      method: "POST", headers: { Authorization: "Bearer " + t, "Content-Type": "application/json" }, body: JSON.stringify(body)
    });
    if (res.status === 401) { clearGToken(t); throw new Error("GOOGLE_AUTH"); }
    if (res.status === 403) throw new Error("이 캘린더에 일정을 만들 권한이 없습니다.");
    if (!res.ok) throw new Error("Calendar " + res.status + ": " + (await res.text()).slice(0, 160));
    return res.json();
  }
  // 일정 삭제 (구글 캘린더에서 지움 — 되돌릴 수 없어 누르기 전에 한 번 확인)
  async function gcalDelete(calId, eventId) {
    const t = storedGToken() || gTokenMem;
    if (!t) throw new Error("GOOGLE_AUTH");
    const res = await fetch(GCAL + "/calendars/" + encodeURIComponent(calId) + "/events/" + encodeURIComponent(eventId), { method: "DELETE", headers: { Authorization: "Bearer " + t } });
    if (res.status === 401) { clearGToken(t); throw new Error("GOOGLE_AUTH"); }
    if (res.status === 403) throw new Error("이 캘린더의 일정을 삭제할 권한이 없습니다.");
    if (res.status === 404 || res.status === 410) return; // 이미 없음
    if (!res.ok) throw new Error("Calendar " + res.status + ": " + (await res.text()).slice(0, 160));
  }
  const calAlias = (id, name) => ((GCFG.calendarNames || {})[id] || name);
  const calName = (c) => calAlias(c.id, String(c.summaryOverride || c.summary || "").trim());
  async function calendarList() {
    return ((await gcal("/users/me/calendarList?minAccessRole=reader&maxResults=250")).items || [])
      .map((c) => ({ id: c.id, name: calName(c), color: c.backgroundColor, primary: !!c.primary, role: c.accessRole }));
  }
  // 일정에 보일 캘린더
  //   관리자: 관리자 설정에서 고른 캘린더 (모든 사용자의 기본값)
  //   일반 사용자: 설정 메뉴에서 고른 '내 캘린더'. 고른 적 없으면 기본 이름(관리자 설정 이름 + config: 1팀, 김현성, 최수진)과 같은 내 캘린더
  //   (일반 사용자의 기본값은 관리자 설정 파일이 아니라 config의 공용 기본 이름만 사용 — 관리자 개인 캘린더 선택이 다른 사용자에게 번지지 않게)
  const defaultCalNames = () => (isAdmin() ? [...new Set([...((appSettings.google || {}).calendars || []).map((c) => c.name), ...(GCFG.calendars || [])])] : [...(GCFG.calendars || [])]);
  async function scheduleCalendars() {
    if (!isAdmin()) {
      const p = await loadMyPrefs();
      if (p && Array.isArray(p.calendars) && p.calendars.length) return p.calendars;
      const want = defaultCalNames();
      return (await calendarList()).filter((c) => want.includes(c.name));
    }
    const saved = ((appSettings.google || {}).calendars || []).filter((c) => c && c.id);
    if (saved.length) return saved;
    const want = GCFG.calendars || [];
    return (await calendarList()).filter((c) => want.includes(c.name));
  }

  // ---------- 내 설정 (일반 사용자): 내 드라이브 Claude/app/신한FC/설정/shlife-내설정.json (+ 이 기기 사본) ----------
  const MY_FILE = "shlife-내설정.json";
  const MY_OWNER_FOLDER = "개인설정"; // 소유자 전용(공유 안 함). 다른 사용자는 각자 자기 Drive의 Claude/app/신한FC/설정 (예전과 같음)
  let myPrefs = null, myDir = null;
  const myCacheKey = () => "shlife.myprefs." + me();
  async function resolveMyDir(create) {
    if (myDir) return myDir;
    await resolveBase();
    if (!GD) { await resolveAppDir(); return (myDir = appDir); }
    if (base.owner) {
      // 소유자(관리자)의 개인 설정은 김경미·최수진에게 읽기 공유된 '설정' 폴더가 아니라, 같은 상위 폴더 아래 공유하지 않는 '개인설정' 폴더에 둔다 (2026-11-11)
      await resolveAppDir();
      const par = ((await gdFetch(GDAPI + "/files/" + encodeURIComponent(appDir.id) + "?fields=parents&supportsAllDrives=true")).parents || [])[0];
      if (!par) throw new Error("설정 폴더의 상위 폴더를 찾지 못했습니다.");
      const ex = (await gdQuery("'" + qv(par) + "' in parents and name='" + qv(MY_OWNER_FOLDER) + "' and mimeType='" + GD_FOLDER + "' and trashed=false"))[0];
      if (ex) return (myDir = { id: ex.id });
      if (!create) return null;
      const mk = await gdFetch(GDAPI + "/files?fields=id&supportsAllDrives=true", { method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ name: MY_OWNER_FOLDER, mimeType: GD_FOLDER, parents: [par] }) });
      return (myDir = { id: mk.id });
    }
    let id = "root";
    for (const part of CFG.appFolder.split("/").filter(Boolean)) {
      const r = await gdQuery("'" + qv(id) + "' in parents and name='" + qv(part) + "' and mimeType='" + GD_FOLDER + "' and trashed=false");
      if (r.length) { id = r[0].id; continue; }
      if (!create) return null;
      id = (await gdFetch(GDAPI + "/files?fields=id", { method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ name: part, mimeType: GD_FOLDER, parents: [id] }) })).id;
    }
    return (myDir = { id });
  }
  async function loadMyPrefs() {
    if (myPrefs) return myPrefs;
    try { myPrefs = JSON.parse(localStorage.getItem(myCacheKey()) || "null"); } catch (e) { myPrefs = null; }
    try {
      const d = await resolveMyDir(false);
      const p = d ? await stReadJsonIn(d, MY_FILE) : null;
      if (p) { myPrefs = p; try { localStorage.setItem(myCacheKey(), JSON.stringify(p)); } catch (e) { /* 무시 */ } }
      if (!embedTheme && p && (p.theme === "dark" || p.theme === "light") && themePref() !== p.theme) { try { localStorage.setItem(THEME_KEY, p.theme); } catch (e) { /* 무시 */ } themeApply(); const c = document.getElementById("darkMode"); if (c) c.checked = p.theme === "dark"; } // Drive에 저장된 다크/라이트가 있으면 그것을 따름
    } catch (e) { console.error(e); }
    return myPrefs || {};
  }
  async function saveMyPrefs(p) {
    myPrefs = p;
    try { localStorage.setItem(myCacheKey(), JSON.stringify(p)); } catch (e) { /* 무시 */ }
    const d = await resolveMyDir(true);
    if (d && d.id) await stWriteJsonIn(d, MY_FILE, p);
  }
  const withAlias = (c) => ({ ...c, name: calAlias(c.id, c.name) });
  async function fetchEvents(from, to) {
    const cals = await scheduleCalendars();
    if (!cals.length) throw new Error("NO_CALENDAR");
    const q = `?singleEvents=true&orderBy=startTime&maxResults=250&timeMin=${encodeURIComponent(from.toISOString())}&timeMax=${encodeURIComponent(to.toISOString())}`;
    schedCals = cals.map(withAlias);
    const sp = salesArr(isAdmin() ? (appSettings.google || {}).sales : (myPrefs || {}).sales).map((c) => c.id).filter((id) => schedCals.some((c) => c.id === id));
    schedSalesIds = new Set(sp.length ? sp : defaultSalesIds(schedCals));
    const ipRaw = isAdmin() ? (appSettings.google || {}).info : (myPrefs || {}).info;
    const ip = salesArr(ipRaw).map((c) => c.id).filter((id) => schedCals.some((c) => c.id === id));
    schedInfoIds = new Set(Array.isArray(ipRaw) ? ip : [...schedSalesIds]); // 정보를 고른 적 없으면 = 일정 캘린더
    const lists = await Promise.all(schedCals.map(async (c) => {
      // 볼 권한이 없는 캘린더(다른 사람 캘린더 등)는 건너뜀
      let d;
      try { d = await gcal("/calendars/" + encodeURIComponent(c.id) + "/events" + q); } catch (e) {
        if (e.message === "GOOGLE_AUTH") throw e;
        console.warn("calendar skipped", c.name, e.message);
        return [];
      }
      return (d.items || []).filter((e) => e.status !== "cancelled").map((e) => ({
        id: c.id + "|" + e.id, cal: c, calId: c.id, eventId: e.id,
        title: String(e.summary || "(제목 없음)").trim(),
        desc: String(e.description || ""),
        location: String(e.location || "").trim(),
        priv: ((e.extendedProperties || {}).private) || {}, updated: e.updated || "", raw: { id: e.id, start: e.start, end: e.end, summary: e.summary, updated: e.updated, creator: e.creator }, creator: String((e.creator || {}).email || "").toLowerCase(), // v559: 캘린더 고객 연결(calcust.js)
        pick: (() => { const l = CalCust.linkOf(((e.extendedProperties || {}).private) || {}); return l.manual ? l.value : ""; })(), // 사람이 고른 연결(선택·무시)만. 자동 연결은 화면에서 다시 찾음
        author: String(((e.extendedProperties || {}).shared || {}).shlifeAuthorName || ""), // v382: 관리자가 일정 추가 때 고른 작성자 이름
        sales: String(((e.extendedProperties || {}).private || {}).shlifeSales || ""), // 영업일정 표시: "1"=영업(확정), "0"=비영업(복귀), 없음=기본(아래 salesOf 규칙)
        created: e.created ? new Date(e.created) : null,
        allDay: !!e.start.date,
        start: e.start.dateTime ? new Date(e.start.dateTime) : ymdToDate(e.start.date),
        end: e.end.dateTime ? new Date(e.end.dateTime) : ymdToDate(e.end.date)
      }));
    }));
    return lists.flat().sort((a, b) => (b.allDay - a.allDay) || (a.start - b.start) || a.cal.name.localeCompare(b.cal.name, "ko"));
  }

  // ---------- 날짜 도우미 ----------
  const ymdToDate = (s) => { const [y, m, d] = s.split("-").map(Number); return new Date(y, m - 1, d); };
  const dayStart = (d) => new Date(d.getFullYear(), d.getMonth(), d.getDate());
  const addDays = (d, n) => new Date(d.getFullYear(), d.getMonth(), d.getDate() + n);
  const sameDay = (a, b) => a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
  const hm = (d) => String(d.getHours()).padStart(2, "0") + ":" + String(d.getMinutes()).padStart(2, "0");
  const WD = ["일", "월", "화", "수", "목", "금", "토"];
  const dayLabel = (d) => `${d.getMonth() + 1}월 ${d.getDate()}일 (${WD[d.getDay()]})`;
  const weekStart = (d) => addDays(dayStart(d), -((d.getDay() + 6) % 7)); // 월요일 시작
  const timeText = (e) => (e.allDay ? "종일" : `${hm(e.start)} ~ ${hm(e.end)}`);
  const onDay = (e, d) => e.start < addDays(dayStart(d), 1) && e.end > dayStart(d) || (e.allDay && sameDay(e.start, d));

  // ---------- 일정 제목 → 고객 찾기 ----------
  // 지역이 없으면 광주(광주광역시/전남광주통합특별시의 구)로 판단
  const COUNTIES = ["담양", "곡성", "구례", "고흥", "보성", "화순", "장흥", "강진", "해남", "영암", "무안", "함평", "영광", "장성", "완도", "진도", "신안"];
  const CITIES = ["목포", "여수", "순천", "나주", "광양", ...COUNTIES, "서울", "부산", "대구", "인천", "대전", "울산", "세종", "전주", "익산", "군산", "정읍", "남원", "김제", "제주"];
  // 제목의 지역: 괄호 안 지역 "(광양)" → 그 외 알려진 시·군 이름 → 없으면 광주
  const titleRegion = (t) => { const p = (t.match(/\(([가-힣]{2,4})\)/) || [])[1]; return (p && !/^\d|차$/.test(p) && p !== "광주" ? p.replace(/(시|군)$/, "") : "") || CITIES.find((c) => t.includes(c)) || ""; };
  const custRegionCity = (region) => { const m = String(region || "").match(/(\S+)$/); return region === "광주" || !region ? "" : (m ? m[1] : ""); };
  const entryAddr = (e) => { const a = e.fields.find(([k]) => k.trim() === "주소"); return a ? a[1] : ""; };

  // 일정 제목에 '이름 + 고객'(예: "김현성 고객 미팅", 붙여 쓴 "김현성고객" 포함)이 있을 때만 그 이름의 기존 고객을 찾는다(v403, 2026-11-17 확정안 2절). 목록에 없으면 아무것도 붙이지 않고, 이름만 적힌 일정에는 고객 카드가 붙지 않는다. 같은 이름이 여러 명이면 모두(최근 등록 순) → 일정 카드의 [선택]
  const NAME_RULE_FROM = new Date(2026, 10, 18); // v406(2026-11-17 사용자 "지난 일정 유지"): 이 날(2026-11-18)부터 시작하는 일정만 '이름 + 고객' 규칙. 그 전 일정은 예전처럼 이름만 있어도 찾음
  function matchCustomers(title, records, contracts, start) {
    const nt = norm(title), strict = false; // v410: 11-18부터의 '이름+고객' 규칙을 풀고 예전처럼 이름만으로 찾음 (NAME_RULE_FROM은 기록용)
    const fcs = new Set(records.map((r) => norm(r.fc)).filter(Boolean));
    let names = [...new Set(records.map((r) => baseName(r.name)))].filter((n) => norm(n).length >= 2 && nt.includes(norm(n) + (strict ? "고객" : "")));
    let part = false;
    if (!names.length) { // v579(2026-11-25 사용자 지시 "*단어* 형태로"): 제목의 낱말이 이름 안에 들어 있으면 추천(예: "Keiko" → Okamoto Keiko)
      const STOP = new Set(["미팅", "상담", "고객", "방문", "계약", "청약", "점심", "저녁", "회의", "통화", "전화", "증권", "전달", "서명", "보장", "분석", "설명", "사무실", "회사", "교육", "약속", "예정", "확인", "가족", "부부", "님", "씨"]);
      const toks = [...new Set((String(title).match(/[A-Za-z]{3,}|[가-힣]{2,}/g) || []).map((w) => w.toLowerCase()).filter((w) => !STOP.has(w) && ![...fcs].some((f) => f.includes(w))))]; // FC 이름(앞 글자 포함)은 빼고
      if (toks.length) {
        const cand = [...new Set(records.map((r) => baseName(r.name)))].filter((n) => { const nn = norm(n); return nn.length >= 2 && !fcs.has(nn) && toks.some((w) => nn.includes(w)); });
        if (cand.length && cand.length <= 8) { names = cand; part = true; }
      }
    }
    if (!names.length) return [];
    // FC 이름이 아닌 것, 긴 이름 우선
    names.sort((a, b) => (fcs.has(norm(a)) - fcs.has(norm(b))) || norm(b).length - norm(a).length);
    const name = names[0];
    const city = titleRegion(title);
    const fcHead = (title.match(/^\s*([^\s)]{1,3})\)/) || [])[1] || ""; // "최) …" → 최수진 FC 고객 우선
    const pickSet = new Set((part ? names : [name]).map(norm));
    const recs = records.filter((r) => pickSet.has(norm(baseName(r.name)))); // "이영숙(640927)"도 포함 (v579: 낱말로 찾은 경우는 찾은 이름 모두)
    const score = (r) => {
      const fcMiss = fcHead && r.fc && !String(r.fc).startsWith(fcHead) ? 4 : 0;
      const a = entryAddr(r);
      if (!a) return 1 + fcMiss;
      return (custRegionCity(regionFromAddress(a)) === city ? 0 : 2) + fcMiss;
    };
    // 같은 사람 묶기: 고객ID가 같으면 한 사람, ID가 없는 옛 행은 이름+생년월일
    const list = groupCustomers(recs);
    list.forEach((g) => {
      g.entries.sort((a, b) => prio(a.type) - prio(b.type) || score(a) - score(b));
      g.name = g.entries[0].name;
      g.city = city;
      g.recent = recencyOf(g, contracts);
    });
    // 표시 순서: 시트 우선순위(계약 > 이관 > 따행 > DB > 추가), 같으면 최근 등록 순 → 지역
    const top = (g) => Math.min(...g.entries.map((e) => prio(e.type)));
    const sc = (g) => Math.min(...g.entries.map(score));
    return list.sort((a, b) => top(a) - top(b) || (b.recent || "").localeCompare(a.recent || "") || sc(a) - sc(b));
  }
  // 날짜 글자 → 비교용 YYYYMMDD ("2026.05", "2026-08-31", "2005.01.01")
  const dateKey = (v) => {
    const d = String(v || "").replace(/\D/g, "");
    if (/^(19|20)\d{6}/.test(d)) return d.slice(0, 8);
    if (/^(19|20)\d{4}$/.test(d)) return d + "00";
    return "";
  };
  const RECENT_COLS = ["최초계약일", "계약일", "이관일자", "이관일", "이관월", "최근 보장분석"];
  function recencyOf(g, contracts) {
    let best = "";
    const take = (v) => { const k = dateKey(v); if (k > best) best = k; };
    g.entries.forEach((e) => e.fields.forEach(([k, v]) => { if (RECENT_COLS.includes(k.trim())) take(v); }));
    const fcs = g.entries.map((e) => e.fc).filter(Boolean);
    (contracts || []).forEach((c) => {
      if (!c.names.includes(norm(g.name))) return;
      if (fcs.length && !fcs.includes(c.fc) && !fcs.includes(c.srcFc)) return;
      RECENT_COLS.forEach((k) => take(c.row[k]));
    });
    return best;
  }

  // ---------- 카카오: 주소 → 좌표, 길찾기 ----------
  // 카카오 REST API 키는 공개 저장소(config.js)가 아니라 OneDrive 설정 파일(관리자 › 설정)에 보관
  const KAKAO = { get restKey() { return appSettings.kakaoRestKey || ""; } };
  const kakaoH = () => ({ Authorization: "KakaoAK " + KAKAO.restKey });
  const GEO_KEY = "shlife.geo";
  let geoCache = {};
  try { geoCache = JSON.parse(localStorage.getItem(GEO_KEY) || "{}"); } catch (e) { geoCache = {}; }
  async function geocode(q) {
    q = String(q || "").trim();
    if (!q || !KAKAO.restKey) return null;
    if (geoCache[q] !== undefined) return geoCache[q];
    let hit = null;
    try {
      const a = await (await fetch("https://dapi.kakao.com/v2/local/search/address.json?query=" + encodeURIComponent(q), { headers: kakaoH() })).json();
      const d = (a.documents || [])[0];
      if (d) hit = { x: +d.x, y: +d.y };
      if (!hit) {
        const k = await (await fetch("https://dapi.kakao.com/v2/local/search/keyword.json?query=" + encodeURIComponent(q), { headers: kakaoH() })).json();
        const p = (k.documents || [])[0];
        if (p) hit = { x: +p.x, y: +p.y };
      }
    } catch (e) { return null; }
    geoCache[q] = hit;
    try { localStorage.setItem(GEO_KEY, JSON.stringify(geoCache)); } catch (e) {}
    return hit;
  }
  const cityHall = (city) => (!city ? "광주광역시청" : COUNTIES.includes(city) ? city + "군청" : ["서울", "부산", "대구", "인천", "대전", "울산", "세종"].includes(city) ? city + "시청" : city + "시청");
  async function route(a, b) {
    if (!a || !b) return null;
    if (Math.abs(a.x - b.x) < 1e-5 && Math.abs(a.y - b.y) < 1e-5) return { km: 0, min: 0 };
    try {
      const r = await (await fetch(`https://apis-navi.kakaomobility.com/v1/directions?origin=${a.x},${a.y}&destination=${b.x},${b.y}`, { headers: kakaoH() })).json();
      const s = r.routes && r.routes[0] && r.routes[0].summary;
      return s ? { km: s.distance / 1000, min: Math.round(s.duration / 60) } : null;
    } catch (e) { return null; }
  }

  // ---------- 일정 화면 ----------
  let schedMode = "day";
  let schedDate = dayStart(new Date());
  let schedReq = 0;
  const schedCustCache = new Map(); // 일정 id → { cust, extra }

  (() => { // v408(2026-11-17 사용자 지시): 일정 화면을 좌우로 밀면 Today는 전날·다음날, Week는 전주·다음주, Month는 전월·다음월
    const v = $("scheduleView"); if (!v) return;
    let sx = 0, sy = 0, st = 0, ok = false;
    // v459(2026-11-19 사용자 지시): 일정 카드 위가 아니라 빈 바탕(화면 아래 남는 곳)을 밀어도 되게 화면 전체에서 받음 (위 막대·아래 메뉴·팝업 위는 제외)
    document.addEventListener("touchstart", (e) => { const t = e.touches[0]; sx = t.clientX; sy = t.clientY; st = Date.now(); ok = !v.hidden && !(e.target.closest && e.target.closest("input, textarea, select, .ev-form, #schedAddBox, .week-scroll, #mainNav, header, .topbar, .modal, .ai-sheet, [role='dialog']")); }, { passive: true });
    document.addEventListener("touchend", (e) => {
      if (!ok) return; const t = e.changedTouches[0], dx = t.clientX - sx, dy = t.clientY - sy;
      if (Date.now() - st > 700 || Math.abs(dx) < 70 || Math.abs(dy) > Math.abs(dx) * 0.6) return;
      const b = $(dx < 0 ? "schedNext" : "schedPrev"); if (b) b.click();
    }, { passive: true });
  })();
  function scheduleShell() {
    return `<div class="sched-top"><div class="sched-modes">${[["day", "Today"], ["week", "Week"], ["month", "Month"]].map(([m, l]) =>
      `<button type="button" data-sm="${m}" class="${schedMode === m ? "on" : ""}">${l}</button>`).join("")}</div></div>
      <div class="sched-nav"><span class="sched-rt"><span class="sched-seg"><button type="button" class="link-btn" id="schedAdd" aria-label="일정 추가" title="일정 추가">추가</button></span></span><button type="button" id="schedPrev" aria-label="이전"><svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M15 6l-6 6 6 6"/></svg></button>
        <button type="button" id="schedLabel" class="sched-label"></button>
        <button type="button" id="schedNext" aria-label="다음"><svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M9 6l6 6-6 6"/></svg></button></div>
      <div id="schedAddBox" hidden></div>
      <div id="schedBody"></div>`;
  }

  async function openSchedule() {
    show("scheduleView");
    const v = $("scheduleView");
    if (!v.dataset.ready) {
      v.innerHTML = scheduleShell();
      v.dataset.ready = "1";
      v.querySelectorAll("[data-sm]").forEach((b) => b.addEventListener("click", () => {
        schedMode = b.dataset.sm;
        if (schedMode === "day") schedDate = dayStart(new Date()); // Today 버튼: 오늘로
        v.querySelectorAll("[data-sm]").forEach((x) => x.classList.toggle("on", x === b));
        renderSchedule();
      }));
      $("schedPrev").addEventListener("click", () => moveSched(-1));
      $("schedNext").addEventListener("click", () => moveSched(1));
      $("schedLabel").addEventListener("click", () => { schedDate = dayStart(new Date()); renderSchedule(); });
      $("schedAdd").addEventListener("click", openSchedAdd);
    }
    v.querySelectorAll("[data-sm]").forEach((x) => x.classList.toggle("on", x.dataset.sm === schedMode));
    const setTbh = () => document.documentElement.style.setProperty("--tbh", Math.round(topbarH()) + "px"); // 날짜 줄을 상단바 바로 아래에 고정
    setTbh();
    if (!window.__tbhBound) { window.__tbhBound = true; window.addEventListener("resize", setTbh); }
    renderSchedule();
  }
  function moveSched(n) {
    if (schedMode === "day") schedDate = addDays(schedDate, n);
    else if (schedMode === "week") schedDate = addDays(schedDate, 7 * n);
    else schedDate = new Date(schedDate.getFullYear(), schedDate.getMonth() + n, 1);
    renderSchedule();
  }
  function setMode(m, d) {
    schedMode = m; schedDate = dayStart(d);
    $("scheduleView").querySelectorAll("[data-sm]").forEach((x) => x.classList.toggle("on", x.dataset.sm === m));
    renderSchedule();
  }

  function needGoogle(msg) {
    schedResumeY = null; schedHold = 0;
    $("schedBody").innerHTML = `<div class="card center"><p class="lead">구글 캘린더 연결</p>
      <p class="hint">${esc(msg || "일정을 보려면 구글 계정으로 한 번 로그인해 주세요.")}${isAdmin() && appSettings.google && appSettings.google.account ? "<br>연결 계정: " + esc(appSettings.google.account) : ""}</p>
      <button type="button" class="primary" id="gConnect">구글 계정 연결</button></div>`;
    $("gConnect").addEventListener("click", async () => {
      try { await requestGToken(); renderSchedule(); } catch (e) { toast(e.message); }
    });
  }

  async function renderSchedule() {
    const req = ++schedReq;
    notifyWhereNow(); // 날짜를 넘기면 말만해요에 보고 있는 날짜를 알림
    { const pr0 = $("pickReset"); if (pr0) { pr0.hidden = false; pr0.dataset.done = ""; pr0.onclick = () => toast("이 날은 초기화할 내용이 없습니다."); } } // v409: 초기화 버튼은 항상 보임
    schedHold = Infinity; // 그리는 동안은 스크롤 위치를 저장하지 않음 (다 그린 뒤 schedApplyResume이 풀어 줌)
    const body = $("schedBody");
    let from, to, label;
    if (schedMode === "day") {
      from = schedDate; to = addDays(from, 1);
      label = (sameDay(from, new Date()) ? "오늘 · " : "") + dayLabel(from);
    } else if (schedMode === "week") {
      from = weekStart(schedDate); to = addDays(from, 7);
      const e = addDays(to, -1);
      label = `${from.getMonth() + 1}/${from.getDate()} ~ ${e.getMonth() + 1}/${e.getDate()}`;
    } else {
      from = new Date(schedDate.getFullYear(), schedDate.getMonth(), 1);
      to = new Date(schedDate.getFullYear(), schedDate.getMonth() + 1, 1);
      label = `${from.getFullYear()}년 ${from.getMonth() + 1}월`;
    }
    $("schedLabel").textContent = label;
    if (!GCFG.clientId) { body.innerHTML = '<p class="hint pad">config.js에 구글 clientId가 없습니다.</p>'; return; }
    if (!storedGToken() && !gTokenMem) return needGoogle();
    body.innerHTML = '<p class="hint pad">일정을 불러오는 중…</p>';
    let events;
    try {
      // 월 보기: 앞뒤 주까지 포함해서 달력 칸을 채움
      const f = schedMode === "month" ? weekStart(from) : from;
      const t = schedMode === "month" ? addDays(weekStart(addDays(to, -1)), 7) : to;
      events = await fetchEvents(f, t);
      if (schedMode === "month") { from = f; to = t; }
    } catch (e) {
      if (req !== schedReq) return;
      if (e.message === "GOOGLE_AUTH") return needGoogle("구글 연결 시간이 지났습니다. 다시 연결해 주세요.");
      if (e.message === "NO_CALENDAR") { body.innerHTML = '<p class="hint pad">보여줄 캘린더가 없습니다. 관리자 › 설정에서 캘린더를 선택해 주세요.</p>'; return; }
      body.innerHTML = ""; return handleError(e);
    }
    if (req !== schedReq) return;
    if (schedMode === "day") return renderDay(events, req);
    if (schedMode === "week") { renderWeek(events, from); return schedApplyResume(); }
    renderMonth(events, from, to); return schedApplyResume();
  }

  const calDot = (e) => `<span class="cal-dot" style="background:${esc(e.cal.color || "var(--brand)")}"></span>`;
  // "10:00 ~ 11:00 스케줄명 - 캘린더" 한 줄
  const evLine = (e) => `<span class="ev-time">${esc(timeText(e))}</span><span class="ev-title">${esc(e.title)}</span> <span class="ev-cal">- ${esc(e.cal.name)}</span>`;

  function renderWeek(events, from) {
    let html = "";
    for (let i = 0; i < 7; i++) {
      const d = addDays(from, i);
      const evs = events.filter((e) => onDay(e, d));
      html += `<div class="card week-day${sameDay(d, new Date()) ? " today" : ""}"><button type="button" class="wd-head" data-day="${+d}">${esc(dayLabel(d))}<span class="hint">${evs.length ? evs.length + "건" : ""}</span><span class="chev">›</span></button>
        ${evs.map((e) => `<div class="ev-row">${calDot(e)}${evLine(e)}</div>`).join("") || '<p class="hint">일정 없음</p>'}</div>`;
    }
    $("schedBody").innerHTML = html;
    $("schedBody").querySelectorAll("[data-day]").forEach((b) => b.addEventListener("click", () => setMode("day", new Date(+b.dataset.day))));
    loadWeekTasks(from, schedReq);
  }

  function renderMonth(events, from, to) {
    const month = schedDate.getMonth();
    let cells = "";
    for (let d = from; d < to; d = addDays(d, 1)) {
      const evs = events.filter((e) => onDay(e, d));
      cells += `<button type="button" class="m-cell${d.getMonth() !== month ? " other" : ""}${sameDay(d, new Date()) ? " today" : ""}" data-day="${+d}">
        <span class="m-num">${d.getDate()}</span>
        ${evs.slice(0, 3).map((e) => `<span class="m-ev" style="border-color:${esc(e.cal.color || "var(--brand)")}">${esc(e.title)}</span>`).join("")}
        ${evs.length > 3 ? `<span class="m-more">+${evs.length - 3}</span>` : ""}</button>`;
    }
    $("schedBody").innerHTML = `<div class="card month"><div class="m-grid">${WD.slice(1).concat(WD[0]).map((w) => `<span class="m-wd">${w}</span>`).join("")}${cells}</div></div>`;
    $("schedBody").querySelectorAll("[data-day]").forEach((b) => b.addEventListener("click", () => setMode("day", new Date(+b.dataset.day))));
  }

  // 고객 정보·이동 시간은 팀 캘린더(1팀) 일정만. 개인 일정(김현성·최수진)은 일정 내용만
  // 영업일정 캘린더 일정만 고객 매칭·거리·시간 (설정에서 지정, 없으면 1팀과 같은 이름)
  let schedSalesIds = new Set(), schedInfoIds = new Set();
  // ---- 영업일정 / 비영업일정: 거리·예상 시간은 '영업일정'만. 일정 카드의 [확정]을 누르면 영업일정, [복귀]는 그날 영업일정을 모두 비영업일정으로.
  //   표시는 구글 일정의 비공개 속성 shlifeSales("1"/"0")에 저장(모든 기기 공유, 캘린더는 그대로). 쓸 권한이 없으면 이 기기에만 기억.
  //   표시가 없는 일정: 2026-11-14까지 만들어진 1팀 캘린더 일정은 영업일정(예전 방식 유지), 그 밖의 모든 일정은 비영업일정
  const SALES_LOCAL_KEY = "shlife.salesMark", SALES_CUTOFF = new Date("2026-11-15T00:00:00+09:00");
  const readSalesLocal = () => { try { return JSON.parse(localStorage.getItem(SALES_LOCAL_KEY) || "{}") || {}; } catch (e) { return {}; } };
  function salesOf(e) {
    if (e.allDay) return false;
    const m = e.sales || readSalesLocal()[e.id] || "";
    if (m === "1") return true;
    if (m === "0") return false;
    const team = (GCFG && GCFG.teamCalendar) || "1팀";
    return schedSalesIds.has(e.cal.id) && e.cal.name === team && !!e.created && e.created < SALES_CUTOFF;
  }
  async function setSalesMark(e, v) { // v = "1" | "0"
    const loc = readSalesLocal(); loc[e.id] = v;
    try { localStorage.setItem(SALES_LOCAL_KEY, JSON.stringify(loc)); } catch (er) { /* 무시 */ }
    try {
      await ensureGWrite();
      await gcalPatch(e.calId, e.eventId, { extendedProperties: { private: { shlifeSales: v } } });
      e.sales = v;
      const l2 = readSalesLocal(); delete l2[e.id]; try { localStorage.setItem(SALES_LOCAL_KEY, JSON.stringify(l2)); } catch (er) { /* 무시 */ }
      return true;
    } catch (er) { console.warn("sales mark saved on this device only", er.message); return false; } // 수정 권한이 없는 캘린더 등: 이 기기에만
  }
  const isTeamEvent = (e) => salesOf(e); // 거리·예상 시간 (영업일정만)
  const isInfoEvent = (e) => schedInfoIds.has(e.cal.id) || schedSalesIds.has(e.cal.id); // 고객 정보 (v410: 종일 일정도 고객 찾기)

  // 동명이인 확인: 일정에 주소를 입력한 고객으로 지정 (이 기기에 저장, '초기화'로 지움)
  const PICK_KEY = "shlife.eventPick";
  // 선택 키: 고객ID가 있으면 ID(동명이인·생년월일 없음도 구분), 없으면 이름|생년월일. 옛 키(이름|생년월일)도 읽을 수 있게 둔다
  const custKeyOld = (c) => norm(baseName(c.name)) + "|" + String(c.birth || "").replace(/\D/g, "");
  const custKey = (c) => { const id = (c.entries || []).map(idOf).find(Boolean); return id ? "id:" + id : custKeyOld(c); };
  function readPicks() { try { return JSON.parse(localStorage.getItem(PICK_KEY) || "{}"); } catch (e) { return {}; } }
  function savePick(eventKey, ck) {
    const p = readPicks();
    if (ck) p[eventKey] = ck; else delete p[eventKey];
    try { localStorage.setItem(PICK_KEY, JSON.stringify(p)); } catch (e) { /* 저장 못 해도 이번 화면엔 반영 */ }
  }

  // 구글 일정 수정 전: 수정 권한이 있는 토큰이 없으면 (버튼을 누른 그 순간에) 로그인 창
  async function ensureGWrite() {
    if (storedGToken() || gTokenMem) return;
    await requestGToken();
  }

  // ---------- 구글 할 일(Tasks): 마감일이 있는 할 일을 일정 화면에 보여 주고, 체크하면 완료로 바꿈 ----------
  //   캘린더·드라이브 로그인과 따로, '할 일 연결' 버튼을 눌러 한 번 허용하면 이 폰에 1시간 단위로 기억 (만료되면 다시 연결)
  const TASKS_KEY = "shlife.gtasks", TASKS_SCOPE = "https://www.googleapis.com/auth/tasks", GTASKS = "https://tasks.googleapis.com/tasks/v1";
  function storedTasksToken() {
    try { const t = JSON.parse(localStorage.getItem(TASKS_KEY) || "null"); if (t && t.exp > Date.now() + 60000) return t.token; } catch (e) { /* 아래로 */ }
    const m = readGAuth(); if (m && m.exp > Date.now() + 60000 && gHasScopes(m.scopes) && m.scopes.some((x) => /(^|\/)tasks$/.test(x))) return m.token; // 로그인 때 할 일 권한도 받은 경우 그 토큰을 같이 씀
    return null;
  }
  function requestTasksToken(prompt) {
    return new Promise((resolve, reject) => {
      const g = window.google && window.google.accounts && window.google.accounts.oauth2;
      if (!g) return reject(new Error("구글 로그인 모듈을 불러오지 못했습니다. 잠시 후 다시 시도하세요."));
      const client = g.initTokenClient({
        client_id: GCFG.clientId, scope: TASKS_SCOPE, hint: (GD ? gAccount() : undefined) || undefined,
        callback: (r) => {
          if (r.error) return reject(new Error("구글 연결 실패: " + r.error));
          if (!g.hasGrantedAllScopes(r, TASKS_SCOPE)) return reject(new Error("할 일 권한을 허용해야 보입니다."));
          try { localStorage.setItem(TASKS_KEY, JSON.stringify({ token: r.access_token, exp: Date.now() + (r.expires_in || 3600) * 1000 })); } catch (e) { /* 이번엔 사용 */ }
          resolve(r.access_token);
        },
        error_callback: (e) => reject(new Error(e && e.type === "popup_closed" ? "구글 로그인 창이 닫혔습니다." : "구글 로그인 창을 열 수 없습니다. 팝업 차단을 확인하세요."))
      });
      client.requestAccessToken({ prompt: prompt === undefined ? "" : prompt });
    });
  }
  // 연결이 풀렸을 때 버튼을 누르기 전에 조용히 한 번 다시 연결 시도 (구글·브라우저가 허락하면 성공, 아니면 그대로 연결 버튼). 5분에 한 번만.
  let tasksSilentBusy = false;
  async function trySilentTasks() {
    if (tasksSilentBusy) return false;
    let last = 0; try { last = Number(sessionStorage.getItem("shlife.tasksSilent") || 0); } catch (e) { /* 무시 */ }
    if (Date.now() - last < 300000) return false;
    try { sessionStorage.setItem("shlife.tasksSilent", String(Date.now())); } catch (e) { /* 무시 */ }
    tasksSilentBusy = true;
    try { await requestTasksToken("none"); return !!storedTasksToken(); } catch (e) { return false; } finally { tasksSilentBusy = false; }
  }
  async function gtasks(path, opt) {
    const t = storedTasksToken();
    if (!t) throw new Error("TASKS_AUTH");
    const res = await fetch(GTASKS + path, { ...(opt || {}), headers: { Authorization: "Bearer " + t, "Content-Type": "application/json" } });
    if (res.status === 401) { try { localStorage.removeItem(TASKS_KEY); } catch (e) {} throw new Error("TASKS_AUTH"); } // 연결 시간이 지남
    if (!res.ok) {
      const txt = await res.text();
      // 403: 구글 클라우드에서 Tasks API를 아직 안 켠 경우와, 권한을 허용하지 않은 경우를 구분
      if (res.status === 403 && /accessNotConfigured|SERVICE_DISABLED|has not been used|is disabled/i.test(txt)) throw new Error("TASKS_API_OFF");
      if (res.status === 403) { try { localStorage.removeItem(TASKS_KEY); } catch (e) {} throw new Error("TASKS_AUTH"); }
      throw new Error("Tasks " + res.status + ": " + txt.slice(0, 120));
    }
    return res.status === 204 ? null : res.json();
  }
  // from(포함) ~ to(미포함) 마감인 할 일, 모든 목록. due 는 날짜(UTC 0시)라 앞 10글자가 날짜
  async function fetchTasks(from, to) {
    const lists = ((await gtasks("/users/@me/lists?maxResults=100")).items) || [];
    const min = ymdStr(from) + "T00:00:00.000Z", max = ymdStr(to) + "T00:00:00.000Z", all = [];
    await Promise.all(lists.map(async (l) => {
      const q = `?showCompleted=true&showHidden=true&maxResults=100&dueMin=${encodeURIComponent(min)}&dueMax=${encodeURIComponent(max)}`;
      const r = await gtasks("/lists/" + encodeURIComponent(l.id) + "/tasks" + q);
      (r.items || []).forEach((t) => { if (t.due) all.push({ id: t.id, listId: l.id, list: l.title || "", title: t.title || "(제목 없음)", done: t.status === "completed", due: String(t.due).slice(0, 10), owner: taskOwnerOf(t.notes), pick: taskPickOf(t.notes) }); });
    }));
    return all.filter((t) => t.due >= ymdStr(from) && t.due < ymdStr(to)).sort((a, b) => a.due.localeCompare(b.due) || a.title.localeCompare(b.title, "ko"));
  }
  // v323: 미완료 할 일 전부(마감 조건 없음). Today 카드의 '밀린 할 일'·'날짜 없음' 묶음용
  async function fetchOpenTasks() {
    const lists = ((await gtasks("/users/@me/lists?maxResults=100")).items) || [], all = [];
    await Promise.all(lists.map(async (l) => {
      let tok = "", n = 0;
      do {
        const r = await gtasks("/lists/" + encodeURIComponent(l.id) + "/tasks?showCompleted=false&showHidden=true&maxResults=100" + (tok ? "&pageToken=" + encodeURIComponent(tok) : ""));
        (r.items || []).forEach((t) => { if (t.status !== "completed" && !t.deleted && String(t.title || "").trim()) all.push({ id: t.id, listId: l.id, list: l.title || "", title: t.title || "(제목 없음)", done: false, owner: taskOwnerOf(t.notes), pick: taskPickOf(t.notes), due: t.due ? String(t.due).slice(0, 10) : "" }); });
        tok = r.nextPageToken || "";
      } while (tok && ++n < 20);
    }));
    return all;
  }
  const taskConnectHtml = (msg) => `<div class="card task-card"><h3 class="sec-title">할 일</h3><p class="hint">${esc(msg || "구글 할 일을 보려면 한 번 연결해 주세요.")}</p><button type="button" class="ev-edit" data-taskconnect>할 일 연결</button></div>`;
  // v431(2026-11-18 사용자 지시): 구글 할 일은 로그인한 구글 계정 것이라 '사용자'는 메모 맨 윗줄의 '작성: 이름'으로 기록(구글 할 일 앱에서도 보임). 관리자가 추가·수정할 때 사용자를 고르고, 본인으로 두면 이 줄은 없음
  const taskOwnerOf = (notes) => { const m = /^작성:[ \t]*(.+)$/m.exec(String(notes || "")); return m ? m[1].trim() : ""; };
  const taskNotesWith = (notes, name) => { const rest = String(notes || "").replace(/^작성:.*(\r?\n)?/m, "").replace(/^\s+/, ""); return name ? "작성: " + name + (rest ? "\n" + rest : "") : rest; };
  // v437(2026-11-18 사용자 지시): 할 일 제목의 이름으로 고객 찾기 → 동명이인 선택·해당 없음은 할 일 메모의 '고객연결: 키' 줄에 저장(모든 기기 공유, 구글 할 일 앱에서도 보임)
  const taskPickOf = (notes) => { const m = /^고객연결:[ \t]*(.+)$/m.exec(String(notes || "")); return m ? m[1].trim() : ""; };
  const taskNotesPick = (notes, key) => { const rest = String(notes || "").replace(/^고객연결:.*(\r?\n)?/m, "").replace(/\s+$/, ""); return key ? (rest ? rest + "\n" : "") + "고객연결: " + key : rest; };
  const taskMyEmail = () => String(account().username).toLowerCase();
  const taskSelfOpt = () => `<option value="${esc(taskMyEmail())}" selected>${esc((permNamesCache || {})[taskMyEmail()] || taskMyEmail().split("@")[0])} (관리자)</option>`;
  // v436(2026-11-18 사용자 지시): 사용자·작성자 선택지는 '그 캘린더에 쓸 수 있는 사람'만. 구글 캘린더 공유 목록(ACL)은 지금 받은 권한 범위로는 읽을 수 없어서(403),
  // 대신 '내 캘린더 목록에서 내가 쓸 수 있는(owner·writer) 캘린더의 주인'을 씀(개인 캘린더의 id = 그 사람 이메일). 목록을 못 받으면 본인만.
  async function writablePeople() {
    try { return new Set((await calendarList()).filter((c) => c.role === "owner" || c.role === "writer").map((c) => String(c.id).toLowerCase())); } catch (e) { return new Set(); }
  }
  function fillTaskUsers(sel, curName) { // 관리자: 본인 + 등록된 사용자 (일정 추가의 작성자와 같은 사람들)
    Promise.all([readPerms(), writablePeople()]).then(([pm, wp]) => {
      if (!sel) return;
      const my = taskMyEmail(), nm = (em) => (pm.names || {})[em] || em.split("@")[0];
      Object.keys(pm.users || {}).filter((em) => em !== my && !/@email\.pending$/.test(em) && wp.has(String(em).toLowerCase())).sort((a, b) => nm(a).localeCompare(nm(b), "ko"))
        .forEach((em) => sel.insertAdjacentHTML("beforeend", `<option value="${esc(em)}">${esc(nm(em))}</option>`));
      if (curName) { const o = [...sel.options].find((x) => x.textContent === curName); if (o) sel.value = o.value; else { sel.insertAdjacentHTML("beforeend", `<option value="__cur">${esc(curName)}</option>`); sel.value = "__cur"; } }
    }).catch(() => {});
  }
  const taskOwnerPick = (sel) => (!sel || sel.value === taskMyEmail() ? "" : sel.selectedOptions[0].textContent.replace(/ \(관리자\)$/, ""));
  let taskOwnerCal = "";
  // 할 일 소유주 = 앱에 등록된 사용자 이름(관리자 화면의 이름), 없으면 내 기본 캘린더 이름
  async function taskOwnerName() {
    if (taskOwnerCal) return taskOwnerCal;
    let n = ""; try { n = (await myFcName()) || ""; } catch (e) { n = ""; }
    if (!n) { try { const pc = (await calendarList()).find((c) => c.primary); n = pc ? pc.name : ""; } catch (e) { n = ""; } }
    return n;
  }
  let taskShowList = true, taskReload = null;
  const taskItemHtml = (t, grp) => `<li class="task-item${t.done ? " done" : ""}"${grp ? ` data-tgrp="${grp}"` : ""} data-tid="${esc(t.listId)}|${esc(t.id)}" data-ttitle="${esc(t.title)}" data-tdue="${esc(t.due || "")}" data-towner="${esc(t.owner || "")}" data-tpick="${esc(t.pick || "")}"><label><input type="checkbox" class="cb" data-task="${esc(t.listId)}|${esc(t.id)}" ${t.done ? "checked" : ""}>${grp === "late" && t.due ? ` <span class="task-due">${esc(t.due.slice(5))}</span>` : ""} <span class="task-title">${esc(t.title)}</span>${(t.owner || taskOwnerCal) ? ` <span class="ev-cal">- ${esc(t.owner || taskOwnerCal)}</span>` : ""}</label>${t.list && taskShowList ? `<span class="hint task-list-name">${esc(t.list)}</span>` : ""}<span class="task-acts"><button type="button" class="ev-edit task-act" data-ttoev title="이 할 일을 캘린더 일정으로 옮깁니다">일정으로</button><button type="button" class="ev-edit task-act" data-tedit>수정</button><button type="button" class="ev-edit ev-del task-act" data-tdel>삭제</button></span></li>`;
  // v449(2026-11-19 사용자 지시): 할 일도 제목의 이름으로 고객을 찾아 '일정에 붙는 고객 카드'를 그대로 붙임
  //   (이름·분류·생년월일·연락처·주소·미팅장소·상담기록, 동명이인이면 [선택] → [다시 선택], 관계없으면 [해당 없음] → [되돌리기]). 고른 값은 할 일 메모의 '고객연결: 키' 줄
  const TASK_CUST = new Map(); // 할 일 id → { all, custs }
  async function taskCustFill(box, req) {
    const lis = [...box.querySelectorAll(".task-item")]; if (!lis.length) return;
    let records = [], contracts = [];
    try { [records, contracts] = await Promise.all([loadLists(), loadContracts().catch(() => [])]); } catch (e) { return; }
    if (req !== schedReq) return;
    const jobs = [];
    lis.forEach((li, n) => {
      const tid = li.dataset.tid, pk = li.dataset.tpick || "";
      const all = matchCustomers(li.dataset.ttitle || "", records, contracts, new Date());
      if (!all.length) return;
      let custs = all, pinned = false;
      const ignored = pk === IGNORE_KEY;
      if (ignored) custs = [];
      else if (pk) {
        let one = all.filter((c) => custKey(c) === pk);
        if (!one.length) { const old = all.filter((c) => custKeyOld(c) === pk); if (old.length === 1) one = old; }
        if (one.length) { custs = one.slice(0, 1); pinned = true; }
      }
      TASK_CUST.set(tid, { all, custs });
      const ti = "t" + n; // 일정 카드의 번호 자리에 't번호'를 써서 같은 카드 모양을 그대로 씀
      const acts = ignored ? `<small class="hint">검색된 고객 ${all.length}명을 숨겼어요</small><button type="button" class="ev-edit" data-tunign>되돌리기</button>` : pinned ? "" : `<button type="button" class="ev-edit" data-tign>해당 없음</button>`;
      li.insertAdjacentHTML("beforeend", `<div class="task-cust${pinned ? " linked" : ""}" data-ti="${ti}"><div class="task-cust-acts">${acts}</div>${custs.map((c, k) => custBlockHtml(c, ti, k, !pinned, pinned)).join("")}</div>`);
      custs.forEach((c, k) => jobs.push((async () => { // 미팅장소·상담기록
        const fc = (c.entries.find((e) => e.fc) || {}).fc || "";
        try { c.extra = await loadExtraFor(fc, c.name, c); } catch (e) { c.extra = null; }
        if (req === schedReq && $(`cs${ti}-${k}`)) fillCust(null, c, ti, k);
      })()));
    });
    await Promise.all(jobs);
  }
  document.addEventListener("click", async (ev) => {
    const b = ev.target && ev.target.closest && ev.target.closest(".task-cust [data-open], .task-cust [data-pick], .task-cust [data-unpick], .task-cust [data-tign], .task-cust [data-tunign]");
    if (!b) return;
    const li = b.closest(".task-item"); if (!li) return;
    const tc = TASK_CUST.get(li.dataset.tid) || { all: [], custs: [] };
    if (b.hasAttribute("data-open")) { const c = tc.custs[+b.dataset.open.split("-")[1]]; if (c) openCustomer(c); return; }
    const key = b.hasAttribute("data-pick") ? custKey(tc.custs[+b.dataset.pick.split("-")[1]]) : b.hasAttribute("data-tign") ? IGNORE_KEY : ""; // [다시 선택]·[되돌리기] = 연결 값 지움
    b.disabled = true; b.textContent = "저장 중";
    try {
      const path = taskPath(li.dataset.tid), cur = await gtasks(path), nn = taskNotesPick(cur.notes, key);
      await gtasks(path, { method: "PATCH", body: JSON.stringify({ notes: nn || null }) });
      if (taskReload) taskReload();
    } catch (e) { b.disabled = false; taskErr(e); }
  });
  // 하루 보기 맨 위 '할 일' 카드
  async function loadDayTasks(box, d, req) {
    if (!box) return;
    if (!storedTasksToken()) { box.innerHTML = taskConnectHtml(); trySilentTasks().then((ok) => { if (ok && req === schedReq) loadDayTasks(box, d, req); }); return; }
    try {
      const today = ymdStr(dayStart(new Date())), isToday = ymdStr(d) === today;
      const [list, open] = await Promise.all([fetchTasks(d, addDays(d, 1)), isToday ? fetchOpenTasks().catch((e) => { console.warn("tasks open", e.message); return []; }) : Promise.resolve([])]);
      if (req !== schedReq) return;
      const late = open.filter((t) => t.due && t.due < today).sort((a, b) => a.due.localeCompare(b.due) || a.title.localeCompare(b.title, "ko"));
      const nodue = open.filter((t) => !t.due).sort((a, b) => a.title.localeCompare(b.title, "ko"));
      // 소유주: 구글 할 일은 로그인한 구글 계정 것이라, 앱에 등록된 내 이름(없으면 구글 계정)을 제목에 붙임
      taskOwnerCal = await taskOwnerName();
      if (req !== schedReq) return;
      const grpHtml = (key, label, arr) => arr.length ? `<div class="task-grp" data-tgrpbox="${key}"><div class="task-grp-h">${label} <small>${arr.length}</small></div><ul class="task-list">${arr.map((t) => taskItemHtml(t, key)).join("")}</ul></div>` : "";
      const total = list.length + late.length + nodue.length;
      taskShowList = new Set([...list, ...late, ...nodue].map((t) => t.listId)).size > 1;
      taskReload = () => loadDayTasks(box, d, schedReq);
      box.innerHTML = total ? `<div class="card task-card"><h3 class="sec-title">할 일 <small class="hint task-cnt">${list.filter((t) => !t.done).length + late.length + nodue.length}/${total}</small></h3>${list.length ? `<ul class="task-list">${list.map((t) => taskItemHtml(t)).join("")}</ul>` : ""}${grpHtml("late", "밀린 할 일", late)}${grpHtml("nodue", "날짜 없음", nodue)}</div>` : "";
      if (total) taskCustFill(box, req);
    } catch (e) {
      if (req !== schedReq) return;
      box.innerHTML = e.message === "TASKS_AUTH" ? taskConnectHtml("할 일 연결 시간이 지났거나 할 일 권한을 허용하지 않았습니다. 다시 연결해 주세요.")
        : e.message === "TASKS_API_OFF" ? taskConnectHtml("구글 클라우드에서 'Google Tasks API'를 아직 켜지 않아 할 일을 불러올 수 없습니다. 관리자가 사용 설정을 한 뒤(몇 분 걸릴 수 있음) 다시 눌러 주세요.")
        : taskConnectHtml("할 일을 불러오지 못했습니다: " + e.message);
      if (e.message !== "TASKS_AUTH") console.warn("tasks", e.message);
    }
  }
  // 주간 보기: 요일 카드마다 그날 마감인 할 일을 한 줄씩 덧붙임 (보기만)
  async function loadWeekTasks(from, req) {
    if (!storedTasksToken()) { trySilentTasks().then((ok) => { if (ok && req === schedReq) loadWeekTasks(from, req); }); return; }
    try {
      const list = await fetchTasks(from, addDays(from, 7));
      taskOwnerCal = await taskOwnerName();
      if (req !== schedReq) return;
      const cards = [...document.querySelectorAll("#schedBody .week-day")];
      list.forEach((t) => {
        const i = Math.round((new Date(t.due + "T00:00:00") - from) / 86400000);
        const card = cards[i]; if (!card) return;
        const none = card.querySelector("p.hint"); if (none && none.textContent.trim() === "일정 없음") none.remove();
        card.insertAdjacentHTML("beforeend", `<div class="ev-row task-row${t.done ? " done" : ""}">${t.done ? "☑" : "☐"} ${esc(t.title)}${(t.owner || taskOwnerCal) ? ` <span class="ev-cal">- ${esc(t.owner || taskOwnerCal)}</span>` : ""}</div>`);
      });
    } catch (e) { /* 연결이 없거나 만료면 주간 보기엔 아무 것도 덧붙이지 않음 */ }
  }
  // v450(2026-11-19 사용자 지시): 할 일 → 일정. 메모에 예전 시간('시간: 09:00~10:00', 일정에서 옮길 때 적힌 것)이 있으면 그 시간으로, 없으면 종일 일정으로 만들고 할 일은 지움
  async function taskToEvent(li, btn) {
    const tid = li.dataset.tid, title = li.dataset.ttitle || "", cals = schedCals || [];
    if (!cals.length) return toast("일정을 불러온 뒤에 옮길 수 있습니다.");
    btn.disabled = true;
    let made = false;
    try {
      const cur = await gtasks(taskPath(tid)), notes = String(cur.notes || "");
      const p2 = (v) => v.replace(/^(\d):/, "0$1:");
      let tm = /^시간:\s*(\d{1,2}:\d{2})\s*~\s*(\d{1,2}:\d{2})/m.exec(notes);
      if (tm && p2(tm[2]) <= p2(tm[1])) tm = null;
      const loc = ((/^장소:\s*(.+)$/m.exec(notes) || [])[1] || "").trim();
      const cn = (/\(캘린더 '([^']+)' (?:종일 )?일정에서 옮김\)/.exec(notes) || [])[1] || "";
      const team = (GCFG && GCFG.teamCalendar) || "";
      const cal = cals.find((c) => c.name === cn) || cals.find((c) => c.name === team) || cals[0];
      const date = (cur.due ? String(cur.due).slice(0, 10) : "") || li.dataset.tdue || ymdStr(schedDate || new Date());
      if (!confirm(`'${title}' 할 일을 일정으로 옮길까요?` + String.fromCharCode(10) + `${date} ${tm ? p2(tm[1]) + "~" + p2(tm[2]) : "종일"} · '${cal.name}' 캘린더에 일정을 만들고, 이 할 일은 지웁니다.`)) { btn.disabled = false; return; }
      btn.textContent = "옮기는 중";
      const desc = notes.split(/\r?\n/).filter((l) => !/^(작성|고객연결|시간|장소):/.test(l) && !/^\(캘린더 '.*일정에서 옮김\)$/.test(l.trim())).join("\n").trim();
      const body = { summary: cur.title || title };
      if (loc) body.location = loc;
      if (desc) body.description = desc;
      if (tm) { const tz = (Intl.DateTimeFormat().resolvedOptions() || {}).timeZone || "Asia/Seoul"; body.start = { dateTime: `${date}T${p2(tm[1])}:00`, timeZone: tz }; body.end = { dateTime: `${date}T${p2(tm[2])}:00`, timeZone: tz }; }
      else { body.start = { date }; body.end = { date: ymdStr(addDays(ymdToDate(date), 1)) }; }
      const owner = taskOwnerOf(notes), pick = taskPickOf(notes);
      if (owner || pick) body.extendedProperties = { ...(owner ? { shared: { shlifeAuthorName: owner } } : {}), ...(pick ? { private: { shlifeCust: pick } } : {}) }; // 작성자·고객 연결도 이어 줌
      await ensureGWrite();
      await gcalInsert(cal.id, body);
      made = true;
      await gtasks(taskPath(tid), { method: "DELETE" });
      toast(`일정으로 옮겼습니다 (${tm ? p2(tm[1]) + "~" + p2(tm[2]) : "종일"}).`);
      renderSchedule();
    } catch (e) {
      btn.disabled = false; btn.textContent = "일정으로";
      const m = e.message === "TASKS_AUTH" ? "할 일 연결이 만료됐습니다. 다시 연결해 주세요." : e.message === "GOOGLE_AUTH" ? "구글 연결이 끊겼습니다. 다시 눌러 주세요." : e.message;
      toast(made ? "일정은 만들었지만 할 일을 지우지 못했습니다: " + m : m);
      if (made) renderSchedule();
    }
  }
  document.addEventListener("click", (ev) => { const b = ev.target && ev.target.closest && ev.target.closest(".task-item [data-ttoev]"); if (b && !b.disabled) taskToEvent(b.closest(".task-item"), b); });
  // v324: 할 일 수정·삭제
  const taskPath = (tid) => { const [l, t] = tid.split("|"); return "/lists/" + encodeURIComponent(l) + "/tasks/" + encodeURIComponent(t); };
  const taskErr = (e) => toast(e.message === "TASKS_AUTH" ? "할 일 연결이 만료됐습니다. 다시 연결해 주세요." : e.message);
  document.addEventListener("click", async (ev) => {
    const t = ev.target && ev.target.closest && ev.target.closest("[data-tedit],[data-tdel],[data-tsave],[data-tcancel]");
    if (!t) return;
    const li = t.closest(".task-item"); if (!li) return;
    const tid = li.dataset.tid, card = li.closest(".task-card");
    if (t.hasAttribute("data-tedit")) {
      li.dataset.html = li.innerHTML; li.classList.add("editing");
      li.innerHTML = `<input type="text" class="task-in" value="${esc(li.dataset.ttitle)}" aria-label="할 일 제목"><span class="task-ed"><input type="text" data-dp="date" class="task-date" value="${esc(li.dataset.tdue)}" aria-label="마감일"><button type="button" class="ev-edit task-act" data-tcancel>취소</button><button type="button" class="ev-edit ev-done task-act" data-tsave>저장</button></span>`;
      if (isAdmin()) { // 관리자: 사용자 선택(기본 = 지금 사용자)
        const sel = document.createElement("select"); sel.className = "task-owner"; sel.setAttribute("aria-label", "사용자"); sel.innerHTML = taskSelfOpt();
        li.querySelector(".task-ed").insertBefore(sel, li.querySelector("[data-tcancel]")); /* 마감일 → 사용자 → 취소·저장 */ fillTaskUsers(sel, li.dataset.towner || "");
      }
      const inp = li.querySelector(".task-in"); inp.focus(); inp.setSelectionRange(inp.value.length, inp.value.length);
      return;
    }
    if (t.hasAttribute("data-tcancel")) { li.innerHTML = li.dataset.html || ""; li.classList.remove("editing"); return; }
    if (t.hasAttribute("data-tsave")) {
      const title = li.querySelector(".task-in").value.trim(), due = li.querySelector(".task-date").value;
      if (!title) { toast("제목을 적어 주세요."); return; }
      li.querySelectorAll("button,input").forEach((x) => { x.disabled = true; });
      try {
        const patch = { title, due: due ? due + "T00:00:00.000Z" : null };
        const osel = li.querySelector(".task-owner");
        if (osel) { // 사용자를 바꿨으면 메모 맨 윗줄 '작성: 이름'만 바꾸고 나머지 메모는 그대로
          const cur = await gtasks(taskPath(tid)), nn = taskNotesWith(cur.notes, taskOwnerPick(osel));
          if (nn !== String(cur.notes || "")) patch.notes = nn || null;
        }
        await gtasks(taskPath(tid), { method: "PATCH", body: JSON.stringify(patch) });
        toast("할 일을 고쳤어요"); if (taskReload) taskReload();
      } catch (e) { li.querySelectorAll("button,input").forEach((x) => { x.disabled = false; }); taskErr(e); }
      return;
    }
    // 삭제
    if (!confirm(`'${li.dataset.ttitle}' 할 일을 지울까요?`)) return;
    t.disabled = true;
    try {
      const old = await gtasks(taskPath(tid));
      await gtasks(taskPath(tid), { method: "DELETE" });
      const box = li.closest(".task-grp"); li.remove(); if (box && !box.querySelector("li")) box.hidden = true;
      if (card) taskCount(card);
      undoToast("할 일을 지웠어요", async () => {
        try {
          const body = { title: old.title || li.dataset.ttitle }; if (old.notes) body.notes = old.notes; if (old.due) body.due = old.due; if (old.status === "completed") body.status = "completed";
          await gtasks("/lists/" + encodeURIComponent(tid.split("|")[0]) + "/tasks", { method: "POST", body: JSON.stringify(body) });
          if (taskReload) taskReload();
        } catch (e) { taskErr(e); }
      });
    } catch (e) { t.disabled = false; taskErr(e); }
  });
  // 카드 제목 카운트 = 미완료/전체 (밀린 할 일·날짜 없음 포함, 빠진 줄은 제외)
  function taskCount(card) {
    const el = card.querySelector(".task-cnt"), lis = [...card.querySelectorAll(".task-item")];
    if (el) el.textContent = lis.filter((li) => !li.classList.contains("done")).length + "/" + lis.length;
    card.querySelectorAll(".task-grp").forEach((g) => { const n = g.querySelectorAll(".task-item").length, sm = g.querySelector(".task-grp-h small"); if (sm) sm.textContent = n; });
  }
  // 체크 → 완료/미완료 저장, '할 일 연결' 버튼
  document.addEventListener("change", async (ev) => {
    const cb = ev.target && ev.target.closest && ev.target.closest("[data-task]");
    if (!cb) return;
    const [listId, taskId] = cb.dataset.task.split("|"), want = cb.checked;
    cb.disabled = true;
    try {
      await gtasks("/lists/" + encodeURIComponent(listId) + "/tasks/" + encodeURIComponent(taskId), { method: "PATCH", body: JSON.stringify(want ? { status: "completed" } : { status: "needsAction", completed: null }) });
      const li = cb.closest("li"); if (li) li.classList.toggle("done", want);
      const grp = li && li.dataset.tgrp, card = cb.closest(".task-card");
      if (want && grp && card) { // 밀린 할 일·날짜 없음: 완료되면 묶음에서 빠짐
        const box = li.closest(".task-grp"), ul = li.parentNode, next = li.nextSibling;
        li.remove(); if (box && !box.querySelector("li")) box.hidden = true;
        taskCount(card);
        undoToast("완료했어요", async () => {
          try {
            await gtasks("/lists/" + encodeURIComponent(listId) + "/tasks/" + encodeURIComponent(taskId), { method: "PATCH", body: JSON.stringify({ status: "needsAction", completed: null }) });
            cb.checked = false; li.classList.remove("done"); ul.insertBefore(li, next && next.parentNode === ul ? next : null); if (box) box.hidden = false; taskCount(card);
          } catch (e) { toast(e.message === "TASKS_AUTH" ? "할 일 연결이 만료됐습니다. 다시 연결해 주세요." : e.message); }
        });
      } else if (card) taskCount(card);
    } catch (e) { cb.checked = !want; toast(e.message === "TASKS_AUTH" ? "할 일 연결이 만료됐습니다. 다시 연결해 주세요." : e.message); }
    cb.disabled = false;
  });
  document.addEventListener("click", async (ev) => {
    const b = ev.target && ev.target.closest && ev.target.closest("[data-taskconnect]");
    if (!b) return;
    b.disabled = true;
    try { await requestTasksToken(); renderSchedule(); } catch (e) { toast(e.message); b.disabled = false; }
  });
  async function renderDay(events, req) {
    const d = schedDate;
    const evs = events.filter((e) => onDay(e, d));
    const body = $("schedBody");
    if (!evs.length) { body.innerHTML = '<div id="taskBox"></div><div class="card center"><p class="hint">일정이 없습니다.</p></div>'; loadDayTasks($("taskBox"), d, req); schedApplyResume(); return; }
    body.innerHTML = '<div id="taskBox"></div><div id="dayList"></div>';
    loadDayTasks($("taskBox"), d, req); // 맨 위에 그날 마감인 할 일 (일정을 기다리지 않고 따로 불러옴)
    const pr = $("pickReset"); // 복귀 버튼: 고객 지정·무시·일정 장소 입력처럼 처리한 일정이 있을 때만 보임 (아래에서 판단, 처음엔 추가 버튼만)
    // 초기화: 동명이인 지정을 지우고, 고객 주소·미팅장소로 '입력'해 둔 일정 장소도 지움
    let dayItems = [];
    if (pr) pr.onclick = async () => { // v440: 일정 위쪽 [초기화] 버튼은 없앰(버튼이 없으면 연결 안 함)
      if (!pr.dataset.done) return toast("이 날은 초기화할 내용이 없습니다.");
      const filled = dayItems.filter((it) => it.e.location && it.all.some((c) =>
        [currentValue(c, ["주소"]), meetingOf(c)].map((v) => String(v || "").trim()).filter(Boolean).includes(it.e.location.trim())));
      const pinned = dayItems.filter((it) => it.e.pick);
      const salesNow = dayItems.filter((it) => salesOf(it.e));
      const msg = (salesNow.length ? `영업일정 ${salesNow.length}건을 모두 비영업일정으로 바꾸고, ` : "") + "동명이인 확인 내용을 지우고, 찾은 고객을 모두 다시 보여줄까요?" +
        (filled.length ? `\n고객 정보로 입력한 일정 장소 ${filled.length}건도 지웁니다.` : "");
      if (!confirm(msg)) return;
      for (const it of salesNow) await setSalesMark(it.e, "0"); // 영업일정 → 비영업일정
      try { localStorage.removeItem(PICK_KEY); } catch (e) {}
      if (pinned.length) {
        try {
          await ensureGWrite();
          for (const it of pinned) await calLinkSave(it, { clear: true, how: "해제" });
        } catch (e) { /* 연결이 끊겼으면 이 기기 기록만 지움 */ }
      }
      if (filled.length) {
        try {
          await ensureGWrite();
          for (const it of filled) await gcalPatch(it.e.calId, it.e.eventId, { location: "" });
          toast(`일정 장소 ${filled.length}건을 지웠습니다.`);
        } catch (e) { toast(e.message === "GOOGLE_AUTH" ? "구글 연결이 끊겼습니다. 다시 눌러 주세요." : e.message); }
      }
      renderSchedule();
    };
    // 고객 찾기 (고객 리스트 엑셀 + 계약일로 최근 등록 순)
    let records = [], contracts = [];
    if (evs.some((e) => isInfoEvent(e) || e.allDay)) { // v449: 종일 일정은 캘린더와 상관없이 고객 찾기
      try { [records, contracts] = await Promise.all([loadLists(), loadContracts().catch(() => [])]); } catch (e) { handleError(e); }
    }
    if (req !== schedReq) return;
    const picks = readPicks();
    const items = evs.map((e) => {
      const all = (isInfoEvent(e) || e.allDay) ? matchCustomers(e.title, records, contracts, e.start) : [];
      let custs = all, pinned = false; // pinned: 이 일정이 고객 한 명(고객ID)과 연결된 상태 (검색된 고객이 한 명이어도 '선택'으로 확인해 연결)
      const pk = e.pick || picks[e.id]; // 구글 일정에 저장된 값(모든 기기 공유)을 먼저
      const ignored = pk === IGNORE_KEY && all.length > 0;
      if (ignored) custs = [];
      else if (pk && custs.length) {
        let one = custs.filter((c) => custKey(c) === pk);
        if (!one.length) { const old = custs.filter((c) => custKeyOld(c) === pk); if (old.length === 1) one = old; } // 예전에 저장된 키(한 명으로 좁혀질 때만)
        if (one.length) { custs = one.slice(0, 1); pinned = true; }
      }
      return { e, custs, all, ignored, pinned };
    });
    dayItems = items;
    setTimeout(() => calAutoLink(items, req), 0); // v559: 자동 연결·옛 값 고치기를 일정 숨은 칸과 기록 파일에 (화면은 기다리지 않음)
    { // 처리한 일이 하나라도 있으면 복귀 버튼 표시
      const done = items.some((it) => salesOf(it.e) || it.pinned || it.ignored || (it.e.location && it.all.some((c) => [currentValue(c, ["주소"]), meetingOf(c)].map((v) => String(v || "").trim()).filter(Boolean).includes(it.e.location.trim()))));
      if (pr) { pr.hidden = false; pr.dataset.done = done ? "1" : ""; } // v409: 항상 보이고, 처리한 일이 없으면 눌러도 안내만
    }
    items.forEach((it) => { it.cust = it.custs[0] || null; });
    const list = $("dayList");
    list.innerHTML = items.map((it, i) => `<div class="travel" id="tr${i}" hidden></div>${dayCardHtml(it, i)}`).join("") +
      `<div class="travel" id="tr${items.length}" hidden></div>`;
    items.forEach((it, i) => bindDayCard(it, i, req));
    schedApplyResume();
    // 추가정보(미팅장소, 상담기록)
    await Promise.all(items.map(async (it, i) => {
      await Promise.all(it.custs.map(async (c, j) => {
        const fc = (c.entries.find((e) => e.fc) || {}).fc || "";
        try { c.extra = await loadExtraFor(fc, c.name, c); } catch (e) { c.extra = null; }
        if (req !== schedReq) return;
        fillCust(it, c, i, j);
      }));
      it.extra = it.cust && it.cust.extra;
    }));
    if (req !== schedReq) return;
    renderTravel(items, req);
  }

  // 일정 카드의 한 줄 메모 → 그 고객의 상담기록(고객 정보 › 상담과 같은 곳)에 날짜·시간과 함께 추가
  const memoFcOf = (c) => primaryFc(c.entries) || (c.entries.find((x) => x.fc) || {}).fc || "";
  // 마이크 버튼: 누르면 말하는 내용이 바로 메모 칸에 글자로 들어감 (브라우저 음성 인식, 다시 누르면 멈춤)
  let micRec = null, micBtn = null;
  function micStop() { if (micRec) { try { micRec.stop(); } catch (e) { /* 이미 멈춤 */ } } }
  function toggleMic(i, j, btn) { dictate($("memo" + i + "-" + j), btn); }
  function dictate(inp, btn, onEnd) {
    if (!inp) return;
    if (micRec && micBtn === btn) return micStop(); // 녹음 중 다시 누르면 멈춤
    micStop();
    wakeStop(); // '몰리야' 듣기와 마이크가 겹치지 않게
    const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SR) return toast("이 브라우저는 음성 입력을 지원하지 않습니다. 키보드의 마이크 버튼을 써 주세요.");
    const rec = new SR();
    rec.lang = "ko-KR"; rec.continuous = true; rec.interimResults = true;
    const base = inp.value.trim();
    let finalT = "";
    const iconHtml = btn.innerHTML;
    const stopOnTouch = () => micStop(); // 글씨 쓰는 칸을 터치하면 녹음을 멈추고 그대로 직접 고칠 수 있게
    const put = (interim) => {
      inp.value = [base, (finalT + interim).trim()].filter(Boolean).join(" ");
      inp.dispatchEvent(new Event("input", { bubbles: true })); // 저장 버튼 켜기
    };
    rec.onresult = (ev) => {
      let interim = "";
      for (let k = ev.resultIndex; k < ev.results.length; k++) {
        const t = ev.results[k][0].transcript;
        if (ev.results[k].isFinal) finalT += t; else interim += t;
      }
      put(interim);
    };
    rec.onerror = (ev) => {
      toast(ev.error === "not-allowed" || ev.error === "service-not-allowed" ? "마이크 사용을 허용해 주세요." : ev.error === "no-speech" ? "말소리가 들리지 않았습니다." : "음성 입력을 하지 못했습니다.");
    };
    rec.onend = () => {
      setTimeout(wakeStart, 1500); // '몰리야' 듣기 다시
      btn.classList.remove("rec"); btn.setAttribute("aria-pressed", "false"); btn.innerHTML = iconHtml; // 마이크 모양으로 되돌림
      inp.removeEventListener("pointerdown", stopOnTouch);
      if (micRec === rec) { micRec = null; micBtn = null; }
      if (typeof onEnd === "function") onEnd(); // 예: 고객 검색은 말이 끝나면 바로 검색
    };
    try {
      rec.start();
      micRec = rec; micBtn = btn;
      btn.classList.add("rec"); btn.setAttribute("aria-pressed", "true");
      btn.textContent = "완료"; // 녹음 중에는 '완료' 버튼: 누르면 녹음을 끝냄
      inp.addEventListener("pointerdown", stopOnTouch);
      toast("말씀하세요. 끝나면 '완료'를 누르세요.");
    } catch (e) { toast("음성 입력을 시작하지 못했습니다."); }
  }
  // ---- 전화·문자를 누른 뒤 앱으로 돌아오면 "상담 내용을 남길까요?" 물어보기 (추가 내용은 글자나 마이크로 입력) ----
  const PC_KEY = "shlife.pendingContact";
  let pcAway = false;
  document.addEventListener("click", (ev) => {
    const a = ev.target.closest && ev.target.closest('a[href^="tel:"], a[href^="sms:"]');
    if (!a || a.closest("#salesView, #contactSheet")) return; // 세일즈 화면은 따로 '시도'를 기록함
    let name = "", fc = "";
    const ctx = a.closest("[data-cn]");
    if (ctx) { name = ctx.dataset.cn; fc = ctx.dataset.cfc || ""; }
    else if (a.closest("#customerView") && current && current.cust) { name = current.cust.name; fc = current.fc || ""; }
    if (!name || (fc && isRO(fc))) return;
    const kind = a.getAttribute("href").startsWith("tel:") ? "전화" : "문자";
    try { localStorage.setItem(PC_KEY, JSON.stringify({ name, fc, kind, t: Date.now() })); } catch (e) { /* 저장 못 하면 묻지 않음 */ }
    pcAway = false;
  }, true);
  function askContactNote() {
    let pc = null;
    try { pc = JSON.parse(localStorage.getItem(PC_KEY) || "null"); } catch (e) { pc = null; }
    if (!pc) return;
    if (Date.now() - pc.t > 45 * 60 * 1000) { try { localStorage.removeItem(PC_KEY); } catch (e) {} return; }
    if (document.getElementById("contactSheet") || $("loginView") && !$("loginView").hidden) return;
    const box = document.createElement("div");
    box.id = "contactSheet"; box.className = "modal";
    box.innerHTML = `<div class="modal-box"><h3 class="sec-title">상담 내용을 남길까요?</h3>
      <p class="hint">${esc(pc.name)} · ${esc(pc.kind)}</p>
      <div class="ev-memo"><input type="text" id="pcText" maxlength="300" placeholder="추가로 남길 내용 (없으면 비워 두세요)" autocomplete="off"><button type="button" class="mic-btn" id="pcMic" aria-label="음성으로 입력"><svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="9" y="3" width="6" height="11" rx="3"/><path d="M5 11a7 7 0 0 0 14 0M12 18v3"/></svg></button></div>
      <div class="file-actions"><button type="button" class="secondary" id="pcNo">안 남기기</button><button type="button" class="primary" id="pcYes">남기기</button></div></div>`;
    document.body.appendChild(box);
    const close = () => { micStop(); box.remove(); try { localStorage.removeItem(PC_KEY); } catch (e) {} };
    $("pcMic").addEventListener("click", (e) => dictate($("pcText"), e.currentTarget));
    $("pcNo").addEventListener("click", close);
    $("pcYes").addEventListener("click", async () => {
      micStop();
      const note = $("pcText").value.trim();
      const text = `[${pc.kind}]${note ? " " + note : pc.kind === "전화" ? " 통화" : " 문자 발송"}`;
      const btn = $("pcYes"); btn.disabled = true; btn.textContent = "저장 중…";
      try {
        const list = await saveConsults(pc.fc, pc.name, [{ text, when: new Date() }], []);
        if (current && current.cust && current.cust.name === pc.name) current.extra = { ...(current.extra || {}), consults: list };
        toast("상담기록에 남겼습니다.");
        close();
      } catch (e) { btn.disabled = false; btn.textContent = "남기기"; saveError(e); }
    });
  }
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "hidden") { pcAway = !!localStorage.getItem(PC_KEY); return; }
    if (pcAway) { pcAway = false; setTimeout(askContactNote, 400); } // 전화·문자 앱에 다녀온 뒤에만 묻기
  });
  // 고객 검색칸의 마이크: 말하면 글자로 들어가고, 끝나면(완료·칸 터치) 바로 검색
  document.addEventListener("click", (ev) => {
    const b = ev.target.closest && ev.target.closest("#topMic, #pageMic");
    if (!b) return;
    const top = b.id === "topMic", inp = $(top ? "q" : "q2"), form = $(top ? "topSearch" : "pageSearch");
    dictate(inp, b, () => { if (inp.value.trim()) form.requestSubmit(); });
  });
  async function saveMemo(it, i, j, btn) {
    micStop();
    const c = it.custs[j], inp = $("memo" + i + "-" + j);
    const text = inp ? inp.value.trim() : "";
    if (!c || !text) return;
    const fc = memoFcOf(c);
    btn.disabled = true; btn.textContent = "저장 중…"; inp.disabled = true;
    try {
      const list = await saveConsults(fc, c.name, [{ text, when: new Date() }], []);
      c.extra = { ...(c.extra || {}), consults: list };
      inp.value = "";
      fillCust(it, c, i, j);
      btn.textContent = "저장됨";
      toast("상담기록에 저장했습니다.");
      setTimeout(() => { if (btn.textContent === "저장됨") btn.textContent = "저장"; }, 1800);
    } catch (e) {
      btn.disabled = false; btn.textContent = "저장";
      saveError(e);
    } finally { inp.disabled = false; }
  }
  const putBtn = (i, j, kind) => `<button type="button" class="put-btn" data-put="${i}-${j}-${kind}" aria-label="일정 장소에 입력">입력</button>`;
  // multi: 동명이인 여러 명 → '선택' / pinned: 동명이인 중 한 명으로 묶인 상태 → '다시 선택'(되돌리기)
  // v429(2026-11-18 사용자 지시): 일정의 고객 카드도 영업과 같은 상담기록 줄(최근 한 줄 + 수정·삭제·더보기, 고객 정보 칸을 터치하면 추가칸 + 최근 6개월 목록)
  function smemoOfCust(c) {
    const fc = memoFcOf(c), list = (c.extra && c.extra.consults) || [];
    SMEMO_C.set(fc + "|" + c.name, c);
    return smemoHtml(c.name, fc, isRO(fc), list, 5);
  }
  function custBlockHtml(c, i, j, multi, pinned) {
    const phone = (c.entries.find((x) => x.phone) || {}).phone || "";
    const addr = currentValue(c, ["주소"]);
    const types = [...new Set(c.entries.map((x) => x.type))];
    const g = genderOf(c);
    return `<div class="ev-cust">
        <div class="ev-namerow"><button type="button" class="ev-name" data-open="${i}-${j}"><span class="cname ${gradeClass(gradeOf(c))}">${esc(c.name)}</span>
          ${types.map((t) => `<span class="tag t-${esc(t)}">${esc(typeLabel(t))}</span>`).join("")}<span class="chev">›</span></button>
        ${multi ? `<button type="button" class="pick-btn" data-pick="${i}-${j}">선택</button>` : ""}${pinned ? `<button type="button" class="pick-btn" data-unpick="${i}">다시 선택</button>` : ""}</div>
        <dl class="ev-info">
          ${c.birth ? `<dt>생년월일</dt><dd>${esc(c.birth)}${g ? ` (${esc(g)})` : ""}</dd>` : ""}
          ${phone ? (() => { const num = (String(phone).match(MOBILE_RE) || [phone])[0], dg = String(num).replace(/[^\d+]/g, ""); return `<dt>연락처</dt><dd class="put-row" data-cn="${esc(c.name)}" data-cfc="${esc(memoFcOf(c))}"><span><a class="nowrap" href="tel:${esc(dg)}">${esc(phone)}</a></span><span class="ev-acts"><a class="ev-act" href="tel:${esc(dg)}">전화</a><a class="ev-act" href="${esc(smsHref(dg, ""))}">문자</a></span></dd>`; })() : ""}
          ${addr ? `<dt>주소</dt><dd class="put-row"><span>${esc(addr)}</span><span class="ev-acts">${mapBtn(addr)}${putBtn(i, j, "addr")}</span></dd>` : ""}
          <dt id="mpl${i}-${j}" hidden>미팅장소</dt><dd id="mp${i}-${j}" class="put-row" hidden></dd>
        </dl>
        <div id="cs${i}-${j}">${smemoOfCust(c)}</div></div>`;
  }
  function dayCardHtml(it, i) {
    const e = it.e;
    const linked = it.pinned && it.custs && it.custs.length && !it.ignored; // v447: [선택]을 눌러 고객 한 명과 묶은 일정만 왼쪽 파란 선 // v439: 고객과 묶인 일정은 왼쪽에 색 띠
    return `<div class="card ev-card${linked ? " linked" + (it.custs.length > 1 ? " multi" : "") : ""}" id="ev${i}"><div class="ev-top">${calDot(e)}<div class="ev-head"><span class="ev-time">${esc(timeText(e))}</span></div>
        <button type="button" class="ev-edit ev-totask" data-evtask="${i}" title="이 일정을 구글 할 일로 옮깁니다">할 일로</button>${e.allDay ? "" : salesOf(e) ? '<button type="button" class="ev-edit ev-sales" data-evunsales title="누르면 비영업일정으로 바꿉니다">영업</button>' : `<button type="button" class="ev-edit ev-confirm" data-evconfirm="${i}" title="누르면 영업일정으로 바꿉니다">비영업</button>`}<span class="ev-edit-acts" id="evActs${i}" hidden><button type="button" class="ev-edit" data-f-cancel>취소</button><button type="button" class="ev-edit ev-done" data-f-save>완료</button></span></div>
      <div class="ev-line2"><span class="ev-title-l"><span class="ev-title">${esc(e.title)}</span> <span class="ev-cal">- ${esc(e.cal.name)}${e.author ? " · 작성 " + esc(e.author) : ""}</span></span>${it.all.length ? (it.ignored
        ? `<button type="button" class="ev-edit" data-unignore="${i}">되돌리기</button>`
        : it.pinned ? "" : `<button type="button" class="ev-edit" data-ignore="${i}">해당 없음</button>`) : ""}</div>
      <div class="ev-btnrow" id="evBtns${i}" hidden><button type="button" class="ev-edit" data-evedit="${i}">수정</button><button type="button" class="ev-edit ev-del" data-evdel="${i}">삭제</button></div>
      ${it.ignored ? `<p class="hint ev-hid">검색된 고객 ${it.all.length}명을 숨겼어요</p>` : ""}
      <p class="hint ev-loc" id="loc${i}" ${e.location ? "" : "hidden"}><svg class="pin" viewBox="0 0 24 24" width="14" height="14" aria-hidden="true" style="vertical-align:-2px;margin-right:4px;fill:none;stroke:currentColor;stroke-width:1.5;stroke-linecap:round;stroke-linejoin:round"><path d="M20 10c0 6-8 12-8 12s-8-6-8-12a8 8 0 0 1 16 0z"/><circle cx="12" cy="10" r="3"/></svg>${esc(e.location)} ${mapBtn(e.location)}</p>
      <div id="evForm${i}"></div>
      ${it.custs.map((c, j) => custBlockHtml(c, i, j, !it.pinned, it.pinned)).join("")}</div>`;
  }
  function bindDayCard(it, i, req) {
    it.custs.forEach((c, j) => {
      const b = document.querySelector(`[data-open="${i}-${j}"]`);
      if (b) b.addEventListener("click", () => openCustomer(c));
    });
    const card = $("ev" + i);
    card.addEventListener("click", (ev) => {
      const put = ev.target.closest("[data-put]");
      if (put) { const [, j, kind] = put.dataset.put.split("-"); putLocation(it, i, +j, kind, put); }
      const pk = ev.target.closest("[data-pick]");
      if (pk) pickCustomer(it, it.custs[+pk.dataset.pick.split("-")[1]], pk);
      const up = ev.target.closest("[data-unpick]");
      if (up) unpickCustomer(it, up);
      const ig = ev.target.closest("[data-ignore]");
      if (ig) setIgnore(it, true, ig);
      const ui = ev.target.closest("[data-unignore]");
      if (ui) setIgnore(it, false, ui);
      const mc = ev.target.closest("[data-mic]");
      if (mc) { const [ci, cj] = mc.dataset.mic.split("-").map(Number); toggleMic(ci, cj, mc); }
      const mm = ev.target.closest("[data-memo]");
      if (mm) { const [mi, mj] = mm.dataset.memo.split("-").map(Number); saveMemo(it, mi, mj, mm); }
    });
    card.addEventListener("input", (ev) => {
      const m = ev.target.closest("[data-memoin]");
      if (!m) return;
      const b = card.querySelector(`[data-memo="${m.dataset.memoin}"]`);
      if (b && b.textContent === "저장") b.disabled = !m.value.trim();
    });
    card.addEventListener("keydown", (ev) => {
      const m = ev.target.closest && ev.target.closest("[data-memoin]");
      if (m && ev.key === "Enter" && !ev.isComposing) { ev.preventDefault(); const b = card.querySelector(`[data-memo="${m.dataset.memoin}"]`); if (b && !b.disabled) b.click(); }
    });
    const cf = card.querySelector("[data-evconfirm]");
    if (cf) cf.addEventListener("click", async () => { cf.disabled = true; cf.textContent = "저장 중"; await setSalesMark(it.e, "1"); renderSchedule(); }); // 영업일정으로 확정
    const us = card.querySelector("[data-evunsales]");
    if (us) us.addEventListener("click", async () => { us.disabled = true; us.textContent = "저장 중"; await setSalesMark(it.e, "0"); renderSchedule(); }); // 다시 비영업일정으로
    card.querySelector("[data-evedit]").addEventListener("click", () => toggleEventForm(it, i));
    card.querySelector("[data-evdel]").addEventListener("click", (ev) => deleteEvent(it, ev.currentTarget));
    const tk = card.querySelector("[data-evtask]");
    if (tk) tk.addEventListener("click", () => eventToTask(it, tk));
  }
  // v412(2026-11-18 사용자 지시): 종일 일정 → 구글 할 일. 그 날짜가 마감인 할 일(기본 목록)을 만들고, 만들어지면 캘린더 일정은 지움
  // (지운 일정은 카드의 '원위치'로 되살릴 수 있음 — 새로 고치기 전까지. 만든 할 일은 화면 위 할 일 카드에서 지우거나 완료)
  async function eventToTask(it, btn) {
    const e = it.e, card = btn.closest(".ev-card");
    if (!confirm(`'${e.title}' ${e.allDay ? "종일 " : ""}일정을 구글 할 일로 옮길까요?` + String.fromCharCode(10) + `${ymdStr(e.start)} 마감인 할 일을 만들고, 이 일정은 캘린더에서 지웁니다.`)) return;
    btn.disabled = true; btn.textContent = "옮기는 중";
    let made = false;
    try {
      if (!storedTasksToken()) await requestTasksToken(""); // 할 일 연결이 안 돼 있으면 한 번 연결
      const notes = [e.allDay ? "" : `시간: ${hm(e.start)}~${hm(e.end)}`, e.location ? "장소: " + e.location : "", e.desc ? e.desc.replace(/<[^>]+>/g, " ").trim() : "", `(캘린더 '${e.cal.name}' ${e.allDay ? "종일 " : ""}일정에서 옮김)`].filter(Boolean).join("\n");
      await gtasks("/lists/%40default/tasks", { method: "POST", body: JSON.stringify({ title: e.title, notes, due: ymdStr(e.start) + "T00:00:00.000Z" }) });
      made = true;
      await ensureGWrite();
      await gcalDelete(e.calId, e.eventId);
      if (card) card.classList.add("deleted");
      btn.textContent = "할 일로 옮김";
      const db = card && card.querySelector("[data-evdel]"); if (db) { db.dataset.restore = "1"; db.textContent = "원위치"; } // 일정만 되살리기(할 일은 그대로)
      toast("할 일로 옮겼습니다. 일정은 '원위치'로 되살릴 수 있습니다.");
      const tb = $("taskBox"); if (tb && schedDate) loadDayTasks(tb, schedDate, schedReq); // 위쪽 할 일 카드 새로
    } catch (err) {
      btn.disabled = false; btn.textContent = "할 일로";
      const m = err.message === "TASKS_AUTH" ? "할 일 연결이 필요합니다. 다시 눌러 주세요." : err.message === "TASKS_API_OFF" ? "구글 할 일(Tasks) 기능이 꺼져 있어 만들 수 없습니다." : err.message === "GOOGLE_AUTH" ? "구글 연결이 끊겼습니다. 다시 눌러 주세요." : err.message;
      toast(made ? "할 일은 만들었지만 일정을 지우지 못했습니다: " + m : m);
      if (made) { btn.disabled = true; btn.textContent = "할 일 만듦"; }
    }
  }
  function fillCust(it, c, i, j) {
    const x = c.extra || {};
    const mp = meetingOf(c, x);
    if (mp) {
      const dd = $(`mp${i}-${j}`);
      dd.innerHTML = `<span>${esc(mp)}</span><span class="ev-acts">${mapBtn(mp)}${putBtn(i, j, "mp")}</span>`;
      dd.hidden = false; $(`mpl${i}-${j}`).hidden = false;
    }
    const box = $(`cs${i}-${j}`); // 상담기록 줄: 불러온 목록으로 다시 그림(쓰던 글이 있으면 그대로 둠)
    if (box) {
      const old = box.querySelector("[data-smin]"), v = old ? old.value : "";
      box.innerHTML = smemoOfCust(c);
      if (v) { const ni = box.querySelector("[data-smin]"); if (ni) { ni.value = v; ni.dispatchEvent(new Event("input", { bubbles: true })); } }
    }
  }

  // ---- v559 캘린더 일정 ↔ 고객 연결: 공통 모듈(calcust.js, 규칙 app\_공통\캘린더고객연결.md)로 숨은 칸과 기록 파일(고객\<FC>\일정매칭)에 남김 ----
  if (window.CalCust) CalCust.config({ token: () => storedGToken() || gTokenMem || "" });
  const calFailed = new Set(); // 이 세션에서 캘린더 쓰기가 막힌 일정(다시 시도하지 않음)
  let calFolderMap = null;
  async function calFolders() { if (!calFolderMap) { const P = await readPerms(); calFolderMap = { ...((P && P.fcFolders) || {}), ...fcFolderIds }; } return calFolderMap; }
  function calInfo(it, o) {
    const e = it.e, c = o.c || null;
    // 담당 FC = 일정을 만든 계정(creator)을 권한 파일 names로 바꾼 이름(사용자 지시: 붙임말은 보기용, 실질 담당은 소유자). 붙임말 FC는 참고 칸만
    return { fc: CalCust.fcOfCreator(e.creator, permNamesCache || {}), creator: e.creator || "", tagFc: CalCust.fcOfTitle(e.title, permPrefixCache || {}), custId: c ? custIdOf(c) : "", custFc: c ? primaryFc(c.entries || []) : "", kind: CalCust.kindOf(e.title), how: o.how };
  }
  async function calWrite(it, o, interactive) { // 숨은 칸 쓰기. 바뀌었으면 e.priv·e.updated를 새 값으로
    const e = it.e, inf = calInfo(it, o);
    if (calFailed.has(e.id) && !interactive) return { inf, ok: false };
    try {
      if (interactive) await ensureGWrite(); else if (!(storedGToken() || gTokenMem)) return { inf, ok: false };
      const r = await CalCust.write({ calId: e.calId, eventId: e.eventId, custId: inf.custId, ignore: !!o.ignore, clear: !!o.clear, how: inf.how, fc: inf.fc, custFc: inf.custFc, kind: inf.kind, priv: e.priv });
      if (r) { e.priv = (r.extendedProperties || {}).private || {}; e.updated = r.updated || e.updated; if (e.raw) e.raw.updated = e.updated; }
      return { inf, ok: true };
    } catch (er) { calFailed.add(e.id); console.warn("캘린더 고객 연결을 일정에 쓰지 못함(기록 파일만)", er.message); return { inf, ok: false }; }
  }
  const calRecSig = new Map(); // 일정ID → 마지막으로 기록 파일에 넘긴 내용(같으면 다시 안 넘김)
  async function calRecord(list) { // list = [{it, inf}]
    try {
      const fm = await calFolders(), mine = await myFcName(), acc = me(), groups = {};
      list.forEach(({ it, inf }) => {
        const e = it.e, row = CalCust.rowOf({ calId: e.calId, eventId: e.eventId, calName: e.cal && e.cal.name, ev: e.raw || {}, title: e.title, fc: inf.fc, creator: inf.creator, tagFc: inf.tagFc, custId: inf.custId, custFc: inf.custFc, kind: inf.kind, how: inf.how, updated: e.updated, account: acc });
        const sig = JSON.stringify([row.날짜, row.시작, row.끝, row.제목, row.FC, row.만든계정, row.붙임말FC, row.고객ID, row.고객FC, row.종류, row.방식, row.일정수정시각]);
        if (calRecSig.get(row.일정ID) === sig) return;
        const dest = (inf.fc && fm[inf.fc]) || fm[mine] || ""; // 일정 FC 폴더, FC 미정이면 내 FC 폴더
        if (!dest) return;
        (groups[dest] = groups[dest] || []).push([row, sig, inf.fc && fm[inf.fc] ? fm[mine] : ""]);
      });
      for (const [dest, rows] of Object.entries(groups)) {
        const put = (fid) => CalCust.recordPut(fid, rows.map((x) => x[0]), acc);
        try { await put(dest); } catch (er) { const alt = rows[0][2]; if (alt && alt !== dest) await put(alt); else throw er; } // 그 FC 폴더에 못 쓰면 내 FC 폴더
        rows.forEach(([row, sig]) => calRecSig.set(row.일정ID, sig));
      }
    } catch (er) { console.warn("일정매칭 기록 파일을 쓰지 못함", er.message); }
  }
  async function calLinkSave(it, o) { const w = await calWrite(it, o, true); calRecord([{ it, inf: w.inf }]); return w.ok; }
  async function calAutoLink(items, req) { // 화면에 보인 일정: 자동(한 명)·선택·무시를 숨은 칸과 기록 파일에 맞춤
    if (!window.CalCust || EMBED) return;
    try { if (!permPrefixCache || !permNamesCache) await readPerms(); } catch (e) { /* 이름표 없이 진행(담당 미정) */ }
    const done = [];
    for (const it of items) {
      if (req !== schedReq) return;
      let o = null;
      if (it.ignored) o = { ignore: true, how: "무시" };
      else if (it.pinned && it.custs.length === 1 && custIdOf(it.custs[0])) o = { c: it.custs[0], how: "선택" };
      else if (!it.pinned && it.all.length === 1 && it.custs.length === 1 && custIdOf(it.custs[0])) o = { c: it.custs[0], how: "자동" };
      if (!o) continue;
      const w = await calWrite(it, o, false);
      done.push({ it, inf: w.inf });
    }
    if (done.length) calRecord(done);
  }
  // 동명이인 '선택': 이 일정을 그 고객으로 묶음 → 구글 일정의 비공개 속성에 저장(모든 기기·사용자 공유) + 이 기기에도 기록
  async function pickCustomer(it, c, btn) {
    if (!c) return;
    const key = custKey(c);
    btn.disabled = true; btn.textContent = "저장 중";
    savePick(it.e.id, key);
    try {
      await ensureGWrite();
      if (custIdOf(c)) { if (!(await calLinkSave(it, { c, how: "선택" }))) throw new Error("캘린더에 쓰지 못함"); } // v559: 공통 모듈로 고객ID·FC·종류 함께
      else await gcalPatch(it.e.calId, it.e.eventId, { extendedProperties: { private: { shlifeCust: key } } }); // 고객ID 없는 옛 고객
      it.e.pick = key;
    } catch (e) {
      console.warn("pick saved on this device only", e.message); // 수정 권한이 없는 캘린더 등: 이 기기에만
    }
    renderSchedule();
  }

  // '무시': 검색된 고객이 이 일정과 상관없을 때 모두 숨김 / '되돌리기': 다시 보여 줌 (구글 일정에 저장해 모든 기기 공유)
  async function setIgnore(it, on, btn) {
    btn.disabled = true; btn.textContent = on ? "숨기는 중" : "되돌리는 중";
    savePick(it.e.id, on ? IGNORE_KEY : "");
    try {
      await ensureGWrite();
      if (!(await calLinkSave(it, on ? { ignore: true, how: "무시" } : { clear: true, how: "해제" }))) throw new Error("캘린더에 쓰지 못함");
      it.e.pick = on ? IGNORE_KEY : "";
    } catch (e) {
      toast("구글 일정에 저장하지 못해 이 기기에만 적용했습니다.");
      if (!on) it.e.pick = "";
    }
    renderSchedule();
  }

  // '다시 선택': 이 일정의 고객 지정을 풀고 동명이인을 모두 다시 보여 줌 (일정 장소는 그대로)
  async function unpickCustomer(it, btn) {
    btn.disabled = true; btn.textContent = "푸는 중";
    savePick(it.e.id, "");
    if (it.e.pick) {
      try {
        await ensureGWrite();
        if (!(await calLinkSave(it, { clear: true, how: "해제" }))) throw new Error("캘린더에 쓰지 못함");
      } catch (e) { console.warn("unpick on this device only", e.message); }
      it.e.pick = "";
    }
    renderSchedule();
  }

  // 주소/미팅장소 '입력' → 그 일정의 장소(위치)에 저장 + 동명이인이면 이 고객으로 지정
  async function putLocation(it, i, j, kind, btn) {
    const c = it.custs[j];
    const text = kind === "mp" ? meetingOf(c) : currentValue(c, ["주소"]);
    if (!text) return;
    btn.disabled = true; btn.textContent = "저장 중";
    try {
      await ensureGWrite();
      await gcalPatch(it.e.calId, it.e.eventId, { location: text });
      it.e.location = text;
      const loc = $("loc" + i);
      loc.innerHTML = `<svg class="pin" viewBox="0 0 24 24" width="14" height="14" aria-hidden="true" style="vertical-align:-2px;margin-right:4px;fill:none;stroke:currentColor;stroke-width:1.5;stroke-linecap:round;stroke-linejoin:round"><path d="M20 10c0 6-8 12-8 12s-8-6-8-12a8 8 0 0 1 16 0z"/><circle cx="12" cy="10" r="3"/></svg>${esc(text)} ${mapBtn(text)}`; loc.hidden = false;
      toast("일정 장소에 입력했습니다.");
      if (!it.pinned) { await pickCustomer(it, c, btn); return; }
      btn.textContent = "입력됨";
    } catch (e) {
      btn.disabled = false; btn.textContent = "입력";
      toast(e.message === "GOOGLE_AUTH" ? "구글 연결이 끊겼습니다. 다시 눌러 주세요." : e.message);
    }
  }

  async function deleteEvent(it, btn) {
    const e = it.e, card = btn.closest(".ev-card");
    const authMsg = (err) => (err.message === "GOOGLE_AUTH" ? "구글 연결이 끊겼습니다. 다시 눌러 주세요." : err.message);
    if (btn.dataset.restore) { // 삭제한 뒤의 '원위치' 버튼: 구글 캘린더에서 지운 일정을 되살림 (새로고침하기 전까지)
      btn.disabled = true; btn.textContent = "되돌리는 중";
      try {
        await ensureGWrite();
        await gcalPatch(e.calId, e.eventId, { status: "confirmed" });
        delete btn.dataset.restore; card.classList.remove("deleted"); btn.textContent = "삭제";
        toast("일정을 원위치했습니다.");
      } catch (err) { btn.textContent = "원위치"; toast(authMsg(err)); }
      btn.disabled = false;
      return;
    }
    if (!confirm(`'${e.title}' 일정을 구글 캘린더에서 정말 삭제할까요?` + String.fromCharCode(10) + "삭제 뒤에는 '원위치' 버튼으로 되돌릴 수 있지만, 화면을 새로 고치면 더는 되돌릴 수 없습니다.")) return;
    btn.disabled = true; btn.textContent = "삭제 중";
    try {
      await ensureGWrite();
      await gcalDelete(e.calId, e.eventId);
      card.classList.add("deleted"); btn.dataset.restore = "1"; btn.textContent = "원위치"; // 삭제 버튼 → 원위치 버튼
      toast("일정을 삭제했습니다. '원위치'를 누르면 되돌립니다.");
    } catch (err) { btn.textContent = "삭제"; toast(authMsg(err)); }
    btn.disabled = false;
  }

  // 일정 추가 (일정 화면 위 [추가]): 제목 · 날짜 · 시간(또는 종일) · 캘린더 · 장소를 넣어 구글 캘린더에 새 일정 만들기
  function closeSchedAdd() { const b = $("schedAddBox"); if (b) { b.innerHTML = ""; b.hidden = true; } }
  function openSchedAdd() {
    const box = $("schedAddBox"); if (!box) return;
    if (!box.hidden) return closeSchedAdd();
    const cals = schedCals || [];
    if (!cals.length) return toast("일정을 불러온 뒤에 추가할 수 있습니다.");
    const team = (GCFG && GCFG.teamCalendar) || "";
    const def = (cals.find((c) => c.name === team) || cals[0]).id;
    const now = new Date(), h0 = sameDay(schedDate, now) ? Math.min(now.getHours() + 1, 22) : 9;
    // v430(2026-11-18 사용자 지시): 맨 위에서 [일정] / [할 일]을 고르고, 고른 것에 따라 입력란이 달라짐 (일정 = 캘린더·시간·장소 / 할 일 = 마감일·메모)
    box.innerHTML = `<div class="card ev-add"><div class="ev-form">
        <div class="prog-bar ev-kind" role="group" aria-label="추가할 종류"><button type="button" class="on" data-a-kind="ev" aria-pressed="true">일정</button><button type="button" data-a-kind="task" aria-pressed="false">할 일</button></div>
        <label class="fld"><span>제목</span><input data-a="title" autocomplete="off" placeholder="예: 김미례 고객 AP"></label>
        <div data-a-ev><label class="fld"><span>캘린더</span><select data-a="cal">${cals.map((c) => `<option value="${esc(c.id)}" ${c.id === def ? "selected" : ""}>${esc(c.name)}</option>`).join("")}</select></label>
        </div>
        <label class="fld"><span data-a-datel>날짜</span><input data-a="date" type="text" data-dp="date" value="${ymdStr(schedDate)}"></label>
        <div data-a-ev><label class="ev-allday"><input type="checkbox" class="cb" data-a="all"> 종일</label>
        <div class="fld-2" data-a-times><label class="fld"><span>시작</span><input data-a="start" type="time" value="${pad2(h0)}:00"></label>
          <label class="fld"><span>종료</span><input data-a="end" type="time" value="${pad2(Math.min(h0 + 1, 23))}:00"></label></div>
        <label class="fld"><span>장소</span><input data-a="loc" autocomplete="off"></label>
        ${isAdmin() ? `<label class="fld"><span>작성자</span><select data-a="author"><option value="${esc(String(account().username).toLowerCase())}" selected>${esc((permNamesCache || {})[String(account().username).toLowerCase()] || String(account().username).split("@")[0])} (관리자)</option></select></label>` : ""}</div>
        <div data-a-task hidden>${isAdmin() ? `<label class="fld"><span>사용자</span><select data-a="tauthor">${taskSelfOpt()}</select></label>` : ""}<label class="fld"><span>메모</span><textarea data-a="notes" rows="3" placeholder="할 일 메모 (없으면 비워 두세요)"></textarea></label></div>
        <div class="file-actions"><button type="button" class="secondary" data-a-cancel>취소</button><button type="button" class="primary" data-a-save>등록</button></div>
      </div></div>`;
    box.hidden = false;
    const q = (k) => box.querySelector(`[data-a="${k}"]`);
    q("all").addEventListener("change", () => { box.querySelector("[data-a-times]").hidden = q("all").checked; });
    let kind = "ev";
    box.querySelectorAll("[data-a-kind]").forEach((b) => b.addEventListener("click", () => { // 일정 ↔ 할 일: 입력란 바꾸기
      kind = b.dataset.aKind;
      box.querySelectorAll("[data-a-kind]").forEach((x) => { const on = x === b; x.classList.toggle("on", on); x.setAttribute("aria-pressed", String(on)); });
      box.querySelectorAll("[data-a-ev]").forEach((x) => { x.hidden = kind !== "ev"; });
      box.querySelectorAll("[data-a-task]").forEach((x) => { x.hidden = kind !== "task"; });
      box.querySelector("[data-a-datel]").textContent = kind === "task" ? "마감일" : "날짜";
      q("title").placeholder = kind === "task" ? "예: 청약서 보내기" : "예: 김미례 고객 AP";
      box.querySelector("[data-a-save]").textContent = kind === "task" ? "할 일 등록" : "등록";
    }));
    box.querySelector("[data-a-cancel]").addEventListener("click", closeSchedAdd);
    q("title").focus();
    if (q("tauthor")) fillTaskUsers(q("tauthor"), ""); // v431: 할 일의 사용자(관리자)
    if (q("author")) { // v382(2026-11-17 사용자 지시): 관리자는 작성자를 사용자 중에서 고름(기본 = 관리자 본인). v436: 내가 쓸 수 있는 캘린더의 사람만
      Promise.all([readPerms(), writablePeople()]).then(([pm, wp]) => {
        const sel = q("author"); if (!sel) return;
        const my = String(account().username).toLowerCase(), nm = (em) => (pm.names || {})[em] || em.split("@")[0];
        Object.keys(pm.users || {}).filter((em) => em !== my && !/@email\.pending$/.test(em) && wp.has(String(em).toLowerCase())).sort((a, b) => nm(a).localeCompare(nm(b), "ko"))
          .forEach((em) => sel.insertAdjacentHTML("beforeend", `<option value="${esc(em)}">${esc(nm(em))}</option>`));
      }).catch(() => {});
    }
    box.querySelector("[data-a-save]").addEventListener("click", async (ev) => {
      const btn = ev.currentTarget, title = q("title").value.trim(), date = q("date").value, all = q("all").checked;
      if (!title) return toast("제목을 입력해 주세요.");
      if (!date) return toast(kind === "task" ? "마감일을 입력해 주세요." : "날짜를 입력해 주세요.");
      if (kind === "task") { // 구글 할 일(기본 목록)에 그 날짜가 마감인 할 일 만들기
        btn.disabled = true; btn.textContent = "등록 중…";
        try {
          if (!storedTasksToken()) await requestTasksToken(""); // 할 일 연결이 안 돼 있으면 한 번 연결
          const notes = taskNotesWith(q("notes").value.trim(), taskOwnerPick(q("tauthor")));
          await gtasks("/lists/%40default/tasks", { method: "POST", body: JSON.stringify({ title, ...(notes ? { notes } : {}), due: date + "T00:00:00.000Z" }) });
          toast("할 일을 등록했습니다: " + title);
          closeSchedAdd();
          schedDate = dayStart(ymdToDate(date)); schedMode = "day"; renderSchedule();
        } catch (err) {
          btn.disabled = false; btn.textContent = "할 일 등록";
          toast(err.message === "TASKS_AUTH" ? "할 일 연결이 필요합니다. 다시 눌러 주세요." : err.message === "TASKS_API_OFF" ? "구글 할 일(Tasks) 기능이 꺼져 있어 만들 수 없습니다." : err.message);
        }
        return;
      }
      const body = { summary: title };
      if (q("loc").value.trim()) body.location = q("loc").value.trim();
      if (all) { body.start = { date }; body.end = { date: ymdStr(addDays(ymdToDate(date), 1)) }; }
      else {
        const st = q("start").value, en = q("end").value;
        if (!st || !en) return toast("시간을 입력해 주세요.");
        if (en <= st) return toast("종료 시간이 시작보다 빠릅니다.");
        const tz = (Intl.DateTimeFormat().resolvedOptions() || {}).timeZone || "Asia/Seoul";
        body.start = { dateTime: `${date}T${st}:00`, timeZone: tz }; body.end = { dateTime: `${date}T${en}:00`, timeZone: tz };
      }
      if (q("author") && q("author").value) { const o = q("author").selectedOptions[0]; body.extendedProperties = { shared: { shlifeAuthor: q("author").value, shlifeAuthorName: o.textContent.replace(/ \(관리자\)$/, "") } }; }
      btn.disabled = true; btn.textContent = "등록 중…";
      try {
        await ensureGWrite();
        await gcalInsert(q("cal").value, body);
        toast("일정을 등록했습니다: " + title);
        closeSchedAdd();
        schedDate = dayStart(ymdToDate(date)); schedMode = "day"; renderSchedule();
      } catch (err) {
        btn.disabled = false; btn.textContent = "등록";
        toast(err.message === "GOOGLE_AUTH" ? "구글 연결이 끊겼습니다. 다시 눌러 주세요." : err.message);
      }
    });
  }
  // 일정 수정 모드: 제목 · 날짜 · 시간 · 장소 · 설명
  const ymdStr = (d) => `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
  // 일정 수정 모드: 이 일정만 '수정' 대신 '취소 · 완료'(맨 위 오른쪽), 다른 일정의 '수정'은 숨김
  function setEventEditing(i, on) {
    const box = $("evForm" + i), acts = $("evActs" + i);
    if (!on) { box.innerHTML = ""; const tl = $("ev" + i).querySelector(".ev-title-l"); if (tl && tl.dataset.orig != null) { tl.innerHTML = tl.dataset.orig; delete tl.dataset.orig; } } // 제목 입력칸을 원래 제목으로
    acts.hidden = !on;
    if (on) { const br = $("evBtns" + i); if (br) br.hidden = true; }
    document.querySelectorAll("#dayList [data-evedit], #dayList [data-evdel]").forEach((b) => { b.hidden = on; });
    $("ev" + i).classList.toggle("editing", on);
  }
  // v561: 날짜 칸은 공통 날짜 달력(datepick.js, app\_공통\날짜달력.md)이 칸 아래에 띄움(v558의 브라우저 기본 달력 대신)
  function toggleEventForm(it, i) {
    const box = $("evForm" + i);
    if (box.innerHTML) { setEventEditing(i, false); return; }
    const e = it.e;
    const endDay = e.allDay ? addDays(e.end, -1) : e.end;
    box.innerHTML = `<div class="ev-form">
        <label class="fld"><span>캘린더</span><select data-f="cal">${schedCals.filter((c) => (GCFG.calendars || []).includes(c.name) || c.id === e.calId).map((c) =>
          `<option value="${esc(c.id)}" ${c.id === e.calId ? "selected" : ""}>${esc(c.name)}</option>`).join("")}</select></label>
        <label class="fld"><span>날짜</span><input data-f="date" type="text" data-dp="date" data-dp-range="ev" value="${ymdStr(e.start)}"></label>
        <label class="fld chk"><span>종일</span><input data-f="allday" type="checkbox" ${e.allDay ? "checked" : ""}></label>
        <label class="fld" data-ad-only ${e.allDay ? "" : "hidden"}><span>종료 날짜</span><input data-f="endDate" type="text" data-dp="date" data-dp-range="ev" data-dp-end value="${ymdStr(endDay)}"></label>
        <div class="fld-2" data-tm-only ${e.allDay ? "hidden" : ""}><label class="fld"><span>시작</span><input data-f="start" type="time" value="${e.allDay ? "09:00" : hm(e.start)}"></label>
             <label class="fld"><span>종료</span><input data-f="end" type="time" value="${e.allDay ? "10:00" : hm(e.end)}"></label></div>
        <label class="fld"><span>장소</span><input data-f="location" value="${esc(e.location)}" autocomplete="off"></label>
        <label class="fld"><span>설명</span><textarea data-f="desc" rows="3">${esc(e.desc)}</textarea></label>
        ${isAdmin() ? `<label class="fld"><span>작성자</span><select data-f="author">${taskSelfOpt()}</select></label>` : ""}
      </div>`;
    { // v432(2026-11-18 사용자 지시): 관리자는 일정을 수정할 때 작성자도 바꿈(기본 = 지금 작성자, 없으면 본인)
      const asel = box.querySelector('[data-f="author"]');
      if (asel) { const me = asel.selectedOptions[0].textContent.replace(/ \(관리자\)$/, ""); fillTaskUsers(asel, e.author && e.author !== me ? e.author : ""); }
    }
    setEventEditing(i, true);
    { // v448(2026-11-19 사용자 지시): 종일 체크를 풀면 시간 입력, 체크하면 종료 날짜 입력
      const ck = box.querySelector('[data-f="allday"]');
      if (ck) ck.addEventListener("change", () => {
        box.querySelector("[data-ad-only]").hidden = !ck.checked; box.querySelector("[data-tm-only]").hidden = ck.checked;
        if (ck.checked) { const dEl = box.querySelector('[data-f="date"]'), nEl = box.querySelector('[data-f="endDate"]'); if (!nEl.value || nEl.value < dEl.value) nEl.value = dEl.value; }
      });
    }
    { // 수정 모드에 들어가면 카드 위의 일정 제목이 바로 수정 칸이 됨
      const tl = $("ev" + i).querySelector(".ev-title-l");
      if (tl) { tl.dataset.orig = tl.innerHTML; tl.innerHTML = `<input data-f="title" class="ev-title-in" value="${esc(e.title)}" autocomplete="off" aria-label="일정 제목">`; }
    }
    { // v380(2026-11-17 사용자 지시): 시작 시간을 바꾸면 길이(종료-시작)를 유지하도록 종료 시간도 자동 계산
      const sIn = box.querySelector('[data-f="start"]'), eIn = box.querySelector('[data-f="end"]');
      if (sIn && eIn) {
        const toMin = (v) => { const m = /^(\d{1,2}):(\d{2})$/.exec(v || ""); return m ? +m[1] * 60 + +m[2] : null; };
        let prev = toMin(sIn.value);
        sIn.addEventListener("change", () => {
          const ns = toMin(sIn.value), oe = toMin(eIn.value);
          if (ns == null) return;
          let dur = prev != null && oe != null && oe > prev ? oe - prev : 60;
          prev = ns;
          const ne = Math.min(ns + dur, 23 * 60 + 59);
          eIn.value = String(Math.floor(ne / 60)).padStart(2, "0") + ":" + String(ne % 60).padStart(2, "0");
        });
      }
    }
    const acts = $("evActs" + i);
    const val = (k) => ((box.querySelector(`[data-f="${k}"]`) || $("ev" + i).querySelector(`[data-f="${k}"]`)) || {}).value;
    acts.querySelector("[data-f-cancel]").onclick = () => setEventEditing(i, false);
    acts.querySelector("[data-f-save]").onclick = async (ev) => {
      const btn = ev.currentTarget;
      const patch = { summary: val("title").trim(), location: val("location").trim(), description: val("desc") };
      if (!val("date")) return toast("날짜를 입력해 주세요.");
      { // 작성자를 바꿨을 때만 일정의 공유 속성을 고침(본인으로 되돌리면 지움)
        const asel = box.querySelector('[data-f="author"]');
        if (asel && asel.value !== "__cur") {
          const nm = taskOwnerPick(asel);
          const meNm = asel.options[0].textContent.replace(/ \(관리자\)$/, "");
          if (nm !== (e.author && e.author !== meNm ? e.author : "")) patch.extendedProperties = { shared: nm ? { shlifeAuthor: asel.value, shlifeAuthorName: nm } : { shlifeAuthor: null, shlifeAuthorName: null } };
        }
      }
      const allDayNow = !!(box.querySelector('[data-f="allday"]') || {}).checked;
      if (allDayNow) {
        const endD = val("endDate") || val("date");
        if (endD < val("date")) return toast("종료 날짜가 시작보다 빠릅니다.");
        patch.start = { date: val("date"), dateTime: null, timeZone: null };
        patch.end = { date: ymdStr(addDays(ymdToDate(endD), 1)), dateTime: null, timeZone: null };
      } else {
        if (!val("start") || !val("end")) return toast("시간을 입력해 주세요.");
        const st = new Date(val("date") + "T" + val("start") + ":00");
        let en = new Date(val("date") + "T" + val("end") + ":00");
        if (en <= st) return toast("종료 시간이 시작보다 빠릅니다.");
        const tz = Intl.DateTimeFormat().resolvedOptions().timeZone || "Asia/Seoul";
        patch.start = { dateTime: st.toISOString(), timeZone: tz, date: null };
        patch.end = { dateTime: en.toISOString(), timeZone: tz, date: null };
      }
      btn.disabled = true; btn.textContent = "저장 중…";
      try {
        await ensureGWrite();
        await gcalPatch(e.calId, e.eventId, patch);
        // 캘린더를 바꿨으면 그 캘린더로 옮기기 (동명이인 지정은 새 일정 id로 이어서)
        const dest = val("cal");
        if (dest && dest !== e.calId) {
          const moved = await gcalMove(e.calId, e.eventId, dest);
          const pk = readPicks()[e.id];
          if (pk) { savePick(e.id, ""); savePick(dest + "|" + (moved.id || e.eventId), pk); }
        }
        toast("일정을 저장했습니다.");
        renderSchedule();
      } catch (err) {
        btn.disabled = false; btn.textContent = "완료";
        toast(err.message === "GOOGLE_AUTH" ? "구글 연결이 끊겼습니다. 다시 눌러 주세요." : err.message);
      }
    };
  }

  // 이동 거리·시간 (1팀 일정만): 회사 → 일정1 → 일정2 … → 회사
  //   장소: 미팅장소 → 고객 주소 → 일정 장소 → 일정 지역(제목)의 시청·군청
  async function renderTravel(items, req) {
    const timed = items.map((it, i) => ({ ...it, i })).filter((it) => isTeamEvent(it.e));
    if (!timed.length || !KAKAO.restKey) return;
    const company = appSettings.companyAddress ? await geocode(appSettings.companyAddress) : null;
    const pointOf = async (it) => {
      const tries = [];
      if (it.cust) {
        tries.push(meetingOf(it.cust, it.extra));
        tries.push(currentValue(it.cust, ["주소"]));
      }
      tries.push(it.e.location);
      for (const q of tries) { const p = q && await geocode(q); if (p) return { p }; }
      const city = (it.cust && it.cust.city) || titleRegion(it.e.title);
      const p = await geocode(cityHall(city));
      return p ? { p } : null;
    };
    const pts = [];
    for (const it of timed) pts.push(await pointOf(it));
    if (req !== schedReq) return;
    const show = (slot, html) => { const el = $("tr" + slot); if (el) { el.innerHTML = html; el.hidden = false; } };
    const legTxt = (r) => r ? `🚗 ${r.km.toFixed(1)}km · ${r.min}분` : '🚗 <span class="hint">경로를 계산할 수 없음</span>';
    // 회사 출발
    if (company && pts[0]) {
      const r = await route(company, pts[0].p);
      const leave = r ? new Date(timed[0].e.start - r.min * 60000) : null;
      show(timed[0].i, `<span class="tl">${leave ? `🏢 회사 출발 ${hm(leave)}` : "🏢 회사에서"}</span><span class="tr">${legTxt(r)}</span>`); // 왼쪽: 회사 출발 / 오른쪽: 🚗 거리·시간
    }
    for (let k = 1; k < timed.length; k++) {
      if (req !== schedReq) return;
      const a = pts[k - 1], b = pts[k];
      const r = a && b ? await route(a.p, b.p) : null;
      const gap = Math.round((timed[k].e.start - timed[k - 1].e.end) / 60000);
      const late = r && gap < r.min ? `<span class="warn" title="이동시간 부족 (여유 ${Math.max(gap, 0)}분)" aria-label="이동시간 부족">⚠</span> ` : ""; // 문구 없이 빨간 ⚠ 표시만, 맨 앞에
      show(timed[k].i, `<span class="tr">${late}${legTxt(r)}</span>`);
    }
    // 회사 복귀: 마지막 1팀 일정 뒤
    const last = timed[timed.length - 1], lp = pts[pts.length - 1];
    if (company && lp) {
      const r = await route(lp.p, company);
      const back = r ? new Date(+last.e.end + r.min * 60000) : null;
      show(last.i + 1, `<span class="tl">${back ? `🏢 회사 도착 ${hm(back)}` : ""}</span><span class="tr">${legTxt(r)}</span>`);
    }
  }


  // ---------- 관리자 설정: 구글 계정·캘린더 선택 ----------
  async function loadCalendarChoices() {
    const box = $("gCalBox");
    if (!box) return;
    if (!storedGToken() && !gTokenMem) {
      if (GD) return handleError(new Error("GOOGLE_AUTH")); // Google Drive 모드: 로그인 계정 = 캘린더 계정
      box.innerHTML = `<button type="button" class="secondary" id="gAdminConnect">구글 계정 연결</button>`;
      $("gAdminConnect").addEventListener("click", async () => {
        try { await requestGToken("select_account"); loadCalendarChoices(); } catch (e) { toast(e.message); }
      });
      return;
    }
    box.innerHTML = '<p class="hint">캘린더 목록을 불러오는 중…</p>';
    let cals;
    try { cals = await calendarList(); } catch (e) {
      if (e.message === "GOOGLE_AUTH") { gTokenMem = null; if (GD) return handleError(e); return loadCalendarChoices(); }
      box.innerHTML = ""; return handleError(e);
    }
    const primary = cals.find((c) => c.primary);
    const g = appSettings.google || {};
    const saved = (g.calendars || []).map((c) => c.id);
    const want = GCFG.calendars || [];
    const isOn = (c) => (saved.length ? saved.includes(c.id) : want.includes(c.name));
    box.dataset.account = primary ? primary.id : "";
    // Google Drive 모드에서는 계정을 바꾸면 Drive 연결도 바뀌므로 '다른 계정' 버튼 없음 (로그인 계정의 캘린더 사용)
    box.innerHTML = `<p class="hint">${GD ? "캘린더 계정(로그인 계정)" : "로그인한 계정"}: ${esc(primary ? primary.id : "-")}
        ${GD ? "" : '<button type="button" class="link-btn" id="gSwitch">다른 계정</button>'}</p>
      <p class="hint">관리자 일정 화면의 캘린더입니다. 다른 사용자는 '설정'에서 각자 고르며, 고르기 전에는 같은 이름의 캘린더가 기본으로 보입니다.</p>
      ${calPickerHtml(cals, isOn)}`;
    box._cals = cals;
    const sIds = salesArr(g.sales).length ? salesArr(g.sales).map((c) => c.id) : defaultSalesIds(cals);
    bindCalPicker(box, cals, sIds, Array.isArray(g.info) ? g.info.map((c) => c.id) : sIds);
    if ($("gSwitch")) $("gSwitch").addEventListener("click", async () => {
      try { await requestGToken("select_account"); loadCalendarChoices(); } catch (e) { toast(e.message); }
    });
    box.dispatchEvent(new Event("change", { bubbles: true }));
  }
  function readCalendarChoices() {
    const box = $("gCalBox");
    if (!box || !box._cals) return null;
    return { account: box.dataset.account || "", ...readCalPicker(box, box._cals) };
  }

  // ---------- 캘린더 고르기 (관리자 설정·사용자 설정 공통) ----------
  //   보여질 캘린더(여러 개) + 영업일정 캘린더(보여질 캘린더 중 여러 개: 약속 장소 사이 거리·예상 시간, 고객 정보)
  //   보여질 캘린더를 하나도 고르지 않으면 영업일정 캘린더도 고를 수 없음
  //   저장 모양: sales = [{ id, name }]  (예전 { id, name } 하나도 읽음)
  const IGNORE_KEY = "__ignore__"; // 일정에 저장하는 고객 지정 값: 이 일정은 검색된 고객을 모두 숨김
  const salesArr = (s) => (Array.isArray(s) ? s : s && s.id ? [s] : []);
  // 한 줄에 캘린더 하나: 캘린더 | Show(보여질 캘린더) | 일정(거리·예상 시간) | 정보(고객 정보)
  //   Show가 꺼지면 일정·정보 비활성. 일정을 켜면 정보도 켜짐, 정보를 끄면 일정도 꺼짐
  function calPickerHtml(cals, isOn) {
    return `<div class="cal-pick cp4">
        <div class="cp-head"><span>캘린더</span><span>Show</span><span>일정</span><span>정보</span></div>
        ${cals.map((c, i) => `<div class="cp-row">
          <span class="cp-name">${calDot({ cal: c })}${esc(c.name)}</span>
          <label class="cp-cell" aria-label="${esc(c.name)} 보이기"><input type="checkbox" data-pcal="${i}" ${isOn(c) ? "checked" : ""}></label>
          <label class="cp-cell cp-right" aria-label="${esc(c.name)} 영업일정"><input type="checkbox" data-scal="${esc(c.id)}"></label>
          <label class="cp-cell" aria-label="${esc(c.name)} 고객 정보"><input type="checkbox" data-ical="${esc(c.id)}"></label>
        </div>`).join("")}
      </div>`;
  }
  function bindCalPicker(root, cals, salesIds, infoIds) {
    const want = new Set(salesIds || []);
    const wantInfo = new Set(infoIds || salesIds || []);
    root.querySelectorAll(".cp-row").forEach((row) => {
      const show = row.querySelector("[data-pcal]"), sales = row.querySelector("[data-scal]"), info = row.querySelector("[data-ical]");
      const sync = () => {
        sales.disabled = info.disabled = !show.checked;
        if (!show.checked) sales.checked = info.checked = false;
        row.classList.toggle("off", !show.checked);
      };
      sales.checked = show.checked && want.has(sales.dataset.scal);
      info.checked = show.checked && (wantInfo.has(info.dataset.ical) || sales.checked);
      show.addEventListener("change", sync);
      sales.addEventListener("change", () => { if (sales.checked) info.checked = true; });
      info.addEventListener("change", () => { if (!info.checked) sales.checked = false; });
      sync();
    });
  }
  function readCalPicker(root, cals) {
    const list = [...root.querySelectorAll("[data-pcal]")].filter((x) => x.checked).map((x) => cals[+x.dataset.pcal])
      .map((c) => ({ id: c.id, name: c.name, color: c.color }));
    const ids = [...root.querySelectorAll("[data-scal]")].filter((x) => x.checked).map((x) => x.dataset.scal);
    const iids = [...root.querySelectorAll("[data-ical]")].filter((x) => x.checked).map((x) => x.dataset.ical);
    return { calendars: list, sales: list.filter((c) => ids.includes(c.id)).map((c) => ({ id: c.id, name: c.name })),
      info: list.filter((c) => iids.includes(c.id)).map((c) => ({ id: c.id, name: c.name })) };
  }
  // 고른 적 없을 때 영업일정 캘린더: 관리자가 정한 영업일정 캘린더 이름들 → config 기본(1팀)과 같은 이름
  function defaultSalesIds(cals) {
    const names = isAdmin() ? salesArr((appSettings.google || {}).sales).map((s) => s.name) : [];
    const want = names.length ? names : [GCFG.teamCalendar || "1팀"];
    return cals.filter((x) => want.includes(x.name)).map((x) => x.id);
  }



  // ---- AI 도우미 창 (v214: 화면 틀만. 말하기·입력칸·예시. 말을 알아듣고 실행하는 것은 다음 단계) ----
  // 상태: aiIsOpen(열려 있음), aiClosing(내려가는 중), aiHist(뒤로가기용으로 쌓아 둔 기록이 있음)
  var aiIsOpen = false, aiClosing = null, aiHist = false, aiExpectPop = false, aiSkipOther = false; // var: 앞쪽 코드(show)에서 먼저 불려도 안전
  function aiFinish() { // 내려가는 애니메이션이 끝났을 때(또는 바로 다시 열 때) 완전히 숨김
    if (!aiClosing) return;
    clearTimeout(aiClosing); aiClosing = null;
    const sh = $("aiSheet");
    $("aiScrim").hidden = true; sh.hidden = true; sh.style.transform = ""; sh.style.transition = ""; sh.style.opacity = "";
    document.documentElement.classList.remove("ai-open");
  }
  // ---------- 오른쪽 아래 떠 있는 AI 버튼 (v408, 2026-11-17 사용자 지시): 몰리(두더지)를 단순한 선 그림으로, 녹색 테두리, 반투명이라 아래 내용이 비침 ----------
  // 말만해요 안(embed)에서는 그 자리(오른쪽 아래 88×88)를 비워 두므로 만들지 않음. 로그인 화면·AI 창이 열려 있을 때는 숨김
  if (!EMBED) {
    const fab = document.createElement("button");
    fab.type = "button"; fab.id = "aiFab"; fab.className = "ai-fab"; fab.setAttribute("aria-label", "AI 도우미 열기(몰리)");
    fab.innerHTML = `<svg viewBox="0 0 48 48" width="38" height="38" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M13 42V25C13 15.5 17.8 10.2 24.6 10.2S36.4 15.5 36.4 25V42"/><path d="M10.5 42h28.4"/><path d="M18.6 11.6C19 7.6 30.2 7.4 30.6 11.4"/><circle cx="24.6" cy="7" r="1" fill="currentColor" stroke="none"/><circle cx="21" cy="19.4" r="1.25" fill="currentColor" stroke="none"/><circle cx="28.4" cy="19.4" r="1.25" fill="currentColor" stroke="none"/><ellipse cx="24.7" cy="24" rx="3.9" ry="2.8"/><ellipse cx="24.7" cy="22.9" rx="1.3" ry=".9" fill="currentColor" stroke="none"/><path d="M13 29.5C9.2 27.6 7 23.6 7.4 19.6c.3-2 2.6-2 3 0 .5 2.6 1.4 4.6 2.6 5.6"/><path d="M36.4 28.5c2.2 1.2 2.6 4 .4 5.6"/></svg>`;
    // v410(2026-11-18 사용자 지시): 버튼을 꾹 끌어 원하는 자리로 옮기면 그 자리를 이 기기에 기억(화면 크기 비율로 저장해 가로·세로가 바뀌어도 화면 안에 둠)
    const FAB_KEY = "shlife.fabPos";
    const fabLoad = () => { try { const o = JSON.parse(localStorage.getItem(FAB_KEY) || "null"); return o && isFinite(o.x) && isFinite(o.y) ? o : null; } catch (e) { return null; } };
    const fabPlace = (o) => {
      if (!o) { fab.style.left = fab.style.top = fab.style.right = fab.style.bottom = ""; return; }
      const w = fab.offsetWidth || 56, h = fab.offsetHeight || 56, W = innerWidth, H = innerHeight;
      fab.style.left = Math.max(4, Math.min(W - w - 4, o.x * W)) + "px";
      fab.style.top = Math.max(4, Math.min(H - h - 4, o.y * H)) + "px";
      fab.style.right = fab.style.bottom = "auto";
    };
    let drag = null, dragged = false;
    fab.addEventListener("pointerdown", (e) => { const r = fab.getBoundingClientRect(); drag = { id: e.pointerId, sx: e.clientX, sy: e.clientY, ox: r.left, oy: r.top }; dragged = false; });
    fab.addEventListener("pointermove", (e) => {
      if (!drag || e.pointerId !== drag.id) return;
      const dx = e.clientX - drag.sx, dy = e.clientY - drag.sy;
      if (!dragged && Math.hypot(dx, dy) < 8) return;
      if (!dragged) { dragged = true; try { fab.setPointerCapture(e.pointerId); } catch (x) { /* 무시 */ } fab.classList.add("dragging"); }
      fabPlace({ x: (drag.ox + dx) / innerWidth, y: (drag.oy + dy) / innerHeight });
    });
    const fabEnd = () => {
      if (!drag) return; drag = null; fab.classList.remove("dragging");
      if (!dragged) return;
      const r = fab.getBoundingClientRect(), o = { x: r.left / innerWidth, y: r.top / innerHeight };
      try { localStorage.setItem(FAB_KEY, JSON.stringify(o)); } catch (e) { /* 이번에만 그 자리 */ }
    };
    fab.addEventListener("pointerup", fabEnd); fab.addEventListener("pointercancel", fabEnd);
    fab.addEventListener("click", (e) => { if (dragged) { dragged = false; e.preventDefault(); return; } if (typeof aiOpen !== "function") return; wakeStop(); aiOpen(); setTimeout(() => { const b = $("aiMic"); if (b && aiIsOpen && !micRec) b.click(); }, 350); }); // v413(2026-11-18 사용자 지시): 누르면 AI 창을 열고 바로 듣기(마이크) 시작
    fab.addEventListener("dblclick", (e) => e.preventDefault());
    document.body.appendChild(fab);
    fabPlace(fabLoad());
    window.addEventListener("resize", () => fabPlace(fabLoad()));
  }
  // ---------- “몰리야”로 AI 부르기 (v407 시험용, 2026-11-17 사용자 지시) ----------
  // 설정에서 켠 기기에서만, 앱 화면이 앞에 있고 AI 창·다른 음성 입력이 쓰이지 않을 때 브라우저 음성 인식으로 계속 듣다가 '몰리야'가 들리면 AI 창을 띄움.
  // '몰리야 ○○' 처럼 이어 말한 내용은 AI 입력칸에 넣어 보내고, 이름만 부르면 AI 창의 마이크가 바로 켜짐. 화면이 꺼지거나 다른 앱으로 가면 멈춤
  const WAKE_KEY = "shlife.wake";
  const wakeOn = () => { try { return localStorage.getItem(WAKE_KEY) === "1"; } catch (e) { return false; } };
  const WAKE_RE = /몰\s*리\s*(야|아)/;
  let wakeRec = null, wakeTimer = null;
  function wakeBadge() {
    let b = document.getElementById("wakeBadge");
    if (!wakeOn() || !account() || EMBED) { if (b) b.remove(); return; }
    if (!b) { b = document.createElement("button"); b.type = "button"; b.id = "wakeBadge"; b.className = "wake-badge"; b.title = "누르면 ‘몰리야’ 듣기를 끕니다"; document.body.appendChild(b); b.addEventListener("click", () => { setWake(false); toast("‘몰리야’ 듣기를 껐습니다. 설정 > 화면에서 다시 켤 수 있습니다."); }); }
    b.classList.toggle("on", !!wakeRec);
    b.textContent = wakeRec ? "몰리 듣는 중" : "몰리 쉬는 중";
  }
  function wakeCanRun() { return wakeOn() && !!account() && !EMBED && document.visibilityState === "visible" && !aiIsOpen && !micRec; }
  function wakeStart() {
    clearTimeout(wakeTimer);
    if (wakeRec || !wakeCanRun()) { wakeBadge(); return; }
    const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SR) { setWake(false); toast("이 브라우저는 ‘몰리야’ 듣기를 지원하지 않습니다(안드로이드 크롬에서 써 주세요)."); return; }
    const rec = new SR();
    rec.lang = "ko-KR"; rec.continuous = true; rec.interimResults = true;
    rec.onresult = (ev) => {
      for (let k = ev.resultIndex; k < ev.results.length; k++) {
        const t = ev.results[k][0].transcript, m = WAKE_RE.exec(t);
        if (!m) continue;
        if (!ev.results[k].isFinal && t.slice(m.index + m[0].length).trim()) continue; // 이어 말하는 중이면 끝날 때까지 기다림
        wakeStop(); wakeFire(t.slice(m.index + m[0].length).trim());
        return;
      }
    };
    rec.onerror = (ev) => { if (ev.error === "not-allowed" || ev.error === "service-not-allowed") { setWake(false); toast("마이크 사용이 막혀 ‘몰리야’ 듣기를 껐습니다."); } };
    rec.onend = () => { if (wakeRec !== rec) return; wakeRec = null; wakeBadge(); wakeTimer = setTimeout(wakeStart, 1200); }; // 크롬이 잠시 뒤 스스로 끝내면 다시 시작
    try { rec.start(); wakeRec = rec; } catch (e) { wakeRec = null; wakeTimer = setTimeout(wakeStart, 3000); }
    wakeBadge();
  }
  function wakeStop() { clearTimeout(wakeTimer); const r = wakeRec; wakeRec = null; if (r) { r.onend = null; try { r.abort(); } catch (e) { /* 이미 멈춤 */ } } wakeBadge(); }
  function wakeFire(rest) {
    aiOpen();
    const inp = $("aiInp");
    if (rest && inp) { inp.value = rest; const f = $("aiForm"); if (f && f.requestSubmit) f.requestSubmit(); return; }
    setTimeout(() => { const b = $("aiMic"); if (b && aiIsOpen && !micRec) b.click(); }, 350); // 이름만 불렀으면 AI 마이크를 바로 켬
  }
  function setWake(on) { try { localStorage.setItem(WAKE_KEY, on ? "1" : "0"); } catch (e) { /* 무시 */ } const c = document.getElementById("wakeMode"); if (c) c.checked = on; if (on) wakeStart(); else wakeStop(); wakeBadge(); }
  document.addEventListener("change", (ev) => { if (ev.target && ev.target.id === "wakeMode") { setWake(ev.target.checked); if (ev.target.checked) toast("‘몰리야’ 듣기를 켰습니다. 앱 화면이 열려 있는 동안만 듣습니다."); } });
  document.addEventListener("visibilitychange", () => { if (document.visibilityState === "visible") setTimeout(wakeStart, 800); else wakeStop(); });
  function aiOpen() {
    if (!account() || aiIsOpen) return; // 로그인 전에는 열지 않음
    wakeStop();
    aiFinish();
    aiIsOpen = true;
    $("aiScrim").hidden = false; $("aiSheet").hidden = false;
    $("aiSheet").style.transform = "";
    $("navAi").classList.add("on"); $("navAi").setAttribute("aria-expanded", "true");
    document.documentElement.classList.add("ai-open");
    history.pushState({ v: "ai" }, ""); aiHist = true; // 폰의 뒤로가기로 닫을 수 있게 기록 한 칸 쌓기
  }
  // 닫기: 끌기·AI 버튼·바깥·Esc·뒤로가기 모두 같은 방식(항상 아래로 내려감). fromPop=뒤로가기로 이미 기록이 빠진 경우
  function aiClose(fromPop) {
    if (!aiIsOpen) return;
    aiIsOpen = false;
    micStop();
    const sh = $("aiSheet");
    sh.style.transition = "transform 200ms var(--ease, ease-out)"; sh.style.transform = "translateY(100%)";
    $("navAi").classList.remove("on"); $("navAi").setAttribute("aria-expanded", "false");
    $("aiMsg").hidden = true; $("aiConfirm").hidden = true; aiPending = null;
    aiClosing = setTimeout(aiFinish, 220);
    setTimeout(wakeStart, 1200); // '몰리야' 듣기 다시
    if (aiHist) { // 쌓아 둔 기록 정리: 뒤로가기로 닫은 게 아니면 한 칸 되돌려 이중 이동을 막음
      aiHist = false;
      if (fromPop !== "pop") { aiExpectPop = true; history.back(); }
    }
  }
  window.addEventListener("popstate", () => { // 다른 뒤로가기 처리(고객 정보 → 목록)보다 먼저 등록되어 먼저 실행됨
    if (aiExpectPop) { aiExpectPop = false; aiSkipOther = true; return; } // 우리가 되돌린 기록: 다른 처리는 건너뜀
    if (aiIsOpen) { aiClose("pop"); aiSkipOther = true; } // 폰 뒤로가기: AI 창만 닫고 화면은 그대로
  });
  $("navAi").addEventListener("click", () => { if (aiIsOpen) aiClose(); else aiOpen(); }); // 아이콘 다시 누르면 닫힘
  $("aiScrim").addEventListener("click", () => aiClose());
  document.addEventListener("keydown", (e) => { if (e.key === "Escape" && aiIsOpen) aiClose(); });
  // 아래로 끌어서 닫기: 팝업 안 어디를 잡고 내려도 됨. 세로로 8px 넘게 끌 때부터 끌기로 보고(짧은 탭은 그대로 동작),
  // 80px 넘게 끌면 닫히고 덜 끌면 제자리로. 내용이 스크롤되는 중이면(맨 위가 아니면) 스크롤이 먼저. 입력칸에서 시작한 터치는 건드리지 않음.
  { const sh = $("aiSheet"), bd = sh.querySelector(".ai-body");
    let y0 = null, dy = 0, drag = false, mouse = false;
    const begin = (y, tgt, isMouse) => { if (tgt && tgt.closest && tgt.closest("input")) return; y0 = y; dy = 0; drag = false; mouse = isMouse; };
    const move = (y, ev) => {
      if (y0 === null) return;
      const d = y - y0;
      if (!drag) {
        if (d > 8 && bd.scrollTop <= 0) { drag = true; sh.style.transition = "none"; } // 맨 위에서 아래로 끌 때만
        else if (d < -8 || bd.scrollTop > 0) { y0 = null; return; } // 위로 끌거나 스크롤 중이면 끌기 아님
        else return;
      }
      dy = Math.max(0, d); sh.style.transform = "translateY(" + dy + "px)";
      if (ev && ev.cancelable) ev.preventDefault(); // 폰 브라우저의 당겨서 새로고침·스크롤과 겹치지 않게
    };
    const end = () => { if (y0 === null) return; const was = drag; y0 = null; drag = false; sh.style.transition = ""; if (!was) return; if (dy > 80) aiClose(); else sh.style.transform = ""; };
    sh.addEventListener("touchstart", (e) => begin(e.touches[0].clientY, e.target, false), { passive: true });
    sh.addEventListener("touchmove", (e) => move(e.touches[0].clientY, e), { passive: false });
    sh.addEventListener("touchend", end); sh.addEventListener("touchcancel", end);
    // 마우스(PC 화면 시험용)
    sh.addEventListener("pointerdown", (e) => { if (e.pointerType === "mouse") begin(e.clientY, e.target, true); });
    window.addEventListener("pointermove", (e) => { if (mouse) move(e.clientY, null); });
    window.addEventListener("pointerup", () => { if (mouse) { end(); mouse = false; } });
  }
  $("aiMic").addEventListener("click", (e) => dictate($("aiInp"), e.currentTarget));
  $("aiChips").addEventListener("click", (e) => { const b = e.target.closest("button"); if (!b) return; $("aiInp").value = b.textContent; $("aiInp").focus(); });

  // ---- AI 도우미: 말을 규칙으로 해석 → 확인 카드 → 실행 (v224: 고객 검색 · 일정 추가 · 일정 보기. 외부 AI 없음, 이름이 폰 밖으로 나가지 않음) ----
  const AI_NUM = { 열두: 12, 열한: 11, 열: 10, 아홉: 9, 여덟: 8, 일곱: 7, 여섯: 6, 다섯: 5, 네: 4, 세: 3, 두: 2, 한: 1 };
  const AI_DOW = "일월화수목금토";
  const aiDateLabel = (d) => `${d.getMonth() + 1}월 ${d.getDate()}일(${AI_DOW[d.getDay()]})`;
  function aiParseDate(t, now) {
    const today = dayStart(now);
    let date = null, rel = "", rest = t, m;
    if ((m = rest.match(/(\d{1,2})\s*월\s*(\d{1,2})\s*일/))) {
      date = new Date(today.getFullYear(), +m[1] - 1, +m[2]);
      if (date < today) date = new Date(today.getFullYear() + 1, +m[1] - 1, +m[2]);
      rest = rest.replace(m[0], " ");
    } else if ((m = rest.match(/(다음\s*주|담주|이번\s*주|금주)?\s*([월화수목금토일])요일/))) {
      const want = AI_DOW.indexOf(m[2]), cur = today.getDay();
      let diff = (want - cur + 7) % 7; // 말 안 하면 오늘부터 가장 가까운 그 요일
      if (/다음|담주/.test(m[1] || "")) { const mon = addDays(today, -((cur + 6) % 7)); diff = Math.round((addDays(mon, 7 + ((want + 6) % 7)) - today) / 86400000); } // 다음 주(월요일 시작) 그 요일
      date = addDays(today, diff);
      rest = rest.replace(m[0], " ");
    } else if (/글피/.test(rest)) { date = addDays(today, 3); rest = rest.replace("글피", " "); }
    else if (/모레/.test(rest)) { date = addDays(today, 2); rest = rest.replace("모레", " "); }
    else if (/내일/.test(rest)) { date = addDays(today, 1); rest = rest.replace("내일", " "); }
    else if (/오늘/.test(rest)) { date = today; rest = rest.replace("오늘", " "); }
    else if ((m = rest.match(/(?:^|\s)(\d{1,2})\s*일(?!정)/))) {
      date = new Date(today.getFullYear(), today.getMonth(), +m[1]);
      if (date < today) date = new Date(today.getFullYear(), today.getMonth() + 1, +m[1]);
      rest = rest.replace(m[0], " ");
    }
    return { date, rest };
  }
  function aiParseTime(t) {
    let rest = t.replace(/(열두|열한|열|아홉|여덟|일곱|여섯|다섯|네|세|두|한)\s*시/g, (x, n) => AI_NUM[n] + "시");
    const per = (rest.match(/오전|아침|새벽|오후|저녁|밤|점심/) || [""])[0];
    const m = rest.match(/(\d{1,2})\s*시\s*(?:(\d{1,2})\s*분|(반))?/);
    if (!m) return { time: null, rest };
    let h = +m[1]; const min = m[3] ? 30 : (m[2] ? +m[2] : 0);
    if (/오후|저녁|밤/.test(per)) { if (h < 12) h += 12; }
    else if (/오전|아침|새벽/.test(per)) { if (h === 12) h = 0; }
    else if (/점심/.test(per)) { if (h < 11) h += 12; }
    else if (h >= 1 && h <= 6) h += 12; // 말 안 하면 업무 시간으로 보고 1~6시는 오후 (확인 카드에 보임)
    if (h > 23 || min > 59) return { time: null, rest };
    rest = rest.replace(m[0], " ").replace(/오전|아침|새벽|오후|저녁|밤|점심/g, " ");
    return { time: { h, min }, rest };
  }
  function aiInterpret(raw) {
    const t = raw.replace(/\s+/g, " ").replace(/[.!?。]+$/g, "").trim();
    if (!t) return null;
    let m;
    // 일정 보기: "오늘 일정 보여줘"
    if (/일정/.test(t) && /(보여|알려|확인|열어)/.test(t) && !/(등록|추가|잡아|넣어)/.test(t)) {
      const { date } = aiParseDate(t, new Date());
      return { kind: "showDay", date: date || dayStart(new Date()) };
    }
    // 고객 검색: "홍길동 검색" / "홍길동 찾아줘" / "홍길동 고객 조회"
    if ((m = t.match(/^(.+?)\s*(?:고객\s*)?(?:검색|조회|찾아\s*줘|찾아|보여\s*줘)(?:\s*해?\s*줘)?$/)) && !/일정/.test(t)) {
      const name = m[1].replace(/\s*고객$/, "").trim();
      if (name) return { kind: "search", name };
    }
    // 일정 추가: 날짜·시간·제목을 뽑음
    if (/(등록|추가|잡아|넣어|일정)/.test(t)) {
      const now = new Date();
      // "내 일정(에다가)", "내 캘린더에": 1팀이 아니라 로그인한 사람 본인 캘린더(내 FC 이름과 같은 이름)에 등록
      let mine = false, t0 = t;
      const mm = t0.match(/(?:^|\s)내\s*(?:일정|캘린더)\s*(?:에다가|에다|에서|에)?(?=\s|$)/);
      if (mm) { mine = true; t0 = t0.replace(mm[0], " ").replace(/\s+/g, " ").trim(); }
      const d = aiParseDate(t0, now);
      const tm = aiParseTime(d.rest);
      let title = tm.rest
        .replace(/일정\s*(?:을|를)?\s*(?:등록|추가|잡아|넣어)?\s*(?:해?\s*줘|해줘|해)?/g, " ")
        .replace(/(?:등록|추가|잡아|넣어)\s*(?:해?\s*줘|해줘|해)?/g, " ")
        .replace(/(?:^|\s)(?:에|에서|를|을|으로|로)(?=\s|$)/g, " ")
        .replace(/\s+/g, " ").trim();
      if (!title) return { kind: "needTitle" };
      return { kind: "addEvent", title, date: d.date || dayStart(now), hasDate: !!d.date, time: tm.time, mine };
    }
    return null;
  }
  // 일정 구분: 글자만 적으면 뒤에 ')'를 붙이고, ')'(또는 '）')까지 적었으면 하나만 남김. 예: 최 → 최) · 최) → 최) · 최)) → 최)
  const fixPrefix = (v) => { const t = String(v || "").trim().replace(/[)）\s]+$/, ""); return t ? t + ")" : ""; };
  async function aiPrefix() {
    try { if (!permPrefixCache) await readPerms(); const fc = await myFcName(); return fixPrefix((permPrefixCache || {})[fc]); } catch (e) { return ""; }
  }
  async function aiTeamCalendar(nameOpt) {
    const name = nameOpt || (GCFG && GCFG.teamCalendar) || "1팀";
    const list = await calendarList();
    const c = list.find((x) => x.name === name);
    if (!c) throw new Error("'" + name + "' 캘린더를 찾지 못했습니다. 일정 화면의 캘린더 설정을 확인해 주세요.");
    return c;
  }
  // ---------- Gemini(외부 AI) 연결: 규칙으로 못 알아들은 말만, 고객 이름·민감 숫자를 가린 글로 '의도'만 물음 ----------
  //   키는 말만해요 설정 파일(Drive의 말만해요_설정.json: geminiKey·model)을 읽어 메모리에만 둠. 키가 없거나 실패하면 규칙만(지금처럼).
  const AI_PH = (n) => "[고객" + n + "]";
  // 보낼 글에서 고객 이름(고객 목록과 맞는 이름, ○○ 고객·김OO 꼴)과 민감 숫자·주소를 자리표시로 바꿈. names = 고객 이름 목록
  function aiMaskText(text, names) {
    const map = new Map(); // "[고객1]" → 실제 이름
    let n = 0, t = String(text || "");
    const ph = (real) => { for (const [k, v] of map) if (v === real) return k; const k = AI_PH(++n); map.set(k, real); return k; };
    [...new Set(names || [])].filter((x) => x && x.length >= 2).sort((a, b) => b.length - a.length).forEach((nm) => { if (t.includes(nm)) t = t.split(nm).join(ph(nm)); }); // 긴 이름부터
    t = t.replace(/([가-힣]{2,4})(\s*)(고객|님|씨)/g, (m, nm, sp, suf) => (/^(저|제|내|이|그|새|옛|각|한|우리|기존|신규|가망|해당|모든|전체|어떤|그냥|해당|이번|다음|지난)$/.test(nm) ? m : ph(nm) + sp + suf)); // 목록에 없어도 "OO 고객·님·씨" 꼴은 이름으로 봄
    t = t.replace(/[가-힣][OoXx○●*＊ㅇ]{1,2}(?![A-Za-z가-힣])/g, (m) => ph(m)); // 김OO 같은 가림 이름
    t = t.replace(/[\w.+-]+@[\w-]+\.[\w.-]+/g, "[이메일]"); // 이메일
    t = t.replace(/\d[\d\s-]{6,}\d/g, "[숫자]"); // 주민번호·전화·계좌·카드 등 긴 숫자
    t = t.replace(/(?:[가-힣]{2,6}(?:특별시|광역시|특별자치시|특별자치도|도)\s+)?(?:[가-힣]{1,5}(?:시|군|구)\s+)+[가-힣0-9]{1,6}(?:읍|면|동|로|길|가)(?:\s*\d+(?:-\d+)?(?:번지|번길)?)?/g, "[주소]"); // 주소 (시·군·구 + 동·로·길)
    return { text: t, map };
  }
  const aiUnmask = (str, map) => { let r = String(str == null ? "" : str); for (const [k, v] of map) r = r.split(k).join(v); return r; };
  let aiGemCache = null; // { t, v } 설정을 읽은 시각·결과 (키는 메모리에만)
  async function aiGeminiConfig() {
    if (aiGemCache && Date.now() - aiGemCache.t < (aiGemCache.v ? 3600000 : 60000)) return aiGemCache.v;
    let v = null;
    try {
      const fs = await gdQuery("name = '말만해요_설정.json' and trashed = false");
      for (const f of fs) {
        const j = await gdFetch(GDAPI + "/files/" + encodeURIComponent(f.id) + "?alt=media&supportsAllDrives=true");
        if (j && typeof j.geminiKey === "string" && j.geminiKey.trim()) { v = { key: j.geminiKey.trim(), model: String(j.model || "gemini-2.5-flash").trim() || "gemini-2.5-flash" }; break; }
      }
    } catch (e) { v = null; }
    aiGemCache = { t: Date.now(), v };
    return v;
  }
  const AI_GEM_PROMPT = (today, dow) => "너는 보험 설계사 앱의 말 해석기다. 사용자가 한 말을 아래 JSON 하나로만 답하라(설명 금지).\n" +
    '{"kind":"search|addEvent|showDay|none","name":"고객 이름(검색일 때)","title":"일정 제목(일정 추가일 때)","date":"YYYY-MM-DD 또는 빈 문자열","time":"HH:MM(24시간) 또는 빈 문자열","mine":false}\n' +
    "kind: 고객을 찾아 달라는 말=search, 일정을 만들거나 약속을 잡는 말=addEvent, 일정을 보여 달라는 말=showDay, 그 밖에는 none.\n" +
    "[고객1] 같은 자리표시는 고객 이름이다. 바꾸지 말고 그대로 name·title에 넣어라. 말에 없는 것은 지어내지 말고 빈 문자열로 둬라.\n" +
    "'내 일정에/내 캘린더에'라고 하면 mine=true. 오늘은 " + today + "(" + dow + "요일)이다. 상대 날짜(내일·다음 주 화요일 등)는 날짜로 바꿔라. 시간을 말하지 않으면 time은 빈 문자열.";
  // 규칙으로 못 알아들은 말을 Gemini에 물음 → 확인 카드용 항목 (실패·키 없음이면 null / {noKey:true})
  async function aiAskGemini(raw) {
    const cfg = await aiGeminiConfig();
    if (!cfg) return { noKey: true };
    let names = [];
    try { names = [...new Set((await loadLists()).map((r) => baseName(r.name)).filter(Boolean))]; } catch (e) { names = []; }
    const m = aiMaskText(raw.replace(/\s+/g, " ").trim(), names);
    const now = new Date(), today = `${now.getFullYear()}-${pad2(now.getMonth() + 1)}-${pad2(now.getDate())}`;
    console.info("[AI→Gemini 보내는 글]", m.text); // 확인용: 이름·민감 숫자가 가려졌는지
    let j;
    try {
      const gemReq = { systemInstruction: { parts: [{ text: AI_GEM_PROMPT(today, AI_DOW[now.getDay()]) }] }, contents: [{ role: "user", parts: [{ text: m.text }] }], generationConfig: { responseMimeType: "application/json", temperature: 0 } };
      let res = await gemCall(cfg, gemReq, null); // v565: 모델 기억·없는 모델 건너뛰기·20초 제한·생각 끄기
      if (!res || !res.ok) return null;
      const d = await res.json();
      const txt = ((((d.candidates || [])[0] || {}).content || {}).parts || []).filter((x) => !x.thought).map((x) => x.text || "").join("");
      j = JSON.parse(txt.replace(/^```json\s*|\s*```$/g, ""));
    } catch (e) { return null; }
    if (!j || typeof j !== "object") return null;
    const dm = String(j.date || "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
    const date = dm ? new Date(+dm[1], +dm[2] - 1, +dm[3]) : null;
    if (j.kind === "search") { const name = aiUnmask(j.name, m.map).trim(); return name ? { kind: "search", name, viaAi: true } : null; }
    if (j.kind === "showDay") return { kind: "showDay", date: date && !isNaN(+date) ? date : dayStart(now), viaAi: true };
    if (j.kind === "addEvent") {
      const title = aiUnmask(j.title, m.map).replace(/\s+/g, " ").trim();
      if (!title) return { kind: "needTitle" };
      const tm = String(j.time || "").match(/^(\d{1,2}):(\d{2})$/);
      const time = tm && +tm[1] < 24 && +tm[2] < 60 ? { h: +tm[1], min: +tm[2] } : null;
      return { kind: "addEvent", title, date: date && !isNaN(+date) ? date : dayStart(now), hasDate: !!(date && !isNaN(+date)), time, mine: !!j.mine, viaAi: true };
    }
    return null;
  }
  let aiPending = null;
  function aiShowMsg(html) { const m = $("aiMsg"); m.innerHTML = html; m.hidden = false; $("aiConfirm").hidden = true; aiPending = null; }
  async function aiSubmit() {
    const txt = $("aiInp").value;
    let it = aiInterpret(txt);
    let noKey = false;
    if (!it && txt.trim()) { // 규칙으로 못 알아들은 말만 Gemini에 물음 (이름·민감 숫자는 가려서)
      aiShowMsg("생각하는 중…");
      const g = await aiAskGemini(txt);
      if (g && g.noKey) noKey = true; else it = g;
    }
    if (!it) return aiShowMsg("못 알아들었습니다. 이렇게 말해 보세요.<br>• “홍길동 검색”<br>• “내일 오후 3시 홍길동 상담 일정 추가”<br>• “오늘 일정 보여줘”" + (noKey ? "<br><small>말만해요 설정에서 Gemini 키를 넣으면 더 많은 말을 알아들어요.</small>" : ""));
    if (it.kind === "needTitle") return aiShowMsg("무슨 일정인지 알려 주세요. 예: “내일 오후 3시 홍길동 상담 일정 추가”");
    $("aiMsg").hidden = true;
    let head, line1, line2 = "";
    if (it.kind === "search") { head = "검색"; line1 = esc(it.name); if (it.viaAi) line2 = "AI가 해석"; }
    else if (it.kind === "showDay") { head = "일정 보기"; line1 = aiDateLabel(it.date); }
    else {
      const pre = it.mine ? "" : await aiPrefix(); // 내 일정에는 앞 글자를 붙이지 않음(팀 캘린더에서 누구 것인지 구분하려는 글자라서)
      it.summary = pre + it.title;
      if (it.mine) { it.calName = await myFcName(); if (!it.calName) return aiShowMsg("내 FC 이름이 아직 정해지지 않아 ‘내 일정’을 찾을 수 없습니다. 관리자에게 이름 등록을 요청해 주세요."); }
      else it.calName = (GCFG && GCFG.teamCalendar) || "1팀";
      const when = it.time ? `${it.time.h < 12 ? "오전" : "오후"} ${it.time.h % 12 || 12}:${pad2(it.time.min)}` : "종일";
      head = "일정 추가"; line1 = `${aiDateLabel(it.date)} ${when}`;
      line2 = `${esc(it.summary)} · ${esc(it.calName)} 캘린더` + (it.hasDate ? "" : " · 날짜를 말하지 않아 오늘로 했습니다") + (it.viaAi ? " · AI가 해석" : "");
    }
    aiPending = it;
    $("aiConfirm").innerHTML = `<div class="ai-ch">이렇게 이해했습니다</div><div class="ai-cm"><b>${head}</b> · ${line1}${line2 ? `<small>${line2}</small>` : ""}</div><div class="ai-cb"><button type="button" class="ai-no" id="aiNo">아니요</button><button type="button" class="ai-go" id="aiGo">실행</button></div>`;
    $("aiConfirm").hidden = false;
    $("aiNo").addEventListener("click", () => { aiPending = null; $("aiConfirm").hidden = true; });
    $("aiGo").addEventListener("click", aiRun);
  }
  async function aiRun() {
    const it = aiPending; if (!it) return;
    const go = $("aiGo"); if (go) { go.disabled = true; go.textContent = "실행 중…"; }
    try {
      if (it.kind === "search") {
        aiClose(); await new Promise((r) => setTimeout(r, 320)); // AI 창이 내려가고 뒤로가기 기록이 정리된 뒤에 이동
        $("aiInp").value = ""; runTopSearch(it.name, true);
      } else if (it.kind === "showDay") {
        aiClose(); await new Promise((r) => setTimeout(r, 320));
        $("aiInp").value = ""; schedMode = "day"; schedDate = dayStart(it.date); openSchedule();
      } else {
        const cal = await aiTeamCalendar(it.calName);
        const tz = (Intl.DateTimeFormat().resolvedOptions() || {}).timeZone || "Asia/Seoul";
        const ymd = (d) => `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
        let body;
        if (it.time) {
          const sH = it.time.h, sM = it.time.min, end = new Date(it.date.getFullYear(), it.date.getMonth(), it.date.getDate(), sH + 1, sM);
          body = { summary: it.summary, start: { dateTime: `${ymd(it.date)}T${pad2(sH)}:${pad2(sM)}:00`, timeZone: tz }, end: { dateTime: `${ymd(end)}T${pad2(end.getHours())}:${pad2(end.getMinutes())}:00`, timeZone: tz } };
        } else {
          body = { summary: it.summary, start: { date: ymd(it.date) }, end: { date: ymd(addDays(it.date, 1)) } };
        }
        await gcalInsert(cal.id, body);
        toast("일정을 등록했습니다: " + it.summary);
        aiClose(); await new Promise((r) => setTimeout(r, 320));
        $("aiInp").value = ""; schedMode = "day"; schedDate = dayStart(it.date); openSchedule();
      }
    } catch (e) {
      const msg = e && e.message === "GOOGLE_AUTH" ? "구글 연결이 풀렸습니다. 일정 화면에서 다시 연결해 주세요." : (e && e.message) || "실행하지 못했습니다.";
      aiShowMsg(esc(msg));
    }
  }
  $("aiForm").addEventListener("submit", (e) => { e.preventDefault(); micStop(); aiSubmit(); });

  // 큰 메뉴: 일정 / 고객 / 관리자
  $("mainNav").querySelectorAll("button[data-section]").forEach((b) => b.addEventListener("click", async () => {
    aiClose(true); // 다른 메뉴로 가면 AI 창은 닫음
    if (!(await confirmLeaveEdit())) return;
    const s = b.dataset.section;
    if (s === "schedule") { schedMode = "day"; schedDate = dayStart(new Date()); openSchedule(); } // 항상 오늘 Today
    else if (s === "admin") openAdmin();
    else if (s === "dash") openDash();
    else if (s === "sales") { SALES.fromNtf = false; openSales(); }
    else if (s === "team") openTeam();
    else resetSearch(); // 고객: 검색을 지우고 처음 화면
  }));
  window.addEventListener("popstate", () => { if (aiSkipOther) { aiSkipOther = false; return; } if (!$("customerView").hidden) { if (returnView === "scheduleView") openSchedule(); else if (returnView === "salesView") renderSales(SALES.cardAnchor || true); else if (returnView === "introView") { show("introView"); renderIntro(); } else show("searchView"); } });

  function resetSearch() {
    $("q").value = ""; $("q2").value = "";
    $("results").innerHTML = ""; $("searchStatus").textContent = "";
    $("recentList").hidden = true;
    show("searchView");
  }

  // (검색 가능한 FC 문구는 표시하지 않음)
  function showScope() { $("scopeHint").textContent = ""; }

  // ---------- 문자 문구 (세일즈 화면의 [문자] 버튼에 미리 채움) ----------
  //   내 문구: 내설정 파일(shlife-내설정.json)의 sms.{transfer,dda}(저장 안 했으면 config.js smsDefaults)
  const SMS_VARS = "{고객명} {FC명} {상품명} {이관일}";
  const isIOS = () => /iPhone|iPad|iPod/i.test(navigator.userAgent || "") || (navigator.platform === "MacIntel" && navigator.maxTouchPoints > 1);
  const smsCfgDefault = (kind) => String((CFG.smsDefaults || {})[kind] || "");
  // 기본 문구 = 관리자가 설정 파일에 정해 둔 팀 문구(비어 있지 않으면), 없으면 config.js 기본. 개인이 저장한 문구가 있으면 그것이 우선
  const smsDefault = (kind) => { const t = ((appSettings && appSettings.sms) || {})[kind]; return typeof t === "string" && t.trim() ? t : smsCfgDefault(kind); };
  // 개인이 저장한 값(지워서 빈 문구로 저장한 것 포함)이 있으면 그것, 아직 손대지 않았으면 config 기본 문구
  const smsTemplate = (kind) => { const m = (myPrefs && myPrefs.sms) || {}; return typeof m[kind] === "string" ? m[kind] : smsDefault(kind); };
  const idateText = (v) => { const k = dateKey(v); return k ? `${+k.slice(4, 6)}월 ${+k.slice(6, 8)}일` : ""; };
  const fillSms = (tpl, r, fcName) => String(tpl || "")
    .replace(/\{고객명\}/g, r.holder || "").replace(/\{FC명\}/g, r.fc || fcName || "").replace(/\{상품명\}/g, r.product || "").replace(/\{이관일\}/g, idateText(r.idate));
  // 아이폰은 sms:번호&body=, 안드로이드·그 외는 sms:번호?body= (표준). 문구가 없으면 번호만.
  const smsHref = (phone, body) => (body ? `sms:${phone}${isIOS() ? "&" : "?"}body=${encodeURIComponent(body)}` : `sms:${phone}`);

  // ---------- 관리자: FC별 일정 앞 글자 (AI 도우미로 일정을 등록하면 제목 앞에 붙음) ----------
  // ---------- 관리자: 사용자별 일정 구분 (AI 도우미로 일정을 등록하면 제목 앞에 붙음). 권한 파일 prefixes{FC명: 글자} ----------
  function bindPfxRows() {
    const body = $("adminBody");
    body.querySelectorAll("[data-pfx]").forEach((inp) => {
      const em = inp.dataset.pfx, btn = body.querySelector(`[data-pfx-save="${CSS.escape(em)}"]`);
      let base = inp.value;
      inp.addEventListener("input", () => { btn.disabled = inp.value.trim() === base.trim(); });
      btn.addEventListener("click", async () => {
        btn.disabled = true; btn.textContent = "저장 중…";
        try {
          const perms = await readPerms();
          const nm = (perms.names || {})[em];
          if (!nm) throw new Error("이름(FC명)이 없습니다.");
          const next = { ...(perms.prefixes || {}) };
          const v = fixPrefix(inp.value);
          if (v) next[nm] = v; else delete next[nm];
          perms.prefixes = next;
          await writePerms(perms);
          adminState.perms = { ...adminState.perms, prefixes: next };
          inp.value = v; base = v; toast(`${nm}: 일정 구분을 저장했습니다.${v ? " (일정 제목 앞에 “" + v + "”가 붙습니다)" : ""}`);
        } catch (e) { btn.disabled = false; handleError(e); }
        btn.textContent = "저장";
      });
    });
  }
  // ---------- 시작 설정: 앱을 열 때 처음 보일 화면 + AI 도우미 바로 열기 (이 기기·이 계정에만 적용) ----------
  const START_PAGES = [["schedule", "일정"], ["customer", "고객"], ["sales", "세일즈"]]; // 하단 메뉴 순서
  const START_FIXED = "dash"; // v520(2026-11-21 사용자 지시): 처음 화면은 대시보드로 고정하고 시작 화면 고르기는 숨김(나중에 살리려면 ""로)
  function startPrefKey() { return "shlife.start." + (gAccount() || "me"); }
  function readStartPref() {
    let v = null;
    try { v = JSON.parse(localStorage.getItem(startPrefKey()) || "null"); } catch (e) { v = null; }
    if (!v) { try { v = JSON.parse(localStorage.getItem("shlife.startAdmin." + me()) || "null"); } catch (e) { v = null; } } // 내가 안 골랐으면 관리자가 정한 값
    v = v || {};
    return { page: START_PAGES.some((p) => p[0] === v.page) ? v.page : "customer", ai: !!v.ai };
  }
  function writeStartPref(v) { try { localStorage.setItem(startPrefKey(), JSON.stringify(v)); return true; } catch (e) { return false; } }
  function startCardHtml() { return ""; } // v337: '화면' 카드(rhCardHtml)로 합침
  function startCardHtmlOld() {
    const cur = readStartPref();
    return `<div class="card" id="startCard"><h3 class="sec-title">시작 화면</h3>
      <p class="hint">앱을 열면 처음 보일 화면을 고릅니다. 내 FC 폴더에 저장되어 관리자도 볼 수 있습니다.</p>
      <div class="st-opts" role="radiogroup" aria-label="시작 화면">${START_PAGES.map(([k, l]) =>
        `<label class="st-opt"><input type="radio" name="startPage" value="${k}" ${cur.page === k ? "checked" : ""}><span>${l}</span></label>`).join("")}</div>
      <label class="st-ai"><input type="checkbox" id="startAi" ${cur.ai ? "checked" : ""}><span>시작하면 AI 도우미를 바로 열기</span></label>
      <div class="file-actions"><button type="button" class="primary" id="startSave" disabled>저장</button></div></div>`;
  }
  function bindStartCard() { // v337: 화면 카드의 AI·시작 화면은 바꾸면 바로 저장
    const sel = $("startPage"), ai = $("startAi"); if (!sel || !ai) return;
    const save = async () => {
      const v = { page: sel.value, ai: ai.checked }, ok = writeStartPref(v);
      let shared = false;
      if (!isAdmin()) { try { const perms = await readPerms(); shared = await writeStartShared(perms, me(), { ...v, rh: rhOn() }, perms.fcFolders); } catch (e) { console.error(e); } }
      toast(ok || shared ? "저장했습니다. 다음에 앱을 열 때 적용됩니다." : "저장하지 못했습니다.");
    };
    sel.addEventListener("change", save); ai.addEventListener("change", save);
  }
  function bindStartCardOld() {
    const box = $("startCard"); if (!box) return;
    const read = () => { const r = box.querySelector('input[name="startPage"]:checked'); return { page: r ? r.value : "customer", ai: $("startAi").checked }; };
    let base = JSON.stringify(read());
    const btn = $("startSave");
    box.addEventListener("change", () => { btn.disabled = JSON.stringify(read()) === base; });
    btn.addEventListener("click", async () => {
      const v = read(), ok = writeStartPref(v);
      let shared = false;
      if (!isAdmin()) { try { const perms = await readPerms(); shared = await writeStartShared(perms, me(), { ...v, rh: rhOn() }, perms.fcFolders); } catch (e) { console.error(e); } }
      if (ok || shared) { base = JSON.stringify(v); btn.disabled = true; }
      toast(shared ? "시작 화면을 저장했습니다. 관리자도 확인할 수 있고, 다음에 앱을 열 때 적용됩니다." : ok ? "이 폰에 저장했습니다. (내 FC 폴더에 쓸 수 없어 관리자 화면에는 보이지 않습니다.) 다음에 앱을 열 때 적용됩니다." : "저장하지 못했습니다.");
    });
  }
  // 사용자별 시작 화면 파일: 그 사용자가 수정 권한을 가진 FC 폴더 안의 shlife-시작-<아이디>.json (사용자·관리자 모두 읽고 쓸 수 있는 곳)
  const startFile = (em) => "shlife-시작-" + String(em).toLowerCase() + ".json";
  function startDirOf(perms, em, fcMap) {
    const map = { ...(perms.fcFolders || {}), ...(fcMap || {}) };
    for (const g of (perms.users[em] || [])) { if (g === "*" || grantRo(g)) continue; const id = map[grantFc(g)]; if (id) return { id }; }
    return null;
  }
  // 고객 파일과 섞이지 않게 FC 폴더 안의 '_앱설정' 하위 폴더에 둠 (앱은 이 폴더를 고객 폴더로 읽지 않음)
  const START_SUB = "_앱설정";
  async function startSubDir(d, create) {
    const r = await gdQuery("'" + qv(d.id) + "' in parents and name='" + qv(START_SUB) + "' and mimeType='" + GD_FOLDER + "' and trashed=false");
    if (r.length) return { id: r[0].id };
    if (!create) return null;
    const mk = await gdFetch(GDAPI + "/files?fields=id&supportsAllDrives=true", { method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ name: START_SUB, mimeType: GD_FOLDER, parents: [d.id] }) });
    return { id: mk.id };
  }
  async function readStartShared(perms, em, fcMap) {
    const d0 = startDirOf(perms, em, fcMap); if (!d0) return null;
    const d = await startSubDir(d0, false); if (!d) return null;
    const v = await stReadJsonIn(d, startFile(em));
    return v && typeof v === "object" && START_PAGES.some((x) => x[0] === v.page) ? { page: v.page, ai: !!v.ai, rh: typeof v.rh === "boolean" ? v.rh : null } : null;
  }
  async function writeStartShared(perms, em, v, fcMap) {
    const d0 = startDirOf(perms, em, fcMap); if (!d0) return false;
    const d = await startSubDir(d0, true);
    await stWriteJsonIn(d, startFile(em), { page: v.page, ai: !!v.ai, ...(typeof v.rh === "boolean" ? { rh: v.rh } : {}), by: me(), at: new Date().toISOString() });
    return true;
  }
  async function syncStartShared() { // 관리자·다른 폰에서 바꾼 값을 이 폰에 반영 (다음 실행부터 적용)
    try {
      const perms = await readPerms(); const v = await readStartShared(perms, me(), perms.fcFolders);
      if (v) writeStartPref({ page: v.page, ai: v.ai });
      if (!v || v.rh !== rhOn()) await writeStartShared(perms, me(), { ...(v || readStartPref()), rh: rhOn() }, perms.fcFolders); // 관리자 조직도에 보이게 (v333)
    } catch (e) { /* 무시 */ }
  }
  // 관리자: 각 사용자 카드 안의 '시작 화면' (그 사용자 FC 폴더의 설정 파일에 저장 — 사용자가 직접 저장한 값도 여기에 보이고, 여기서 바꾸면 그 사용자 폰에 따라감)
  function startRowHtml(perms, em) {
    const st = (perms.starts || {})[em] || {}, pg = START_PAGES.some((p) => p[0] === st.page) ? st.page : "customer";
    return `<div class="perm-start" data-sa="${esc(em)}"><label class="pfx-l"${START_FIXED ? " hidden" : ""}><span>시작 화면</span><select data-sa-page aria-label="${esc(em)} 시작 화면">${START_PAGES.map(([k, l]) => `<option value="${k}" ${pg === k ? "selected" : ""}>${l}</option>`).join("")}</select></label>
      <label class="sa-ai"><input type="checkbox" data-sa-ai ${st.ai ? "checked" : ""}> AI 바로 열기</label><button type="button" class="secondary" data-sa-save="${esc(em)}" disabled>저장</button></div>`;
  }
  function bindStartRows() {
    const body = $("adminBody"), perms0 = adminState.perms;
    body.querySelectorAll("[data-sa]").forEach((box) => {
      const em = box.dataset.sa, sel = box.querySelector("[data-sa-page]"), ai = box.querySelector("[data-sa-ai]"), btn = box.querySelector("[data-sa-save]");
      const read = () => JSON.stringify([sel.value, ai.checked]);
      let base = read(), ready = false;
      let rhKeep = null;
      readStartShared(perms0, em, adminState.fcIds).then((v) => { if (v) { sel.value = v.page; ai.checked = v.ai; rhKeep = v.rh; } base = read(); ready = true; }).catch(() => { ready = true; });
      box.addEventListener("change", () => { btn.disabled = !ready || read() === base; });
      btn.addEventListener("click", async () => {
        btn.disabled = true; btn.textContent = "저장 중…";
        try {
          const perms = await readPerms(), v = { page: sel.value, ai: ai.checked, rh: rhKeep };
          let ok = false;
          try { ok = await writeStartShared(perms, em, v, adminState.fcIds); } catch (e) { console.error(e); }
          orgStartCache = null;
          const next = { ...(perms.starts || {}) };
          if (ok) delete next[em]; else next[em] = v; // FC 폴더가 없으면 권한 파일에 저장
          perms.starts = next; await writePerms(perms);
          adminState.perms = { ...adminState.perms, starts: next };
          base = read(); toast(`${(perms.names || {})[em] || em}: 시작 화면을 저장했습니다.${ok ? "" : " (FC 폴더가 없어 권한 파일에 저장)"}`);
        } catch (e) { btn.disabled = false; handleError(e); }
        btn.textContent = "저장";
      });
    });
  }
  // 관리자: 지원 기능 박스 (2026-11-17 사용자 지시) — '문자 문구 (팀 기본)' 바로 위. [고객 등록]은 새 고객 등록 페이지(addcustom.html)를 새 탭으로 엶
  function supportCardHtml() {
    return `<div class="card" id="supportCard"><div class="sec-title">지원 기능</div><div class="support-row"><a class="ev-edit support-link" href="addcustom.html" target="_blank" rel="noopener">고객 등록</a></div></div>`;
  }
  // 관리자: 팀 기본 문자 문구 (설정 파일 sms) — 사용자가 자기 문구를 저장하지 않았으면 이 문구가 기본
  function smsTeamCardHtml() {
    const ta = (id, label) => `<label class="fld"><span>${label}</span><textarea id="${id}" rows="4" maxlength="600"></textarea></label>`;
    return `<div class="card" id="smsTeamCard"><details class="fold"><summary class="sec-title">문자 문구 (팀 기본)</summary>
      <p class="hint">사용자가 자기 문구를 따로 저장하지 않았으면 이 문구가 기본으로 나갑니다. 사용 가능 항목: ${SMS_VARS}</p>
      ${ta("smsTeamTransfer", "이관 문자 문구")}${ta("smsTeamDda", "따행 문자 문구")}
      <div class="file-actions"><button type="button" class="primary" id="smsTeamSave" disabled>팀 문구 저장</button></div></details></div>`;
  }
  function bindSmsTeamCard() {
    const KINDS = { smsTeamTransfer: "transfer", smsTeamDda: "dda" }, ids = Object.keys(KINDS);
    if (!$("smsTeamCard")) return;
    ids.forEach((id) => { $(id).value = smsDefault(KINDS[id]); });
    const read = () => JSON.stringify(ids.map((id) => $(id).value.trim()));
    let base = read();
    ids.forEach((id) => $(id).addEventListener("input", () => { $("smsTeamSave").disabled = read() === base; }));
    $("smsTeamSave").addEventListener("click", async () => {
      const b = $("smsTeamSave"); b.disabled = true; b.textContent = "저장 중…";
      try {
        const sms = { ...((appSettings && appSettings.sms) || {}) };
        ids.forEach((id) => { const v = $(id).value.trim(); if (v && v !== smsCfgDefault(KINDS[id])) sms[KINDS[id]] = v; else delete sms[KINDS[id]]; });
        await saveSettings({ ...appSettings, sms });
        base = read(); toast("팀 기본 문구를 저장했습니다.");
      } catch (e) { handleError(e); }
      b.textContent = "팀 문구 저장"; b.disabled = read() === base;
    });
  }
  function smsCardHtml() {
    const ta =(id, label) => `<label class="fld"><span>${label}</span><textarea id="${id}" rows="5" maxlength="600"></textarea>
        <span class="smsrow"><button type="button" class="link-btn" data-sms-reset="${id}">기본값으로 되돌리기</button><small class="cnt" data-for="${id}"></small></span></label>`;
    return `<div class="card" id="smsCard"><details class="fold"><summary class="sec-title">문자 문구</summary>
      ${ta("smsMyTransfer", "이관 문자 문구")}
      ${ta("smsMyDda", "따행 문자 문구")}
      <p class="hint">내가 저장한 문구가 있으면 그것을 쓰고, 없으면 관리자가 정해 준 기본 문구를 씁니다.</p>
      <p class="hint" id="smsPreview"></p>
      <div class="file-actions"><button type="button" class="primary" id="smsSave" disabled>문구 저장</button></div></details></div>`;
  }
  async function bindSmsCard() {
    const KINDS = { smsMyTransfer: "transfer", smsMyDda: "dda" };
    const ids = Object.keys(KINDS), el = (id) => $(id);
    const sample = { holder: "홍길동", fc: "최수진", product: "종신보험", idate: "2026.11.09" };
    const cnt = () => ids.forEach((id) => { const c = document.querySelector(`.cnt[data-for="${id}"]`); if (c) c.textContent = `${el(id).value.length}자` + (el(id).value.length > 45 ? " · 45자 초과(장문 문자)" : ""); });
    const read = () => JSON.stringify(ids.map((id) => el(id).value.trim()));
    const preview = (id) => { $("smsPreview").textContent = el(id).value.trim() ? "미리보기: " + fillSms(el(id).value, sample, "최수진") : ""; };
    try { await Promise.all([loadMyPrefs(), loadSettings()]); } catch (e) { /* 기본값 */ }
    if (!$("smsCard")) return;
    ids.forEach((id) => { el(id).value = smsTemplate(KINDS[id]); });
    let base = read(); cnt();
    ids.forEach((id) => {
      el(id).addEventListener("input", () => { cnt(); preview(id); $("smsSave").disabled = read() === base; });
      el(id).addEventListener("focus", () => preview(id));
    });
    $("smsCard").querySelectorAll("[data-sms-reset]").forEach((b) => b.addEventListener("click", () => {
      const id = b.dataset.smsReset; el(id).value = smsDefault(KINDS[id]); cnt(); preview(id); $("smsSave").disabled = read() === base;
    }));
    $("smsSave").addEventListener("click", async () => {
      const b = $("smsSave"); b.disabled = true; b.textContent = "저장 중…";
      try {
        const sms = {};
        ids.forEach((id) => { const v = el(id).value.trim(); if (v !== smsDefault(KINDS[id])) sms[KINDS[id]] = v; }); // 기본 문구와 같으면 저장하지 않음(나중에 기본이 바뀌면 따라감), 지우면 빈 문구로 저장
        await saveMyPrefs({ ...(myPrefs || {}), sms });
        base = read(); toast("문구를 저장했습니다.");
      } catch (e) { handleError(e); }
      b.textContent = "문구 저장"; b.disabled = read() === base;
    });
  }

  // ---------- 세일즈 (이달이관 · 이달 따행) + 전화·문자 '시도' 기록 ----------
  //   이관: 계약 단위(계약 시트 최근통화일시·최근문자일시 + 이력 시트에 계약ID와 함께 한 줄) / 따행: 고객 단위(고객 시트 최근통화일시·최근문자일시 + 이력에 계약ID 비움)
  const SALES = { tab: "transfer", ym: "", showRef: false, arch: { transfer: false, dda: false, valueup: false }, data: null, df: { send: false, resp: false, coupon: false }, nf: { call: false, sms: false }, open: new Set(), dct: "", dcls: "", dact: true };
  const ymNow = () => { const d = new Date(); return d.getFullYear() + pad2(d.getMonth() + 1); };
  const ymShift = (ym, n) => { let y = +ym.slice(0, 4), m = +ym.slice(4) + n; while (m > 12) { m -= 12; y++; } while (m < 1) { m += 12; y--; } return y + pad2(m); };
  const ymLabel = (ym) => ym.slice(0, 4) + "년 " + (+ym.slice(4)) + "월";
  const monthOf = (v) => dateKey(v).slice(0, 6);
  const firstPhone = (s) => { const m = String(s || "").match(/\+?\d[\d -]{7,}/); return m ? m[0].replace(/[^\d+]/g, "") : ""; };
  const cleanAddr = (a) => String(a || "").replace(/^\s*\[[^\]]*\]\s*/, "").replace(/^\(?\d{5}\)?\s*/, "").trim();
  const DEAD_STATUS = ["해약", "사고소멸", "순수보장만기소멸", "연금소멸"];
  // 'YYYY-MM-DD HH:mm' 또는 엑셀 표시 문자열 → 비교용 숫자열 / 표시용 'M/D HH:mm'
  const tsDigits = (v) => String(v || "").replace(/\D/g, "").slice(0, 12);
  const tsShow = (v) => { const d = tsDigits(v); return d.length >= 8 ? `${+d.slice(4, 6)}/${+d.slice(6, 8)}` + (d.length >= 12 ? ` ${d.slice(8, 10)}:${d.slice(10, 12)}` : "") : ""; };
  const nowStamp = () => { const d = new Date(); return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())} ${pad2(d.getHours())}:${pad2(d.getMinutes())}`; };

  // 이력 시트 (읽기): 파일마다 한 번
  let histCache = null;
  function loadHistory() {
    if (histCache) return histCache;
    histCache = (async () => {
      const { files } = await scanFiles();
      const spec = CFG.lists.find((l) => l.fileRegex || l.fileKeyword);
      const out = [];
      await Promise.all(pickSourceFiles(files, spec).map(async (f) => {
        const wb = await readBook(f.id);
        const rows = sheetRows(wb, "이력");
        if (!rows) return;
        const h = headIndex(rows, "고객ID");
        if (h < 0) return;
        const head = rows[h].map((c) => String(c).trim());
        rows.slice(h + 1).forEach((r) => {
          const o = {};
          head.forEach((k, j) => { if (k) o[k] = r[j]; });
          if (String(o["고객ID"] || "").trim()) out.push({ fc: f.fc, fileId: f.id, ...o });
        });
      }));
      return out;
    })();
    histCache.catch(() => { histCache = null; });
    return histCache;
  }

  // ---- 따행 시트 (B 구조): 고객ID × 대상년월 1행. 열 이름은 SOL 화면 이름 그대로 (대상년월·따행분류·접촉현황·캠페인대상여부·활동기한시작일/종료일·방문약속일·방문완료일·설문발송일·설문응답일·쿠폰발송·쿠폰발송일·선물·전달여부)
  let followCache = null;
  function loadFollow() {
    if (followCache) return followCache;
    followCache = (async () => {
      const { files } = await scanFiles();
      const spec = CFG.lists.find((l) => l.fileRegex || l.fileKeyword);
      const out = { files: new Set(), rows: [] };
      await Promise.all(pickSourceFiles(files, spec).map(async (f) => {
        const wb = await readBook(f.id);
        const rows = sheetRows(wb, "따행");
        if (!rows) return;
        const h = headIndex(rows, "고객ID");
        if (h < 0) return;
        out.files.add(f.id); // 이 파일의 고객은 따행 시트 기준(B), 없는 파일은 고객 시트 필드 기준(A)
        const head = rows[h].map((c) => String(c).trim());
        rows.slice(h + 1).forEach((r) => {
          const o = {};
          head.forEach((k, j) => { if (k) o[k] = String(r[j] ?? "").trim(); });
          if (o["고객ID"] && o["대상년월"]) out.rows.push({ ...o, _fc: f.fc, _fileId: f.id });
        });
      }));
      return out;
    })();
    followCache.catch(() => { followCache = null; });
    return followCache;
  }
  // v571(2026-11-24 사용자 승인 'DB 시트'): 밸류업DB 배분 1건 = 'DB' 시트 1줄(유입FC 파일에만). 전환 기간: 읽기 = DB 시트 줄이 있는 FC는 DB 시트, 없으면 옛 칸 / 쓰기 = DB 시트에만
  let dbsCache = null;
  const DBS_BY_ID = new Map(); // 고객ID → DB 시트 줄들(화면·쓰기 공용, 같은 객체)
  function loadDbs() {
    if (dbsCache) return dbsCache;
    dbsCache = (async () => {
      const { files } = await scanFiles();
      const spec = CFG.lists.find((l) => l.fileRegex || l.fileKeyword);
      const out = { rows: [] };
      await Promise.all(pickSourceFiles(files, spec).map(async (f) => {
        const wb = await readBook(f.id);
        const rows = sheetRows(wb, "DB");
        if (!rows) return;
        const h = headIndex(rows, "DB ID");
        if (h < 0) return;
        const head = rows[h].map((c) => String(c).trim());
        rows.slice(h + 1).forEach((r) => {
          const o = {};
          head.forEach((k, j) => { if (k) o[k] = String(r[j] ?? "").trim(); });
          if (o["DB ID"] && o["고객ID"]) out.rows.push({ ...o, _fc: f.fc, _fileId: f.id });
        });
      }));
      DBS_BY_ID.clear(); out.rows.forEach((r) => (DBS_BY_ID.get(r["고객ID"]) || DBS_BY_ID.set(r["고객ID"], []).get(r["고객ID"])).push(r));
      return out;
    })();
    dbsCache.catch(() => { dbsCache = null; });
    return dbsCache;
  }
  const ymOfFollow = (v) => monthOf(v); // '2026-11' → '202611'
  // ---------- 세일즈 > 계약진행 (4단계 ①, 2026-11-17 확정안 5-1) ----------
  // 계약진행 표(FC별 data.json): 한 줄 = 한 회차. 단계 고객등록 > 전화 > AP > PC > 계약, 상태 진행/보관/삭제/완료(단계가 계약이면 자동 완료). 단계·상태가 바뀔 때마다 이력에 '계약진행' 한 줄(진행ID로 이음)
  let progCache = null;
  function loadProg() {
    if (progCache) return progCache;
    progCache = (async () => {
      const { files } = await scanFiles();
      const spec = CFG.lists.find((l) => l.fileRegex || l.fileKeyword);
      const out = { rows: [] };
      await Promise.all(pickSourceFiles(files, spec).map(async (f) => {
        const wb = await readBook(f.id);
        const rows = sheetRows(wb, "계약진행");
        if (!rows) return;
        const h = headIndex(rows, "진행ID");
        if (h < 0) return;
        const head = rows[h].map((c) => String(c).trim());
        rows.slice(h + 1).forEach((r) => {
          const o = {};
          head.forEach((k, j) => { if (k) o[k] = String(r[j] ?? "").trim(); });
          if (o["진행ID"] && o["고객ID"]) out.rows.push({ ...o, _fc: f.fc, _fileId: f.id });
        });
      }));
      return out;
    })();
    progCache.catch(() => { progCache = null; });
    return progCache;
  }
  const PG_STAGES = [["고객등록", "new"], ["전화", "send"], ["AP", "resp"], ["PC", "cpn"], ["계약", "done"]];
  const PG_VIEWS = ["진행", "보관", "완료"]; // v411(2026-11-18 사용자 지시): '삭제'를 '완료'로 합침 — 예전에 삭제로 둔 줄(상태=삭제)도 완료에 보임
  const pgView = (r) => (r.status === "삭제" ? "완료" : r.status);
  function progRows(data, fc) {
    const out = [];
    ((data.prog && data.prog.rows) || []).forEach((p) => {
      if (fc && p._fc !== fc) return;
      const id = p["고객ID"], ents = data.custById.get(id) || [], ent = ents.find((e) => e.fc === p._fc) || ents[0] || null;
      out.push({
        pid: p["진행ID"], p, custId: id, fc: p._fc, fileId: p._fileId, ents, name: ent ? ent.name : id, route: ent ? routeShow(fieldOf(ent, "최초유입경로")) : "", from: p["출발"] || "",
        stage: p["단계"] || "고객등록", status: p["상태"] || "진행", start: p["시작일"] || "", stageDay: p["단계일"] || "", end: p["종료일"] || "", prod: p["설계상품"] || ""
      });
    });
    out.sort((a, b) => dateKey(b.start || "0").localeCompare(dateKey(a.start || "0")) || a.name.localeCompare(b.name, "ko")); // 시작일 최신 먼저(시작일 모르는 줄은 아래), 같은 날은 이름순
    return out;
  }
  // ---------- 영업 고객별 메모 칸 (v411, 2026-11-18 사용자 지시) ----------
  // 설정 > 화면의 [영업 메모 칸]을 켜면 세일즈(이달목표·계약진행)의 고객마다 맨 아래에 최근 메모(상담기록)와 한 줄 입력칸이 나옴.
  // 저장하면 일정 화면의 한 줄 메모처럼 고객추가정보의 '상담기록'에 들어가 고객 정보 > 상담 탭에서도 보임. 켜고 끔은 내 설정 파일(드라이브)에 저장
  const smemoOn = () => true; // v420(2026-11-18 사용자 지시): 설정 옵션을 없애고 항상 사용
  const SMEMO = { x: null, at: 0, map: new Map() };
  async function smemoPrep() { // 세일즈 화면을 그리기 전에 상담기록 파일을 한 번 읽어 둠(1분 안에는 다시 읽지 않음)
    if (!smemoOn()) return;
    if (SMEMO.x && Date.now() - SMEMO.at < 60000) return;
    try { SMEMO.x = await readExtra(); } catch (e) { SMEMO.x = null; }
    SMEMO.at = Date.now(); SMEMO.map = new Map();
  }
  const smemoFc = (custId, fb) => { const ents = (SALES.data && SALES.data.custById && SALES.data.custById.get(custId)) || []; return (ents.length && memoFcOf({ entries: ents })) || fb || ""; };
  function smemoList(fc, name) {
    const k = fc + "|" + name;
    if (!SMEMO.map.has(k)) SMEMO.map.set(k, SMEMO.x ? readConsults(SMEMO.x.wb, fc, name) : []);
    return SMEMO.map.get(k);
  }
  const SMEMO_OPEN = new Set(); // 입력칸이 열려 있는 고객(이름|FC) — 다시 그려도 유지
  const SMEMO_C = new Map(); // 일정 화면: 이름|FC → 고객 카드(저장 뒤 c.extra.consults를 같이 고치려고)
  function smemoHtml(name, fc, ro, given, max) {
    max = max || 1; // 일정 화면은 5: 닫혀 있을 때 최근 상담기록을 날짜와 함께 최대 5줄
    if (!smemoOn() || !name) return "";
    const list = given || smemoList(fc, name), last = list[0], k = fc + "|" + name;
    const md = (d) => esc(String(d || "").slice(5, 10));
    const keyAgo = (m) => { const d = new Date(); d.setMonth(d.getMonth() - m); return (d.getFullYear() * 10000 + (d.getMonth() + 1) * 100 + d.getDate()) * 10000; };
    const recentN = list.filter((r) => r.key >= keyAgo(12)).length; // [더보기]는 최근 1년 안에 2건 이상일 때만
    const six = list.filter((r) => r.key >= keyAgo(6)); // v421(2026-11-18 사용자 지시): 펼치면 최근 6개월 상담기록
    const attrOf = (r) => `data-smname="${esc(name)}" data-smfc="${esc(fc)}" data-smk="${r.key}" data-smo="${esc(r.text)}"`;
    const btns = (r) => ro ? "" : `<button type="button" class="sm-mini" data-smedit ${attrOf(r)}>수정</button><button type="button" class="sm-mini del" data-smdel ${attrOf(r)}>삭제</button>`;
    const editRow = (r) => ro ? "" : `<div class="sm-edit" hidden><input type="text" maxlength="200" value="${esc(r.text)}" autocomplete="off" enterkeyhint="done" data-smeditin aria-label="${esc(name)} 상담기록 수정"><button type="button" class="sm-save" data-smeditsave ${attrOf(r)}>저장</button><button type="button" class="sm-mini" data-smeditcancel>취소</button></div>`;
    // 닫혀 있을 때: 최근 상담기록 한 줄 (펼치면 아래 목록에 들어 있으므로 숨김)
    const moreBtn = `<button type="button" class="sm-more" data-smmore="${esc(name)}" data-smfc="${esc(fc)}" aria-label="${esc(name)} 상담기록 ${list.length}건 모두 보기">더보기</button>`;
    const top = list.slice(0, max).map((r, n, a) => `<div class="sm-rec top"><div class="sm-rv sm-last"><span class="sm-d">${md(r.date)}</span><span class="sm-t">${esc(r.text)}</span>${btns(r)}${n === a.length - 1 && (max > 1 ? list.length > max : recentN > 1) ? moreBtn : ""}</div>${editRow(r)}</div>`).join("");
    const add = ro ? "" : `<div class="sm-in"><input type="text" maxlength="200" placeholder="상담기록 추가" autocomplete="off" enterkeyhint="done" data-smin aria-label="${esc(name)} 상담기록 추가"><button type="button" class="mic-btn" data-smmic aria-label="음성으로 입력(녹음)"><svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="9" y="3" width="6" height="11" rx="3"/><path d="M5 11a7 7 0 0 0 14 0M12 18v3"/></svg></button><button type="button" class="sm-save" data-smsave="${esc(name)}" data-smfc="${esc(fc)}" disabled>저장</button></div>`;
    const recs = six.length ? `<div class="sm-recs">${six.map((r) => `<div class="sm-rec"><div class="sm-rv li"><span class="sm-d">${esc(String(r.date || "").slice(0, 10))}</span><span class="sm-tx">${esc(r.text).replace(/\n/g, "<br>")}</span>${btns(r)}</div>${editRow(r)}</div>`).join("")}</div>` : ""; // v422: 기록이 없을 때 안내문구는 두지 않음
    const older = list.length - six.length;
    const more = older > 0 ? `<button type="button" class="sm-more wide" data-smmore="${esc(name)}" data-smfc="${esc(fc)}">6개월 이전 ${older}건 더보기</button>` : "";
    const open = SMEMO_OPEN.has(k);
    return `<div class="sline smemo${open ? " open" : ""}">${top}<div class="sm-slide${open ? " open" : ""}" data-smkey="${esc(k)}"><div class="sm-slide-i">${add}${recs}${more}</div></div></div>`;
  }
  // 저장하며 돌려받은 최신 목록을 바로 쓰고(파일을 다시 읽지 않음), 영업은 화면을 다시 그리고 일정은 그 고객의 상담기록 칸만 다시 그림
  function smemoAfter(sm, fc, name, list) {
    const k = fc + "|" + name;
    SMEMO.map.set(k, list); SMEMO.at = Date.now();
    const c = SMEMO_C.get(k); if (c) c.extra = { ...(c.extra || {}), consults: list };
    if (sm && sm.closest("#scheduleView")) { sm.outerHTML = smemoHtml(name, fc, isRO(fc), list, 5); return; }
    renderSales(true);
  }
  // v541(2026-11-21 사용자 지시): 고객에게 전화·문자를 시도하면 상담기록에 "<페이지>에서 전화 시도" / "…문자 시도"를 남김
  //   눌러서 전화·문자 앱이 열리는 동안 저장이 끊기지 않게 기기 대기열에 먼저 넣고, 앱으로 돌아오면(또는 3초 뒤) 저장
  const cqKey = () => "shlife.consultQueue." + me();
  const cqRead = () => { try { return JSON.parse(localStorage.getItem(cqKey()) || "[]") || []; } catch (e) { return []; } };
  const cqWrite = (q) => { try { localStorage.setItem(cqKey(), JSON.stringify(q)); } catch (e) { /* 무시 */ } };
  function logAttempt(page, kind, name, fc) {
    name = String(name || "").trim(); if (!name || !fc || isRO(fc)) return;
    const q = cqRead(); q.push({ uid: Date.now() + "-" + Math.random().toString(36).slice(2, 7), fc, name, text: `${page}에서 ${kind === "sms" ? "문자" : "전화"} 시도`, at: Date.now() }); cqWrite(q);
    setTimeout(flushConsultQ, 3000);
  }
  let cqBusy = false;
  async function flushConsultQ() {
    if (cqBusy || !account()) return; cqBusy = true;
    try {
      for (const it of cqRead()) {
        try { await saveConsults(it.fc, it.name, [{ text: it.text, when: new Date(it.at) }], []); cqWrite(cqRead().filter((x) => x.uid !== it.uid)); }
        catch (e) { console.warn("상담기록(시도) 저장 실패, 다음에 다시", e); if (/NO_EXTRA_FILE/.test(String(e && e.message))) cqWrite(cqRead().filter((x) => x.uid !== it.uid)); break; }
      }
    } finally { cqBusy = false; }
  }
  document.addEventListener("visibilitychange", () => { if (document.visibilityState === "visible") setTimeout(flushConsultQ, 2000); });
  window.addEventListener("pageshow", () => setTimeout(flushConsultQ, 3000));
  const attemptPage = () => { // 지금 화면 이름
    const v = document.documentElement.dataset.view;
    if (v === "salesView") return SALES.sub === "material" ? "보고서" : SALES.sub === "data" ? "자료" : SALES.sub === "prog" ? "계약진행" : "이달목표 " + ({ transfer: "이관", dda: "따행", valueup: "DB" }[SALES.tab] || "");
    return { scheduleView: "일정", searchView: "고객 목록", customerView: "고객 정보", dashView: "대시보드" }[v] || "앱";
  };
  document.addEventListener("click", (ev) => { // 확인창이 없는 전화·문자 링크(고객 정보·고객 목록·일정 등). 세일즈 이달목표 버튼·연체/실효 리스트는 확인창 뒤에 따로 남김
    const a = ev.target.closest && ev.target.closest('a[href^="tel:"], a[href^="sms:"]');
    if (!a || a.closest("[data-contact]") || a.hasAttribute("data-odtel") || a.hasAttribute("data-contact")) return;
    const kind = /^sms:/.test(a.getAttribute("href")) ? "sms" : "call", box = a.closest("[data-cn]");
    let name = box ? box.dataset.cn : "", fc = box ? box.dataset.cfc : "";
    if (!name && current && !$("customerView").hidden) { name = current.name; fc = memoFcOf(current); }
    logAttempt(attemptPage(), kind, name, fc);
  }, true);
  async function smemoSave(btn) {
    const box = btn.closest(".smemo"), inp = box && box.querySelector("[data-smin]");
    const text = inp ? inp.value.trim() : "", name = btn.dataset.smsave, fc = btn.dataset.smfc;
    if (!text || !name) return;
    if (isRO(fc)) return toast("보기 전용 권한이라 저장할 수 없습니다.");
    micStop();
    btn.disabled = true; btn.textContent = "저장 중"; inp.disabled = true;
    try {
      const list = await saveConsults(fc, name, [{ text, when: new Date() }], []);
      inp.value = "";
      toast("상담기록에 저장했습니다.");
      smemoAfter(box, fc, name, list);
    } catch (e) { btn.disabled = false; btn.textContent = "저장"; inp.disabled = false; saveError(e); }
  }
  document.addEventListener("input", (ev) => { const m = ev.target.closest && ev.target.closest(":is(#salesView, #scheduleView) [data-smin]"); if (!m) return; const b = m.parentElement.querySelector("[data-smsave]"); if (b && b.textContent === "저장") b.disabled = !m.value.trim(); });
  document.addEventListener("keydown", (ev) => { const m = ev.target.closest && ev.target.closest(":is(#salesView, #scheduleView) [data-smin]"); if (m && ev.key === "Enter" && !ev.isComposing) { ev.preventDefault(); const b = m.parentElement.querySelector("[data-smsave]"); if (b && !b.disabled) b.click(); } });
  document.addEventListener("click", (ev) => {
    const b = ev.target.closest && ev.target.closest(":is(#salesView, #scheduleView) [data-smsave]"); if (b && !b.disabled) smemoSave(b);
    const m = ev.target.closest && ev.target.closest(":is(#salesView, #scheduleView) [data-smmore]"); if (m) smemoView(m.dataset.smmore, m.dataset.smfc);
    const ed = ev.target.closest && ev.target.closest(":is(#salesView, #scheduleView) [data-smedit]"); if (ed) { const rec = ed.closest(".sm-rec"); rec.querySelector(".sm-rv").hidden = true; const e = rec.querySelector(".sm-edit"); e.hidden = false; const i = e.querySelector("input"); i.focus(); i.setSelectionRange(i.value.length, i.value.length); }
    const ec = ev.target.closest && ev.target.closest(":is(#salesView, #scheduleView) [data-smeditcancel]"); if (ec) { const rec = ec.closest(".sm-rec"); rec.querySelector(".sm-edit").hidden = true; rec.querySelector(".sm-rv").hidden = false; }
    const es = ev.target.closest && ev.target.closest(":is(#salesView, #scheduleView) [data-smeditsave]"); if (es && !es.disabled) smemoEditSave(es);
    const dl = ev.target.closest && ev.target.closest(":is(#salesView, #scheduleView) [data-smdel]"); if (dl && !dl.disabled) smemoDelete(dl);
    const mc = ev.target.closest && ev.target.closest(":is(#salesView, #scheduleView) [data-smmic]"); if (mc) dictate(mc.parentElement.querySelector("[data-smin]"), mc);
  });
  const smemoWhen = (key) => { const day = Math.floor(key / 10000), hm = key % 10000; return new Date(Math.floor(day / 10000), (Math.floor(day / 100) % 100) - 1, day % 100, Math.floor(hm / 100), hm % 100); }; // 수정해도 처음 날짜·시간 그대로
  async function smemoEditSave(btn) {
    const box = btn.closest(".sm-rec"), sm = btn.closest(".smemo"), inp = box.querySelector("[data-smeditin]");
    const text = inp.value.trim(), name = btn.dataset.smname, fc = btn.dataset.smfc, key = +btn.dataset.smk, old = btn.dataset.smo;
    if (!text) return toast("내용을 입력해 주세요.");
    if (text === old) { box.querySelector(".sm-edit").hidden = true; box.querySelector(".sm-rv").hidden = false; return; }
    if (isRO(fc)) return toast("보기 전용 권한이라 수정할 수 없습니다.");
    btn.disabled = true; btn.textContent = "저장 중"; inp.disabled = true;
    try {
      const list = await saveConsults(fc, name, [{ text, when: smemoWhen(key) }], [{ key, text: old }]);
      toast("상담기록을 수정했습니다.");
      smemoAfter(sm, fc, name, list);
    } catch (e) { btn.disabled = false; btn.textContent = "저장"; inp.disabled = false; saveError(e); }
  }
  async function smemoDelete(btn) {
    const sm = btn.closest(".smemo"), name = btn.dataset.smname, fc = btn.dataset.smfc, key = +btn.dataset.smk, old = btn.dataset.smo;
    if (isRO(fc)) return toast("보기 전용 권한이라 삭제할 수 없습니다.");
    if (!confirm(`${name} 고객의 이 상담기록을 삭제할까요?\n\n${old}\n\n(엑셀에서 지워지며, 되돌릴 수 없어요)`)) return;
    btn.disabled = true; btn.textContent = "삭제 중";
    try {
      const list = await saveConsults(fc, name, [], [{ key, text: old }]);
      toast("상담기록을 삭제했습니다.");
      smemoAfter(sm, fc, name, list);
    } catch (e) { btn.disabled = false; btn.textContent = "삭제"; saveError(e); }
  }
  // v438(2026-11-18 사용자 지시): 일정 카드의 [할 일로]·[영업/비영업]·[수정]·[삭제]는 평소엔 접어 두고, 일정 칸(시간·제목 줄)을 터치하면 한 줄 아래 오른쪽 끝으로 펼침
  document.addEventListener("click", (ev) => {
    const card = ev.target.closest && ev.target.closest("#dayList .ev-card");
    if (!card || !ev.target.closest(".ev-top, .ev-line2") || ev.target.closest("a, button, input, select, textarea, label") || card.classList.contains("editing")) return;
    const r = card.querySelector(".ev-btnrow"); if (r) r.hidden = !r.hidden;
  });
  // 일정 화면: 고객 카드의 정보 칸(버튼·링크가 아닌 곳)을 터치하면 같은 방식으로 펼침/접힘
  document.addEventListener("click", (ev) => {
    const cu = ev.target.closest && ev.target.closest("#scheduleView .ev-cust");
    if (!cu || ev.target.closest("a, button, input, label, select, .smemo")) return;
    const sl = cu.querySelector(".sm-slide"); if (!sl) return;
    const on = sl.classList.toggle("open"), k = sl.dataset.smkey;
    sl.closest(".smemo").classList.toggle("open", on);
    if (on) SMEMO_OPEN.add(k); else SMEMO_OPEN.delete(k);
  });
  // 고객 줄(정보 칸)을 터치하면 상담기록 입력칸이 아래로 내려오고, 한 번 더 터치하면 올라가며 사라짐 (버튼·링크·입력칸을 누른 것은 제외)
  document.addEventListener("click", (ev) => {
    const row = ev.target.closest && ev.target.closest("#salesView .srow");
    if (!row || ev.target.closest("a, button, input, label, select, .sbtns, .smemo")) return;
    const sl = row.querySelector(".sm-slide"); if (!sl) return;
    const on = sl.classList.toggle("open"), k = sl.dataset.smkey;
    sl.closest(".smemo").classList.toggle("open", on);
    if (on) SMEMO_OPEN.add(k); else SMEMO_OPEN.delete(k);
  });
  function smemoView(name, fc) { // 상담기록 전체 보기(보기 전용 팝업): 바깥·[닫기]·뒤로(Esc)로 닫음
    const old = document.getElementById("smemoPop"); if (old) old.remove();
    const list = smemoList(fc, name);
    const box = document.createElement("div");
    box.id = "smemoPop"; box.className = "modal sm-pop";
    box.innerHTML = `<div class="modal-box" role="dialog" aria-modal="true" aria-label="${esc(name)} 상담기록"><div class="sm-ph"><h3 class="sec-title">${esc(name)} 상담기록</h3><span class="hint">${list.length}건 · 보기 전용</span></div>
      <ol class="sm-list">${list.map((r) => `<li><div class="sm-d">${esc(r.date)}</div><div class="sm-tx">${esc(r.text).replace(/\n/g, "<br>")}${r.voice ? ' <small class="hint">(녹음 있음)</small>' : ""}</div></li>`).join("")}</ol>
      <div class="file-actions"><button type="button" class="secondary" data-smclose>닫기</button></div></div>`;
    document.body.appendChild(box);
    const close = () => { box.remove(); document.removeEventListener("keydown", onKey); };
    const onKey = (e) => { if (e.key === "Escape") close(); };
    box.addEventListener("click", (e) => { if (e.target === box || e.target.closest("[data-smclose]")) close(); });
    document.addEventListener("keydown", onKey);
  }
  function progRowsHtml(rows, view) {
    const mdd = (v) => (v ? esc(String(v).slice(5, 10)) : "");
    return `<div class="slist">` + rows.map((r, i) => {
      const ro = isRO(r.fc), idx = PG_STAGES.findIndex(([lab]) => lab === r.stage);
      const steps = PG_STAGES.map(([lab, cls], k) => `<button type="button" class="dc ${cls} ${k === idx ? "on" : k < idx ? "off past" : "off"}${ro ? " lock" : ""}" data-pgs="${esc(r.pid)}" data-v="${lab}" aria-pressed="${k === idx ? "true" : "false"}">${lab}</button>`).join("");
      const sep = `<i class="ar-sep" aria-hidden="true"></i>`;
      const acts = ro ? "" : view === "진행" ? `${sep}<button type="button" class="ar-btn back" data-pgst="${esc(r.pid)}" data-to="보관">보관으로</button>`
        : view === "보관" ? `${sep}<button type="button" class="ar-btn back" data-pgst="${esc(r.pid)}" data-to="진행">진행으로</button><button type="button" class="ar-btn go" data-pgst="${esc(r.pid)}" data-to="완료">완료로</button>`
          : `${sep}<button type="button" class="ar-btn back" data-pgst="${esc(r.pid)}" data-to="진행">진행으로</button><button type="button" class="ar-btn go" data-pgst="${esc(r.pid)}" data-to="보관">보관으로</button>`;
      const meta = [r.start ? `시작 ${mdd(r.start)}` : "시작일 모름", r.stageDay ? `단계 ${mdd(r.stageDay)}` : "", r.end && view !== "진행" ? `${view} ${mdd(r.end)}` : "", r.from ? `출발 ${esc(r.from)}` : ""].filter(Boolean).join(" · ");
      return `<div class="srow prow"><div class="sline l1"><span class="no">${i + 1}</span><span class="nmw n50 ${gradeClass(gradeOf({ entries: r.ents })) || "grade-good"}">${r.custId ? `<a href="#" data-cust="${esc(r.custId)}">${esc(r.name)}</a>` : esc(r.name)}${r.route ? ` <small class="ins">${esc(r.route)}</small>` : ""}</span><span class="rt">${acts}</span></div>
        <div class="sline prow-steps">${steps}</div>
        <div class="sline l3"><span class="si">${meta}</span>${r.prod ? `<span class="si pd">${esc(r.prod)}</span>` : ""}</div>${smemoHtml(r.name, smemoFc(r.custId, r.fc), ro)}</div>`;
    }).join("") + `</div>`;
  }
  function renderProgPage(box, data, fcSel, userHtml, fcHtml) {
    const all = progRows(data, fcSel), view = SALES.pgv && PG_VIEWS.includes(SALES.pgv) ? SALES.pgv : "진행";
    const cnt = Object.fromEntries(PG_VIEWS.map((v) => [v, all.filter((r) => pgView(r) === v).length]));
    const rows = all.filter((r) => pgView(r) === view);
    const bar = `<span class="prog-bar" role="group" aria-label="계약진행 보기">${PG_VIEWS.map((v) => `<button type="button" data-pgv="${v}" class="${v === view ? "on" : ""}">${v} ${cnt[v]}</button>`).join("")}</span>`;
    box.innerHTML = `<div class="sales-top"><div class="sline2 sl3${userHtml || fcHtml ? "" : " center"}"><div class="ymnav"><b>계약진행</b></div><span class="sl-rt">${userHtml}${fcHtml}</span></div></div>
      <div id="salesBody"><div class="sales-info"><span>${rows.length}명</span>${bar}</div>` +
      (rows.length ? progRowsHtml(rows, view) : `<p class="hint pad">${fcSel ? esc(fcSel) + " " : ""}계약진행 ${esc(view)} 고객이 없습니다.</p>`) +
      `<p class="hint pad">단계를 눌러 바꿉니다(건너뛰어도 됩니다). ‘계약’을 누르면 완료로 옮겨집니다. 바꿀 때마다 이력에 남습니다.</p></div>`;
    box._prog = rows;
  }
  // 단계·상태 바꾸기: 화면 먼저, 저장은 뒤에서 차례로, 실패하면 되돌림, 6초 안에 [되돌리기]
  function pgApply(row, nx, result) {
    const p = row.p, prev = { 단계: p["단계"] || "", 단계일: p["단계일"] || "", 상태: p["상태"] || "", 종료일: p["종료일"] || "" };
    const set = (v) => { p["단계"] = v.단계; p["단계일"] = v.단계일; p["상태"] = v.상태; p["종료일"] = v.종료일; row.stage = v.단계 || "고객등록"; row.status = v.상태 || "진행"; row.stageDay = v.단계일; row.end = v.종료일; };
    set(nx); renderSales(true);
    progBusy(1);
    progQ = progQ.then(() => saveProgRow(row, nx, result)).then(() => true, (err) => { set(prev); renderSales(true); handleError(err); return false; }).finally(() => progBusy(-1));
    return prev;
  }
  function progChange(row, kind, to) {
    if (isRO(row.fc)) return toast("보기 전용 권한이라 수정할 수 없습니다.");
    const p = row.p, today = todayStr(), nx = { 단계: p["단계"] || "", 단계일: p["단계일"] || "", 상태: p["상태"] || "진행", 종료일: p["종료일"] || "" };
    let result, msg;
    if (kind === "stage") {
      if ((p["단계"] || "고객등록") === to) return;
      nx.단계 = to; nx.단계일 = today; result = to; msg = `${row.name} 고객을 ‘${to}’ 단계로 바꿨습니다`;
      if (to === "계약") { nx.상태 = "완료"; nx.종료일 = today; } else if (nx.상태 === "완료") { nx.상태 = "진행"; nx.종료일 = ""; }
    } else {
      if ((p["상태"] || "진행") === to) return;
      nx.상태 = to; nx.종료일 = to === "진행" ? "" : today; result = to === "진행" ? "원복" : to; msg = `${row.name} 고객을 ${to === "진행" ? "진행으로 되돌렸습니다" : to + "(으)로 옮겼습니다"}`;
    }
    const prev = pgApply(row, nx, result);
    undoToast(msg, () => { pgApply(row, prev, "취소"); toast("되돌렸습니다."); });
  }
  async function saveProgRow(row, nx, result) {
    const wb = await stGetBook(row.fileId);
    const info = sheetInfo(wb, "계약진행", "진행ID");
    if (!info) throw new Error("계약진행 시트를 찾지 못했습니다.");
    const kc = info.colOf("진행ID");
    let r0 = -1;
    for (let r = info.hr + 1; r <= info.rg.e.r; r++) if (info.txt(r, kc) === row.pid) { r0 = r; break; }
    if (r0 < 0) throw new Error("계약진행 표에서 이 줄을 찾지 못했습니다.");
    const setc = (col, val) => { const c = ensureColumn(info, col); if (val === "") delete info.ws[info.enc(r0, c)]; else info.ws[info.enc(r0, c)] = { t: "s", v: val }; };
    setc("단계", nx.단계); setc("단계일", nx.단계일); setc("상태", nx.상태); setc("종료일", nx.종료일);
    const hist = sheetInfo(wb, "이력", "고객ID"), stamp = tsDigits(nowStamp());
    if (hist) appendRow(hist, { 일시: { stamp }, 고객ID: row.custId, 진행ID: row.pid, 담당FC: row.fc || "", 유형: "계약진행", 결과: result, 내용: "앱에서 변경 (세일즈 > 계약진행)", 입력일시: { stamp } });
    await stPutBook(row.fileId, wb);
    cacheBook(row.fileId, wb);
    progCache = null; listCache = null; histCache = null;
  }

  // ---- 기기 대기 기록 (누르면 전화·문자 앱으로 넘어가 웹앱이 끊길 수 있어 먼저 기기에 저장 → 앱이 다시 보일 때 Drive에 반영) ----
  const queueKey = () => "shlife.contactQueue." + me();
  const readQueue = () => { try { return JSON.parse(localStorage.getItem(queueKey()) || "[]"); } catch (e) { return []; } };
  const writeQueue = (q) => { try { localStorage.setItem(queueKey(), JSON.stringify(q)); } catch (e) { /* 무시 */ } };
  let flushing = false;

  // ---- 엑셀 편집 도우미 (최신 파일을 다시 읽어 바뀐 칸만) ----
  function sheetInfo(wb, name, keyCol) {
    const sn = wb.SheetNames.find((n) => n.trim() === name);
    if (!sn) return null;
    const ws = wb.Sheets[sn];
    if (!ws || !ws["!ref"]) return null;
    const enc = (r, c) => XLSX.utils.encode_cell({ r, c });
    const info = { ws, enc, rg: XLSX.utils.decode_range(ws["!ref"]), hr: -1 };
    info.txt = (r, c) => { const cell = ws[enc(r, c)]; return cell ? String(cell.w ?? cell.v).trim() : ""; };
    for (let r = info.rg.s.r; r <= Math.min(info.rg.e.r, info.rg.s.r + 10); r++) {
      for (let c = info.rg.s.c; c <= info.rg.e.c; c++) if (info.txt(r, c) === keyCol) { info.hr = r; break; }
      if (info.hr >= 0) break;
    }
    if (info.hr < 0) return null;
    info.colOf = (n) => { for (let c = info.rg.s.c; c <= info.rg.e.c; c++) if (info.txt(info.hr, c) === n) return c; return -1; };
    return info;
  }
  function shiftColumns(info, at) {
    const ws = info.ws, moves = [];
    Object.keys(ws).forEach((k) => { if (k[0] === "!") return; const a = XLSX.utils.decode_cell(k); if (a.c >= at) moves.push([k, a]); });
    moves.sort((x, y) => y[1].c - x[1].c).forEach(([k, a]) => { ws[XLSX.utils.encode_cell({ r: a.r, c: a.c + 1 })] = ws[k]; delete ws[k]; });
    info.rg.e.c += 1;
    ws["!ref"] = XLSX.utils.encode_range(info.rg);
    if (ws["!cols"]) ws["!cols"].splice(at, 0, { wch: 14 });
    if (ws["!merges"]) ws["!merges"].forEach((m) => { if (m.s.c >= at) m.s.c++; if (m.e.c >= at) m.e.c++; });
  }
  // 열이 없으면 만듦: after 열이 있으면 그 다음, 아니면 '입력일시' 앞(없으면 맨 끝)
  function ensureColumn(info, name, after) {
    let c = info.colOf(name);
    if (c >= 0) return c;
    const a = after ? info.colOf(after) : -1;
    const sc = info.colOf(STAMP_COL);
    const at = a >= 0 ? a + 1 : sc >= 0 ? sc : info.rg.e.c + 1;
    if (at <= info.rg.e.c) shiftColumns(info, at); else { info.rg.e.c = at; info.ws["!ref"] = XLSX.utils.encode_range(info.rg); }
    info.ws[info.enc(info.hr, at)] = { t: "s", v: name };
    return at;
  }
  const stampCell = (digits) => { // 'YYYYMMDDHHmm' → 엑셀 날짜·시간 칸
    const y = +digits.slice(0, 4), m = +digits.slice(4, 6), d = +digits.slice(6, 8), hh = +digits.slice(8, 10) || 0, mi = +digits.slice(10, 12) || 0;
    return { t: "n", v: toExcelDateTime(new Date(y, m - 1, d, hh, mi)), z: "yyyy-mm-dd hh:mm" };
  };
  function setRowCells(info, keyCol, keyVal, values) {
    const kc = info.colOf(keyCol);
    if (kc < 0) return false;
    let row = -1;
    for (let r = info.hr + 1; r <= info.rg.e.r; r++) if (info.txt(r, kc) === keyVal) { row = r; break; }
    if (row < 0) return false;
    Object.entries(values).forEach(([k, v]) => { const c = ensureColumn(info, k); if (tsDigits(info.txt(row, c)) > tsDigits(v)) return; /* 최근 일시만 남김 */ info.ws[info.enc(row, c)] = stampCell(tsDigits(v)); });
    return true;
  }
  function appendRow(info, values) { // values: { 열이름: 문자열 | {stamp:'YYYYMMDDHHmm'} }
    let last = info.hr;
    for (let r = info.rg.e.r; r > info.hr; r--) { let any = false; for (let c = info.rg.s.c; c <= info.rg.e.c; c++) if (info.ws[info.enc(r, c)]) { any = true; break; } if (any) { last = r; break; } }
    const row = last + 1;
    Object.entries(values).forEach(([k, v]) => {
      const c = ensureColumn(info, k, k === "계약ID" ? "고객ID" : undefined);
      info.ws[info.enc(row, c)] = v && v.stamp ? stampCell(v.stamp) : { t: "s", v: String(v) };
    });
    if (row > info.rg.e.r) info.rg.e.r = row;
    info.ws["!ref"] = XLSX.utils.encode_range(info.rg);
    return row;
  }
  // ---- 새 구조(v2) 쓰기: 새 칸이 있으면 새 칸에, 없으면 옛 칸에 (2026-11-17 데이터구조 확정안 1단계) ----
  const isV2 = (wb) => wb.SheetNames.some((n) => n.trim() === "계약진행"); // 계약진행 표가 있으면 새 구조 파일
  const STATE_NEW = { "이관보관": "이관상태", "따행보관": "따행상태" }; // 날짜 있음 = 보관, 비움 = 목표(이력에 날짜가 남음)
  const sttCells = (info, wb, col, value) => { // 이관보관·따행보관: 옛 칸(있으면)과 새 칸(이관상태·따행상태) 둘 다에 씀(2026-11-17 프팀장: 옛 칸을 없앨 때까지)
    const nw = STATE_NEW[col]; if (!nw) return [{ col, val: value }];
    const v2 = isV2(wb), out = [];
    if (info.colOf(col) >= 0 || !v2) out.push({ col, val: value });
    if (v2 || info.colOf(nw) >= 0) out.push({ col: nw, val: hellV(value) ? "삭제" : value ? "보관" : "목표" });
    return out.map((t) => (t.col === col ? { col, val: hellDay(value) } : t)); // 옛 칸에는 날짜만
  };
  const hellV = (v) => /^D:/.test(String(v || "")); // v416: 지옥(=저장 값 '삭제')
  const hellDay = (v) => (hellV(v) ? String(v).slice(2) : String(v || ""));
  function progressWrite(wb, e, key, v) { // 영업단계·진행제외 → 계약진행 표의 이 고객 가장 최근 줄(없으면 새 줄) + 이력 '계약진행' 한 줄
    const info = sheetInfo(wb, "계약진행", "진행ID"); if (!info) return;
    const cid = idOf(e), today = todayStr(), stamp = tsDigits(nowStamp()), hi = sheetInfo(wb, "이력", "고객ID");
    const kc = info.colOf("고객ID"), sc = info.colOf("시작일");
    let r0 = -1, best = "";
    for (let r = info.hr + 1; r <= info.rg.e.r; r++) if (info.txt(r, kc) === cid) { const d = sc >= 0 ? info.txt(r, sc) : ""; if (r0 < 0 || d >= best) { r0 = r; best = d; } }
    const setc = (r, col, val) => { const c = ensureColumn(info, col); if (val === "") delete info.ws[info.enc(r, c)]; else info.ws[info.enc(r, c)] = { t: "s", v: val }; };
    const getc = (r, col) => { const c = info.colOf(col); return c < 0 ? "" : info.txt(r, c); };
    const route = fieldOf(e, "최초유입경로"), from = route === "이관" ? "이관" : route === "DB" ? "DB" : "신규";
    let pid = r0 >= 0 ? getc(r0, "진행ID") : "", result = "";
    let nRows = 0; for (let r = info.hr + 1; r <= info.rg.e.r; r++) if (info.txt(r, kc) === cid) nRows++; // 진행ID 형식(데이터수집): 'J' + 고객ID 숫자 + '-회차'(예: J00829-1)
    const mk = (stage, status, end) => { pid = "J" + cid.replace(/\D/g, "") + "-" + (nRows + 1); appendRow(info, { 진행ID: pid, 고객ID: cid, 출발: from, 시작일: today, 단계: stage, 단계일: today, 상태: status, ...(end ? { 종료일: end } : {}), 입력일시: { stamp } }); };
    if (key === "영업단계") {
      const stage = { "통화": "전화", "AP": "AP", "PC": "PC", "계약": "계약", "": "고객등록" }[v];
      if (!stage) return;
      const status = stage === "계약" ? "완료" : "진행";
      if (r0 < 0) { if (!v) return; mk(stage, status, status === "완료" ? today : ""); }
      else { setc(r0, "단계", stage); setc(r0, "단계일", today); setc(r0, "상태", status); setc(r0, "종료일", status === "완료" ? today : ""); }
      result = stage;
    } else if (key === "진행제외") {
      if (r0 < 0) { if (!v) return; mk("고객등록", "보관", v); }
      else if (v) { setc(r0, "상태", "보관"); setc(r0, "종료일", v); }
      else { setc(r0, "상태", "진행"); setc(r0, "종료일", ""); }
      result = v ? "보관" : "원복";
    } else return;
    if (hi) appendRow(hi, { 일시: { stamp }, 고객ID: cid, 진행ID: pid, 담당FC: e.fc || "", 유형: "계약진행", 결과: result, 내용: "앱에서 변경", 입력일시: { stamp } });
  }
  const juminOf = (b, back) => { b = String(b || "").replace(/\D/g, ""); back = String(back || "").replace(/\D/g, ""); if (b.length === 8 && back.length >= 7) return b.slice(2) + back.slice(0, 7); if (b.length === 8 && back.length >= 1) return b.slice(2) + back[0]; if (b.length === 8) return b; return back.slice(0, 1); };
  // writeEntryCells용: 값 하나를 쓸 곳들 [[열, 값], …]. 옛 칸과 새 칸이 같이 있으면 둘 다에 쓴다(옛 칸은 사라질 때까지)
  function writeTargets(wb, hasCol, k, v, e) {
    const v2 = isV2(wb), old = hasCol(k) || !v2, out = [];
    const nw = STATE_NEW[k];
    if (nw) { if (old) out.push([k, v]); if (v2 || hasCol(nw)) out.push([nw, v ? "보관" : "목표"]); return out; }
    if (k === "최근통화일시" || k === "최근문자일시" || k === "진행" || k === "우선") return old ? [[k, v]] : []; // 새 구조에서 옛 칸이 없으면 이력 시트에만 남김
    if (k === "영업단계" || k === "진행제외") { if (v2) progressWrite(wb, e, k, v); return old ? [[k, v]] : []; }
    if (k === "DB상태") { out.push([k, v]); if (hasCol("DB제외")) { if (v === "보관") out.push(["DB제외", todayStr()]); else if (v === "신규") out.push(["DB제외", ""]); } return out; }
    if (k === "주민번호") { // v396: 옛 칸(생년월일·성별·후번호)이 아직 있는 파일에는 함께 맞춰 씀
      out.push([k, v]); const pj = parseJumin(v);
      if (hasCol("생년월일") && pj.birth) out.push(["생년월일", pj.birth]); if (hasCol("성별") && pj.sex) out.push(["성별", pj.sex]); if (hasCol("후번호")) out.push(["후번호", pj.back || ""]);
      return out;
    }
    if (k === "후번호") { if (old) out.push([k, v]); if (v2 || hasCol("주민번호")) out.push(["주민번호", juminOf(fieldOf(e, "생년월일"), v)]); return out; }
    return [[k, v]];
  }
  function histHas(info, v) { // 같은 일시·고객ID·계약ID·유형이 이미 있으면 true (재시도 중복 방지)
    const cs = { t: info.colOf("일시"), u: info.colOf("고객ID"), k: info.colOf("계약ID"), y: info.colOf("유형") };
    if (cs.t < 0 || cs.u < 0) return false;
    for (let r = info.hr + 1; r <= info.rg.e.r; r++) {
      if (tsDigits(info.txt(r, cs.t)) === tsDigits(v.일시) && info.txt(r, cs.u) === v.고객ID && (cs.k < 0 ? "" : info.txt(r, cs.k)) === (v.계약ID || "") && (cs.y < 0 ? "" : info.txt(r, cs.y)) === v.유형) return true;
    }
    return false;
  }
  async function writeContacts(fileId, items) {
    const wb = await stGetBook(fileId); // 최신 파일
    const stamp = tsDigits(nowStamp());
    const hist = sheetInfo(wb, "이력", "고객ID");
    const ct = sheetInfo(wb, "계약", "계약ID");
    const cu = sheetInfo(wb, CUST_SHEET, ID_COL);
    for (const it of items) {
      const colName = it.kind === "call" ? "최근통화일시" : "최근문자일시";
      const v = { 일시: it.ts, 고객ID: it.custId || "", 계약ID: it.cid || "", 유형: it.kind === "call" ? "통화" : "문자" };
      if (hist && !histHas(hist, v)) appendRow(hist, { 일시: { stamp: tsDigits(it.ts) }, 고객ID: v.고객ID, 계약ID: v.계약ID, 담당FC: it.fc || "", 유형: v.유형, 결과: "시도", 내용: it.src || "세일즈 화면", 입력일시: { stamp } });
      if (it.cid && ct && (ct.colOf(colName) >= 0 || !isV2(wb))) setRowCells(ct, "계약ID", it.cid, { [colName]: it.ts }); // 새 구조는 이력 시트에만 남김
      if (!it.cid && it.custId && cu && (cu.colOf(colName) >= 0 || !isV2(wb))) setRowCells(cu, ID_COL, it.custId, { [colName]: it.ts });
    }
    await stPutBook(fileId, wb);
    delete bookCache[fileId];
  }
  // ===== v550(2026-11-22 사용자 "이 안대로") 데이터 수집 중 잠금 =====
  //   PC가 FC 폴더에 .수집중.lock 을 두는 동안: 맨 위 경고 띠, 그 FC data.json 고치기는 막음(jsonPutDoc), 상담기록 수정·삭제도 막음
  //   새 기록(전화·문자 시도 이력)은 그 FC 폴더 _앱대기\YYYYMMDD-HHMMSS-<계정>.json 에 쌓았다가, 잠금이 풀리면 앱이 붙이고 그 파일을 지움(PC와 합의한 모양)
  //   상담기록 추가는 고객추가정보(PC가 쓰지 않는 파일)라 그대로 저장
  const SYNC = { parents: new Set(), fcs: new Set(), seen: new Set(), timer: null };
  const PEND_DIR = "_앱대기";
  function fcOfParent(parent) { // 그 폴더의 FC 이름(data.json 안 fc 값 → 권한 파일 FC 폴더 ID)
    for (const x of jsonInfo.values()) if (x.name === "data.json" && x.parent === parent && x.meta && x.meta.fc) return x.meta.fc;
    for (const [fc, id] of Object.entries(fcFolderIds || {})) if (id === parent) return fc;
    return "";
  }
  function syncMark(parent, on) {
    const had = SYNC.parents.has(parent);
    if (on) SYNC.parents.add(parent); else SYNC.parents.delete(parent);
    if (had !== on) syncBand();
  }
  const syncLockedFc = (fc) => { for (const p of SYNC.parents) if (fcOfParent(p) === fc) return true; return false; };
  function syncBand() {
    let el = $("syncBand");
    const on = SYNC.parents.size > 0;
    document.documentElement.classList.toggle("sync-lock", on);
    if (!on) { if (el) el.remove(); return; }
    const names = [...SYNC.parents].map(fcOfParent).filter(Boolean);
    if (!el) { el = document.createElement("div"); el.id = "syncBand"; el.className = "sync-band"; el.setAttribute("role", "status"); const main = document.querySelector("main"); if (main) main.prepend(el); }
    el.innerHTML = `<span class="dot" aria-hidden="true"></span><span><b>데이터 업데이트 중</b>${names.length ? " (" + esc(names.join(", ")) + ")" : ""} · 지금은 보기와 새 기록(전화·문자·상담기록 추가)만 됩니다<small>수정·완료·단계 변경은 업데이트가 끝난 뒤 다시 열립니다</small></span>`;
  }
  async function syncCheck() { // 읽어 둔 FC 폴더마다 수집중 잠금 확인 → 띠, 풀린 폴더는 _앱대기 붙이기
    if (!GD || !account() || document.visibilityState === "hidden") return;
    const parents = [...new Set([...jsonInfo.values()].filter((x) => x.name === "data.json" && x.parent).map((x) => x.parent))];
    for (const p of parents) {
      let on = false;
      try { on = (await lockList(p)).some((f) => f.name === LOCK_COLLECT && f.fresh); } catch (e) { continue; }
      syncMark(p, on);
      if (!on && !SYNC.seen.has(p)) { SYNC.seen.add(p); pendMerge(p).catch((e) => console.warn("대기 기록 붙이기 실패", e)); }
      if (on) SYNC.seen.delete(p); // 풀리면 다시 붙이기
    }
  }
  function syncStart() { if (SYNC.timer) return; setTimeout(syncCheck, 6000); SYNC.timer = setInterval(syncCheck, 60000); }
  document.addEventListener("visibilitychange", () => { if (document.visibilityState === "visible" && SYNC.timer) setTimeout(syncCheck, 1000); });
  async function pendDir(parent, make) {
    const ex = await gdFindIn(parent, PEND_DIR);
    if (ex || !make) return ex ? ex.id : "";
    const f = await gdFetch(GDAPI + "/files?supportsAllDrives=true&fields=id", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ name: PEND_DIR, mimeType: GD_FOLDER, parents: [parent] }) });
    return f.id;
  }
  async function pendWrite(parent, items) { // 전화·문자 시도 기록을 _앱대기 파일 하나로
    const dir = await pendDir(parent, true);
    const adds = items.map((it) => ({ sheet: "이력", row: { 일시: it.ts, 고객ID: it.custId || "", 계약ID: it.cid || "", 담당FC: it.fc || "", 유형: it.kind === "call" ? "통화" : "문자", 결과: "시도", 내용: it.src || "", 입력일시: nowStamp() }, item: it }));
    const body = JSON.stringify({ at: new Date().toISOString(), by: me(), adds });
    const name = tsName() + "-" + me().split("@")[0] + ".json", b = "shlife" + Date.now();
    const blob = new Blob(["--" + b + "\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n" + JSON.stringify({ name, parents: [dir], mimeType: "application/json" }) + "\r\n--" + b + "\r\nContent-Type: application/json\r\n\r\n" + body + "\r\n--" + b + "--"]);
    await gdFetch(GDUP + "/files?uploadType=multipart&supportsAllDrives=true&fields=id", { method: "POST", headers: { "Content-Type": "multipart/related; boundary=" + b }, body: blob });
  }
  async function pendMerge(parent) { // 잠금이 풀린 폴더의 _앱대기 파일을 data.json에 붙이고 지움(이력에 같은 일시·고객·유형이 있으면 건너뜀)
    const dir = await pendDir(parent, false); if (!dir) return;
    const files = (await gdQuery("'" + qv(dir) + "' in parents and trashed=false")).filter((f) => f.file && /\.json$/.test(f.name));
    if (!files.length) return;
    const dataId = [...jsonInfo.entries()].find(([, x]) => x.name === "data.json" && x.parent === parent);
    if (!dataId) return;
    let n = 0;
    for (const f of files) {
      let d; try { d = parseDoc(await stGet0(f.id)); } catch (e) { continue; }
      const items = (d.adds || []).filter((a) => a.sheet === "이력").map((a) => a.item || { uid: "p" + f.id + Math.random(), kind: a.row && a.row.유형 === "문자" ? "sms" : "call", ts: a.row.일시, cid: a.row.계약ID, custId: a.row.고객ID, fc: a.row.담당FC, src: a.row.내용 });
      if (items.length) { await writeContacts(dataId[0], items); n += items.length; }
      try { await gdFetch(GDAPI + "/files/" + encodeURIComponent(f.id) + "?supportsAllDrives=true", { method: "DELETE" }); } catch (e) { await stDel(f.id).catch(() => {}); }
    }
    if (n) { histCache = null; contractCache = null; toast(`업데이트 중에 남긴 전화·문자 기록 ${n}건을 저장했습니다.`); }
  }
  async function flushContacts() {
    if (flushing || !account()) return;
    flushing = true;
    try {
      const q = readQueue().filter((it) => it.fileId);
      if (!q.length) return;
      const byFile = new Map();
      q.forEach((it) => { if (!byFile.has(it.fileId)) byFile.set(it.fileId, []); byFile.get(it.fileId).push(it); });
      let wrote = false;
      for (const [fileId, items] of byFile) {
        try {
          const par = (jsonInfo.get(fileId) || {}).parent;
          if (par && SYNC.parents.has(par)) { await pendWrite(par, items); const done = new Set(items.map((i) => i.uid)); writeQueue(readQueue().filter((i) => !done.has(i.uid))); continue; } // v550: 수집 중이면 _앱대기로
          try { await writeContacts(fileId, items); }
          catch (e) { if (e && e.lockKind === LOCK_COLLECT && par) { await pendWrite(par, items); const done = new Set(items.map((i) => i.uid)); writeQueue(readQueue().filter((i) => !done.has(i.uid))); continue; } throw e; }
          addSent(items); // v509: 되돌린 뒤 다시 반영할 수 있게 14일 기억
          const done = new Set(items.map((i) => i.uid));
          writeQueue(readQueue().filter((i) => !done.has(i.uid)));
          wrote = true;
        } catch (e) { console.error("전화·문자 기록 반영 실패", e); }
      }
      if (wrote) { histCache = null; contractCache = null; listCache = null; }
      if (!$("salesView").hidden) renderSales(true);
    } finally { flushing = false; }
  }
  document.addEventListener("visibilitychange", () => { if (document.visibilityState === "visible") setTimeout(flushContacts, 1500); });
  window.addEventListener("pageshow", () => setTimeout(flushContacts, 2500));

  // ---- 데이터 → 목록 ----
  async function loadSalesData() {
    const [lists, contracts, hist, follow, prog, dbs] = await Promise.all([loadLists(), loadContracts(), loadHistory().catch(() => []), loadFollow().catch(() => ({ files: new Set(), rows: [] })), loadProg().catch(() => ({ rows: [] })), loadDbs().catch(() => ({ rows: [] }))]);
    const custById = new Map();
    lists.forEach((e) => { const id = idOf(e); if (id) (custById.get(id) || custById.set(id, []).get(id)).push(e); });
    return { lists, contracts, hist, custById, follow, prog, dbs };
  }
  // ym 달에 시도한 기록: key('c:계약ID' | 'u:고객ID') → { call: ts, sms: ts }
  function contactMarks(data, ym) {
    const m = new Map();
    const put = (key, kind, ts) => {
      if (!key || monthOf(ts) !== ym) return;
      const cur = m.get(key) || {};
      if (!cur[kind] || tsDigits(ts) > tsDigits(cur[kind])) cur[kind] = ts;
      m.set(key, cur);
    };
    data.hist.forEach((h) => { if (String(h["결과"] || "").trim() !== "시도") return; /* 앱의 '시도' 기록만 (실제 통화·상담 기록·취소(초기화)는 제외) */ const k = String(h["계약ID"] || "").trim() ? "c:" + String(h["계약ID"]).trim() : "u:" + String(h["고객ID"]).trim(); put(k, String(h["유형"]).trim() === "통화" ? "call" : String(h["유형"]).trim() === "문자" ? "sms" : "", h["일시"]); });
    data.contracts.forEach((c) => { const id = String(c.row["계약ID"] || "").trim(); if (id) { put("c:" + id, "call", c.row["최근통화일시"]); put("c:" + id, "sms", c.row["최근문자일시"]); } });
    data.lists.forEach((e) => { const id = idOf(e); if (id) { put("u:" + id, "call", fieldOf(e, "최근통화일시")); put("u:" + id, "sms", fieldOf(e, "최근문자일시")); } });
    readQueue().forEach((it) => put(it.cid ? "c:" + it.cid : "u:" + it.custId, it.kind, it.ts)); // 아직 Drive에 못 올린 기록도 바로 표시
    return m;
  }
  function transferRows(data, ym, showRef) {
    const out = [], stMode = data.contracts.some((c) => c.row && "이관상태" in c.row);
    data.contracts.forEach((c) => {
      if (c.type !== "이관") return;
      const rr = c.row, idate = rr["이관일자"] || rr["이관일"] || "";
      if (stMode) { const st = String(rr["이관상태"] || "").trim(); if (st !== "목표" && st !== "보관" && st !== "삭제") return; } // 목표 상태 기준(보관·삭제 포함해 보관 목록에서 봄)
      else if (monthOf(idate) !== ym) return;
      const ref = !!(c.fc && c.srcFc && c.fc !== c.srcFc); // 담당 FC가 이 파일의 FC가 아니면 참고
      if (ref && !showRef) return;
      const cid = String(rr["계약자ID"] || "").trim();
      const ce = cid && data.custById.get(cid) ? data.custById.get(cid).find((e) => e.fc === c.srcFc) || data.custById.get(cid)[0] : null;
      const holder = nameShow(rr["계약자"] || (ce && ce.name) || "");
      const insured = nameShow(rr["주피보험자"] || rr["피보"] || "");
      out.push({
        key: String(rr["계약ID"] || "").trim() ? "c:" + String(rr["계약ID"]).trim() : "", cid: String(rr["계약ID"] || "").trim(), custId: cid, fc: c.fc, srcFc: c.srcFc, ref,
        fileId: c.src && c.src.fileId, holder, ins: norm(insured) && norm(insured) !== norm(holder) ? insured : "",
        product: productName(rr["상품명"]), fee: rr["합계보험료"] || rr["보험료"] || "", cdate: rr["최초계약일"] || rr["계약일"] || "",
        bday: (ce && fieldOf(ce, "생년월일")) || rr["생일"] || "", phone: firstPhone((ce && (ce.phone || fieldOf(ce, "연락처"))) || rr["연락처"]),
        phoneText: (ce && (ce.phone || fieldOf(ce, "연락처"))) || rr["연락처"] || "", job: (ce && fieldOf(ce, "직업")) || rr["직업"] || "",
        addr: cleanAddr((ce && fieldOf(ce, "주소")) || rr["주소"]), recFc: rr["모집FC"] || "", status: rr["계약상태"] || rr["상태"] || "", idate,
        ent: ce && ce.sheet === CUST_SHEET ? ce : null, stage: ce && ce.sheet === CUST_SHEET ? stageOfPerson(data, cid, fieldOf(ce, "영업단계")) : undefined, // 영업단계 칩(통화·AP·PC·계약): 고객 시트의 '영업단계' 칸
        ct: c, archDay: String(rr["이관보관"] || "").trim(), // 보관(계약 단위): 계약 시트 '이관보관' 날짜
        del: stMode && String(rr["이관상태"] || "").trim() === "삭제", canHell: stMode, // v416: 이관상태 '삭제' = 지옥
        grp: stMode ? (monthOf(idate) === ymNow() ? 0 : 1) : undefined // 목표 기준 목록: 0 = 이번 달 이관, 1 = 지난 이관(보관·삭제 전)
      });
    });
    out.sort((a, b) => ((a.grp || 0) - (b.grp || 0)) || a.holder.localeCompare(b.holder, "ko") || dateKey(b.cdate).localeCompare(dateKey(a.cdate)));
    return out;
  }
  // 밸류업DB: 고객 시트의 최초유입경로가 'DB'인 고객 (담당 FC 행 기준, 월과 무관한 전체 목록). 카드는 이달이관·따행과 같은 모양
  // v515(2026-11-21 사용자 규칙 "다른 곳에서 연결되면 상태가 따라가게"): 영업단계 칩은 그 줄 파일에 값이 없으면 그 고객ID의 모든 FC 파일 계약진행 중 가장 최근 줄로(표시만, 쓰기는 그대로)
  function stageOfPerson(data, id, own) {
    if (own) return own;
    if (!data._stageBy) {
      const m = new Map(), key = new Map();
      ((data.prog && data.prog.rows) || []).forEach((r) => { const cid = r["고객ID"], k = (r["시작일"] || "") + "|" + (r["단계일"] || ""); if (cid && (!key.has(cid) || k >= key.get(cid))) { key.set(cid, k); m.set(cid, r["단계"] || ""); } });
      data._stageBy = m;
    }
    const st = data._stageBy.get(id) || "";
    return st && st !== "고객등록" && STAGE_NEW_OLD[st] != null ? STAGE_NEW_OLD[st] : own;
  }
  function dbSexAge(ssn) { // 주민번호(앞 6 + 뒤 첫 자리) → '남 45' (만 나이). 모르면 ""
    const m = /^(\d{2})(\d{2})(\d{2})-?([1-8])/.exec(String(ssn || "").trim()); if (!m) return "";
    const c = +m[4], y = (c === 1 || c === 2 || c === 5 || c === 6 ? 1900 : 2000) + +m[1], now = new Date();
    let age = now.getFullYear() - y; if (now.getMonth() + 1 < +m[2] || (now.getMonth() + 1 === +m[2] && now.getDate() < +m[3])) age--;
    return (c % 2 ? "남" : "여") + " " + age;
  }
  // v574(2026-11-24 사용자 지시 "이달목표 DB 리스트에 지역 … 유입 오른쪽에"): 보고서 '지역'과 같은 값(지역 칸, 없으면 주소 앞 두 낱말)
  const dbAreaOf = (r) => String((r.ent && (fieldOf(r.ent, "지역") || fieldOf(r.ent, "주소"))) || r.addr || "").replace(/^\[[^\]]*\]\s*/, "").split(/\s+/).slice(0, 2).join(" ").trim();
  function valueupRows(data) {
    const out = [], seen = new Set();
    const dbRows = (data.dbs && data.dbs.rows) || [], dbFcs = new Set(dbRows.map((r) => r._fc)); // DB 시트 줄이 있는 FC 파일 = 새 방식
    const perId = new Map(); dbRows.forEach((r) => perId.set(r["고객ID"], (perId.get(r["고객ID"]) || 0) + 1));
    dbRows.forEach((d) => {
      const id = d["고객ID"], mine = data.custById.get(id) || []; if (!mine.length) return;
      const own = mine.find((x) => x.fc === primaryFc(mine)) || mine[0], e = mine.find((x) => x.fc === d._fc && x.sheet === CUST_SHEET) || own;
      const st = String(d["DB상태"] || "").trim() || "신규";
      const cts = data.contracts.filter((c) => String(c.row["계약자ID"] || "").trim() === id && !DEAD_STATUS.includes(c.row["계약상태"] || c.row["상태"] || ""));
      out.push({
        key: perId.get(id) > 1 ? "d:" + d["DB ID"] : "u:" + id, cid: "", custId: id, fc: d._fc, srcFc: d["유입FC"] || d._fc, ref: false, fileId: own.src && own.src.fileId, holder: own.name, ins: "", ent: own, stage: stageOfPerson(data, id, fieldOf(own, "영업단계")),
        product: String(d["유입DB명"] || "").replace(/^[\s■□▪▫◼◾]+/, "") || "DB 고객", fee: "", cdLabel: "배분일", cdate: d["FC 배분일"] || fieldOf(e, "유입일"),
        comp: d["유입회사"] || "", inAtReal: d["유입일시"] || "", inAt: d["유입일시"] || d["FC 배분일"] || "", sexAge: dbSexAge(fieldOf(own, "주민번호")), dbMemo: d["유입메모"] || "",
        db: true, dbRow: d, dbAuto: cts.length > 0 || st === "계약" || d["계약체결"] === "Y", dbOffM: st === "보관", stCtr: st === "계약", del: st === "삭제", delDay: st === "삭제" ? d["완료일"] : "", dbDay: st === "보관" ? d["완료일"] : "",
        bday: fieldOf(own, "생년월일"), phone: firstPhone(own.phone || fieldOf(own, "연락처")), phoneText: own.phone || fieldOf(own, "연락처"), job: fieldOf(own, "직업"), addr: cleanAddr(fieldOf(own, "주소")), recFc: "", status: ""
      });
    });
    const arDays = new Map(), delDays = new Map(); // 고객ID → 가장 최근 '이동(보관)' / '삭제' 이력 일시
    (data.hist || []).forEach((h) => { const t = String(h["유형"] || "").trim(); const mp = t === "이동(보관)" ? arDays : t === "삭제" ? delDays : null; if (!mp) return; const id = String(h["고객ID"] || "").trim(), d = String(h["일시"] || "").trim(); if (id && d && (!mp.has(id) || dateKey(d) > dateKey(mp.get(id)))) mp.set(id, d); });
    data.lists.forEach((e) => {
      const id = idOf(e);
      if (!id || e.sheet !== CUST_SHEET || seen.has(id)) return;
      const mine = data.custById.get(id) || [e];
      const own = mine.find((x) => x.fc === primaryFc(mine)) || e;
      // v514(2026-11-21 사용자 승인 '유입FC'): DB를 받은 FC(유입FC)가 있으면 그 FC 파일의 줄이 DB 줄 → 유입FC의 DB 목록에 나오고, DB상태도 그 줄을 읽고 씀.
      //   영업단계·전화·문자·메모는 담당FC(메인FC) 줄로(진행은 다른 FC가 할 수 있음). 유입FC가 없거나 그 FC 파일에 줄이 없으면 예전처럼 담당FC 줄 기준
      const dbE = dbLineOf(mine);
      if (dbE ? dbE !== e : (own !== e || fieldOf(e, "최초유입경로") !== "DB")) return;
      if (dbFcs.has(e.fc) || DBS_BY_ID.has(id)) return; // v571: 이 FC는 DB 시트 기준(옛 칸 줄은 읽지 않음)
      seen.add(id);
      const cts = data.contracts.filter((c) => String(c.row["계약자ID"] || "").trim() === id && !DEAD_STATUS.includes(c.row["계약상태"] || c.row["상태"] || ""));
      out.push({
        key: "u:" + id, cid: "", custId: id, fc: e.fc, srcFc: e.fc, ref: false, fileId: own.src && own.src.fileId, holder: e.name, ins: "", ent: own, stage: stageOfPerson(data, id, fieldOf(own, "영업단계")), // 영업단계 칩(통화·AP·PC·계약)은 이관 목록과 같은 방식
        product: fieldOf(e, "유입DB명").replace(/^[\s■□▪▫◼◾]+/, "") || (cts.length ? "계약 " + cts.length + "건" : "DB 고객"), fee: "", cdLabel: "배분일", cdate: fieldOf(e, "유입일"), comp: fieldOf(e, "유입회사"), inAtReal: fieldOf(e, "유입일시"), inAt: fieldOf(e, "유입일시") || fieldOf(e, "유입일"), sexAge: fieldOf(e, "성별나이") || fieldOf(e, "성별/나이") || dbSexAge(fieldOf(e, "주민번호")), dbMemo: fieldOf(e, "유입메모") || fieldOf(e, "메모"), // v569(사용자 다시 정함): DB 줄 = 이름·연락처·유입회사·성별/나이·메모·유입일시. 새 칸(유입회사·유입일시·성별나이·유입메모)이 생기면 바로 읽고, 없으면 유입일·주민번호·메모로 대신
        db: true, dbAuto: cts.length > 0 || mine.some(dbAutoRec), dbOffM: dbState(e) === "보관", stCtr: dbState(e) === "계약", del: (dbE ? [e] : mine).some((x) => dbState(x) === "삭제"), delDay: delDays.get(id) || "", dbDay: fieldOf(e, "DB제외") || arDays.get(id) || "", // DB 보관: 'DB상태'=보관 또는 계약 고객(자동). 날짜는 이력 '이동(보관)' 줄(옛 DB제외 날짜가 있으면 그것)
        bday: fieldOf(e, "생년월일"), phone: firstPhone(e.phone || fieldOf(e, "연락처")), phoneText: e.phone || fieldOf(e, "연락처"),
        job: fieldOf(e, "직업"), addr: cleanAddr(fieldOf(e, "주소")), recFc: fieldOf(e, "원모집FC") || "", status: ""
      });
    });
    out.forEach((r) => { r.ctr = !r.del && (r.dbAuto || r.stCtr); r.dbAuto = r.ctr; r.dbOffM = r.dbOffM && !r.ctr; r.dbOff = r.del || r.ctr || r.dbOffM; r.archDay = r.del ? r.delDay : r.dbOffM ? r.dbDay : ""; });
    // v572(2026-11-24 사용자 결정 '김현성 DB+진행표시', 리더 전달): 배분받은 FC와 다른 FC가 계약진행(고객ID 연결)을 맡으면 DB 줄에 '진행: FC' 표시. 목록·건수·자동완료는 배분받은 FC 기준 그대로
    const pgBy = new Map();
    ((data.prog && data.prog.rows) || []).forEach((p) => { const id = p["고객ID"]; if (!id) return; const k = (p["상태"] === "진행" || !p["상태"] ? "1" : "0") + (p["시작일"] || "") + "|" + (p["단계일"] || ""); (pgBy.get(id) || pgBy.set(id, []).get(id)).push({ fc: p._fc, k }); });
    out.forEach((r) => { const c = (pgBy.get(r.custId) || []).filter((x) => x.fc && x.fc !== r.fc).sort((a, b) => b.k.localeCompare(a.k))[0]; r.pgFc = c ? c.fc : ""; });
    // 유입일 최신순, 같은 날은 이름 가나다순 (삭제 목록은 삭제한 날 최신순, 계약 자동은 유입일)
    // v569(2026-11-24 사용자 지시, 데이터정리 전달): 최신 유입일시가 맨 위, 같으면 이름 오름순(가나다). 유입일시가 빈 줄은 맨 아래에 FC 배분일(유입일) 최신 → 이름 오름순. 보관·지옥 목록은 예전처럼 옮긴 날 최신순
    const day = (r) => dateKey(r.dbOffM && r.dbDay ? r.dbDay : r.cdate);
    const nmDesc = (a, b) => a.holder.localeCompare(b.holder, "ko"); // 사용자 지시(같은 날): 이름은 오름순(가나다)
    out.sort((a, b) => {
      if ((a.dbOffM || a.del) || (b.dbOffM || b.del)) return day(b).localeCompare(day(a)) || nmDesc(a, b);
      const ia = String(a.inAtReal || ""), ib = String(b.inAtReal || "");
      if (!ia !== !ib) return ia ? -1 : 1; // 유입일시 있는 줄이 먼저
      return ib.localeCompare(ia) || day(b).localeCompare(day(a)) || nmDesc(a, b);
    });
    return out;
  }
  // 따행 대상(v384): 새 칸 '따행상태'가 있으면 그것('목표'면 대상, 빈칸이면 활동기한이 안 끝난 줄만 대상, 보관·삭제는 대상 아님), 없으면 옛 칸 '캠페인대상여부'
  const ddaTarget = (f) => {
    if (!f || !("따행상태" in f)) return !f || !f["캠페인대상여부"] || f["캠페인대상여부"] === "활동대상";
    const st = String(f["따행상태"] || "").trim();
    if (st === "목표") return true;
    if (st) return false;
    const to = String(f["활동기한종료일"] || "").slice(0, 10);
    return !!to && to >= todayStr();
  };
  function ddaRows(data, ym) {
    const out = [], seen = new Set(), follow = data.follow || { files: new Set(), rows: [] };
    const base = (e, id) => {
      const cts = data.contracts.filter((c) => String(c.row["계약자ID"] || "").trim() === id && !DEAD_STATUS.includes(c.row["계약상태"] || c.row["상태"] || ""));
      const prods = cts.map((c) => ({ name: productName(c.row["상품명"]), fee: c.row["합계보험료"] || c.row["보험료"] || "" })).filter((p) => p.name);
      const recFcs = [...new Set(cts.map((c) => String(c.row["모집FC"] || "").trim()).filter(Boolean))];
      const cd = cts.map((c) => c.row["최초계약일"] || c.row["계약일"] || "").sort((a, b) => dateKey(b).localeCompare(dateKey(a)))[0] || "";
      return {
        key: "u:" + id, cid: "", custId: id, fc: e.fc, srcFc: e.fc, ref: false, ent: e, fileId: e.src && e.src.fileId, holder: e.name, ins: "",
        prods, product: prods.length === 1 ? prods[0].name : prods.length ? "" : "계약 정보 없음(조회 제한)", // 계약 2건 이상이면 이름 아래 줄마다 상품명·보험료
        fee: prods.length === 1 ? prods[0].fee : "", cdate: cd, bday: fieldOf(e, "생년월일"), phone: firstPhone(e.phone || fieldOf(e, "연락처")), phoneText: e.phone || fieldOf(e, "연락처"),
        job: fieldOf(e, "직업"), addr: cleanAddr(fieldOf(e, "주소")), recFc: fieldOf(e, "원모집FC") || recFcs.join(", "), status: ""
      };
    };
    const ownOf = (id) => { const mine = data.custById.get(id); return mine ? (mine.find((x) => x.fc === primaryFc(mine)) || mine[0]) : null; };
    // B: 따행 시트가 있는 파일 — 선택한 대상년월의 행
    const foState = follow.rows.some((f) => "따행상태" in f); // 새 구조: 따행상태 '목표'(보관·삭제 포함) 기준, 달이 바뀌어도 남음
    const fRows = foState ? follow.rows.filter((f) => ["목표", "보관", "삭제"].includes(String(f["따행상태"] || "").trim())).slice().sort((a, b) => ymOfFollow(b["대상년월"]).localeCompare(ymOfFollow(a["대상년월"]))) : follow.rows.filter((f) => ymOfFollow(f["대상년월"]) === ymOfFollow(ym));
    fRows.forEach((f) => {
      const id = f["고객ID"], e = ownOf(id), sk = foState ? id + "|" + ymOfFollow(f["대상년월"]) : id;
      if (!e || seen.has(sk)) return;
      seen.add(sk);
      const r = base(e, id);
      r.fileId = f._fileId; r.follow = f;
      if (foState) r.grp = ymOfFollow(f["대상년월"]) === ymOfFollow(ymNow()) ? 0 : 1;
      r.del = foState && String(f["따행상태"] || "").trim() === "삭제"; r.canHell = foState; // v416: 따행상태 '삭제' = 지옥
      r.archDay = String(f["따행보관"] || "").trim(); // 보관(그 달 대상 단위): 따행 시트 '따행보관'
      r.dda = { done: f["방문완료일"] || "", /* v384: 방문완료는 방문완료일만 봄(접촉현황은 보지 않음) */ send: f["설문발송일"], resp: f["설문응답일"], coupon: f["쿠폰발송일"], assign: "", apply: "", due: f["방문약속일"], gift: f["선물"], deliver: f["전달여부"], cls: f["따행분류"], active: f["캠페인대상여부"], contact: f["접촉현황"], target: ddaTarget(f) };
      out.push(r);
    });
    // A: 따행 시트가 없는 파일 — 고객 시트의 따행 필드 (따행대상='예' 전체, 월 무관)
    data.lists.forEach((e) => {
      const id = idOf(e);
      if (!id || e.sheet !== CUST_SHEET || seen.has(id)) return;
      if (follow.files.has(e.src && e.src.fileId)) return; // B 파일은 위에서 처리
      const mine = data.custById.get(id);
      const own = mine.find((x) => x.fc === primaryFc(mine)) || e; // 담당 FC 행
      if (own !== e) return;
      if (fieldOf(e, "따행대상") !== "예") return;
      seen.add(id);
      const r = base(e, id);
      r.archDay = fieldOf(e, "따행보관"); // 따행 시트가 없는 파일: 고객 시트 '따행보관'
      r.dda = { send: fieldOf(e, "설문발송일"), resp: fieldOf(e, "설문응답일"), coupon: fieldOf(e, "쿠폰발송일"), assign: fieldOf(e, "따행배정일"), apply: fieldOf(e, "따행고객신청"), due: fieldOf(e, "따행방문예정일"), done: fieldOf(e, "따행방문완료일"), gift: fieldOf(e, "따행선물"), deliver: fieldOf(e, "따행전달"), cls: "", active: "", contact: "" };
      out.push(r);
    });
    out.sort((a, b) => ((a.grp || 0) - (b.grp || 0)) || a.holder.localeCompare(b.holder, "ko") || dateKey(a.bday).localeCompare(dateKey(b.bday)) || dateKey(a.dda.assign).localeCompare(dateKey(b.dda.assign))); // 이름 오름차순(가나다), 같은 이름은 생년월일 → 배정일 빠른 순
    return out;
  }
  // 따행 필터: 활동대상(기본) → 접촉현황(방문완료 칸 값 유무) → 따행분류. totals는 활동대상 기준 전체(SOL 조회와 같은 73·16·57), classes는 접촉현황 필터 후의 분류별 건수
  function ddaFilter(baseRows, bMode, st) {
    const afterAct = baseRows.filter((r) => !bMode || !st.dact || r.arch || r.dda.target !== false); // v415: 보관(따행상태 보관·삭제)은 활동대상이 아니어도 보관 목록에 보임 — 전에는 건수만 나오고 목록이 비었음
    const afterCt = afterAct.filter((r) => !st.dct || (st.dct === "done" ? !!r.dda.done : !r.dda.done));
    const cc = new Map(); afterCt.forEach((r) => { if (r.dda.cls) cc.set(r.dda.cls, (cc.get(r.dda.cls) || 0) + 1); });
    return {
      rows: st.dcls ? afterCt.filter((r) => r.dda.cls === st.dcls) : afterCt,
      classes: [...cc.entries()].sort((x, y) => y[1] - x[1] || x[0].localeCompare(y[0], "ko")),
      totals: { all: afterAct.length, done: afterAct.filter((r) => r.dda.done).length, open: afterAct.filter((r) => !r.dda.done).length }
    };
  }
  // 따행 필터 막대 (대상년월은 위 월 이동): 접촉현황 · 따행분류 · 활동대상 · 분류별 건수
  function ddaBarHtml(st) { // st: { dct, dcls, dact, classes: [[이름, 건수]], totals }
    const seg = [["", "전체"], ["done", "방문완료"], ["open", "미활동"]].map(([v, l]) => `<button type="button" data-dct="${v}" class="${st.dct === v ? "on" : ""}">${l}</button>`).join("");
    const opts = [`<option value="">따행분류: 전체</option>`].concat(st.classes.map(([n, c]) => `<option value="${esc(n)}"${st.dcls === n ? " selected" : ""}>${esc(n)} (${c})</option>`)).join("");
    const chips = st.classes.map(([n, c]) => `<button type="button" class="cls-chip${st.dcls === n ? " on" : ""}" data-dcls="${esc(n)}">${esc(n)} ${c}</button>`).join("");
    return `<div class="dbar"><div class="dseg">${seg}</div><div class="drow"><select id="dclsSel" aria-label="따행분류">${opts}</select><button type="button" id="dactBtn" class="act-pill${st.dact ? " on" : ""}" aria-pressed="${st.dact ? "true" : "false"}">활동대상</button></div></div>` +
      `<div class="dsum"><b>${st.totals.all}명</b> · 방문완료 <b>${st.totals.done}</b> · 미활동 <b>${st.totals.open}</b></div>` + (chips ? `<div class="dchips">${chips}</div>` : "");
  }
  // ---- 화면 ----
  const salesBtn = (r, kind, mark, ctx) => {
    const lab = kind === "call" ? "전화" : "문자", has = !!r.phone, ts = mark && mark[kind];
    const href = has ? (kind === "call" ? `tel:${r.phone}` : smsHref(r.phone, fillSms((ctx && ctx.tpl) || "", r, ctx && ctx.fcName))) : "";
    return `<a class="sbtn ${kind}${ts ? " done" : ""}${has ? "" : " off"}" ${has ? `href="${esc(href)}" data-contact="${kind}" data-key="${esc(r.key)}"` : ""}>${lab}</a>`;
  };
  // 한 건 = 3줄: ① 계약자 · 피보 · 상품명 · 보험료 ② 계약일 · 생일 · 연락처 [전화][문자] · 직업 ③ 주소 · 모집FC (줄 안에서 항목이 나란히, 길면 줄바꿈)
  // 따행 칩 4개 (필드 값이 있으면 ON): 방문완료=따행방문완료일 · 설문발송=설문발송일 · 설문응답=설문응답일 · 쿠폰발송=쿠폰발송일
  const DCHIPS = [["done", "방문완료", "done", "따행방문완료일"], ["send", "설문발송", "send", "설문발송일"], ["resp", "설문응답", "resp", "설문응답일"], ["coupon", "쿠폰발송", "cpn", "쿠폰발송일"]];
  // v93d 배치: ① 순번·이름(→피보) [이관 50% / 따행 30%] + 오른쪽 끝(이관 상태 칩 / 따행 칩 4개) ② 상품·보험료(아래 줄, 계약마다) ③ 계약일·생일·연락처[전화][문자]·직업(+따행은 모집FC) ④ 주소(+이관은 모집FC)
  function salesRowsHtml(rows, mark, ctx, openSet) {
    const it = (cls, lb, v) => (v === "" || v == null ? "" : `<span class="si ${cls || ""}">${lb ? `<i class="lb">${lb}</i>` : ""}${v}</span>`);
    const grpOn = rows.some((r) => r.grp === 1); let pg;
    return `<div class="slist">` + rows.map((r, idx) => {
      const gh = grpOn && r.grp !== pg ? `<div class="sgrp">${r.grp === 1 ? "지난 목표" : "이번 달"}</div>` : ""; pg = r.grp;
      const mk = mark.get(r.key) || {};
      const isD = !!r.dda, lock = isD && isRO(r.fc);
      const st = r.status && r.status !== "유지" ? `<span class="st${r.status === "완납" ? " st-ok" : r.status === "실효" ? " st-bad" : ""}">${esc(r.status === "납입면제" ? "납면" : r.status)}</span>` : ""; // 완납=초록, 실효=빨강
      const times = mk.call ? `전화 시도 ✓ ${tsShow(mk.call)}` : ""; // v491: 문자 시도 줄은 표시 안 함([문자] 버튼 색·필터는 그대로)
      const stageChips = !isD && r.stage !== undefined ? STAGE_CHIPS.map(([lab, cls]) => `<button type="button" class="dc ${cls} ${r.stage === lab ? "on" : "off"}" data-stg="${lab}" data-key="${esc(r.key)}" aria-pressed="${r.stage === lab ? "true" : "false"}">${lab}</button>`).join("") : "";
      // 따행: 방문완료 > 설문발송 > 설문응답 > 쿠폰발송 중 가장 마지막에 달성한 칩만 색(값이 있는 앞 칩은 'past': 색 없이, 누르면 그대로 켬/끔)
      const lastD = isD ? DCHIPS.reduce((acc, [f], i) => (r.dda[f] ? i : acc), -1) : -1;
      const chips = stageChips ? stageChips : isD ? DCHIPS.map(([f, lab, cls], i) => `<button type="button" class="dc ${cls} ${i === lastD ? "on" : r.dda[f] ? "off past" : "off"}${lock ? " lock" : ""}" data-chip="${f}" data-key="${esc(r.key)}" aria-pressed="${r.dda[f] ? "true" : "false"}">${lab}</button>`).join("") : "";
      // [이동]/[원복]: 담당이 다른 FC 계약(참고)·계약ID 없는 계약·보기 전용은 없음. 계약 고객이 된 DB 고객은 '계약'(자동 보관)
      const arDay = r.arch && r.archDay ? `<span class="r-day">${r.db ? (r.del ? "지옥" : "보관") : r.del ? "지옥" : "완료"} ${esc(dshow(r.archDay).slice(5, 10))}</span>` : "";
      const arBtn = r.db && r.del ? (r.ref || !r.key || isRO(r.fc) ? "" : `<i class="ar-sep" aria-hidden="true"></i><button type="button" class="ar-btn back" data-ardelnew="${esc(r.key)}" aria-label="${esc(r.holder)} 신규로 이동">신규로</button><button type="button" class="ar-btn" data-ardelx="${esc(r.key)}" aria-label="${esc(r.holder)} 보관으로 이동">보관으로</button>`) : r.db && r.arch && r.dbAuto ? `<i class="ar-sep" aria-hidden="true"></i><span class="ar-auto">계약</span>` : (r.ref || !r.key || isRO(r.fc) || r.arch === undefined) ? "" : `<i class="ar-sep" aria-hidden="true"></i>${r.db ? `<button type="button" class="ar-btn${r.arch ? " back" : ""}" data-arx="${esc(r.key)}" aria-label="${esc(r.holder)} ${r.arch ? "신규로 이동" : "보관으로 이동"}">${r.arch ? "신규로" : "보관으로"}</button>${r.arch ? `<button type="button" class="ar-btn del" data-ardel="${esc(r.key)}" aria-label="${esc(r.holder)} 지옥으로 이동">지옥으로</button>` : ""}`
        : !r.arch ? `<button type="button" class="ar-btn" data-arx="${esc(r.key)}" data-to="done" aria-label="${esc(r.holder)} 완료로 이동">완료</button>` // v416: 이관·따행 = 신규 · 완료 · 지옥
        : `<button type="button" class="ar-btn back" data-arx="${esc(r.key)}" data-to="new" aria-label="${esc(r.holder)} 신규로 이동">신규로</button>${r.del ? `<button type="button" class="ar-btn" data-arx="${esc(r.key)}" data-to="done" aria-label="${esc(r.holder)} 완료로 이동">완료로</button>` : r.canHell ? `<button type="button" class="ar-btn del" data-arx="${esc(r.key)}" data-to="hell" aria-label="${esc(r.holder)} 지옥으로 이동">지옥으로</button>` : ""}`}`;
      const arHtml = arDay ? `<span class="ar-act">${arDay}</span>` : "";
      // 왼쪽 굵은 선 = 지금 단계 색 (이관·DB = 영업단계 칩, 따행 = 마지막 달성 칩). 단계가 없으면 선 없음
      const lv = isD ? (lastD >= 0 ? DCHIPS[lastD][2] : "") : ((STAGE_CHIPS.find(([lab]) => lab === r.stage) || [])[1] || "");
      const prodLines = isD
        ? (r.prods && r.prods.length ? r.prods : [{ name: "계약 정보 없음(조회 제한)", fee: "" }])
        : [{ name: r.product, fee: r.fee }];
      const detail = isD ? `<div class="sline sub detail">방문완료 ${esc(r.dda.done || "-")} · 설문발송일 ${esc(r.dda.send || "-")} · 설문응답일 ${esc(r.dda.resp || "-")} · 쿠폰발송일 ${esc(r.dda.coupon || "-")}</div>` : "";
      return gh + `<div class="srow${lv ? " lv lv-" + lv : ""}${r.ref ? " ref" : ""}${isD ? (r.dda.done ? " d-done" : " d-open") + " fold" + (openSet && openSet.has(r.key) ? " open" : "") : ""}" data-fold="${esc(r.key)}">
        <div class="sline l1"><span class="no">${idx + 1}</span><span class="nmw ${isD ? "n30" : "n50"} ${gradeClass(gradeOf({ entries: (SALES.data && SALES.data.custById && SALES.data.custById.get(r.custId)) || [] })) || "grade-good"}">${r.custId ? `<a href="#" data-cust="${esc(r.custId)}">${esc(r.holder)}</a>` : esc(r.holder)}${r.ins ? ` <small class="ins">→ ${esc(r.ins)}</small>` : ""}${r.db && r.sexAge ? ` <small class="dbsa">${esc(r.sexAge)}</small>` : ""}${r.db && r.pgFc ? ` <small class="dbpg">진행: ${esc(r.pgFc)}</small>` : ""}</span><span class="rt">${chips}${arBtn}</span></div>
        ${(r.db ? [] : prodLines).map((p) => `<div class="sline prod"><span class="si pd">${esc(p.name)}${isD ? "" : st}</span><span class="si fe">${esc(p.fee ? wonText(p.fee) : "")}</span></div>`).join("")}
        ${r.db ? `<div class="sline l2 dbl"><span class="si ph">${esc(fmtPhoneText(r.phoneText) || "연락처 없음")}<span class="sbtns">${salesBtn(r, "call", mk, ctx)}${salesBtn(r, "sms", mk, ctx)}</span></span></div>
        <div class="sline l3 dbl">${r.comp ? `<b class="dbco">${esc(r.comp)}</b>` : ""}${it("", "유입", esc(String(r.inAt || "").slice(0, 10) || "-"))}${dbAreaOf(r) ? `<span class="dbar">${esc(dbAreaOf(r))}</span>` : ""}${arHtml}</div>
        ${r.dbMemo ? `<div class="sline dbm"><small>유입메모</small>${esc(r.dbMemo)}</div>` : ""}` : `<div class="sline l2">${it("", r.cdLabel || "계약일", esc(r.cdate))}${it("", "생일", esc(r.bday))}<span class="si ph">${esc(fmtPhoneText(r.phoneText))} ${r.ref ? `<span class="hint">참고(기록 안 함)</span>` : ""}<span class="sbtns">${salesBtn(r, "call", mk, ctx)}${salesBtn(r, "sms", mk, ctx)}</span></span></div>
        <div class="sline l3">${it("", "직업", esc(r.job))}${it("", "모집", esc(r.recFc))}${arHtml}</div>
        <div class="sline l4">${it("ad", "", esc(r.addr))}${mapBtn(r.addr)}</div>`}
        ${times ? `<div class="sline sub ok">${times}</div>` : ""}${detail}${r.ref || !r.custId ? "" : smemoHtml(r.holder, smemoFc(r.custId, r.fc || r.srcFc), isRO(r.fc || r.srcFc))}</div>`;
    }).join("") + `</div>`;
  }
  // ---- 영업단계 칩(이관 목록): 통화 · AP · PC · 계약. 지금 값인 칩만 색이 들어오고, 누르면 그 고객의 '영업단계' 칸에 칩 글자를 저장 ----
  const STAGE_CHIPS = [["통화", "send"], ["AP", "resp"], ["PC", "cpn"], ["계약", "done"]];
  async function setStage(row, value) {
    if (!row.ent) return;
    if (isRO(row.fc)) return toast("보기 전용 권한이라 수정할 수 없습니다.");
    if (row.stage === value) { // 이미 켜진 칩을 다시 누르면: 확인 후 영업단계를 비움
      if (!confirm("고객의 영업상태를 제거하시겠습니까?")) return;
      value = "";
    }
    const f0 = row.ent.fields.find(([k]) => k === "영업단계"), prev = f0 ? f0[1] : "";
    const put = (v) => { const f = row.ent.fields.find(([k]) => k === "영업단계"); if (f) f[1] = v; else row.ent.fields.push(["영업단계", v]); };
    put(value); renderSales(true); // 화면 먼저
    try { await writeEntryCells([row.ent], { 영업단계: value }); }
    catch (err) { put(prev); renderSales(true); handleError(err); }
  }
  // ---- 진행 고객에서 빼기 / 되돌리기: 고객 기록은 그대로 두고 고객 시트의 '진행제외' 칸에 날짜를 적어 진행 목록에서만 뺌. 이력에 한 줄 ----
  function progEnts(cust) {
    const ents = cust.entries.filter((e) => e.sheet === CUST_SHEET && e.src && (e.type === "진행" || e.type === "진행제외"));
    if (!ents.length) { toast("진행 고객 정보를 찾지 못했습니다."); return null; }
    if (ents.some((e) => isRO(e.fc))) { toast("보기 전용 권한이라 수정할 수 없습니다."); return null; }
    return ents;
  }
  const ARCH_ASK = (name, what) => `${name} 고객${what ? "의 " + what : ""}을 완료로 옮길까요?\n고객 정보는 그대로예요.`;
  const stLabel = (v) => (hellV(v) ? "지옥" : v ? "완료" : "신규"); // v416: 이관·따행 보기 이름(저장 값은 삭제·보관·목표)
  // 이관·따행 보관: 한 시트의 한 줄(이관 = 계약 시트의 계약ID 줄 '이관보관', 따행 = 따행 시트의 고객ID+대상년월 줄 또는 고객 시트의 고객ID 줄 '따행보관')에 날짜, 이력 한 줄 → 한 번 받고 한 번 저장
  // 고객의 이달 이관 계약(계약자 = 이 고객, 담당 FC 계약만, 계약ID 있는 것). 상태: 없음 null / 모두 보관 "off" / 그 밖 "on"
  // 이달 이관만(이관일자가 이번 달인 계약 = 세일즈 > 이달목표 > 이관에 나오는 계약). 지난 달 이관 계약은 보관하지 않음
  const trCtsOf = (ids, cts) => cts.filter((c) => c.type === "이관" && monthOf(c.row["이관일자"] || c.row["이관일"] || "") === ymNow() && ids.has(String(c.row["계약자ID"] || "").trim()) && !(c.fc && c.srcFc && c.fc !== c.srcFc) && String(c.row["계약ID"] || "").trim() && c.src && c.src.fileId);
  const trStateOf = (list) => (!list.length ? null : list.every((c) => String(c.row["이관보관"] || "").trim()) ? "off" : "on");
  // 이관 계약 여러 건 보관/원복: 파일마다 한 번 받고 한 번 저장(계약 시트 '이관보관', 이력은 계약마다 한 줄)
  async function saveTrCts(list, value) {
    const byFile = new Map();
    list.forEach((c) => { if (!byFile.has(c.src.fileId)) byFile.set(c.src.fileId, []); byFile.get(c.src.fileId).push(c); });
    for (const [fileId, cs] of byFile) {
      const wb = await stGetBook(fileId);
      const info = sheetInfo(wb, "계약", "계약ID");
      if (!info) throw new Error("계약 시트를 찾지 못했습니다.");
      const tgs = sttCells(info, wb, "이관보관", value).map((t) => ({ c: ensureColumn(info, t.col), val: t.val })), kc = info.colOf("계약ID"), hi = sheetInfo(wb, "이력", "고객ID"), stamp = tsDigits(nowStamp());
      cs.forEach((c) => {
        const cid = String(c.row["계약ID"]).trim();
        for (let r = info.hr + 1; r <= info.rg.e.r; r++) if (info.txt(r, kc) === cid) { tgs.forEach((t) => { if (t.val === "") delete info.ws[info.enc(r, t.c)]; else info.ws[info.enc(r, t.c)] = { t: "s", v: t.val }; }); break; }
        const text = (value ? "이관 → 보관" : "보관 → 이관") + ` (계약 ${productName(c.row["상품명"]) || cid}, 고객 화면에서 ${value ? "이동" : "원복"})`;
        if (hi) appendRow(hi, { 일시: { stamp }, 고객ID: String(c.row["계약자ID"] || "").trim(), 계약ID: cid, 담당FC: c.srcFc || c.fc || "", 유형: value ? "이동(보관)" : "원복", 결과: "입력", 내용: text, 입력일시: { stamp } });
      });
      await stPutBook(fileId, wb);
      cacheBook(fileId, wb);
    }
  }
  // 고객 단위 이관 이동/원복: 화면 먼저, 저장은 뒤에서 차례로, 실패하면 되돌림 + 6초 되돌리기
  function trArchive(name, list, on, after) {
    if (!list.length) return toast("보관할 수 있는 이관 계약이 없습니다.");
    if (list.some((c) => isRO(c.srcFc || c.fc))) return toast("보기 전용 권한이라 수정할 수 없습니다.");
    const prev = list.map((c) => String(c.row["이관보관"] || "").trim());
    const setAll = (vals) => list.forEach((c, i) => { c.row["이관보관"] = vals[i]; });
    const run = (vals) => {
      const before = list.map((c) => String(c.row["이관보관"] || "").trim());
      setAll(vals); if (after) after();
      progBusy(1);
      const targets = list.filter((c, i) => before[i] !== vals[i]);
      progQ = progQ.then(async () => { for (const v of [...new Set(targets.map((c) => vals[list.indexOf(c)]))]) await saveTrCts(targets.filter((c) => vals[list.indexOf(c)] === v), v); })
        .then(() => true, (err) => { setAll(before); if (after) after(); handleError(err); return false; }).finally(() => progBusy(-1));
    };
    const today = todayStr(), onVals = list.map(() => today), offVals = list.map(() => "");
    if (on) {
      if (!confirm(ARCH_ASK(name, list.length > 1 ? `이관 계약 ${list.length}건` : "이관 계약"))) return;
      run(onVals);
      undoToast(`${name} 고객을 보관으로 이동했습니다`, () => { run(prev); toast("원복했습니다."); });
    } else {
      run(offVals);
      undoToast(`${name} 고객을 원복했습니다`, () => { run(prev); toast("다시 보관했습니다."); });
    }
  }
  async function saveArchCell(fileId, sheet, keyCol, match, col, value, hist) {
    const wb = await stGetBook(fileId);
    const info = sheetInfo(wb, sheet, keyCol);
    if (!info) throw new Error(sheet + " 시트를 찾지 못했습니다.");
    let r0 = -1;
    for (let r = info.hr + 1; r <= info.rg.e.r; r++) if (match(info, r)) { r0 = r; break; }
    if (r0 < 0) throw new Error(sheet + " 시트에서 이 줄을 찾지 못했습니다.");
    sttCells(info, wb, col, value).map((t) => ({ c: ensureColumn(info, t.col), val: t.val })).forEach((t) => { if (t.val === "") delete info.ws[info.enc(r0, t.c)]; else info.ws[info.enc(r0, t.c)] = { t: "s", v: t.val }; });
    const hi = sheetInfo(wb, "이력", "고객ID"), stamp = tsDigits(nowStamp());
    if (hi) appendRow(hi, { 일시: { stamp }, 고객ID: hist.custId, 계약ID: hist.cid || "", 담당FC: hist.fc || "", 유형: hellV(value) ? "이동(삭제)" : value ? "이동(보관)" : "원복", 결과: "입력", 내용: hist.text, 입력일시: { stamp } });
    await stPutBook(fileId, wb);
    cacheBook(fileId, wb); // 방금 저장한 내용이 최신
    if (histCache) histCache.then((a) => a.push({ fc: hist.fc, fileId, 일시: nowStamp(), 고객ID: hist.custId, 계약ID: hist.cid || "", 담당FC: hist.fc || "", 유형: hellV(value) ? "이동(삭제)" : value ? "이동(보관)" : "원복", 결과: "입력", 내용: hist.text })).catch(() => {});
  }
  // 영업 이관·따행 줄의 이동/원복: 화면 먼저, 저장은 뒤에서 차례로, 실패하면 화면 되돌림
  // v433(2026-11-18 사용자 지시): 지옥으로 보낸 계약·계약자는 성향을 Shit으로 바꿈(되돌리기 하면 예전 성향으로)
  function hellShit(ents) {
    const l = (ents || []).filter((e) => e && e.src && !isRO(e.fc));
    if (!l.length || l.every((e) => String(e.grade || "").toLowerCase() === "shit")) return null;
    const prev = l.map((e) => e.grade || "");
    const put = (g) => { l.forEach((e) => (e.grade = g)); progBusy(1); progQ = progQ.then(() => writeEntryCells(l, { [GRADE_COL]: g })).then(() => true, (err) => { l.forEach((e, i) => (e.grade = prev[i])); handleError(err); return false; }).finally(() => progBusy(-1)); };
    put("Shit");
    return () => put(prev[0] || "Good");
  }
  function archSalesRow(row, on, after) {
    const isT = SALES.tab === "transfer", today = todayStr();
    let setUi, save;
    if (isT) {
      if (!row.ct || !row.cid) return toast("계약ID가 없는 계약이라 보관할 수 없습니다.");
      setUi = (v) => { row.ct.row["이관보관"] = hellDay(v); if ("이관상태" in row.ct.row) row.ct.row["이관상태"] = hellV(v) ? "삭제" : v ? "보관" : "목표"; };
      save = (v) => saveArchCell(row.fileId, "계약", "계약ID", (info, r) => info.txt(r, info.colOf("계약ID")) === row.cid, "이관보관", v,
        { custId: row.custId, cid: row.cid, fc: row.fc, text: `이관 → ${stLabel(v)} (계약 ${row.product || row.cid}, 앱에서 이동)` });
    } else if (row.follow) {
      const fo = row.follow;
      setUi = (v) => { fo["따행보관"] = hellDay(v); if ("따행상태" in fo) fo["따행상태"] = hellV(v) ? "삭제" : v ? "보관" : "목표"; };
      save = (v) => saveArchCell(row.fileId, "따행", "고객ID", (info, r) => info.txt(r, info.colOf("고객ID")) === row.custId && ymOfFollow(info.txt(r, info.colOf("대상년월"))) === ymOfFollow(fo["대상년월"]), "따행보관", v,
        { custId: row.custId, fc: row.fc, text: `따행 → ${stLabel(v)} (대상 ${fo["대상년월"]}, 앱에서 이동)` });
    } else {
      if (!row.ent) return toast("고객 정보를 찾지 못했습니다.");
      setUi = (v) => { setEntField(row.ent, "따행보관", v); };
      save = (v) => saveArchCell(row.fileId, CUST_SHEET, ID_COL, (info, r) => info.txt(r, info.colOf(ID_COL)) === row.custId, "따행보관", v,
        { custId: row.custId, fc: row.fc, text: (v ? "따행 → 보관" : "보관 → 따행") + ` (앱에서 ${v ? "이동" : "원복"})` });
    }
    if (isRO(row.fc)) return toast("보기 전용 권한이라 수정할 수 없습니다.");
    const run = (v, prev) => {
      setUi(v); if (after) after();
      progBusy(1);
      progQ = progQ.then(() => save(v)).then(() => true, (err) => { setUi(prev); if (after) after(); handleError(err); return false; }).finally(() => progBusy(-1));
    };
    const to = on === true ? "done" : on === false ? "new" : on; // v416: new(신규) / done(완료) / hell(지옥)
    const prev = row.del ? "D:" + (row.archDay || today) : row.archDay || "";
    if (to === "hell" && !row.canHell) return toast("이 파일은 지옥(삭제) 상태를 저장할 수 없습니다.");
    const v = to === "new" ? "" : to === "hell" ? "D:" + today : (row.archDay && !row.del ? row.archDay : today);
    if (v === prev) return;
    if (to === "done" && !row.arch && !confirm(ARCH_ASK(row.holder, isT ? "이 계약(" + (row.product || "") + ")" : ""))) return;
    if (to === "hell" && !confirm(`${row.holder} 고객${isT ? "의 이 계약(" + (row.product || "") + ")" : ""}을 지옥으로 보낼까요?\n목록에서 옮겨지고 고객 성향이 Shit으로 바뀌어요. 지옥 칩에서 다시 볼 수 있어요.`)) return;
    run(v, prev); if (to === "hell") hellFx();
    const unShit = to === "hell" ? hellShit((SALES.data && SALES.data.custById.get(row.custId)) || (row.ent ? [row.ent] : [])) : null;
    undoToast(`${row.holder} 고객을 ${stLabel(v)}(으)로 옮겼습니다`, () => { run(prev, v); if (unShit) unShit(); toast("되돌렸습니다."); });
  }
  // 목록에서 빼기/되돌리기 공통(진행 = '진행제외' 칸, DB = 'DB제외' 칸): 칸에 날짜를 적고 이력 한 줄 → 파일마다 한 번 받고 한 번 저장
  const EXCL = {
    progress: { field: "진행제외", on: "이동(보관)", off: "원복", onText: "진행 → 보관 (앱에서 이동, 고객 정보는 그대로)", offText: "보관 → 진행 (앱에서 원복)", type: true },
    db: { field: "DB상태", val: "보관", offVal: "신규", on: "이동(보관)", off: "원복", onText: "DB → 보관 (앱에서 이동, 고객 정보는 그대로)", offText: "보관 → DB (앱에서 원복)", type: false },
    dbdel: { field: "DB상태", val: "삭제", offVal: "보관", on: "삭제", off: "삭제취소", onText: "DB 보관 → 삭제 표시 (앱에서 삭제, 자료는 그대로 두고 목록에서만 숨김)", offText: "DB 삭제 표시 취소 (앱에서 되돌림, 보관 목록으로)", type: false } // v364: 'DB상태' 칸에 '삭제' (v363의 DB삭제 칸은 더 쓰지 않음)
  };
  function dbSheetTarget(ents) { // 이 고객의 DB 시트 줄(여럿이면 계약 아닌 것 중 배분일 최신)
    const id = idOf(ents[0]), rs = (DBS_BY_ID.get(id) || []).filter((r) => r["DB상태"] !== "계약");
    return rs.sort((a, b) => String(b["FC 배분일"] || "").localeCompare(String(a["FC 배분일"] || "")))[0] || (DBS_BY_ID.get(id) || [])[0] || null;
  }
  async function saveDbSheet(mode, ents, exclude, d) {
    const M = EXCL[mode], e0 = ents[0], stamp = tsDigits(nowStamp()), val = exclude ? M.val : M.offVal, done = val === "보관" || val === "삭제";
    const wb = await stGetBook(d._fileId), info = sheetInfo(wb, "DB", "DB ID");
    if (!info) throw new Error("DB 시트를 찾지 못했습니다.");
    const kc = info.colOf("DB ID"); let r0 = -1;
    for (let r = info.hr + 1; r <= info.rg.e.r; r++) if (info.txt(r, kc) === d["DB ID"]) { r0 = r; break; }
    if (r0 < 0) throw new Error("DB 시트에서 이 줄을 찾지 못했습니다.");
    const setc = (col, v) => { const c = ensureColumn(info, col); if (v === "") delete info.ws[info.enc(r0, c)]; else info.ws[info.enc(r0, c)] = { t: "s", v }; };
    setc("DB상태", val === "신규" ? "" : val); setc("완료일", done ? todayStr() : ""); setc("완료사유", done ? "사용자" : "");
    const kind = exclude ? M.on : M.off, text = (exclude ? M.onText : M.offText) + " DB:" + d["DB ID"];
    const hist = sheetInfo(wb, "이력", "고객ID");
    if (hist) appendRow(hist, { 일시: { stamp }, 고객ID: idOf(e0), 계약ID: "", 담당FC: d._fc || "", 유형: kind, 결과: "입력", 내용: text, 입력일시: { stamp } });
    await stPutBook(d._fileId, wb); cacheBook(d._fileId, wb);
    if (histCache) histCache.then((a) => a.push({ fc: d._fc, fileId: d._fileId, 일시: nowStamp(), 고객ID: idOf(e0), 계약ID: "", 담당FC: d._fc || "", 유형: kind, 결과: "입력", 내용: text })).catch(() => {});
  }
  async function saveExclude(mode, ents, exclude) {
    if (mode === "db" || mode === "dbdel") { const d = dbSheetTarget(ents); if (d) return saveDbSheet(mode, ents, exclude, d); } // v571: DB 시트 줄이 있으면 거기에만 씀(옛 칸에는 안 씀)
    const M = EXCL[mode], e0 = ents[0], stamp = tsDigits(nowStamp());
    const kind = exclude ? M.on : M.off, text = exclude ? M.onText : M.offText;
    await writeEntryCells(ents, { [M.field]: exclude ? (M.val || todayStr()) : (M.offVal || "") }, (wb, fileId) => {
      if (fileId !== e0.src.fileId) return;
      try {
        const hist = sheetInfo(wb, "이력", "고객ID");
        if (hist) appendRow(hist, { 일시: { stamp }, 고객ID: idOf(e0), 계약ID: "", 담당FC: e0.fc || "", 유형: kind, 결과: "입력", 내용: text, 입력일시: { stamp } });
      } catch (err) { console.warn("이력 기록 실패", err); }
    });
    // 이력 캐시는 버리지 않고 방금 넣은 한 줄만 더함(버리면 목록을 다시 그릴 때 파일을 또 받아 1초쯤 늦어짐)
    if (histCache) histCache.then((a) => a.push({ fc: e0.fc, fileId: e0.src.fileId, 일시: nowStamp(), 고객ID: idOf(e0), 계약ID: "", 담당FC: e0.fc || "", 유형: kind, 결과: "입력", 내용: text })).catch(() => {});
  }
  // 화면 먼저 바꾸고 저장은 뒤에서. 저장은 하나씩 차례로(되돌리기를 빨리 눌러도 순서대로 저장)
  let progQ = Promise.resolve(), progBusyN = 0;
  function progBusy(d) {
    progBusyN = Math.max(0, progBusyN + d);
    let el = $("saveInd");
    if (!el) { el = document.createElement("div"); el.id = "saveInd"; el.className = "save-ind"; el.setAttribute("role", "status"); el.textContent = "저장 중…"; document.body.appendChild(el); }
    el.hidden = !progBusyN;
  }
  function exclApply(mode, ents, exclude, after) {
    const M = EXCL[mode];
    const prev = ents.map((e) => [e.type, fieldOf(e, M.field)]);
    const dRow = (mode === "db" || mode === "dbdel") ? dbSheetTarget(ents) : null, dPrev = dRow ? [dRow["DB상태"], dRow["완료일"]] : null;
    const setUi = (i, type, day) => { const e = ents[i]; if (M.type) e.type = type; if (dRow) { if (i === 0) { dRow["DB상태"] = day === "신규" ? "" : day; dRow["완료일"] = day === "보관" || day === "삭제" ? todayStr() : ""; } return; } setEntField(e, M.field, day); };
    ents.forEach((e, i) => setUi(i, exclude ? "진행제외" : "진행", exclude ? (M.val || todayStr()) : (M.offVal || "")));
    if (after) after();
    progBusy(1);
    progQ = progQ.then(() => saveExclude(mode, ents, exclude)).then(() => true, (err) => {
      if (dRow) { dRow["DB상태"] = dPrev[0]; dRow["완료일"] = dPrev[1]; } else ents.forEach((e, i) => setUi(i, prev[i][0], prev[i][1])); // 저장 실패 → 화면 되돌림
      if (after) after();
      handleError(err);
      return false;
    }).finally(() => progBusy(-1));
    return progQ;
  }
  const progApply = (cust, ents, exclude, after) => exclApply("progress", ents, exclude, after);
  // DB 고객: 고객 시트의 최초유입경로가 'DB'(또는 현재단계 DB). 'DB제외' 날짜가 있거나 계약 고객이 되면 DB 보관 목록(계약은 자동, 화면 계산)
  const isDbRec = (r) => (r.type === "DB" || fieldOf(r, "최초유입경로") === "DB") && (!fieldOf(r, "유입FC") || fieldOf(r, "유입FC") === r.fc); // v514: 유입FC가 다른 FC면 그 줄은 DB로 세지 않음(유입FC 파일 줄이 DB 줄)
  // v514: 이 사람의 DB 줄 = 유입FC 파일의 고객 시트 줄(유입FC가 비었거나 그 파일에 줄이 없으면 null → 예전 방식)
  function dbLineOf(entries) {
    const inflow = (entries || []).map((x) => fieldOf(x, "유입FC")).find(Boolean);
    return inflow ? (entries.find((x) => x.fc === inflow && x.sheet === CUST_SHEET) || null) : null;
  }
  const dbAutoRec = (r) => r.type === "계약" || fieldOf(r, "계약현황") === "유효" || fieldOf(r, "DB상태") === "계약";
  // v364: DB 상태 한 칸 'DB상태'(신규/보관/삭제, 빈칸=신규). 아직 옮기지 않은 고객은 옛 칸으로 판단(DB삭제 '삭제' = 삭제, DB제외 날짜 = 보관)
  const dbState = (r) => { const v = String(fieldOf(r, "DB상태") || "").trim(); if (v === "신규" || v === "보관" || v === "삭제" || v === "계약") return v; if (fieldOf(r, "DB삭제") === "삭제") return "삭제"; return fieldOf(r, "DB제외") ? "보관" : "신규"; };
  const dbOffRec = (r) => dbState(r) !== "신규" || dbAutoRec(r);
  function dbEnts(entries) {
    const all = entries.filter((e) => e.sheet === CUST_SHEET && e.src), dbE = dbLineOf(all);
    const ents = dbE ? [dbE] : all; // v514: 유입FC가 있으면 DB상태는 그 FC 파일 줄에만 씀
    if (!ents.length) { toast("DB 고객 정보를 찾지 못했습니다."); return null; }
    if (ents.some((e) => isRO(e.fc))) { toast("보기 전용 권한이라 수정할 수 없습니다."); return null; }
    return ents;
  }
  function dbRemove(name, entries, after) {
    if (!confirm(ARCH_ASK(name))) return;
    const ents = dbEnts(entries); if (!ents) return;
    exclApply("db", ents, true, after);
    undoToast(`${name} 고객을 보관으로 이동했습니다`, () => { exclApply("db", ents, false, after); toast("원복했습니다."); });
  }
  function dbRestore(name, entries, after) {
    const ents = dbEnts(entries); if (!ents) return;
    exclApply("db", ents, false, after);
    undoToast(`${name} 고객을 원복했습니다`, () => { exclApply("db", ents, true, after); toast("다시 보관했습니다."); });
  }
  // v491(2026-11-19 사용자 지시): [지옥으로]를 누르면 옮긴 뒤 앱 폴더(웹)의 hellgate.gif를 1초 전체 화면으로. 처음 한 번만 받아 둠
  let hellGifP = null;
  function hellGifUrl() { // 웹(앱과 같은 주소)의 hellgate.gif를 미리 받아 둠
    if (!hellGifP) hellGifP = new Promise((ok) => { const im = new Image(); im.onload = () => ok("hellgate.gif"); im.onerror = () => { hellGifP = null; ok(null); }; im.src = "hellgate.gif"; });
    return hellGifP;
  }
  async function hellFx() {
    const url = await hellGifUrl(); if (!url) return;
    const d = document.createElement("div"); d.className = "hell-fx"; d.setAttribute("aria-hidden", "true");
    const im = new Image(); im.alt = ""; im.src = url + "#" + Date.now(); // 매번 처음부터 재생
    d.appendChild(im); document.body.appendChild(d);
    setTimeout(() => d.remove(), 1000);
  }
  function dbDelete(name, entries, after) { // v363: 보관 목록에서도 숨김(자료는 지우지 않고 'DB삭제' 칸에 '삭제'만 표시)
    if (!confirm(`${name} 고객을 지옥으로 보낼까요?
보관 목록에서도 안 보이게 됩니다.
(자료는 지워지지 않고 지옥 표시만 남고, 고객 성향이 Shit으로 바뀌어요. 지옥 칩에서 다시 볼 수 있어요.)`)) return;
    const ents = dbEnts(entries); if (!ents) return;
    exclApply("dbdel", ents, true, after); hellFx();
    const unShit = hellShit(ownEntries(entries.filter((e) => e.sheet === CUST_SHEET && e.src))); // v514: 성향은 담당FC 줄에(DB상태만 유입FC 줄)
    undoToast(`${name} 고객을 지옥으로 보냈습니다`, () => { exclApply("dbdel", ents, false, after); if (unShit) unShit(); toast("지옥 이동을 취소했습니다."); });
  }
  const PROG_ASK = (name) => ARCH_ASK(name);
  // 진행 목록·고객 정보에서 쓰는 공통 동작: 확인 → (화면 바로 바뀜, 저장은 뒤에서) → 6.5초 되돌리기
  function progressRemove(cust, after) {
    if (!confirm(PROG_ASK(cust.name))) return;
    const ents = progEnts(cust); if (!ents) return;
    progApply(cust, ents, true, after);
    undoToast(`${cust.name} 고객을 보관으로 이동했습니다`, () => { progApply(cust, ents, false, after); toast("원복했습니다."); });
  }
  function progressRestore(cust, after) { // 확인 없이 바로 진행으로(원복) → 6.5초 되돌리기(다시 보관)
    const ents = progEnts(cust); if (!ents) return;
    progApply(cust, ents, false, after);
    undoToast(`${cust.name} 고객을 원복했습니다`, () => { progApply(cust, ents, true, after); toast("다시 보관했습니다."); });
  }
  async function renderStageAct(cust) { // 고객 정보 머리 카드 아래(탭 줄 오른쪽): 보관할 수 있는 목록(진행·DB·이관)마다 [보관으로 이동] / [원복]
    const box = $("stageAct");
    if (!box) return;
    const req = (renderStageAct.req = (renderStageAct.req || 0) + 1);
    const ro = cust.entries.some((e) => isRO(e.fc));
    const ids = new Set(cust.entries.map(idOf).filter(Boolean));
    const acts = [];
    const cs = cust.entries.filter((e) => e.sheet === CUST_SHEET);
    const re = () => { if (current && current.cust === cust) renderStageAct(cust); };
    if (cs.some((e) => e.type === "진행" || e.type === "진행제외")) {
      const out = cs.some((e) => e.type === "진행제외");
      acts.push({ k: "진행", out, go: () => (out ? progressRestore(cust, re) : progressRemove(cust, re)) });
    }
    let cts = [];
    try { cts = await loadContracts(); } catch (e) { cts = []; }
    if (req !== renderStageAct.req) return;
    const dbE = cs.filter(isDbRec);
    if (dbE.length && !cust.entries.some(dbAutoRec)) {
      const out = dbE.some((e) => dbState(e) === "보관");
      acts.push({ k: "DB", out, go: () => (out ? dbRestore : dbRemove)(cust.name, cust.entries, re) });
    }
    const trList = ids.size ? trCtsOf(ids, cts) : [];
    const trSt = trStateOf(trList);
    if (trSt) acts.push({ k: "이관", out: trSt === "off", go: () => trArchive(cust.name, trList, trSt !== "off", re) });
    box.hidden = !acts.length || ro;
    if (box.hidden) { box.innerHTML = ""; return; }
    const one = acts.length === 1;
    box.innerHTML = acts.map((a, i) => `<button type="button" class="secondary${a.out ? " back" : ""}" data-sa="${i}">${one ? (a.out ? "원복" : "보관으로 이동") : a.k + " " + (a.out ? "원복" : "보관")}</button>`).join("");
    box.querySelectorAll("[data-sa]").forEach((b) => { b.onclick = () => acts[+b.dataset.sa].go(); });
  }

  // ---- 따행 칩 터치: 필드에 오늘 날짜(방문완료는 '방문완료')를 쓰거나 비움, 이력에 '수동 변경' 기록 ----
  const todayStr = () => { const d = new Date(); return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`; };
  function setTextCell(info, keyCol, keyVal, col, value) {
    const kc = info.colOf(keyCol);
    if (kc < 0) return false;
    for (let r = info.hr + 1; r <= info.rg.e.r; r++) {
      if (info.txt(r, kc) !== keyVal) continue;
      const c = ensureColumn(info, col);
      if (value === "") delete info.ws[info.enc(r, c)]; else info.ws[info.enc(r, c)] = { t: "s", v: value };
      return true;
    }
    return false;
  }
  async function writeFollowField(row, field, value, prev) {
    const wb = await stGetBook(row.fileId); // 저장 직전 최신 파일
    const cu = sheetInfo(wb, CUST_SHEET, ID_COL);
    if (!cu || !setTextCell(cu, ID_COL, row.custId, field, value)) throw new Error("고객 시트에서 이 고객 줄을 찾지 못했습니다.");
    const hist = sheetInfo(wb, "이력", "고객ID");
    if (hist) {
      const stamp = tsDigits(nowStamp());
      appendRow(hist, { 일시: { stamp }, 고객ID: row.custId, 계약ID: "", 담당FC: row.fc || "", 유형: "따행필드", 결과: value ? "입력" : "비움", 내용: `${field} ${value ? "입력 " + value : "비움(이전 " + (prev || "-") + ")"} (세일즈 화면 수동 변경)`, 입력일시: { stamp } });
    }
    await stPutBook(row.fileId, wb);
    delete bookCache[row.fileId];
    histCache = null;
  }
  const setEntField = (e, k, v) => { const f = e.fields.find(([x]) => x.trim() === k); if (v === "") { if (f) e.fields.splice(e.fields.indexOf(f), 1); } else if (f) f[1] = v; else e.fields.push([k, v]); };
  function undoToast(msg, onUndo) {
    const el = document.createElement("div");
    el.className = "toast2";
    el.innerHTML = `<span>${esc(msg)}</span><b role="button" tabindex="0">되돌리기</b>`;
    document.body.appendChild(el);
    const done = () => el.remove();
    el.querySelector("b").addEventListener("click", () => { done(); onUndo(); });
    setTimeout(done, 6500);
  }
  async function applyFollowField(row, field, value, prev) {
    const e = row.ent;
    setEntField(e, field, value); // 화면은 바로 바꾸고(낙관적), 저장이 실패하면 되돌림
    renderSales(true);
    try { await writeFollowField(row, field, value, prev); return true; }
    catch (err) { setEntField(e, field, prev || ""); renderSales(true); handleError(err); return false; }
  }
  // B: 칩 → 따행 시트 열 (방문완료는 접촉현황도 함께, 쿠폰발송은 쿠폰발송 글자도 함께)
  function followCells(f, value) {
    if (f === "done") return { "방문완료일": value, "접촉현황": value ? "방문완료" : "미활동" };
    if (f === "send") return { "설문발송일": value };
    if (f === "resp") return { "설문응답일": value };
    return { "쿠폰발송일": value }; // 쿠폰발송 열은 읽지도 쓰지도 않음 (활동 계획 화면의 버튼 글자)
  }
  // logs: [{ col: 따행 시트 열 이름, value: 새 값, prev: 이전 값 }] — 이력에 남길 변경
  async function writeFollowRow(row, cells, logs) {
    const fo = row.follow, wb = await stGetBook(row._fileId || fo._fileId); // 저장 직전 최신 파일
    const info = sheetInfo(wb, "따행", "고객ID");
    if (!info) throw new Error("따행 시트를 찾지 못했습니다.");
    const kc = info.colOf("고객ID"), mc = info.colOf("대상년월");
    let r0 = -1;
    for (let r = info.hr + 1; r <= info.rg.e.r; r++) if (info.txt(r, kc) === row.custId && ymOfFollow(info.txt(r, mc)) === ymOfFollow(fo["대상년월"])) { r0 = r; break; }
    if (r0 < 0) throw new Error("따행 시트에서 이 고객의 해당 대상년월 줄을 찾지 못했습니다.");
    Object.entries(cells).forEach(([col, v]) => { if (col === "접촉현황" && info.colOf(col) < 0) return; /* 접촉현황 칸이 없어진 파일(새 구조)에는 쓰지 않음 */ const c = ensureColumn(info, col); if (v === "") delete info.ws[info.enc(r0, c)]; else info.ws[info.enc(r0, c)] = { t: "s", v }; });
    const hist = sheetInfo(wb, "이력", "고객ID");
    if (hist) {
      const stamp = tsDigits(nowStamp());
      // 이력 내용은 '<따행 시트 열 이름> 입력|비움 … 대상 YYYY-MM (세일즈 화면 수동 변경)' 형식 (데이터 갱신 스크립트가 필드 이름으로 읽어 보호). 바뀐 열마다 한 줄
      logs.forEach(({ col, value, prev }) => appendRow(hist, { 일시: { stamp }, 고객ID: row.custId, 계약ID: "", 담당FC: row.fc || "", 유형: "따행필드", 결과: value ? "입력" : "비움", 내용: `${col} ${value ? "입력 " + value : "비움(이전 " + (prev || "-") + ")"} 대상 ${fo["대상년월"]} (세일즈 화면 수동 변경)`, 입력일시: { stamp } }));
    }
    await stPutBook(row._fileId || fo._fileId, wb);
    delete bookCache[row._fileId || fo._fileId];
    histCache = null;
  }
  async function applyChip(row, f, value, prev) {
    const field = DCHIPS.find((c) => c[0] === f)[3];
    if (!row.follow) return applyFollowField(row, field, value, prev);
    const fo = row.follow, cells = followCells(f, value), before = {};
    Object.keys(cells).forEach((k) => { before[k] = fo[k] || ""; });
    Object.assign(fo, cells); // 화면 먼저
    renderSales(true);
    try { await writeFollowRow(row, cells, [{ col: { done: "방문완료일", send: "설문발송일", resp: "설문응답일", coupon: "쿠폰발송일" }[f], value, prev }]); return true; }
    catch (err) { Object.assign(fo, before); renderSales(true); handleError(err); return false; }
  }
  async function toggleChip(row, f) {
    const [, lab, , field] = DCHIPS.find((c) => c[0] === f);
    if (isRO(row.fc)) return toast("보기 전용 권한이라 수정할 수 없습니다.");
    const prev = String(row.dda[f] || "");
    if (!prev) { // OFF → ON: 설문·쿠폰은 오늘 날짜, 방문완료는 '방문완료'
      await applyChip(row, f, f === "done" ? "방문완료" : todayStr(), "");
      return;
    }
    // ON → OFF: 지우기 전에 확인 (SOL에서 가져온 값일 수 있음) + 6초 되돌리기
    if (!confirm(`${field} ${prev} 값을 지울까요?\nSOL에서 가져온 값일 수 있습니다. 지우면 칩이 꺼집니다.`)) return;
    if (await applyChip(row, f, "", prev)) undoToast(`${lab} 지웠습니다`, () => applyChip(row, f, prev, ""));
  }
  // ---- FC 선택 · 초기화 ----
  const salesFcKey = () => "shlife.salesFc." + me();
  const salesUserKey = () => "shlife.salesUser." + me();
  async function myFcName() { try { if (!permNamesCache) await readPerms(); return String((permNamesCache || {})[me()] || ""); } catch (e) { return ""; } }
  // 이 달·이 탭·이 FC의 '시도' 표시를 없던 것으로: 시트 최근*일시를 비우고 이력의 해당 행은 지우지 않고 결과를 '취소(초기화)'로 바꿈
  function clearContacts(wbInfo, targets, ym) {
    const { ct, cu, hist } = wbInfo;
    let n = 0;
    const clearCell = (info, keyCol, keyVal) => {
      const kc = info && info.colOf(keyCol);
      if (!info || kc < 0) return;
      for (let r = info.hr + 1; r <= info.rg.e.r; r++) {
        if (info.txt(r, kc) !== keyVal) continue;
        ["최근통화일시", "최근문자일시"].forEach((col) => { const c = info.colOf(col); if (c >= 0 && monthOf(info.txt(r, c)) === ym) { delete info.ws[info.enc(r, c)]; n++; } });
        break;
      }
    };
    targets.forEach((t) => { if (t.cid) clearCell(ct, "계약ID", t.cid); else clearCell(cu, ID_COL, t.custId); });
    if (hist) {
      const cs = { t: hist.colOf("일시"), u: hist.colOf("고객ID"), k: hist.colOf("계약ID"), y: hist.colOf("유형"), res: hist.colOf("결과") };
      if (cs.t >= 0 && cs.u >= 0 && cs.res >= 0) {
        for (let r = hist.hr + 1; r <= hist.rg.e.r; r++) {
          const y = cs.y >= 0 ? hist.txt(r, cs.y) : "";
          if (monthOf(hist.txt(r, cs.t)) !== ym || !["통화", "문자"].includes(y) || hist.txt(r, cs.res) !== "시도") continue; // 앱의 '시도' 행만 취소 처리 (다른 통화·상담 기록은 그대로)
          const kid = cs.k >= 0 ? hist.txt(r, cs.k) : "", uid = hist.txt(r, cs.u);
          if (targets.some((t) => (t.cid ? kid === t.cid : kid === "" && uid === t.custId))) { hist.ws[hist.enc(r, cs.res)] = { t: "s", v: "취소(초기화)" }; n++; }
        }
      }
    }
    return n;
  }
  // 이력 시트의 한 줄을 지우고 아래 줄을 한 칸씩 올림
  function deleteSheetRow(info, r) {
    const { ws, enc, rg } = info;
    for (let rr = r; rr < rg.e.r; rr++) for (let c = rg.s.c; c <= rg.e.c; c++) { const nx = ws[enc(rr + 1, c)]; if (nx) ws[enc(rr, c)] = nx; else delete ws[enc(rr, c)]; }
    for (let c = rg.s.c; c <= rg.e.c; c++) delete ws[enc(rg.e.r, c)];
    rg.e.r -= 1;
    ws["!ref"] = XLSX.utils.encode_range(rg);
  }
  // 한 건·한 종류(전화/문자)의 이 달 '시도' 기록 삭제: 기기 대기 기록 + 시트 최근통화/문자일시 + 이력의 '시도' 행(앱 기록만, 실제 통화·상담 기록은 그대로)
  async function cancelContact(row, kind, ym) {
    const sameKey = (it) => (row.cid ? it.cid === row.cid : !it.cid && it.custId === row.custId);
    writeQueue(readQueue().filter((it) => !(it.kind === kind && monthOf(it.ts) === ym && sameKey(it))));
    if (!row.fileId) return;
    for (let i = 0; flushing && i < 25; i++) await sleep(200); // 대기 기록을 올리는 중이면 끝나길 잠깐 기다림
    const wb = await stGetBook(row.fileId); // 최신 파일
    const col = kind === "call" ? "최근통화일시" : "최근문자일시", yname = kind === "call" ? "통화" : "문자";
    const info = row.cid ? sheetInfo(wb, "계약", "계약ID") : sheetInfo(wb, CUST_SHEET, ID_COL), keyCol = row.cid ? "계약ID" : ID_COL, keyVal = row.cid || row.custId;
    if (info) {
      const kc = info.colOf(keyCol), cc = info.colOf(col);
      if (kc >= 0 && cc >= 0) for (let r = info.hr + 1; r <= info.rg.e.r; r++) if (info.txt(r, kc) === keyVal) { if (monthOf(info.txt(r, cc)) === ym) delete info.ws[info.enc(r, cc)]; break; }
    }
    const hist = sheetInfo(wb, "이력", "고객ID");
    if (hist) {
      const cs = { t: hist.colOf("일시"), u: hist.colOf("고객ID"), k: hist.colOf("계약ID"), y: hist.colOf("유형"), res: hist.colOf("결과") };
      if (cs.t >= 0 && cs.u >= 0 && cs.y >= 0 && cs.res >= 0) {
        for (let r = hist.rg.e.r; r > hist.hr; r--) { // 아래에서 위로 (지우면 줄이 올라오므로)
          if (monthOf(hist.txt(r, cs.t)) !== ym || hist.txt(r, cs.y) !== yname || hist.txt(r, cs.res) !== "시도") continue;
          const kid = cs.k >= 0 ? hist.txt(r, cs.k) : "", uid = hist.txt(r, cs.u);
          if (row.cid ? kid === row.cid : kid === "" && uid === row.custId) deleteSheetRow(hist, r);
        }
      }
    }
    await stPutBook(row.fileId, wb);
    delete bookCache[row.fileId];
    histCache = null; contractCache = null; listCache = null;
  }
  async function resetSalesMarks(rows, mark, ym, tabLabel, fcSel) {
    const targets = rows.filter((r) => r.key && mark.get(r.key) && r.fileId && !r.ref);
    const pending = readQueue().filter((it) => monthOf(it.ts) === ym && rows.some((r) => r.key === (it.cid ? "c:" + it.cid : "u:" + it.custId)));
    if (!targets.length && !pending.length) return toast("초기화할 시도 표시가 없습니다.");
    if (!confirm(`${ymLabel(ym)} · ${fcSel || "전체"} · ${tabLabel}의 전화·문자 '시도' 표시 ${new Set([...targets.map((r) => r.key), ...pending.map((it) => (it.cid ? "c:" + it.cid : "u:" + it.custId))]).size}건을 초기화할까요?\n\n· 목록의 ✓ 표시가 없어지고, 시트의 최근통화·최근문자 일시가 비워집니다.\n· 이력 시트의 기록은 지우지 않고 결과를 '취소(초기화)'로 바꿔 보존합니다.\n· 아직 저장 안 된 이 기기의 기록도 함께 지웁니다.`)) return;
    try {
      const byFile = new Map();
      targets.forEach((r) => { if (!byFile.has(r.fileId)) byFile.set(r.fileId, []); byFile.get(r.fileId).push(r); });
      for (const [fileId, ts] of byFile) {
        const wb = await stGetBook(fileId);
        const info = { ct: sheetInfo(wb, "계약", "계약ID"), cu: sheetInfo(wb, CUST_SHEET, ID_COL), hist: sheetInfo(wb, "이력", "고객ID") };
        clearContacts(info, ts, ym);
        await stPutBook(fileId, wb);
        delete bookCache[fileId];
      }
      const keys = new Set(rows.map((r) => r.key));
      writeQueue(readQueue().filter((it) => !(monthOf(it.ts) === ym && keys.has(it.cid ? "c:" + it.cid : "u:" + it.custId))));
      histCache = null; contractCache = null; listCache = null;
      toast("초기화했습니다.");
    } catch (e) { handleError(e); }
    renderSales(true);
  }

  // ---- 세일즈 화면 스크롤 복귀: 화면 맨 위에 보이는 행(key)과 그 행의 위치(오프셋)를 기준으로 저장·복원. 월·필터·탭·FC 변경은 맨 위 ----
  const SALES_SAVE_KEY = "shlife.salesResume";
  let SALES_BOOT = null; // 이 탭이 새로 로드됐을 때(브라우저가 탭을 내렸다가 다시 열 때 등) 직전 세일즈 화면 상태 (30분 이내, 같은 계정)
  try { const s = JSON.parse(sessionStorage.getItem(SALES_SAVE_KEY) || "null"); if (s && s.acc === gAccount() && Date.now() - s.t < 30 * 60 * 1000) SALES_BOOT = s; } catch (e) { /* 무시 */ }
  try { history.scrollRestoration = "manual"; } catch (e) { /* 무시 */ }
  const topbarH = () => { const t = document.querySelector(".topbar"); return t ? t.getBoundingClientRect().height : 0; };
  function salesAnchor() {
    const h = topbarH(), rows = [...document.querySelectorAll("#salesView .srow")];
    const i = rows.findIndex((el) => el.getBoundingClientRect().bottom > h + 1);
    if (i < 0) return { key: "", off: 0, after: [], y: Math.round(window.scrollY) };
    return { key: rows[i].dataset.fold || "", off: Math.round(rows[i].getBoundingClientRect().top - h), after: rows.slice(i + 1, i + 7).map((el) => el.dataset.fold || "").filter(Boolean), y: Math.round(window.scrollY) };
  }
  let salesRestoreStop = null;
  function salesRestore(a) {
    if (!a) return;
    if (salesRestoreStop) salesRestoreStop();
    let stop = false;
    const evs = ["touchstart", "wheel", "keydown", "mousedown"];
    const onUser = () => { stop = true; };
    evs.forEach((n) => window.addEventListener(n, onUser, { once: true, passive: true }));
    const halt = () => { stop = true; cleanup(); };
    const cleanup = () => { evs.forEach((n) => window.removeEventListener(n, onUser)); SALES.restoring = false; if (salesRestoreStop === halt) salesRestoreStop = null; };
    salesRestoreStop = halt;
    SALES.restoring = true;
    const apply = () => {
      if (stop) return;
      const els = [...document.querySelectorAll("#salesView .srow")];
      let el = null, off = a.off || 0;
      for (const k of [a.key, ...(a.after || [])]) { if (!k) continue; el = els.find((e) => e.dataset.fold === k); if (el) { if (k !== a.key) off = 0; break; } } // 앵커 행이 사라졌으면 그 다음 행
      const y = el ? el.getBoundingClientRect().top + window.scrollY - topbarH() - off : a.y || 0;
      window.scrollTo(0, Math.max(0, Math.round(y)));
    };
    apply();
    setTimeout(() => { apply(); setTimeout(() => { apply(); cleanup(); }, 150); }, 50); // 레이아웃이 자리 잡은 뒤 한 번 더 맞춤 (타이머: 화면이 가려져 있어도 멈추지 않음)
  }
  function salesSaveState() { // 새로 로드돼도 이어 보도록 (세일즈 화면을 보고 있을 때만)
    if ($("salesView").hidden || SALES.resume || SALES.fromNtf) return; // v564: 알림으로 온 화면은 남기지 않음
    try { sessionStorage.setItem(SALES_SAVE_KEY, JSON.stringify({ acc: gAccount(), t: Date.now(), tab: SALES.tab, ym: SALES.ym, fc: SALES.fc, dct: SALES.dct, dcls: SALES.dcls, dact: SALES.dact, df: SALES.df, open: [...SALES.open].slice(0, 300), a: salesAnchor() })); } catch (e) { /* 무시 */ }
  }
  function salesClearState() { try { sessionStorage.removeItem(SALES_SAVE_KEY); } catch (e) { /* 무시 */ } }
  let salesSaveTimer = 0;
  window.addEventListener("scroll", () => { clearTimeout(salesSaveTimer); salesSaveTimer = setTimeout(salesSaveState, 300); }, { passive: true });
  document.addEventListener("visibilitychange", () => { if (document.visibilityState === "hidden") salesSaveState(); });
  window.addEventListener("pagehide", salesSaveState);

  // ---- 일정 화면 이어보기: 새로고침해도 일정 화면(보던 날짜·Today/Week/Month)과 스크롤 위치로 돌아옴 (30분 이내, 같은 계정) ----
  const SCHED_SAVE_KEY = "shlife.schedResume";
  let SCHED_BOOT = null;
  try { const s = JSON.parse(sessionStorage.getItem(SCHED_SAVE_KEY) || "null"); if (s && s.acc === gAccount() && Date.now() - s.t < 30 * 60 * 1000) SCHED_BOOT = s; } catch (e) { /* 무시 */ }
  let schedHold = 0;       // 이 시각 전에는 저장하지 않음 (일정을 새로 그리는 동안 스크롤이 0으로 튀는 것을 저장하지 않게)
  let schedResumeY = null; // 새로 로드된 뒤 돌아갈 스크롤 위치
  function schedSaveState() {
    if ($("scheduleView").hidden || schedResumeY != null || Date.now() < schedHold) return;
    try { sessionStorage.setItem(SCHED_SAVE_KEY, JSON.stringify({ acc: gAccount(), t: Date.now(), mode: schedMode, d: +schedDate, y: Math.round(window.scrollY) })); } catch (e) { /* 무시 */ }
  }
  function schedClearState() { try { sessionStorage.removeItem(SCHED_SAVE_KEY); } catch (e) { /* 무시 */ } }
  function schedApplyResume() { // 일정이 그려진 뒤 호출: 내용이 자리 잡는 동안 몇 번 맞춤, 사용자가 손대면 멈춤
    schedHold = Date.now() + 1500;
    if (schedResumeY == null) return;
    const y = schedResumeY; schedResumeY = null;
    let stop = false;
    const evs = ["touchstart", "wheel", "keydown", "mousedown"];
    const onUser = () => { stop = true; };
    evs.forEach((n) => window.addEventListener(n, onUser, { once: true, passive: true }));
    const apply = () => { if (!stop) window.scrollTo(0, y); };
    apply();
    [60, 250, 700, 1500].forEach((ms) => setTimeout(apply, ms));
    setTimeout(() => { evs.forEach((n) => window.removeEventListener(n, onUser)); }, 1600);
  }
  let schedSaveTimer = 0;
  window.addEventListener("scroll", () => { clearTimeout(schedSaveTimer); schedSaveTimer = setTimeout(schedSaveState, 300); }, { passive: true });
  document.addEventListener("visibilitychange", () => { if (document.visibilityState === "hidden") schedSaveState(); });
  window.addEventListener("pagehide", schedSaveState);

  let salesReq = 0;
  // keep: true = 지금 위치 그대로(칩·초기화·자동 반영), 위치 객체 = 그 위치로(고객 카드에서 돌아올 때), 없으면 맨 위(월·필터·탭·FC 변경)
  async function renderSales(keep) { // 다시 그리는 동안 화면 높이를 유지해, 세로 스크롤바가 사라졌다 생기며 화면이 출렁이지 않게 함
    const box0 = $("salesView");
    if (box0 && box0.offsetHeight > 0) box0.style.minHeight = box0.offsetHeight + "px";
    try { await renderSalesInner(keep); } finally { if (box0) box0.style.minHeight = ""; }
  }
  async function renderSalesInner(keep) {
    const req = ++salesReq;
    let anchor = keep && keep !== true ? keep : keep ? salesAnchor() : null;
    if (!anchor && SALES.resume) anchor = SALES.resume;
    SALES.resume = null;
    show("salesView", !!keep);
    const box = $("salesView");
    if (SALES.sub === "material") { await renderMaterial(req); return; } // 세일즈 > 보고서: 조건 보고서 (v360)
    if (SALES.sub === "data") { await renderMat(req); return; } // 세일즈 > 자료 (v554)
    if (!SALES.ym) SALES.ym = ymNow();
    let ym = SALES.ym;
    const top = (fcHtml) => `<div class="sales-top">
        <div class="seg"><button type="button" data-stab="dda" class="${SALES.tab === "dda" ? "on" : ""}">따행</button><button type="button" data-stab="transfer" class="${SALES.tab === "transfer" ? "on" : ""}">이관</button><button type="button" data-stab="valueup" class="${SALES.tab === "valueup" ? "on" : ""}">DB</button></div>
        <div class="sline2 sl3${SALES.userHtml || fcHtml ? "" : " center"}"><div class="ymnav">${SALES.stateTab ? `<b>이달목표 · ${ymLabel(ym)}</b>` : `<button type="button" data-ym="-1" aria-label="이전 달"><svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M15 6l-6 6 6 6"/></svg></button><b>${ymLabel(ym)}</b><button type="button" data-ym="1" aria-label="다음 달"><svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M9 6l6 6-6 6"/></svg></button>`}${top.note || ""}</div><span class="sl-rt">${SALES.userHtml || ""}${fcHtml}<select id="salesTab" aria-label="따행·이관·DB">${[["dda", "따행"], ["transfer", "이관"], ["valueup", "DB"]].map(([k, l]) => `<option value="${k}" ${SALES.tab === k ? "selected" : ""}>${l}</option>`).join("")}</select></span></div></div>`; // v348: 넓은 화면 = 월 이동 가운데, 오른쪽 [사용자][FC][따행]. v345(2026-11-15 사용자 지시): 따행·이관·DB 탭 → 월 줄 오른쪽 드롭박스, 사용자·FC 선택도 같은 줄
    if (!keep || !box.querySelector("#salesBody .slist, #salesBody .sales-info")) box.innerHTML = top(SALES.fcHtml || "") + `<div id="salesBody"><p class="hint pad">불러오는 중…</p></div>`; // 불러오는 동안에도 지난번 FC 선택창을 그대로 두어 월 줄이 위아래로 흔들리지 않게 // keep이면 기존 목록을 그대로 두고 새 내용으로 바꿈
    try {
      const data = await loadSalesData();
      if (req !== salesReq) return;
      await loadMyPrefs().catch(() => {});
      await refreshSalesExtra();
      await smemoPrep(); // v411: 메모 칸을 켠 경우만 상담기록 파일을 읽음
      // v394(2026-11-17, 확정안 5-3·5-4): 이관상태·따행상태 칸이 있는 파일은 '목표' 상태 기준 목록(달 이동 없음, 이번 달 것이 위·지난 목표가 아래)
      { const hasTrState = data.contracts.some((c) => c.row && "이관상태" in c.row), hasFoState = (data.follow ? data.follow.rows : []).some((f) => "따행상태" in f);
        SALES.stateTab = (SALES.tab === "transfer" && hasTrState) || (SALES.tab === "dda" && hasFoState);
        if (SALES.stateTab) { SALES.ym = ym = ymNow(); } }
      const myFc = await myFcName();
      if (req !== salesReq) return;
      // FC 선택: 접근 가능한 FC들(고객이 등록된 FC). 기본은 내 FC(권한 파일 이름), 없으면 첫 FC. 선택은 이 기기에 기억
      const fcs = [...new Set([...data.lists.map((e) => e.fc), ...data.contracts.map((c) => c.srcFc), ...knownFcs].filter(Boolean))].sort((a, b) => { const o = CFG.fcOrder || [], ia = o.indexOf(a), ib = o.indexOf(b); return (ia < 0 ? 999 : ia) - (ib < 0 ? 999 : ib) || a.localeCompare(b, "ko"); }); // 직급 먼저 → 사번순(config.fcOrder)
      let saved = ""; try { saved = localStorage.getItem(salesFcKey()) || ""; } catch (e) { /* 무시 */ }
      // 선택창에 나오는 FC: 모든 계정(관리자 포함) = 내 FC + 관리자 화면에서 '세일즈 페이지 접근 허용'에 체크한 FC (고객 폴더 접근 권한이 있는 FC만)
      let selFcs = [...new Set([myFc, ...salesExtra])].filter((f) => f && fcs.includes(f)); // 일반 사용자: 본인 FC + 관리자 화면 '세일즈 페이지 접근 허용'에 체크한 FC
      let userHtml = "";
      if (isAdmin() && salesPermsAll) { // v325(2026-11-15 사용자 지시): 관리자는 모든 사용자를 고르고, 고른 사용자의 '세일즈 페이지 접근 허용' FC만 봄
        const P = salesPermsAll, meE = me(), ord = CFG.fcOrder || [];
        const ownOf = (em) => (em === meE ? myFc : String((P.names || {})[em] || "").trim());
        const fcsOf = (em) => [...new Set([ownOf(em), ...(Array.isArray((P.salesFcs || {})[em]) ? P.salesFcs[em] : [])])].filter((f) => f && fcs.includes(f));
        const users = [...new Set([meE, ...Object.keys(P.users || {}), ...Object.keys(P.names || {}), ...Object.keys(P.salesFcs || {})])].filter(Boolean)
          .map((em) => ({ em, name: ownOf(em) || em.split("@")[0], fcs: fcsOf(em) })).filter((u) => u.fcs.length)
          .reduce((a, u, i, all) => { a.push(u); if (i === all.length - 1) fcs.filter((f) => !all.some((x) => x.fcs.includes(f))).forEach((f) => a.push({ em: "fc:" + f, name: f, fcs: [f] })); return a; }, []) // v512: 계정(이메일)이 없는 팀원 FC도 관리자가 고를 수 있게
          .sort((a, b) => (b.em === meE) - (a.em === meE) || ((ord.indexOf(a.name) + 1 || 999) - (ord.indexOf(b.name) + 1 || 999)) || a.name.localeCompare(b.name, "ko"));
        let su = SALES.user; if (!su) { try { su = localStorage.getItem(salesUserKey()) || ""; } catch (e) { /* 무시 */ } }
        if (SALES.goFc) { const w = SALES.goFc, cur = users.find((u) => u.em === su); if (!cur || !cur.fcs.includes(w)) { const hit = users.find((u) => u.name === w) || users.find((u) => u.fcs.includes(w)); if (hit) su = hit.em; } } // 말만해요에서 FC를 지정해 연 경우
        if (!users.some((u) => u.em === su)) su = (users.find((u) => u.em === meE) || users[0] || {}).em || "";
        SALES.user = su;
        const cu = users.find((u) => u.em === su);
        if (cu) { selFcs = cu.fcs; if (SALES.fc === "__own") SALES.fc = cu.fcs[0]; }
        const dup = (n) => users.filter((u) => u.name === n).length > 1;
        if (users.length > 1) userHtml = `<select id="salesUser" class="sales-user" aria-label="사용자 선택">${users.map((u) => `<option value="${esc(u.em)}"${u.em === su ? " selected" : ""}>${esc(u.name + (dup(u.name) ? " (" + u.em.split("@")[0] + ")" : ""))}</option>`).join("")}</select>`;
      }
      if (SALES.fc === "__own") SALES.fc = "";
      if (selFcs.length <= 1) SALES.fc = selFcs[0] || (fcs.includes(myFc) ? myFc : fcs[0] || ""); // 추가 FC가 없으면 선택창 없이 본인 FC
      else if (!SALES.fc || !selFcs.includes(SALES.fc)) SALES.fc = selFcs.includes(saved) ? saved : selFcs.includes(myFc) ? myFc : selFcs[0] || "";
      if (SALES.goFc !== undefined) { // 말만해요에서 연 경우: 저장해 둔 선택 대신 지정한 FC(없으면 로그인한 본인 FC). 저장해 둔 선택은 덮어쓰지 않음
        const want = SALES.goFc || myFc; SALES.goFc = undefined;
        if (selFcs.includes(want)) SALES.fc = want; else if (selFcs.includes(myFc)) SALES.fc = myFc;
      }
      const fcSel = SALES.fc;
      if (SALES.sub === "prog") { // 세일즈 > 계약진행 (이달목표·자료와 같은 FC 선택을 씀)
        const fcHtmlP = selFcs.length > 1 ? `<select id="salesFc" aria-label="FC 선택">${selFcs.map((f) => `<option value="${esc(f)}"${f === fcSel ? " selected" : ""}>${esc(f)}</option>`).join("")}</select>` : "";
        SALES.data = data; SALES.userHtml = userHtml; SALES.fcHtml = fcHtmlP;
        renderProgPage(box, data, fcSel, userHtml, fcHtmlP);
        if (anchor) salesRestore(anchor); else if (!keep) window.scrollTo(0, 0);
        return;
      }
      let baseRows = (SALES.tab === "transfer" ? transferRows(data, ym, SALES.showRef) : SALES.tab === "valueup" ? valueupRows(data) : ddaRows(data, ym)).filter((r) => !fcSel || r.srcFc === fcSel);
      baseRows.forEach((r) => { r.arch = SALES.tab === "valueup" ? !!r.dbOff : !!r.archDay || !!r.del; });
      const isDbTab = SALES.tab === "valueup";
      const dbv = isDbTab ? (SALES.dbv || (SALES.arch.valueup ? "arch" : "new")) : ""; // DB 탭: new(신규) / ctr(계약) / arch(보관) / del(삭제)
      const arState = isDbTab ? dbv !== "new" : (SALES.arch[SALES.tab] || false); // 이 탭에서 보관(·삭제) 목록을 보는 중인지
      const arCnt = isDbTab ? { on: baseRows.filter((r) => !r.dbOff).length, ctr: baseRows.filter((r) => r.ctr).length, off: baseRows.filter((r) => r.dbOff && !r.del && !r.ctr).length, del: baseRows.filter((r) => r.del).length } : { on: baseRows.filter((r) => !r.arch).length, off: baseRows.filter((r) => r.arch && !r.del).length, del: baseRows.filter((r) => r.arch && r.del).length };
      baseRows = isDbTab ? baseRows.filter((r) => dbv === "del" ? r.del : dbv === "ctr" ? r.ctr : dbv === "arch" ? r.dbOff && !r.del && !r.ctr : !r.dbOff) : baseRows.filter((r) => (arState === "del" ? r.arch && r.del : arState ? r.arch && !r.del : !r.arch)); // v416: 신규 / 완료 / 지옥 // [이관/따행/DB N] [보관 M] (DB는 [삭제 K] 포함) 칩으로 바꿔 봄
      if (dbv === "del") { const dk = (r) => tsDigits(r.delDay) || "99999999999999"; baseRows.sort((a, b) => dk(b).localeCompare(dk(a)) || a.holder.localeCompare(b.holder, "ko")); } // 삭제 목록: 가장 최근에 삭제한 고객이 맨 위(일시 내림차순, 방금 삭제해 이력이 아직 없으면 맨 위)
      const arLabel = SALES.tab === "transfer" ? "이관" : SALES.tab === "dda" ? "따행" : "DB";
      const arBar = isDbTab ? `<span class="prog-bar" role="group" aria-label="DB 보기"><button type="button" data-dbv="new" class="${dbv === "new" ? "on" : ""}">신규 ${arCnt.on}</button><button type="button" data-dbv="ctr" class="${dbv === "ctr" ? "on" : ""}">계약 ${arCnt.ctr}</button><button type="button" data-dbv="arch" class="${dbv === "arch" ? "on" : ""}">보관 ${arCnt.off}</button><button type="button" data-dbv="del" class="${dbv === "del" ? "on" : ""}">지옥 ${arCnt.del}</button></span>` : `<span class="prog-bar" role="group" aria-label="${arLabel} 보기"><button type="button" data-arview="on" class="${arState ? "" : "on"}">신규 ${arCnt.on}</button><button type="button" data-arview="off" class="${arState === true ? "on" : ""}">완료 ${arCnt.off}</button><button type="button" data-arview="del" class="${arState === "del" ? "on" : ""}">지옥 ${arCnt.del}</button></span>`;
      // 따행: 활동대상 → 접촉현황 → 따행분류 (SOL 조회 조건과 같은 기준: 대상년월 · 활동대상 기본)
      const bMode = SALES.tab === "dda" && baseRows.some((r) => r.follow);
      let allRows = baseRows, ddaTotals = null, classes = [];
      if (SALES.tab === "dda") ({ rows: allRows, totals: ddaTotals, classes } = ddaFilter(baseRows, bMode, { dct: "", dcls: "", dact: true })); // v94: 접촉현황·분류 전체, 활동대상 기본 고정 (화면에서 바꾸지 않음)
      const ymNote = "";
      // 따행 필터: 설문발송·설문응답·쿠폰(체크한 것 모두 채워진 고객만, AND)
      const df = SALES.df, dfOn = SALES.tab === "dda" && (df.send || df.resp || df.coupon);
      const rows0 = dfOn ? allRows.filter((r) => (!df.send || r.dda.send) && (!df.resp || r.dda.resp) && (!df.coupon || r.dda.coupon)) : allRows;
      const mark = contactMarks(data, ym);
      // 전화·문자 필터: 체크한 것을 아직 안 한(시도 ✓ 없는) 고객만 (둘 다 체크하면 둘 다 안 한 고객)
      const nf = SALES.nf, nfOn = nf.call || nf.sms;
      const rows = nfOn ? rows0.filter((r) => { const mk = (r.key && mark.get(r.key)) || {}; return (!nf.call || !mk.call) && (!nf.sms || !mk.sms); }) : rows0;
      const nfBox = (k, l) => `<label><input type="checkbox" data-nf="${k}" ${nf[k] ? "checked" : ""}> ${l}</label>`;
      const nfHtml = `<div class="nfilter"><span class="dfl">${nfBox("call", "전화")}${nfBox("sms", "문자")}</span></div>`;
      const ctx = { tpl: smsTemplate(SALES.tab === "transfer" ? "transfer" : "dda"), fcName: fcSel || myFc };
      const holders = new Set(rows.map((r) => r.custId || r.holder)).size;
      const refN = SALES.tab === "transfer" ? data.contracts.filter((c) => c.type === "이관" && monthOf(c.row["이관일자"] || c.row["이관일"]) === ym && c.fc && c.srcFc && c.fc !== c.srcFc && (!fcSel || c.srcFc === fcSel)).length : 0;
      const marked = rows.filter((r) => r.key && mark.get(r.key)).length;
      SALES.data = data;
      top.note = ymNote;
      SALES.userHtml = userHtml; SALES.fcHtml = (selFcs.length > 1 ? `<select id="salesFc" aria-label="FC 선택">${selFcs.map((f) => `<option value="${esc(f)}"${f === fcSel ? " selected" : ""}>${esc(f)}</option>`).join("")}</select>` : "");
      box.innerHTML = top(SALES.fcHtml) +
        `<div id="salesBody">${SALES.tab === "dda" ? `<div class="sales-info"><span>${arState === "del" ? "지옥 " : arState ? "완료 " : ""}${rows.length}명</span><span class="dfl dfl-in"><span class="dfg"><b class="dfk">설문</b><label><input type="checkbox" data-df="send" ${df.send ? "checked" : ""}> 발송</label><label><input type="checkbox" data-df="resp" ${df.resp ? "checked" : ""}> 응답</label></span><i class="dsep" aria-hidden="true"></i><span class="dfg"><label><input type="checkbox" data-df="coupon" ${df.coupon ? "checked" : ""}> 쿠폰</label></span><i class="dsep" aria-hidden="true"></i><span class="dfg">${nfBox("call", "전화")}${nfBox("sms", "문자")}</span></span>${arBar}</div>` : `<div class="sales-info${dbv === "arch" || dbv === "del" ? " db4" : ""}"><span>${SALES.tab === "dda" ? (bMode ? `${(SALES.dct || SALES.dcls || dfOn) ? `필터 ${rows.length}명` : `표시 ${rows.length}명`}` : `전체 ${allRows.length}명(미완료 ${allRows.filter((r) => !(r.dda && r.dda.done)).length}명)${dfOn ? ` · 필터 ${rows.length}명` : ""}`) : (SALES.tab === "valueup" ? `${dbv === "del" ? "삭제 " : dbv === "ctr" ? "계약 " : arState ? "보관 " : ""}${rows.length}명` : `${arState ? "보관 " : ""}${rows.length}건 · ${holders}명`)}${marked ? ` · 시도 ✓ ${marked}` : ""}${SALES.tab === "dda" ? ` <span class="hint">(시도 ✓ 기준 ${ymLabel(ym)})</span>` : ""}</span>${SALES.tab === "valueup" ? `<span class="sinfo-r"><span class="nfl">${nfBox("call", "전화")}${nfBox("sms", "문자")}</span>${arBar}</span>` : ""}${SALES.tab === "transfer" ? `<span class="sinfo-r"><span class="nfl">${nfBox("call", "전화")}${nfBox("sms", "문자")}</span>${arBar}${refN ? `<label class="ref-toggle"><input type="checkbox" id="salesRef" ${SALES.showRef ? "checked" : ""}> 참고 포함(${refN})</label>` : ""}</span>` : ""}</div>`}` +
        
        (rows.length ? salesRowsHtml(rows, mark, ctx, SALES.open) : `<p class="hint pad">${SALES.tab === "transfer" ? ymLabel(ym) + " " : ""}${fcSel ? fcSel + " " : ""}${SALES.tab === "transfer" ? "이관" : SALES.tab === "valueup" ? (dbv === "new" ? "신규 DB" : "DB") : "따행"}${dbv === "del" ? " 지옥" : dbv === "ctr" ? " 계약" : isDbTab ? (arState ? " 보관" : "") : arState === "del" ? " 지옥" : arState ? " 완료" : " 신규"} 고객이 없습니다.${nfOn ? " (전화·문자 필터를 확인해 주세요)" : ""}</p>`) +
        (SALES.tab === "valueup" || !readQueue().length ? "" : `<p class="hint pad">아직 저장 안 된 기록 ${readQueue().length}건은 앱이 다시 열리면 반영합니다.</p>`) /* v422: '시도' 안내문구는 없앰 */ + `</div>`;
      box._rows = rows; box._mark = mark;
      if (anchor) salesRestore(anchor); else if (!keep) window.scrollTo(0, 0);
    } catch (e) { if (req === salesReq) { $("salesBody").innerHTML = ""; handleError(e); } }
  }
  // ---------- 세일즈 > 보고서: 조건 보고서 (v360, 2026-11-16 사용자 지시) ----------
  // 사용자(내 권한 범위) · 따행/이관/DB(복수) · 보관 포함 · 년월 → 건수를 보여 준 뒤 [진행] → 보고서를 화면에 보여 주고 아래에 엑셀·PDF 다운, PDF 열람, 이메일(자신에게)
  // 고객 정보는 이 기기 안에서만 만들고 어디에도 올리지 않음(이메일·공유는 사용자가 직접 누를 때만)
  const RPT = { opt: "", types: { dda: true, transfer: true, valueup: true }, st: { new: true, ctr: false, done: false, hell: false }, span: 1, y: 0, m: 0, built: null, data: null, opts: [], cache: null };
  const RPT_S = [["new", "신규"], ["ctr", "계약"], ["done", "완료"], ["hell", "지옥"]]; // v424: 계약 = DB의 계약 고객(DB 탭의 [계약] 칩과 같음, 이관·따행에는 없음)
  const rptStOf = (type, r) => (r.del ? "hell" : type === "valueup" && r.ctr ? "ctr" : (type === "valueup" ? r.dbOff : r.arch) ? "done" : "new"); // DB: 계약 = 계약, 보관 = 완료, 삭제 = 지옥
  const RPT_T = [["dda", "따행"], ["transfer", "이관"], ["valueup", "DB"]];
  const RPT_COLS = ["구분", "FC", "이름", "피보", "상품", "보험료", "기준일", "계약일", "생일", "연락처", "직업", "주소", "모집FC", "상태", "진행", "단계"];
  async function rptOptions(data) {
    const myFc = await myFcName();
    const fcs = [...new Set([...data.lists.map((e) => e.fc), ...data.contracts.map((c) => c.srcFc), ...knownFcs].filter(Boolean))];
    let opts = [];
    if (isAdmin() && salesPermsAll) { // 관리자: 모든 사용자 (고른 사용자의 '세일즈 페이지 접근 허용' FC까지)
      const P = salesPermsAll, meE = me(), ord = CFG.fcOrder || [];
      const ownOf = (em) => (em === meE ? myFc : String((P.names || {})[em] || "").trim());
      const fcsOf = (em) => [...new Set([ownOf(em), ...(Array.isArray((P.salesFcs || {})[em]) ? P.salesFcs[em] : [])])].filter((f) => f && fcs.includes(f));
      opts = [...new Set([meE, ...Object.keys(P.users || {}), ...Object.keys(P.names || {}), ...Object.keys(P.salesFcs || {})])].filter(Boolean)
        .map((em) => ({ id: em, label: ownOf(em) || em.split("@")[0], fcs: fcsOf(em) })).filter((u) => u.fcs.length)
        .reduce((a, u, i, all) => { a.push(u); if (i === all.length - 1) fcs.filter((f) => !all.some((x) => x.fcs.includes(f))).forEach((f) => a.push({ id: "fc:" + f, label: f, fcs: [f] })); return a; }, []) // v512: 계정 없는 팀원 FC도
        .sort((a, b) => (b.id === meE) - (a.id === meE) || ((ord.indexOf(a.label) + 1 || 999) - (ord.indexOf(b.label) + 1 || 999)) || a.label.localeCompare(b.label, "ko"));
    } else { // 일반 사용자: 내 FC + 관리자가 '세일즈 페이지 접근 허용'한 FC
      opts = [...new Set([myFc, ...salesExtra])].filter((f) => f && fcs.includes(f)).map((f) => ({ id: f, label: f, fcs: [f] }));
    }
    if (!opts.length && fcs.length) opts = [{ id: fcs[0], label: fcs[0], fcs: [fcs[0]] }];
    if (opts.length > 1) opts.unshift({ id: "__all", label: "전체", fcs: [...new Set(opts.flatMap((o) => o.fcs))] }); // v426(2026-11-18 사용자 지시): '전체'가 맨 위
    return opts;
  }
  const rptYm = () => RPT.y + pad2(RPT.m);
  const rptMonths = () => Array.from({ length: RPT.span || 1 }, (_, i) => ymShift(rptYm(), -i)); // 끝 달(고른 달)부터 거꾸로
  const rptPeriod = () => { const ms = rptMonths(); return ms.length === 1 ? ymLabel(ms[0]) : `${ymLabel(ms[ms.length - 1])} ~ ${ymLabel(ms[0])} (${RPT.span === 12 ? "1년" : RPT.span + "개월"})`; };
  const rptPeriodKey = () => { const ms = rptMonths(); return ms.length === 1 ? ms[0] : ms[ms.length - 1] + "-" + ms[0]; };
  function rptCollect() { // 선택한 사용자·기간의 따행/이관/DB 행(신규·완료·지옥 모두). 같은 조건이면 다시 만들지 않음
    const o = RPT.opts.find((x) => x.id === RPT.opt) || { id: "", fcs: [] }, ms = rptMonths(), key = o.id + "|" + ms.join(",");
    if (RPT.cache && RPT.cache.key === key) return RPT.cache.rows;
    const data = RPT.data, inF = (r) => o.fcs.includes(r.srcFc), mset = new Set(ms);
    const uniq = (arr) => { const seen = new Set(); return arr.filter((r) => { const k = (r.key || [r.srcFc, r.holder, r.idate || "", r.product || ""].join("/")) + "|" + (r.follow ? ymOfFollow(r.follow["대상년월"]) : ""); if (seen.has(k)) return false; seen.add(k); return true; }); };
    // 새 구조(상태 칸)는 달과 상관없이 목표·보관·삭제 전체를 주므로 여기서 기간으로 거름: 이관 = 이관일, 따행 = 대상년월
    const dda = uniq(ms.flatMap((ym) => ddaRows(data, ym))).filter((r) => inF(r) && (!r.follow || mset.has(ymOfFollow(r.follow["대상년월"]))));
    dda.forEach((r) => { r.arch = !!r.archDay || !!r.del; });
    const tr = uniq(ms.flatMap((ym) => transferRows(data, ym, false))).filter((r) => inF(r) && mset.has(monthOf(r.idate)));
    tr.forEach((r) => { r.arch = !!r.archDay || !!r.del; });
    const db = valueupRows(data).filter(inF); db.forEach((r) => { r.arch = !!r.dbOff; });
    const rows = { dda: ddaFilter(dda, dda.some((r) => r.follow), { dct: "", dcls: "", dact: true }).rows, transfer: tr, valueup: db };
    RPT.cache = { key, rows };
    return rows;
  }
  function rptSelected() { // 체크한 유형 × 체크한 단계(신규·완료·지옥)
    const all = rptCollect(), out = {};
    RPT_T.forEach(([k]) => { if (RPT.types[k]) out[k] = all[k].filter((r) => RPT.st[rptStOf(k, r)]); });
    return out;
  }
  const dshow = (v) => { const k = dateKey(v); return k.length >= 8 ? k.slice(0, 4) + "-" + k.slice(4, 6) + "-" + k.slice(6, 8) : String(v == null ? "" : v).trim(); };
  function rptRec(type, r, ymLab) {
    const nz = (v) => { const t = String(v == null ? "" : v).trim(); const n = Number(t.replace(/[^\d.-]/g, "")); return t && isFinite(n) && /\d/.test(t) ? n : t; };
    let fee = nz(r.fee), prod = r.product || "";
    if (r.prods && r.prods.length > 1) { prod = r.prods.map((p) => p.name).join(" / "); fee = r.prods.reduce((s, p) => s + (typeof nz(p.fee) === "number" ? nz(p.fee) : 0), 0); }
    const prog = type === "dda" ? ((DCHIPS.slice().reverse().find(([f]) => r.dda && r.dda[f]) || [])[1] || "") : (r.stage || "");
    const sk = rptStOf(type, r), sd = type === "valueup" ? (r.del ? dshow(r.delDay) : r.dbAuto ? "" : dshow(r.dbDay)) : dshow(r.archDay);
    const arch = sk === "new" ? "신규" : sk === "ctr" ? "계약" : (sk === "hell" ? "지옥" : "완료") + (sd ? " " + sd : "");
    return { 구분: RPT_T.find((t) => t[0] === type)[1], FC: r.srcFc || "", 이름: r.holder || "", 피보: r.ins || "", 상품: prod, 보험료: fee, 기준일: type === "transfer" ? dshow(r.idate) : type === "valueup" ? dshow(r.cdate) : (r.follow && r.follow["대상년월"] ? ymLabel(ymOfFollow(r.follow["대상년월"])) : ymLab),
      ...(type === "valueup" ? { "성별/나이": r.sexAge || "", 유입회사: r.comp || "", 유입일: String(r.inAt || "").slice(0, 10), 유입메모: r.dbMemo || "", 지역: String((r.ent && (fieldOf(r.ent, "지역") || fieldOf(r.ent, "주소"))) || r.addr || "").replace(/^\[[^\]]*\]\s*/, "").split(/\s+/).slice(0, 2).join(" "), 유입DB명: prod, 메모: "", 성향: "", 통화: "", 미팅: "" } : {}), // v570: 밸류업DB 보고서 칸(사용자 지시)
      계약일: type === "valueup" ? "" : dshow(r.cdate), 생일: dshow(r.bday), 연락처: fmtPhoneText(r.phoneText || r.phone || ""), 직업: r.job || "", 주소: r.addr || "", 모집FC: r.recFc || "", 상태: r.status || "", 진행: prog, 단계: arch };
  }
  function rptBuild() {
    const sel = rptSelected(), o = RPT.opts.find((x) => x.id === RPT.opt) || { label: "", fcs: [] }, ymLab = ymLabel(rptYm()), stLab = RPT_S.filter(([k]) => RPT.st[k]).map(([, l]) => l).join("·");
    // v575(2026-11-25 사용자 지시): 유형마다 고른 단계별로 따로 만듦(신규 → 계약 → 완료 → 지옥 순), 제목에 단계 표시 "2026년 11월 밸류업DB 리스트 - 신규"
    const full = { dda: "따행", transfer: "이관", valueup: "밸류업DB" };
    let groups = RPT_T.filter(([k]) => sel[k]).flatMap(([k, l]) => RPT_S.filter(([sk]) => RPT.st[sk] && (k === "valueup" || sk !== "ctr")).map(([sk, sl]) => ({ type: k, st: sk, label: l + "-" + sl, title: `${rptPeriod()} ${full[k]} 리스트 - ${sl}`, recs: sel[k].filter((r) => rptStOf(k, r) === sk).map((r) => rptRec(k, r, ymLab)) })));
    if (groups.some((g) => g.recs.length)) groups = groups.filter((g) => g.recs.length); // 빈 단계는 빼고(모두 비면 그대로 두어 '없음' 표시)
    const total = groups.reduce((s, g) => s + g.recs.length, 0);
    const cond = `${o.label} · ${rptPeriod()} · ${groups.map((g) => g.label).join("·")} · ${stLab}`;
    const now = new Date(), stamp = now.getFullYear() + "-" + pad2(now.getMonth() + 1) + "-" + pad2(now.getDate()) + " " + pad2(now.getHours()) + ":" + pad2(now.getMinutes());
    return { title: "세일즈 보고서", cond, stamp, user: o.label, ym: rptPeriodKey(), groups, total, multiFc: o.fcs.length > 1 };
  }
  const rptView = (rep, type) => type === "valueup" ? ["구분", ...(rep.multiFc ? ["FC"] : []), "이름", "성별/나이", "연락처", "유입회사", "유입일", "유입메모", "지역", "진행", "단계"] : ["구분", ...(rep.multiFc ? ["FC"] : []), "이름", "상품", "보험료", "기준일", "연락처", "주소", "진행", "단계"];
  const RPT_DBP = ["메모", "No", "이름", "성별/나이", "연락처", "유입회사", "유입일", "유입메모", "지역", "성향", "통화", "미팅"]; // v573(2026-11-24 사용자 지시, 리더 전달): 유입DB명 칸 삭제, 유입메모·지역은 잘리지 않게 줄바꿈 // v570: 밸류업DB 인쇄(엑셀·PDF) 칸 — 맨 왼쪽 손글씨 메모, 성향·통화·미팅은 손으로 적는 칸
  const rptCell = (c, v) => (c === "보험료" && typeof v === "number" ? v.toLocaleString("ko-KR") : esc(String(v == null ? "" : v)));
  function rptTableHtml(rep) {
    return rep.groups.map((g) => { const cols = rptView(rep, g.type).filter((c) => c !== "구분"); return `<div class="rpt-grp"><div class="rpt-gh"><b>${esc(g.title || g.label)}</b> <span>${g.recs.length}건</span></div>` + (g.recs.length
      ? `<div class="rpt-tw"><table class="rpt-t${g.type === "valueup" ? " rpt-db" : ""}"><thead><tr><th>#</th>${cols.map((c) => `<th>${c}</th>`).join("")}</tr></thead><tbody>${g.recs.map((r, i) => `<tr><td>${i + 1}</td>${cols.map((c) => `<td class="${c === "보험료" ? "n" : c === "유입메모" ? "dbmemo" : c === "지역" ? "dbarea" : ""}">${rptCell(c, r[c])}</td>`).join("")}</tr>`).join("")}</tbody></table></div>`
      : `<p class="hint pad">해당 고객이 없습니다.</p>`) + `</div>`; }).join("");
  }
  const rptFileName = (rep, ext) => { const n = new Date(); return `${n.getFullYear()}${pad2(n.getMonth() + 1)}${pad2(n.getDate())} - 세일즈 보고서 - ${rep.user} - ${rep.ym}.${ext}`.replace(/[\\/:*?"<>|]/g, ""); };
  function rptSave(blob, name) { const a = document.createElement("a"), u = URL.createObjectURL(blob); a.href = u; a.download = name; document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(u), 60000); }
  function rptXlsxBlob(rep) {
    if (typeof XLSX === "undefined") throw new Error("엑셀 도구를 아직 불러오지 못했어요. 잠시 뒤 다시 눌러 주세요.");
    const wb = XLSX.utils.book_new();
    const sum = [["세일즈 보고서"], ["조건", rep.cond], ["만든 시각", rep.stamp], [], ["구분", "건수"], ...rep.groups.map((g) => [g.label, g.recs.length]), ["합계", rep.total]];
    const ws0 = XLSX.utils.aoa_to_sheet(sum); ws0["!cols"] = [{ wch: 12 }, { wch: 48 }]; XLSX.utils.book_append_sheet(wb, ws0, "요약");
    rep.groups.forEach((g) => {
      if (g.type === "valueup") { // v570: 밸류업DB = 인쇄 양식 칸(손글씨 메모·성향·통화·미팅은 빈 칸), 가로·폭 한 쪽
        const cols = [...(rep.multiFc ? ["FC"] : []), ...RPT_DBP];
        const ws = XLSX.utils.aoa_to_sheet([cols, ...g.recs.map((r, i) => cols.map((c) => (c === "No" ? i + 1 : c === "FC" ? r.FC : r[c] == null ? "" : r[c])))]);
        ws["!cols"] = cols.map((c) => ({ wch: { 메모: 22, No: 4, 이름: 9, "성별/나이": 8, 연락처: 14, 유입회사: 9, 유입일: 11, 유입메모: 56, 지역: 16, 성향: 5, 통화: 5, 미팅: 5, FC: 7 }[c] || 9 }));
        ws["!rows"] = [{ hpt: 18 }, ...g.recs.map((r) => ({ hpt: Math.max(30, 15 * Math.ceil(String(r["유입메모"] || "").length / 34)) }))]; // 긴 유입메모는 줄 높이를 늘림
        ws["!pageSetup"] = { orientation: "landscape", fitToWidth: 1, fitToHeight: 0, paperSize: 9 };
        XLSX.utils.book_append_sheet(wb, ws, g.label);
        return;
      }
      const ws = XLSX.utils.aoa_to_sheet([RPT_COLS, ...g.recs.map((r) => RPT_COLS.map((c) => r[c]))]);
      ws["!cols"] = RPT_COLS.map((c) => ({ wch: { 주소: 36, 상품: 28, 이름: 10, 연락처: 14, 보험료: 11 }[c] || 11 }));
      const fi = RPT_COLS.indexOf("보험료"); g.recs.forEach((r, i) => { const cell = ws[XLSX.utils.encode_cell({ r: i + 1, c: fi })]; if (cell && cell.t === "n") cell.z = "#,##0"; });
      XLSX.utils.book_append_sheet(wb, ws, g.label);
    });
    return new Blob([XLSX.write(wb, { type: "array", bookType: "xlsx", compression: true })], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  }
  const rptLibs = {};
  const rptLoad = (u) => rptLibs[u] || (rptLibs[u] = new Promise((res, rej) => { const s = document.createElement("script"); s.src = u; s.onload = res; s.onerror = () => { delete rptLibs[u]; rej(new Error("PDF 도구를 불러오지 못했어요. 인터넷 연결을 확인해 주세요.")); }; document.head.appendChild(s); }));
  async function rptPdfBlob(rep) { // 보고서를 화면 글꼴 그대로 그려 쪽마다 이미지로 담음(한글 글꼴 파일 없이)
    await rptLoad("https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js");
    await rptLoad("https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js");
    const { jsPDF } = window.jspdf, doc = new jsPDF({ orientation: "landscape", unit: "pt", format: "a4", compress: true });
    const host = document.createElement("div"); host.style.cssText = "position:fixed;left:-20000px;top:0;"; document.body.appendChild(host);
    const pageHtml = (pg, i, n) => {
      if (pg.g.type === "valueup") { // v570: 밸류업DB 인쇄 양식(시안 ③)
        const dc = [...(rep.multiFc ? ["FC"] : []), ...RPT_DBP], box = '<span class="dbx"></span>';
        return `<div class="rpt-pdf rpt-dbp"><div class="rpt-ph"><b>${esc(pg.g.title || pg.g.label)}</b><span>${esc(rep.user)} · ${pg.g.recs.length}명 · ${esc(rep.stamp)}</span></div><table><thead><tr>${dc.map((c) => `<th class="c-${esc(c.replace("/", ""))}">${c}</th>`).join("")}</tr></thead><tbody>${pg.rows.map((r, k) => `<tr>${dc.map((c) => `<td class="c-${esc(c.replace("/", ""))}">${c === "No" ? pg.from + k + 1 : c === "통화" || c === "미팅" ? box : c === "FC" ? esc(r.FC) : esc(String(r[c] == null ? "" : r[c]))}</td>`).join("")}</tr>`).join("")}</tbody></table><div class="rpt-pf">성별/나이 = 주민번호 앞 7자리로 계산 · 유입일 = 회사 유입 날짜 · 유입메모 = SOL DB 메모 · ${i + 1} / ${n}</div></div>`;
      }
      const th = rptView(rep, pg.g.type).filter((c) => c !== "구분");
      return `<div class="rpt-pdf"><div class="rpt-ph"><b>${esc(pg.g.title || pg.g.label)}</b><span>${esc(rep.user)} · ${pg.g.recs.length}건 · ${esc(rep.stamp)}</span></div><table><thead><tr><th>#</th>${th.map((c) => `<th>${c}</th>`).join("")}</tr></thead><tbody>${pg.rows.map((r, k) => `<tr><td>${pg.from + k + 1}</td>${th.map((c) => `<td class="${c === "보험료" ? "n" : ""}">${rptCell(c, r[c])}</td>`).join("")}</tr>`).join("")}</tbody></table><div class="rpt-pf">${esc(rep.stamp)} · ${i + 1} / ${n}</div></div>`;
    };
    // v575(2026-11-25 사용자 지시 "아래쪽 빈 공간 없게"): 줄 높이를 실제로 재서 한 쪽에 들어가는 만큼 → 쪽 수를 정하고 쪽마다 고르게 나눔 → 남는 높이는 줄 높이를 늘려 채움
    const pages = [];
    for (const g of rep.groups) {
      if (!g.recs.length) { pages.push({ g, rows: [], from: 0, hs: [], avail: 0 }); continue; }
      host.innerHTML = pageHtml({ g, rows: g.recs, from: 0 }, 0, 1);
      const box = host.firstChild.getBoundingClientRect(), tb = host.querySelector("tbody").getBoundingClientRect();
      const avail = box.top + 794 - 46 - tb.top, hs = [...host.querySelectorAll("tbody tr")].map((tr) => tr.getBoundingClientRect().height);
      const chunk = (cap) => { const out = []; let i = 0; while (i < hs.length) { let sum = 0, c = 0; while (i + c < hs.length && c < cap && (c === 0 || sum + hs[i + c] <= avail)) { sum += hs[i + c]; c++; } out.push([i, c]); i += c; } return out; };
      const n = chunk(1e9).length, per = Math.ceil(hs.length / n);
      chunk(per).forEach(([i, c]) => pages.push({ g, rows: g.recs.slice(i, i + c), from: i, hs: hs.slice(i, i + c), avail }));
    }
    if (!pages.length) pages.push({ g: { label: "", recs: [] }, rows: [], from: 0, hs: [], avail: 0 });
    try {
      for (let i = 0; i < pages.length; i++) {
        const pg = pages[i];
        host.innerHTML = pageHtml(pg, i, pages.length);
        if (pg.hs.length) { // 남는 높이를 줄마다 나눠 줄 높이를 늘림(한 줄이 너무 커지지 않게 최대치)
          const used = pg.hs.reduce((a, b) => a + b, 0), add = Math.max(0, (pg.avail - used) / pg.hs.length), max = pg.g.type === "valueup" ? 34 : 28; // v577(2026-11-25 사용자 지시 "Row 크기 30% 줄여"): 줄을 늘리는 최대치도 낮춤
          [...host.querySelectorAll("tbody tr")].forEach((tr, k) => { const h = Math.min(max, pg.hs[k] + add); if (h > pg.hs[k]) tr.style.height = h + "px"; });
        }
        const cv = await window.html2canvas(host.firstChild, { scale: 1.6, backgroundColor: "#ffffff", logging: false });
        if (i) doc.addPage();
        doc.addImage(cv.toDataURL("image/jpeg", 0.88), "JPEG", 0, 0, 841.89, 595.28);
      }
    } finally { host.remove(); }
    return doc.output("blob");
  }
  function rptFormHtml() {
    const yNow = new Date().getFullYear(), ys = []; for (let y = yNow - 3; y <= yNow + 1; y++) ys.push(y);
    return `<div class="card rpt"><h3 class="sec-title">조건 보고서</h3>
      <div class="rpt-row"><span class="rpt-k">사용자</span><select id="rptUser" aria-label="사용자">${RPT.opts.map((o) => `<option value="${esc(o.id)}"${o.id === RPT.opt ? " selected" : ""}>${esc(o.label)}</option>`).join("")}</select></div>
      <div class="rpt-row"><span class="rpt-k">유형</span><span class="rpt-chk">${RPT_T.map(([k, l]) => `<label><input type="checkbox" data-rt="${k}"${RPT.types[k] ? " checked" : ""}> ${l}</label>`).join("")}</span></div>
      <div class="rpt-row"><span class="rpt-k">단계</span><span class="rpt-chk">${RPT_S.map(([k, l]) => `<label><input type="checkbox" data-rs="${k}"${RPT.st[k] ? " checked" : ""}> ${l}</label>`).join("")}</span></div>
      <div class="rpt-row"><span class="rpt-k">기간</span><span class="rpt-ym"><select id="rptY" aria-label="년">${ys.map((y) => `<option value="${y}"${y === RPT.y ? " selected" : ""}>${y}년</option>`).join("")}</select><select id="rptM" aria-label="월">${Array.from({ length: 12 }, (_, i) => `<option value="${i + 1}"${i + 1 === RPT.m ? " selected" : ""}>${i + 1}월</option>`).join("")}</select><button type="button" class="ev-edit rpt-span${RPT.span === 3 ? " on" : ""}" data-rpt="m3" aria-pressed="${RPT.span === 3}">3개월</button><button type="button" class="ev-edit rpt-span${RPT.span === 12 ? " on" : ""}" data-rpt="y1" aria-pressed="${RPT.span === 12}">1년</button></span></div>
      <div class="rpt-go"><span id="rptCount" class="rpt-count"></span><button type="button" class="ev-edit ev-done" id="rptGo" data-rpt="go">진행</button></div>
</div><div id="rptOut"></div>`;
  }
  function rptCounts() {
    const el = $("rptCount"); if (!el) return;
    const sel = rptSelected(), parts = RPT_T.filter(([k]) => RPT.types[k]).map(([k, l]) => ({ l, n: sel[k].length }));
    const total = parts.reduce((s, p) => s + p.n, 0);
    el.innerHTML = !parts.length ? "유형을 하나 이상 골라 주세요" : !RPT_S.some(([k]) => RPT.st[k]) ? "단계를 하나 이상 골라 주세요" : parts.map((p) => `${p.l} <b>${p.n}</b>건`).join(" · ") + ` · 합계 <b>${total}</b>건`;
    $("rptGo").disabled = !total;
  }
  // ===== v529(2026-11-21 사용자 지시) 세일즈 > 보고서 > 연체/실효 리스트 =====
  //   데이터(데이터수집이 확인일마다 씀): 각 FC 폴더 _보고서\연체실효\ — 목록.json · YYYY-MM-DD.json(연체) · 실효-YYYY-MM.json(월별 실효) · PDF
  //   그 FC 몫만 그 FC 폴더에 있어 드라이브에서 직접 열어도 남의 것이 안 보임. 보기: 각 FC는 본인 것만, ESL(관리자·직급 ESL)은 팀 전체
  //   연체 = 가장 최근 정식 확인일 리스트("11월 15일자 연체리스트") + PDF, 실효 = 지난달(최근) 실효리스트 PDF. 정식 기록 전에는 시험 기록을 '시험'으로 표시
  // v531(2026-11-21 사용자 규칙 "새 페이지는 디자인을 먼저 보여 주고 코딩"): 디자인 확인 전까지 숨김. 이 PC 시험에서만 주소에 ?od=1 을 붙여 켬
  const OD_ON = true; // v542(2026-11-21 사용자 "이대로 좋음"): 디자인 확인 끝 → 보임
  const OD = { opt: "", dirs: new Map(), cache: new Map(), data: null, opts: [], built: null };
  const OD_COLS = ["신규", "모집", "회차", "계약자", "주피", "연체", "상품명", "보험료", "최종납입", "연락처", "증권번호"]; // v553(2026-11-22 사용자 지시): 회차를 모집 오른쪽으로 // v547(2026-11-21 사용자 지시): 신규 옆 모집(직접·이관)
  const OD_MINN = 2, OD_MAXN = 24; // v537(2026-11-21 사용자 지시) 연체리스트: 최종납입회차 2~24회만(회차 모름은 넣음)
  const odInRange = (r) => { const v = r["납입회차"]; if (v == null || v === "") return true; const n = +v; return !(n < OD_MINN || n > OD_MAXN); };
  const OL_COLS = ["모집", "회차", "계약자", "주피", "상품명", "보험료", "연락처", "증권번호"]; // v553: 회차를 모집 오른쪽으로 // v547: 계약자 왼쪽에 모집
  function odFcDir(fc) { // FC 폴더 ID: 읽어 둔 그 FC 고객 줄의 data.json 부모 폴더 → 없으면 권한 파일의 FC 폴더 ID
    const e = (OD.data ? OD.data.lists : []).find((x) => x.fc === fc && x.src && x.src.fileId && jsonInfo.get(x.src.fileId));
    if (e) return jsonInfo.get(e.src.fileId).parent;
    return (fcFolderIds || {})[fc] || "";
  }
  async function odFolder(fc) { // { files: {이름: id}, list: 목록.json 내용 } (없으면 null)
    if (OD.dirs.has(fc)) return OD.dirs.get(fc);
    const p = (async () => {
      const dir = odFcDir(fc); if (!dir) return null;
      const a = await gdFindIn(dir, "_보고서"); if (!a) return null;
      const b = await gdFindIn(a.id, "연체실효"); if (!b) return null;
      const files = {}; (await gdQuery("'" + qv(b.id) + "' in parents and trashed=false")).forEach((f) => { if (f.file) files[f.name] = f.id; });
      const list = files["목록.json"] ? await odJson(files["목록.json"]).catch(() => null) : null;
      return { files, list };
    })().catch((e) => { console.warn("연체 폴더 못 읽음", fc, e); return null; });
    OD.dirs.set(fc, p); return p;
  }
  async function odJson(id) {
    if (!OD.cache.has(id)) OD.cache.set(id, stGet0(id).then((b) => parseDoc(b)).catch((e) => { OD.cache.delete(id); throw e; }));
    return OD.cache.get(id);
  }
  const odRowsOf = (d) => (Array.isArray(d) ? d : (d && d.rows) || []);
  function odPick(list) { // 연체 항목(정식 최근, 없으면 가장 최근 시험)과 그 바로 전 연체 항목, 실효 항목
    const items = (list && list.목록) || [], od = items.filter((x) => x.종류 === "연체").sort((a, b) => String(a.확인일).localeCompare(String(b.확인일)));
    const cur = (list && list.최근연체 && od.find((x) => x.파일 === list.최근연체)) || od[od.length - 1] || null;
    const prev = cur ? od.filter((x) => x.확인일 < cur.확인일 && x.정식 !== false).pop() || null : null;
    const lp = items.filter((x) => x.종류 === "실효").sort((a, b) => String(a.월).localeCompare(String(b.월)));
    const lapse = (list && list.최근실효 && lp.find((x) => x.파일 === list.최근실효)) || lp[lp.length - 1] || null;
    return { cur, prev, lapse };
  }
  const odTitle = (it) => { if (!it) return "연체리스트"; if (it.정식 === false || /시험/.test(String(it.제목 || ""))) { const m = /^(\d{4})-(\d\d)/.exec(String(it.확인일 || "")); return m ? `${m[1]}년 ${+m[2]}월 연체리스트` : "연체리스트"; } return String(it.제목 || "연체리스트").replace(/^\d{4}-\d\d-\d\d /, ""); };
  const odFcsOf = () => { const o = OD.opts.find((x) => x.id === OD.opt); return o ? o.fcs : []; };
  const odKey = (r) => String(r["계약ID"] || r["증권번호"] || "");
  function odJoin(fc, r) { // 줄에 이름·상품·보험료·연락처가 없으면 data.json에서 붙임
    const data = OD.data, ent = (cid) => { const es = data.custById.get(String(cid || "")) || []; return es.find((x) => x.fc === fc) || es[0] || null; };
    const kid = String(r["계약ID"] || ""), pol = String(r["증권번호"] || "");
    const ct = data.contracts.find((c) => (kid && String(c.row["계약ID"] || "") === kid) || (pol && String(c.row["증권번호"] || "") === pol));
    const own = ent(r["고객ID"]), ins = ent(r["주피보험자ID"]);
    const num = (v) => +(String(v == null ? "" : v).replace(/[^\d]/g, "")) || "";
    return {
      계약자: r["계약자"] || (ct && ct.row["계약자"]) || (own && own.name) || r["고객ID"] || "", 주피: r["피보험자"] || (ct && (ct.row["주피보험자"] || ct.row["피보"])) || (ins && ins.name) || "",
      상품명: productName(r["상품명"] || (ct && ct.row["상품명"]) || ""), 보험료: num(r["보험료"]) || (ct ? num(ct.row["합계보험료"] || ct.row["보험료"]) : ""),
      최종납입: String(r["최종납입년월"] || "").replace("-", "."), 연락처: r["연락처"] || (own ? own.phone || fieldOf(own, "연락처") : ""), 증권번호: pol,
      _fc: fc, _cust: String(r["고객ID"] || ""), _cid: kid,
      모집: (() => { const t = String((ct && (ct.row["모집유형"] || ct.row["계약분류"])) || r["모집유형"] || r["계약분류"] || ""); return /타인|이관/.test(t) ? "이관" : /직접/.test(t) ? "직접" : ""; })()
    };
  }
  async function odBuild() { // 고른 범위의 FC별 연체 줄·실효 줄
    const groups = await Promise.all(odFcsOf().map(async (fc) => { // FC마다 동시에
      const fo = await odFolder(fc), pk = odPick(fo && fo.list), g = { fc, cur: pk.cur, lapse: pk.lapse, rows: [], lrows: [], files: (fo && fo.files) || {} };
      const nByKey = new Map();
      if (pk.cur && g.files[pk.cur.파일]) {
        const rows = odRowsOf(await odJson(g.files[pk.cur.파일]));
        rows.forEach((r) => { if (r["납입회차"] != null && r["납입회차"] !== "") nByKey.set(odKey(r), r["납입회차"]); });
        const rows2 = rows, prevIds = pk.prev && g.files[pk.prev.파일] ? new Set(odRowsOf(await odJson(g.files[pk.prev.파일])).map(odKey)) : null;
        g.excl = rows2.filter((r) => String(r["상태"] || "연체") === "연체" && !odInRange(r)).length;
        g.rows = rows2.filter((r) => String(r["상태"] || "연체") === "연체" && odInRange(r)).map((r) => { const st = String(r["상태"] || ""), isNew = !!(prevIds && !prevIds.has(odKey(r))), j = odJoin(fc, r); // v535: 연체리스트는 최종납입 24회 이내만(데이터는 그대로, 화면·PDF에서만 거름)
          return { ...j, 상태: st, 신규: isNew ? "신규" : "", _new: isNew, 회차: r["납입회차"] != null && r["납입회차"] !== "" ? r["납입회차"] : "(회차 모름)",
            연체: r["연체월수"] != null ? r["연체월수"] + "개월" : "" };
        }).sort((a, b) => { const na = typeof a.회차 === "number" ? a.회차 : +a.회차 || 1e9, nb = typeof b.회차 === "number" ? b.회차 : +b.회차 || 1e9; return na - nb || String(a.계약자).localeCompare(String(b.계약자), "ko"); }); // v547: 낮은 회차가 위
      }
      if (pk.lapse && g.files[pk.lapse.파일]) g.lrows = odRowsOf(await odJson(g.files[pk.lapse.파일])).map((r) => ({ ...odJoin(fc, r), 실효일: String(r["실효일"] || "").replace(/-/g, "."), 지금상태: r["지금상태"] || "", 회차: r["납입회차"] != null && r["납입회차"] !== "" ? r["납입회차"] : (nByKey.get(odKey(r)) != null ? nByKey.get(odKey(r)) : ""), _lp: true })).sort((a, b) => String(a.실효일).localeCompare(String(b.실효일)) || String(a.계약자).localeCompare(String(b.계약자), "ko"));
      return g;
    }));
    return groups;
  }
  const odCell = (c, v) => (c === "보험료" && typeof v === "number" ? v.toLocaleString("ko-KR") : esc(String(v == null ? "" : v)));
  const odSum = (rows, lrows) => `연체 <b>${rows.length}</b>${lrows ? ` · 실효 <b>${lrows.length}</b>` : ""} · 신규 <b>${rows.filter((r) => r._new).length}</b>`; // v533: 실효 수 = 아래 실효리스트 건수
  // v540(2026-11-21 사용자 지시): 연락처 번호를 누르면 전화(세일즈 화면 전화 버튼과 같게: 한 번 더 묻고, 그 FC 이력에 '통화·시도' 기록)
  const odPhone = (r) => { const t = firstPhone(r.연락처); return t ? `<a href="tel:${esc(t)}" class="od-tel" data-odtel="${esc(t)}" data-fc="${esc(r._fc || "")}" data-cust="${esc(r._cust || "")}" data-cid="${esc(r._cid || "")}" data-nm="${esc(r.계약자 || "")}" data-pg="${r._lp ? "실효리스트" : "연체리스트"}">${esc(r.연락처)}</a>` : ""; };
  const odTable = (cols, rows, cls) => `<div class="rpt-tw"><table class="rpt-t od-t"><thead><tr><th>#</th>${cols.map((c) => `<th>${c}</th>`).join("")}</tr></thead><tbody>${rows.map((r, i) => `<tr class="${cls ? cls(r) : ""}"><td>${i + 1}</td>${cols.map((c) => `<td class="${c === "보험료" ? "n" : ""}">${c === "연락처" ? odPhone(r) : odCell(c, r[c])}</td>`).join("")}</tr>`).join("")}</tbody></table></div>`;
  async function odRender() {
    const box = $("odCard"); if (!box) return;
    const sel = OD.opts.length > 1 ? `<select id="odUser" aria-label="FC">${OD.opts.map((o) => `<option value="${esc(o.id)}"${o.id === OD.opt ? " selected" : ""}>${esc(o.label)}</option>`).join("")}</select>` : "";
    const head = `<div class="od-h"><h3 class="sec-title">연체/실효 리스트</h3>${sel}</div>`;
    box.innerHTML = head + `<p class="hint">불러오는 중…</p>`;
    const groups = await odBuild(), multi = odFcsOf().length > 1;
    const cur = (groups.find((g) => g.cur && g.cur.정식 !== false) || groups.find((g) => g.cur) || {}).cur || null;
    const lapse = (groups.find((g) => g.lapse) || {}).lapse || null;
    OD.built = { groups, multi, cur, lapse, label: (OD.opts.find((o) => o.id === OD.opt) || {}).label || "" };
    const all = groups.flatMap((g) => g.rows);
    const odHtml = !cur ? `<p class="hint">아직 연체리스트가 없습니다.</p>` :
      `<div class="od-th"><p class="od-title"><b>${esc(odTitle(cur))}</b></p>${all.length ? `<button type="button" class="ev-edit" data-od="od" aria-label="연체리스트 PDF 다운">PDF 다운</button>` : ""}</div>` + // v564(사용자 지시): 연체가 없으면 PDF 다운 버튼 숨김 // v549: PDF 다운은 제목 줄 오른쪽 끝 // v546(2026-11-21 사용자 지시): '입니다' 빼고, 시험 기록은 'YYYY년 M월 연체리스트'

      `<div class="od-sum">${odSum(all, groups.flatMap((g) => g.lrows))}${multi ? ` <small>· FC ${groups.length}명</small>` : ""}</div>` +
      (multi && !groups.some((g) => g.rows.length) ? `<p class="hint pad">${OD_MINN}~${OD_MAXN}회 연체가 없습니다.</p>` : "") +
      groups.filter((g) => !multi || g.rows.length).map((g) => `<div class="od-grp">${multi ? `<div class="rpt-gh"><b>${esc(g.fc)}</b> <span>${g.cur ? odSum(g.rows, g.lrows) : "기록 없음"}</span></div>` : ""}${g.rows.length ? odTable(OD_COLS, g.rows, (r) => (r._new ? "od-new" : "")) : g.cur ? `<p class="hint pad">${OD_MINN}~${OD_MAXN}회 연체가 없습니다.</p>` : ""}</div>`).join("") +
      "";
    const nL = groups.reduce((a, g) => a + g.lrows.length, 0);
    // v532(2026-11-21 사용자 지시): 실효리스트도 표로 보여 주고 그 밑에 PDF 다운
    const lpHtml = lapse ? `<div class="od-lp"><div class="od-th"><p class="od-title"><b>${esc(lapse.제목 || "실효리스트")}</b></p>${nL ? `<button type="button" class="ev-edit" data-od="lapse" aria-label="${esc(lapse.제목 || "실효리스트")} PDF 다운">PDF 다운</button>` : ""}</div>` + // v564: 실효가 없으면 숨김
      (nL ? groups.filter((g) => g.lapse && (!multi || g.lrows.length)).map((g) => `<div class="od-grp">${multi ? `<div class="rpt-gh"><b>${esc(g.fc)}</b> <span>${g.lrows.length}건</span></div>` : ""}${g.lrows.length ? odTable(OL_COLS, g.lrows, (r) => (r.지금상태 && r.지금상태 !== "실효" ? "od-revived" : "")) : multi ? `<p class="hint pad">실효 계약이 없습니다.</p>` : ""}</div>`).join("") : `<p class="hint">실효 계약이 없습니다.</p>`) +
      `</div>` : "";
    box.innerHTML = head + odHtml + lpHtml + `<p class="hint" id="odMsg"></p>`; // v538(2026-11-21 사용자 지시): 회색 안내 글은 모두 뺌(메시지 자리만 남김)
  }
  // v548(2026-11-21 사용자 지시): '전체' PDF = 앞부분 ESL용 묶음(모든 FC를 이어서, FC마다 제목 줄) + 뒷부분 FC별 나눠 주기용(FC마다 새 쪽)
  async function odPdfBlob(title, sub, groups, cols, rowsOf, combined) { // 쪽마다 18줄(가로 A4). 화면 글꼴 그대로 이미지로
    await rptLoad("https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js");
    await rptLoad("https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js");
    const { jsPDF } = window.jspdf, doc = new jsPDF({ orientation: "landscape", unit: "pt", format: "a4", compress: true });
    const PER = 18, pages = [];
    if (combined && groups.length > 1) { // 묶음: FC 제목 줄도 한 줄로 세어 이어서 나눔
      const items = []; groups.forEach((g) => { const rs = rowsOf(g); items.push({ hdr: g.fc, txt: sub(g, rs) }); rs.forEach((r, k) => items.push({ r, no: k + 1 })); });
      const all = groups.flatMap((g) => rowsOf(g)), n = Math.max(1, Math.ceil(items.length / PER));
      for (let p = 0; p < n; p++) pages.push({ comb: true, items: items.slice(p * PER, p * PER + PER), p: p + 1, n, total: all.length });
    }
    groups.forEach((g) => { const rs = rowsOf(g), n = Math.max(1, Math.ceil(rs.length / PER)); for (let p = 0; p < n; p++) pages.push({ g, rs, rows: rs.slice(p * PER, p * PER + PER), from: p * PER, p: p + 1, n }); });
    const nComb = pages.filter((x) => x.comb).length, tOf = (g) => (typeof title === "function" ? title(g) : title); // v548: FC별 쪽은 그 FC 기록의 제목
    const cell = (r) => cols.map((c) => `<td class="${c === "보험료" ? "n" : ""}">${odCell(c, r[c])}</td>`).join("");
    const host = document.createElement("div"); host.style.cssText = "position:fixed;left:-20000px;top:0;"; document.body.appendChild(host);
    try {
      for (let i = 0; i < pages.length; i++) {
        const pg = pages[i], head = `<thead><tr><th>#</th>${cols.map((c) => `<th>${c}</th>`).join("")}</tr></thead>`;
        host.innerHTML = pg.comb
          ? `<div class="rpt-pdf"><div class="rpt-ph"><b>${esc(tOf(null))} · 전체(FC별 묶음)</b><span>FC ${groups.length}명 · ${pg.total}건</span></div><table>${head}<tbody>${pg.items.map((it) => it.hdr ? `<tr class="od-ph"><td colspan="${cols.length + 1}"><b>${esc(it.hdr)}</b> <span>${esc(it.txt)}</span></td></tr>` : `<tr><td>${it.no}</td>${cell(it.r)}</tr>`).join("")}</tbody></table><div class="rpt-pf">묶음 ${pg.p} / ${pg.n} · 전체 ${i + 1} / ${pages.length}</div></div>`
          : `<div class="rpt-pdf"><div class="rpt-ph"><b>${esc(tOf(pg.g))} · ${esc(pg.g.fc)}</b><span>${esc(sub(pg.g, pg.rs))}</span></div><table>${head}<tbody>${pg.rows.length ? pg.rows.map((r, k) => `<tr><td>${pg.from + k + 1}</td>${cell(r)}</tr>`).join("") : `<tr><td colspan="${cols.length + 1}">해당 없음</td></tr>`}</tbody></table><div class="rpt-pf">${esc(pg.g.fc)} ${pg.p} / ${pg.n} · 전체 ${i + 1} / ${pages.length}${nComb ? " (FC별 나눠 주기용)" : ""}</div></div>`;
        const cv = await window.html2canvas(host.firstChild, { scale: 1.6, backgroundColor: "#ffffff", logging: false });
        if (i) doc.addPage();
        doc.addImage(cv.toDataURL("image/jpeg", 0.88), "JPEG", 0, 0, 841.89, 595.28);
      }
    } finally { host.remove(); }
    return doc.output("blob");
  }
  async function odInit(data) {
    OD.data = data; OD.dirs = new Map(); OD.cache = new Map();
    const myFc = await myFcName(), ord = CFG.fcOrder || [];
    const pm = salesPermsAll || (await readPerms().catch(() => null)), esl = isAdmin() || (!!pm && titleOf(pm, me()) === "ESL");
    const fcs = [...new Set([...data.lists.map((e) => e.fc), ...knownFcs].filter(Boolean))].sort((a, b) => ((ord.indexOf(a) + 1 || 999) - (ord.indexOf(b) + 1 || 999)) || a.localeCompare(b, "ko"));
    OD.opts = esl ? [{ id: "__all", label: "전체", fcs }, ...fcs.map((f) => ({ id: f, label: f, fcs: [f] }))] : (fcs.includes(myFc) ? [{ id: myFc, label: myFc, fcs: [myFc] }] : []);
    if (!OD.opts.some((o) => o.id === OD.opt)) OD.opt = (OD.opts[0] || {}).id || "";
    if (!OD.opts.length) { $("odCard").innerHTML = `<h3 class="sec-title">연체/실효 리스트</h3><p class="hint">볼 수 있는 FC가 없습니다.</p>`; return; }
    try { await odRender(); } catch (e) { $("odCard").innerHTML = `<h3 class="sec-title">연체/실효 리스트</h3><p class="hint">불러오지 못했습니다(${esc(String(e.message || e).slice(0, 60))}).</p>`; }
  }
  $("salesView").addEventListener("click", (ev) => {
    const a = ev.target.closest && ev.target.closest("[data-odtel]"); if (!a) return;
    if (!confirm(`${a.dataset.nm ? a.dataset.nm + "님께 " : ""}정말로 전화를 거시겠습니까?`)) { ev.preventDefault(); return; }
    logAttempt(a.dataset.pg || "연체리스트", "call", a.dataset.nm, smemoFc(a.dataset.cust, a.dataset.fc)); // v541: 상담기록에 "연체리스트에서 전화 시도"
    const fc = a.dataset.fc, e = (OD.data ? OD.data.lists : []).find((x) => x.fc === fc && x.src && x.src.fileId && jsonInfo.get(x.src.fileId));
    if (!fc || !e || !a.dataset.cust || isRO(fc)) return; // 기록할 곳이 없거나 보기 전용이면 전화만
    const q = readQueue();
    q.push({ uid: Date.now() + "-" + Math.random().toString(36).slice(2, 7), kind: "call", ts: nowStamp(), cid: a.dataset.cid || "", custId: a.dataset.cust, fc, fileId: e.src.fileId, src: "세일즈 보고서(연체/실효)" });
    writeQueue(q); setTimeout(flushContacts, 4000);
  });
  $("salesView").addEventListener("change", (ev) => { const t = ev.target; if (!t || t.id !== "odUser") return; OD.opt = t.value; odRender().catch(handleError); });
  $("salesView").addEventListener("click", async (ev) => {
    const btn = ev.target.closest && ev.target.closest("[data-od]"); if (!btn || !OD.built) return;
    const b = OD.built, msg = $("odMsg"), say = (t) => { if (msg) msg.textContent = t; }, kind = btn.dataset.od, old = btn.textContent;
    btn.disabled = true; btn.textContent = "만드는 중…";
    try {
      const it = kind === "od" ? b.cur : b.lapse, g0 = b.groups[0];
      const title = kind === "od" ? odTitle(it) : String((it && it.제목) || "실효리스트");
      const fileName = (kind === "od" ? `${String(it && it.확인일 || "").replace(/-/g, "")} - ${title}` : title) + ` - ${b.label}.pdf`;
      // 한 FC면 데이터수집이 만든 PDF를 그대로 받음(있을 때), 전체거나 없으면 앱에서 FC별로 나눠 만듦
      const pc = !b.multi && g0 && (kind === "od" ? g0.cur : g0.lapse) && g0.files[(kind === "od" ? g0.cur : g0.lapse).pdf];
      const blob = pc ? new Blob([await stGet0(pc)], { type: "application/pdf" })
        : kind === "od" ? await odPdfBlob((g) => { if (g) return odTitle(g.cur); const ds = new Set(b.groups.filter((x) => x.cur && x.rows.length).map((x) => x.cur.확인일)); return ds.size === 1 ? odTitle(b.groups.find((x) => x.cur && x.rows.length).cur) : "연체리스트"; }, (g, rs) => `확인일 ${g.cur ? g.cur.확인일 : "-"} · ${odSum(rs).replace(/<\/?b>/g, "")}`, b.groups.filter((g) => g.cur && (!b.multi || g.rows.length)), OD_COLS, (g) => g.rows, b.multi)
          : await odPdfBlob(title, (g, rs) => `${rs.length}건 · 실효일이 그 달인 계약 · 지금상태는 확인한 날 기준`, b.groups.filter((g) => g.lapse && (!b.multi || g.lrows.length)), OL_COLS, (g) => g.lrows, b.multi);
      rptSave(blob, fileName.replace(/[\\/:*?"<>|]/g, "")); say("PDF를 내려받았어요.");
    } catch (e) { say("실패: " + (e && e.message ? e.message : e)); }
    finally { btn.disabled = false; btn.textContent = old; }
  });
  async function renderMaterial(req) { // 세일즈 > 보고서
    const box = $("salesView");
    box.innerHTML = `<div class="card"><h3 class="sec-title">조건 보고서</h3><p class="hint">불러오는 중…</p></div>`;
    try {
      const data = await loadSalesData(); if (req !== salesReq) return;
      await loadMyPrefs().catch(() => {}); await refreshSalesExtra();
      const opts = await rptOptions(data); if (req !== salesReq) return;
      RPT.data = data; RPT.opts = opts; RPT.cache = null; RPT.built = null;
      if (!RPT.y) { const n = new Date(); RPT.y = n.getFullYear(); RPT.m = n.getMonth() + 1; }
      if (!opts.some((o) => o.id === RPT.opt)) RPT.opt = (opts.find((o) => o.id !== "__all") || opts[0] || {}).id || ""; // 처음 고르는 값은 지금처럼 본인(전체는 맨 위에 있지만 기본은 아님)
      box.innerHTML = rptFormHtml(); rptCounts();
      if (OD_ON) { box.insertAdjacentHTML("beforeend", `<div class="card od" id="odCard"><h3 class="sec-title">연체/실효 리스트</h3><p class="hint">불러오는 중…</p></div>`); odInit(data); } // v529
    } catch (e) { if (req === salesReq) { box.innerHTML = ""; handleError(e); } }
  }
  function rptOutHtml(rep) {
    let share = false; try { share = !!(navigator.canShare && navigator.canShare({ files: [new File(["x"], "a.pdf", { type: "application/pdf" })] })); } catch (e) { share = false; }
    return `<div class="card rpt-out"><div class="rpt-oh"><b>${esc(rep.title)}</b><span class="hint">${esc(rep.cond)} · ${rep.total}건 · ${esc(rep.stamp)}</span></div>${rptTableHtml(rep)}
      <div class="rpt-acts"><button type="button" class="ev-edit" data-rpt="xlsx">엑셀 다운</button><button type="button" class="ev-edit" data-rpt="pdf">PDF 다운</button><button type="button" class="ev-edit" data-rpt="pdfopen">PDF 열람</button><button type="button" class="ev-edit" data-rpt="mail">이메일(나에게)</button>${share ? `<button type="button" class="ev-edit" data-rpt="share">파일 공유</button>` : ""}</div>
      <p class="hint" id="rptMsg"></p></div>`;
  }
  async function rptAct(kind, btn) {
    const rep = RPT.built; if (!rep) return;
    const msg = $("rptMsg"), say = (t) => { if (msg) msg.textContent = t; };
    if (kind === "mail") {
      const to = me(), lines = rep.groups.flatMap((g) => g.recs.map((r) => `[${g.label}] ${r.이름}${r.상품 ? " · " + r.상품 : ""}${r.연락처 ? " · " + r.연락처 : ""}`));
      let body = `${rep.title}\n${rep.cond}\n${rep.groups.map((g) => g.label + " " + g.recs.length + "건").join(" / ")} (합계 ${rep.total}건)\n\n`;
      for (const ln of lines) { if ((body + ln).length > 1700) { body += "…(나머지는 앱에서 엑셀·PDF로 받으세요)"; break; } body += ln + "\n"; }
      location.href = "mailto:" + encodeURIComponent(to) + "?subject=" + encodeURIComponent(`${rep.title} ${rep.user} ${rep.ym}`) + "&body=" + encodeURIComponent(body);
      say("메일 앱이 열립니다. 파일 첨부는 안 되어 본문에 명단만 담았어요. 파일은 다운 후 직접 첨부하세요.");
      return;
    }
    const old = btn && btn.textContent; if (btn) { btn.disabled = true; btn.textContent = "만드는 중…"; }
    let w = null; if (kind === "pdfopen") { try { w = window.open("", "_blank"); } catch (e) { w = null; } }
    try {
      if (kind === "xlsx") { rptSave(rptXlsxBlob(rep), rptFileName(rep, "xlsx")); say("엑셀을 내려받았어요."); }
      else {
        const blob = await rptPdfBlob(rep), name = rptFileName(rep, "pdf");
        if (kind === "pdf") { rptSave(blob, name); say("PDF를 내려받았어요."); }
        else if (kind === "pdfopen") { const u = URL.createObjectURL(blob); if (w) { w.location.href = u; say("PDF를 새 창으로 열었어요."); } else { rptSave(blob, name); say("새 창이 막혀 PDF를 내려받았어요."); } }
        else if (kind === "share") { await navigator.share({ files: [new File([blob], name, { type: "application/pdf" })], title: rep.title }); say("공유 창을 열었어요."); }
      }
    } catch (e) { if (w) { try { w.close(); } catch (e2) { /* 무시 */ } } if (!(e && e.name === "AbortError")) say("실패: " + (e && e.message ? e.message : e)); }
    finally { if (btn) { btn.disabled = false; btn.textContent = old; } }
  }
  $("salesView").addEventListener("change", (ev) => {
    const t = ev.target; if (!t || !$("rptCount")) return;
    if (t.id === "rptUser") RPT.opt = t.value; else if (t.id === "rptY" || t.id === "rptM") { RPT[t.id === "rptY" ? "y" : "m"] = +t.value; RPT.span = 1; document.querySelectorAll("#salesView .rpt-span").forEach((b) => { b.classList.remove("on"); b.setAttribute("aria-pressed", "false"); }); }
    else if (t.dataset && t.dataset.rs) RPT.st[t.dataset.rs] = t.checked; else if (t.dataset && t.dataset.rt) RPT.types[t.dataset.rt] = t.checked; else return;
    RPT.built = null; $("rptOut").innerHTML = ""; rptCounts();
  });
  $("salesView").addEventListener("click", (ev) => {
    const b = ev.target.closest && ev.target.closest("[data-rpt]"); if (!b) return;
    const k = b.dataset.rpt;
    if (k === "m3" || k === "y1") { // 이번 달까지 3개월·1년 → 바로 보고서
      const n = new Date(); RPT.y = n.getFullYear(); RPT.m = n.getMonth() + 1; RPT.span = k === "m3" ? 3 : 12;
      $("rptY").value = RPT.y; $("rptM").value = RPT.m;
      document.querySelectorAll("#salesView .rpt-span").forEach((x) => { const on = x.dataset.rpt === k; x.classList.toggle("on", on); x.setAttribute("aria-pressed", String(on)); });
      rptCounts();
      if ($("rptGo").disabled) { RPT.built = null; $("rptOut").innerHTML = `<p class="hint pad">이 조건에 맞는 고객이 없습니다.</p>`; return; }
      RPT.built = rptBuild(); $("rptOut").innerHTML = rptOutHtml(RPT.built); $("rptOut").scrollIntoView({ behavior: "smooth", block: "start" });
    }
    else if (k === "go") { RPT.built = rptBuild(); $("rptOut").innerHTML = rptOutHtml(RPT.built); $("rptOut").scrollIntoView({ behavior: "smooth", block: "start" }); }
    else rptAct(k, b);
  });
  // ---------- v554(2026-11-22 사용자 지시 "세일즈 보고서 오른쪽에 자료 페이지 추가", "버튼 누르면 바로 실행되게") 세일즈 > 자료 ----------
  //   설정 폴더 › 자료 안의 폴더 하나 = 분류 하나(폴더 이름 앞 숫자는 순서용, 화면에선 뺌), 자료 폴더 바로 아래 파일 = '기타'
  //   줄을 누르면 바로 열기, [공유]는 바로 휴대폰 공유창(파일째), [올리기]는 관리자만(고른 분류 폴더에, '전체'면 자료 폴더 바로 아래)
  const MAT_DIR = "자료", MAT_ETC = "기타";
  const MAT = { cat: "전체", q: "", groups: null, dir: null, blob: {} };
  async function matLoad() {
    await resolveAppDir();
    if (!appDir || !appDir.id) return [];
    if (!MAT.dir) {
      const r = await gdQuery("'" + qv(appDir.id) + "' in parents and name='" + qv(MAT_DIR) + "' and mimeType='" + GD_FOLDER + "' and trashed=false");
      if (r.length) MAT.dir = { id: r[0].id };
      else if (isAdmin()) { const mk = await gdFetch(GDAPI + "/files?fields=id&supportsAllDrives=true", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ name: MAT_DIR, mimeType: GD_FOLDER, parents: [appDir.id] }) }); MAT.dir = { id: mk.id }; }
    }
    if (!MAT.dir) return [];
    const subs = (await gdQuery("'" + qv(MAT.dir.id) + "' in parents and mimeType='" + GD_FOLDER + "' and trashed=false")).sort((a, b) => a.name.localeCompare(b.name, "ko", { numeric: true }));
    const lists = (await Promise.all([MAT.dir, ...subs].map((d) => docsList({ id: d.id })))).map((l) => l.sort((a, b) => String(b.mod).slice(0, 10).localeCompare(String(a.mod).slice(0, 10)) || b.name.localeCompare(a.name, "ko", { numeric: true }))); // 최근 날짜 → 같은 날이면 이름 큰 것(11월 > 10월)부터
    const groups = subs.map((d, i) => ({ name: d.name, id: d.id, files: lists[i + 1] }));
    if (lists[0].length) groups.push({ name: MAT_ETC, id: MAT.dir.id, files: lists[0] });
    return groups;
  }
  const matCat = (s) => String(s).replace(/^\d+[.\s_-]+/, "") || String(s);
  const matNm = (f) => String(f.name).replace(/\.[A-Za-z0-9]{1,5}$/, "").replace(/^\d{8}\s*-\s*/, "");
  const matTy = (f) => { const e = docsExt(f); return e === "PDF" ? "PDF" : /^(DOCX?|HWPX?|TXT|RTF)$/.test(e) ? "DOC" : /^(XLSX?|CSV)$/.test(e) ? "XLS" : /^PPTX?$/.test(e) ? "PPT" : docsIsImg(f) ? "IMG" : /google-apps/.test(f.mime || "") ? "GDOC" : e.slice(0, 4); };
  const matSize = (n) => (n >= 1048576 ? (n / 1048576).toFixed(1) + "MB" : n >= 1024 ? Math.round(n / 1024) + "KB" : n ? n + "B" : "");
  async function renderMat(req, again) {
    const box = $("salesView");
    if (!MAT.groups || again) {
      if (!MAT.groups) box.innerHTML = `<div class="card"><p class="hint">자료를 불러오는 중…</p></div>`;
      try { const g = await matLoad(); if (req !== salesReq) return; MAT.groups = g; }
      catch (e) { if (req === salesReq) box.innerHTML = `<div class="card"><p class="hint">자료를 불러오지 못했습니다.</p></div>`; handleError(e); return; }
    }
    const adm = isAdmin() && GD;
    box.innerHTML = `<div class="mt-bar"><input type="search" id="matQ" placeholder="자료 검색" value="${esc(MAT.q)}" autocomplete="off">${adm ? `<button type="button" class="ev-edit" data-mat-up>올리기</button><input type="file" id="matFile" multiple hidden>` : ""}</div><div id="matList"></div>`;
    matDraw();
  }
  function matDraw() {
    const el = $("matList"); if (!el) return;
    const G = MAT.groups || [];
    if (MAT.cat !== "전체" && !G.some((g) => g.name === MAT.cat)) MAT.cat = "전체";
    const q = MAT.q.trim().toLowerCase(), hit = (f) => !q || f.name.toLowerCase().includes(q);
    const all = G.reduce((n, g) => n + g.files.length, 0);
    const cats = [["전체", all], ...G.map((g) => [g.name, g.files.length])].map(([c, n]) => `<button type="button" class="${c === MAT.cat ? "on" : ""}" data-mat-cat="${esc(c)}">${esc(c === "전체" ? c : matCat(c))}<small>${n}</small></button>`).join("");
    const shown = G.map((g, gi) => ({ g, gi, files: g.files.map((f, fi) => ({ f, fi })).filter((x) => hit(x.f)) })).filter((x) => (MAT.cat === "전체" || x.g.name === MAT.cat) && x.files.length);
    const body = shown.map(({ g, gi, files }) => `<div class="card mt-card"><div class="mt-gl"><b>${esc(matCat(g.name))}</b>${files.length}개</div>${files.map(({ f, fi }) => { const ty = matTy(f); return `<div class="mt-row"><button type="button" class="mt-open" data-mat-open="${gi}:${fi}"><span class="mt-ty${ty === "PDF" ? " pdf" : ""}">${esc(ty)}</span><span class="mt-nm"><b>${esc(matNm(f))}</b><span>${esc(String(f.mod).slice(0, 10))}${matSize(f.size) ? " · " + matSize(f.size) : ""}</span></span></button><button type="button" class="mt-sh" data-mat-share="${gi}:${fi}">${MAT.blob[f.id] ? "보내기" : "공유"}</button></div>`; }).join("")}</div>`).join("");
    el.innerHTML = `<div class="mt-cats">${cats}</div>${body || `<div class="card"><p class="hint">${all ? "찾는 자료가 없습니다." : `아직 자료가 없습니다.${isAdmin() ? " [올리기]로 파일을 올리면 설정 폴더 › 자료에 저장됩니다. 분류는 자료 폴더 안에 폴더를 만들면 생깁니다." : ""}`}</p></div>`}<p class="hint mt-tip">줄을 누르면 열립니다 · [공유]는 휴대폰 공유창(카톡·문자)으로 파일을 보냅니다</p>`;
  }
  const matFile = (k) => { const [gi, fi] = String(k).split(":").map(Number), g = (MAT.groups || [])[gi]; return g ? g.files[fi] : null; };
  async function matShare(f, btn) {
    const old = btn.textContent;
    try {
      let file = MAT.blob[f.id];
      if (!file) {
        if (/google-apps/.test(f.mime || "")) { if (navigator.share && f.webUrl) await navigator.share({ title: matNm(f), url: f.webUrl }); else if (f.webUrl) window.open(f.webUrl, "_blank", "noopener"); return; }
        btn.disabled = true; btn.textContent = "준비 중";
        file = new File([await stGet0(f.id)], f.name, { type: f.mime || "application/octet-stream" });
      }
      if (navigator.canShare && navigator.canShare({ files: [file] })) {
        try { await navigator.share({ files: [file], title: matNm(f) }); delete MAT.blob[f.id]; }
        catch (e) { if (e && e.name === "NotAllowedError") { MAT.blob[f.id] = file; btn.disabled = false; btn.textContent = "보내기"; toast("파일을 받았어요. [보내기]를 누르면 공유창이 열립니다"); return; } if (!(e && e.name === "AbortError")) throw e; }
      } else { const u = URL.createObjectURL(file), a = document.createElement("a"); a.href = u; a.download = f.name; document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(u), 60000); toast("이 기기는 공유창을 쓸 수 없어 내려받았습니다"); }
      btn.textContent = "공유";
    } catch (e) { btn.textContent = old; handleError(e); }
    finally { btn.disabled = false; }
  }
  async function matUpload(files, btn) {
    const G = MAT.groups || [], g = G.find((x) => x.name === MAT.cat), dir = g ? { id: g.id } : MAT.dir;
    if (!dir) { toast("자료 폴더를 찾지 못했습니다"); return; }
    const ok = files.filter((f) => f.size <= DOCS_MAX); let done = 0; const failed = [];
    if (!ok.length) { toast("50MB가 넘는 파일은 올릴 수 없습니다"); return; }
    btn.disabled = true;
    try {
      const have = new Set((await docsList(dir)).map((f) => f.name));
      for (const f of ok) {
        btn.textContent = `올리는 중 ${done + 1}/${ok.length}`;
        const base = String(f.name || "파일").replace(/[\\/:*?"<>|]/g, "").trim() || "파일", m = base.match(/^(.*?)(\.[A-Za-z0-9]{1,5})?$/); let name = base, k = 2;
        while (have.has(name)) name = `${m[1]} (${k++})${m[2] || ""}`;
        try { await docsUpload(dir, name, f); have.add(name); done++; } catch (e) { if (String(e.message).includes("GOOGLE_AUTH")) throw e; console.error(e); failed.push(f.name); }
      }
    } catch (e) { handleError(e); }
    toast([done ? `${done}개 올림` : "", failed.length ? `${failed.length}개 실패` : "", files.length > ok.length ? `${files.length - ok.length}개는 50MB 초과` : ""].filter(Boolean).join(", ") || "올리지 못했습니다");
    if (SALES.sub === "data") await renderMat(salesReq, true);
  }
  $("salesView").addEventListener("input", (ev) => { if (ev.target && ev.target.id === "matQ") { MAT.q = ev.target.value; matDraw(); } });
  $("salesView").addEventListener("change", (ev) => { const t = ev.target; if (t && t.id === "matFile" && t.files && t.files.length) { const fs0 = [...t.files]; t.value = ""; matUpload(fs0, $("salesView").querySelector("[data-mat-up]")); } });
  $("salesView").addEventListener("click", (ev) => {
    if (SALES.sub !== "data") return;
    const c = ev.target.closest("[data-mat-cat]"); if (c) { MAT.cat = c.dataset.matCat; matDraw(); return; }
    const o = ev.target.closest("[data-mat-open]"); if (o) { const f = matFile(o.dataset.matOpen); if (f) docsView(f); return; }
    const sh = ev.target.closest("[data-mat-share]"); if (sh) { const f = matFile(sh.dataset.matShare); if (f) matShare(f, sh); return; }
    const up = ev.target.closest("[data-mat-up]"); if (up) { const i = $("matFile"); if (i) i.click(); }
  });
  function openSales() { renderSales(); flushContacts(); }
  $("salesSub").addEventListener("click", (ev) => { // 세일즈 서브메뉴: 이달목표 / 자료 (세일즈 화면 위, 화면 바깥 줄)
    const b = ev.target.closest("[data-ssub]");
    if (!b) return;
    SALES.sub = b.dataset.ssub === "material" ? "material" : b.dataset.ssub === "data" ? "data" : b.dataset.ssub === "prog" ? "prog" : "check";
    renderSales();
  });
  $("salesView").addEventListener("click", (ev) => {
    const pgv = ev.target.closest("[data-pgv]");
    if (pgv) { SALES.pgv = pgv.dataset.pgv; return renderSales(true); }
    const pgs = ev.target.closest("[data-pgs]"), pgt = ev.target.closest("[data-pgst]");
    if (pgs || pgt) { const pid = (pgs || pgt).dataset[pgs ? "pgs" : "pgst"], row = (($("salesView")._prog) || []).find((r) => r.pid === pid); if (row) progChange(row, pgs ? "stage" : "status", pgs ? pgs.dataset.v : pgt.dataset.to); return; }
    const arv = ev.target.closest("[data-arview]");
    if (arv) { SALES.arch[SALES.tab] = arv.dataset.arview === "del" ? "del" : arv.dataset.arview === "off"; return renderSales(); }
    const dvb = ev.target.closest("[data-dbv]");
    if (dvb) { SALES.dbv = dvb.dataset.dbv; SALES.arch.valueup = SALES.dbv !== "new"; return renderSales(); }
    const ardn = ev.target.closest("[data-ardelnew]");
    if (ardn) { // v381: 삭제 목록 줄의 [신규로]: 삭제 표시와 보관 표시를 모두 풀어 신규로
      const row = (($("salesView")._rows) || []).find((r) => r.key === ardn.dataset.ardelnew);
      if (!row || SALES.tab !== "valueup") return;
      const el = ardn.closest(".srow"); let first = true;
      const after = () => { if (first && el) el.style.display = "none"; first = false; renderSales(true); };
      const ents = dbEnts(((SALES.data && SALES.data.custById.get(row.custId)) || [row.ent]).filter(Boolean)); if (!ents) return;
      exclApply("dbdel", ents, false); exclApply("db", ents, false, after);
      undoToast(`${row.holder} 고객을 신규로 이동했습니다`, () => { exclApply("db", ents, true); exclApply("dbdel", ents, true, after); toast("다시 삭제했습니다."); });
      return;
    }
    const ardx = ev.target.closest("[data-ardelx]");
    if (ardx) { // 삭제 목록 줄의 [복구]: 삭제 표시를 풀어 보관 목록으로
      const row = (($("salesView")._rows) || []).find((r) => r.key === ardx.dataset.ardelx);
      if (!row || SALES.tab !== "valueup") return;
      const el = ardx.closest(".srow"); let first = true;
      const after = () => { if (first && el) el.style.display = "none"; first = false; renderSales(true); };
      const ents = dbEnts(((SALES.data && SALES.data.custById.get(row.custId)) || [row.ent]).filter(Boolean)); if (!ents) return;
      exclApply("dbdel", ents, false, after);
      undoToast(`${row.holder} 고객을 보관으로 이동했습니다`, () => { exclApply("dbdel", ents, true, after); toast("다시 삭제했습니다."); });
      return;
    }
    const ard = ev.target.closest("[data-ardel]");
    if (ard || ev.target.closest('[data-arx][data-to="hell"]')) hellGifUrl();
    if (ard) { // DB 보관 줄의 [삭제]: 'DB삭제' 칸에 '삭제' 표시 → 보관 목록에서도 숨김
      const row = (($("salesView")._rows) || []).find((r) => r.key === ard.dataset.ardel);
      if (!row || SALES.tab !== "valueup") return;
      const el = ard.closest(".srow"); let first = true;
      const after = () => { if (first && el) el.style.display = "none"; first = false; renderSales(true); };
      dbDelete(row.holder, ((SALES.data && SALES.data.custById.get(row.custId)) || [row.ent]).filter(Boolean), after);
      return;
    }
    const arx = ev.target.closest("[data-arx]");
    if (arx) { // 줄의 [이동] / [원복]
      const row = (($("salesView")._rows) || []).find((r) => r.key === arx.dataset.arx);
      if (!row) return;
      const el = arx.closest(".srow"); let first = true;
      const after = () => { if (first && el) el.style.display = "none"; first = false; renderSales(true); }; // 누르는 즉시 줄을 숨기고, 목록(건수 칩 포함)은 이어서 다시 그림
      if (SALES.tab === "valueup") {
        const ents = ((SALES.data && SALES.data.custById.get(row.custId)) || [row.ent]).filter(Boolean);
        if (row.arch) dbRestore(row.holder, ents, after); else dbRemove(row.holder, ents, after);
      } else archSalesRow(row, arx.dataset.to || (row.arch ? "new" : "done"), after);
      return;
    }
    const stg = ev.target.closest("[data-stg]");
    if (stg) { const row = (($("salesView")._rows) || []).find((r) => r.key === stg.dataset.key); if (row) setStage(row, stg.dataset.stg); return; }
    const chip = ev.target.closest("[data-chip]");
    if (chip) { // 칩 터치 = ON/OFF 전환만 (접힘/펼침·전화·문자와 별개)
      const row = (($("salesView")._rows) || []).find((r) => r.key === chip.dataset.key);
      if (row && row.dda) toggleChip(row, chip.dataset.chip);
      return;
    }
    const fold = ev.target.closest(".srow.fold");
    if (fold && !ev.target.closest("a, button, input, label, .sbtns, .sline.l1, .smemo")) { // 칩이 있는 이름 줄은 접힘/펼침 제외 // 따행 행 접힘/펼침 (전화·문자 버튼과 이름 링크는 제외)
      const k = fold.dataset.fold, on = fold.classList.toggle("open");
      if (on) SALES.open.add(k); else SALES.open.delete(k);
      return;
    }
    const dct = ev.target.closest("[data-dct]");
    if (dct) { SALES.dct = dct.dataset.dct; return renderSales(); }
    const dcl = ev.target.closest("[data-dcls]");
    if (dcl) { SALES.dcls = SALES.dcls === dcl.dataset.dcls ? "" : dcl.dataset.dcls; return renderSales(); }
    if (ev.target.closest("#dactBtn")) { SALES.dact = !SALES.dact; return renderSales(); }
    const st = ev.target.closest("[data-stab]");
    if (st) { SALES.tab = ["dda", "valueup"].includes(st.dataset.stab) ? st.dataset.stab : "transfer"; return renderSales(); }
    const ym = ev.target.closest("[data-ym]");
    if (ym) { SALES.ym = ymShift(SALES.ym || ymNow(), +ym.dataset.ym); return renderSales(); }
    if (ev.target.closest("#salesReset")) { const b = $("salesView"); return resetSalesMarks(b._rows || [], b._mark || new Map(), SALES.ym, SALES.tab === "transfer" ? "이달이관" : "이달 따행", SALES.fc); }
    const cu = ev.target.closest("[data-cust]");
    if (cu) {
      ev.preventDefault();
      const es = (SALES.data && SALES.data.custById.get(cu.dataset.cust)) || [];
      if (es.length) { const g = groupCustomers(es); g.forEach((x) => { x.entries.sort((a, b) => prio(a.type) - prio(b.type)); x.name = x.entries[0].name; }); openCustomer(g[0]); }
      return;
    }
    const ct = ev.target.closest("[data-contact]");
    if (ct && !ct.classList.contains("off")) { // 전화·문자: 링크는 그대로 열리게 두고, 그 전에 기기에 대기 기록을 먼저 저장
      const row = ((($("salesView")._rows) || []).find((r) => r.key === ct.dataset.key));
      if (!row) return;
      if (ct.classList.contains("done")) { // 이미 시도 표시된 버튼을 다시 누르면: 전화·문자 앱은 열지 않고 표시와 기록을 지움
        ev.preventDefault();
        if (isRO(row.fc)) return toast("보기 전용 권한이라 수정할 수 없습니다.");
        const kind = ct.dataset.contact, ymNow = SALES.ym;
        ct.classList.remove("done"); ct.textContent = ct.textContent.replace(/\s*✓\s*$/, "");
        toast("시도 표시를 취소하고 기록을 지우는 중…");
        cancelContact(row, kind, ymNow).then(() => { toast("시도 표시와 기록을 지웠습니다."); renderSales(true); })
          .catch((e) => { handleError(e); renderSales(true); });
        return;
      }
      // 한 번 더 확인: 잘못 눌러 전화가 걸리거나 문자 창이 열리는 것을 막음
      if (!confirm(`${row.holder ? row.holder + "님께 " : ""}${ct.dataset.contact === "call" ? "정말로 전화를 거시겠습니까?" : "정말로 문자를 보내시겠습니까?"}`)) { ev.preventDefault(); return; }
      logAttempt(attemptPage(), ct.dataset.contact, row.holder, smemoFc(row.custId, row.fc)); // v541: 상담기록에도 남김
      if (row.ref || !row.fileId || !row.key) { toast(row.ref ? "담당이 다른 FC 계약이라 기록하지 않습니다." : "새 구조 파일의 계약만 기록합니다."); return; }
      const q = readQueue();
      q.push({ uid: Date.now() + "-" + Math.random().toString(36).slice(2, 7), kind: ct.dataset.contact, ts: nowStamp(), cid: row.cid, custId: row.custId, fc: row.fc, fileId: row.fileId, src: SALES.tab === "dda" ? "세일즈 화면(따행)" : "세일즈 화면(이관)" });
      writeQueue(q);
      ct.classList.add("done");
      setTimeout(flushContacts, 4000);
    }
  });
  document.addEventListener("click", (ev) => { // 고객 카드 따행 섹션: 대상년월 선택
    const b = ev.target.closest && ev.target.closest("[data-fidx]");
    if (b && current && !$("customerView").hidden) { current.followIdx = +b.dataset.fidx; selectTab("basic"); }
  });
  document.addEventListener("click", (ev) => { // 고객 FC 선택 (전체 / FC 이름)
    const b = ev.target.closest && ev.target.closest("#fcOpts [data-fco]");
    if (!b) return;
    try { localStorage.setItem(listFcKey(), b.dataset.fco); } catch (e) { /* 저장 못 해도 이번엔 적용 */ }
    search(curQuery(), true);
  });
  document.addEventListener("change", (ev) => { // 고객 FC 선택창(FC가 6명 이상일 때)
    if (!ev.target || ev.target.id !== "fcoSel") return;
    try { localStorage.setItem(listFcKey(), ev.target.value); } catch (e) { /* 저장 못 해도 이번엔 적용 */ }
    search(curQuery(), true);
  });
  document.addEventListener("click", (ev) => { // 성향 버튼: Good → Bad → Shit → Good
    if (!(ev.target.closest && ev.target.closest("#gradeOpt"))) return;
    gradeFilter = gradeFilter === "Good" ? "Bad" : gradeFilter === "Bad" ? "Shit" : "Good";
    syncGradeOpt();
    search(curQuery(), true);
  });
  document.addEventListener("click", (ev) => { // 고객 범위 옵션 (전체 = 해제하고 바로 검색 / 본인·이관·DB·따행·가망 = 다중 선택)
    const b = ev.target.closest && ev.target.closest("#searchOpts [data-opt]");
    if (!b) return;
    const id = b.dataset.opt;
    if (id === "progress") progShowRemoved = false; // 진행 버튼을 누르면 기본(진행 고객)부터
    if (id === "db") dbShowRemoved = false; // DB 버튼도 기본(DB 고객)부터
    if (id === "transfer") trShowRemoved = false; // 이관 버튼도 기본(이관 고객)부터
    if (id === "all") optSel.clear(); else if (optSel.has(id)) optSel.delete(id); else optSel.add(id);
    syncOpts();
    search(curQuery(), true);
  });
  document.addEventListener("change", (ev) => {
    if (!ev.target) return;
    if (ev.target.id === "salesRef") { SALES.showRef = ev.target.checked; renderSales(); }
    if (ev.target.id === "dclsSel") { SALES.dcls = ev.target.value; renderSales(); }
    if (ev.target.dataset && ev.target.dataset.nf) { SALES.nf[ev.target.dataset.nf] = ev.target.checked; renderSales(); return; }
    if (ev.target.dataset && ev.target.dataset.df) { SALES.df[ev.target.dataset.df] = ev.target.checked; renderSales(); }
    if (ev.target.id === "salesUser") { SALES.user = ev.target.value; SALES.fc = "__own"; try { localStorage.setItem(salesUserKey(), SALES.user); } catch (e) { /* 무시 */ } renderSales(); return; } // 사용자를 바꾸면 그 사용자 본인 FC부터
    if (ev.target.id === "salesTab") { const v = ev.target.value; SALES.tab = ["dda", "valueup"].includes(v) ? v : "transfer"; renderSales(); return; }
    if (ev.target.id === "salesFc") { SALES.fc = ev.target.value; try { localStorage.setItem(salesFcKey(), SALES.fc); } catch (e) { /* 무시 */ } renderSales(); }
  });

  // ---------- 말만해요 연동 ----------
  //  주소 #go=customers&q=검색어 (고객 검색 목록) / #go=customer&id=고객키 (고객 정보). 검색어·고객 키는 # 뒤에 둬서 서버로 안 나감.
  //  ?embed=1 이면 말만해요 안에 모듈로 떠 있는 모드. 로그인 뒤에 이어서 처리하고, 처리하면 주소에서 # 뒤를 지움
  let GO0 = null, appStarted = false;
  try {
    const h = new URLSearchParams(location.hash.replace(/^#/, "")), go = h.get("go");
    if (go) {
      GO0 = { go, q: h.get("q") || "", id: h.get("id") || "", kind: h.get("kind") || "", tab: h.get("tab") || "" };
      history.replaceState(history.state, "", location.pathname + location.search); // 새로고침해도 다시 처리하지 않게 주소에서 지움
    }
  } catch (e) { GO0 = null; }
  async function goCustomer(id, tab) {
    try {
      const gs = groupCustomers(await loadLists());
      const c = gs.find((x) => custKey(x) === id) || gs.find((x) => custKeyOld(x) === id);
      if (!c) { toast("고객을 찾지 못했습니다."); return; }
      c.entries.sort((a, b) => prio(a.type) - prio(b.type)); c.name = c.entries[0].name;
      await openCustomer(c); if (tab) selectTab(tab);
    } catch (e) { handleError(e); }
  }
  // 이미 앱 안에 있을 때는 새로 열지 않고 바로 처리. 아직 로그인 전이면 로그인 뒤에 처리
  window.blue10kpGo = function (o) {
    if (!o || (o.go !== "customers" && o.go !== "customer" && o.go !== "schedule" && o.go !== "sales")) return;
    if (!appStarted) { GO0 = { go: o.go, q: String(o.q || ""), id: String(o.id || ""), kind: String(o.kind || ""), tab: String(o.tab || "") }; return; }
    if (typeof helpClose === "function" && helpIsOpen) helpClose();
    if (!$("aiSheet").hidden && typeof aiClose === "function") aiClose();
    if (o.go === "sales") { // q = 탭(transfer=이달이관, dda=따행), id = 달(YYYY-MM, 없으면 이번 달)
      const mm = String(o.id || "").match(/^(\d{4})-?(\d{2})$/);
      SALES.sub = "check"; // 말만해요에서 열면 이달목표
      SALES.tab = o.q === "dda" ? "dda" : o.q === "valueup" ? "valueup" : "transfer";
      SALES.goFc = String(o.kind || "").trim(); // FC 이름(없으면 본인 FC)
      SALES.ym = mm && +mm[2] >= 1 && +mm[2] <= 12 ? mm[1] + mm[2] : ymNow();
      SALES.arch[SALES.tab] = o.tab === "archive"; // tab=archive: 그 탭의 보관 목록
      loadSettings().then(() => openSales());
      return;
    }
    if (o.go === "schedule") { // q = YYYY-MM-DD (없거나 틀리면 오늘): 일정 화면의 하루 보기
      const m = String(o.q || "").match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
      const d = m ? new Date(+m[1], +m[2] - 1, +m[3]) : null;
      schedMode = "day"; schedDate = d && !isNaN(+d) ? dayStart(d) : dayStart(new Date());
      loadSettings().then(() => openSchedule());
      return;
    }
    if (o.go === "customer") { const id = String(o.id || ""); if (id) { if (!$("customerView").hidden) show("searchView"); goCustomer(id, String(o.tab || "")); } return; }
    const q = String(o.q || "").trim();
    let kd = String(o.kind || "").trim();
    progShowRemoved = kd === "progress_removed" || kd === "progress_archive"; if (progShowRemoved) kd = "progress"; // kind=progress_archive(예전 progress_removed): 진행 보관 고객 보기
    dbShowRemoved = kd === "db_archive"; if (dbShowRemoved) kd = "db"; // kind=db_archive: DB 보관 고객 보기
    trShowRemoved = kd === "transfer_archive"; if (trShowRemoved) kd = "transfer"; // kind=transfer_archive: 이관 보관 고객 보기
    // 고객 범위 옵션(contract·transfer·db·ddahaeng·plan): 그 옵션만 켜고 보여 줌. q가 비어도 그 분류 전체 목록
    if (OPT_LABEL[kd]) { optSel.clear(); optSel.add(kd); syncOpts(); }
    if (!q) {
      if (OPT_LABEL[kd]) { show("searchView"); $("q").value = ""; $("q2").value = ""; search("", true); }
      else show("searchView");
      return;
    }
    goKeepList = q; goTabPending = String(o.tab || "");
    runTopSearch(q, !(o.memo === "1" || o.memo === 1 || o.memo === true)); // 말만해요의 memo=1이면 전체검색(메모·상담기록·일정 포함), 아니면 고객검색
  };
  // 말만해요 안에 모듈(iframe)로 떠 있을 때: 부모 말만해요에 지금 위치를 알림 (고객 이름이 든 값은 같은 주소(origin)에만 보냄)
  function notifyWhere(w) {
    if (!EMBED) return;
    try { parent.postMessage({ type: "blue10kp-where", app: "shlife", go: w.go, q: w.q || "", id: w.id || "", kind: w.kind || "", tab: w.tab || "", canBack: !!w.canBack }, location.origin); } catch (e) { /* 무시 */ }
  }
  const whereNow = () => {
    if (!$("customerView").hidden && current && current.cust) return { go: "customer", id: custKey(current.cust), tab: ((document.querySelector("#tabs button.on") || {}).dataset || {}).tab || "", canBack: true }; // 고객 정보에서는 앱 안 뒤로(목록)가 있음. 검색 목록이 첫 화면이면 canBack=false
    if (!$("searchView").hidden) return { go: "customers", q: curQuery(), kind: optSel.size === 1 ? ([...optSel][0] === "progress" && progShowRemoved ? "progress_archive" : [...optSel][0] === "db" && dbShowRemoved ? "db_archive" : [...optSel][0] === "transfer" && trShowRemoved ? "transfer_archive" : [...optSel][0]) : "" };
    if (!$("salesView").hidden) return { go: "sales", q: SALES.tab === "dda" ? "dda" : SALES.tab === "valueup" ? "valueup" : "transfer", id: SALES.ym ? SALES.ym.slice(0, 4) + "-" + SALES.ym.slice(4) : "", kind: SALES.fc || "", tab: SALES.arch[SALES.tab] ? "archive" : "" }; // 탭·달·보관 목록 (앱으로 열기용)
    if (!$("scheduleView").hidden) return { go: "schedule", q: `${schedDate.getFullYear()}-${pad2(schedDate.getMonth() + 1)}-${pad2(schedDate.getDate())}` }; // 보고 있는 날짜(앱으로 열기용)
    return { go:!$("salesView").hidden ? "sales" : !$("adminView").hidden ? "settings" : "" };
  };
  notifyWhereNow = () => { if (EMBED) { clearTimeout(notifyWhereNow._t); notifyWhereNow._t = setTimeout(() => { const w = whereNow(); if (w.go) notifyWhere(w); }, 150); } };
  // 오른쪽 아래에 떠 있는 버튼(수정·닫기 줄, 취소·확인, 토스트, 아래 시트 등)이 말만해요 버튼 자리(오른쪽 88px)를 가리면,
  // 그 가장 위쪽 끝이 화면 바닥에서 몇 px인지 부모에 알려 말만해요가 버튼을 그만큼 올리게 함 (없으면 0)
  if (EMBED) {
    let lastLift = -1;
    setInterval(() => {
      let lift = 0;
      const IW = innerWidth, IH = innerHeight;
      document.querySelectorAll(".cust-bar button, .toast:not([hidden]), .toast2, .help-sheet").forEach((el) => {
        const cs = getComputedStyle(el);
        if (cs.display === "none" || cs.visibility === "hidden" || el.closest("[hidden]")) return;
        const r = el.getBoundingClientRect();
        if (r.width <= 0 || r.height <= 0 || r.bottom < IH - 240 || r.right < IW - 96 || r.top >= IH) return; // 오른쪽 아래 구역과 겹칠 때만
        lift = Math.max(lift, Math.round(IH - r.top));
      });
      if (lift === lastLift) return;
      lastLift = lift;
      try { parent.postMessage({ type: "blue10kp-fab", lift }, location.origin); } catch (e) { /* 무시 */ }
    }, 250);
  }
  // 부모 말만해요가 보내는 새 요청 받기 (보낸 곳 origin 확인)
  window.addEventListener("message", (ev) => {
    if (ev.origin !== location.origin || !ev.data) return;
    if (ev.data.type === "blue10kp-back") { // 말만해요의 왼쪽 위 뒤로가기: 고객 정보에서는 목록으로 한 단계 뒤로
      if (!$("customerView").hidden) confirmLeaveEdit().then((ok) => { if (ok) history.back(); });
      return;
    }
    if (ev.data.type === "blue10kp-go") window.blue10kpGo(ev.data);
    if (ev.data.type === "blue10kp-theme" && (ev.data.theme === "dark" || ev.data.theme === "light")) { // 말만해요 설정에서 모드를 바꿈: 열려 있는 앱도 바로 다시 칠함(저장하지 않음)
      embedTheme = ev.data.theme; themeApply();
      const c = document.getElementById("darkMode"); if (c) c.checked = embedTheme === "dark";
    }
  });

  // 서비스워커: 앱 설치(홈 화면에 추가) 조건. 실패해도 앱은 그대로 동작
  try { if ("serviceWorker" in navigator && (location.protocol === "https:" || location.hostname === "localhost" || location.hostname === "127.0.0.1")) navigator.serviceWorker.register("sw.js").catch(() => {}); } catch (e) { /* 무시 */ }
  (async function init() {
    if (CFG.clientId === "YOUR-CLIENT-ID") {
      show("loginView");
      $("loginBtn").disabled = true;
      toast("config.js에 clientId를 넣어 주세요.");
      return;
    }
    if (GD) {
      // 구글 토큰이 살아 있으면 바로 시작, 아니면 로그인(다시 연결) 버튼 (팝업은 버튼을 눌러야 열림)
      migrateLegacyGToken();
      await am6Check(); // v557: 6시 전에 받은 출입증이면 지우고 로그인 화면
      const sh = readGAuth();
      if (sh && sh.email && storedGToken()) { // 다른 앱에서 한 로그인도 인정: 계정이 다르면 이전 계정 자료를 비우고 새 계정으로 시작
        if (gAccount() && gAccount() !== String(sh.email).toLowerCase()) await photoCacheClear();
        try { localStorage.setItem(GACC_KEY, String(sh.email).toLowerCase()); } catch (e) { /* 이번만 */ }
        startApp();
      } else showGLogin();
      return;
    }
    if (!msal) { // OneDrive 옛 모드는 지원 중단(구글 드라이브만 씀) — 시작 속도를 위해 msal 라이브러리를 더 이상 받지 않는다
      show("loginView"); $("loginBtn").disabled = true; toast("이 버전은 Google Drive만 지원합니다."); return;
    }
    await msal.initialize();
    try { await msal.handleRedirectPromise(); } catch (e) { handleError(e); }
    if (msal.getAllAccounts().length) startApp();
    else show("loginView");
  })();

  // 화면이 다시 보일 때 공통 출입증을 다시 읽음: 다른 앱에서 새로 로그인했으면 이어서 쓰고, 계정이 바뀌었으면 새로 시작
  if (GD) window.addEventListener("storage", (e) => { if (e.key === GAUTH_KEY && !$("loginView").hidden) checkSharedLogin(); }); // 말만해요 안에서 보일 때: 부모가 로그인하면 바로 이어서
  if (GD) document.addEventListener("visibilitychange", () => { if (document.visibilityState === "visible") am6Check().then((out) => { if (!out) { checkSharedLogin(); if (appStarted && $("loginView").hidden) ntfCheck(); } }); });
  if (GD) { setInterval(() => { am6Check(); }, 60000); window.addEventListener("focus", () => { am6Check(); }); } // v557: 켜 둔 동안 1분마다
  async function checkSharedLogin() {
    const sh = readGAuth();
    if (!sh || !sh.email || !storedGToken()) return;
    const em = String(sh.email).toLowerCase();
    if (!$("loginView").hidden) { // 로그인 화면이던 중 다른 앱에서 로그인함
      if (checkSharedLogin.busy) return;
      checkSharedLogin.busy = true;
      if (gAccount() && gAccount() !== em) await photoCacheClear();
      try { localStorage.setItem(GACC_KEY, em); } catch (e) { /* 이번만 */ }
      startApp();
    } else if (gAccount() && gAccount() !== em && !editMode) { // 다른 앱에서 계정을 바꿨음
      await photoCacheClear();
      try { localStorage.setItem(GACC_KEY, em); } catch (e) { /* 이번만 */ }
      location.reload();
    }
  }

  // 말만해요 해석용 색인: 이 사용자가 볼 수 있는 고객의 이름·고객ID·담당FC·분류·생년(연도만)만 담아 사용자 Drive Claude/app/신한FC/설정/말만해요_색인.json에 저장
  //   (주민번호·계좌·전화·주소·이메일·메모는 넣지 않음). 목록을 읽을 때 만들고, 내용이 같으면 다시 쓰지 않음.
  async function syncMalIndex(records) {
    if (!GD || !records || !records.length) return;
    const seen = new Map();
    groupCustomers(records).forEach((c) => {
      const e = c.entries.slice().sort((a, b) => prio(a.type) - prio(b.type))[0];
      const id = idOf(e) || "";
      const key = id || (baseName(e.name) + "|" + String(c.birth || ""));
      if (seen.has(key)) return;
      seen.set(key, { 이름: baseName(e.name), 고객ID: id, 담당FC: [...new Set(c.entries.map((x) => x.fc).filter(Boolean))].join(","), 분류: e.type || "", 생년: birthYear(c.birth) || "" });
    });
    const items = [...seen.values()].sort((a, b) => a.이름.localeCompare(b.이름, "ko") || String(a.고객ID).localeCompare(String(b.고객ID)));
    const body = JSON.stringify(items);
    let h = 0; for (let i = 0; i < body.length; i++) h = (h * 31 + body.charCodeAt(i)) | 0;
    const hk = "shlife.malIdx." + me() + "." + items.length + "." + h;
    try { if (localStorage.getItem("shlife.malIdxHash." + me()) === hk) return; } catch (e) { /* 무시 */ }
    const d = await resolveMyDir(true); // 내 Drive의 Claude/app/신한FC/설정
    if (!d || !d.id) return;
    await stWriteJsonIn(d, "말만해요_색인.json", { v: 1, 갱신: new Date().toISOString(), 계정: me(), 고객: items });
    try { localStorage.setItem("shlife.malIdxHash." + me(), hk); } catch (e) { /* 무시 */ }
  }

  // v512→v513(2026-11-21 사용자 규칙): 관리자 > FC 사용 설정의 로그인 토글이 '승인'이 아닌 계정은 로그인 화면에서 멈춤(토글 도입 전 파일은 users에 있으면 통과)
  //   권한 파일을 못 읽는(공유받지 않은) 계정도 등록 안 된 것으로 봄. 통신 오류처럼 판단할 수 없을 때는 막지 않음
  async function isRegistered() {
    if (!GD || isAdmin()) return true;
    try {
      await resolveAppDir();
      if (base && base.owner) return true;
      if (!appDir || !appDir.id) return false;
      const p = await stReadJsonIn(appDir, PERM_FILE);
      return loginOn(p, me());
    } catch (e) {
      const m = String(e && e.message);
      if (m === "NO_SHARE" || /Drive 40[34]/.test(m)) return false;
      if (m === "GOOGLE_AUTH") throw e;
      console.warn("등록 확인 못 함, 그대로 진행", e); return true;
    }
  }
  // ---------- v557 전체 알림 (2026-11-24 사용자 "이대로 코딩": 로그인 직후 알림 창 + 관리자 > 전체 알림) ----------
  //   알림 목록 = 앱 설정 폴더 shlife-알림.json {v, system, items:[{id, kind:"date"|"dow"|"system", sys, text, from, to, dows:[0~6], target:["*"]|[FC…], go, on}]}
  //   확인 기록 = 그 사람 FC 폴더의 알림확인.json {이메일: {확인키: 시각}} + 이 기기(localStorage). Drive에 못 남겨도 이 기기 기록으로 동작
  //   시스템 알림 = 새 이관·새 DB(내 FC, 지난번 확인 뒤 새로 생긴 줄). 처음 한 번은 지금 목록을 기억만 함
  const NTF_FILE = "shlife-알림.json", NTF_ACK = "알림확인.json";
  const NTF_GO = [["", "이동 없음"], ["sales:od", "세일즈 › 보고서 › 연체/실효"], ["sales:transfer", "세일즈 › 이달목표 › 이관"], ["sales:dda", "세일즈 › 이달목표 › 따행"], ["sales:valueup", "세일즈 › 이달목표 › DB"], ["sales:prog", "세일즈 › 계약진행"], ["sales:material", "세일즈 › 보고서"], ["sales:data", "세일즈 › 자료"], ["dash", "대시보드"], ["schedule", "일정"], ["customer", "고객"]];
  const NTF_SYS = [["newOverdue", "{t}가 나왔습니다. 확인해 주세요.", "sales:od"], ["newLapse", "{t}가 나왔습니다. 확인해 주세요.", "sales:od"], ["newTransfer", "새 이관 계약 {n}건이 이달목표에 올라왔습니다.", "sales:transfer"], ["newDb", "새 밸류업DB {n}건이 배분되었습니다.", "sales:valueup"]]; // v564: 연체·실효 리스트가 나온 날 알림(목록 위쪽 = 이동 우선)
  const NTF_DOW = "일월화수목금토";
  const ntfGoLabel = (g) => (NTF_GO.find((x) => x[0] === (g || "")) || NTF_GO[0])[1];
  const ntfDay = () => am6Today();
  const ntfMd = (s) => { const m = String(s || "").match(/^\d{4}-(\d{2})-(\d{2})/); return m ? +m[1] + "월 " + +m[2] + "일" : ""; };
  function ntfNorm(d) {
    d = d && typeof d === "object" ? d : {};
    const items = (Array.isArray(d.items) ? d.items : []).filter((x) => x && x.id);
    NTF_SYS.slice().reverse().forEach(([k, t, go]) => { if (!items.some((x) => x.kind === "system" && x.sys === k)) (k === "newOverdue" || k === "newLapse" ? items.unshift.bind(items) : items.push.bind(items))({ id: "sys-" + k, kind: "system", sys: k, text: t, go, target: ["*"], on: true }); });
    return { v: 1, system: d.system !== false, items };
  }
  async function ntfLoad() { await resolveAppDir(); return ntfNorm(await stReadJsonIn(appDir, NTF_FILE)); }
  async function ntfSave(doc) { await resolveAppDir(); await stWriteJsonIn(appDir, NTF_FILE, { v: 1, system: doc.system, items: doc.items, at: new Date().toISOString(), by: me() }); }
  function ntfToday(it, day, dow) {
    if (it.on === false) return false;
    if (it.kind === "date") { const f = String(it.from || ""), t = String(it.to || it.from || ""); return !!f && f <= day && day <= t; }
    if (it.kind === "dow") return (it.dows || []).map(Number).includes(dow);
    return false;
  }
  const ntfTargets = (it) => (it.target && it.target.length ? it.target : ["*"]);
  const ntfForMe = (it, fc) => { const t = ntfTargets(it); return t.includes("*") || (!!fc && t.includes(fc)); };
  const ntfAckKey = (it, day) => (it.kind === "date" ? it.id : it.id + "@" + day); // 날짜 알림은 한 번 확인하면 끝, 요일·시스템은 그날만
  const ntfLsKey = () => "shlife.ntfAck." + (me() || "me");
  function ntfLocal() { try { return JSON.parse(localStorage.getItem(ntfLsKey()) || "{}") || {}; } catch (e) { return {}; } }
  function ntfLocalPut(o) { const ks = Object.keys(o); if (ks.length > 300) ks.sort((a, b) => String(o[a]).localeCompare(String(o[b]))).slice(0, ks.length - 300).forEach((k) => delete o[k]); try { localStorage.setItem(ntfLsKey(), JSON.stringify(o)); } catch (e) { /* 기기에 못 남겨도 진행 */ } }
  async function ntfFcDir(fc) { try { const P = await readPerms(); const id = ((P && P.fcFolders) || {})[fc] || fcFolderIds[fc]; return id ? { id } : null; } catch (e) { return null; } }
  const NTF_SEEN = () => "shlife.ntfSeen." + (me() || "me");
  async function ntfSystem(doc, fc) {
    if (!doc.system || !fc) return [];
    const on = doc.items.filter((x) => x.kind === "system" && x.on !== false && ntfForMe(x, fc));
    if (!on.length) return [];
    let data; try { data = await loadSalesData(); } catch (e) { return []; }
    const keys = {
      newTransfer: () => transferRows(data, ymNow(), false).filter((r) => r.srcFc === fc && !r.archDay && !r.del && r.key).map((r) => r.key),
      newDb: () => valueupRows(data).filter((r) => r.srcFc === fc && !r.dbOff).map((r) => r.key)
    };
    let seen = {}; try { seen = JSON.parse(localStorage.getItem(NTF_SEEN()) || "{}") || {}; } catch (e) { seen = {}; }
    const out = []; let first = false;
    // v564: 연체·실효 리스트 — 나온 지 7일(실효는 그 달·지난달) 안이고 이 기기에서 아직 확인 안 한 파일이면 알림(처음 기억만 하는 규칙 없이)
    const odOn = on.filter((x) => x.sys === "newOverdue" || x.sys === "newLapse");
    if (odOn.length && OD_ON) {
      try {
        if (!OD.data) OD.data = data;
        const f = await odFolder(fc), pk = f && f.list ? odPick(f.list) : {};
        const ago = new Date(Date.now() - 7 * 864e5), lim = ago.getFullYear() + "-" + String(ago.getMonth() + 1).padStart(2, "0") + "-" + String(ago.getDate()).padStart(2, "0");
        const pm = new Date(); pm.setDate(1); pm.setMonth(pm.getMonth() - 1); const lm = pm.getFullYear() + "-" + String(pm.getMonth() + 1).padStart(2, "0");
        odOn.forEach((it) => {
          const x = it.sys === "newOverdue" ? pk.cur : pk.lapse; if (!x || x.정식 === false) return;
          const fresh = it.sys === "newOverdue" ? String(x.확인일 || "") >= lim : String(x.월 || "") >= lm;
          const key = String(x.파일 || x.확인일 || x.월 || ""), prev = Array.isArray(seen[it.sys]) ? seen[it.sys] : [];
          if (!key || !fresh || prev.includes(key)) return;
          const t = it.sys === "newOverdue" ? odTitle(x) : String(x.제목 || "실효리스트");
          out.push({ ...it, text: String(it.text || "").replace("{t}", t).replace("{n}", "1"), _now: [...prev, key].slice(-24) });
        });
      } catch (e) { console.warn("연체 알림 확인 못 함", e); }
    }
    on.filter((x) => x.sys !== "newOverdue" && x.sys !== "newLapse").forEach((it) => {
      const f = keys[it.sys]; if (!f) return;
      const now = f(), prev = seen[it.sys];
      if (!Array.isArray(prev)) { seen[it.sys] = now; first = true; return; } // 처음 = 지금 목록을 기억만
      const nw = now.filter((k) => !prev.includes(k));
      if (nw.length) out.push({ ...it, text: String(it.text || "").replace("{n}", nw.length), _now: now });
    });
    if (first) { try { localStorage.setItem(NTF_SEEN(), JSON.stringify(seen)); } catch (e) { /* 무시 */ } }
    return out;
  }
  let ntfShown = "";
  async function ntfCheck() {
    if (EMBED || !GD) return;
    const day = ntfDay(); if (ntfShown === day || document.getElementById("ntfPop")) return; ntfShown = day;
    try {
      const doc = await ntfLoad(), fc = await myFcName(), dow = new Date().getDay();
      const dir = fc ? await ntfFcDir(fc) : null;
      let rem = null; if (dir) { try { rem = await stReadJsonIn(dir, NTF_ACK); } catch (e) { rem = null; } }
      const done = Object.assign({}, ntfLocal(), (rem || {})[me()] || {});
      const list = doc.items.filter((it) => it.kind !== "system" && ntfToday(it, day, dow) && ntfForMe(it, fc) && !done[ntfAckKey(it, day)]);
      const all = [...(await ntfSystem(doc, fc)), ...list];
      if (!all.length) return;
      ntfPopup(all, fc, () => ntfAck(all, day, dir));
    } catch (e) { ntfShown = ""; console.warn("알림 확인 못 함", e); }
  }
  async function ntfAck(all, day, dir) {
    const stamp = new Date().toISOString(), loc = ntfLocal();
    all.forEach((it) => { loc[ntfAckKey(it, day)] = stamp; });
    ntfLocalPut(loc);
    const sys = all.filter((it) => it._now);
    if (sys.length) { let seen = {}; try { seen = JSON.parse(localStorage.getItem(NTF_SEEN()) || "{}") || {}; } catch (e) { seen = {}; } sys.forEach((it) => { seen[it.sys] = it._now; }); try { localStorage.setItem(NTF_SEEN(), JSON.stringify(seen)); } catch (e) { /* 무시 */ } }
    const go = (all.find((x) => x.kind === "system" && x.go) || all.find((x) => x.go) || {}).go; // 이동: 시스템 알림 → 목록 위쪽
    if (go) ntfGo(go);
    if (dir) { // 확인 기록(관리자 화면 '확인 N명'용). 보기 전용 FC 등 못 쓰면 이 기기 기록만
      try { const cur = (await stReadJsonIn(dir, NTF_ACK)) || {}; const m = cur[me()] || {}; all.forEach((it) => { m[ntfAckKey(it, day)] = stamp; }); cur[me()] = m; await stWriteJsonIn(dir, NTF_ACK, cur); } catch (e) { console.warn("알림 확인 기록을 Drive에 남기지 못함", e); }
    }
  }
  function ntfGo(g) {
    const [a, b] = String(g).split(":");
    salesClearState(); SALES.fromNtf = true; // v564(사용자 지시): 알림을 보고 간 화면은 이어 보기에 남기지 않음 → 다음에 열면 첫 화면(대시보드)부터
    if (a === "sales" && b === "od") { // 세일즈 › 보고서의 연체/실효 리스트로
      SALES.sub = "material";
      loadSettings().then(() => openSales());
      let n = 0; const tm = setInterval(() => { const c = $("odCard"); if (c && !/불러오는 중/.test(c.textContent)) { clearInterval(tm); c.scrollIntoView({ block: "start", behavior: "smooth" }); } else if (++n > 40) clearInterval(tm); }, 300);
      return;
    }
    if (a === "sales") {
      SALES.sub = ["prog", "material", "data"].includes(b) ? b : "check";
      if (SALES.sub === "check") { SALES.tab = ["dda", "valueup"].includes(b) ? b : "transfer"; SALES.ym = ymNow(); }
      loadSettings().then(() => openSales());
    } else if (a === "dash") openDash();
    else if (a === "schedule") { schedMode = "day"; schedDate = dayStart(new Date()); loadSettings().then(() => openSchedule()); }
    else if (a === "customer") resetSearch();
  }
  function ntfMeta(it, fc) {
    const t = ntfTargets(it), who = t.includes("*") ? (it.kind === "system" ? fc || "내 FC" : "전체") : t.join(" · ");
    const when = it.kind === "system" ? `<span class="ntf-sys">시스템</span>` : it.kind === "dow" ? `매주 ${(it.dows || []).map((n) => NTF_DOW[n]).join("·")}요일` : esc(ntfMd(it.from) + (it.to && it.to !== it.from ? " ~ " + ntfMd(it.to) : ""));
    return `${when} · ${esc(who)}`;
  }
  // 닫기·건너뛰기 없이 [확인]만 (바깥 터치·뒤로로 닫히지 않음)
  function ntfPopup(list, fc, onOk) {
    const d = new Date(), goIt = list.find((x) => x.kind === "system" && x.go) || list.find((x) => x.go);
    const box = document.createElement("div");
    box.id = "ntfPop"; box.className = "modal ntf-pop";
    box.innerHTML = `<div class="modal-box ntf-box" role="alertdialog" aria-modal="true" aria-labelledby="ntfH">
      <h3 id="ntfH"><span class="ntf-dot" aria-hidden="true"></span>중요 알림${list.length > 1 ? " " + list.length + "건" : ""}</h3>
      <p class="ntf-sub">${d.getMonth() + 1}월 ${d.getDate()}일 (${NTF_DOW[d.getDay()]})</p>
      <div class="ntf-list">${list.map((it) => `<div class="ntf-al"><div class="ntf-m">${ntfMeta(it, fc)}</div><p>${esc(it.text)}</p>${it === goIt ? `<div class="ntf-go">확인하면 이동 · ${esc(ntfGoLabel(it.go))}</div>` : ""}</div>`).join("")}</div>
      <button type="button" class="ntf-ok">확인</button></div>`;
    box.querySelector(".ntf-ok").addEventListener("click", () => { box.remove(); onOk(); });
    document.body.appendChild(box);
    setTimeout(() => { const b = box.querySelector(".ntf-ok"); if (b) b.focus(); }, 50);
  }
  // 관리자 > 전체 알림
  const NTFA = { doc: null, f: "all", edit: null, acks: null };
  async function openNtfAdmin() {
    $("adminBody").innerHTML = '<p class="hint pad">알림을 불러오는 중…</p>';
    window.scrollTo(0, 0);
    try {
      NTFA.doc = await ntfLoad(); NTFA.acks = null; NTFA.edit = null;
      renderNtfAdmin();
      ntfLoadAcks().then((a) => { NTFA.acks = a; if (document.querySelector(".ntf-adm")) renderNtfAdmin(); }).catch(() => {});
    } catch (e) { handleError(e); }
  }
  async function ntfLoadAcks() {
    const P = await readPerms().catch(() => null), out = {};
    const ids = [...new Set(Object.values((P && P.fcFolders) || {}).filter(Boolean))];
    await Promise.all(ids.map(async (id) => {
      try { const d = await stReadJsonIn({ id }, NTF_ACK); Object.entries(d || {}).forEach(([em, m]) => Object.keys(m || {}).forEach((k) => (out[k] || (out[k] = new Set())).add(em))); } catch (e) { /* 그 폴더는 건너뜀 */ }
    }));
    return out;
  }
  const ntfSw = (attr, on) => `<button type="button" class="ntf-sw${on ? " on" : ""}" ${attr} role="switch" aria-checked="${on}" aria-label="켜기·끄기"></button>`;
  function renderNtfAdmin() {
    const doc = NTFA.doc, f = NTFA.f, day = ntfDay(), sysOff = !doc.system;
    const shown = doc.items.filter((it) => f === "all" || it.kind === f);
    const md = (s) => String(s || "").slice(5);
    const when = (it) => (it.kind === "date" ? `<b>날짜</b> ${esc(md(it.from))}${it.to && it.to !== it.from ? " ~ " + esc(md(it.to)) : ""}` : it.kind === "dow" ? `<b>요일</b> 매주 ${(it.dows || []).map((n) => NTF_DOW[n]).join("·") || "—"}` : `<span class="ntf-sys">시스템</span>`);
    const tgt = (it) => { const t = ntfTargets(it); return t.includes("*") ? (it.kind === "system" ? "해당 FC" : "전체") : t.join(" · "); };
    const ackN = (it) => { if (it.kind === "system") return ""; if (!NTFA.acks) return `<span class="ntf-ack">확인 …</span>`; const s = NTFA.acks[ntfAckKey(it, day)]; return `<span class="ntf-ack">확인 ${s ? s.size : 0}명${it.kind === "dow" ? " (오늘)" : ""}</span>`; };
    $("adminBody").innerHTML = `<div class="ntf-adm">
      <div class="ntf-head"><button type="button" class="ev-edit org-back-btn" data-ntf-back>‹ 관리자</button><h3 class="sec-title">전체 알림</h3></div>
      <div class="ntf-bar"><div><b>시스템 알림</b><small>새 이관·DB처럼 앱이 자동으로 만드는 알림</small></div>${ntfSw("data-ntf-sys", doc.system)}</div>
      <div class="ntf-fl"><span class="ntf-seg">${[["all", "전체"], ["date", "날짜"], ["dow", "요일"], ["system", "시스템"]].map(([k, l]) => `<button type="button" data-ntf-f="${k}" class="${f === k ? "on" : ""}">${l}${k === "all" ? " " + doc.items.length : ""}</button>`).join("")}</span><button type="button" class="ntf-add" data-ntf-add>+ 추가</button></div>
      ${shown.map((it) => `<div class="ntf-card${it.on === false || (it.kind === "system" && sysOff) ? " off" : ""}"><p>${esc(it.text)}</p>${ntfSw(`data-ntf-on="${esc(it.id)}"`, it.on !== false)}
        <div class="ntf-tg"><span>${when(it)}</span><span><b>대상</b> ${esc(tgt(it))}</span><span><b>이동</b> ${esc(ntfGoLabel(it.go).replace(/^세일즈 › /, ""))}</span></div>
        <div class="ntf-act"><button type="button" data-ntf-edit="${esc(it.id)}">수정</button>${it.kind === "system" ? "" : `<button type="button" class="del" data-ntf-del="${esc(it.id)}">삭제</button>`}${ackN(it)}</div></div>`).join("") || '<p class="hint pad">이 종류의 알림이 없습니다.</p>'}
      <p class="hint ntf-note">켜진 알림 중 오늘에 해당하는 것이 FC 로그인 직후 한 창에 모두 보입니다. [확인]을 누르면 시스템 알림 → 목록 위쪽 알림의 페이지로 이동합니다.</p></div>`;
  }
  function ntfSheet() {
    const it = NTFA.edit, sys = it.kind === "system", fcs0 = [...((adminState && adminState.fcs) || [])], nm = new Set(Object.values(((adminState && adminState.perms) || {}).names || {})), fcs = nm.size ? fcs0.filter((x) => nm.has(x)) : fcs0, tg = ntfTargets(it);
    const old = $("ntfSheet"); if (old) old.remove();
    const box = document.createElement("div");
    box.id = "ntfSheet"; box.className = "modal ntf-shw";
    box.innerHTML = `<div class="ntf-sh" role="dialog" aria-modal="true" aria-label="알림 ${it._new ? "추가" : "수정"}"><h3>알림 ${it._new ? "추가" : "수정"}<button type="button" data-ntfs-x>취소</button></h3>
      <label class="ntf-f"><span>알림 내용${sys ? " ({n} 자리에 새 건수)" : ""}</span><textarea data-ntfs-text rows="2" maxlength="200">${esc(it.text || "")}</textarea></label>
      ${sys ? "" : `<div class="ntf-f"><span>언제 보일까</span><div class="ntf-chips">${[["date", "날짜"], ["dow", "요일"]].map(([k, l]) => `<button type="button" data-ntfs-kind="${k}" class="${it.kind === k ? "on" : ""}">${l}</button>`).join("")}</div></div>
      <div class="ntf-row2" data-ntfs-datebox ${it.kind === "date" ? "" : "hidden"}><label class="ntf-f"><span>시작</span><input type="text" data-dp="date" data-dp-range="ntf" data-ntfs-from value="${esc(it.from || "")}"></label><label class="ntf-f"><span>끝 (비우면 하루)</span><input type="text" data-dp="date" data-dp-range="ntf" data-dp-end data-ntfs-to value="${esc(it.to || "")}"></label></div>
      <div class="ntf-f" data-ntfs-dowbox ${it.kind === "dow" ? "" : "hidden"}><span>요일</span><div class="ntf-chips">${[1, 2, 3, 4, 5, 6, 0].map((n) => `<button type="button" data-ntfs-dow="${n}" class="${(it.dows || []).map(Number).includes(n) ? "on" : ""}">${NTF_DOW[n]}</button>`).join("")}</div></div>`}
      <div class="ntf-f"><span>대상</span><div class="ntf-chips">${["*", ...fcs].map((x) => `<button type="button" data-ntfs-tg="${esc(x)}" class="${tg.includes(x) ? "on" : ""}">${x === "*" ? (sys ? "해당 FC 모두" : "전체") : esc(x)}</button>`).join("")}</div></div>
      <label class="ntf-f"><span>확인하면 이동할 페이지</span><select data-ntfs-go>${NTF_GO.map(([k, l]) => `<option value="${k}" ${(it.go || "") === k ? "selected" : ""}>${l}</option>`).join("")}</select></label>
      <p class="hint">같은 날 알림이 여러 개면 한 창에 모두 보이고, 이동은 시스템 알림 → 목록 위쪽 순입니다.</p>
      <button type="button" class="ntf-save" data-ntfs-save>저장</button></div>`;
    document.body.appendChild(box);
  }
  async function ntfCommit(next, msg) {
    try { await ntfSave(next); NTFA.doc = next; toast(msg); } catch (e) { handleError(e); }
    renderNtfAdmin();
  }
  async function ntfSaveEdit(btn) {
    const box = $("ntfSheet"), it = { ...NTFA.edit };
    it.text = box.querySelector("[data-ntfs-text]").value.trim();
    if (!it.text) return toast("알림 내용을 입력해 주세요.");
    if (it.kind !== "system") {
      it.kind = (box.querySelector("[data-ntfs-kind].on") || {}).dataset ? box.querySelector("[data-ntfs-kind].on").dataset.ntfsKind : "date";
      if (it.kind === "date") {
        it.from = box.querySelector("[data-ntfs-from]").value; it.to = box.querySelector("[data-ntfs-to]").value || ""; delete it.dows;
        if (!it.from) return toast("시작 날짜를 넣어 주세요.");
        if (it.to && it.to < it.from) return toast("끝 날짜가 시작 날짜보다 앞입니다.");
      } else {
        it.dows = [...box.querySelectorAll("[data-ntfs-dow].on")].map((b) => +b.dataset.ntfsDow); delete it.from; delete it.to;
        if (!it.dows.length) return toast("요일을 하나 이상 골라 주세요.");
      }
    }
    const tg = [...box.querySelectorAll("[data-ntfs-tg].on")].map((b) => b.dataset.ntfsTg);
    it.target = !tg.length || tg.includes("*") ? ["*"] : tg;
    it.go = box.querySelector("[data-ntfs-go]").value;
    const isNew = !!it._new; delete it._new; if (it.on == null) it.on = true;
    btn.disabled = true; btn.textContent = "저장 중";
    const doc = NTFA.doc, items = isNew ? [it, ...doc.items] : doc.items.map((x) => (x.id === it.id ? it : x)); // 새 알림은 맨 위(이동 우선)
    NTFA.edit = null; box.remove();
    await ntfCommit({ ...doc, items }, isNew ? "알림을 추가했습니다." : "알림을 고쳤습니다.");
  }
  document.addEventListener("click", async (ev) => {
    const t = ev.target.closest && ev.target.closest("[data-ntf-open], [data-ntf-back], [data-ntf-sys], [data-ntf-f], [data-ntf-add], [data-ntf-on], [data-ntf-edit], [data-ntf-del], [data-ntfs-x], [data-ntfs-kind], [data-ntfs-dow], [data-ntfs-tg], [data-ntfs-save]");
    if (!t) { const sh = ev.target.id === "ntfSheet" ? ev.target : null; if (sh) { NTFA.edit = null; sh.remove(); } return; }
    const d = t.dataset;
    if (d.ntfOpen != null) return openNtfAdmin();
    if (d.ntfBack != null) { renderAdmin(); window.scrollTo(0, 0); return; }
    if (d.ntfF) { NTFA.f = d.ntfF; return renderNtfAdmin(); }
    if (d.ntfSys != null) return ntfCommit({ ...NTFA.doc, system: !NTFA.doc.system }, NTFA.doc.system ? "시스템 알림을 모두 껐습니다." : "시스템 알림을 켰습니다.");
    if (d.ntfOn) return ntfCommit({ ...NTFA.doc, items: NTFA.doc.items.map((x) => (x.id === d.ntfOn ? { ...x, on: x.on === false } : x)) }, "알림을 " + ((NTFA.doc.items.find((x) => x.id === d.ntfOn) || {}).on === false ? "켰습니다." : "껐습니다."));
    if (d.ntfAdd != null) { NTFA.edit = { _new: true, id: "n" + Date.now().toString(36), kind: "date", text: "", from: ntfDay(), to: "", target: ["*"], go: "", on: true }; return ntfSheet(); }
    if (d.ntfEdit) { const it = NTFA.doc.items.find((x) => x.id === d.ntfEdit); if (it) { NTFA.edit = { ...it }; ntfSheet(); } return; }
    if (d.ntfDel) { const it = NTFA.doc.items.find((x) => x.id === d.ntfDel); if (!it || !confirm(`이 알림을 삭제할까요?\n\n${it.text}`)) return; return ntfCommit({ ...NTFA.doc, items: NTFA.doc.items.filter((x) => x.id !== d.ntfDel) }, "알림을 삭제했습니다."); }
    if (d.ntfsX != null) { NTFA.edit = null; const b = $("ntfSheet"); if (b) b.remove(); return; }
    if (d.ntfsKind) { const box = $("ntfSheet"); box.querySelectorAll("[data-ntfs-kind]").forEach((b) => b.classList.toggle("on", b === t)); box.querySelector("[data-ntfs-datebox]").hidden = d.ntfsKind !== "date"; box.querySelector("[data-ntfs-dowbox]").hidden = d.ntfsKind !== "dow"; return; }
    if (d.ntfsDow != null) { t.classList.toggle("on"); return; }
    if (d.ntfsTg != null) { const box = $("ntfSheet"), all = d.ntfsTg === "*"; if (all) box.querySelectorAll("[data-ntfs-tg]").forEach((b) => b.classList.toggle("on", b === t)); else { t.classList.toggle("on"); box.querySelector('[data-ntfs-tg="*"]').classList.toggle("on", !box.querySelector("[data-ntfs-tg].on:not([data-ntfs-tg='*'])")); } return; }
    if (d.ntfsSave != null) return ntfSaveEdit(t);
  });

  // ---------- v568(2026-11-24 사용자 "그대로 코딩 진행"): 설정 맨 위 '다양한 기능' 상자 + [주소록 관리] ----------
  //   로그인한 사람 본인 구글 주소록(People API)에 이달이관 고객을 'YYYYMM이관' 라벨로 넣고(번호가 같은 연락처는 라벨·빈 칸만), 앱이 넣은 연락처('신한FC' 라벨) 중 고른 것만 지움.
  //   권한(contacts)은 [주소록 관리]에서 처음 넣기·정리할 때만 따로 요청, 출입증은 이 앱 전용 키(shlife.gcontacts). 연락처 값은 화면(메모리)과 구글에만 둠. 안: 작업데이터\20261124 - 주소록 연동 안.md
  const CT_SCOPE = "https://www.googleapis.com/auth/contacts", CT_KEY = "shlife.gcontacts", CT_API = "https://people.googleapis.com/v1/", CT_APP = "신한FC", CT_IDT = "신한FC고객ID";
  const featCardHtml = () => `<div class="card feat-card"><h3 class="sec-title">다양한 기능</h3><div class="feat-btns"><button type="button" class="secondary" data-feat="contacts">주소록 관리</button></div></div>`;
  document.addEventListener("click", (ev) => { const b = ev.target.closest && ev.target.closest("[data-feat]"); if (b && b.dataset.feat === "contacts") openContacts(); });
  function ctStored() { try { const t = JSON.parse(localStorage.getItem(CT_KEY) || "null"); if (t && t.exp > Date.now() + 60000) return t.token; } catch (e) { /* 무시 */ } return ""; }
  function ctRequest() {
    return new Promise((resolve, reject) => {
      const g = window.google && window.google.accounts && window.google.accounts.oauth2;
      if (!g) return reject(new Error("구글 로그인 모듈을 불러오지 못했습니다. 잠시 후 다시 시도하세요."));
      const client = g.initTokenClient({
        client_id: GCFG.clientId, scope: CT_SCOPE, hint: gAccount() || undefined,
        callback: (r) => {
          if (r.error) return reject(new Error("구글 연결 실패: " + r.error));
          if (!g.hasGrantedAllScopes(r, CT_SCOPE)) return reject(new Error("주소록 권한을 허용해야 넣을 수 있습니다."));
          try { localStorage.setItem(CT_KEY, JSON.stringify({ token: r.access_token, exp: Date.now() + (r.expires_in || 3600) * 1000 })); } catch (e) { /* 이번엔 사용 */ }
          resolve(r.access_token);
        },
        error_callback: (e) => reject(new Error(e && e.type === "popup_closed" ? "구글 로그인 창이 닫혔습니다." : "구글 로그인 창을 열 수 없습니다. 팝업 차단을 확인하세요."))
      });
      client.requestAccessToken({ prompt: "" });
    });
  }
  async function ctApi(path, opt) {
    const t = ctStored() || (await ctRequest());
    const o = opt || {};
    const res = await fetch(CT_API + path, { method: o.method || "GET", headers: { Authorization: "Bearer " + t, ...(o.body ? { "Content-Type": "application/json" } : {}) }, body: o.body ? JSON.stringify(o.body) : undefined });
    if (res.status === 401) { try { localStorage.removeItem(CT_KEY); } catch (e) { /* 무시 */ } throw new Error("주소록 연결 시간이 지났습니다. 다시 눌러 주세요."); }
    if (res.status === 403) { const tx = await res.text().catch(() => ""); throw new Error(/SERVICE_DISABLED|has not been used|disabled/i.test(tx) ? "구글 콘솔에서 People API를 켜야 합니다(관리자 할 일)." : "주소록 권한이 없습니다(403)."); }
    if (!res.ok) throw new Error("주소록 " + res.status + ": " + (await res.text().catch(() => "")).slice(0, 120));
    return res.status === 204 ? {} : res.json();
  }
  const ctDigits = (v) => { let d = String(v || "").replace(/\D/g, ""); if (d.startsWith("82")) d = "0" + d.slice(2); return d; };
  const CT = { fc: "", ym: "", rows: [], people: [], groups: [], mode: "", sel: new Set(), busy: false, msg: "" };
  async function ctLoadAll() { // 내 주소록 전체(번호·라벨·고객ID·메모·주소)와 라벨 목록
    let people = [], pt = "";
    do { const d = await ctApi("people/me/connections?pageSize=1000&personFields=names,phoneNumbers,memberships,externalIds,biographies,addresses" + (pt ? "&pageToken=" + encodeURIComponent(pt) : "")); people = people.concat(d.connections || []); pt = d.nextPageToken || ""; } while (pt);
    const g = await ctApi("contactGroups?pageSize=1000");
    CT.people = people; CT.groups = (g.contactGroups || []).filter((x) => x.groupType === "USER_CONTACT_GROUP" || !x.groupType);
  }
  const ctGroupName = (rn) => (CT.groups.find((g) => g.resourceName === rn) || {}).name || "";
  const ctLabels = (p) => (p.memberships || []).map((m) => m.contactGroupMembership && ctGroupName(m.contactGroupMembership.contactGroupResourceName)).filter(Boolean);
  const ctCustId = (p) => ((p.externalIds || []).find((x) => x.type === CT_IDT) || {}).value || "";
  const ctIsApp = (p) => ctLabels(p).includes(CT_APP) || !!ctCustId(p);
  async function ctGroup(name) { // 라벨 찾기·없으면 만들기 → resourceName
    const g = CT.groups.find((x) => x.name === name); if (g) return g.resourceName;
    const r = await ctApi("contactGroups", { method: "POST", body: { contactGroup: { name } } });
    CT.groups.push(r); return r.resourceName;
  }
  async function openContacts() {
    show("adminView");
    const body = $("adminBody");
    body.innerHTML = `<div class="card ct-wrap"><p class="hint">불러오는 중…</p></div>`;
    try {
      const data = await loadSalesData(), myFc = await myFcName();
      const fcs = [...new Set(data.lists.map((e) => e.fc).filter(Boolean))].sort((a, b) => a.localeCompare(b, "ko"));
      CT.fcs = fcs; CT.fc = CT.fc && fcs.includes(CT.fc) ? CT.fc : fcs.includes(myFc) ? myFc : fcs[0] || ""; CT.ym = ymNow(); CT.data = data;
      CT.mode = ""; CT.sel = new Set(); CT.msg = ""; ctRows(); renderContacts(); window.scrollTo(0, 0);
    } catch (e) { body.innerHTML = ""; handleError(e); }
  }
  function ctRows() { // 이달이관(이 FC, 보관·삭제 아닌 줄) → 고객 한 명씩
    const seen = new Set();
    CT.rows = transferRows(CT.data, CT.ym, false).filter((r) => r.srcFc === CT.fc && !r.archDay && !r.del).filter((r) => { const k = r.custId || r.holder; if (seen.has(k)) return false; seen.add(k); return true; })
      .map((r) => ({ custId: r.custId, name: r.holder, phone: firstPhone(r.phoneText || r.phone || ""), addr: r.addr || "", note: `${CT_APP} 이관 ${CT.ym.slice(0, 4)}-${CT.ym.slice(4)} · ${r.product || ""}${r.fee ? " · " + wonText(r.fee) : ""} · 담당 ${CT.fc}` }));
  }
  const ctLabel = () => CT.ym + "이관";
  function renderContacts() {
    const body = $("adminBody"), acc = gAccount();
    const fcSel = CT.fcs.length > 1 ? `<select id="ctFc" aria-label="FC">${CT.fcs.map((f) => `<option${f === CT.fc ? " selected" : ""}>${esc(f)}</option>`).join("")}</select>` : `<b>${esc(CT.fc)}</b>`;
    let list = "";
    if (CT.mode === "add") {
      const items = CT.rows.map((r, i) => { const d = ctDigits(r.phone), hit = d ? CT.people.find((p) => (p.phoneNumbers || []).some((x) => ctDigits(x.value) === d)) : null, has = hit && ctLabels(hit).includes(ctLabel()); return { r, i, d, hit, st: !d ? "번호 없음" : has ? "이미 있음" : hit ? "라벨만 더함" : "새로 넣음" }; });
      CT.addItems = items;
      const n = { new: items.filter((x) => x.st === "새로 넣음").length, lab: items.filter((x) => x.st === "라벨만 더함").length, no: items.filter((x) => x.st === "번호 없음").length, has: items.filter((x) => x.st === "이미 있음").length };
      list = `<div class="ct-sum">라벨 <b>${esc(ctLabel())}</b> · 새로 ${n.new} · 라벨만 더함 ${n.lab}${n.has ? ` · 이미 있음 ${n.has}` : ""}${n.no ? ` · 번호 없음 ${n.no}` : ""}</div>
        <div class="ct-list">${items.map((x) => `<label class="ct-it${x.st === "번호 없음" || x.st === "이미 있음" ? " off" : ""}"><input type="checkbox" class="cb" data-ct-sel="${x.i}" ${CT.sel.has(x.i) ? "checked" : ""} ${x.st === "번호 없음" || x.st === "이미 있음" ? "disabled" : ""}><span class="ct-nm">${esc(x.r.name)}<small>${x.r.phone ? esc(String(x.r.phone).replace(/^(\d{3})(\d{3,4})(\d{4})$/, "$1-$2-$3")) : "번호가 없어 넣을 수 없음"}${x.hit ? " · 이미 있는 연락처" : ""}</small></span><span class="ct-tg ${x.st === "새로 넣음" ? "new" : x.st === "번호 없음" ? "no" : ""}">${x.st}</span></label>`).join("") || '<p class="hint">이번 달 이관 고객이 없습니다.</p>'}</div>
        <p class="hint ct-note">이름·휴대전화·주소와 메모 한 줄(상품·보험료·담당)을 넣습니다. 이미 있는 사람은 새로 만들지 않고 라벨과 빈 칸만 채웁니다.</p>
        <button type="button" class="ntf-save" data-ct-go="add" ${CT.sel.size && !CT.busy ? "" : "disabled"}>${CT.busy ? esc(CT.busy) : CT.sel.size + "명 넣기"}</button>`;
    } else if (CT.mode === "clean") {
      const keep = new Set(CT.rows.map((r) => r.custId).filter(Boolean));
      const items = CT.people.filter(ctIsApp).map((p, i) => { const id = ctCustId(p), labs = ctLabels(p).filter((l) => /^\d{6}/.test(l)); return { p, i, id, labs, cand: !id || !keep.has(id) }; }).sort((a, b) => (b.cand - a.cand) || String(a.labs[0] || "").localeCompare(String(b.labs[0] || "")));
      CT.cleanItems = items;
      list = `<div class="ct-sum">앱이 넣은 연락처 ${items.length}명 · 이달목표에서 빠진 고객 ${items.filter((x) => x.cand).length}명을 먼저 골라 두었습니다</div>
        <div class="ct-list">${items.map((x, k) => `<label class="ct-it"><input type="checkbox" class="cb" data-ct-sel="${k}" ${CT.sel.has(k) ? "checked" : ""}><span class="ct-nm">${esc(((x.p.names || [])[0] || {}).displayName || "(이름 없음)")}<small>${esc(x.labs.join(" · ") || CT_APP)}${x.cand ? " · 이달목표에 없음" : " · 이달목표에 있음"}</small></span><span class="ct-tg ${x.cand ? "del" : ""}">${x.cand ? "지울 후보" : "남김"}</span></label>`).join("") || '<p class="hint">앱이 넣은 연락처가 없습니다.</p>'}</div>
        <button type="button" class="ntf-save ct-del" data-ct-go="clean" ${CT.sel.size && !CT.busy ? "" : "disabled"}>${CT.busy ? esc(CT.busy) : CT.sel.size + "명 지우기"}</button>`;
    }
    body.innerHTML = `<div class="card ct-wrap"><div class="ct-head"><button type="button" class="ev-edit org-back-btn" data-ct-back>‹ 설정</button><h3 class="sec-title">주소록 관리</h3></div>
      <p class="hint ct-acc">내 구글 주소록${acc ? `(${esc(acc)})` : ""}에 넣고 지웁니다. 앱이 넣은 연락처만 지울 수 있습니다.</p>
      <div class="ct-bar"><span>이관 · ${fcSel} · ${esc(CT.ym.slice(0, 4))}-${esc(CT.ym.slice(4))} · ${CT.rows.length}명</span></div><div class="ct-acts"><button type="button" class="ct-b${CT.mode === "add" ? " on" : ""}" data-ct-mode="add">주소록 넣기</button><button type="button" class="ct-b${CT.mode === "clean" ? " on" : ""}" data-ct-mode="clean">주소록 정리</button></div>
      ${CT.msg ? `<div class="ct-msg">${CT.msg}</div>` : ""}${list}</div>`;
  }
  async function ctMode(m) {
    CT.mode = m; CT.sel = new Set(); CT.msg = ""; CT.busy = "주소록을 읽는 중…"; renderContacts();
    try { await ctLoadAll(); }
    catch (e) { CT.mode = ""; CT.busy = false; CT.msg = esc(e.message || String(e)) + ` <button type="button" class="ev-edit" data-ct-mode="${m}">다시 연결</button>`; renderContacts(); return; }
    CT.busy = false;
    if (m === "add") { renderContacts(); (CT.addItems || []).forEach((x) => { if (x.st === "새로 넣음" || x.st === "라벨만 더함") CT.sel.add(x.i); }); }
    else { renderContacts(); (CT.cleanItems || []).forEach((x, k) => { if (x.cand) CT.sel.add(k); }); }
    renderContacts();
  }
  async function ctAdd() {
    const items = (CT.addItems || []).filter((x) => CT.sel.has(x.i)); if (!items.length) return;
    let ok = 0, lab = 0; const fail = [];
    try {
      const gLab = await ctGroup(ctLabel()), gApp = await ctGroup(CT_APP);
      for (let k = 0; k < items.length; k++) { // 같은 사용자 쓰기는 차례로(People API 안내)
        const x = items[k]; CT.busy = `넣는 중 ${k + 1}/${items.length}`; renderContacts();
        try {
          if (x.hit) {
            const add = [gLab, gApp].filter((g) => !(x.hit.memberships || []).some((m) => m.contactGroupMembership && m.contactGroupMembership.contactGroupResourceName === g));
            for (const g of add) await ctApi(g + "/members:modify", { method: "POST", body: { resourceNamesToAdd: [x.hit.resourceName] } });
            const upd = {}, fields = [];
            if (!(x.hit.addresses || []).length && x.r.addr) { upd.addresses = [{ formattedValue: x.r.addr, type: "home" }]; fields.push("addresses"); }
            if (!(x.hit.biographies || []).length) { upd.biographies = [{ value: x.r.note, contentType: "TEXT_PLAIN" }]; fields.push("biographies"); }
            if (!ctCustId(x.hit) && x.r.custId) { upd.externalIds = [...(x.hit.externalIds || []), { value: x.r.custId, type: CT_IDT }]; fields.push("externalIds"); }
            if (fields.length) await ctApi(x.hit.resourceName + ":updateContact?updatePersonFields=" + fields.join(","), { method: "PATCH", body: { etag: x.hit.etag, ...upd } });
            lab++;
          } else {
            await ctApi("people:createContact?personFields=names", { method: "POST", body: { names: [{ unstructuredName: x.r.name }], phoneNumbers: [{ value: x.r.phone, type: "mobile" }], ...(x.r.addr ? { addresses: [{ formattedValue: x.r.addr, type: "home" }] } : {}), biographies: [{ value: x.r.note, contentType: "TEXT_PLAIN" }], ...(x.r.custId ? { externalIds: [{ value: x.r.custId, type: CT_IDT }] } : {}), memberships: [{ contactGroupMembership: { contactGroupResourceName: gLab } }, { contactGroupMembership: { contactGroupResourceName: gApp } }] } });
            ok++;
          }
        } catch (e) { fail.push(x.r.name + "(" + String(e.message || e).slice(0, 40) + ")"); if (/People API|시간이 지났|권한/.test(e.message)) break; }
      }
    } catch (e) { fail.push(String(e.message || e)); }
    CT.busy = false; CT.mode = ""; CT.sel = new Set();
    CT.msg = `넣음 ${ok} · 라벨 더함 ${lab}${fail.length ? ` · 실패 ${fail.length}: ${esc(fail.join(", "))}` : ""}`;
    renderContacts(); toast(fail.length ? "일부를 넣지 못했습니다." : "주소록에 넣었습니다.");
  }
  async function ctClean() {
    const items = (CT.cleanItems || []).filter((x, k) => CT.sel.has(k)); if (!items.length) return;
    if (!confirm(`${items.length}명을 주소록에서 지울까요?\n\n구글 주소록에서 영구 삭제되며 되돌릴 수 없습니다. 고객 정보(앱 데이터)는 지워지지 않습니다.`)) return;
    let ok = 0; const fail = [];
    for (let k = 0; k < items.length; k++) {
      CT.busy = `지우는 중 ${k + 1}/${items.length}`; renderContacts();
      try { await ctApi(items[k].p.resourceName + ":deleteContact", { method: "DELETE" }); ok++; } catch (e) { fail.push((((items[k].p.names || [])[0] || {}).displayName || "") + "(" + String(e.message || e).slice(0, 40) + ")"); }
    }
    CT.busy = false; CT.mode = ""; CT.sel = new Set();
    CT.msg = `지움 ${ok}${fail.length ? ` · 실패 ${fail.length}: ${esc(fail.join(", "))}` : ""}`;
    renderContacts(); toast(fail.length ? "일부를 지우지 못했습니다." : "주소록에서 지웠습니다.");
  }
  document.addEventListener("click", (ev) => {
    const t = ev.target.closest && ev.target.closest("[data-ct-back],[data-ct-mode],[data-ct-go]"); if (!t || CT.busy) return;
    if (t.dataset.ctBack != null) { if (isAdmin()) openAdmin(); else openMySettings(); return; }
    if (t.dataset.ctMode) return ctMode(t.dataset.ctMode);
    if (t.dataset.ctGo === "add") return ctAdd();
    if (t.dataset.ctGo === "clean") return ctClean();
  });
  document.addEventListener("change", (ev) => {
    const t = ev.target;
    if (t && t.id === "ctFc") { CT.fc = t.value; CT.mode = ""; CT.sel = new Set(); CT.msg = ""; ctRows(); renderContacts(); return; }
    const s = t && t.closest && t.closest("[data-ct-sel]"); if (!s) return;
    const i = +s.dataset.ctSel; s.checked ? CT.sel.add(i) : CT.sel.delete(i);
    const b = document.querySelector("[data-ct-go]"); if (b) { b.disabled = !CT.sel.size; b.textContent = CT.sel.size + (CT.mode === "add" ? "명 넣기" : "명 지우기"); }
  });
  function startApp() {
    isRegistered().then((ok) => {
      if (ok) return startApp0();
      show("loginView");
      const msg = $("lmsg"); if (msg) msg.textContent = `${me()} 계정은 아직 신한FC 로그인이 승인되지 않았습니다. 관리자에게 승인을 요청해 주세요.`;
      const ob = $("loginOtherBtn"); if (ob) ob.hidden = false;
    }, (e) => handleError(e));
  }
  function startApp0() {
    appStarted = true;
    syncStart(); // v550: 데이터 수집 중 잠금 확인(1분마다)
    const goReq = GO0; GO0 = null;
    if (goReq) { // 말만해요에서 넘어옴: 첫 화면은 고객, 이어서 검색
      schedMode = "day"; schedDate = dayStart(new Date());
      resetSearch();
      setTimeout(() => window.blue10kpGo(goReq), 0);
    } else if (!GD && sessionStorage.getItem("openAdmin") && isAdmin()) {
      sessionStorage.removeItem("openAdmin");
      openAdmin();
    } else if (SALES_BOOT) { // 새로 로드됐을 때 세일즈 화면을 보던 중이었으면 그 탭·월·FC·필터·위치로
      const s = SALES_BOOT; SALES_BOOT = null;
      Object.assign(SALES, { tab: s.tab || "transfer", ym: s.ym || "", fc: s.fc || "", dct: "", dcls: "", dact: true, /* v94: 따행 필터는 이어보기에서도 기본값 */ df: { send: !!(s.df || {}).send, resp: !!(s.df || {}).resp, coupon: !!(s.df || {}).coupon }, open: new Set(s.open || []), resume: s.a || null });
      loadSettings().then(() => openSales());
    } else if (SCHED_BOOT) { // 새로 로드됐을 때 일정 화면을 보던 중이었으면 그 날짜·보기·위치로
      const s = SCHED_BOOT; SCHED_BOOT = null;
      schedMode = ["day", "week", "month"].includes(s.mode) ? s.mode : "day";
      schedDate = s.d ? dayStart(new Date(s.d)) : dayStart(new Date());
      if (isNaN(+schedDate)) schedDate = dayStart(new Date());
      if (schedMode === "day" && !sameDay(schedDate, new Date())) { schedDate = dayStart(new Date()); s.y = 0; } // Today는 새로 열면 항상 오늘 (다른 날을 보던 중이었어도)
      schedResumeY = s.y > 0 ? s.y : null;
      loadSettings().then(() => openSchedule());
    } else {
      schedMode = "day"; schedDate = dayStart(new Date());
      const sp = START_FIXED || readStartPref().page; // 설정 > 시작 화면(v520: 대시보드 고정)
      if (sp === "dash") openDash();
      else if (sp === "schedule") loadSettings().then(() => openSchedule());
      else if (sp === "sales") loadSettings().then(() => openSales());
      else {
        resetSearch(); // 첫 화면: 고객 (일정은 하단 메뉴에서 열 때 불러옴 — 캘린더 호출이 첫 화면을 늦추지 않게)
        $("searchStatus").textContent = "고객 목록을 불러오는 중…";
      }
    }
    setTimeout(() => { if (!goReq) checkNewCustomers(); }, 2500); // 로그인 때 새 고객 확인
    setTimeout(() => { if (!goReq) ntfCheck(); }, 1200); // v557: 로그인 직후 중요 알림(그날 한 번)
    setTimeout(wakeStart, 4000); // '몰리야' 듣기(설정에서 켠 기기만)
    setTimeout(() => { loadMyPrefs().catch(() => {}); }, 1500); // 내 설정(다크/라이트 포함)을 일찍 읽어 둠
    if (!isAdmin()) syncStartShared();
    if (readStartPref().ai && !EMBED) setTimeout(aiOpen, 400); // 설정 > 시작 화면 > AI 도우미 바로 열기
    setTimeout(askContactNote, 1500); // 전화·문자 중에 앱이 닫혔다가 다시 열린 경우
    loadSettings(); // 관리자 설정 (회사 주소, 데이터 절약 모드 등)
    if (!isAdmin()) loadMyPrefs(); // 일반 사용자 설정 (캘린더·데이터 절약)
    loadLists().then((recs) => { syncMalIndex(recs).catch(() => {}); renderFcOpts(recs).catch(() => {}); showScope(); if ($("searchStatus").textContent === "고객 목록을 불러오는 중…") $("searchStatus").textContent = ""; }).catch(handleError); // 미리 불러오기
  }
})();
