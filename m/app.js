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
    if (!jsonDocP.has(id)) jsonDocP.set(id, stGet0(id).then((buf) => JSON.parse(new TextDecoder().decode(buf).replace(/^\uFEFF/, ""))).catch((e) => { jsonDocP.delete(id); throw e; }));
    return jsonDocP.get(id);
  }
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
      jsonInfo.set(i.id, { name: i.name, meta: null, ver: null, mod: i.lastModifiedDateTime || "", parent: id });
      virt.push({ ...i, name: "99999999" + (i.name === "data.json" ? " - JSON" : "") + suffix });
    });
    return virt.length ? virt.concat(items.filter((i) => !(i.file && JSON_FILES[i.name]))) : items;
  }
  async function stGet(id) {
    const info = jsonInfo.get(id);
    if (!info) return stGet0(id);
    const m = await gdFetch(GDAPI + "/files/" + encodeURIComponent(id) + "?fields=version&supportsAllDrives=true");
    const doc = JSON.parse(new TextDecoder().decode(await stGet0(id)).replace(/^\uFEFF/, ""));
    info.ver = m.version; info.meta = { fc: doc.fc, source: doc.source };
    return XLSX.write(jsonToBook(doc), { type: "array", bookType: "xlsx" });
  }
  async function stPut(id, data, mime) {
    const info = jsonInfo.get(id);
    if (!info) return stPut0(id, data, mime);
    const m = await gdFetch(GDAPI + "/files/" + encodeURIComponent(id) + "?fields=version&supportsAllDrives=true");
    if (info.ver != null && String(m.version) !== String(info.ver)) throw new Error("다른 곳에서 먼저 수정되었습니다. 새로고침한 뒤 다시 시도해 주세요.");
    // 데이터 반영(수집) 중에는 FC 폴더에 .반영중.lock 이 있다 → 저장 직전에만 확인해 막는다(읽기는 그대로)
    if (info.name === "data.json" && info.parent && (await gdFindIn(info.parent, ".반영중.lock"))) throw new Error("데이터 반영 중이라 저장하지 못했습니다. 잠시 후 다시 시도해 주세요.");
    const doc = bookToDoc(XLSX.read(data, { type: "array", cellNF: true }), info.meta || {});
    const out = await stPut0(id, new Blob([JSON.stringify(doc)], { type: "application/json" }), "application/json");
    jsonDocP.delete(id);
    if (out && out.lastModifiedDateTime) { info.mod = out.lastModifiedDateTime; renderDataStamp(); } // 저장하면 데이터 시각도 바로 갱신
    const m2 = await gdFetch(GDAPI + "/files/" + encodeURIComponent(id) + "?fields=version&supportsAllDrives=true");
    info.ver = m2.version;
    return out;
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

  // ---------- 화면 전환 ----------
  // 오른손 모드: 설정에서 켜면(이 폰에만 저장) 하단 메뉴·탭·선택 버튼을 오른쪽으로 모음 (모양은 design.css의 .rh)
  const RH_KEY = "shlife.rightHand";
  const rhOn = () => { try { return localStorage.getItem(RH_KEY) === "1"; } catch (e) { return false; } };
  const rhApply = () => document.documentElement.classList.toggle("rh", rhOn());
  rhApply();
  const rhOptHtml = () => `<div class="fld rh-fld">
        <label class="rh-opt"><span>오른손 모드</span><input type="checkbox" class="cb" id="rhMode" ${rhOn() ? "checked" : ""}></label>
      </div>`;
  const rhCardHtml = () => `<div class="card"><h3 class="sec-title">화면</h3>${rhOptHtml()}</div>`;
  document.addEventListener("change", (ev) => {
    if (!ev.target || ev.target.id !== "rhMode") return;
    try { localStorage.setItem(RH_KEY, ev.target.checked ? "1" : "0"); } catch (e) { /* 저장 못 해도 이번엔 적용 */ }
    rhApply();
    toast(ev.target.checked ? "오른손 모드를 켰습니다." : "오른손 모드를 껐습니다.");
  });
  const views = ["loginView", "scheduleView", "searchView", "customerView", "salesView", "adminView"];
  const APP_NAME = "신한FC m";
  const TITLES = { scheduleView: "일정", searchView: "고객", customerView: "고객 정보", salesView: "영업", adminView: "설정" };
  const SECTION = { scheduleView: "schedule", searchView: "customer", customerView: "customer", salesView: "sales", adminView: "admin" };
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
    // 관리자: '관리자' 메뉴 / 일반 사용자: '설정' 메뉴 (내 캘린더 고르기)
    $("navAdmin").hidden = !account();
    $("navAdminLabel").textContent = "설정";
    if (view === "searchView") $("pageSearch").hidden = $("results").children.length > 0; // 보이는 고객 목록이 없으면 화면 안 검색창
    if (view !== "salesView" && view !== "customerView" && view !== "loginView") salesClearState(); // 영업 화면을 떠나면 이어보기 상태 삭제
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
  // 영업 화면이 매번 권한 파일을 다시 읽지 않도록 'FC 이름(names)'만 세션 캐시. 권한 파일을 읽거나 쓸 때마다 갱신 (관리자 화면은 열 때 최신을 읽음). 접근 FC 판단(myAllowedFcs)은 캐시를 쓰지 않음
  let permNamesCache = null, permPrefixCache = null; // permPrefixCache: { FC이름: 일정 제목 앞 글자 }
  let salesExtraAt = 0;
  // 영업 화면을 열 때(1분에 한 번) 권한 파일을 다시 읽어, 관리자가 방금 바꾼 '추가 FC'도 앱을 다시 열지 않고 반영
  async function refreshSalesExtra() { if (Date.now() - salesExtraAt < 60000) return; salesExtraAt = Date.now(); try { await readPerms(); } catch (e) { /* 이전 값 유지 */ } }
  let salesExtra = []; // 영업(이달이관·따행) 화면의 FC 선택창에 추가로 보이는 FC (관리자 화면에서 계정마다 지정)
  async function writePerms(perms) {
    await resolveAppDir();
    await stWriteJsonIn(appDir, PERM_FILE, perms);
    permNamesCache = { ...(perms.names || {}) };
    permPrefixCache = { ...(perms.prefixes || {}) };
  }
  const me = () => String((account() || {}).username || "").toLowerCase();

  async function readPerms() {
    await resolveAppDir();
    const p = await stReadJsonIn(appDir, PERM_FILE);
    const out = { users: (p && p.users) || {}, fcFolders: (p && p.fcFolders) || {}, names: (p && p.names) || {}, prefixes: (p && p.prefixes && typeof p.prefixes === "object") ? p.prefixes : {}, sensitive: Array.isArray(p && p.sensitive) ? p.sensitive : [], salesFcs: (p && p.salesFcs && typeof p.salesFcs === "object") ? p.salesFcs : {}, starts: (p && p.starts && typeof p.starts === "object") ? p.starts : {} }; // salesFcs: { 이메일: [영업 화면에서 추가로 볼 수 있는 FC 이름] } / sensitive: 민감정보 접근을 허용한 이메일 목록 (다시 저장해도 지워지지 않게 읽어 둠)
    permNamesCache = { ...out.names };
    permPrefixCache = { ...out.prefixes };
    try { const k = "shlife.startAdmin." + me(), v = out.starts[me()]; if (v) localStorage.setItem(k, JSON.stringify(v)); else localStorage.removeItem(k); } catch (e) { /* 무시 */ } // 관리자가 정한 시작 화면 (다음 실행부터)
    salesExtra = Array.isArray(out.salesFcs[me()]) ? out.salesFcs[me()].slice() : [];
    salesExtraAt = Date.now();
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
  function scanBase() { return scanCache || (scanCache = scanRun(true).catch((e) => { scanCache = null; throw e; })); }
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
  async function readBook(id) {
    if (!bookCache[id]) {
      bookCache[id] = jsonInfo.has(id)
        ? readJsonDoc(id).then((doc) => { const inf = jsonInfo.get(id); if (inf && !inf.meta) inf.meta = { fc: doc.fc, source: doc.source }; return jsonToBook(doc); }) // JSON을 앱 구조(시트)로 바로 읽기
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
  const stageType = (row) => STAGE_TYPE[String(row["현재단계"] || "").trim()] || "가망";
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
  const contractLabel = (mt) => { const v = String(mt || "").trim(); return v === "타인모집" ? "이관" : /^타인모집/.test(v) ? "타인모집" : "계약"; };
  const isNewContractSheet = (head) => head.includes("계약자ID") && !head.includes("계약자");
  // </schema>
  // 앱 버전 표시 (화면 맨 아래 작은 글씨) — 새 버전이 모든 기기에 반영됐는지 확인용
  const APP_VER = (((document.querySelector('script[src*="app.js"]') || {}).src || "").match(/[?&]v=(\d+)/) || [])[1] || "";
  if (APP_VER) { const el = document.createElement("div"); el.className = "app-ver"; el.textContent = "v" + APP_VER; document.body.appendChild(el); }

  const pick = (row, names) => { for (const n of names) if (row[n] != null && String(row[n]).trim()) return String(row[n]).trim(); return ""; };
  const NAME_COLS = ["이름", "계약자", "고객명"];
  const BIRTH_COLS = ["생년월일", "생일"];
  const PHONE_COLS = ["연락처", "휴대폰", "핸드폰", "전화번호"];

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
          const name = pick(row, NAME_COLS);
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
    listCache = out;
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
  const gradeOf = (c) => { const g = (c.entries || []).map((e) => e.grade).find(Boolean) || ""; return GRADES.find((x) => x.id.toLowerCase() === g.toLowerCase()) ? g : ""; };
  const gradeClass = (g) => (g ? "grade-" + g.toLowerCase() : "");

  // entries의 원래 줄(파일·시트·행)에 values({열이름: 값})를 기록. 열이 없으면 머리글 끝에 새로 만듦
  async function writeEntryCells(entries, values) {
    entries = ownEntries(entries); // 담당 FC 행에만 저장
    if (entries.some((e) => isRO(e.fc))) throw new Error("보기 전용 권한이라 수정할 수 없습니다.");
    if (entries.some((e) => isRO(e.fc))) throw new Error("보기 전용 권한이라 수정할 수 없습니다.");
    const byFile = new Map();
    entries.filter((e) => e.src).forEach((e) => { if (!byFile.has(e.src.fileId)) byFile.set(e.src.fileId, []); byFile.get(e.src.fileId).push(e); });
    for (const [fileId, list] of byFile) {
      const wb = XLSX.read(await stGet(fileId), { type: "array", cellNF: true }); // 최신 파일 (cellNF: 숫자·날짜 서식 유지)
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
        for (const [k, v] of Object.entries(values)) {
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
          if (k === GRADE_COL) e.grade = String(v);
          if (PHONE_COLS.includes(k)) e.phone = String(v);
        }
      }
      await stPut(fileId, XLSX.write(wb, { type: "array", bookType: "xlsx", compression: true }),
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      delete bookCache[fileId];
    }
  }

  // ---------- 검색 ----------
  let mode = "all";
  const norm = (s) => String(s).replace(/\s+/g, "").toLowerCase();
  const custPart = (n) => n.split(" - ")[0]; // 폴더명 "고객명 - 담당자명"
  // 동명이인 표시를 뺀 이름: "이영숙(640927)" → "이영숙" (일정 제목·계약 매칭은 이 이름으로)
  const baseName = (n) => String(n || "").replace(/\(\d{6}\)\s*$/, "").trim();
  // ---- 새 구조: 고객ID·담당FC ----
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
  const LIST_PRIORITY = CFG.lists.map((l) => l.label);
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
  const OPT_LABEL = { contract: "본인", transfer: "이관", db: "DB", ddahaeng: "따행", plan: "가망" };
  function syncOpts() { document.querySelectorAll("#searchOpts [data-opt]").forEach((b) => b.classList.toggle("on", b.dataset.opt === "all" ? optSel.size === 0 : optSel.has(b.dataset.opt))); }
  const typeOfOpt = (id) => (CFG.lists.find((l) => l.id === id) || {}).label;
  // 고객 리스트의 FC 선택: 전체 / 각 FC. 기본은 내 FC(권한 파일의 내 이름), 고른 값은 이 폰에 기억. FC가 하나뿐이면 선택칸을 숨김
  const FC_ALL = "__all__";
  const listFcKey = () => "shlife.listFc." + me();
  async function renderFcOpts(records) {
    const box = $("fcOpts");
    if (!box) return FC_ALL;
    const fcs = [...new Set(records.map((r) => r.fc).filter(Boolean))].sort((a, b) => { const o = CFG.fcOrder || [], ia = o.indexOf(a), ib = o.indexOf(b); return (ia < 0 ? 999 : ia) - (ib < 0 ? 999 : ib) || a.localeCompare(b, "ko"); });
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
  let goKeepList = null; // 말만해요에서 넘어온 검색어: 한 명만 나와도 목록을 보여 줌
  async function search(q, force) {
    q = q.trim();
    const keepList = goKeepList !== null && goKeepList === q; goKeepList = null;
    const status = $("searchStatus");
    const list = $("results");
    list.innerHTML = "";
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
      const meets = await loadMeetingPlaces().catch(() => []);
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
      const hits = records.filter((r) => {
        if (fcPick !== FC_ALL && r.fc !== fcPick) return false;
        if (want && r.type !== want.label) return false;
        if (optTypes.length && !optTypes.some((id) => id === "ddahaeng" ? (r.type === typeOfOpt(id) || fieldOf(r, "따행대상") === "예") : r.type === typeOfOpt(id))) return false;
        if (!q) return true;
        if (polIds && (polIds.has(idOf(r)) || (!idOf(r) && polNames.includes(norm(baseName(r.name)) + "|" + r.fc)))) return true;
        if (norm(r.name).includes(nq)) return true;
        if (dq.length >= 3 && dq === q.replace(/[\s-]/g, "") && digits(r.phone).includes(dq)) return true;
        if (nq.length < 2) return false;
        return norm(entryAddr(r)).includes(nq) || norm(meetOf(r)).includes(nq);
      });
      // 같은 사람 묶기: 고객ID가 같으면 한 고객, ID가 없는 옛 행은 같은 이름 + 같은 생년월일
      const groups = groupCustomers(hits);
      // 고객 안의 시트는 우선순위 순 (계약 > 이관 > 따행 > DB > 가망)
      groups.forEach((c) => { c.entries.sort((a, b) => prio(a.type) - prio(b.type)); c.name = c.entries[0].name; });
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
      status.textContent = custs.length ? label + custs.length + "명 찾음" : (q ? "'" + q + "' " : "") + label + "고객이 없습니다.";
      $("pageSearch").hidden = custs.length > 0 || !!force; // 못 찾으면 화면 안 검색창으로 다시 검색
      if (custs.length === 1 && q && !keepList) return openCustomer(custs[0]);
      list.innerHTML = custs.map((c, i) => {
        const types = [...new Set(c.entries.map((e) => e.type))];
        const fcs = [...new Set(c.entries.map((e) => e.fc).filter(Boolean))];
        return `<li><button data-i="${i}"><span class="avatar sm">${esc(c.name[0])}</span>
          <span class="r-name"><span><span class="cname ${gradeClass(gradeOf(c))}">${esc(c.name)}</span>${resultInfo(c) ? ` <span class="rinfo">${esc(resultInfo(c))}</span>` : ""}</span>${fcs.length ? `<span class="r-sub">${esc(fcs.join(", "))}</span>` : ""}</span>
          ${churnTagHtml(churnMap.get(c))}${types.map((t) => `<span class="tag t-${esc(t)}">${esc(t)}</span>`).join("")}<span class="chev">›</span></button>${insLineHtml(c)}</li>`;
      }).join("");
      list.querySelectorAll("button").forEach((b) =>
        b.addEventListener("click", () => openCustomer(custs[+b.dataset.i])));
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
    const wb = XLSX.read(await stGet(file.id), { type: "array", cellNF: true }); // 서식 유지 (다시 저장할 때 #,##0·날짜 서식이 사라지지 않게)
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
    const out = XLSX.write(x.wb, { type: "array", bookType: "xlsx", compression: true });
    await stPut(x.file.id, out,
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
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
    if (!$("salesView").hidden) SALES.cardAnchor = salesAnchor(); // 영업 화면에서 열면 돌아올 때 이 위치로
    returnView = !$("scheduleView").hidden ? "scheduleView" : !$("salesView").hidden ? "salesView" : "searchView";
    show("customerView");
    $("custName").textContent = cust.name;
    editMode = false;
    renderGrade(cust);
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
      const allFolders = await loadCustomerFolders();
      const byLink = linked.length ? allFolders.filter((f) => linked.includes(norm(f.name))) : [];
      const byName = byLink.length ? byLink : allFolders.filter((f) => names.has(norm(custPart(f.name))));
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
    const out = current.files.filter((f) => kws.some((k) => f.path.toLowerCase().includes(k)) && !(ex && ex.test(f.name)));
    if (tab.id !== "analysis") return out;
    // 분석 탭: 워드(docx) 파일은 숨기고, ① AI/클로드 분석 보고서 ② 스마트보장분석 ③ 손해담보 ④ 그 밖의 순으로 (같은 종류는 파일 이름 앞 날짜가 최근인 것부터)
    const rank = (f) => { const n = f.name.replace(/\s/g, ""); return /AI.*분석|클로드|claude|보장분석보고서|심화분석/i.test(n) ? 0 : /스마트보장분석/.test(n) ? 1 : /손해담보/.test(n) ? 2 : 3; };
    return out.filter((f) => kind(f.name) !== "word")
      .map((f, i) => [f, i]).sort((a, b) => rank(a[0]) - rank(b[0]) || String(b[0].name).localeCompare(String(a[0].name)) || a[1] - b[1]).map((x) => x[0]);
  }

  function renderTabs() {
    $("tabs").innerHTML = CFG.tabs.map((t) => {
      const n = t.id === "contact" || t.id === "all" || t.id === "basic" || t.id === "claim" ? "" :
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

  async function selectTab(id) {
    $("tabs").querySelectorAll("button").forEach((b) =>
      b.classList.toggle("on", b.dataset.tab === id));
    const tab = CFG.tabs.find((t) => t.id === id);
    const body = $("tabBody");

    if (id === "contact") return renderContact();
    if (id === "consult") { body.innerHTML = consultHtml(); bindConsult(); return; }
    if (id === "claim") { body.innerHTML = '<div class="card"><p class="hint">청구 기능은 준비 중입니다.</p></div>'; return; } // 보험금 청구: 나중에 따로 만듦

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
  const BASIC_EXCLUDE = ["상품명 - 보험료", "상품명", "보험료", "이관월", "이관일", "미팅장소", "고객 폴더", "분류", "No", "순번", "메모", "모집FC", "담당FC", "담당 FC", "우편번호", "성별", "폴더", "성향", "우선", "입력일시", "고객ID", "고객 ID", "SOL고객번호", "SOL 고객번호", "현재단계", "현재 단계", "고객분류", "고객 분류", "고객등급", "고객 등급", "마케팅동의 만료일", "마케팅 동의 만료일", "계약안내 수신", "통신사", "보장요약", "생일구분", "생일 구분", "차량번호", "자동차보험갱신일"]; // 차량번호·갱신일은 아래 '차량' 줄 하나로 합쳐 보여 줌 // 정보 탭 보기에서 숨기는 항목
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
  function addrWithMap(v) {
    const a = String(v || "").trim();
    if (!a) return "";
    return `<div class="val-row"><span class="val">${esc(a)}</span><span class="acts"><a class="act-btn" href="https://map.kakao.com/link/search/${encodeURIComponent(a)}" target="_blank" rel="noopener noreferrer" aria-label="지도에서 보기" title="지도에서 보기"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M20 10c0 6-8 12-8 12S4 16 4 10a8 8 0 0 1 16 0z"/><circle cx="12" cy="10" r="3"/></svg><span>지도</span></a></span></div>`;
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
        return `<div class="n">${carrier ? `<span class="carrier">${esc(carrier)}</span>` : ""}${esc(num)}</div><span class="acts">` +
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
    return `<span class="sub-l">${left}</span><span class="sub-tags">${types.map((t) => `<span class="tag t-${esc(t)}">${esc(t)}</span>`).join(" ")}</span>`;
  }

  const mdText = (v) => { const p = parseMonthDay(v); return p ? `${p[0]}월 ${p[1]}일` : String(v || ""); };

  function basicHtml() {
    return editMode ? basicFormHtml() : basicViewHtml();
  }

  // 주민번호 = 생년월일(앞 6자리) + 후번호(뒤 7자리). 민감정보 접근이 허용된 계정에게만 보임. 후번호만 있으면 후번호로 표시
  function ssnRow(c, backs, row) {
    const bd = String(c.birth || "").replace(/\D/g, ""), front = bd.length >= 8 ? bd.slice(2, 8) : bd.slice(0, 6);
    const back = String((backs[0] || [])[1] || "").replace(/\D/g, "");
    if (front.length === 6 && back) return row("주민번호", `<div>${esc(front + "-" + back)}</div>`);
    if (back) return row("후번호", `<div>${esc(back)}</div>`);
    return "";
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
        ${canSens() && accts.length ? row("계좌", accts.map(([, v]) => `<div>${esc(v)}</div>`).join("")) : ""}
        ${introContracts.length ? row("소개 계약", `<div>소개${introNames ? " · " + esc(introNames) : ""} <span class="hint">(${introContracts.length}건)</span></div>`) : ""}
        ${emails.map(([k, v]) => row("이메일", valueWithActions(k, v))).join("")}
        ${addrs.length ? `<dt>주소</dt><dd class="dd-left">${addrs.map(([, v]) => addrWithMap(v)).join("")}</dd>` : ""}
        ${carRows.length ? row("차량", carRows.map((r) => `<div>${esc(r.num || "-")}${r.d ? ` · 보험 갱신 ${esc(mdText(r.d))}` : ""}</div>`).join("")) : ""}
        ${chips ? row("추가", `<div class="chips">${chips}</div>`) : ""}
        ${rest.map(([k, vs]) => /가입\s*설계\s*동의\s*만료일/.test(k) ? row("설계 동의 만료", vs.map((v) => `<div>${esc(v)}</div>`).join("")) : k === "신한라이프 유지계약" ? row("유지 계약", vs.map((v) => `<div>${esc(/^\d+$/.test(String(v).trim()) ? String(v).trim() + "건" : v)}</div>`).join("")) : LAND_COLS.test(k.trim()) ? row(k, vs.map((v) => valueWithActions(k, v)).join("")) : row(k, vs.map((v) => `<div>${esc(v)}</div>`).join(""))).join("")}
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
    { label: "연락처", cols: PHONE_COLS, type: "tel", pair: "후번호" },
    { label: "후번호", cols: ["후번호"], type: "tel", paired: true },
    { label: "주소", cols: ["주소"], type: "text", multi: true },
    { label: "계좌", cols: ["계좌"], type: "text", ph: "예: 신한 110-123-456789" },
    { label: "이메일", cols: EMAIL_COLS, type: "email" },
    { label: "유입경로", col: "최초유입경로", cols: ["최초유입경로"], type: "text", select: ["DB", "이관", "지인", "소개"] }, // 직접 쓰지 않고 선택창(드롭박스)에서 고름
    { label: "소개자", col: "소개자명", cols: ["소개자명"], type: "text", ph: "소개자 이름" }
  ];
  function currentValue(c, cols) {
    for (const e of c.entries) for (const k of cols) { const f = e.fields.find(([fk]) => fk.trim() === k); if (f && f[1]) return f[1]; }
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
      <select class="car-mm" aria-label="갱신 월"><option value="">월</option>${Array.from({ length: 12 }, (_, i) => `<option value="${i + 1}" ${p[0] === i + 1 ? "selected" : ""}>${i + 1}월</option>`).join("")}</select>
      <select class="car-dd" aria-label="갱신 일"><option value="">일</option>${Array.from({ length: 31 }, (_, i) => `<option value="${i + 1}" ${p[1] === i + 1 ? "selected" : ""}>${i + 1}일</option>`).join("")}</select>
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
        const mm = r.querySelector(".car-mm").value, dd = r.querySelector(".car-dd").value;
        return { num: r.querySelector(".car-num").value.trim().replace(/\|/g, ""), d: mm && dd ? pad2(mm) + pad2(dd) : "" };
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
      else { row.querySelector(".car-num").value = ""; row.querySelector(".car-mm").value = ""; row.querySelector(".car-dd").value = ""; }
      sync();
    });
  }
  function basicFormHtml() {
    const c = current.cust;
    const x = current.extra || { values: {}, fields: [] };
    const ev = x.values || {};
    const canList = c.entries.some((e) => e.src);
    const selOpts = (f) => { const cur = currentValue(c, f.cols); const list = cur && !f.select.includes(cur) ? [cur, ...f.select] : f.select; return `<option value=""></option>` + list.map((o) => `<option value="${esc(o)}"${o === cur ? " selected" : ""}>${esc(o)}</option>`).join(""); };
    const fld = (f) => `<label class="fld"><span>${esc(f.label)}</span>
        ${f.select ? `<select data-edit="${esc(f.label)}">${selOpts(f)}</select>` : f.multi ? `<textarea data-edit="${esc(f.label)}" rows="2">${esc(currentValue(c, f.cols))}</textarea>`
          : `<input data-edit="${esc(f.label)}" value="${esc(currentValue(c, f.cols))}" type="${f.type}" autocomplete="off"${f.list ? ` list="${f.list}"` : ""}${f.ph ? ` placeholder="${esc(f.ph)}"` : ""}>`}</label>`;
    const field = (k) => {
      const v = esc(ev[k] || "");
      let input;
      if (EXTRA.monthDayCols.includes(k)) {
        const p = parseMonthDay(ev[k]) || [0, 0];
        input = `<div class="mmdd">
            <select data-mm aria-label="${esc(k)} 월"><option value="">월</option>${Array.from({ length: 12 }, (_, i) => `<option value="${i + 1}" ${p[0] === i + 1 ? "selected" : ""}>${i + 1}월</option>`).join("")}</select>
            <select data-dd aria-label="${esc(k)} 일"><option value="">일</option>${Array.from({ length: 31 }, (_, i) => `<option value="${i + 1}" ${p[1] === i + 1 ? "selected" : ""}>${i + 1}일</option>`).join("")}</select>
            <input type="hidden" data-ex="${esc(k)}" value="${p[0] ? pad2(p[0]) + pad2(p[1]) : v}">
          </div>`;
      } else if (/메모/.test(k)) input = `<textarea data-ex="${esc(k)}" rows="4">${v}</textarea>`;
      else if (k === EXTRA.carrierCol) { // 통신사는 목록에서 고르기만 함 (예전에 다른 글자로 저장된 값은 그대로 보이게 목록에 추가)
        const cur = String(ev[k] || "").trim(), base = ["SKT", "KT", "LG U+", "알뜰폰"], opts = cur && !base.includes(cur) ? [cur, ...base] : base;
        input = `<select data-ex="${esc(k)}"><option value=""></option>${opts.map((o) => `<option value="${esc(o)}"${o === cur ? " selected" : ""}>${esc(o)}</option>`).join("")}</select>`;
      }
      else input = `<input data-ex="${esc(k)}" value="${v}" autocomplete="off">`;
      return `<label class="fld ${EXTRA.monthDayCols.includes(k) ? "fld-md" : ""}"><span>${esc(k)}</span>${input}</label>`;
    };
    const fields = x.fields || [];
    // 통신사는 연락처 왼쪽 같은 줄에, 주민번호(앞 6자리-뒤 7자리)는 연락처 아래 줄에 (주민번호는 민감정보 접근 허용 계정만)
    //   앞 6자리는 생년월일에서 가져온 값을 보여 주기만 함 (생년월일은 고객 구분·사진 키로 쓰여 여기서 고치지 않음), 뒤 7자리는 '후번호' 칸에 저장
    const carrierInRow = canList && fields.includes(EXTRA.carrierCol) && !(x.missing || x.error);
    const ssnFld = () => {
      const bd = String(c.birth || "").replace(/\D/g, ""), front = bd.length >= 8 ? bd.slice(2, 8) : bd.slice(0, 6);
      // 뒤 7자리의 첫 숫자는 성별로 정해짐(남 1·여 2, 2000년 이후 출생은 3·4) → 고정, 나머지 6자리만 입력. 성별을 모르면 7자리 모두 입력
      const sx = genderOf(c), yr = +birthYear(c.birth) || 0, d1 = sx === "남" ? (yr >= 2000 ? "3" : "1") : sx === "여" ? (yr >= 2000 ? "4" : "2") : "";
      const cur = String(currentValue(c, ["후번호"])).replace(/\D/g, "");
      const backHtml = d1
        ? `<span class="ssn-d1" aria-label="뒤 첫 숫자(성별)">${d1}</span><input class="ssn-rest" value="${esc(cur.length === 7 ? cur.slice(1) : "")}" type="tel" maxlength="6" inputmode="numeric" autocomplete="off" placeholder="뒤 6자리" aria-label="뒤 6자리"><input type="hidden" data-edit="후번호" data-d1="${d1}" value="${esc(cur.length === 7 ? d1 + cur.slice(1) : cur)}">`
        : `<input data-edit="후번호" value="${esc(cur)}" type="tel" maxlength="7" inputmode="numeric" autocomplete="off" placeholder="뒤 7자리" aria-label="뒤 7자리">`;
      return `<div class="fld ssn-fld"><span>주민번호</span><div class="ssn-in"><input class="ssn-front" value="${esc(front)}" readonly placeholder="생년월일" inputmode="numeric" aria-label="앞 6자리(생년월일)"><b>-</b>${backHtml}</div></div>`;
    };
    const listFields = canList ? EDIT_FIELDS.filter((f) => !f.paired && (canSens() || f.label !== "계좌")).map((f) => f.label === "연락처"
      ? (carrierInRow ? `<div class="fld-2 carrier-phone">${field(EXTRA.carrierCol)}${fld(f)}</div>` : fld(f)) + (canSens() ? ssnFld() : "")
      : fld(f)).join("") : "";
    const hasAddr = !!currentValue(c, ["주소"]); // 주소가 있으면 지역 항목은 수정 화면에서도 숨김
    const more = fields.filter((k) => !EXTRA.primary.includes(k) && k !== CAR_NUM && k !== CAR_DATE && !(carrierInRow && k === EXTRA.carrierCol) && !(hasAddr && k === "지역")); // 미팅장소·메모는 정보 화면에서 바로 입력 (차량은 여러 대 입력 칸으로 따로)
    const hasCars = fields.includes(CAR_NUM) || fields.includes(CAR_DATE);
    return `<div class="card" id="infoForm">
        ${listFields}
        ${x.missing || x.error ? `<p class="hint">추가 정보 엑셀을 읽지 못해 통신사·차량번호·갱신일은 수정할 수 없습니다.</p>` : more.map(field).join("") + (hasCars ? carsFormHtml(ev) : "")}
        ${followFormHtml()}
        <datalist id="routeList"><option value="DB"><option value="이관"><option value="지인"><option value="소개"></datalist>
        <datalist id="carrierList"><option value="SKT"><option value="KT"><option value="LG U+"><option value="알뜰폰"></datalist>
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
        if (i.value.trim() === currentValue(c, f.cols).trim()) return;
        const col = c.entries.flatMap((e) => e.fields.map(([k]) => k.trim())).find((k) => f.cols.includes(k)) || f.col || f.label;
        listValues[col] = i.value.trim();
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
      if (!Object.keys(listValues).length && !extraDirty && !Object.keys(followChanges).length) { done(); return true; }
      btn.disabled = true; btn.textContent = "저장 중…"; $("editSave").disabled = true; $("editSave").textContent = "저장 중…";
      try {
        if (Object.keys(listValues).length) {
          listValues[PRIORITY_COL] = 1;
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
  const CONTRACT_PDF = /^(\d{8}) - .+? - .+? - (.+) - [\d,]+원\.pdf$/i;
  const normP = (s) => String(s || "").replace(/무배당|[\s()（）\[\]·,.\-_*]/g, "").toLowerCase();
  function findContractPdf(date, product) {
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
  // 분석 탭 맨 위 '보장 요약': 고객 시트의 '보장요약' 칸(JSON 글자)을 읽어 보여 줌. 스마트보장분석에서 미리 뽑아 저장해 둔 값이라 따로 불러오지 않음.
  //   형식: {"기준":"2026-10-05","실손":{"세대":"3세대","상품":"…","가입일":"2018-05","혜택":["…"]},"담보":{"진단비":"…","수술비":"…","입원비":"…"}}
  function coverageSummaryHtml() {
    const raw = (current.cust.entries.map((e) => fieldOf(e, "보장요약")).find(Boolean)) || "";
    if (!raw) return "";
    let j = null;
    try { j = JSON.parse(raw); } catch (e) { return ""; }
    if (!j || typeof j !== "object") return "";
    const sil = j["실손"];
    let silHtml = "";
    if (sil && typeof sil === "object" && (sil["세대"] || sil["상품"])) {
      const meta = [sil["상품"], sil["가입일"] ? sil["가입일"] + " 가입" : ""].filter(Boolean).join(" · ");
      const ben = Array.isArray(sil["혜택"]) ? sil["혜택"].filter(Boolean) : [];
      silHtml = `<div class="cov-sil"><b>실손${sil["세대"] ? " " + esc(sil["세대"]) : ""}</b>${meta ? ` <span class="hint">${esc(meta)}</span>` : ""}${ben.length ? `<ul>${ben.map((b) => `<li>${esc(b)}</li>`).join("")}</ul>` : ""}</div>`;
    } else if (sil === "없음" || sil === null || sil === "") {
      silHtml = `<div class="cov-sil"><b>실손</b> <span class="hint">없음</span></div>`;
    }
    // 담보 표: 암·뇌·심 진단비, 수술비, 입원비를 고정 순서로 (값이 없으면 '-'), 그 밖의 항목은 아래에 이어 붙임
    const dam = j["담보"] && typeof j["담보"] === "object" ? j["담보"] : null;
    const FIX = [["암 진단비", ["암진단비", "암 진단비", "암"]], ["뇌 진단비", ["뇌진단비", "뇌 진단비", "뇌"]], ["심 진단비", ["심진단비", "심 진단비", "심", "심장진단비", "심장 진단비"]], ["수술비", ["수술비"]], ["입원비", ["입원비"]]];
    let cov = [];
    if (dam) {
      const used = new Set();
      cov = FIX.map(([label, keys]) => { const k = keys.find((x) => dam[x] != null && String(dam[x]).trim()); if (k) used.add(k); return [label, k ? String(dam[k]).trim() : "-"]; });
      Object.entries(dam).forEach(([k, v]) => { if (!used.has(k) && !FIX.some(([, ks]) => ks.includes(k)) && String(v == null ? "" : v).trim()) cov.push([k, String(v).trim()]); });
    }
    if (!silHtml && !cov.length) return "";
    return `<div class="card cov"><h3 class="sec-title">보장 요약</h3>${silHtml}${cov.length ? `<table class="cov-table"><thead><tr><th>항목</th><th>보장금액</th></tr></thead><tbody>${cov.map(([k, v]) => `<tr><td>${esc(k)}</td><td${v === "-" ? ' class="none"' : ""}>${esc(v)}</td></tr>`).join("")}</tbody></table>` : ""}${j["기준"] ? `<p class="hint">스마트보장분석 ${esc(j["기준"])} 기준</p>` : ""}</div>`;
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
      const pdf = findContractPdf(val(["최초계약일", "계약일"]), c.row["상품명"]);
      const ref = isRefContract(c);
      return `<div class="card${ref ? " ref-card" : ""}">
        <div class="entry-head"><span class="tag t-${esc(c.type)}">${esc(c.type)}</span>
          ${status ? `<span class="tag">${esc(status)}</span>` : ""}
          ${ref ? `<span class="tag ref-tag">담당 ${esc(c.fc)} · 참고</span>` : ""}
          ${c.row["이관일자"] ? `<span class="hint">${esc("이관 " + c.row["이관일자"])}</span>` : ""}
          ${!ref && c.fc ? `<span class="hint rec-fc">${esc("모집 FC: " + c.fc)}</span>` : ""}</div>
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
    body.innerHTML = `<div class="card" id="myCalCard"><details class="fold">
        <summary class="sec-title">구글 캘린더</summary>
        <div id="myCalBox"><p class="hint">캘린더 목록을 불러오는 중…</p></div>
        <div class="file-actions"><button type="button" class="primary" id="myCalSave" disabled>저장</button></div></details></div>` + startCardHtml() + smsCardHtml() + rhCardHtml();
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

  function renderAdmin() {
    const { fcs, perms } = adminState;
    const users = [...new Set([...Object.keys(perms.users), ...extraUsers])].sort();
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
    const ACC = GD ? "구글 계정" : "아이디";
    const pfxHtml = (em) => { const nm = nameOfEm(em); return `<label class="pfx-l"><span>일정 구분</span><input type="text" class="perm-name" data-pfx="${esc(em)}" maxlength="10" value="${esc(nm ? (perms.prefixes || {})[nm] || "" : "")}" placeholder="예: 최" autocomplete="off" ${nm ? "" : "disabled"} aria-label="${esc(em)} 일정 구분"><button type="button" class="secondary" data-pfx-save="${esc(em)}" disabled>저장</button></label>${nm ? "" : '<p class="hint">이름(FC명)을 먼저 입력·저장하면 정할 수 있습니다.</p>'}`; };

    // 순서: 사용 FC 현황 → 사용자 추가 → 설정  (Drive 공유는 '변경 저장' 때 자동 처리, 표에는 표시 안 함)
    const personal = `<div class="card" id="adminCalCard"><details class="fold" id="gFold"><summary class="sec-title">구글 캘린더</summary>
        <div id="gCalBox"><p class="hint">캘린더 목록을 불러오는 중…</p></div>
        <div class="file-actions"><button type="button" class="primary" id="gCalSave" disabled>저장</button></div></details></div>` + startCardHtml() + smsCardHtml() + rhCardHtml();
    $("adminBody").innerHTML = personal + `<h2 class="admin-sec">관리자</h2>
      <div class="card">
        <h3 class="sec-title">FC 사용 설정</h3>
        ${(() => {
          const sh = adminState.shares, adm = (CFG.admins || []).map((a) => a.toLowerCase());
          const hits = ((sh && sh.anc) || []).flatMap((a) => a.perms.filter((p) => p.role !== "owner" && !adm.includes(p.email)).map((p) => esc((perms.names || {})[p.email] || p.email) + " ← '" + esc(a.name) + "' 폴더(" + p.role + ")"));
          return hits.length ? `<div class="perm-warn"><b>⚠ 상위 폴더 공유 감지</b><br>${hits.join("<br>")}<br><span class="hint">이 계정은 FC 체크와 상관없이 상위 폴더 전체를 Drive에서 열 수 있습니다. 해당 계정 카드의 [저장]을 누르면 상위 공유를 해제하고 FC 폴더만 공유합니다.</span></div>` : "";
        })()}
        <div class="perm-list">
          ${(CFG.admins || []).map((a) => `<div class="perm-admin"><span class="perm-em">${esc(a)}</span> <span class="tag">관리자</span> <span class="hint">전체 FC</span>
            <span class="adm-name"><input type="text" class="perm-name" data-admin-name="${esc(a.toLowerCase())}" placeholder="내 FC 이름(예: 김현성)" maxlength="20" value="${esc(nameOfEm(a.toLowerCase()))}" aria-label="관리자 FC 이름"><button type="button" class="secondary" data-admin-save="${esc(a.toLowerCase())}" disabled>저장</button></span></div><div class="perm-pfx perm-pfx-adm">${pfxHtml(a.toLowerCase())}</div>`).join("")}
          ${users.map((em) => `<div class="perm-row${permOpen.has(em) ? " open" : ""}" data-em="${esc(em)}">
            <div class="perm-head"><button type="button" class="perm-fold" data-pfold aria-expanded="${permOpen.has(em)}" aria-label="${esc(nameOfEm(em) || em)} 펼치기·접기"><svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M9 6l6 6-6 6"/></svg></button>
              <span class="perm-id"><input type="text" class="perm-name" data-name placeholder="이름(FC명) 입력" value="${esc(nameOfEm(em))}" maxlength="20" aria-label="${esc(em)} 이름"><span class="perm-mail">${esc(em)}</span></span>
              <span class="perm-btns"><button type="button" class="secondary" data-rowedit="${esc(em)}">수정</button><button type="button" class="del-btn" data-del="${esc(em)}" aria-label="${esc(em)} 삭제">삭제</button><button type="button" class="secondary" data-rowcancel="${esc(em)}" hidden>취소</button><button type="button" class="primary row-save" data-rowsave="${esc(em)}" hidden>저장</button></span>
            </div>
            <div class="perm-body"><div class="perm-pfx">${pfxHtml(em)}</div>${startRowHtml(perms, em)}
            <div class="perm-sens"><label class="chip"><input type="checkbox" class="cb" data-sens="${esc(em)}" ${(perms.sensitive || []).includes(em) ? "checked" : ""} aria-label="${esc(em)} 민감정보 접근"> 민감정보 접근 허용</label></div>
            <div class="perm-sales"><p class="hint perm-lbl">• 영업 페이지 접근 허용</p><div class="perm-fcs">${fcsFor(em).map((fc) => { const own = !!nameOfEm(em) && fc === nameOfEm(em); return `<span class="chip"><label><input type="checkbox" class="cb" data-sf="${esc(em)}|${esc(fc)}" ${own ? 'data-own="1" checked disabled' : ((perms.salesFcs || {})[em] || []).includes(fc) ? "checked" : ""} aria-label="${esc(em)} 영업 ${esc(fc)}"> ${esc(fc)}${own ? ' <small class="hint">(본인)</small>' : ""}</label></span>`; }).join("")}</div></div>
            <p class="hint perm-lbl">• 고객 페이지 접근 허용</p>
            <div class="perm-fcs fc-lines">
              <label class="chip chip-all" hidden title="전체(*)는 폐지 — 해제하고 FC를 체크하세요"><input type="checkbox" class="cb" data-all ${isAll(em) ? "checked" : ""} aria-label="${esc(em)} 전체"> 전체(폐지 예정)</label>
              <div class="fc-line fc-head"><span class="fc-nm">FC</span><span class="fc-c">보기</span><span class="fc-c">수정</span></div>
              ${fcsFor(em).map((fc) => { const own = fc === ownFc(em); return `<div class="fc-line${own ? " own" : ""}"><span class="fc-nm">${esc(fc)}${own ? ' <small class="hint">(본인)</small>' : ""}</span><label class="fc-c"><input type="checkbox" class="cb" data-fc="${esc(fc)}" ${has(em, fc) ? "checked" : ""} ${own ? 'data-own="1" disabled' : ""} aria-label="${esc(em)} ${esc(fc)} 보기"></label><label class="fc-c"><input type="checkbox" class="cb ro" data-edit="${esc(fc)}" ${has(em, fc) && canEdit(em, fc) ? "checked" : ""} ${own || isAll(em) ? 'data-own="1" disabled' : ""} aria-label="${esc(em)} ${esc(fc)} 수정"></label></div>`; }).join("")}
            </div></div>
          </div>`).join("")}
        </div>
        ${users.length ? "" : '<p class="hint">아직 권한을 준 사용자가 없습니다. 아래에서 이메일을 추가하세요.</p>'}
        <div class="add-user"><h4 class="sub-title">사용자 추가</h4>
        <form id="addUserForm" class="search">
          <input id="newName" type="text" placeholder="이름(FC명)" autocomplete="off" maxlength="20" style="max-width:130px">
          <input id="newUser" type="email" placeholder="${GD ? "구글 계정 이메일 (예: name@gmail.com)" : "Microsoft 계정 이메일"}" autocomplete="off" required>
          <button class="primary" type="submit">추가</button>
        </form>
        </div>
        <div class="file-actions">
          <button class="secondary" id="resetPerms" disabled>되돌리기</button>
          <button class="primary" id="savePerms" disabled>변경 저장</button>
        </div>
      </div>` + smsTeamCardHtml() + settingsCardHtml(); // 오른손 모드는 '설정' 카드 안에 있음

    $("addUserForm").addEventListener("submit", (ev) => {
      ev.preventDefault();
      const em = $("newUser").value.trim().toLowerCase();
      if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(em)) return toast("이메일 형식을 확인해 주세요.");
      if ((CFG.admins || []).map((a) => a.toLowerCase()).includes(em)) return toast("관리자는 이미 모든 폴더를 볼 수 있습니다.");
      const pending = collectRows(); // 표에서 아직 저장 안 한 체크 상태 유지
      const nm = $("newName").value.trim();
      extraUsers.add(em); permOpen.add(em);
      renderAdmin();
      applyRows(pending);
      if (nm) { const nr = $("adminBody").querySelector(`.perm-row[data-em="${CSS.escape(em)}"] [data-name]`); if (nr) { nr.value = nm; nr.dispatchEvent(new Event("input", { bubbles: true })); } }
    });

    const rows = [...$("adminBody").querySelectorAll(".perm-row")];
    const same = (a, b) => a.length === b.length && a.every((x) => b.includes(x));
    const sel = (row) => {
      const all = row.querySelector("input[data-all]");
      return all.checked ? ["*"] : [...row.querySelectorAll("input[data-fc]")].filter((b) => b.checked).map((b) => {
        const ed = row.querySelector(`input[data-edit="${CSS.escape(b.dataset.fc)}"]`);
        return b.dataset.fc + (ed && ed.checked ? "" : ":r"); // 수정 체크 없으면 보기만
      });
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
      btn.addEventListener("click", async () => { // 관리자 본인 FC 이름: 권한 파일 names에 저장 (영업 화면 FC 선택 기본값·문자 {FC명}에 사용)
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
      r.querySelector(".perm-head").addEventListener("click", (ev) => { // 머리 줄(이름 입력칸·버튼 제외)을 누르면 펼침/접힘
        if (ev.target.closest("input, [data-rowedit], [data-del], [data-rowcancel], [data-rowsave]")) return;
        setOpen(r, !r.classList.contains("open"));
      });
      r.querySelector("[data-rowedit]").addEventListener("click", () => { setOpen(r, true); rowLock(r, false); });
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

  // 등록된 아이디 삭제 (권한 파일에서 바로 제거)
  async function deleteUser(em) {
    const shared = GD && shareState(em) && shareState(em) !== "none";
    if (!confirm(`${em} 계정을 삭제할까요?\n삭제하면 이 계정은 앱에서 고객을 검색할 수 없습니다.` +
      (shared ? "\nGoogle Drive 폴더 공유도 함께 해제합니다." : ""))) return;
    try {
      const perms = await readPerms();
      delete perms.users[em];
      if (perms.names) delete perms.names[em];
      perms.sensitive = (perms.sensitive || []).filter((x) => x !== em); // 삭제한 계정의 민감정보 허용도 함께 정리
      if (perms.salesFcs) delete perms.salesFcs[em];
      extraUsers.delete(em);
      await writePerms(perms);
      adminState.perms = perms;
      if (shared) await unshareFolders(em);
      toast(`${em} 계정을 삭제했습니다.`);
      renderAdmin();
    } catch (e) {
      handleError(e);
    }
  }

  async function savePerms(changes) {
    if (!changes.length) return;
    if (changes.some(([, l]) => l.includes("*"))) return toast("전체(*)는 더 이상 지원하지 않습니다. FC를 하나씩 체크해 주세요.");
    // Google Drive: FC를 체크한 계정은 폴더 공유, 체크를 모두 뺀 계정은 공유 해제 (변경 저장 한 번으로)
    const auto = GD && adminState.shares;
    const toShare = auto ? changes.filter(([em, list]) => list.length && !shareMatches(em, list)).map(([em]) => em) : [];
    const toUnshare = auto ? changes.filter(([em, list]) => !list.length && shareState(em) !== "none").map(([em]) => em) : [];
    const listOfChange = (em) => (changes.find((c) => c[0] === em) || [em, []])[1];
    if (toShare.length || toUnshare.length) {
      const msg = ["저장하면서 Google Drive 공유도 바꿉니다."];
      if (toShare.length) msg.push("· 공유(체크한 FC 폴더만 편집·보기만, 설정 폴더 보기): " + toShare.map((em) => em + " [" + listOfChange(em).join(", ") + "]").join(" / "));
      const ancMsg = [...toShare, ...toUnshare].flatMap((em) => ancShares(em).map((a) => em + " ← '" + a.folder + "' 폴더 공유(" + a.perm.role + ")"));
      if (ancMsg.length) msg.push("· 상위 폴더에 직접 공유된 것도 함께 해제합니다(이 공유가 있으면 FC 제한이 Drive에서 안 걸림): " + ancMsg.join(", "));
      if (toUnshare.length) msg.push("· 공유 해제: " + toUnshare.join(", "));
      if (!confirm(msg.join("\n"))) return;
    }
    const btn = $("savePerms");
    btn.disabled = true;
    btn.textContent = "저장 중…";
    try {
      // 저장 직전에 최신 파일을 다시 읽고, 바뀐 아이디만 반영 (다른 설정 보존)
      const perms = await readPerms();
      perms.names = perms.names || {};
      changes.forEach(([em, list, nm]) => {
        if (list.length) { perms.users[em] = list; if (nm !== undefined) { if (nm) perms.names[em] = nm; else delete perms.names[em]; } }
        else { delete perms.users[em]; delete perms.names[em]; extraUsers.delete(em); }
      });
      perms.fcFolders = adminState.fcIds || perms.fcFolders || {}; // 사용자 앱이 FC 폴더를 ID로 바로 열 수 있게
      await writePerms(perms);
      adminState.perms = perms;
      const failed = [], tAll = { cust: false, app: false, fc: [], anc: [] };
      for (const em of toShare) { try { const tt = await grantFolders(em, listOfChange(em)); tAll.cust = tAll.cust || tt.cust; tAll.app = tAll.app || tt.app; tAll.fc.push(...tt.fc); tAll.anc.push(...tt.anc); } catch (e) { console.error(e); failed.push(em); } }
      for (const em of toUnshare) { try { await unshareFolders(em); } catch (e) { console.error(e); failed.push(em); } }
      if (toShare.length || toUnshare.length) { // 실패가 있으면 상태를 확실히 알 수 없으니 전체를 다시 읽고, 모두 성공이면 바뀐 폴더만 갱신
        if (failed.length) adminState.shares = await loadShares(); else await refreshSharesFor(tAll);
      }
      toast(failed.length ? `저장했습니다. 공유하지 못한 계정: ${failed.join(", ")} (구글 계정 이메일인지 확인)` : `저장했습니다 (${changes.length}개 계정${toShare.length ? ", 폴더 공유 " + toShare.length : ""}${toUnshare.length ? ", 공유 해제 " + toUnshare.length : ""}).`);
      renderAdmin();
    } catch (e) {
      btn.disabled = false;
      btn.textContent = "변경 저장";
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
      await requestGToken(other ? "select_account" : undefined, other);
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
    $("adminBody").querySelectorAll("[data-pfx-save], [data-sa-save]").forEach((b) => list.push(b));
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
  // 영업 화면 위쪽 줄(이달이관·따행 선택, 월 이동, FC 선택)의 높이를 재 두어, 그 아래 '설문 발송' 줄이 바로 밑에 붙어 고정되게 함
  new MutationObserver(() => { const st = document.querySelector("#salesView .sales-top"); if (st) document.documentElement.style.setProperty("--salestop", Math.round(st.getBoundingClientRect().height) + "px"); }).observe($("salesView"), { subtree: true, childList: true });
  const fsWait = async (b) => { for (let i = 0; i < 160 && b.isConnected && /저장 중/.test(b.textContent); i++) await new Promise((r) => setTimeout(r, 150)); };
  // ---------- 설명서: 관리자·설정 화면의 설명을 한곳에 모음 (아래 '설명서' 버튼) ----------
  function helpHtml() {
    const iconBox = `<div class="help-icon"><img src="icon.svg?v=3" width="72" height="72" alt="신한FC m 앱 아이콘"><div><b>앱 아이콘 안내</b><p>고객의 보장을 지키는 <b>방패</b> 안에, 계약과 실적이 꾸준히 <b>오르는 화살표</b>를 담았습니다. 보험 설계사의 일(고객을 지키고, 함께 성장하는 것)을 한 그림으로 줄였습니다.</p><p>연한 하늘색 바탕과 파란 선은 신뢰감 있고 차분한 업무용 앱의 느낌을 주고, 굵은 윤곽선 방패는 작은 홈 화면 아이콘에서도 또렷하게 보이도록 단순하게 만들었습니다. 신한FC m(모바일)에서는 아래 메뉴 맨 왼쪽의 이 아이콘을 누르면 AI 도우미가 열리고, PC 화면 왼쪽 위에는 "신한FC"로 표시됩니다.</p></div></div>`;
    const sec = (t, items) => `<h4>${t}</h4><ul>${items.map((x) => `<li>${x}</li>`).join("")}</ul>`;
    const grp = (t) => `<h3 class="help-grp">${t}</h3>`; // 하단 메뉴 순서대로 묶음: AI · 일정 · (고객) · 영업 · 설정/관리자
    const s_start = sec("시작 화면", ["<b>처음 보일 화면</b>: 설정(관리자는 관리자) 화면의 '시작 화면' 카드에서 앱을 열었을 때 처음 보일 화면을 고릅니다. 일정 · 고객(기본) · 영업 중 하나를 고르고 아래 [저장]을 누르면 다음에 앱을 열 때부터 적용됩니다. 저장한 값은 내 FC 폴더 안 '_앱설정' 폴더에도 기록되어 관리자가 확인·수정할 수 있고, 다른 폰에도 따라갑니다. 관리자는 설정 화면 아래 ‘관리자’의 각 사용자 카드(눌러서 펼침) 안 ‘시작 화면’에서 그 사용자의 값을 보고 바꿀 수 있습니다.", "<b>AI 도우미 바로 열기</b>: '시작하면 AI 도우미를 바로 열기'를 켜면 앱을 열자마자 AI 도우미 창이 올라옵니다.", "내 FC 폴더에 쓸 수 없는 계정은 이 폰에만 저장됩니다. 앱을 쓰다 잠깐 나갔다 30분 안에 돌아오면 보던 화면(일정·영업)을 이어서 보여 주는 기능이 먼저 적용될 수 있습니다."]);
    const s_ai = sec("AI 도우미", ["<b>열기·닫기</b>: 아래 메뉴 맨 왼쪽의 앱 아이콘을 누르면 AI 도우미가 아래에서 올라옵니다. 아래로 끌어내리거나, 아이콘을 한 번 더 누르거나, 바깥 어두운 곳을 누르거나, 폰의 뒤로가기를 누르면 내려갑니다.", "<b>말하기·입력</b>: 큰 마이크 버튼을 눌러 말하면 입력칸에 글자로 들어가고, 키보드로 직접 써도 됩니다. 아래 예시 문구를 누르면 입력칸에 들어갑니다.", "<b>지금 알아듣는 말</b>: ① “홍길동 검색” → 고객 검색 ② “내일 오후 3시 홍길동 상담 일정 추가” → 구글 캘린더(1팀)에 1시간짜리 일정 등록 (“내 일정에 …”이라고 하면 내 FC 이름과 같은 이름의 내 캘린더에 등록, 이때는 앞 글자를 붙이지 않음) (시간을 말하지 않으면 종일 일정, 날짜를 말하지 않으면 오늘) ③ “오늘 일정 보여줘” → 일정 화면. 말한 뒤에는 \"이렇게 이해했습니다\" 확인 카드가 뜨고, [실행]을 눌러야 실제로 처리됩니다. 못 알아들으면 말하는 방법을 안내합니다.", "<b>일정 구분(제목 앞 글자)</b>: 관리자가 사용자마다 정해 두면(설정 화면 아래 ‘관리자’의 각 사용자 카드 안 ‘일정 구분’) 등록할 때 제목 맨 앞에 붙습니다. 예: 최)홍길동 상담. 칸에는 글자만 적어도 뒤에 ‘)’가 자동으로 붙고, ‘)’까지 적어도 하나만 붙습니다", "<b>이름 보호</b>: 지금은 외부 AI를 쓰지 않고 이 폰 안에서만 해석하므로 고객 이름이 밖으로 나가지 않습니다. 쓰면서 알아듣는 말을 계속 늘려 갑니다."]);
    const s_task = sec("구글 할 일 (일정 화면)", ["그날 마감인 구글 할 일이 일정 화면 맨 위에 나오고, 체크하면 완료로 바뀝니다.", "로그인할 때 할 일 권한도 함께 허용하면 따로 연결하지 않아도 됩니다. 연결이 풀리면(약 1시간 뒤) 앱이 먼저 조용히 다시 연결을 시도하고, 안 되면 나오는 [연결] 버튼을 한 번 누르세요."]);
    const s_cal = sec("구글 캘린더 (표)", ["<b>Show</b>: 일정 화면에 보일 캘린더를 고릅니다. 꺼 두면 일정·정보를 켤 수 없습니다.", "<b>일정</b>: 영업 일정으로 쓸 캘린더입니다. 고객 매칭을 하고, 이동 거리·예상 시간·회사 출발·도착은 그 일정 카드에서 <b>확정</b>을 누른 영업일정에만 계산됩니다(예전에 1팀 캘린더에 만든 일정은 그대로 영업일정). 켜면 정보도 함께 켜집니다.", "<b>정보</b>: 이 캘린더의 일정에서 고객 정보를 찾아 붙입니다. 끄면 일정도 꺼집니다."]);
    const s_addev = sec("일정 추가 (일정 화면)", ["일정 화면 오른쪽 위의 <b>추가</b> 버튼을 누르면 제목·캘린더·날짜·시간(또는 종일)·장소를 넣어 새 일정을 구글 캘린더에 등록할 수 있습니다. 처음 열 때 캘린더는 1팀이 선택되어 있습니다.", "일정은 처음에 모두 <b>비영업일정</b>입니다. 카드의 <b>확정</b>(수정 왼쪽)을 누르면 <b>영업일정</b>이 되어 이동 거리·시간이 나옵니다(영업일정 카드에는 ‘영업’ 버튼이 나오고, 누르면 다시 비영업일정으로 돌아갑니다). <b>복귀</b> 버튼(추가의 왼쪽)은 그날 영업일정을 모두 비영업일정으로 바꾸고, 동명이인 지정·고객 정보로 넣은 일정 장소도 되돌립니다(영업일정이나 처리한 일정이 있는 날에만 보임). 새로 만든 일정과 1팀 외 캘린더의 일정은 비영업일정으로 시작합니다."]);
    const s_nf = sec("전화·문자 필터 (이달이관 · 따행)", ["목록 위의 <b>전화</b>·<b>문자</b> 칸을 체크하면 그 달에 아직 전화(또는 문자) 시도 ✓가 없는 고객만 보입니다. 둘 다 체크하면 전화도 문자도 안 한 고객만 보입니다.", "체크를 끄면 전체가 다시 보입니다. 따행 줄은 ‘설문 ☐발송 ☐응답 | ☐쿠폰 | ☐전화 ☐문자’ 순서입니다. 발송·응답·쿠폰은 체크하면 그 날짜가 있는 고객만, 전화·문자는 반대로 아직 안 한 고객만 보여 주며, 함께 쓸 수 있습니다."]);
    const s_sms = sec("문자 문구", ["영업 화면에서 [문자]를 누르면 이 문구가 채워진 문자 작성 화면이 열립니다. 내가 저장한 문구가 있으면 그것을 쓰고, 없으면 관리자가 정해 준 팀 기본 문구(관리자 화면의 ‘문자 문구 (팀 기본)’)가 미리 들어 있습니다. 고치거나 지우면 내 설정으로 저장됩니다. 지우면 문구 없이 작성 화면만 열립니다.", `쓸 수 있는 항목: <b>${SMS_VARS}</b> — {고객명}=계약자 이름, {FC명}=담당 FC 이름, {상품명}, {이관일}(예: 11월 9일, 따행 목록에서는 비어 있음)`, "문자는 작성 화면을 연 뒤 직접 '보내기'를 눌러야 발송되며, 앱에는 '시도'만 기록됩니다. 일부 폰·문자앱·카카오톡 안의 브라우저에서는 문구가 채워지지 않을 수 있고, 45자를 넘으면 장문 문자로 나가거나 잘릴 수 있습니다."]);
    const s_login = sec("로그인 (공통 로그인)", ["한 번 로그인하면 같은 브라우저(크롬·사파리 등)에서 다른 내 앱(차차차·항상일상·말만해요·메뉴메이커)도 다시 로그인하지 않고 바로 열립니다. 로그인은 약 1시간 유지되고, 끊기면 아무 앱에서 [다시 연결]을 한 번 누르면 나머지도 이어집니다.", "다른 앱에서만 로그인했고 이 앱에 필요한 권한이 모자라면 버튼이 ‘구글 연결 (권한 추가)’로 바뀝니다. 한 번 눌러 허용하세요.", "<b>로그아웃은 모든 앱에서</b> 함께 됩니다(오른쪽 위 전원 버튼). 아이폰은 홈 화면에 추가한 앱마다 저장 공간이 따로라 서로 나눠 쓰지 못합니다."]);
    const s_go = sec("말만해요 연결", ["말만해요에서 “김현성 고객 검색해줘”처럼 말하면 말만해요 화면 안에 이 앱의 고객 검색 목록이 나타납니다. 한 명만 나와도 목록이 먼저 보이고, 고객을 누르면 그 안에서 고객 정보가 열립니다. 이때는 앱 전체가 아니라 그 페이지만 보입니다(위쪽 바·아래 메뉴 없음). 목록→고객 정보, 정보·계약·분석·상담 탭 이동은 그대로이고, 뒤로가기는 말만해요의 왼쪽 위 버튼이 한 단계씩 되돌립니다.", "말만해요의 [앱으로 열기]를 누르면 보던 바로 그 화면(검색 목록 또는 그 고객 정보)으로 신한FC m이 전체 화면으로 열립니다. 말만해요 안에서는 로그아웃 버튼이 없고(로그아웃은 말만해요에서), 로그인이 풀렸으면 ‘말만해요에서 다시 로그인’ 안내가 나옵니다.", "“오늘 일정 보여줘”처럼 일정을 물으면 일정 화면(하루 보기)만 말만해요 안에 나타납니다. 날짜 이동·추가 버튼은 그대로 쓰고, [앱으로 열기]를 누르면 보던 날짜로 열립니다.", "검색한 이름은 주소의 # 뒤에만 실려 서버로 전송되지 않습니다."]);
    const s_map = sec("고객 주소 지도", ["고객 정보의 주소 오른쪽 <b>위치 아이콘</b>을 누르면 카카오맵에서 그 주소를 찾은 지도가 새 창으로 열립니다. 누를 때만 주소가 카카오맵에 전달됩니다."]);
    const s_rh = sec("오른손 모드", ["하단 메뉴, 탭, 선택 버튼을 오른쪽으로 모아 엄지로 누르기 쉽게 합니다. 이 폰에만 적용됩니다."]);
    const s_save = sec("저장", ["저장할 내용이 하나라도 있으면 아래 [저장]이 켜집니다. 누르면 바뀐 모든 카드(시작 화면·문자 문구·캘린더·사용자별 일정 구분과 시작 화면·FC 권한 등)를 한꺼번에 저장합니다. 각 카드 안의 [저장]을 눌러도 됩니다."]);
    const s_perm = sec("FC 사용 설정 (권한)", ["이름(FC명)을 적고, <b>고객 페이지 접근 허용</b>에서 볼 FC를 체크하세요. 'FC 체크 = 보기', 옆의 '수정'도 체크하면 수정까지 가능합니다." + (GD ? " 저장하면 Google Drive 폴더 공유도 함께 바뀝니다." : ""), "<b>구글 캘린더</b>: 설정 화면 맨 위 ‘구글 캘린더’ 카드에서 관리자의 기본 캘린더와 영업·정보 캘린더를 고르고 그 카드의 [저장]을 누릅니다.", "<b>민감정보 접근 허용</b>: 체크한 사용자만 주민번호·계좌를 보고 고칠 수 있습니다.", "<b>영업 페이지 접근 허용</b>: 영업(이달이관·따행) 선택창에 나올 FC입니다. 아래 고객 페이지 접근을 허용한 FC만 고를 수 있습니다.", "관리자는 모든 FC를 볼 수 있어서 별도 체크가 필요 없습니다."]);
    const s_add = sec("사용자 추가", ["이름과 구글 계정 이메일을 넣고 [추가]를 누르면 위 표에 줄이 생깁니다. 볼 FC를 체크하고 [저장]을 눌러야 적용됩니다."]);
    const s_set = sec("설정", ["<b>회사 주소</b>: 일정 이동시간 계산의 출발·도착 기준입니다.", "<b>카카오 REST API 키</b>: 이동 거리·시간 계산에 씁니다. 실수로 바뀌지 않게 잠겨 있고, [수정]을 눌러야 고칠 수 있습니다. 저장하면 다시 잠깁니다."]);
    const parts = [grp("AI (아래 메뉴 맨 왼쪽)"), s_ai, grp("일정"), s_addev, s_task, s_cal, grp("영업"), s_nf, s_sms, grp("설정"), s_start, s_rh, s_save, s_login, s_go, s_map];
    if (isAdmin()) parts.push(grp("관리자 (설정 화면 아래쪽)"), s_perm, s_add, s_set);
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
  // 새로고침 버튼(로그아웃 왼쪽): 이 화면을 다시 불러옴. 영업·일정 화면은 보던 자리로 돌아옴
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
  function runTopSearch(q) {
    if (editMode && !$("customerView").hidden) { confirmLeaveEdit().then((ok) => { if (ok) runTopSearch(q); }); return; }
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
    runTopSearch(q);
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
  function writeGAuth(a) { try { localStorage.setItem(GAUTH_KEY, JSON.stringify({ v: 1, token: a.token, exp: a.exp, scopes: a.scopes || [], email: a.email || "" })); } catch (e) { /* 저장 못 해도 이번엔 사용 */ } }
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
      writeGAuth({ token: t.token, exp: t.exp, scopes: sc, email: gAccount() });
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
          writeGAuth({ token: r.access_token, exp: Date.now() + (r.expires_in || 3600) * 1000, scopes: String(r.scope || "").split(/\s+/).filter(Boolean), email: !noHint && prev ? prev.email : "" });
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
      .map((c) => ({ id: c.id, name: calName(c), color: c.backgroundColor, primary: !!c.primary }));
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
        pick: String(((e.extendedProperties || {}).private || {}).shlifeCust || ""), // 동명이인 선택(모든 기기 공유)
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

  // 일정 제목에 들어 있는 고객 이름 → 고객 (같은 이름이 여러 명이면 모두, 최근 등록 순)
  function matchCustomers(title, records, contracts) {
    const nt = norm(title);
    const fcs = new Set(records.map((r) => norm(r.fc)).filter(Boolean));
    const names = [...new Set(records.map((r) => baseName(r.name)))].filter((n) => norm(n).length >= 2 && nt.includes(norm(n)));
    if (!names.length) return [];
    // FC 이름이 아닌 것, 긴 이름 우선
    names.sort((a, b) => (fcs.has(norm(a)) - fcs.has(norm(b))) || norm(b).length - norm(a).length);
    const name = names[0];
    const city = titleRegion(title);
    const fcHead = (title.match(/^\s*([^\s)]{1,3})\)/) || [])[1] || ""; // "최) …" → 최수진 FC 고객 우선
    const recs = records.filter((r) => norm(baseName(r.name)) === norm(name)); // "이영숙(640927)"도 포함
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

  function scheduleShell() {
    return `<div class="sched-top"><div class="sched-modes">${[["day", "Today"], ["week", "Week"], ["month", "Month"]].map(([m, l]) =>
      `<button type="button" data-sm="${m}" class="${schedMode === m ? "on" : ""}">${l}</button>`).join("")}</div></div>
      <div class="sched-nav"><span class="sched-rt"><span class="sched-seg"><button type="button" class="link-btn" id="pickReset" hidden>복귀</button><button type="button" class="link-btn" id="schedAdd" aria-label="일정 추가" title="일정 추가">추가</button></span></span><button type="button" id="schedPrev" aria-label="이전"><svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M15 6l-6 6 6 6"/></svg></button>
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
    { const pr0 = $("pickReset"); if (pr0) pr0.hidden = true; }
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
  const isInfoEvent = (e) => !e.allDay && (schedInfoIds.has(e.cal.id) || schedSalesIds.has(e.cal.id)); // 고객 정보

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
      (r.items || []).forEach((t) => { if (t.due) all.push({ id: t.id, listId: l.id, list: l.title || "", title: t.title || "(제목 없음)", done: t.status === "completed", due: String(t.due).slice(0, 10) }); });
    }));
    return all.filter((t) => t.due >= ymdStr(from) && t.due < ymdStr(to)).sort((a, b) => a.due.localeCompare(b.due) || a.title.localeCompare(b.title, "ko"));
  }
  const taskConnectHtml = (msg) => `<div class="card task-card"><h3 class="sec-title">할 일</h3><p class="hint">${esc(msg || "구글 할 일을 보려면 한 번 연결해 주세요.")}</p><button type="button" class="ev-edit" data-taskconnect>할 일 연결</button></div>`;
  let taskOwnerCal = "";
  // 할 일 소유주 = 앱에 등록된 사용자 이름(관리자 화면의 이름), 없으면 내 기본 캘린더 이름
  async function taskOwnerName() {
    if (taskOwnerCal) return taskOwnerCal;
    let n = ""; try { n = (await myFcName()) || ""; } catch (e) { n = ""; }
    if (!n) { try { const pc = (await calendarList()).find((c) => c.primary); n = pc ? pc.name : ""; } catch (e) { n = ""; } }
    return n;
  }
  const taskItemHtml = (t) => `<li class="task-item${t.done ? " done" : ""}"><label><input type="checkbox" class="cb" data-task="${esc(t.listId)}|${esc(t.id)}" ${t.done ? "checked" : ""}> <span class="task-title">${esc(t.title)}</span>${taskOwnerCal ? ` <span class="ev-cal">- ${esc(taskOwnerCal)}</span>` : ""}</label>${t.list ? `<span class="hint task-list-name">${esc(t.list)}</span>` : ""}</li>`;
  // 하루 보기 맨 위 '할 일' 카드
  async function loadDayTasks(box, d, req) {
    if (!box) return;
    if (!storedTasksToken()) { box.innerHTML = taskConnectHtml(); trySilentTasks().then((ok) => { if (ok && req === schedReq) loadDayTasks(box, d, req); }); return; }
    try {
      const list = await fetchTasks(d, addDays(d, 1));
      if (req !== schedReq) return;
      // 소유주: 구글 할 일은 로그인한 구글 계정 것이라, 앱에 등록된 내 이름(없으면 구글 계정)을 제목에 붙임
      taskOwnerCal = await taskOwnerName();
      if (req !== schedReq) return;
      box.innerHTML = list.length ? `<div class="card task-card"><h3 class="sec-title">할 일 <small class="hint">${list.filter((t) => !t.done).length}/${list.length}</small></h3><ul class="task-list">${list.map(taskItemHtml).join("")}</ul></div>` : "";
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
        card.insertAdjacentHTML("beforeend", `<div class="ev-row task-row${t.done ? " done" : ""}">${t.done ? "☑" : "☐"} ${esc(t.title)}${taskOwnerCal ? ` <span class="ev-cal">- ${esc(taskOwnerCal)}</span>` : ""}</div>`);
      });
    } catch (e) { /* 연결이 없거나 만료면 주간 보기엔 아무 것도 덧붙이지 않음 */ }
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
    pr.onclick = async () => {
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
          for (const it of pinned) await gcalPatch(it.e.calId, it.e.eventId, { extendedProperties: { private: { shlifeCust: "" } } }).catch(() => {});
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
    if (evs.some(isInfoEvent)) {
      try { [records, contracts] = await Promise.all([loadLists(), loadContracts().catch(() => [])]); } catch (e) { handleError(e); }
    }
    if (req !== schedReq) return;
    const picks = readPicks();
    const items = evs.map((e) => {
      const all = isInfoEvent(e) ? matchCustomers(e.title, records, contracts) : [];
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
    { // 처리한 일이 하나라도 있으면 복귀 버튼 표시
      const done = items.some((it) => salesOf(it.e) || it.pinned || it.ignored || (it.e.location && it.all.some((c) => [currentValue(c, ["주소"]), meetingOf(c)].map((v) => String(v || "").trim()).filter(Boolean).includes(it.e.location.trim()))));
      if (pr) pr.hidden = !done;
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
    if (!a || a.closest("#salesView, #contactSheet")) return; // 영업 화면은 따로 '시도'를 기록함
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
  function custBlockHtml(c, i, j, multi, pinned) {
    const phone = (c.entries.find((x) => x.phone) || {}).phone || "";
    const addr = currentValue(c, ["주소"]);
    const types = [...new Set(c.entries.map((x) => x.type))];
    const g = genderOf(c);
    return `<div class="ev-cust">
        <div class="ev-namerow"><button type="button" class="ev-name" data-open="${i}-${j}"><span class="cname ${gradeClass(gradeOf(c))}">${esc(c.name)}</span>
          ${types.map((t) => `<span class="tag t-${esc(t)}">${esc(t)}</span>`).join("")}<span class="chev">›</span></button>
        ${multi ? `<button type="button" class="pick-btn" data-pick="${i}-${j}">선택</button>` : ""}${pinned ? `<button type="button" class="pick-btn" data-unpick="${i}">다시 선택</button>` : ""}</div>
        <dl class="ev-info">
          ${c.birth ? `<dt>생년월일</dt><dd>${esc(c.birth)}${g ? ` (${esc(g)})` : ""}</dd>` : ""}
          ${phone ? (() => { const num = (String(phone).match(MOBILE_RE) || [phone])[0], dg = String(num).replace(/[^\d+]/g, ""); return `<dt>연락처</dt><dd class="put-row" data-cn="${esc(c.name)}" data-cfc="${esc(memoFcOf(c))}"><span><a class="nowrap" href="tel:${esc(dg)}">${esc(phone)}</a></span><span class="ev-acts"><a class="ev-act" href="tel:${esc(dg)}">전화</a><a class="ev-act" href="${esc(smsHref(dg, ""))}">문자</a></span></dd>`; })() : ""}
          ${addr ? `<dt>주소</dt><dd class="put-row"><span>${esc(addr)}</span>${putBtn(i, j, "addr")}</dd>` : ""}
          <dt id="mpl${i}-${j}" hidden>미팅장소</dt><dd id="mp${i}-${j}" class="put-row" hidden></dd>
        </dl>
        ${isRO(memoFcOf(c)) ? "" : `<div class="ev-memo"><input type="text" id="memo${i}-${j}" data-memoin="${i}-${j}" maxlength="200" placeholder="한 줄 메모 (상담기록에 저장)" autocomplete="off" enterkeyhint="done"><button type="button" class="mic-btn" data-mic="${i}-${j}" aria-label="음성으로 입력(녹음)"><svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="9" y="3" width="6" height="11" rx="3"/><path d="M5 11a7 7 0 0 0 14 0M12 18v3"/></svg></button><button type="button" class="memo-btn" data-memo="${i}-${j}" disabled>저장</button></div>`}
        <div id="cs${i}-${j}"></div></div>`;
  }
  function dayCardHtml(it, i) {
    const e = it.e;
    return `<div class="card ev-card" id="ev${i}"><div class="ev-top">${calDot(e)}<div class="ev-head"><span class="ev-time">${esc(timeText(e))}</span></div>
        ${e.allDay ? "" : salesOf(e) ? '<button type="button" class="ev-edit ev-sales" data-evunsales="${i}" title="누르면 비영업일정으로 바꿉니다">영업</button>' : `<button type="button" class="ev-edit ev-confirm" data-evconfirm="${i}">확정</button>`}<button type="button" class="ev-edit" data-evedit="${i}">수정</button><button type="button" class="ev-edit ev-del" data-evdel="${i}">삭제</button>
        <span class="ev-edit-acts" id="evActs${i}" hidden><button type="button" class="ev-edit" data-f-cancel>취소</button><button type="button" class="ev-edit ev-done" data-f-save>완료</button></span></div>
      <div class="ev-line2"><span class="ev-title-l"><span class="ev-title">${esc(e.title)}</span> <span class="ev-cal">- ${esc(e.cal.name)}</span></span>${it.all.length ? (it.ignored
        ? `<button type="button" class="ev-edit" data-unignore="${i}">되돌리기</button>`
        : `<button type="button" class="ev-edit" data-ignore="${i}">해당 없음</button>`) : ""}</div>
      ${it.ignored ? `<p class="hint ev-hid">검색된 고객 ${it.all.length}명을 숨겼어요</p>` : ""}
      <p class="hint ev-loc" id="loc${i}" ${e.location ? "" : "hidden"}><svg class="pin" viewBox="0 0 24 24" width="14" height="14" aria-hidden="true" style="vertical-align:-2px;margin-right:4px;fill:none;stroke:currentColor;stroke-width:1.5;stroke-linecap:round;stroke-linejoin:round"><path d="M20 10c0 6-8 12-8 12s-8-6-8-12a8 8 0 0 1 16 0z"/><circle cx="12" cy="10" r="3"/></svg>${esc(e.location)}</p>
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
  }
  function fillCust(it, c, i, j) {
    const x = c.extra || {};
    const mp = meetingOf(c, x);
    if (mp) {
      const dd = $(`mp${i}-${j}`);
      dd.innerHTML = `<span>${esc(mp)}</span>${putBtn(i, j, "mp")}`;
      dd.hidden = false; $(`mpl${i}-${j}`).hidden = false;
    }
    const cs = (x.consults || []).slice(0, 10);
    const box = $(`cs${i}-${j}`);
    if (box && cs.length) {
      box.innerHTML = `<details class="ev-consult"><summary>최근 상담 ${cs.length}건</summary>
        ${cs.map((r) => `<div class="consult-item"><div class="hint">${esc(r.date)}</div><div>${esc(r.text).replace(/\n/g, "<br>")}</div></div>`).join("")}</details>`;
    }
  }

  // 동명이인 '선택': 이 일정을 그 고객으로 묶음 → 구글 일정의 비공개 속성에 저장(모든 기기·사용자 공유) + 이 기기에도 기록
  async function pickCustomer(it, c, btn) {
    if (!c) return;
    const key = custKey(c);
    btn.disabled = true; btn.textContent = "저장 중";
    savePick(it.e.id, key);
    try {
      await ensureGWrite();
      await gcalPatch(it.e.calId, it.e.eventId, { extendedProperties: { private: { shlifeCust: key } } });
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
      await gcalPatch(it.e.calId, it.e.eventId, { extendedProperties: { private: { shlifeCust: on ? IGNORE_KEY : "" } } });
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
        await gcalPatch(it.e.calId, it.e.eventId, { extendedProperties: { private: { shlifeCust: "" } } });
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
      loc.innerHTML = `<svg class="pin" viewBox="0 0 24 24" width="14" height="14" aria-hidden="true" style="vertical-align:-2px;margin-right:4px;fill:none;stroke:currentColor;stroke-width:1.5;stroke-linecap:round;stroke-linejoin:round"><path d="M20 10c0 6-8 12-8 12s-8-6-8-12a8 8 0 0 1 16 0z"/><circle cx="12" cy="10" r="3"/></svg>${esc(text)}`; loc.hidden = false;
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
    box.innerHTML = `<div class="card ev-add"><div class="ev-form">
        <label class="fld"><span>제목</span><input data-a="title" autocomplete="off" placeholder="예: 김미례 고객 AP"></label>
        <label class="fld"><span>캘린더</span><select data-a="cal">${cals.map((c) => `<option value="${esc(c.id)}" ${c.id === def ? "selected" : ""}>${esc(c.name)}</option>`).join("")}</select></label>
        <label class="fld"><span>날짜</span><input data-a="date" type="date" value="${ymdStr(schedDate)}"></label>
        <label class="ev-allday"><input type="checkbox" class="cb" data-a="all"> 종일</label>
        <div class="fld-2" data-a-times><label class="fld"><span>시작</span><input data-a="start" type="time" value="${pad2(h0)}:00"></label>
          <label class="fld"><span>종료</span><input data-a="end" type="time" value="${pad2(Math.min(h0 + 1, 23))}:00"></label></div>
        <label class="fld"><span>장소</span><input data-a="loc" autocomplete="off"></label>
        <div class="file-actions"><button type="button" class="secondary" data-a-cancel>취소</button><button type="button" class="primary" data-a-save>등록</button></div>
      </div></div>`;
    box.hidden = false;
    const q = (k) => box.querySelector(`[data-a="${k}"]`);
    q("all").addEventListener("change", () => { box.querySelector("[data-a-times]").hidden = q("all").checked; });
    box.querySelector("[data-a-cancel]").addEventListener("click", closeSchedAdd);
    q("title").focus();
    box.querySelector("[data-a-save]").addEventListener("click", async (ev) => {
      const btn = ev.currentTarget, title = q("title").value.trim(), date = q("date").value, all = q("all").checked;
      if (!title) return toast("제목을 입력해 주세요.");
      if (!date) return toast("날짜를 입력해 주세요.");
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
    document.querySelectorAll("#dayList [data-evedit], #dayList [data-evdel]").forEach((b) => { b.hidden = on; });
    $("ev" + i).classList.toggle("editing", on);
  }
  function toggleEventForm(it, i) {
    const box = $("evForm" + i);
    if (box.innerHTML) { setEventEditing(i, false); return; }
    const e = it.e;
    const endDay = e.allDay ? addDays(e.end, -1) : e.end;
    box.innerHTML = `<div class="ev-form">
        <label class="fld"><span>캘린더</span><select data-f="cal">${schedCals.filter((c) => (GCFG.calendars || []).includes(c.name) || c.id === e.calId).map((c) =>
          `<option value="${esc(c.id)}" ${c.id === e.calId ? "selected" : ""}>${esc(c.name)}</option>`).join("")}</select></label>
        <label class="fld"><span>날짜</span><input data-f="date" type="date" value="${ymdStr(e.start)}"></label>
        ${e.allDay ? `<label class="fld"><span>종료 날짜</span><input data-f="endDate" type="date" value="${ymdStr(endDay)}"></label>`
          : `<div class="fld-2"><label class="fld"><span>시작</span><input data-f="start" type="time" value="${hm(e.start)}"></label>
             <label class="fld"><span>종료</span><input data-f="end" type="time" value="${hm(e.end)}"></label></div>`}
        <label class="fld"><span>장소</span><input data-f="location" value="${esc(e.location)}" autocomplete="off"></label>
        <label class="fld"><span>설명</span><textarea data-f="desc" rows="3">${esc(e.desc)}</textarea></label>
      </div>`;
    setEventEditing(i, true);
    { // 수정 모드에 들어가면 카드 위의 일정 제목이 바로 수정 칸이 됨
      const tl = $("ev" + i).querySelector(".ev-title-l");
      if (tl) { tl.dataset.orig = tl.innerHTML; tl.innerHTML = `<input data-f="title" class="ev-title-in" value="${esc(e.title)}" autocomplete="off" aria-label="일정 제목">`; }
    }
    const acts = $("evActs" + i);
    const val = (k) => ((box.querySelector(`[data-f="${k}"]`) || $("ev" + i).querySelector(`[data-f="${k}"]`)) || {}).value;
    acts.querySelector("[data-f-cancel]").onclick = () => setEventEditing(i, false);
    acts.querySelector("[data-f-save]").onclick = async (ev) => {
      const btn = ev.currentTarget;
      const patch = { summary: val("title").trim(), location: val("location").trim(), description: val("desc") };
      if (!val("date")) return toast("날짜를 입력해 주세요.");
      if (e.allDay) {
        const endD = val("endDate") || val("date");
        if (endD < val("date")) return toast("종료 날짜가 시작보다 빠릅니다.");
        patch.start = { date: val("date") };
        patch.end = { date: ymdStr(addDays(ymdToDate(endD), 1)) };
      } else {
        if (!val("start") || !val("end")) return toast("시간을 입력해 주세요.");
        const st = new Date(val("date") + "T" + val("start") + ":00");
        let en = new Date(val("date") + "T" + val("end") + ":00");
        if (en <= st) return toast("종료 시간이 시작보다 빠릅니다.");
        const tz = Intl.DateTimeFormat().resolvedOptions().timeZone || "Asia/Seoul";
        patch.start = { dateTime: st.toISOString(), timeZone: tz };
        patch.end = { dateTime: en.toISOString(), timeZone: tz };
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
      show(timed[0].i, `<span class="tl">${leave ? `<b>🏢 회사 출발 ${hm(leave)}</b>` : "🏢 회사에서"}</span><span class="tr">${legTxt(r)}</span>`); // 왼쪽: 회사 출발 / 오른쪽: 🚗 거리·시간
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
      show(last.i + 1, `<span class="tl">${back ? `<b>🏢 회사 도착 ${hm(back)}</b>` : ""}</span><span class="tr">${legTxt(r)}</span>`);
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
  function aiOpen() {
    if (!account() || aiIsOpen) return; // 로그인 전에는 열지 않음
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
  let aiPending = null;
  function aiShowMsg(html) { const m = $("aiMsg"); m.innerHTML = html; m.hidden = false; $("aiConfirm").hidden = true; aiPending = null; }
  async function aiSubmit() {
    const txt = $("aiInp").value;
    const it = aiInterpret(txt);
    if (!it) return aiShowMsg("못 알아들었습니다. 이렇게 말해 보세요.<br>• “홍길동 검색”<br>• “내일 오후 3시 홍길동 상담 일정 추가”<br>• “오늘 일정 보여줘”");
    if (it.kind === "needTitle") return aiShowMsg("무슨 일정인지 알려 주세요. 예: “내일 오후 3시 홍길동 상담 일정 추가”");
    $("aiMsg").hidden = true;
    let head, line1, line2 = "";
    if (it.kind === "search") { head = "고객 검색"; line1 = esc(it.name); }
    else if (it.kind === "showDay") { head = "일정 보기"; line1 = aiDateLabel(it.date); }
    else {
      const pre = it.mine ? "" : await aiPrefix(); // 내 일정에는 앞 글자를 붙이지 않음(팀 캘린더에서 누구 것인지 구분하려는 글자라서)
      it.summary = pre + it.title;
      if (it.mine) { it.calName = await myFcName(); if (!it.calName) return aiShowMsg("내 FC 이름이 아직 정해지지 않아 ‘내 일정’을 찾을 수 없습니다. 관리자에게 이름 등록을 요청해 주세요."); }
      else it.calName = (GCFG && GCFG.teamCalendar) || "1팀";
      const when = it.time ? `${it.time.h < 12 ? "오전" : "오후"} ${it.time.h % 12 || 12}:${pad2(it.time.min)}` : "종일";
      head = "일정 추가"; line1 = `${aiDateLabel(it.date)} ${when}`;
      line2 = `${esc(it.summary)} · ${esc(it.calName)} 캘린더` + (it.hasDate ? "" : " · 날짜를 말하지 않아 오늘로 했습니다");
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
        $("aiInp").value = ""; runTopSearch(it.name);
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
    else if (s === "sales") openSales();
    else resetSearch(); // 고객: 검색을 지우고 처음 화면
  }));
  window.addEventListener("popstate", () => { if (aiSkipOther) { aiSkipOther = false; return; } if (!$("customerView").hidden) { if (returnView === "scheduleView") openSchedule(); else if (returnView === "salesView") renderSales(SALES.cardAnchor || true); else show("searchView"); } });

  function resetSearch() {
    $("q").value = ""; $("q2").value = "";
    $("results").innerHTML = ""; $("searchStatus").textContent = "";
    $("recentList").hidden = true;
    show("searchView");
  }

  // (검색 가능한 FC 문구는 표시하지 않음)
  function showScope() { $("scopeHint").textContent = ""; }

  // ---------- 문자 문구 (영업 화면의 [문자] 버튼에 미리 채움) ----------
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
  const START_PAGES = [["schedule", "일정"], ["customer", "고객"], ["sales", "영업"]]; // 하단 메뉴 순서
  function startPrefKey() { return "shlife.start." + (gAccount() || "me"); }
  function readStartPref() {
    let v = null;
    try { v = JSON.parse(localStorage.getItem(startPrefKey()) || "null"); } catch (e) { v = null; }
    if (!v) { try { v = JSON.parse(localStorage.getItem("shlife.startAdmin." + me()) || "null"); } catch (e) { v = null; } } // 내가 안 골랐으면 관리자가 정한 값
    v = v || {};
    return { page: START_PAGES.some((p) => p[0] === v.page) ? v.page : "customer", ai: !!v.ai };
  }
  function writeStartPref(v) { try { localStorage.setItem(startPrefKey(), JSON.stringify(v)); return true; } catch (e) { return false; } }
  function startCardHtml() {
    const cur = readStartPref();
    return `<div class="card" id="startCard"><h3 class="sec-title">시작 화면</h3>
      <p class="hint">앱을 열면 처음 보일 화면을 고릅니다. 내 FC 폴더에 저장되어 관리자도 볼 수 있습니다.</p>
      <div class="st-opts" role="radiogroup" aria-label="시작 화면">${START_PAGES.map(([k, l]) =>
        `<label class="st-opt"><input type="radio" name="startPage" value="${k}" ${cur.page === k ? "checked" : ""}><span>${l}</span></label>`).join("")}</div>
      <label class="st-ai"><input type="checkbox" id="startAi" ${cur.ai ? "checked" : ""}><span>시작하면 AI 도우미를 바로 열기</span></label>
      <div class="file-actions"><button type="button" class="primary" id="startSave" disabled>저장</button></div></div>`;
  }
  function bindStartCard() {
    const box = $("startCard"); if (!box) return;
    const read = () => { const r = box.querySelector('input[name="startPage"]:checked'); return { page: r ? r.value : "customer", ai: $("startAi").checked }; };
    let base = JSON.stringify(read());
    const btn = $("startSave");
    box.addEventListener("change", () => { btn.disabled = JSON.stringify(read()) === base; });
    btn.addEventListener("click", async () => {
      const v = read(), ok = writeStartPref(v);
      let shared = false;
      if (!isAdmin()) { try { const perms = await readPerms(); shared = await writeStartShared(perms, me(), v, perms.fcFolders); } catch (e) { console.error(e); } }
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
    return v && typeof v === "object" && START_PAGES.some((x) => x[0] === v.page) ? { page: v.page, ai: !!v.ai } : null;
  }
  async function writeStartShared(perms, em, v, fcMap) {
    const d0 = startDirOf(perms, em, fcMap); if (!d0) return false;
    const d = await startSubDir(d0, true);
    await stWriteJsonIn(d, startFile(em), { page: v.page, ai: !!v.ai, by: me(), at: new Date().toISOString() });
    return true;
  }
  async function syncStartShared() { // 관리자·다른 폰에서 바꾼 값을 이 폰에 반영 (다음 실행부터 적용)
    try { const perms = await readPerms(); const v = await readStartShared(perms, me(), perms.fcFolders); if (v) writeStartPref(v); } catch (e) { /* 무시 */ }
  }
  // 관리자: 각 사용자 카드 안의 '시작 화면' (그 사용자 FC 폴더의 설정 파일에 저장 — 사용자가 직접 저장한 값도 여기에 보이고, 여기서 바꾸면 그 사용자 폰에 따라감)
  function startRowHtml(perms, em) {
    const st = (perms.starts || {})[em] || {}, pg = START_PAGES.some((p) => p[0] === st.page) ? st.page : "customer";
    return `<div class="perm-start" data-sa="${esc(em)}"><label class="pfx-l"><span>시작 화면</span><select data-sa-page aria-label="${esc(em)} 시작 화면">${START_PAGES.map(([k, l]) => `<option value="${k}" ${pg === k ? "selected" : ""}>${l}</option>`).join("")}</select></label>
      <label class="sa-ai"><input type="checkbox" data-sa-ai ${st.ai ? "checked" : ""}> AI 바로 열기</label><button type="button" class="secondary" data-sa-save="${esc(em)}" disabled>저장</button></div>`;
  }
  function bindStartRows() {
    const body = $("adminBody"), perms0 = adminState.perms;
    body.querySelectorAll("[data-sa]").forEach((box) => {
      const em = box.dataset.sa, sel = box.querySelector("[data-sa-page]"), ai = box.querySelector("[data-sa-ai]"), btn = box.querySelector("[data-sa-save]");
      const read = () => JSON.stringify([sel.value, ai.checked]);
      let base = read(), ready = false;
      readStartShared(perms0, em, adminState.fcIds).then((v) => { if (v) { sel.value = v.page; ai.checked = v.ai; } base = read(); ready = true; }).catch(() => { ready = true; });
      box.addEventListener("change", () => { btn.disabled = !ready || read() === base; });
      btn.addEventListener("click", async () => {
        btn.disabled = true; btn.textContent = "저장 중…";
        try {
          const perms = await readPerms(), v = { page: sel.value, ai: ai.checked };
          let ok = false;
          try { ok = await writeStartShared(perms, em, v, adminState.fcIds); } catch (e) { console.error(e); }
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
    return `<div class="card" id="smsCard"><details class="fold" open><summary class="sec-title">문자 문구</summary>
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

  // ---------- 영업 (이달이관 · 이달 따행) + 전화·문자 '시도' 기록 ----------
  //   이관: 계약 단위(계약 시트 최근통화일시·최근문자일시 + 이력 시트에 계약ID와 함께 한 줄) / 따행: 고객 단위(고객 시트 최근통화일시·최근문자일시 + 이력에 계약ID 비움)
  const SALES = { tab: "transfer", ym: "", showRef: false, data: null, df: { send: false, resp: false, coupon: false }, nf: { call: false, sms: false }, open: new Set(), dct: "", dcls: "", dact: true };
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
  const ymOfFollow = (v) => monthOf(v); // '2026-11' → '202611'

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
  function histHas(info, v) { // 같은 일시·고객ID·계약ID·유형이 이미 있으면 true (재시도 중복 방지)
    const cs = { t: info.colOf("일시"), u: info.colOf("고객ID"), k: info.colOf("계약ID"), y: info.colOf("유형") };
    if (cs.t < 0 || cs.u < 0) return false;
    for (let r = info.hr + 1; r <= info.rg.e.r; r++) {
      if (tsDigits(info.txt(r, cs.t)) === tsDigits(v.일시) && info.txt(r, cs.u) === v.고객ID && (cs.k < 0 ? "" : info.txt(r, cs.k)) === (v.계약ID || "") && (cs.y < 0 ? "" : info.txt(r, cs.y)) === v.유형) return true;
    }
    return false;
  }
  async function writeContacts(fileId, items) {
    const wb = XLSX.read(await stGet(fileId), { type: "array", cellNF: true }); // 최신 파일
    const stamp = tsDigits(nowStamp());
    const hist = sheetInfo(wb, "이력", "고객ID");
    const ct = sheetInfo(wb, "계약", "계약ID");
    const cu = sheetInfo(wb, CUST_SHEET, ID_COL);
    for (const it of items) {
      const colName = it.kind === "call" ? "최근통화일시" : "최근문자일시";
      const v = { 일시: it.ts, 고객ID: it.custId || "", 계약ID: it.cid || "", 유형: it.kind === "call" ? "통화" : "문자" };
      if (hist && !histHas(hist, v)) appendRow(hist, { 일시: { stamp: tsDigits(it.ts) }, 고객ID: v.고객ID, 계약ID: v.계약ID, 담당FC: it.fc || "", 유형: v.유형, 결과: "시도", 내용: it.src || "영업 화면", 입력일시: { stamp } });
      if (it.cid && ct) setRowCells(ct, "계약ID", it.cid, { [colName]: it.ts });
      if (!it.cid && it.custId && cu) setRowCells(cu, ID_COL, it.custId, { [colName]: it.ts });
    }
    await stPut(fileId, XLSX.write(wb, { type: "array", bookType: "xlsx", compression: true }), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    delete bookCache[fileId];
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
          await writeContacts(fileId, items);
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
    const [lists, contracts, hist, follow] = await Promise.all([loadLists(), loadContracts(), loadHistory().catch(() => []), loadFollow().catch(() => ({ files: new Set(), rows: [] }))]);
    const custById = new Map();
    lists.forEach((e) => { const id = idOf(e); if (id) (custById.get(id) || custById.set(id, []).get(id)).push(e); });
    return { lists, contracts, hist, custById, follow };
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
    const out = [];
    data.contracts.forEach((c) => {
      if (c.type !== "이관") return;
      const rr = c.row, idate = rr["이관일자"] || rr["이관일"] || "";
      if (monthOf(idate) !== ym) return;
      const ref = !!(c.fc && c.srcFc && c.fc !== c.srcFc); // 담당 FC가 이 파일의 FC가 아니면 참고
      if (ref && !showRef) return;
      const cid = String(rr["계약자ID"] || "").trim();
      const ce = cid && data.custById.get(cid) ? data.custById.get(cid).find((e) => e.fc === c.srcFc) || data.custById.get(cid)[0] : null;
      const holder = rr["계약자"] || (ce && ce.name) || "";
      const insured = rr["주피보험자"] || rr["피보"] || "";
      out.push({
        key: String(rr["계약ID"] || "").trim() ? "c:" + String(rr["계약ID"]).trim() : "", cid: String(rr["계약ID"] || "").trim(), custId: cid, fc: c.fc, srcFc: c.srcFc, ref,
        fileId: c.src && c.src.fileId, holder, ins: norm(insured) && norm(insured) !== norm(holder) ? insured : "",
        product: productName(rr["상품명"]), fee: rr["합계보험료"] || rr["보험료"] || "", cdate: rr["최초계약일"] || rr["계약일"] || "",
        bday: (ce && fieldOf(ce, "생년월일")) || rr["생일"] || "", phone: firstPhone((ce && (ce.phone || fieldOf(ce, "연락처"))) || rr["연락처"]),
        phoneText: (ce && (ce.phone || fieldOf(ce, "연락처"))) || rr["연락처"] || "", job: (ce && fieldOf(ce, "직업")) || rr["직업"] || "",
        addr: cleanAddr((ce && fieldOf(ce, "주소")) || rr["주소"]), recFc: rr["모집FC"] || "", status: rr["계약상태"] || rr["상태"] || "", idate
      });
    });
    out.sort((a, b) => a.holder.localeCompare(b.holder, "ko") || dateKey(b.cdate).localeCompare(dateKey(a.cdate)));
    return out;
  }
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
    follow.rows.filter((f) => ymOfFollow(f["대상년월"]) === ymOfFollow(ym)).forEach((f) => {
      const id = f["고객ID"], e = ownOf(id);
      if (!e || seen.has(id)) return;
      seen.add(id);
      const r = base(e, id);
      r.fileId = f._fileId; r.follow = f;
      r.dda = { done: f["방문완료일"] || (["방문완료", "설문발송", "설문응답", "쿠폰발송"].includes(String(f["접촉현황"] || "").trim()) ? "방문완료" : ""), send: f["설문발송일"], resp: f["설문응답일"], coupon: f["쿠폰발송일"], assign: "", apply: "", due: f["방문약속일"], gift: f["선물"], deliver: f["전달여부"], cls: f["따행분류"], active: f["캠페인대상여부"], contact: f["접촉현황"] };
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
      r.dda = { send: fieldOf(e, "설문발송일"), resp: fieldOf(e, "설문응답일"), coupon: fieldOf(e, "쿠폰발송일"), assign: fieldOf(e, "따행배정일"), apply: fieldOf(e, "따행고객신청"), due: fieldOf(e, "따행방문예정일"), done: fieldOf(e, "따행방문완료일"), gift: fieldOf(e, "따행선물"), deliver: fieldOf(e, "따행전달"), cls: "", active: "", contact: "" };
      out.push(r);
    });
    out.sort((a, b) => a.holder.localeCompare(b.holder, "ko") || dateKey(a.bday).localeCompare(dateKey(b.bday)) || dateKey(a.dda.assign).localeCompare(dateKey(b.dda.assign))); // 이름 오름차순(가나다), 같은 이름은 생년월일 → 배정일 빠른 순
    return out;
  }
  // 따행 필터: 활동대상(기본) → 접촉현황(방문완료 칸 값 유무) → 따행분류. totals는 활동대상 기준 전체(SOL 조회와 같은 73·16·57), classes는 접촉현황 필터 후의 분류별 건수
  function ddaFilter(baseRows, bMode, st) {
    const afterAct = baseRows.filter((r) => !bMode || !st.dact || !r.dda.active || r.dda.active === "활동대상");
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
    return `<div class="slist">` + rows.map((r, idx) => {
      const mk = mark.get(r.key) || {};
      const isD = !!r.dda, lock = isD && isRO(r.fc);
      const st = r.status && r.status !== "유지" ? `<span class="st">${esc(r.status === "납입면제" ? "납면" : r.status)}</span>` : "";
      const times = [mk.call ? `전화 시도 ✓ ${tsShow(mk.call)}` : "", mk.sms ? `문자 시도 ✓ ${tsShow(mk.sms)}` : ""].filter(Boolean).join(" · ");
      const chips = isD ? DCHIPS.map(([f, lab, cls]) => `<button type="button" class="dc ${cls} ${r.dda[f] ? "on" : "off"}${lock ? " lock" : ""}" data-chip="${f}" data-key="${esc(r.key)}" aria-pressed="${r.dda[f] ? "true" : "false"}">${lab}</button>`).join("") : st;
      const prodLines = isD
        ? (r.prods && r.prods.length ? r.prods : [{ name: "계약 정보 없음(조회 제한)", fee: "" }])
        : [{ name: r.product, fee: r.fee }];
      const detail = isD ? `<div class="sline sub detail">방문완료 ${esc(r.dda.done || "-")} · 설문발송일 ${esc(r.dda.send || "-")} · 설문응답일 ${esc(r.dda.resp || "-")} · 쿠폰발송일 ${esc(r.dda.coupon || "-")}</div>` : "";
      return `<div class="srow${r.ref ? " ref" : ""}${isD ? (r.dda.done ? " d-done" : " d-open") + " fold" + (openSet && openSet.has(r.key) ? " open" : "") : ""}" data-fold="${esc(r.key)}">
        <div class="sline l1"><span class="no">${idx + 1}</span><span class="nmw ${isD ? "n30" : "n50"} ${gradeClass(gradeOf({ entries: (SALES.data && SALES.data.custById && SALES.data.custById.get(r.custId)) || [] })) || "grade-good"}">${r.custId ? `<a href="#" data-cust="${esc(r.custId)}">${esc(r.holder)}</a>` : esc(r.holder)}${r.ins ? ` <small class="ins">→ ${esc(r.ins)}</small>` : ""}</span><span class="rt">${chips}</span></div>
        ${prodLines.map((p) => `<div class="sline prod"><span class="si pd">${esc(p.name)}</span><span class="si fe">${esc(p.fee ? wonText(p.fee) : "")}</span></div>`).join("")}
        <div class="sline l2">${it("", "계약일", esc(r.cdate))}${it("", "생일", esc(r.bday))}<span class="si ph">${esc(r.phoneText)} ${r.ref ? `<span class="hint">참고(기록 안 함)</span>` : ""}<span class="sbtns">${salesBtn(r, "call", mk, ctx)}${salesBtn(r, "sms", mk, ctx)}</span></span></div>
        <div class="sline l3">${it("", "직업", esc(r.job))}${it("", "모집", esc(r.recFc))}</div>
        <div class="sline l4">${it("ad", "", esc(r.addr))}</div>
        ${times ? `<div class="sline sub ok">${times}</div>` : ""}${detail}</div>`;
    }).join("") + `</div>`;
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
    const wb = XLSX.read(await stGet(row.fileId), { type: "array", cellNF: true }); // 저장 직전 최신 파일
    const cu = sheetInfo(wb, CUST_SHEET, ID_COL);
    if (!cu || !setTextCell(cu, ID_COL, row.custId, field, value)) throw new Error("고객 시트에서 이 고객 줄을 찾지 못했습니다.");
    const hist = sheetInfo(wb, "이력", "고객ID");
    if (hist) {
      const stamp = tsDigits(nowStamp());
      appendRow(hist, { 일시: { stamp }, 고객ID: row.custId, 계약ID: "", 담당FC: row.fc || "", 유형: "따행필드", 결과: value ? "입력" : "비움", 내용: `${field} ${value ? "입력 " + value : "비움(이전 " + (prev || "-") + ")"} (영업 화면 수동 변경)`, 입력일시: { stamp } });
    }
    await stPut(row.fileId, XLSX.write(wb, { type: "array", bookType: "xlsx", compression: true }), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
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
    const fo = row.follow, wb = XLSX.read(await stGet(row._fileId || fo._fileId), { type: "array", cellNF: true }); // 저장 직전 최신 파일
    const info = sheetInfo(wb, "따행", "고객ID");
    if (!info) throw new Error("따행 시트를 찾지 못했습니다.");
    const kc = info.colOf("고객ID"), mc = info.colOf("대상년월");
    let r0 = -1;
    for (let r = info.hr + 1; r <= info.rg.e.r; r++) if (info.txt(r, kc) === row.custId && ymOfFollow(info.txt(r, mc)) === ymOfFollow(fo["대상년월"])) { r0 = r; break; }
    if (r0 < 0) throw new Error("따행 시트에서 이 고객의 해당 대상년월 줄을 찾지 못했습니다.");
    Object.entries(cells).forEach(([col, v]) => { const c = ensureColumn(info, col); if (v === "") delete info.ws[info.enc(r0, c)]; else info.ws[info.enc(r0, c)] = { t: "s", v }; });
    const hist = sheetInfo(wb, "이력", "고객ID");
    if (hist) {
      const stamp = tsDigits(nowStamp());
      // 이력 내용은 '<따행 시트 열 이름> 입력|비움 … 대상 YYYY-MM (영업 화면 수동 변경)' 형식 (데이터 갱신 스크립트가 필드 이름으로 읽어 보호). 바뀐 열마다 한 줄
      logs.forEach(({ col, value, prev }) => appendRow(hist, { 일시: { stamp }, 고객ID: row.custId, 계약ID: "", 담당FC: row.fc || "", 유형: "따행필드", 결과: value ? "입력" : "비움", 내용: `${col} ${value ? "입력 " + value : "비움(이전 " + (prev || "-") + ")"} 대상 ${fo["대상년월"]} (영업 화면 수동 변경)`, 입력일시: { stamp } }));
    }
    await stPut(row._fileId || fo._fileId, XLSX.write(wb, { type: "array", bookType: "xlsx", compression: true }), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
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
    const wb = XLSX.read(await stGet(row.fileId), { type: "array", cellNF: true }); // 최신 파일
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
    await stPut(row.fileId, XLSX.write(wb, { type: "array", bookType: "xlsx", compression: true }), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
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
        const wb = XLSX.read(await stGet(fileId), { type: "array", cellNF: true });
        const info = { ct: sheetInfo(wb, "계약", "계약ID"), cu: sheetInfo(wb, CUST_SHEET, ID_COL), hist: sheetInfo(wb, "이력", "고객ID") };
        clearContacts(info, ts, ym);
        await stPut(fileId, XLSX.write(wb, { type: "array", bookType: "xlsx", compression: true }), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        delete bookCache[fileId];
      }
      const keys = new Set(rows.map((r) => r.key));
      writeQueue(readQueue().filter((it) => !(monthOf(it.ts) === ym && keys.has(it.cid ? "c:" + it.cid : "u:" + it.custId))));
      histCache = null; contractCache = null; listCache = null;
      toast("초기화했습니다.");
    } catch (e) { handleError(e); }
    renderSales(true);
  }

  // ---- 영업 화면 스크롤 복귀: 화면 맨 위에 보이는 행(key)과 그 행의 위치(오프셋)를 기준으로 저장·복원. 월·필터·탭·FC 변경은 맨 위 ----
  const SALES_SAVE_KEY = "shlife.salesResume";
  let SALES_BOOT = null; // 이 탭이 새로 로드됐을 때(브라우저가 탭을 내렸다가 다시 열 때 등) 직전 영업 화면 상태 (30분 이내, 같은 계정)
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
  function salesSaveState() { // 새로 로드돼도 이어 보도록 (영업 화면을 보고 있을 때만)
    if ($("salesView").hidden || SALES.resume) return;
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
  async function renderSales(keep) {
    const req = ++salesReq;
    let anchor = keep && keep !== true ? keep : keep ? salesAnchor() : null;
    if (!anchor && SALES.resume) anchor = SALES.resume;
    SALES.resume = null;
    show("salesView", !!keep);
    const box = $("salesView");
    if (!SALES.ym) SALES.ym = ymNow();
    const ym = SALES.ym;
    const top = (fcHtml) => `<div class="sales-top">
        <div class="seg"><button type="button" data-stab="transfer" class="${SALES.tab === "transfer" ? "on" : ""}">이달이관</button><button type="button" data-stab="dda" class="${SALES.tab === "dda" ? "on" : ""}">따행</button></div>
        <div class="sline2${fcHtml ? "" : " center"}"><div class="ymnav"><button type="button" data-ym="-1" aria-label="이전 달"><svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M15 6l-6 6 6 6"/></svg></button><b>${ymLabel(ym)}</b><button type="button" data-ym="1" aria-label="다음 달"><svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M9 6l6 6-6 6"/></svg></button>${top.note || ""}</div>${fcHtml}</div></div>`;
    if (!keep || !box.querySelector("#salesBody .slist, #salesBody .sales-info")) box.innerHTML = top("") + `<div id="salesBody"><p class="hint pad">불러오는 중…</p></div>`; // keep이면 기존 목록을 그대로 두고 새 내용으로 바꿈
    try {
      const data = await loadSalesData();
      if (req !== salesReq) return;
      await loadMyPrefs().catch(() => {});
      await refreshSalesExtra();
      const myFc = await myFcName();
      if (req !== salesReq) return;
      // FC 선택: 접근 가능한 FC들(고객이 등록된 FC). 기본은 내 FC(권한 파일 이름), 없으면 첫 FC. 선택은 이 기기에 기억
      const fcs = [...new Set([...data.lists.map((e) => e.fc), ...data.contracts.map((c) => c.srcFc)].filter(Boolean))].sort((a, b) => { const o = CFG.fcOrder || [], ia = o.indexOf(a), ib = o.indexOf(b); return (ia < 0 ? 999 : ia) - (ib < 0 ? 999 : ib) || a.localeCompare(b, "ko"); }); // 직급 먼저 → 사번순(config.fcOrder)
      let saved = ""; try { saved = localStorage.getItem(salesFcKey()) || ""; } catch (e) { /* 무시 */ }
      // 선택창에 나오는 FC: 관리자 = 모두 / 일반 계정 = 내 FC + 관리자가 허용한 추가 FC (고객 폴더 접근 권한이 있는 FC만)
      const selFcs = isAdmin() ? fcs : [...new Set([myFc, ...salesExtra])].filter((f) => f && fcs.includes(f));
      if (!isAdmin() && selFcs.length <= 1) SALES.fc = fcs.includes(myFc) ? myFc : fcs[0] || ""; // 추가 FC가 없으면 선택창 없이 본인 FC
      else if (!SALES.fc || !selFcs.includes(SALES.fc)) SALES.fc = selFcs.includes(saved) ? saved : selFcs.includes(myFc) ? myFc : selFcs[0] || "";
      const fcSel = SALES.fc;
      const baseRows = (SALES.tab === "transfer" ? transferRows(data, ym, SALES.showRef) : ddaRows(data, ym)).filter((r) => !fcSel || r.srcFc === fcSel);
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
      const ctx = { tpl: smsTemplate(SALES.tab === "dda" ? "dda" : "transfer"), fcName: fcSel || myFc };
      const holders = new Set(rows.map((r) => r.custId || r.holder)).size;
      const refN = SALES.tab === "transfer" ? data.contracts.filter((c) => c.type === "이관" && monthOf(c.row["이관일자"] || c.row["이관일"]) === ym && c.fc && c.srcFc && c.fc !== c.srcFc && (!fcSel || c.srcFc === fcSel)).length : 0;
      const marked = rows.filter((r) => r.key && mark.get(r.key)).length;
      SALES.data = data;
      top.note = ymNote;
      box.innerHTML = top(selFcs.length > 1 ? `<select id="salesFc" aria-label="FC 선택">${selFcs.map((f) => `<option value="${esc(f)}"${f === fcSel ? " selected" : ""}>${esc(f)}</option>`).join("")}</select>` : "") +
        `<div id="salesBody">${SALES.tab === "dda" ? "" : `<div class="sales-info"><span>${SALES.tab === "dda" ? (bMode ? `${(SALES.dct || SALES.dcls || dfOn) ? `필터 ${rows.length}명` : `표시 ${rows.length}명`}` : `전체 ${allRows.length}명(미완료 ${allRows.filter((r) => !(r.dda && r.dda.done)).length}명)${dfOn ? ` · 필터 ${rows.length}명` : ""}`) : `${rows.length}건 · ${holders}명`}${marked ? ` · 시도 ✓ ${marked}` : ""}${SALES.tab === "dda" ? ` <span class="hint">(시도 ✓ 기준 ${ymLabel(ym)})</span>` : ""}</span>${SALES.tab === "transfer" ? `<span class="sinfo-r">${refN ? `<label class="ref-toggle"><input type="checkbox" id="salesRef" ${SALES.showRef ? "checked" : ""}> 참고 포함(${refN})</label>` : ""}<button type="button" class="secondary reset-btn" id="salesReset"${marked || readQueue().length ? "" : " disabled"}>초기화</button></span>` : ""}</div>${nfHtml}`}` +
        (SALES.tab === "dda" ? `<div class="dfilter"><span class="dfl"><span class="dfg"><b class="dfk">설문</b><label><input type="checkbox" data-df="send" ${df.send ? "checked" : ""}> 발송</label><label><input type="checkbox" data-df="resp" ${df.resp ? "checked" : ""}> 응답</label></span><i class="dsep" aria-hidden="true"></i><span class="dfg"><label><input type="checkbox" data-df="coupon" ${df.coupon ? "checked" : ""}> 쿠폰</label></span><i class="dsep" aria-hidden="true"></i><span class="dfg">${nfBox("call", "전화")}${nfBox("sms", "문자")}</span></span></div>` : "") +
        (rows.length ? salesRowsHtml(rows, mark, ctx, SALES.open) : `<p class="hint pad">${SALES.tab === "transfer" ? ymLabel(ym) + " " : ""}${fcSel ? fcSel + " " : ""}${SALES.tab === "transfer" ? "이관" : "따행"} 고객이 없습니다.${nfOn ? " (전화·문자 필터를 확인해 주세요)" : ""}</p>`) +
        `<p class="hint pad">전화·문자는 버튼을 누른 '시도'만 기록됩니다(실제 통화·발송 여부는 알 수 없음).${readQueue().length ? ` 아직 저장 안 된 기록 ${readQueue().length}건은 앱이 다시 열리면 반영합니다.` : ""}</p></div>`;
      box._rows = rows; box._mark = mark;
      if (anchor) salesRestore(anchor); else if (!keep) window.scrollTo(0, 0);
    } catch (e) { if (req === salesReq) { $("salesBody").innerHTML = ""; handleError(e); } }
  }
  function openSales() { renderSales(); flushContacts(); }
  $("salesView").addEventListener("click", (ev) => {
    const chip = ev.target.closest("[data-chip]");
    if (chip) { // 칩 터치 = ON/OFF 전환만 (접힘/펼침·전화·문자와 별개)
      const row = (($("salesView")._rows) || []).find((r) => r.key === chip.dataset.key);
      if (row && row.dda) toggleChip(row, chip.dataset.chip);
      return;
    }
    const fold = ev.target.closest(".srow.fold");
    if (fold && !ev.target.closest("a, button, input, label, .sbtns, .sline.l1")) { // 칩이 있는 이름 줄은 접힘/펼침 제외 // 따행 행 접힘/펼침 (전화·문자 버튼과 이름 링크는 제외)
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
    if (st) { SALES.tab = st.dataset.stab; return renderSales(); }
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
      if (row.ref || !row.fileId || !row.key) { toast(row.ref ? "담당이 다른 FC 계약이라 기록하지 않습니다." : "새 구조 파일의 계약만 기록합니다."); return; }
      const q = readQueue();
      q.push({ uid: Date.now() + "-" + Math.random().toString(36).slice(2, 7), kind: ct.dataset.contact, ts: nowStamp(), cid: row.cid, custId: row.custId, fc: row.fc, fileId: row.fileId, src: SALES.tab === "dda" ? "영업 화면(따행)" : "영업 화면(이관)" });
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
    if (ev.target.id === "salesFc") { SALES.fc = ev.target.value; try { localStorage.setItem(salesFcKey(), SALES.fc); } catch (e) { /* 무시 */ } renderSales(); }
  });

  // ---------- 말만해요 연동 ----------
  //  주소 #go=customers&q=검색어 (고객 검색 목록) / #go=customer&id=고객키 (고객 정보). 검색어·고객 키는 # 뒤에 둬서 서버로 안 나감.
  //  ?embed=1 이면 말만해요 안에 모듈로 떠 있는 모드. 로그인 뒤에 이어서 처리하고, 처리하면 주소에서 # 뒤를 지움
  let GO0 = null, appStarted = false;
  try {
    const h = new URLSearchParams(location.hash.replace(/^#/, "")), go = h.get("go");
    if (go) {
      GO0 = { go, q: h.get("q") || "", id: h.get("id") || "" };
      history.replaceState(history.state, "", location.pathname + location.search); // 새로고침해도 다시 처리하지 않게 주소에서 지움
    }
  } catch (e) { GO0 = null; }
  async function goCustomer(id) {
    try {
      const gs = groupCustomers(await loadLists());
      const c = gs.find((x) => custKey(x) === id) || gs.find((x) => custKeyOld(x) === id);
      if (!c) { toast("고객을 찾지 못했습니다."); return; }
      c.entries.sort((a, b) => prio(a.type) - prio(b.type)); c.name = c.entries[0].name;
      openCustomer(c);
    } catch (e) { handleError(e); }
  }
  // 이미 앱 안에 있을 때는 새로 열지 않고 바로 처리. 아직 로그인 전이면 로그인 뒤에 처리
  window.blue10kpGo = function (o) {
    if (!o || (o.go !== "customers" && o.go !== "customer" && o.go !== "schedule")) return;
    if (!appStarted) { GO0 = { go: o.go, q: String(o.q || ""), id: String(o.id || "") }; return; }
    if (typeof helpClose === "function" && helpIsOpen) helpClose();
    if (!$("aiSheet").hidden && typeof aiClose === "function") aiClose();
    if (o.go === "schedule") { // q = YYYY-MM-DD (없거나 틀리면 오늘): 일정 화면의 하루 보기
      const m = String(o.q || "").match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
      const d = m ? new Date(+m[1], +m[2] - 1, +m[3]) : null;
      schedMode = "day"; schedDate = d && !isNaN(+d) ? dayStart(d) : dayStart(new Date());
      loadSettings().then(() => openSchedule());
      return;
    }
    if (o.go === "customer") { const id = String(o.id || ""); if (id) { if (!$("customerView").hidden) show("searchView"); goCustomer(id); } return; }
    const q = String(o.q || "").trim();
    if (!q) { show("searchView"); return; }
    goKeepList = q;
    runTopSearch(q);
  };
  // 말만해요 안에 모듈(iframe)로 떠 있을 때: 부모 말만해요에 지금 위치를 알림 (고객 이름이 든 값은 같은 주소(origin)에만 보냄)
  function notifyWhere(w) {
    if (!EMBED) return;
    try { parent.postMessage({ type: "blue10kp-where", app: "shlife", go: w.go, q: w.q || "", id: w.id || "", canBack: !!w.canBack }, location.origin); } catch (e) { /* 무시 */ }
  }
  const whereNow = () => {
    if (!$("customerView").hidden && current && current.cust) return { go: "customer", id: custKey(current.cust), canBack: true }; // 고객 정보에서는 앱 안 뒤로(목록)가 있음. 검색 목록이 첫 화면이면 canBack=false
    if (!$("searchView").hidden) return { go: "customers", q: curQuery() };
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
  });

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
  if (GD) document.addEventListener("visibilitychange", () => { if (document.visibilityState === "visible") checkSharedLogin(); });
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

  function startApp() {
    appStarted = true;
    const goReq = GO0; GO0 = null;
    if (goReq) { // 말만해요에서 넘어옴: 첫 화면은 고객, 이어서 검색
      schedMode = "day"; schedDate = dayStart(new Date());
      resetSearch();
      setTimeout(() => window.blue10kpGo(goReq), 0);
    } else if (!GD && sessionStorage.getItem("openAdmin") && isAdmin()) {
      sessionStorage.removeItem("openAdmin");
      openAdmin();
    } else if (SALES_BOOT) { // 새로 로드됐을 때 영업 화면을 보던 중이었으면 그 탭·월·FC·필터·위치로
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
      const sp = readStartPref().page; // 설정 > 시작 화면
      if (sp === "schedule") loadSettings().then(() => openSchedule());
      else if (sp === "sales") loadSettings().then(() => openSales());
      else {
        resetSearch(); // 첫 화면: 고객 (일정은 하단 메뉴에서 열 때 불러옴 — 캘린더 호출이 첫 화면을 늦추지 않게)
        $("searchStatus").textContent = "고객 목록을 불러오는 중…";
      }
    }
    if (!isAdmin()) syncStartShared();
    if (readStartPref().ai && !EMBED) setTimeout(aiOpen, 400); // 설정 > 시작 화면 > AI 도우미 바로 열기
    setTimeout(askContactNote, 1500); // 전화·문자 중에 앱이 닫혔다가 다시 열린 경우
    loadSettings(); // 관리자 설정 (회사 주소, 데이터 절약 모드 등)
    if (!isAdmin()) loadMyPrefs(); // 일반 사용자 설정 (캘린더·데이터 절약)
    loadLists().then((recs) => { renderFcOpts(recs).catch(() => {}); showScope(); if ($("searchStatus").textContent === "고객 목록을 불러오는 중…") $("searchStatus").textContent = ""; }).catch(handleError); // 미리 불러오기
  }
})();
