(function () {
  "use strict";
  const CFG = window.APP_CONFIG;
  // Files.Read.All: 내 파일 + 다른 사람이 공유해 준 폴더까지 읽기 (읽기 전용)
  const SCOPES = ["Files.Read.All", "User.Read"];
  const GRAPH = "https://graph.microsoft.com/v1.0";
  const $ = (id) => document.getElementById(id);

  // ---------- 로그인 (MSAL) ----------
  const redirectUri = location.origin + location.pathname.replace(/index\.html$/, "");
  const msal = new window.msal.PublicClientApplication({
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
  const account = () => msal.getAllAccounts()[0];
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

  // ---------- 화면 전환 ----------
  const views = ["loginView", "scheduleView", "searchView", "customerView", "adminView"];
  const APP_NAME = "신한FC";
  const TITLES = { scheduleView: "일정", searchView: "고객", customerView: "고객 정보", adminView: "관리자" };
  const SECTION = { scheduleView: "schedule", searchView: "customer", customerView: "customer", adminView: "admin" };
  function show(view) {
    views.forEach((v) => ($(v).hidden = v !== view));
    // 고객 정보 화면: "‹ 고객 정보" 전체가 뒤로 가기 버튼
    $("backBtn").hidden = view !== "customerView";
    $("title").hidden = view === "customerView";
    $("logoutBtn").hidden = view === "loginView";
    $("topSearch").hidden = view === "loginView"; // 상단 검색창은 로그인 후 항상
    $("title").textContent = TITLES[view] || APP_NAME;
    $("mainNav").hidden = view === "loginView";
    $("navAdmin").hidden = !isAdmin();
    $("mainNav").querySelectorAll("button").forEach((b) => b.classList.toggle("on", b.dataset.section === SECTION[view]));
    window.scrollTo(0, 0);
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
    try {
      const path = "/me/drive/root:/" + CFG.baseFolder.split("/").map(encodeURIComponent).join("/");
      const it = await graph(path + "?$select=id,parentReference");
      return (base = { driveId: it.parentReference.driveId, id: it.id, owner: true });
    } catch (e) {
      if (!String(e.message).includes("404")) throw e;
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
  async function resolveAppDir() {
    if (appDir) return appDir;
    await resolveBase();
    try {
      const path = "/me/drive/root:/" + CFG.appFolder.split("/").map(encodeURIComponent).join("/");
      const it = await graph(path + "?$select=id,parentReference");
      return (appDir = { driveId: it.parentReference.driveId, id: it.id });
    } catch (e) {
      if (!String(e.message).includes("404")) throw e;
    }
    return (appDir = { driveId: CFG.owner.driveId, id: CFG.owner.appDirId });
  }
  const appFilePath = (name) => "/drives/" + appDir.driveId + "/items/" + appDir.id + ":/" + encodeURIComponent(name) + ":/content";

  // ---------- 관리자 설정 (앱 저장 폴더의 설정 파일, 모든 사용자가 같은 설정 사용) ----------
  //   { companyAddress, dataSaver, google: { account, calendars: [] } }
  const SETTINGS_FILE = CFG.settingsFile || "shlife-설정.json";
  let appSettings = { companyAddress: "", dataSaver: false, google: { account: "", calendars: [] } };
  const settingsPath = () => appFilePath(SETTINGS_FILE);
  async function loadSettings() {
    try {
      await resolveAppDir();
      const s = await graph(settingsPath() + "?x=" + Date.now());
      appSettings = { ...appSettings, ...s, google: { ...appSettings.google, ...(s.google || {}) } };
    } catch (e) { /* 설정 파일이 아직 없으면 기본값 */ }
    return appSettings;
  }
  async function saveSettings(next) {
    await resolveAppDir();
    await gx("PUT", settingsPath(), next);
    appSettings = next;
  }
  const permPath = () => appFilePath(PERM_FILE);
  const me = () => String((account() || {}).username || "").toLowerCase();

  async function readPerms() {
    await resolveAppDir();
    try {
      const p = await graph(permPath() + "?x=" + Date.now());
      return { users: (p && p.users) || {} };
    } catch (e) {
      if (String(e.message).includes("404")) return { users: {} };
      throw e;
    }
  }

  // 지금 로그인한 아이디가 볼 수 있는 FC: "*" 또는 FC 이름 목록
  async function myAllowedFcs() {
    if (isAdmin() || base.owner) return "*";
    const perms = await readPerms();
    const list = perms.users[me()] || [];
    return list.includes("*") ? "*" : list;
  }

  // 고객 폴더 구조 읽기
  //   고객/김현성/_계약/<고객폴더>, 고객/김현성/계약고객리스트.xlsx   ← FC별 폴더 (현재 구조)
  //   고객/_계약/<고객폴더>, 고객/계약고객리스트.xlsx               ← FC 폴더 없이 (예전 구조)
  // "_"로 시작하는 폴더 = 분류(계약/이관/진행), 그 안의 폴더 = 고객 폴더
  let scanCache = null; // Promise<{ folders, files, fcs }>
  function scanBase() {
    if (scanCache) return scanCache;
    scanCache = (async () => {
      await resolveBase();
      const allowed = await myAllowedFcs();
      const ok = (name) => allowed === "*" || allowed.includes(name);
      const folders = [], files = [];
      async function scan(id, fc) {
        const items = await listChildren(itemPath(id) + "/children");
        // FC 폴더 바깥(고객 폴더 바로 아래) 파일·분류폴더는 '전체' 권한만
        const rootOk = !!fc || allowed === "*";
        if (rootOk) items.filter((i) => i.file && i.name !== PERM_FILE).forEach((f) => files.push({ ...f, fc }));
        await Promise.all(items.filter((i) => i.folder).map(async (d) => {
          if (d.name.startsWith("_")) {
            if (!rootOk) return;
            (await listChildren(itemPath(d.id) + "/children"))
              .filter((i) => i.folder)
              .forEach((f) => folders.push({ ...f, group: d.name.slice(1), fc }));
          } else if (!fc && ok(d.name)) {
            await scan(d.id, d.name); // 허용된 FC 폴더만
          }
        }));
      }
      try {
        await scan(base.id, "");
      } catch (e) {
        if (!base.owner && /40[34]/.test(String(e.message))) throw new Error("NO_SHARE");
        throw e;
      }
      const fcs = [];
      folders.forEach((f) => f.fc && !fcs.includes(f.fc) && fcs.push(f.fc));
      files.forEach((f) => f.fc && !fcs.includes(f.fc) && fcs.push(f.fc));
      return { folders, files, fcs: fcs.sort((a, b) => a.localeCompare(b, "ko")), all: allowed === "*" };
    })();
    scanCache.catch(() => { scanCache = null; });
    return scanCache;
  }
  const loadCustomerFolders = async () => (await scanBase()).folders;

  // ---------- 고객 리스트 엑셀 (계약고객리스트 / 이관고객리스트) ----------
  const bookCache = {};
  async function readBook(id) {
    if (!bookCache[id]) {
      bookCache[id] = graph(itemPath(id) + "/content", true).then((buf) => XLSX.read(buf, { type: "array" }));
    }
    return bookCache[id];
  }

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
    const { files, folders } = await scanBase();
    const sources = CFG.lists.filter((l) => l.fileRegex || l.fileKeyword).flatMap((l) =>
      pickSourceFiles(files, l).map((f) => ({ ...l, file: f })));
    const out = [];
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
      sheetsOf(wb, s).forEach((sheet) => {
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
            type: s.label, name, fc, sheet,
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
    const byFile = new Map();
    entries.filter((e) => e.src).forEach((e) => { if (!byFile.has(e.src.fileId)) byFile.set(e.src.fileId, []); byFile.get(e.src.fileId).push(e); });
    for (const [fileId, list] of byFile) {
      const wb = XLSX.read(await graph(itemPath(fileId) + "/content", true), { type: "array" }); // 최신 파일
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
          if (c < 0) { c = rg.e.c + 1; ws[enc(hr, c)] = { t: "s", v: k }; rg.e.c = c; } // 새 열
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
      await gx("PUT", itemPath(fileId) + "/content", XLSX.write(wb, { type: "array", bookType: "xlsx" }),
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      delete bookCache[fileId];
    }
  }

  // ---------- 검색 ----------
  let mode = "all";
  const norm = (s) => String(s).replace(/\s+/g, "").toLowerCase();
  const custPart = (n) => n.split(" - ")[0]; // 폴더명 "고객명 - 담당자명"

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

  async function search(q) {
    q = q.trim();
    const status = $("searchStatus");
    const list = $("results");
    list.innerHTML = "";
    if (!q) { status.textContent = ""; return; }
    status.textContent = "찾는 중…";
    try {
      const records = await loadLists();
      const want = CFG.lists.find((l) => l.id === mode);
      if (want && !records.some((r) => r.type === want.label)) {
        status.textContent = want.folderGroup
          ? `'_${want.folderGroup}' 폴더에 고객 폴더가 없습니다.`
          : `'${want.label}' 엑셀 파일이 없습니다. FC 폴더에 이름에 '${want.fileKeyword}'가 들어간 엑셀을 넣어 주세요.`;
        return;
      }
      // 고객 이름 · 연락처 · 주소 · 미팅장소로 검색
      const nq = norm(q), dq = digits(q);
      const meets = await loadMeetingPlaces().catch(() => []);
      const meetOf = (r) => meets.filter((m) => norm(m.name) === norm(r.name) && sameFc(m.fc, r.fc)).map((m) => m.mp).join(" ");
      const hits = records.filter((r) => {
        if (want && r.type !== want.label) return false;
        if (norm(r.name).includes(nq)) return true;
        if (dq.length >= 3 && dq === q.replace(/[\s-]/g, "") && digits(r.phone).includes(dq)) return true;
        if (nq.length < 2) return false;
        return norm(entryAddr(r)).includes(nq) || norm(meetOf(r)).includes(nq);
      });
      // 같은 이름 + 같은 생년월일은 한 고객으로 묶기
      const map = new Map();
      hits.forEach((r) => {
        // 생년월일 형식 통일: "1980-10-21" → 801021, "851002-1" → 851002
        const d = String(r.birth).replace(/\D/g, "");
        const ymd = /^(19|20)\d{6}/.test(d) ? d.slice(2, 8) : d.slice(0, 6);
        const key = norm(r.name) + "|" + ymd;
        if (!map.has(key)) map.set(key, { name: r.name, birth: r.birth, entries: [] });
        map.get(key).entries.push(r);
      });
      const custs = [...map.values()].sort((a, b) => a.name.localeCompare(b.name, "ko"));
      const label = want ? want.label + " " : "";
      status.textContent = custs.length ? label + custs.length + "명 찾음" : "'" + q + "' " + label + "고객이 없습니다.";
      if (custs.length === 1) return openCustomer(custs[0]);
      list.innerHTML = custs.map((c, i) => {
        const types = [...new Set(c.entries.map((e) => e.type))];
        const fcs = [...new Set(c.entries.map((e) => e.fc).filter(Boolean))];
        return `<li><button data-i="${i}"><span class="avatar sm">${esc(c.name[0])}</span>
          <span class="r-name"><span><span class="cname ${gradeClass(gradeOf(c))}">${esc(c.name)}</span>${regionOf(c) ? ` <span class="region">(${esc(regionOf(c))})</span>` : ""}</span><span class="r-sub">${esc([c.birth, fcs.join(", ")].filter(Boolean).join(" · "))}</span></span>
          ${types.map((t) => `<span class="tag t-${esc(t)}">${esc(t)}</span>`).join("")}<span class="chev">›</span></button></li>`;
      }).join("");
      list.querySelectorAll("button").forEach((b) =>
        b.addEventListener("click", () => openCustomer(custs[+b.dataset.i])));
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
      const { files } = await scanBase();
      const sources = (CFG.contractLists || []).flatMap((l) =>
        pickSourceFiles(files, l).map((f) => ({ ...l, file: f })));
      const out = [];
      await Promise.all(sources.map(async (s) => {
        const wb = await readBook(s.file.id);
        const fc = s.file.fc || (s.file.name.split(" - ")[1] || "").trim();
        sheetsOf(wb, s).forEach((sheet) => {
          const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheet], { header: 1, defval: "", raw: false });
          const h = rows.findIndex((r) => r.some((c) => String(c).trim() === "계약자"));
          if (h < 0) return;
          const head = rows[h].map((c) => String(c).trim());
          rows.slice(h + 1).forEach((r) => {
            const fields = head.map((k, i) => [k, String(r[i] ?? "").trim()]).filter(([k, v]) => k && v);
            if (!fields.length) return;
            const row = Object.fromEntries(fields);
            const names = CONTRACT_NAME_COLS.map((c) => row[c]).filter(Boolean).map(norm);
            // 줄에 담당FC가 있으면 그 FC로 표시 (다른 FC 파일에 같이 들어 있는 계약)
            const rowFc = String(row["담당FC"] || row["담당 FC"] || "").trim() || fc;
            if (names.length) out.push({ type: s.label, fc: rowFc, srcFc: fc, sheet, row, fields, names });
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
    const items = await listChildren(itemPath(base.id) + "/children");
    const file = items.find((i) => i.file && /\.xlsx$/i.test(i.name) && nospace(i.name).includes(nospace(EXTRA.fileKeyword)));
    if (!file) return null;
    const wb = XLSX.read(await graph(itemPath(file.id) + "/content", true), { type: "array" });
    const sheetName = wb.SheetNames.includes(EXTRA.sheet) ? EXTRA.sheet : wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: false, blankrows: true });
    const head = (rows[0] || []).map((c) => String(c).trim());
    return { file, wb, ws, rows, head };
  }

  const extraCols = (x) => x.head.filter((h) => h && h !== EXTRA.fcCol && h !== EXTRA.nameCol);
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

  async function loadExtraFor(fc, name) {
    const x = await readExtra();
    if (!x) return { missing: true, fields: [], values: {}, consults: [] };
    const r = extraRowIndex(x, fc, name);
    const fields = extraCols(x);
    const values = {};
    fields.forEach((k) => { const c = x.head.indexOf(k); values[k] = r > 0 ? String(x.rows[r][c] ?? "").trim() : ""; });
    return { found: r > 0, fields, values, fileName: x.file.name, consults: readConsults(x.wb, fc, name) };
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
    const out = XLSX.write(x.wb, { type: "array", bookType: "xlsx" });
    await gx("PUT", itemPath(x.file.id) + "/content", out,
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    delete bookCache[x.file.id];
  }

  async function saveExtraFor(fc, name, values, before) {
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
    if (isNew) { setCell(EXTRA.fcCol, fc); setCell(EXTRA.nameCol, name); }
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
      const buf = await graph(itemPath(id) + "/content", true);
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
      `<button type="button" class="grade-opt${x.id === g ? " on" : ""}" data-grade="${x.id}" aria-label="성향 ${x.label}"><span class="grade-dot" style="background:${x.color}"></span><span>${x.label}</span></button>`).join("") : "";
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
    returnView = $("scheduleView").hidden ? "searchView" : "scheduleView";
    show("customerView");
    $("custName").textContent = cust.name;
    editMode = false;
    renderGrade(cust);
    $("avatar").textContent = cust.name[0];
    $("avatar").style.backgroundImage = "";
    editMode = false;
    $("editBtn").textContent = "수정";
    $("editBtn").classList.remove("on");
    $("photoMenu").hidden = true;
    $("avatar").classList.remove("editable");
    consultShowAll = false; // 상담기록은 처음엔 최근 10건
    $("avatar").classList.remove("has-photo");
    $("custSub").textContent = custSubText(cust);
    $("tabBody").innerHTML = '<p class="hint pad">고객 정보를 불러오는 중…</p>';
    $("tabs").innerHTML = "";
    history.pushState({ v: "customer" }, "");
    current = { cust, folders: [], files: [], phones: null, contracts: [], fc: "", extra: { values: {} }, profilePhoto: null, photoUrl: null };
    try {
      const byName = (await loadCustomerFolders()).filter((f) => norm(custPart(f.name)) === norm(cust.name));
      // 같은 이름이 여러 FC에 있으면 이 고객의 담당 FC 폴더만
      const fcs = cust.entries.map((e) => e.fc).filter(Boolean);
      const sameFc = byName.filter((f) => f.fc && fcs.includes(f.fc));
      const folders = sameFc.length ? sameFc : byName;
      let files = [];
      for (const f of folders) files = files.concat(await collectFiles(f.id, folders.length > 1 ? f.name + "/" : "", 0));
      current.folders = folders;
      current.files = files;
      current.profilePhoto = findProfilePhoto(files);
      if (current.profilePhoto && !appSettings.dataSaver) loadProfilePhoto(); // 머리 부분 동그라미에 사진 표시 (데이터 절약 모드면 안 함)
      current.fc = fcs[0] || (folders[0] && folders[0].fc) || "";
      const [contracts, extra] = await Promise.all([
        loadContracts().catch((e) => { console.error(e); return []; }),
        loadExtraFor(current.fc, cust.name).catch((e) => { console.error(e); return { error: true, values: {} }; })
      ]);
      current.contracts = contracts.filter((c) => c.names.includes(norm(cust.name)) && (!fcs.length || fcs.includes(c.fc) || fcs.includes(c.srcFc)));
      current.extra = extra;
    } catch (e) {
      handleError(e);
    }
    renderTabs();
    selectTab("basic");
  }

  // 고객 폴더 안 파일 (하위 폴더 2단계까지)
  async function collectFiles(id, prefix, depth) {
    const items = await listChildren(itemPath(id) + "/children");
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
    return current.files.filter((f) => kws.some((k) => f.path.toLowerCase().includes(k)) && !(ex && ex.test(f.name)));
  }

  function renderTabs() {
    $("tabs").innerHTML = CFG.tabs.map((t) => {
      const n = t.id === "contact" || t.id === "all" || t.id === "basic" ? "" :
        t.id === "consult" ? ((current.extra && current.extra.consults) || []).length :
        filesForTab(t).length + (t.id === "contract" ? current.contracts.length : 0);
      return `<button data-tab="${t.id}">${esc(t.label)}${n !== "" ? `<small>${n}</small>` : ""}</button>`;
    }).join("");
    $("tabs").querySelectorAll("button").forEach((b) =>
      b.addEventListener("click", () => selectTab(b.dataset.tab)));
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

    const files = filesForTab(tab);
    const listHtml = id === "basic" ? basicHtml() : id === "contract" ? contractHtml() : "";
    if (!files.length) {
      body.innerHTML = listHtml; // 파일이 없으면 안내 문구 없이 비워 둠
      if (id === "basic") bindExtra();
      return;
    }
    body.innerHTML = listHtml + files.map((f, i) => {
      const k = kind(f.name);
      return `<div class="card file" data-i="${i}">
        <div class="file-row">
          <span class="badge ${k}">${KIND_LABEL[k]}</span>
          <div class="file-name">${fileLabel(f)}</div>
        </div>
        <div class="file-actions">
          ${k === "excel" ? `<button class="secondary" data-act="view">내용 보기</button>` : ""}
          ${k === "pdf" ? `<button class="secondary" data-open="${esc(f.id)}">열기</button>`
            : `<a class="secondary" href="${esc(f.webUrl)}" target="_blank" rel="noopener">열기</a>`}
        </div>
        <div class="sheet" hidden></div>
      </div>`;
    }).join("");
    body.querySelectorAll('[data-act="view"]').forEach((b) =>
      b.addEventListener("click", () => {
        const card = b.closest(".file");
        toggleSheet(card, files[+card.dataset.i], b);
      }));
    // 엑셀이 하나뿐이면 바로 펼치기
    const excels = body.querySelectorAll('[data-act="view"]');
    if (excels.length === 1) excels[0].click();
    if (id === "basic") bindExtra();
  }

  // 정보 탭
  //   보기: 연락처 · (이메일) · 주소 · 통신사/차량번호/자동차보험갱신일 · (그 밖의 고객리스트 항목) / 미팅장소 · 메모
  //   수정(맨 위 '수정' 버튼): 같은 항목을 한 화면에서 고치고 저장
  //     연락처·주소·이메일 → 고객리스트 엑셀 + '우선' 1 / 나머지 → 고객추가정보 엑셀
  const BASIC_EXCLUDE = ["상품명 - 보험료", "상품명", "보험료", "이관월", "고객 폴더", "분류", "No", "순번", "메모", "모집FC", "담당FC", "담당 FC", "우편번호", "성별", "폴더", "성향", "우선"];
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
  const MOBILE_RE = /01[016789][-.\s]?\d{3,4}[-.\s]?\d{4}/;
  function valueWithActions(k, v) {
    let btns = "";
    const m = PHONE_COLS.includes(k.trim()) && String(v).match(MOBILE_RE);
    if (m) {
      const num = m[0].replace(/\D/g, "");
      btns = `<a class="act-btn" href="tel:${num}" aria-label="전화 걸기">📞<span>Call</span></a>
        <a class="act-btn" href="sms:${num}" aria-label="문자 보내기">💬<span>SMS</span></a>`;
    } else if (/이메일|e-?mail/i.test(k) && /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(String(v).trim())) {
      btns = `<a class="act-btn" href="mailto:${esc(String(v).trim())}" aria-label="메일 보내기">✉️<span>@</span></a>`;
    }
    return btns ? `<div class="val-row"><span class="val ${m ? "nowrap" : ""}">${esc(v)}</span><span class="acts">${btns}</span></div>`
      : `<div>${esc(v)}</div>`;
  }

  // 맨 위: 생년월일 (남) · 계약 · 담당: 최수진FC
  function custSubText(c) {
    const g = genderOf(c);
    const types = [...new Set(c.entries.map((e) => e.type))];
    const fcs = [...new Set(c.entries.map((e) => e.fc).filter(Boolean))];
    return [c.birth ? c.birth + (g ? ` (${g})` : "") : "", types.join(" · "), fcs.length ? "담당: " + fcs.map((f) => f + "FC").join(", ") : ""]
      .filter(Boolean).join(" · ");
  }

  const mdText = (v) => { const p = parseMonthDay(v); return p ? `${p[0]}월 ${p[1]}일` : String(v || ""); };

  function basicHtml() {
    return editMode ? basicFormHtml() : basicViewHtml();
  }

  function basicViewHtml() {
    const c = current.cust;
    const x = current.extra || { values: {}, fields: [] };
    const ev = x.values || {};
    // 고객리스트 항목 (같은 항목은 한 번만)
    const merged = new Map();
    c.entries.forEach((e) => e.fields.forEach(([k, v]) => {
      const key = k.trim();
      if (BASIC_EXCLUDE.includes(key) || NAME_COLS.includes(key) || BIRTH_KEYS.includes(key)) return;
      if (!merged.has(key)) merged.set(key, []);
      if (!merged.get(key).includes(v)) merged.get(key).push(v);
    }));
    const take = (cols) => { const out = []; cols.forEach((k) => { (merged.get(k) || []).forEach((v) => out.push([k, v])); merged.delete(k); }); return out; };
    const phones = take(PHONE_COLS), emails = take(EMAIL_COLS), addrs = take(["주소"]);
    const row = (label, html) => `<dt>${esc(label)}</dt><dd>${html}</dd>`;
    // 통신사 / 차량번호 / 자동차보험갱신일 (+ 미팅장소·메모가 아닌 추가 항목) — 주소 밑 한 줄
    const extraKeys = (x.fields || []).filter((k) => !EXTRA.primary.includes(k));
    const chips = extraKeys.filter((k) => (ev[k] || "").trim()).map((k) =>
      `<span class="chip"><small>${esc(k === "자동차보험갱신일" ? "보험갱신" : k)}</small> ${esc(EXTRA.monthDayCols.includes(k) ? mdText(ev[k]) : ev[k])}</span>`).join("");
    const rest = [...merged];
    const info = `<div class="card" id="infoCard"><dl class="kv">
        ${phones.length ? row("연락처", phones.map(([k, v]) => valueWithActions(k, v)).join("")) : row("연락처", '<span class="hint">-</span>')}
        ${emails.map(([k, v]) => row("이메일", valueWithActions(k, v))).join("")}
        ${addrs.length ? row("주소", addrs.map(([, v]) => `<div>${esc(v)}</div>`).join("")) : ""}
        ${chips ? row("추가", `<div class="chips">${chips}</div>`) : ""}
        ${rest.map(([k, vs]) => row(k, vs.map((v) => `<div>${esc(v)}</div>`).join(""))).join("")}
      </dl></div>`;
    const prim = (x.fields || []).filter((k) => EXTRA.primary.includes(k));
    const memo = x.missing ? `<div class="card"><p class="hint">'${esc(EXTRA.fileKeyword)}' 엑셀 파일을 찾지 못했습니다.</p></div>`
      : prim.length ? `<div class="card"><dl class="kv">${prim.map((k) => row(k, (ev[k] || "").trim()
        ? `<div class="pre">${esc(ev[k])}</div>` : '<span class="hint">-</span>')).join("")}</dl></div>` : "";
    return info + memo;
  }

  // 정보 수정: 연락처·주소·이메일 → 고객리스트 엑셀 + '우선' 1
  const EDIT_FIELDS = [
    { label: "연락처", cols: PHONE_COLS, type: "tel" },
    { label: "주소", cols: ["주소"], type: "text", multi: true },
    { label: "이메일", cols: EMAIL_COLS, type: "email" }
  ];
  function currentValue(c, cols) {
    for (const e of c.entries) for (const k of cols) { const f = e.fields.find(([fk]) => fk.trim() === k); if (f && f[1]) return f[1]; }
    return "";
  }

  function basicFormHtml() {
    const c = current.cust;
    const x = current.extra || { values: {}, fields: [] };
    const ev = x.values || {};
    const canList = c.entries.some((e) => e.src);
    const listFields = canList ? EDIT_FIELDS.map((f) => `<label class="fld"><span>${esc(f.label)}</span>
        ${f.multi ? `<textarea data-edit="${esc(f.label)}" rows="2">${esc(currentValue(c, f.cols))}</textarea>`
          : `<input data-edit="${esc(f.label)}" value="${esc(currentValue(c, f.cols))}" type="${f.type}" autocomplete="off">`}</label>`).join("") : "";
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
      else if (k === EXTRA.carrierCol) input = `<input data-ex="${esc(k)}" value="${v}" list="carrierList" autocomplete="off" placeholder="SKT">`;
      else input = `<input data-ex="${esc(k)}" value="${v}" autocomplete="off">`;
      return `<label class="fld ${EXTRA.monthDayCols.includes(k) ? "fld-md" : ""}"><span>${esc(k)}</span>${input}</label>`;
    };
    const fields = x.fields || [];
    const more = fields.filter((k) => !EXTRA.primary.includes(k));
    const prim = fields.filter((k) => EXTRA.primary.includes(k));
    return `<div class="card" id="infoForm">
        ${listFields}
        ${x.missing || x.error ? `<p class="hint">추가 정보 엑셀을 읽지 못해 통신사·차량·미팅장소·메모는 수정할 수 없습니다.</p>` : more.map(field).join("") + prim.map(field).join("")}
        <datalist id="carrierList"><option value="SKT"><option value="KT"><option value="LG U+"><option value="알뜰폰"></datalist>
        <div class="file-actions">
          <button type="button" class="secondary" id="infoCancel">취소</button>
          <button type="button" class="primary" id="infoSave">저장</button>
        </div>
        ${canList ? `<p class="hint">연락처·주소·이메일을 바꾸면 고객리스트 엑셀에 반영되고 '${PRIORITY_COL}' 칸에 1이 표시됩니다.</p>` : ""}
      </div>`;
  }

  function bindExtra() {
    if (!editMode) return;
    const body = $("tabBody");
    // 월/일 선택 → 'MMDD'
    body.querySelectorAll(".mmdd").forEach((box) => {
      const mm = box.querySelector("[data-mm]"), dd = box.querySelector("[data-dd]"), hid = box.querySelector("[data-ex]");
      const upd = () => { hid.value = mm.value && dd.value ? pad2(mm.value) + pad2(dd.value) : (!mm.value && !dd.value ? "" : hid.value); };
      mm.addEventListener("change", upd);
      dd.addEventListener("change", upd);
    });
    $("infoCancel").addEventListener("click", () => setEditMode(false));
    $("infoSave").addEventListener("click", async () => {
      const c = current.cust;
      const btn = $("infoSave");
      // 고객리스트: 바뀐 칸만
      const listValues = {};
      body.querySelectorAll("[data-edit]").forEach((i) => {
        const f = EDIT_FIELDS.find((x) => x.label === i.dataset.edit);
        if (i.value.trim() === currentValue(c, f.cols).trim()) return;
        const col = c.entries.flatMap((e) => e.fields.map(([k]) => k.trim())).find((k) => f.cols.includes(k)) || f.label;
        listValues[col] = i.value.trim();
      });
      // 고객추가정보: 바뀐 칸이 있으면
      const inputs = [...body.querySelectorAll("[data-ex]")];
      const orig = { ...((current.extra && current.extra.values) || {}) };
      const v = Object.fromEntries(inputs.map((i) => [i.dataset.ex, i.value.trim()]));
      const extraDirty = inputs.some((i) => i.value.trim() !== (orig[i.dataset.ex] || ""));
      if (!Object.keys(listValues).length && !extraDirty) return setEditMode(false);
      btn.disabled = true; btn.textContent = "저장 중…";
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
        current.phones = null;
        toast("저장했습니다.");
        setEditMode(false);
      } catch (e) {
        btn.disabled = false; btn.textContent = "저장";
        saveError(e);
      }
    });
  }

  // 맨 위 '수정' 버튼: 성향 버튼 · 사진 바꾸기 · 정보 수정 화면
  function setEditMode(on) {
    editMode = on;
    $("editBtn").textContent = on ? "닫기" : "수정";
    $("editBtn").classList.toggle("on", on);
    $("photoMenu").hidden = true;
    renderGrade(current.cust);
    $("avatar").classList.toggle("editable", on);
    selectTab("basic");
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
    const d = new Date();
    const ymd = `${d.getFullYear()}${pad2(d.getMonth() + 1)}${pad2(d.getDate())}`;
    const fc = current.fc || folder.fc || "";
    const name = `${ymd} - ${fc} - ${current.cust.name} - 프로필사진.jpg`.replace(/[\\/:*?"<>|]/g, "");
    const item = await gx("PUT", itemPath(folder.id) + ":/" + encodeURIComponent(name) + ":/content",
      await blob.arrayBuffer(), "image/jpeg");
    current.files = current.files.filter((f) => f.id !== item.id).concat({ ...item, path: item.name });
    current.profilePhoto = item;
    if (current.photoUrl) URL.revokeObjectURL(current.photoUrl);
    current.photoUrl = URL.createObjectURL(blob);
    setAvatarPhoto();
  }

  // 프로필 사진 삭제 (이 고객 폴더의 '프로필사진' 이미지 모두 → OneDrive 휴지통)
  async function deleteProfilePhotos() {
    const list = current.files.filter((f) => /프로필사진/.test(f.name) && /\.(jpe?g|png|heic|webp)$/i.test(f.name));
    for (const f of list) await gx("DELETE", itemPath(f.id));
    current.files = current.files.filter((f) => !list.includes(f));
    current.profilePhoto = null;
    if (current.photoUrl) URL.revokeObjectURL(current.photoUrl);
    current.photoUrl = null;
    const a = $("avatar");
    a.style.backgroundImage = ""; a.classList.remove("has-photo"); a.textContent = current.cust.name[0];
  }

  function findProfilePhoto(files) {
    return files.filter((f) => /프로필사진/.test(f.name) && /\.(jpe?g|png|heic|webp)$/i.test(f.name))
      .sort((a, b) => (b.name > a.name ? 1 : -1) || (new Date(b.lastModifiedDateTime) - new Date(a.lastModifiedDateTime)))[0] || null;
  }

  async function loadProfilePhoto() {
    const c = current;
    if (!c.profilePhoto) return;
    try {
      const buf = await graph(itemPath(c.profilePhoto.id) + "/content", true);
      if (c !== current) return; // 다른 고객으로 넘어감
      c.photoUrl = URL.createObjectURL(new Blob([buf], { type: "image/jpeg" }));
      setAvatarPhoto();
      const t = $("photoThumb");
      if (t && !t.querySelector("img")) t.innerHTML = `<img src="${c.photoUrl}" alt="프로필 사진">`;
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
    ["이체일", ["이체일"]]
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
    ? `<div class="file-actions"><button class="secondary" data-open="${esc(f.id)}">📄 계약 열기</button></div>`
    : "";

  // 상품명에서 '무배당'이 들어간 괄호 삭제: "신한통합건강보험 원(ONE)(무배당, 해약환급금 미지급형)" → "신한통합건강보험 원(ONE)"
  const productName = (s) => String(s || "").replace(/\s*[(（][^()（）]*무배당[^()（）]*[)）]/g, "").replace(/\s{2,}/g, " ").trim();

  function contractHtml() {
    const list = current.contracts;
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
        .map(([k, v]) => [k, k === "합계보험료" && !/원\s*$/.test(v) ? v.trim() + "원" : v]); // 79,644 → 79,644원
      const pdf = findContractPdf(val(["최초계약일", "계약일"]), c.row["상품명"]);
      return `<div class="card">
        <div class="entry-head"><span class="tag t-${esc(c.type)}">${esc(c.type)}</span>
          ${status ? `<span class="tag">${esc(status)}</span>` : ""}
          <span class="hint">${esc([c.fc, c.row["이관일자"] ? "이관 " + c.row["이관일자"] : ""].filter(Boolean).join(" · "))}</span></div>
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
        <h3 class="sec-title">설정</h3>
        <label class="fld"><span>회사 주소 (일정 이동시간 계산의 출발·도착 기준)</span>
          <input id="setCompany" value="${esc(s.companyAddress || "")}" placeholder="예: 광주광역시 서구 상무중앙로 98" autocomplete="off"></label>
        <label class="fld"><span>카카오 REST API 키 (이동 거리·시간 계산용)</span>
          <input id="setKakao" value="${esc(s.kakaoRestKey || "")}" placeholder="developers.kakao.com › 앱 › 플랫폼 키" autocomplete="off" spellcheck="false"></label>
        <label class="chk-line"><input type="checkbox" id="setSaver" ${s.dataSaver ? "checked" : ""}>
          <span><b>데이터 절약 모드</b><br><span class="hint">켜면 고객 사진을 불러오지 않습니다.</span></span></label>
        <details class="fold" id="gFold"><summary>구글 캘린더 설정 <span class="hint">${esc(((s.google || {}).calendars || []).map((c) => c.name).join(", ") || "기본(1팀·김현성·최수진)")}</span></summary>
          <div id="gCalBox"></div></details>
        <div class="file-actions"><button type="button" class="primary" id="setSave" disabled>설정 저장</button></div>
      </div>`;
  }
  function bindSettings() {
    const btn = $("setSave");
    if (!btn) return;
    const read = () => {
      const g = readCalendarChoices();
      return { ...appSettings, companyAddress: $("setCompany").value.trim(), dataSaver: $("setSaver").checked,
        kakaoRestKey: $("setKakao").value.trim(),
        google: g ? { ...(appSettings.google || {}), ...g } : appSettings.google };
    };
    const calKey = (g) => JSON.stringify(((g || {}).calendars || []).map((c) => c.id).sort());
    const dirty = () => { const v = read(); return v.companyAddress !== (appSettings.companyAddress || "") || v.kakaoRestKey !== (appSettings.kakaoRestKey || "") || v.dataSaver !== !!appSettings.dataSaver || calKey(v.google) !== calKey(appSettings.google); };
    ["input", "change"].forEach((ev) => $("settingsCard").addEventListener(ev, () => { btn.disabled = !dirty(); }));
    btn.addEventListener("click", async () => {
      btn.disabled = true; btn.textContent = "저장 중…";
      try {
        await saveSettings(read());
        toast("설정을 저장했습니다.");
      } catch (e) { handleError(e); }
      btn.textContent = "설정 저장";
      btn.disabled = !dirty();
    });
    $("gFold").addEventListener("toggle", () => { if ($("gFold").open && !$("gCalBox").dataset.loaded) { $("gCalBox").dataset.loaded = "1"; loadCalendarChoices(); } });
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

  async function loadAdminState() {
    await resolveBase();
    const items = await listChildren(itemPath(base.id) + "/children");
    const fcs = items.filter((i) => i.folder && !i.name.startsWith("_")).map((f) => f.name)
      .sort((a, b) => a.localeCompare(b, "ko"));
    const perms = await readPerms();
    adminState = { fcs, perms };
  }

  async function openAdmin() {
    if (!isAdmin()) return show("searchView");
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
    const has = (em, fc) => isAll(em) || listOf(em).includes(fc);

    $("adminBody").innerHTML = settingsCardHtml() + `
      <div class="card">
        <h3 class="sec-title">사용 FC 현황</h3>
        <p class="hint">체크박스로 바로 바꾸고 '변경 저장'을 누르세요.</p>
        <div class="table-wrap"><table class="matrix">
          <tr><th>아이디</th><th>전체</th>${fcs.map((fc) => `<th>${esc(fc)}</th>`).join("")}<th></th></tr>
          ${(CFG.admins || []).map((a) => `<tr class="admin-row"><td>${esc(a)} <span class="tag">관리자</span></td>
            <td class="ok">●</td>${fcs.map(() => `<td class="ok">●</td>`).join("")}<td></td></tr>`).join("")}
          ${users.map((em) => `<tr class="perm-row" data-em="${esc(em)}">
            <td>${esc(em)}</td>
            <td><input type="checkbox" class="cb" data-all ${isAll(em) ? "checked" : ""} aria-label="${esc(em)} 전체"></td>
            ${fcs.map((fc) => `<td><input type="checkbox" class="cb" data-fc="${esc(fc)}" ${has(em, fc) ? "checked" : ""} aria-label="${esc(em)} ${esc(fc)}"></td>`).join("")}
            <td><button type="button" class="del-btn" data-del="${esc(em)}" aria-label="${esc(em)} 삭제">삭제</button></td>
          </tr>`).join("")}
        </table></div>
        ${users.length ? "" : '<p class="hint">아직 권한을 준 사용자가 없습니다. 아래에서 이메일을 추가하세요.</p>'}
        <div class="file-actions">
          <button class="secondary" id="resetPerms" disabled>되돌리기</button>
          <button class="primary" id="savePerms" disabled>변경 저장</button>
        </div>
      </div>

      <div class="card">
        <h3 class="sec-title">사용자 추가</h3>
        <form id="addUserForm" class="search">
          <input id="newUser" type="email" placeholder="Microsoft 계정 이메일" autocomplete="off" required>
          <button class="primary" type="submit">추가</button>
        </form>
        <p class="hint">추가하면 위 표에 줄이 생깁니다. 권한을 체크하고 '변경 저장'을 누르세요.</p>
      </div>`;

    $("addUserForm").addEventListener("submit", (ev) => {
      ev.preventDefault();
      const em = $("newUser").value.trim().toLowerCase();
      if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(em)) return toast("이메일 형식을 확인해 주세요.");
      if ((CFG.admins || []).map((a) => a.toLowerCase()).includes(em)) return toast("관리자는 이미 모든 폴더를 볼 수 있습니다.");
      const pending = collectRows(); // 표에서 아직 저장 안 한 체크 상태 유지
      extraUsers.add(em);
      renderAdmin();
      applyRows(pending);
    });

    const rows = [...$("adminBody").querySelectorAll(".perm-row")];
    const same = (a, b) => a.length === b.length && a.every((x) => b.includes(x));
    const sel = (row) => {
      const all = row.querySelector("input[data-all]");
      return all.checked ? ["*"] : [...row.querySelectorAll("input[data-fc]")].filter((b) => b.checked).map((b) => b.dataset.fc);
    };
    const changed = () => rows.filter((r) => !same(sel(r), listOf(r.dataset.em)));
    const refresh = () => {
      rows.forEach((r) => r.classList.toggle("changed", !same(sel(r), listOf(r.dataset.em))));
      const n = changed().length;
      $("savePerms").disabled = !n;
      $("resetPerms").disabled = !n;
      $("savePerms").textContent = n ? `변경 저장 (${n})` : "변경 저장";
    };
    rows.forEach((row) => {
      const all = row.querySelector("input[data-all]");
      const boxes = [...row.querySelectorAll("input[data-fc]")];
      const syncAll = () => { all.indeterminate = !all.checked && boxes.some((b) => b.checked); };
      all.addEventListener("change", () => { boxes.forEach((b) => (b.checked = all.checked)); syncAll(); refresh(); });
      boxes.forEach((b) => b.addEventListener("change", () => {
        if (!b.checked) all.checked = false;
        else if (boxes.every((x) => x.checked)) all.checked = true;
        syncAll(); refresh();
      }));
      syncAll();
    });
    refresh();
    $("resetPerms").addEventListener("click", () => renderAdmin());
    bindSettings();
    $("adminBody").querySelectorAll("[data-del]").forEach((b) => b.addEventListener("click", () => deleteUser(b.dataset.del)));
    $("savePerms").addEventListener("click", () => savePerms(changed().map((r) => [r.dataset.em, sel(r)])));

    // 사용자 추가 시 저장 전 체크 상태 보존용
    function collectRows() { return rows.map((r) => [r.dataset.em, sel(r)]); }
    function applyRows(pending) {
      pending.forEach(([em, list]) => {
        const row = $("adminBody").querySelector(`.perm-row[data-em="${CSS.escape(em)}"]`);
        if (!row) return;
        const all = row.querySelector("input[data-all]");
        all.checked = list.includes("*");
        row.querySelectorAll("input[data-fc]").forEach((b) => (b.checked = all.checked || list.includes(b.dataset.fc)));
        all.dispatchEvent(new Event("change"));
        if (!all.checked) row.querySelectorAll("input[data-fc]").forEach((b) => b.dispatchEvent(new Event("change")));
      });
    }
  }

  // 등록된 아이디 삭제 (권한 파일에서 바로 제거)
  async function deleteUser(em) {
    if (!confirm(`${em} 아이디를 삭제할까요?\n삭제하면 이 아이디는 앱에서 고객을 검색할 수 없습니다.`)) return;
    try {
      const perms = await readPerms();
      delete perms.users[em];
      extraUsers.delete(em);
      await gx("PUT", permPath(), perms);
      adminState.perms = perms;
      toast(`${em} 아이디를 삭제했습니다.`);
      renderAdmin();
    } catch (e) {
      handleError(e);
    }
  }

  async function savePerms(changes) {
    if (!changes.length) return;
    const btn = $("savePerms");
    btn.disabled = true;
    btn.textContent = "저장 중…";
    try {
      // 저장 직전에 최신 파일을 다시 읽고, 바뀐 아이디만 반영 (다른 설정 보존)
      const perms = await readPerms();
      changes.forEach(([em, list]) => {
        if (list.length) perms.users[em] = list; else { delete perms.users[em]; extraUsers.delete(em); }
      });
      await gx("PUT", permPath(), perms);
      toast(`저장했습니다 (${changes.length}개 아이디).`);
      adminState.perms = perms;
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
    if (m === "NO_SHARE") {
      $("scopeHint").textContent = "아직 고객 폴더에 접근할 수 없습니다. 관리자에게 OneDrive '고객' 폴더 공유를 요청하세요.";
      return toast("고객 폴더 공유가 필요합니다. 관리자에게 요청하세요.");
    }
    if (m.includes("404")) return toast("OneDrive에서 '" + CFG.baseFolder + "' 폴더를 찾을 수 없습니다.");
    if (m.includes("403")) return toast("권한이 없습니다. 관리자에게 요청하세요.");
    toast("오류: " + m.slice(0, 120));
  }

  // ---------- 시작 ----------
  $("loginBtn").addEventListener("click", () => msal.loginRedirect({ scopes: SCOPES }));
  $("logoutBtn").addEventListener("click", () => msal.logoutRedirect());
  // 상단 검색창: 어느 화면에서든 검색 → 고객 검색 결과 화면으로
  function runTopSearch(q) {
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
  $("backBtn").addEventListener("click", () => history.back());
  // 고객 사진을 누르면 크게 보기, 다시 누르면 닫기
  $("avatar").addEventListener("click", () => {
    if (!current) return;
    if (editMode) {
      if (!current.folders.length) return toast("이 고객의 폴더가 없어 사진을 저장할 수 없습니다.");
      if (current.profilePhoto) { $("photoMenu").hidden = !$("photoMenu").hidden; return; }
      return $("photoInput").click();
    }
    if (!current.photoUrl) return;
    $("photoViewerImg").src = current.photoUrl;
    $("photoViewer").hidden = false;
  });
  $("editBtn").addEventListener("click", () => { if (current) setEditMode(!editMode); });
  $("photoPick").addEventListener("click", () => { $("photoMenu").hidden = true; $("photoInput").click(); });
  $("photoMenuClose").addEventListener("click", () => { $("photoMenu").hidden = true; });
  $("photoDel").addEventListener("click", async () => {
    $("photoMenu").hidden = true;
    if (!confirm("이 고객의 프로필 사진을 삭제할까요? (OneDrive 휴지통으로 이동)")) return;
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
  // 검색 범위 버튼: 전체 + config.lists 순서대로
  $("modes").innerHTML = `<button type="button" data-mode="all" class="on">전체</button>` +
    CFG.lists.map((l) => `<button type="button" data-mode="${esc(l.id)}">${esc(l.label)}</button>`).join("");
  $("modes").style.gridTemplateColumns = `repeat(${CFG.lists.length + 1}, 1fr)`;
  $("modes").querySelectorAll("button").forEach((b) => b.addEventListener("click", () => {
    mode = b.dataset.mode;
    $("modes").querySelectorAll("button").forEach((x) => x.classList.toggle("on", x === b));
    if ($("q").value.trim()) search($("q").value);
  }));
  // ---------- 구글 캘린더 (일정) ----------
  const GCFG = CFG.google || {};
  const GCAL = "https://www.googleapis.com/calendar/v3";
  const GTOKEN_KEY = "shlife.gtoken";
  function storedGToken() {
    try {
      const t = JSON.parse(localStorage.getItem(GTOKEN_KEY) || "null");
      return t && t.exp > Date.now() + 60000 ? t.token : null;
    } catch (e) { return null; }
  }
  // 구글 로그인 창(팝업) → 캘린더 읽기 권한 토큰 (1시간 유효, 버튼을 눌렀을 때만 호출)
  function requestGToken(prompt) {
    return new Promise((resolve, reject) => {
      const g = window.google && window.google.accounts && window.google.accounts.oauth2;
      if (!g) return reject(new Error("구글 로그인 모듈을 불러오지 못했습니다. 잠시 후 다시 시도하세요."));
      const client = g.initTokenClient({
        client_id: GCFG.clientId,
        scope: "https://www.googleapis.com/auth/calendar.readonly",
        hint: (appSettings.google && appSettings.google.account) || undefined,
        callback: (r) => {
          if (r.error) return reject(new Error("구글 연결 실패: " + r.error));
          try { localStorage.setItem(GTOKEN_KEY, JSON.stringify({ token: r.access_token, exp: Date.now() + (r.expires_in || 3600) * 1000 })); } catch (e) { /* 저장 못 해도 이번엔 사용 */ }
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
    if (res.status === 401) { try { localStorage.removeItem(GTOKEN_KEY); } catch (e) {} gTokenMem = null; throw new Error("GOOGLE_AUTH"); }
    if (!res.ok) throw new Error("Calendar " + res.status + ": " + (await res.text()).slice(0, 160));
    return res.json();
  }
  const calAlias = (id, name) => ((GCFG.calendarNames || {})[id] || name);
  const calName = (c) => calAlias(c.id, String(c.summaryOverride || c.summary || "").trim());
  async function calendarList() {
    return ((await gcal("/users/me/calendarList?minAccessRole=reader&maxResults=250")).items || [])
      .map((c) => ({ id: c.id, name: calName(c), color: c.backgroundColor, primary: !!c.primary }));
  }
  // 관리자 설정에서 고른 캘린더 (없으면 config의 기본 이름: 1팀, 김현성, 최수진)
  async function scheduleCalendars() {
    const saved = ((appSettings.google || {}).calendars || []).filter((c) => c && c.id);
    if (saved.length) return saved;
    const want = GCFG.calendars || [];
    return (await calendarList()).filter((c) => want.includes(c.name));
  }
  const withAlias = (c) => ({ ...c, name: calAlias(c.id, c.name) });
  async function fetchEvents(from, to) {
    const cals = await scheduleCalendars();
    if (!cals.length) throw new Error("NO_CALENDAR");
    const q = `?singleEvents=true&orderBy=startTime&maxResults=250&timeMin=${encodeURIComponent(from.toISOString())}&timeMax=${encodeURIComponent(to.toISOString())}`;
    const lists = await Promise.all(cals.map(withAlias).map(async (c) => {
      const d = await gcal("/calendars/" + encodeURIComponent(c.id) + "/events" + q);
      return (d.items || []).filter((e) => e.status !== "cancelled").map((e) => ({
        id: c.id + "|" + e.id, cal: c,
        title: String(e.summary || "(제목 없음)").trim(),
        location: String(e.location || "").trim(),
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
  const LIST_PRIORITY = ["진행", "DB", "이관", "계약"]; // 같은 이름이 여러 리스트에 있으면 이 순서로
  const prio = (type) => { const i = LIST_PRIORITY.indexOf(type); return i < 0 ? 99 : i; };
  const entryAddr = (e) => { const a = e.fields.find(([k]) => k.trim() === "주소"); return a ? a[1] : ""; };

  // 일정 제목에 들어 있는 고객 이름 → 고객 (같은 이름이 여러 명이면 모두, 최근 등록 순)
  function matchCustomers(title, records, contracts) {
    const nt = norm(title);
    const fcs = new Set(records.map((r) => norm(r.fc)).filter(Boolean));
    const names = [...new Set(records.map((r) => r.name))].filter((n) => norm(n).length >= 2 && nt.includes(norm(n)));
    if (!names.length) return [];
    // FC 이름이 아닌 것, 긴 이름 우선
    names.sort((a, b) => (fcs.has(norm(a)) - fcs.has(norm(b))) || norm(b).length - norm(a).length);
    const name = names[0];
    const city = titleRegion(title);
    const fcHead = (title.match(/^\s*([^\s)]{1,3})\)/) || [])[1] || ""; // "최) …" → 최수진 FC 고객 우선
    const recs = records.filter((r) => norm(r.name) === norm(name));
    const score = (r) => {
      const fcMiss = fcHead && r.fc && !String(r.fc).startsWith(fcHead) ? 4 : 0;
      const a = entryAddr(r);
      if (!a) return 1 + fcMiss;
      return (custRegionCity(regionFromAddress(a)) === city ? 0 : 2) + fcMiss;
    };
    // 같은 사람(이름+생년월일)으로 묶기
    const groups = new Map();
    recs.forEach((r) => {
      const d = String(r.birth).replace(/\D/g, "");
      const key = /^(19|20)\d{6}/.test(d) ? d.slice(2, 8) : d.slice(0, 6);
      if (!groups.has(key)) groups.set(key, { name: r.name, birth: r.birth, entries: [] });
      groups.get(key).entries.push(r);
    });
    const list = [...groups.values()];
    list.forEach((g) => {
      g.entries.sort((a, b) => score(a) - score(b) || prio(a.type) - prio(b.type));
      g.city = city;
      g.recent = recencyOf(g, contracts);
    });
    // 최근 등록 순 (계약일·이관일·보장분석일 중 가장 최근), 같으면 지역·리스트 우선순위
    const best = (g) => Math.min(...g.entries.map((e) => score(e) * 100 + prio(e.type)));
    return list.sort((a, b) => (b.recent || "").localeCompare(a.recent || "") || best(a) - best(b));
  }
  // 날짜 글자 → 비교용 YYYYMMDD ("2026.05", "2026-08-31", "2005.01.01")
  const dateKey = (v) => {
    const d = String(v || "").replace(/\D/g, "");
    if (/^(19|20)\d{6}/.test(d)) return d.slice(0, 8);
    if (/^(19|20)\d{4}$/.test(d)) return d + "00";
    return "";
  };
  const RECENT_COLS = ["최초계약일", "계약일", "이관일자", "이관월", "최근 보장분석"];
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
    return `<div class="sched-modes">${[["day", "Today"], ["week", "Week"], ["month", "Month"]].map(([m, l]) =>
      `<button type="button" data-sm="${m}" class="${schedMode === m ? "on" : ""}">${l}</button>`).join("")}</div>
      <div class="sched-nav"><button type="button" id="schedPrev" aria-label="이전">‹</button>
        <button type="button" id="schedLabel" class="sched-label"></button>
        <button type="button" id="schedNext" aria-label="다음">›</button></div>
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
        v.querySelectorAll("[data-sm]").forEach((x) => x.classList.toggle("on", x === b));
        renderSchedule();
      }));
      $("schedPrev").addEventListener("click", () => moveSched(-1));
      $("schedNext").addEventListener("click", () => moveSched(1));
      $("schedLabel").addEventListener("click", () => { schedDate = dayStart(new Date()); renderSchedule(); });
    }
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
    $("schedBody").innerHTML = `<div class="card center"><p class="lead">구글 캘린더 연결</p>
      <p class="hint">${esc(msg || "일정을 보려면 구글 계정으로 한 번 로그인해 주세요.")}${appSettings.google && appSettings.google.account ? "<br>연결 계정: " + esc(appSettings.google.account) : ""}</p>
      <button type="button" class="primary" id="gConnect">구글 계정 연결</button></div>`;
    $("gConnect").addEventListener("click", async () => {
      try { await requestGToken(); renderSchedule(); } catch (e) { toast(e.message); }
    });
  }

  async function renderSchedule() {
    const req = ++schedReq;
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
    if (schedMode === "week") return renderWeek(events, from);
    return renderMonth(events, from, to);
  }

  const calDot = (e) => `<span class="cal-dot" style="background:${esc(e.cal.color || "var(--brand)")}"></span>`;
  // "10:00 ~ 11:00 스케줄명 - 캘린더" 한 줄
  const evLine = (e) => `<span class="ev-time">${esc(timeText(e))}</span> <span class="ev-title">${esc(e.title)}</span> <span class="ev-cal">- ${esc(e.cal.name)}</span>`;

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
  const isTeamEvent = (e) => !e.allDay && e.cal.name === (GCFG.teamCalendar || "1팀");

  async function renderDay(events, req) {
    const d = schedDate;
    const evs = events.filter((e) => onDay(e, d));
    const body = $("schedBody");
    if (!evs.length) { body.innerHTML = '<div class="card center"><p class="hint">일정이 없습니다.</p></div>'; return; }
    body.innerHTML = '<div id="dayList"></div>';
    // 고객 찾기 (고객 리스트 엑셀 + 계약일로 최근 등록 순)
    let records = [], contracts = [];
    if (evs.some(isTeamEvent)) {
      try { [records, contracts] = await Promise.all([loadLists(), loadContracts().catch(() => [])]); } catch (e) { handleError(e); }
    }
    if (req !== schedReq) return;
    const items = evs.map((e) => ({ e, custs: isTeamEvent(e) ? matchCustomers(e.title, records, contracts) : [] }));
    items.forEach((it) => { it.cust = it.custs[0] || null; });
    const list = $("dayList");
    list.innerHTML = items.map((it, i) => `<div class="travel" id="tr${i}" hidden></div>${dayCardHtml(it, i)}`).join("") +
      `<div class="travel" id="tr${items.length}" hidden></div>`;
    items.forEach((it, i) => bindDayCard(it, i));
    // 추가정보(미팅장소, 상담기록)
    await Promise.all(items.map(async (it, i) => {
      await Promise.all(it.custs.map(async (c, j) => {
        const fc = (c.entries.find((e) => e.fc) || {}).fc || "";
        try { c.extra = await loadExtraFor(fc, c.name); } catch (e) { c.extra = null; }
        if (req !== schedReq) return;
        fillCust(c, i, j);
      }));
      it.extra = it.cust && it.cust.extra;
    }));
    if (req !== schedReq) return;
    renderTravel(items, req);
  }

  function custBlockHtml(c, i, j) {
    const phone = (c.entries.find((x) => x.phone) || {}).phone || "";
    const addr = currentValue(c, ["주소"]);
    const types = [...new Set(c.entries.map((x) => x.type))];
    const g = genderOf(c);
    return `<div class="ev-cust">
        <button type="button" class="ev-name" data-open="${i}-${j}"><span class="cname ${gradeClass(gradeOf(c))}">${esc(c.name)}</span>
          ${types.map((t) => `<span class="tag t-${esc(t)}">${esc(t)}</span>`).join("")}<span class="chev">›</span></button>
        <dl class="ev-info">
          ${c.birth ? `<dt>생년월일</dt><dd>${esc(c.birth)}${g ? ` (${esc(g)})` : ""}</dd>` : ""}
          ${phone ? `<dt>연락처</dt><dd><a class="nowrap" href="tel:${esc(phone.replace(/[^\d+]/g, ""))}">${esc(phone)}</a></dd>` : ""}
          ${addr ? `<dt>주소</dt><dd>${esc(addr)}</dd>` : ""}
          <dt id="mpl${i}-${j}" hidden>미팅장소</dt><dd id="mp${i}-${j}" hidden></dd>
        </dl>
        <div id="cs${i}-${j}"></div></div>`;
  }
  function dayCardHtml(it, i) {
    const e = it.e;
    return `<div class="card ev-card">${calDot(e)}<div class="ev-head">${evLine(e)}</div>
      ${e.location ? `<p class="hint">📍 ${esc(e.location)}</p>` : ""}${it.custs.map((c, j) => custBlockHtml(c, i, j)).join("")}</div>`;
  }
  function bindDayCard(it, i) {
    it.custs.forEach((c, j) => {
      const b = document.querySelector(`[data-open="${i}-${j}"]`);
      if (b) b.addEventListener("click", () => openCustomer(c));
    });
  }
  function fillCust(c, i, j) {
    const x = c.extra || {};
    const mp = ((x.values || {})["미팅장소"] || "").trim();
    if (mp) { $(`mp${i}-${j}`).textContent = mp; $(`mp${i}-${j}`).hidden = false; $(`mpl${i}-${j}`).hidden = false; }
    const cs = (x.consults || []).slice(0, 10);
    const box = $(`cs${i}-${j}`);
    if (box && cs.length) {
      box.innerHTML = `<details class="ev-consult"><summary>최근 상담 ${cs.length}건</summary>
        ${cs.map((r) => `<div class="consult-item"><div class="hint">${esc(r.date)}</div><div>${esc(r.text).replace(/\n/g, "<br>")}</div></div>`).join("")}</details>`;
    }
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
        tries.push((it.extra && it.extra.values && it.extra.values["미팅장소"]) || "");
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
      show(timed[0].i, `${leave ? `<b>🏢 회사 출발 ${hm(leave)}</b> · ` : "🏢 회사에서 "}${legTxt(r)}`);
    }
    for (let k = 1; k < timed.length; k++) {
      if (req !== schedReq) return;
      const a = pts[k - 1], b = pts[k];
      const r = a && b ? await route(a.p, b.p) : null;
      const gap = Math.round((timed[k].e.start - timed[k - 1].e.end) / 60000);
      const late = r && gap < r.min ? ` <span class="warn">⚠ 이동시간 부족 (여유 ${Math.max(gap, 0)}분)</span>` : "";
      show(timed[k].i, legTxt(r) + late);
    }
    // 회사 복귀: 마지막 1팀 일정 뒤
    const last = timed[timed.length - 1], lp = pts[pts.length - 1];
    if (company && lp) {
      const r = await route(lp.p, company);
      const back = r ? new Date(+last.e.end + r.min * 60000) : null;
      show(last.i + 1, `${legTxt(r)}${back ? ` · <b>🏢 회사 도착 ${hm(back)}</b>` : ""}`);
    }
  }


  // ---------- 관리자 설정: 구글 계정·캘린더 선택 ----------
  async function loadCalendarChoices() {
    const box = $("gCalBox");
    if (!box) return;
    if (!storedGToken() && !gTokenMem) {
      box.innerHTML = `<button type="button" class="secondary" id="gAdminConnect">구글 계정 연결</button>`;
      $("gAdminConnect").addEventListener("click", async () => {
        try { await requestGToken("select_account"); loadCalendarChoices(); } catch (e) { toast(e.message); }
      });
      return;
    }
    box.innerHTML = '<p class="hint">캘린더 목록을 불러오는 중…</p>';
    let cals;
    try { cals = await calendarList(); } catch (e) {
      if (e.message === "GOOGLE_AUTH") { gTokenMem = null; return loadCalendarChoices(); }
      box.innerHTML = ""; return handleError(e);
    }
    const primary = cals.find((c) => c.primary);
    const saved = ((appSettings.google || {}).calendars || []).map((c) => c.id);
    const want = GCFG.calendars || [];
    const isOn = (c) => (saved.length ? saved.includes(c.id) : want.includes(c.name));
    box.dataset.account = primary ? primary.id : "";
    box.innerHTML = `<p class="hint">로그인한 계정: ${esc(primary ? primary.id : "-")}
        <button type="button" class="link-btn" id="gSwitch">다른 계정</button></p>
      ${cals.map((c, i) => `<label class="chk-line"><input type="checkbox" data-cal="${i}" ${isOn(c) ? "checked" : ""}>
        <span>${calDot({ cal: c })}${esc(c.name)}</span></label>`).join("")}`;
    box._cals = cals;
    $("gSwitch").addEventListener("click", async () => {
      try { await requestGToken("select_account"); loadCalendarChoices(); } catch (e) { toast(e.message); }
    });
    box.dispatchEvent(new Event("change", { bubbles: true }));
  }
  function readCalendarChoices() {
    const box = $("gCalBox");
    if (!box || !box._cals) return null;
    const cals = [...box.querySelectorAll("[data-cal]")].filter((x) => x.checked).map((x) => box._cals[+x.dataset.cal])
      .map((c) => ({ id: c.id, name: c.name, color: c.color }));
    return { account: box.dataset.account || "", calendars: cals };
  }

  // 큰 메뉴: 일정 / 고객 / 관리자
  $("mainNav").querySelectorAll("button").forEach((b) => b.addEventListener("click", () => {
    const s = b.dataset.section;
    if (s === "schedule") openSchedule();
    else if (s === "admin") openAdmin();
    else show("searchView");
  }));
  window.addEventListener("popstate", () => { if (!$("customerView").hidden) { if (returnView === "scheduleView") openSchedule(); else show("searchView"); } });

  // (검색 가능한 FC 문구는 표시하지 않음)
  function showScope() { $("scopeHint").textContent = ""; }

  (async function init() {
    if (CFG.clientId === "YOUR-CLIENT-ID") {
      show("loginView");
      $("loginBtn").disabled = true;
      toast("config.js에 clientId를 넣어 주세요.");
      return;
    }
    await msal.initialize();
    try { await msal.handleRedirectPromise(); } catch (e) { handleError(e); }
    if (msal.getAllAccounts().length) {
      if (sessionStorage.getItem("openAdmin") && isAdmin()) {
        sessionStorage.removeItem("openAdmin");
        openAdmin();
      } else {
        schedMode = "day"; schedDate = dayStart(new Date());
        loadSettings().then(() => openSchedule()); // 첫 화면: 일정 Today (설정을 읽은 뒤)
      }
      loadSettings(); // 관리자 설정 (회사 주소, 데이터 절약 모드 등)
      loadLists().then(showScope).catch(handleError); // 미리 불러오기
    } else {
      show("loginView");
    }
  })();
})();
