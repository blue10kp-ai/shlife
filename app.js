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
  const TITLES = { scheduleView: "일정", searchView: "고객 찾기", customerView: "고객 정보", adminView: "관리자" };
  const SECTION = { scheduleView: "schedule", searchView: "customer", customerView: "customer", adminView: "admin" };
  function show(view) {
    views.forEach((v) => ($(v).hidden = v !== view));
    $("backBtn").hidden = view !== "customerView";
    $("logoutBtn").hidden = view === "loginView";
    $("title").textContent = TITLES[view] || "고객 찾기";
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
  const permPath = () => itemPath(base.id) + ":/" + encodeURIComponent(PERM_FILE) + ":/content";
  const me = () => String((account() || {}).username || "").toLowerCase();

  async function readPerms() {
    await resolveBase();
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
  async function loadLists() {
    if (listCache) return listCache;
    const { files, folders } = await scanBase();
    const sources = CFG.lists.filter((l) => l.fileKeyword).flatMap((l) =>
      files.filter((f) => f.name.includes(l.fileKeyword)).map((f) => ({ ...l, file: f })));
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
      wb.SheetNames.forEach((sheet) => {
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheet], { header: 1, defval: "", raw: false });
        // 제목 줄이 있을 수 있으니 '이름/계약자' 칸이 있는 줄을 머리글로 사용
        const h = rows.findIndex((r) => r.some((c) => NAME_COLS.includes(String(c).trim())));
        if (h < 0) return;
        const head = rows[h].map((c) => String(c).trim());
        rows.slice(h + 1).forEach((r) => {
          const row = {};
          head.forEach((k, i) => { if (k) row[k] = r[i]; });
          const name = pick(row, NAME_COLS);
          if (!name) return;
          out.push({
            type: s.label, name, fc, sheet,
            birth: pick(row, BIRTH_COLS),
            phone: pick(row, PHONE_COLS),
            fields: head.map((k, i) => [k, String(r[i] ?? "").trim()]).filter(([k, v]) => k && v)
          });
        });
      });
    }));
    listCache = out;
    return out;
  }

  // ---------- 검색 ----------
  let mode = "all";
  const norm = (s) => String(s).replace(/\s+/g, "").toLowerCase();
  const custPart = (n) => n.split(" - ")[0]; // 폴더명 "고객명 - 담당자명"

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
      const hits = records.filter((r) => (!want || r.type === want.label) && norm(r.name).includes(norm(q)));
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
          <span class="r-name">${esc(c.name)}<span class="r-sub">${esc([c.birth, fcs.join(", ")].filter(Boolean).join(" · "))}</span></span>
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
        files.filter((f) => f.name.includes(l.fileKeyword)).map((f) => ({ ...l, file: f })));
      const out = [];
      await Promise.all(sources.map(async (s) => {
        const wb = await readBook(s.file.id);
        const fc = s.file.fc || (s.file.name.split(" - ")[1] || "").trim();
        wb.SheetNames.forEach((sheet) => {
          const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheet], { header: 1, defval: "", raw: false });
          const h = rows.findIndex((r) => r.some((c) => String(c).trim() === "계약자"));
          if (h < 0) return;
          const head = rows[h].map((c) => String(c).trim());
          rows.slice(h + 1).forEach((r) => {
            const fields = head.map((k, i) => [k, String(r[i] ?? "").trim()]).filter(([k, v]) => k && v);
            if (!fields.length) return;
            const row = Object.fromEntries(fields);
            const names = CONTRACT_NAME_COLS.map((c) => row[c]).filter(Boolean).map(norm);
            if (names.length) out.push({ type: s.label, fc, sheet, row, fields, names });
          });
        });
      }));
      return out;
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
    consult: { sheet: "상담기록", nameCol: "고객명", fcCol: "담당 FC", dateCol: "저장날짜", textCol: "내용" }
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

  // 상담기록: 날짜 값 → "2026.11.08"
  function consultDate(cell) {
    if (!cell) return { key: 0, text: "" };
    const v = cell.v;
    let y, m, d;
    if (typeof v === "number" && v > 19000000) { y = Math.floor(v / 10000); m = Math.floor(v / 100) % 100; d = v % 100; }
    else if (typeof v === "number") { const t = new Date(Date.UTC(1899, 11, 30) + v * 86400000); y = t.getUTCFullYear(); m = t.getUTCMonth() + 1; d = t.getUTCDate(); }
    else { const s = String(v).replace(/\D/g, ""); if (s.length >= 8) { y = +s.slice(0, 4); m = +s.slice(4, 6); d = +s.slice(6, 8); } else return { key: 0, text: String(v) }; }
    return { key: y * 10000 + m * 100 + d, text: `${y}.${pad2(m)}.${pad2(d)}` };
  }

  function readConsults(wb, fc, name) {
    const C = EXTRA.consult;
    const ws = wb.Sheets[C.sheet];
    if (!ws || !ws["!ref"]) return [];
    const rg = XLSX.utils.decode_range(ws["!ref"]);
    const head = [];
    for (let c = rg.s.c; c <= rg.e.c; c++) { const cell = ws[XLSX.utils.encode_cell({ r: rg.s.r, c })]; head[c] = cell ? String(cell.v).trim() : ""; }
    const col = (h) => head.indexOf(h);
    const [ni, fi, di, ti] = [col(C.nameCol), col(C.fcCol), col(C.dateCol), col(C.textCol)];
    if (ni < 0 || ti < 0) return [];
    const at = (r, c) => c < 0 ? null : ws[XLSX.utils.encode_cell({ r, c })];
    const out = [];
    for (let r = rg.s.r + 1; r <= rg.e.r; r++) {
      const nm = at(r, ni), tx = at(r, ti);
      if (!nm || !tx || norm(nm.v) !== norm(name)) continue;
      if (!sameFc(at(r, fi) ? at(r, fi).v : "", fc)) continue;
      const d = consultDate(at(r, di));
      out.push({ row: r, date: d.text, key: d.key, text: String(tx.v) });
    }
    return out.sort((a, b) => b.key - a.key || b.row - a.row); // 최근 날짜가 위로
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

  // 상담기록 시트 맨 아래에 [고객명, 담당 FC, 저장날짜(YYYYMMDD), 내용] 추가
  async function addConsult(fc, name, text) {
    const C = EXTRA.consult;
    const x = await readExtra();
    if (!x) throw new Error("NO_EXTRA_FILE");
    let ws = x.wb.Sheets[C.sheet];
    if (!ws) {
      ws = XLSX.utils.aoa_to_sheet([[C.nameCol, C.fcCol, C.dateCol, C.textCol]]);
      XLSX.utils.book_append_sheet(x.wb, ws, C.sheet);
    }
    const rg = XLSX.utils.decode_range(ws["!ref"] || "A1:D1");
    const head = [];
    for (let c = rg.s.c; c <= rg.e.c; c++) { const cell = ws[XLSX.utils.encode_cell({ r: rg.s.r, c })]; head[c] = cell ? String(cell.v).trim() : ""; }
    // 마지막으로 내용이 있는 줄 다음
    let last = rg.s.r;
    for (let r = rg.s.r + 1; r <= rg.e.r; r++) for (let c = rg.s.c; c <= rg.e.c; c++) if (ws[XLSX.utils.encode_cell({ r, c })]) last = r;
    const r = last + 1;
    const now = new Date();
    const ymd = now.getFullYear() * 10000 + (now.getMonth() + 1) * 100 + now.getDate();
    const put = (h, cell) => { const c = head.indexOf(h); if (c >= 0) ws[XLSX.utils.encode_cell({ r, c })] = cell; };
    put(C.nameCol, { t: "s", v: name });
    put(C.fcCol, { t: "s", v: fc });
    put(C.dateCol, { t: "n", v: ymd });
    put(C.textCol, { t: "s", v: text });
    rg.e.r = Math.max(rg.e.r, r);
    ws["!ref"] = XLSX.utils.encode_range(rg);
    await uploadExtra(x);
    return readConsults(x.wb, fc, name);
  }

  // PDF는 OneDrive 웹 대신 휴대폰의 PDF 보기 앱/브라우저로 열기 (임시 다운로드 주소 사용)
  async function openFileDirect(id) {
    const w = window.open("", "_blank"); // 팝업 차단을 피하려고 먼저 창을 엶
    try {
      const it = await graph(itemPath(id) + "?select=id,name,@microsoft.graph.downloadUrl");
      const url = it["@microsoft.graph.downloadUrl"];
      if (!url) throw new Error("다운로드 주소를 받지 못했습니다.");
      if (w) w.location.href = url; else location.href = url;
    } catch (e) {
      if (w) w.close();
      handleError(e);
    }
  }

  // ---------- 고객 화면 ----------
  let current = null; // { cust, folders, files, phones, contracts, fc, extra }

  async function openCustomer(cust) {
    show("customerView");
    $("custName").textContent = cust.name;
    $("avatar").textContent = cust.name[0];
    $("custSub").textContent = [cust.birth, ...new Set(cust.entries.map((e) => e.type))].filter(Boolean).join(" · ");
    $("tabBody").innerHTML = '<p class="hint pad">고객 폴더를 찾는 중…</p>';
    $("tabs").innerHTML = "";
    history.pushState({ v: "customer" }, "");
    current = { cust, folders: [], files: [], phones: null, contracts: [], fc: "", extra: { values: {} } };
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
      current.fc = fcs[0] || (folders[0] && folders[0].fc) || "";
      const [contracts, extra] = await Promise.all([
        loadContracts().catch((e) => { console.error(e); return []; }),
        loadExtraFor(current.fc, cust.name).catch((e) => { console.error(e); return { error: true, values: {} }; })
      ]);
      current.contracts = contracts.filter((c) => c.names.includes(norm(cust.name)) && (!fcs.length || fcs.includes(c.fc)));
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
          <div class="file-name">${esc(f.path)}<div class="hint">${new Date(f.lastModifiedDateTime).toLocaleDateString("ko-KR")}</div></div>
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

  // 기본정보: 고객리스트의 인적 정보만 (계약·이관 내용, 상품명, 우편번호, 성별 제외 → 성별은 생년월일 뒤에)
  //   + 추가 정보 엑셀의 통신사(연락처 앞), 차량번호·자동차보험갱신일(주소 밑)
  const BASIC_EXCLUDE = ["상품명 - 보험료", "상품명", "보험료", "이관월", "고객 폴더", "분류", "No", "순번", "메모", "모집FC", "담당FC", "담당 FC", "우편번호", "성별"];
  const BIRTH_KEYS = ["생년월일", "생일"];
  let extraOpen = false;

  function genderMark(v) {
    const s = String(v || "").trim();
    if (/^(남|남자|M|male)$/i.test(s)) return "남";
    if (/^(여|여자|F|female)$/i.test(s)) return "여";
    return s;
  }

  function basicHtml() {
    const c = current.cust;
    const x = current.extra || { values: {}, fields: [] };
    const ev = x.values || {};
    const merged = new Map();
    let gender = "";
    c.entries.forEach((e) => e.fields.forEach(([k, v]) => {
      if (k.trim() === "성별" && !gender) gender = genderMark(v);
      if (BASIC_EXCLUDE.includes(k.trim())) return;
      if (!merged.has(k)) merged.set(k, []);
      if (!merged.get(k).includes(v)) merged.get(k).push(v);
    }));
    // 성별을 생년월일 뒤에: 1980-10-21 (남)
    if (gender) BIRTH_KEYS.forEach((k) => { if (merged.has(k)) merged.set(k, merged.get(k).map((v) => `${v} (${gender})`)); });
    // 통신사를 연락처 앞에
    const carrier = (ev[EXTRA.carrierCol] || "").trim();
    if (carrier) PHONE_COLS.forEach((k) => { if (merged.has(k)) merged.set(k, merged.get(k).map((v) => carrier + " " + v)); });
    // 차량번호·자동차보험갱신일을 주소 밑에 (갱신일은 "3월 26일")
    const car = EXTRA.afterAddress.filter((k) => (ev[k] || "").trim())
      .map((k) => [k, [EXTRA.monthDayCols.includes(k) ? monthDayText(ev[k]) : ev[k]]]);
    let rows = [...merged];
    if (car.length) {
      const ai = rows.findIndex(([k]) => k === "주소");
      rows = ai >= 0 ? [...rows.slice(0, ai + 1), ...car, ...rows.slice(ai + 1)] : rows.concat(car);
    }
    const types = [...new Set(c.entries.map((e) => e.type))];
    const fcs = [...new Set(c.entries.map((e) => e.fc).filter(Boolean))];
    const infoCard = `<div class="card">
        <div class="entry-head">${types.map((t) => `<span class="tag t-${esc(t)}">${esc(t)}</span>`).join("")}
          <span class="hint">${esc(fcs.length ? "담당 FC: " + fcs.join(", ") : "")}</span></div>
        ${rows.length ? `<dl class="kv">${rows.map(([k, vs]) => `<dt>${esc(k)}</dt><dd>${vs.map(esc).join("<br>")}</dd>`).join("")}</dl>`
          : '<p class="hint">고객리스트에 기본정보가 없습니다.</p>'}
      </div>`;

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
    const prim = (x.fields || []).filter((k) => EXTRA.primary.includes(k));
    const more = (x.fields || []).filter((k) => !EXTRA.primary.includes(k));
    const moreFilled = more.filter((k) => (ev[k] || "").trim()).length;
    const extraCard = `<div class="card">
        <h3 class="sec-title">추가 정보</h3>
        ${x.missing ? `<p class="hint">'${esc(EXTRA.fileKeyword)}' 엑셀 파일을 찾지 못했습니다.</p>` :
          x.error ? '<p class="hint">추가 정보 엑셀을 읽지 못했습니다.</p>' : `
          ${prim.map(field).join("")}
          ${more.length ? `
            <button type="button" class="more-btn" id="exMore" aria-expanded="${extraOpen}">
              <span>${extraOpen ? "추가 항목 접기" : "추가 항목 펼치기"}</span>
              <span class="hint">${esc(more.join(" · "))}${moreFilled ? ` (${moreFilled}개 입력됨)` : ""}</span>
              <span class="chev">${extraOpen ? "▴" : "▾"}</span>
            </button>
            <div id="exMoreBox" ${extraOpen ? "" : "hidden"}>${more.map(field).join("")}</div>` : ""}
          <datalist id="carrierList"><option value="SKT"><option value="KT"><option value="LG U+"><option value="알뜰폰"></datalist>
          <div class="file-actions"><button class="primary" id="exSave" disabled>저장</button></div>`}
      </div>`;
    return infoCard + extraCard;
  }

  function bindExtra() {
    const btn = $("exSave");
    if (!btn) return;
    const more = $("exMore");
    if (more) more.addEventListener("click", () => {
      extraOpen = !extraOpen;
      $("exMoreBox").hidden = !extraOpen;
      more.setAttribute("aria-expanded", extraOpen);
      more.querySelector("span").textContent = extraOpen ? "추가 항목 접기" : "추가 항목 펼치기";
      more.querySelector(".chev").textContent = extraOpen ? "▴" : "▾";
    });
    // 월/일 선택 → 'MMDD'
    $("tabBody").querySelectorAll(".mmdd").forEach((box) => {
      const mm = box.querySelector("[data-mm]"), dd = box.querySelector("[data-dd]"), hid = box.querySelector("[data-ex]");
      const upd = () => {
        hid.value = mm.value && dd.value ? pad2(mm.value) + pad2(dd.value) : (!mm.value && !dd.value ? "" : hid.value);
        hid.dispatchEvent(new Event("input"));
      };
      mm.addEventListener("change", upd);
      dd.addEventListener("change", upd);
    });
    const inputs = [...$("tabBody").querySelectorAll("[data-ex]")];
    const vals = () => Object.fromEntries(inputs.map((i) => [i.dataset.ex, i.value.trim()]));
    const orig = { ...(current.extra.values || {}) };
    const dirty = () => inputs.some((i) => i.value.trim() !== (orig[i.dataset.ex] || ""));
    inputs.forEach((i) => i.addEventListener("input", () => { btn.disabled = !dirty(); }));
    btn.addEventListener("click", async () => {
      btn.disabled = true;
      btn.textContent = "저장 중…";
      const v = vals();
      try {
        await saveExtraFor(current.fc, current.cust.name, v, orig);
        current.extra = { ...current.extra, found: true, values: v };
        current.phones = null; // 연락하기 탭 다시 계산
        toast("추가 정보를 엑셀에 저장했습니다.");
        selectTab("basic"); // 통신사·차량 정보 반영해서 다시 표시
      } catch (e) {
        btn.disabled = false;
        btn.textContent = "저장";
        saveError(e);
      }
    });
  }

  function saveError(e) {
    const m = String(e && e.message);
    if (/423|locked/i.test(m)) return toast("엑셀 파일이 다른 곳에서 열려 있어 저장하지 못했습니다. 닫고 다시 시도하세요.");
    if (/403/.test(m)) return toast("이 아이디는 엑셀을 수정할 권한이 없습니다.");
    handleError(e);
  }

  // 상담기록: '상담기록' 시트에서 이 고객(고객명+담당 FC) 기록, 최근 날짜가 위로 + 새 기록 입력
  function consultHtml() {
    const x = current.extra || {};
    if (x.missing) return `<p class="hint pad">'${esc(EXTRA.fileKeyword)}' 엑셀 파일을 찾지 못했습니다.</p>`;
    const list = x.consults || [];
    return `<div class="card">
        <h3 class="sec-title">상담기록</h3>
        ${list.length ? `<ul class="consult-list">${list.map((c) => `<li><b>${esc(c.date)}</b>: ${esc(c.text).replace(/\n/g, "<br>")}</li>`).join("")}</ul>`
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
    ta.addEventListener("input", () => { btn.disabled = !ta.value.trim(); });
    btn.addEventListener("click", async () => {
      btn.disabled = true;
      btn.textContent = "저장 중…";
      try {
        const list = await addConsult(current.fc, current.cust.name, ta.value.trim());
        current.extra = { ...current.extra, consults: list };
        toast("상담기록을 엑셀에 저장했습니다.");
        renderTabs();
        selectTab("consult");
      } catch (e) {
        btn.disabled = false;
        btn.textContent = "저장";
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

  function contractHtml() {
    const list = current.contracts;
    if (!list.length) {
      const lines = [...new Set(current.cust.entries.flatMap((e) =>
        e.fields.filter(([k]) => k.includes("상품명")).flatMap(([, v]) => v.split(/\n/)).map((s) => s.trim()).filter(Boolean)))];
      return lines.length ? lines.map((l) => {
          const product = l.replace(/\s*-\s*[\d,]+원\s*$/, "");
          return `<div class="card"><div class="contract-title">${esc(l)}</div>${pdfButton(findContractPdf("", product))}</div>`;
        }).join("") + '<p class="hint pad">계약리스트·이관리스트에서 이 고객의 계약을 찾지 못해 고객리스트의 상품 정보를 표시합니다.</p>'
        : '<p class="hint pad">계약리스트·이관리스트에서 이 고객의 계약을 찾지 못했습니다.</p>';
    }
    return list.map((c) => {
      const status = c.row["계약상태"] || c.row["상태"] || "";
      const val = (cols) => { for (const k of cols) if (c.row[k]) return c.row[k]; return ""; };
      const shown = CONTRACT_SHOW.map(([label, cols]) => [label, val(cols)]).filter(([, v]) => v);
      const pdf = findContractPdf(val(["최초계약일", "계약일"]), c.row["상품명"]);
      return `<div class="card">
        <div class="entry-head"><span class="tag t-${esc(c.type)}">${esc(c.type)}</span>
          ${status ? `<span class="tag">${esc(status)}</span>` : ""}
          <span class="hint">${esc([c.fc, c.type === "이관" ? c.sheet : ""].filter(Boolean).join(" · "))}</span></div>
        <div class="contract-title">${esc(c.row["상품명"] || "(상품명 없음)")}</div>
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
    // 참고용: OneDrive에서 '고객' 폴더(또는 상위 폴더)를 누구와 공유했는지
    const shares = [];
    try {
      const p = (await graph(itemPath(base.id) + "/permissions")).value || [];
      p.filter((x) => !(x.roles || []).includes("owner")).forEach((x) =>
        shares.push({ email: emailOf(x), name: nameOf(x), link: !!x.link, inherited: !!x.inheritedFrom }));
    } catch (e) { /* 공유 목록을 못 읽어도 권한 관리는 가능 */ }
    adminState = { fcs, perms, shares };
  }

  async function openAdmin() {
    if (!isAdmin()) return show("searchView");
    show("adminView");
    $("adminBody").innerHTML = '<p class="hint pad">권한 정보를 불러오는 중…</p>';
    try {
      await loadAdminState();
      renderAdmin();
    } catch (e) {
      $("adminBody").innerHTML = "";
      handleError(e);
    }
  }

  function renderAdmin() {
    const { fcs, perms, shares } = adminState;
    const users = [...new Set([...Object.keys(perms.users), ...extraUsers])].sort();
    const listOf = (em) => perms.users[em] || [];
    const isAll = (em) => listOf(em).includes("*");
    const has = (em, fc) => isAll(em) || listOf(em).includes(fc);
    const shared = (em) => shares.some((s) => s.email === em);

    $("adminBody").innerHTML = `
      <div class="card">
        <h3 class="sec-title">현재 매칭 현황</h3>
        <p class="hint">체크박스로 바로 바꾸고 '변경 저장'을 누르세요.</p>
        <div class="table-wrap"><table class="matrix">
          <tr><th>아이디</th><th>전체</th>${fcs.map((fc) => `<th>${esc(fc)}</th>`).join("")}</tr>
          ${(CFG.admins || []).map((a) => `<tr class="admin-row"><td>${esc(a)} <span class="tag">관리자</span></td>
            <td class="ok">●</td>${fcs.map(() => `<td class="ok">●</td>`).join("")}</tr>`).join("")}
          ${users.map((em) => `<tr class="perm-row" data-em="${esc(em)}">
            <td>${esc(em)}${shared(em) ? "" : ' <span class="warn-dot" title="OneDrive에서 고객 폴더 공유가 확인되지 않음">⚠️</span>'}</td>
            <td><input type="checkbox" class="cb" data-all ${isAll(em) ? "checked" : ""} aria-label="${esc(em)} 전체"></td>
            ${fcs.map((fc) => `<td><input type="checkbox" class="cb" data-fc="${esc(fc)}" ${has(em, fc) ? "checked" : ""} aria-label="${esc(em)} ${esc(fc)}"></td>`).join("")}
          </tr>`).join("")}
        </table></div>
        ${users.length ? "" : '<p class="hint">아직 권한을 준 사용자가 없습니다. 아래에서 이메일을 추가하세요.</p>'}
        <div class="file-actions">
          <button class="secondary" id="resetPerms" disabled>되돌리기</button>
          <button class="primary" id="savePerms" disabled>변경 저장</button>
        </div>
        <p class="hint">'전체'는 나중에 생기는 FC 폴더도 포함 · 모두 체크 해제 후 저장하면 그 아이디는 목록에서 빠집니다 · ⚠️ OneDrive '고객' 폴더 공유가 확인되지 않은 아이디</p>
      </div>

      <div class="card">
        <h3 class="sec-title">사용자 추가</h3>
        <form id="addUserForm" class="search">
          <input id="newUser" type="email" placeholder="Microsoft 계정 이메일" autocomplete="off" required>
          <button class="primary" type="submit">추가</button>
        </form>
        <p class="hint">추가하면 위 표에 줄이 생깁니다. 권한을 체크하고 '변경 저장'을 누르세요.</p>
      </div>

      <div class="card">
        <h3 class="sec-title">OneDrive '고객' 폴더 공유 현황 (참고)</h3>
        ${shares.length ? `<ul class="plain">${shares.map((s) => `<li>${esc(s.email || s.name || "(알 수 없음)")}${s.link ? " · 링크 공유" : ""}${s.inherited ? " · 상위 폴더에서 공유" : ""}</li>`).join("")}</ul>`
          : '<p class="hint">아직 아무와도 공유되지 않았습니다.</p>'}
        <p class="hint">다른 사용자가 앱을 쓰려면 OneDrive에서 '${esc(CFG.baseFolder)}' 폴더를 그 아이디와 한 번 공유(보기 가능)해야 합니다. 어떤 FC를 보여줄지는 위 표의 앱 권한으로 정합니다.</p>
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
  $("searchForm").addEventListener("submit", (ev) => { ev.preventDefault(); $("q").blur(); search($("q").value); });
  $("backBtn").addEventListener("click", () => history.back());
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
  // 큰 메뉴: 일정 / 고객 / 관리자
  $("mainNav").querySelectorAll("button").forEach((b) => b.addEventListener("click", () => {
    const s = b.dataset.section;
    if (s === "schedule") show("scheduleView");
    else if (s === "admin") openAdmin();
    else show("searchView");
  }));
  window.addEventListener("popstate", () => { if (!$("customerView").hidden) show("searchView"); });

  // 검색 화면에 "검색 가능한 FC" 표시
  async function showScope() {
    try {
      const { fcs } = await scanBase();
      $("scopeHint").textContent = fcs.length
        ? "검색 가능한 FC: " + fcs.join(", ") + ((await scanBase()).all ? " (전체)" : "")
        : "검색 가능한 FC 폴더가 없습니다. 관리자에게 권한을 요청하세요.";
    } catch (e) { /* handleError에서 처리 */ }
  }

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
        show("searchView");
      }
      loadLists().then(showScope).catch(handleError); // 미리 불러오기
    } else {
      show("loginView");
    }
  })();
})();
