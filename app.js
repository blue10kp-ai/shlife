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
      headers: { Authorization: "Bearer " + token }
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

  async function gx(method, path, body) {
    const acc = account();
    let token;
    try {
      token = (await msal.acquireTokenSilent({ scopes: ADMIN_SCOPES, account: acc })).accessToken;
    } catch (e) {
      sessionStorage.setItem("openAdmin", "1");
      await msal.acquireTokenRedirect({ scopes: ADMIN_SCOPES, account: acc });
      throw e;
    }
    const res = await fetch(GRAPH + path, {
      method,
      headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
      body: body ? JSON.stringify(body) : undefined
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
  //   소유자(관리자): 내 OneDrive의 Claude/고객 → 전체 검색
  //   다른 사용자: 소유자의 고객 폴더 중 "나에게 공유된 FC 폴더"만 검색
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

  // 다른 사용자가 접근할 수 있는 FC 폴더 후보 (config 목록 + 나에게 공유된 폴더)
  async function sharedFcCandidates() {
    const cands = new Map((CFG.fcFolders || []).map((f) => [f.itemId, f.name]));
    try {
      const shared = await graph("/me/drive/sharedWithMe");
      (shared.value || []).forEach((s) => {
        const r = s.remoteItem;
        if (r && r.folder && r.parentReference && r.parentReference.id === base.id) cands.set(r.id, r.name);
      });
    } catch (e) { /* 목록을 못 읽어도 config 목록으로 진행 */ }
    return [...cands];
  }

  // 고객 폴더 구조 읽기
  //   고객/김현성/_계약/<고객폴더>, 고객/김현성/계약고객리스트.xlsx   ← FC별 폴더 (현재 구조)
  //   고객/_계약/<고객폴더>, 고객/계약고객리스트.xlsx               ← FC 폴더 없이 (예전 구조)
  // "_"로 시작하는 폴더 = 분류(계약/이관/진행), 그 안의 폴더 = 고객 폴더
  let scanCache = null; // Promise<{ folders: [...], files: [...] }>
  function scanBase() {
    if (scanCache) return scanCache;
    scanCache = (async () => {
      await resolveBase();
      const folders = [], files = [];
      async function scan(id, fc) {
        const items = await listChildren(itemPath(id) + "/children");
        items.filter((i) => i.file).forEach((f) => files.push({ ...f, fc }));
        await Promise.all(items.filter((i) => i.folder).map(async (d) => {
          if (d.name.startsWith("_")) {
            (await listChildren(itemPath(d.id) + "/children"))
              .filter((i) => i.folder)
              .forEach((f) => folders.push({ ...f, group: d.name.slice(1), fc }));
          } else if (!fc) {
            await scan(d.id, d.name); // FC 폴더
          }
        }));
      }
      const fcs = [];
      if (base.owner) {
        await scan(base.id, "");
      } else {
        // 고객 폴더 전체를 공유받았으면 전체, 아니면 공유받은 FC 폴더만 (권한 없는 폴더는 OneDrive가 거부)
        let whole = false;
        try { await graph(itemPath(base.id) + "?$select=id"); whole = true; } catch (e) { /* 전체 공유 아님 */ }
        if (whole) await scan(base.id, "");
        else await Promise.all((await sharedFcCandidates()).map(([id, name]) =>
          scan(id, name).then(() => fcs.push(name)).catch(() => {})));
      }
      folders.forEach((f) => f.fc && !fcs.includes(f.fc) && fcs.push(f.fc));
      files.forEach((f) => f.fc && !fcs.includes(f.fc) && fcs.push(f.fc));
      return { folders, files, fcs: fcs.sort((a, b) => a.localeCompare(b, "ko")) };
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

  // ---------- 고객 화면 ----------
  let current = null; // { cust, folders, files, phones }

  async function openCustomer(cust) {
    show("customerView");
    $("custName").textContent = cust.name;
    $("avatar").textContent = cust.name[0];
    $("custSub").textContent = [cust.birth, ...new Set(cust.entries.map((e) => e.type))].filter(Boolean).join(" · ");
    $("folderLink").hidden = true;
    $("tabBody").innerHTML = '<p class="hint pad">고객 폴더를 찾는 중…</p>';
    $("tabs").innerHTML = "";
    history.pushState({ v: "customer" }, "");
    current = { cust, folders: [], files: [], phones: null };
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
      if (folders[0]) { $("folderLink").href = folders[0].webUrl; $("folderLink").hidden = false; }
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
    return current.files.filter((f) => kws.some((k) => f.path.toLowerCase().includes(k)));
  }

  function renderTabs() {
    $("tabs").innerHTML = CFG.tabs.map((t) => {
      const n = t.id === "contact" || t.id === "all" ? "" :
        filesForTab(t).length + (t.id === "basic" ? current.cust.entries.length : 0);
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

    const files = filesForTab(tab);
    // 기본정보: 고객리스트 엑셀에서 찾은 내용을 먼저 보여줌
    const listHtml = id === "basic" ? current.cust.entries.map((e) => `<div class="card">
        <div class="entry-head"><span class="tag t-${esc(e.type)}">${esc(e.type)}</span>
          <span class="hint">${esc([e.fc, e.sheet].filter(Boolean).join(" · "))}</span></div>
        <dl class="kv">${e.fields.map(([k, v]) => `<dt>${esc(k)}</dt><dd>${esc(v)}</dd>`).join("")}</dl>
      </div>`).join("") : "";
    const noFolder = !current.folders.length ? '<p class="hint pad">이 고객 이름의 폴더를 찾지 못했습니다.</p>' : "";
    if (!files.length) {
      body.innerHTML = listHtml + (noFolder || (id === "basic" && listHtml ? "" :
        `<p class="hint pad">이름에 ${tab.keywords.map((k) => "'" + esc(k) + "'").join(", ")}이(가) 들어간 파일이 없습니다.<br>'전체파일'에서 찾아보세요.</p>`));
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
          <a class="secondary" href="${esc(f.webUrl)}" target="_blank" rel="noopener">열기</a>
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
      return `<div class="card">
        <div class="phone">${pretty}</div>
        <div class="file-actions">
          <a class="primary" href="tel:${p}">전화</a>
          <a class="secondary" href="sms:${p}">문자</a>
        </div>
      </div>`;
    }).join("");
  }

  // ---------- 관리자 페이지: 아이디별 FC 폴더 공유 ----------
  const emailOf = (p) => String(
    (p.invitation && p.invitation.email) ||
    (p.grantedToV2 && p.grantedToV2.user && p.grantedToV2.user.email) ||
    (p.grantedTo && p.grantedTo.user && p.grantedTo.user.email) ||
    (p.grantedToIdentitiesV2 && p.grantedToIdentitiesV2[0] && p.grantedToIdentitiesV2[0].user && p.grantedToIdentitiesV2[0].user.email) ||
    (p.grantedToV2 && p.grantedToV2.user && p.grantedToV2.user.displayName) || ""
  ).toLowerCase();

  let adminState = null; // { fcs: [{id,name}], users: Map(email → { [fcId]: perm }), warnings: [] }
  const extraUsers = new Set();

  async function loadAdminState() {
    await resolveBase();
    const me = String(account().username).toLowerCase();
    const items = await listChildren(itemPath(base.id) + "/children");
    const fcs = items.filter((i) => i.folder && !i.name.startsWith("_")).map((f) => ({ id: f.id, name: f.name }));
    const users = new Map();
    const warnings = [];
    // 상위 폴더(Claude, 고객)가 공유되어 있으면 FC 제한이 소용없으므로 경고
    const baseItem = await graph(itemPath(base.id) + "?$select=id,name,parentReference");
    const parents = [{ id: base.id, name: CFG.baseFolder }];
    if (baseItem.parentReference && baseItem.parentReference.id) parents.push({ id: baseItem.parentReference.id, name: CFG.baseFolder.split("/")[0] });
    for (const p of parents) {
      const perms = (await gx("GET", itemPath(p.id) + "/permissions")).value || [];
      perms.filter((x) => !(x.roles || []).includes("owner")).forEach((x) =>
        warnings.push(`'${p.name}' 폴더 전체가 ${x.link ? "링크로" : (emailOf(x) || "다른 사용자") + "에게"} 공유되어 있습니다.`));
    }
    await Promise.all(fcs.map(async (fc) => {
      const perms = (await gx("GET", itemPath(fc.id) + "/permissions")).value || [];
      perms.forEach((p) => {
        if ((p.roles || []).includes("owner")) return;
        if (p.link) { warnings.push(`'${fc.name}' 폴더가 링크로 공유되어 있습니다(링크를 가진 사람 모두 접근).`); return; }
        const em = emailOf(p);
        if (!em || em === me) return;
        if (!users.has(em)) users.set(em, {});
        users.get(em)[fc.id] = p;
      });
    }));
    extraUsers.forEach((em) => { if (!users.has(em)) users.set(em, {}); });
    adminState = { fcs, users, warnings };
  }

  async function openAdmin() {
    if (!isAdmin()) return show("searchView");
    show("adminView");
    $("adminBody").innerHTML = '<p class="hint pad">공유 상태를 불러오는 중…</p>';
    try {
      await loadAdminState();
      renderAdmin();
    } catch (e) {
      $("adminBody").innerHTML = "";
      handleError(e);
    }
  }

  function renderAdmin() {
    const { fcs, users, warnings } = adminState;
    const rows = [...users.keys()].sort();
    const has = (em, fc) => !!users.get(em)[fc.id];
    const allOf = (em) => fcs.length > 0 && fcs.every((fc) => has(em, fc));

    $("adminBody").innerHTML = `
      ${warnings.length ? `<div class="card warn">${warnings.map((w) => `<p>⚠️ ${esc(w)}</p>`).join("")}
        <p class="hint">상위 폴더 공유가 있으면 FC별 제한이 적용되지 않습니다. OneDrive에서 해당 공유를 해제해 주세요.</p></div>` : ""}

      <div class="card">
        <h3 class="sec-title">현재 매칭 현황</h3>
        <div class="table-wrap"><table class="matrix">
          <tr><th>아이디</th>${fcs.map((fc) => `<th>${esc(fc.name)}</th>`).join("")}</tr>
          ${(CFG.admins || []).map((a) => `<tr><td>${esc(a)} <span class="tag">관리자</span></td>${fcs.map(() => `<td class="ok">●</td>`).join("")}</tr>`).join("")}
          ${rows.filter((em) => fcs.some((fc) => has(em, fc))).map((em) => `<tr><td>${esc(em)}</td>${fcs.map((fc) =>
            `<td class="${has(em, fc) ? "ok" : "no"}">${has(em, fc) ? "●" : "–"}</td>`).join("")}</tr>`).join("")}
        </table></div>
        <p class="hint">● 검색 가능 · – 접근 불가 · 관리자는 항상 전체</p>
      </div>

      <div class="card">
        <h3 class="sec-title">사용자 추가</h3>
        <form id="addUserForm" class="search">
          <input id="newUser" type="email" placeholder="Microsoft 계정 이메일" autocomplete="off" required>
          <button class="primary" type="submit">추가</button>
        </form>
      </div>

      <h3 class="sec-title pad-x">아이디별 FC 폴더 권한</h3>
      ${rows.length ? rows.map((em) => `<div class="card user-card" data-em="${esc(em)}">
          <div class="user-head"><span class="avatar sm">${esc(em[0].toUpperCase())}</span><span class="r-name">${esc(em)}</span></div>
          <label class="chk all"><input type="checkbox" data-all ${allOf(em) ? "checked" : ""}> 전체</label>
          <div class="chk-grid">${fcs.map((fc) => {
            const p = users.get(em)[fc.id];
            const inherited = p && p.inheritedFrom;
            return `<label class="chk"><input type="checkbox" data-fc="${esc(fc.id)}" ${p ? "checked" : ""} ${inherited ? "disabled" : ""}> ${esc(fc.name)}${inherited ? " (상위 공유)" : ""}</label>`;
          }).join("")}</div>
          <div class="file-actions"><button class="primary" data-save disabled>저장</button></div>
        </div>`).join("") : '<p class="hint pad">아직 권한을 준 사용자가 없습니다. 위에서 이메일을 추가하세요.</p>'}
      <p class="hint pad">저장하면 체크한 FC 폴더가 OneDrive에서 해당 아이디와 읽기 전용으로 공유되고, 체크를 뺀 폴더는 공유가 해제됩니다. 새로 공유하면 OneDrive가 상대방에게 알림 메일을 보냅니다. '전체'는 지금 있는 FC 폴더 모두를 체크합니다.</p>`;

    $("addUserForm").addEventListener("submit", (ev) => {
      ev.preventDefault();
      const em = $("newUser").value.trim().toLowerCase();
      if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(em)) return toast("이메일 형식을 확인해 주세요.");
      if ((CFG.admins || []).map((a) => a.toLowerCase()).includes(em)) return toast("관리자는 이미 모든 폴더를 볼 수 있습니다.");
      extraUsers.add(em);
      if (!users.has(em)) users.set(em, {});
      renderAdmin();
    });

    $("adminBody").querySelectorAll(".user-card").forEach((card) => {
      const em = card.dataset.em;
      const boxes = [...card.querySelectorAll("input[data-fc]")];
      const all = card.querySelector("input[data-all]");
      const save = card.querySelector("[data-save]");
      const dirty = () => boxes.some((b) => b.checked !== !!users.get(em)[b.dataset.fc]);
      const sync = () => {
        all.checked = boxes.every((b) => b.checked);
        all.indeterminate = !all.checked && boxes.some((b) => b.checked);
        save.disabled = !dirty();
      };
      all.addEventListener("change", () => { boxes.forEach((b) => { if (!b.disabled) b.checked = all.checked; }); sync(); });
      boxes.forEach((b) => b.addEventListener("change", sync));
      sync();
      save.addEventListener("click", () => saveUser(em, boxes, save));
    });
  }

  async function saveUser(em, boxes, btn) {
    const perms = adminState.users.get(em);
    const name = (id) => adminState.fcs.find((f) => f.id === id).name;
    const add = boxes.filter((b) => b.checked && !perms[b.dataset.fc]).map((b) => b.dataset.fc);
    const del = boxes.filter((b) => !b.checked && perms[b.dataset.fc] && !perms[b.dataset.fc].inheritedFrom).map((b) => b.dataset.fc);
    if (!add.length && !del.length) return;
    const msg = [`${em} 권한을 변경할까요?`,
      add.length ? "공유 추가: " + add.map(name).join(", ") + "\n(OneDrive가 상대방에게 알림 메일을 보냅니다)" : "",
      del.length ? "공유 해제: " + del.map(name).join(", ") : ""].filter(Boolean).join("\n");
    if (!confirm(msg)) return;
    btn.disabled = true;
    btn.textContent = "저장 중…";
    const failed = [];
    for (const id of add) {
      try {
        await gx("POST", itemPath(id) + "/invite", {
          recipients: [{ email: em }], roles: ["read"], requireSignIn: true, sendInvitation: true,
          message: "고객찾기 앱(shlife)에서 '" + name(id) + "' 고객 폴더를 검색할 수 있도록 공유합니다."
        });
      } catch (e) { console.error(e); failed.push(name(id) + " 추가"); }
    }
    for (const id of del) {
      try { await gx("DELETE", itemPath(id) + "/permissions/" + perms[id].id); }
      catch (e) { console.error(e); failed.push(name(id) + " 해제"); }
    }
    extraUsers.add(em);
    toast(failed.length ? "일부 실패: " + failed.join(", ") : "저장했습니다.");
    try { await loadAdminState(); renderAdmin(); } catch (e) { handleError(e); }
  }

  // ---------- 오류 처리 ----------
  function handleError(e) {
    console.error(e);
    const m = String(e && e.message);
    if (m === "NOT_SIGNED_IN") return show("loginView");
    if (m.includes("404")) return toast("OneDrive에서 '" + CFG.baseFolder + "' 폴더를 찾을 수 없습니다.");
    if (m.includes("403")) return toast("권한이 없습니다. 관리자에게 요청하세요.");
    toast("오류: " + m.slice(0, 120));
  }

  // ---------- 시작 ----------
  $("loginBtn").addEventListener("click", () => msal.loginRedirect({ scopes: SCOPES }));
  $("logoutBtn").addEventListener("click", () => msal.logoutRedirect());
  $("searchForm").addEventListener("submit", (ev) => { ev.preventDefault(); $("q").blur(); search($("q").value); });
  $("backBtn").addEventListener("click", () => history.back());
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
        ? "검색 가능한 FC: " + fcs.join(", ") + (base.owner ? " (전체)" : "")
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
