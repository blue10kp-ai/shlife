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

  // ---------- 화면 전환 ----------
  const views = ["loginView", "searchView", "customerView"];
  function show(view) {
    views.forEach((v) => ($(v).hidden = v !== view));
    $("backBtn").hidden = view !== "customerView";
    $("logoutBtn").hidden = view === "loginView";
    $("title").textContent = view === "customerView" ? "고객 정보" : "고객 찾기";
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
  // 고객 폴더 위치 찾기: ① config의 공유 폴더 → ② 내 OneDrive 경로 → ③ 나에게 공유된 폴더
  let base = null; // { driveId, id }
  async function resolveBase() {
    if (base) return base;
    const sf = CFG.sharedFolder || {};
    if (sf.driveId && sf.itemId) return (base = { driveId: sf.driveId, id: sf.itemId });
    try {
      const path = "/me/drive/root:/" + CFG.baseFolder.split("/").map(encodeURIComponent).join("/");
      const it = await graph(path + "?$select=id,parentReference");
      return (base = { driveId: it.parentReference.driveId, id: it.id });
    } catch (e) {
      if (!String(e.message).includes("404")) throw e;
    }
    const leaf = CFG.baseFolder.split("/").pop();
    const shared = await graph("/me/drive/sharedWithMe");
    const hit = (shared.value || []).find((s) => s.remoteItem && s.remoteItem.folder && s.name === leaf);
    if (!hit) throw new Error("Graph 404: base folder not found");
    return (base = { driveId: hit.remoteItem.parentReference.driveId, id: hit.remoteItem.id });
  }
  const itemPath = (id) => "/drives/" + base.driveId + "/items/" + id;

  let folderCache = null;
  async function loadCustomerFolders() {
    if (folderCache) return folderCache;
    await resolveBase();
    const items = (await listChildren(itemPath(base.id) + "/children")).filter((i) => i.folder);
    // "_계약", "_이관", "_진행"처럼 _로 시작하는 폴더는 분류 폴더 → 그 안의 폴더가 고객
    const groups = items.filter((i) => i.name.startsWith("_"));
    const direct = items.filter((i) => !i.name.startsWith("_")).map((f) => ({ ...f, group: "" }));
    const nested = await Promise.all(groups.map(async (g) =>
      (await listChildren(itemPath(g.id) + "/children"))
        .filter((i) => i.folder)
        .map((f) => ({ ...f, group: g.name.slice(1) }))));
    folderCache = direct.concat(...nested);
    return folderCache;
  }

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
    await resolveBase();
    const files = (await listChildren(itemPath(base.id) + "/children")).filter((f) => f.file);
    const sources = CFG.lists.flatMap((l) =>
      files.filter((f) => f.name.includes(l.fileKeyword)).map((f) => ({ ...l, file: f })));
    const out = [];
    await Promise.all(sources.map(async (s) => {
      const wb = await readBook(s.file.id);
      const fc = (s.file.name.split(" - ")[1] || "").trim();
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
      const folders = (await loadCustomerFolders()).filter((f) => norm(custPart(f.name)) === norm(cust.name));
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

  // ---------- 오류 처리 ----------
  function handleError(e) {
    console.error(e);
    const m = String(e && e.message);
    if (m === "NOT_SIGNED_IN") return show("loginView");
    if (m.includes("404")) return toast("OneDrive에서 '" + CFG.baseFolder + "' 폴더를 찾을 수 없습니다.");
    toast("오류: " + m.slice(0, 120));
  }

  // ---------- 시작 ----------
  $("loginBtn").addEventListener("click", () => msal.loginRedirect({ scopes: SCOPES }));
  $("logoutBtn").addEventListener("click", () => msal.logoutRedirect());
  $("searchForm").addEventListener("submit", (ev) => { ev.preventDefault(); $("q").blur(); search($("q").value); });
  $("backBtn").addEventListener("click", () => history.back());
  $("modes").querySelectorAll("button").forEach((b) => b.addEventListener("click", () => {
    mode = b.dataset.mode;
    $("modes").querySelectorAll("button").forEach((x) => x.classList.toggle("on", x === b));
    if ($("q").value.trim()) search($("q").value);
  }));
  window.addEventListener("popstate", () => { if (!$("customerView").hidden) show("searchView"); });

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
      show("searchView");
      $("q").focus();
      loadLists().then(() => loadCustomerFolders()).catch(handleError); // 미리 불러오기
    } else {
      show("loginView");
    }
  })();
})();
