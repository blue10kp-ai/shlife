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
    const items = await listChildren(itemPath(base.id) + "/children");
    folderCache = items.filter((i) => i.folder);
    return folderCache;
  }

  async function search(q) {
    q = q.trim();
    const status = $("searchStatus");
    const list = $("results");
    list.innerHTML = "";
    if (!q) { status.textContent = ""; return; }
    status.textContent = "찾는 중…";
    try {
      const folders = await loadCustomerFolders();
      const norm = (s) => s.replace(/\s+/g, "").toLowerCase();
      const hits = folders.filter((f) => norm(f.name).includes(norm(q)));
      status.textContent = hits.length ? hits.length + "명 찾음" : "'" + q + "' 고객 폴더가 없습니다.";
      if (hits.length === 1) return openCustomer(hits[0]);
      list.innerHTML = hits.map((f, i) =>
        `<li><button data-i="${i}"><span class="avatar sm">${esc(f.name[0])}</span>${esc(f.name)}<span class="chev">›</span></button></li>`
      ).join("");
      list.querySelectorAll("button").forEach((b) =>
        b.addEventListener("click", () => openCustomer(hits[+b.dataset.i])));
    } catch (e) {
      status.textContent = "";
      handleError(e);
    }
  }

  // ---------- 고객 화면 ----------
  let current = null; // { folder, files, phones }

  async function openCustomer(folder) {
    show("customerView");
    $("custName").textContent = folder.name;
    $("avatar").textContent = folder.name[0];
    $("folderLink").href = folder.webUrl;
    $("tabBody").innerHTML = '<p class="hint pad">파일을 불러오는 중…</p>';
    $("tabs").innerHTML = "";
    history.pushState({ v: "customer" }, "");
    try {
      const files = await collectFiles(folder.id, "", 0);
      current = { folder, files, phones: null, sheetCache: {} };
      renderTabs();
      selectTab("basic");
    } catch (e) {
      handleError(e);
    }
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
      const n = t.id === "contact" || t.id === "all" ? "" : filesForTab(t).length;
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
    if (!files.length) {
      body.innerHTML = `<p class="hint pad">이름에 ${tab.keywords.map((k) => "'" + esc(k) + "'").join(", ")}이(가) 들어간 파일이 없습니다.<br>'전체파일'에서 찾아보세요.</p>`;
      return;
    }
    body.innerHTML = files.map((f, i) => {
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

  async function readWorkbook(f) {
    if (current.sheetCache[f.id]) return current.sheetCache[f.id];
    const buf = await graph(itemPath(f.id) + "/content", true);
    const wb = XLSX.read(buf, { type: "array" });
    current.sheetCache[f.id] = wb;
    return wb;
  }

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
      loadCustomerFolders().catch(handleError); // 미리 불러오기
    } else {
      show("loginView");
    }
  })();
})();
