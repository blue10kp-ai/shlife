/* 날짜 달력 공통 모듈 (규칙: app\_공통\날짜달력.md, 2026-11-24 사용자 지시 "날짜는 칸 옆에 뜨는 달력에서 고르기, 직접 넣지 않도록")
 * 앱 폴더에 datepick.js로 복사해 <script src="datepick.js?v=N">로 불러온다. 원본은 app\_공통\날짜달력.js 하나.
 * 쓰는 법: <input data-dp="date|md|month|year" value="저장 값">  (기간: 두 칸에 data-dp-range="이름", 끝 칸에 data-dp-end)
 *   저장 값: date = YYYY-MM-DD, md = MM-DD(data-dp-format="MMDD"면 MMDD), month = YYYY-MM, year = YYYY
 *   원래 칸은 숨기고 그 자리에 읽기 전용 칸(같은 class)을 보여 준다. 앱은 원래 칸의 value를 그대로 읽고 쓰면 된다(코드로 value를 바꿔도 보이는 칸이 따라감).
 *   값이 바뀌면 원래 칸에 input·change 이벤트를 보낸다. 색은 앱의 --surface --accent --accent-ink --accent-weak --line --t1 --t2 --t3 --bad --font를 따른다.
 */
(function () {
  "use strict";
  if (window.DatePick) return;
  const WD = "일월화수목금토", pad = (n) => String(n).padStart(2, "0");
  const css = `.dp-face{cursor:pointer;caret-color:transparent}
.dp-pop{position:absolute;z-index:200;background:var(--surface,#fff);color:var(--t1,#151a22);border:1px solid var(--line,rgba(0,0,0,.1));border-radius:14px;box-shadow:0 8px 24px rgba(0,0,0,.18);padding:10px;font-family:var(--font,sans-serif);user-select:none;-webkit-user-select:none}
.dp-h{display:flex;align-items:center;gap:6px;margin-bottom:6px}
.dp-t{font:600 15px var(--font,sans-serif);background:none;border:0;color:inherit;padding:4px 2px;cursor:pointer;min-height:0;box-shadow:none}
.dp-t small{font-size:12px;color:var(--t3,#6b7382);font-weight:500;margin-left:6px}
.dp-nv{margin-left:auto;display:flex;gap:4px}
.dp-nv button{width:30px;height:30px;min-height:0;padding:0;border-radius:8px;border:1px solid var(--line,rgba(0,0,0,.1));background:var(--surface,#fff);color:var(--t2,#4b5361);font-size:15px;line-height:1;cursor:pointer;box-shadow:none}
.dp-wd,.dp-g{display:grid;grid-template-columns:repeat(7,1fr);text-align:center}
.dp-wd span{font-size:11px;color:var(--t3,#6b7382);padding:2px 0}.dp-wd span:first-child{color:var(--bad,#c0392b)}
.dp-g button{height:36px;min-height:0;padding:0;border:0;background:none;color:inherit;font:14px var(--font,sans-serif);font-variant-numeric:tabular-nums;border-radius:10px;cursor:pointer;box-shadow:none}
.dp-g button.o{color:var(--t3,#6b7382);opacity:.5}
.dp-g button.t{box-shadow:inset 0 0 0 1px var(--accent,#1f4e8c);color:var(--accent,#1f4e8c)}
.dp-g button.s,.dp-m button.s{background:var(--accent,#1f4e8c);color:var(--accent-ink,#fff);font-weight:600}
.dp-g button.r{background:var(--accent-weak,rgba(31,78,140,.08));border-radius:0}
.dp-g button.k,.dp-m button.k{outline:2px solid var(--accent,#1f4e8c);outline-offset:-2px}
.dp-g button.r.s{background:var(--accent,#1f4e8c);color:var(--accent-ink,#fff);font-weight:600}
.dp-g button.rs{border-radius:10px 0 0 10px}.dp-g button.re{border-radius:0 10px 10px 0}.dp-g button.rs.re{border-radius:10px}
.dp-m{display:grid;grid-template-columns:repeat(4,1fr);gap:6px}
.dp-m button{height:40px;min-height:0;padding:0;border-radius:10px;border:1px solid var(--line,rgba(0,0,0,.1));background:none;color:inherit;font:14px var(--font,sans-serif);cursor:pointer;box-shadow:none}
.dp-m button.t{border-color:var(--accent,#1f4e8c);color:var(--accent,#1f4e8c)}
.dp-f{display:flex;align-items:center;gap:6px;margin-top:6px;border-top:1px solid var(--line,rgba(0,0,0,.1));padding-top:8px}
.dp-f button{font:500 13px var(--font,sans-serif);color:var(--t2,#4b5361);background:none;border:0;padding:5px 8px;min-height:0;border-radius:8px;cursor:pointer;box-shadow:none}
.dp-f .dp-ok{margin-left:auto;background:var(--accent,#1f4e8c);color:var(--accent-ink,#fff);font-weight:600;padding:5px 14px}
.dp-f .dp-hint{font-size:11.5px;color:var(--t3,#6b7382)}`;
  const st = document.createElement("style"); st.textContent = css; document.head.appendChild(st);

  // ---- 값 읽기·쓰기 ----
  const today = () => { const d = new Date(); return { y: d.getFullYear(), m: d.getMonth() + 1, d: d.getDate() }; };
  function parse(el, v) {
    const k = el.dataset.dp, s = String(v == null ? el.value : v).trim();
    let m;
    if (k === "date") { m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(s); return m ? { y: +m[1], m: +m[2], d: +m[3] } : null; }
    if (k === "month") { m = /^(\d{4})[-.](\d{1,2})/.exec(s); return m ? { y: +m[1], m: +m[2] } : null; }
    if (k === "year") { m = /^(\d{4})/.exec(s); return m ? { y: +m[1] } : null; }
    if (k === "md") { m = /^(\d{1,2})[-./](\d{1,2})$/.exec(s) || (/^\d{3,4}$/.test(s) ? [0, s.slice(0, s.length - 2), s.slice(-2)] : null) || /(\d{1,2})\s*월\s*(\d{1,2})/.exec(s); return m ? { m: +m[1], d: +m[2] } : null; }
    return null;
  }
  function store(el, p) {
    const k = el.dataset.dp;
    if (!p) return "";
    if (k === "date") return p.y + "-" + pad(p.m) + "-" + pad(p.d);
    if (k === "month") return p.y + "-" + pad(p.m);
    if (k === "year") return String(p.y);
    if (k === "md") return el.dataset.dpFormat === "MMDD" ? pad(p.m) + pad(p.d) : pad(p.m) + "-" + pad(p.d);
    return "";
  }
  function show(el) {
    const p = parse(el), k = el.dataset.dp;
    if (!p) return el.value ? String(el.value) : "";
    if (k === "date") return p.y + "-" + pad(p.m) + "-" + pad(p.d) + " (" + WD[new Date(p.y, p.m - 1, p.d).getDay()] + ")";
    if (k === "month") return p.y + "." + pad(p.m);
    if (k === "year") return String(p.y);
    if (k === "md") return p.m + "월 " + p.d + "일";
    return "";
  }
  const desc = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value");
  function setVal(el, v) { desc.set.call(el, v); refresh(el); el.dispatchEvent(new Event("input", { bubbles: true })); el.dispatchEvent(new Event("change", { bubbles: true })); }
  function refresh(el) { const f = el._dpFace; if (f) { desc.set.call(f, show(el)); f.disabled = el.disabled; } }

  // ---- 칸 붙이기: 원래 칸은 숨기고 같은 모양의 읽기 전용 칸을 보여 줌 ----
  function attach(el) {
    if (el._dpFace || !el.dataset.dp) return;
    const f = document.createElement("input");
    f.type = "text"; f.readOnly = true; f.className = (el.className ? el.className + " " : "") + "dp-face";
    f.setAttribute("inputmode", "none"); f.setAttribute("autocomplete", "off");
    f.placeholder = el.placeholder || { date: "날짜 고르기", md: "월·일 고르기", month: "연월 고르기", year: "연도 고르기" }[el.dataset.dp] || "";
    const al = el.getAttribute("aria-label"); if (al) f.setAttribute("aria-label", al);
    if (el.style.cssText) f.style.cssText = el.style.cssText;
    f._dpFor = el; el._dpFace = f;
    el.insertAdjacentElement("afterend", f);
    el.style.display = "none"; el.tabIndex = -1;
    Object.defineProperty(el, "value", { configurable: true, get() { return desc.get.call(this); }, set(v) { desc.set.call(this, v); refresh(this); } }); // 앱이 코드로 값을 바꿔도 보이는 칸이 따라감
    refresh(el);
  }
  function scan(root) { (root.querySelectorAll ? root.querySelectorAll("input[data-dp]") : []).forEach(attach); if (root.matches && root.matches("input[data-dp]")) attach(root); }
  const mo = new MutationObserver((ms) => ms.forEach((m) => m.addedNodes.forEach((n) => { if (n.nodeType === 1) scan(n); })));
  const start = () => { scan(document); mo.observe(document.body, { childList: true, subtree: true }); };
  if (document.body) start(); else document.addEventListener("DOMContentLoaded", start);

  // ---- 기간 짝 찾기 ----
  function pairOf(el) {
    const name = el.dataset.dpRange; if (!name) return null;
    const sel = 'input[data-dp-range="' + CSS.escape(name) + '"]';
    let a = el.parentElement;
    while (a && a.querySelectorAll(sel).length < 2) a = a.parentElement;
    if (!a) return null;
    const all = [...a.querySelectorAll(sel)];
    return { s: all.find((x) => !x.hasAttribute("data-dp-end")), e: all.find((x) => x.hasAttribute("data-dp-end")) };
  }

  // ---- 달력 창 ----
  let pop = null, cur = null; // cur = { el, view:"d"|"m"|"y", y, m, pair, step }
  function close() { if (pop) { pop.remove(); pop = null; } document.removeEventListener("keydown", onKey, true); window.removeEventListener("resize", close); cur = null; }
  // 키보드(PC): 화살표로 옮기고 Enter로 고름, PageUp/PageDown = 앞뒤 달(연), Esc = 닫기
  const onKey = (e) => {
    if (!cur) return;
    if (e.key === "Escape") { e.preventDefault(); const f = cur.el._dpFace; close(); if (f) f.focus(); return; }
    const c = cur, mv = { ArrowLeft: -1, ArrowRight: 1, ArrowUp: c.view === "d" ? -7 : -4, ArrowDown: c.view === "d" ? 7 : 4 }[e.key];
    if (mv == null && e.key !== "Enter" && e.key !== " " && e.key !== "PageUp" && e.key !== "PageDown") return;
    e.preventDefault();
    if (c.view === "d") {
      const yy = c.k === "md" ? 2024 : c.y, n = new Date(yy, c.m, 0).getDate();
      if (e.key === "Enter" || e.key === " ") { if (c.k === "md") return pick({ m: c.m, d: c.fd }); return pick({ y: c.y, m: c.m, d: c.fd }); }
      if (e.key === "PageUp" || e.key === "PageDown") { c.m += e.key === "PageUp" ? -1 : 1; if (c.m < 1) { c.m = 12; c.y--; } if (c.m > 12) { c.m = 1; c.y++; } c.fd = Math.min(c.fd, new Date(c.k === "md" ? 2024 : c.y, c.m, 0).getDate()); render(); place(); return; }
      let d = c.fd + mv;
      if (d < 1) { c.m--; if (c.m < 1) { c.m = 12; c.y--; } d += new Date(c.k === "md" ? 2024 : c.y, c.m, 0).getDate(); }
      else if (d > n) { d -= n; c.m++; if (c.m > 12) { c.m = 1; c.y++; } }
      c.fd = d; render(); place(); return;
    }
    if (c.view === "m") {
      if (e.key === "Enter" || e.key === " ") { c.m = c.fm; if (c.k === "month") return pick({ y: c.y, m: c.m }); c.view = "d"; c.fd = c.fd || 1; render(); place(); return; }
      if (e.key === "PageUp" || e.key === "PageDown") { if (c.k !== "md") c.y += e.key === "PageUp" ? -1 : 1; render(); return; }
      c.fm = Math.min(12, Math.max(1, (c.fm || c.m) + mv)); render(); return;
    }
    if (e.key === "Enter" || e.key === " ") { c.y = c.fy; if (c.k === "year") return pick({ y: c.y }); c.view = "m"; c.fm = c.m; render(); place(); return; }
    if (e.key === "PageUp" || e.key === "PageDown") { c.fy += e.key === "PageUp" ? -12 : 12; c.y = c.fy; render(); return; }
    c.fy = (c.fy || c.y) + mv; c.y = c.fy; render();
  };
  function open(el) {
    close();
    const k = el.dataset.dp, p = parse(el) || {}, t = today(), pair = pairOf(el);
    const ref = pair && el === pair.e && !parse(el) && pair.s ? parse(pair.s) || {} : p;
    cur = { el, k, pair, y: ref.y || t.y, m: ref.m || t.m, view: k === "month" ? "m" : k === "year" ? "y" : "d", step: pair ? (el === pair.e ? "e" : "s") : "" };
    cur.fd = ref.d || (ref.y === t.y && ref.m === t.m ? t.d : 1); cur.fm = ref.m || t.m; cur.fy = ref.y || t.y; // 키보드 자리
    pop = document.createElement("div"); pop.className = "dp-pop"; pop.setAttribute("role", "dialog"); pop.setAttribute("aria-label", "날짜 고르기");
    document.body.appendChild(pop);
    render(); place();
    document.addEventListener("keydown", onKey, true); window.addEventListener("resize", close);
  }
  function place() {
    const f = cur.el._dpFace || cur.el, r = f.getBoundingClientRect(), vw = document.documentElement.clientWidth, vh = window.innerHeight;
    const w = Math.min(300, vw - 16); pop.style.width = w + "px";
    let left = Math.min(Math.max(8, r.left), vw - w - 8);
    const h = pop.offsetHeight, below = vh - r.bottom, top = below >= h + 8 || below >= r.top ? r.bottom + 6 : r.top - h - 6;
    let tp = Math.max(8, top);
    // 말만해요 안(?embed=1)에서는 오른쪽 아래 말만해요 버튼 자리(88×88)를 비움
    const emb = document.documentElement.classList.contains("embed") || /[?&]embed=1/.test(location.search);
    if (emb && left + w > vw - 96 && tp + h > vh - 96) { if (left > 8 + 0 && vw - 96 - w >= 8) left = vw - 96 - w; else tp = Math.max(8, vh - 96 - h); }
    pop.style.left = left + window.scrollX + "px"; pop.style.top = tp + window.scrollY + "px";
  }
  const btn = (cls, txt, data) => `<button type="button" class="${cls}" ${data || ""}>${txt}</button>`;
  function render() {
    const c = cur, t = today(), k = c.k;
    const sel = parse(c.el) || {}, ps = c.pair ? parse(c.pair.s) : null, pe = c.pair ? parse(c.pair.e) : null;
    let head = "", body = "";
    if (c.view === "d") {
      head = k === "md" ? `${c.m}월<small>연도 없음</small>` : `${c.y}년 ${c.m}월`;
      const yy = k === "md" ? 2024 : c.y, first = new Date(yy, c.m - 1, 1).getDay(), n = new Date(yy, c.m, 0).getDate(), pn = new Date(yy, c.m - 1, 0).getDate();
      let g = "";
      if (k !== "md") for (let i = 0; i < first; i++) g += btn("o", pn - first + 1 + i, 'data-pm="-1" data-d="' + (pn - first + 1 + i) + '"');
      for (let d = 1; d <= n; d++) {
        const key = c.y * 10000 + c.m * 100 + d, cls = [];
        if (k !== "md" && c.y === t.y && c.m === t.m && d === t.d) cls.push("t");
        if (c.pair && ps && pe) { const a = ps.y * 10000 + ps.m * 100 + ps.d, b = pe.y * 10000 + pe.m * 100 + pe.d; if (key >= a && key <= b) cls.push("r", "s"); if (key === a) cls.push("rs"); if (key === b) cls.push("re"); if (key > a && key < b) cls.splice(cls.indexOf("s"), 1); }
        else if (c.pair && ps && ps.y === c.y && ps.m === c.m && ps.d === d) cls.push("s");
        else if (sel.d === d && sel.m === c.m && (k === "md" || sel.y === c.y)) cls.push("s");
        if (d === c.fd) cls.push("k");
        g += btn(cls.join(" "), d, 'data-d="' + d + '"');
      }
      if (k !== "md") { const tot = first + n; for (let i = 1; tot + i - 1 < Math.ceil(tot / 7) * 7; i++) g += btn("o", i, 'data-pm="1" data-d="' + i + '"'); }
      body = (k === "md" ? "" : `<div class="dp-wd">${[...WD].map((w) => `<span>${w}</span>`).join("")}</div>`) + `<div class="dp-g">${g}</div>`;
    } else if (c.view === "m") {
      head = k === "md" ? "월 고르기" : `${c.y}년`;
      body = `<div class="dp-m">${Array.from({ length: 12 }, (_, i) => i + 1).map((mm) => btn(((k === "month" && sel.y === c.y && sel.m === mm) || (k !== "month" && mm === c.m) ? "s" : k !== "md" && c.y === t.y && mm === t.m ? "t" : "") + (mm === c.fm ? " k" : ""), mm + "월", 'data-m="' + mm + '"')).join("")}</div>`;
    } else {
      const y0 = c.y - ((c.y - 1900) % 12);
      head = `${y0} ~ ${y0 + 11}`;
      body = `<div class="dp-m">${Array.from({ length: 12 }, (_, i) => y0 + i).map((yy) => btn((sel.y === yy && k === "year" ? "s" : yy === t.y ? "t" : "") + (yy === c.fy ? " k" : ""), yy, 'data-y="' + yy + '"')).join("")}</div>`;
    }
    const hint = c.pair ? `<span class="dp-hint">${c.step === "e" ? "끝 날짜를 누르세요" : "시작 날짜를 누르세요"}</span>` : "";
    const todayBtn = k === "date" || k === "month" || k === "year" ? btn("", "오늘", "data-today") : "";
    pop.innerHTML = `<div class="dp-h"><button type="button" class="dp-t" data-head>${head}</button><span class="dp-nv">${btn("", "‹", 'data-nav="-1" aria-label="이전"')}${btn("", "›", 'data-nav="1" aria-label="다음"')}</span></div>${body}<div class="dp-f">${hint}${todayBtn}${btn("", "지우기", "data-clear")}${btn("dp-ok", "닫기", "data-close")}</div>`;
  }
  function pick(p) {
    const c = cur;
    if (c.pair && c.k === "date") {
      const key = (q) => (q ? q.y * 10000 + q.m * 100 + q.d : 0), ps = parse(c.pair.s);
      if (c.step === "s" || !ps || key(p) < key(ps)) { setVal(c.pair.s, store(c.pair.s, p)); const pe = parse(c.pair.e); if (pe && key(pe) < key(p)) setVal(c.pair.e, ""); if (c.pair.e && c.pair.e._dpFace && c.pair.e._dpFace.offsetParent !== null) { c.step = "e"; render(); return; } close(); return; }
      setVal(c.pair.e, store(c.pair.e, p)); close(); return;
    }
    if (c.pair && c.k === "month") { // 연월 기간: 시작 → 끝
      const key = (q) => (q ? q.y * 100 + q.m : 0), ps = parse(c.pair.s);
      if (c.step === "s" || !ps || key(p) < key(ps)) { setVal(c.pair.s, store(c.pair.s, p)); const pe = parse(c.pair.e); if (pe && key(pe) < key(p)) setVal(c.pair.e, ""); c.step = "e"; c.el = c.pair.e; render(); return; }
      setVal(c.pair.e, store(c.pair.e, p)); close(); return;
    }
    setVal(c.el, store(c.el, p)); close();
  }
  document.addEventListener("click", (ev) => {
    const face = ev.target.closest && ev.target.closest(".dp-face");
    if (face && face._dpFor) { if (face.disabled) return; ev.preventDefault(); if (cur && cur.el === face._dpFor) { close(); return; } open(face._dpFor); return; }
    if (!pop) return;
    if (!pop.contains(ev.target)) { close(); return; }
    const b = ev.target.closest("button"); if (!b) return;
    const c = cur, t = today(), d = b.dataset;
    if (d.close != null) return close();
    if (d.clear != null) { if (c.pair) { setVal(c.pair.s, ""); if (c.pair.e) setVal(c.pair.e, ""); } else setVal(c.el, ""); return close(); }
    if (d.today != null) { if (c.k === "date") return pick({ y: t.y, m: t.m, d: t.d }); if (c.k === "month") return pick({ y: t.y, m: t.m }); return pick({ y: t.y }); }
    if (d.nav) { const n = +d.nav; if (c.view === "d") { c.m += n; if (c.m < 1) { c.m = 12; c.y--; } if (c.m > 12) { c.m = 1; c.y++; } } else if (c.view === "m") { if (c.k === "md") return; c.y += n; } else c.y += 12 * n; render(); place(); return; }
    if (d.head != null) { if (c.view === "d") c.view = "m"; else if (c.view === "m" && c.k !== "md") c.view = "y"; render(); place(); return; }
    if (d.y) { c.y = +d.y; if (c.k === "year") return pick({ y: c.y }); c.view = "m"; render(); place(); return; }
    if (d.m) { c.m = +d.m; if (c.k === "month") return pick({ y: c.y, m: c.m }); c.view = "d"; render(); place(); return; }
    if (d.d) { let y = c.y, m = c.m; if (d.pm) { m += +d.pm; if (m < 1) { m = 12; y--; } if (m > 12) { m = 1; y++; } } if (c.k === "md") return pick({ m, d: +d.d }); return pick({ y, m, d: +d.d }); }
  }, true);
  document.addEventListener("keydown", (e) => { // 칸에 커서가 있을 때 Enter·Space·↓ = 달력 열기, Delete·Backspace = 지우기
    const f = e.target && e.target.classList && e.target.classList.contains("dp-face") ? e.target : null;
    if (!f || !f._dpFor || f.disabled || pop) return;
    if (e.key === "Enter" || e.key === " " || e.key === "ArrowDown") { e.preventDefault(); open(f._dpFor); }
    else if (e.key === "Delete" || e.key === "Backspace") { e.preventDefault(); setVal(f._dpFor, ""); }
  });
  window.addEventListener("popstate", close);
  window.DatePick = { version: 1, attach, refresh, close, scan };
})();
