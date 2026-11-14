// 신한FC m 서비스워커: 설치(홈 화면 추가) 조건용. 화면 파일(html)은 네트워크 먼저, 같은 주소의 js·css·아이콘은 저장본 먼저(?v=N이 바뀌면 새 주소라 새로 받음).
// 구글 로그인·Drive·Calendar·Gemini·카카오 등 다른 주소 요청은 가로채지 않고, 고객 정보·기록 데이터는 저장하지 않음.
const CACHE = "shlife-m-v1";
self.addEventListener("install", () => self.skipWaiting());
self.addEventListener("activate", (e) => {
  e.waitUntil((async () => {
    for (const k of await caches.keys()) if (k !== CACHE) await caches.delete(k);
    await self.clients.claim();
  })());
});
self.addEventListener("fetch", (e) => {
  const req = e.request, url = new URL(req.url);
  if (req.method !== "GET" || url.origin !== self.location.origin) return; // 다른 주소(구글 등)·쓰기 요청은 그대로 통과
  if (!url.pathname.startsWith(new URL(self.registration.scope).pathname)) return; // 이 앱 폴더(m/) 밖은 통과
  const isDoc = req.mode === "navigate" || (req.headers.get("accept") || "").includes("text/html");
  if (isDoc) { // 화면 파일: 네트워크 먼저 (새 버전이 늦게 뜨지 않게), 오프라인일 때만 저장본
    e.respondWith((async () => {
      try {
        const res = await fetch(req);
        if (res && res.ok) { const c = await caches.open(CACHE); c.put(new Request(url.origin + url.pathname), res.clone()); } // 쿼리(말만해요 embed·시각)를 뺀 주소로 저장
        return res;
      } catch (err) {
        const hit = await caches.match(new Request(url.origin + url.pathname));
        if (hit) return hit;
        throw err;
      }
    })());
    return;
  }
  // js·css·아이콘·manifest: 저장본 먼저, 없으면 받아서 저장 (주소에 ?v=N이 있어 버전이 바뀌면 자동으로 새로 받음)
  e.respondWith((async () => {
    const hit = await caches.match(req);
    if (hit) return hit;
    const res = await fetch(req);
    if (res && res.ok && res.type === "basic") { const c = await caches.open(CACHE); c.put(req, res.clone()); }
    return res;
  })());
});
