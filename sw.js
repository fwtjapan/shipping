// FWT JAPAN 雲端集運 PWA Service Worker
//
// 策略：
//  - /static/ 與圖示：cache-first（快、節省流量）
//  - 一切 API（/api/*）：network-only，永不快取（資料即時性）
//  - 頁面 (HTML / 導覽請求)：network-first，網路爆掉時 fallback 到 offline 頁
//  - SW 換版時自動 skipWaiting + claim → 下次重新整理就用新版

// 改動 static/ 底下的檔案時，記得同步更新這個 VERSION，否則舊快取不會失效。
const VERSION = "2026-09-15.1";
const STATIC_CACHE = `fwt-static-${VERSION}`;
const OFFLINE_URL = "/static/offline.html";

const PRECACHE = [
  "/static/offline.html",
  "/static/icons/icon-192.png",
  "/static/icons/icon-512.png",
];

// === Install：預先快取必備 static + offline page ===
self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(STATIC_CACHE).then((cache) => cache.addAll(PRECACHE))
  );
  self.skipWaiting();
});

// === Activate：清掉舊版本快取 ===
self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== STATIC_CACHE).map((k) => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

// === Fetch：依路徑採用不同策略 ===
self.addEventListener("fetch", (event) => {
  const req = event.request;

  // 只處理 GET（POST/PUT/DELETE 都直接通過）
  if (req.method !== "GET") return;

  const url = new URL(req.url);

  // 略過 chrome-extension:// 等非 http 請求
  if (!url.protocol.startsWith("http")) return;

  // 略過跨網域請求（Shopify CDN、Google fonts 等）
  if (url.origin !== self.location.origin) return;

  // ---- API：永不快取，網路一定要新鮮 ----
  if (url.pathname.startsWith("/api/")) {
    return; // 不攔截 → 走預設網路請求
  }

  // ---- /static/：cache-first ----
  if (url.pathname.startsWith("/static/")) {
    event.respondWith(
      caches.match(req).then((cached) => {
        if (cached) return cached;
        return fetch(req).then((resp) => {
          if (resp.ok) {
            const clone = resp.clone();
            caches.open(STATIC_CACHE).then((c) => c.put(req, clone));
          }
          return resp;
        });
      })
    );
    return;
  }

  // ---- 導覽請求 / HTML：network-first，網路掛了用 offline ----
  if (req.mode === "navigate" || (req.headers.get("accept") || "").includes("text/html")) {
    event.respondWith(
      fetch(req).catch(() => caches.match(OFFLINE_URL))
    );
    return;
  }

  // 其他資源：嘗試網路，失敗看快取
  event.respondWith(
    fetch(req).catch(() => caches.match(req))
  );
});

// === 收到 client 訊息：清快取（debug 用）===
self.addEventListener("message", (event) => {
  if (event.data === "skipWaiting") self.skipWaiting();
  if (event.data === "clearCache") {
    caches.keys().then((keys) => Promise.all(keys.map((k) => caches.delete(k))));
  }
});
