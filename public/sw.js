/* Service worker — bản tối giản, đủ để app cài được và mở nhanh.
 *
 * Nguyên tắc chọn chiến lược cache:
 *   - Trang HTML: LUÔN ưu tiên mạng. Nếu cache trước, người dùng sẽ mắc kẹt
 *     ở phiên bản cũ sau mỗi lần deploy — lỗi kinh điển của PWA.
 *   - Tệp có mã băm (/assets/index-a1b2c3.js): cache thoải mái, vì đổi nội
 *     dung là đổi tên tệp.
 *   - Gọi Supabase: KHÔNG cache. Dữ liệu công nợ, kho hàng phải luôn mới.
 *
 * Đổi CACHE_VERSION mỗi lần muốn xóa sạch cache cũ.
 */

const CACHE_VERSION = "v1";
const STATIC_CACHE = `ch54-static-${CACHE_VERSION}`;

const PRECACHE = [
  "/",
  "/manifest.webmanifest",
  "/icon-192.png",
  "/icon-512.png",
];

self.addEventListener("install", (e) => {
  e.waitUntil(
    caches.open(STATIC_CACHE)
      .then((c) => c.addAll(PRECACHE))
      .then(() => self.skipWaiting())
      .catch(() => self.skipWaiting())
  );
});

self.addEventListener("activate", (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(
        keys.filter((k) => k !== STATIC_CACHE).map((k) => caches.delete(k))
      ))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("message", (e) => {
  if (e.data === "SKIP_WAITING") self.skipWaiting();
});

self.addEventListener("fetch", (e) => {
  const req = e.request;
  if (req.method !== "GET") return;

  const url = new URL(req.url);

  // Dữ liệu Supabase và mọi tên miền khác: để mạng lo, không đụng vào
  if (url.origin !== self.location.origin) return;

  // Điều hướng trang: mạng trước, hỏng mạng thì lấy bản đã lưu
  if (req.mode === "navigate") {
    e.respondWith(
      fetch(req)
        .then((res) => {
          const copy = res.clone();
          caches.open(STATIC_CACHE).then((c) => c.put("/", copy));
          return res;
        })
        .catch(() => caches.match("/").then((r) => r || Response.error()))
    );
    return;
  }

  // Tệp tĩnh có mã băm: lấy cache trước cho nhanh, đồng thời làm mới nền
  if (url.pathname.startsWith("/assets/") ||
      /\.(png|svg|ico|woff2?|xlsx)$/.test(url.pathname)) {
    e.respondWith(
      caches.match(req).then((cached) => {
        const network = fetch(req)
          .then((res) => {
            if (res && res.status === 200) {
              const copy = res.clone();
              caches.open(STATIC_CACHE).then((c) => c.put(req, copy));
            }
            return res;
          })
          .catch(() => cached);
        return cached || network;
      })
    );
  }
});
