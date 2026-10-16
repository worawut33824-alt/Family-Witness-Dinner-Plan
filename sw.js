// Wedding Planner – Service Worker
// กลยุทธ์ผสม เพื่อให้ "สดเสมอ" + "เปิดไว":
//  • HTML (หน้าเว็บ)        → NETWORK-FIRST : ได้หน้าล่าสุดเสมอ (อัปเดตขึ้นเอง ไม่ค้าง)
//  • รูป/เพลง/ฟอนต์/js/css → STALE-WHILE-REVALIDATE : เสิร์ฟจาก cache ทันที (ไว)
//                              แล้วอัปเดต cache เบื้องหลังไว้ใช้รอบหน้า
//  • Google Apps Script/Drive → เน็ตล้วน ไม่แคช
// รูปที่เปลี่ยนบ่อย (cover/countdown) ใส่ ?v=N ต่อท้าย → เปลี่ยนเลข = URL ใหม่ = โหลดสดทันที

const CACHE = 'wedding-planner-v91';
const CORE = [
  './',
  './index.html',
  './invite.html',
  './logo.png',
  './icon-192.png',
  './icon-512.png',
  './apple-touch-icon.png'
];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(CORE)).catch(() => {}));
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

function isHTML(req) {
  if (req.mode === 'navigate') return true;
  const u = new URL(req.url);
  return u.pathname.endsWith('.html') || u.pathname.endsWith('/');
}

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = req.url;

  // backend/Drive — เน็ตล้วน
  if (url.includes('script.google.com') || url.includes('googleusercontent.com') || url.includes('drive.google.com')) {
    e.respondWith(fetch(req).catch(() => new Response('offline', { status: 503 })));
    return;
  }

  // HTML — network-first (สดเสมอ), ออฟไลน์ค่อย fallback cache
  if (isHTML(req)) {
    e.respondWith(
      fetch(req)
        .then(res => {
          if (res && res.status === 200) {
            const clone = res.clone();
            caches.open(CACHE).then(c => c.put(req, clone)).catch(() => {});
          }
          return res;
        })
        .catch(() => caches.match(req).then(c => c || caches.match('./index.html')))
    );
    return;
  }

  // asset อื่น (รูป/เพลง/ฟอนต์/js/css) — stale-while-revalidate (เสิร์ฟ cache ทันที + อัปเดตเบื้องหลัง)
  e.respondWith(
    caches.match(req).then(cached => {
      const network = fetch(req).then(res => {
        if (res && res.status === 200 && (res.type === 'basic' || res.type === 'cors')) {
          const clone = res.clone();
          caches.open(CACHE).then(c => c.put(req, clone)).catch(() => {});
        }
        return res;
      }).catch(() => cached);
      return cached || network;   // มี cache = คืนทันที (ไว) ไม่มี = รอเน็ต
    })
  );
});
