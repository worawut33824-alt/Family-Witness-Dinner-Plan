// Wedding Planner – Service Worker
// กลยุทธ์: NETWORK-FIRST — ดึงของใหม่จากเน็ตก่อนเสมอ แล้วค่อย fallback cache ตอนออฟไลน์
// => อัปเดตหน้าเว็บ/รูป ขึ้นเองทุกครั้ง ไม่ต้องลบแคช/บั๊มเวอร์ชันทุกครั้ง
// (ของเดิมเป็น cache-first จึงค้างรูป/หน้าเก่า)

const CACHE = 'wedding-planner-v90';
const CORE = [
  './',
  './index.html',
  './invite.html',
  './logo.png',
  './icon-192.png',
  './icon-512.png',
  './apple-touch-icon.png'
];

// ── Install: เก็บ core ไว้ให้เปิดออฟไลน์ได้ แล้วข้ามไปทำงานทันที ──
self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE).then(c => c.addAll(CORE)).catch(() => {})
  );
  self.skipWaiting();
});

// ── Activate: ลบ cache เก่าทุกเวอร์ชัน + เข้าควบคุมหน้าเปิดอยู่ทันที ──
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

// ── Fetch: NETWORK-FIRST ──
// เน็ตก่อนเสมอ → ได้ของล่าสุด; สำเนาลง cache ไว้เผื่อออฟไลน์; เน็ตล่ม → ใช้ cache
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;

  const url = req.url;
  // เรียก Google Apps Script / รูป Drive → เน็ตล้วน ไม่แคช
  if (url.includes('script.google.com') || url.includes('googleusercontent.com') || url.includes('drive.google.com')) {
    e.respondWith(fetch(req).catch(() => new Response('offline', { status: 503 })));
    return;
  }

  e.respondWith(
    fetch(req)
      .then(res => {
        // เก็บสำเนา response ที่ดีไว้ใช้ตอนออฟไลน์
        if (res && res.status === 200 && (res.type === 'basic' || res.type === 'cors')) {
          const clone = res.clone();
          caches.open(CACHE).then(c => c.put(req, clone)).catch(() => {});
        }
        return res;
      })
      .catch(() => caches.match(req).then(c => c || caches.match('./index.html')))
  );
});
