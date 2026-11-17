# 💍 Family Witness Dinner — แป้ง × เอ็ม

เว็บแอปงานสักขีพยาน (Wedding PWA) — **การ์ดเชิญออนไลน์ + ระบบจัดการงานหลังบ้าน**
📅 วันเสาร์ที่ 28 พฤศจิกายน 2569 · 16:00 น. · 📍 ธาราเทอเรส (TARA Terrace) นครปฐม

## 🔗 ลิงก์ใช้งาน (Live · GitHub Pages)
| หน้า | ลิงก์ | สำหรับ |
|------|------|--------|
| 💌 การ์ดเชิญ | [invite.html](https://worawut33824-alt.github.io/Family-Witness-Dinner-Plan/invite.html) | แขก |
| 📊 แดชบอร์ดจัดการงาน | [index.html](https://worawut33824-alt.github.io/Family-Witness-Dinner-Plan/index.html) | เจ้าภาพ |
| 🖥️ จอแสดงคำอวยพร (วันงาน) | [wishes.html](https://worawut33824-alt.github.io/Family-Witness-Dinner-Plan/wishes.html) | เปิดบนจอ/โปรเจคเตอร์ |
| ✍️ ฟอร์มส่งคำอวยพร (QR) | [bless.html](https://worawut33824-alt.github.io/Family-Witness-Dinner-Plan/bless.html) | แขกสแกน QR ในงาน |
| 🎴 จัดรูปพื้นหลังการ์ด | [bg-editor.html](https://worawut33824-alt.github.io/Family-Witness-Dinner-Plan/bg-editor.html) | เจ้าภาพ |
| ✨ หน้าโปรโมท/ฟีเจอร์ | [features.html](https://worawut33824-alt.github.io/Family-Witness-Dinner-Plan/features.html) | ทั่วไป |

## 🏗️ สถาปัตยกรรม
- **Frontend** — ไฟล์ HTML เดี่ยว (PWA ภาษาไทย) โฮสต์บน **GitHub Pages** (deploy อัตโนมัติจาก branch `master`)
- **Backend** — **Google Apps Script** Web App (`code.gs`) + ฐานข้อมูลเป็น **Google Sheets**
- **บริการที่เชื่อม** — Gemini Vision (OCR อ่านยอดจากสลิป), Google Calendar + Gmail (ส่ง invite/.ics), Google Drive (เก็บรูป/สลิป)

## 📁 โครงสร้างไฟล์หลัก
| ไฟล์ | หน้าที่ |
|------|--------|
| `invite.html` | การ์ดเชิญ (cover, story, couple, กำหนดการ, แผนที่, RSVP, คำอวยพร, donate ฯลฯ) |
| `index.html` | แดชบอร์ดหลังบ้าน — 11 แท็บ (Dashboard, ค่าใช้จ่าย, งบประมาณ, การสนับสนุน, Checklist, ผู้ติดต่อ, แขก, คำอวยพร, ลิงก์เชิญ, ตั้งค่าการ์ด, ซิงค์) |
| `wishes.html` | จอแสดงคำอวยพรแบบสไลด์โชว์ (เปิดบนจอ/โปรเจคเตอร์วันงาน · ดึงคำอวยพรสดจาก backend · 2 โหมด/2 ธีม/รูปคู่) |
| `bless.html` | หน้าฟอร์มส่งคำอวยพรสำหรับสแกน QR ในงาน (เปิดกล้องถ่ายรูป + ชื่อ/ฝั่ง/คำอวยพร → POST `action=bless` → ขึ้นจอ) |
| `bg-editor.html` | เครื่องมือ WYSIWYG จัดตำแหน่ง+ซูมรูปพื้นหลัง/รูปคู่บ่าวสาว |
| `code.gs` | Backend Apps Script (อ่าน/เขียนชีต, OCR, ปฏิทิน, อีเมล) |
| `apps-script-invite-addon.js` | ไฟล์สำรองสร้าง .ics/อีเมลเชิญ (ไม่ได้ deploy คู่กับ code.gs) |
| `sw.js` · `manifest.json` | Service Worker + PWA manifest |
| `scripts/deploy-gs.mjs` · `package.json` | ตั้ง clasp สำหรับ deploy `code.gs` อัตโนมัติ |
| `.github/workflows/keepwarm.yml` | cron อุ่น backend (ลด cold start) |

## 🚀 การ Deploy
**Frontend** (invite.html, index.html, bg-editor.html, sw.js, รูป): `git push` ขึ้น `master` → GitHub Pages rebuild เองภายใน ~1 นาที
> เปลี่ยนไฟล์ HTML/asset แล้วอย่าลืม bump เลข cache ใน `sw.js` (`const CACHE = 'wedding-planner-vNN'`)

**Backend** (`code.gs`): ต้อง deploy เข้า Apps Script เอง
```bash
git pull
npm install        # ครั้งแรกครั้งเดียว
npm run deploy:gs  # ใช้ clasp (ต้อง clasp login + มี .clasp.json บนเครื่อง)
```
หรือเปิด Apps Script editor → วางเนื้อหา `code.gs` ใหม่ทับ → Deploy

## 📝 ความคืบหน้าล่าสุด (Changelog)

### รอบ 17 พ.ย. 2026 (เพิ่ม) — ฟอร์มส่งคำอวยพรสำหรับ QR (กล้องในหน้าเว็บ) + หน้ารวมลิงค์ ✍️📷
- **หน้าใหม่ `bless.html`** — หน้าฟอร์มส่งคำอวยพรเฉพาะสำหรับสแกน QR ในงาน (QR บน `wishes.html` ชี้มาที่นี่ แทน deep-link เข้าการ์ด)
  - 📷 **เปิดกล้องในหน้าเว็บ** (เซลฟี่ + สลับกล้องหน้า/หลัง + ถ่าย) · สำรองเลือกรูปจากเครื่อง · ย่อรูป ≤800px
  - ฟิลด์เหมือนฟอร์มในการ์ดเชิญ: ชื่อ · ฝั่ง (เอ็ม/แป้ง) · คำอวยพร · รูป → POST `action=bless` (no-cors) → ขึ้นจอแสดงผล
  - มือถือเป็นหลัก · ธีมฟ้า-ขาว + โลโก้ · มีหน้า "ส่งสำเร็จ / ส่งอีกคน"
- 🔗 **ย้ายลิงก์ฮับ "เปิดหน้าต่างๆ ของงาน" ไปหน้า "รวมลิงค์"** (เปลี่ยนชื่อเมนูจาก "ลิงก์เชิญ") + เพิ่มปุ่ม bless.html
- bump SW cache **v104** + precache `bless.html`

### รอบ 17 พ.ย. 2026 (เพิ่ม) — จอแสดงคำอวยพรวันงาน (Live Wishes Display) 🖥️
- **หน้าใหม่ `wishes.html`** — สไลด์โชว์คำอวยพรสำหรับเปิดบนจอ/โปรเจคเตอร์วันงาน (แยกจากแดชบอร์ด)
- 🎞️ 2 โหมด: **การ์ดเดี่ยว** (ทีละคน เลื่อนเอง) · **ผนังลอย** (หลายใบพร้อมกัน) — สลับได้หน้างาน
- 🌗 2 ธีม: **ฟ้า-ขาว / กลางคืน** + **เปิด/ปิดรูปคู่** (โลโก้จริง + รูปคู่เบลอเป็นพื้นหลัง · การ์ดฝ้าให้อ่านชัด)
- 🔄 ดึงคำอวยพรสดจาก `?action=blessings` + **รีเฟรชอัตโนมัติทุก 60 วิ** (คนส่งใหม่เด้งขึ้นจอ) · แคชออฟไลน์
- 💗 สีแยกฝั่งแป้ง(ชมพู)/เอ็ม(ฟ้า) · **QR มุมจอ** → เปิดฟอร์มส่งคำอวยพรอัตโนมัติ (`invite.html#bless` deep-link เปิดซอง+ฟอร์มให้เลย) · เต็มจอ + คุมด้วยคีย์บอร์ด/รีโมท · wake-lock กันจอดับ
- ➕ ปุ่ม **"เปิดจอแสดงคำอวยพร"** ในแท็บกล่องคำอวยพร + **ลิงก์ฮับ "เปิดหน้าต่างๆ ของงาน"** บนหน้า Dashboard (การ์ดเชิญ · จอคำอวยพร · จัดรูปพื้นหลัง · หน้าฟีเจอร์) · ขยาย QR บนจอให้สแกนง่ายขึ้น · bump SW cache **v101**

### รอบ 17 พ.ย. 2026 — ยกเครื่องดีไซน์แดชบอร์ด (Soft UI) 🎨
- **รีดีไซน์ `index.html` ครบทุกแท็บ** (CSS-forward · คงฟังก์ชัน/page id/handler/localStorage เดิมทั้งหมด)
- 🧭 **Sidebar ฟ้าเข้มหรูหรา** + เมนูเป็น **ไอคอน SVG** (เลิกใช้อีโมจิในเมนู) + active ไฮไลต์เส้นทอง
- 🪟 **สไตล์ Soft UI** — พื้นหลังนิวทรัลเย็น, เงานุ่มมีมิติ, มุมโค้งใหญ่ขึ้น, การ์ด/ตาราง/ฟอร์ม/ปุ่มสะอาดตา
- 📊 **KPI/stat cards** แถบไล่สีด้านบน + hover ยก · countdown hero เปลี่ยน watermark เป็น soft glow
- 🩷💙 **สีแยกฝั่ง: แป้ง = ชมพู · เอ็ม = ฟ้า** (กราฟโดนัท/pictogram/badge แขก)
- ♿ a11y: `focus-visible` + รองรับ `prefers-reduced-motion` · bump SW cache **v98**
- *(ใช้แนวทางจากสกิล ui-ux-pro-max — สไตล์ Soft UI Evolution, คงโทนแบรนด์ฟ้า-เซจ-ทอง)*

### รอบ 13 พ.ย. 2026 — ระบบอีเมลเตือนแขก (Guest Reminder) 📨
- **ระบบส่งอีเมลเตือนอัตโนมัติ** (`code.gs`) — ส่งถึงแขกที่ RSVP "มาร่วมงาน" + มีอีเมล (26 คน) และคู่บ่าวสาว (เอ็ม/แป้ง) รวม 28 คน ผ่าน Gmail ของเจ้าของสคริปต์ ชื่อผู้ส่ง "แป้ง & เอ็ม"
  - `sendGuestReminders()` ส่งจริง — personalize ชื่อ (`เรียน คุณ...`, ตัด "คุณ" ซ้ำอัตโนมัติ) + "อีก N วัน" คำนวณสดตอนส่ง + **กันส่งซ้ำ** (sent log ใน Script Properties)
  - `previewGuestReminders()` ดูรายชื่อผู้รับก่อนส่ง · `sendReminderTestToMe()` ส่งทดสอบเข้าเมลตัวเอง · `installReminderTrigger()` / `removeReminderTrigger()` ตั้ง/ยกเลิก time-based trigger
  - ⏰ **ตั้งเวลายิงอัตโนมัติ 2 รอบ: จันทร์ 23 + ศุกร์ 27 พฤศจิกายน 2569 18:00 น. (เวลาไทย)** — ติดตั้ง trigger แล้ว (deploy ผ่าน Apps Script editor)
  - 🔁 แต่ละรอบส่งใหม่ครบทุกคน (กันส่งซ้ำแยกตามวัน `reminder_sent_YYYY-MM-DD`) · "อีก N วัน" คำนวณสดตอนส่ง (23 พ.ย.→5 วัน, 27 พ.ย.→1 วัน) · คนที่ลงทะเบียนเพิ่มก่อนเวลาส่งจะได้รับอัตโนมัติ
  - 🔣 ใช้ HTML entity แทนอีโมจิในเนื้ออีเมล (กันอักขระเพี้ยนตอนส่งผ่าน Gmail)
- 🎨 **ดีไซน์อีเมล** — โทนฟ้า-ขาว (เข้าชุดโมโนแกรมทางการ P M) + โลโก้ `logo.png` ด้านบนสุด + รูปคู่บ่าวสาว `รูปประกอบการ์ดเชิญ/couple-reminder.jpg` + การ์ดวัน/เวลา/สถานที่ + ปุ่มดูการ์ดเชิญ/แผนที่ · ออกแบบ email-safe (table + inline style) เปิดได้ทุกเมลไคลเอนต์
- 🖼️ เพิ่มรูป `couple-reminder.jpg` (1000×1500) โฮสต์บน GitHub Pages สำหรับใช้ในอีเมล

### รอบ 20 ต.ค. 2026 — รีดีไซน์หลังบ้าน + เพิ่มความเร็ว/ความแม่นยำ
- 🎨 **รีดีไซน์แดชบอร์ด `index.html`** — เมนูเป็น **sidebar ซ้าย** (เดสก์ท็อป) / top bar + drawer (มือถือ), วางระบบ **design token** ใหม่ (สี/ระยะ/รัศมี/เงานุ่ม), ฟอนต์ **Sarabun + Cormorant Garamond**, รีสไตล์ card/ตาราง/ฟอร์ม/ปุ่ม/modal/countdown ให้โมเดิร์น สะอาดตา
- 📐 **จัด layout สมดุล + ยืดหยุ่นทุกจอ** — เนื้อหา center ในพื้นที่หลัง sidebar (เลิกเหลือที่ว่างขวาบนจอกว้าง)
- 📊 **กราฟสรุปแขกใหม่** — โดนัทแยกฝั่งเอ็ม/แป้ง + เลขที่นั่งรวมตรงกลาง (ดูง่าย) ข้างๆ pictogram เดิม
- 🐛 **แก้บั๊ก/ความแม่นยำ (หน้าเว็บ)** — เลิกขึ้น "ซิงค์ไม่ได้" หลอกทุก 60 วิ (ลบ `renderGuestStats` ที่ไม่มีจริง) · อ่านยอดเงินทน comma ("2,000" ไม่เป็น 0) · escape HTML ชื่อแขก/ผู้ติดต่อ (กัน XSS)
- ⚡ **ความเร็ว (หน้าเว็บ)** — auto-refresh วาดใหม่เฉพาะหน้าที่เปิด + เฉพาะเมื่อข้อมูลเปลี่ยนจริง · bump SW cache v92
- ⚙️ **Backend `code.gs`** — ใส่ **LockService** ครอบทุกการเขียน (กันข้อมูลชนกัน/หายตอนส่ง RSVP/สลิปพร้อมกัน) · `saveDonation` strip comma · `getChecklist` อ่านคอลัมน์ทีเดียว (เลิก N+1)
- 👥 **แก้ schema แขกชนกัน** — ตัดช่อง "อาหาร" (ไม่ใช้แล้ว), `getGuests`/`saveGuests` ใช้ 11 คอลัมน์ให้ตรงกับ RSVP — ไม่ทับ/ล้าง อีเมล/คำอวยพร/รูป/วันที่ อีกต่อไป (มีผลเต็มเมื่อ deploy `code.gs`)
- 💌 **ปรับปรุงการ์ดเชิญ** — lazy-load รูป (Love Tree ~7.8MB) + preload cover + bgm `preload=none` + ฟอนต์ (โหลดไวขึ้น), `prefers-reduced-motion`, ปลดล็อกซูม, ปุ่มตอบรับกดด้วยคีย์บอร์ดได้ + label ผูก input (`for`/`id`), scroll bar ใช้ rAF
- 🎴 **รีดีไซน์การ์ด "มาร่วมงาน / ไม่มาร่วมงาน"** ให้เข้าธีม (ขาวครีม + เส้นขอบทอง + sage/ฟ้าหม่น)
- 🛠️ **เพิ่มสคริปต์ย่อรูป** `optimize-images.mjs` (`npm run optimize:img`, ต้อง `npm i sharp`) — ย่อรูปที่ใหญ่เกินจริง, สำรองต้นฉบับใน `images-backup/` *(ยังไม่รัน — รันบนเครื่องที่มี Node ได้เลย)*
- ✨ **ขัดเกลาการ์ดเพิ่ม** — ต้นไม้ผู้ร่วมงานเล่นอนิเมชันโดยไม่ rebuild DOM (เบาขึ้น) + จัดระยะ section เป็น token 3 ชั้น (`--sp-sm/md/lg`) ให้จังหวะสม่ำเสมอ

### ก่อนหน้า (16–18 ต.ค. 2026)
- เพิ่ม `bg-editor.html` (จัดรูปพื้นหลัง + ซูม + เปลี่ยนรูปคู่บ่าวสาว), เปลี่ยนรูป cover/countdown, ปรับ perf (เพลงไม่โหลดอัตโนมัติ, SW network-first, keep-warm cron)

### ก่อนหน้า (ก.ย.–ต.ค. 2026)
- ระบบเปิด/ปิด section การ์ด, เปิด donate กลับมา, OCR อ่านสลิปด้วย Gemini, ตั้ง clasp auto-deploy, ปรับกำหนดการ (ลบพิธีสงฆ์, เวลาเริ่ม 16:00, สลับไทม์ไลน์เริ่มจากซ้าย)

## ⚠️ สิ่งที่ค้าง / Known Issues
1. ✅ **Schema แขกชนกัน — แก้แล้ว (20 ต.ค. 2026)** — ตัดช่อง "อาหาร" (ไม่ใช้แล้ว), `getGuests`/`saveGuests` ใช้ 11 คอลัมน์ ไม่ทับ/ล้างข้อมูล RSVP อีกต่อไป · **ต้อง deploy `code.gs` เพื่อให้ฝั่ง backend มีผลเต็ม**
2. **ปรับปรุง backend เพิ่มเติม (แนะนำทำตอน login clasp + ทดสอบได้)** — CacheService ลดการอ่านชีตซ้ำ, กันนับ RSVP ซ้ำ (idempotency), ย้าย action ที่เขียน/ลบชีตจาก GET → POST (ความปลอดภัย), ทบทวนรายชื่อ Gemini model
3. **Multi-device** — การ sync ยังเป็นแบบ push ทั้งก้อน (last-writer-wins) หลายเครื่องแก้พร้อมกันอาจทับกัน — LockService ช่วยกันระดับ request แล้ว แต่การ merge ราย row ยังเป็นงานเฟสถัดไป

## 🔧 หมายเหตุเทคนิค
- Push ขึ้น repo ใช้บัญชี GitHub **worawut33824-alt** (เจ้าของ repo) · Pages deploy จาก `master`
- ไฟล์ที่ไม่ push: `New folder/`, `tryout/`, `node_modules/`, `.clasp*.json` (ดู `.gitignore`)
- Apps Script: ตั้ง `GEMINI_API_KEY` ใน Script Properties (เมนู "🤖 Gemini") สำหรับ OCR อ่านสลิป

---
🤖 เอกสารนี้สร้าง/อัปเดตด้วย [Claude Code](https://claude.com/claude-code)
