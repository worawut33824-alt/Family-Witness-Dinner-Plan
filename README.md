# 💍 Family Witness Dinner — แป้ง × เอ็ม

เว็บแอปงานสักขีพยาน (Wedding PWA) — **การ์ดเชิญออนไลน์ + ระบบจัดการงานหลังบ้าน**
📅 วันเสาร์ที่ 28 พฤศจิกายน 2569 · 16:00 น. · 📍 ธาราเทอเรส (TARA Terrace) นครปฐม

## 🔗 ลิงก์ใช้งาน (Live · GitHub Pages)
| หน้า | ลิงก์ | สำหรับ |
|------|------|--------|
| 💌 การ์ดเชิญ | [invite.html](https://worawut33824-alt.github.io/Family-Witness-Dinner-Plan/invite.html) | แขก |
| 📊 แดชบอร์ดจัดการงาน | [index.html](https://worawut33824-alt.github.io/Family-Witness-Dinner-Plan/index.html) | เจ้าภาพ |
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

### รอบ 20 ต.ค. 2026 — รีดีไซน์หลังบ้าน + เพิ่มความเร็ว/ความแม่นยำ
- 🎨 **รีดีไซน์แดชบอร์ด `index.html`** — เมนูเป็น **sidebar ซ้าย** (เดสก์ท็อป) / top bar + drawer (มือถือ), วางระบบ **design token** ใหม่ (สี/ระยะ/รัศมี/เงานุ่ม), ฟอนต์ **Sarabun + Cormorant Garamond**, รีสไตล์ card/ตาราง/ฟอร์ม/ปุ่ม/modal/countdown ให้โมเดิร์น สะอาดตา
- 📐 **จัด layout สมดุล + ยืดหยุ่นทุกจอ** — เนื้อหา center ในพื้นที่หลัง sidebar (เลิกเหลือที่ว่างขวาบนจอกว้าง)
- 📊 **กราฟสรุปแขกใหม่** — โดนัทแยกฝั่งเอ็ม/แป้ง + เลขที่นั่งรวมตรงกลาง (ดูง่าย) ข้างๆ pictogram เดิม
- 🐛 **แก้บั๊ก/ความแม่นยำ (หน้าเว็บ)** — เลิกขึ้น "ซิงค์ไม่ได้" หลอกทุก 60 วิ (ลบ `renderGuestStats` ที่ไม่มีจริง) · อ่านยอดเงินทน comma ("2,000" ไม่เป็น 0) · escape HTML ชื่อแขก/ผู้ติดต่อ (กัน XSS)
- ⚡ **ความเร็ว (หน้าเว็บ)** — auto-refresh วาดใหม่เฉพาะหน้าที่เปิด + เฉพาะเมื่อข้อมูลเปลี่ยนจริง · bump SW cache v92
- ⚙️ **Backend `code.gs`** — ใส่ **LockService** ครอบทุกการเขียน (กันข้อมูลชนกัน/หายตอนส่ง RSVP/สลิปพร้อมกัน) · `saveDonation` strip comma · `getChecklist` อ่านคอลัมน์ทีเดียว (เลิก N+1)
- 👥 **แก้ schema แขกชนกัน** — ตัดช่อง "อาหาร" (ไม่ใช้แล้ว), `getGuests`/`saveGuests` ใช้ 11 คอลัมน์ให้ตรงกับ RSVP — ไม่ทับ/ล้าง อีเมล/คำอวยพร/รูป/วันที่ อีกต่อไป (มีผลเต็มเมื่อ deploy `code.gs`)

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
