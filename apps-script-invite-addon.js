// ═══════════════════════════════════════════════════════════════════
//  PM WEDDING INVITE — Apps Script Add-on
//  เพิ่ม code นี้ต่อจาก function doPost(e) / doGet(e) ที่มีอยู่แล้ว
//  หรือ merge เข้าไปใน doPost / doGet เดิม
// ═══════════════════════════════════════════════════════════════════

// ── เพิ่มใน doPost(e) ──────────────────────────────────────────────
// รับ POST จาก invite.html (action: rsvp | bless)
/*
  ใน doPost(e) เดิม เพิ่ม:

  if (action === 'rsvp')   return handleRSVP(params);
  if (action === 'bless')  return handleBlessing(params);
*/

// ── เพิ่มใน doGet(e) ───────────────────────────────────────────────
// รับ GET จาก invite.html (action: stats)
/*
  ใน doGet(e) เดิม เพิ่ม:

  if (action === 'stats')  return getInviteStats();
*/

// ═══════════════════════════════════════════════════════════════════
//  บันทึก RSVP → sheet "RSVP"
// ═══════════════════════════════════════════════════════════════════
function handleRSVP(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // สร้าง sheet RSVP ถ้ายังไม่มี
  let sheet = ss.getSheetByName('RSVP');
  if (!sheet) {
    sheet = ss.insertSheet('RSVP');
    sheet.appendRow([
      'วันที่/เวลา', 'ชื่อ-นามสกุล', 'จำนวน',
      'เบอร์โทร', 'Email', 'ฝั่ง', 'ความสัมพันธ์', 'หมายเหตุ'
    ]);
    sheet.setFrozenRows(1);
  }

  sheet.appendRow([
    new Date().toLocaleString('th-TH', { timeZone: 'Asia/Bangkok' }),
    params.name    || '',
    params.count   || 1,
    params.phone   || '',
    params.email   || '',
    params.side    || '',
    params.relation|| '',
    params.note    || ''
  ]);

  // ส่ง Calendar Invite ถ้ามี email
  if (params.email && params.email.includes('@')) {
    try { sendCalendarInvite(params.email, params.name); } catch(e) {}
  }

  // คืนค่า guest count ล่าสุด
  const totalCount = getTotalGuestCount();

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, guestCount: totalCount }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ═══════════════════════════════════════════════════════════════════
//  บันทึกคำอวยพร → sheet "Blessings"
// ═══════════════════════════════════════════════════════════════════
function handleBlessing(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName('Blessings');
  if (!sheet) {
    sheet = ss.insertSheet('Blessings');
    sheet.appendRow(['วันที่/เวลา', 'ชื่อ', 'คำอวยพร']);
    sheet.setFrozenRows(1);
  }

  sheet.appendRow([
    new Date().toLocaleString('th-TH', { timeZone: 'Asia/Bangkok' }),
    params.name    || 'ไม่ระบุชื่อ',
    params.message || ''
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ═══════════════════════════════════════════════════════════════════
//  ดึงสถิติ: จำนวนแขก + คำอวยพรล่าสุด
// ═══════════════════════════════════════════════════════════════════
function getInviteStats() {
  const guestCount = getTotalGuestCount();
  const blessings  = getLatestBlessings(10);

  return ContentService
    .createTextOutput(JSON.stringify({ guestCount, blessings }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getTotalGuestCount() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('RSVP');
  if (!sheet || sheet.getLastRow() <= 1) return 0;

  const data = sheet.getRange(2, 3, sheet.getLastRow() - 1, 1).getValues();
  return data.reduce((sum, row) => sum + (parseInt(row[0]) || 0), 0);
}

function getLatestBlessings(n) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Blessings');
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const last  = sheet.getLastRow();
  const start = Math.max(2, last - n + 1);
  const data  = sheet.getRange(start, 1, last - start + 1, 3).getValues();

  return data.reverse().map(row => ({
    date:    row[0] ? row[0].toString() : '',
    name:    row[1] || '',
    message: row[2] || ''
  }));
}

// ═══════════════════════════════════════════════════════════════════
//  ส่ง Calendar Invite (.ics) ทาง Gmail
// ═══════════════════════════════════════════════════════════════════
function sendCalendarInvite(email, guestName) {
  // ── ปรับเวลาเริ่ม/จบตามที่ต้องการ ──
  // DTSTART = 2026-11-28 14:00 (Bangkok UTC+7 = 07:00 UTC)
  // DTEND   = 2026-11-28 20:00 (Bangkok UTC+7 = 13:00 UTC)
  const ics = [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    'PRODID:-//PM Wedding//TH',
    'METHOD:REQUEST',
    'BEGIN:VEVENT',
    'UID:pmwedding-20261128@tara-terrace',
    'DTSTART:20261128T070000Z',   // 14:00 ICT = 07:00 UTC
    'DTEND:20261128T130000Z',     // 20:00 ICT = 13:00 UTC
    'SUMMARY:💍 งานสักขีพยาน Family Witness Dinner – เอ็ม & แป้ง',
    'LOCATION:TARA Terrace\\, Nakhon Pathom',
    'DESCRIPTION:ยินดีต้อนรับสู่งานสักขีพยาน Family Witness Dinner\\n'
      + 'ลงทะเบียน 13:30 น. · เริ่ม 14:00 น.\\n'
      + 'TARA Terrace · จ.นครปฐม',
    'STATUS:CONFIRMED',
    'SEQUENCE:0',
    'END:VEVENT',
    'END:VCALENDAR'
  ].join('\r\n');

  const htmlBody = `
    <div style="font-family:sans-serif;max-width:520px;margin:0 auto;background:#FFF8F3;border-radius:12px;overflow:hidden;border:1px solid #F0DEB8;">
      <div style="background:linear-gradient(135deg,#3d6a91,#7B9CBF);padding:28px;text-align:center;color:white;">
        <div style="font-size:2rem;margin-bottom:8px;">💍</div>
        <div style="font-family:Georgia,serif;font-size:1.4rem;font-style:italic;">เอ็ม &amp; แป้ง</div>
        <div style="font-size:0.85rem;opacity:0.85;margin-top:4px;letter-spacing:2px;">FAMILY WITNESS DINNER</div>
      </div>
      <div style="padding:28px 24px;text-align:center;">
        <p style="font-size:1rem;color:#2c3a4a;">สวัสดีคุณ <strong>${guestName}</strong>,</p>
        <p style="color:#5a7080;font-size:0.9rem;line-height:1.7;">
          ขอบคุณมากที่ยืนยันร่วมงานสักขีพยาน<br>
          เราตั้งตารอพบคุณในวันงาน 🥂
        </p>
        <div style="background:#EEF5FC;border-radius:10px;padding:16px;margin:20px 0;text-align:left;">
          <div style="font-weight:700;color:#2c3a4a;margin-bottom:8px;">📅 รายละเอียดงาน</div>
          <div style="color:#5a7080;font-size:0.9rem;line-height:1.8;">
            วันเสาร์ที่ 28 พฤศจิกายน 2569<br>
            ลงทะเบียน 13:30 น. · เริ่ม 14:00 น.<br>
            TARA Terrace · จ.นครปฐม
          </div>
        </div>
        <p style="color:#b09070;font-size:0.8rem;">
          ไฟล์ .ics แนบมาด้วยสำหรับบันทึกใน Calendar ของคุณ
        </p>
        <p style="color:#C9A96E;font-size:0.9rem;margin-top:16px;">ด้วยความรัก เอ็ม &amp; แป้ง 💕</p>
      </div>
    </div>
  `;

  GmailApp.sendEmail(
    email,
    '💍 คุณได้รับเชิญ · งานสักขีพยาน เอ็ม & แป้ง · 28 พ.ย. 69',
    `คุณ ${guestName} ได้รับเชิญร่วมงานสักขีพยาน Family Witness Dinner\n28 พฤศจิกายน 2569 · 14:00 น. · TARA Terrace`,
    {
      htmlBody,
      attachments: [
        Utilities.newBlob(ics, 'text/calendar;method=REQUEST', 'PM_Wedding_Invite.ics')
      ]
    }
  );
}
