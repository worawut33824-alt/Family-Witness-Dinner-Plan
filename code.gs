// ============================================================
//  Family Witness Dinner Plan - Google Apps Script Web App
//  Execute as: Me  |  Who has access: Anyone
// ============================================================

// --- ROUTER ------------------------------------------------------------------
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'all';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var result = { status: 'ok' };

  try {
    if (action === 'ping') {
      result.message = 'connected';
    } else if (action === 'all') {
      result.expenses  = getExpenses(ss);
      result.contacts  = getContacts(ss);
      result.checklist = getChecklist(ss);
      result.guests    = getGuests(ss);
      result.budget    = getBudget(ss);
    } else if (action === 'expenses') {
      result.expenses  = getExpenses(ss);
    } else if (action === 'contacts') {
      result.contacts  = getContacts(ss);
    } else if (action === 'checklist') {
      result.checklist = getChecklist(ss);
    } else if (action === 'guests') {
      result.guests    = getGuests(ss);
    } else if (action === 'budget') {
      result.budget    = getBudget(ss);
    } else if (action === 'format_budget') {
      var r1 = formatBudgetSheet(ss);
      result.ok  = r1.ok;
      result.msg = r1.msg;
    } else if (action === 'format_summary') {
      var r2 = formatSummarySheet(ss);
      result.ok  = r2.ok;
      result.msg = r2.msg;
    } else if (action === 'format_all') {
      var rb = formatBudgetSheet(ss);
      var rs = formatSummarySheet(ss);
      result.budget_msg  = rb.msg;
      result.summary_msg = rs.msg;
    } else if (action === 'update_summary') {
      var rsum = fixDashboardSummaryFormulas(ss);
      result.ok  = rsum.ok;
      result.msg = rsum.msg;
    } else if (action === 'update_dashboard') {
      var rdash = fixDashboardSummaryFormulas(ss);
      result.ok  = rdash.ok;
      result.msg = rdash.msg;
    } else if (action === 'update_summary_dashboard') {
      var rfixed = fixDashboardSummaryFormulas(ss);
      result.ok            = rfixed.ok;
      result.summary_msg   = rfixed.msg;
      result.dashboard_msg = rfixed.msg;
    } else if (action === 'stats') {
      var sc = countSeatsBySide(ss);
      result.guestCount = sc.total;
      result.sideM      = sc.sideM;
      result.sideP      = sc.sideP;
      result.blessings  = getBlessings(ss);

    } else if (action === 'rsvp') {
      saveRsvpGuest(ss, e.parameter);
      var rc = countSeatsBySide(ss);
      result.guestCount = rc.total;
      result.sideM      = rc.sideM;
      result.sideP      = rc.sideP;

    } else if (action === 'bless') {
      saveBlessing(ss, e.parameter);

    } else if (action === 'blessings') {
      result.blessings = getBlessingsForPage(ss);

    } else if (action === 'links') {
      result.links = getLinks(ss);

    } else if (action === 'check_link') {
      var slug = e.parameter.slug || '';
      result   = checkLink(ss, slug);

    } else {
      result.message = 'unknown action';
    }
  } catch (err) {
    result.status = 'error';
    result.error  = err.toString();
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var result = { status: 'ok' };

  try {
    var body = JSON.parse(e.postData.contents);
    var ss   = SpreadsheetApp.getActiveSpreadsheet();

    // --- invite page actions ---
    if (body.action === 'rsvp') {
      saveRsvpGuest(ss, body);
      var rc2 = countSeatsBySide(ss);
      result.guestCount = rc2.total;
      result.sideM      = rc2.sideM;
      result.sideP      = rc2.sideP;
    } else if (body.action === 'bless') {
      saveBlessing(ss, body);

    } else if (body.action === 'links') {
      saveLinks(ss, body.links || []);

    } else if (body.action === 'donation') {
      var donResult = saveDonation(ss, body);
      result.amount  = donResult.amount;
      result.slipUrl = donResult.slipUrl;

    } else if (body.action === 'readSlip') {
      // OCR เท่านั้น ไม่บันทึก
      result.amount = readAmountFromSlip(body.slipBase64 || '', body.slipMime || 'image/jpeg');

    // --- planner app actions ---
    } else {
      if (body.expenses  !== undefined) saveExpenses(ss, body.expenses);
      if (body.contacts  !== undefined) saveContacts(ss, body.contacts);
      if (body.checklist !== undefined) saveChecklist(ss, body.checklist);
      if (body.guests    !== undefined) saveGuests(ss, body.guests);
      if (body.budget    !== undefined) saveBudget(ss, body.budget);
    }
  } catch (err) {
    result.status = 'error';
    result.error  = err.toString();
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// --- COUNT SEATS — นับที่นั่งแยกตามฝั่ง --------------------------
function countAllSeats(ss) {
  return countSeatsBySide(ss).total;
}

function countSeatsBySide(ss) {
  var sheet = ss.getSheetByName('แขก');
  if (!sheet) return { total: 0, sideM: 0, sideP: 0 };
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { total: 0, sideM: 0, sideP: 0 };
  // อ่าน column B (ฝั่ง) และ E (ที่นั่ง)
  var data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  var total = 0, sideM = 0, sideP = 0;
  for (var i = 0; i < data.length; i++) {
    var side  = String(data[i][1] || '').trim();  // column B
    var seats = parseInt(data[i][4]) || 0;         // column E
    if (seats <= 0) continue;
    total += seats;
    if (side.indexOf('เอ็ม') >= 0) sideM += seats;
    else if (side.indexOf('แป้ง') >= 0) sideP += seats;
  }
  return { total: total, sideM: sideM, sideP: sideP };
}

// --- SETUP — รันครั้งเดียวเพื่อสร้าง sheet ที่จำเป็น ----------------
var GUEST_HEADERS = [
  'ชื่อ','ฝั่ง','เบอร์โทร','ความสัมพันธ์',
  'ที่นั่ง','สถานะ','อีเมล','คำอวยพร','รูป URL','วันที่','หมายเหตุ'
];
var BLESS_HEADERS = ['วันที่','ชื่อ','ข้อความ','รูป URL','สถานะเข้าร่วม','ฝั่ง'];

function doSetup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // สร้าง/อัปเดต sheet คำอวยพร
  var s = ss.getSheetByName('คำอวยพร') || ss.insertSheet('คำอวยพร');
  s.getRange(1, 1, 1, BLESS_HEADERS.length).setValues([BLESS_HEADERS])
    .setFontWeight('bold').setBackground('#f4d9d0');
  s.setFrozenRows(1);
  [150,160,420,200,140].forEach(function(w,i){ s.setColumnWidth(i+1,w); });

  // สร้าง/อัปเดต sheet แขก
  var g = ss.getSheetByName('แขก') || ss.insertSheet('แขก');
  g.getRange(1, 1, 1, GUEST_HEADERS.length).setValues([GUEST_HEADERS])
    .setFontWeight('bold').setBackground('#f4d9d0');
  g.setFrozenRows(1);

  // สร้าง folder รูปภาพ ใน Drive
  var folderName = 'Wedding Photos – Family Witness Dinner';
  if (!DriveApp.getFoldersByName(folderName).hasNext()) {
    DriveApp.createFolder(folderName);
  }

  SpreadsheetApp.getUi().alert(
    '✅ Setup เรียบร้อย!\n\n' +
    '• Sheet "แขก" — ' + GUEST_HEADERS.length + ' คอลัมน์\n' +
    '• Sheet "คำอวยพร" — ' + BLESS_HEADERS.length + ' คอลัมน์\n' +
    '• Google Drive folder "' + folderName + '"'
  );
}

// --- PHOTO: บันทึกรูปลง Google Drive → return URL ----------------
function savePictureToDrive(base64Data, filename) {
  try {
    if (!base64Data || base64Data.length < 100) return '';
    var folderName = 'Wedding Photos – Family Witness Dinner';
    var folders = DriveApp.getFoldersByName(folderName);
    var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

    var contentType = 'image/jpeg';
    var raw = base64Data;
    if (raw.indexOf(',') >= 0) {
      var prefix = raw.split(',')[0];
      raw = raw.split(',')[1];
      if (prefix.indexOf('png')  >= 0) contentType = 'image/png';
      if (prefix.indexOf('webp') >= 0) contentType = 'image/webp';
      if (prefix.indexOf('gif')  >= 0) contentType = 'image/gif';
    }

    var bytes = Utilities.base64Decode(raw);
    var blob  = Utilities.newBlob(bytes, contentType, filename || 'photo.jpg');
    var file  = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var id = file.getId();
    return 'https://drive.google.com/thumbnail?id=' + id + '&sz=w800';
  } catch(e) {
    Logger.log('savePictureToDrive error: ' + e);
    return '';
  }
}

// --- RSVP (invite page — บันทึกแขกรายบุคคล) -------------------------
function saveRsvpGuest(ss, body) {
  var sheet = ss.getSheetByName('แขก');
  if (!sheet) {
    sheet = ss.insertSheet('แขก');
    sheet.getRange(1, 1, 1, GUEST_HEADERS.length).setValues([GUEST_HEADERS])
      .setFontWeight('bold').setBackground('#f4d9d0');
    sheet.setFrozenRows(1);
  }

  // อัปโหลดรูปถ้ามี
  var photoUrl = '';
  var photoBase64 = body.photoBase64 || body.photo || '';
  if (photoBase64 && photoBase64.length > 100) {
    var ts = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd_HHmmss');
    photoUrl = savePictureToDrive(photoBase64, 'rsvp_' + ts + '.jpg');
  }

  var dateStr = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm');
  var lastRow = sheet.getLastRow();
  // ชื่อ | ฝั่ง | เบอร์โทร | ความสัมพันธ์ | ที่นั่ง | สถานะ | อีเมล | คำอวยพร | รูป URL | วันที่ | หมายเหตุ
  sheet.getRange(lastRow + 1, 1, 1, 11).setValues([[
    body.name     || '',
    body.side     || '',
    body.phone    || '',
    body.relation || '',
    parseInt(body.count) || 1,
    'confirmed',
    body.email    || '',
    body.blessing || '',
    photoUrl,
    dateStr,
    body.note     || ''
  ]]);

  // บันทึกคำอวยพรลง sheet คำอวยพร ด้วย (ถ้ามี)
  var blessingText = String(body.blessing || '').trim();
  if (blessingText) {
    saveBlessingRow(ss, body.name, blessingText, photoUrl, 'confirmed', body.side || '');
  }

  // ส่ง Google Calendar invite + Email ถ้ามี email
  var guestEmail = String(body.email || '').trim();
  if (guestEmail) {
    // Step A: เพิ่มแขกเข้า Calendar event (แยก try-catch เพื่อไม่ให้ block Email)
    try {
      addGuestToWeddingEvent_(guestEmail);
    } catch(calErr) {
      Logger.log('addGuestToWeddingEvent error: ' + calErr);
    }

    // Step B: ส่ง HTML Email + .ics แนบ (ทำงานอิสระจาก Calendar)
    try {
      sendWeddingEmail_(guestEmail, body.name || '', parseInt(body.count) || 1);
    } catch(mailErr) {
      Logger.log('sendWeddingEmail error: ' + mailErr);
    }
  }
}

// --- CALENDAR INVITE (Hybrid) ------------------------------------------------
// ทำงาน 2 ขั้นตอนพร้อมกัน:
// 1) เพิ่มแขกเข้า event เดียวบนปฏิทินเอ็ม ผ่าน Calendar Advanced Service
//    → Gmail users ได้รับ invite อัตโนมัติ กด Accept ได้ทันที
// 2) ส่ง HTML email + .ics แนบ ให้ทุก email client
//    → มีปุ่ม Google Calendar / Apple Calendar / Outlook ให้กดชัดเจน
// -----------------------------------------------------------------------------

var WEDDING_EVENT_ID_KEY = 'WEDDING_EVENT_ID';
var WEDDING_CAL_ID       = 'primary';
var WEDDING_START_UTC    = '20261128T090000Z'; // 16:00 Bangkok
var WEDDING_END_UTC      = '20261128T130000Z'; // 20:00 Bangkok
var WEDDING_UID          = 'pmwedding-20261128@family-witness-dinner';

// สร้าง event ครั้งแรกหรือดึง event เดิม — คืนค่า eventId
// ลำดับการค้นหา: 1) Script Property → 2) ค้นหาจากชื่องานใน Calendar → 3) สร้างใหม่
function getOrCreateWeddingEvent_() {
  var props   = PropertiesService.getScriptProperties();
  var eventId = props.getProperty(WEDDING_EVENT_ID_KEY);

  // 1) ตรวจ Script Property ก่อน
  if (eventId) {
    try {
      Calendar.Events.get(WEDDING_CAL_ID, eventId);
      return eventId; // event ยังอยู่ — ใช้ได้เลย
    } catch(e) {
      Logger.log('getOrCreateWeddingEvent_: stored event not found, searching... ' + e);
      props.deleteProperty(WEDDING_EVENT_ID_KEY);
    }
  }

  // 2) ค้นหา event เดิมจากชื่องานและวันที่ (ป้องกันสร้างซ้ำ)
  try {
    var existing = Calendar.Events.list(WEDDING_CAL_ID, {
      timeMin:      '2026-11-28T00:00:00+07:00',
      timeMax:      '2026-11-28T23:59:59+07:00',
      singleEvents: true,
      maxResults:   20
    });
    var items = (existing && existing.items) ? existing.items : [];
    for (var i = 0; i < items.length; i++) {
      if (items[i].summary && items[i].summary.indexOf('Family Witness Dinner') >= 0) {
        Logger.log('getOrCreateWeddingEvent_: found existing event ' + items[i].id);
        props.setProperty(WEDDING_EVENT_ID_KEY, items[i].id);
        return items[i].id;
      }
    }
  } catch(searchErr) {
    Logger.log('getOrCreateWeddingEvent_: search error ' + searchErr);
  }

  // 3) ไม่พบ event เดิม — สร้างใหม่ครั้งเดียว
  var newEvent = Calendar.Events.insert({
    summary:     'เอ็ม x แป้ง — Family Witness Dinner',
    location:    'ธาราเทอเรส (TARA Terrace), นครปฐม',
    description: 'งานสักขีพยาน Family Witness Dinner\nเอ็ม & แป้ง\n28 พฤศจิกายน 2569 16:00-20:00 น.\nธาราเทอเรส นครปฐม',
    start:       { dateTime: '2026-11-28T16:00:00+07:00', timeZone: 'Asia/Bangkok' },
    end:         { dateTime: '2026-11-28T20:00:00+07:00', timeZone: 'Asia/Bangkok' },
    status:      'confirmed'
  }, WEDDING_CAL_ID, { sendUpdates: 'none' });

  props.setProperty(WEDDING_EVENT_ID_KEY, newEvent.id);
  Logger.log('getOrCreateWeddingEvent_: created new event ' + newEvent.id);
  return newEvent.id;
}

// ── ลบ event ซ้ำใน Calendar ของเจ้าของ (รันครั้งเดียวจาก Apps Script Editor) ──
// เก็บ event ที่มีแขกมากที่สุดไว้ ลบอันที่เหลือ
function cleanupDuplicateEvents() {
  try {
    var existing = Calendar.Events.list(WEDDING_CAL_ID, {
      timeMin:      '2026-11-28T00:00:00+07:00',
      timeMax:      '2026-11-28T23:59:59+07:00',
      singleEvents: true,
      maxResults:   50
    });
    var items = (existing && existing.items) ? existing.items : [];
    var matches = items.filter(function(ev) {
      return ev.summary && ev.summary.indexOf('Family Witness Dinner') >= 0;
    });

    if (matches.length <= 1) {
      Logger.log('cleanupDuplicateEvents: ไม่มี event ซ้ำ (พบ ' + matches.length + ' event)');
      SpreadsheetApp.getUi().alert('✅ ไม่มี event ซ้ำ พบ ' + matches.length + ' event เท่านั้น');
      return;
    }

    // เรียงให้ event ที่มีแขกมากสุดอยู่ก่อน — เก็บอันนั้นไว้
    matches.sort(function(a, b) {
      return ((b.attendees || []).length) - ((a.attendees || []).length);
    });

    var keepEvent = matches[0];
    var props     = PropertiesService.getScriptProperties();
    props.setProperty(WEDDING_EVENT_ID_KEY, keepEvent.id);
    Logger.log('cleanupDuplicateEvents: เก็บ event ' + keepEvent.id +
               ' (แขก ' + (keepEvent.attendees || []).length + ' คน)');

    var deleted = [];
    for (var i = 1; i < matches.length; i++) {
      try {
        Calendar.Events.remove(WEDDING_CAL_ID, matches[i].id);
        deleted.push(matches[i].id);
        Logger.log('cleanupDuplicateEvents: ลบ event ' + matches[i].id);
      } catch(delErr) {
        Logger.log('cleanupDuplicateEvents: ลบไม่ได้ ' + matches[i].id + ' — ' + delErr);
      }
    }

    SpreadsheetApp.getUi().alert(
      '✅ ล้าง event ซ้ำเสร็จ\n\n' +
      '• เก็บไว้: 1 event (แขก ' + (keepEvent.attendees || []).length + ' คน)\n' +
      '• ลบออก: ' + deleted.length + ' event\n\n' +
      'Event ID ที่ใช้งาน:\n' + keepEvent.id
    );
  } catch(err) {
    Logger.log('cleanupDuplicateEvents error: ' + err);
    SpreadsheetApp.getUi().alert('❌ เกิดข้อผิดพลาด: ' + err);
  }
}

// เพิ่มแขกเข้า event เดิม + ส่ง invite อัตโนมัติ (Gmail)
function addGuestToWeddingEvent_(toEmail) {
  var eventId = getOrCreateWeddingEvent_();

  // ดึง attendees ปัจจุบัน
  var ev         = Calendar.Events.get(WEDDING_CAL_ID, eventId);
  var attendees  = ev.attendees || [];

  // ตรวจไม่ให้ซ้ำ
  var alreadyIn  = attendees.some(function(a) { return a.email === toEmail; });
  if (!alreadyIn) {
    attendees.push({ email: toEmail });
    Calendar.Events.patch(
      { attendees: attendees },
      WEDDING_CAL_ID,
      eventId,
      { sendUpdates: 'none' }
      // 'none' = ไม่ให้ Calendar ส่ง email ของตัวเอง — ใช้ sendWeddingEmail_ แทน
    );
  }
}

// สร้างเนื้อหา .ics — ใช้ UID เดิมเสมอ เพื่อให้บันทึกทับ event เดิมในปฏิทินแขก
function buildIcs_(guestName, seats) {
  var seatsLabel = seats > 1 ? guestName + ' +' + (seats - 1) : guestName;
  return [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    'PRODID:-//PM Wedding//Family Witness Dinner//TH',
    'METHOD:REQUEST',
    'BEGIN:VEVENT',
    'UID:' + WEDDING_UID,
    'DTSTART:' + WEDDING_START_UTC,
    'DTEND:'   + WEDDING_END_UTC,
    'SUMMARY:เอ็ม x แป้ง — Family Witness Dinner',
    'LOCATION:ธาราเทอเรส (TARA Terrace)\\, นครปฐม',
    'DESCRIPTION:ยินดีต้อนรับคุณ ' + seatsLabel + '\\n' +
      'งานสักขีพยาน Family Witness Dinner\\n' +
      'ลงทะเบียน 16:00 น. · เริ่มงาน 16:00-20:00 น.\\n' +
      'ธาราเทอเรส (TARA Terrace) นครปฐม',
    'STATUS:CONFIRMED',
    'SEQUENCE:0',
    'END:VEVENT',
    'END:VCALENDAR'
  ].join('\r\n');
}

// Google Calendar link (เปิดหน้าเพิ่มลง calendar โดยตรง)
function buildGoogleCalLink_() {
  return 'https://calendar.google.com/calendar/render?action=TEMPLATE' +
    '&text=' + encodeURIComponent('เอ็ม x แป้ง — Family Witness Dinner') +
    '&dates=' + WEDDING_START_UTC + '/' + WEDDING_END_UTC +
    '&details=' + encodeURIComponent('งานสักขีพยาน Family Witness Dinner\nเอ็ม & แป้ง\n28 พฤศจิกายน 2569 16:00-20:00 น.\nธาราเทอเรส นครปฐม') +
    '&location=' + encodeURIComponent('ธาราเทอเรส (TARA Terrace), นครปฐม') +
    '&sf=true&output=xml';
}

// sendCalendarInvite ยังคงไว้เพื่อ backward-compat (ไม่มีใครเรียกแล้ว แต่เผื่อ)
function sendCalendarInvite(toEmail, guestName, seats) {
  try { addGuestToWeddingEvent_(toEmail); } catch(e) { Logger.log(e); }
  try { sendWeddingEmail_(toEmail, guestName, seats); } catch(e) { Logger.log(e); }
}

// ── ส่ง HTML Email + .ics แนบ ────────────────────────────────────────────────
function sendWeddingEmail_(toEmail, guestName, seats) {
  var seatsLabel   = seats > 1 ? guestName + ' และอีก ' + (seats - 1) + ' ท่าน' : guestName;
  var googleCalUrl = buildGoogleCalLink_();

  // ── สร้าง .ics สำหรับแนบ email ───────────────────────────────────────────
  var icsContent = buildIcs_(guestName, seats);
  var icsBlob    = Utilities.newBlob(icsContent, 'text/calendar;method=REQUEST', 'PM_Wedding_Invite.ics');

  // ── Step 3: ส่ง HTML email ──────────────────────────────────────────────
  var htmlBody =
    '<div style="background:#f4f4f4;padding:24px 0;font-family:Arial,sans-serif;">' +
    '<div style="background:#ffffff;border-radius:16px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,0.12);max-width:560px;margin:0 auto;">' +

    // Header
    '<div style="background:linear-gradient(135deg,#4a7fa5 0%,#6b9bbe 50%,#8fb5d0 100%);padding:36px 32px 28px;text-align:center;color:#fff;">' +
    '<div style="font-family:Georgia,serif;font-size:26px;font-style:italic;font-weight:bold;letter-spacing:1px;margin-bottom:4px;">เอ็ม &amp; แป้ง</div>' +
    '<div style="font-size:11px;letter-spacing:4px;opacity:0.85;text-transform:uppercase;">Family Witness Dinner</div>' +
    '<div style="margin-top:16px;background:rgba(255,255,255,0.18);border-radius:20px;display:inline-block;padding:5px 18px;font-size:12px;letter-spacing:1px;">28 พฤศจิกายน 2569</div>' +
    '</div>' +

    // Body
    '<div style="padding:28px 32px 0;">' +
    '<p style="font-size:15px;color:#2c3a4a;margin:0 0 6px;">เรียน คุณ <strong>' + seatsLabel + '</strong>,</p>' +
    '<p style="color:#5a7080;font-size:14px;line-height:1.8;margin:0 0 24px;">' +
    'ขอบคุณมากที่ยืนยันร่วมงานสักขีพยาน Family Witness Dinner ของเรา<br>' +
    'เราตั้งตารอพบคุณในวันงานอย่างยิ่งค่ะ' +
    '</p>' +

    // Event Details
    '<div style="background:#f0f6fb;border-radius:12px;padding:20px 22px;margin-bottom:22px;border-left:4px solid #4a7fa5;">' +
    '<div style="font-weight:700;color:#2c3a4a;font-size:14px;margin-bottom:12px;">รายละเอียดงาน</div>' +
    '<table style="width:100%;border-collapse:collapse;font-size:13.5px;">' +
    '<tr><td style="padding:5px 0;color:#7a9ab0;width:90px;vertical-align:top;">วันที่</td>' +
    '<td style="padding:5px 0;color:#2c3a4a;font-weight:600;">วันเสาร์ที่ 28 พฤศจิกายน 2569<br>' +
    '<span style="font-size:11px;color:#999;font-weight:normal;">Saturday, 28 November 2026</span></td></tr>' +
    '<tr><td style="padding:5px 0;color:#7a9ab0;vertical-align:top;">เวลา</td>' +
    '<td style="padding:5px 0;color:#2c3a4a;font-weight:600;">16:00 – 20:00 น.<br>' +
    '<span style="font-size:11px;color:#999;font-weight:normal;">ลงทะเบียน 16:00 น.</span></td></tr>' +
    '<tr><td style="padding:5px 0;color:#7a9ab0;vertical-align:top;">สถานที่</td>' +
    '<td style="padding:5px 0;color:#2c3a4a;font-weight:600;">ธาราเทอเรส (TARA Terrace)<br>' +
    '<span style="font-size:11px;color:#999;font-weight:normal;">นครปฐม</span></td></tr>' +
    '<tr><td style="padding:5px 0;color:#7a9ab0;vertical-align:top;">จำนวน</td>' +
    '<td style="padding:5px 0;color:#2c3a4a;font-weight:600;">' + seats + ' ท่าน</td></tr>' +
    '</table>' +
    '</div>' +

    // Calendar Section
    '<div style="background:#fff8f0;border:1.5px solid #f0dcc8;border-radius:14px;padding:22px;margin-bottom:24px;text-align:center;">' +
    '<div style="font-size:13px;color:#7a6050;font-weight:700;margin-bottom:4px;">บันทึกลงปฏิทินของคุณ</div>' +
    '<div style="font-size:12px;color:#aaa;margin-bottom:18px;">เลือกวิธีที่สะดวกสำหรับคุณได้เลยค่ะ</div>' +

    // Button row: Google + Apple
    '<table style="width:100%;border-collapse:collapse;margin-bottom:10px;"><tr>' +
    '<td style="padding:0 5px 0 0;width:50%;">' +
    '<a href="' + googleCalUrl + '" target="_blank" ' +
    'style="display:block;background:#4285F4;color:#fff;text-decoration:none;padding:11px 10px;border-radius:8px;font-size:13px;font-weight:600;text-align:center;">' +
    'Google Calendar</a></td>' +
    '<td style="padding:0 0 0 5px;width:50%;">' +
    '<a href="cid:wedding.ics" ' +
    'style="display:block;background:#1c1c1e;color:#fff;text-decoration:none;padding:11px 10px;border-radius:8px;font-size:13px;font-weight:600;text-align:center;">' +
    'Apple Calendar</a></td>' +
    '</tr></table>' +

    // Outlook button
    '<a href="cid:wedding.ics" ' +
    'style="display:block;background:#0078d4;color:#fff;text-decoration:none;padding:11px 20px;border-radius:8px;font-size:13px;font-weight:600;text-align:center;max-width:300px;margin:0 auto;">' +
    'Outlook / ปฏิทินอื่น — เปิดไฟล์แนบ .ics</a>' +

    '<div style="margin-top:14px;font-size:11.5px;color:#bbb;line-height:1.6;">' +
    'ไฟล์ <strong>PM_Wedding_Invite.ics</strong> แนบมาใน email นี้ด้วยค่ะ<br>' +
    'เปิดไฟล์แนบเพื่อเพิ่มลงปฏิทินได้เลย' +
    '</div>' +
    '</div>' + // end calendar section
    '</div>' + // end body

    // Footer
    '<div style="padding:0 32px 28px;text-align:center;border-top:1px solid #f5efe9;padding-top:22px;">' +
    '<img src="https://drive.google.com/thumbnail?id=148E0apdHMhde4NV87DZFym-LutL8Mt3z&sz=w300" ' +
    'alt="เอ็ม x แป้ง" style="width:160px;height:auto;border-radius:10px;margin-bottom:16px;display:block;margin-left:auto;margin-right:auto;" />' +
    '<div style="color:#c9a97a;font-size:15px;margin-bottom:8px;">ด้วยความรักและตั้งตารอ</div>' +
    '<div style="font-family:Georgia,serif;font-size:20px;font-style:italic;color:#4a7fa5;font-weight:bold;">เอ็ม &amp; แป้ง</div>' +
    '<div style="font-size:11px;color:#bbb;margin-top:10px;letter-spacing:1px;">FAMILY WITNESS DINNER · 28.11.2569</div>' +
    '</div>' +

    '</div>' + // end email container
    '</div>';  // end outer wrapper

  GmailApp.sendEmail(
    toEmail,
    'ยืนยันร่วมงาน Family Witness Dinner — เอ็ม x แป้ง',
    'ขอบคุณที่ยืนยันร่วมงานสักขีพยาน Family Witness Dinner\nวันเสาร์ที่ 28 พ.ย. 2569 | 16:00-20:00 น. | ธาราเทอเรส นครปฐม\nบันทึกลงปฏิทิน: ' + googleCalUrl,
    {
      htmlBody:    htmlBody,
      attachments: [icsBlob]
    }
  );
}

// --- BLESSINGS -------------------------------------------------------
function getBlessings(ss) {
  var sheet = ss.getSheetByName('คำอวยพร');
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var cols = Math.max(sheet.getLastColumn(), 6);
  var data = sheet.getRange(2, 1, lastRow - 1, cols).getValues();
  var list = [];
  for (var i = data.length - 1; i >= 0; i--) {
    if (!data[i][1]) continue;
    list.push({
      date:       data[i][0] instanceof Date ? Utilities.formatDate(data[i][0], 'Asia/Bangkok', 'dd/MM/yyyy HH:mm') : String(data[i][0] || ''),
      name:       String(data[i][1] || ''),
      message:    String(data[i][2] || ''),
      photoUrl:   String(data[i][3] || ''),
      attendance: String(data[i][4] || ''),
      side:       String(data[i][5] || '')
    });
  }
  return list.slice(0, 50);
}

// คืนข้อมูลทั้งหมดสำหรับหน้า blessings (ไม่จำกัด 50)
function getBlessingsForPage(ss) {
  var sheet = ss.getSheetByName('คำอวยพร');
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var cols = Math.max(sheet.getLastColumn(), 6);
  var data = sheet.getRange(2, 1, lastRow - 1, cols).getValues();
  var list = [];
  for (var i = data.length - 1; i >= 0; i--) {
    if (!data[i][1]) continue;
    list.push({
      date:       data[i][0] instanceof Date ? Utilities.formatDate(data[i][0], 'Asia/Bangkok', 'dd/MM/yyyy HH:mm') : String(data[i][0] || ''),
      name:       String(data[i][1] || ''),
      message:    String(data[i][2] || ''),
      photoUrl:   String(data[i][3] || ''),
      attendance: String(data[i][4] || ''),
      side:       String(data[i][5] || '')
    });
  }
  return list;
}

// helper: เขียน row ลง sheet คำอวยพร
function saveBlessingRow(ss, name, message, photoUrl, attendance, side) {
  var sheet = ss.getSheetByName('คำอวยพร');
  if (!sheet) {
    sheet = ss.insertSheet('คำอวยพร');
    sheet.getRange(1, 1, 1, BLESS_HEADERS.length).setValues([BLESS_HEADERS])
      .setFontWeight('bold').setBackground('#f4d9d0');
    sheet.setFrozenRows(1);
    [150,160,420,200,140,120].forEach(function(w,i){ sheet.setColumnWidth(i+1,w); });
  }
  var dateStr = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm');
  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1).setNumberFormat('@');
  sheet.getRange(lastRow + 1, 1, 1, 6).setValues([[
    dateStr,
    String(name    || '').substring(0, 80),
    String(message || '').substring(0, 500),
    String(photoUrl || ''),
    String(attendance || ''),
    String(side || '')
  ]]);
}

function saveBlessing(ss, body) {
  // อัปโหลดรูปถ้ามี
  var photoUrl = '';
  var photoBase64 = body.photoBase64 || body.photo || '';
  if (photoBase64 && photoBase64.length > 100) {
    var ts = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd_HHmmss');
    photoUrl = savePictureToDrive(photoBase64, 'bless_' + ts + '.jpg');
  }

  // บันทึกลง sheet คำอวยพร
  saveBlessingRow(ss, body.name, body.message, photoUrl, 'declined', body.side || '');

  // บันทึกลง sheet แขก ด้วย (สถานะ = declined, ที่นั่ง = 0)
  var gSheet = ss.getSheetByName('แขก');
  if (!gSheet) {
    gSheet = ss.insertSheet('แขก');
    gSheet.getRange(1, 1, 1, GUEST_HEADERS.length).setValues([GUEST_HEADERS])
      .setFontWeight('bold').setBackground('#f4d9d0');
    gSheet.setFrozenRows(1);
  }
  var dateStr = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm');
  var lastRow = gSheet.getLastRow();
  gSheet.getRange(lastRow + 1, 1, 1, 11).setValues([[
    String(body.name    || ''),
    String(body.side    || ''),   // ฝั่ง
    '',   // เบอร์
    '',   // ความสัมพันธ์
    0,    // ที่นั่ง = 0
    'declined',
    '',   // อีเมล
    String(body.message || ''),
    photoUrl,
    dateStr,
    ''
  ]]);
}

// --- EXPENSES ----------------------------------------------------------------
function getExpenses(ss) {
  var sheet = ss.getSheetByName('บันทึกค่าใช้จ่าย');
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 4) return [];

  var data = sheet.getRange(4, 1, lastRow - 3, 5).getValues();
  var expenses = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (!row[0] && !row[2]) continue;
    var dateStr = '';
    if (row[0] instanceof Date) {
      dateStr = Utilities.formatDate(row[0], 'Asia/Bangkok', 'yyyy-MM-dd');
    } else if (row[0]) {
      dateStr = String(row[0]).substring(0, 10);
    }
    expenses.push({
      id:     i + 1,
      date:   dateStr,
      cat:    String(row[1] || '').trim(),
      item:   String(row[2] || '').trim(),
      amount: parseFloat(row[3]) || 0,
      note:   String(row[4] || '').trim()
    });
  }
  return expenses.filter(function(e) { return e.date || e.item; });
}

function saveExpenses(ss, expenses) {
  var sheet = ss.getSheetByName('บันทึกค่าใช้จ่าย');
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  if (lastRow >= 4) sheet.getRange(4, 1, lastRow - 3, 5).clearContent();
  if (!expenses || !expenses.length) return;

  var rows = [];
  for (var i = 0; i < expenses.length; i++) {
    var exp = expenses[i];
    rows.push([exp.date, exp.cat, exp.item, exp.amount, exp.note]);
  }
  sheet.getRange(4, 1, rows.length, 5).setValues(rows);
}

// --- CONTACTS ----------------------------------------------------------------
function getContacts(ss) {
  var sheet = ss.getSheetByName('Contact ทีมงาน');
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  var contacts = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (!row[0] && !row[1]) continue;
    contacts.push({
      id:          i + 1,
      role:        String(row[0] || '').trim(),
      name:        String(row[1] || '').trim(),
      phone:       String(row[2] || '').trim(),
      email:       String(row[3] || '').trim(),
      social:      String(row[4] || '').trim(),
      contactedBy: String(row[5] || '').trim(),
      note:        ''
    });
  }
  return contacts;
}

function saveContacts(ss, contacts) {
  var sheet = ss.getSheetByName('Contact ทีมงาน');
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) sheet.getRange(2, 1, lastRow - 1, 6).clearContent();
  if (!contacts || !contacts.length) return;

  // ตรวจสอบ / สร้าง header row 1
  var h1 = String(sheet.getRange(1, 6).getValue()).trim();
  if (!h1) {
    sheet.getRange(1, 6).setValue('ติดต่อโดย').setFontWeight('bold');
  }

  var rows = [];
  for (var i = 0; i < contacts.length; i++) {
    var c = contacts[i];
    rows.push([c.role, c.name, c.phone, c.email, c.social, c.contactedBy || '']);
  }
  sheet.getRange(2, 1, rows.length, 6).setValues(rows);
}

// --- CHECKLIST ---------------------------------------------------------------
var CHECKLIST_ROWS = [3,4,5,6,7,8,9,10, 12,13,14,15,16,17, 19,20,21,22, 24,25,26,27, 29,30,31,32];

function getChecklist(ss) {
  var sheet = ss.getSheetByName('Check List');
  if (!sheet) return {};
  var result = {};

  for (var i = 0; i < CHECKLIST_ROWS.length; i++) {
    var row  = CHECKLIST_ROWS[i];
    var cell = String(sheet.getRange(row, 1).getValue()).trim();
    var status = 'none';
    if (cell === '✓') status = 'done';
    else if (cell === '✗') status = 'skipped';
    result['c' + (i + 1)] = status;
  }
  return result;
}

function saveChecklist(ss, checklist) {
  var sheet = ss.getSheetByName('Check List');
  if (!sheet) return;

  var statusMap = {};
  if (Array.isArray(checklist)) {
    for (var pi = 0; pi < checklist.length; pi++) {
      var period = checklist[pi];
      for (var ii = 0; ii < period.items.length; ii++) {
        var item = period.items[ii];
        statusMap[item.id] = item.status;
      }
    }
  } else if (typeof checklist === 'object') {
    statusMap = checklist;
  }

  for (var i = 0; i < CHECKLIST_ROWS.length; i++) {
    var row = CHECKLIST_ROWS[i];
    var id  = 'c' + (i + 1);
    var appStatus = statusMap[id];
    if (appStatus === undefined) continue;
    var symbol = '';
    if (appStatus === 'done')    symbol = '✓';
    if (appStatus === 'skipped') symbol = '✗';
    sheet.getRange(row, 1).setValue(symbol);
  }
}

// --- BUDGET ------------------------------------------------------------------
// Layout หลัง format_budget:
//   Row 1=Title  Row 2=Info  Row 3=Sep  Row 4=Headers  Row 5+=Data
// Columns: A=category  B=catBudget  C=itemName  D=itemBudget  E=actual(formula)  F=remain(formula)
// แถวหมวดหมู่: A=cat, B=catBudget, C='', D=''
// แถวรายการ:   A=cat, B='',        C=item, D=itemBudget

function budgetDataStart(sheet) {
  if (!sheet) return 2;
  var last = sheet.getLastRow();
  if (last >= 4) {
    var a4 = String(sheet.getRange(4, 1).getValue()).trim();
    if (a4 === 'หมวดหมู่') return 5;
  }
  return 2;
}

// แปลง raw rows → โครงสร้าง {cat, budget, items[]}
function parseBudgetRows(rawRows) {
  var catMap   = {};
  var catOrder = [];
  for (var i = 0; i < rawRows.length; i++) {
    var row        = rawRows[i];
    var cat        = String(row[0] || '').trim();
    var catBudget  = parseFloat(row[1]) || 0;
    var itemName   = String(row[2] || '').trim();
    var itemBudget = parseFloat(row[3]) || 0;
    if (!cat) continue;
    if (!catMap[cat]) {
      catMap[cat] = { cat: cat, budget: catBudget, items: [] };
      catOrder.push(cat);
    } else if (catBudget > 0 && catMap[cat].budget === 0) {
      catMap[cat].budget = catBudget;
    }
    if (itemName) {
      catMap[cat].items.push({ name: itemName, budget: itemBudget });
    }
  }
  return catOrder.map(function(c) { return catMap[c]; });
}

function getBudget(ss) {
  var sheet = ss.getSheetByName('งบประมาณ');
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  var dataRow = budgetDataStart(sheet);
  if (lastRow < dataRow) return [];

  var rawRows = sheet.getRange(dataRow, 1, lastRow - dataRow + 1, 4).getValues();
  return parseBudgetRows(rawRows);
}

function saveBudget(ss, budget) {
  if (!budget || !Array.isArray(budget) || !budget.length) return;

  var sheet = ss.getSheetByName('งบประมาณ');

  // ── สร้าง sheet ใหม่ ──────────────────────────────────────────────────────
  if (!sheet) {
    sheet = ss.insertSheet('งบประมาณ');
    // simple header (ใช้ format_budget ตกแต่งทีหลัง)
    sheet.getRange(1, 1, 1, 4)
      .setValues([['หมวดหมู่','งบรวมหมวด','รายการ','งบรายการ']])
      .setFontWeight('bold');
    sheet.setFrozenRows(1);

    var newRows = [];
    for (var ni = 0; ni < budget.length; ni++) {
      var nc = budget[ni];
      newRows.push([nc.cat, nc.budget, '', '']);           // category header
      for (var nj = 0; nj < nc.items.length; nj++) {
        newRows.push([nc.cat, '', nc.items[nj].name, nc.items[nj].budget]);
      }
    }
    if (newRows.length) sheet.getRange(2, 1, newRows.length, 4).setValues(newRows);
    return;
  }

  // ── อัปเดตเฉพาะตัวเลข col B (catBudget) และ D (itemBudget) ──────────────
  // ไม่แตะ col A, C, E (formula จ่ายจริง), F (formula คงเหลือ)
  var dataRow = budgetDataStart(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow < dataRow) return;

  var data = sheet.getRange(dataRow, 1, lastRow - dataRow + 1, 4).getValues();

  var catBudgetMap  = {};
  var itemBudgetMap = {};
  for (var bi = 0; bi < budget.length; bi++) {
    var bc = budget[bi];
    catBudgetMap[bc.cat] = bc.budget;
    for (var bj = 0; bj < bc.items.length; bj++) {
      itemBudgetMap[bc.cat + '|' + bc.items[bj].name] = bc.items[bj].budget;
    }
  }

  var newColB = [];
  var newColD = [];

  for (var i = 0; i < data.length; i++) {
    var row      = data[i];
    var cat      = String(row[0] || '').trim();
    var curB     = row[1];
    var itemName = String(row[2] || '').trim();
    var curD     = row[3];

    // col B: อัปเดตเฉพาะแถวหมวดหมู่ (C ว่าง = header row)
    if (cat && !itemName && catBudgetMap.hasOwnProperty(cat)) {
      newColB.push([catBudgetMap[cat]]);
    } else {
      newColB.push([curB]);
    }

    // col D: อัปเดตแถว item
    var key = cat + '|' + itemName;
    if (cat && itemName && itemBudgetMap.hasOwnProperty(key)) {
      newColD.push([itemBudgetMap[key]]);
    } else {
      newColD.push([curD]);
    }
  }

  sheet.getRange(dataRow, 2, newColB.length, 1).setValues(newColB);
  sheet.getRange(dataRow, 4, newColD.length, 1).setValues(newColD);
}

// ─────────────────────────────────────────────────────────────────────────────
// FORMAT BUDGET SHEET
// Columns: A=หมวดหมู่  B=งบหมวด  C=รายการ  D=งบรายการ  E=จ่ายจริง(SUMIF)  F=คงเหลือ
// ─────────────────────────────────────────────────────────────────────────────
function formatBudgetSheet(ss) {
  var sheet = ss.getSheetByName('งบประมาณ');
  if (!sheet) return { ok: false, msg: 'ไม่พบ sheet งบประมาณ — sync ข้อมูลก่อน' };

  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return { ok: false, msg: 'sheet ว่างเปล่า — sync ข้อมูลก่อน' };

  // ── 1. อ่าน + parse ข้อมูล ──────────────────────────────────────────────
  var dataStart = budgetDataStart(sheet);
  var rawRows   = [];
  if (lastRow >= dataStart) {
    rawRows = sheet.getRange(dataStart, 1, lastRow - dataStart + 1, 4).getValues();
  }
  var cats = parseBudgetRows(rawRows);
  if (cats.length === 0) return { ok: false, msg: 'ไม่มีข้อมูลงบประมาณ — sync ข้อมูลก่อน' };

  // ── 2. สร้าง newRows โครงสร้างใหม่ (header row + item rows) ─────────────
  var newRows = [];
  for (var ci = 0; ci < cats.length; ci++) {
    var c = cats[ci];
    newRows.push([c.cat, c.budget, '', '']);              // category header
    for (var ii = 0; ii < c.items.length; ii++) {
      newRows.push([c.cat, '', c.items[ii].name, c.items[ii].budget]);
    }
  }

  // ── 3. ล้าง sheet ────────────────────────────────────────────────────────
  sheet.clear();
  sheet.clearFormats();
  try { sheet.getRange(1, 1, 1, 6).breakApart(); } catch(ex) {}

  // ── 4. Column widths (6 คอลัมน์) ─────────────────────────────────────────
  sheet.setColumnWidth(1, 195);   // A หมวดหมู่
  sheet.setColumnWidth(2, 140);   // B งบหมวด
  sheet.setColumnWidth(3, 235);   // C รายการ
  sheet.setColumnWidth(4, 140);   // D งบรายการ
  sheet.setColumnWidth(5, 130);   // E จ่ายจริง
  sheet.setColumnWidth(6, 130);   // F คงเหลือ

  // ── 5. Row heights ───────────────────────────────────────────────────────
  sheet.setRowHeight(1, 48);
  sheet.setRowHeight(2, 32);
  sheet.setRowHeight(3, 6);
  sheet.setRowHeight(4, 30);

  // ── 6. Row 1: Title ───────────────────────────────────────────────────────
  sheet.getRange(1, 1, 1, 6).merge()
    .setValue('💍  งบประมาณการจัดงานสักขีพยาน  ·  Family Witness Dinner')
    .setBackground('#7a3f4a').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(14)
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(false);

  // ── 7. Row 2: Event info + งบรวม ─────────────────────────────────────────
  var totalBudget = 0;
  for (var ti = 0; ti < cats.length; ti++) totalBudget += cats[ti].budget;
  sheet.getRange(2, 1, 1, 6).merge()
    .setValue('📅  28 พฤศจิกายน 2569  ·  TARA Terrace  ·  14:00 น.  ·  งบรวม ' +
              totalBudget.toLocaleString() + ' บาท')
    .setBackground('#b5838d').setFontColor('#ffffff')
    .setFontSize(11).setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(false);

  // ── 8. Row 3: เส้นคั่น ────────────────────────────────────────────────────
  sheet.getRange(3, 1, 1, 6).merge().setBackground('#e8c4c4');

  // ── 9. Row 4: หัวคอลัมน์ ──────────────────────────────────────────────────
  sheet.getRange(4, 1, 1, 6)
    .setValues([['หมวดหมู่', 'งบประมาณหมวด', 'รายการ', 'งบประมาณรายการ', 'จ่ายจริง', 'คงเหลือ']])
    .setBackground('#c9a84c').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setFrozenRows(4);

  // ── 10. เขียนข้อมูล row 5+ ────────────────────────────────────────────────
  sheet.getRange(5, 1, newRows.length, 4).setValues(newRows);

  // ── 11. สีและ formula แต่ละแถว ────────────────────────────────────────────
  var C_CAT_BG   = '#f4d9d0';
  var C_TEXT     = '#3d2c2c';
  var C_ODD      = '#ffffff';
  var C_EVEN     = '#fdf8f6';
  var C_ACT_CAT  = '#fef3d0';   // จ่ายจริง แถวหมวด
  var C_ACT_ITEM = '#f9f9f9';   // จ่ายจริง แถว item
  var C_REM_POS  = '#d4f5e2';   // คงเหลือ > 0
  var C_REM_NEG  = '#fde8e8';   // คงเหลือ < 0 (สูตรแสดงสีอัตโนมัติ)
  var C_BORDER   = '#d4a0a8';
  var SN         = "'บันทึกค่าใช้จ่าย'";  // ชื่อ sheet ค่าใช้จ่าย

  var itemIdx = 0;

  for (var i = 0; i < newRows.length; i++) {
    var rn   = i + 5;
    var row  = newRows[i];
    var cat  = String(row[0] || '').trim();
    var item = String(row[2] || '').trim();

    sheet.setRowHeight(rn, 22);

    // ── สูตร จ่ายจริง (col E) ──────────────────────────────────────────────
    // แถวหมวด: SUMIFS category เท่านั้น
    // แถว item: SUMIFS category + item
    var fActual, fRemain;
    if (!item) {
      // category header row
      fActual = '=IFERROR(SUMIFS(' + SN + '!$D:$D,' + SN + '!$B:$B,$A' + rn + '),0)';
      fRemain = '=$B' + rn + '-$E' + rn;
    } else {
      // item row
      fActual = '=IFERROR(SUMIFS(' + SN + '!$D:$D,' + SN + '!$B:$B,$A' + rn + ',' + SN + '!$C:$C,$C' + rn + '),0)';
      fRemain = '=$D' + rn + '-$E' + rn;
    }
    sheet.getRange(rn, 5).setFormula(fActual);
    sheet.getRange(rn, 6).setFormula(fRemain);

    // ── สีแต่ละแถว ────────────────────────────────────────────────────────
    if (!item) {
      // category header
      sheet.getRange(rn, 1, 1, 4)
        .setBackground(C_CAT_BG).setFontColor(C_TEXT).setFontWeight('bold').setFontSize(10);
      sheet.getRange(rn, 5)
        .setBackground(C_ACT_CAT).setFontColor(C_TEXT).setFontWeight('bold').setFontSize(10)
        .setHorizontalAlignment('right').setNumberFormat('#,##0');
      sheet.getRange(rn, 6)
        .setBackground(C_CAT_BG).setFontColor(C_TEXT).setFontWeight('bold').setFontSize(10)
        .setHorizontalAlignment('right').setNumberFormat('#,##0');
      sheet.getRange(rn, 1).setHorizontalAlignment('left');
      sheet.getRange(rn, 2).setHorizontalAlignment('right').setNumberFormat('#,##0');
      sheet.getRange(rn, 3).setHorizontalAlignment('left');
      sheet.getRange(rn, 4).setHorizontalAlignment('right').setNumberFormat('#,##0');
      itemIdx = 0;
    } else {
      // item row
      var bg = (itemIdx % 2 === 0) ? C_ODD : C_EVEN;
      sheet.getRange(rn, 1, 1, 4)
        .setBackground(bg).setFontColor(C_TEXT).setFontWeight('normal').setFontSize(10);
      sheet.getRange(rn, 5)
        .setBackground(C_ACT_ITEM).setFontColor(C_TEXT).setFontWeight('normal').setFontSize(10)
        .setHorizontalAlignment('right').setNumberFormat('#,##0');
      sheet.getRange(rn, 6)
        .setBackground(bg).setFontColor(C_TEXT).setFontWeight('normal').setFontSize(10)
        .setHorizontalAlignment('right').setNumberFormat('#,##0');
      sheet.getRange(rn, 1).setHorizontalAlignment('left');
      sheet.getRange(rn, 2).setHorizontalAlignment('right').setNumberFormat('#,##0');
      sheet.getRange(rn, 3).setHorizontalAlignment('left');
      sheet.getRange(rn, 4).setHorizontalAlignment('right').setNumberFormat('#,##0');
      itemIdx++;
    }
  }

  // ── 12. เส้นกรอบ ──────────────────────────────────────────────────────────
  sheet.getRange(4, 1, newRows.length + 1, 6)
    .setBorder(true, true, true, true, true, true,
               C_BORDER, SpreadsheetApp.BorderStyle.SOLID);

  // ── 13. แถวสรุปรวม ────────────────────────────────────────────────────────
  var sumRow = newRows.length + 5;
  sheet.setRowHeight(sumRow, 28);
  var sumRange = sheet.getRange(sumRow, 1, 1, 6);
  sumRange.setBackground('#7a3f4a').setFontColor('#ffffff').setFontWeight('bold').setFontSize(11);

  // สูตรรวมทั้งหมด (ใช้ SUMPRODUCT กรองเฉพาะแถว category header ที่ C="")
  var dataRange5  = 5;
  var dataRangeN  = newRows.length + 4;
  var fTotalBudget = totalBudget;  // ค่าตรง (เร็วกว่า)
  var fTotalActual = '=IFERROR(SUM(' + SN + '!D4:D10000),0)';
  var fTotalRemain = '=' + fTotalBudget + '-' + 'F' + sumRow;

  sheet.getRange(sumRow, 1).setValue('💰  รวมทั้งหมด').setHorizontalAlignment('left');
  sheet.getRange(sumRow, 2).setValue(fTotalBudget).setHorizontalAlignment('right').setNumberFormat('#,##0');
  sheet.getRange(sumRow, 3).setValue('');
  sheet.getRange(sumRow, 4).setValue('');
  sheet.getRange(sumRow, 5).setFormula(fTotalActual).setHorizontalAlignment('right').setNumberFormat('#,##0');
  sheet.getRange(sumRow, 6).setFormula('=B' + sumRow + '-E' + sumRow)
    .setHorizontalAlignment('right').setNumberFormat('#,##0');

  return {
    ok:  true,
    msg: '✅ จัดรูปแบบ sheet งบประมาณ เสร็จสมบูรณ์ (' + newRows.length + ' แถวข้อมูล + SUMIF จ่ายจริง)'
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// FORMAT SUMMARY SHEET  (สรุปภาพรวม)
// ─────────────────────────────────────────────────────────────────────────────
function formatSummarySheet(ss) {
  var budgetSheet = ss.getSheetByName('งบประมาณ');
  var expSheet    = ss.getSheetByName('บันทึกค่าใช้จ่าย');

  if (!budgetSheet) return { ok: false, msg: 'ไม่พบ sheet งบประมาณ — รัน format_budget ก่อน' };
  if (!expSheet)    return { ok: false, msg: 'ไม่พบ sheet บันทึกค่าใช้จ่าย' };

  // อ่าน budget categories
  var dataStart = budgetDataStart(budgetSheet);
  var lastRow   = budgetSheet.getLastRow();
  if (lastRow < dataStart) return { ok: false, msg: 'sheet งบประมาณว่างเปล่า — sync ข้อมูลก่อน' };

  var rawRows = budgetSheet.getRange(dataStart, 1, lastRow - dataStart + 1, 4).getValues();
  var cats    = parseBudgetRows(rawRows);
  if (cats.length === 0) return { ok: false, msg: 'ไม่มีข้อมูลงบประมาณ' };

  // ── สร้าง / ล้าง sheet สรุปภาพรวม ────────────────────────────────────────
  var sheet = ss.getSheetByName('สรุปภาพรวม');
  if (!sheet) {
    sheet = ss.insertSheet('สรุปภาพรวม');
  } else {
    sheet.clear();
    sheet.clearFormats();
  }

  // ── column widths ─────────────────────────────────────────────────────────
  sheet.setColumnWidth(1, 215);   // A หมวดหมู่
  sheet.setColumnWidth(2, 145);   // B งบประมาณ
  sheet.setColumnWidth(3, 145);   // C จ่ายจริง
  sheet.setColumnWidth(4, 145);   // D คงเหลือ
  sheet.setColumnWidth(5, 100);   // E % ใช้ไป
  sheet.setColumnWidth(6, 110);   // F สถานะ

  // ── Row 1: Title ──────────────────────────────────────────────────────────
  sheet.setRowHeight(1, 48);
  sheet.getRange(1, 1, 1, 6).merge()
    .setValue('📊  สรุปภาพรวมงบประมาณ  ·  Family Witness Dinner')
    .setBackground('#7a3f4a').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(14)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  // ── Row 2: Event info ─────────────────────────────────────────────────────
  sheet.setRowHeight(2, 32);
  sheet.getRange(2, 1, 1, 6).merge()
    .setValue('📅  28 พฤศจิกายน 2569  ·  TARA Terrace  ·  14:00 น.')
    .setBackground('#b5838d').setFontColor('#ffffff')
    .setFontSize(11).setHorizontalAlignment('center').setVerticalAlignment('middle');

  // ── Row 3: เส้นคั่น ───────────────────────────────────────────────────────
  sheet.setRowHeight(3, 6);
  sheet.getRange(3, 1, 1, 6).merge().setBackground('#e8c4c4');

  // ── Rows 4-7: กล่องสรุปภาพรวม ─────────────────────────────────────────────
  var SN_B = "'งบประมาณ'";
  var SN_E = "'บันทึกค่าใช้จ่าย'";
  var totalBudget = 0;
  for (var ti = 0; ti < cats.length; ti++) totalBudget += cats[ti].budget;

  sheet.setRowHeight(4, 26);
  sheet.getRange(4, 1, 1, 6).merge()
    .setValue('💡 ภาพรวมทั้งหมด')
    .setBackground('#9a6472').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(11).setHorizontalAlignment('left');

  // Row 5: Labels
  sheet.setRowHeight(5, 22);
  sheet.getRange(5, 1, 1, 3)
    .setValues([['งบประมาณรวม (บาท)', 'จ่ายจริงรวม (บาท)', 'คงเหลือ (บาท)']])
    .setBackground('#f4d9d0').setFontColor('#3d2c2c')
    .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');
  sheet.getRange(5, 4, 1, 3)
    .setValues([['% ใช้งบแล้ว', 'รายการค่าใช้จ่าย', 'อัปเดตล่าสุด']])
    .setBackground('#f4d9d0').setFontColor('#3d2c2c')
    .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');

  // Row 6: Values with formulas
  sheet.setRowHeight(6, 28);
  sheet.getRange(6, 1).setValue(totalBudget)
    .setBackground('#fdf8f6').setFontColor('#3d2c2c').setFontWeight('bold').setFontSize(13)
    .setHorizontalAlignment('center').setNumberFormat('#,##0');
  sheet.getRange(6, 2).setFormula('=IFERROR(SUM(' + SN_E + '!D4:D10000),0)')
    .setBackground('#fef3d0').setFontColor('#3d2c2c').setFontWeight('bold').setFontSize(13)
    .setHorizontalAlignment('center').setNumberFormat('#,##0');
  sheet.getRange(6, 3).setFormula('=A6-B6')
    .setBackground('#fdf8f6').setFontColor('#3d2c2c').setFontWeight('bold').setFontSize(13)
    .setHorizontalAlignment('center').setNumberFormat('#,##0');
  sheet.getRange(6, 4).setFormula('=IFERROR(B6/A6,0)')
    .setBackground('#fef3d0').setFontColor('#3d2c2c').setFontWeight('bold').setFontSize(13)
    .setHorizontalAlignment('center').setNumberFormat('0.0%');
  sheet.getRange(6, 5).setFormula('=IFERROR(COUNTA(' + SN_E + '!C4:C10000),0)')
    .setBackground('#fdf8f6').setFontColor('#3d2c2c').setFontWeight('bold').setFontSize(13)
    .setHorizontalAlignment('center');
  sheet.getRange(6, 6).setValue(Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm'))
    .setBackground('#fef3d0').setFontColor('#7a5c5c').setFontSize(10)
    .setHorizontalAlignment('center');

  // Row 7: เส้นคั่น
  sheet.setRowHeight(7, 10);
  sheet.getRange(7, 1, 1, 6).merge().setBackground('#e8c4c4');

  // ── Row 8: หัวตาราง category ─────────────────────────────────────────────
  sheet.setRowHeight(8, 28);
  sheet.getRange(8, 1, 1, 6)
    .setValues([['หมวดหมู่', 'งบประมาณ', 'จ่ายจริง', 'คงเหลือ', '% ใช้ไป', 'สถานะ']])
    .setBackground('#c9a84c').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center');

  sheet.setFrozenRows(8);

  // ── Row 9+: ข้อมูลแต่ละหมวดหมู่ ──────────────────────────────────────────
  var C_BORDER = '#d4a0a8';
  var C_ODD    = '#ffffff';
  var C_EVEN   = '#fdf8f6';

  for (var ci = 0; ci < cats.length; ci++) {
    var c  = cats[ci];
    var rn = ci + 9;
    sheet.setRowHeight(rn, 24);

    var bg = (ci % 2 === 0) ? C_ODD : C_EVEN;
    sheet.getRange(rn, 1, 1, 6)
      .setBackground(bg).setFontColor('#3d2c2c').setFontSize(10);

    // A: ชื่อหมวด
    sheet.getRange(rn, 1).setValue(c.cat)
      .setFontWeight('bold').setHorizontalAlignment('left');

    // B: งบประมาณ (ค่าตรงจาก cats ที่อ่านมาแล้ว)
    sheet.getRange(rn, 2).setValue(c.budget)
      .setHorizontalAlignment('right').setNumberFormat('#,##0');

    // C: จ่ายจริง → SUMIF จาก บันทึกค่าใช้จ่าย
    var fAct = '=IFERROR(SUMIFS(' + SN_E + '!$D:$D,' + SN_E + '!$B:$B,"' + c.cat + '"),0)';
    sheet.getRange(rn, 3).setFormula(fAct)
      .setHorizontalAlignment('right').setNumberFormat('#,##0');

    // D: คงเหลือ = B - C
    sheet.getRange(rn, 4).setFormula('=B' + rn + '-C' + rn)
      .setHorizontalAlignment('right').setNumberFormat('#,##0');

    // E: % ใช้ไป = C / B
    sheet.getRange(rn, 5).setFormula('=IFERROR(C' + rn + '/B' + rn + ',0)')
      .setHorizontalAlignment('center').setNumberFormat('0.0%');

    // F: สถานะ → ข้อความตามเปอร์เซ็นต์
    sheet.getRange(rn, 6)
      .setFormula('=IF(E' + rn + '>1,"🔴 เกินงบ",IF(E' + rn + '>0.9,"⚠️ ใกล้เต็ม","✅ ปกติ"))')
      .setHorizontalAlignment('center').setFontSize(10);
  }

  // ── แถวรวม ────────────────────────────────────────────────────────────────
  var totalRow = cats.length + 9;
  sheet.setRowHeight(totalRow, 26);
  sheet.getRange(totalRow, 1, 1, 6)
    .setBackground('#9a6472').setFontColor('#ffffff').setFontWeight('bold').setFontSize(11);
  sheet.getRange(totalRow, 1).setValue('💰  รวมทั้งหมด').setHorizontalAlignment('left');
  sheet.getRange(totalRow, 2).setValue(totalBudget).setHorizontalAlignment('right').setNumberFormat('#,##0');
  sheet.getRange(totalRow, 3)
    .setFormula('=SUM(C9:C' + (totalRow - 1) + ')')
    .setHorizontalAlignment('right').setNumberFormat('#,##0');
  sheet.getRange(totalRow, 4)
    .setFormula('=B' + totalRow + '-C' + totalRow)
    .setHorizontalAlignment('right').setNumberFormat('#,##0');
  sheet.getRange(totalRow, 5)
    .setFormula('=IFERROR(C' + totalRow + '/B' + totalRow + ',0)')
    .setHorizontalAlignment('center').setNumberFormat('0.0%');
  sheet.getRange(totalRow, 6).setValue('');

  // ── เส้นกรอบ ──────────────────────────────────────────────────────────────
  sheet.getRange(8, 1, cats.length + 2, 6)
    .setBorder(true, true, true, true, true, true,
               C_BORDER, SpreadsheetApp.BorderStyle.SOLID);

  // ── กล่อง summary row 5-6 border ─────────────────────────────────────────
  sheet.getRange(5, 1, 2, 6)
    .setBorder(true, true, true, true, true, false,
               C_BORDER, SpreadsheetApp.BorderStyle.SOLID);

  return {
    ok:  true,
    msg: '✅ จัดรูปแบบ sheet สรุปภาพรวม เสร็จสมบูรณ์ (' + cats.length + ' หมวดหมู่)'
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// UPDATE "Summary" AND "Dashboard" SHEETS
// ทั้งสองใช้ layout เดียวกัน: KPI รวม (row 5-6) + ตารางแยกหมวด (row 9+)
// สูตร SUMIF ดึงตรงจาก 'บันทึกค่าใช้จ่าย' — ไม่ขึ้นกับโครงสร้าง งบประมาณ
// ─────────────────────────────────────────────────────────────────────────────
function updateSummarySheetGS(ss) {
  return rebuildOverviewSheet(ss, 'Summary',
    '📊  Summary – งบประมาณ vs ค่าใช้จ่ายจริง  ·  Family Witness Dinner');
}

function updateDashboardSheetGS(ss) {
  return rebuildOverviewSheet(ss, 'Dashboard',
    '📊  Dashboard – ภาพรวมค่าใช้จ่าย  ·  Family Witness Dinner');
}

function rebuildOverviewSheet(ss, sheetName, title) {
  var expSheet = ss.getSheetByName('บันทึกค่าใช้จ่าย');
  if (!expSheet) return { ok: false, msg: 'ไม่พบ sheet บันทึกค่าใช้จ่าย' };

  var cats = getBudget(ss);
  if (!cats.length) return { ok: false, msg: 'ไม่มีข้อมูลงบประมาณ — sync ข้อมูลก่อน' };

  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
    sheet.clearFormats();
    try { sheet.getRange(1, 1, 1, 6).breakApart(); } catch(ex) {}
  }

  var SN_E = "'บันทึกค่าใช้จ่าย'";
  var totalBudget = 0;
  for (var ti = 0; ti < cats.length; ti++) totalBudget += cats[ti].budget;

  // ── คอลัมน์ ─────────────────────────────────────────────────────────────
  sheet.setColumnWidth(1, 215);  // หมวดหมู่
  sheet.setColumnWidth(2, 145);  // งบประมาณ
  sheet.setColumnWidth(3, 145);  // จ่ายจริง
  sheet.setColumnWidth(4, 145);  // คงเหลือ
  sheet.setColumnWidth(5, 100);  // %
  sheet.setColumnWidth(6, 120);  // สถานะ

  // ── Row 1: Title ──────────────────────────────────────────────────────────
  sheet.setRowHeight(1, 48);
  sheet.getRange(1, 1, 1, 6).merge()
    .setValue(title)
    .setBackground('#7a3f4a').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(14)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  // ── Row 2: Event info ─────────────────────────────────────────────────────
  sheet.setRowHeight(2, 32);
  sheet.getRange(2, 1, 1, 6).merge()
    .setValue('📅  28 พฤศจิกายน 2569  ·  TARA Terrace  ·  14:00 น.  ·  งบรวม ' +
              totalBudget.toLocaleString() + ' บาท')
    .setBackground('#b5838d').setFontColor('#ffffff')
    .setFontSize(11).setHorizontalAlignment('center').setVerticalAlignment('middle');

  // ── Row 3: Separator ──────────────────────────────────────────────────────
  sheet.setRowHeight(3, 6);
  sheet.getRange(3, 1, 1, 6).merge().setBackground('#e8c4c4');

  // ── Row 4: KPI section label ──────────────────────────────────────────────
  sheet.setRowHeight(4, 26);
  sheet.getRange(4, 1, 1, 6).merge()
    .setValue('💡 ภาพรวมทั้งหมด')
    .setBackground('#9a6472').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(11).setHorizontalAlignment('left');

  // ── Row 5: KPI labels ─────────────────────────────────────────────────────
  sheet.setRowHeight(5, 22);
  sheet.getRange(5, 1, 1, 6)
    .setValues([['งบประมาณรวม (บาท)', 'จ่ายจริงรวม (บาท)', 'คงเหลือ (บาท)',
                 '% ใช้งบแล้ว', 'จำนวนรายการ', 'อัปเดต']])
    .setBackground('#f4d9d0').setFontColor('#3d2c2c')
    .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');

  // ── Row 6: KPI values ─────────────────────────────────────────────────────
  sheet.setRowHeight(6, 30);
  sheet.getRange(6, 1).setValue(totalBudget)
    .setBackground('#fdf8f6').setFontColor('#7a3f4a').setFontWeight('bold').setFontSize(15)
    .setHorizontalAlignment('center').setNumberFormat('#,##0');
  sheet.getRange(6, 2).setFormula('=IFERROR(SUM(' + SN_E + '!D4:D10000),0)')
    .setBackground('#fef3d0').setFontColor('#9a6472').setFontWeight('bold').setFontSize(15)
    .setHorizontalAlignment('center').setNumberFormat('#,##0');
  sheet.getRange(6, 3).setFormula('=A6-B6')
    .setBackground('#fdf8f6').setFontColor('#5a9e6f').setFontWeight('bold').setFontSize(15)
    .setHorizontalAlignment('center').setNumberFormat('#,##0');
  sheet.getRange(6, 4).setFormula('=IFERROR(B6/A6,0)')
    .setBackground('#fef3d0').setFontColor('#c9a84c').setFontWeight('bold').setFontSize(15)
    .setHorizontalAlignment('center').setNumberFormat('0.0%');
  sheet.getRange(6, 5).setFormula('=IFERROR(COUNTA(' + SN_E + '!C4:C10000),0)')
    .setBackground('#fdf8f6').setFontColor('#3d2c2c').setFontWeight('bold').setFontSize(13)
    .setHorizontalAlignment('center');
  sheet.getRange(6, 6)
    .setValue(Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm'))
    .setBackground('#fef3d0').setFontColor('#7a5c5c').setFontSize(10)
    .setHorizontalAlignment('center');

  // ── Row 7: Separator ──────────────────────────────────────────────────────
  sheet.setRowHeight(7, 10);
  sheet.getRange(7, 1, 1, 6).merge().setBackground('#e8c4c4');

  // ── Row 8: ตารางหัวคอลัมน์ ────────────────────────────────────────────────
  sheet.setRowHeight(8, 28);
  sheet.getRange(8, 1, 1, 6)
    .setValues([['หมวดหมู่', 'งบประมาณ', 'จ่ายจริง', 'คงเหลือ', '% ใช้ไป', 'สถานะ']])
    .setBackground('#c9a84c').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center');
  sheet.setFrozenRows(8);

  // ── Row 9+: ข้อมูลแต่ละหมวดหมู่ ──────────────────────────────────────────
  var C_ODD    = '#ffffff';
  var C_EVEN   = '#fdf8f6';
  var C_BORDER = '#d4a0a8';

  for (var ci = 0; ci < cats.length; ci++) {
    var c  = cats[ci];
    var rn = ci + 9;
    sheet.setRowHeight(rn, 24);

    var bg = (ci % 2 === 0) ? C_ODD : C_EVEN;
    sheet.getRange(rn, 1, 1, 6).setBackground(bg).setFontColor('#3d2c2c').setFontSize(10);

    sheet.getRange(rn, 1).setValue(c.cat).setFontWeight('bold').setHorizontalAlignment('left');
    sheet.getRange(rn, 2).setValue(c.budget).setHorizontalAlignment('right').setNumberFormat('#,##0');

    // จ่ายจริง: SUMIF จาก บันทึกค่าใช้จ่าย ตรงๆ
    var fAct = '=IFERROR(SUMIFS(' + SN_E + '!$D:$D,' + SN_E + '!$B:$B,"' + c.cat + '"),0)';
    sheet.getRange(rn, 3).setFormula(fAct).setHorizontalAlignment('right').setNumberFormat('#,##0');

    // คงเหลือ = งบ - จ่ายจริง
    sheet.getRange(rn, 4).setFormula('=B' + rn + '-C' + rn)
      .setHorizontalAlignment('right').setNumberFormat('#,##0');

    // % ใช้ไป
    sheet.getRange(rn, 5).setFormula('=IFERROR(C' + rn + '/B' + rn + ',0)')
      .setHorizontalAlignment('center').setNumberFormat('0.0%');

    // สถานะ
    sheet.getRange(rn, 6)
      .setFormula('=IF(E' + rn + '>1,"🔴 เกินงบ",IF(E' + rn + '>0.9,"⚠️ ใกล้เต็ม","✅ ปกติ"))')
      .setHorizontalAlignment('center').setFontSize(10);
  }

  // ── แถวรวม ────────────────────────────────────────────────────────────────
  var totalRow = cats.length + 9;
  sheet.setRowHeight(totalRow, 26);
  sheet.getRange(totalRow, 1, 1, 6)
    .setBackground('#9a6472').setFontColor('#ffffff').setFontWeight('bold').setFontSize(11);
  sheet.getRange(totalRow, 1).setValue('💰  รวมทั้งหมด').setHorizontalAlignment('left');
  sheet.getRange(totalRow, 2).setValue(totalBudget)
    .setHorizontalAlignment('right').setNumberFormat('#,##0');
  sheet.getRange(totalRow, 3)
    .setFormula('=SUM(C9:C' + (totalRow - 1) + ')')
    .setHorizontalAlignment('right').setNumberFormat('#,##0');
  sheet.getRange(totalRow, 4)
    .setFormula('=B' + totalRow + '-C' + totalRow)
    .setHorizontalAlignment('right').setNumberFormat('#,##0');
  sheet.getRange(totalRow, 5)
    .setFormula('=IFERROR(C' + totalRow + '/B' + totalRow + ',0)')
    .setHorizontalAlignment('center').setNumberFormat('0.0%');
  sheet.getRange(totalRow, 6).setValue('');

  // ── เส้นกรอบ ──────────────────────────────────────────────────────────────
  sheet.getRange(8, 1, cats.length + 2, 6)
    .setBorder(true, true, true, true, true, true,
               C_BORDER, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(5, 1, 2, 6)
    .setBorder(true, true, true, true, true, false,
               C_BORDER, SpreadsheetApp.BorderStyle.SOLID);

  return {
    ok:  true,
    msg: '✅ อัปเดต sheet "' + sheetName + '" เสร็จ (' + cats.length +
         ' หมวดหมู่ + SUMIF จ่ายจริง/คงเหลือ)'
  };
}

// ── Trigger จาก Apps Script Editor ───────────────────────────────────────────
function doFormatBudget() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var r  = formatBudgetSheet(ss);
  Logger.log(r.msg);
  SpreadsheetApp.getUi().alert(r.msg);
}

function doFormatSummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var r  = formatSummarySheet(ss);
  Logger.log(r.msg);
  SpreadsheetApp.getUi().alert(r.msg);
}

function doFormatAll() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rb = formatBudgetSheet(ss);
  var rs = formatSummarySheet(ss);
  Logger.log(rb.msg);
  Logger.log(rs.msg);
  SpreadsheetApp.getUi().alert(rb.msg + '\n' + rs.msg);
}

function doUpdateSummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var r  = fixDashboardSummaryFormulas(ss);
  Logger.log(r.msg);
  SpreadsheetApp.getUi().alert(r.msg);
}

function doUpdateDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var r  = fixDashboardSummaryFormulas(ss);
  Logger.log(r.msg);
  SpreadsheetApp.getUi().alert(r.msg);
}

function doUpdateSummaryDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var r  = fixDashboardSummaryFormulas(ss);
  Logger.log(r.msg);
  SpreadsheetApp.getUi().alert(r.msg);
}

// ─────────────────────────────────────────────────────────────────────────────
// FIX FORMULAS ใน "Dashboard" และ "Summary" — ไม่ rebuild, scan & patch เท่านั้น
// แก้: งบประมาณ (SUMIF จาก งบประมาณ!B), จ่ายจริง (SUMIF จาก บันทึกค่าใช้จ่าย!D),
//       คงเหลือ (งบ-จ่าย), % (จ่าย/งบ), สถานะ (IF)
// ─────────────────────────────────────────────────────────────────────────────
function fixDashboardSummaryFormulas(ss) {
  var SN_B = "'งบประมาณ'";
  var SN_E = "'บันทึกค่าใช้จ่าย'";
  var cats = [
    'สถานที่จัดงานและอาหาร','ชุดและความงาม','ผู้แลงาน',
    'งานภาพและวีดีโอ','การ์ดและของขวัญ','ความบันเทิง'
  ];

  function colLetter(n) {
    var s = '';
    while (n > 0) { s = String.fromCharCode(64+(n%26||26))+s; n=Math.floor((n-1)/26); }
    return s;
  }

  // งบรวม = SUMIF หมวดหมู่จาก งบประมาณ!B (เฉพาะ category header row ที่ C='')
  var fTotalBudget = '=IFERROR(' + cats.map(function(c) {
    return 'SUMIF('+SN_B+'!$A:$A,"'+c+'",'+SN_B+'!$B:$B)';
  }).join('+') + ',0)';

  // จ่ายจริงรวม = SUM จาก บันทึกค่าใช้จ่าย
  var fTotalActual = '=IFERROR(SUM('+SN_E+'!D4:D10000),0)';

  var results = [];

  ['Dashboard','Summary'].forEach(function(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) { results.push(sheetName+': ไม่พบ sheet'); return; }

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (!lastRow || !lastCol) { results.push(sheetName+': ว่างเปล่า'); return; }

    var data = sheet.getRange(1,1,lastRow,lastCol).getValues();

    // ── A: แก้ KPI section (Dashboard: งบประมาณรวม / จ่ายไปแล้ว / คงเหลือ / %) ──
    if (sheetName === 'Dashboard') {
      var kpi = {};
      for (var r=0; r<data.length; r++) {
        for (var c=0; c<data[r].length; c++) {
          var v = String(data[r][c]).trim();
          if (v.indexOf('งบประมาณรวม') >= 0)  kpi.budget = {r:r,c:c};
          if (v.indexOf('จ่ายไปแล้ว')  >= 0)  kpi.actual = {r:r,c:c};
          if (v.indexOf('คงเหลือ')     >= 0 && v.indexOf('(บาท)') >= 0) kpi.remain = {r:r,c:c};
          if (v.indexOf('% งบที่ใช้ไป') >= 0) kpi.pct    = {r:r,c:c};
        }
      }
      if (kpi.budget) {
        var vCol   = kpi.budget.c + 2;   // 1-based column ของ value cell
        var vLet   = colLetter(vCol);
        var budR   = kpi.budget.r + 1;
        sheet.getRange(budR, vCol).setFormula(fTotalBudget).setNumberFormat('#,##0');
        if (kpi.actual) {
          var actR = kpi.actual.r + 1;
          sheet.getRange(actR, vCol).setFormula(fTotalActual).setNumberFormat('#,##0');
          if (kpi.remain) {
            var remR = kpi.remain.r + 1;
            sheet.getRange(remR, vCol)
              .setFormula('='+vLet+budR+'-'+vLet+actR).setNumberFormat('#,##0');
          }
          if (kpi.pct) {
            var pctR = kpi.pct.r + 1;
            sheet.getRange(pctR, vCol)
              .setFormula('=IFERROR('+vLet+actR+'/'+vLet+budR+',0)').setNumberFormat('0.0%');
          }
        }
      }
    }

    // ── B: หา header row ของตารางงบประมาณ (มี หมวดหมู่/งบประมาณ/จ่ายจริง/คงเหลือ) ──
    var headerRow=-1, ci=-1, bi=-1, ai=-1, ri=-1, pi=-1, si=-1;
    for (var r2=0; r2<data.length; r2++) {
      var tc=-1,tb=-1,ta=-1,tr=-1,tp=-1,ts=-1;
      for (var c2=0; c2<data[r2].length; c2++) {
        var v2 = String(data[r2][c2]).trim();
        if (v2==='หมวดหมู่')                        tc=c2;
        if (v2==='งบประมาณ')                        tb=c2;
        if (v2==='จ่ายจริง')                        ta=c2;
        if (v2==='คงเหลือ'||v2==='ส่วนต่าง')       tr=c2;
        if (v2.indexOf('%')>=0 && v2.length<=15)    tp=c2;
        if (v2==='สถานะ')                            ts=c2;
      }
      if (tc>=0 && tb>=0 && ta>=0 && tr>=0) {
        headerRow=r2; ci=tc; bi=tb; ai=ta; ri=tr; pi=tp; si=ts; break;
      }
    }
    if (headerRow<0) { results.push(sheetName+': ไม่พบหัวตาราง'); return; }

    var budL=colLetter(bi+1), actL=colLetter(ai+1), remL=colLetter(ri+1);
    var pctL=pi>=0?colLetter(pi+1):null, statL=si>=0?colLetter(si+1):null;
    var catL=colLetter(ci+1);
    var updated=0;

    // ── C: วนแต่ละแถวข้อมูลใต้ header ──────────────────────────────────────
    for (var r3=headerRow+1; r3<data.length; r3++) {
      var catVal = String(data[r3][ci]).trim();
      var isTotal = (catVal==='รวมทั้งหมด');
      var isCat   = (cats.indexOf(catVal)>=0);
      if (!isTotal && !isCat) continue;

      var rn = r3+1;
      var catRef = catL+rn;

      // งบประมาณ
      sheet.getRange(rn, bi+1)
        .setFormula(isTotal ? fTotalBudget :
          '=IFERROR(SUMIF('+SN_B+'!$A:$A,'+catRef+','+SN_B+'!$B:$B),0)')
        .setNumberFormat('#,##0');

      // จ่ายจริง
      sheet.getRange(rn, ai+1)
        .setFormula(isTotal ? fTotalActual :
          '=IFERROR(SUMIF('+SN_E+'!$B:$B,'+catRef+','+SN_E+'!$D:$D),0)')
        .setNumberFormat('#,##0');

      // คงเหลือ
      sheet.getRange(rn, ri+1)
        .setFormula('='+budL+rn+'-'+actL+rn)
        .setNumberFormat('#,##0');

      // %
      if (pi>=0) {
        sheet.getRange(rn, pi+1)
          .setFormula('=IFERROR('+actL+rn+'/'+budL+rn+',0)')
          .setNumberFormat('0.0%');
      }

      // สถานะ
      if (si>=0 && pi>=0) {
        sheet.getRange(rn, si+1)
          .setFormula('=IF('+pctL+rn+'>1,"🔴 เกินงบ",IF('+pctL+rn+'>0.9,"⚠ ใกล้เต็ม","✅ ปกติ"))');
      }

      updated++;
    }

    results.push(sheetName+': อัปเดตสูตร '+updated+' แถว ✅');
  });

  return { ok:true, msg: results.join('\n') };
}

// --- GUESTS ------------------------------------------------------------------
function getGuests(ss) {
  var sheet = ss.getSheetByName('แขก');
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var data   = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  var guests = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    guests.push({
      id:       i + 1,
      name:     String(row[0] || ''),
      side:     String(row[1] || 'เอ็ม'),
      phone:    String(row[2] || ''),
      relation: String(row[3] || ''),
      seats:    parseInt(row[4]) || 1,
      status:   String(row[5] || 'pending'),
      food:     String(row[6] || ''),
      note:     String(row[7] || '')
    });
  }
  return guests;
}

function saveGuests(ss, guests) {
  var sheet = ss.getSheetByName('แขก');
  if (!sheet) {
    sheet = ss.insertSheet('แขก');
    sheet.getRange(1, 1, 1, 8).setValues([[
      'ชื่อ', 'ฝั่ง', 'เบอร์โทร', 'ความสัมพันธ์',
      'ที่นั่ง', 'สถานะ', 'อาหาร/แพ้อาหาร', 'หมายเหตุ'
    ]]);
    sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  var lastRow = sheet.getLastRow();
  var lastCol = Math.max(sheet.getLastColumn(), 11);
  if (lastRow >= 2) sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  if (!guests || !guests.length) return;

  var rows = [];
  for (var i = 0; i < guests.length; i++) {
    var g = guests[i];
    rows.push([g.name, g.side, g.phone, g.relation, g.seats, g.status, g.food, g.note]);
  }
  sheet.getRange(2, 1, rows.length, 8).setValues(rows);
}

// --- INVITE LINKS -------------------------------------------------------
var LINK_HEADERS = ['slug', 'ชื่อกลุ่ม', 'สถานะ', 'วันที่สร้าง'];

function getLinks(ss) {
  var sheet = ss.getSheetByName('ลิงก์เชิญ');
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  var links = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    links.push({
      slug:      String(row[0] || '').trim(),
      name:      String(row[1] || '').trim(),
      active:    String(row[2] || '').trim() === 'open',
      createdAt: String(row[3] || '')
    });
  }
  return links;
}

function saveLinks(ss, links) {
  var sheet = ss.getSheetByName('ลิงก์เชิญ');
  if (!sheet) {
    sheet = ss.insertSheet('ลิงก์เชิญ');
    sheet.getRange(1, 1, 1, LINK_HEADERS.length)
      .setValues([LINK_HEADERS])
      .setFontWeight('bold')
      .setBackground('#f4d9d0');
    sheet.setFrozenRows(1);
    [160, 200, 80, 180].forEach(function(w, i) { sheet.setColumnWidth(i + 1, w); });
  }
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) sheet.getRange(2, 1, lastRow - 1, 4).clearContent();
  if (!links || !links.length) return;
  var rows = links.map(function(lk) {
    return [
      lk.slug || '',
      lk.name || '',
      lk.active ? 'open' : 'closed',
      lk.createdAt || ''
    ];
  });
  sheet.getRange(2, 1, rows.length, 4).setValues(rows);
}

// ตรวจสอบ slug เดียว — ใช้จาก invite.html ผ่าน ?action=check_link&slug=xxx
function checkLink(ss, slug) {
  if (!slug) return { status: 'ok', active: false, reason: 'no_slug' };
  var links = getLinks(ss);
  var found = null;
  for (var i = 0; i < links.length; i++) {
    if (links[i].slug === slug) { found = links[i]; break; }
  }
  if (!found)        return { status: 'ok', active: false, reason: 'not_found' };
  if (!found.active) return { status: 'ok', active: false, reason: 'closed', name: found.name };
  return { status: 'ok', active: true, name: found.name };
}

// ============================================================
//  DONATION — รับสลิปโอนเงิน + OCR ด้วย Gemini
// ============================================================

var DONATE_HEADERS = ['วันที่', 'ชื่อผู้โอน', 'ยอด (บาท)', 'ลิงก์สลิป', 'หมายเหตุ'];

// ── อ่านยอดจากสลิปด้วย Gemini Vision ──────────────────────────
function readAmountFromSlip(base64, mimeType) {
  try {
    var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey || !base64 || base64.length < 100) return '';

    var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey;
    var prompt = 'Look at this Thai bank transfer slip image. Find the transfer amount (ยอดโอนเงิน). ' +
      'It is the large prominent number on the slip (NOT account number, NOT date, NOT reference number). ' +
      'Reply with ONLY the number in digits, no commas, no "บาท", no spaces. Example: 399';

    var payload = {
      contents: [{
        parts: [
          { inline_data: { mime_type: mimeType, data: base64 } },
          { text: prompt }
        ]
      }],
      generationConfig: { temperature: 0, maxOutputTokens: 32 }
    };

    var resp = UrlFetchApp.fetch(url, {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    Logger.log('Gemini HTTP status: ' + resp.getResponseCode());
    var json = JSON.parse(resp.getContentText());

    if (json.candidates && json.candidates[0] && json.candidates[0].content) {
      var text = json.candidates[0].content.parts[0].text.trim();
      Logger.log('Gemini slip OCR raw: [' + text + ']');
      var m = text.replace(/,/g, '').match(/\d+(?:\.\d+)?/);
      return m ? String(Math.round(parseFloat(m[0]))) : '';
    } else {
      Logger.log('Gemini error: ' + resp.getContentText().substring(0, 500));
    }
  } catch(e) {
    Logger.log('readAmountFromSlip error: ' + e);
  }
  return '';
}

// ── บันทึกการโอนเงิน ──────────────────────────────────────────
function saveDonation(ss, body) {
  // 1. อัปโหลดสลิปลง Drive
  var slipUrl = '';
  if (body.slipBase64 && body.slipBase64.length > 100) {
    var ts  = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd_HHmmss');
    var ext = (body.slipMime || '').indexOf('png') >= 0 ? 'png' : 'jpg';
    slipUrl = savePictureToDrive(body.slipBase64, 'slip_' + ts + '.' + ext);
  }

  // 2. OCR ยอด (ถ้ายังไม่มี)
  var amount = String(body.amount || '').trim();
  if (!amount && body.slipBase64 && body.slipBase64.length > 100) {
    amount = readAmountFromSlip(body.slipBase64, body.slipMime || 'image/jpeg');
  }

  // 3. บันทึกลง sheet "โอนเงิน"
  var sheet = ss.getSheetByName('โอนเงิน');
  if (!sheet) {
    sheet = ss.insertSheet('โอนเงิน');
    sheet.getRange(1, 1, 1, DONATE_HEADERS.length).setValues([DONATE_HEADERS])
      .setFontWeight('bold').setBackground('#c8e6d4');
    sheet.setFrozenRows(1);
    [160, 200, 120, 340, 200].forEach(function(w, i) { sheet.setColumnWidth(i + 1, w); });
  }

  var dateStr = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm');
  var lastRow = sheet.getLastRow();
  var amountNum = parseFloat(amount) || '';

  sheet.getRange(lastRow + 1, 1, 1, 5).setValues([[
    dateStr,
    String(body.name || ''),
    amountNum,
    slipUrl,
    ''
  ]]);

  // ตั้ง hyperlink สลิป
  if (slipUrl) {
    sheet.getRange(lastRow + 1, 4)
      .setFormula('=HYPERLINK("' + slipUrl + '","ดูสลิป")');
  }

  return { amount: String(amountNum || ''), slipUrl: slipUrl };
}

// ── ตั้ง Gemini API Key (รันครั้งเดียวจาก Script Editor) ──────
function setGeminiApiKey() {
  var key = 'AIzaSy_REPLACE_WITH_YOUR_KEY'; // ← ใส่ key จริงที่นี่
  PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', key);
  SpreadsheetApp.getUi().alert('✅ บันทึก Gemini API Key เรียบร้อยแล้ว');
}
