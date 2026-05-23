// ============================================================
//  Family Witness Dinner Plan – Google Apps Script Web App
//  วาง code นี้ทั้งหมดใน Apps Script แล้ว Deploy เป็น Web App
//  Execute as: Me  |  Who has access: Anyone
// ============================================================

// ─── ROUTER ──────────────────────────────────────────────────────────────────
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'all';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let result = { status: 'ok' };

  try {
    if (action === 'ping') {
      result.message = 'connected';
    } else if (action === 'all') {
      result.expenses  = getExpenses(ss);
      result.contacts  = getContacts(ss);
      result.checklist = getChecklist(ss);
      result.guests    = getGuests(ss);
    } else if (action === 'expenses')  result.expenses  = getExpenses(ss);
    else if (action === 'contacts')    result.contacts  = getContacts(ss);
    else if (action === 'checklist')   result.checklist = getChecklist(ss);
    else if (action === 'guests')      result.guests    = getGuests(ss);
    else result.message = 'unknown action';
  } catch(err) {
    result.status = 'error';
    result.error  = err.toString();
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let result = { status: 'ok' };

  try {
    const body = JSON.parse(e.postData.contents);
    const ss   = SpreadsheetApp.getActiveSpreadsheet();

    if (body.expenses  !== undefined) saveExpenses(ss, body.expenses);
    if (body.contacts  !== undefined) saveContacts(ss, body.contacts);
    if (body.checklist !== undefined) saveChecklist(ss, body.checklist);
    if (body.guests    !== undefined) saveGuests(ss, body.guests);
  } catch(err) {
    result.status = 'error';
    result.error  = err.toString();
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── EXPENSES ────────────────────────────────────────────────────────────────
// Sheet: บันทึกค่าใช้จ่าย
// Row 1: Title, Row 2: Subtitle, Row 3: Headers, Row 4+: Data
// Columns: A=วันที่  B=หมวดหมู่  C=รายการ  D=จำนวนเงิน  E=หมายเหตุ

function getExpenses(ss) {
  const sheet = ss.getSheetByName('บันทึกค่าใช้จ่าย');
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 4) return [];

  const data = sheet.getRange(4, 1, lastRow - 3, 5).getValues();
  const expenses = [];

  data.forEach(function(row, i) {
    if (!row[0] && !row[2]) return;
    var dateStr = '';
    if (row[0] instanceof Date) {
      dateStr = Utilities.formatDate(row[0], 'Asia/Bangkok', 'yyyy-MM-dd');
    } else if (row[0]) {
      dateStr = row[0].toString().substring(0, 10);
    }
    expenses.push({
      id:     i + 1,
      date:   dateStr,
      cat:    (row[1] || '').toString().trim(),
      item:   (row[2] || '').toString().trim(),
      amount: parseFloat(row[3]) || 0,
      note:   (row[4] || '').toString().trim()
    });
  });

  return expenses.filter(function(e) { return e.date || e.item; });
}

function saveExpenses(ss, expenses) {
  const sheet = ss.getSheetByName('บันทึกค่าใช้จ่าย');
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow >= 4) sheet.getRange(4, 1, lastRow - 3, 5).clearContent();
  if (!expenses.length) return;

  const rows = expenses.map(function(exp) {
    return [exp.date, exp.cat, exp.item, exp.amount, exp.note];
  });
  sheet.getRange(4, 1, rows.length, 5).setValues(rows);
}

// ─── CONTACTS ────────────────────────────────────────────────────────────────
// Sheet: Contact ทีมงาน
// Row 1: Headers, Row 2+: Data
// Columns: A=งาน  B=ชื่อ  C=เบอร์โทร  D=E-mail  E=Page/Social

function getContacts(ss) {
  const sheet = ss.getSheetByName('Contact ทีมงาน');
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  const contacts = [];

  data.forEach(function(row, i) {
    if (!row[0] && !row[1]) return;
    contacts.push({
      id:     i + 1,
      role:   (row[0] || '').toString().trim(),
      name:   (row[1] || '').toString().trim(),
      phone:  (row[2] || '').toString().trim(),
      email:  (row[3] || '').toString().trim(),
      social: (row[4] || '').toString().trim(),
      note:   ''
    });
  });

  return contacts;
}

function saveContacts(ss, contacts) {
  const sheet = ss.getSheetByName('Contact ทีมงาน');
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) sheet.getRange(2, 1, lastRow - 1, 5).clearContent();
  if (!contacts.length) return;

  const rows = contacts.map(function(c) {
    return [c.role, c.name, c.phone, c.email, c.social];
  });
  sheet.getRange(2, 1, rows.length, 5).setValues(rows);
}

// ─── CHECKLIST ───────────────────────────────────────────────────────────────
// Sheet: Check List
// Col A = สถานะ (✓ / ✗ / blank), Col B = รายการ
// Returns map: { "taskText": "done"|"skipped"|"none" }

function getChecklist(ss) {
  const sheet = ss.getSheetByName('Check List');
  if (!sheet) return {};
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(1, 1, lastRow, 2).getValues();

  var statusMap = {};
  data.forEach(function(row) {
    var status = (row[0] || '').toString().trim();
    var task   = (row[1] || '').toString().trim();
    if (!task) return;
    // Skip header rows and period headers
    if (task === 'สถานะ' || task.indexOf('เดือน') !== -1 || task.indexOf('สัปดาห์') !== -1) return;

    var mapped = 'none';
    if (status === '✓') mapped = 'done';
    else if (status === '✗') mapped = 'skipped';
    statusMap[task] = mapped;
  });

  return statusMap;
}

function saveChecklist(ss, checklist) {
  const sheet = ss.getSheetByName('Check List');
  if (!sheet) return;

  // Build task→status map from the app's checklist structure
  var statusMap = {};
  if (Array.isArray(checklist)) {
    checklist.forEach(function(period) {
      period.items.forEach(function(item) {
        statusMap[item.text] = item.status;
      });
    });
  } else if (typeof checklist === 'object') {
    statusMap = checklist;
  }

  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(1, 1, lastRow, 2).getValues();

  data.forEach(function(row, i) {
    var task = (row[1] || '').toString().trim();
    if (!task || task === 'สถานะ' || task.indexOf('เดือน') !== -1 || task.indexOf('สัปดาห์') !== -1) return;

    var appStatus = statusMap[task];
    if (appStatus === undefined) return;
    var symbol = '';
    if (appStatus === 'done')    symbol = '✓';
    if (appStatus === 'skipped') symbol = '✗';
    sheet.getRange(i + 1, 1).setValue(symbol);
  });
}

// ─── GUESTS ──────────────────────────────────────────────────────────────────
// Sheet: แขก (สร้างอัตโนมัติถ้ายังไม่มี)
// Row 1: Headers, Row 2+: Data
// Columns: A=ชื่อ  B=ฝั่ง  C=เบอร์โทร  D=ความสัมพันธ์  E=ที่นั่ง  F=สถานะ  G=อาหาร  H=หมายเหตุ

function getGuests(ss) {
  var sheet = ss.getSheetByName('แขก');
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  const guests = [];

  data.forEach(function(row, i) {
    if (!row[0]) return;
    guests.push({
      id:       i + 1,
      name:     (row[0] || '').toString(),
      side:     (row[1] || 'bride').toString(),
      phone:    (row[2] || '').toString(),
      relation: (row[3] || '').toString(),
      seats:    parseInt(row[4]) || 1,
      status:   (row[5] || 'pending').toString(),
      food:     (row[6] || '').toString(),
      note:     (row[7] || '').toString()
    });
  });

  return guests;
}

function saveGuests(ss, guests) {
  var sheet = ss.getSheetByName('แขก');
  if (!sheet) {
    sheet = ss.insertSheet('แขก');
    var headers = [['ชื่อ','ฝั่ง','เบอร์โทร','ความสัมพันธ์','ที่นั่ง','สถานะ','อาหาร/แพ้อาหาร','หมายเหตุ']];
    sheet.getRange(1, 1, 1, 8).setValues(headers);
    sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) sheet.getRange(2, 1, lastRow - 1, 8).clearContent();
  if (!guests.length) return;

  const rows = guests.map(function(g) {
    return [g.name, g.side, g.phone, g.relation, g.seats, g.status, g.food, g.note];
  });
  sheet.getRange(2, 1, rows.length, 8).setValues(rows);
}
