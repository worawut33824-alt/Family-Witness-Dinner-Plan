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
      result.budget    = getBudget(ss);
    } else if (action === 'expenses')  result.expenses  = getExpenses(ss);
    else if (action === 'contacts')    result.contacts  = getContacts(ss);
    else if (action === 'checklist')   result.checklist = getChecklist(ss);
    else if (action === 'guests')      result.guests    = getGuests(ss);
    else if (action === 'budget')      result.budget    = getBudget(ss);
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
    if (body.budget    !== undefined) saveBudget(ss, body.budget);
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
// Sheet: Check List  Col A = สถานะ (✓ / ✗ / blank), Col B = รายการ
// ใช้ row number แทน text matching เพื่อความแม่นยำ
// Row numbers ของแต่ละ item (ข้าม header row และ period row)
var CHECKLIST_ROWS = [3,4,5,6,7,8,9,10, 12,13,14,15,16,17, 19,20,21,22, 24,25,26,27, 29,30,31,32];
// ตรงกับ id: c1, c2, ... c26 ในแอพ

function getChecklist(ss) {
  var sheet = ss.getSheetByName('Check List');
  if (!sheet) return {};
  var result = {};
  CHECKLIST_ROWS.forEach(function(row, i) {
    var cell = sheet.getRange(row, 1).getValue().toString().trim();
    var status = 'none';
    if (cell === '✓') status = 'done';
    else if (cell === '✗') status = 'skipped';
    result['c' + (i + 1)] = status;
  });
  return result;
}

function saveChecklist(ss, checklist) {
  var sheet = ss.getSheetByName('Check List');
  if (!sheet) return;

  // Build id→status map { c1: 'done', c2: 'skipped', ... }
  var statusMap = {};
  if (Array.isArray(checklist)) {
    checklist.forEach(function(period) {
      period.items.forEach(function(item) {
        statusMap[item.id] = item.status;
      });
    });
  } else if (typeof checklist === 'object') {
    statusMap = checklist;
  }

  CHECKLIST_ROWS.forEach(function(row, i) {
    var id = 'c' + (i + 1);
    var appStatus = statusMap[id];
    if (appStatus === undefined) return;
    var symbol = '';
    if (appStatus === 'done')    symbol = '✓';
    if (appStatus === 'skipped') symbol = '✗';
    sheet.getRange(row, 1).setValue(symbol);
  });
}

// ─── BUDGET ──────────────────────────────────────────────────────────────────
// Sheet: งบประมาณ  (สร้างอัตโนมัติถ้ายังไม่มี)
// Row 1: Headers, Row 2+: Data
// Columns: A=หมวดหมู่  B=งบรวมหมวด  C=รายการ  D=งบรายการ

function getBudget(ss) {
  var sheet = ss.getSheetByName('งบประมาณ');
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  var catMap = {};
  var catOrder = [];

  data.forEach(function(row) {
    if (!row[0] && !row[2]) return;
    var cat       = (row[0] || '').toString().trim();
    var catBudget = parseFloat(row[1]) || 0;
    var itemName  = (row[2] || '').toString().trim();
    var itemBudget = parseFloat(row[3]) || 0;

    if (cat) {
      if (!catMap[cat]) {
        catMap[cat] = { cat: cat, budget: catBudget, items: [] };
        catOrder.push(cat);
      } else if (row[1] !== '') {
        catMap[cat].budget = catBudget;
      }
      if (itemName) {
        catMap[cat].items.push({ name: itemName, budget: itemBudget });
      }
    }
  });

  return catOrder.map(function(c) { return catMap[c]; });
}

function saveBudget(ss, budget) {
  var sheet = ss.getSheetByName('งบประมาณ');
  if (!sheet) {
    sheet = ss.insertSheet('งบประมาณ');
    sheet.getRange(1, 1, 1, 4).setValues([['หมวดหมู่', 'งบรวมหมวด', 'รายการ', 'งบรายการ']]);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) sheet.getRange(2, 1, lastRow - 1, 4).clearContent();
  if (!budget || !budget.length) return;

  var rows = [];
  budget.forEach(function(cat) {
    cat.items.forEach(function(item) {
      rows.push([cat.cat, cat.budget, item.name, item.budget]);
    });
  });

  if (rows.length) sheet.getRange(2, 1, rows.length, 4).setValues(rows);
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
