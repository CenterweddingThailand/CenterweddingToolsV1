// ════════════════════════════════════════════════════════════
//  CW Club — Budget Advisor: Google Apps Script Backend
//  วางทั้งไฟล์นี้ใน Google Apps Script แล้ว Deploy as Web App
//  Execute as: Me | Who has access: Anyone
// ════════════════════════════════════════════════════════════

const SHEET_LEADS     = 'Leads';      // ชีทเก็บ quotation requests
const SHEET_ANALYTICS = 'Analytics'; // ชีทเก็บทุก event

// ── Headers ────────────────────────────────────────────────
const LEAD_HEADERS = [
  'Timestamp','วันที่','เวลา',
  'ชื่อ','เบอร์โทร','LINE ID',
  'LINE User ID','LINE Display Name',
  'สถานที่','ประเภทงาน','จำนวนแขก','ระดับงาน',
  'งบประมาณ (THB)','Source','Device'
];
const ANALYTICS_HEADERS = [
  'Timestamp','วันที่','เวลา',
  'Action','สถานที่','ประเภทงาน',
  'จำนวนแขก','ระดับงาน','งบประมาณ (THB)',
  'Source','Device'
];

// ── doGet — รับทุก event (analytics + leads) ───────────────
function doGet(e) {
  const p      = e.parameter;
  const action = p.action || 'open';
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const now    = new Date();
  const date   = Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM-dd');
  const time   = Utilities.formatDate(now, 'Asia/Bangkok', 'HH:mm:ss');

  if (action === 'lead') {
    // บันทึก quotation request ลง Leads sheet
    const sheet = getOrCreateSheet(ss, SHEET_LEADS, LEAD_HEADERS);
    sheet.appendRow([
      now, date, time,
      p.name       || '',
      p.phone      || '',
      p.lineId     || '',
      p.lineUserId || '',
      p.lineName   || '',
      p.loc        || '',
      p.type       || '',
      p.guests     || 0,
      p.tier       || '',
      p.budget     || 0,
      p.source     || 'direct',
      p.device     || 'unknown'
    ]);
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 13).setNumberFormat('#,##0');
  } else {
    // บันทึก analytics event
    const sheet = getOrCreateSheet(ss, SHEET_ANALYTICS, ANALYTICS_HEADERS);
    sheet.appendRow([
      now, date, time,
      action,
      p.loc    || '',
      p.type   || '',
      p.guests || 0,
      p.tier   || '',
      p.budget || 0,
      p.source || 'direct',
      p.device || 'unknown'
    ]);
  }

  // Return transparent 1x1 GIF
  return ContentService.createTextOutput('OK');
}

// ── doPost — รับ lead form submissions ─────────────────────
function doPost(e) {
  try {
    const p = e.parameter;

    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    const now  = new Date();

    if (p.action === 'lead') {
      const sheet = getOrCreateSheet(ss, SHEET_LEADS, LEAD_HEADERS);
      sheet.appendRow([
        now,
        Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM-dd'),
        Utilities.formatDate(now, 'Asia/Bangkok', 'HH:mm:ss'),
        p.name        || '',
        p.phone       || '',
        p.lineId      || '',
        p.lineUserId  || '',
        p.lineName    || '',
        p.loc         || '',
        p.type        || '',
        p.guests      || 0,
        p.tier        || '',
        p.budget      || 0,
        p.source      || 'direct',
        p.device      || 'unknown'
      ]);

      // Auto-format budget column as number
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow, 13).setNumberFormat('#,##0');
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', msg: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Helper: get or create sheet with headers ───────────────
function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    const header = sheet.getRange(1, 1, 1, headers.length);
    header.setValues([headers]);
    header.setFontWeight('bold');
    header.setBackground('#1C2B3A');
    header.setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160); // Timestamp
    sheet.setColumnWidth(4, 140); // ชื่อ
    sheet.setColumnWidth(5, 120); // เบอร์
  }
  return sheet;
}

// ── Auto summary sheet (optional: run once manually) ───────
function createSummarySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Dashboard');
  if (!sheet) sheet = ss.insertSheet('Dashboard');

  sheet.clear();

  // Title
  sheet.getRange('A1').setValue('CW Club Budget Advisor — Summary Dashboard');
  sheet.getRange('A1').setFontSize(14).setFontWeight('bold');
  sheet.getRange('A2').setValue('Auto-updated from Leads & Analytics sheets');
  sheet.getRange('A2').setFontColor('#6b7280').setFontSize(11);

  // Metrics (formulas pulling from other sheets)
  const metrics = [
    ['A4', 'Metric', 'B4', 'Value'],
    ['A5', 'Total Opens',           'B5', "=COUNTA(Analytics!A:A)-1"],
    ['A6', 'Total Calculations',    'B6', "=COUNTIF(Analytics!D:D,\"calc\")"],
    ['A7', 'Total Leads (Quotations)','B7',"=COUNTA(Leads!A:A)-1"],
    ['A8', 'Conversion Rate',       'B8', "=IFERROR(B7/B6,0)"],
    ['A10','Popular Event Type',    'B10',"=IFERROR(INDEX(Analytics!F:F,MATCH(MAX(COUNTIF(Analytics!F2:F999,Analytics!F2:F999)),COUNTIF(Analytics!F2:F999,Analytics!F2:F999),0)+1),\"-\")"],
    ['A11','Avg Budget (THB)',      'B11',"=IFERROR(AVERAGEIF(Analytics!D:D,\"calc\",Analytics!I:I),0)"],
    ['A12','Popular Location',      'B12',"=IFERROR(INDEX(Analytics!E:E,MATCH(MAX(COUNTIF(Analytics!E2:E999,Analytics!E2:E999)),COUNTIF(Analytics!E2:E999,Analytics!E2:E999),0)+1),\"-\")"],
  ];

  metrics.forEach(([cellA, labelA, cellB, valB]) => {
    sheet.getRange(cellA).setValue(labelA).setFontWeight('bold');
    sheet.getRange(cellB).setValue(valB);
  });

  // Format conversion rate as %
  sheet.getRange('B8').setNumberFormat('0.0%');
  sheet.getRange('B11').setNumberFormat('#,##0');

  // Header formatting
  sheet.getRange('A4:B4').setBackground('#1C2B3A').setFontColor('#ffffff').setFontWeight('bold');
  sheet.getRange('A10:B10').setBackground('#f4f5f7');
  sheet.getRange('A11:B11').setBackground('#f4f5f7');
  sheet.getRange('A12:B12').setBackground('#f4f5f7');

  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 180);

  SpreadsheetApp.getActiveSpreadsheet().toast('Dashboard sheet created!', 'Done', 3);
}
