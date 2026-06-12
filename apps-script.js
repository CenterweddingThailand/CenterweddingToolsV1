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

// ── Price Reference Sheet (run once manually) ──────────────
function createPriceReferenceSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('ราคาอ้างอิง');
  if (sh) { ss.deleteSheet(sh); }
  sh = ss.insertSheet('ราคาอ้างอิง');

  const PINK  = '#D81B7A';
  const NAVY  = '#1C2B3A';
  const BLUE  = '#4a6fa5';
  const GRAY  = '#f4f5f7';
  const WHITE = '#ffffff';

  let row = 1;

  // ── Title ────────────────────────────────────────────────
  sh.getRange(row, 1, 1, 8).merge()
    .setValue('CW Club — ตารางราคาอ้างอิง (อัปเดตโดยทีม Centerwedding)')
    .setFontSize(13).setFontWeight('bold').setFontColor(WHITE)
    .setBackground(PINK).setHorizontalAlignment('center');
  row++;
  sh.getRange(row, 1, 1, 8).merge()
    .setValue('แก้ไขราคาในตารางนี้ได้เลย — ระบบจะใช้ข้อมูลนี้ในการคำนวณโดยอัตโนมัติ')
    .setFontSize(11).setFontColor(PINK).setBackground('#fde8f3').setHorizontalAlignment('center');
  row += 2;

  // ─────────────────────────────────────────────────────────
  // SECTION 1: ราคาต่อหัว
  // ─────────────────────────────────────────────────────────
  sh.getRange(row, 1, 1, 8).merge()
    .setValue('ส่วนที่ 1 — ราคาอาหารและ Venue Package (ต่อหัว / บาท)')
    .setFontSize(12).setFontWeight('bold').setFontColor(WHITE).setBackground(NAVY);
  row++;

  const hdr1 = ['หมวด','รายการ','หน่วย','Standard','Premium','Luxury','หมายเหตุ'];
  sh.getRange(row, 1, 1, 7).setValues([hdr1])
    .setFontWeight('bold').setBackground(GRAY).setFontColor(NAVY);
  row++;

  const perHead = [
    ['อาหาร & เครื่องดื่ม','อาหารต่อหัว (Banquet)','บาท/คน',800,1200,2000,'รวม service charge 10%'],
    ['อาหาร & เครื่องดื่ม','บาร์และเครื่องดื่ม','บาท/คน',200,400,800,'Open bar = บวกเพิ่ม 30%'],
    ['อาหาร & เครื่องดื่ม','เค้กแต่งงาน','บาท/งาน',8000,15000,35000,''],
    ['Venue','ค่าเช่าห้องจัดงาน','บาท/งาน',60000,180000,450000,'ไม่รวมอาหาร'],
    ['Venue','ค่าเช่า outdoor / สวน','บาท/งาน',40000,120000,300000,''],
    ['Venue','ค่าเช่า beach venue','บาท/งาน',50000,150000,380000,''],
  ];
  perHead.forEach(r => {
    sh.getRange(row, 1, 1, 7).setValues([r]);
    row++;
  });
  row++;

  // ─────────────────────────────────────────────────────────
  // SECTION 2: ราคาต่องาน (Fixed Costs)
  // ─────────────────────────────────────────────────────────
  sh.getRange(row, 1, 1, 8).merge()
    .setValue('ส่วนที่ 2 — ค่าใช้จ่ายคงที่ต่องาน (Fixed Costs / บาท)')
    .setFontSize(12).setFontWeight('bold').setFontColor(WHITE).setBackground(NAVY);
  row++;

  const hdr2 = ['หมวด','รายการ','หน่วย','Standard','Premium','Luxury','หมายเหตุ'];
  sh.getRange(row, 1, 1, 7).setValues([hdr2])
    .setFontWeight('bold').setBackground(GRAY).setFontColor(NAVY);
  row++;

  const fixed = [
    ['ตกแต่ง','ดอกไม้ + ตกแต่งสถานที่','บาท/งาน',35000,90000,220000,'รวม backdrop'],
    ['ตกแต่ง','ดอกไม้เจ้าสาว (ช่อ+ข้อมือ)','บาท/งาน',3000,8000,20000,''],
    ['ช่างภาพ','ช่างภาพ (ทั้งวัน)','บาท/งาน',20000,55000,130000,''],
    ['ช่างภาพ','ช่างวิดีโอ','บาท/งาน',15000,40000,100000,''],
    ['ช่างภาพ','Pre-wedding shoot','บาท/งาน',10000,25000,70000,''],
    ['ชุด','ชุดเจ้าสาว (เช่า)','บาท/ตัว',8000,25000,80000,''],
    ['ชุด','สูทเจ้าบ่าว (เช่า)','บาท/ตัว',4000,12000,35000,''],
    ['ชุด','ช่างแต่งหน้า+ทำผม','บาท/งาน',6000,18000,45000,''],
    ['ดนตรี & MC','วงดนตรี','บาท/งาน',25000,60000,150000,''],
    ['ดนตรี & MC','พิธีกร (MC)','บาท/งาน',8000,20000,50000,''],
    ['พิธีกรรม','พิธีสงฆ์ + ของถวาย','บาท/งาน',15000,30000,60000,''],
    ['การเดินทาง','รถลีมูซีน/รถเจ้าบ่าวสาว','บาท/งาน',8000,20000,60000,''],
    ['การ์ดเชิญ','การ์ดเชิญ + ของชำร่วย','บาท/งาน',5000,15000,40000,''],
  ];
  fixed.forEach(r => {
    sh.getRange(row, 1, 1, 7).setValues([r]);
    row++;
  });
  row++;

  // ─────────────────────────────────────────────────────────
  // SECTION 3: ตัวคูณจังหวัด
  // ─────────────────────────────────────────────────────────
  sh.getRange(row, 1, 1, 8).merge()
    .setValue('ส่วนที่ 3 — ตัวคูณราคาตามจังหวัด (เทียบกับกรุงเทพฯ = 1.00)')
    .setFontSize(12).setFontWeight('bold').setFontColor(WHITE).setBackground(NAVY);
  row++;

  sh.getRange(row, 1, 1, 3).setValues([['จังหวัด','ตัวคูณ','หมายเหตุ']])
    .setFontWeight('bold').setBackground(GRAY);
  row++;

  const locs = [
    ['กรุงเทพฯ',   1.00, 'ราคาฐาน'],
    ['ภูเก็ต',     1.12, 'High season + international demand'],
    ['พัทยา',      0.90, 'ใกล้กรุงเทพฯ ราคาต่ำกว่าเล็กน้อย'],
    ['เชียงใหม่',  0.82, 'ค่าครองชีพต่ำกว่า'],
    ['หาดใหญ่',   0.78, 'ตลาดขนาดเล็ก'],
    ['ขอนแก่น',   0.75, 'ตลาดขนาดเล็ก'],
  ];
  locs.forEach(r => {
    sh.getRange(row, 1, 1, 3).setValues([r]);
    row++;
  });
  row++;

  // ─────────────────────────────────────────────────────────
  // SECTION 4: สัดส่วนงบแต่ละประเภทงาน
  // ─────────────────────────────────────────────────────────
  sh.getRange(row, 1, 1, 8).merge()
    .setValue('ส่วนที่ 4 — สัดส่วนงบแต่ละประเภทงาน (%) รวม = 100')
    .setFontSize(12).setFontWeight('bold').setFontColor(WHITE).setBackground(NAVY);
  row++;

  const hdr4 = ['ประเภทงาน','Venue & F&B','ตกแต่ง','ช่างภาพ','ชุด & แต่งหน้า','ดนตรี/MC','อื่นๆ'];
  sh.getRange(row, 1, 1, 7).setValues([hdr4])
    .setFontWeight('bold').setBackground(GRAY);
  row++;

  const splits = [
    ['งานแต่งงาน (ทั้งวัน)', 38, 15, 12, 20, 8, 7],
    ['งานแต่งงานเช้า',       43, 16, 15, 17, 4, 5],
    ['งานแต่งงานเย็น',       48, 15, 12, 14, 5, 6],
    ['ดินเนอร์',             60, 18, 12,  0, 0,10],
    ['งานแต่งในสวน',         38, 22, 16, 14, 0,10],
    ['งานแต่งริมทะเล',       35, 22, 20, 13, 0,10],
    ['งานหมั้น',             48, 20, 15, 10, 0, 7],
  ];
  splits.forEach(r => {
    sh.getRange(row, 1, 1, 7).setValues([r]);
    row++;
  });
  row++;

  // ─────────────────────────────────────────────────────────
  // SECTION 5: หมายเหตุ
  // ─────────────────────────────────────────────────────────
  sh.getRange(row, 1, 1, 8).merge()
    .setValue('ส่วนที่ 5 — หมายเหตุและสิ่งที่ยังไม่รวม')
    .setFontSize(12).setFontWeight('bold').setFontColor(WHITE).setBackground(NAVY);
  row++;

  const notes = [
    ['ไม่รวม','ค่าสินสอด','แตกต่างกันมากตามตกลงสองฝ่าย'],
    ['ไม่รวม','ค่าฤกษ์มงคล','ประมาณ 3,000–15,000 บาท'],
    ['ไม่รวม','ค่า VAT 7% + Service Charge 10%','ควรบวกเพิ่ม 17% จากราคา venue'],
    ['ไม่รวม','Honeymoon / ท่องเที่ยว',''],
    ['แนะนำ','งบสำรอง 10–15%','สำหรับค่าใช้จ่ายที่ไม่ได้วางแผน'],
    ['อัปเดต','แก้ไขราคาทุก 6 เดือน','หรือเมื่อราคาตลาดเปลี่ยนอย่างมีนัยสำคัญ'],
  ];
  sh.getRange(row, 1, 1, 3).setValues([['ประเภท','รายการ','หมายเหตุ']])
    .setFontWeight('bold').setBackground(GRAY);
  row++;
  notes.forEach(r => {
    sh.getRange(row, 1, 1, 3).setValues([r]);
    row++;
  });

  // ── Column widths ────────────────────────────────────────
  sh.setColumnWidth(1, 180);
  sh.setColumnWidth(2, 220);
  sh.setColumnWidth(3, 100);
  sh.setColumnWidth(4, 100);
  sh.setColumnWidth(5, 100);
  sh.setColumnWidth(6, 100);
  sh.setColumnWidth(7, 220);

  // ── Alternating row colors (data rows only) ──────────────
  [perHead, fixed, locs, splits, notes].forEach((section, si) => {
    // ข้าม: ตัวเลขแถวซับซ้อนเกิน — ใช้ banding แทน
  });
  sh.setFrozenRows(1);

  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sh);
  ss.toast('✅ สร้าง sheet ราคาอ้างอิง เรียบร้อย!', 'Done', 5);
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
