// ╔══════════════════════════════════════════════════════════════╗
// ║  Google Apps Script — ระบบตรวจเช็คเครื่องมือแพทย์ ER       ║
// ║  วิธีใช้:                                                    ║
// ║  1. เปิด script.google.com → New Project                    ║
// ║  2. วางโค้ดนี้ทั้งหมด                                        ║
// ║  3. Deploy → New Deployment → Web App                       ║
// ║     - Execute as: Me                                        ║
// ║     - Who has access: Anyone                                ║
// ║  4. Copy URL ไปใส่ใน index.html ที่ SCRIPT_URL              ║
// ╚══════════════════════════════════════════════════════════════╝

const SHEET_ID   = '1yEXZQdP3C_OAKgMxiP-DZXKTA_XMD8_kryZInyvC440';
const SHEET_DATA = 'ข้อมูลการตรวจเช็ค';
const HEADERS    = [
  'Timestamp','วันที่','ผู้ตรวจสอบ','ระดับความเสี่ยง',
  'ชื่อเครื่องมือ','หน่วยที่','สถานะ',
  'รายการผ่าน','รายการไม่ผ่าน','คะแนน'
];

// ─── POST: รับข้อมูลจากฟอร์ม บันทึกลง Sheet ───────────────────
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(15000);

  try {
    const body = JSON.parse(e.postData.contents);
    const rows = body.rows || [];

    if (rows.length === 0) {
      return jsonResponse({ success: false, error: 'No data' });
    }

    const sheet = getOrCreateSheet();

    const values = rows.map(r => [
      r.timestamp,
      r.date,
      r.inspector,
      translateRisk(r.riskLevel),
      r.equipmentName,
      r.unitNo,
      translateStatus(r.status),
      r.checksPassed  || '',
      r.checksFailed  || '',
      r.score         || 'N/A'
    ]);

    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, values.length, HEADERS.length).setValues(values);

    // Style rows by status
    values.forEach((row, idx) => {
      const rowNum = startRow + idx;
      const status = rows[idx].status;
      const color  = status === 'cancel' ? '#FFEBEE'
                   : status === 'partial' ? '#FFF3E0'
                   : '#F1F8E9';
      sheet.getRange(rowNum, 1, 1, HEADERS.length).setBackground(color);
    });

    return jsonResponse({ success: true, added: rows.length });

  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

// ─── GET: ส่งข้อมูลให้ Dashboard ─────────────────────────────────
function doGet(e) {
  const action = (e.parameter && e.parameter.action) || 'getData';
  const days   = parseInt((e.parameter && e.parameter.days) || '30', 10);

  if (action === 'getData') {
    return getData(days);
  }
  return jsonResponse({ error: 'Unknown action' });
}

function getData(days) {
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_DATA);

    if (!sheet || sheet.getLastRow() <= 1) {
      return jsonResponse({ rows: [] });
    }

    const data    = sheet.getDataRange().getValues();
    const cutoff  = new Date();
    cutoff.setDate(cutoff.getDate() - days);

    const rows = data.slice(1)
      .filter(row => {
        const ts = row[0] ? new Date(row[0]) : null;
        return ts && ts >= cutoff;
      })
      .map(row => ({
        timestamp:     row[0] ? new Date(row[0]).toISOString() : '',
        date:          row[1] ? formatDate(row[1]) : '',
        inspector:     row[2] || '',
        riskLevel:     reverseRisk(row[3]),
        equipmentName: row[4] || '',
        unitNo:        row[5] || '',
        status:        reverseStatus(row[6]),
        checksPassed:  row[7] || '',
        checksFailed:  row[8] || '',
        score:         row[9] || ''
      }));

    return jsonResponse({ rows, total: rows.length });

  } catch (err) {
    return jsonResponse({ rows: [], error: err.toString() });
  }
}

// ─── HELPERS ─────────────────────────────────────────────────────
function getOrCreateSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_DATA);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_DATA);

    // Headers
    const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setValues([HEADERS]);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#0D47A1');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setFontSize(12);
    sheet.setFrozenRows(1);

    // Column widths
    const widths = [180, 100, 160, 140, 220, 60, 120, 300, 300, 80];
    widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  }

  return sheet;
}

function translateRisk(r) {
  const m = { high: 'ความเสี่ยงสูง', medium: 'ความเสี่ยงปานกลาง', low: 'ความเสี่ยงต่ำ' };
  return m[r] || r;
}
function reverseRisk(r) {
  const m = { 'ความเสี่ยงสูง':'high','ความเสี่ยงปานกลาง':'medium','ความเสี่ยงต่ำ':'low' };
  return m[r] || r;
}
function translateStatus(s) {
  const m = { ready:'พร้อมใช้/มี', partial:'ไม่พร้อมใช้ แต่ยังใช้', cancel:'ยกเลิกการใช้งาน' };
  return m[s] || s;
}
function reverseStatus(s) {
  if (s === 'พร้อมใช้/มี' || s === 'ready') return 'ready';
  if (s === 'ยกเลิกการใช้งาน' || s === 'cancel') return 'cancel';
  return 'partial';
}
function formatDate(d) {
  if (!d) return '';
  const dt = new Date(d);
  return isNaN(dt) ? String(d) : dt.toISOString().slice(0, 10);
}
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
