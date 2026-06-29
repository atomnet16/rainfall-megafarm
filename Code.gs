// ============================================================
// MEGA FARM CAMBODIA — Rainfall App · Google Apps Script
// Container-bound script (Extensions → Apps Script จาก Sheet)
// Deploy: Execute as Me · Access: Anyone
// ============================================================
//
// โครงสร้าง Sheet "Rain 2026":
//   A   : Date  (format D/M/YY เช่น 12/6/26)
//   B   : WS#1  Office
//   C   : WS#3  C2-JF2
//   D   : WS#4  IRR-A2
//   E   : WS#5  T4S
//   F   : WS#6  T2SW
//   G   : WS#8  M4(P9)
//   H   : WS#15 CNBD1 Gate
//   I   : WS#16 T4N
//   J   : WS#2  M2          ← ลำดับไม่ตรงกับ code!
//   K   : WS#9  A3
//   L   : WS#17 A2/C2 PPA
// ============================================================

const CFG_SHEET_NAME  = 'AppConfig';
const LOG_SHEET_NAME  = 'RainfallLog';
const DEFAULT_PIN     = '1234';
const TZ              = 'Asia/Phnom_Penh';
const DATA_START_ROW  = 2; // row 1 = header
const LINE_TOKEN_KEY  = 'LINE_TOKEN';
const LINE_TARGET_KEY = 'LINE_TARGET'; // userId or groupId

function dataSheetName(year) {
  return 'Rain ' + year; // "Rain 2025", "Rain 2026", …
}

// ── Station definitions (ลำดับ = คอลัมน์จริงใน Sheet) ────────
const STATIONS = [
  { code:'WS#1',  name:'Office',     col:2  },  // B
  { code:'WS#3',  name:'C2-JF2',     col:3  },  // C
  { code:'WS#4',  name:'IRR-A2',     col:4  },  // D
  { code:'WS#5',  name:'T4S',        col:5  },  // E
  { code:'WS#6',  name:'T2SW',       col:6  },  // F
  { code:'WS#8',  name:'M4(P9)',     col:7  },  // G
  { code:'WS#15', name:'CNBD1 Gate', col:8  },  // H
  { code:'WS#16', name:'T4N',        col:9  },  // I
  { code:'WS#2',  name:'M2',         col:10 },  // J ← อยู่คอลัมน์ J ไม่ใช่ B
  { code:'WS#9',  name:'A3',         col:11 },  // K
  { code:'WS#17', name:'A2/C2 PPA',  col:12 },  // L
];

// ── Response ─────────────────────────────────────────────────
function out(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── GET ──────────────────────────────────────────────────────
function doGet(e) {
  const p = e.parameter || {};
  try {
    switch (p.action) {
      case 'getLatestDay':
        return out(getLatestDay());
      case 'getMonthly': {
        const now = new Date();
        return out(getMonthlyTable(
          parseInt(p.year  || now.getFullYear()),
          parseInt(p.month || now.getMonth() + 1)
        ));
      }
      case 'getRecentLog':
        return out(getRecentLog(parseInt(p.limit || 60)));
      case 'getStations':
        return out({ stations: STATIONS.map(s => ({ code: s.code, name: s.name })) });
      case 'getYearly': {
        const now2 = new Date();
        return out(getYearly(
          parseInt(p.year  || now2.getFullYear()),
          parseInt(p.monthFrom || 1),
          parseInt(p.monthTo   || 12)
        ));
      }
      case 'getAvailableYears': {
        const ss2 = SpreadsheetApp.getActiveSpreadsheet();
        const years = ss2.getSheets()
          .map(sh => { const m = sh.getName().match(/^Rain (\d{4})$/); return m ? parseInt(m[1]) : null; })
          .filter(Boolean).sort((a,b) => b-a);
        return out({ years });
      }
      case 'getFilledStations': {
        const date = p.date; // "2026-06-14"
        if (!date) return out({ filled: [] });
        // Use RainfallLog as source of truth — avoids counting stale/test values in sheet
        const logRows = getLogSheet().getDataRange().getValues();
        const filledSet = new Set();
        for (let i = 1; i < logRows.length; i++) {
          const r = logRows[i];
          if (r[8] === 'DELETED' || r[8] === 'DELETE') continue;
          if (r[8] !== 'ADD') continue;
          const rd = r[2];
          const ds = rd instanceof Date
            ? `${rd.getFullYear()}-${String(rd.getMonth()+1).padStart(2,'0')}-${String(rd.getDate()).padStart(2,'0')}`
            : String(rd);
          if (ds === date) filledSet.add(String(r[3]));
        }
        return out({ filled: [...filledSet] });
      }
      case 'ping':
        return out({ ok: true, time: new Date().toISOString() });
      case 'debugSheets': {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheets = ss.getSheets().map(s => s.getName());
        // ดู 5 rows แรกของ Rain 2025
        const sh25 = ss.getSheetByName('Rain 2025');
        const sample = sh25 ? sh25.getRange(1, 1, Math.min(6, sh25.getLastRow()), 3).getValues() : [];
        return out({ sheets, sample });
      }
      default:
        return out({ error: 'unknown action: ' + p.action });
    }
  } catch (err) {
    return out({ error: err.message + ' | ' + err.stack });
  }
}

// ── POST ─────────────────────────────────────────────────────
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    switch (body.action) {
      case 'addRecord':    return out(addRecord(body.data, body.pin));
      case 'deleteRecord': return out(deleteRecord(body.id, body.pin));
      case 'verifyPin':    return out({ valid: verifyPin(body.pin) });
      case 'changePin':    return out(changePin(body.oldPin, body.newPin));
      case 'saveConfig':   setCfgValue(body.key, body.value); return out({ success: true });
      default:             return out({ error: 'unknown action' });
    }
  } catch (err) {
    return out({ error: err.message });
  }
}

// ══════════════════════════════════════════════════════════════
// SHEET ACCESS
// ══════════════════════════════════════════════════════════════
function getDataSheet(year) {
  const name = dataSheetName(year || new Date().getFullYear());
  const sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('ไม่พบ Sheet "' + name + '"');
  return sh;
}

function getCfgSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh   = ss.getSheetByName(CFG_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(CFG_SHEET_NAME);
    sh.appendRow(['Key', 'Value']);
    sh.appendRow(['PIN', DEFAULT_PIN]);
    sh.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#7BA7D4').setFontColor('#fff');
    sh.setFrozenRows(1);
  }
  const vals = sh.getDataRange().getValues();
  if (!vals.find(r => r[0] === 'PIN')) sh.appendRow(['PIN', DEFAULT_PIN]);
  return sh;
}

function getLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh   = ss.getSheetByName(LOG_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(LOG_SHEET_NAME);
    const hdr = ['ID','Timestamp','Date','StationCode','StationName','OldValue_mm','NewValue_mm','EnteredBy','Action'];
    sh.appendRow(hdr);
    sh.getRange(1, 1, 1, hdr.length).setFontWeight('bold').setBackground('#7BA7D4').setFontColor('#fff');
    sh.setFrozenRows(1);
  }
  return sh;
}

// ══════════════════════════════════════════════════════════════
// DATE HELPERS
// ══════════════════════════════════════════════════════════════

// แปลง "2026-06-12" → "12/6/26"  (format ในไฟล์)
function appDateToSheetDate(iso) {
  const [y, m, d] = iso.split('-');
  return `${parseInt(d)}/${parseInt(m)}/${y.slice(2)}`;
}

// แปลงค่าจาก cell (Date object หรือ string) → "12/6/26"
function cellToSheetDate(cell) {
  if (!cell) return '';
  if (cell instanceof Date) {
    return `${cell.getDate()}/${cell.getMonth()+1}/${String(cell.getFullYear()).slice(2)}`;
  }
  return String(cell).trim();
}

// แปลง "12/6/26" → "2026-06-12"
function sheetDateToIso(sheetDate) {
  const parts = String(sheetDate).trim().split('/');
  if (parts.length !== 3) return '';
  const [d, m, y] = parts;
  const fullYear = parseInt(y) + 2000;
  return `${fullYear}-${String(m).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
}

// หาหมายเลข row (1-based) จาก date string "D/M/YY"
function findRowByDate(sh, sheetDateStr) {
  const lastRow = sh.getLastRow();
  if (lastRow < DATA_START_ROW) return -1;
  const dateCells = sh.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 1).getValues();
  for (let i = 0; i < dateCells.length; i++) {
    if (cellToSheetDate(dateCells[i][0]) === sheetDateStr) return i + DATA_START_ROW;
  }
  return -1;
}

// ══════════════════════════════════════════════════════════════
// PIN
// ══════════════════════════════════════════════════════════════
function getStoredPin() {
  const rows = getCfgSheet().getDataRange().getValues();
  const r = rows.find(r => r[0] === 'PIN');
  return r ? String(r[1]) : DEFAULT_PIN;
}
function verifyPin(pin) { return String(pin) === getStoredPin(); }
function changePin(oldPin, newPin) {
  if (!verifyPin(oldPin))              return { error: 'PIN เดิมไม่ถูกต้อง' };
  if (!/^\d{4}$/.test(String(newPin))) return { error: 'PIN ใหม่ต้องเป็นตัวเลข 4 หลัก' };
  const sh   = getCfgSheet();
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === 'PIN') { sh.getRange(i+1, 2).setValue(String(newPin)); return { success: true }; }
  }
  sh.appendRow(['PIN', String(newPin)]);
  return { success: true };
}

// ══════════════════════════════════════════════════════════════
// ADD RECORD  →  เขียนตรงลงเซลล์ที่ถูกต้องใน Rain 2026
// ══════════════════════════════════════════════════════════════
function addRecord(data, pin) {
  if (!verifyPin(pin)) return { error: 'PIN ไม่ถูกต้อง' };

  // ตรวจสอบสถานี
  const station = STATIONS.find(s => s.code === data.stationCode);
  if (!station) return { error: 'ไม่พบสถานี: ' + data.stationCode };

  // ตรวจสอบค่า mm
  const mm = parseFloat(data.rainfall_mm);
  if (isNaN(mm) || mm < 0) return { error: 'ปริมาณน้ำฝนไม่ถูกต้อง' };

  // แปลงวันที่
  const isoDate   = data.date; // "2026-06-12"
  const sheetDate = appDateToSheetDate(isoDate);
  const recYear   = parseInt(isoDate.slice(0, 4));

  // หา row ในชีตของปีนั้น
  const sh  = getDataSheet(recYear);
  const row = findRowByDate(sh, sheetDate);
  if (row === -1) return { error: `ไม่พบวันที่ ${sheetDate} ใน Sheet "${dataSheetName(recYear)}"` };

  // ตรวจสอบ column ของสถานี
  const col = station.col;

  // อ่านค่าเดิม
  const oldVal = parseFloat(sh.getRange(row, col).getValue()) || 0;

  // เขียนค่าใหม่
  sh.getRange(row, col).setValue(mm);

  // บันทึก audit log
  const logId = String(Date.now());
  const now   = new Date();
  getLogSheet().appendRow([
    logId, now.toISOString(), isoDate,
    station.code, station.name, oldVal, mm,
    data.enteredBy || 'App', 'ADD'
  ]);

  checkAndSendLineSummary(isoDate);

  return {
    success:   true,
    id:        logId,
    sheetDate: sheetDate,
    row:       row,
    col:       col,
    colLetter: String.fromCharCode(64 + col), // A=65
    station:   station.code,
    name:      station.name,
    oldValue:  oldVal,
    newValue:  mm
  };
}

// ══════════════════════════════════════════════════════════════
// DELETE RECORD  →  reset cell กลับเป็น 0
// ══════════════════════════════════════════════════════════════
function deleteRecord(id, pin) {
  if (!verifyPin(pin)) return { error: 'PIN ไม่ถูกต้อง' };

  // หาใน audit log ว่า record นี้คือ date + station อะไร
  const logSh = getLogSheet();
  const rows  = logSh.getDataRange().getValues();
  let logRow  = -1;
  let rec     = null;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id) && rows[i][8] !== 'DELETE') {
      logRow = i + 1;
      const rawDate = rows[i][2];
      const dateStr = rawDate instanceof Date
        ? `${rawDate.getFullYear()}-${String(rawDate.getMonth()+1).padStart(2,'0')}-${String(rawDate.getDate()).padStart(2,'0')}`
        : String(rawDate);
      rec = { date: dateStr, stationCode: rows[i][3] };
      break;
    }
  }
  if (!rec) return { error: 'ไม่พบรายการ ID: ' + id + ' (อาจถูกลบไปแล้ว)' };

  // หาสถานี
  const station = STATIONS.find(s => s.code === rec.stationCode);
  if (!station) return { error: 'ไม่พบสถานี' };

  // หา row ในชีตของปีนั้น
  const recYear   = parseInt(String(rec.date).slice(0, 4));
  const sh        = getDataSheet(recYear);
  const sheetDate = appDateToSheetDate(rec.date);
  const dataRow   = findRowByDate(sh, sheetDate);
  if (dataRow === -1) return { error: 'ไม่พบวันที่ใน Sheet' };

  // อ่านค่าปัจจุบัน
  const currentVal = parseFloat(sh.getRange(dataRow, station.col).getValue()) || 0;

  // Reset เป็น 0
  sh.getRange(dataRow, station.col).setValue(0);

  // บันทึก audit log
  const now = new Date();
  logSh.appendRow([
    String(Date.now()), now.toISOString(), rec.date,
    station.code, station.name, currentVal, 0, 'System', 'DELETE'
  ]);

  // Mark original log row เป็น deleted
  if (logRow > 0) logSh.getRange(logRow, 9).setValue('DELETED');

  return { success: true, sheetDate, station: station.code, resetTo: 0 };
}

// ══════════════════════════════════════════════════════════════
// GET LATEST DAY
// ══════════════════════════════════════════════════════════════
function getLatestDay() {
  const sh      = getDataSheet(new Date().getFullYear());
  const lastRow = sh.getLastRow();
  if (lastRow < DATA_START_ROW) return { date: null, summary: emptySummary() };

  const all = sh.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 12).getValues();

  // หา row ล่าสุดที่มีค่า > 0 (ไม่นับ row ที่ทุกช่องว่างหรือ 0)
  let latestRowIdx = -1;
  let latestIso    = '';
  for (let i = all.length - 1; i >= 0; i--) {
    const row     = all[i];
    const dateStr = cellToSheetDate(row[0]);
    if (!dateStr) continue;
    const iso = sheetDateToIso(dateStr);
    if (!iso)  continue;
    // ตรวจว่ามีค่า > 0 อย่างน้อย 1 สถานี
    const hasData = STATIONS.some(s => parseFloat(row[s.col - 1]) > 0);
    if (hasData) {
      latestRowIdx = i;
      latestIso    = iso;
      break;
    }
  }

  if (latestRowIdx === -1) return { date: null, summary: emptySummary() };

  const rowData = all[latestRowIdx];
  const summary = STATIONS.map(s => ({
    code:     s.code,
    name:     s.name,
    total_mm: parseFloat(rowData[s.col - 1]) || 0
  }));

  return { date: latestIso, summary, updatedAt: new Date().toISOString() };
}

// ══════════════════════════════════════════════════════════════
// GET MONTHLY TABLE
// ══════════════════════════════════════════════════════════════
function getMonthlyTable(year, month) {
  const sh      = getDataSheet(year);
  const lastRow = sh.getLastRow();
  if (lastRow < DATA_START_ROW) return emptyMonthly(year, month);

  const all    = sh.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 12).getValues();
  const matrix = {};
  const totals = {};
  STATIONS.forEach(s => { totals[s.code] = 0; });

  // prefix เช่น "2026-06"
  const prefix = `${year}-${String(month).padStart(2, '0')}`;

  for (let i = 0; i < all.length; i++) {
    const row     = all[i];
    const dateStr = cellToSheetDate(row[0]);
    if (!dateStr) continue;
    const iso = sheetDateToIso(dateStr);
    if (!iso || !iso.startsWith(prefix)) continue;

    const day = iso.slice(8, 10); // "01"–"31"
    if (!matrix[day]) matrix[day] = {};

    STATIONS.forEach(s => {
      const mm = parseFloat(row[s.col - 1]) || 0;
      if (!matrix[day][s.code]) matrix[day][s.code] = 0;
      matrix[day][s.code] = +((matrix[day][s.code] + mm).toFixed(2));
      totals[s.code]       = +((totals[s.code]       + mm).toFixed(2));
    });
  }

  const daysInMonth = new Date(year, month, 0).getDate();
  const days = Array.from({ length: daysInMonth }, (_, i) => String(i + 1).padStart(2, '0'));

  return {
    year, month, prefix, days,
    stations: STATIONS.map(s => ({ code: s.code, name: s.name })),
    matrix, totals,
    updatedAt: new Date().toISOString()
  };
}

// ══════════════════════════════════════════════════════════════
// GET RECENT LOG (audit trail)
// ══════════════════════════════════════════════════════════════
function getRecentLog(limit) {
  const sh   = getLogSheet();
  const rows = sh.getDataRange().getValues();
  const recs = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (r[8] === 'DELETED' || r[8] === 'DELETE') continue;
    const ts = r[1];
    const timeStr = ts instanceof Date
      ? ts.toTimeString().slice(0,5)
      : (String(ts).length >= 16 ? String(ts).slice(11,16) : '');
    const rawDate = r[2];
    const dateStr = rawDate instanceof Date
      ? `${rawDate.getFullYear()}-${String(rawDate.getMonth()+1).padStart(2,'0')}-${String(rawDate.getDate()).padStart(2,'0')}`
      : String(rawDate);
    recs.push({
      id:          String(r[0]),
      timestamp:   ts,
      time:        timeStr,
      date:        dateStr,
      stationCode: r[3],
      stationName: r[4],
      oldValue:    parseFloat(r[5]) || 0,
      rainfall_mm: parseFloat(r[6]) || 0,
      enteredBy:   r[7],
      action:      r[8]
    });
  }
  recs.reverse();
  return { records: recs.slice(0, limit || 60), updatedAt: new Date().toISOString() };
}

// ══════════════════════════════════════════════════════════════
// HELPERS
// ══════════════════════════════════════════════════════════════
// ══════════════════════════════════════════════════════════════
// LINE NOTIFY
// ══════════════════════════════════════════════════════════════
function getCfgValue(key) {
  const rows = getCfgSheet().getDataRange().getValues();
  const r = rows.find(r => r[0] === key);
  return r ? String(r[1]).trim() : '';
}

function setCfgValue(key, value) {
  const sh   = getCfgSheet();
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === key) { sh.getRange(i+1, 2).setValue(value); return; }
  }
  sh.appendRow([key, value]);
}

function checkAndSendLineSummary(isoDate) {
  const token  = getCfgValue(LINE_TOKEN_KEY);
  const target = getCfgValue(LINE_TARGET_KEY);
  if (!token || !target) return;

  // ตรวจจาก RainfallLog — เหมือน getFilledStations เพื่อความสอดคล้อง
  const logRows = getLogSheet().getDataRange().getValues();
  const reportedMap = {}; // stationCode → mm value (ล่าสุด)
  for (let i = 1; i < logRows.length; i++) {
    const r = logRows[i];
    if (r[8] === 'DELETED' || r[8] === 'DELETE') continue;
    if (r[8] !== 'ADD') continue;
    const rd = r[2];
    const ds = rd instanceof Date
      ? `${rd.getFullYear()}-${String(rd.getMonth()+1).padStart(2,'0')}-${String(rd.getDate()).padStart(2,'0')}`
      : String(rd);
    if (ds === isoDate) reportedMap[String(r[3])] = r[6]; // NewValue_mm
  }

  // ถ้ายังมีสถานีที่ยังไม่รายงาน → ยังไม่ส่ง
  const allFilled = STATIONS.every(s => reportedMap.hasOwnProperty(s.code));
  if (!allFilled) return;

  const stationVals = STATIONS.map(s => ({ station: s, val: reportedMap[s.code] }));

  // สร้าง summary message
  const [y, m, d] = isoDate.split('-');
  const dateLabel = `${parseInt(d)}/${parseInt(m)}/${y}`;
  let totalMm = 0;
  let lines = '';
  stationVals.forEach(sv => {
    const mm = parseFloat(sv.val) || 0;
    totalMm += mm;
    lines += `${mm > 0 ? '💧' : '○'} ${sv.station.code} ${sv.station.name}: ${mm.toFixed(1)} mm\n`;
  });
  const avgMm = (totalMm / STATIONS.length).toFixed(1);
  const msg = `🌧 Mega Farm Daily Rainfall\n📅 ${dateLabel} — All stations reported\n\n${lines}\n📊 Farm avg: ${avgMm} mm  |  Total: ${totalMm.toFixed(1)} mm`;

  try {
    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + token },
      payload: JSON.stringify({ to: target, messages: [{ type: 'text', text: msg }] }),
      muteHttpExceptions: true
    });
  } catch(e) {
    console.log('LINE summary error: ' + e.message);
  }
}

// ══════════════════════════════════════════════════════════════
// GET YEARLY  — monthly totals + rain days per station
// ══════════════════════════════════════════════════════════════
function getYearly(year, monthFrom, monthTo) {
  const mFrom = Math.max(1,  monthFrom || 1);
  const mTo   = Math.min(12, monthTo   || 12);

  let sh, all;
  try {
    sh  = getDataSheet(year);
    const lastRow = sh.getLastRow();
    if (lastRow < DATA_START_ROW) return emptyYearly(year, mFrom, mTo);
    all = sh.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 12).getValues();
  } catch(e) {
    return emptyYearly(year, mFrom, mTo);
  }

  const monthData = {};
  for (let m = mFrom; m <= mTo; m++) {
    monthData[m] = { totals: {}, rainDays: {} };
    STATIONS.forEach(s => { monthData[m].totals[s.code] = 0; monthData[m].rainDays[s.code] = 0; });
  }

  const seenRainDay = {};

  for (let i = 0; i < all.length; i++) {
    const row     = all[i];
    const dateStr = cellToSheetDate(row[0]);
    if (!dateStr) continue;
    const iso = sheetDateToIso(dateStr);
    if (!iso) continue;
    const rowYear = parseInt(iso.slice(0, 4));
    if (rowYear !== year) continue;
    const month = parseInt(iso.slice(5, 7));
    if (month < mFrom || month > mTo) continue;
    const day = iso.slice(8, 10);

    STATIONS.forEach(s => {
      const mm = parseFloat(row[s.col - 1]) || 0;
      monthData[month].totals[s.code] = +((monthData[month].totals[s.code] + mm).toFixed(2));
      if (mm > 0) {
        const key = `${month}-${s.code}-${day}`;
        if (!seenRainDay[key]) { seenRainDay[key] = true; monthData[month].rainDays[s.code]++; }
      }
    });
  }

  const stationTotals   = {};
  const stationRainDays = {};
  STATIONS.forEach(s => {
    stationTotals[s.code]   = 0;
    stationRainDays[s.code] = 0;
    for (let m = mFrom; m <= mTo; m++) {
      stationTotals[s.code]   = +((stationTotals[s.code]   + (monthData[m].totals[s.code]   || 0)).toFixed(2));
      stationRainDays[s.code] +=                                monthData[m].rainDays[s.code] || 0;
    }
  });

  const months = [];
  for (let m = mFrom; m <= mTo; m++) {
    months.push({ month: m, totals: monthData[m].totals, rainDays: monthData[m].rainDays });
  }

  return {
    year, monthFrom: mFrom, monthTo: mTo,
    months,
    stations:         STATIONS.map(s => ({ code: s.code, name: s.name })),
    stationTotals,
    stationRainDays,
    updatedAt: new Date().toISOString()
  };
}

function emptyYearly(year, mFrom, mTo) {
  const months = [];
  for (let m = mFrom; m <= mTo; m++) {
    const t = {}, rd = {};
    STATIONS.forEach(s => { t[s.code] = 0; rd[s.code] = 0; });
    months.push({ month: m, totals: t, rainDays: rd });
  }
  const st = {}, srd = {};
  STATIONS.forEach(s => { st[s.code] = 0; srd[s.code] = 0; });
  return {
    year, monthFrom: mFrom, monthTo: mTo,
    months, stations: STATIONS.map(s => ({ code: s.code, name: s.name })),
    stationTotals: st, stationRainDays: srd,
    updatedAt: new Date().toISOString()
  };
}

function emptySummary() {
  return STATIONS.map(s => ({ code: s.code, name: s.name, total_mm: 0 }));
}
function emptyMonthly(year, month) {
  const daysInMonth = new Date(year, month, 0).getDate();
  return {
    year, month,
    prefix:   `${year}-${String(month).padStart(2,'0')}`,
    days:     Array.from({ length: daysInMonth }, (_, i) => String(i+1).padStart(2,'0')),
    stations: STATIONS.map(s => ({ code: s.code, name: s.name })),
    matrix:   {},
    totals:   Object.fromEntries(STATIONS.map(s => [s.code, 0])),
    updatedAt: new Date().toISOString()
  };
}
