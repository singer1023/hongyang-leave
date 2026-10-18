/**
 * 虹揚廣告 出勤／休假登記 — Google Apps Script 線上版
 * 資料存在本試算表；休假會自動寫入公司 Google 行事曆（當日 8:00~9:00，只在早上標示，不佔整天）。
 */
const CALENDAR_ID = 'e16b2186b5dc4cf8e75427aff9851bd8f2951270ecf7d29c0a13786643fa6d4b@group.calendar.google.com';
const TZ = 'Asia/Taipei';
const MARK_START = '08:00';   // 行事曆標示起
const MARK_END = '09:00';     // 行事曆標示迄
const SHEET_REC = '出勤紀錄';
const SHEET_SET = '設定';
const HEADERS = ['ID', '登記時間', '員工名字', '假別', '服務工地', '年度', '週次', '休假日期起', '休假日期迄', '備註', '行事曆活動ID'];

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** 通行碼：未設定（預設）＝不需要；設定後，PWA 版每次呼叫都要帶正確通行碼。見下方 setAccessCode */
function keyOk_(p) {
  const code = PropertiesService.getScriptProperties().getProperty('ACCESS_CODE');
  if (!code) return true;
  return String(p.k || '') === code;
}

/** 一次性設定通行碼：把下面引號內的文字改成您要的通行碼，在編輯器選這個函式按「執行」。要取消通行碼，請改執行 clearAccessCode */
function setAccessCode() {
  PropertiesService.getScriptProperties().setProperty('ACCESS_CODE', '請改成您要的通行碼');
  console.log('通行碼已設定');
}
function clearAccessCode() {
  PropertiesService.getScriptProperties().deleteProperty('ACCESS_CODE');
  console.log('通行碼已取消');
}

/** 網址帶 ?api=… → 回傳 JSON（給 PWA 使用）；沒有 api 參數 → 照舊回傳網頁 */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.api) {
    try {
      if (!keyOk_(p)) return jsonOut_({ error: 'AUTH' });
      if (p.api === 'data') return jsonOut_(getData());
      if (p.api === 'submit') return jsonOut_({ records: submitLeave(JSON.parse(p.d)) });
      if (p.api === 'delete') return jsonOut_({ records: deleteLeave(String(p.id)) });
      return jsonOut_({ error: '未知的指令' });
    } catch (err) {
      return jsonOut_({ error: String(err && err.message || err) });
    }
  }
  setup_();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('虹揚廣告出勤表')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 第一次執行時自動建立工作表 */
function setup_() {
  const ss = SpreadsheetApp.getActive();
  let rec = ss.getSheetByName(SHEET_REC);
  if (!rec) {
    rec = ss.insertSheet(SHEET_REC);
    rec.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#fbeeee');
    rec.getRange('A:K').setNumberFormat('@'); // 全部存純文字，避免日期被自動轉換
    rec.setFrozenRows(1);
  }
  let set = ss.getSheetByName(SHEET_SET);
  if (!set) {
    set = ss.insertSheet(SHEET_SET);
    const rows = [
      ['員工名字（下方空白格可填入新員工）', '服務工地（下方空白格可填入新工地）'],
      ['楊明勳', '村聲'], ['盧冠鎮', '蘭田遇'], ['劉子廉', ''], ['謝蕙婷', ''],
      ['黃潔', ''], ['林芯如', ''], ['', ''], ['', ''], ['', '']
    ];
    set.getRange(1, 1, rows.length, 2).setValues(rows);
    set.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#fbeeee');
    set.setColumnWidths(1, 2, 260);
  }
}

const MONTHLY_NAMES = ['楊明勳', '盧冠鎮', '劉子廉'];   // 月薪制人員（第一次自動帶入「設定」工作表，之後以工作表為準，可自行修改）
const DEFAULT_SALARY = 36000;                           // 出勤制預設月薪

/** 「設定」工作表增加 C 欄「薪制」、D 欄「月薪金額」；已填的不會被覆蓋 */
function ensurePayCols_() {
  const set = SpreadsheetApp.getActive().getSheetByName(SHEET_SET);
  const lastRow = Math.max(set.getLastRow(), 2);
  const hdr = set.getRange(1, 3, 1, 2).getValues()[0];
  if (!hdr[0]) set.getRange(1, 3).setValue('薪制（月薪制／出勤制）');
  if (!hdr[1]) set.getRange(1, 4).setValue('月薪金額（出勤制用，可逐人修改）');
  set.getRange(1, 3, 1, 2).setFontWeight('bold').setBackground('#fbeeee');
  set.setColumnWidths(3, 2, 260);
  const names = set.getRange(2, 1, lastRow - 1, 1).getValues();
  const cur = set.getRange(2, 3, lastRow - 1, 2).getValues();
  let changed = false;
  names.forEach((r, i) => {
    const n = String(r[0]).trim();
    if (!n) return;
    if (!String(cur[i][0]).trim()) { cur[i][0] = MONTHLY_NAMES.indexOf(n) >= 0 ? '月薪制' : '出勤制'; changed = true; }
    if (String(cur[i][1]).trim() === '' && cur[i][0] === '出勤制') { cur[i][1] = DEFAULT_SALARY; changed = true; }
  });
  if (changed) set.getRange(2, 3, cur.length, 2).setValues(cur);
}

/** 網頁載入時取得選單、薪制設定與全部紀錄 */
function getData() {
  setup_();
  try { ensurePayCols_(); } catch (e) { console.warn('ensurePayCols_ 失敗：' + e); }
  const ss = SpreadsheetApp.getActive();
  const set = ss.getSheetByName(SHEET_SET).getDataRange().getValues().slice(1);
  const emps = set.map(r => String(r[0]).trim()).filter(Boolean);
  const sites = set.map(r => String(r[1]).trim()).filter(Boolean);
  const pay = {};
  set.forEach(r => {
    const n = String(r[0]).trim();
    if (!n) return;
    pay[n] = { type: String(r[2]).trim() === '月薪制' ? '月薪制' : '出勤制', salary: Number(r[3]) || 0 };
  });
  return { emps, sites, pay, records: readRecords_() };
}

function readRecords_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_REC);
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, HEADERS.length).getDisplayValues()
    .filter(r => r[0])
    .map(r => ({ id: r[0], created: r[1], emp: r[2], leave: r[3], site: r[4], year: +r[5], week: +r[6], start: r[7], end: r[8], note: r[9] }));
}

/** 新增一筆休假：寫入試算表並在公司行事曆逐日建立 8:00~9:00 活動（含半天休） */
function submitLeave(r) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(r.start) || !/^\d{4}-\d{2}-\d{2}$/.test(r.end) || r.end < r.start) throw new Error('日期格式錯誤');
    const cal = CalendarApp.getCalendarById(CALENDAR_ID);
    if (!cal) throw new Error('找不到公司行事曆，請確認部署帳號有此行事曆的「變更活動」權限');

    const eventIds = [];
    const title = r.emp + '（' + r.leave + '）';
    const desc = '虹揚廣告出勤登記\n員工：' + r.emp + '\n假別：' + r.leave + '\n服務工地：' + r.site + (r.note ? '\n備註：' + r.note : '');
    for (let d = r.start; d <= r.end; d = nextDay_(d)) {
      const s = Utilities.parseDate(d + ' ' + MARK_START, TZ, 'yyyy-MM-dd HH:mm');
      const e = Utilities.parseDate(d + ' ' + MARK_END, TZ, 'yyyy-MM-dd HH:mm');
      eventIds.push(cal.createEvent(title, s, e, { description: desc, location: r.site }).getId());
    }

    const id = Utilities.getUuid();
    const created = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
    SpreadsheetApp.getActive().getSheetByName(SHEET_REC).appendRow(
      [id, created, r.emp, r.leave, r.site, String(r.year), String(r.week), r.start, r.end, r.note || '', eventIds.join(',')]);
    return readRecords_();
  } finally {
    lock.releaseLock();
  }
}

/** 刪除一筆休假，同時移除行事曆活動 */
function deleteLeave(id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_REC);
    const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getDisplayValues().map(r => r[0]);
    const row = ids.indexOf(id) + 1;
    if (row > 1) {
      const evIds = String(sh.getRange(row, 11).getDisplayValue()).split(',').filter(Boolean);
      const cal = CalendarApp.getCalendarById(CALENDAR_ID);
      evIds.forEach(eid => { try { const ev = cal.getEventById(eid); if (ev) ev.deleteEvent(); } catch (e) {} });
      sh.deleteRow(row);
    }
    return readRecords_();
  } finally {
    lock.releaseLock();
  }
}

function nextDay_(s) {
  const d = new Date(s + 'T00:00:00Z');
  d.setUTCDate(d.getUTCDate() + 1);
  return d.toISOString().slice(0, 10);
}

/**
 * 一次性整理：把行事曆上「已登記」的休假活動，全部改成當日 8:00~9:00。
 * 使用方式：在編輯器上方選「convertExistingEvents」，按「執行」一次即可；到「執行記錄」看結果。
 * 可重複執行，已經是 8:00~9:00 的不會再動。
 */
function convertExistingEvents() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_REC);
  const cal = CalendarApp.getCalendarById(CALENDAR_ID);
  if (!cal) throw new Error('找不到公司行事曆，請確認帳號有此行事曆的「變更活動」權限');
  const n = sh.getLastRow() - 1;
  if (n < 1) { console.log('沒有任何登記紀錄'); return; }
  const ids = sh.getRange(2, 11, n, 1).getDisplayValues();
  let changed = 0, same = 0, missing = 0;
  ids.forEach(row => {
    String(row[0]).split(',').filter(Boolean).forEach(eid => {
      try {
        const ev = cal.getEventById(eid.trim());
        if (!ev) { missing++; return; }
        const d = Utilities.formatDate(ev.getStartTime(), TZ, 'yyyy-MM-dd');
        const s = Utilities.parseDate(d + ' ' + MARK_START, TZ, 'yyyy-MM-dd HH:mm');
        const e = Utilities.parseDate(d + ' ' + MARK_END, TZ, 'yyyy-MM-dd HH:mm');
        if (ev.getStartTime().getTime() === s.getTime() && ev.getEndTime().getTime() === e.getTime()) { same++; return; }
        ev.setTime(s, e);
        changed++;
      } catch (err) { missing++; console.warn('處理失敗 ' + eid + '：' + err); }
    });
  });
  console.log('已改為 8:00~9:00：' + changed + ' 筆｜原本就是：' + same + ' 筆｜找不到或失敗：' + missing + ' 筆');
}
