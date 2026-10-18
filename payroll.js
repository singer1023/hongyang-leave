/* 虹揚廣告 出勤／薪資請款表 輸出（ExcelJS）
 * 規則集中在 RULES，要改計薪方式只需改這裡。
 */
var RULES = {
  DAYS_DIVISOR: 30,                 // 日薪 = 月薪 ÷ 30
  DEFAULT_SALARY: 36000,            // 出勤制預設月薪（可在 Google 試算表「設定」工作表 D 欄逐人修改）
  PAID_REST_PER_WEEK: 1,            // 每週給薪的例行休假天數（超過的部分扣薪）；設很大的數字（如 99）等於不限制
  REST_LEAVES: ["例行休假", "例行上休", "例行下休"],
  PAID_LEAVES: ["特休", "婚假", "喪假", "病假", "例行休假", "例行上休", "例行下休"],   // 有薪（照算薪水）
  UNPAID_LEAVES: ["事假", "產假"]   // 無薪（每天扣日薪）
};
var HALF_UNITS = { "例行上休": 0.5, "例行下休": 0.5 };

(function (root) {
  var FONT = "微軟正黑體";
  var GREEN = "FFE2EFDA", WHITE = "FFFFFFFF", HEAD = "FFB0191E", SOFT = "FFFBEEEE", TAB = "FFC6E0B4";
  var WD = ["日", "一", "二", "三", "四", "五", "六"];
  var pad = function (n) { return String(n).padStart(2, "0"); };
  var iso = function (d) { return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate()); };
  var parse = function (s) { var p = s.split("-").map(Number); return new Date(p[0], p[1] - 1, p[2]); };
  var addDays = function (d, n) { return new Date(d.getFullYear(), d.getMonth(), d.getDate() + n); };
  var unit = function (l) { return HALF_UNITS[l] || 1; };
  var rocShort = function (s) { var p = s.split("-").map(Number); return (p[0] - 1911) + "/" + p[1] + "/" + p[2]; };
  var fmt = function (n) { return Math.round(n * 100) / 100; };
  var money = function (n) { return Math.round(n); };

  function mondayKey(d) { return iso(addDays(d, -((d.getDay() + 6) % 7))); }

  /** 展開某員工某月的每日請假 → [{ds, leave, u}] */
  function dayList(records, emp, ym) {
    var y = +ym.split("-")[0], m = +ym.split("-")[1];
    var ms = iso(new Date(y, m - 1, 1)), me = iso(new Date(y, m, 0)), out = [];
    records.filter(function (r) { return r.emp === emp; }).forEach(function (r) {
      var s = r.start > ms ? r.start : ms, e = r.end < me ? r.end : me;
      if (s > e) return;
      for (var d = parse(s); iso(d) <= e; d = addDays(d, 1)) out.push({ ds: iso(d), leave: r.leave, u: unit(r.leave) });
    });
    return out;
  }

  /** 計算一位出勤制員工的薪資 */
  function calcHourly(name, salary, days) {
    var by = {}, weeks = {};
    days.forEach(function (x) {
      by[x.leave] = (by[x.leave] || 0) + x.u;
      if (RULES.REST_LEAVES.indexOf(x.leave) >= 0) { var k = mondayKey(parse(x.ds)); weeks[k] = (weeks[k] || 0) + x.u; }
    });
    var excess = 0;
    Object.keys(weeks).forEach(function (k) { excess += Math.max(0, weeks[k] - RULES.PAID_REST_PER_WEEK); });
    var g = function (a) { return a.reduce(function (s, l) { return s + (by[l] || 0); }, 0); };
    var rest = g(RULES.REST_LEAVES);
    var unpaid = g(RULES.UNPAID_LEAVES);
    var daily = salary / RULES.DAYS_DIVISOR;
    var deductDays = unpaid + excess;
    var deduct = money(daily * deductDays);
    return {
      name: name, salary: salary, daily: daily,
      annual: g(["特休"]), wedfun: g(["婚假", "喪假"]), rest: rest, sick: g(["病假"]),
      personal: g(["事假"]), maternity: g(["產假"]), excess: excess,
      deductDays: deductDays, deduct: deduct, pay: money(salary) - deduct
    };
  }

  function style(c, o) {
    o = o || {};
    c.font = { name: FONT, size: o.size || 11, bold: !!o.bold, color: { argb: o.color || "FF222222" } };
    c.alignment = { vertical: "middle", horizontal: o.h || "center", wrapText: o.wrap !== false };
    if (o.fill) c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: o.fill } };
    if (o.border) { var t = { style: "thin", color: { argb: "FF999999" } }; c.border = { top: t, left: t, bottom: t, right: t }; }
    if (o.fmt) c.numFmt = o.fmt;
  }

  function fillRest(ws, from, to, cols) {   // 把無資料部分網底設為全白
    for (var r = from; r <= to; r++) for (var c = 1; c <= cols; c++) {
      var cell = ws.getCell(r, c);
      if (!cell.fill || cell.fill.type !== "pattern") cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: WHITE } };
    }
  }

  /** ctx: {ym:"2026-10", emps:[], pay:{name:{type,salary}}, records:[], today:"2026-10-31"} */
  function build(ExcelJS, ctx) {
    var ym = ctx.ym, y = +ym.split("-")[0], m = +ym.split("-")[1], title = "虹揚廣告有限公司　" + (y - 1911) + "年" + m + "月";
    var records = ctx.records.filter(function (r) { return dayList([r], r.emp, ym).length > 0; });
    var names = ctx.emps.slice(); ctx.records.forEach(function (r) { if (names.indexOf(r.emp) < 0) names.push(r.emp); });
    var payOf = function (n) { var p = (ctx.pay || {})[n] || {}; return { type: p.type === "月薪制" ? "月薪制" : "出勤制", salary: +p.salary || 0 }; };
    var monthly = names.filter(function (n) { return payOf(n).type === "月薪制"; });
    var hourly = names.filter(function (n) { return payOf(n).type === "出勤制"; });

    var wb = new ExcelJS.Workbook(); wb.creator = "虹揚廣告"; wb.created = new Date();

    /* ========= 工作表1：薪資請款表（列印用） ========= */
    var ws = wb.addWorksheet("薪資請款表", { properties: { tabColor: { argb: TAB } }, views: [{ showGridLines: false }] });
    var widths = [11, 11, 11, 8, 8, 9, 8, 11, 11, 11, 9, 11, 13, 14];
    widths.forEach(function (w, i) { ws.getColumn(i + 1).width = w; });
    var COLS = 14;
    ws.mergeCells(1, 1, 1, COLS); ws.getCell(1, 1).value = title + "　薪資請款表";
    style(ws.getCell(1, 1), { size: 20, bold: true, color: HEAD }); ws.getRow(1).height = 36;
    ws.mergeCells(2, 1, 2, COLS);
    ws.getCell(2, 1).value = "製表日期：" + (ctx.today ? rocShort(ctx.today) : "") + "　｜　計薪方式：出勤制 日薪＝月薪÷" + RULES.DAYS_DIVISOR + "；事假、產假無薪；每週例行休假超過 " + RULES.PAID_REST_PER_WEEK + " 天者，超出部分扣薪";
    style(ws.getCell(2, 1), { size: 10, h: "left", color: "FF666666" }); ws.getRow(2).height = 22;

    var r = 4;
    ws.mergeCells(r, 1, r, COLS); ws.getCell(r, 1).value = "壹、月薪制人員（依標準薪資表，不扣薪）";
    style(ws.getCell(r, 1), { size: 13, bold: true, h: "left", color: HEAD }); ws.getRow(r).height = 24;
    r++;
    var h1 = [[1, 3, "員工"], [4, 11, "本月請假紀錄（僅供參考，不影響薪資）"], [12, 12, "扣薪"], [13, 13, "應發薪資"], [14, 14, "簽名"]];
    h1.forEach(function (x) { if (x[1] > x[0]) ws.mergeCells(r, x[0], r, x[1]); ws.getCell(r, x[0]).value = x[2]; });
    for (var c = 1; c <= COLS; c++) style(ws.getCell(r, c), { bold: true, color: "FFFFFFFF", fill: HEAD, border: true });
    ws.getRow(r).height = 24; r++;
    var summary = monthly.map(function (n) {
      var d = dayList(ctx.records, n, ym), by = {};
      d.forEach(function (x) { by[x.leave] = (by[x.leave] || 0) + x.u; });
      var parts = Object.keys(by).map(function (k) { return k + " " + fmt(by[k]) + " 天"; });
      return n + "：" + (parts.length ? parts.join("、") : "無請假");
    }).join("\n");
    ws.mergeCells(r, 1, r, 3); ws.mergeCells(r, 4, r, 11);
    ws.getCell(r, 1).value = monthly.length ? monthly.join("／") : "（無）";
    ws.getCell(r, 4).value = monthly.length ? summary : "";
    ws.getCell(r, 12).value = "不扣薪"; ws.getCell(r, 13).value = "依標準薪資表"; ws.getCell(r, 14).value = "";
    for (c = 1; c <= COLS; c++) style(ws.getCell(r, c), { fill: GREEN, border: true, h: c === 4 ? "left" : "center", bold: c === 1 });
    ws.getRow(r).height = Math.max(40, 22 * Math.max(1, monthly.length)); r += 2;

    ws.mergeCells(r, 1, r, COLS); ws.getCell(r, 1).value = "貳、依實際出勤計算薪水人員";
    style(ws.getCell(r, 1), { size: 13, bold: true, h: "left", color: HEAD }); ws.getRow(r).height = 24; r++;
    var heads = ["員工", "月薪", "日薪", "特休", "婚喪假", "例行休假", "病假", "事假\n（無薪）", "產假\n（無薪）", "超休\n（扣薪）", "扣薪天數", "扣薪金額", "應發薪資", "簽名"];
    heads.forEach(function (t, i) { ws.getCell(r, i + 1).value = t; style(ws.getCell(r, i + 1), { bold: true, color: "FFFFFFFF", fill: HEAD, border: true }); });
    ws.getRow(r).height = 36; var headRow = r; r++;
    var first = r, rows = hourly.map(function (n) { return calcHourly(n, payOf(n).salary || RULES.DEFAULT_SALARY, dayList(ctx.records, n, ym)); });
    rows.forEach(function (x, i) {
      var vals = [x.name, x.salary, x.daily, fmt(x.annual), fmt(x.wedfun), fmt(x.rest), fmt(x.sick), fmt(x.personal), fmt(x.maternity), fmt(x.excess), null, null, null, ""];
      vals.forEach(function (v, j) { if (v !== null) ws.getCell(r, j + 1).value = v; });
      ws.getCell(r, 3).value = { formula: "ROUND(B" + r + "/" + RULES.DAYS_DIVISOR + ",2)", result: fmt(x.daily) };
      ws.getCell(r, 11).value = { formula: "H" + r + "+I" + r + "+J" + r, result: fmt(x.deductDays) };
      ws.getCell(r, 12).value = { formula: "ROUND(C" + r + "*K" + r + ",0)", result: x.deduct };
      ws.getCell(r, 13).value = { formula: "B" + r + "-L" + r, result: x.pay };
      for (c = 1; c <= COLS; c++) style(ws.getCell(r, c), { fill: i % 2 === 0 ? GREEN : WHITE, border: true, bold: c === 1 || c === 13, fmt: (c === 2 || c === 3 || c === 12 || c === 13) ? "#,##0" : (c >= 4 && c <= 11 ? "0.0;-0.0;;@" : undefined) });
      ws.getRow(r).height = 30; r++;
    });
    if (!rows.length) { ws.mergeCells(r, 1, r, COLS); ws.getCell(r, 1).value = "（本月無出勤制人員）"; style(ws.getCell(r, 1), { border: true }); r++; }
    var last = r - 1;
    ws.getCell(r, 1).value = "合計";
    ws.mergeCells(r, 1, r, 1);
    var sums = { 2: "B", 12: "L", 13: "M" };
    Object.keys(sums).forEach(function (k) {
      var L = sums[k], tot = rows.reduce(function (s, x) { return s + (k == 2 ? x.salary : k == 12 ? x.deduct : x.pay); }, 0);
      ws.getCell(r, +k).value = rows.length ? { formula: "SUBTOTAL(109," + L + first + ":" + L + last + ")", result: tot } : 0;
    });
    for (c = 1; c <= COLS; c++) style(ws.getCell(r, c), { bold: true, fill: SOFT, border: true, fmt: (c === 2 || c === 12 || c === 13) ? "#,##0" : undefined });
    ws.getRow(r).height = 28; var totalRow = r; r += 2;

    ws.mergeCells(r, 1, r, COLS);
    ws.getCell(r, 1).value = "備註：①有薪假別（照算薪水）：特休、婚假、喪假、病假、例行休假（含上休／下休各 0.5 天）。②無薪假別：事假、產假，每天扣一日薪。③每週（週一至週日）例行休假超過 " + RULES.PAID_REST_PER_WEEK + " 天，超出部分列入「超休」扣薪。④日薪、扣薪金額已設公式，月薪可直接修改。";
    style(ws.getCell(r, 1), { size: 10, h: "left", color: "FF555555" }); ws.getRow(r).height = 44; r += 2;
    ["製表", "覆核", "核准"].forEach(function (t, i) {
      var c0 = 1 + i * 5; ws.mergeCells(r, c0, r, c0 + 3); ws.getCell(r, c0).value = t + "：＿＿＿＿＿＿＿＿＿";
      style(ws.getCell(r, c0), { h: "left", size: 12 });
    });
    ws.getRow(r).height = 34; var endRow = r;
    fillRest(ws, 1, endRow + 2, COLS);
    ws.pageSetup = { paperSize: 9, orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 1, horizontalCentered: true,
      margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 }, printArea: "A1:N" + (endRow + 1) };
    ws.headerFooter = { oddFooter: "&L&\"微軟正黑體\"虹揚廣告有限公司&R&\"微軟正黑體\"第 &P 頁" };

    /* ========= 工作表2：請假明細 ========= */
    var w2 = wb.addWorksheet("請假明細", { views: [{ showGridLines: false, state: "frozen", ySplit: 3 }] });
    var heads2 = ["員工", "薪制", "假別", "服務工地", "週次", "休假日期（起）", "休假日期（迄）", "本月天數", "有薪／無薪", "備註", "登記時間"];
    [10, 9, 10, 12, 14, 17, 17, 10, 11, 22, 17].forEach(function (w, i) { w2.getColumn(i + 1).width = w; });
    w2.mergeCells(1, 1, 1, heads2.length); w2.getCell(1, 1).value = title + "　請假明細";
    style(w2.getCell(1, 1), { size: 16, bold: true, color: HEAD, h: "left" }); w2.getRow(1).height = 30;
    heads2.forEach(function (t, i) { w2.getCell(3, i + 1).value = t; style(w2.getCell(3, i + 1), { bold: true, color: "FFFFFFFF", fill: HEAD, border: true }); });
    w2.getRow(3).height = 26;
    var sorted = records.slice().sort(function (a, b) { return a.emp.localeCompare(b.emp, "zh-Hant") || a.start.localeCompare(b.start); });
    var rr = 4;
    sorted.forEach(function (x, i) {
      var days = dayList([x], x.emp, ym).reduce(function (s, d) { return s + d.u; }, 0);
      var vals = [x.emp, payOf(x.emp).type, x.leave, x.site, x.year - 1911 + "年第" + x.week + "週", rocShort(x.start), rocShort(x.end), days,
        RULES.PAID_LEAVES.indexOf(x.leave) >= 0 ? "有薪" : "無薪", x.note || "", x.created || ""];
      vals.forEach(function (v, j) { w2.getCell(rr, j + 1).value = v; style(w2.getCell(rr, j + 1), { fill: i % 2 === 0 ? GREEN : WHITE, border: true, h: j === 9 ? "left" : "center" }); });
      rr++;
    });
    if (!sorted.length) { w2.mergeCells(rr, 1, rr, heads2.length); w2.getCell(rr, 1).value = "本月尚無登記"; style(w2.getCell(rr, 1), { border: true }); rr++; }
    for (c = 1; c <= heads2.length; c++) style(w2.getCell(rr, c), { bold: true, fill: SOFT, border: true });
    w2.getCell(rr, 1).value = "合計";
    w2.getCell(rr, 8).value = sorted.length ? { formula: "SUBTOTAL(109,H4:H" + (rr - 1) + ")", result: sorted.reduce(function (s, x) { return s + dayList([x], x.emp, ym).reduce(function (a, d) { return a + d.u; }, 0); }, 0) } : 0;
    w2.getCell(rr, 3).value = { formula: "SUBTOTAL(103,A4:A" + (rr - 1) + ")", result: sorted.length }; w2.getCell(rr, 4).value = "筆";
    if (sorted.length) w2.autoFilter = { from: { row: 3, column: 1 }, to: { row: rr - 1, column: heads2.length } };
    fillRest(w2, 1, rr + 20, heads2.length + 3);
    w2.pageSetup = { paperSize: 9, orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0, horizontalCentered: true, printTitlesRow: "3:3",
      margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 } };

    /* ========= 工作表3：每日總覽 ========= */
    var w3 = wb.addWorksheet("每日總覽", { properties: { tabColor: { argb: TAB } }, views: [{ showGridLines: false, state: "frozen", xSplit: 1, ySplit: 4 }] });
    var lastDay = new Date(y, m, 0).getDate();
    w3.getColumn(1).width = 11; for (var d = 1; d <= lastDay; d++) w3.getColumn(d + 1).width = 7;
    w3.mergeCells(1, 1, 1, lastDay + 1); w3.getCell(1, 1).value = title + "　每日總覽";
    style(w3.getCell(1, 1), { size: 16, bold: true, color: HEAD, h: "left" }); w3.getRow(1).height = 30;
    w3.getCell(3, 1).value = "員工"; w3.getCell(4, 1).value = "";
    for (d = 1; d <= lastDay; d++) { w3.getCell(3, d + 1).value = d; w3.getCell(4, d + 1).value = WD[new Date(y, m - 1, d).getDay()]; }
    for (c = 1; c <= lastDay + 1; c++) { style(w3.getCell(3, c), { bold: true, color: "FFFFFFFF", fill: HEAD, border: true }); style(w3.getCell(4, c), { bold: true, color: "FFFFFFFF", fill: HEAD, border: true }); }
    names.forEach(function (n, i) {
      var row = 5 + i, dl = dayList(ctx.records, n, ym), map = {};
      dl.forEach(function (x) { map[x.ds] = (map[x.ds] ? map[x.ds] + "/" : "") + x.leave.replace("例行", "例"); });
      w3.getCell(row, 1).value = n; style(w3.getCell(row, 1), { bold: true, fill: i % 2 === 0 ? GREEN : WHITE, border: true });
      for (d = 1; d <= lastDay; d++) {
        var ds = iso(new Date(y, m - 1, d)), wk = new Date(y, m - 1, d).getDay();
        var cell = w3.getCell(row, d + 1); cell.value = map[ds] || "";
        style(cell, { size: 9, border: true, fill: map[ds] ? "FFFFF2CC" : (i % 2 === 0 ? GREEN : WHITE), color: wk === 0 || wk === 6 ? "FFB0191E" : "FF222222" });
      }
      w3.getRow(row).height = 30;
    });
    fillRest(w3, 1, 5 + names.length + 20, lastDay + 4);
    w3.pageSetup = { paperSize: 9, orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0, horizontalCentered: true,
      margins: { left: 0.3, right: 0.3, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 } };

    return { wb: wb, rows: rows, monthly: monthly, hourly: hourly, totalRow: totalRow };
  }

  root.PayrollBuild = { build: build, calcHourly: calcHourly, dayList: dayList };
  if (typeof module !== "undefined" && module.exports) module.exports = root.PayrollBuild;
})(typeof window !== "undefined" ? window : globalThis);
