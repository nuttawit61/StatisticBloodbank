/* =====================================================================
 *  autoload_luminex.js — ตัวเชื่อมข้อมูลเฉพาะห้อง DNA Luminex
 *  อ่านชีต "Luminex" จากไฟล์ สถิติ 69.xlsx (แต่ละเดือน = 3 คอลัมน์:
 *  ศิริราช / ร.พ.อื่น ๆ / EQAS) แล้วแมปเข้าโครงสร้าง LUM_DATA[2569] ของ Dashboard
 *  - ค่ารวมต่อเดือน = ศิริราช + ร.พ.อื่น ๆ
 *  - src_siriraj/src_other = แยกแหล่งของแถว Total
 *  - scr_siriraj/scr_other = แยกแหล่งของแถว HLA Antibody Screening
 *  อัพเดทเฉพาะปีปัจจุบัน (2569) — ปีเก่าโครงสร้างต่างกัน คงข้อมูลเดิมไว้
 *  ตรวจสอบแล้ว: ฟิลด์ผลรวมตรงกับ LUM_2569 เดิมครบทุกช่อง
 * ===================================================================== */
(function () {
  "use strict";
  var CURRENT_YEAR = 2569;

  function tonum(v) {
    if (v === null || v === undefined || v === "") return null;
    if (typeof v === "number") return v;
    var m = String(v).match(/-?\d+(?:\.\d+)?/);
    return m ? Number(m[0]) : null;
  }

  function parse(wb, H) {
    var norm = H.norm, XLSX = window.XLSX, out = {};
    if (!wb.Sheets["Luminex"]) return out;
    var g = H.sheetGrid(wb, "Luminex", XLSX);
    var hr = H.findRow(g, function (row) { return row.some(function (x) { return norm(x) === "ม.ค."; }); });
    if (hr < 0) return out;
    var base = -1;
    for (var j = 0; j < g[hr].length; j++) { if (norm(g[hr][j]) === "ม.ค.") { base = j; break; } }
    if (base < 0) return out;

    function sir(row, m) { var c = base + 3 * m; return c < row.length ? tonum(row[c]) : null; }
    function oth(row, m) { var c = base + 3 * m + 1; return c < row.length ? tonum(row[c]) : null; }
    function summ(row) {
      var o = [];
      for (var m = 0; m < 12; m++) {
        var a = sir(row, m), b = oth(row, m);
        o.push((a === null && b === null) ? null : (a || 0) + (b || 0));
      }
      return o;
    }
    function pick(row, fn) { var o = []; for (var m = 0; m < 12; m++) o.push(fn(row, m)); return o; }
    function find(pred) {
      for (var i = 0; i < g.length; i++) {
        var lab = (g[i] && g[i][0] != null) ? norm(g[i][0]) : "";
        if (lab && pred(lab)) return g[i];
      }
      return null;
    }
    function set(key, row, fn) { if (row) out[key] = fn(row); }

    var rScreen = find(function (n) { return n.indexOf("antibodyscreening") >= 0; });
    set("screening", rScreen, summ);
    set("scr_siriraj", rScreen, function (r) { return pick(r, sir); });
    set("scr_other", rScreen, function (r) { return pick(r, oth); });

    set("pra1", find(function (n) { return n.indexOf("specificprahlaclassi") >= 0 && n.indexOf("classii") < 0; }), summ);
    set("pra2", find(function (n) { return n.indexOf("specificprahlaclassii") >= 0; }), summ);
    set("sa1", find(function (n) { return n.indexOf("singleantigenhlaclassi") >= 0 && n.indexOf("classii") < 0 && n.indexOf("c1q") < 0; }), summ);
    set("sa2", find(function (n) { return n.indexOf("singleantigenhlaclassii") >= 0 && n.indexOf("c1q") < 0; }), summ);
    set("sa1_c1q", find(function (n) { return n.indexOf("singleantigenhlaclassi") >= 0 && n.indexOf("classii") < 0 && n.indexOf("c1q") >= 0; }), summ);
    set("sa2_c1q", find(function (n) { return n.indexOf("singleantigenhlaclassii") >= 0 && n.indexOf("c1q") >= 0; }), summ);
    set("mica_geno", find(function (n) { return n.indexOf("micagenotype") >= 0; }), summ);
    set("mica_ab", find(function (n) { return n.indexOf("micaantibody") >= 0; }), summ);
    set("trali", find(function (n) { return n.indexOf("trali") >= 0; }), summ);

    var rTotal = find(function (n) { return n === "total"; });
    set("total", rTotal, summ);
    set("src_siriraj", rTotal, function (r) { return pick(r, sir); });
    set("src_other", rTotal, function (r) { return pick(r, oth); });
    return out;
  }

  function matchYear(filename) { return /^สถิติ\s*69\.xlsx$/i.test(String(filename)) ? CURRENT_YEAR : null; }

  // Dashboard เก็บข้อมูลใน LUM_DATA (เปิดเป็น window.LUM_DATA) — เขียนแบบ merge in-place
  // เพื่อให้ตัวแปร LUM (ที่อ้างถึง object เดียวกัน) เห็นค่าล่าสุดทันที
  function applyYear(year, data) {
    if (year !== CURRENT_YEAR || !window.LUM_DATA) return;
    var obj = window.LUM_DATA[year] || (window.LUM_DATA[year] = {});
    Object.keys(data).forEach(function (k) { obj[k] = data[k]; });
  }

  function render() { if (typeof window.refresh === "function") window.refresh(); }

  DashboardAutoload.register({
    mode: "directory",
    storeKey: "luminex",
    snapshotFile: "live_luminex.js",
    defaultYear: CURRENT_YEAR,
    match: matchYear,
    parse: parse,
    applyYear: applyYear,
    render: render
  });
})();
