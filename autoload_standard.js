/* =====================================================================
 *  autoload_standard.js — ตัวเชื่อมข้อมูลเฉพาะห้อง Standard Laboratory
 *  อ่านชีต "Standard Lab" และ "Standard Test" จากไฟล์ สถิติ XX.xlsx
 *  แล้วแมปเข้าโครงสร้าง STD ของ Dashboard (จับข้อมูลตาม "หัวข้อ/label"
 *  ไม่ใช่เลขแถวตายตัว จึงทนต่อการเพิ่ม/ย้ายแถวในไฟล์)
 *  ตรวจสอบแล้วว่าให้ผลตรงกับ STD69 เดิมครบทุกช่อง
 * ===================================================================== */
(function () {
  "use strict";

  var CURRENT_YEAR = 2569; // ไฟล์ สถิติ 69.xlsx = ปีปฏิทิน 2569

  function parse(wb, H) {
    var norm = H.norm, num = H.num, findRow = H.findRow;
    var ABBR = H.THAI_ABBR, FULL = H.THAI_FULL;
    var out = {};

    // ---------------- Standard Test ----------------
    var gt = H.sheetGrid(wb, "Standard Test", window.XLSX);
    var hdr = findRow(gt, function (row) {
      var hasABO = row.some(function (x) { return norm(x) === "abo"; });
      var hasXM = row.some(function (x) { return norm(x) === "x-match"; });
      return hasABO && hasXM;
    });
    var hrow = gt[hdr] || [];
    function colOf(name) {
      for (var j = 0; j < hrow.length; j++) if (norm(hrow[j]) === norm(name)) return j;
      return -1;
    }
    var testCols = {
      ABO: colOf("ABO"), Rh: colOf("Rh"), DAT: colOf("DAT"), IAT: colOf("IAT"),
      "IAT AB/cells": colOf("IAT AB/cells"), "X-match": colOf("X-match"), Iden: colOf("Iden")
    };
    var test = {}; Object.keys(testCols).forEach(function (k) { test[k] = newArr(); });
    gt.forEach(function (row) {
      var mi = FULL.indexOf((row[0] === null || row[0] === undefined) ? "" : String(row[0]).trim());
      if (mi >= 0) {
        Object.keys(testCols).forEach(function (k) {
          var c = testCols[k];
          if (c >= 0 && c < row.length) test[k][mi] = num(row[c]);
        });
      }
    });
    var tmap = { test_abo: "ABO", test_rh: "Rh", test_dat: "DAT", test_iat: "IAT",
      test_iat_ab_cells: "IAT AB/cells", test_xmatch: "X-match", test_iden: "Iden" };
    Object.keys(tmap).forEach(function (dk) { out[dk] = test[tmap[dk]]; });

    // ---------------- Standard Lab ----------------
    var gl = H.sheetGrid(wb, "Standard Lab", window.XLSX);
    var mh = findRow(gl, function (row) {
      return row.some(function (x) { return norm(x) === "ม.ค."; }) &&
             row.some(function (x) { return norm(x) === "ธ.ค."; });
    });
    var monthCols = [];
    ABBR.forEach(function (ab) {
      var hrow2 = gl[mh] || [];
      for (var j = 0; j < hrow2.length; j++) { if (norm(hrow2[j]) === norm(ab)) { monthCols.push(j); break; } }
    });
    function valsAt(i) {
      var row = gl[i] || [];
      return monthCols.map(function (c) { return c < row.length ? num(row[c]) : null; });
    }
    function findHeader(sub, start, exc) {
      for (var i = (start || 0); i < gl.length; i++) {
        var lab = (gl[i] && gl[i][0] != null) ? String(gl[i][0]) : "";
        if (lab.toLowerCase().indexOf(sub.toLowerCase()) >= 0 &&
            (!exc || lab.toLowerCase().indexOf(exc.toLowerCase()) < 0)) return i;
      }
      return -1;
    }
    var s1 = findHeader("Coombs");
    var s2 = findHeader("ANC", s1 + 1);
    var s3 = -1;
    for (var i = s2 + 1; i < gl.length; i++) {
      var lab = (gl[i] && gl[i][0] != null) ? String(gl[i][0]) : "";
      if (lab.toLowerCase().indexOf("blood group") >= 0 &&
          lab.toLowerCase().indexOf("coombs") < 0 && /^\s*\d+\./.test(lab)) { s3 = i; break; }
    }
    // ส่วนที่ 4: Blood group (หน่วยงานเปลี่ยนอวัยวะ) — หัวข้อมีคำว่า "อวัยวะ"
    var sOrg = -1;
    for (var io = s3 + 1; io < gl.length; io++) {
      var labO = (gl[io] && gl[io][0] != null) ? String(gl[io][0]) : "";
      if (labO.indexOf("อวัยวะ") >= 0) { sOrg = io; break; }
    }
    var s4 = findHeader("Standard cells");
    var bounds = [s1, s2, s3, sOrg, s4, gl.length].filter(function (x) { return x >= 0; }).sort(function (a, b) { return a - b; });
    function rng(s) {
      var nxt = gl.length;
      bounds.forEach(function (b) { if (b > s && b < nxt) nxt = b; });
      return [s, nxt];
    }
    function matchIn(r, pred) {
      for (var k = r[0]; k < r[1]; k++) {
        if (pred(norm(gl[k] && gl[k][0]))) return valsAt(k);
      }
      return newArr();
    }
    var r1 = rng(s1), r2 = rng(s2), r3 = rng(s3), rOrg = (sOrg >= 0 ? rng(sOrg) : [gl.length, gl.length]), r4 = rng(s4);
    function P() { var a = [].slice.call(arguments); return function (l) { return a.some(function (s) { return l.indexOf(norm(s)) === 0; }); }; }
    function EQ(s) { return function (l) { return l === norm(s); }; }
    function C(s) { return function (l) { return l.indexOf(norm(s)) >= 0; }; }

    var lm = {
      lab_abo_rh_hdfn: [r1, EQ("ABO, Rh")], lab_dat: [r1, EQ("DAT")], lab_iat_abo: [r1, EQ("IAT-ABO")],
      lab_iat_ab_screen: [r1, C("screen")], lab_xmatch_hdfn: [r1, P("X-match")], lab_ab_iden_hdfn: [r1, P("Ab-Ident")],
      lab_exchange: [r1, P("Exchange")], lab_total_case_hdfn: [r1, P("Total case")], lab_intra_uterine: [r1, P("Intra")],
      lab_anc_abo_rh: [r2, P("ABO, Rh")], lab_anc_rh_ve: [r2, P("Rh-ve")], lab_anc_rh_del: [r2, EQ("RhDel")], lab_anc_weak_d: [r2, C("Weak D")], lab_anc_ab_iden: [r2, P("Ab Ident", "Ab-Ident")],
      lab_bg_abo: [r3, EQ("ABO")], lab_bg_rh: [r3, EQ("Rh")], lab_bg_rh_del: [r3, EQ("RhDel")], lab_bg_weak_d: [r3, C("Weak D")], lab_bg_ab_screen: [r3, P("Ab screen")], lab_bg_ab_iden: [r3, P("Ab-Ident")],
      lab_org_abo: [rOrg, EQ("ABO")], lab_org_rh: [rOrg, EQ("Rh")], lab_org_rh_del: [rOrg, EQ("RhDel")], lab_org_weak_d: [rOrg, C("Weak D")], lab_org_ab_screen: [rOrg, P("Ab screen")], lab_org_ab_iden: [rOrg, P("Ab-Ident")],
      lab_cell_alsever: [r4, P("Alsever")], lab_cell_5_glycerol: [r4, P("5%")], lab_cell_12_glycerol: [r4, P("12%")], lab_cell_40_glycerol: [r4, P("40%")],
      lab_cell_ph_6_0: [r4, C("6.0")], lab_cell_ph_5_4: [r4, C("5.4")], lab_cell_1_papain: [r4, P("1%")], lab_cell_abo: [r4, P("ABO Cell")],
      lab_cell_coombs_ctrl: [r4, C("Control")], lab_cell_screen_o3: [r4, C("O3")], lab_cell_panel11: [r4, P("Panel")], lab_cell_papain_panel11: [r4, C("Papainized")], lab_cell_diluent2: [r4, P("Diluent")]
    };
    Object.keys(lm).forEach(function (k) { out[k] = matchIn(lm[k][0], lm[k][1]); });
    return out;

    function newArr() { return [null, null, null, null, null, null, null, null, null, null, null, null]; }
  }

  // คำนวณจำนวนเดือนที่มีข้อมูลจริง -> [0..max]
  function computeMonths(data) {
    var maxIdx = -1;
    Object.keys(data).forEach(function (k) {
      var arr = data[k];
      if (!Array.isArray(arr)) return;
      for (var i = 0; i < arr.length; i++) if (arr[i] !== null && arr[i] !== undefined) { if (i > maxIdx) maxIdx = i; }
    });
    if (maxIdx < 0) return null;
    var out = []; for (var j = 0; j <= maxIdx; j++) out.push(j); return out;
  }

  // ชื่อไฟล์ "สถิติ NN.xlsx" (ปีปฏิทิน) -> พ.ศ. 25NN  (ข้าม "ปีงบประมาณ"/"(Autosaved)")
  function matchYear(filename) {
    var m = String(filename).match(/^สถิติ\s*(\d{2})\.xlsx$/i);
    if (!m) return null;
    var year = 2500 + parseInt(m[1], 10);
    return (year >= 2564 && year <= 2569) ? year : null;
  }

  // Dashboard เวอร์ชันใหม่อ่านข้อมูลจาก window.__LIVE__.standard โดยตรง (ตัวแปร DATA)
  // จึงเขียนข้อมูลปีนั้นเข้า __LIVE__.standard แบบ "merge" (คงฟิลด์ที่ตัวอ่านไม่ได้สร้างไว้)
  function applyYear(year, data) {
    var key = String(year);
    window.__LIVE__ = window.__LIVE__ || {};
    window.__LIVE__.standard = window.__LIVE__.standard || {};
    var obj = window.__LIVE__.standard[key] || (window.__LIVE__.standard[key] = {});
    Object.keys(data).forEach(function (k) { obj[k] = data[k]; });
  }

  // เรนเดอร์ใหม่หนึ่งครั้งหลังอัพเดทครบทุกไฟล์ที่เปลี่ยน (ใช้ฟังก์ชัน render() ของหน้า)
  function render() {
    if (typeof window.render === "function") window.render();
  }

  DashboardAutoload.register({
    mode: "directory",
    storeKey: "standard",        // คีย์สำหรับบันทึกข้อมูลถาวร
    snapshotFile: "live_standard.js",  // เขียนข้อมูลล่าสุดลงไฟล์นี้
    defaultYear: CURRENT_YEAR,   // ใช้กับเบราว์เซอร์ที่ต้องเลือกไฟล์เดียว
    match: matchYear,
    parse: parse,
    applyYear: applyYear,
    render: render
  });
})();
