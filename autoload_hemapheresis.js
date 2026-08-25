/* =====================================================================
 *  autoload_hemapheresis.js — ตัวเชื่อมข้อมูลเฉพาะห้อง Hemapheresis
 *  อ่านชีต "Hemapheresis" จากไฟล์ สถิติ XX.xlsx แล้วแมปเข้าโครงสร้าง D ของ Dashboard
 *
 *  หมายเหตุสำคัญ: โครงสร้างชีตห้องนี้ "เปลี่ยนไปในแต่ละปี" (ปีเก่าไม่มี Trima/Amicus,
 *  เกณฑ์ QC คนละแบบ ฯลฯ) และข้อมูลปีเก่าในหน้าถูกต้องอยู่แล้ว
 *  ดังนั้นระบบนี้จะ "อัพเดทเฉพาะปีปัจจุบัน (2569)" ส่วนปีเก่าคงข้อมูลเดิมไว้
 *  - จับ QC ตาม "ตำแหน่ง" (ไม่ผูกกับตัวเลขเกณฑ์) และแยก ครั้ง/ราย ตามหัวข้อ
 *  - ตรวจสอบแล้ว: อ่าน สถิติ 69.xlsx ครบ 28 ฟิลด์ และตรง 3 เดือนแรกกับข้อมูลเดิม
 *    (ที่ต่างคือ Excel มีเดือน เม.ย./พ.ค. เพิ่ม และช่องว่างในหน้าเดิม = 0 จริงในไฟล์)
 * ===================================================================== */
(function () {
  "use strict";

  var CURRENT_YEAR = 2569;

  function toNum(v) {
    if (v === null || v === undefined || v === "") return null;
    if (typeof v === "number") return v;
    var m = String(v).match(/-?\d+(?:\.\d+)?/);   // ดึงตัวเลขนำ เช่น "9(3)" -> 9
    return m ? Number(m[0]) : null;
  }

  function parse(wb, H) {
    var norm = H.norm, ABBR = H.THAI_ABBR;
    var g = H.sheetGrid(wb, "Hemapheresis", window.XLSX);
    if (!g.length) return {};

    var hr = H.findRow(g, function (row) {
      return row.some(function (x) { return norm(x) === "ม.ค."; }) &&
             row.some(function (x) { return norm(x) === "ธ.ค."; });
    });
    if (hr < 0) return {};
    var hrow = g[hr];
    var lc = 1; for (var j = 0; j < hrow.length; j++) { if (norm(hrow[j]) === "รายการ") { lc = j; break; } }
    var mcols = ABBR.map(function (ab) {
      for (var k = 0; k < hrow.length; k++) if (norm(hrow[k]) === norm(ab)) return k;
      return null;
    });
    function vals(i) {
      var row = g[i] || [];
      return mcols.map(function (c) { return (c !== null && c < row.length) ? toNum(row[c]) : null; });
    }

    // เก็บแถวข้อมูล (i, normLabel)
    var rows = [];
    for (var r = hr + 1; r < g.length; r++) {
      var lab = (g[r] && g[r][lc] != null) ? g[r][lc] : "";
      if (String(lab).trim() !== "") rows.push([r, norm(lab)]);
    }
    function findIdx(pred) { for (var t = 0; t < rows.length; t++) if (pred(rows[t][1])) return rows[t][0]; return null; }

    var C = function (s) { return function (l) { return l.indexOf(norm(s)) >= 0; }; };
    var EQ = function (s) { return function (l) { return l === norm(s); }; };
    var P = function () { var a = [].slice.call(arguments); return function (l) { return a.some(function (s) { return l.indexOf(norm(s)) === 0; }); }; };

    var out = {};
    function setf(name, pred) { var i = findIdx(pred); if (i !== null) out[name] = vals(i); }

    setf("platelet", P("platelet")); setf("trima", EQ("trima")); setf("amicus", EQ("amicus"));
    setf("product", P("product")); setf("unitRandom", P("unitr"));

    // QC bins ตามตำแหน่ง: แถวระหว่าง unitRandom กับ QC% (ข้าม total/รวม)
    var ur = findIdx(P("unitr")), qp = findIdx(EQ("qc%"));
    if (ur !== null && qp !== null) {
      var bins = [];
      for (var b = 0; b < rows.length; b++) {
        var ri = rows[b][0], rl = rows[b][1];
        if (ri > ur && ri < qp && rl.indexOf("total") < 0 && rl.indexOf("รวม") < 0) bins.push(ri);
      }
      var names = ["qcBelow", "qcLow", "qcMid", "qcHigh", "qcVhigh"];
      for (var n = 0; n < bins.length && n < names.length; n++) out[names[n]] = vals(bins[n]);
    }
    setf("qcPct", EQ("qc%")); setf("fail", EQ("fail"));

    // Stem / Leuk / Plasma : รองรับทั้งแบบแยก (ครั้ง)/(ราย) และแบบรวม
    var sK = findIdx(C("stemcell(ครั้ง)")), sR = findIdx(C("stemcell(ราย)")), sOne = findIdx(EQ("stemcell"));
    if (sK !== null) out.stemOut = vals(sK);
    if (sR !== null) out.stemIn = vals(sR);
    if (sOne !== null) { out.stemOut = vals(sOne); out.stemIn = vals(sOne); }

    setf("cartCell", C("car-t"));

    var lK = findIdx(C("leukapheresis(ครั้ง)")), lR = findIdx(C("leukapheresis(ราย)")), lOne = findIdx(EQ("leukapheresis"));
    if (lK !== null) out.leukOut = vals(lK);
    if (lR !== null) out.leukIn = vals(lR);
    if (lOne !== null) out.leukOut = vals(lOne);

    var pK = findIdx(C("plasmaexchange(ครั้ง)")), pR = findIdx(C("plasmaexchange(ราย)"));
    var pOne = findIdx(function (l) { return l.indexOf("plasmaexchange") === 0 && l.indexOf("(") < 0; });
    if (pK !== null) out.plasmaOut = vals(pK);
    if (pR !== null) out.plasmaIn = vals(pR);
    if (pOne !== null) out.plasmaOut = vals(pOne);

    setf("granulocyte", C("granulocyte"));
    setf("washFrozen", C("washfrozen"));
    setf("pi", P("pi("));
    setf("dli", EQ("dli"));
    setf("sdpCase", function (l) { return l.indexOf("จำนวนจอง") >= 0 || l.indexOf("จำนวนราย") >= 0; });
    setf("sdpSend", C("ที่จอง"));
    setf("sdpReceive", C("ที่ได้รับ"));
    setf("outsideService", C("นอกเวลา"));
    setf("donorReaction", C("donorreaction"));
    return out;
  }

  function computeMonths(data) {
    var maxIdx = -1;
    Object.keys(data).forEach(function (k) {
      var arr = data[k]; if (!Array.isArray(arr)) return;
      for (var i = 0; i < arr.length; i++) if (arr[i] !== null && arr[i] !== undefined) { if (i > maxIdx) maxIdx = i; }
    });
    if (maxIdx < 0) return null;
    var out = []; for (var j = 0; j <= maxIdx; j++) out.push(j); return out;
  }

  function matchYear(filename) {
    // อัพเดทเฉพาะปีปัจจุบัน (สถิติ 69.xlsx) — ปีเก่าโครงสร้างต่างกัน คงข้อมูลเดิม
    return /^สถิติ\s*69\.xlsx$/i.test(String(filename)) ? CURRENT_YEAR : null;
  }

  function applyYear(year, data) {
    if (year !== CURRENT_YEAR) return;
    if (typeof D69 !== "undefined") {
      Object.keys(data).forEach(function (k) { if (D69.hasOwnProperty(k)) D69[k] = data[k]; });
    }
    if (!window.__autoMonthsByYear) window.__autoMonthsByYear = {};
    var mo = computeMonths(data);
    if (mo) window.__autoMonthsByYear[year] = mo;
  }

  function render() {
    if (!window.__gaiPatchedHema && typeof getAvailableIdx === "function") {
      var _gai = getAvailableIdx;
      // eslint-disable-next-line no-global-assign
      getAvailableIdx = function (y) {
        var by = window.__autoMonthsByYear;
        if (by && by[y]) return by[y];
        return _gai(y);
      };
      window.__gaiPatchedHema = true;
    }
    if (typeof refresh === "function") refresh();
    if (typeof buildCharts === "function") buildCharts();
  }

  DashboardAutoload.register({
    mode: "directory",
    storeKey: "hemapheresis",     // คีย์สำหรับบันทึกข้อมูลถาวร
    snapshotFile: "live_hemapheresis.js",  // เขียนข้อมูลล่าสุดลงไฟล์นี้
    defaultYear: CURRENT_YEAR,
    match: matchYear,
    parse: parse,
    applyYear: applyYear,
    render: render
  });
})();
