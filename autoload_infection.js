/* =====================================================================
 *  autoload_infection.js — ตัวเชื่อมข้อมูลเฉพาะห้อง Infectious Serology
 *  อ่านชีต "Infec." จากไฟล์ สถิติ XX.xlsx แล้วแมปเข้าโครงสร้าง
 *  BLOOD_INFECTION_DATA[ปี] ของ Dashboard
 *
 *  จับข้อมูลตาม "หัวข้อ/label" ในคอลัมน์ A (ไม่ใช่เลขแถวตายตัว) จึงทนต่อ
 *  การเพิ่ม/ย้ายแถว — ตรวจสอบแล้วว่าโครงสร้างแถวของแต่ละปีไม่เหมือนกัน
 *  (เช่น ปี 2564/2565 ส่วน "จำหน่ายทิ้ง" เริ่มคนละแถวกับปี 2566)
 *
 *  ค่าที่ได้ตรวจสอบแล้วว่าตรงกับ blood_infection_data.js เดิมครบทุกช่อง
 * ===================================================================== */
(function () {
  "use strict";

  var CURRENT_YEAR = 2569; // ไฟล์ สถิติ 69.xlsx = ปีปฏิทิน 2569

  function newArr() { return [null, null, null, null, null, null, null, null, null, null, null, null]; }

  function parse(wb, H) {
    var norm = H.norm, num = H.num, findRow = H.findRow, ABBR = H.THAI_ABBR;
    var g = H.sheetGrid(wb, "Infec.", window.XLSX);

    // ---- หาแถวหัวเดือน (มี "ม.ค." และ "ธ.ค.") เพื่อระบุคอลัมน์ของแต่ละเดือน ----
    var mh = findRow(g, function (row) {
      return row.some(function (x) { return norm(x) === "ม.ค."; }) &&
             row.some(function (x) { return norm(x) === "ธ.ค."; });
    });
    var monthCols = [];
    ABBR.forEach(function (ab) {
      var hrow = g[mh] || [];
      var found = -1;
      for (var j = 0; j < hrow.length; j++) { if (norm(hrow[j]) === norm(ab)) { found = j; break; } }
      monthCols.push(found);
    });

    function valsAt(i) {
      var row = g[i] || [];
      return monthCols.map(function (c) { return (c >= 0 && c < row.length) ? num(row[c]) : null; });
    }
    function labelAt(i) { return (g[i] && g[i][0] != null) ? String(g[i][0]) : ""; }

    // หาแถวที่ label (คอลัมน์ A) ตรงเงื่อนไข ภายในช่วง [s,e)
    function findInRange(s, e, pred) {
      for (var i = s; i < e && i < g.length; i++) { if (pred(norm(labelAt(i)))) return i; }
      return -1;
    }
    // หาแถวหัวข้อ (section header) จาก substring — คืน index แถวนั้น
    function findHeader(sub, start) {
      for (var i = (start || 0); i < g.length; i++) {
        if (norm(labelAt(i)).indexOf(norm(sub)) >= 0) return i;
      }
      return -1;
    }
    function C(s) { return function (l) { return l.indexOf(norm(s)) >= 0; }; }       // contains
    function EQ(s) { return function (l) { return l === norm(s); }; }                 // exact
    function START(s) { return function (l) { return l.indexOf(norm(s)) === 0; }; }   // startsWith

    function rowVals(s, e, pred) {
      var i = findInRange(s, e, pred);
      return i >= 0 ? valsAt(i) : newArr();
    }

    // ---- ระบุขอบเขตของแต่ละ section ตามหัวข้อ (ทนต่อการเลื่อนแถว) ----
    var hSero = findHeader("วิธี Serology");            // "1. จำนวนตัวอย่าง...Serology"
    var hNat  = findHeader("วิธี NAT", hSero + 1);      // "2. จำนวนตัวอย่าง...NAT"
    var hDisc = findHeader("จำหน่ายทิ้ง", hNat + 1);    // "3. ปริมาณเลือดที่จำหน่ายทิ้ง"
    var hTP   = findHeader("True Positive", (hDisc >= 0 ? hDisc : hNat) + 1);

    var END = g.length;
    var eSero = hNat  >= 0 ? hNat  : END;
    var eNat  = hDisc >= 0 ? hDisc : END;
    var eDisc = hTP   >= 0 ? hTP   : END;
    var eTP   = END;

    var out = { misc: {}, truePositive: {}, serology: {}, nat: {}, discarded: {}, plateletSterility: null };

    // ---------------- 1) Serology ----------------
    out.serology.inHouse      = rowVals(hSero, eSero, C("ภายในสถานที่"));
    out.serology.pltApheresis = rowVals(hSero, eSero, START("pltapheresis")); // "Plt Apheresis+PBSC*+plasma"
    out.serology.mobile       = rowVals(hSero, eSero, C("หน่วยเคลื่อนที่"));
    out.serology.bmt          = rowVals(hSero, eSero, START("bmt"));      // "BMT/นมแม่/prescreen/เนื้อเยื่อ"
    out.serology.total        = rowVals(hSero, eSero, EQ("รวม"));

    // ---------------- 2) NAT ----------------
    out.nat.inHouse    = rowVals(hNat, eNat, C("ภายในสถานที่"));
    out.nat.mobile     = rowVals(hNat, eNat, C("หน่วยเคลื่อนที่"));
    out.nat.totalExBmt = rowVals(hNat, eNat, START("รวม(ไม่รวม"));        // "รวม (ไม่รวม BMT,plt,ID)"
    out.nat.bmt        = rowVals(hNat, eNat, START("bmt"));               // "BMT (Stem cell)"
    out.nat.prescreen  = rowVals(hNat, eNat, START("prescreen"));         // "Prescreen (plt apheresis)"
    out.nat.breastMilk = rowVals(hNat, eNat, EQ("นมแม่"));
    out.nat.sampleOnly = rowVals(hNat, eNat, START("sampleonly"));        // "Sample only"
    out.nat.seroWeak   = rowVals(hNat, eNat, START("seroweak"));          // "sero weak/lookback"
    out.nat.total      = rowVals(hNat, eNat, EQ("รวม"));

    // ---------------- 3) Discarded (จำหน่ายทิ้ง) ----------------
    out.discarded.HBsAg    = rowVals(hDisc, eDisc, EQ("hbsag"));
    out.discarded.Syphilis = rowVals(hDisc, eDisc, EQ("syphilis"));
    out.discarded.AntiHIV  = rowVals(hDisc, eDisc, START("anti-hiv"));
    out.discarded.AntiHCV  = rowVals(hDisc, eDisc, START("anti-hcv"));
    out.discarded.HBV_NAT  = rowVals(hDisc, eDisc, START("hbv"));         // "HBV (NAT)"
    out.discarded.HCV_NAT  = rowVals(hDisc, eDisc, START("hcv"));         // "HCV (NAT)"
    out.discarded.HIV_NAT  = rowVals(hDisc, eDisc, START("hiv(nat)"));    // "HIV (NAT)"

    // ---------------- misc (อยู่ใต้ section จำหน่ายทิ้ง ก่อน True Positive) ----------------
    // หัวข้อมีเลขนำหน้า เช่น "4 delay (วัน)", "5 failure run (batch)", "6 จำนวน run (NAT)...", "7 clot error"
    out.misc.delayDays  = rowVals(hDisc, eDisc, C("delay"));
    out.misc.failureRun = rowVals(hDisc, eDisc, C("failurerun"));
    out.misc.natRuns    = rowVals(hDisc, eDisc, C("จำนวนrun"));
    out.misc.clotError  = rowVals(hDisc, eDisc, C("cloterror"));

    // ---------------- True Positive ----------------
    out.truePositive.HBsAg    = rowVals(hTP, eTP, EQ("hbsag"));
    out.truePositive.Syphilis = rowVals(hTP, eTP, EQ("syphilis"));
    out.truePositive.AntiHIV  = rowVals(hTP, eTP, START("anti-hiv"));
    out.truePositive.AntiHCV  = rowVals(hTP, eTP, START("anti-hcv"));
    out.truePositive.HBV_NAT  = rowVals(hTP, eTP, START("hbv"));
    out.truePositive.HCV_NAT  = rowVals(hTP, eTP, START("hcv"));
    // หมายเหตุ: ตาราง True Positive ไม่มีแถว "HIV (NAT)" ในบางปี — ใส่เท่าที่มี
    var hivTP = rowVals(hTP, eTP, START("hiv(nat)"));
    if (hivTP.some(function (x) { return x !== null; })) out.truePositive.HIV_NAT = hivTP;

    // ---------------- Platelet Sterility (อ่านจากชีตแยก "Plt.Sterility test") ----------------
    // ชีตนี้วางเดือนไว้ในคอลัมน์ A (ไม่ใช่หัวเดือนแนวนอน) มี 2 บล็อค: "LDPC" และ "Single Donor Platelet"
    // หัวคอลัมน์: จำนวนที่ตรวจ / จำนวน Positive / True positive / false positive / Identification result
    out.plateletSterility = parseSterility();

    function parseSterility() {
      var gs;
      try { gs = H.sheetGrid(wb, "Plt.Sterility test", window.XLSX); } catch (e) { gs = null; }
      if (!gs || !gs.length) return null;

      function labelA(i) { return (gs[i] && gs[i][0] != null) ? String(gs[i][0]) : ""; }
      function findLabel(sub, start) {
        for (var i = (start || 0); i < gs.length; i++) {
          if (norm(labelA(i)).indexOf(norm(sub)) >= 0) return i;
        }
        return -1;
      }
      // หาแถวหัวคอลัมน์ (มี "จำนวนที่ตรวจ") ภายใน 4 แถวถัดจาก label ของบล็อค
      function headerAfter(r) {
        for (var i = r; i < gs.length && i < r + 5; i++) {
          var row = gs[i] || [];
          for (var j = 0; j < row.length; j++) {
            if (norm(row[j]).indexOf(norm("จำนวนที่ตรวจ")) >= 0) return i;
          }
        }
        return -1;
      }
      function readBlock(hr, stopAt) {
        if (hr < 0) return null;
        var hdr = gs[hr] || [];
        function colOf(pred) { for (var j = 0; j < hdr.length; j++) { if (pred(norm(hdr[j]))) return j; } return -1; }
        var cTested = colOf(function (l) { return l.indexOf(norm("จำนวนที่ตรวจ")) >= 0; });
        var cPos    = colOf(function (l) { return l.indexOf("positive") >= 0 && l.indexOf("true") < 0 && l.indexOf("false") < 0; });
        var cTrue   = colOf(function (l) { return l.indexOf("true") >= 0 && l.indexOf("positive") >= 0; });
        var cFalse  = colOf(function (l) { return l.indexOf("false") >= 0 && l.indexOf("positive") >= 0; });
        var cOrg    = colOf(function (l) { return l.indexOf("identification") >= 0; });
        var tested = newArr(), positive = newArr(), truePositive = newArr(), falsePositive = newArr(), organisms = newArr();
        var end = (stopAt > hr ? stopAt : gs.length);
        for (var i = hr + 1; i < end; i++) {
          var na = norm(labelA(i));
          if (na.indexOf(norm("รวม")) >= 0) break;           // ถึงแถว "รวม" -> จบบล็อค
          var mi = -1;
          for (var k = 0; k < ABBR.length; k++) { if (norm(ABBR[k]) === na) { mi = k; break; } }
          if (mi < 0) continue;                               // ไม่ใช่แถวเดือน -> ข้าม
          var row = gs[i] || [];
          if (cTested >= 0) tested[mi]        = num(row[cTested]);
          if (cPos    >= 0) positive[mi]      = num(row[cPos]);
          if (cTrue   >= 0) truePositive[mi]  = num(row[cTrue]);
          if (cFalse  >= 0) falsePositive[mi] = num(row[cFalse]);
          if (cOrg    >= 0) {
            var v = row[cOrg]; var s = (v == null) ? "" : String(v).trim();
            organisms[mi] = (s === "" || s === "-") ? null : s;
          }
        }
        return { tested: tested, positive: positive, truePositive: truePositive, falsePositive: falsePositive, organisms: organisms };
      }

      var rLdpc = findLabel("LDPC");
      var rSdp  = findLabel("Single Donor");
      if (rLdpc < 0 && rSdp < 0) return null;
      var res = {};
      var ldpc = readBlock(headerAfter(rLdpc), (rSdp > rLdpc ? rSdp : -1));
      var sdp  = readBlock(headerAfter(rSdp), -1);
      if (ldpc) res.ldpc = ldpc;
      if (sdp)  res.sdp  = sdp;
      // ถ้าทั้งคู่ไม่มีข้อมูลจริงเลย ให้คืน null เพื่อ fallback ไปใช้ค่าเดิม
      if (!res.ldpc && !res.sdp) return null;
      return res;
    }

    return out;
  }

  // ชื่อไฟล์ "สถิติ NN.xlsx" (ปีปฏิทิน) -> พ.ศ. 25NN  (ข้าม "ปีงบประมาณ"/"(Autosaved)")
  function matchYear(filename) {
    var m = String(filename).match(/^สถิติ\s*(\d{2})\.xlsx$/i);
    if (!m) return null;
    var year = 2500 + parseInt(m[1], 10);
    return (year >= 2564 && year <= 2569) ? year : null;
  }

  // ใส่ข้อมูลของปีนั้นเข้า BLOOD_INFECTION_DATA[ปี]  (getAvailableIdx จะคำนวณเดือนเอง)
  function applyYear(year, data) {
    if (typeof BLOOD_INFECTION_DATA === "undefined") return;
    var key = String(year);
    var prev = BLOOD_INFECTION_DATA[key] || {};
    // รักษา plateletSterility เดิมไว้ ถ้าไฟล์ Excel ไม่มีข้อมูลส่วนนี้
    if (data.plateletSterility == null && prev.plateletSterility != null) {
      data.plateletSterility = prev.plateletSterility;
    }
    BLOOD_INFECTION_DATA[key] = data;
  }

  function render() {
    if (typeof refresh === "function") refresh();
    if (typeof buildCharts === "function") buildCharts();
  }

  DashboardAutoload.register({
    mode: "directory",
    storeKey: "infection",
    snapshotFile: "live_infection.js",
    defaultYear: CURRENT_YEAR,
    match: matchYear,
    parse: parse,
    applyYear: applyYear,
    render: render
  });
})();
