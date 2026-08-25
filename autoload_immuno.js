/* =====================================================================
 *  autoload_immuno.js — ตัวเชื่อมข้อมูลเฉพาะห้อง Immunohematology
 *  อ่านชีต "Immunohematology" (เดือนเป็นคอลัมน์) และ "Immunohematology2"
 *  (เดือนเป็นแถว, ผลิตภัณฑ์เป็นคอลัมน์คู่ ครั้ง/unit) จากไฟล์ สถิติ XX.xlsx
 *  แล้วแมปเข้าโครงสร้าง BLOOD_IMMUNO_DATA[ปี] ของ Dashboard
 *
 *  จับข้อมูลตาม "หัวข้อ/label" (ไม่ใช่เลขแถว/คอลัมน์ตายตัว) — ตรวจพบว่า
 *  หัวคอลัมน์ของ Immunohematology2 เปลี่ยนไปตามปี และบางปีไม่มี section
 *  "5.Request (Stat)"  ค่าที่ได้ตรวจสอบแล้วว่าตรงกับ blood_immuno_data.js เดิม
 *
 *  หมายเหตุ: รับเฉพาะปี 2566 ขึ้นไป (โครงสร้างชีตของ 2564/2565 ต่างออกไป
 *  และไฟล์ถูกแก้ภายหลัง) — ปีเก่าจะใช้ข้อมูล snapshot เดิม
 * ===================================================================== */
(function () {
  "use strict";

  var CURRENT_YEAR = 2569; // ไฟล์ สถิติ 69.xlsx = ปีปฏิทิน 2569

  function newArr() { return [null, null, null, null, null, null, null, null, null, null, null, null]; }

  /* ---------------------------------------------------------------------
   *  ฝั่ง "เตรียม" (production) — สำหรับหน้า Dashboard เตรียม/จ่าย เท่านั้น
   *  อ่านชีต "Blood Component" (+ SDP จากชีต "Hemapheresis" แถว Product (bag))
   *  แล้วแมปเป็นโครงสร้างเดียวกับ COMPONENT_PRODUCTION_DATA[ปี]
   *  (ตรรกะ label ตรงกับ autoload_bloodcomponent.js — เก็บแยกเพื่อให้ไฟล์นี้
   *   ทำงานได้ในตัวเองโดยไม่ต้องพึ่ง autoload ห้องอื่น)
   * ------------------------------------------------------------------- */
  function tonumP(v) {
    if (v === null || v === undefined || v === "") return null;
    if (typeof v === "number") return v;
    var s = String(v); if (s.toLowerCase().indexOf("div/0") >= 0) return null;
    var m = s.match(/-?\d+(?:\.\d+)?/); return m ? Number(m[0]) : null;
  }
  function parseProduction(wb, H) {
    var norm = H.norm;
    var g = H.sheetGrid(wb, "Blood Component", window.XLSX);
    if (!g.length) return {};
    var hr = H.findRow(g, function (row) { return row.some(function (x) { return norm(x) === "ม.ค."; }); });
    if (hr < 0) return {};
    var ms = -1, hrow = g[hr];
    for (var j = 0; j < hrow.length; j++) if (norm(hrow[j]) === "ม.ค.") { ms = j; break; }
    var mcols = []; for (var k = 0; k < 12; k++) mcols.push(ms + k);
    var lc = 0;
    function vals(i) {
      var row = g[i] || [];
      return mcols.map(function (c) { return (c < row.length) ? tonumP(row[c]) : null; });
    }
    var rows = [];
    for (var r = hr + 1; r < g.length; r++) { var lab = (g[r] && g[r][lc] != null) ? g[r][lc] : ""; if (String(lab).trim() !== "") rows.push([r, norm(lab)]); }
    function find(pred) {
      var t;
      for (t = 0; t < rows.length; t++) if (pred(rows[t][1])) { var v = vals(rows[t][0]); if (v.some(function (x) { return x !== null; })) return v; }
      for (t = 0; t < rows.length; t++) if (pred(rows[t][1])) return vals(rows[t][0]);
      return null;
    }
    var EQ = function (s) { return function (l) { return l === norm(s); }; };
    var C = function (s) { return function (l) { return l.indexOf(norm(s)) >= 0; }; };
    var P = function (s) { return function (l) { return l.indexOf(norm(s)) === 0; }; };
    var out = {}; function put(k, v) { if (v !== null && v !== undefined) out[k] = v; }
    put("donor", find(EQ("donor"))); put("wb", find(P("wholeblood"))); put("pfb", find(function (l) { return l === "pfb" || l === "ldprc"; }));
    put("ldppc", find(C("leukocytedepleted"))); put("stemcell", find(P("stemcell"))); put("washedCell", find(P("washedcell")));
    put("granulocyte", find(C("pooledgranulocyte"))); put("ffp", find(EQ("ffp"))); put("cryoSingle", find(P("cryoppt.(single)")));
    put("ffpStock_O", find(EQ("o"))); put("ffpStock_A", find(EQ("a"))); put("ffpStock_B", find(EQ("b"))); put("ffpStock_AB", find(EQ("ab")));
    put("lowVol", find(EQ("lowvolume"))); put("overVol", find(EQ("overvolume"))); put("overtime", find(EQ("overtime"))); put("shortBag", find(EQ("shortbag")));
    put("clotted", find(C("clotted(ใช้"))); put("clottedUnused", find(C("clotted(ไม่ใช้"))); put("contaminated", find(EQ("contaminated")));
    put("donorNew", find(EQ("donorใน"))); put("donorOld", find(EQ("donorหน่วย")));
    put("quadBag", find(EQ("quadruplebag"))); put("tripleBag", find(EQ("triplebag"))); put("singleBag", find(EQ("singlebag"))); put("doubleBag", find(EQ("doublebag")));

    // SDP (Single Donor Platelets) — จากชีต "Hemapheresis" แถว "Product (bag)"
    try {
      var gh = H.sheetGrid(wb, "Hemapheresis", window.XLSX);
      if (gh.length) {
        var hhr = H.findRow(gh, function (row) {
          return row.some(function (x) { return norm(x) === "ม.ค."; }) &&
                 row.some(function (x) { return norm(x) === "ธ.ค."; });
        });
        if (hhr >= 0) {
          var hrow2 = gh[hhr], ABBR = H.THAI_ABBR;
          var hmcols = ABBR.map(function (ab) {
            for (var kk = 0; kk < hrow2.length; kk++) if (norm(hrow2[kk]) === norm(ab)) return kk;
            return null;
          });
          var hlc = 1;
          for (var jj = 0; jj < hrow2.length; jj++) { if (norm(hrow2[jj]) === "รายการ") { hlc = jj; break; } }
          var sdpRow = -1;
          for (var rr = hhr + 1; rr < gh.length; rr++) {
            var l2 = (gh[rr] && gh[rr][hlc] != null) ? norm(gh[rr][hlc]) : "";
            if (l2.indexOf("product") === 0) { sdpRow = rr; break; }
          }
          if (sdpRow >= 0) {
            var srow = gh[sdpRow] || [];
            var sdpVals = hmcols.map(function (c) { return (c !== null && c < srow.length) ? tonumP(srow[c]) : null; });
            if (sdpVals.some(function (x) { return x !== null; })) put("sdp", sdpVals);
          }
        }
      }
    } catch (e) { /* ไม่มีชีต Hemapheresis ก็ข้าม คงค่า sdp เดิม */ }

    return out;
  }

  function parse(wb, H) {
    var norm = H.norm, num = H.num, ABBR = H.THAI_ABBR;

    var out = {
      antibodyId: {}, occurrence: {},
      bloodProductsPaid: { dispense: {}, receive: {} },
      bloodProductsUsage: { granulocyte: {}, pc: {}, totalPlatelet: {}, sdp: {}, cryo: {},
                            lppc: {}, bovineThrombin: {}, crp: {}, dli: {}, ffp: {} },
      requestStat: {}, requestEmergency: {}, request: {}, plateletExpired: null
    };

    /* ===== ส่วนที่ 1: ชีต "Immunohematology" (เดือนเป็นคอลัมน์ B..M) ===== */
    var g1 = H.sheetGrid(wb, "Immunohematology", window.XLSX);

    function labelAt(g, i) { return (g[i] && g[i][0] != null) ? String(g[i][0]) : ""; }

    var monthCols1 = null;
    for (var r = 0; r < g1.length && !monthCols1; r++) {
      var row = g1[r] || [];
      var hasJan = row.some(function (x) { return norm(x) === "ม.ค."; });
      var hasDec = row.some(function (x) { return norm(x) === "ธ.ค."; });
      if (hasJan && hasDec) {
        monthCols1 = ABBR.map(function (ab) {
          for (var j = 0; j < row.length; j++) if (norm(row[j]) === norm(ab)) return j;
          return -1;
        });
      }
    }
    if (!monthCols1) monthCols1 = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12];

    function vals1(i) {
      var row = g1[i] || [];
      return monthCols1.map(function (c) { return (c >= 0 && c < row.length) ? num(row[c]) : null; });
    }
    function findSec(sub, start) {
      for (var i = (start || 0); i < g1.length; i++) {
        if (norm(labelAt(g1, i)).indexOf(norm(sub)) >= 0) return i;
      }
      return -1;
    }
    function rowIn(s, e, pred) {
      for (var i = s; i < e && i < g1.length; i++) { if (pred(norm(labelAt(g1, i)))) return i; }
      return -1;
    }
    function vals1In(s, e, pred) { var i = rowIn(s, e, pred); return i >= 0 ? vals1(i) : newArr(); }
    function EQ(x) { return function (l) { return l === norm(x); }; }
    function C(x) { return function (l) { return l.indexOf(norm(x)) >= 0; }; }

    var END1 = g1.length;
    var sReq   = findSec("1.request");
    var sPay   = findSec("จ่ายbloodproduct", sReq + 1);
    var sOcc   = findSec("3.occurrence", sPay + 1);
    var sAb    = findSec("antibody", sOcc + 1);
    var sStat  = findSec("request(stat)", sAb + 1);
    var sEmerg = findSec("request(emergency)", (sStat >= 0 ? sStat : sAb) + 1);

    function nextStart(after) {
      var cand = [sPay, sOcc, sAb, sStat, sEmerg].filter(function (x) { return x > after; });
      return cand.length ? Math.min.apply(null, cand) : END1;
    }
    var eReq   = nextStart(sReq);
    var ePay   = nextStart(sPay);
    var eOcc   = nextStart(sOcc);
    var eAb    = nextStart(sAb);
    var eStat  = sStat >= 0 ? nextStart(sStat) : -1;
    var eEmerg = END1;

    // ----- 1.Request : ในเวลา=receive, นอกเวลา=dispense -----
    var inReq  = rowIn(sReq, eReq, C("ในเวลา"));
    var outReq = rowIn(sReq, eReq, C("นอกเวลา"));
    out.request.receiveSpecimens   = inReq  >= 0 ? vals1(inReq)      : newArr();
    out.request.receiveCrossmatch  = inReq  >= 0 ? vals1(inReq + 1)  : newArr();
    out.request.dispenseSpecimens  = outReq >= 0 ? vals1(outReq)     : newArr();
    out.request.dispenseCrossmatch = outReq >= 0 ? vals1(outReq + 1) : newArr();
    out.request.totalSpecimens     = vals1In(sReq, eReq, C("รวมในเวลา"));
    out.request.totalCrossmatch    = vals1In(sReq, eReq, C("cmรวม"));
    out.request.transfusionRbc     = vals1In(sReq, eReq, C("transfusionrbc"));
    out.request.ctRatio = vals1In(sReq, eReq, EQ("c:t")).map(function (v) {
      return (v === null) ? null : Math.round(v * 1000) / 1000;
    });
    out.bloodProductsUsage.rbc = out.request.transfusionRbc.slice();

    // ----- 2.จ่าย Blood Product : ในเวลา=receive, นอกเวลา=dispense -----
    var inPay  = rowIn(sPay, ePay, C("ในเวลา"));
    var outPay = rowIn(sPay, ePay, C("นอกเวลา"));
    function payBlock(start, end, target) {
      target.pltConcPat   = vals1In(start, end, C("plt.concentrate(ครั้ง"));
      target.pltConcDon   = vals1In(start, end, C("plt.concentrate(ยู"));
      target.totalPltConc = vals1In(start, end, C("รวมplt"));
      target.ffp          = vals1In(start, end, C("ffp"));  // แถวชื่อ "FFP (unit)" — ใช้แบบมีคำว่า ffp
      target.cryoPool     = vals1In(start, end, C("cryo(pool"));
      target.cryoUnit     = vals1In(start, end, C("cryo(unit"));
    }
    payBlock(inPay  >= 0 ? inPay  : sPay, outPay >= 0 ? outPay : ePay, out.bloodProductsPaid.receive);
    payBlock(outPay >= 0 ? outPay : sPay, ePay, out.bloodProductsPaid.dispense);

    // ----- 3.Occurrence -----
    out.occurrence.transfusionReaction   = vals1In(sOcc, eOcc, C("ผิดหมู่"));
    out.occurrence.wrongBloodInTube      = vals1In(sOcc, eOcc, C("wrongbloodintube"));
    out.occurrence.inappropriateSpecimen = vals1In(sOcc, eOcc, C("ปฏิเสธสิ่งส่งตรวจ"));
    out.occurrence.total                 = vals1In(sOcc, eOcc, C("total"));

    // ----- 4.Antibody Identification -----
    out.antibodyId.unidentified = vals1In(sAb, eAb, C("unidentified"));
    out.antibodyId.nonSpecific  = vals1In(sAb, eAb, C("nonspecific"));
    out.antibodyId.autoAb       = vals1In(sAb, eAb, C("autoab"));
    out.antibodyId.unexpectedAb = vals1In(sAb, eAb, C("unexpected"));
    out.antibodyId.total        = vals1In(sAb, eAb, C("total"));
    out.antibodyId.patient      = newArr();

    // ----- 5.Request (Stat) / (Emergency) -----
    out.requestStat.total      = (sStat >= 0) ? vals1In(sStat, eStat, C("จำนวนใบrequest")) : newArr();
    out.requestEmergency.total = vals1In(sEmerg, eEmerg, C("จำนวนใบrequest"));

    /* ===== ส่วนที่ 2: ชีต "Immunohematology2" (เดือนเป็นแถว, ผลิตภัณฑ์คอลัมน์คู่) ===== */
    var g2 = H.sheetGrid(wb, "Immunohematology2", window.XLSX);

    var hdrRow = -1;
    for (var i2 = 0; i2 < g2.length; i2++) {
      if (norm(g2[i2] && g2[i2][0]) === "เดือน") { hdrRow = i2; break; }
    }
    var monthRow = {};
    var ABBRN = ABBR.map(norm);
    for (var k = 0; k < g2.length; k++) {
      var a = norm(g2[k] && g2[k][0]);
      var mi = ABBRN.indexOf(a);
      if (mi >= 0) monthRow[mi] = k;
    }
    function colByHeader(names) {
      if (hdrRow < 0) return -1;
      var hrow = g2[hdrRow] || [];
      for (var j = 0; j < hrow.length; j++) {
        var h = norm(hrow[j]);
        for (var n = 0; n < names.length; n++) if (h === norm(names[n])) return j;
      }
      return -1;
    }
    function colVals(col) {
      var arr = newArr();
      if (col < 0) return arr;
      for (var m = 0; m < 12; m++) {
        if (monthRow[m] != null) {
          var row = g2[monthRow[m]] || [];
          arr[m] = (col < row.length) ? num(row[col]) : null;
        }
      }
      return arr;
    }
    function pairAt(names) {
      var c = colByHeader(names);
      return { patient: colVals(c), unit: c >= 0 ? colVals(c + 1) : newArr() };
    }
    function single(names) { return colVals(colByHeader(names)); }

    var U = out.bloodProductsUsage;
    U.wb  = single(["WB"]);
    U.prc = single(["PRC"]);
    U.lpb = single(["LPB"]);
    U.pfb = single(["PFB"]);
    var rbc2 = single(["RBC"]);
    if (rbc2.some(function (x) { return x !== null; })) U.rbc = rbc2;

    U.pc             = pairAt(["PC", "Plt.conc."]);
    U.lppc           = pairAt(["LPPC/LDPPC", "Plt.opti."]);
    U.sdp            = pairAt(["SDP", "Plt.Pheresis"]);
    U.cryo           = pairAt(["Cryo"]);
    U.ffp            = pairAt(["FFP"]);
    U.crp            = pairAt(["CRP", "FFP cryo.removed"]);
    U.granulocyte    = pairAt(["Granulocyte"]);
    U.dli            = pairAt(["DLI"]);
    U.bovineThrombin = pairAt(["Bovine Thrombin"]);

    var totalCol = colByHeader(["รวม"]);
    U.totalPlatelet = { patient: colVals(totalCol), unit: totalCol >= 0 ? colVals(totalCol + 1) : newArr() };

    U.transfusionReactionCases = single(["T.R."]);

    // ข้อมูลฝั่ง "เตรียม" (production) แนบไว้ใน __production
    // ใช้เฉพาะหน้า Dashboard เตรียม/จ่าย — หน้า Immunohematology ไม่แตะฟิลด์นี้
    try { out.__production = parseProduction(wb, H); } catch (e) { /* ข้ามถ้าอ่านไม่ได้ */ }

    // เกล็ดเลือดหมดอายุ (Platelet Expired) — อ่านชีต "Platelet Exp."
    try { var pe = parsePlateletExpired(wb, H); if (pe) out.plateletExpired = pe; }
    catch (e) { /* ไม่มีชีต Platelet Exp. ก็ข้าม คงค่าเดิม */ }

    return out;
  }

  /* ---------------------------------------------------------------------
   *  เกล็ดเลือดหมดอายุ (Platelet Expired) — ชีต "Platelet Exp."
   *  โครงสร้าง: บล็อกละ 1 เดือน — แถวหัว = ชื่อเดือน + O A B AB,
   *  ตามด้วยแถว LDPPC / LDPPC(A) / SDP / SDP(A) (ค่าคือจำนวนตามหมู่เลือด)
   *  รวม O+A+B+AB ต่อแถว เก็บเป็น array index แบบปีปฏิทิน (0=ม.ค. .. 11=ธ.ค.)
   *  ให้ตรงกับข้อมูลส่วนอื่นของ Dashboard (MONTHS_FY เป็นลำดับปฏิทิน)
   * ------------------------------------------------------------------- */
  function parsePlateletExpired(wb, H) {
    var g = H.sheetGrid(wb, "Platelet Exp.", window.XLSX);
    if (!g || !g.length) return null;

    // ชื่อเดือน → index ปีปฏิทิน (ม.ค.=0)
    var MONTH_CAL = {
      jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5,
      jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11
    };
    function monthIdx(label) {
      var s = String(label).toLowerCase().replace(/[^a-z]/g, "");
      if (!s) return -1;
      // สะกดผิดในไฟล์ต้นฉบับ: "janaury" → jan
      if (s.indexOf("jana") === 0) return 0;
      var keys = Object.keys(MONTH_CAL);
      for (var i = 0; i < keys.length; i++) if (s.indexOf(keys[i]) === 0) return MONTH_CAL[keys[i]];
      return -1;
    }
    // ชื่อชนิด → key (ตรวจ (A) ก่อน เพราะ ldppc/sdp เป็น substring)
    function typeKey(label) {
      var s = String(label).toLowerCase().replace(/\s/g, "");
      if (s.indexOf("ldppc(a)") === 0) return "ldppcA";
      if (s.indexOf("ldppc") === 0) return "ldppc";
      if (s.indexOf("sdp(a)") === 0) return "sdpA";
      if (s.indexOf("sdp") === 0) return "sdp";
      return null;
    }

    var out = {
      ldppc: newArr12(), ldppcA: newArr12(), sdp: newArr12(),
      sdpA: newArr12(), total: newArr12()
    };
    function newArr12() { return [0,0,0,0,0,0,0,0,0,0,0,0]; }

    var curIdx = -1, found = false;
    for (var r = 0; r < g.length; r++) {
      var row = g[r] || [];
      var c0 = (row[0] != null) ? row[0] : "";
      var mi = monthIdx(c0);
      if (mi >= 0) { curIdx = mi; continue; }
      if (curIdx < 0) continue;
      var tk = typeKey(c0);
      if (!tk) continue;
      var sum = 0;
      for (var c = 1; c <= 4; c++) { var v = tonumP(row[c]); if (v) sum += v; }
      out[tk][curIdx] = sum;
      if (sum) found = true;
    }
    // อัปเดต total ต่อเดือน
    for (var k = 0; k < 12; k++) {
      out.total[k] = (out.ldppc[k] || 0) + (out.ldppcA[k] || 0) + (out.sdp[k] || 0) + (out.sdpA[k] || 0);
    }
    void found; // พบชีตแล้วคืนค่าเสมอ — ให้ข้อมูลต้นทางเป็นตัวตัดสิน
    return out;
  }

  function matchYear(filename) {
    var m = String(filename).match(/^สถิติ\s*(\d{2})\.xlsx$/i);
    if (!m) return null;
    var year = 2500 + parseInt(m[1], 10);
    // รับเฉพาะปี 2566 ขึ้นไป (โครงสร้าง 2564/2565 ต่างออกไป และถูกแก้ภายหลัง)
    return (year >= 2566 && year <= 2569) ? year : null;
  }

  function applyYear(year, data) {
    // --- ฝั่ง "เตรียม" (production) : เขียนลง COMPONENT_PRODUCTION_DATA ถ้าหน้านั้นมีตัวแปรนี้ ---
    // (คง __production ไว้ใน data เพื่อให้ snapshot live_immuno.js พกพา/รีเพลย์ตอนเปิดหน้าได้)
    if (data && data.__production &&
        typeof COMPONENT_PRODUCTION_DATA !== "undefined" && COMPONENT_PRODUCTION_DATA) {
      var ptgt = COMPONENT_PRODUCTION_DATA[String(year)];
      if (ptgt) {
        Object.keys(data.__production).forEach(function (k) {
          if (ptgt.hasOwnProperty(k)) ptgt[k] = data.__production[k];
        });
      }
    }

    if (typeof BLOOD_IMMUNO_DATA === "undefined") return;
    var key = String(year);
    var prev = BLOOD_IMMUNO_DATA[key] || {};
    if (data.plateletExpired == null && prev.plateletExpired != null) {
      data.plateletExpired = prev.plateletExpired;
    }
    if (prev.antibodyId && prev.antibodyId.patient &&
        data.antibodyId && data.antibodyId.patient &&
        !data.antibodyId.patient.some(function (x) { return x !== null; })) {
      data.antibodyId.patient = prev.antibodyId.patient;
    }
    BLOOD_IMMUNO_DATA[key] = data;
  }

  function render() {
    if (typeof refresh === "function") refresh();
    if (typeof buildCharts === "function") buildCharts();
  }

  DashboardAutoload.register({
    mode: "directory",
    storeKey: "immuno",
    snapshotFile: "live_immuno.js",
    defaultYear: CURRENT_YEAR,
    match: matchYear,
    parse: parse,
    applyYear: applyYear,
    render: render
  });
})();
