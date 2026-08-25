/* =====================================================================
 *  autoload_hla.js — ตัวเชื่อมข้อมูลเฉพาะห้อง HLA Laboratory
 *  อ่านชีต "HLA1" (หรือ "HLA") -> โครงสร้าง HLA2569 ของ Dashboard
 *  จับตามเลขข้อ (1.1, 2.1…) + ตำแหน่งของ EQA ; รองรับเซลล์สูตรข้อความ "2*4=8"
 *  อัพเดทเฉพาะปีปัจจุบัน (2569) — ปีเก่าโครงสร้างต่างกัน คงข้อมูลเดิมไว้
 *  ตรวจสอบแล้ว: สถิติ 69 ตรงกับ HLA2569 ครบ 46 ฟิลด์
 * ===================================================================== */
(function () {
  "use strict";
  var CURRENT_YEAR = 2569;

  function tonum(v) {
    if (v === null || v === undefined || v === "") return null;
    if (typeof v === "number") return v;
    var s = String(v); if (s.indexOf("=") >= 0) s = s.split("=").pop();
    var m = s.match(/-?\d+(?:\.\d+)?/); return m ? Number(m[0]) : null;
  }

  function parse(wb, H) {
    var norm = H.norm, XLSX = window.XLSX;
    var sheet = wb.Sheets["HLA1"] ? "HLA1" : (wb.Sheets["HLA"] ? "HLA" : null);
    if (!sheet) return {};
    var g = H.sheetGrid(wb, sheet, XLSX);
    if (!g.length) return {};
    var hr = H.findRow(g, function (row) { return row.some(function (x) { return norm(x) === "ม.ค."; }); });
    if (hr < 0) return {};
    var ms = -1, hrow = g[hr];
    for (var j = 0; j < hrow.length; j++) if (norm(hrow[j]) === "ม.ค.") { ms = j; break; }
    var mcols = []; for (var k = 0; k < 12; k++) mcols.push(ms + k);
    function vals(i) { var row = g[i] || []; return mcols.map(function (c) { return c < row.length ? tonum(row[c]) : null; }); }
    var labs = g.map(function (row) {
      var a = (row && row[0] != null) ? norm(row[0]) : "";
      var c = (row && row[2] != null) ? norm(row[2]) : "";
      return a + "|" + c;
    });
    function find(pred, start) { for (var i = (start || 0); i < g.length; i++) if (pred(labs[i])) return i; return -1; }
    var C = function (s) { return function (l) { return l.indexOf(norm(s)) >= 0; }; };
    var out = {};
    function setf(k, pred) { var i = find(pred); if (i >= 0) out[k] = vals(i); }

    setf("hla_lympho_patient", C("1.1")); setf("hla_lympho_donor", C("1.2")); setf("hla_lympho_panel", C("1.3")); setf("hla_lympho_granu", C("1.4"));
    setf("hla_class1_kt", C("2.1")); setf("hla_class1_bmt", C("2.2")); setf("hla_class1_plt_pher", C("2.3")); setf("hla_class1_pt_tt", C("2.4"));
    setf("hla_b27_siriraj", C("ศิริราช")); setf("hla_b27_center", C("ศูนย์การแพทย์"));
    setf("hla_cdc_kt", C("4.1")); setf("hla_cdc_re", C("4.2")); setf("hla_cdc_granu", C("4.3"));
    setf("hla_fcxm_kt", C("5.1")); setf("hla_fcxm_re", C("5.2"));
    setf("hla_plt_ab_patient", C("patient(ราย")); setf("hla_plt_ab_neg", function (l) { return l.indexOf("negative(ร") >= 0 && l.indexOf("unit") < 0; }); setf("hla_plt_ab_pos", function (l) { return l.indexOf("positive(ราย") >= 0; });
    setf("hla_plt_xm_case", C("7.platelet")); setf("hla_plt_xm_unit_neg", C("negative(unit")); setf("hla_plt_xm_unit_pos", C("positive(unit"));
    setf("hla_serum_auto", C("8.1")); setf("hla_serum_allo", C("8.2")); setf("hla_serum_prp", C("8.3"));
    setf("hla_serum_src_opd", C("opd")); setf("hla_serum_src_ward", function (l) { return l.indexOf("-ward") >= 0; }); setf("hla_serum_src_siph", function (l) { return l.indexOf("-siph") >= 0; });
    setf("hla_rc_geno_rhd", C("9.1")); setf("hla_rc_geno_rhdel", C("9.2")); setf("hla_rc_geno_kidd", C("9.3")); setf("hla_rc_geno_duffy", C("9.4"));
    setf("hla_rc_geno_fc_rhdel", C("9.5")); setf("hla_rc_geno_rhce", C("9.6")); setf("hla_rc_geno_dombrock", C("9.7"));
    var iag = find(C("rbc-ag"));
    if (iag >= 0) {
      out.hla_rbc_ag_kell = vals(iag);
      for (var jj = iag + 1; jj < Math.min(iag + 3, g.length); jj++) {
        if ((g[jj] || []).some(function (x) { return norm(x).indexOf("k+") >= 0; })) { out.hla_rbc_kell_pos = vals(jj); break; }
      }
    }
    setf("hla_prep_hla_ab", C("hla-ab(ชุด")); setf("hla_prep_hla_b27", C("hla-b27(ชุด"));
    var eqa = find(C("eqa"));
    if (eqa >= 0) {
      var names = ["hla_eqa_class1", "hla_eqa_b27", "hla_eqa_fcxm", "hla_eqa_si_ra", "hla_eqa_plt_ab", "hla_eqa_plt_xm"], cnt = 0;
      for (var e = eqa + 1; e < g.length && cnt < 6; e++) {
        var aa = labs[e].split("|")[0];
        if (aa.indexOf("13.") >= 0) break;
        if (aa.charAt(0) === "-") { out[names[cnt]] = vals(e); cnt++; }
      }
    }
    setf("hla_reject", C("13.")); setf("hla_critical", C("14."));
    return out;
  }

  function computeMonths(data) {
    var maxIdx = -1;
    Object.keys(data).forEach(function (kk) { var a = data[kk]; if (!Array.isArray(a)) return; for (var i = 0; i < a.length; i++) if (a[i] !== null && a[i] !== undefined) { if (i > maxIdx) maxIdx = i; } });
    if (maxIdx < 0) return null; var o = []; for (var j = 0; j <= maxIdx; j++) o.push(j); return o;
  }
  function matchYear(filename) { return /^สถิติ\s*69\.xlsx$/i.test(String(filename)) ? CURRENT_YEAR : null; }
  function applyYear(year, data) {
    if (year !== CURRENT_YEAR || typeof HLA2569 === "undefined") return;
    Object.keys(data).forEach(function (k) { if (HLA2569.hasOwnProperty(k)) HLA2569[k] = data[k]; });
    if (!window.__autoMonthsByYear) window.__autoMonthsByYear = {};
    var mo = computeMonths(data); if (mo) window.__autoMonthsByYear[year] = mo;
  }
  function render() {
    if (!window.__gaiPatchedHLA && typeof getAvailableIdx === "function") {
      var _gai = getAvailableIdx;
      // eslint-disable-next-line no-global-assign
      getAvailableIdx = function (y) { var by = window.__autoMonthsByYear; if (by && by[y]) return by[y]; return _gai(y); };
      window.__gaiPatchedHLA = true;
    }
    if (typeof refresh === "function") refresh();
    if (typeof buildCharts === "function") buildCharts();
  }
  DashboardAutoload.register({
    mode: "directory", storeKey: "hla", snapshotFile: "live_hla.js",
    defaultYear: CURRENT_YEAR, match: matchYear, parse: parse, applyYear: applyYear, render: render
  });
})();
