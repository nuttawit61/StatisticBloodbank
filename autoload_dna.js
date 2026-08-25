/* =====================================================================
 *  autoload_dna.js — ตัวเชื่อมข้อมูลเฉพาะห้อง DNA Laboratory
 *  อ่าน 2 ชีต:
 *    "DNA"     — ตารางแบบใหม่ (กลุ่ม BMT/Solid Organ/Post KT + Method summary
 *                + ADR + KPI) -> คีย์ dna_* ของ DNA2569
 *    "Luminex" — เดือนละ 3 คอลัมน์ (Service ศิริราช + ร.พ.อื่น ๆ + EQAS)
 *                -> คีย์ lum_* + lum_eqas
 *  อัพเดทเฉพาะปีปัจจุบัน (2569) — ปีเก่าโครงสร้างต่างกัน คงข้อมูลเดิมไว้
 *  ตรวจสอบแล้ว: สถิติ 69.xlsx ตรงกับ DNA2569 ครบทุกฟิลด์
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
  function methodSuffix(cn) {
    if (cn === "ngs") return "ngs";
    if (cn === "sso") return "sso";
    if (cn === "ssp") return "ssp";
    if (cn.indexOf("pcr") >= 0) return "pcr";   // RT-PCR
    return null;
  }

  function parse(wb, H) {
    var norm = H.norm, XLSX = window.XLSX, out = {};

    // ---- DNA (รองรับชื่อชีตใหม่ "DNA" และของเดิม "DNA1") ----
    var dnaName = wb.Sheets["DNA"] ? "DNA" : (wb.Sheets["DNA1"] ? "DNA1" : null);
    if (dnaName) {
      var g = H.sheetGrid(wb, dnaName, XLSX);
      // หาแถวหัวเดือน + คอลัมน์ ม.ค.
      var hr = -1, ms = -1;
      for (var i = 0; i < g.length && hr < 0; i++) {
        var row = g[i] || [];
        for (var j = 0; j < row.length; j++) {
          if (norm(row[j]) === norm("ม.ค.")) { hr = i; ms = j; break; }
        }
      }
      if (hr >= 0) {
        var mcols = []; for (var k = 0; k < 12; k++) mcols.push(ms + k);
        var vals = function (r) {
          var row = g[r] || [];
          return mcols.map(function (c) { return c < row.length ? tonum(row[c]) : null; });
        };
        var curGroup = null, pastMethod = false, adrStarted = false;
        for (var r = 0; r < g.length; r++) {
          var rw = g[r] || [];
          var A = rw.length > 0 ? rw[0] : null;
          var B = rw.length > 1 ? rw[1] : null;
          var C = rw.length > 2 ? rw[2] : null;
          var an = norm(A), bn = norm(B), cn = norm(C);

          // อัพเดทกลุ่มปัจจุบันจากคอลัมน์ B (เซลล์ผสาน — ปรากฏครั้งเดียวต่อกลุ่ม)
          if (bn) {
            if (bn.indexOf("pt.bmt") >= 0) curGroup = "dna_bmt_pt";
            else if (bn.indexOf("dbmt") >= 0) curGroup = "dna_bmt_d";
            else if (bn.indexOf("kt") >= 0 && bn.indexOf("pt") >= 0) curGroup = "dna_kt_pt";
            else if (bn.indexOf("lrd") >= 0) curGroup = "dna_lrd";
            else if (bn.indexOf("dqa") >= 0) curGroup = "__dqa";
            else if (bn.indexOf("dpa") >= 0 || bn.indexOf("dpb") >= 0) curGroup = "__dpa";
          }

          // KPI (จับจากคอลัมน์ A — ป้ายไม่ซ้ำ)
          if (an.indexOf("firstpass") >= 0) out.dna_kpi_firstpass = vals(r);
          else if (an.indexOf("repeat") >= 0) out.dna_kpi_repeat = vals(r);

          // ADR / HLA disease association (จับจากคอลัมน์ C — ป้ายไม่ซ้ำ)
          var key = cn.replace(/[^a-z0-9]/g, "");
          if (cn.indexOf("hlafor") >= 0 || key === "2569") adrStarted = true;
          if (key === "b1502") { out.dna_adr_b1502 = vals(r); adrStarted = true; }
          else if (key === "b5801") { out.dna_adr_b5801 = vals(r); adrStarted = true; }
          else if (key === "b5701") { out.dna_adr_b5701 = vals(r); adrStarted = true; }
          else if (key === "b51") { out.dna_adr_b51 = vals(r); adrStarted = true; }

          // หัวตาราง Method summary
          if (cn === "method") { pastMethod = true; continue; }

          var suf = methodSuffix(cn);
          if (!pastMethod) {
            // ส่วนรายกลุ่ม (ก่อนถึง Method summary)
            if (suf && curGroup) {
              if (curGroup === "__dqa") out.dna_post_dqa = vals(r);
              else if (curGroup === "__dpa") out.dna_post_dpadpb = vals(r);
              else out[curGroup + "_" + suf] = vals(r);
            }
          } else {
            // ส่วน Method summary (dna_m_* + dna_total)
            if (suf && !bn) out["dna_m_" + suf] = vals(r);
            else if (cn === "total" && !bn && !adrStarted && !out.dna_total) out.dna_total = vals(r);
          }
        }
      }
    }

    // ---- Luminex ----
    if (wb.Sheets["Luminex"]) {
      var gl = H.sheetGrid(wb, "Luminex", XLSX);
      var hr2 = H.findRow(gl, function (row) { return row.some(function (x) { return norm(x) === norm("ม.ค."); }); });
      if (hr2 >= 0) {
        var base = -1; for (var b = 0; b < gl[hr2].length; b++) if (norm(gl[hr2][b]) === norm("ม.ค.")) { base = b; break; }
        // ผลรวม Service (ศิริราช + ร.พ.อื่น ๆ) = คอลัมน์ที่ 1+2 ของแต่ละเดือน
        var lsum = function (i) {
          var row = gl[i] || [], o = [];
          for (var m = 0; m < 12; m++) {
            var c1 = base + 3 * m, c2 = base + 3 * m + 1;
            var a = c1 < row.length ? tonum(row[c1]) : null, bb = c2 < row.length ? tonum(row[c2]) : null;
            o.push((a === null && bb === null) ? null : (a || 0) + (bb || 0));
          }
          return o;
        };
        // EQAS = คอลัมน์ที่ 3 ของแต่ละเดือน
        var leqas = function (i) {
          var row = gl[i] || [], o = [];
          for (var m = 0; m < 12; m++) {
            var c3 = base + 3 * m + 2;
            o.push(c3 < row.length ? tonum(row[c3]) : null);
          }
          return o;
        };
        var llabs = gl.map(function (row) { return (row && row[0] != null) ? norm(row[0]) : ""; });
        var occ = function (sub) { var n = norm(sub), r = []; for (var i = 0; i < llabs.length; i++) if (llabs[i].indexOf(n) >= 0) r.push(i); return r; };
        var one = function (kk, sub) { var r = occ(sub); if (r.length) out[kk] = lsum(r[0]); };
        one("lum_screening", "hlaantibodyscreening");
        var pr = occ("specificprahlaclassi"); if (pr[0] != null) out.lum_pra1 = lsum(pr[0]); if (pr[1] != null) out.lum_pra2 = lsum(pr[1]);
        var sa = occ("singleantigenhlaclass"); ["lum_sa1", "lum_sa2", "lum_sa1_c1q", "lum_sa2_c1q"].forEach(function (kk, idx) { if (sa[idx] != null) out[kk] = lsum(sa[idx]); });
        one("lum_mica_geno", "micagenotype"); one("lum_mica_ab", "micaantibody"); one("lum_trali", "trali");
        var tt = occ("total"); if (tt.length) { for (var z = 0; z < tt.length; z++) if (llabs[tt[z]] === "total") { out.lum_total = lsum(tt[z]); out.lum_eqas = leqas(tt[z]); break; } }
      }
    }
    return out;
  }

  function computeMonths(data) {
    var maxIdx = -1, inds = ["dna_m_ngs", "dna_adr_b1502", "dna_total", "lum_screening"];
    inds.forEach(function (key) { var a = data[key]; if (!Array.isArray(a)) return; for (var i = 0; i < a.length; i++) if (a[i] !== null && a[i] !== undefined) { if (i > maxIdx) maxIdx = i; } });
    if (maxIdx < 0) return null; var o = []; for (var j = 0; j <= maxIdx; j++) o.push(j); return o;
  }
  function matchYear(filename) { return /^สถิติ\s*69\.xlsx$/i.test(String(filename)) ? CURRENT_YEAR : null; }
  function applyYear(year, data) {
    if (year !== CURRENT_YEAR || typeof DNA2569 === "undefined") return;
    Object.keys(data).forEach(function (k) { if (DNA2569.hasOwnProperty(k)) DNA2569[k] = data[k]; });
    if (!window.__autoMonthsByYear) window.__autoMonthsByYear = {};
    var mo = computeMonths(data); if (mo) window.__autoMonthsByYear[year] = mo;
  }
  function render() {
    if (!window.__gaiPatchedDNA && typeof getAvailableIdx === "function") {
      var _gai = getAvailableIdx;
      // eslint-disable-next-line no-global-assign
      getAvailableIdx = function (y) { var by = window.__autoMonthsByYear; if (by && by[y]) return by[y]; return _gai(y); };
      window.__gaiPatchedDNA = true;
    }
    if (typeof refresh === "function") refresh();
    if (typeof buildCharts === "function") buildCharts();
  }
  DashboardAutoload.register({
    mode: "directory", storeKey: "dna", snapshotFile: "live_dna.js",
    defaultYear: CURRENT_YEAR, match: matchYear, parse: parse, applyYear: applyYear, render: render
  });
})();
