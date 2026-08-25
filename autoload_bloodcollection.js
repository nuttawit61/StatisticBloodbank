/* =====================================================================
 *  autoload_bloodcollection.js — ตัวเชื่อมข้อมูลเฉพาะห้อง Blood Collection
 *  อ่านชีต "Blood Collection" จากไฟล์ สถิติ XX.xlsx -> โครงสร้าง BC ของ Dashboard
 *  จับข้อมูลตาม "section (ลำดับในคอลัมน์ A)" + หัวข้อย่อย (ในสถานที่/หน่วยฯ/Total/ชาย/หญิง)
 *  ตรวจสอบแล้ว: ตรงกับ BC64–BC69 (ยกเว้น 2 ฟิลด์ที่หน้าเดิมเว้นว่างแต่ Excel มีค่าจริง
 *  คือ directOnsite, voluntaryExclMobile — ระบบจะแสดงค่าจริงจากไฟล์)
 * ===================================================================== */
(function () {
  "use strict";

  var YEARS = { 2564: "BC64", 2565: "BC65", 2566: "BC66", 2567: "BC67", 2568: "BC68", 2569: "BC69" };

  function tonum(v) {
    if (v === null || v === undefined || v === "") return null;
    if (typeof v === "number") return v;
    var s = String(v);
    if (s.toLowerCase().indexOf("div/0") >= 0) return null;
    var m = s.match(/-?\d+(?:\.\d+)?/);
    return m ? Number(m[0]) : null;
  }

  function parse(wb, H) {
    var norm = H.norm;
    var g = H.sheetGrid(wb, "Blood Collection", window.XLSX);
    if (!g.length) return {};
    var hr = H.findRow(g, function (row) { return row.some(function (x) { return norm(x) === "ม.ค."; }); });
    if (hr < 0) return {};
    var ms = -1, hrow = g[hr];
    for (var j = 0; j < hrow.length; j++) if (norm(hrow[j]) === "ม.ค.") { ms = j; break; }
    var mcols = []; for (var k = 0; k < 12; k++) mcols.push(ms + k);
    var lc = 1; for (var j2 = 0; j2 < hrow.length; j2++) if (norm(hrow[j2]) === "รายการ") { lc = j2; break; }

    function vals(i) { var row = g[i] || []; return mcols.map(function (c) { return c < row.length ? tonum(row[c]) : null; }); }

    // sections = แถวที่คอลัมน์ A เป็นจำนวนเต็ม
    var secs = [];
    for (var r = hr + 1; r < g.length; r++) {
      var a = g[r] ? g[r][0] : null;
      if (typeof a === "number" && Math.floor(a) === a) secs.push([r, norm(g[r][lc])]);
    }
    secs.push([g.length, ""]);
    function rngOf(idx) { return [secs[idx][0], secs[idx + 1][0]]; }
    function findSecs(pred) { var o = []; for (var t = 0; t < secs.length - 1; t++) if (pred(secs[t][1])) o.push(t); return o; }
    function hasSub(rg, pred) { for (var i = rg[0]; i < rg[1]; i++) if (pred(norm(g[i][lc]))) return true; return false; }
    function sub(rg, pred) {
      for (var i = rg[0] + 1; i < rg[1]; i++) {
        var lab = norm(g[i][lc]); if (lab.charAt(0) === "%") continue;
        if (pred(lab)) return vals(i);
      }
      return null;
    }
    function secrow(rg) { return vals(rg[0]); }
    var C = function (s) { return function (l) { return l.indexOf(norm(s)) >= 0; }; };
    var EQ = function (s) { return function (l) { return l === norm(s); }; };
    var onsite = function (l) { return l.indexOf("ในสถานที่") >= 0; };
    var mobile = function (l) { return l.indexOf("หน่วยฯ") >= 0 || l === "หน่วย"; };
    var total = function (l) { return l === "total"; };

    var out = {};
    function put(name, arr) { if (arr) out[name] = arr; }

    function trio(secpred, base) {
      var ks = findSecs(secpred).filter(function (k) { return hasSub(rngOf(k), total); });
      if (!ks.length) return;
      var rg = rngOf(ks[0]);
      put(base + "Onsite", sub(rg, onsite)); put(base + "Mobile", sub(rg, mobile)); put(base + "Total", sub(rg, total));
    }
    trio(EQ("donorvisit"), "visit");
    trio(EQ("defer"), "defer");
    trio(function (l) { return l.indexOf("donation") === 0; }, "donation");
    trio(EQ("ผู้บริจาคใหม่"), "new");
    trio(EQ("ผู้บริจาคเก่า"), "old");
    trio(EQ("replacement"), "replace");

    var ks = findSecs(EQ("direct")); if (ks.length) put("directOnsite", sub(rngOf(ks[0]), onsite));
    ks = findSecs(EQ("autologous")); if (ks.length) put("autologous", secrow(rngOf(ks[0])));
    ks = findSecs(function (l) { return l.indexOf("voluntary(unit)") === 0; });
    if (ks.length) { var rv = rngOf(ks[0]); put("voluntary", secrow(rv)); put("voluntaryExclMobile", sub(rv, C("ไม่รวมหน่วย"))); }
    ks = findSecs(C("bloodletting")); if (ks.length) put("bloodLetting", secrow(rngOf(ks[0])));
    ks = findSecs(EQ("donorreaction"));
    if (ks.length) { var rr = rngOf(ks[0]); put("reactionOnsite", sub(rr, onsite)); put("reactionMobile", sub(rr, mobile)); put("reactionTotal", sub(rr, total)); }

    // แยกเพศ: section ที่ไม่มี Total แต่มี "ชาย"
    function genderSecs(label) { return findSecs(EQ(label)).filter(function (k) { return hasSub(rngOf(k), function (l) { return l === "ชาย"; }); }); }
    var go = genderSecs("ผู้บริจาคเก่า"), gn = genderSecs("ผู้บริจาคใหม่");
    if (go.length) { var rgo = rngOf(go[0]); put("oldMaleOnsite", sub(rgo, EQ("ชาย"))); put("oldFemaleOnsite", sub(rgo, EQ("หญิง"))); put("oldMaleMobile", sub(rgo, C("ชาย(หน่วย)"))); put("oldFemaleMobile", sub(rgo, C("หญิง(หน่วย)"))); }
    if (gn.length) { var rgn = rngOf(gn[0]); put("newMaleOnsite", sub(rgn, EQ("ชาย"))); put("newFemaleOnsite", sub(rgn, EQ("หญิง"))); put("newMaleMobile", sub(rgn, C("ชาย(หน่วย)"))); put("newFemaleMobile", sub(rgn, C("หญิง(หน่วย)"))); }

    ks = findSecs(C("จำนวนหน่วยเคลื่อนที่")); if (ks.length) put("drivesConducted", secrow(rngOf(ks[0])));
    ks = findSecs(C("นัดหมายออนไลน์")); if (ks.length) put("drivesMeetingTarget", secrow(rngOf(ks[0])));
    ks = findSecs(C("บุคลากรศิริราช"));
    if (ks.length) { var rf = rngOf(ks[0]); put("facultyDonors", secrow(rf)); put("externalDonors", sub(rf, C("หน่วยเคลื่อนที่"))); }
    return out;
  }

  // นับเดือนจากฟิลด์ดิบ (visitOnsite/Mobile ที่ว่าง = null ไม่ใช่ 0)
  function computeMonths(data) {
    var maxIdx = -1, inds = ["visitOnsite", "visitMobile", "donationOnsite", "donationMobile"];
    inds.forEach(function (key) {
      var arr = data[key]; if (!Array.isArray(arr)) return;
      for (var i = 0; i < arr.length; i++) if (arr[i] !== null && arr[i] !== undefined) { if (i > maxIdx) maxIdx = i; }
    });
    if (maxIdx < 0) return null;
    var o = []; for (var j = 0; j <= maxIdx; j++) o.push(j); return o;
  }

  function matchYear(filename) {
    var m = String(filename).match(/^สถิติ\s*(\d{2})\.xlsx$/i);
    if (!m) return null;
    var y = 2500 + parseInt(m[1], 10);
    return YEARS[y] ? y : null;
  }
  function objFor(year) {
    switch (year) {
      case 2564: return (typeof BC64 !== "undefined") ? BC64 : null;
      case 2565: return (typeof BC65 !== "undefined") ? BC65 : null;
      case 2566: return (typeof BC66 !== "undefined") ? BC66 : null;
      case 2567: return (typeof BC67 !== "undefined") ? BC67 : null;
      case 2568: return (typeof BC68 !== "undefined") ? BC68 : null;
      case 2569: return (typeof BC69 !== "undefined") ? BC69 : null;
    }
    return null;
  }

  function applyYear(year, data) {
    if (!YEARS[year]) return;
    var obj = objFor(year);
    if (obj) Object.keys(data).forEach(function (k) { if (obj.hasOwnProperty(k)) obj[k] = data[k]; });
    if (!window.__autoMonthsByYear) window.__autoMonthsByYear = {};
    var mo = computeMonths(data); if (mo) window.__autoMonthsByYear[year] = mo;
  }

  function render() {
    if (!window.__gaiPatchedBC && typeof getAvailableIdx === "function") {
      var _gai = getAvailableIdx;
      // eslint-disable-next-line no-global-assign
      getAvailableIdx = function (y) { var by = window.__autoMonthsByYear; if (by && by[y]) return by[y]; return _gai(y); };
      window.__gaiPatchedBC = true;
    }
    if (typeof refresh === "function") refresh();
    if (typeof buildCharts === "function") buildCharts();
  }

  DashboardAutoload.register({
    mode: "directory", storeKey: "bloodcollection", snapshotFile: "live_bloodcollection.js",
    defaultYear: 2569, match: matchYear, parse: parse, applyYear: applyYear, render: render
  });
})();
