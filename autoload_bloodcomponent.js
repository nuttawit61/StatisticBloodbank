/* =====================================================================
 *  autoload_bloodcomponent.js — ตัวเชื่อมข้อมูลเฉพาะห้อง Blood Component
 *  อ่านชีต "Blood Component" -> โครงสร้าง RAW_DATA[ปี] ของ Dashboard
 *  จับตาม label (คอลัมน์ A) ; ตรวจสอบแล้วตรงกับ RAW_DATA ทุกปี
 *  (clottedUnused ปีเก่าจะแสดงค่าจริงจากไฟล์แทน 0 ที่หน้าเดิมบันทึกไว้)
 * ===================================================================== */
(function () {
  "use strict";

  var YEAR_KEYS = [2564, 2565, 2566, 2567, 2568, 2569];

  function tonum(v) {
    if (v === null || v === undefined || v === "") return null;
    if (typeof v === "number") return v;
    var s = String(v); if (s.toLowerCase().indexOf("div/0") >= 0) return null;
    var m = s.match(/-?\d+(?:\.\d+)?/); return m ? Number(m[0]) : null;
  }

  function parse(wb, H) {
    var norm = H.norm;
    var g = H.sheetGrid(wb, "Blood Component", window.XLSX);
    if (!g.length) return {};
    var hr = H.findRow(g, function (row) { return row.some(function (x) { return norm(x) === "ม.ค."; }); });
    if (hr < 0) return {};
    var ms = -1, hrow = g[hr];
    for (var j = 0; j < hrow.length; j++) if (norm(hrow[j]) === "ม.ค.") { ms = j; break; }
    var mcols = []; for (var k = 0; k < 12; k++) mcols.push(ms + k);
    var lc = 0;
    function vals(i, text) {
      var row = g[i] || [];
      return mcols.map(function (c) {
        if (c >= row.length) return text ? "" : null;
        return text ? (row[c] == null ? "" : String(row[c])) : tonum(row[c]);
      });
    }
    var rows = [];
    for (var r = hr + 1; r < g.length; r++) { var lab = (g[r] && g[r][lc] != null) ? g[r][lc] : ""; if (String(lab).trim() !== "") rows.push([r, norm(lab)]); }
    function find(pred, text) {
      var t;
      for (t = 0; t < rows.length; t++) if (pred(rows[t][1])) { var v = vals(rows[t][0], text); if (text || v.some(function (x) { return x !== null; })) return v; }
      for (t = 0; t < rows.length; t++) if (pred(rows[t][1])) return vals(rows[t][0], text);
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
    put("contaminatedCause", find(C("สาเหตุ"), true));

    // SDP (Single Donor Platelets) — ไม่ได้อยู่ในชีต Blood Component
    // ค่าจริงมาจากชีต "Hemapheresis" แถว "Product (bag)"
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
            var sdpVals = hmcols.map(function (c) { return (c !== null && c < srow.length) ? tonum(srow[c]) : null; });
            if (sdpVals.some(function (x) { return x !== null; })) put("sdp", sdpVals);
          }
        }
      }
    } catch (e) { /* ไม่มีชีต Hemapheresis ก็ข้ามไป คงค่า sdp เดิม */ }

    return out;
  }

  function monthCount(data) {
    var arr = data.donor || data.pfb || data.ffp; if (!Array.isArray(arr)) return null;
    var maxIdx = -1; for (var i = 0; i < arr.length; i++) if (arr[i] !== null && arr[i] !== undefined) maxIdx = i;
    return maxIdx + 1;
  }

  function matchYear(filename) {
    var m = String(filename).match(/^สถิติ\s*(\d{2})\.xlsx$/i);
    if (!m) return null; var y = 2500 + parseInt(m[1], 10);
    return YEAR_KEYS.indexOf(y) >= 0 ? y : null;
  }

  function applyYear(year, data) {
    if (typeof RAW_DATA === "undefined" || !RAW_DATA[year]) return;
    var obj = RAW_DATA[year];
    Object.keys(data).forEach(function (k) { if (obj.hasOwnProperty(k)) obj[k] = data[k]; });
    var mc = monthCount(data); if (mc && mc > 0) obj.available = mc;
  }

  function render() { if (typeof renderAll === "function") renderAll(); }

  DashboardAutoload.register({
    mode: "directory", storeKey: "bloodcomponent", snapshotFile: "live_bloodcomponent.js",
    defaultYear: 2569, match: matchYear, parse: parse, applyYear: applyYear, render: render
  });
})();
