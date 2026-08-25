/* =====================================================================
 *  dashboard_autoload.js
 *  ระบบอัพเดท Dashboard อัตโนมัติจากไฟล์ Excel (ไม่ต้องรันโปรแกรมใด ๆ)
 *
 *  รองรับ 2 โหมด:
 *   - mode "directory" : เชื่อม "ทั้งโฟลเดอร์" ครั้งเดียว แล้วเฝ้าดูทุกไฟล์
 *                        สถิติ XX.xlsx -> ปีไหนถูกแก้ ปีนั้นอัพเดทเอง (แนะนำ)
 *   - mode "file"      : เชื่อมไฟล์เดียว
 *
 *  ทำงานด้วย File System Access API ของ Chrome/Edge + SheetJS (xlsx.full.min.js)
 *  ตรวจทุก ~3 วินาทีว่าไฟล์ถูกแก้ไข (เซฟ) หรือยัง ถ้าใช่ -> อ่านใหม่ + เรนเดอร์ใหม่
 *
 *  หน้า Dashboard เรียกใช้ผ่าน:
 *     DashboardAutoload.register({ mode, match, parse, applyYear, render })   // directory
 *     DashboardAutoload.register({ parse, apply })                            // file
 * ===================================================================== */
(function () {
  "use strict";

  // ---------- ตัวช่วยอ่านชีต (ใช้ร่วมกันทุกห้อง) ----------
  var THAI_FULL = ["มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน",
                   "กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม"];
  var THAI_ABBR = ["ม.ค.","ก.พ.","มี.ค.","เม.ย.","พ.ค.","มิ.ย.",
                   "ก.ค.","ส.ค.","ก.ย.","ต.ค.","พ.ย.","ธ.ค."];

  function norm(s) {
    if (s === null || s === undefined) return "";
    return String(s).trim().toLowerCase().replace(/\s+/g, "");
  }
  function num(v) {
    if (v === null || v === undefined || v === "") return null;
    if (typeof v === "number") return v;
    var n = Number(String(v).replace(/,/g, ""));
    return isNaN(n) ? null : n;
  }
  function sheetGrid(wb, name, XLSX) {
    var ws = wb.Sheets[name];
    if (!ws) return [];
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  }
  function findRow(grid, pred, start) {
    for (var i = (start || 0); i < grid.length; i++) if (pred(grid[i] || [])) return i;
    return -1;
  }
  function monthIndexFull(v) {
    var n = (v === null || v === undefined) ? "" : String(v).trim();
    return THAI_FULL.indexOf(n);
  }
  var H = { THAI_FULL: THAI_FULL, THAI_ABBR: THAI_ABBR, norm: norm, num: num,
    sheetGrid: sheetGrid, findRow: findRow, monthIndexFull: monthIndexFull };

  // ---------- IndexedDB เก็บ handle ----------
  var DB_NAME = "dashboard-autoload", STORE = "handles";
  function idb() {
    return new Promise(function (res, rej) {
      var r = indexedDB.open(DB_NAME, 1);
      r.onupgradeneeded = function () { r.result.createObjectStore(STORE); };
      r.onsuccess = function () { res(r.result); };
      r.onerror = function () { rej(r.error); };
    });
  }
  function idbSet(k, v) { return idb().then(function (db) { return new Promise(function (res, rej) {
    var tx = db.transaction(STORE, "readwrite"); tx.objectStore(STORE).put(v, k);
    tx.oncomplete = function () { res(); }; tx.onerror = function () { rej(tx.error); }; }); }); }
  function idbGet(k) { return idb().then(function (db) { return new Promise(function (res, rej) {
    var tx = db.transaction(STORE, "readonly"); var g = tx.objectStore(STORE).get(k);
    g.onsuccess = function () { res(g.result); }; g.onerror = function () { rej(g.error); }; }); }); }
  function idbDel(k) { return idb().then(function (db) { return new Promise(function (res) {
    var tx = db.transaction(STORE, "readwrite"); tx.objectStore(STORE).delete(k);
    tx.oncomplete = function () { res(); }; }); }); }

  // คีย์ "ร่วมกันทุกหน้า" — เชื่อมโฟลเดอร์ครั้งเดียวที่หน้าใดก็ได้ ทุก Dashboard ใช้โฟลเดอร์เดียวกัน
  var STORAGE_KEY = "dashboard-folder-handle";

  // ---------- บันทึกข้อมูลล่าสุดถาวร (localStorage) ----------
  // ทำให้หน้ายังแสดงข้อมูลอัพเดทล่าสุด แม้ยังไม่เชื่อม/ตัดการเชื่อมแล้ว/ปิดเปิดใหม่
  function snapKey(cfg) { return "dash_snap:" + (cfg.storeKey || location.pathname); }
  function loadSnap(cfg) {
    try { var s = localStorage.getItem(snapKey(cfg)); return s ? JSON.parse(s) : {}; }
    catch (e) { return {}; }
  }
  function saveSnap(cfg, snap) {
    try { localStorage.setItem(snapKey(cfg), JSON.stringify(snap)); } catch (e) {}
  }
  function clearSnap(cfg) { try { localStorage.removeItem(snapKey(cfg)); } catch (e) {} }

  // เขียนข้อมูลล่าสุดกลับลงไฟล์ live_<lab>.js ในโฟลเดอร์ (ฝังถาวร/แชร์/ย้ายเครื่องได้)
  function writeSnapshotFile(cfg, snap) {
    if (!cfg.snapshotFile || !dirHandleRef || !dirHandleRef.getFileHandle) return Promise.resolve();
    var content =
      "/* ไฟล์นี้สร้างอัตโนมัติโดยระบบอัพเดท Dashboard — เก็บข้อมูลล่าสุดจาก Excel\n" +
      "   เปิดหน้า Dashboard ที่เครื่องใดก็จะเห็นข้อมูลนี้ทันที โดยไม่ต้องเชื่อมโฟลเดอร์\n" +
      "   อัพเดทล่าสุด: " + new Date().toLocaleString("th-TH") + " */\n" +
      "window.__LIVE__ = window.__LIVE__ || {};\n" +
      "window.__LIVE__[" + JSON.stringify(cfg.storeKey) + "] = " + JSON.stringify(snap) + ";\n";
    return dirHandleRef.getFileHandle(cfg.snapshotFile, { create: true })
      .then(function (fh) { return fh.createWritable(); })
      .then(function (w) { return w.write(content).then(function () { return w.close(); }); })
      .catch(function () { /* ไม่มีสิทธิ์เขียน — ยังมี localStorage สำรองอยู่ */ });
  }

  // ---------- UI ----------
  var bar, statusEl, btnEl, hintEl;
  function buildUI(cfg) {
    bar = document.createElement("div");
    bar.id = "autoload-bar";
    bar.innerHTML =
      '<style>' +
      '#autoload-bar{position:fixed;left:50%;bottom:18px;transform:translateX(-50%);z-index:99999;' +
      'display:flex;align-items:center;gap:12px;background:rgba(20,24,40,.92);color:#fff;' +
      'padding:10px 16px;border-radius:30px;box-shadow:0 8px 28px rgba(0,0,0,.35);' +
      'font-family:"Sarabun","Segoe UI",sans-serif;font-size:14px;backdrop-filter:blur(8px);' +
      'border:1px solid rgba(255,255,255,.12);max-width:94vw}' +
      '#autoload-bar .al-dot{width:9px;height:9px;border-radius:50%;background:#9aa3b2;flex:0 0 auto}' +
      '#autoload-bar.live .al-dot{background:#28d17c;animation:alPulse 2s infinite}' +
      '@keyframes alPulse{0%{box-shadow:0 0 0 0 rgba(40,209,124,.55)}70%{box-shadow:0 0 0 9px rgba(40,209,124,0)}100%{box-shadow:0 0 0 0 rgba(40,209,124,0)}}' +
      '#autoload-bar button{cursor:pointer;border:0;border-radius:20px;padding:8px 16px;font-weight:600;' +
      'font-family:inherit;font-size:14px;background:#3b82f6;color:#fff}' +
      '#autoload-bar button:hover{background:#2563eb}' +
      '#autoload-bar .al-link{cursor:pointer;color:#9ec5ff;text-decoration:underline;font-size:12.5px}' +
      '#autoload-bar .al-status{white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:48vw}' +
      '#autoload-bar.flash{animation:alFlash .9s ease}' +
      '@keyframes alFlash{0%{background:rgba(40,209,124,.95)}100%{background:rgba(20,24,40,.92)}}' +
      '</style>' +
      '<span class="al-dot"></span>' +
      '<span class="al-status" id="al-status">ยังไม่ได้เชื่อมข้อมูล</span>' +
      '<button id="al-btn">📂 เชื่อมโฟลเดอร์ข้อมูล</button>' +
      '<span class="al-link" id="al-hint" style="display:none"></span>';
    document.body.appendChild(bar);
    statusEl = document.getElementById("al-status");
    btnEl = document.getElementById("al-btn");
    hintEl = document.getElementById("al-hint");
    if (cfg.mode !== "directory") btnEl.textContent = "📂 เชื่อมไฟล์ Excel";
    btnEl.addEventListener("click", function () { onConnectClick(cfg); });
  }
  function setStatus(t, live) { if (statusEl) statusEl.textContent = t; if (bar) bar.classList.toggle("live", !!live); }
  function flash() { if (!bar) return; bar.classList.remove("flash"); void bar.offsetWidth; bar.classList.add("flash"); }
  function fmtTime(d) { function p(n){return (n<10?"0":"")+n;} return p(d.getHours())+":"+p(d.getMinutes())+":"+p(d.getSeconds()); }

  // ---------- core ----------
  var cfgRef = null, pollTimer = null;
  var fileHandleRef = null, lastModifiedSingle = 0;        // file mode
  var dirHandleRef = null, fileMeta = {};                  // directory mode: name -> lastModified

  /* ===== file mode ===== */
  function fileApply(buf) {
    var wb = XLSX.read(buf, { type: "array" });
    var data = cfgRef.parse(wb, H);
    cfgRef.apply(data);
    var snap = loadSnap(cfgRef); snap.__file__ = data; saveSnap(cfgRef, snap);
    flash(); setStatus("🟢 อัพเดทแล้ว " + fmtTime(new Date()) + " · " + (fileHandleRef ? fileHandleRef.name : ""), true);
  }
  function fileReadApply() {
    return fileHandleRef.getFile().then(function (f) { lastModifiedSingle = f.lastModified; return f.arrayBuffer(); }).then(fileApply);
  }

  /* ===== directory mode ===== */
  function dirScanAll(initial) {
    var changed = 0, years = [], snap = loadSnap(cfgRef);
    return (function () {
      var it = dirHandleRef.entries();
      function step() {
        return it.next().then(function (r) {
          if (r.done) return;
          var name = r.value[0], handle = r.value[1];
          if (handle.kind === "file") {
            var year = cfgRef.match(name);
            if (year) {
              return handle.getFile().then(function (file) {
                if (fileMeta[name] !== file.lastModified) {
                  fileMeta[name] = file.lastModified;
                  return file.arrayBuffer().then(function (buf) {
                    try {
                      var wb = XLSX.read(buf, { type: "array" });
                      var data = cfgRef.parse(wb, H);
                      cfgRef.applyYear(year, data);
                      snap[year] = data;            // เก็บไว้บันทึกถาวร
                      changed++; years.push(year);
                    } catch (e) { /* ข้ามไฟล์ที่อ่านไม่ได้ */ }
                  });
                }
              }).then(step).catch(step);
            }
          }
          return step();
        });
      }
      return step();
    })().then(function () {
      if (changed > 0) {
        saveSnap(cfgRef, snap);                     // บันทึกลงเครื่อง (localStorage)
        writeSnapshotFile(cfgRef, snap);            // เขียนถาวรลงไฟล์ live_<lab>.js
        if (cfgRef.render) cfgRef.render();
        flash();
        years.sort();
        setStatus("🟢 อัพเดท+บันทึกลงไฟล์แล้ว " + fmtTime(new Date()) + " · ปี " + years.map(function (y) { return y; }).join(", "), true);
      } else if (initial) {
        setStatus("🟢 เฝ้าดูโฟลเดอร์: " + dirHandleRef.name + " · พร้อมอัพเดทอัตโนมัติ", true);
      }
    });
  }

  function startPolling() {
    if (pollTimer) clearInterval(pollTimer);
    pollTimer = setInterval(function () {
      if (cfgRef.mode === "directory") { if (dirHandleRef) dirScanAll(false).catch(function () {}); }
      else if (fileHandleRef) {
        fileHandleRef.getFile().then(function (f) {
          if (f.lastModified !== lastModifiedSingle) { lastModifiedSingle = f.lastModified; return f.arrayBuffer().then(fileApply); }
        }).catch(function () {});
      }
    }, 3000);
  }

  function showDisconnect(cfg) {
    hintEl.style.display = "inline";
    hintEl.textContent = "ตัดการเชื่อม";
    hintEl.onclick = function () {
      idbDel(STORAGE_KEY); dirHandleRef = null; fileHandleRef = null; fileMeta = {};
      if (pollTimer) clearInterval(pollTimer);
      setStatus("ตัดการเชื่อมแล้ว", false);
      btnEl.textContent = cfg.mode === "directory" ? "📂 เชื่อมโฟลเดอร์ข้อมูล" : "📂 เชื่อมไฟล์ Excel";
      btnEl.onclick = function () { onConnectClick(cfg); };
      hintEl.style.display = "none";
    };
  }

  function beginDirectory(cfg) {
    btnEl.textContent = "เปลี่ยนโฟลเดอร์";
    showDisconnect(cfg);
    setStatus("กำลังอ่านไฟล์ในโฟลเดอร์ " + dirHandleRef.name + " …", true);
    dirScanAll(true).then(startPolling).catch(function (e) { setStatus("อ่านโฟลเดอร์ไม่สำเร็จ: " + e.message, false); });
  }
  function beginFile(cfg) {
    btnEl.textContent = "เปลี่ยนไฟล์";
    showDisconnect(cfg);
    setStatus("กำลังอ่านไฟล์ " + fileHandleRef.name + " …", true);
    fileReadApply().then(startPolling).catch(function (e) { setStatus("อ่านไฟล์ไม่สำเร็จ: " + e.message, false); });
  }

  function onConnectClick(cfg) {
    cfgRef = cfg;
    if (!window.showOpenFilePicker) { fallbackPicker(); return; }
    if (cfg.mode === "directory") {
      window.showDirectoryPicker({ mode: "readwrite" }).then(function (h) {
        dirHandleRef = h; fileMeta = {}; return idbSet(STORAGE_KEY, h);
      }).then(function () { beginDirectory(cfg); })
        .catch(function (e) { if (!(e && e.name === "AbortError")) setStatus("เกิดข้อผิดพลาด: " + e.message, false); });
    } else {
      var opts = { types: [{ description: "Excel", accept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"] } }], multiple: false };
      window.showOpenFilePicker(opts).then(function (hs) { fileHandleRef = hs[0]; return idbSet(STORAGE_KEY, fileHandleRef); })
        .then(function () { beginFile(cfg); })
        .catch(function (e) { if (!(e && e.name === "AbortError")) setStatus("เกิดข้อผิดพลาด: " + e.message, false); });
    }
  }

  // เบราว์เซอร์ที่ไม่รองรับ (Firefox/Safari) -> โหลดครั้งเดียว (ไฟล์เดียว)
  function fallbackPicker() {
    var inp = document.createElement("input");
    inp.type = "file"; inp.accept = ".xlsx,.xls";
    inp.onchange = function () {
      var f = inp.files[0]; if (!f) return;
      setStatus("กำลังอ่านไฟล์ " + f.name + " …", true);
      f.arrayBuffer().then(function (buf) {
        var wb = XLSX.read(buf, { type: "array" });
        if (cfgRef.mode === "directory") {
          var year = cfgRef.match(f.name) || (cfgRef.defaultYear || null);
          if (year) { cfgRef.applyYear(year, cfgRef.parse(wb, H)); if (cfgRef.render) cfgRef.render(); }
        } else { cfgRef.apply(cfgRef.parse(wb, H)); }
        flash();
        setStatus("🟡 โหลดข้อมูลแล้ว (เบราว์เซอร์นี้ไม่รองรับอัพเดทอัตโนมัติ — ใช้ Chrome/Edge)", false);
      });
    };
    inp.click();
  }

  // เปิดหน้าใหม่: พยายามเชื่อมของเดิมอัตโนมัติ
  function tryRestore(cfg) {
    cfgRef = cfg;
    if (!window.showOpenFilePicker) { setStatus("เบราว์เซอร์นี้ไม่รองรับอัพเดทอัตโนมัติ — แนะนำ Chrome หรือ Edge", false); return; }
    idbGet(STORAGE_KEY).then(function (h) {
      if (!h) return;
      var isDir = (h.kind === "directory");
      var pmode = isDir ? "readwrite" : "read";   // โฟลเดอร์ต้องเขียนไฟล์ live ได้
      h.queryPermission({ mode: pmode }).then(function (perm) {
        if (perm === "granted") {
          if (isDir) { dirHandleRef = h; beginDirectory(cfg); } else { fileHandleRef = h; beginFile(cfg); }
        } else {
          setStatus("คลิกเพื่อเชื่อมต่อ " + (isDir ? "โฟลเดอร์" : "ไฟล์") + "เดิม: " + h.name, false);
          btnEl.textContent = "🔄 เชื่อมต่อของเดิม";
          btnEl.onclick = function () {
            h.requestPermission({ mode: pmode }).then(function (p) {
              if (p === "granted") {
                btnEl.onclick = function () { onConnectClick(cfg); };
                if (isDir) { dirHandleRef = h; beginDirectory(cfg); } else { fileHandleRef = h; beginFile(cfg); }
              } else setStatus("ไม่ได้รับสิทธิ์เข้าถึงข้อมูล", false);
            });
          };
        }
      });
    }).catch(function () {});
  }

  // ตอนเปิดหน้า: ถ้ามีข้อมูลที่บันทึกไว้ ให้แสดงทันที (แม้ยังไม่เชื่อม/ตัดการเชื่อมแล้ว)
  // ลำดับความสำคัญ: ไฟล์ live_<lab>.js (แชร์/ย้ายเครื่องได้) ก่อน แล้วค่อย localStorage
  function applyOfflineSnapshot(cfg) {
    var fileSnap = (window.__LIVE__ && cfg.storeKey) ? window.__LIVE__[cfg.storeKey] : null;
    var snap = fileSnap || loadSnap(cfg);
    if (fileSnap) saveSnap(cfg, fileSnap);   // ซิงค์เข้าเครื่องด้วย
    var keys = snap ? Object.keys(snap) : [];
    if (!keys.length) return false;
    var applied = 0;
    if (cfg.mode === "directory") {
      keys.forEach(function (y) { if (y === "__file__") return; try { cfgRef = cfg; cfg.applyYear(Number(y), snap[y]); applied++; } catch (e) {} });
    } else if (snap.__file__) { try { cfg.apply(snap.__file__); applied++; } catch (e) {} }
    if (applied && cfg.render) cfg.render();
    if (applied) setStatus("📁 แสดงข้อมูลที่บันทึกไว้ล่าสุด · กดเชื่อมเพื่ออัพเดทต่อ", false);
    return applied > 0;
  }

  // ---------- public ----------
  window.DashboardAutoload = {
    helpers: H,
    register: function (cfg) {
      cfgRef = cfg;
      function go() { buildUI(cfg); applyOfflineSnapshot(cfg); tryRestore(cfg); }
      if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", go);
      else go();
    },
    // ให้หน้ารวมเรียกเปิดตัวเลือกโฟลเดอร์ได้ (เชื่อมครั้งเดียวใช้ทุก Dashboard)
    connect: function () { if (cfgRef) onConnectClick(cfgRef); },
    isConnected: function () { return !!(dirHandleRef || fileHandleRef); }
  };
})();
