/* =====================================================================
 *  autoload_hub.js — หน้ารวม (index)
 *  ปุ่มเชื่อมโฟลเดอร์ "ครั้งเดียวใช้ทุก Dashboard"
 *  หน้ารวมไม่อ่านข้อมูลเอง แค่เลือก+จำโฟลเดอร์ (เก็บไว้ร่วมกันทุกหน้า)
 *  จากนั้นเปิด Dashboard ห้องไหนก็จะเชื่อมโฟลเดอร์เดิมให้อัตโนมัติ
 * ===================================================================== */
(function () {
  "use strict";

  DashboardAutoload.register({
    mode: "directory",
    storeKey: "hub",
    match: function () { return null; },   // หน้ารวมไม่อ่านไฟล์ใด ๆ
    parse: function () { return {}; },
    applyYear: function () {},
    render: function () {}
  });

  function wire() {
    var btn = document.getElementById("connectAllBtn");
    var status = document.getElementById("connectAllStatus");
    if (!btn) return;

    if (!window.showDirectoryPicker) {
      btn.disabled = true;
      if (status) status.textContent = "⚠️ ใช้ได้กับ Chrome หรือ Edge เท่านั้น";
      return;
    }
    if (DashboardAutoload.isConnected()) {
      if (status) status.textContent = "🟢 เชื่อมโฟลเดอร์แล้ว — เปิด Dashboard ห้องใดก็อัพเดทอัตโนมัติ";
    }
    btn.addEventListener("click", function () {
      if (status) status.textContent = "กำลังเลือกโฟลเดอร์ …";
      DashboardAutoload.connect();
      // เช็คสถานะหลังผู้ใช้เลือก/อนุญาต
      var tries = 0;
      var t = setInterval(function () {
        tries++;
        if (DashboardAutoload.isConnected()) {
          if (status) status.textContent = "🟢 เชื่อมแล้ว! เปิด Dashboard ห้องใดก็ได้ ระบบจะอัพเดทให้อัตโนมัติ";
          clearInterval(t);
        } else if (tries > 40) { clearInterval(t); }
      }, 500);
    });
  }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", wire);
  else wire();
})();
