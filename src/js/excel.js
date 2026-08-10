/**
 * excel.js — Excelへの押印(画像挿入)。Excel外ではPNGダウンロードにフォールバック。
 */
(function (global) {
  "use strict";

  var inExcel = false;

  function init() {
    return Office.onReady().then(function (info) {
      inExcel = !!(info && info.host === Office.HostType.Excel);
      return inExcel;
    });
  }

  function mmToPt(mm) { return mm * 72 / 25.4; }

  /**
   * 選択セルの中央に印影PNGを挿入し、監査メタデータを記録する。
   * meta: { stampId, name, date, stampedAt }
   */
  function insertStamp(pngDataUrl, diameterMm, meta) {
    var base64 = pngDataUrl.split(",")[1];
    var sizePt = mmToPt(diameterMm);
    return Excel.run(function (ctx) {
      var sheet = ctx.workbook.worksheets.getActiveWorksheet();
      var range = ctx.workbook.getSelectedRange();
      range.load(["left", "top", "width", "height"]);
      return ctx.sync().then(function () {
        var shape = sheet.shapes.addImage(base64);
        shape.left = range.left + Math.max(0, (range.width - sizePt) / 2);
        shape.top = range.top + Math.max(0, (range.height - sizePt) / 2);
        shape.width = sizePt;
        shape.height = sizePt;
        shape.name = "stamp_" + meta.stampId + "_" + meta.date;
        shape.altTextTitle = "デート印: " + meta.name;
        shape.altTextDescription = JSON.stringify(meta);
        return ctx.sync();
      });
    });
  }

  function downloadPNG(pngDataUrl, filename) {
    var a = document.createElement("a");
    a.href = pngDataUrl;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  }

  global.XL = {
    init: init,
    isExcel: function () { return inExcel; },
    insertStamp: insertStamp,
    downloadPNG: downloadPNG
  };
})(window);
