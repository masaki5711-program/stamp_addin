/**
 * excel.js — Excel/Wordへの押印(画像挿入)。Office外ではPNGダウンロードにフォールバック。
 *  - Excel: 選択セルの中央に図形(画像)として挿入
 *  - Word : カーソル位置にインライン画像として挿入
 */
(function (global) {
  "use strict";

  var hostType = null; // "excel" | "word" | null

  function init() {
    return Office.onReady().then(function (info) {
      if (info && info.host === Office.HostType.Excel) hostType = "excel";
      else if (info && info.host === Office.HostType.Word) hostType = "word";
      else hostType = null;
      return hostType;
    });
  }

  function mmToPt(mm) { return mm * 72 / 25.4; }

  function insertStampExcel(base64, sizePt, meta) {
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

  function insertStampWord(base64, sizePt, meta) {
    return Word.run(function (ctx) {
      var sel = ctx.document.getSelection();
      var pic = sel.insertInlinePictureFromBase64(base64, Word.InsertLocation.replace);
      pic.width = sizePt;
      pic.height = sizePt;
      pic.altTextTitle = "デート印: " + meta.name;
      pic.altTextDescription = JSON.stringify(meta);
      return ctx.sync();
    });
  }

  /**
   * 押印する。meta: { stampId, name, date, stampedAt }
   */
  function insertStamp(pngDataUrl, diameterMm, meta) {
    var base64 = pngDataUrl.split(",")[1];
    var sizePt = mmToPt(diameterMm);
    if (hostType === "excel") return insertStampExcel(base64, sizePt, meta);
    if (hostType === "word") return insertStampWord(base64, sizePt, meta);
    return Promise.reject(new Error("Office上で実行されていません"));
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
    hostType: function () { return hostType; },
    isOffice: function () { return hostType !== null; },
    insertStamp: insertStamp,
    downloadPNG: downloadPNG
  };
})(window);
