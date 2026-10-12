/**
 * excel.js — Excel/Wordへの押印(画像挿入)。Office外ではPNGダウンロードにフォールバック。
 *  - Excel: 選択セルの中央に図形(画像)として挿入
 *  - Word : カーソル位置にインライン画像として挿入
 */
(function (global) {
  "use strict";

  var hostType = null; // "excel" | "word" | null
  var MAX_STAMPS_PER_RUN = 100;

  function init() {
    return Office.onReady().then(function (info) {
      if (info && info.host === Office.HostType.Excel) hostType = "excel";
      else if (info && info.host === Office.HostType.Word) hostType = "word";
      else hostType = null;
      return hostType;
    });
  }

  function mmToPt(mm) { return mm * 72 / 25.4; }

  function supports(set, version) {
    try { return Office.context.requirements.isSetSupported(set, version); } catch (e) { return false; }
  }

  // 開いている文書のファイル名(未保存などで取れないときは空文字)
  function docName() {
    try {
      var url = Office.context.document.url || "";
      return decodeURIComponent(url.split(/[\\\/]/).pop().split("?")[0]);
    } catch (e) { return ""; }
  }

  /**
   * 選択範囲に押印する。perCell=true なら選択した各セルに1つずつ、false なら選択範囲(先頭の領域)の中央に1つ。
   * 戻り値: 押した場所の配列(例 ["Sheet1!B3"])
   */
  function insertStampExcel(base64, wPt, hPt, meta, perCell) {
    return Excel.run(function (ctx) {
      var sheet = ctx.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
      var areas = ctx.workbook.getSelectedRanges().areas;
      areas.load("left,top,width,height,address,rowCount,columnCount");
      return ctx.sync().then(function () {
        var targets = [];
        if (perCell) {
          var total = areas.items.reduce(function (n, a) { return n + a.rowCount * a.columnCount; }, 0);
          if (total > MAX_STAMPS_PER_RUN) {
            throw new Error("選択セルが多すぎます(" + total + "セル)。一度に押せるのは" + MAX_STAMPS_PER_RUN + "セルまでです");
          }
          areas.items.forEach(function (a) {
            for (var r = 0; r < a.rowCount; r++) {
              for (var c = 0; c < a.columnCount; c++) {
                var cell = a.getCell(r, c);
                cell.load("left,top,width,height,address");
                targets.push(cell);
              }
            }
          });
        } else {
          targets.push(areas.items[0]);
        }
        return ctx.sync().then(function () {
          var movable = supports("ExcelApi", "1.10");
          targets.forEach(function (t, i) {
            var shape = sheet.shapes.addImage(base64);
            shape.lockAspectRatio = false;
            shape.width = wPt;
            shape.height = hPt;
            shape.left = t.left + Math.max(0, (t.width - wPt) / 2);
            shape.top = t.top + Math.max(0, (t.height - hPt) / 2);
            shape.lockAspectRatio = true;
            // セルと一緒に移動するが、列幅・行高の変更では伸び縮みしない
            if (movable) shape.placement = "OneCell";
            shape.name = "stamp_" + meta.stampId + "_" + meta.date + (i ? "_" + (i + 1) : "");
            shape.altTextTitle = "デート印: " + meta.name;
            shape.altTextDescription = JSON.stringify(meta);
          });
          return ctx.sync().then(function () {
            return targets.map(function (t) { return t.address; });
          });
        });
      });
    });
  }

  function insertStampWord(base64, wPt, hPt, meta) {
    return Word.run(function (ctx) {
      var sel = ctx.document.getSelection();
      var pic = sel.insertInlinePictureFromBase64(base64, Word.InsertLocation.replace);
      pic.lockAspectRatio = false;
      pic.width = wPt;
      pic.height = hPt;
      pic.altTextTitle = "デート印: " + meta.name;
      pic.altTextDescription = JSON.stringify(meta);
      return ctx.sync().then(function () { return ["本文"]; });
    });
  }

  /**
   * 押印する。size: {w, h}(mm)、meta: { stampId, name, date, stampedAt, vcode }
   * opts.perCell — Excelで選択した各セルに押す
   * 戻り値(Promise): 押した場所の配列
   */
  function insertStamp(pngDataUrl, size, meta, opts) {
    var base64 = pngDataUrl.split(",")[1];
    var wPt = mmToPt(size.w), hPt = mmToPt(size.h);
    if (hostType === "excel") return insertStampExcel(base64, wPt, hPt, meta, !!(opts && opts.perCell));
    if (hostType === "word") return insertStampWord(base64, wPt, hPt, meta);
    return Promise.reject(new Error("Office上で実行されていません"));
  }

  /**
   * 承認欄を作る(Excel)。アクティブセルを左上として「見出し行+押印行」×人数分の枠を作り、
   * 押印行を選択した状態で返す(続けて「各セルに押印」できる)。
   * 列幅・行の高さは印が収まる大きさに変更する。
   */
  function createApprovalBox(labels, size) {
    var wPt = mmToPt(size.w), hPt = mmToPt(size.h);
    return Excel.run(function (ctx) {
      var box = ctx.workbook.getActiveCell().getResizedRange(1, labels.length - 1);
      var head = box.getRow(0), body = box.getRow(1);
      box.load("address");
      head.values = [labels];
      head.format.horizontalAlignment = "Center";
      head.format.verticalAlignment = "Center";
      head.format.font.size = 9;
      box.format.columnWidth = Math.max(wPt + 10, 48);
      body.format.rowHeight = hPt + 10;
      ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideHorizontal", "InsideVertical"].forEach(function (edge) {
        var b = box.format.borders.getItem(edge);
        b.style = "Continuous";
        b.weight = "Thin";
      });
      body.select();
      return ctx.sync().then(function () { return box.address; });
    });
  }

  /**
   * 文書内の押印メタデータ(altTextのJSON)を収集する。各要素に where(場所)を付ける。
   * Excel: 全シートの図形 / Word: 本文のインライン画像
   */
  function collectStampMetas() {
    function parse(list) {
      var out = [];
      list.forEach(function (item) {
        if (!item.text) return;
        try {
          var m = JSON.parse(item.text);
          if (m && m.stampId) { m.where = item.where; out.push(m); }
        } catch (e) { /* 押印以外の画像は無視 */ }
      });
      return out;
    }
    if (hostType === "excel") {
      return Excel.run(function (ctx) {
        var sheets = ctx.workbook.worksheets;
        sheets.load("items/name");
        return ctx.sync().then(function () {
          sheets.items.forEach(function (s) { s.shapes.load("items/altTextDescription"); });
          return ctx.sync().then(function () {
            var all = [];
            sheets.items.forEach(function (s) {
              s.shapes.items.forEach(function (sh) { all.push({ text: sh.altTextDescription, where: s.name }); });
            });
            return parse(all);
          });
        });
      });
    }
    if (hostType === "word") {
      return Word.run(function (ctx) {
        var pics = ctx.document.body.inlinePictures;
        pics.load("items/altTextDescription");
        return ctx.sync().then(function () {
          return parse(pics.items.map(function (p) { return { text: p.altTextDescription, where: "本文" }; }));
        });
      });
    }
    return Promise.reject(new Error("Office上で実行されていません"));
  }

  function downloadText(text, filename, type) {
    var blob = new Blob([text], { type: type || "application/json" });
    var a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(a.href); }, 1000);
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
    supports: supports,
    docName: docName,
    insertStamp: insertStamp,
    createApprovalBox: createApprovalBox,
    collectStampMetas: collectStampMetas,
    downloadPNG: downloadPNG,
    downloadText: downloadText
  };
})(window);
