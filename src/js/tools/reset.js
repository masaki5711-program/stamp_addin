/**
 * tools/reset.js — ツール②: お作法リセット(納品前の体裁リセット・Excel専用)
 * 全シートの選択をA1に戻し、必要ならウィンドウ枠固定解除・先頭シート表示・保存を行う。
 */
(function () {
  "use strict";

  var $ = function (id) { return document.getElementById(id); };

  function run() {
    var optUnfreeze = $("chk-reset-unfreeze").checked;
    var optFirst = $("chk-reset-first").checked;
    var optSave = $("chk-reset-save").checked;

    $("btn-reset-run").disabled = true;
    Excel.run(function (ctx) {
      var sheets = ctx.workbook.worksheets;
      sheets.load("items/name,visibility,position");
      var active = ctx.workbook.worksheets.getActiveWorksheet();
      active.load("name");
      return ctx.sync().then(function () {
        var visible = sheets.items
          .filter(function (s) { return s.visibility === "Visible"; })
          .sort(function (a, b) { return a.position - b.position; });
        visible.forEach(function (s) {
          if (optUnfreeze) s.freezePanes.unfreeze();
          s.getRange("A1").select();
        });
        // 最後に表示するシート: 先頭 or 元のシート
        var last = optFirst ? visible[0] : visible.find(function (s) { return s.name === active.name; }) || visible[0];
        if (last) last.getRange("A1").select();
        return ctx.sync().then(function () {
          if (optSave && Office.context.requirements.isSetSupported("ExcelApi", "1.11")) {
            ctx.workbook.save(Excel.SaveBehavior.save);
            return ctx.sync();
          }
        }).then(function () {
          UI.status(visible.length + "枚のシートをA1に整えました" + (optSave ? "(保存済み)" : ""));
        });
      });
    }).catch(function (e) {
      UI.status("リセットに失敗しました: " + e.message, true);
    }).finally(function () {
      $("btn-reset-run").disabled = false;
    });
  }

  UI.registerTool({
    id: "reset",
    name: "お作法リセット",
    desc: "全シートをA1選択に整えて納品前の体裁に",
    hosts: ["excel"],
    unavailableNote: "Excel専用のツールです",
    init: function () {
      $("btn-reset-run").addEventListener("click", run);
    }
  });
})();
