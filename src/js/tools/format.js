/**
 * tools/format.js — ツール④: 定型書式(Excel専用)
 * 選択範囲に「見出し・ゼブラ・罫線・列幅」の定型書式をワンクリック適用。
 * プリセット(組込3種+マイプリセット保存)対応。
 */
(function () {
  "use strict";

  var $ = function (id) { return document.getElementById(id); };
  var KEY = "formatPresets.v1";
  var MAX_BAND_ROWS = 500;

  var BUILTINS = [
    { name: "標準(グレー)", header: true, headerFill: "#EDEAE3", headerFont: "#2A2622", band: true, bandFill: "#F7F5F0", border: "thin", autofit: true },
    { name: "ブルー", header: true, headerFill: "#DCE6F2", headerFont: "#1F3A5F", band: true, bandFill: "#F1F5FA", border: "thin", autofit: true },
    { name: "朱", header: true, headerFill: "#F7E8E4", headerFont: "#A93223", band: true, bandFill: "#FBF3F1", border: "thin", autofit: true }
  ];

  function readSettings() {
    return {
      header: $("chk-fmt-header").checked,
      headerFill: $("fmt-header-fill").value,
      headerFont: $("fmt-header-font").value,
      band: $("chk-fmt-band").checked,
      bandFill: $("fmt-band-fill").value,
      border: $("fmt-border").value,
      autofit: $("chk-fmt-autofit").checked
    };
  }

  function fillSettings(p) {
    $("chk-fmt-header").checked = !!p.header;
    $("fmt-header-fill").value = p.headerFill || "#EDEAE3";
    $("fmt-header-font").value = p.headerFont || "#2A2622";
    $("chk-fmt-band").checked = !!p.band;
    $("fmt-band-fill").value = p.bandFill || "#F7F5F0";
    $("fmt-border").value = p.border || "thin";
    $("chk-fmt-autofit").checked = !!p.autofit;
  }

  function renderPresets() {
    var host = $("fmt-presets");
    host.innerHTML = "";
    var mine = Store.readKey(KEY, []);
    BUILTINS.concat(mine).forEach(function (p, idx) {
      var isMine = idx >= BUILTINS.length;
      var chip = document.createElement("span");
      chip.className = "chip-wrap";
      var btn = document.createElement("button");
      btn.type = "button";
      btn.className = "chip";
      btn.textContent = p.name;
      btn.title = "クリックで設定を読み込み";
      btn.addEventListener("click", function () {
        fillSettings(p);
        UI.status("プリセット「" + p.name + "」を読み込みました。「選択範囲に適用」で反映します");
      });
      chip.appendChild(btn);
      if (isMine) {
        var del = document.createElement("button");
        del.type = "button";
        del.className = "chip-del";
        del.textContent = "×";
        del.title = "このプリセットを削除";
        del.addEventListener("click", function () {
          var list = Store.readKey(KEY, []).filter(function (x) { return x.name !== p.name; });
          Store.writeKey(KEY, list);
          renderPresets();
        });
        chip.appendChild(del);
      }
      host.appendChild(chip);
    });
  }

  function savePreset() {
    var name = $("fmt-preset-name").value.trim();
    if (!name) { UI.status("プリセット名を入力してください", true); return; }
    var p = readSettings();
    p.name = name;
    var list = Store.readKey(KEY, []).filter(function (x) { return x.name !== name; });
    list.push(p);
    Store.writeKey(KEY, list);
    $("fmt-preset-name").value = "";
    renderPresets();
    UI.status("プリセット「" + name + "」を保存しました");
  }

  function apply() {
    var p = readSettings();
    $("btn-fmt-apply").disabled = true;
    Excel.run(function (ctx) {
      var rng = ctx.workbook.getSelectedRange();
      rng.load("rowCount,columnCount,cellCount");
      return ctx.sync().then(function () {
        if (rng.cellCount < 2) throw new Error("表の範囲を選択してから実行してください");
        if (p.border !== "none") {
          ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideHorizontal", "InsideVertical"]
            .forEach(function (edge) {
              var b = rng.format.borders.getItem(edge);
              b.style = p.border === "medium" ? "Continuous" : "Continuous";
              b.weight = p.border === "medium" ? "Medium" : "Thin";
              b.color = "#A9A296";
            });
        }
        if (p.header) {
          var head = rng.getRow(0);
          head.format.fill.color = p.headerFill;
          head.format.font.color = p.headerFont;
          head.format.font.bold = true;
        }
        if (p.band) {
          var start = p.header ? 1 : 0;
          var rows = Math.min(rng.rowCount, MAX_BAND_ROWS);
          for (var r = start; r < rows; r++) {
            if ((r - start) % 2 === 1) rng.getRow(r).format.fill.color = p.bandFill;
          }
        }
        if (p.autofit) rng.format.autofitColumns();
        return ctx.sync().then(function () {
          UI.status("書式を適用しました(" + rng.rowCount + "行×" + rng.columnCount + "列)");
        });
      });
    }).catch(function (e) {
      UI.status("適用に失敗しました: " + e.message, true);
    }).finally(function () {
      $("btn-fmt-apply").disabled = false;
    });
  }

  UI.registerTool({
    id: "format",
    name: "定型書式",
    desc: "選択範囲を見出し・縞・罫線つきの表にワンクリック整形",
    hosts: ["excel"],
    unavailableNote: "Excel専用のツールです",
    init: function () {
      renderPresets();
      fillSettings(BUILTINS[0]);
      $("btn-fmt-apply").addEventListener("click", apply);
      $("btn-fmt-save").addEventListener("click", savePreset);
    }
  });
})();
