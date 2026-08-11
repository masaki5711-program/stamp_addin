/**
 * tools/cleanup.js — ツール⑥: ブック掃除(Excel専用)
 * 名前定義・非表示シート・カスタムスタイルの棚卸しと削除。
 */
(function () {
  "use strict";

  var $ = function (id) { return document.getElementById(id); };

  function scan() {
    $("btn-clean-scan").disabled = true;
    var styleOk = Office.context.requirements.isSetSupported("ExcelApi", "1.7");
    Excel.run(function (ctx) {
      var names = ctx.workbook.names;
      names.load("items/name,visible");
      var sheets = ctx.workbook.worksheets;
      sheets.load("items/name,visibility");
      var styles = null;
      if (styleOk) {
        styles = ctx.workbook.styles;
        styles.load("items/name,builtIn");
      }
      return ctx.sync().then(function () {
        renderNames(names.items.map(function (n) { return { name: n.name, visible: n.visible }; }));
        renderSheets(sheets.items
          .filter(function (s) { return s.visibility !== "Visible"; })
          .map(function (s) { return s.name; }));
        renderStyles(styleOk ? styles.items.filter(function (st) { return !st.builtIn; }).map(function (st) { return st.name; }) : null);
        UI.status("スキャンしました");
      });
    }).catch(function (e) {
      UI.status("スキャンに失敗しました: " + e.message, true);
    }).finally(function () {
      $("btn-clean-scan").disabled = false;
    });
  }

  function renderNames(items) {
    var host = $("clean-names");
    host.innerHTML = "";
    $("clean-names-count").textContent = items.length + "件";
    if (items.length === 0) { host.innerHTML = '<p class="empty">名前定義はありません。</p>'; return; }
    items.forEach(function (n) {
      var row = document.createElement("div");
      row.className = "clean-row";
      row.innerHTML = '<span class="clean-label"></span><button type="button" class="danger">削除</button>';
      row.querySelector(".clean-label").textContent = n.name + (n.visible ? "" : "(非表示)");
      row.querySelector("button").addEventListener("click", function () {
        Excel.run(function (ctx) {
          ctx.workbook.names.getItem(n.name).delete();
          return ctx.sync();
        }).then(function () { row.remove(); UI.status("名前定義「" + n.name + "」を削除しました"); })
          .catch(function (e) { UI.status("削除に失敗: " + e.message, true); });
      });
      host.appendChild(row);
    });
  }

  function deleteAllNames() {
    if (!confirm("すべての名前定義を削除しますか?\n数式が名前を参照している場合、その数式はエラーになります。")) return;
    Excel.run(function (ctx) {
      var names = ctx.workbook.names;
      names.load("items/name");
      return ctx.sync().then(function () {
        names.items.forEach(function (n) { n.delete(); });
        return ctx.sync().then(function () {
          UI.status(names.items.length + "件の名前定義を削除しました");
          scan();
        });
      });
    }).catch(function (e) { UI.status("削除に失敗: " + e.message, true); });
  }

  function renderSheets(hiddenNames) {
    var host = $("clean-sheets");
    host.innerHTML = "";
    $("clean-sheets-count").textContent = hiddenNames.length + "件";
    if (hiddenNames.length === 0) { host.innerHTML = '<p class="empty">非表示シートはありません。</p>'; return; }
    hiddenNames.forEach(function (name) {
      var row = document.createElement("div");
      row.className = "clean-row";
      row.innerHTML = '<span class="clean-label"></span><button type="button">再表示</button>';
      row.querySelector(".clean-label").textContent = name;
      row.querySelector("button").addEventListener("click", function () {
        Excel.run(function (ctx) {
          ctx.workbook.worksheets.getItem(name).visibility = "Visible";
          return ctx.sync();
        }).then(function () { row.remove(); UI.status("シート「" + name + "」を再表示しました"); })
          .catch(function (e) { UI.status("再表示に失敗: " + e.message, true); });
      });
      host.appendChild(row);
    });
  }

  function renderStyles(customNames) {
    var host = $("clean-styles");
    if (customNames === null) {
      host.innerHTML = '<p class="empty">この環境ではスタイル操作に対応していません。</p>';
      $("btn-clean-styles").hidden = true;
      return;
    }
    $("clean-styles-count").textContent = customNames.length + "件";
    host.innerHTML = customNames.length === 0
      ? '<p class="empty">カスタムスタイルはありません。</p>'
      : '<p class="hint">' + customNames.slice(0, 10).join("、") + (customNames.length > 10 ? " ほか" : "") + "</p>";
    $("btn-clean-styles").hidden = customNames.length === 0;
  }

  function deleteCustomStyles() {
    if (!confirm("カスタムスタイルをすべて削除しますか?\n該当スタイルを使っているセルは標準スタイルに戻ります。")) return;
    Excel.run(function (ctx) {
      var styles = ctx.workbook.styles;
      styles.load("items/name,builtIn");
      return ctx.sync().then(function () {
        var customs = styles.items.filter(function (st) { return !st.builtIn; });
        customs.forEach(function (st) { st.delete(); });
        return ctx.sync().then(function () {
          UI.status(customs.length + "件のカスタムスタイルを削除しました");
          scan();
        });
      });
    }).catch(function (e) { UI.status("削除に失敗: " + e.message, true); });
  }

  UI.registerTool({
    id: "cleanup",
    name: "ブック掃除",
    desc: "名前定義・非表示シート・不要スタイルを棚卸し",
    hosts: ["excel"],
    unavailableNote: "Excel専用のツールです",
    init: function () {
      $("btn-clean-scan").addEventListener("click", scan);
      $("btn-clean-names").addEventListener("click", deleteAllNames);
      $("btn-clean-styles").addEventListener("click", deleteCustomStyles);
    }
  });
})();
