/**
 * tools/cleanup.js — ツール⑥: ブック掃除(Excel専用)
 * 名前定義・非表示シート・カスタムスタイルの棚卸しと削除。
 * 加えて、外部ブックへのリンク・コメント・作成者情報・書式だけが残った範囲を点検する。
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
        // シートスコープの名前定義も収集(ブック全体の名前と別枠で管理されている)
        var perSheet = sheets.items.map(function (s) { return { sheetName: s.name, names: s.names }; });
        perSheet.forEach(function (p) { p.names.load("items/name,visible"); });
        return ctx.sync().then(function () {
          var all = names.items.map(function (n) { return { name: n.name, visible: n.visible, scope: null }; });
          perSheet.forEach(function (p) {
            p.names.items.forEach(function (n) {
              all.push({ name: n.name, visible: n.visible, scope: p.sheetName });
            });
          });
          renderNames(all);
          renderSheets(sheets.items
            .filter(function (s) { return s.visibility !== "Visible"; })
            .map(function (s) { return s.name; }));
          renderStyles(styleOk ? styles.items.filter(function (st) { return !st.builtIn; }).map(function (st) { return st.name; }) : null);
          UI.status("スキャンしました");
        });
      });
    }).then(scanExtras).catch(function (e) {
      UI.status("スキャンに失敗しました: " + e.message, true);
    }).finally(function () {
      $("btn-clean-scan").disabled = false;
    });
  }

  function deleteName(item) {
    return Excel.run(function (ctx) {
      var coll = item.scope
        ? ctx.workbook.worksheets.getItem(item.scope).names
        : ctx.workbook.names;
      coll.getItem(item.name).delete();
      return ctx.sync();
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
      row.querySelector(".clean-label").textContent =
        n.name + (n.scope ? "(シート: " + n.scope + ")" : "") + (n.visible ? "" : "(非表示)");
      row.querySelector("button").addEventListener("click", function () {
        deleteName(n)
          .then(function () { row.remove(); UI.status("名前定義「" + n.name + "」を削除しました"); })
          .catch(function (e) { UI.status("削除に失敗: " + e.message, true); });
      });
      host.appendChild(row);
    });
  }

  function deleteAllNames() {
    // 数式が名前を参照している場合、その数式はエラーになる
    if (!UI.confirmClick($("btn-clean-names"))) return;
    Excel.run(function (ctx) {
      var names = ctx.workbook.names;
      names.load("items/name");
      var sheets = ctx.workbook.worksheets;
      sheets.load("items/name");
      return ctx.sync().then(function () {
        var perSheet = sheets.items.map(function (s) { return s.names; });
        perSheet.forEach(function (nc) { nc.load("items/name"); });
        return ctx.sync().then(function () {
          var count = 0;
          names.items.forEach(function (n) { n.delete(); count++; });
          perSheet.forEach(function (nc) {
            nc.items.forEach(function (n) { n.delete(); count++; });
          });
          return ctx.sync().then(function () {
            UI.status(count + "件の名前定義を削除しました");
            scan();
          });
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
    // 該当スタイルを使っているセルは標準スタイルに戻る
    if (!UI.confirmClick($("btn-clean-styles"))) return;
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

  /* ---- 追加の棚卸し: 外部リンク・コメント・作成者情報・書式だけの範囲 ----
   * 本体のスキャンとは別の Excel.run で行い、どれかが失敗しても他の結果は表示する。
   */

  var MAX_SCAN_CELLS = 200000; // 1シートあたり数式を読む上限
  var MAX_LINKS = 50;
  // 外部ブック参照([Book.xlsx]Sheet!A1 の形)。テーブルの構造化参照 Table1[列] は "!" が続かないので除外される
  var EXTERNAL_REF = /\[[^\[\]]+\][^!\[\]']*'?!/;
  var PROP_LABELS = {
    author: "作成者", lastAuthor: "最終更新者", company: "会社", manager: "管理者",
    title: "タイトル", subject: "件名", keywords: "キーワード", category: "分類", comments: "コメント欄"
  };
  var CLEARABLE_PROPS = ["author", "company", "manager", "title", "subject", "keywords", "category", "comments"];

  function colLetter(c) {
    var s = "";
    c += 1;
    while (c > 0) { var m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = (c - m - 1) / 26; }
    return s;
  }

  // formulas(2次元配列)から外部ブック参照を含むセルを拾う
  function findExternalRefs(sheetName, formulas, rowIndex, columnIndex, out) {
    for (var r = 0; r < formulas.length; r++) {
      for (var c = 0; c < formulas[r].length; c++) {
        var f = formulas[r][c];
        if (typeof f === "string" && f.charAt(0) === "=" && EXTERNAL_REF.test(f)) {
          out.push({ sheet: sheetName, r: rowIndex + r, c: columnIndex + c, formula: f });
        }
      }
    }
  }

  function scanExtras() {
    var commentsOk = XL.supports("ExcelApi", "1.10");
    var notesOk = XL.supports("ExcelApi", "1.18");
    var propsOk = XL.supports("ExcelApi", "1.7");
    return Excel.run(function (ctx) {
      var sheets = ctx.workbook.worksheets;
      sheets.load("items/name");
      var commentCount = commentsOk ? ctx.workbook.comments.getCount() : null;
      var noteCount = notesOk ? ctx.workbook.notes.getCount() : null;
      var props = null;
      if (propsOk) {
        props = ctx.workbook.properties;
        props.load(Object.keys(PROP_LABELS).join(","));
      }
      return ctx.sync().then(function () {
        var per = sheets.items.map(function (s) {
          var all = s.getUsedRangeOrNullObject(false);
          var vals = s.getUsedRangeOrNullObject(true);
          all.load("isNullObject,rowCount,columnCount,address");
          vals.load("isNullObject,rowCount,columnCount,rowIndex,columnIndex,address");
          return { name: s.name, all: all, vals: vals };
        });
        return ctx.sync().then(function () {
          var bloat = [], skipped = [];
          per.forEach(function (p) {
            var allCells = p.all.isNullObject ? 0 : p.all.rowCount * p.all.columnCount;
            var valCells = p.vals.isNullObject ? 0 : p.vals.rowCount * p.vals.columnCount;
            // 値のある範囲より大幅に広い = 書式だけが遠くまで残っている(ファイル肥大・スクロールバー縮小の原因)
            if (allCells - valCells > 10000 && allCells > valCells * 4) {
              bloat.push({
                name: p.name,
                all: p.all.address.split("!").pop(),
                vals: p.vals.isNullObject ? "(値なし)" : p.vals.address.split("!").pop()
              });
            }
            if (valCells > MAX_SCAN_CELLS) { skipped.push(p.name); p.formulas = false; }
            else if (valCells > 0) { p.vals.load("formulas"); p.formulas = true; }
          });
          return ctx.sync().then(function () {
            var links = [];
            per.forEach(function (p) {
              if (p.formulas) findExternalRefs(p.name, p.vals.formulas, p.vals.rowIndex, p.vals.columnIndex, links);
            });
            renderLinks(links, skipped);
            renderComments(commentCount ? commentCount.value : null, noteCount ? noteCount.value : null);
            renderProps(props);
            renderBloat(bloat);
          });
        });
      });
    }).catch(function (e) {
      UI.status("追加項目のスキャンに失敗しました: " + e.message, true);
    });
  }

  function renderLinks(links, skipped) {
    var host = $("clean-links");
    host.innerHTML = "";
    $("clean-links-count").textContent = links.length + "件";
    if (links.length === 0) host.innerHTML = '<p class="empty">外部ブックを参照する数式はありません。</p>';
    links.slice(0, MAX_LINKS).forEach(function (l) {
      var row = document.createElement("button");
      row.type = "button";
      row.className = "diff-row";
      row.innerHTML = '<span class="diff-addr"></span><span class="diff-vals"></span>';
      row.querySelector(".diff-addr").textContent = l.sheet + "!" + colLetter(l.c) + (l.r + 1);
      row.querySelector(".diff-vals").textContent = l.formula;
      row.addEventListener("click", function () {
        Excel.run(function (ctx) {
          var sheet = ctx.workbook.worksheets.getItem(l.sheet);
          sheet.activate();
          sheet.getCell(l.r, l.c).select();
          return ctx.sync();
        }).catch(function () {});
      });
      host.appendChild(row);
    });
    var notes = [];
    if (links.length > MAX_LINKS) notes.push("先頭" + MAX_LINKS + "件を表示しています。");
    if (skipped.length) notes.push("大きすぎるため未確認のシート: " + skipped.join("、"));
    if (notes.length) {
      var p = document.createElement("p");
      p.className = "hint";
      p.textContent = notes.join(" ");
      host.appendChild(p);
    }
  }

  function renderComments(comments, notes) {
    var host = $("clean-comments");
    if (comments === null) {
      host.innerHTML = '<p class="empty">この環境ではコメントの確認に対応していません。</p>';
      $("btn-clean-comments").hidden = true;
      return;
    }
    var total = comments + (notes || 0);
    $("clean-comments-count").textContent = total + "件";
    var p = document.createElement("p");
    p.className = total ? "hint" : "empty";
    p.textContent = total === 0
      ? "コメントはありません。"
      : "コメント " + comments + "件" + (notes === null ? "(メモはこの環境では確認できません)" : " / メモ " + notes + "件");
    host.innerHTML = "";
    host.appendChild(p);
    $("btn-clean-comments").hidden = total === 0;
  }

  function deleteAllComments() {
    if (!UI.confirmClick($("btn-clean-comments"))) return;
    var notesOk = XL.supports("ExcelApi", "1.18");
    Excel.run(function (ctx) {
      var comments = ctx.workbook.comments;
      comments.load("items");
      var notes = notesOk ? ctx.workbook.notes : null;
      if (notes) notes.load("items");
      return ctx.sync().then(function () {
        var n = comments.items.length + (notes ? notes.items.length : 0);
        comments.items.forEach(function (c) { c.delete(); });
        if (notes) notes.items.forEach(function (c) { c.delete(); });
        return ctx.sync().then(function () { return n; });
      });
    }).then(function (n) {
      UI.status(n + "件のコメント・メモを削除しました");
      scanExtras();
    }).catch(function (e) { UI.status("削除に失敗: " + e.message, true); });
  }

  function renderProps(props) {
    var host = $("clean-props");
    host.innerHTML = "";
    if (!props) {
      host.innerHTML = '<p class="empty">この環境ではプロパティの確認に対応していません。</p>';
      $("btn-clean-props").hidden = true;
      return;
    }
    var clearable = 0;
    Object.keys(PROP_LABELS).forEach(function (k) {
      var v = props[k];
      if (!v) return;
      if (CLEARABLE_PROPS.indexOf(k) >= 0) clearable++;
      var row = document.createElement("div");
      row.className = "clean-row";
      row.innerHTML = '<span class="clean-label"><span class="dt-label"></span><span class="dt-text"></span></span>';
      row.querySelector(".dt-label").textContent = PROP_LABELS[k];
      row.querySelector(".dt-text").textContent = v;
      host.appendChild(row);
    });
    if (!host.firstChild) host.innerHTML = '<p class="empty">作成者などの情報は入っていません。</p>';
    $("btn-clean-props").hidden = clearable === 0;
  }

  function clearProps() {
    if (!UI.confirmClick($("btn-clean-props"))) return;
    Excel.run(function (ctx) {
      var props = ctx.workbook.properties;
      CLEARABLE_PROPS.forEach(function (k) { props[k] = ""; });
      return ctx.sync();
    }).then(function () {
      UI.status("作成者などの情報を消去しました(最終更新者は保存時にExcelが書き込みます)");
      scanExtras();
    }).catch(function (e) { UI.status("消去に失敗: " + e.message, true); });
  }

  function renderBloat(items) {
    var host = $("clean-bloat");
    host.innerHTML = "";
    $("clean-bloat-count").textContent = items.length + "件";
    if (items.length === 0) { host.innerHTML = '<p class="empty">書式だけが遠くまで残っているシートはありません。</p>'; return; }
    items.forEach(function (b) {
      var row = document.createElement("div");
      row.className = "clean-row";
      row.innerHTML = '<span class="clean-label"><span class="dt-label"></span><span class="dt-text"></span></span>';
      row.querySelector(".dt-label").textContent = b.name;
      row.querySelector(".dt-text").textContent = "使用範囲 " + b.all + " / 値があるのは " + b.vals;
      host.appendChild(row);
    });
  }

  UI.registerTool({
    id: "cleanup",
    name: "ブック掃除",
    desc: "名前定義・非表示シート・不要スタイルを棚卸し",
    hosts: ["excel"],
    unavailableNote: "Excel専用のツールです",
    _test: { findExternalRefs: findExternalRefs },
    init: function () {
      $("btn-clean-scan").addEventListener("click", scan);
      $("btn-clean-names").addEventListener("click", deleteAllNames);
      $("btn-clean-styles").addEventListener("click", deleteCustomStyles);
      $("btn-clean-comments").addEventListener("click", deleteAllComments);
      $("btn-clean-props").addEventListener("click", clearProps);
    }
  });
})();
