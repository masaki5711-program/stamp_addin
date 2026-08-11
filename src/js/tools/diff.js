/**
 * tools/diff.js — ツール③: シート差分比較(Excel専用)
 * 2シートの値をセル単位で比較し、一覧表示+シートB側をハイライトする。
 */
(function () {
  "use strict";

  var $ = function (id) { return document.getElementById(id); };
  var MAX_CELLS = 50000;   // 比較セル数の上限
  var MAX_LIST = 200;      // 一覧表示の上限
  var MAX_HIGHLIGHT = 500; // ハイライトの上限
  var HIGHLIGHT_COLOR = "#FFE79E";
  var lastHighlight = null; // { sheetName, cells: [{r,c}] }

  function colLetter(c) {
    var s = "";
    c += 1;
    while (c > 0) { var m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = (c - m - 1) / 26; }
    return s;
  }

  function addr(r, c) { return colLetter(c) + (r + 1); }

  /**
   * 2つの usedRange 情報 {values, rowIndex, columnIndex} を絶対座標で比較する。
   * 戻り値: [{r, c, a, b}](絶対0始まり)
   */
  function diffGrids(A, B) {
    function bounds(g) {
      var rows = g.values.length, cols = rows ? g.values[0].length : 0;
      return { r0: g.rowIndex, c0: g.columnIndex, r1: g.rowIndex + rows, c1: g.columnIndex + cols };
    }
    function val(g, r, c) {
      var rr = r - g.rowIndex, cc = c - g.columnIndex;
      if (rr < 0 || cc < 0 || rr >= g.values.length || cc >= (g.values[0] || []).length) return "";
      var v = g.values[rr][cc];
      return v === null || v === undefined ? "" : v;
    }
    var ba = bounds(A), bb = bounds(B);
    var r0 = Math.min(ba.r0, bb.r0), c0 = Math.min(ba.c0, bb.c0);
    var r1 = Math.max(ba.r1, bb.r1), c1 = Math.max(ba.c1, bb.c1);
    if ((r1 - r0) * (c1 - c0) > MAX_CELLS) {
      throw new Error("比較範囲が大きすぎます(" + (r1 - r0) + "行×" + (c1 - c0) + "列)。上限は" + MAX_CELLS + "セルです");
    }
    var out = [];
    for (var r = r0; r < r1; r++) {
      for (var c = c0; c < c1; c++) {
        var a = val(A, r, c), b = val(B, r, c);
        if (String(a) !== String(b)) out.push({ r: r, c: c, a: a, b: b });
      }
    }
    return out;
  }

  function loadSheetNames() {
    return Excel.run(function (ctx) {
      var sheets = ctx.workbook.worksheets;
      sheets.load("items/name,visibility");
      return ctx.sync().then(function () {
        return sheets.items
          .filter(function (s) { return s.visibility === "Visible"; })
          .map(function (s) { return s.name; });
      });
    });
  }

  function refreshSheets() {
    loadSheetNames().then(function (names) {
      ["diff-a", "diff-b"].forEach(function (id, idx) {
        var sel = $(id);
        var prev = sel.value;
        sel.innerHTML = "";
        names.forEach(function (n) {
          var o = document.createElement("option");
          o.value = n; o.textContent = n;
          sel.appendChild(o);
        });
        sel.value = names.indexOf(prev) >= 0 ? prev : (names[Math.min(idx, names.length - 1)] || "");
      });
    }).catch(function (e) { UI.status("シート一覧の取得に失敗: " + e.message, true); });
  }

  function usedRangeOf(ctx, name) {
    var rng = ctx.workbook.worksheets.getItem(name).getUsedRangeOrNullObject(true);
    rng.load("values,rowIndex,columnIndex,isNullObject");
    return rng;
  }

  function compare() {
    var nameA = $("diff-a").value, nameB = $("diff-b").value;
    if (!nameA || !nameB) { UI.status("シートを選択してください", true); return; }
    if (nameA === nameB) { UI.status("異なるシートを選択してください", true); return; }

    $("btn-diff-run").disabled = true;
    Excel.run(function (ctx) {
      var ra = usedRangeOf(ctx, nameA);
      var rb = usedRangeOf(ctx, nameB);
      return ctx.sync().then(function () {
        var A = ra.isNullObject ? { values: [], rowIndex: 0, columnIndex: 0 } : ra;
        var B = rb.isNullObject ? { values: [], rowIndex: 0, columnIndex: 0 } : rb;
        var diffs = diffGrids(A, B);
        renderResult(nameA, nameB, diffs);
        if ($("chk-diff-highlight").checked && diffs.length > 0) {
          var sheetB = ctx.workbook.worksheets.getItem(nameB);
          var targets = diffs.slice(0, MAX_HIGHLIGHT);
          targets.forEach(function (d) {
            sheetB.getCell(d.r, d.c).format.fill.color = HIGHLIGHT_COLOR;
          });
          lastHighlight = { sheetName: nameB, cells: targets.map(function (d) { return { r: d.r, c: d.c }; }) };
          $("btn-diff-clear").hidden = false;
          return ctx.sync();
        }
      });
    }).catch(function (e) {
      UI.status("比較に失敗しました: " + e.message, true);
    }).finally(function () {
      $("btn-diff-run").disabled = false;
    });
  }

  function renderResult(nameA, nameB, diffs) {
    var sum = $("diff-summary");
    var list = $("diff-list");
    list.innerHTML = "";
    if (diffs.length === 0) {
      sum.textContent = "差分はありません(「" + nameA + "」と「" + nameB + "」は値が一致)";
      return;
    }
    sum.textContent = "差分 " + diffs.length + "件" +
      (diffs.length > MAX_LIST ? "(先頭" + MAX_LIST + "件を表示)" : "") +
      " — クリックでそのセルへ移動";
    diffs.slice(0, MAX_LIST).forEach(function (d) {
      var li = document.createElement("button");
      li.type = "button";
      li.className = "diff-row";
      li.innerHTML = '<span class="diff-addr"></span><span class="diff-vals"></span>';
      li.querySelector(".diff-addr").textContent = addr(d.r, d.c);
      li.querySelector(".diff-vals").textContent =
        (d.a === "" ? "(空)" : d.a) + " → " + (d.b === "" ? "(空)" : d.b);
      li.addEventListener("click", function () {
        Excel.run(function (ctx) {
          ctx.workbook.worksheets.getItem(nameB).getCell(d.r, d.c).select();
          return ctx.sync();
        }).catch(function () {});
      });
      list.appendChild(li);
    });
  }

  function clearHighlight() {
    if (!lastHighlight) return;
    var h = lastHighlight;
    Excel.run(function (ctx) {
      var sheet = ctx.workbook.worksheets.getItem(h.sheetName);
      h.cells.forEach(function (cell) {
        sheet.getCell(cell.r, cell.c).format.fill.clear();
      });
      return ctx.sync();
    }).then(function () {
      lastHighlight = null;
      $("btn-diff-clear").hidden = true;
      UI.status("色付けを解除しました");
    }).catch(function (e) { UI.status("解除に失敗: " + e.message, true); });
  }

  UI.registerTool({
    id: "diff",
    name: "シート差分比較",
    desc: "2つのシートの値の違いを色付け・一覧化",
    hosts: ["excel"],
    unavailableNote: "Excel専用のツールです",
    init: function () {
      $("btn-diff-refresh").addEventListener("click", refreshSheets);
      $("btn-diff-run").addEventListener("click", compare);
      $("btn-diff-clear").addEventListener("click", clearHighlight);
      if (UI.host() === "excel") refreshSheets();
    },
    _test: { diffGrids: diffGrids, colLetter: colLetter }
  });

  // nodeテスト用
  if (typeof module !== "undefined") window.ToolDiffTest = { diffGrids: diffGrids, colLetter: colLetter };
})();
