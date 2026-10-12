/**
 * tools/diff.js — ツール③: シート差分比較(Excel専用)
 * 2シートをセル単位で比較し、一覧表示+シートB側をハイライトする。
 *  - 比較する内容: 表示上の文字列 / 数式
 *  - 行の挿入・削除を考慮した対応付け(行単位のLCS)
 *  - ハイライトは元の塗りつぶし色ごと記録し、消すときに元へ戻す。再比較時は前回分を自動で消す
 *  - 結果の新規シート出力、別ブックのシート取り込み
 */
(function () {
  "use strict";

  var $ = function (id) { return document.getElementById(id); };
  var MAX_CELLS = 50000;      // 比較セル数の上限
  var MAX_LIST = 200;         // 一覧表示の上限
  var MAX_HIGHLIGHT = 500;    // ハイライトの上限
  var MAX_EXPORT = 5000;      // シート出力の上限
  var MAX_ALIGN = 4000000;    // 行の対応付けを行う上限(A行数×B行数)
  var MAX_IMPORT_BYTES = 5 * 1024 * 1024;
  var HIGHLIGHT_COLOR = "#FFE79E";
  var HL_KEY = "diffHighlight.v1";

  // { docKey, sheetName, cells: [{r,c,p}], on: bool }  p = 色付け前の塗りつぶし色(なしは "")
  var hl = null;
  // 直近の比較結果(シート出力用)
  var last = null;

  function docKey() {
    try {
      return (Office.context && Office.context.document && Office.context.document.url) || "doc";
    } catch (e) { return "doc"; }
  }

  function saveHl() { Store.writeKey(HL_KEY, hl); }

  function loadHl() {
    var s = Store.readKey(HL_KEY, null);
    if (s && s.docKey === docKey() && s.cells && s.cells.length) hl = s;
  }

  function colLetter(c) {
    var s = "";
    c += 1;
    while (c > 0) { var m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = (c - m - 1) / 26; }
    return s;
  }

  function addr(r, c) { return colLetter(c) + (r + 1); }

  /* ---- 比較(純粋関数) ---- */

  function bounds(g) {
    var rows = g.values.length, cols = rows ? g.values[0].length : 0;
    return { r0: g.rowIndex, c0: g.columnIndex, r1: g.rowIndex + rows, c1: g.columnIndex + cols };
  }

  function val(g, r, c) {
    var rr = r - g.rowIndex, cc = c - g.columnIndex;
    if (rr < 0 || cc < 0 || rr >= g.values.length || cc >= (g.values[0] || []).length) return "";
    var v = g.values[rr][cc];
    return v === null || v === undefined ? "" : String(v);
  }

  function span(A, B) {
    var ba = bounds(A), bb = bounds(B);
    var s = {
      r0: Math.min(ba.r0, bb.r0), c0: Math.min(ba.c0, bb.c0),
      r1: Math.max(ba.r1, bb.r1), c1: Math.max(ba.c1, bb.c1)
    };
    if ((s.r1 - s.r0) * (s.c1 - s.c0) > MAX_CELLS) {
      throw new Error("比較範囲が大きすぎます(" + (s.r1 - s.r0) + "行×" + (s.c1 - s.c0) + "列)。上限は" + MAX_CELLS + "セルです");
    }
    return s;
  }

  /**
   * 2つのグリッド {values, rowIndex, columnIndex} を同じ位置どうしで比較する。
   * 戻り値: 差分の配列 [{kind:"cell", rA, rB, c, a, b}](絶対0始まり)
   */
  function diffGrids(A, B) {
    var s = span(A, B);
    var out = [];
    for (var r = s.r0; r < s.r1; r++) {
      for (var c = s.c0; c < s.c1; c++) {
        var a = val(A, r, c), b = val(B, r, c);
        if (a !== b) out.push({ kind: "cell", rA: r, rB: r, c: c, a: a, b: b });
      }
    }
    return out;
  }

  // 文字列配列どうしの最長共通部分列。対応する添字の組 [[i, j], ...] を返す
  function lcsPairs(a, b) {
    var pre = 0;
    while (pre < a.length && pre < b.length && a[pre] === b[pre]) pre++;
    var suf = 0;
    while (suf < a.length - pre && suf < b.length - pre && a[a.length - 1 - suf] === b[b.length - 1 - suf]) suf++;
    var n = a.length - pre - suf, m = b.length - pre - suf;
    var pairs = [], i, j;
    for (i = 0; i < pre; i++) pairs.push([i, i]);
    if (n > 0 && m > 0) {
      var w = m + 1;
      var dp = new Uint32Array((n + 1) * w);
      for (i = n - 1; i >= 0; i--) {
        for (j = m - 1; j >= 0; j--) {
          dp[i * w + j] = a[pre + i] === b[pre + j]
            ? dp[(i + 1) * w + j + 1] + 1
            : Math.max(dp[(i + 1) * w + j], dp[i * w + j + 1]);
        }
      }
      i = 0; j = 0;
      while (i < n && j < m) {
        if (a[pre + i] === b[pre + j]) { pairs.push([pre + i, pre + j]); i++; j++; }
        else if (dp[(i + 1) * w + j] >= dp[i * w + j + 1]) i++;
        else j++;
      }
    }
    for (i = 0; i < suf; i++) pairs.push([a.length - suf + i, b.length - suf + i]);
    return pairs;
  }

  /**
   * 行の挿入・削除を考慮して比較する。内容が完全に一致する行どうしを対応付け、
   * その間に残った行は順に「変更された行」として突き合わせ、余りを追加行・削除行とする。
   * 戻り値: [{kind:"cell", rA, rB, c, a, b} | {kind:"added", rB} | {kind:"deleted", rA}]
   * 対応付けの上限を超える場合は null(呼び出し側で位置比較にフォールバック)
   */
  function diffAligned(A, B) {
    var s = span(A, B);
    var ba = bounds(A), bb = bounds(B);
    var nA = ba.r1 - ba.r0, nB = bb.r1 - bb.r0;
    if (nA * nB > MAX_ALIGN) return null;

    function sig(g, r) {
      var cells = [];
      for (var c = s.c0; c < s.c1; c++) cells.push(val(g, r, c));
      return cells.join("\u0001");
    }
    var sa = [], sb = [], i;
    for (i = 0; i < nA; i++) sa.push(sig(A, ba.r0 + i));
    for (i = 0; i < nB; i++) sb.push(sig(B, bb.r0 + i));

    var out = [];
    function gap(a0, a1, b0, b1) { // 対応が付かなかった区間 A[a0,a1) と B[b0,b1)
      var k = Math.min(a1 - a0, b1 - b0), j, c;
      for (j = 0; j < k; j++) {
        var rA = ba.r0 + a0 + j, rB = bb.r0 + b0 + j;
        for (c = s.c0; c < s.c1; c++) {
          var a = val(A, rA, c), b = val(B, rB, c);
          if (a !== b) out.push({ kind: "cell", rA: rA, rB: rB, c: c, a: a, b: b });
        }
      }
      for (j = a0 + k; j < a1; j++) if (sa[j].replace(/\u0001/g, "") !== "") out.push({ kind: "deleted", rA: ba.r0 + j });
      for (j = b0 + k; j < b1; j++) if (sb[j].replace(/\u0001/g, "") !== "") out.push({ kind: "added", rB: bb.r0 + j });
    }
    var pa = 0, pb = 0;
    lcsPairs(sa, sb).forEach(function (p) {
      gap(pa, p[0], pb, p[1]);
      pa = p[0] + 1; pb = p[1] + 1;
    });
    gap(pa, nA, pb, nB);
    return out;
  }

  // 差分1件の表示用テキスト { where, text }
  function describe(d) {
    if (d.kind === "added") return { where: (d.rB + 1) + "行", text: "比較先(B)で追加された行" };
    if (d.kind === "deleted") return { where: (d.rA + 1) + "行", text: "比較元(A)にだけある行(削除)" };
    return {
      where: addr(d.rB, d.c) + (d.rA !== d.rB ? "(A: " + addr(d.rA, d.c) + ")" : ""),
      text: (d.a === "" ? "(空)" : d.a) + " → " + (d.b === "" ? "(空)" : d.b)
    };
  }

  /* ---- ハイライトの塗り・消し ---- */

  function setFills(state, on) {
    return Excel.run(function (ctx) {
      var sheet = ctx.workbook.worksheets.getItemOrNullObject(state.sheetName);
      return ctx.sync().then(function () {
        if (sheet.isNullObject) return; // シートが消えていたら何もしない
        state.cells.forEach(function (cell) {
          var f = sheet.getCell(cell.r, cell.c).format.fill;
          if (on) f.color = HIGHLIGHT_COLOR;
          else if (cell.p) f.color = cell.p; // 元の塗りつぶし色へ戻す
          else f.clear();
        });
        return ctx.sync();
      });
    });
  }

  function updateToggleButton() {
    var btn = $("btn-diff-toggle");
    if (!hl || !hl.cells.length) { btn.hidden = true; return; }
    btn.hidden = false;
    btn.textContent = hl.on ? "色付けを消す" : "色付けを付け直す";
  }

  function toggleHighlight() {
    if (!hl) return;
    var next = !hl.on;
    $("btn-diff-toggle").disabled = true;
    setFills(hl, next).then(function () {
      hl.on = next;
      saveHl();
      updateToggleButton();
      UI.status(next ? "差分セルを色付けしました" : "色付けを消しました(元の色に戻しました)");
    }).catch(function (e) {
      UI.status("切り替えに失敗: " + e.message, true);
    }).finally(function () {
      $("btn-diff-toggle").disabled = false;
    });
  }

  /* ---- シート一覧・比較 ---- */

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
    return loadSheetNames().then(function (names) {
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

  // 表示上の文字列(text)で比べるのが既定 — 生の値で比べると丸め表示や書式違いが誤検出されるため
  function usedRangeOf(ctx, name, prop) {
    var rng = ctx.workbook.worksheets.getItem(name).getUsedRangeOrNullObject(true);
    rng.load(prop + ",rowIndex,columnIndex,isNullObject");
    return rng;
  }

  function toGrid(rng, prop) {
    return rng.isNullObject
      ? { values: [], rowIndex: 0, columnIndex: 0 }
      : { values: rng[prop], rowIndex: rng.rowIndex, columnIndex: rng.columnIndex };
  }

  function compare() {
    var nameA = $("diff-a").value, nameB = $("diff-b").value;
    if (!nameA || !nameB) { UI.status("シートを選択してください", true); return; }
    if (nameA === nameB) { UI.status("異なるシートを選択してください", true); return; }
    var prop = $("diff-mode").value === "formulas" ? "formulas" : "text";
    var align = $("chk-diff-align").checked;

    $("btn-diff-run").disabled = true;
    // 前回のハイライトを必ず消してから比較する(直した差分の色が残らないように)
    var clearOld = (hl && hl.cells.length && hl.on) ? setFills(hl, false) : Promise.resolve();
    clearOld.then(function () {
      hl = null;
      saveHl();
      return Excel.run(function (ctx) {
        var ra = usedRangeOf(ctx, nameA, prop);
        var rb = usedRangeOf(ctx, nameB, prop);
        return ctx.sync().then(function () {
          var A = toGrid(ra, prop), B = toGrid(rb, prop);
          var diffs = align ? diffAligned(A, B) : null;
          var fellBack = align && diffs === null;
          if (diffs === null) diffs = diffGrids(A, B);
          last = { nameA: nameA, nameB: nameB, diffs: diffs };
          renderResult(nameA, nameB, diffs, fellBack);

          var targets = diffs.filter(function (d) { return d.kind === "cell"; }).slice(0, MAX_HIGHLIGHT);
          if (!$("chk-diff-highlight").checked || targets.length === 0) return;
          var sheetB = ctx.workbook.worksheets.getItem(nameB);
          var fills = targets.map(function (d) {
            var f = sheetB.getCell(d.rB, d.c).format.fill;
            f.load("color");
            return f;
          });
          return ctx.sync().then(function () {
            hl = {
              docKey: docKey(),
              sheetName: nameB,
              // 塗りつぶしなしのセルは白(#FFFFFF)として返るため、白は「なし」として扱う
              cells: targets.map(function (d, i) {
                var p = fills[i].color;
                return { r: d.rB, c: d.c, p: !p || /^#?ffffff$/i.test(p) ? "" : p };
              }),
              on: true
            };
            fills.forEach(function (f) { f.color = HIGHLIGHT_COLOR; });
            saveHl();
            return ctx.sync();
          });
        });
      });
    }).catch(function (e) {
      UI.status("比較に失敗しました: " + e.message, true);
    }).finally(function () {
      updateToggleButton();
      $("btn-diff-run").disabled = false;
    });
  }

  function renderResult(nameA, nameB, diffs, fellBack) {
    var sum = $("diff-summary");
    var list = $("diff-list");
    list.innerHTML = "";
    $("btn-diff-export").hidden = diffs.length === 0;
    if (diffs.length === 0) {
      sum.textContent = "差分はありません(「" + nameA + "」と「" + nameB + "」は一致)";
      return;
    }
    sum.textContent = "差分 " + diffs.length + "件" +
      (diffs.length > MAX_LIST ? "(先頭" + MAX_LIST + "件を表示)" : "") +
      " — クリックでそのセルへ移動" +
      (fellBack ? " ※行数が多いため、行の対応付けは行わず同じ位置どうしで比較しました" : "");
    diffs.slice(0, MAX_LIST).forEach(function (d) {
      var t = describe(d);
      var li = document.createElement("button");
      li.type = "button";
      li.className = "diff-row";
      li.innerHTML = '<span class="diff-addr"></span><span class="diff-vals"></span>';
      li.querySelector(".diff-addr").textContent = t.where;
      li.querySelector(".diff-vals").textContent = t.text;
      li.addEventListener("click", function () {
        Excel.run(function (ctx) {
          var sheet = ctx.workbook.worksheets.getItem(d.kind === "deleted" ? nameA : nameB);
          sheet.getCell(d.kind === "deleted" ? d.rA : d.rB, d.kind === "cell" ? d.c : 0).select();
          return ctx.sync();
        }).catch(function () {});
      });
      list.appendChild(li);
    });
  }

  /* ---- 結果を新しいシートへ ---- */

  function exportResult() {
    if (!last || !last.diffs.length) return;
    var rows = [["場所(比較先B)", "場所(比較元A)", "種類", "比較元(A) " + last.nameA, "比較先(B) " + last.nameB]];
    last.diffs.slice(0, MAX_EXPORT).forEach(function (d) {
      if (d.kind === "added") rows.push([(d.rB + 1) + "行", "", "行の追加", "", ""]);
      else if (d.kind === "deleted") rows.push(["", (d.rA + 1) + "行", "行の削除", "", ""]);
      else rows.push([addr(d.rB, d.c), addr(d.rA, d.c), "変更", d.a, d.b]);
    });
    Excel.run(function (ctx) {
      var sheet = ctx.workbook.worksheets.add();
      var rng = sheet.getRangeByIndexes(0, 0, rows.length, rows[0].length);
      // 数式や数値として解釈されないよう、文字列書式にしてから書き込む
      rng.numberFormat = rows.map(function (r) { return r.map(function () { return "@"; }); });
      rng.values = rows;
      rng.getRow(0).format.font.bold = true;
      rng.format.autofitColumns();
      sheet.activate();
      return ctx.sync();
    }).then(function () {
      refreshSheets();
      UI.status("比較結果を新しいシートに出力しました" + (last.diffs.length > MAX_EXPORT ? "(先頭" + MAX_EXPORT + "件)" : ""));
    }).catch(function (e) { UI.status("出力に失敗しました: " + e.message, true); });
  }

  /* ---- 別ブックのシートを取り込む ---- */

  function importWorkbook(file) {
    if (!XL.supports("ExcelApi", "1.13")) { UI.status("このExcelでは別ブックの取り込みに対応していません", true); return; }
    if (file.size > MAX_IMPORT_BYTES) { UI.status("ファイルが大きすぎます(5MBまで)", true); return; }
    var reader = new FileReader();
    reader.onload = function () {
      var base64 = String(reader.result).split(",")[1];
      Excel.run(function (ctx) {
        ctx.workbook.insertWorksheetsFromBase64(base64, { positionType: "End" });
        return ctx.sync();
      }).then(function () {
        return refreshSheets();
      }).then(function () {
        UI.status("「" + file.name + "」のシートをこのブックの末尾に取り込みました。比較後、不要なら削除してください");
      }).catch(function (e) { UI.status("取り込みに失敗しました: " + e.message, true); });
    };
    reader.readAsDataURL(file);
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
      $("btn-diff-toggle").addEventListener("click", toggleHighlight);
      $("btn-diff-export").addEventListener("click", exportResult);
      $("diff-import").addEventListener("change", function () {
        if (this.files[0]) importWorkbook(this.files[0]);
        this.value = "";
      });
      if (UI.host() === "excel") {
        refreshSheets();
        loadHl();
        updateToggleButton();
      }
    },
    _test: { diffGrids: diffGrids, diffAligned: diffAligned, lcsPairs: lcsPairs, colLetter: colLetter }
  });
})();
