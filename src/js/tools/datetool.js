/**
 * tools/datetool.js — ツール⑦: 日付ツール(Excel/Word)
 * 和暦変換・各種書式の一覧と挿入・営業日計算。
 */
(function () {
  "use strict";

  var $ = function (id) { return document.getElementById(id); };

  var ERAS = [
    { name: "令和", initial: "R", start: new Date(2019, 4, 1) },
    { name: "平成", initial: "H", start: new Date(1989, 0, 8) },
    { name: "昭和", initial: "S", start: new Date(1926, 11, 25) }
  ];

  function toWareki(d) {
    for (var i = 0; i < ERAS.length; i++) {
      if (d >= ERAS[i].start) {
        var y = d.getFullYear() - ERAS[i].start.getFullYear() + 1;
        return { era: ERAS[i].name, initial: ERAS[i].initial, year: y, yearText: y === 1 ? "元" : String(y) };
      }
    }
    return null;
  }

  function pad2(n) { return String(n).padStart(2, "0"); }
  var YOBI = ["日", "月", "火", "水", "木", "金", "土"];

  function formatsOf(d) {
    var y = d.getFullYear(), m = d.getMonth() + 1, day = d.getDate();
    var w = toWareki(d);
    var MMM = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"][m - 1];
    var out = [
      { label: "西暦(スラッシュ)", text: y + "/" + pad2(m) + "/" + pad2(day) },
      { label: "西暦(ハイフン)", text: y + "-" + pad2(m) + "-" + pad2(day) },
      { label: "西暦(日本語)", text: y + "年" + m + "月" + day + "日" },
      { label: "曜日つき", text: y + "年" + m + "月" + day + "日(" + YOBI[d.getDay()] + ")" }
    ];
    if (w) {
      out.push({ label: "和暦", text: w.era + w.yearText + "年" + m + "月" + day + "日" });
      out.push({ label: "和暦(短縮)", text: w.initial + w.year + "." + m + "." + day });
    }
    out.push({ label: "英語(英式)", text: pad2(day) + " " + MMM + " " + y });
    out.push({ label: "英語(米式)", text: MMM + " " + pad2(day) + ", " + y });
    return out;
  }

  /** 営業日加算(土日を除く。nは正負どちらも可、0はそのまま) */
  function addBusinessDays(d, n) {
    var out = new Date(d.getTime());
    var step = n >= 0 ? 1 : -1;
    var remain = Math.abs(n);
    while (remain > 0) {
      out.setDate(out.getDate() + step);
      var dow = out.getDay();
      if (dow !== 0 && dow !== 6) remain--;
    }
    return out;
  }

  function insertText(text) {
    var host = XL.hostType();
    if (host === "excel") {
      return Excel.run(function (ctx) {
        ctx.workbook.getActiveCell().values = [[text]];
        return ctx.sync();
      });
    }
    if (host === "word") {
      return Word.run(function (ctx) {
        ctx.document.getSelection().insertText(text, Word.InsertLocation.replace);
        return ctx.sync();
      });
    }
    return navigator.clipboard.writeText(text);
  }

  function currentDate() {
    var v = $("dt-date").value;
    if (!v) return new Date();
    var p = v.split("-");
    return new Date(parseInt(p[0], 10), parseInt(p[1], 10) - 1, parseInt(p[2], 10));
  }

  function renderFormats() {
    var host = $("dt-formats");
    host.innerHTML = "";
    var inOffice = XL.isOffice();
    formatsOf(currentDate()).forEach(function (f) {
      var row = document.createElement("div");
      row.className = "clean-row";
      row.innerHTML =
        '<span class="clean-label"><span class="dt-label"></span><span class="dt-text"></span></span>' +
        '<button type="button"></button>';
      row.querySelector(".dt-label").textContent = f.label;
      row.querySelector(".dt-text").textContent = f.text;
      var btn = row.querySelector("button");
      btn.textContent = inOffice ? "挿入" : "コピー";
      btn.addEventListener("click", function () {
        insertText(f.text)
          .then(function () { UI.status(inOffice ? "「" + f.text + "」を挿入しました" : "コピーしました"); })
          .catch(function (e) { UI.status("挿入に失敗: " + e.message, true); });
      });
      host.appendChild(row);
    });
  }

  function calcBiz() {
    var n = parseInt($("dt-biz-n").value, 10) || 0;
    var r = addBusinessDays(currentDate(), n);
    var w = toWareki(r);
    $("dt-biz-result").textContent =
      r.getFullYear() + "/" + pad2(r.getMonth() + 1) + "/" + pad2(r.getDate()) +
      "(" + YOBI[r.getDay()] + ")" + (w ? " ・ " + w.era + w.yearText + "年" : "");
    $("btn-dt-biz-insert").dataset.text =
      r.getFullYear() + "/" + pad2(r.getMonth() + 1) + "/" + pad2(r.getDate());
    $("btn-dt-biz-insert").hidden = false;
  }

  function todayStr() {
    var d = new Date();
    return d.getFullYear() + "-" + pad2(d.getMonth() + 1) + "-" + pad2(d.getDate());
  }

  UI.registerTool({
    id: "datetool",
    name: "日付ツール",
    desc: "和暦変換・書式いろいろ・営業日計算",
    hosts: ["excel", "word", "browser"],
    init: function () {
      $("dt-date").value = todayStr();
      renderFormats();
      $("dt-date").addEventListener("change", function () { renderFormats(); $("btn-dt-biz-insert").hidden = true; $("dt-biz-result").textContent = ""; });
      $("btn-dt-biz-calc").addEventListener("click", calcBiz);
      $("btn-dt-biz-insert").addEventListener("click", function () {
        var t = this.dataset.text;
        if (!t) return;
        insertText(t)
          .then(function () { UI.status(XL.isOffice() ? "挿入しました" : "コピーしました"); })
          .catch(function (e) { UI.status("挿入に失敗: " + e.message, true); });
      });
      $("btn-dt-today").addEventListener("click", function () {
        $("dt-date").value = todayStr();
        renderFormats();
      });
    },
    _test: { toWareki: toWareki, addBusinessDays: addBusinessDays, formatsOf: formatsOf }
  });
})();
