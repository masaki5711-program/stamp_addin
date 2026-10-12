/**
 * tools/snippets.js — ツール⑤: 定型文ライブラリ(Excel/Word)
 * よく使う文例を保存し、セル/カーソル位置にワンクリック挿入。
 */
(function () {
  "use strict";

  var $ = function (id) { return document.getElementById(id); };
  var KEY = "snippets.v1";
  var editingId = null;

  var SAMPLES = [
    { id: "snip_sample1", title: "お礼(社内)", text: "お疲れさまです。ご対応ありがとうございました。" },
    { id: "snip_sample2", title: "確認依頼", text: "お手数ですが、下記についてご確認をお願いいたします。" }
  ];

  function load() {
    var list = Store.readKey(KEY, null);
    if (list === null) { list = SAMPLES; Store.writeKey(KEY, list); }
    return list;
  }

  function saveList(list) { Store.writeKey(KEY, list); }

  var YOBI = ["日", "月", "火", "水", "木", "金", "土"];
  var NAME_KEY = "profile.name";

  function pad2(n) { return String(n).padStart(2, "0"); }

  // 例: 令和8年10月12日(元号と年はブラウザの和暦カレンダーから取る)
  function wareki(d) {
    var era = "", year = "";
    new Intl.DateTimeFormat("ja-JP-u-ca-japanese", { era: "long", year: "numeric" }).formatToParts(d).forEach(function (p) {
      if (p.type === "era") era = p.value;
      if (p.type === "year") year = p.value;
    });
    return era + (year === "1" ? "元" : year) + "年" + (d.getMonth() + 1) + "月" + d.getDate() + "日";
  }

  // 本文中の {今日} {和暦} {曜日} {時刻} {氏名} を挿入時の値に置き換える
  function expand(text, now) {
    now = now || new Date();
    var vars = {
      "今日": now.getFullYear() + "/" + pad2(now.getMonth() + 1) + "/" + pad2(now.getDate()),
      "和暦": wareki(now),
      "曜日": YOBI[now.getDay()],
      "時刻": pad2(now.getHours()) + ":" + pad2(now.getMinutes()),
      "氏名": Store.readKey(NAME_KEY, "")
    };
    return text.replace(/\{(今日|和暦|曜日|時刻|氏名)\}/g, function (_, k) { return vars[k]; });
  }

  function insertText(text) {
    text = expand(text);
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
    return navigator.clipboard.writeText(text); // ブラウザモードはコピー
  }

  function render() {
    var host = $("snippet-list");
    var list = load();
    host.innerHTML = "";
    if (list.length === 0) {
      host.innerHTML = '<p class="empty">文例がありません。下のフォームから追加してください。</p>';
      return;
    }
    var inOffice = XL.isOffice();
    list.forEach(function (s) {
      var item = document.createElement("div");
      item.className = "lib-item";
      item.innerHTML =
        '<div class="lib-main"><div class="sname"></div><div class="snip-preview"></div>' +
        '<div class="lib-btns">' +
        '<button data-act="ins" class="ins"></button>' +
        '<button data-act="edit">編集</button>' +
        '<button data-act="del" class="danger">削除</button>' +
        "</div></div>";
      item.querySelector(".sname").textContent = s.title;
      item.querySelector(".snip-preview").textContent = s.text.length > 60 ? s.text.slice(0, 60) + "…" : s.text;
      item.querySelector(".ins").textContent = inOffice ? "挿入" : "コピー";
      item.addEventListener("click", function (ev) {
        var act = ev.target.dataset && ev.target.dataset.act;
        if (!act) return;
        if (act === "ins") {
          insertText(s.text)
            .then(function () { UI.status(inOffice ? "「" + s.title + "」を挿入しました" : "コピーしました"); })
            .catch(function (e) { UI.status("挿入に失敗: " + e.message, true); });
        } else if (act === "edit") {
          editingId = s.id;
          $("snip-title").value = s.title;
          $("snip-text").value = s.text;
          $("btn-snip-save").textContent = "更新";
          $("snip-title").focus();
        } else if (act === "del") {
          if (UI.confirmClick(ev.target)) {
            saveList(load().filter(function (x) { return x.id !== s.id; }));
            if (editingId === s.id) resetForm();
            render();
          }
        }
      });
      host.appendChild(item);
    });
  }

  function resetForm() {
    editingId = null;
    $("snip-title").value = "";
    $("snip-text").value = "";
    $("btn-snip-save").textContent = "追加";
  }

  function save() {
    var title = $("snip-title").value.trim();
    var text = $("snip-text").value;
    if (!title || !text.trim()) { UI.status("タイトルと本文を入力してください", true); return; }
    var list = load();
    if (editingId) {
      var s = list.find(function (x) { return x.id === editingId; });
      if (s) { s.title = title; s.text = text; }
    } else {
      list.push({ id: "snip_" + Date.now().toString(36), title: title, text: text });
    }
    saveList(list);
    resetForm();
    render();
    UI.status("文例を保存しました");
  }

  UI.registerTool({
    id: "snippets",
    name: "定型文",
    desc: "よく使う文例を保存してワンクリック挿入",
    hosts: ["excel", "word", "browser"],
    _test: { expand: expand },
    init: function () {
      render();
      $("btn-snip-save").addEventListener("click", save);
      $("btn-snip-new").addEventListener("click", function () { resetForm(); $("snip-title").focus(); });
      $("snip-profile-name").value = Store.readKey(NAME_KEY, "");
      $("snip-profile-name").addEventListener("change", function () {
        Store.writeKey(NAME_KEY, this.value.trim().slice(0, 60));
        UI.status("差し込み用の氏名を保存しました");
      });
    }
  });
})();
