/**
 * app.js — ツール①: デート印(押印 / デザイナ / ライブラリ / 履歴)
 */
(function () {
  "use strict";

  var DEFAULT_KEY = "defaultStamp"; // 最後に選んだ印。リボンの「押印」ボタンもこの印を使う

  var editing = Stamp.defaultDef(); // デザイナで編集中の定義
  var selectedId = Store.readKey(DEFAULT_KEY, null); // 押印タブで選択中の印ID
  var previewGen = 0;

  var $ = function (id) { return document.getElementById(id); };
  var status = function (m, e) { UI.status(m, e); };

  function pad2(n) { return String(n).padStart(2, "0"); }
  function ymd(d) { return d.getFullYear() + "-" + pad2(d.getMonth() + 1) + "-" + pad2(d.getDate()); }

  /* ---- タブ ---- */

  function showTab(name) {
    document.querySelectorAll("#view-stamp .tab").forEach(function (b) {
      b.classList.toggle("active", b.dataset.tab === name);
    });
    document.querySelectorAll("#view-stamp .panel").forEach(function (p) {
      p.hidden = p.id !== "panel-" + name;
    });
    if (name === "stamp") renderStampList();
    if (name === "library") renderLibList();
    if (name === "history") renderHistory();
  }

  /* ---- デザイナ: フォーム ⇔ 定義 ---- */

  function fillForm(def) {
    $("f-name").value = def.name || "";
    $("f-shape").value = Stamp.SHAPES[def.frame.shape] ? def.frame.shape : "circle";
    $("f-diameter").value = def.frame.diameterMm;
    $("f-stroke").value = def.frame.strokeWidthMm;
    $("f-dividers").checked = !!def.frame.dividers;
    $("f-double-frame").checked = !!def.frame.doubleFrame;
    $("f-double-dividers").checked = !!def.frame.doubleDividers;
    $("f-code-on-frame").checked = !!def.options.codeOnFrame;
    $("f-ink").value = def.color.ink;
    $("f-tilt-on-stamp").value = def.options.randomTiltOnStampDeg;

    [0, 1, 2].forEach(function (i) {
      var r = def.rows[i];
      $("f-r" + i + "-enabled").checked = !!r.enabled;
      if (i !== 1) {
        $("f-r" + i + "-text").value = r.text || "";
        syncLineScales(i, r.lineScales);
      } else {
        $("f-r1-format").value = r.format;
      }
      $("f-r" + i + "-font").value = r.font;
      $("f-r" + i + "-ratio").value = r.heightRatio;
      $("f-r" + i + "-tilt").value = r.tiltDeg || 0;
      $("f-r" + i + "-spacing").value = r.letterSpacing || 0;
      $("f-r" + i + "-bold").checked = r.bold === undefined ? r.role === "date" : !!r.bold;
    });
  }

  function readForm() {
    editing.name = $("f-name").value.trim();
    editing.frame.shape = $("f-shape").value;
    editing.frame.diameterMm = clamp(parseFloat($("f-diameter").value) || 12, 8, 30);
    editing.frame.strokeWidthMm = clamp(parseFloat($("f-stroke").value) || 0.35, 0.2, 1.2);
    editing.frame.dividers = $("f-dividers").checked;
    editing.frame.doubleFrame = $("f-double-frame").checked;
    editing.frame.doubleDividers = $("f-double-dividers").checked;
    editing.options.codeOnFrame = $("f-code-on-frame").checked;
    editing.color.ink = $("f-ink").value;
    editing.options.randomTiltOnStampDeg = clamp(parseFloat($("f-tilt-on-stamp").value) || 0, 0, 5);

    [0, 1, 2].forEach(function (i) {
      var r = editing.rows[i];
      r.enabled = $("f-r" + i + "-enabled").checked;
      if (i !== 1) {
        r.text = $("f-r" + i + "-text").value;
        r.lineScales = readLineScales(i);
      } else {
        r.format = $("f-r1-format").value;
      }
      r.font = $("f-r" + i + "-font").value;
      r.heightRatio = clamp(parseFloat($("f-r" + i + "-ratio").value) || 0.3, 0.15, 0.5);
      r.tiltDeg = clamp(parseFloat($("f-r" + i + "-tilt").value) || 0, -30, 30);
      r.letterSpacing = clamp(parseFloat($("f-r" + i + "-spacing").value) || 0, 0, 0.5);
      r.bold = $("f-r" + i + "-bold").checked;
    });
  }

  function clamp(v, lo, hi) { return Math.min(hi, Math.max(lo, v)); }

  /* 上段・下段の行ごと大きさスライダー(2行以上のときだけ表示) */

  function syncLineScales(i, scalesFromDef) {
    var host = $("f-r" + i + "-lines");
    var lines = $("f-r" + i + "-text").value.split(/\r?\n/);
    var current = scalesFromDef || readLineScales(i);
    host.innerHTML = "";
    if (lines.length < 2) return;
    lines.forEach(function (line, idx) {
      var label = document.createElement("label");
      label.className = "line-scale";
      var name = document.createElement("span");
      name.textContent = (idx + 1) + "行目の大きさ" + (line.trim() ? "「" + line.trim().slice(0, 6) + "」" : "");
      var input = document.createElement("input");
      input.type = "range";
      input.min = "0.5"; input.max = "1.5"; input.step = "0.05";
      input.value = current[idx] || 1;
      label.appendChild(name);
      label.appendChild(input);
      host.appendChild(label);
    });
  }

  function readLineScales(i) {
    var sliders = $("f-r" + i + "-lines").querySelectorAll("input");
    if (sliders.length === 0) return [1];
    return Array.from(sliders).map(function (inp) { return parseFloat(inp.value) || 1; });
  }

  /* ---- デザイナ: 操作 ---- */

  function renderPreview() {
    var gen = ++previewGen;
    var now = new Date();
    if (editing.wear.seed) {
      var dateText = Stamp.formatDate(Stamp.dateFormatOf(editing), now);
      Stamp.verificationCode(editing, dateText).then(function (code) {
        if (gen !== previewGen) return; // 古いプレビューは破棄
        $("preview").innerHTML = Stamp.renderSVG(editing, { date: now, microtext: Stamp.formatCode(code) });
      });
    } else {
      $("preview").innerHTML = Stamp.renderSVG(editing, { date: now });
    }
    $("seed-text").textContent = editing.wear.seed
      ? editing.wear.seed
      : "(未生成 — 保存時に欠けを自動生成します)";
  }

  function onDesignerInput(ev) {
    if (ev && ev.target) {
      if (ev.target.id === "f-r0-text") syncLineScales(0);
      if (ev.target.id === "f-r2-text") syncLineScales(2);
    }
    readForm();
    renderPreview();
  }

  function newStamp() {
    editing = Stamp.defaultDef();
    fillForm(editing);
    renderPreview();
    status("新しい印の編集を開始しました");
  }

  function reseed() {
    editing.wear.seed = Stamp.newSeedHex();
    renderPreview();
    status("欠けを再抽選しました(新しい個体になります)");
  }

  function saveStamp() {
    readForm();
    if (!editing.name) { status("印の名前を入力してください", true); return; }
    if (!editing.rows.some(function (r) { return r.enabled; })) { status("上段・中段・下段のどれかは使用してください", true); return; }
    if (!editing.wear.seed) editing.wear.seed = Stamp.newSeedHex();
    if (!editing.createdAt) editing.createdAt = new Date().toISOString();
    Store.save(JSON.parse(JSON.stringify(editing)));
    renderPreview();
    status("「" + editing.name + "」を保存しました");
  }

  /* ---- 押印タブ ---- */

  function selectedDef() {
    var list = Store.list();
    if (list.length === 0) return null;
    if (!selectedId || !Store.get(selectedId)) select(list[0].id);
    return Store.get(selectedId);
  }

  function select(id) {
    selectedId = id;
    Store.writeKey(DEFAULT_KEY, id);
  }

  function renderStampList() {
    var list = Store.list();
    var host = $("stamp-list");
    host.innerHTML = "";
    if (list.length === 0) {
      host.innerHTML = '<p class="empty">保存済みの印がありません。「デザイナ」タブで作成してください。</p>';
      return;
    }
    selectedDef();
    list.forEach(function (def) {
      var item = document.createElement("label");
      item.className = "stamp-item" + (def.id === selectedId ? " selected" : "");
      item.innerHTML =
        '<input type="radio" name="stamp-sel">' +
        '<span class="thumb">' + Stamp.renderSVG(def, { date: new Date() }) + "</span>" +
        '<span class="sname"></span>';
      item.querySelector(".sname").textContent = def.name;
      var radio = item.querySelector("input");
      radio.checked = def.id === selectedId;
      radio.addEventListener("change", function () {
        select(def.id);
        renderStampList();
      });
      host.appendChild(item);
    });
  }

  // 押印日(未指定なら今日)。日付だけを使うので時刻は正午に固定する
  function stampDate() {
    var v = $("opt-stamp-date").value;
    if (!v) return new Date();
    var p = v.split("-");
    return new Date(+p[0], +p[1] - 1, +p[2], 12, 0, 0);
  }

  function doStamp() {
    var def = selectedDef();
    if (!def) { status("印を選択してください", true); return; }
    var when = stampDate();
    var backdated = ymd(when) !== ymd(new Date());

    $("btn-stamp").disabled = true;
    StampAction.run(def, {
      date: when,
      randomTilt: $("opt-random-tilt").checked,
      perCell: XL.hostType() === "excel" && $("opt-per-cell").checked
    })
      .then(function (r) {
        if (!XL.isOffice()) { status("ブラウザモードのためPNGをダウンロードしました"); return; }
        status("押印しました(" + r.meta.date + (r.places.length > 1 ? " ・ " + r.places.length + "か所" : "") + ")" +
          (backdated ? " ※押印日を指定しています" : ""));
      })
      .catch(function (e) { status("押印に失敗しました: " + e.message, true); })
      .finally(function () { $("btn-stamp").disabled = false; });
  }

  /* ---- 承認欄 ---- */

  function createApprovalBox() {
    var labels = $("approval-labels").value.split(/[,、,]/)
      .map(function (s) { return s.trim(); })
      .filter(function (s) { return s; });
    if (labels.length === 0 || labels.length > 10) { status("見出しは1〜10個で指定してください", true); return; }
    var def = selectedDef() || Stamp.defaultDef();
    $("btn-approval").disabled = true;
    XL.createApprovalBox(labels, Stamp.sizeMm(def))
      .then(function (address) {
        $("opt-per-cell").checked = true;
        status("承認欄を作成しました(" + address + ")。押印行を選択中です");
      })
      .catch(function (e) { status("承認欄の作成に失敗しました: " + e.message, true); })
      .finally(function () { $("btn-approval").disabled = false; });
  }

  /* ---- 押印の検証 ---- */

  var VERIFY_LABEL = { ok: "正常", ng: "不一致", unknown: "検証不可(手元にない印)", legacy: "検証不可(旧形式)" };

  function verifyStamps() {
    $("btn-verify").disabled = true;
    var host = $("verify-list");
    host.innerHTML = "";
    XL.collectStampMetas()
      .then(function (metas) {
        if (metas.length === 0) {
          status("この" + (XL.hostType() === "word" ? "文書" : "ブック") + "に押印が見つかりません");
          return;
        }
        var checks = metas.map(function (m) {
          var def = Store.get(m.stampId);
          if (!def) return Promise.resolve({ m: m, result: "unknown" });
          if (!m.vcode) return Promise.resolve({ m: m, result: "legacy" });
          return Stamp.verificationCode(def, m.date).then(function (code) {
            return { m: m, result: code === m.vcode ? "ok" : "ng" };
          });
        });
        return Promise.all(checks).then(function (results) {
          var count = function (keys) {
            return results.filter(function (r) { return keys.indexOf(r.result) >= 0; }).length;
          };
          var ok = count(["ok"]), ng = count(["ng"]), unknown = count(["unknown", "legacy"]);
          var msg = "検証: 正常 " + ok + "件";
          if (ng) msg += " / 改ざんの疑い " + ng + "件";
          if (unknown) msg += " / 検証不可 " + unknown + "件";
          status(msg, ng > 0);
          results.forEach(function (r) {
            var row = document.createElement("div");
            row.className = "clean-row" + (r.result === "ng" ? " ng" : "");
            row.innerHTML = '<span class="clean-label"><span class="dt-label"></span><span class="dt-text"></span></span><span class="verify-result"></span>';
            row.querySelector(".dt-label").textContent = r.m.where || "";
            row.querySelector(".dt-text").textContent = (r.m.name || "") + " " + (r.m.date || "");
            row.querySelector(".verify-result").textContent = VERIFY_LABEL[r.result];
            host.appendChild(row);
          });
        });
      })
      .catch(function (e) { status("検証に失敗しました: " + e.message, true); })
      .finally(function () { $("btn-verify").disabled = false; });
  }

  /* ---- ライブラリタブ ---- */

  // 書き出し: ダウンロードを試みつつ、内容をテキスト欄にも出す
  // (Officeのタスクペインではファイルのダウンロードが始まらない環境があるため)
  function exportText(text, filename) {
    $("export-box").value = text;
    $("export-wrap").hidden = false;
    try { XL.downloadText(text, filename); } catch (e) { /* テキスト欄からコピーできる */ }
  }

  function renderLibList() {
    var list = Store.list();
    var host = $("lib-list");
    host.innerHTML = "";
    if (list.length === 0) {
      host.innerHTML = '<p class="empty">保存済みの印がありません。</p>';
      return;
    }
    list.forEach(function (def) {
      var size = Stamp.sizeMm(def);
      var item = document.createElement("div");
      item.className = "lib-item";
      item.innerHTML =
        '<span class="thumb">' + Stamp.renderSVG(def, { date: new Date() }) + "</span>" +
        '<div class="lib-main"><div class="sname"></div><div class="smeta"></div>' +
        '<div class="lib-btns">' +
        '<button data-act="edit">編集</button>' +
        '<button data-act="dup">複製</button>' +
        '<button data-act="export">書き出し</button>' +
        '<button data-act="del" class="danger">削除</button>' +
        "</div></div>";
      item.querySelector(".sname").textContent = def.name;
      item.querySelector(".smeta").textContent =
        (size.w === size.h ? size.w : size.w.toFixed(1) + "×" + size.h.toFixed(1)) + "mm / seed " + (def.wear.seed || "").slice(0, 8) + "…";
      item.addEventListener("click", function (ev) {
        var act = ev.target.dataset && ev.target.dataset.act;
        if (!act) return;
        if (act === "edit") {
          editing = JSON.parse(JSON.stringify(def));
          fillForm(editing);
          renderPreview();
          showTab("designer");
          status("「" + def.name + "」を編集中(欠けは維持されます)");
        } else if (act === "dup") {
          var copy = JSON.parse(JSON.stringify(def));
          copy.id = "stamp_" + Stamp.newSeedHex().slice(0, 12);
          copy.name = def.name + "(複製)";
          copy.wear.seed = Stamp.newSeedHex(); // 複製は別個体
          copy.createdAt = new Date().toISOString();
          Store.save(copy);
          renderLibList();
          status("複製しました(欠けは新しい個体として再抽選)");
        } else if (act === "export") {
          exportText(Store.exportJSON(def), def.name + ".stamp.json");
          status("書き出しました。※定義ファイルを渡すと欠けごと複製できてしまう点に注意");
        } else if (act === "del") {
          if (UI.confirmClick(ev.target)) {
            Store.remove(def.id);
            renderLibList();
            status("削除しました");
          }
        }
      });
      host.appendChild(item);
    });
  }

  // 印定義(1個)とバックアップ(一式)のどちらも受け付ける
  function importText(text) {
    try {
      var obj = JSON.parse(text);
      if (obj && obj.app === "stampset") {
        var r = Store.importAll(text);
        status("バックアップを取り込みました(印 " + r.stamps + "件・定型文 " + r.snippets + "件)。画面を読み込み直します");
        setTimeout(function () { UI.reload(); }, 1500); // 他ツールの一覧にも反映する
        return;
      }
      var def = Store.importJSON(text);
      if (Store.get(def.id)) def.id = "stamp_" + Stamp.newSeedHex().slice(0, 12);
      Store.save(def);
      renderLibList();
      status("「" + def.name + "」を取り込みました");
    } catch (e) {
      status(e instanceof SyntaxError ? "JSONとして読み取れませんでした" : e.message, true);
    }
  }

  function importFile(file) {
    var reader = new FileReader();
    reader.onload = function () { importText(reader.result); };
    reader.readAsText(file);
  }

  /* ---- 履歴タブ ---- */

  function fmtAt(iso) {
    var d = new Date(iso);
    if (isNaN(d)) return iso;
    return ymd(d).replace(/-/g, "/") + " " + pad2(d.getHours()) + ":" + pad2(d.getMinutes());
  }

  function renderHistory() {
    var list = Store.history();
    var host = $("history-list");
    host.innerHTML = "";
    $("btn-history-export").hidden = $("btn-history-clear").hidden = list.length === 0;
    if (list.length === 0) {
      host.innerHTML = '<p class="empty">まだ押印の記録がありません。</p>';
      return;
    }
    list.slice(0, 200).forEach(function (h) {
      var row = document.createElement("div");
      row.className = "clean-row";
      row.innerHTML = '<span class="clean-label"><span class="dt-label"></span><span class="dt-text"></span></span>';
      row.querySelector(".dt-label").textContent = fmtAt(h.at) + " ・ " + (h.doc || "(未保存の文書)") + " ・ " + h.where;
      row.querySelector(".dt-text").textContent = h.name + " " + h.date;
      host.appendChild(row);
    });
  }

  var HISTORY_HEAD = ["押印日時", "印の名前", "印面の日付", "検証コード", "文書", "場所"];

  function historyRows() {
    return Store.history().map(function (h) { return [fmtAt(h.at), h.name, h.date, h.vcode, h.doc, h.where]; });
  }

  function exportHistory() {
    var rows = [HISTORY_HEAD].concat(historyRows());
    if (XL.hostType() === "excel") {
      Excel.run(function (ctx) {
        var sheet = ctx.workbook.worksheets.add();
        // 先頭が "=" などの文字列が数式として解釈されないよう、文字列書式にしてから書き込む
        var rng = sheet.getRangeByIndexes(0, 0, rows.length, HISTORY_HEAD.length);
        rng.numberFormat = rows.map(function (r) { return r.map(function () { return "@"; }); });
        rng.values = rows;
        rng.getRow(0).format.font.bold = true;
        rng.format.autofitColumns();
        sheet.activate();
        return ctx.sync();
      }).then(function () { status("押印履歴を新しいシートに出力しました"); })
        .catch(function (e) { status("出力に失敗しました: " + e.message, true); });
      return;
    }
    var csv = rows.map(function (r) {
      return r.map(function (v) { return '"' + String(v || "").replace(/"/g, '""') + '"'; }).join(",");
    }).join("\r\n");
    XL.downloadText("﻿" + csv, "stamp-history_" + ymd(new Date()) + ".csv", "text/csv");
    status("押印履歴をCSVで書き出しました");
  }

  /* ---- 初期化 ---- */

  function init(host) {
    var sel = $("f-r1-format");
    Object.keys(Stamp.DATE_FORMATS).forEach(function (k) {
      var o = document.createElement("option");
      o.value = k;
      o.textContent = Stamp.DATE_FORMATS[k].label;
      sel.appendChild(o);
    });
    var shapeSel = $("f-shape");
    Object.keys(Stamp.SHAPES).forEach(function (k) {
      var o = document.createElement("option");
      o.value = k;
      o.textContent = Stamp.SHAPES[k].label;
      shapeSel.appendChild(o);
    });
    var inkSel = $("f-ink-preset");
    Object.keys(Stamp.INK_PRESETS).forEach(function (k) {
      var o = document.createElement("option");
      o.value = Stamp.INK_PRESETS[k];
      o.textContent = { shu: "朱(標準)", koshu: "濃朱", usushu: "薄朱", blue: "藍", black: "墨" }[k];
      inkSel.appendChild(o);
    });
    inkSel.addEventListener("change", function () {
      $("f-ink").value = inkSel.value;
      onDesignerInput();
    });

    document.querySelectorAll("#view-stamp .tab").forEach(function (b) {
      b.addEventListener("click", function () { showTab(b.dataset.tab); });
    });
    $("designer-form").addEventListener("input", onDesignerInput);
    $("btn-new").addEventListener("click", newStamp);
    $("btn-reseed").addEventListener("click", reseed);
    $("btn-save").addEventListener("click", saveStamp);
    $("btn-stamp").addEventListener("click", doStamp);
    $("btn-stamp-today").addEventListener("click", function () { $("opt-stamp-date").value = ymd(new Date()); });
    $("btn-verify").addEventListener("click", verifyStamps);
    $("btn-approval").addEventListener("click", createApprovalBox);

    $("f-import").addEventListener("change", function () {
      if (this.files[0]) importFile(this.files[0]);
      this.value = "";
    });
    $("btn-import-text").addEventListener("click", function () {
      var t = $("import-text").value.trim();
      if (!t) { status("取り込む内容を貼り付けてください", true); return; }
      importText(t);
      $("import-text").value = "";
    });
    $("btn-backup-export").addEventListener("click", function () {
      exportText(Store.exportAll(), "stampset-backup_" + ymd(new Date()) + ".json");
      status("バックアップを書き出しました");
    });
    $("btn-export-copy").addEventListener("click", function () {
      UI.copyFrom($("export-box"));
      status("コピーしました。メモ帳などに貼り付けて .json で保存してください");
    });
    $("btn-history-export").addEventListener("click", exportHistory);
    $("btn-history-clear").addEventListener("click", function () {
      if (!UI.confirmClick(this)) return;
      Store.clearHistory();
      renderHistory();
      status("押印履歴を消去しました");
    });

    $("opt-stamp-date").value = ymd(new Date());
    if (host === "excel") $("btn-stamp").textContent = "選択セルに押印";
    else if (host === "word") $("btn-stamp").textContent = "カーソル位置に押印";
    else {
      $("btn-stamp").textContent = "押印(PNGをダウンロード)";
      $("btn-verify").hidden = true;
    }
    // 各セルへの一括押印と承認欄はExcelのみ
    $("opt-per-cell-wrap").hidden = $("approval-box").hidden = host !== "excel";

    fillForm(editing);
    renderPreview();
    showTab("stamp");
  }

  UI.registerTool({
    id: "stamp",
    name: "デート印",
    desc: "日付が自動で入る電子日付印を作成・押印",
    hosts: ["excel", "word", "browser"],
    init: init
  });
})();
