/**
 * app.js — ツール①: デート印(押印 / デザイナ / ライブラリ)
 */
(function () {
  "use strict";

  var editing = Stamp.defaultDef(); // デザイナで編集中の定義
  var selectedId = null;            // 押印タブで選択中の印ID
  var previewGen = 0;

  var $ = function (id) { return document.getElementById(id); };
  var status = function (m, e) { UI.status(m, e); };

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
  }

  /* ---- デザイナ: フォーム ⇔ 定義 ---- */

  function fillForm(def) {
    $("f-name").value = def.name || "";
    $("f-diameter").value = def.frame.diameterMm;
    $("f-stroke").value = def.frame.strokeWidthMm;
    $("f-dividers").checked = !!def.frame.dividers;
    $("f-ink").value = def.color.ink;
    $("f-tilt-on-stamp").value = def.options.randomTiltOnStampDeg;

    [0, 1, 2].forEach(function (i) {
      var r = def.rows[i];
      if (i !== 1) {
        $("f-r" + i + "-enabled").checked = !!r.enabled;
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
    editing.frame.diameterMm = clamp(parseFloat($("f-diameter").value) || 12, 8, 30);
    editing.frame.strokeWidthMm = clamp(parseFloat($("f-stroke").value) || 0.35, 0.2, 1.2);
    editing.frame.dividers = $("f-dividers").checked;
    editing.color.ink = $("f-ink").value;
    editing.options.randomTiltOnStampDeg = clamp(parseFloat($("f-tilt-on-stamp").value) || 0, 0, 5);

    [0, 1, 2].forEach(function (i) {
      var r = editing.rows[i];
      if (i !== 1) {
        r.enabled = $("f-r" + i + "-enabled").checked;
        r.text = $("f-r" + i + "-text").value;
        r.lineScales = readLineScales(i);
      } else {
        r.enabled = true;
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
      var dateText = Stamp.formatDate(editing.rows[1].format, now);
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
    if (!editing.wear.seed) editing.wear.seed = Stamp.newSeedHex();
    if (!editing.createdAt) editing.createdAt = new Date().toISOString();
    Store.save(JSON.parse(JSON.stringify(editing)));
    renderPreview();
    status("「" + editing.name + "」を保存しました");
  }

  /* ---- 押印タブ ---- */

  function renderStampList() {
    var list = Store.list();
    var host = $("stamp-list");
    host.innerHTML = "";
    if (list.length === 0) {
      host.innerHTML = '<p class="empty">保存済みの印がありません。「デザイナ」タブで作成してください。</p>';
      return;
    }
    if (!selectedId || !Store.get(selectedId)) selectedId = list[0].id;
    list.forEach(function (def) {
      var item = document.createElement("label");
      item.className = "stamp-item" + (def.id === selectedId ? " selected" : "");
      item.innerHTML =
        '<input type="radio" name="stamp-sel" value="' + def.id + '"' + (def.id === selectedId ? " checked" : "") + ">" +
        '<span class="thumb">' + Stamp.renderSVG(def, { date: new Date() }) + "</span>" +
        '<span class="sname"></span>';
      item.querySelector(".sname").textContent = def.name;
      item.querySelector("input").addEventListener("change", function () {
        selectedId = def.id;
        renderStampList();
      });
      host.appendChild(item);
    });
  }

  function doStamp() {
    var def = Store.get(selectedId);
    if (!def) { status("印を選択してください", true); return; }
    var now = new Date();
    var maxTilt = def.options.randomTiltOnStampDeg || 0;
    var tilt = $("opt-random-tilt").checked && maxTilt > 0 ? (Math.random() * 2 - 1) * maxTilt : 0;
    var dateText = Stamp.formatDate(def.rows[1].format, now);

    $("btn-stamp").disabled = true;
    Stamp.verificationCode(def, dateText)
      .then(function (code) {
        var svg = Stamp.renderSVG(def, { date: now, tiltDeg: tilt, microtext: Stamp.formatCode(code) });
        return Stamp.toPNG(svg, def.frame.diameterMm, 600).then(function (png) {
          return { png: png, code: code };
        });
      })
      .then(function (r) {
        var meta = {
          stampId: def.id,
          name: def.name,
          date: dateText,
          stampedAt: now.toISOString(),
          vcode: r.code
        };
        if (XL.isOffice()) {
          return XL.insertStamp(r.png, def.frame.diameterMm, meta).then(function () {
            status("押印しました(" + meta.date + ")");
          });
        }
        XL.downloadPNG(r.png, def.name + "_" + now.toISOString().slice(0, 10) + ".png");
        status("ブラウザモードのためPNGをダウンロードしました");
      })
      .catch(function (e) { status("押印に失敗しました: " + e.message, true); })
      .finally(function () { $("btn-stamp").disabled = false; });
  }

  /* ---- 押印の検証 ---- */

  function verifyStamps() {
    $("btn-verify").disabled = true;
    XL.collectStampMetas()
      .then(function (metas) {
        if (metas.length === 0) {
          status("この" + (XL.hostType() === "word" ? "文書" : "シート") + "に押印が見つかりません");
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
          var ok = results.filter(function (r) { return r.result === "ok"; }).length;
          var ng = results.filter(function (r) { return r.result === "ng"; });
          var unknown = results.filter(function (r) { return r.result === "unknown" || r.result === "legacy"; }).length;
          var msg = "検証: 正常 " + ok + "件";
          if (ng.length) msg += " / 改ざんの疑い " + ng.length + "件";
          if (unknown) msg += " / 検証不可 " + unknown + "件(手元にない印・旧形式)";
          status(msg, ng.length > 0);
          if (ng.length) {
            alert("改ざんの疑いがある押印:\n" + ng.map(function (r) {
              return "・" + r.m.name + " " + r.m.date + "(コード不一致)";
            }).join("\n"));
          }
        });
      })
      .catch(function (e) { status("検証に失敗しました: " + e.message, true); })
      .finally(function () { $("btn-verify").disabled = false; });
  }

  /* ---- ライブラリタブ ---- */

  function renderLibList() {
    var list = Store.list();
    var host = $("lib-list");
    host.innerHTML = "";
    if (list.length === 0) {
      host.innerHTML = '<p class="empty">保存済みの印がありません。</p>';
      return;
    }
    list.forEach(function (def) {
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
        def.frame.diameterMm + "mm / seed " + (def.wear.seed || "").slice(0, 8) + "…";
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
          var blob = new Blob([Store.exportJSON(def)], { type: "application/json" });
          var a = document.createElement("a");
          a.href = URL.createObjectURL(blob);
          a.download = def.name + ".stamp.json";
          a.click();
          URL.revokeObjectURL(a.href);
          status("書き出しました。※定義ファイルを渡すと欠けごと複製できてしまう点に注意");
        } else if (act === "del") {
          if (confirm("「" + def.name + "」を削除しますか?")) {
            Store.remove(def.id);
            renderLibList();
            status("削除しました");
          }
        }
      });
      host.appendChild(item);
    });
  }

  function importStamp(file) {
    var reader = new FileReader();
    reader.onload = function () {
      try {
        var def = Store.importJSON(reader.result);
        if (Store.get(def.id)) def.id = "stamp_" + Stamp.newSeedHex().slice(0, 12);
        Store.save(def);
        renderLibList();
        status("「" + def.name + "」を取り込みました");
      } catch (e) {
        status(e.message, true);
      }
    };
    reader.readAsText(file);
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
    $("btn-verify").addEventListener("click", verifyStamps);
    $("f-import").addEventListener("change", function () {
      if (this.files[0]) importStamp(this.files[0]);
      this.value = "";
    });

    if (host === "excel") $("btn-stamp").textContent = "選択セルに押印";
    else if (host === "word") $("btn-stamp").textContent = "カーソル位置に押印";
    else {
      $("btn-stamp").textContent = "押印(PNGをダウンロード)";
      $("btn-verify").hidden = true;
    }

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
