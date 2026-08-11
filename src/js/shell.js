/**
 * shell.js — 便利アドインセットのシェル(ホーム画面・ツール切替・共通UI)
 * 各ツールは UI.registerTool({id, name, desc, hosts, init}) で登録する。
 *  hosts: "excel" | "word" | "browser" の配列(利用可能なホスト)
 */
(function (global) {
  "use strict";

  var tools = [];
  var host = "browser";
  var statusTimer = null;

  function $(id) { return document.getElementById(id); }

  function status(msg, isError) {
    var el = $("status");
    el.textContent = msg;
    el.className = "status" + (isError ? " error" : "");
    if (statusTimer) clearTimeout(statusTimer);
    if (msg) statusTimer = setTimeout(function () { el.textContent = ""; }, 8000);
  }

  function registerTool(t) { tools.push(t); }

  function show(toolId) {
    document.querySelectorAll(".tool-view").forEach(function (v) {
      v.hidden = v.id !== "view-" + toolId;
    });
    var home = toolId === "home";
    $("view-home").hidden = !home;
    $("btn-back").hidden = home;
    var t = tools.find(function (x) { return x.id === toolId; });
    $("shell-title").textContent = home ? "ホーム" : (t ? t.name : "");
    status("");
  }

  function buildHome() {
    var grid = $("home-grid");
    grid.innerHTML = "";
    tools.forEach(function (t) {
      var ok = t.hosts.indexOf(host) >= 0;
      var card = document.createElement("button");
      card.type = "button";
      card.className = "tool-card" + (ok ? "" : " disabled");
      card.disabled = !ok;
      card.innerHTML =
        '<span class="tool-icon"></span>' +
        '<span class="tool-card-body"><span class="tool-card-name"></span>' +
        '<span class="tool-card-desc"></span></span>';
      card.querySelector(".tool-icon").textContent = t.name.charAt(0);
      card.querySelector(".tool-card-name").textContent = t.name;
      card.querySelector(".tool-card-desc").textContent = ok ? t.desc : t.unavailableNote || "このアプリでは使えません";
      if (ok) card.addEventListener("click", function () { show(t.id); });
      grid.appendChild(card);
    });
  }

  function bootstrap() {
    XL.init().then(function (h) {
      host = h || "browser";
      $("host-mode").textContent =
        host === "excel" ? "Excel に接続中" :
        host === "word" ? "Word に接続中" :
        "ブラウザモード(一部機能のみ)";
      buildHome();
      tools.forEach(function (t) {
        if (t.init) {
          try { t.init(host); } catch (e) { console.error("tool init failed: " + t.id, e); }
        }
      });
      show("home");
    });
    $("btn-back").addEventListener("click", function () { show("home"); });
  }

  global.UI = {
    status: status,
    registerTool: registerTool,
    show: show,
    host: function () { return host; },
    bootstrap: bootstrap
  };

  if (global.Office && Office.onReady) {
    Office.onReady().then(function () { bootstrap(); });
  } else {
    document.addEventListener("DOMContentLoaded", bootstrap);
  }
})(window);
