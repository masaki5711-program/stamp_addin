/**
 * storage.js — 印ライブラリの保存(プロトタイプは localStorage)
 * 本番では OfficeRuntime.storage への差し替えを想定した薄いラッパ。
 */
(function (global) {
  "use strict";

  var PREFIX = "stampset.";
  var KEY = PREFIX + "stamps.v1";
  var HISTORY_KEY = "history.v1";
  var HISTORY_MAX = 500;
  // 一括バックアップの対象(端末・ブック固有の一時状態は含めない)
  var BACKUP_KEYS = ["stamps.v1", "defaultStamp", "snippets.v1", "profile.name", "formatPresets.v1", "holidays.custom.v1", HISTORY_KEY];

  function readAll() {
    try {
      return JSON.parse(localStorage.getItem(KEY) || "[]");
    } catch (e) {
      return [];
    }
  }

  function writeAll(list) {
    localStorage.setItem(KEY, JSON.stringify(list));
  }

  function list() { return readAll(); }

  function get(id) {
    return readAll().find(function (s) { return s.id === id; }) || null;
  }

  function save(def) {
    var all = readAll();
    var i = all.findIndex(function (s) { return s.id === def.id; });
    if (i >= 0) all[i] = def; else all.push(def);
    writeAll(all);
  }

  function remove(id) {
    writeAll(readAll().filter(function (s) { return s.id !== id; }));
  }

  function exportJSON(def) {
    return JSON.stringify(def, null, 2);
  }

  function importJSON(text) {
    return Stamp.sanitizeDef(JSON.parse(text));
  }

  /* 汎用キー保存(他ツールのプリセット・文例など) */

  function readKey(key, fallback) {
    try {
      var v = localStorage.getItem(PREFIX + key);
      return v === null ? fallback : JSON.parse(v);
    } catch (e) {
      return fallback;
    }
  }

  function writeKey(key, value) {
    localStorage.setItem(PREFIX + key, JSON.stringify(value));
  }

  /* 押印履歴(新しい順・上限あり) */

  function history() {
    var v = readKey(HISTORY_KEY, []);
    return Array.isArray(v) ? v : [];
  }

  function addHistory(entries) {
    writeKey(HISTORY_KEY, entries.concat(history()).slice(0, HISTORY_MAX));
  }

  function clearHistory() { writeKey(HISTORY_KEY, []); }

  /* 一括バックアップ(印・定型文・書式プリセット・会社休日・履歴) */

  function exportAll() {
    var data = {};
    BACKUP_KEYS.forEach(function (k) {
      var v = readKey(k, undefined);
      if (v !== undefined) data[k] = v;
    });
    return JSON.stringify({ app: "stampset", backupVersion: 1, exportedAt: new Date().toISOString(), data: data }, null, 2);
  }

  function str(v, max) { return String(v === null || v === undefined ? "" : v).slice(0, max); }

  // 取り込む値はキーごとに形を検査し、文字列・色などを安全な値に整える
  var CLEANERS = {
    "stamps.v1": function (v) {
      return v.map(function (d) { return Stamp.sanitizeDef(d); });
    },
    "defaultStamp": function (v) { return str(v, 60); },
    "profile.name": function (v) { return str(v, 60); },
    "snippets.v1": function (v) {
      return v.filter(function (s) { return s && typeof s === "object"; }).map(function (s, i) {
        return { id: /^snip_[0-9a-z]+$/.test(s.id) ? s.id : "snip_" + Date.now().toString(36) + i, title: str(s.title, 100), text: str(s.text, 5000) };
      });
    },
    "formatPresets.v1": function (v) {
      function color(c, d) { return /^#[0-9a-fA-F]{6}$/.test(c) ? c : d; }
      return v.filter(function (p) { return p && typeof p === "object"; }).map(function (p) {
        return {
          name: str(p.name, 40), header: !!p.header, headerFill: color(p.headerFill, "#EDEAE3"), headerFont: color(p.headerFont, "#2A2622"),
          band: !!p.band, bandFill: color(p.bandFill, "#F7F5F0"),
          border: ["thin", "medium", "none"].indexOf(p.border) >= 0 ? p.border : "thin", autofit: !!p.autofit
        };
      });
    },
    "holidays.custom.v1": function (v) {
      return v.filter(function (k) { return /^\d{4}-\d{2}-\d{2}$/.test(k); });
    },
    "history.v1": function (v) {
      return v.filter(function (h) { return h && typeof h === "object"; }).slice(0, HISTORY_MAX).map(function (h) {
        return {
          at: str(h.at, 30), stampId: str(h.stampId, 60), name: str(h.name, 60), date: str(h.date, 30),
          vcode: str(h.vcode, 16), doc: str(h.doc, 200), where: str(h.where, 200)
        };
      });
    }
  };

  /**
   * バックアップを取り込む。印と定型文は「同じIDは上書き・それ以外は追加」、その他は置き換え。
   * 戻り値: { stamps: 件数, snippets: 件数 }
   */
  function importAll(text) {
    var obj = JSON.parse(text);
    if (!obj || obj.app !== "stampset" || !obj.data || typeof obj.data !== "object") {
      throw new Error("バックアップファイルの形式が正しくありません");
    }
    var result = { stamps: 0, snippets: 0 };
    BACKUP_KEYS.forEach(function (k) {
      if (!(k in obj.data)) return;
      var raw = obj.data[k];
      var isList = k !== "defaultStamp" && k !== "profile.name";
      if (isList && !Array.isArray(raw)) return;
      var clean = CLEANERS[k](raw);
      if (k === "stamps.v1" || k === "snippets.v1") {
        var cur = readKey(k, []);
        if (!Array.isArray(cur)) cur = [];
        clean.forEach(function (item) {
          var i = cur.findIndex(function (x) { return x.id === item.id; });
          if (i >= 0) cur[i] = item; else cur.push(item);
        });
        writeKey(k, cur);
        result[k === "stamps.v1" ? "stamps" : "snippets"] = clean.length;
      } else if (k === HISTORY_KEY) {
        // 履歴は重複(同じ時刻・同じ場所)を除いて結合する
        var seen = {};
        var merged = history().concat(clean).filter(function (h) {
          var id = h.at + "|" + h.where;
          if (seen[id]) return false;
          seen[id] = true;
          return true;
        }).sort(function (a, b) { return a.at < b.at ? 1 : -1; }).slice(0, HISTORY_MAX);
        writeKey(k, merged);
      } else {
        writeKey(k, clean);
      }
    });
    return result;
  }

  global.Store = {
    list: list, get: get, save: save, remove: remove,
    exportJSON: exportJSON, importJSON: importJSON,
    readKey: readKey, writeKey: writeKey,
    history: history, addHistory: addHistory, clearHistory: clearHistory,
    exportAll: exportAll, importAll: importAll
  };
})(window);
