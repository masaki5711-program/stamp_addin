/**
 * storage.js — 印ライブラリの保存(プロトタイプは localStorage)
 * 本番では OfficeRuntime.storage への差し替えを想定した薄いラッパ。
 */
(function (global) {
  "use strict";

  var KEY = "stampset.stamps.v1";

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
    var def = JSON.parse(text);
    if (!def || def.schema !== 1 || !def.frame || !def.rows) {
      throw new Error("印定義ファイルの形式が正しくありません");
    }
    return def;
  }

  /* 汎用キー保存(他ツールのプリセット・文例など) */

  function readKey(key, fallback) {
    try {
      var v = localStorage.getItem("stampset." + key);
      return v === null ? fallback : JSON.parse(v);
    } catch (e) {
      return fallback;
    }
  }

  function writeKey(key, value) {
    localStorage.setItem("stampset." + key, JSON.stringify(value));
  }

  global.Store = {
    list: list, get: get, save: save, remove: remove,
    exportJSON: exportJSON, importJSON: importJSON,
    readKey: readKey, writeKey: writeKey
  };
})(window);
