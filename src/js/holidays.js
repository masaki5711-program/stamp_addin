/**
 * holidays.js — 日本の祝日(計算で求める)と、利用者が登録する会社休日
 * 対象は現行の祝日法(2020年以降。2020・2021年の五輪特例を含む)。春分・秋分は近似式(〜2099年)。
 */
(function (global) {
  "use strict";

  var CUSTOM_KEY = "holidays.custom.v1";
  var cache = {};

  function pad2(n) { return String(n).padStart(2, "0"); }
  function key(y, m, d) { return y + "-" + pad2(m) + "-" + pad2(d); }
  function keyOf(date) { return key(date.getFullYear(), date.getMonth() + 1, date.getDate()); }

  // m月の第n月曜日
  function nthMonday(y, m, n) {
    var first = new Date(y, m - 1, 1).getDay();
    return 1 + ((8 - first) % 7) + (n - 1) * 7;
  }

  function springEquinox(y) {
    return Math.floor(20.8431 + 0.242194 * (y - 1980) - Math.floor((y - 1980) / 4));
  }

  function autumnEquinox(y) {
    return Math.floor(23.2488 + 0.242194 * (y - 1980) - Math.floor((y - 1980) / 4));
  }

  /** その年の祝日 { "YYYY-MM-DD": 名称 }(振替休日・国民の休日を含む) */
  function ofYear(y) {
    if (cache[y]) return cache[y];
    var h = {};
    function add(m, d, name) { h[key(y, m, d)] = name; }

    add(1, 1, "元日");
    add(1, nthMonday(y, 1, 2), "成人の日");
    add(2, 11, "建国記念の日");
    add(2, 23, "天皇誕生日");
    add(3, springEquinox(y), "春分の日");
    add(4, 29, "昭和の日");
    add(5, 3, "憲法記念日");
    add(5, 4, "みどりの日");
    add(5, 5, "こどもの日");
    if (y === 2020) { add(7, 23, "海の日"); add(7, 24, "スポーツの日"); add(8, 10, "山の日"); }
    else if (y === 2021) { add(7, 22, "海の日"); add(7, 23, "スポーツの日"); add(8, 8, "山の日"); }
    else {
      add(7, nthMonday(y, 7, 3), "海の日");
      add(8, 11, "山の日");
      add(10, nthMonday(y, 10, 2), "スポーツの日");
    }
    add(9, nthMonday(y, 9, 3), "敬老の日");
    add(9, autumnEquinox(y), "秋分の日");
    add(11, 3, "文化の日");
    add(11, 23, "勤労感謝の日");

    // 振替休日: 祝日が日曜なら、その後の最初の「祝日でない日」
    Object.keys(h).sort().forEach(function (k) {
      var p = k.split("-");
      var d = new Date(+p[0], +p[1] - 1, +p[2]);
      if (d.getDay() !== 0) return;
      do { d.setDate(d.getDate() + 1); } while (h[keyOf(d)]);
      h[keyOf(d)] = "振替休日";
    });
    // 国民の休日: 前日と翌日が祝日に挟まれた平日
    Object.keys(h).sort().forEach(function (k) {
      var p = k.split("-");
      var mid = new Date(+p[0], +p[1] - 1, +p[2] + 1);
      var next = new Date(+p[0], +p[1] - 1, +p[2] + 2);
      if (!h[keyOf(mid)] && h[keyOf(next)] && mid.getDay() !== 0) h[keyOf(mid)] = "国民の休日";
    });

    cache[y] = h;
    return h;
  }

  /* ---- 会社休日(利用者登録) ---- */

  // テキスト(1行1日、YYYY-MM-DD / YYYY/M/D、後ろにメモ可)から日付キーの配列にする
  function parseCustom(text) {
    var out = [];
    String(text || "").split(/\r?\n/).forEach(function (line) {
      var m = line.match(/(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/);
      if (!m) return;
      var d = new Date(+m[1], +m[2] - 1, +m[3]);
      if (d.getMonth() !== +m[2] - 1) return; // 2/30 などの不正日付
      var k = keyOf(d);
      if (out.indexOf(k) < 0) out.push(k);
    });
    return out.sort();
  }

  function customList() {
    var v = global.Store ? Store.readKey(CUSTOM_KEY, []) : [];
    return Array.isArray(v) ? v : [];
  }

  function saveCustom(text) {
    var list = parseCustom(text);
    Store.writeKey(CUSTOM_KEY, list);
    return list;
  }

  /** 休日名を返す(祝日名 / "会社休日" / 平日は null)。土日は含まない */
  function nameOf(date) {
    var jp = ofYear(date.getFullYear())[keyOf(date)];
    if (jp) return jp;
    return customList().indexOf(keyOf(date)) >= 0 ? "会社休日" : null;
  }

  function isBusinessDay(date, useHolidays) {
    var dow = date.getDay();
    if (dow === 0 || dow === 6) return false;
    return !(useHolidays && nameOf(date));
  }

  global.Holidays = {
    ofYear: ofYear,
    nameOf: nameOf,
    isBusinessDay: isBusinessDay,
    parseCustom: parseCustom,
    customList: customList,
    saveCustom: saveCustom,
    keyOf: keyOf
  };
})(window);
