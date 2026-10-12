/**
 * stamp.js — 印影の定義・描画・欠け生成・日付書式
 * 座標系: 1単位 = 0.1mm(直径12mmの印 → viewBox 0 0 120 120)
 */
(function (global) {
  "use strict";

  var FONT_MAP = {
    mincho: '"Yu Mincho","YuMincho","Hiragino Mincho ProN","MS PMincho",serif',
    gothic: '"Yu Gothic","YuGothic","Hiragino Sans","Meiryo",sans-serif'
  };

  var INK_PRESETS = {
    shu: "#C13B2A",     // 朱(標準)
    koshu: "#A32418",   // 濃朱
    usushu: "#D96A54",  // 薄朱
    blue: "#1F4E9C",    // 藍(事務印)
    black: "#33302C"    // 墨
  };

  var DATE_FORMATS = {
    "'YY.MM.DD": { label: "'26.08.10(日付印標準)" },
    "YYYY.MM.DD": { label: "2026.08.10" },
    "YYYY/MM/DD": { label: "2026/08/10" },
    "RY.M.D": { label: "R8.8.10(和暦)" },
    "令和Y.M.D": { label: "令和8.8.10" },
    "DD MMM YYYY": { label: "10 AUG 2026(英式)" },
    "MMM DD, YYYY": { label: "AUG 10, 2026(米式)" },
    "MM/DD/YYYY": { label: "08/10/2026(米式)" },
    "DD/MM/YYYY": { label: "10/08/2026(欧州式)" },
    "YYYY-MM-DD": { label: "2026-08-10(ISO)" }
  };

  function pad2(n) { return String(n).padStart(2, "0"); }

  function formatDate(fmt, d) {
    var y = d.getFullYear(), m = d.getMonth() + 1, day = d.getDate();
    var reiwa = y - 2018; // 令和元年 = 2019
    var MMM = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"][m - 1];
    switch (fmt) {
      case "'YY.MM.DD": return "'" + pad2(y % 100) + "." + pad2(m) + "." + pad2(day);
      case "YYYY.MM.DD": return y + "." + pad2(m) + "." + pad2(day);
      case "YYYY/MM/DD": return y + "/" + pad2(m) + "/" + pad2(day);
      case "RY.M.D": return "R" + reiwa + "." + m + "." + day;
      case "令和Y.M.D": return "令和" + reiwa + "." + m + "." + day;
      case "DD MMM YYYY": return pad2(day) + " " + MMM + " " + y;
      case "MMM DD, YYYY": return MMM + " " + pad2(day) + ", " + y;
      case "MM/DD/YYYY": return pad2(m) + "/" + pad2(day) + "/" + y;
      case "DD/MM/YYYY": return pad2(day) + "/" + pad2(m) + "/" + y;
      case "YYYY-MM-DD": return y + "-" + pad2(m) + "-" + pad2(day);
      default: return y + "." + pad2(m) + "." + pad2(day);
    }
  }

  /* ---- 乱数 ---- */

  function newSeedHex() {
    var b = new Uint8Array(16);
    crypto.getRandomValues(b);
    return Array.from(b).map(function (x) { return x.toString(16).padStart(2, "0"); }).join("");
  }

  // 128bitシード全体を使う sfc32
  function rngFromSeed(seedHex) {
    var w = [];
    for (var i = 0; i < 4; i++) w.push(parseInt(seedHex.slice(i * 8, i * 8 + 8), 16) >>> 0);
    var a = w[0], b = w[1], c = w[2], d = w[3];
    return function () {
      a >>>= 0; b >>>= 0; c >>>= 0; d >>>= 0;
      var t = (a + b) | 0;
      a = b ^ (b >>> 9);
      b = (c + (c << 3)) | 0;
      c = (c << 21) | (c >>> 11);
      d = (d + 1) | 0;
      t = (t + d) | 0;
      c = (c + t) | 0;
      return (t >>> 0) / 4294967296;
    };
  }

  /* ---- 印定義 ---- */

  function defaultDef() {
    return {
      schema: 1,
      id: "stamp_" + newSeedHex().slice(0, 12),
      name: "",
      createdAt: null,
      frame: { shape: "circle", diameterMm: 12, strokeWidthMm: 0.35, dividers: true, doubleFrame: false, doubleDividers: false },
      color: { ink: INK_PRESETS.shu },
      rows: [
        { role: "text", enabled: true, text: "営業部", font: "mincho", heightRatio: 0.28, tiltDeg: 0, letterSpacing: 0.08, lineScales: [1], bold: false },
        { role: "date", enabled: true, format: "'YY.MM.DD", font: "gothic", heightRatio: 0.34, tiltDeg: 0, letterSpacing: 0, bold: true },
        { role: "text", enabled: true, text: "藤川", font: "mincho", heightRatio: 0.28, tiltDeg: 0, letterSpacing: 0.15, lineScales: [1], bold: false }
      ],
      wear: { seed: null },
      options: { randomTiltOnStampDeg: 3, codeOnFrame: false }
    };
  }

  /* ---- 枠の寸法 ----
   * innerR は文字・区切り線を収める内側の半径。
   * 二重枠(frame.doubleFrame)のときは外枠の内側にすき間を空けて細い内枠を引く。
   */
  function frameGeom(def) {
    var S = def.frame.diameterMm * 10;
    var c = S / 2;
    var strokeW = def.frame.strokeWidthMm * 10;
    var g = { S: S, c: c, strokeW: strokeW, frameR: c - strokeW / 2, innerR: c - strokeW, inner: null, gapR: 0 };
    if (def.frame.doubleFrame) {
      var gap = S * 0.035, w2 = strokeW * 0.5;
      g.gapR = g.innerR - gap / 2;
      g.inner = { r: g.innerR - gap - w2 / 2, w: w2 };
      g.innerR = g.inner.r - w2 / 2;
    }
    return g;
  }

  /* ---- 欠けの導出(シードから決定論的) ----
   * 実物の摩耗に似せ、枠の縁に沿った浅い楕円ニックのみ生成する。
   * 各ニック: {x, y, rx, ry, rot}(rotは度)
   */

  function deriveWear(def) {
    if (!def.wear || !def.wear.seed) return [];
    var rnd = rngFromSeed(def.wear.seed);
    var g = frameGeom(def);
    var c = g.c, strokeW = g.strokeW, frameR = g.frameR, innerR = g.innerR;
    var nicks = [];

    // 枠のニック: 3〜6箇所。外縁/内縁のどちらかを接線方向に浅く欠く
    var nFrame = 3 + Math.floor(rnd() * 4);
    for (var i = 0; i < nFrame; i++) {
      var th = rnd() * Math.PI * 2;
      var edge = rnd() < 0.5 ? 1 : -1;
      var rr = frameR + edge * strokeW * 0.55;
      nicks.push({
        x: c + rr * Math.cos(th),
        y: c + rr * Math.sin(th),
        rx: strokeW * (0.6 + rnd() * 1.2),   // 接線方向の長さ
        ry: strokeW * (0.25 + rnd() * 0.45), // 径方向の深さ
        rot: (th * 180 / Math.PI) + 90
      });
    }
    // 区切り線のニック: 0〜2箇所。線の上下どちらかの縁を浅く欠く
    var nDiv = Math.floor(rnd() * 3);
    var divYs = dividerYs(def);
    for (var j = 0; j < nDiv && divYs.length > 0; j++) {
      var y = divYs[Math.floor(rnd() * divYs.length)];
      var dy = y - c;
      var half = Math.sqrt(Math.max(0, innerR * innerR - dy * dy));
      nicks.push({
        x: c + (rnd() * 2 - 1) * half * 0.8,
        y: y + (rnd() < 0.5 ? 1 : -1) * strokeW * 0.3,
        rx: strokeW * (0.5 + rnd() * 0.8),
        ry: strokeW * (0.2 + rnd() * 0.35),
        rot: 0
      });
    }
    return nicks;
  }

  /* ---- 検証コード ----
   * SHA-256(印ID | シード | 日付文字列) の先頭8桁(16進大文字)。
   * シードを知らない第三者は、日付を変えた押印に正しいコードを付けられない。
   */
  function verificationCode(def, dateText) {
    var data = new TextEncoder().encode(def.id + "|" + (def.wear && def.wear.seed) + "|" + dateText);
    return crypto.subtle.digest("SHA-256", data).then(function (buf) {
      return Array.from(new Uint8Array(buf).slice(0, 4))
        .map(function (b) { return b.toString(16).padStart(2, "0"); })
        .join("").toUpperCase();
    });
  }

  function formatCode(code) {
    return code.slice(0, 4) + "·" + code.slice(4, 8);
  }

  /* ---- レイアウト ---- */

  function enabledRows(def) {
    return def.rows.filter(function (r) { return r.enabled; });
  }

  // 有効行の帯(バンド)の上端/下端を返す
  function bands(def) {
    var g = frameGeom(def);
    var c = g.c, innerR = g.innerR;
    var top = c - innerR * 0.96, bottom = c + innerR * 0.96;
    var rows = enabledRows(def);
    var total = rows.reduce(function (s, r) { return s + r.heightRatio; }, 0) || 1;
    var out = [], y = top;
    rows.forEach(function (r) {
      var h = (bottom - top) * (r.heightRatio / total);
      out.push({ row: r, y0: y, y1: y + h });
      y += h;
    });
    return out;
  }

  function dividerYs(def) {
    var bs = bands(def);
    var ys = [];
    for (var i = 1; i < bs.length; i++) ys.push(bs[i].y0);
    return ys;
  }

  /* ---- 文字幅のはめ込み ---- */

  var measureCtx = null;

  // フォントサイズ1あたりの文字列幅(字送り込み)。canvasが使えない環境は概算にフォールバック
  function widthPerEm(text, family, weight, letterSpacing) {
    var n = Array.from(text).length;
    var w;
    try {
      if (!measureCtx) measureCtx = document.createElement("canvas").getContext("2d");
      measureCtx.font = weight + " 100px " + family;
      w = measureCtx.measureText(text).width / 100;
    } catch (e) {
      w = Array.from(text).reduce(function (a, ch) { return a + (ch.charCodeAt(0) < 0x2000 ? 0.6 : 1); }, 0);
    }
    return w + n * letterSpacing;
  }

  // 円の内側(その高さでの弦の幅)に収まる最大のフォントサイズを返す。収まっていれば fontSize のまま
  function fitFontSize(fontSize, perEm, cy, c, innerR) {
    if (perEm <= 0) return fontSize;
    var f = fontSize;
    for (var i = 0; i < 8; i++) {
      var d = Math.abs(cy - c) + f * 0.45; // 文字の上下端のうち中心から遠い側
      var avail = 2 * Math.sqrt(Math.max(0, innerR * innerR - d * d)) * 0.94;
      if (perEm * f <= avail) break;
      f = Math.min(f * 0.95, avail / perEm);
    }
    return f;
  }

  /* ---- SVG描画 ---- */

  function esc(s) {
    return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }

  /**
   * renderSVG(def, opts)
   *  opts.date       — Dateオブジェクト(date行の描画に使用。省略時は今日)
   *  opts.tiltDeg    — 印全体の傾き(押印時ランダム傾き用)
   *  opts.showWear   — 欠けを描くか(既定 true)
   * 欠けはSVGマスク(透過)で抜くため、背景色に依存しない。
   */
  function renderSVG(def, opts) {
    opts = opts || {};
    var date = opts.date || new Date();
    var tilt = opts.tiltDeg || 0;
    var showWear = opts.showWear !== false;

    var g = frameGeom(def);
    var S = g.S, c = g.c, strokeW = g.strokeW, frameR = g.frameR, innerR = g.innerR;
    var ink = def.color.ink;

    var nicks = showWear ? deriveWear(def) : [];
    var maskId = "wear_" + (def.wear && def.wear.seed ? def.wear.seed.slice(0, 8) : "none");

    var s = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ' + S + " " + S + '" width="' + S + '" height="' + S + '">';
    s += "<defs><mask id=\"" + maskId + "\">";
    s += '<rect x="0" y="0" width="' + S + '" height="' + S + '" fill="#fff"/>';
    nicks.forEach(function (ch) {
      s += '<ellipse cx="' + ch.x.toFixed(2) + '" cy="' + ch.y.toFixed(2) + '" rx="' + ch.rx.toFixed(2) + '" ry="' + ch.ry.toFixed(2) +
        '" transform="rotate(' + ch.rot.toFixed(1) + " " + ch.x.toFixed(2) + " " + ch.y.toFixed(2) + ')" fill="#000"/>';
    });
    s += "</mask></defs>";

    s += '<g mask="url(#' + maskId + ')" transform="rotate(' + tilt.toFixed(2) + " " + c + " " + c + ')">';
    s += '<circle cx="' + c + '" cy="' + c + '" r="' + frameR + '" fill="none" stroke="' + ink + '" stroke-width="' + strokeW + '"/>';
    if (g.inner) {
      s += '<circle cx="' + c + '" cy="' + c + '" r="' + g.inner.r.toFixed(2) + '" fill="none" stroke="' + ink + '" stroke-width="' + g.inner.w.toFixed(2) + '"/>';
    }

    if (def.frame.dividers) {
      // 二重線(frame.doubleDividers)は細線2本を上下に振り分ける
      var dbl = !!def.frame.doubleDividers;
      var divW = strokeW * (dbl ? 0.4 : 0.7);
      dividerYs(def).forEach(function (y0) {
        (dbl ? [y0 - strokeW * 0.55, y0 + strokeW * 0.55] : [y0]).forEach(function (y) {
          var dy = y - c;
          var half = Math.sqrt(Math.max(0, innerR * innerR - dy * dy)) * 0.99;
          s += '<line x1="' + (c - half).toFixed(2) + '" y1="' + y.toFixed(2) + '" x2="' + (c + half).toFixed(2) + '" y2="' + y.toFixed(2) +
            '" stroke="' + ink + '" stroke-width="' + divW.toFixed(2) + '"/>';
        });
      });
    }

    bands(def).forEach(function (b) {
      var row = b.row;
      var bandH = b.y1 - b.y0;
      var bandCy = (b.y0 + b.y1) / 2;
      // bold未指定の旧定義は従来どおり(日付のみ太字)
      var weight = row.bold === undefined
        ? (row.role === "date" ? 700 : 500)
        : (row.bold ? 700 : 500);
      // 上段・下段は複数行(改行区切り)に対応。行ごとの大きさは lineScales の比率で配分する。
      var lines = row.role === "date"
        ? [formatDate(row.format, date)]
        : String(row.text || "").split(/\r?\n/);
      var scales = row.lineScales || [];
      var weights = lines.map(function (_, i) {
        return Math.min(2, Math.max(0.4, scales[i] || 1));
      });
      var totalW = weights.reduce(function (a, w) { return a + w; }, 0) || 1;
      var inner = "";
      var y = b.y0;
      lines.forEach(function (line, i) {
        var h = bandH * weights[i] / totalW;
        var cy = y + h / 2;
        y += h;
        if (!line.trim()) return;
        var family = FONT_MAP[row.font || "gothic"];
        // 文字数が多い行は、円からはみ出さない大きさまで縮める
        var fontSize = fitFontSize(h * 0.68, widthPerEm(line, family, weight, row.letterSpacing || 0), cy, c, innerR);
        inner +='<text x="' + c + '" y="' + (cy + fontSize * 0.34).toFixed(2) + '" text-anchor="middle" fill="' + ink +
          '" font-size="' + fontSize.toFixed(2) + '" font-weight="' + weight +
          '" font-family=\'' + family + "'" +
          ' letter-spacing="' + ((row.letterSpacing || 0) * fontSize).toFixed(2) + '"' +
          ">" + esc(line) + "</text>";
      });
      if (!inner) return;
      s += row.tiltDeg
        ? '<g transform="rotate(' + row.tiltDeg + " " + c + " " + bandCy.toFixed(2) + ')">' + inner + "</g>"
        : inner;
    });

    // マイクロ文字(検証コード): 約0.25mm。options.codeOnFrame のときは枠の下側に沿わせる
    // (二重枠なら外枠と内枠のすき間、単枠なら枠のすぐ内側)。それ以外は日付段の下端に印字
    if (opts.microtext && def.options && def.options.codeOnFrame) {
      var csize = S * 0.021;
      var rb = g.inner ? g.gapR + csize * 0.35 : innerR - csize * 0.5;
      var arcId = maskId + "_code";
      s += '<path id="' + arcId + '" fill="none" d="M' + (c - rb).toFixed(2) + " " + c + " A" + rb.toFixed(2) + " " + rb.toFixed(2) +
        " 0 0 0 " + (c + rb).toFixed(2) + " " + c + '"/>';
      s += '<text fill="' + ink + '" font-size="' + csize.toFixed(2) + '" font-family="Consolas, monospace" letter-spacing="' +
        (csize * 0.25).toFixed(2) + '" text-anchor="middle"><textPath href="#' + arcId + '" startOffset="50%">' +
        esc(opts.microtext) + "</textPath></text>";
    } else if (opts.microtext) {
      var bs = bands(def);
      var target = bs.find(function (b) { return b.row.role === "date"; }) || bs[bs.length - 1];
      if (target) {
        var msize = S * 0.021;
        s += '<text x="' + c + '" y="' + (target.y1 - msize * 0.3).toFixed(2) + '" text-anchor="middle" fill="' + ink +
          '" font-size="' + msize.toFixed(2) + '" font-family="Consolas, monospace" letter-spacing="' + (msize * 0.25).toFixed(2) + '">' +
          esc(opts.microtext) + "</text>";
      }
    }

    s += "</g></svg>";
    return s;
  }

  /**
   * SVG → PNG(dataURL)。dpi指定で実寸解像度を決める(既定300dpi)。背景は透過。
   */
  function toPNG(svgString, diameterMm, dpi) {
    dpi = dpi || 300;
    var px = Math.ceil(diameterMm * dpi / 25.4);
    return new Promise(function (resolve, reject) {
      var img = new Image();
      img.onload = function () {
        var canvas = document.createElement("canvas");
        canvas.width = px; canvas.height = px;
        var ctx = canvas.getContext("2d");
        ctx.drawImage(img, 0, 0, px, px);
        resolve(canvas.toDataURL("image/png"));
      };
      img.onerror = function () { reject(new Error("SVGのPNG変換に失敗しました")); };
      img.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svgString);
    });
  }

  global.Stamp = {
    FONT_MAP: FONT_MAP,
    INK_PRESETS: INK_PRESETS,
    DATE_FORMATS: DATE_FORMATS,
    defaultDef: defaultDef,
    newSeedHex: newSeedHex,
    formatDate: formatDate,
    deriveWear: deriveWear,
    verificationCode: verificationCode,
    formatCode: formatCode,
    renderSVG: renderSVG,
    toPNG: toPNG
  };
})(window);
