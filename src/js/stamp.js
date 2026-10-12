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

  // 枠の形。w/h は「大きさ(長辺)」に対する幅・高さの比、corner は角丸半径の比
  var SHAPES = {
    circle: { label: "丸", w: 1, h: 1, corner: 0 },
    ellipse: { label: "楕円(横長)", w: 1, h: 0.72, corner: 0 },
    oval: { label: "小判(縦長)", w: 0.72, h: 1, corner: 0 },
    square: { label: "角(角丸四角)", w: 1, h: 1, corner: 0.12 }
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

  function clampNum(v, lo, hi, fallback) {
    v = Number(v);
    return isFinite(v) ? Math.min(hi, Math.max(lo, v)) : fallback;
  }

  /**
   * 外部から来た印定義(JSON取り込み・バックアップ復元)を検査し、安全な値だけで組み直す。
   * 色・書体などはSVGに埋め込まれるため、想定外の文字列をそのまま通さない。
   */
  function sanitizeDef(raw) {
    if (!raw || typeof raw !== "object" || raw.schema !== 1 || !raw.frame || !Array.isArray(raw.rows) || raw.rows.length !== 3) {
      throw new Error("印定義ファイルの形式が正しくありません");
    }
    var base = defaultDef();
    var f = raw.frame, o = raw.options || {};
    var def = {
      schema: 1,
      id: /^stamp_[0-9a-f]{6,32}$/.test(raw.id) ? raw.id : base.id,
      name: String(raw.name || "").slice(0, 60),
      createdAt: typeof raw.createdAt === "string" && !isNaN(Date.parse(raw.createdAt)) ? raw.createdAt : null,
      frame: {
        shape: SHAPES[f.shape] ? f.shape : "circle",
        diameterMm: clampNum(f.diameterMm, 8, 30, 12),
        strokeWidthMm: clampNum(f.strokeWidthMm, 0.2, 1.2, 0.35),
        dividers: !!f.dividers,
        doubleFrame: !!f.doubleFrame,
        doubleDividers: !!f.doubleDividers
      },
      color: { ink: raw.color && /^#[0-9a-fA-F]{6}$/.test(raw.color.ink) ? raw.color.ink : INK_PRESETS.shu },
      rows: raw.rows.map(function (r, i) {
        r = r || {};
        var b = base.rows[i];
        var row = {
          role: b.role,
          enabled: r.enabled !== false,
          font: FONT_MAP[r.font] ? r.font : b.font,
          heightRatio: clampNum(r.heightRatio, 0.15, 0.5, b.heightRatio),
          tiltDeg: clampNum(r.tiltDeg, -30, 30, 0),
          letterSpacing: clampNum(r.letterSpacing, 0, 0.5, 0),
          bold: r.bold === undefined ? b.bold : !!r.bold
        };
        if (b.role === "date") {
          row.format = DATE_FORMATS[r.format] ? r.format : b.format;
        } else {
          row.text = String(r.text || "").slice(0, 200);
          row.lineScales = (Array.isArray(r.lineScales) ? r.lineScales : [1]).slice(0, 10)
            .map(function (s) { return clampNum(s, 0.4, 2, 1); });
        }
        return row;
      }),
      wear: { seed: raw.wear && /^[0-9a-f]{32}$/.test(raw.wear.seed) ? raw.wear.seed : null },
      options: {
        randomTiltOnStampDeg: clampNum(o.randomTiltOnStampDeg, 0, 5, 3),
        codeOnFrame: !!o.codeOnFrame
      }
    };
    return def;
  }

  /* ---- 枠の寸法 ----
   * 輪郭は「外形から t だけ内側に寄せた形」{a, b, r}(半幅・半高・角丸半径)で表す。
   * content は文字・区切り線を収める内側の輪郭。
   * 二重枠(frame.doubleFrame)のときは外枠の内側にすき間を空けて細い内枠を引く。
   */
  function shapeOf(def) { return SHAPES[def.frame.shape] || SHAPES.circle; }

  function sizeMm(def) {
    var sh = shapeOf(def);
    return { w: def.frame.diameterMm * sh.w, h: def.frame.diameterMm * sh.h };
  }

  function frameGeom(def) {
    var sh = shapeOf(def);
    var S = def.frame.diameterMm * 10;
    var W = S * sh.w, H = S * sh.h;
    var strokeW = def.frame.strokeWidthMm * 10;
    var g = {
      S: S, W: W, H: H, cx: W / 2, cy: H / 2, strokeW: strokeW,
      rect: sh.corner > 0, corner: S * sh.corner,
      inner: null, gapT: 0
    };
    g.at = function (t) {
      return { a: W / 2 - t, b: H / 2 - t, r: g.rect ? Math.max(0, g.corner - t) : 0 };
    };
    g.frame = g.at(strokeW / 2);
    var contentT = strokeW;
    if (def.frame.doubleFrame) {
      var gap = S * 0.035, w2 = strokeW * 0.5;
      g.gapT = strokeW + gap / 2;
      g.inner = { o: g.at(strokeW + gap + w2 / 2), w: w2 };
      contentT = strokeW + gap + w2;
    }
    g.content = g.at(contentT);
    return g;
  }

  // 輪郭 o の、中心から縦に dy 離れた位置での半幅
  function halfWidthAt(g, o, dy) {
    dy = Math.abs(dy);
    if (dy >= o.b) return 0;
    if (!g.rect) return o.a * Math.sqrt(1 - (dy / o.b) * (dy / o.b));
    var k = dy - (o.b - o.r);
    return k <= 0 ? o.a : o.a - o.r + Math.sqrt(Math.max(0, o.r * o.r - k * k));
  }

  // 輪郭 o 上の、角度 th 方向の点と外向き法線(欠けの配置用)
  function outlinePoint(g, o, th) {
    var cs = Math.cos(th), sn = Math.sin(th);
    if (!g.rect) {
      var nx = cs / o.a, ny = sn / o.b, nl = Math.sqrt(nx * nx + ny * ny) || 1;
      return { x: o.a * cs, y: o.b * sn, nx: nx / nl, ny: ny / nl };
    }
    var k = Math.min(o.a / Math.max(1e-9, Math.abs(cs)), o.b / Math.max(1e-9, Math.abs(sn)));
    var x = cs * k, y = sn * k;
    var ex = Math.abs(x) - (o.a - o.r), ey = Math.abs(y) - (o.b - o.r);
    if (ex > 0 && ey > 0) { // 角丸部分
      var l = Math.sqrt(ex * ex + ey * ey) || 1;
      var ux = (x < 0 ? -1 : 1) * ex / l, uy = (y < 0 ? -1 : 1) * ey / l;
      return {
        x: (x < 0 ? -1 : 1) * (o.a - o.r) + ux * o.r,
        y: (y < 0 ? -1 : 1) * (o.b - o.r) + uy * o.r,
        nx: ux, ny: uy
      };
    }
    return Math.abs(Math.abs(x) - o.a) < 1e-6
      ? { x: x, y: y, nx: x < 0 ? -1 : 1, ny: 0 }
      : { x: x, y: y, nx: 0, ny: y < 0 ? -1 : 1 };
  }

  function outlineSVG(g, o, ink, w) {
    var attrs = ' fill="none" stroke="' + ink + '" stroke-width="' + Number(w.toFixed(2)) + '"/>';
    if (g.rect) {
      return '<rect x="' + (g.cx - o.a).toFixed(2) + '" y="' + (g.cy - o.b).toFixed(2) + '" width="' + (o.a * 2).toFixed(2) +
        '" height="' + (o.b * 2).toFixed(2) + '" rx="' + o.r.toFixed(2) + '"' + attrs;
    }
    return '<ellipse cx="' + g.cx + '" cy="' + g.cy + '" rx="' + Number(o.a.toFixed(2)) + '" ry="' + Number(o.b.toFixed(2)) + '"' + attrs;
  }

  /* ---- 欠けの導出(シードから決定論的) ----
   * 実物の摩耗に似せ、枠の縁に沿った浅い楕円ニックのみ生成する。
   * 各ニック: {x, y, rx, ry, rot}(rotは度)
   */

  function deriveWear(def) {
    if (!def.wear || !def.wear.seed) return [];
    var rnd = rngFromSeed(def.wear.seed);
    var g = frameGeom(def);
    var strokeW = g.strokeW;
    var nicks = [];

    // 枠のニック: 3〜6箇所。外縁/内縁のどちらかを接線方向に浅く欠く
    var nFrame = 3 + Math.floor(rnd() * 4);
    for (var i = 0; i < nFrame; i++) {
      var th = rnd() * Math.PI * 2;
      var edge = rnd() < 0.5 ? 1 : -1;
      var p = outlinePoint(g, g.frame, th);
      var off = edge * strokeW * 0.55;
      nicks.push({
        x: g.cx + p.x + p.nx * off,
        y: g.cy + p.y + p.ny * off,
        rx: strokeW * (0.6 + rnd() * 1.2),   // 接線方向の長さ
        ry: strokeW * (0.25 + rnd() * 0.45), // 径方向の深さ
        rot: (Math.atan2(p.ny, p.nx) * 180 / Math.PI) + 90
      });
    }
    // 区切り線のニック: 0〜2箇所。線の上下どちらかの縁を浅く欠く
    var nDiv = Math.floor(rnd() * 3);
    var divYs = dividerYs(def);
    for (var j = 0; j < nDiv && divYs.length > 0; j++) {
      var y = divYs[Math.floor(rnd() * divYs.length)];
      var half = halfWidthAt(g, g.content, y - g.cy);
      nicks.push({
        x: g.cx + (rnd() * 2 - 1) * half * 0.8,
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

  // 日付段の書式(日付段を使わない印でも検証コード・履歴用に日付文字列は作る)
  function dateFormatOf(def) {
    var row = def.rows.find(function (r) { return r.role === "date"; });
    return row ? row.format : "YYYY.MM.DD";
  }

  // 有効行の帯(バンド)の上端/下端を返す
  function bands(def) {
    var g = frameGeom(def);
    var top = g.cy - g.content.b * 0.96, bottom = g.cy + g.content.b * 0.96;
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

  // 枠の内側(その高さでの幅)に収まる最大のフォントサイズを返す。収まっていれば fontSize のまま
  function fitFontSize(fontSize, perEm, lineCy, g) {
    if (perEm <= 0) return fontSize;
    var f = fontSize;
    for (var i = 0; i < 8; i++) {
      var d = Math.abs(lineCy - g.cy) + f * 0.45; // 文字の上下端のうち中心から遠い側
      var avail = 2 * halfWidthAt(g, g.content, d) * 0.94;
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
    var tilt = Number(opts.tiltDeg) || 0;
    var showWear = opts.showWear !== false;

    var g = frameGeom(def);
    var W = g.W, H = g.H, cx = g.cx, cy = g.cy, strokeW = g.strokeW;
    var ink = /^#[0-9a-fA-F]{6}$/.test(def.color && def.color.ink) ? def.color.ink : INK_PRESETS.shu;

    var nicks = showWear ? deriveWear(def) : [];
    var maskId = "wear_" + (def.wear && /^[0-9a-f]{8}/.test(def.wear.seed) ? def.wear.seed.slice(0, 8) : "none");

    var s = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ' + W + " " + H + '" width="' + W + '" height="' + H + '">';
    s += "<defs><mask id=\"" + maskId + "\">";
    s += '<rect x="0" y="0" width="' + W + '" height="' + H + '" fill="#fff"/>';
    nicks.forEach(function (ch) {
      s += '<ellipse cx="' + ch.x.toFixed(2) + '" cy="' + ch.y.toFixed(2) + '" rx="' + ch.rx.toFixed(2) + '" ry="' + ch.ry.toFixed(2) +
        '" transform="rotate(' + ch.rot.toFixed(1) + " " + ch.x.toFixed(2) + " " + ch.y.toFixed(2) + ')" fill="#000"/>';
    });
    s += "</mask></defs>";

    s += '<g mask="url(#' + maskId + ')" transform="rotate(' + tilt.toFixed(2) + " " + cx + " " + cy + ')">';
    s += outlineSVG(g, g.frame, ink, strokeW);
    if (g.inner) s += outlineSVG(g, g.inner.o, ink, g.inner.w);

    if (def.frame.dividers) {
      // 二重線(frame.doubleDividers)は細線2本を上下に振り分ける
      var dbl = !!def.frame.doubleDividers;
      var divW = strokeW * (dbl ? 0.4 : 0.7);
      dividerYs(def).forEach(function (y0) {
        (dbl ? [y0 - strokeW * 0.55, y0 + strokeW * 0.55] : [y0]).forEach(function (y) {
          var half = halfWidthAt(g, g.content, y - cy) * 0.99;
          s += '<line x1="' + (cx - half).toFixed(2) + '" y1="' + y.toFixed(2) + '" x2="' + (cx + half).toFixed(2) + '" y2="' + y.toFixed(2) +
            '" stroke="' + ink + '" stroke-width="' + divW.toFixed(2) + '"/>';
        });
      });
    }

    var bs = bands(def);
    bs.forEach(function (b) {
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
        return Math.min(2, Math.max(0.4, Number(scales[i]) || 1));
      });
      var totalW = weights.reduce(function (a, w) { return a + w; }, 0) || 1;
      var family = FONT_MAP[row.font] || FONT_MAP.gothic;
      var spacing = Number(row.letterSpacing) || 0;
      var inner = "";
      var y = b.y0;
      lines.forEach(function (line, i) {
        var h = bandH * weights[i] / totalW;
        var lineCy = y + h / 2;
        y += h;
        if (!line.trim()) return;
        // 文字数が多い行は、枠からはみ出さない大きさまで縮める
        var fontSize = fitFontSize(h * 0.68, widthPerEm(line, family, weight, spacing), lineCy, g);
        inner += '<text x="' + cx + '" y="' + (lineCy + fontSize * 0.34).toFixed(2) + '" text-anchor="middle" fill="' + ink +
          '" font-size="' + fontSize.toFixed(2) + '" font-weight="' + weight +
          '" font-family=\'' + family + "'" +
          ' letter-spacing="' + (spacing * fontSize).toFixed(2) + '"' +
          ">" + esc(line) + "</text>";
      });
      if (!inner) return;
      var rowTilt = Number(row.tiltDeg) || 0;
      s += rowTilt
        ? '<g transform="rotate(' + rowTilt + " " + cx + " " + bandCy.toFixed(2) + ')">' + inner + "</g>"
        : inner;
    });

    // マイクロ文字(検証コード): 約0.25mm。options.codeOnFrame のとき、または日付段のない印では
    // 枠の下側に沿わせる(二重枠なら外枠と内枠のすき間、単枠なら枠のすぐ内側)。
    // それ以外は日付段の下端に印字
    if (opts.microtext) {
      var msize = g.S * 0.021;
      var mattrs = ' fill="' + ink + '" font-size="' + msize.toFixed(2) + '" font-family="Consolas, monospace" letter-spacing="' +
        (msize * 0.25).toFixed(2) + '" text-anchor="middle"';
      var dateBand = bs.find(function (b) { return b.row.role === "date"; });
      if ((def.options && def.options.codeOnFrame) || !dateBand) {
        var o = g.at(g.inner ? g.gapT - msize * 0.35 : strokeW + msize * 0.5);
        var arcId = maskId + "_code";
        var d = g.rect
          ? "M" + (cx - o.a + o.r).toFixed(2) + " " + (cy + o.b).toFixed(2) + " L" + (cx + o.a - o.r).toFixed(2) + " " + (cy + o.b).toFixed(2)
          : "M" + (cx - o.a).toFixed(2) + " " + cy + " A" + o.a.toFixed(2) + " " + o.b.toFixed(2) + " 0 0 0 " + (cx + o.a).toFixed(2) + " " + cy;
        s += '<path id="' + arcId + '" fill="none" d="' + d + '"/>';
        s += "<text" + mattrs + '><textPath href="#' + arcId + '" startOffset="50%">' + esc(opts.microtext) + "</textPath></text>";
      } else {
        s += '<text x="' + cx + '" y="' + (dateBand.y1 - msize * 0.3).toFixed(2) + '"' + mattrs + ">" + esc(opts.microtext) + "</text>";
      }
    }

    s += "</g></svg>";
    return s;
  }

  /**
   * SVG → PNG(dataURL)。dpi指定で実寸解像度を決める(既定300dpi)。背景は透過。
   * sizeMm: {w, h}(sizeMm(def) の戻り値)
   */
  function toPNG(svgString, size, dpi) {
    dpi = dpi || 300;
    var pw = Math.ceil(size.w * dpi / 25.4), ph = Math.ceil(size.h * dpi / 25.4);
    return new Promise(function (resolve, reject) {
      var img = new Image();
      img.onload = function () {
        var canvas = document.createElement("canvas");
        canvas.width = pw; canvas.height = ph;
        var ctx = canvas.getContext("2d");
        ctx.drawImage(img, 0, 0, pw, ph);
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
    SHAPES: SHAPES,
    defaultDef: defaultDef,
    sanitizeDef: sanitizeDef,
    sizeMm: sizeMm,
    dateFormatOf: dateFormatOf,
    newSeedHex: newSeedHex,
    formatDate: formatDate,
    deriveWear: deriveWear,
    verificationCode: verificationCode,
    formatCode: formatCode,
    renderSVG: renderSVG,
    toPNG: toPNG
  };
})(window);
