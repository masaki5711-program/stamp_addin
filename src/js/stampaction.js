/**
 * stampaction.js — 押印の実行(タスクペインとリボンの「押印」ボタンの共通処理)
 * 印影の生成 → 文書への挿入 → 押印履歴への記録までを行う。
 */
(function (global) {
  "use strict";

  /**
   * run(def, opts)
   *  opts.date       — 印面に入れる日付(省略時は今日)
   *  opts.randomTilt — 印の設定の範囲でランダムに傾ける
   *  opts.perCell    — Excelで選択した各セルに押す
   * 戻り値(Promise): { meta, places }(places は押した場所の配列。ブラウザモードでは空)
   */
  function run(def, opts) {
    opts = opts || {};
    var when = opts.date || new Date();
    var maxTilt = (def.options && def.options.randomTiltOnStampDeg) || 0;
    var tilt = opts.randomTilt && maxTilt > 0 ? (Math.random() * 2 - 1) * maxTilt : 0;
    var dateText = Stamp.formatDate(Stamp.dateFormatOf(def), when);
    var size = Stamp.sizeMm(def);

    return Stamp.verificationCode(def, dateText)
      .then(function (code) {
        var svg = Stamp.renderSVG(def, { date: when, tiltDeg: tilt, microtext: Stamp.formatCode(code) });
        return Stamp.toPNG(svg, size, 600).then(function (png) { return { png: png, code: code }; });
      })
      .then(function (r) {
        var meta = {
          stampId: def.id,
          name: def.name,
          date: dateText,
          stampedAt: new Date().toISOString(),
          vcode: r.code
        };
        if (!XL.isOffice()) {
          XL.downloadPNG(r.png, def.name + "_" + Holidays.keyOf(when) + ".png");
          return { meta: meta, places: [] };
        }
        return XL.insertStamp(r.png, size, meta, { perCell: opts.perCell }).then(function (places) {
          var doc = XL.docName();
          Store.addHistory(places.map(function (p) {
            return { at: meta.stampedAt, stampId: meta.stampId, name: meta.name, date: meta.date, vcode: meta.vcode, doc: doc, where: p };
          }));
          return { meta: meta, places: places };
        });
      });
  }

  global.StampAction = { run: run };
})(window);
