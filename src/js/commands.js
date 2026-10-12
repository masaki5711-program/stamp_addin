/**
 * commands.js — リボンの「押印」ボタン(タスクペインを開かずに押す)
 * 押印タブで最後に選んだ印を、今日の日付で、選択セル(Wordはカーソル位置)に押す。
 */
(function (global) {
  "use strict";

  function stampDefault(event) {
    XL.init()
      .then(function () {
        var def = Store.get(Store.readKey("defaultStamp", null)) || Store.list()[0];
        // 印が未作成のときは何もしない(リボンコマンドからは画面に案内を出せない)
        if (!def) return null;
        return StampAction.run(def, { date: new Date(), randomTilt: true });
      })
      .catch(function (e) { console.error("stampDefault failed", e); })
      .then(function () { event.completed(); });
  }

  global.stampDefault = stampDefault;
  Office.onReady(function () {
    if (Office.actions && Office.actions.associate) Office.actions.associate("stampDefault", stampDefault);
  });
})(window);
