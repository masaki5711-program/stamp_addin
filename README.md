# Excel/Word 便利アドインセット

タスクペインのホーム画面から選べる7ツール構成:

| ツール | 対応 | 内容 |
|---|---|---|
| デート印 | Excel/Word | 日付入りの電子日付印。丸・楕円・小判・角の枠、二重線、日付なしの認印、押印日の指定、複数セルへの一括押印、承認欄の作成、押印履歴、検証コード |
| お作法リセット | Excel | 全シートA1選択・枠固定解除・先頭シート表示・保存 |
| シート差分比較 | Excel | 2シートの差分(値/数式)を色付け+一覧化。行の挿入・削除を考慮した対応付け、結果のシート出力、別ブックのシート取り込み |
| 定型書式 | Excel | 選択範囲を見出し・縞・罫線つきの表にワンクリック整形(プリセット保存可) |
| 定型文 | Excel/Word | 文例ライブラリからセル/カーソルへ挿入。`{今日}` `{和暦}` `{曜日}` `{時刻}` `{氏名}` の差し込み |
| ブック掃除 | Excel | 名前定義・非表示シート・カスタムスタイル・外部リンク・コメント・作成者情報・書式だけ残った範囲の棚卸し |
| 日付ツール | Excel/Word | 和暦変換・各種書式挿入・営業日計算(祝日・会社休日に対応) |

リボンには「便利ツール」(タスクペインを開く)と「押印」(最後に選んだ印をその場で押す)の2つのボタンが出る。

仕様検討書: https://claude.ai/code/artifact/60062165-f43d-41ae-a519-d9508bf13b58

## 構成

```
manifest.xml                      開発用マニフェスト(https://localhost:3000)
deploy/manifest.prod.template.xml 配布用マニフェストのテンプレート
src/
  taskpane.html                   タスクペイン(ホーム+7ツール)
  commands.html                   リボンの「押印」ボタン用(画面なし)
  css/taskpane.css
  js/stamp.js                     印影の描画・欠け生成・日付書式・定義の検査(中核ロジック)
  js/storage.js                   保存(localStorage)・押印履歴・一括バックアップ
  js/holidays.js                  日本の祝日(計算)と会社休日
  js/excel.js                     Excel/Wordへの画像挿入・承認欄・押印メタデータの収集
  js/stampaction.js               押印の実行(タスクペインとリボンの共通処理)
  js/shell.js                     ホーム画面・ツール切替・共通UI
  js/app.js                       デート印のUIロジック
  js/commands.js                  リボンの「押印」ボタンの処理
  js/tools/*.js                   各ツール
  assets/icon-*.png               リボン用アイコン
```

Officeのタスクペインでは `confirm()` / `alert()` が無効、ファイルのダウンロードも環境によっては始まらない。
そのため確認は「同じボタンをもう一度押す」(`UI.confirmClick`)、書き出しは「テキスト欄に表示してコピー」も併用している。

## 動作要件

- **Microsoft 365 の Excel**(Windows / Mac / Web)。図形挿入に ExcelApi 1.9 以上が必要
- 買い切り版(Excel 2016/2019/2021)は対象外
- 開発には Node.js(npm)

## 開発・動作確認

```powershell
npm install
npm run certs        # 初回のみ: localhost用HTTPS証明書(信頼ダイアログでOKを押す)
npm start            # https://localhost:3000 で配信
```

- **Excelなしで試す**: ブラウザで `https://localhost:3000/taskpane.html` を開く。
  デザイナ・ライブラリは全機能動作し、押印はPNGダウンロードになる
  (証明書を入れない場合は `npm run start:http` で `http://localhost:3000` でも可)。
- **Excelに読み込む(サイドロード)**: `npm start` したまま別ターミナルで:

  ```powershell
  npm run sideload     # Excelが起動し、ホームタブ右端に「デート印」ボタンが出る
  npm run sideload:stop
  ```

  Excel on the web の場合: ブック → 挿入 → アドイン → 自分のアドイン →
  「マニフェストをアップロード」で `manifest.xml` を指定。

## 他の人への配布

Officeアドインの実体は **Webサーバー上のファイル+マニフェストXML** の2点。
利用者に渡すのはマニフェストだけで、本体を更新すればすぐ全員に反映される。

手順:

1. **ホスティング**: `src/` の中身を HTTPS の静的ホスティングに置く
   (GitHub Pages / Azure Static Web Apps / Netlify など無料枠で十分)。
   このリポジトリは `main` への push で GitHub Pages に自動デプロイされる(`.github/workflows/pages.yml`)。
   その際、JS/CSS の読み込みURLにコミットIDを付け(`?v=dev` を置換)、公開URL向けの
   `manifest.prod.xml` と `stampset-setup.zip` も生成して一緒に公開する。
2. **マニフェスト作成**: `deploy/manifest.prod.template.xml` の `{{HOST_URL}}` を
   ホスティングURLに置換して `manifest.prod.xml` を作る。

   ```powershell
   (Get-Content deploy/manifest.prod.template.xml -Raw) -replace '\{\{HOST_URL\}\}','https://example.github.io/addin-excel' |
     Set-Content deploy/manifest.prod.xml
   npx office-addin-manifest validate deploy/manifest.prod.xml
   ```

3. **簡易インストーラーで配る場合**(Windows向け・推奨):

   ```powershell
   .\deploy\make-package.ps1 -HostUrl "https://<ユーザー名>.github.io/<リポジトリ名>"
   ```

   `deploy/stampset-setup.zip` ができるので、これを相手に渡す。
   相手は「zipを展開 → `install.cmd` をダブルクリック → Excel再起動」だけで
   ホームタブに「デート印」ボタンが出る(管理者権限不要。公式のレジストリ
   サイドロード方式を自動化したもの)。`uninstall.cmd` で削除できる。
   zipには受け取った人向けの「はじめにお読みください.txt」を同梱済み。

4. **その他の配布経路**(規模に応じて選択):

   | 経路 | 対象 | 方法 |
   |---|---|---|
   | マニフェスト個別配布 | 少人数・知人 | `manifest.prod.xml` を渡し、各自が Excel on the web「マニフェストをアップロード」or Windowsの共有フォルダカタログで登録 |
   | 集中展開(推奨) | 組織(Microsoft 365) | 管理センター →「統合アプリ」にマニフェストをアップロードし、対象ユーザーへ一括配布 |
   | AppSource | 一般公開 | パートナーセンター登録+審査。将来検討 |

## 既知の制限

- フォントは端末のシステムフォント(游明朝/游ゴシック)。押印時に画像化するので、押した後の見た目は環境に依存しない
- 保存先は localStorage(このPCのOffice内)。キャッシュ削除やPC入れ替えで消えるため、
  デート印の「ライブラリ」→「すべて書き出す」でバックアップする
- 押印の「検証」は、押印に記録された日付・検証コードと印の定義の突き合わせ。印影の画像そのものの書き換えは検出できない
- 日付ソースはローカル時計。押印日は利用者が指定できる(履歴には実際の押印日時も残る)
- 祝日は現行の祝日法から計算する。春分・秋分は計算による予測で、法改正や臨時の休日には追従しない
- リボンの「押印」ボタンを追加・変更した版では、利用者側でマニフェストの入れ直し(再インストール)が必要

## セキュリティに関する注意

印影の「欠け」は個体識別の演出であり、**偽造防止機能ではない**(画像コピーで複製可能)。
真正性の確認は、押印画像の代替テキストに記録される押印メタデータ
(印ID・押印日時)との突き合わせを正とする。電子印影自体に法的効力はない。
また、印定義JSONを共有すると欠けごと複製できるため、定義ファイルは配布しないこと。
