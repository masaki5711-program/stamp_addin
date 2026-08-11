# Excel/Word 便利アドインセット

タスクペインのホーム画面から選べる7ツール構成:

| ツール | 対応 | 内容 |
|---|---|---|
| デート印 | Excel/Word | 日付自動の電子日付印(欠け個体差・マイクロ文字検証コード付き) |
| お作法リセット | Excel | 全シートA1選択・枠固定解除・先頭シート表示・保存 |
| シート差分比較 | Excel | 2シートの値差分を色付け+一覧化、クリックでジャンプ |
| 定型書式 | Excel | 選択範囲を見出し・縞・罫線つきの表にワンクリック整形(プリセット保存可) |
| 定型文 | Excel/Word | 文例ライブラリからセル/カーソルへ挿入 |
| ブック掃除 | Excel | 名前定義・非表示シート・カスタムスタイルの棚卸しと削除 |
| 日付ツール | Excel/Word | 和暦変換・各種書式挿入・営業日計算 |

仕様検討書: https://claude.ai/code/artifact/60062165-f43d-41ae-a519-d9508bf13b58

## 構成

```
manifest.xml                     開発用マニフェスト(https://localhost:3000)
deploy/manifest.prod.template.xml 配布用マニフェストのテンプレート
src/
  taskpane.html                  タスクペイン(押印/デザイナ/ライブラリの3タブ)
  css/taskpane.css
  js/stamp.js                    印影の描画・欠け生成・日付書式(中核ロジック)
  js/storage.js                  印ライブラリ保存(プロトタイプはlocalStorage)
  js/excel.js                    Excelへの画像挿入(Excel外はPNGダウンロード)
  js/app.js                      UIロジック
  assets/icon-*.png              リボン用アイコン
```

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

## プロトタイプの既知の制限(本番で対応)

- フォントは端末のシステムフォント(游明朝/游ゴシック)。本番はオープンソース日本語フォントを同梱し、環境差をなくす
- 保存先が localStorage(本番は `OfficeRuntime.storage` + エクスポート運用)
- 枠形状は正円のみ(楕円・角丸四角は未実装)。かすれ表現も未実装
- 日付ソースはローカル時計。時刻の信頼性が必要な場合はサーバー時刻の導入を検討

## セキュリティに関する注意

印影の「欠け」は個体識別の演出であり、**偽造防止機能ではない**(画像コピーで複製可能)。
真正性の確認は、押印画像の代替テキストに記録される押印メタデータ
(印ID・押印日時)との突き合わせを正とする。電子印影自体に法的効力はない。
また、印定義JSONを共有すると欠けごと複製できるため、定義ファイルは配布しないこと。
