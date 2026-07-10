# Web App版

コミュニティー配布向けの中央管理型Web Appです。
処理本体を管理者側のApps Scriptに置き、利用者はWeb画面だけを開きます。
コードは利用者に渡らないため、コピー・改変されません。

## できること(利用者側の画面)

- メールアドレス + 確認コードでログイン(Usersシートに登録された人だけ)
- 自分のeBayアカウントをOAuthで接続(Refresh Tokenは利用者ごとにUsersシートへ保存)
- ポリシーID一覧の表示(配送・支払い・返品)
- 配送ポリシー一括変更(Trading API): 置換元ポリシーの出品を全件検索 → チェックした出品だけ変更先ポリシーへ変更 + 送料上書き(価格 × 関税率)
- 個別更新(Item ID指定 / Trading API): ドライラン確認 → 更新の2段階
- 同一商品2個目以降の追加送料も同時に設定(送料 × 割合%。画面の入力欄で変更可、デフォルト75%)

更新はすべてTrading API(ReviseFixedPriceItem)で行います。
Inventory API系のエンドポイント(SKU更新・移行)はサーバー側(WebApp.gs)に残っていますが、画面には出していません。

## ファイル

- `WebApp.gs`: ログイン・セッション・利用者管理 + Web画面から呼ばれるエンドポイント
- `EbayApi.gs`: 利用者ごとのeBay API処理層(Code.gsのロジックを利用者別設定で動くよう移植)
- `Index.html`: 利用者向けのタブ式操作画面(日本語)

## 初期設定(管理者)

1. 管理者用Googleスプレッドシートを作ります。
2. `拡張機能 > Apps Script` を開き、`WebApp.gs`、`EbayApi.gs`、`Index.html` の3ファイルを追加します。
3. `appsscript.json` にこのリポジトリのスコープを反映します。
4. スクリプトプロパティに以下を設定します。
5. `setupWebAppSheets` を1回実行します(Usersシートが作られます)。
6. `Users` シートに利用者メールを登録します。
7. `デプロイ > 新しいデプロイ > ウェブアプリ` としてデプロイします。
   - 次のユーザーとして実行: 自分
   - アクセスできるユーザー: 全員

## スクリプトプロパティ

| Key | 内容 |
| --- | --- |
| `WEBAPP_SESSION_SECRET` | ログイン署名用の長いランダム文字列 |
| `ENVIRONMENT` | `PRODUCTION`。Sandboxなら `SANDBOX` |
| `EBAY_CLIENT_ID` | 管理者側eBay Developer AppのClient ID |
| `EBAY_CLIENT_SECRET` | 管理者側eBay Developer AppのClient Secret |
| `EBAY_RUNAME` | 管理者側eBay Developer AppのRuName |

eBay Developer情報は管理者側だけが持ちます。利用者には入力させません。

## Usersシート

`setupWebAppSheets` 実行で全列が作られます。管理者が手で入力するのは以下だけです。

| email | status | expiresAt | displayName |
| --- | --- | --- | --- |
| user@example.com | ACTIVE | 2026-12-31 | 山田さん |

- `status` が `ACTIVE` の利用者だけログインできます。止めたい場合は別の値(例: `STOPPED`)にします。
- `expiresAt` は空なら期限なしです。
- `ebayRefreshToken` などの他の列は、利用者がeBay接続したときに自動で埋まります。

## 利用者の流れ

1. Web App URLを開きます。
2. 登録メールアドレスを入力し、メールに届いた確認コードでログインします。
3. 「eBay接続」タブでMarketplaceを確認し、eBay認証URLを開いて許可 → 許可後のURLを貼り付けて接続完了。
4. あとは「一括変更」「個別更新」「SKU更新・移行」タブから操作します。
5. どの更新も「確認(ドライラン)→ 更新」の2段階です。

## 一括変更の仕組み

- 出品の検索はブラウザから約40秒ずつの小分けで自動的に繰り返されます(Apps Scriptの実行時間制限対策)。検索中はブラウザのタブを開いたままにしてください。
- 更新は10件ずつまとめて実行され、行ごとに成功/エラーが表示されます。

## 規模の目安

- 確認コードのメール送信は、無料のGoogleアカウントで1日100通までです(利用者50人なら十分)。
- 同時に大量の利用者が一括検索を行うと、Apps Scriptの同時実行数上限(約30)に当たる場合があります。その場合は時間をずらして使ってもらってください。
