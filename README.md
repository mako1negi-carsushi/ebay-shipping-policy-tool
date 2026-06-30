# eBay Shipping Policy Tool for Google Apps Script

Google SheetsからeBay Inventory APIを呼び出し、既存出品の送料上書きと配送ポリシー変更を行うGoogle Apps Scriptです。

## できること

- 既存出品のSKUからOfferを取得
- 商品ごとに `shippingCostOverrides` を更新
- `overrideShippingCostUSD` が空なら `priceUSD * dutyRate` で送料上書き額を計算
- eBayの配送・支払い・返品ポリシーID一覧を取得
- 現在の配送ポリシーIDが一致する出品を検索し、指定の配送ポリシーへ一括変更
- OAuth認証URLの作成とRefresh Token保存をGoogle Sheetsメニューから実行

## ファイル

- `Code.gs`: Apps Script本体
- `appsscript.json`: Apps Scriptマニフェスト

## セットアップ

1. Google Sheetsを作成します。
2. `拡張機能 > Apps Script` を開きます。
3. `Code.gs` にこのリポジトリの `Code.gs` を貼り付けて保存します。
4. Apps Scriptの `プロジェクトの設定` で `appsscript.json` を表示する設定をONにします。
5. `appsscript.json` にこのリポジトリの `appsscript.json` を貼り付けて保存します。
6. Apps Scriptの `プロジェクトの設定 > スクリプト プロパティ` に必要な値を設定します。

## スクリプトプロパティ

最初に必要な値:

| Key | Value |
| --- | --- |
| `ENVIRONMENT` | `PRODUCTION` |
| `EBAY_CLIENT_ID` | eBay DeveloperのClient ID |
| `EBAY_CLIENT_SECRET` | eBay DeveloperのClient Secret |
| `EBAY_RUNAME` | eBay DeveloperのRuName |
| `MARKETPLACE_ID` | `EBAY_US` |
| `CURRENCY` | `USD` |
| `CONTENT_LANGUAGE` | `en-US` |

OAuth後に自動保存される値:

| Key | Value |
| --- | --- |
| `EBAY_REFRESH_TOKEN` | OAuthメニューで自動保存 |

任意のデフォルト値:

| Key | Value |
| --- | --- |
| `FULFILLMENT_POLICY_ID` | デフォルト配送ポリシーID |
| `PAYMENT_POLICY_ID` | デフォルト支払いポリシーID |
| `RETURN_POLICY_ID` | デフォルト返品ポリシーID |

秘密情報はGitHubにコミットしないでください。`EBAY_CLIENT_SECRET` と `EBAY_REFRESH_TOKEN` はスクリプトプロパティだけに保存します。

## OAuth

Google Sheetsを再読み込みすると `eBay出品` メニューが表示されます。

1. `eBay出品 > OAuth: 認証URLを表示`
2. 表示されたURLをブラウザで開き、eBayにログインして許可します。
3. 許可後URLに含まれる `code=...` 付きURL全体をコピーします。
4. `eBay出品 > OAuth: codeからRefresh Tokenを保存`
5. コピーしたURLを貼り付けます。

成功すると `EBAY_REFRESH_TOKEN` がスクリプトプロパティに保存されます。

## ポリシーID一覧取得

1. `eBay出品 > ポリシーID一覧を取得`
2. `Policies` シートに以下が出力されます。

| type | 内容 |
| --- | --- |
| `FULFILLMENT` | 配送ポリシー |
| `PAYMENT` | 支払いポリシー |
| `RETURN` | 返品ポリシー |

`Listings` シートの `fulfillmentPolicyId` 列では、取得した配送ポリシーIDを商品ごとに指定できます。

## 既存出品の送料を更新

1. `eBay出品 > シートを初期化`
2. `Listings` シートに既存商品の `sku` を入力します。
3. 送料を直接指定する場合は `overrideShippingCostUSD` に金額を入れます。
4. 自動計算する場合は `priceUSD` と `dutyRate` を入れ、`overrideShippingCostUSD` は空にします。
5. まず `eBay出品 > 既存出品: 選択行をドライラン` を実行します。
6. `requestPreview` を確認します。
7. 問題なければ `eBay出品 > 既存出品: 選択行の送料を更新` を実行します。

## 配送ポリシーを一括置換

現在Aの配送ポリシーで登録されている出品を、指定Bの配送ポリシーへ変更できます。

1. `eBay出品 > ポリシーID一覧を取得`
2. `Policies` シートで置換元と変更先の `FULFILLMENT` の `policyId` を確認します。
3. `eBay出品 > 一括置換: 対象出品を検索`
4. 置換元の配送ポリシーIDと変更先の配送ポリシーIDを入力します。
5. `BulkPolicyChange` シートに候補が作成されます。
6. 更新したい行だけ `approve` をTRUEにします。
7. `eBay出品 > 一括置換: approve=TRUEを更新`

いきなり全件更新しないよう、候補確認と承認の2段階にしています。

## 注意

- 最初は必ず1件だけでテストしてください。
- 本番出品を操作する場合は `ENVIRONMENT=PRODUCTION` を使います。
- Sandboxの認証情報では本番出品は操作できません。
- Seller Hubや古い方式で作成した出品がInventory API側で取得できない場合があります。
- APIエラーは各シートの `lastError` に出力されます。

