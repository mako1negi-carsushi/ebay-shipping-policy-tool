# License Gate

コミュニティー生だけに配布するための、メールアドレス許可制のひな形です。

## 重要

`Code.gs` をそのまま配布すると、受け取った人はApps Scriptの中身を編集できます。
そのため、クライアント側だけのメールチェックは完全なコピー防止にはなりません。

現実的な運用は次の2段階です。

1. 簡易版: 配布用スプレッドシートにライセンス確認を入れる
2. 強化版: eBay API実行部分を管理者側Web Appに寄せ、配布先には操作画面だけ渡す

このフォルダは、まず簡易版を作るためのテンプレートです。

## 管理者側

1. 管理者用Googleスプレッドシートを作ります。
2. Apps Scriptに `LicenseServer.gs` を貼ります。
3. `setupAllowlistSheet` を実行します。
4. `Allowlist` シートに利用許可するメールアドレスを登録します。
5. スクリプトプロパティに `LICENSE_SERVER_SECRET` を設定します。
6. Web Appとしてデプロイします。

`Allowlist` シート:

| email | status | expiresAt | toolId | displayName | notes |
| --- | --- | --- | --- | --- | --- |
| user@example.com | ACTIVE | 2026-12-31 | ebay-shipping-policy-tool | 山田さん |  |

`status` は `ACTIVE` の人だけ利用できます。
`expiresAt` は空なら期限なしです。

## 配布先

配布用の `Code.gs` に `ClientLicenseSnippet.gs` の内容を追加します。

スクリプトプロパティ:

| Key | Value |
| --- | --- |
| `LICENSE_SERVER_URL` | 管理者側Web App URL |
| `LICENSE_SERVER_TOKEN` | 管理者側の `LICENSE_SERVER_SECRET` と同じ値 |
| `LICENSED_EMAIL` | 利用者メール。メニューから保存してもよい |

各メニュー処理の先頭で `requireLicensedUser_();` を呼ぶと、許可されていない人は実行できません。

例:

```js
function prepareBulkTradingPolicyChange() {
  requireLicensedUser_();
  // existing code...
}
```

## 注意

この簡易版では、配布されたApps Scriptを編集できる人ならチェックを外せます。
拡散対策を本気でやる場合は、eBay APIを呼ぶ処理を管理者側Web Appに置く構成にしてください。
