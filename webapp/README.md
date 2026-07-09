# Web App版

コミュニティー配布向けに、中央管理型Web Appへ移行するための雛形です。

## 方針

配布先に `Code.gs` を渡す方式では、利用者がコードを編集できます。
Web App版では、処理本体を管理者側のApps Scriptに置き、利用者はWeb画面だけを開きます。

この雛形では、利用者ごとに以下を管理します。

- 登録メールアドレス
- 利用状態
- 利用期限
- eBay Developer Client ID
- eBay Developer Client Secret
- eBay RuName
- eBay Refresh Token
- Marketplace / Trading Site ID

## 初期設定

1. 管理者用Googleスプレッドシートを作ります。
2. Apps Scriptに `WebApp.gs` と `Index.html` を追加します。
3. `appsscript.json` にこのリポジトリのスコープを反映します。
4. スクリプトプロパティに `WEBAPP_SESSION_SECRET` を設定します。
5. `setupWebAppSheets` を実行します。
6. `Users` シートに利用者メールを登録します。
7. Web Appとしてデプロイします。

## Usersシート

最低限、以下の列を使います。

| email | status | expiresAt | displayName |
| --- | --- | --- | --- |
| user@example.com | ACTIVE | 2026-12-31 | 山田さん |

`status` が `ACTIVE` の利用者だけログインできます。
`expiresAt` は空なら期限なしです。

## 利用者の流れ

1. Web App URLを開きます。
2. 登録メールアドレスを入力します。
3. メールに届いた確認コードでログインします。
4. 自分のeBay Developer情報を入力します。
5. eBay認証URLを作成して許可します。
6. 許可後URLまたは `code` を貼り付け、Refresh Tokenを保存します。

## 重要

この雛形は、まず「ログイン」「利用者別API情報保存」「eBay OAuth」までです。
既存の送料上書き・配送ポリシー変更処理をWeb画面から実行する部分は、次の段階で接続します。

Web App側に処理本体を置くため、配布先にコードを渡す必要がなくなります。
URLが拡散されても、登録メールの受信者でなければログインできません。
