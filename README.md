# Gmailアドオン：メール内容をもとにGoogleカレンダーに予定作成

このスクリプトは **Gmail の任意のメールを開いた状態から、その内容を使って Google カレンダーに予定を作成**する Gmail アドオンです。
**iPhone の Gmail**でも動作し、**作成先カレンダーの選択・開始/終了の分単位指定・説明への本文貼り付け**ができます。作成後は完了カードを表示します。

## できること

* メールを開いた画面下部のアドオンから、**そのメールを元に予定を作成**
* **カレンダー選択**（自分のメイン／所有カレンダーから選択）
* **開始・終了を分単位で指定**
  * メール本文から開始・終了日時を自動補完（Open AI APIを使用）
  * 補完できない場合は今の日時から算出
* **説明欄に本文の先頭とメールのパーマリンク**を自動挿入
* 作成後は **「予定作成完了」カード**を表示

> 補足：Gmail アドオンは仕様上、ペインをプログラムで完全に閉じることはできません。完了カードの表示／ナビゲーションで「閉じた風」の体験を作ります。

## 必要な権限（Scopes）

`appsscript.json` には少なくとも以下のスコープが必要です。

```json
"oauthScopes": [
  "https://www.googleapis.com/auth/gmail.addons.execute",
  "https://www.googleapis.com/auth/gmail.addons.current.message.readonly",
  "https://www.googleapis.com/auth/calendar",
  "https://www.googleapis.com/auth/script.external_request"
]
```

* `gmail.addons.current.message.readonly` … 開いている **その** メールの件名・本文を読むため
* `gmail.addons.execute` … アドオン実行の基本権限
* `calendar` … 予定の作成
* `script.external_request` … Open AI APIにHTTPリクエストを発行するため

## Open AI APIキーの設定

* Google App Scriptの「プロジェクトの設定 > スクリプトプロパティ」に、`OPENAI_API_KEY`を定義する

## 開発環境について

このプロジェクトは、Google Apps Script用のモダンな開発環境を提供する[App Script in IDE (ASIDE)](https://github.com/google/aside/tree/main)を使用しています。

## 使い方（操作フロー）

1. Gmail で予定化したいメールを開く
2. アドオンを起動 → **カレンダー選択**／**タイトル**／**開始**／**終了**／**説明** を確認
3. **［カレンダーに予定を作成］** を押す
4. **「予定作成完了」カード**が出たら終了（Google カレンダーに登録済み）

## TODO

* [ ] IOS版Gmail上で、ボタンを複数回押下できないようにする（多重登録防止）
  * GASでUI制御する方法がよくわからない。。。
  * Windows11, Edge上のGmailならローディング処理が勝手に走る。
* [ ] ユニットテスト実装

## 備考

* 本アドオンはあなたの Google アカウント範囲で動作します。第三者への公開は Google Workspace Marketplace のポリシーに従ってください。
* メール本文はアドオン実行時に一時的に読み取り、予定説明へ挿入します。サーバ等へ保存はしません。
