# Attendance Summary with Forms

Microsoft 365に含まれるFormsから得られた複数のエクセルファイルを「出席簿」として集約するためのWebアプリケーションです。  
オンライン講義などでFormsを用いて受講確認を行う場合などを想定して作成しました。

[demo](https://attendance-summary-with-forms.netlify.app)

「メール」列を軸として，「開始時刻」と「合計点数」から，それぞれ「出席簿」と「得点総計」シートを作成します。  
そのため，受講生はMicrosoft 365にログインしていることが必須となります。

## プライバシー

すべての作業は端末内のみで行われ，他のサーバーなどには一切の情報を送信しません。

## 開発環境など

Angularなどを使って作成されています。

## xlsx（SheetJS）の CDN 依存について

`package.json` の `xlsx` は、npm レジストリではなく SheetJS の公式 CDN（https://cdn.sheetjs.com）の tarball URL を直接指定しています。SheetJS は 0.18.6 以降を npm レジストリに公開しておらず、公式配布を自社 CDN に移行したためです。現在は CDN 上の最新である 0.20.3 を指定しています。

npm レジストリ上の `xlsx` は 0.18.5 で止まっていますが、このバージョンには次の脆弱性が未修正のまま残っています。

- CVE-2023-30533 / GHSA-4r6h-8v6p-xvw6: Prototype Pollution（0.19.3 で修正）
- CVE-2024-22363 / GHSA-5pgg-2g8v-p4x9: ReDoS（0.20.2 で修正）

そのため、npm レジストリ版に戻すことは、0.18.5 へのダウングレードと脆弱性の導入になるので避けてください。

npm 12 以降は `allow-remote` の既定値が `none` になっています。CDN の URL を依存関係として利用するため、リポジトリの `.npmrc` に `allow-remote=root` を設定しています。これがない場合、`npm install` と `npm ci`（Netlify のビルドを含む）は `EALLOWREMOTE` で失敗します。

バージョンを上げるときは、まず https://cdn.sheetjs.com で新版の公開を確認し、`package.json` の `xlsx` の URL を書き換えてから `npm install` を実行してください。0.20.4、0.20.5、0.21.0、0.21.1、0.22.0 は CDN 上で 404 になることを確認済みです。

CDN から直接インストールしたパッケージも `npm audit` の検査対象です。実際に別ディレクトリで CDN から脆弱な 0.19.2 をインストールして `npm audit` を実行したところ、`1 high severity vulnerability` として検出されました。また、`package-lock.json` には `integrity: sha512-oLDq3jw7...` が記録されており、改竄検知が有効です。

なお、cdn.sheetjs.com が停止している場合は `npm ci` が失敗します。この場合、npm レジストリのミラーは利用できません。このリスクは現状受容しています。
