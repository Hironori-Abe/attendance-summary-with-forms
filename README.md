# Attendance Summary with Forms

Microsoft 365に含まれるFormsから得られた複数のエクセルファイルを「出席簿」として集約するためのWebアプリケーションです。  
オンライン講義などでFormsを用いて受講確認を行う場合などを想定して作成しました。

[demo](https://attendance-summary-with-forms.netlify.app)

「メール」列を軸として，「開始時刻」と「合計点数」から，それぞれ「出席簿」と「得点総計」シートを作成します。  
そのため，受講生はMicrosoft 365にログインしていることが必須となります。

## プライバシー

すべての作業は端末内のみで行われ，他のサーバーなどには一切の情報を送信しません。

## 開発環境など

Angular 本体・CLI・Material・CDK は `22.2.3` に揃えています。Node.js は Angular 22 の要件を満たす版（22系なら `22.22.3` 以降）を使用してください。

## テスト

単体テスト18件は Angular CLI の `@angular/build:unit-test` と Vitest で実行します。横幅のはみ出しやスクロール位置の検証を維持するため、Playwright の Chromium をヘッドレスで使用しています。

初回は依存関係とテスト用ブラウザーをインストールしてください。

```sh
npm ci
npx playwright install chromium
npm test
```

既存の Chrome / Chromium を使う場合は、代わりに `CHROME_BIN` を指定できます。

```sh
CHROME_BIN=/usr/bin/chromium npm test
```

カバレッジは `npm test -- --coverage` で取得でき、`coverage/` に出力されます。本番ビルドは `npm run build` で確認できます。

Karma・Jasmine の依存関係と旧設定は削除し、TestBed の初期化は Angular CLI に任せています。テスト用コンポーネントも事前コンパイルできるよう、依存する Material モジュールを明示しています。

## 開発用依存関係のセキュリティ更新

2026-10-22 に Dependabot の警告16件に対応し、`@angular/cli` と `@angular/build` を `22.2.3` に更新しました。対象はいずれも開発用の間接依存です。

| パッケージ | Dependabot の警告番号 | 対応後 |
| --- | --- | --- |
| `piscina` | #153 | `5.3.2`（[修正情報](https://github.com/advisories/GHSA-67c8-pqhq-4rmx)） |
| `source-map-js` | #155 | `1.2.2`（[修正情報](https://github.com/advisories/GHSA-68fv-2mgg-jv7q)） |
| `follow-redirects` | #158, #159 | Karma の削除により依存関係から除去 |
| `brace-expansion` | #156 | Karma 関連パッケージの削除により依存関係から除去 |
| `@modelcontextprotocol/sdk` | #154, #161, #162, #163 | 依存関係から除去。CLI は `@modelcontextprotocol/server@2.3.1` を利用 |
| `hono` | #145, #146, #152, #160 | CLI の更新により依存関係から除去 |
| `ip-address` | #148, #149, #151 | CLI の更新により依存関係から除去 |

旧 MCP SDK 用の override は不要になったため削除しています。一方、`@angular/build@22.2.3` が固定する `vite@8.3.0` にも脆弱性があるため、`overrides` でその配下の Vite を `8.3.5` に固定しています。Angular のビルドツールを更新する際は、override の要否を確認してください。

追加監査で検出した `engine.io` と `qs` も、Karma の削除により依存関係から除去しました。依存関係の再現には、更新した `package-lock.json` と `npm ci` を使用してください。

以前残っていた `braces@3.0.3` 由来の警告6件（High）も、Karma から Vitest への移行によって解消しています。2026-10-22 の更新後の `npm audit` は、開発用依存を含めて0件です。

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
