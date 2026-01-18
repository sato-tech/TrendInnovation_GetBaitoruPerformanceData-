# TrendInnovation GetBaitoru Performance Data

スプレッドシート（Google Sheets / Excel）から検索キーを取得し、Puppeteerで特定メディア管理画面にログインして検索を実行し、結果をCSV出力・集計してスプレッドシートに追記するツールです。

## 機能

- ✅ GoogleスプレッドシートまたはExcelから検索キーを読み取り
- ✅ Puppeteerでメディア管理画面にログイン・検索実行
- ✅ 検索結果をCSVとして出力
- ✅ CSVを集計してスプレッドシートに1行追記
- ✅ リトライ・タイムアウト・ログ機能
- ✅ ヘッドレス/非ヘッドレスモード切替可能

## 必要な環境

- Node.js v20.10.0 以上
- npm 10.2.3 以上
- Windows 10/11

## セットアップ

### 1. リポジトリのクローン

```bash
git clone <repository-url>
cd TrendInnovation_GetBaitoruPerformanceData-
```

### 2. 依存関係のインストール

```bash
npm install
```

### 3. 環境変数の設定

`.env.example`をコピーして`.env`を作成し、必要な値を設定してください。

```bash
# Windows PowerShell
Copy-Item .env.example .env
```

### 4. Google Sheets API認証（Google Sheetsを使用する場合）

1. [Google Cloud Console](https://console.cloud.google.com/)でプロジェクトを作成
2. Google Sheets APIを有効化
3. OAuth 2.0認証情報を作成（デスクトップアプリケーション）
4. 認証情報をダウンロードして`credentials.json`としてプロジェクトルートに配置
5. 初回実行時にブラウザが開き、認証を完了すると`token.json`が自動生成されます

### 5. ビルド

```bash
npm run build
```

## 環境変数

| 変数名 | 説明 | 必須 | デフォルト |
|--------|------|------|-----------|
| `NODE_ENV` | 実行環境 | - | `development` |
| `HEADLESS` | ヘッドレスモード | - | `true` |
| `SPREADSHEET_TYPE` | スプレッドシート種別 (`google` / `excel`) | ✅ | - |
| `SPREADSHEET_ID` | Google Sheets ID (googleの場合) | ⚠️ | - |
| `SPREADSHEET_FILE_PATH` | Excelファイルパス (excelの場合) | ⚠️ | - |
| `SPREADSHEET_RANGE` | 読み取り範囲 | - | `Sheet1!A2:B` |
| `SEARCH_KEY_COLUMN` | 検索キー列 | - | `A` |
| `OUTPUT_SHEET_ID` | 出力先Google Sheets ID | - | - |
| `OUTPUT_SHEET_RANGE` | 出力先範囲 | - | `Sheet1!A1` |
| `GOOGLE_CREDENTIALS_PATH` | Google認証情報パス | ⚠️ | `./credentials.json` |
| `GOOGLE_TOKEN_PATH` | Googleトークンパス | ⚠️ | `./token.json` |
| `LOGIN_URL` | ログインURL | ✅ | - |
| `LOGIN_USERNAME` | ログインユーザー名 | ✅ | - |
| `LOGIN_PASSWORD` | ログインパスワード | ✅ | - |
| `SEARCH_TIMEOUT` | 検索タイムアウト(ms) | - | `30000` |
| `PAGE_LOAD_TIMEOUT` | ページ読み込みタイムアウト(ms) | - | `60000` |
| `MAX_RETRIES` | 最大リトライ回数 | - | `3` |
| `RETRY_DELAY` | リトライ間隔(ms) | - | `2000` |
| `LOG_LEVEL` | ログレベル | - | `info` |
| `LOG_FILE_PATH` | ログファイルパス | - | `./logs/app.log` |

⚠️: 条件付き必須（`SPREADSHEET_TYPE`に応じて）

## 使用方法

### 開発モード（TypeScriptを直接実行）

```bash
npm run dev
```

### 本番モード（ビルド後に実行）

```bash
npm run build
npm start
```

## 実行フロー

1. **検索キー取得**: スプレッドシート（Google Sheets / Excel）から検索キーを読み取り
2. **ブラウザ起動**: Puppeteerでブラウザを起動（ヘッドレス/非ヘッドレス切替可能）
3. **ログイン**: 指定されたURLにログイン
4. **検索実行**: 各検索キーで検索を実行
5. **CSV出力**: 検索結果を`output/`ディレクトリにCSVとして出力
6. **集計**: CSVを集計ルールに従って集計
7. **スプレッドシート追記**: 集計結果をスプレッドシートに1行追記

## ディレクトリ構成

```
TrendInnovation_GetBaitoruPerformanceData-/
├── src/
│   ├── config/              # 設定管理
│   ├── services/
│   │   ├── spreadsheet/     # スプレッドシート読み取り
│   │   ├── scraper/         # スクレイピング処理
│   │   ├── csv/             # CSV出力・集計
│   │   └── retry/            # リトライ処理
│   ├── utils/               # ユーティリティ
│   └── index.ts             # エントリーポイント
├── output/                  # CSV出力先
├── logs/                    # ログ出力先
├── .env                     # 環境変数（.gitignore）
├── credentials.json         # Google認証情報（.gitignore）
├── token.json               # Googleトークン（.gitignore）
└── README.md
```

## カスタマイズ

### ログイン処理のカスタマイズ

`src/services/scraper/login.ts`のセレクタを実際のサイトに合わせて調整してください。

### 検索処理のカスタマイズ

`src/services/scraper/search.ts`のセレクタとデータ取得ロジックを実際のサイトに合わせて調整してください。

### 集計ルールのカスタマイズ

`src/services/csv/aggregator.ts`の`aggregateData`メソッドを実際の要件に合わせて調整してください。

## トラブルシューティング

### ログインに失敗する

- `LOGIN_URL`、`LOGIN_USERNAME`、`LOGIN_PASSWORD`が正しく設定されているか確認
- `HEADLESS=false`に設定してブラウザの動作を確認
- ログファイル（`logs/app.log`）を確認

### 検索結果が取得できない

- `SEARCH_TIMEOUT`を増やす
- `src/services/scraper/search.ts`のセレクタを実際のサイトに合わせて調整
- `HEADLESS=false`に設定してブラウザの動作を確認

### Google Sheets API認証エラー

- `credentials.json`が正しく配置されているか確認
- `token.json`を削除して再認証を試す
- Google Cloud ConsoleでAPIが有効になっているか確認

## ライセンス

MIT
