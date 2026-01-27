# バイトル求人データ自動収集・集計システム

反響事例の取説に基づいた業務フローを完全自動化するNode.js + Puppeteerシステムです。

## 📋 概要

このシステムは、バイトル企業データから求人情報を自動収集し、トレンドデータベースに集計するための自動化ツールです。

## 🛠️ 技術スタック

- **Node.js** (v18以上)
- **Puppeteer** - ブラウザ自動化
- **ExcelJS** - Excelファイル操作
- **csv-parse** - CSVファイル解析
- **Google Sheets API** - Googleスプレッドシートへのデータ書き込み
- **OpenAI API** - AI連携（ドロップダウン選択肢のマッチング）
- **dotenv** - 環境変数管理

## 📁 プロジェクト構造

```
.
├── config/                 # 設定ファイル
│   ├── config.js          # メイン設定ファイル
│   ├── selectors.json     # CSSセレクタ定義
│   ├── excelColumns.json  # Excel列定義
│   ├── jobCategoriesNight.json    # ナイト案件職種カテゴリリスト
│   └── jobCategoriesNormal.json    # 通常案件職種カテゴリリスト
├── src/
│   ├── services/          # ビジネスロジック
│   │   ├── ScrapingService.js      # スクレイピング処理
│   │   ├── ExcelService.js         # Excel操作処理
│   │   ├── GoogleSheetsService.js  # Google Sheets操作処理
│   │   └── AIService.js            # AI連携処理
│   ├── utils/             # ユーティリティ関数
│   │   └── dateUtils.js   # 日付処理
│   ├── index.js           # メインエントリーポイント
│   ├── convert-excel-to-json.js    # Excel→JSON変換ツール
│   ├── debug-selectors.js          # セレクタデバッグツール
│   ├── debug-selectors-by-url.js   # URL別セレクタデバッグツール
│   └── test-*.js          # テストファイル
├── downloads/             # ダウンロードファイル保存先
├── credentials.json       # Googleサービスアカウントキー（要作成）
├── .env                   # 環境変数（要作成）
├── package.json
└── README.md
```

## 🚀 セットアップ

### Gitから取得した後の初回セットアップ

このプロジェクトをGitから取得した後、以下の手順でセットアップを行ってください。

#### ステップ1: リポジトリのクローン（まだの場合）

```bash
git clone <リポジトリURL>
cd TrendInnovation_GetBaitoruPerformanceData-
```

#### ステップ2: 依存関係のインストール

```bash
npm install
```

**注意**: 
- 権限エラーが発生する場合は、ターミナルで直接実行してください
- 初回実行時は時間がかかる場合があります（依存パッケージのダウンロード）

#### ステップ3: Puppeteerのブラウザをインストール

```bash
npm run install-browser
```

または

```bash
npx puppeteer browsers install chrome
```

**注意**: この手順により、Puppeteerが使用するChromeブラウザが自動的にインストールされます。これにより、`.env`ファイルでブラウザパスを指定する必要がなくなります。

#### ステップ4: 環境変数の設定

`.env`ファイルを作成して、必要な設定を行います。

**Windowsの場合（推奨）:**
1. エクスプローラーでプロジェクトフォルダを開く
2. `setup-env.bat`をダブルクリックして実行
3. 完了メッセージが表示されるまで待つ

**macOS/Linuxの場合:**
```bash
chmod +x setup-env.sh
./setup-env.sh
```

**手動で作成する場合:**
プロジェクトルートに`.env`ファイルを作成し、以下の内容を記述してください：

```env
# バイトル企業データ ログイン情報（実際の値に置き換えてください）
BAITORU_LOGIN_URL=https://agent.baitoru.com/top
BAITORU_USERNAME=your_username
BAITORU_PASSWORD=your_password

# ファイルパス設定
INPUT_FILE=【バイトル】8月実績.xlsx

# Puppeteer設定
HEADLESS=false
BROWSER_TIMEOUT=30000
PAGE_TIMEOUT=60000
# ブラウザパスは指定不要（Puppeteerが自動的に見つけます）
# BROWSER_EXECUTABLE_PATH=

# リトライ設定
MAX_RETRIES=3
RETRY_DELAY=2000

# Google Sheets API設定
GOOGLE_SERVICE_ACCOUNT_KEY_PATH=./credentials.json
GOOGLE_SPREADSHEET_ID_NIGHT=your_night_spreadsheet_id
GOOGLE_SPREADSHEET_ID_NORMAL=your_normal_spreadsheet_id
GOOGLE_SHEET_NAME=Sheet1

# OpenAI API設定（オプション：AI機能を使用する場合）
# OPENAI_API_KEY=your_openai_api_key
# OPENAI_MODEL=gpt-4o-mini

# 設定検証をスキップ（テスト用：ログイン情報が未設定でもExcel読み込みテストが可能）
# SKIP_CONFIG_VALIDATION=true
```

#### ステップ5: Google Sheets APIの設定

1. **Google Cloud Consoleでサービスアカウントを作成**
   - [Google Cloud Console](https://console.cloud.google.com/)にアクセス
   - プロジェクトを作成または選択
   - 「APIとサービス」→「認証情報」→「サービスアカウントを作成」
   - サービスアカウント名を入力して作成

2. **サービスアカウントキーをダウンロード**
   - 作成したサービスアカウントを選択
   - 「キー」タブ→「キーを追加」→「新しいキーを作成」
   - JSON形式でダウンロード
   - ダウンロードしたファイルを`credentials.json`としてプロジェクトルートに配置

3. **Google Sheets APIを有効化**
   - 「APIとサービス」→「ライブラリ」
   - 「Google Sheets API」を検索して有効化

4. **スプレッドシートへのアクセス権限を付与**
   - 使用するGoogleスプレッドシートを開く
   - 「共有」ボタンをクリック
   - サービスアカウントのメールアドレス（`credentials.json`内の`client_email`）を追加
   - 「編集者」権限を付与

#### ステップ6: 動作確認

セットアップが完了したら、以下のコマンドで動作確認を行います：

```bash
npm run test:single
```

このコマンドは、入力Excelファイルの最初のデータ行のみを処理して動作確認を行います。

**注意**: 
- 初回実行時は、ログイン情報やGoogle Sheets APIの設定が正しく行われているか確認してください
- エラーが発生した場合は、「🔧 トラブルシューティング」セクションを参照してください

---

### 詳細なセットアップ手順

#### 1. 依存関係のインストール

```bash
npm install
```

**注意**: 権限エラーが発生する場合は、ターミナルで直接実行してください。

#### 2. 環境変数の設定

**重要**: `.env`ファイルを作成する必要があります。

#### 方法1: .env.exampleをコピー（推奨）

```bash
cp .env.example .env
```

その後、`.env`ファイルを編集して実際の値を設定してください。

#### 方法2: セットアップスクリプトを使用

**macOS/Linuxの場合:**
```bash
chmod +x setup-env.sh
./setup-env.sh
```

**Windowsの場合（3つの方法）:**

**方法A: バッチファイルを使用（最も簡単・推奨）**
1. エクスプローラーでプロジェクトフォルダを開く
2. `setup-env.bat`をダブルクリックして実行
3. 完了メッセージが表示されるまで待つ

**方法B: PowerShellで直接実行**
1. エクスプローラーでプロジェクトフォルダを開く
2. アドレスバーに`powershell`と入力してEnterキーを押す（または、Shiftキーを押しながら右クリック → 「PowerShellウィンドウをここで開く」）
3. 以下のコマンドを実行：
   ```powershell
   .\setup-env.ps1
   ```

**方法C: 実行ポリシーをバイパスして実行**
```powershell
powershell -ExecutionPolicy Bypass -File setup-env.ps1
```

**注意**: 
- PowerShellの実行ポリシーでエラーが発生する場合は、以下のコマンドを実行してください：
  ```powershell
  Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
  ```
- パスに日本語（「デスクトップ」など）が含まれている場合、方法A（バッチファイル）が最も確実です

セットアップスクリプトは、Google Sheets APIの基本設定を自動的に`.env`ファイルに追加または更新します。

#### 方法3: 手動で作成

プロジェクトルートに`.env`ファイルを作成し、以下の内容を記述してください：

```env
# バイトル企業データ ログイン情報（実際の値に置き換えてください）
BAITORU_LOGIN_URL=https://your-baitoru-url.com/login
BAITORU_USERNAME=your_username
BAITORU_PASSWORD=your_password

# ファイルパス設定
INPUT_FILE=【バイトル】8月実績.xlsx

# Puppeteer設定
HEADLESS=false
BROWSER_TIMEOUT=30000
PAGE_TIMEOUT=60000
# ブラウザのパスを手動で指定する場合（macOSで起動エラーが発生する場合）
# BROWSER_EXECUTABLE_PATH=/Applications/Google Chrome.app/Contents/MacOS/Google Chrome

# リトライ設定
MAX_RETRIES=3
RETRY_DELAY=2000

# Google Sheets API設定
GOOGLE_SERVICE_ACCOUNT_KEY_PATH=./credentials.json
GOOGLE_SPREADSHEET_ID_NIGHT=your_night_spreadsheet_id
GOOGLE_SPREADSHEET_ID_NORMAL=your_normal_spreadsheet_id
GOOGLE_SHEET_NAME=Sheet1

# OpenAI API設定（オプション：AI機能を使用する場合）
# OPENAI_API_KEY=your_openai_api_key
# OPENAI_MODEL=gpt-4o-mini

# 設定検証をスキップ（テスト用：ログイン情報が未設定でもExcel読み込みテストが可能）
SKIP_CONFIG_VALIDATION=true
```

**注意**: 
- ログイン情報が未設定でも、`SKIP_CONFIG_VALIDATION=true`を設定すればExcelファイルの読み込みテストは可能です
- 実際のスクレイピングを実行する場合は、ログイン情報を正しく設定してください
- Google Sheets APIを使用する場合は、サービスアカウントキーファイル（`credentials.json`）が必要です

### 3. Google Sheets APIの設定

詳細は「ステップ5: Google Sheets APIの設定」を参照してください。

### 4. セレクタの設定

`config/selectors.json`に実際のWebサイトのCSSセレクタを設定してください。

### 5. Excel列の設定

`config/excelColumns.json`に実際のExcelファイルの列構成を設定してください。

### 6. 職種カテゴリリストの管理

職種カテゴリリストは**JSONファイル**で管理します。

#### JSONファイルの場所

- **ナイト案件**: `config/jobCategoriesNight.json`
- **通常案件**: `config/jobCategoriesNormal.json`

#### JSONファイルの形式

```json
[
  {
    "large": "職種大",
    "medium": "職種中",
    "small": "職種小"
  },
  {
    "large": "職種大2",
    "medium": "職種中2",
    "small": "職種小2"
  }
]
```

#### ExcelファイルからJSONファイルへの変換

Excelファイル（`ナイト案件リスト.xlsx`、`通常案件リスト.xlsx`）を更新した場合、以下のコマンドでJSONファイルに変換できます：

```bash
npm run convert:excel-to-json
```

このコマンドは：
- `ナイト案件リスト.xlsx` → `config/jobCategoriesNight.json`
- `通常案件リスト.xlsx` → `config/jobCategoriesNormal.json`

に変換します。

#### JSONファイルの直接編集

JSONファイルは直接編集可能です。Excelファイルを経由せずに、JSONファイルを直接編集して職種カテゴリを管理することもできます。

**注意事項**:
- JSONファイルの形式（配列、オブジェクトの構造）を維持してください
- 各オブジェクトには`large`、`medium`、`small`の3つのプロパティが必要です
- 空の値は空文字列`""`として設定してください

## 📝 使用方法

### 実行

```bash
npm start
```

### 開発モード（ファイル変更を監視）

```bash
npm run dev
```

### 1行のみの動作確認

```bash
npm run test:single
```

このコマンドは【バイトル】8月実績.xlsxの2行目（最初のデータ行）のみを処理して動作確認を行います。

### セレクターのデバッグ

ログイン画面やその他のページのセレクターを特定する場合：

```bash
npm run debug:selectors
```

このコマンドは：
- ログイン画面に遷移します
- ページのHTMLとスクリーンショットを保存します（`debug-login-page.html`、`debug-login-page.png`）
- 入力フィールド、ボタンなどの候補セレクターを表示します
- 30秒間ブラウザを開いたままにして、手動で確認できます

**注意**: 初回実行時はPuppeteerのブラウザをインストールする必要があります：

```bash
npm run install-browser
```

または：

```bash
npx puppeteer browsers install chrome
```

## 🔄 処理フロー（手順①〜⑥）

### 手順①：入力ファイルから企業IDを取得
- 【バイトル】○月実績.xlsxのE列から企業IDを取得
- P列の掲載区分を確認（「ナイト」ならナイト案件、それ以外は通常案件）

### 手順②：ログインとTOPページ移動
- バイトル企業データにログイン
- TOPページに移動

### 手順③：企業IDで検索
- 企業ID欄に入力して検索

### 手順④：選択ボタンをクリック
- 検索結果から選択ボタンをクリック

### 手順⑤：掲載実績をダウンロード
- 掲載実績ページに移動
- 入力エクセルのM・N列（開始日・終了日）を指定
- CSVまたはExcelファイルをダウンロード
- ダウンロード完了を待機（最大30秒）
- ファイル名の衝突を回避（企業IDとタイムスタンプを付与）

### 手順⑥：バリデーション
- ダウンロードしたファイルのE・F列（掲載実績期間）とG・H列（申込期間）をチェック
- **同日の場合**：データを処理
- **同日でない場合**：
  - 開始日・終了日が一致しない場合は、指定期間内の全レコードを合計
  - 該当レコードがない場合はスキップして次の企業へ

### 手順⑦〜⑨：データ抽出と転記
- プラン（H列）を抽出し、AI判定で適切なプランを選択して転記
- PV数・応募数（U~X列）をスプレッドシートの対応列に転記
- 申込期間から週数を計算して転記

### 手順⑩〜⑪：原稿検索とプレビュー
- 仕事Noで原稿を検索
- プレビューボタンをクリックして新しいタブで開く

### 手順⑫〜⑰：詳細情報の抽出と転記
- 勤務地（都道府県・市区町村・最寄り駅）を抽出して転記
- 年月を計算して転記（申込開始日から抽出）
- **AI連携**：都道府県から地方を判定して転記
- 職種（大・中・小）を抽出
- **AI連携**：職種テキストから職種カテゴリリストにマッチング（キーワードマッチング → コサイン類似度判定）
- 給与形態と金額を抽出して転記（ナイト案件と通常案件で処理方法が異なる）

### 手順⑱〜⑳：企業情報の転記
- 企業ID・企業名を転記
- 店名（応募受付先名）を取得して転記
- 申込開始日・終了日を転記

### データ振り分け
- P列が「ナイト」→ Googleスプレッドシート（ナイト案件用）
- それ以外 → Googleスプレッドシート（通常案件用）

**注意**: データはExcelファイルではなく、Googleスプレッドシートに直接書き込まれます。

## ⚙️ 設定ファイル

### config/selectors.json

WebサイトのCSSセレクタを定義します。

### config/excelColumns.json

Excelファイルの列マッピングを定義します。CSVファイルの列名も`downloadFile.csvColumns`に定義します。

## ⚠️ 注意事項

- 初回実行前に、実際のWebサイトの構造に合わせてセレクタを調整してください
- Excelファイルの列構成が取説と異なる場合は、`excelColumns.json`を調整してください
- CSVファイルの列名は`config/excelColumns.json`の`downloadFile.csvColumns`で設定してください
- ダウンロードファイルは`downloads/`フォルダに保存されます（企業IDとタイムスタンプ付き）
- エラーが発生した場合でも、処理済みのデータはGoogleスプレッドシートに保存されます
- **Google Sheets APIについて**：
  - サービスアカウントキーファイル（`credentials.json`）が正しく配置されているか確認してください
  - スプレッドシートIDが正しく設定されているか確認してください
  - サービスアカウントにスプレッドシートへの編集権限が付与されているか確認してください
- **AI機能について**：
  - OpenAI APIキーが設定されていない場合、AI機能は使用されず、直接テキストが入力されます
  - AI機能を使用する場合は、`.env`に`OPENAI_API_KEY`を設定してください
  - モデルは`OPENAI_MODEL`で指定できます（デフォルト: `gpt-4o-mini`）
- **重複チェック機能**：
  - ユニークID列を使用して重複データを自動的にスキップします
  - ユニークIDは「企業ID_企業名_開始日_終了日」の形式で生成されます

## 🔧 トラブルシューティング

### PowerShellスクリプトが実行できない（Windows）

**エラー**: `.\setup-env.ps1 : このシステムではスクリプトの実行が無効になっているため...` または `cannot be loaded because running scripts is disabled`

**原因**: PowerShellの実行ポリシーがスクリプトの実行を禁止している

**対策1: 実行ポリシーを変更（推奨）**
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```
その後、再度スクリプトを実行：
```powershell
.\setup-env.ps1
```

**対策2: 実行ポリシーをバイパスして実行**
```powershell
powershell -ExecutionPolicy Bypass -File setup-env.ps1
```

**対策3: 一時的に実行ポリシーを変更**
```powershell
powershell -ExecutionPolicy RemoteSigned -File setup-env.ps1
```

**エラー**: `パスが見つかりません` または `PathNotFound`

**原因**: パスに日本語（例：「デスクトップ」）が含まれている場合、PowerShellが正しくパスを認識できないことがある

**対策1: バッチファイルを使用（最も簡単）**
1. エクスプローラーでプロジェクトフォルダを開く
2. `setup-env.bat`をダブルクリック
3. これで実行ポリシーやパスの問題を回避できます

**対策2: エクスプローラーからPowerShellを開く**
1. エクスプローラーでプロジェクトフォルダを開く
2. アドレスバーに`powershell`と入力してEnterキーを押す
   - または、Shiftキーを押しながらフォルダ内の空白部分を右クリック → 「PowerShellウィンドウをここで開く」
3. 以下のコマンドを実行：
   ```powershell
   .\setup-env.ps1
   ```

**対策3: 完全パスを指定して実行**
```powershell
powershell -ExecutionPolicy Bypass -File "C:\Users\vagab\OneDrive\デスクトップ\work\Git\TrendInnovation_GetBaitoruPerformanceData-\setup-env.ps1"
```

**対策4: 短いパス名（8.3形式）を使用**
```powershell
# 短いパス名を取得
powershell -Command "(Get-Item 'C:\Users\vagab\OneDrive\デスクトップ').FullName"
# 取得した短いパス名を使用して実行
```

**エラー**: コンソール出力で日本語が文字化けする

**原因**: PowerShellのコンソールエンコーディング設定の問題

**対策**: これは表示の問題のみで、実際の処理には影響しません。`.env`ファイルは正しく作成・更新されています。確認方法：
```powershell
Get-Content .env | Select-String "GOOGLE_"
```

**その他の問題**

- **スクリプトファイルが見つからない**: プロジェクトのルートディレクトリで実行しているか確認してください
- **権限エラー**: PowerShellを「管理者として実行」する必要はありません。通常のユーザー権限で実行できます
- **文字エンコーディングエラー**: スクリプトファイルがUTF-8 with BOMで保存されているか確認してください

### Puppeteerのブラウザ起動エラー

**エラー**: `Failed to launch the browser process!` または `spawn /Applications/Google Chrome.app/Contents/MacOS/Google Chrome ENOENT`

**原因**: 
- `.env`ファイルに異なるプラットフォーム（macOS/Windows/Linux）用のブラウザパスが設定されている
- Windows環境でmacOS用のパス（`/Applications/...`）が設定されている場合、そのパスは存在しない

**対策**:

1. **ブラウザの再インストール（推奨）**:
   ```bash
   npm run install-browser
   ```
   または
   ```bash
   npx puppeteer browsers install chrome
   ```
   これで、Puppeteerが自動的にブラウザを見つけます（`.env`ファイルの`BROWSER_EXECUTABLE_PATH`は不要）

2. **`.env`ファイルから`BROWSER_EXECUTABLE_PATH`を削除**:
   - `.env`ファイルを開く
   - `BROWSER_EXECUTABLE_PATH=...`の行を削除またはコメントアウト（`#`を先頭に追加）
   - Puppeteerが自動的にブラウザを見つけます

3. **プラットフォームに応じたブラウザパスを設定**:
   
   **Windowsの場合**:
   ```env
   BROWSER_EXECUTABLE_PATH=C:\Program Files\Google\Chrome\Application\chrome.exe
   ```
   または
   ```env
   BROWSER_EXECUTABLE_PATH=C:\Program Files (x86)\Google\Chrome\Application\chrome.exe
   ```
   
   **macOSの場合**:
   ```env
   BROWSER_EXECUTABLE_PATH=/Applications/Google Chrome.app/Contents/MacOS/Google Chrome
   ```
   または
   ```env
   BROWSER_EXECUTABLE_PATH=/Applications/Chromium.app/Contents/MacOS/Chromium
   ```
   
   **Linuxの場合**:
   ```env
   BROWSER_EXECUTABLE_PATH=/usr/bin/google-chrome
   ```
   または
   ```env
   BROWSER_EXECUTABLE_PATH=/usr/bin/chromium-browser
   ```

4. **macOSのセキュリティ設定を確認**（macOSのみ）:
   - システム環境設定 > セキュリティとプライバシー
   - Chrome/Chromiumの実行を許可

5. **ヘッドレスモードを無効にして試す**:
   `.env`ファイルに以下を追加:
   ```env
   HEADLESS=false
   ```

**注意**: 
- システムは自動的にプラットフォームを検出し、異なるプラットフォームのパスが設定されている場合は無視します
- ブラウザパスを設定しない場合、Puppeteerが自動的にブラウザを見つけます（推奨）

### 日付変換エラー

**エラー**: `NaN-NaN-NaN` が表示される

**対策**:
- ExcelJSが日付をDateオブジェクトとして読み込む場合がありますが、自動的に処理されます
- 無効な日付の場合は、その行はスキップされます
- 入力Excelファイルの日付形式を確認してください

### Google Sheets APIエラー

**エラー**: `GOOGLE_SERVICE_ACCOUNT_KEY_PATHが設定されていません` または `Google Sheets APIの初期化に失敗しました`

**対策**:
1. `.env`ファイルに`GOOGLE_SERVICE_ACCOUNT_KEY_PATH`が設定されているか確認
2. `credentials.json`ファイルがプロジェクトルートに存在するか確認
3. サービスアカウントキーファイルの内容が正しいか確認
4. Google Sheets APIが有効化されているか確認
5. サービスアカウントにスプレッドシートへの編集権限が付与されているか確認

### その他のエラー

- **ログインエラー**: `.env`ファイルのログイン情報を確認してください
- **セレクタエラー**: `config/selectors.json`のセレクタが正しいか確認してください
- **Excel列エラー**: `config/excelColumns.json`の列マッピングが正しいか確認してください
- **スプレッドシートIDエラー**: `.env`ファイルの`GOOGLE_SPREADSHEET_ID_NIGHT`と`GOOGLE_SPREADSHEET_ID_NORMAL`が正しく設定されているか確認してください

## 📊 実装済み機能

✅ 手順①〜⑳の完全自動化
✅ 通常案件/ナイト案件の自動振り分け
✅ CSV/Excelファイルの両対応
✅ ダウンロード待機処理（最大30秒）
✅ ファイル名衝突回避
✅ バリデーション（同日チェック）
✅ 複数レコードの合計処理（期間内の全レコードを合計）
✅ プレビューページからの詳細情報抽出
  - 勤務地（都道府県・市区町村・最寄り駅）
  - 職種（大・中・小）
  - 給与形態と金額
  - 店名（応募受付先名）
✅ Google Sheets API連携による直接データ書き込み
  - ナイト案件と通常案件で別々のスプレッドシートに自動振り分け
  - 重複チェック機能（ユニークIDによる）
  - 空白行への自動追記
✅ OpenAI API連携によるドロップダウン選択肢の自動マッチング
  - 地方の自動判定
  - 職種カテゴリの自動判定（コサイン類似度による）

## 📄 ライセンス

ISC
