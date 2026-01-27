import dotenv from 'dotenv';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';
import { readFileSync } from 'fs';

// 環境変数の読み込み
dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

// 設定オブジェクト
const config = {
  // ログイン情報
  baitoru: {
    loginUrl: process.env.BAITORU_LOGIN_URL || '',
    username: process.env.BAITORU_USERNAME || '',
    password: process.env.BAITORU_PASSWORD || ''
  },

  // ファイルパス
  files: {
    input: process.env.INPUT_FILE || '【バイトル】8月実績.xlsx',
    outputSurvey: process.env.OUTPUT_SURVEY_SHEET || 'バイトル　反響調査入力シート.xlsx',
    jobCategoryListNight: process.env.JOB_CATEGORY_LIST_NIGHT || join(__dirname, 'jobCategoriesNight.json'),
    jobCategoryListNormal: process.env.JOB_CATEGORY_LIST_NORMAL || join(__dirname, 'jobCategoriesNormal.json'),
    downloadDir: join(__dirname, '../downloads'),
    outputDir: join(__dirname, '../output')
  },
  googleSheets: {
    spreadsheetIdNight: process.env.GOOGLE_SPREADSHEET_ID_NIGHT || '',
    spreadsheetIdNormal: process.env.GOOGLE_SPREADSHEET_ID_NORMAL || '',
    sheetName: process.env.GOOGLE_SHEET_NAME || 'Sheet1',
    serviceAccountKeyPath: process.env.GOOGLE_SERVICE_ACCOUNT_KEY_PATH || ''
  },

  // Puppeteer設定
  puppeteer: {
    headless: process.env.HEADLESS === 'true',
    timeout: parseInt(process.env.BROWSER_TIMEOUT || '30000', 10),
    pageTimeout: parseInt(process.env.PAGE_TIMEOUT || '60000', 10),
    executablePath: process.env.BROWSER_EXECUTABLE_PATH || null,
    args: [
      '--no-sandbox',
      '--disable-setuid-sandbox',
      '--disable-dev-shm-usage',
      '--disable-accelerated-2d-canvas',
      '--disable-gpu'
      // macOSでは--single-processを削除（問題を引き起こす可能性があるため）
    ]
  },

  // リトライ設定
  retry: {
    maxRetries: parseInt(process.env.MAX_RETRIES || '3', 10),
    delay: parseInt(process.env.RETRY_DELAY || '2000', 10)
  },

  // セレクタ設定
  selectors: JSON.parse(
    readFileSync(join(__dirname, 'selectors.json'), 'utf-8')
  ),

  // Excel列設定
  excelColumns: JSON.parse(
    readFileSync(join(__dirname, 'excelColumns.json'), 'utf-8')
  )
};

// 設定の検証
function validateConfig(skipValidation = false) {
  // 環境変数 SKIP_CONFIG_VALIDATION=true が設定されている場合は検証をスキップ
  if (skipValidation || process.env.SKIP_CONFIG_VALIDATION === 'true') {
    console.warn('⚠️  設定の検証をスキップしました（SKIP_CONFIG_VALIDATION=true）');
    return;
  }

  const errors = [];

  if (!config.baitoru.loginUrl) {
    errors.push('BAITORU_LOGIN_URL が設定されていません');
  }
  if (!config.baitoru.username) {
    errors.push('BAITORU_USERNAME が設定されていません');
  }
  if (!config.baitoru.password) {
    errors.push('BAITORU_PASSWORD が設定されていません');
  }

  if (errors.length > 0) {
    throw new Error(`設定エラー:\n${errors.join('\n')}\n\n対策:\n1. .envファイルを作成してください\n2. または、SKIP_CONFIG_VALIDATION=true を設定して検証をスキップできます`);
  }
}

// 初期化時に検証
validateConfig();

export default config;
