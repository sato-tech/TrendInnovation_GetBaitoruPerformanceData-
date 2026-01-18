import dotenv from 'dotenv';
import path from 'path';
import { logger } from '../utils/logger';

// .envファイルを読み込み
dotenv.config();

export interface Config {
  nodeEnv: string;
  headless: boolean;
  spreadsheet: {
    type: 'google' | 'excel';
    id?: string;
    filePath?: string;
    range: string;
    searchKeyColumn: string;
    outputSheetId?: string;
    outputSheetRange: string;
  };
  google: {
    credentialsPath: string;
    tokenPath: string;
  };
  scraper: {
    loginUrl: string;
    username: string;
    password: string;
    searchTimeout: number;
    pageLoadTimeout: number;
  };
  retry: {
    maxRetries: number;
    retryDelay: number;
  };
  log: {
    level: string;
    filePath: string;
  };
}

function getEnv(key: string, defaultValue?: string): string {
  const value = process.env[key];
  if (!value && !defaultValue) {
    throw new Error(`環境変数 ${key} が設定されていません`);
  }
  return value || defaultValue!;
}

function getEnvNumber(key: string, defaultValue: number): number {
  const value = process.env[key];
  if (!value) {
    return defaultValue;
  }
  const num = parseInt(value, 10);
  if (isNaN(num)) {
    logger.warn(`環境変数 ${key} が数値ではないため、デフォルト値 ${defaultValue} を使用します`);
    return defaultValue;
  }
  return num;
}

function getEnvBoolean(key: string, defaultValue: boolean): boolean {
  const value = process.env[key];
  if (!value) {
    return defaultValue;
  }
  return value.toLowerCase() === 'true';
}

export function loadConfig(): Config {
  const spreadsheetType = getEnv('SPREADSHEET_TYPE', 'google') as 'google' | 'excel';
  
  const config: Config = {
    nodeEnv: getEnv('NODE_ENV', 'development'),
    headless: getEnvBoolean('HEADLESS', true),
    spreadsheet: {
      type: spreadsheetType,
      id: spreadsheetType === 'google' ? getEnv('SPREADSHEET_ID') : undefined,
      filePath: spreadsheetType === 'excel' ? getEnv('SPREADSHEET_FILE_PATH') : undefined,
      range: getEnv('SPREADSHEET_RANGE', 'Sheet1!A2:B'),
      searchKeyColumn: getEnv('SEARCH_KEY_COLUMN', 'A'),
      outputSheetId: getEnv('OUTPUT_SHEET_ID', ''),
      outputSheetRange: getEnv('OUTPUT_SHEET_RANGE', 'Sheet1!A1'),
    },
    google: {
      credentialsPath: getEnv('GOOGLE_CREDENTIALS_PATH', './credentials.json'),
      tokenPath: getEnv('GOOGLE_TOKEN_PATH', './token.json'),
    },
    scraper: {
      loginUrl: getEnv('LOGIN_URL'),
      username: getEnv('LOGIN_USERNAME'),
      password: getEnv('LOGIN_PASSWORD'),
      searchTimeout: getEnvNumber('SEARCH_TIMEOUT', 30000),
      pageLoadTimeout: getEnvNumber('PAGE_LOAD_TIMEOUT', 60000),
    },
    retry: {
      maxRetries: getEnvNumber('MAX_RETRIES', 3),
      retryDelay: getEnvNumber('RETRY_DELAY', 2000),
    },
    log: {
      level: getEnv('LOG_LEVEL', 'info'),
      filePath: getEnv('LOG_FILE_PATH', './logs/app.log'),
    },
  };

  // バリデーション
  if (spreadsheetType === 'google' && !config.spreadsheet.id) {
    throw new Error('Google Sheetsを使用する場合、SPREADSHEET_IDが必要です');
  }
  if (spreadsheetType === 'excel' && !config.spreadsheet.filePath) {
    throw new Error('Excelを使用する場合、SPREADSHEET_FILE_PATHが必要です');
  }

  logger.info('設定を読み込みました', {
    nodeEnv: config.nodeEnv,
    headless: config.headless,
    spreadsheetType: config.spreadsheet.type,
  });

  return config;
}

export default loadConfig;
