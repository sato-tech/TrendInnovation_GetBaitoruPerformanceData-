import { loadConfig, Config } from './config';
import { logger } from './utils/logger';
import { RetryHandler } from './services/retry/retryHandler';
import { BrowserService } from './services/scraper/browser';
import { LoginService } from './services/scraper/login';
import { SearchService, SearchResult } from './services/scraper/search';
import { CsvWriter } from './services/csv/writer';
import { CsvAggregator } from './services/csv/aggregator';
import { GoogleSheetsService } from './services/spreadsheet/googleSheets';
import { ExcelService } from './services/spreadsheet/excel';

async function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function main() {
  try {
    logger.info('=== アプリケーション開始 ===');
    
    // 1. 設定読み込み
    const config = loadConfig();
    const retryHandler = new RetryHandler(config.retry.maxRetries, config.retry.retryDelay);

    // 2. スプレッドシートから検索キー取得
    logger.info('検索キーを取得中...');
    const searchKeys = await retryHandler.execute(
      async () => {
        if (config.spreadsheet.type === 'google') {
          const googleSheets = new GoogleSheetsService();
          await googleSheets.initialize(config);
          return await googleSheets.readSearchKeys(config);
        } else {
          const excelService = new ExcelService();
          return await excelService.readSearchKeys(config);
        }
      },
      '検索キー取得'
    );

    if (searchKeys.length === 0) {
      logger.warn('検索キーが見つかりませんでした。処理を終了します。');
      return;
    }

    logger.info(`${searchKeys.length}件の検索キーを取得しました`);

    // 3. Puppeteerブラウザ初期化
    const browserService = new BrowserService();
    const browser = await browserService.init(config.headless);
    const page = await browser.newPage();

    try {
      // 4. ログイン
      const loginService = new LoginService();
      await retryHandler.execute(
        () => loginService.login(page, config),
        'ログイン処理',
        {
          timeout: config.scraper.pageLoadTimeout * 2,
        }
      );

      // 5. 各検索キーで検索実行
      const allResults: SearchResult[] = [];
      const searchService = new SearchService();

      for (let i = 0; i < searchKeys.length; i++) {
        const searchKey = searchKeys[i];
        logger.info(`[${i + 1}/${searchKeys.length}] 検索実行: ${searchKey}`);

        try {
          const result = await retryHandler.execute(
            () => searchService.search(page, searchKey, config),
            `検索処理: ${searchKey}`,
            {
              timeout: config.scraper.searchTimeout,
            }
          );

          allResults.push(result);
          logger.info(`検索完了: ${searchKey} (${result.data.length}件の結果)`);
        } catch (error) {
          logger.error(`検索キー "${searchKey}" の処理に失敗しました`, error);
          // エラーが発生しても次の検索キーに進む
          allResults.push({
            searchKey,
            data: [],
            timestamp: new Date(),
          });
        }

        // レート制限対策（最後の検索キー以外は待機）
        if (i < searchKeys.length - 1) {
          await sleep(1000);
        }
      }

      // 6. CSV出力
      logger.info('CSVファイルを出力中...');
      const csvWriter = new CsvWriter();
      const csvPath = await csvWriter.write(allResults, './output');
      logger.info(`CSV出力完了: ${csvPath}`);

      // 7. CSV集計
      logger.info('CSVファイルを集計中...');
      const csvAggregator = new CsvAggregator();
      const aggregatedData = await csvAggregator.aggregate(csvPath);
      logger.info('集計完了', aggregatedData);

      // 8. スプレッドシートに追記（Google Sheetsの場合のみ）
      if (config.spreadsheet.type === 'google' && config.spreadsheet.outputSheetId) {
        logger.info('スプレッドシートに結果を追記中...');
        const googleSheets = new GoogleSheetsService();
        await googleSheets.initialize(config);
        await retryHandler.execute(
          () => googleSheets.appendResult(aggregatedData, config),
          'スプレッドシート追記'
        );
      } else {
        logger.info('スプレッドシートへの追記はスキップされました');
      }

      logger.info('=== 処理完了 ===');
    } finally {
      await browser.close();
      logger.info('ブラウザを閉じました');
    }
  } catch (error) {
    logger.error('=== 致命的エラー ===', error);
    process.exit(1);
  }
}

// エラーハンドリング
process.on('unhandledRejection', (reason, promise) => {
  logger.error('未処理のPromise拒否', { reason, promise });
  process.exit(1);
});

process.on('uncaughtException', (error) => {
  logger.error('未捕捉の例外', error);
  process.exit(1);
});

// メイン処理を実行
main().catch((error) => {
  logger.error('予期しないエラー', error);
  process.exit(1);
});
