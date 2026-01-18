import { Page } from 'puppeteer';
import { logger } from '../../utils/logger';
import { Config } from '../../config';

export interface SearchResult {
  searchKey: string;
  data: Record<string, any>[];
  timestamp: Date;
}

export class SearchService {
  async search(page: Page, searchKey: string, config: Config): Promise<SearchResult> {
    try {
      logger.info(`検索を実行: ${searchKey}`);

      // 検索ボックスを探して入力（セレクタは実際のサイトに合わせて調整）
      const searchSelectors = [
        'input[type="search"]',
        'input[name="search"]',
        'input[id="search"]',
        'input[placeholder*="検索"]',
        'input[placeholder*="Search"]',
      ];

      let searchInputFound = false;
      for (const selector of searchSelectors) {
        try {
          await page.waitForSelector(selector, { timeout: 5000 });
          await page.type(selector, searchKey, { delay: 100 });
          searchInputFound = true;
          break;
        } catch {
          continue;
        }
      }

      if (!searchInputFound) {
        throw new Error('検索入力欄が見つかりませんでした');
      }

      // 検索ボタンをクリック
      const searchButtonSelectors = [
        'button[type="submit"]',
        'button:has-text("検索")',
        'button:has-text("Search")',
        'input[type="submit"]',
      ];

      let searchButtonFound = false;
      for (const selector of searchButtonSelectors) {
        try {
          await page.waitForSelector(selector, { timeout: 3000 });
          await page.click(selector);
          searchButtonFound = true;
          break;
        } catch {
          continue;
        }
      }

      if (!searchButtonFound) {
        // Enterキーで検索を実行
        await page.keyboard.press('Enter');
      }

      // 検索結果の読み込みを待機
      await page.waitForTimeout(2000);
      await page.waitForSelector('table, .result, .data-table, [class*="result"], [class*="table"]', {
        timeout: config.scraper.searchTimeout,
      });

      // 検索結果を取得（実際のサイト構造に合わせて調整が必要）
      const results = await page.evaluate(() => {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const data: Record<string, any>[] = [];
        
        // テーブル形式の場合
        // @ts-ignore - page.evaluate内ではdocumentが利用可能
        const tables = document.querySelectorAll('table');
        if (tables.length > 0) {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          tables.forEach((table: any) => {
            const rows = table.querySelectorAll('tr');
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            rows.forEach((row: any, rowIndex: number) => {
              if (rowIndex === 0) return; // ヘッダー行をスキップ
              
              const cells = row.querySelectorAll('td, th');
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              const rowData: Record<string, any> = {};
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              cells.forEach((cell: any, cellIndex: number) => {
                const header = table.querySelector(`tr:first-child th:nth-child(${cellIndex + 1}), tr:first-child td:nth-child(${cellIndex + 1})`);
                const key = header?.textContent?.trim() || `column_${cellIndex}`;
                rowData[key] = cell.textContent?.trim() || '';
              });
              if (Object.keys(rowData).length > 0) {
                data.push(rowData);
              }
            });
          });
        } else {
          // テーブル以外の形式の場合（実際のサイトに合わせて調整）
          // @ts-ignore - page.evaluate内ではdocumentが利用可能
          const resultElements = document.querySelectorAll('.result, [class*="result"], [data-result]');
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          resultElements.forEach((element: any) => {
            data.push({
              text: element.textContent?.trim() || '',
              html: element.innerHTML,
            });
          });
        }
        
        return data;
      });

      logger.info(`検索結果を取得: ${results.length}件`);

      return {
        searchKey,
        data: results,
        timestamp: new Date(),
      };
    } catch (error) {
      logger.error(`検索処理でエラーが発生しました: ${searchKey}`, error);
      throw error;
    }
  }
}

export default SearchService;
