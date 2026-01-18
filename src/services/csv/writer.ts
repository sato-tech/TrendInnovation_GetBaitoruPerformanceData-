import fs from 'fs';
import path from 'path';
import { createObjectCsvWriter } from 'csv-writer';
import { logger } from '../../utils/logger';
import { SearchResult } from '../scraper/search';

export class CsvWriter {
  async write(results: SearchResult[], outputDir: string = './output'): Promise<string> {
    try {
      // 出力ディレクトリが存在しない場合は作成
      if (!fs.existsSync(outputDir)) {
        fs.mkdirSync(outputDir, { recursive: true });
      }

      const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
      const filename = `search_results_${timestamp}.csv`;
      const filepath = path.join(outputDir, filename);

      if (results.length === 0) {
        logger.warn('出力するデータがありません');
        return filepath;
      }

      // 全結果から全カラムを収集
      const allColumns = new Set<string>();
      results.forEach((result) => {
        result.data.forEach((row) => {
          Object.keys(row).forEach((key) => allColumns.add(key));
        });
      });

      // 検索キーとタイムスタンプも追加
      allColumns.add('search_key');
      allColumns.add('timestamp');

      const columns = Array.from(allColumns).map((col) => ({
        id: col,
        title: col,
      }));

      const csvWriter = createObjectCsvWriter({
        path: filepath,
        header: columns,
        encoding: 'utf8',
      });

      // データをフラット化
      const flatData: Record<string, any>[] = [];
      results.forEach((result) => {
        result.data.forEach((row) => {
          flatData.push({
            ...row,
            search_key: result.searchKey,
            timestamp: result.timestamp.toISOString(),
          });
        });
      });

      await csvWriter.writeRecords(flatData);
      logger.info(`CSVファイルを出力しました: ${filepath} (${flatData.length}件)`);

      return filepath;
    } catch (error) {
      logger.error('CSV出力でエラーが発生しました', error);
      throw error;
    }
  }
}

export default CsvWriter;
