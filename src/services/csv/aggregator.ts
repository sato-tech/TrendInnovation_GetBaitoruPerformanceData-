import fs from 'fs';
import csv from 'csv-parser';
import { logger } from '../../utils/logger';
import path from 'path';

export interface AggregatedResult {
  [key: string]: any;
}

export class CsvAggregator {
  async aggregate(csvPath: string): Promise<AggregatedResult> {
    try {
      if (!fs.existsSync(csvPath)) {
        throw new Error(`CSVファイルが見つかりません: ${csvPath}`);
      }

      logger.info(`CSVファイルを読み込み中: ${csvPath}`);

      const rows: Record<string, any>[] = [];

      return new Promise((resolve, reject) => {
        fs.createReadStream(csvPath)
          .pipe(csv())
          .on('data', (row) => {
            rows.push(row);
          })
          .on('end', () => {
            try {
              const aggregated = this.aggregateData(rows);
              logger.info('集計が完了しました', { rowCount: rows.length });
              resolve(aggregated);
            } catch (error) {
              reject(error);
            }
          })
          .on('error', (error) => {
            reject(error);
          });
      });
    } catch (error) {
      logger.error('CSV集計でエラーが発生しました', error);
      throw error;
    }
  }

  private aggregateData(rows: Record<string, any>[]): AggregatedResult {
    if (rows.length === 0) {
      return { total_count: 0 };
    }

    // 基本的な集計（実際の要件に合わせて調整が必要）
    const aggregated: AggregatedResult = {
      total_count: rows.length,
      processed_at: new Date().toISOString(),
    };

    // 検索キーごとの集計
    const searchKeyCounts: Record<string, number> = {};
    rows.forEach((row) => {
      const searchKey = row.search_key || 'unknown';
      searchKeyCounts[searchKey] = (searchKeyCounts[searchKey] || 0) + 1;
    });
    aggregated.search_key_counts = searchKeyCounts;

    // 数値カラムの合計・平均を計算（数値として解釈できるカラムのみ）
    const numericColumns: Record<string, number[]> = {};
    rows.forEach((row) => {
      Object.keys(row).forEach((key) => {
        if (key === 'search_key' || key === 'timestamp') return;
        const value = parseFloat(row[key]);
        if (!isNaN(value)) {
          if (!numericColumns[key]) {
            numericColumns[key] = [];
          }
          numericColumns[key].push(value);
        }
      });
    });

    Object.keys(numericColumns).forEach((key) => {
      const values = numericColumns[key];
      aggregated[`${key}_sum`] = values.reduce((a, b) => a + b, 0);
      aggregated[`${key}_avg`] = values.reduce((a, b) => a + b, 0) / values.length;
      aggregated[`${key}_min`] = Math.min(...values);
      aggregated[`${key}_max`] = Math.max(...values);
    });

    return aggregated;
  }
}

export default CsvAggregator;
