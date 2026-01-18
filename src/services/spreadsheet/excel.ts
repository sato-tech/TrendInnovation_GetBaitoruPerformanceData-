import * as XLSX from 'xlsx';
import fs from 'fs';
import { logger } from '../../utils/logger';
import { Config } from '../../config';

export class ExcelService {
  async readSearchKeys(config: Config): Promise<string[]> {
    try {
      if (!config.spreadsheet.filePath) {
        throw new Error('Excelファイルパスが設定されていません');
      }

      if (!fs.existsSync(config.spreadsheet.filePath)) {
        throw new Error(`Excelファイルが見つかりません: ${config.spreadsheet.filePath}`);
      }

      logger.info(`Excelファイルを読み込み中: ${config.spreadsheet.filePath}`);

      const workbook = XLSX.readFile(config.spreadsheet.filePath);
      
      // 範囲からシート名と範囲を抽出（例: "Sheet1!A2:B"）
      const rangeMatch = config.spreadsheet.range.match(/^([^!]+)!(.+)$/);
      const sheetName = rangeMatch ? rangeMatch[1] : workbook.SheetNames[0];
      const rangeStr = rangeMatch ? rangeMatch[2] : 'A2:B';

      const worksheet = workbook.Sheets[sheetName];
      if (!worksheet) {
        throw new Error(`シート "${sheetName}" が見つかりません`);
      }

      // 範囲を解析（例: "A2:B"）
      const range = XLSX.utils.decode_range(rangeStr);
      const columnIndex = this.getColumnIndex(config.spreadsheet.searchKeyColumn);

      const searchKeys: string[] = [];
      for (let row = range.s.r; row <= range.e.r; row++) {
        const cellAddress = XLSX.utils.encode_cell({ r: row, c: columnIndex });
        const cell = worksheet[cellAddress];
        if (cell && cell.v) {
          const value = cell.v.toString().trim();
          if (value !== '') {
            searchKeys.push(value);
          }
        }
      }

      logger.info(`${searchKeys.length}件の検索キーを取得しました`);
      return searchKeys;
    } catch (error) {
      logger.error('検索キーの読み取りに失敗しました', error);
      throw error;
    }
  }

  private getColumnIndex(columnName: string): number {
    // A=0, B=1, C=2, ... の形式に対応
    let index = 0;
    for (let i = 0; i < columnName.length; i++) {
      index = index * 26 + (columnName.charCodeAt(i) - 64);
    }
    return index - 1;
  }
}

export default ExcelService;
