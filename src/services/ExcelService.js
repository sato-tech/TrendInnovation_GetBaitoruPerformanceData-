import ExcelJS from 'exceljs';
import config from '../../config/config.js';
import { readFileSync, promises as fs } from 'fs';
import { parse } from 'csv-parse/sync';
import iconv from 'iconv-lite';
import jschardet from 'jschardet';
import { excelDateToJSDate } from '../utils/dateUtils.js';

/**
 * Excel操作を担当するサービスクラス
 */
class ExcelService {
  /**
   * 入力ファイル（【バイトル】○月実績.xlsx）を読み込む
   * @returns {Promise<ExcelJS.Workbook>}
   */
  async loadInputFile() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(config.files.input);
    return workbook;
  }


  /**
   * ワークブックを保存する
   * @param {ExcelJS.Workbook} workbook - 保存するワークブック
   * @param {string} filePath - 保存先パス
   */
  async saveWorkbook(workbook, filePath) {
    if (!workbook) {
      throw new Error('ワークブックがnullです');
    }
    try {
      console.log(`    [DEBUG] ワークブック保存中: ${filePath}`);
      await workbook.xlsx.writeFile(filePath);
      console.log(`    [DEBUG] ワークブック保存完了: ${filePath}`);
    } catch (error) {
      console.error(`    [DEBUG] ワークブック保存エラー: ${error.message}`);
      throw error;
    }
  }

  /**
   * セルの値を取得する
   * @param {ExcelJS.Worksheet} worksheet - ワークシート
   * @param {string} column - 列（例: 'A', 'B'）
   * @param {number} row - 行番号
   * @returns {any} セルの値
   */
  getCellValue(worksheet, column, row) {
    const cell = worksheet.getCell(`${column}${row}`);
    return cell.value;
  }

  /**
   * セルに値を設定する
   * @param {ExcelJS.Worksheet} worksheet - ワークシート
   * @param {string} column - 列
   * @param {number} row - 行番号
   * @param {any} value - 設定する値
   */
  setCellValue(worksheet, column, row, value) {
    if (!worksheet) {
      console.warn(`⚠️  setCellValue: worksheetがnullです (列: ${column}, 行: ${row})`);
      return;
    }
    try {
      const cell = worksheet.getCell(`${column}${row}`);
      cell.value = value;
      // デバッグ: 重要な列のみログ出力（ログが多すぎるのを防ぐ）
      if (['A', 'B', 'C', 'E', 'F', 'G', 'H', 'I', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'Z', 'AA', 'AB', 'AC', 'AD'].includes(column)) {
        console.log(`    [DEBUG] セル書き込み: ${column}${row} = ${value}`);
      }
    } catch (error) {
      console.error(`❌ setCellValueエラー (列: ${column}, 行: ${row}): ${error.message}`);
    }
  }

  /**
   * ダウンロードファイルを読み込む
   * @param {string} filePath - ダウンロードファイルのパス
   * @returns {Promise<ExcelJS.Workbook>}
   */
  async loadDownloadFile(filePath) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    return workbook;
  }

  /**
   * 期間（週数）を計算する
   * @param {Date} startDate - 開始日
   * @param {Date} endDate - 終了日
   * @returns {number} 週数
   */
  calculateWeeks(startDate, endDate) {
    const diffTime = Math.abs(endDate - startDate);
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    return Math.ceil(diffDays / 7);
  }

  /**
   * CSVファイルを読み込む（文字化けを自動解消）
   * @param {string} filePath - CSVファイルのパス
   * @returns {Promise<Array<Object>>} CSVデータの配列（オブジェクト形式）
   */
  async loadCSVFile(filePath) {
    // ファイルが存在し、読み取り可能か確認（リトライ処理付き）
    let fileExists = false;
    for (let retry = 0; retry < 10; retry++) {
      try {
        const stats = await fs.stat(filePath);
        if (stats.size > 0) {
          fileExists = true;
          break;
        }
      } catch (error) {
        if (retry < 9) {
          console.log(`ファイルの存在確認中... (${retry + 1}/10)`);
          await new Promise(resolve => setTimeout(resolve, 500));
        } else {
          throw new Error(`CSVファイルが見つかりません: ${filePath}`);
        }
      }
    }
    
    if (!fileExists) {
      throw new Error(`CSVファイルが空または存在しません: ${filePath}`);
    }
    // ファイルをバイナリとして読み込む
    const fileBuffer = readFileSync(filePath);
    
    // 文字エンコーディングを自動検出
    let detected = null;
    try {
      detected = jschardet.detect(fileBuffer);
    } catch (error) {
      console.warn(`文字エンコーディング検出エラー: ${error.message}`);
    }
    
    // 検出されたエンコーディングを正規化
    const encodingMap = {
      'Shift_JIS': 'shift_jis',
      'shift-jis': 'shift_jis',
      'SJIS': 'shift_jis',
      'sjis': 'shift_jis',
      'EUC-JP': 'euc-jp',
      'euc-jp': 'euc-jp',
      'ISO-2022-JP': 'iso-2022-jp',
      'iso-2022-jp': 'iso-2022-jp',
      'UTF-8': 'utf8',
      'utf-8': 'utf8',
      'utf8': 'utf8',
      'Windows-31J': 'shift_jis', // WindowsのShift_JIS
      'windows-31j': 'shift_jis',
      'CP932': 'shift_jis', // WindowsのShift_JISの別名
      'cp932': 'shift_jis'
    };
    
    let encoding = detected?.encoding ? (encodingMap[detected.encoding] || detected.encoding.toLowerCase()) : null;
    
    console.log(`CSVファイルの文字エンコーディングを検出: ${detected?.encoding || '不明'} (信頼度: ${detected?.confidence || 0})`);
    
    // 文字エンコーディングを変換してUTF-8に統一
    // 日本語CSVファイルは通常Shift_JISなので、まずShift_JISを試す
    const encodingsToTry = [];
    
    if (encoding && encoding !== 'utf8') {
      encodingsToTry.push(encoding);
    }
    // Shift_JISを優先的に試す（日本語CSVの一般的なエンコーディング）
    if (encoding !== 'shift_jis') {
      encodingsToTry.push('shift_jis');
    }
    // その他の日本語エンコーディング
    encodingsToTry.push('euc-jp', 'iso-2022-jp');
    // 最後にUTF-8を試す
    encodingsToTry.push('utf8');
    
    let fileContent = null;
    let successfulEncoding = null;
    
    for (const tryEncoding of encodingsToTry) {
      try {
        fileContent = iconv.decode(fileBuffer, tryEncoding);
        
        // 文字化けチェック（日本語文字が含まれているか確認）
        const hasJapanese = /[\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FAF]/.test(fileContent);
        
        // 最初の数行をチェックして、正常なCSV形式か確認
        const firstLines = fileContent.split('\n').slice(0, 3);
        const hasValidCSV = firstLines.some(line => line.includes(',') && line.length > 10);
        
        // 文字化け文字（）が含まれていないか確認
        const hasGarbledChars = fileContent.includes('');
        
        if (hasJapanese || (hasValidCSV && !hasGarbledChars)) {
          // 日本語が含まれている、または有効なCSV形式で文字化けがない場合
          successfulEncoding = tryEncoding;
          console.log(`✓ 文字エンコーディング変換成功: ${tryEncoding}`);
          break;
        }
      } catch (error) {
        // このエンコーディングでは変換できなかった
        continue;
      }
    }
    
    if (!fileContent || !successfulEncoding) {
      // デバッグ情報を出力
      console.error('CSVファイルの文字エンコーディング変換に失敗しました。');
      console.error('試行したエンコーディング:', encodingsToTry);
      console.error('ファイルパス:', filePath);
      console.error('ファイルサイズ:', fileBuffer.length, 'bytes');
      
      // 最後の手段として、Shift_JISで強制的に変換を試みる
      console.log('最後の手段として、Shift_JISで強制的に変換を試みます...');
      try {
        fileContent = iconv.decode(fileBuffer, 'shift_jis');
        successfulEncoding = 'shift_jis';
        console.log('✓ Shift_JISでの強制変換に成功しました');
      } catch (forceError) {
        throw new Error(`CSVファイルの文字エンコーディングを変換できませんでした: ${forceError.message}`);
      }
    }
    
    // CSVをパース（エラーハンドリング付き）
    let records = [];
    try {
      records = parse(fileContent, {
        columns: true,
        skip_empty_lines: true,
        encoding: 'utf8' // 既にUTF-8に変換済み
      });
      
      if (!records || records.length === 0) {
        console.warn('⚠️  CSVファイルをパースしましたが、レコードが0件です（ヘッダーのみ）。');
        console.warn('ファイル内容の最初の500文字:', fileContent.substring(0, 500));
        // データがない場合は空の配列を返す（呼び出し側でスキップ処理を行う）
      }
    } catch (parseError) {
      console.error('CSVパースエラー:', parseError.message);
      console.error('ファイル内容の最初の500文字:', fileContent.substring(0, 500));
      throw new Error(`CSVファイルのパースに失敗しました: ${parseError.message}`);
    }
    
    console.log(`✓ CSVファイルを読み込みました: ${records.length}件のレコード (エンコーディング: ${successfulEncoding})`);
    return records;
  }

  /**
   * CSVレコードから列の値を取得する（列名でアクセス）
   * @param {Object} record - CSVレコードオブジェクト
   * @param {string} columnName - 列名（CSVのヘッダー）
   * @returns {any} セルの値
   */
  getCSVValue(record, columnName) {
    return record[columnName] || null;
  }

  /**
   * 日付文字列を比較する（YYYY-MM-DD形式を想定）
   * @param {Date|number|string} date1 - 日付1（Dateオブジェクト、数値、または文字列）
   * @param {Date|number|string} date2 - 日付2（Dateオブジェクト、数値、または文字列）
   * @returns {boolean} 同じ日付かどうか
   */
  isSameDate(date1, date2) {
    if (!date1 || !date2) return false;
    
    // 文字列の場合は、YYYY/MM/DD形式をYYYY-MM-DD形式に正規化して比較
    if (typeof date1 === 'string' && typeof date2 === 'string') {
      const trimmed1 = date1.trim();
      const trimmed2 = date2.trim();
      
      // 空文字列の場合はfalse
      if (trimmed1 === '' || trimmed2 === '') return false;
      
      // YYYY/MM/DD形式をYYYY-MM-DD形式に変換
      let normalized1 = trimmed1.replace(/\//g, '-');
      let normalized2 = trimmed2.replace(/\//g, '-');
      
      // YYYY-MM-DD形式に変換（パディングを追加）
      const normalizeDateString = (dateStr) => {
        // 既にYYYY-MM-DD形式の場合はそのまま
        if (/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
          return dateStr;
        }
        // YYYY-M-D形式の場合はパディング
        const parts = dateStr.split('-');
        if (parts.length === 3) {
          const year = parts[0].padStart(4, '0');
          const month = parts[1].padStart(2, '0');
          const day = parts[2].padStart(2, '0');
          return `${year}-${month}-${day}`;
        }
        return dateStr;
      };
      
      normalized1 = normalizeDateString(normalized1);
      normalized2 = normalizeDateString(normalized2);
      
      // YYYY-MM-DD形式の文字列を直接比較
      if (/^\d{4}-\d{2}-\d{2}$/.test(normalized1) && /^\d{4}-\d{2}-\d{2}$/.test(normalized2)) {
        return normalized1 === normalized2;
      }
    }
    
    // Dateオブジェクトに変換して比較
    const d1 = excelDateToJSDate(date1);
    const d2 = excelDateToJSDate(date2);
    if (!d1 || !d2) return false;
    
    // UTC日付で比較（タイムゾーンの影響を避けるため）
    const year1 = d1.getUTCFullYear();
    const month1 = d1.getUTCMonth();
    const day1 = d1.getUTCDate();
    const year2 = d2.getUTCFullYear();
    const month2 = d2.getUTCMonth();
    const day2 = d2.getUTCDate();
    
    return year1 === year2 && month1 === month2 && day1 === day2;
  }

  /**
   * CSVレコードを集計する（U~X列の数値を合計）
   * @param {Array<Object>} records - CSVレコードの配列
   * @param {Array<string>} performanceDataColumns - 集計する列名の配列
   * @returns {Object} 集計されたデータ
   */
  aggregateCSVRecords(records, performanceDataColumns) {
    if (!records || records.length === 0) {
      return null;
    }

    const aggregated = { ...records[0] }; // 最初のレコードの情報を保持

    // 数値列を合計
    performanceDataColumns.forEach((colName, index) => {
      const sum = records.reduce((acc, record) => {
        const value = parseFloat(record[colName] || 0);
        return acc + (isNaN(value) ? 0 : value);
      }, 0);
      aggregated[colName] = sum;
    });

    return aggregated;
  }

  /**
   * CSVレコードが指定期間内に含まれるかチェック
   * @param {Object} record - CSVレコード
   * @param {Date} startDate - 開始日
   * @param {Date} endDate - 終了日
   * @param {string} startDateColumn - 開始日列名
   * @param {string} endDateColumn - 終了日列名
   * @returns {boolean} 期間内に含まれるかどうか
   */
  isRecordInDateRange(record, startDate, endDate, startDateColumn, endDateColumn) {
    const recordStart = new Date(this.getCSVValue(record, startDateColumn));
    const recordEnd = new Date(this.getCSVValue(record, endDateColumn));

    if (isNaN(recordStart.getTime()) || isNaN(recordEnd.getTime())) {
      return false;
    }

    // レコードの開始日・終了日が指定期間内に含まれるか
    return recordStart >= startDate && recordEnd <= endDate;
  }

  /**
   * CSVレコードから値を取得（列名または列インデックスで）
   * @param {Object} record - CSVレコード
   * @param {string|number} column - 列名または列インデックス
   * @returns {any} セルの値
   */
  getValueFromRecord(record, column) {
    if (typeof column === 'number') {
      // 列インデックスの場合（A=0, B=1, ...）
      const keys = Object.keys(record);
      return keys[column] ? record[keys[column]] : null;
    }
    return this.getCSVValue(record, column);
  }

  /**
   * Excelワークシートからドロップダウン選択肢を取得する
   * @param {ExcelJS.Worksheet} worksheet - ワークシート
   * @param {string} column - 列（例: 'C'）
   * @param {number} startRow - 開始行（デフォルト: 2）
   * @param {number} maxRows - 最大行数（デフォルト: 100）
   * @returns {Array<string>} 選択肢の配列
   */
  getDropdownOptions(worksheet, column, startRow = 2, maxRows = 100) {
    const options = new Set();
    
    for (let row = startRow; row <= startRow + maxRows; row++) {
      const value = this.getCellValue(worksheet, column, row);
      if (value && typeof value === 'string' && value.trim() !== '') {
        options.add(value.trim());
      }
    }

    return Array.from(options).filter(opt => opt !== '');
  }

  /**
   * 年月を日付から抽出する
   * @param {Date} date - 日付オブジェクト
   * @returns {{year: number, month: number}} 年月
   */
  extractYearMonth(date) {
    return {
      year: date.getFullYear(),
      month: date.getMonth() + 1
    };
  }

  /**
   * ワークシートの指定行以降をクリアする
   * @param {ExcelJS.Worksheet} worksheet - ワークシート
   * @param {number} startRow - 開始行（この行以降をクリア）
   */
  clearRowsFrom(worksheet, startRow) {
    const maxRow = worksheet.rowCount;
    const maxCol = worksheet.columnCount || 100; // 最大列数を取得（デフォルト100列）
    
    for (let row = startRow; row <= maxRow; row++) {
      for (let col = 1; col <= maxCol; col++) {
        const cell = worksheet.getCell(row, col);
        cell.value = null;
      }
    }
  }

  /**
   * ワークシートの最初の空白行を見つける
   * @param {ExcelJS.Worksheet} worksheet - ワークシート
   * @param {string} checkColumn - チェックする列（デフォルト: 'A'、年列）
   * @param {number} startRow - 検索開始行（デフォルト: 2、ヘッダー行を除く）
   * @returns {number} 最初の空白行の行番号
   */
  findFirstEmptyRow(worksheet, checkColumn = 'A', startRow = 2) {
    if (!worksheet) {
      return startRow;
    }
    
    // シートの最大行数を取得
    const maxRow = worksheet.rowCount || startRow;
    
    // startRowから順番に見ていって、最初の空白行を見つける
    for (let row = startRow; row <= maxRow + 1; row++) {
      const cellValue = this.getCellValue(worksheet, checkColumn, row);
      // セルの値が空（null、undefined、空文字列）の場合は空白行とみなす
      if (!cellValue || cellValue === '' || cellValue === null || cellValue === undefined) {
        return row;
      }
    }
    
    // すべての行にデータがある場合は、次の行を返す
    return maxRow + 1;
  }

  /**
   * 案件リストファイルを読み込む
   * @param {string} filePath - ファイルパス
   * @returns {Promise<ExcelJS.Workbook>}
   */
  async loadJobCategoryList(filePath) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    return workbook;
  }

  /**
   * JSONファイルから職種カテゴリを読み込む
   * @param {string} filePath - JSONファイルパス
   * @returns {Promise<Array<{large: string, medium: string, small: string, combined: string}>>}
   */
  async loadJobCategoriesFromJSON(filePath) {
    try {
      const jsonData = readFileSync(filePath, 'utf-8');
      const categories = JSON.parse(jsonData);
      
      // combinedプロパティを追加
      return categories.map(category => ({
        large: String(category.large || ''),
        medium: String(category.medium || ''),
        small: String(category.small || ''),
        combined: `${category.large || ''} ${category.medium || ''} ${category.small || ''}`.trim()
      }));
    } catch (error) {
      console.error(`職種カテゴリJSONファイルの読み込みエラー: ${error.message}`);
      return [];
    }
  }

  /**
   * 案件リストファイルから職種カテゴリを読み込む（Excel版 - 後方互換性のため残す）
   * @param {string} filePath - ファイルパス
   * @param {string} largeCol - 職種大の列（デフォルト: 'A'）
   * @param {string} mediumCol - 職種中の列（デフォルト: 'B'）
   * @param {string} smallCol - 職種小の列（デフォルト: 'C'）
   * @param {number} startRow - 開始行（デフォルト: 2）
   * @returns {Promise<Array<{large: string, medium: string, small: string, combined: string}>>}
   */
  async loadJobCategoriesFromList(filePath, largeCol = 'A', mediumCol = 'B', smallCol = 'C', startRow = 2) {
    try {
      const workbook = await this.loadJobCategoryList(filePath);
      const sheet = workbook.getWorksheet(1);
      const categories = [];
      const maxRow = sheet.rowCount || 1000;
      let consecutiveEmptyRows = 0;
      const maxConsecutiveEmptyRows = 10; // 連続して10行空行があれば終了
      
      for (let row = startRow; row <= maxRow; row++) {
        const large = this.getCellValue(sheet, largeCol, row);
        const medium = this.getCellValue(sheet, mediumCol, row);
        const small = this.getCellValue(sheet, smallCol, row);
        
        if (large || medium || small) {
          consecutiveEmptyRows = 0; // データがある行を見つけたらリセット
          const combined = `${large || ''} ${medium || ''} ${small || ''}`.trim();
          if (combined) {
            categories.push({
              large: String(large || ''),
              medium: String(medium || ''),
              small: String(small || ''),
              combined: combined
            });
          }
        } else {
          // 空行をカウント
          consecutiveEmptyRows++;
          // 連続して空行が一定数続いたら終了
          if (consecutiveEmptyRows >= maxConsecutiveEmptyRows) {
            break;
          }
        }
      }
      
      return categories;
    } catch (error) {
      console.error(`案件リストファイルの読み込みエラー: ${error.message}`);
      return [];
    }
  }
}

export default ExcelService;
