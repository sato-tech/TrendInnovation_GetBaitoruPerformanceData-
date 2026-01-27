import { google } from 'googleapis';
import config from '../../config/config.js';

/**
 * Google Sheets API操作を担当するサービスクラス
 */
class GoogleSheetsService {
  constructor() {
    this.auth = null;
    this.sheets = null;
    this.initialized = false;
  }

  /**
   * Google Sheets APIを初期化する
   */
  async initialize() {
    if (this.initialized) {
      return;
    }

    try {
      // サービスアカウント認証
      const credentialsPath = process.env.GOOGLE_SERVICE_ACCOUNT_KEY_PATH || '';
      if (!credentialsPath) {
        throw new Error('GOOGLE_SERVICE_ACCOUNT_KEY_PATHが設定されていません');
      }

      // サービスアカウントキーファイルを読み込む
      const { readFileSync } = await import('fs');
      const credentials = JSON.parse(readFileSync(credentialsPath, 'utf-8'));

      // 認証クライアントを作成
      this.auth = new google.auth.GoogleAuth({
        credentials: credentials,
        scopes: ['https://www.googleapis.com/auth/spreadsheets'],
      });

      const authClient = await this.auth.getClient();
      this.sheets = google.sheets({ version: 'v4', auth: authClient });

      this.initialized = true;
      console.log('✓ Google Sheets APIを初期化しました');
    } catch (error) {
      console.error('❌ Google Sheets APIの初期化エラー:', error.message);
      throw error;
    }
  }

  /**
   * 列名（A, B, C...）を列番号（1, 2, 3...）に変換する
   * @param {string} column - 列名（例: 'A', 'B', 'AA'）
   * @returns {number} 列番号（1から始まる）
   */
  columnToNumber(column) {
    let result = 0;
    for (let i = 0; i < column.length; i++) {
      result = result * 26 + (column.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
    }
    return result;
  }

  /**
   * スプレッドシートのシート名一覧を取得する
   * @param {string} spreadsheetId - スプレッドシートID
   * @returns {Promise<Array<string>>} シート名の配列
   */
  async getSheetNames(spreadsheetId) {
    if (!this.initialized) {
      await this.initialize();
    }

    try {
      const response = await this.sheets.spreadsheets.get({
        spreadsheetId: spreadsheetId,
      });
      
      const sheets = response.data.sheets || [];
      return sheets.map(sheet => sheet.properties.title);
    } catch (error) {
      console.error(`❌ getSheetNamesエラー: ${error.message}`);
      return [];
    }
  }

  /**
   * 実際に使用するシート名を取得する（存在しない場合は最初のシートを使用）
   * @param {string} spreadsheetId - スプレッドシートID
   * @param {string} preferredSheetName - 希望するシート名
   * @returns {Promise<string>} 実際に使用するシート名
   */
  async getActualSheetName(spreadsheetId, preferredSheetName) {
    const sheetNames = await this.getSheetNames(spreadsheetId);
    
    if (sheetNames.length === 0) {
      throw new Error('スプレッドシートにシートが存在しません');
    }
    
    // 希望するシート名が存在する場合はそれを使用
    if (sheetNames.includes(preferredSheetName)) {
      return preferredSheetName;
    }
    
    // 存在しない場合は最初のシートを使用
    console.log(`⚠️  シート名「${preferredSheetName}」が見つかりません。最初のシート「${sheetNames[0]}」を使用します。`);
    return sheetNames[0];
  }

  /**
   * セルに値を設定する
   * @param {string} spreadsheetId - スプレッドシートID
   * @param {string} sheetName - シート名
   * @param {string} column - 列（例: 'A', 'B'）
   * @param {number} row - 行番号
   * @param {any} value - 設定する値
   */
  async setCellValue(spreadsheetId, sheetName, column, row, value) {
    if (!this.initialized) {
      await this.initialize();
    }

    try {
      // 実際に使用するシート名を取得
      const actualSheetName = await this.getActualSheetName(spreadsheetId, sheetName);
      const range = `${actualSheetName}!${column}${row}`;

      await this.sheets.spreadsheets.values.update({
        spreadsheetId: spreadsheetId,
        range: range,
        valueInputOption: 'USER_ENTERED',
        resource: {
          values: [[value]],
        },
      });

      // デバッグ: 重要な列のみログ出力（K列とL列を追加）
      if (['A', 'B', 'C', 'E', 'F', 'G', 'H', 'I', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'Z', 'AA', 'AB', 'AC', 'AD'].includes(column)) {
        console.log(`    [DEBUG] スプレッドシート書き込み: ${column}${row} = ${value}`);
      }
    } catch (error) {
      console.error(`❌ setCellValueエラー (列: ${column}, 行: ${row}): ${error.message}`);
      throw error;
    }
  }

  /**
   * 複数のセルに値を一括で設定する（パフォーマンス向上のため）
   * @param {string} spreadsheetId - スプレッドシートID
   * @param {string} sheetName - シート名
   * @param {Array<{column: string, row: number, value: any}>} cells - セルデータの配列
   */
  async setCellValuesBatch(spreadsheetId, sheetName, cells) {
    if (!this.initialized) {
      await this.initialize();
    }

    try {
      // 実際に使用するシート名を取得
      const actualSheetName = await this.getActualSheetName(spreadsheetId, sheetName);
      
      const updates = cells.map(cell => ({
        range: `${actualSheetName}!${cell.column}${cell.row}`,
        values: [[cell.value]],
      }));

      const data = updates.map(update => ({
        range: update.range,
        values: update.values,
      }));

      await this.sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: spreadsheetId,
        resource: {
          valueInputOption: 'USER_ENTERED',
          data: data,
        },
      });

      console.log(`  ✓ ${cells.length}個のセルを一括書き込みしました`);
    } catch (error) {
      console.error(`❌ setCellValuesBatchエラー: ${error.message}`);
      throw error;
    }
  }

  /**
   * セルの値を取得する
   * @param {string} spreadsheetId - スプレッドシートID
   * @param {string} sheetName - シート名
   * @param {string} column - 列（例: 'A', 'B'）
   * @param {number} row - 行番号
   * @returns {Promise<any>} セルの値
   */
  async getCellValue(spreadsheetId, sheetName, column, row) {
    if (!this.initialized) {
      await this.initialize();
    }

    try {
      // 実際に使用するシート名を取得
      const actualSheetName = await this.getActualSheetName(spreadsheetId, sheetName);
      const range = `${actualSheetName}!${column}${row}`;
      const response = await this.sheets.spreadsheets.values.get({
        spreadsheetId: spreadsheetId,
        range: range,
      });

      const values = response.data.values;
      if (values && values.length > 0 && values[0].length > 0) {
        return values[0][0];
      }
      return null;
    } catch (error) {
      console.error(`❌ getCellValueエラー (列: ${column}, 行: ${row}): ${error.message}`);
      return null;
    }
  }

  /**
   * 最初の空白行を見つける
   * @param {string} spreadsheetId - スプレッドシートID
   * @param {string} sheetName - シート名
   * @param {string} checkColumn - チェックする列（デフォルト: 'A'、年列）
   * @param {number} startRow - 検索開始行（デフォルト: 2、ヘッダー行を除く）
   * @returns {Promise<number>} 最初の空白行の行番号
   */
  async findFirstEmptyRow(spreadsheetId, sheetName, checkColumn = 'A', startRow = 2) {
    if (!this.initialized) {
      await this.initialize();
    }

    try {
      // 実際に使用するシート名を取得
      const actualSheetName = await this.getActualSheetName(spreadsheetId, sheetName);
      
      // 一度に100行ずつチェック
      let currentRow = startRow;
      const batchSize = 100;

      while (true) {
        const range = `${actualSheetName}!${checkColumn}${currentRow}:${checkColumn}${currentRow + batchSize - 1}`;
        const response = await this.sheets.spreadsheets.values.get({
          spreadsheetId: spreadsheetId,
          range: range,
        });

        const values = response.data.values || [];
        
        // 空白行を探す
        for (let i = 0; i < batchSize; i++) {
          const rowIndex = currentRow + i;
          const value = values[i] && values[i][0] ? values[i][0] : null;
          
          if (!value || value === '' || value === null || value === undefined) {
            return rowIndex;
          }
        }

        // すべての行にデータがある場合は次のバッチをチェック
        currentRow += batchSize;
      }
    } catch (error) {
      console.error(`❌ findFirstEmptyRowエラー: ${error.message}`);
      return startRow;
    }
  }

  /**
   * 列番号を列名（A, B, C...）に変換する
   * @param {number} columnNumber - 列番号（1から始まる）
   * @returns {string} 列名（例: 'A', 'B', 'AA'）
   */
  numberToColumn(columnNumber) {
    let result = '';
    let num = columnNumber;
    while (num > 0) {
      num--;
      result = String.fromCharCode(65 + (num % 26)) + result;
      num = Math.floor(num / 26);
    }
    return result;
  }

  /**
   * スプレッドシートの最後の列を取得する
   * @param {string} spreadsheetId - スプレッドシートID
   * @param {string} sheetName - シート名
   * @returns {Promise<string>} 最後の列名（例: 'Z', 'AA'）
   */
  async getLastColumn(spreadsheetId, sheetName) {
    if (!this.initialized) {
      await this.initialize();
    }

    try {
      const actualSheetName = await this.getActualSheetName(spreadsheetId, sheetName);
      
      // 1行目（ヘッダー行）の全列を取得
      const response = await this.sheets.spreadsheets.values.get({
        spreadsheetId: spreadsheetId,
        range: `${actualSheetName}!1:1`,
      });

      const values = response.data.values;
      if (values && values.length > 0 && values[0]) {
        const lastColumnIndex = values[0].length;
        return this.numberToColumn(lastColumnIndex);
      }
      
      // データがない場合はA列を返す
      return 'A';
    } catch (error) {
      console.error(`❌ getLastColumnエラー: ${error.message}`);
      return 'A';
    }
  }

  /**
   * スプレッドシートに列を追加する
   * @param {string} spreadsheetId - スプレッドシートID
   * @param {string} sheetName - シート名
   * @param {number} columnIndex - 追加する列のインデックス（0から始まる、既存の列の後ろに追加）
   * @param {string} headerValue - ヘッダー行に設定する値
   * @returns {Promise<string>} 追加された列名
   */
  async addColumn(spreadsheetId, sheetName, columnIndex, headerValue = '') {
    if (!this.initialized) {
      await this.initialize();
    }

    try {
      const actualSheetName = await this.getActualSheetName(spreadsheetId, sheetName);
      
      // シートIDを取得
      const sheetInfo = await this.sheets.spreadsheets.get({
        spreadsheetId: spreadsheetId,
      });
      
      const sheet = sheetInfo.data.sheets.find(s => s.properties.title === actualSheetName);
      if (!sheet) {
        throw new Error(`シート「${actualSheetName}」が見つかりません`);
      }
      
      const sheetId = sheet.properties.sheetId;
      
      // 列を追加
      // columnIndexは0ベースで、新しい列を挿入する位置を指定
      // シートのグリッドサイズを取得して確認
      const gridProperties = sheet.properties.gridProperties;
      const currentColumnCount = gridProperties.columnCount || 0;
      
      // insertDimensionのstartIndexは、新しい列を挿入する位置を指定（0ベース）
      // inheritFromBeforeがfalseの場合、startIndexはグリッドサイズ未満である必要がある
      // 最後の列の後に追加する場合は、最後の列のインデックス（currentColumnCount - 1）を指定する
      // これにより、最後の列の後に新しい列が追加される
      let startIndex = columnIndex;
      if (startIndex >= currentColumnCount) {
        // グリッドサイズを超える場合は、最後の列のインデックスを使用
        startIndex = Math.max(0, currentColumnCount - 1);
      }
      
      await this.sheets.spreadsheets.batchUpdate({
        spreadsheetId: spreadsheetId,
        resource: {
          requests: [
            {
              insertDimension: {
                range: {
                  sheetId: sheetId,
                  dimension: 'COLUMNS',
                  startIndex: startIndex + 1, // 最後の列の後に追加する場合は、最後の列のインデックス+1
                  endIndex: startIndex + 2,
                },
                inheritFromBefore: false,
              },
            },
          ],
        },
      });

      // ヘッダー行に値を設定
      // 列を追加した後、新しい列のインデックスはstartIndexになる
      const newColumnIndex = startIndex + 1; // 1ベースに変換
      if (headerValue) {
        const newColumnName = this.numberToColumn(newColumnIndex);
        await this.setCellValue(spreadsheetId, sheetName, newColumnName, 1, headerValue);
      }

      const newColumnName = this.numberToColumn(newColumnIndex);
      console.log(`✓ 列を追加しました: ${newColumnName}`);
      return newColumnName;
    } catch (error) {
      console.error(`❌ addColumnエラー: ${error.message}`);
      throw error;
    }
  }

  /**
   * 指定したヘッダー名の列を検索する（列を追加しない）
   * @param {string} spreadsheetId - スプレッドシートID
   * @param {string} sheetName - シート名
   * @param {string} headerValue - 検索するヘッダー名
   * @returns {Promise<string|null>} 列名（見つからない場合はnull）
   */
  async findColumnByName(spreadsheetId, sheetName, headerValue) {
    if (!this.initialized) {
      await this.initialize();
    }

    try {
      const actualSheetName = await this.getActualSheetName(spreadsheetId, sheetName);
      
      // 1行目（ヘッダー行）の全列を取得
      const response = await this.sheets.spreadsheets.values.get({
        spreadsheetId: spreadsheetId,
        range: `${actualSheetName}!1:1`,
      });

      const values = response.data.values;
      if (values && values.length > 0 && values[0]) {
        const headerRow = values[0];
        const existingIndex = headerRow.findIndex(cell => cell === headerValue);
        
        if (existingIndex !== -1) {
          // 見つかった場合はその列名を返す
          const columnName = this.numberToColumn(existingIndex + 1);
          return columnName;
        }
      }
      
      // 見つからない場合はnullを返す
      return null;
    } catch (error) {
      console.error(`❌ findColumnByNameエラー: ${error.message}`);
      return null;
    }
  }

  /**
   * 重複チェック用の列を追加する（既に存在する場合は追加しない）
   * @param {string} spreadsheetId - スプレッドシートID
   * @param {string} sheetName - シート名
   * @param {string} headerValue - ヘッダー行に設定する値（デフォルト: '重複チェック'）
   * @returns {Promise<string>} 重複チェック用の列名
   */
  async ensureDuplicateCheckColumn(spreadsheetId, sheetName, headerValue = '重複チェック') {
    if (!this.initialized) {
      await this.initialize();
    }

    try {
      const actualSheetName = await this.getActualSheetName(spreadsheetId, sheetName);
      
      // 1行目（ヘッダー行）の全列を取得
      const response = await this.sheets.spreadsheets.values.get({
        spreadsheetId: spreadsheetId,
        range: `${actualSheetName}!1:1`,
      });

      const values = response.data.values;
      if (values && values.length > 0 && values[0]) {
        // 既に「重複チェック」列が存在するかチェック
        const headerRow = values[0];
        const existingIndex = headerRow.findIndex(cell => cell === headerValue);
        
        if (existingIndex !== -1) {
          // 既に存在する場合はその列名を返す
          const columnName = this.numberToColumn(existingIndex + 1);
          console.log(`✓ 重複チェック列は既に存在します: ${columnName}`);
          return columnName;
        }
      }
      
      // 存在しない場合は最後の列の後に追加
      const lastColumn = await this.getLastColumn(spreadsheetId, sheetName);
      const lastColumnNumber = this.columnToNumber(lastColumn);
      // columnToNumberは1ベース（A=1）を返すが、addColumnのinsertDimensionは0ベースのインデックスを期待する
      // 最後の列の後に追加する場合は、lastColumnNumber（1ベース）を0ベースに変換する必要がある
      // ただし、insertDimensionのstartIndexは新しい列を挿入する位置なので、最後の列のインデックスを指定する
      const columnIndexForInsert = lastColumnNumber - 1; // 1ベースから0ベースに変換
      const newColumnName = await this.addColumn(spreadsheetId, sheetName, columnIndexForInsert, headerValue);
      
      return newColumnName;
    } catch (error) {
      console.error(`❌ ensureDuplicateCheckColumnエラー: ${error.message}`);
      throw error;
    }
  }

  /**
   * 指定した列の値を検索する
   * @param {string} spreadsheetId - スプレッドシートID
   * @param {string} sheetName - シート名
   * @param {string} column - 検索する列名
   * @param {any} searchValue - 検索する値
   * @param {number} startRow - 検索開始行（デフォルト: 2、ヘッダー行を除く）
   * @returns {Promise<boolean>} 値が存在するかどうか
   */
  async findValueInColumn(spreadsheetId, sheetName, column, searchValue, startRow = 2) {
    if (!this.initialized) {
      await this.initialize();
    }

    try {
      const actualSheetName = await this.getActualSheetName(spreadsheetId, sheetName);
      
      // 一度に100行ずつチェック
      let currentRow = startRow;
      const batchSize = 100;

      while (true) {
        const range = `${actualSheetName}!${column}${currentRow}:${column}${currentRow + batchSize - 1}`;
        const response = await this.sheets.spreadsheets.values.get({
          spreadsheetId: spreadsheetId,
          range: range,
        });

        const values = response.data.values || [];
        
        // 値を検索
        for (let i = 0; i < values.length; i++) {
          const cellValue = values[i] && values[i][0] ? values[i][0] : null;
          
          if (cellValue === searchValue || String(cellValue) === String(searchValue)) {
            return true;
          }
        }

        // すべての行をチェックした場合、次のバッチをチェック
        if (values.length < batchSize) {
          break;
        }
        
        currentRow += batchSize;
      }
      
      return false;
    } catch (error) {
      console.error(`❌ findValueInColumnエラー: ${error.message}`);
      return false;
    }
  }
}

export default GoogleSheetsService;
