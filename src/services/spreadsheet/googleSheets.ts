import { google } from 'googleapis';
import fs from 'fs';
import { logger } from '../../utils/logger';
import { Config } from '../../config';

export class GoogleSheetsService {
  private auth: any;
  private sheets: any;

  async initialize(config: Config): Promise<void> {
    try {
      if (!fs.existsSync(config.google.credentialsPath)) {
        throw new Error(`認証情報ファイルが見つかりません: ${config.google.credentialsPath}`);
      }

      const credentials = JSON.parse(fs.readFileSync(config.google.credentialsPath, 'utf8'));
      
      const { client_secret, client_id, redirect_uris } = credentials.installed || credentials.web;
      const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);

      // トークンファイルが存在する場合は読み込み
      if (fs.existsSync(config.google.tokenPath)) {
        const token = JSON.parse(fs.readFileSync(config.google.tokenPath, 'utf8'));
        oAuth2Client.setCredentials(token);
      } else {
        throw new Error(`トークンファイルが見つかりません: ${config.google.tokenPath}\n認証が必要です。`);
      }

      this.auth = oAuth2Client;
      this.sheets = google.sheets({ version: 'v4', auth: oAuth2Client });
      
      logger.info('Google Sheets APIを初期化しました');
    } catch (error) {
      logger.error('Google Sheets APIの初期化に失敗しました', error);
      throw error;
    }
  }

  async readSearchKeys(config: Config): Promise<string[]> {
    try {
      if (!this.sheets) {
        await this.initialize(config);
      }

      logger.info(`スプレッドシートから読み取り: ${config.spreadsheet.id}, 範囲: ${config.spreadsheet.range}`);
      
      const response = await this.sheets.spreadsheets.values.get({
        spreadsheetId: config.spreadsheet.id,
        range: config.spreadsheet.range,
      });

      const rows = response.data.values || [];
      if (rows.length === 0) {
        logger.warn('データが見つかりませんでした');
        return [];
      }

      // 検索キー列を取得（列名から列インデックスを計算）
      const headerRow = rows[0];
      const columnIndex = this.getColumnIndex(config.spreadsheet.searchKeyColumn);
      const dataRows = rows.slice(1); // ヘッダー行をスキップ

      const searchKeys = dataRows
        .map((row: any[]) => row[columnIndex])
        .filter((key: any) => key && key.toString().trim() !== '');

      logger.info(`${searchKeys.length}件の検索キーを取得しました`);
      return searchKeys;
    } catch (error) {
      logger.error('検索キーの読み取りに失敗しました', error);
      throw error;
    }
  }

  async appendResult(data: Record<string, any>, config: Config): Promise<void> {
    try {
      if (!this.sheets) {
        await this.initialize(config);
      }

      if (!config.spreadsheet.outputSheetId) {
        logger.warn('出力先スプレッドシートIDが設定されていないため、スキップします');
        return;
      }

      // データを配列形式に変換
      const values = [Object.values(data)];

      logger.info(`スプレッドシートに追記: ${config.spreadsheet.outputSheetId}`);

      await this.sheets.spreadsheets.values.append({
        spreadsheetId: config.spreadsheet.outputSheetId,
        range: config.spreadsheet.outputSheetRange,
        valueInputOption: 'USER_ENTERED',
        insertDataOption: 'INSERT_ROWS',
        requestBody: {
          values: values,
        },
      });

      logger.info('スプレッドシートへの追記が完了しました');
    } catch (error) {
      logger.error('スプレッドシートへの追記に失敗しました', error);
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

export default GoogleSheetsService;
