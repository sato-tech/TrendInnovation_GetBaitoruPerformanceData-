/**
 * Google Sheets APIの動作確認用スクリプト
 */
import GoogleSheetsService from './src/services/GoogleSheetsService.js';
import config from './config/config.js';
import dotenv from 'dotenv';

dotenv.config();

async function testGoogleSheets() {
  console.log('=== Google Sheets API動作確認 ===\n');
  
  const googleSheetsService = new GoogleSheetsService();
  
  try {
    // Google Sheets APIを初期化
    console.log('1. Google Sheets APIを初期化中...');
    await googleSheetsService.initialize();
    console.log('✓ 初期化完了\n');
    
    // ナイト案件スプレッドシートのテスト
    if (config.googleSheets.spreadsheetIdNight) {
      console.log('2. ナイト案件スプレッドシートのテスト...');
      console.log(`   スプレッドシートID: ${config.googleSheets.spreadsheetIdNight}`);
      console.log(`   希望シート名: ${config.googleSheets.sheetName}`);
      
      // シート名一覧を取得
      const sheetNames = await googleSheetsService.getSheetNames(config.googleSheets.spreadsheetIdNight);
      console.log(`   利用可能なシート名: ${sheetNames.join(', ')}`);
      
      // 実際に使用するシート名を取得
      const actualSheetName = await googleSheetsService.getActualSheetName(
        config.googleSheets.spreadsheetIdNight,
        config.googleSheets.sheetName
      );
      console.log(`   使用するシート名: ${actualSheetName}`);
      
      // 空白行を検索
      const emptyRowNight = await googleSheetsService.findFirstEmptyRow(
        config.googleSheets.spreadsheetIdNight,
        config.googleSheets.sheetName,
        'A',
        2
      );
      console.log(`   空白行: ${emptyRowNight}`);
      
      // テストデータを書き込み（A列に現在時刻）
      const testValue = `テスト書き込み ${new Date().toISOString()}`;
      await googleSheetsService.setCellValue(
        config.googleSheets.spreadsheetIdNight,
        config.googleSheets.sheetName,
        'A',
        emptyRowNight,
        testValue
      );
      console.log(`   ✓ テストデータを書き込みました: ${testValue}\n`);
    } else {
      console.log('⚠️  ナイト案件のスプレッドシートIDが設定されていません\n');
    }
    
    // 通常案件スプレッドシートのテスト
    if (config.googleSheets.spreadsheetIdNormal) {
      console.log('3. 通常案件スプレッドシートのテスト...');
      console.log(`   スプレッドシートID: ${config.googleSheets.spreadsheetIdNormal}`);
      console.log(`   希望シート名: ${config.googleSheets.sheetName}`);
      
      // シート名一覧を取得
      const sheetNames = await googleSheetsService.getSheetNames(config.googleSheets.spreadsheetIdNormal);
      console.log(`   利用可能なシート名: ${sheetNames.join(', ')}`);
      
      // 実際に使用するシート名を取得
      const actualSheetName = await googleSheetsService.getActualSheetName(
        config.googleSheets.spreadsheetIdNormal,
        config.googleSheets.sheetName
      );
      console.log(`   使用するシート名: ${actualSheetName}`);
      
      // 空白行を検索
      const emptyRowNormal = await googleSheetsService.findFirstEmptyRow(
        config.googleSheets.spreadsheetIdNormal,
        config.googleSheets.sheetName,
        'A',
        2
      );
      console.log(`   空白行: ${emptyRowNormal}`);
      
      // テストデータを書き込み（A列に現在時刻）
      const testValue = `テスト書き込み ${new Date().toISOString()}`;
      await googleSheetsService.setCellValue(
        config.googleSheets.spreadsheetIdNormal,
        config.googleSheets.sheetName,
        'A',
        emptyRowNormal,
        testValue
      );
      console.log(`   ✓ テストデータを書き込みました: ${testValue}\n`);
    } else {
      console.log('⚠️  通常案件のスプレッドシートIDが設定されていません\n');
    }
    
    console.log('=== 動作確認完了 ===');
    console.log('スプレッドシートを確認して、テストデータが書き込まれているか確認してください。');
    
  } catch (error) {
    console.error('❌ エラーが発生しました:', error.message);
    console.error('スタックトレース:', error.stack);
    process.exit(1);
  }
}

testGoogleSheets();
