/**
 * 企業IDを引数で受け取って動作確認するスクリプト
 * 使用方法: npm run test:company <企業ID> [開始日] [終了日] [掲載区分]
 * 例: npm run test:company 70687 2025-08-25 2025-09-21 通常
 * 例: npm run test:company 260342 2025-08-11 2025-09-07 ナイト
 * 
 * エクセルデータを読み込まずに、直接企業IDを指定して処理を実行します
 * 企業ID入力以降の処理は既存のtest-loop.jsと共通です
 */

import ExcelService from './services/ExcelService.js';
import ScrapingService from './services/ScrapingService.js';
import GoogleSheetsService from './services/GoogleSheetsService.js';
import AIService from './services/AIService.js';
import config from '../config/config.js';
import { excelDateToJSDate, formatDateForInput } from './utils/dateUtils.js';
import { promises as fs } from 'fs';
import { join } from 'path';
import { testSingleCompany } from './test-loop.js';

/**
 * メイン処理：企業IDを引数で受け取って処理を実行
 */
async function main() {
  // コマンドライン引数を取得
  const args = process.argv.slice(2);
  
  if (args.length < 1) {
    console.error('❌ エラー: 企業IDを指定してください。');
    console.log('使用方法: npm run test:company <企業ID> [開始日] [終了日] [掲載区分]');
    console.log('例: npm run test:company 70687 2025-08-25 2025-09-21 通常');
    console.log('例: npm run test:company 260342 2025-08-11 2025-09-07 ナイト');
    console.log('\n引数:');
    console.log('  企業ID: 必須（例: 70687）');
    console.log('  開始日: オプション（YYYY-MM-DD形式、例: 2025-08-25）');
    console.log('  終了日: オプション（YYYY-MM-DD形式、例: 2025-09-21）');
    console.log('  掲載区分: オプション（「ナイト」または「通常」、デフォルト: 通常）');
    process.exit(1);
  }

  const companyId = args[0];
  const startDateStr = args[1] || null;
  const endDateStr = args[2] || null;
  const publishingCategory = args[3] || '通常';

  // 企業IDの検証
  if (!companyId || isNaN(parseInt(companyId, 10))) {
    console.error('❌ エラー: 有効な企業IDを指定してください。');
    process.exit(1);
  }

  // 掲載区分の検証
  if (publishingCategory !== 'ナイト' && publishingCategory !== '通常') {
    console.error('❌ エラー: 掲載区分は「ナイト」または「通常」を指定してください。');
    process.exit(1);
  }

  // 日付の検証と変換
  let startDate = null;
  let endDate = null;

  if (startDateStr && endDateStr) {
    // 文字列からDateオブジェクトに変換
    startDate = new Date(startDateStr);
    endDate = new Date(endDateStr);

    if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
      console.error('❌ エラー: 開始日・終了日はYYYY-MM-DD形式で指定してください。');
      console.log('例: 2025-08-25');
      process.exit(1);
    }
  } else if (startDateStr || endDateStr) {
    console.error('❌ エラー: 開始日と終了日の両方を指定してください。');
    process.exit(1);
  } else {
    // 日付が指定されていない場合は、現在の日付から1ヶ月前と現在をデフォルトとする
    const now = new Date();
    endDate = new Date(now);
    startDate = new Date(now);
    startDate.setMonth(startDate.getMonth() - 1);
    console.log('⚠️  開始日・終了日が指定されていません。デフォルト値を使用します。');
    console.log(`   開始日: ${formatDateForInput(startDate)}`);
    console.log(`   終了日: ${formatDateForInput(endDate)}\n`);
  }

  const scrapingService = new ScrapingService();
  const excelService = new ExcelService();
  const googleSheetsService = new GoogleSheetsService();
  const aiService = new AIService();
  
  // Google Sheets APIを初期化（必須）
  try {
    await googleSheetsService.initialize();
    console.log('✓ Google Sheets APIを初期化しました');
  } catch (error) {
    console.error('❌ Google Sheets APIの初期化に失敗しました:', error.message);
    console.error('スプレッドシートIDとサービスアカウントキーの設定を確認してください。');
    process.exit(1);
  }

  try {
    console.log(`\n${'='.repeat(60)}`);
    console.log(`=== 企業ID: ${companyId} の動作確認を開始します ===`);
    console.log(`${'='.repeat(60)}\n`);

    // ブラウザ起動とログイン
    console.log('1. ブラウザを起動中...');
    await scrapingService.launchBrowser();
    console.log('✓ ブラウザを起動しました\n');

    console.log('2. ログイン中...');
    try {
      await scrapingService.login();
      console.log('✓ ログインしました\n');
    } catch (error) {
      console.error(`❌ ログインエラー: ${error.message}`);
      console.log('   環境変数（BAITORU_LOGIN_URL, BAITORU_USERNAME, BAITORU_PASSWORD）を確認してください。\n');
      await scrapingService.closeBrowser();
      process.exit(1);
    }

    console.log('3. TOPページに移動中...');
    try {
      await scrapingService.goToTop();
      console.log('✓ TOPページに移動しました（または既にTOPページにいます）\n');
    } catch (error) {
      console.warn(`⚠️  TOPページ移動で警告: ${error.message}`);
      console.log('（既にTOPページにいる可能性があります。続行します...）\n');
    }

    // 実行ごとのダウンロードフォルダを作成
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, -5); // YYYY-MM-DDTHH-MM-SS形式
    const downloadFolderName = `downloads_${timestamp}`;
    const downloadFolderPath = join(config.files.downloadDir, downloadFolderName);
    
    await fs.mkdir(downloadFolderPath, { recursive: true });
    console.log(`✓ ダウンロードフォルダを作成しました: ${downloadFolderPath}\n`);
    
    // ダウンロードフォルダをScrapingServiceに設定
    scrapingService.setDownloadFolder(downloadFolderPath);

    // 案件リストファイルを読み込み（職種カテゴリ用 - JSON形式）
    let nightJobCategories = [];
    let normalJobCategories = [];
    try {
      nightJobCategories = await excelService.loadJobCategoriesFromJSON(
        config.files.jobCategoryListNight
      );
      console.log(`✓ ナイト案件リストを読み込みました（${nightJobCategories.length}件）`);
    } catch (error) {
      console.warn(`⚠️  ナイト案件リストの読み込みエラー: ${error.message}`);
    }
    
    try {
      normalJobCategories = await excelService.loadJobCategoriesFromJSON(
        config.files.jobCategoryListNormal
      );
      console.log(`✓ 通常案件リストを読み込みました（${normalJobCategories.length}件）`);
    } catch (error) {
      console.warn(`⚠️  通常案件リストの読み込みエラー: ${error.message}`);
    }

    // 企業データを構築（エクセルから読み込まずに直接指定）
    const companyData = {
      companyId: companyId,
      companyName: '', // 企業名は取得できないため空文字
      publishingCategory: publishingCategory,
      startDate: startDate,
      endDate: endDate,
      row: null // エクセル行番号は不要
    };

    // test-loop.jsからtestSingleCompany関数をインポートする必要がある
    // 関数をエクスポートする必要があるため、test-loop.jsを修正するか、
    // ここで同じ処理を実装する必要がある
    
    // とりあえず、test-loop.jsのtestSingleCompany関数を直接呼び出すために
    // 動的インポートを使用するか、関数を共通モジュールに分離する必要がある
    
    // 簡易版として、test-loop.jsのtestSingleCompany関数をコピーして使用
    // または、test-loop.jsを修正して関数をエクスポートする
    
    // ここでは、test-loop.jsのtestSingleCompany関数を直接使用するために
    // 動的インポートを試みる（ただし、test-loop.jsは関数をエクスポートしていない）
    
    // 代替案：testSingleCompany関数を共通モジュールに分離する
    // または、ここで同じ処理を実装する
    
    // 一旦、test-loop.jsを修正してtestSingleCompany関数をエクスポートする方法を採用
    // または、ここで同じ処理を実装する
    
    console.log('4. 企業データを準備しました');
    console.log(`   企業ID: ${companyData.companyId}`);
    console.log(`   掲載区分: ${companyData.publishingCategory}`);
    console.log(`   開始日: ${formatDateForInput(companyData.startDate)}`);
    console.log(`   終了日: ${formatDateForInput(companyData.endDate)}\n`);

    // 失敗ログを記録する配列
    const failureLogs = [];

    // 失敗ログを記録するコールバック関数
    const recordFailure = (companyId, companyName, jobNo, reason) => {
      failureLogs.push({
        companyId: companyId || '',
        companyName: companyName || '',
        jobNo: jobNo || '',
        reason: reason || '',
        timestamp: new Date().toISOString()
      });
    };

    // testSingleCompany関数を呼び出し（エクセル読み込みを行わないため、inputSheetはnull）
    const success = await testSingleCompany(
      companyData,
      1, // loopIndex
      1, // totalLoops
      scrapingService,
      excelService,
      googleSheetsService,
      aiService,
      nightJobCategories,
      normalJobCategories,
      null, // inputSheet（エクセル読み込みを行わないためnull）
      recordFailure
    );

    // ブラウザを閉じる
    await scrapingService.closeBrowser();
    console.log('ブラウザを閉じました\n');

    // 結果を表示
    console.log(`\n${'='.repeat(60)}`);
    console.log('=== 動作確認結果 ===');
    console.log(`${'='.repeat(60)}`);
    if (success) {
      console.log('✓ 処理が正常に完了しました');
    } else {
      console.log('❌ 処理が失敗しました');
    }
    console.log(`${'='.repeat(60)}\n`);

    // 失敗ログをテキストファイルに出力
    if (failureLogs.length > 0) {
      const failureLogPath = join(downloadFolderPath, `failure_log_${timestamp}.txt`);
      let logContent = '='.repeat(80) + '\n';
      logContent += '取得失敗ログ\n';
      logContent += `実行日時: ${new Date().toLocaleString('ja-JP')}\n`;
      logContent += `総失敗数: ${failureLogs.length}件\n`;
      logContent += '='.repeat(80) + '\n\n';

      failureLogs.forEach((log, index) => {
        logContent += `【失敗 ${index + 1}】\n`;
        logContent += `企業ID: ${log.companyId}\n`;
        logContent += `企業名: ${log.companyName}\n`;
        logContent += `仕事No: ${log.jobNo || '(取得できず)'}\n`;
        logContent += `失敗理由: ${log.reason}\n`;
        logContent += `発生時刻: ${new Date(log.timestamp).toLocaleString('ja-JP')}\n`;
        logContent += '-'.repeat(80) + '\n\n';
      });

      try {
        await fs.writeFile(failureLogPath, logContent, 'utf-8');
        console.log(`✓ 失敗ログを保存しました: ${failureLogPath}\n`);
      } catch (writeError) {
        console.error(`❌ 失敗ログの保存エラー: ${writeError.message}`);
      }
    } else {
      console.log('✓ 失敗はありませんでした。\n');
    }

    // 失敗があった場合は終了コード1を返す
    if (!success) {
      process.exit(1);
    }
  } catch (error) {
    console.error('致命的なエラー:', error);
    try {
      await scrapingService.closeBrowser();
    } catch (closeError) {
      // ブラウザが既に閉じられている可能性があるため、エラーを無視
    }
    process.exit(1);
  }
}

// 実行
main().catch(error => {
  console.error('致命的なエラー:', error);
  process.exit(1);
});
