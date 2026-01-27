/**
 * テスト表用スクリプト
 * 使用方法: npm run test:table <企業ID> <掲載開始日> <掲載終了日> <ループ回数>
 * 例: npm run test:table 70687 2025/08/25 2025/09/21 5
 * 
 * 指定された企業ID、掲載開始日、掲載終了日に一致する行を見つけ、
 * その行からループ回数分、次の行を処理します
 */

import ExcelService from './services/ExcelService.js';
import ScrapingService from './services/ScrapingService.js';
import GoogleSheetsService from './services/GoogleSheetsService.js';
import AIService from './services/AIService.js';
import config from '../config/config.js';
import { excelDateToJSDate, formatDateForInput, calculateWeeks } from './utils/dateUtils.js';
import FileSelector from './utils/fileSelector.js';
import { promises as fs } from 'fs';
import { join } from 'path';
import { testSingleCompany } from './test-loop.js';

/**
 * メイン処理：指定された条件に一致する行を見つけ、その行からループ回数分処理を繰り返す
 */
async function main() {
  // コマンドライン引数から企業ID、掲載開始日、掲載終了日、ループ回数を取得
  const companyId = process.argv[2];
  const startDateStr = process.argv[3];
  const endDateStr = process.argv[4];
  const loopCount = parseInt(process.argv[5], 10);

  if (!companyId || !startDateStr || !endDateStr || !loopCount || loopCount < 1 || isNaN(loopCount)) {
    console.error('❌ エラー: 有効な引数を指定してください。');
    console.log('使用方法: npm run test:table <企業ID> <掲載開始日> <掲載終了日> <ループ回数>');
    console.log('例: npm run test:table 70687 2025/08/25 2025/09/21 5');
    process.exit(1);
  }

  // 日付の形式を確認（YYYY/MM/DD形式を想定）
  const startDate = excelDateToJSDate(startDateStr);
  const endDate = excelDateToJSDate(endDateStr);

  if (!startDate || !endDate) {
    console.error('❌ エラー: 日付の形式が正しくありません。YYYY/MM/DD形式で指定してください。');
    console.log('例: 2025/08/25');
    process.exit(1);
  }

  const scrapingService = new ScrapingService();
  const excelService = new ExcelService();
  const googleSheetsService = new GoogleSheetsService();
  const aiService = new AIService();
  const fileSelector = new FileSelector();
  
  // 入力ファイルを選択
  let selectedFilePath;
  try {
    console.log('ファイル選択ダイアログを開いています...');
    selectedFilePath = await fileSelector.selectFile();
    console.log(`✓ 選択されたファイル: ${selectedFilePath}`);
    console.log('ファイル選択が完了しました。処理を続行します...\n');
  } catch (error) {
    console.error(`❌ ファイル選択エラー: ${error.message}`);
    console.error('ファイル選択がキャンセルされたか、エラーが発生しました。');
    process.exit(1);
  }
  
  // Google Sheets APIを初期化（必須）
  try {
    console.log('Google Sheets APIを初期化中...');
    await googleSheetsService.initialize();
    console.log('✓ Google Sheets APIを初期化しました\n');
  } catch (error) {
    console.error('❌ Google Sheets APIの初期化に失敗しました:', error.message);
    console.error('スプレッドシートIDとサービスアカウントキーの設定を確認してください。');
    process.exit(1);
  }

  try {
    console.log(`\n${'='.repeat(60)}`);
    console.log(`=== テスト表処理を開始します ===`);
    console.log(`企業ID: ${companyId}`);
    console.log(`掲載開始日: ${startDateStr}`);
    console.log(`掲載終了日: ${endDateStr}`);
    console.log(`ループ回数: ${loopCount}`);
    console.log(`${'='.repeat(60)}\n`);

    // 入力ファイルを読み込み（選択されたファイルパスを使用）
    console.log('1. 入力ファイルを読み込み中...');
    console.log(`   ファイルパス: ${selectedFilePath}`);
    let inputWorkbook;
    let inputSheet;
    try {
      inputWorkbook = await excelService.loadInputFile(selectedFilePath);
      inputSheet = inputWorkbook.getWorksheet(1);
      console.log('✓ 入力ファイルを読み込みました\n');
    } catch (fileError) {
      console.error(`❌ ファイル読み込みエラー: ${fileError.message}`);
      throw fileError;
    }

    // 指定された企業ID、掲載開始日、掲載終了日に一致する行を見つける
    console.log('2. 指定された条件に一致する行を検索中...');
    let targetRow = null;
    const startRow = 2; // データ行の開始（1行目はヘッダー）
    
    // シートの最大行数を取得（適切な範囲で検索）
    let maxRow = 10000; // デフォルトの最大行数
    try {
      // ExcelJSでは実際のデータ行数を取得するのが難しいため、適切な範囲で検索
      for (let row = startRow; row <= maxRow; row++) {
        const rowCompanyId = excelService.getCellValue(
          inputSheet,
          config.excelColumns.inputSheet.companyId,
          row
        );

        // 企業IDが空の場合は検索終了
        if (!rowCompanyId) {
          break;
        }

        // 企業IDが一致するかチェック
        if (String(rowCompanyId) === String(companyId)) {
          // 掲載開始日と終了日を取得
          const rowStartDateValue = excelService.getCellValue(
            inputSheet,
            config.excelColumns.inputSheet.startDate,
            row
          );
          const rowEndDateValue = excelService.getCellValue(
            inputSheet,
            config.excelColumns.inputSheet.endDate,
            row
          );

          // 日付を変換
          const rowStartDate = excelDateToJSDate(rowStartDateValue);
          const rowEndDate = excelDateToJSDate(rowEndDateValue);

          if (rowStartDate && rowEndDate) {
            // 日付を比較（年月日が一致するかチェック）
            const rowStartDateFormatted = formatDateForInput(rowStartDate);
            const rowEndDateFormatted = formatDateForInput(rowEndDate);
            const targetStartDateFormatted = formatDateForInput(startDate);
            const targetEndDateFormatted = formatDateForInput(endDate);

            if (rowStartDateFormatted === targetStartDateFormatted && 
                rowEndDateFormatted === targetEndDateFormatted) {
              targetRow = row;
              console.log(`✓ 一致する行を見つけました: ${row}行目\n`);
              break;
            }
          }
        }
      }
    } catch (error) {
      console.error(`❌ 行の検索中にエラーが発生しました: ${error.message}`);
      process.exit(1);
    }

    if (!targetRow) {
      console.error(`❌ エラー: 指定された条件（企業ID: ${companyId}, 開始日: ${startDateStr}, 終了日: ${endDateStr}）に一致する行が見つかりませんでした。`);
      process.exit(1);
    }

    // 該当行からループ回数分の企業データを取得
    console.log(`3. ${targetRow}行目から${loopCount}件の企業データを取得中...`);
    const companyDataList = [];
    
    for (let i = 0; i < loopCount; i++) {
      const row = targetRow + i;
      
      const rowCompanyId = excelService.getCellValue(
        inputSheet,
        config.excelColumns.inputSheet.companyId,
        row
      );
      
      if (!rowCompanyId) {
        console.warn(`  ⚠️  ${row}行目に企業IDがありません。スキップします。`);
        continue;
      }

      const companyName = excelService.getCellValue(
        inputSheet,
        config.excelColumns.inputSheet.companyName,
        row
      );
      const publishingCategory = excelService.getCellValue(
        inputSheet,
        config.excelColumns.inputSheet.publishingCategory,
        row
      );
      const startDateValue = excelService.getCellValue(
        inputSheet,
        config.excelColumns.inputSheet.startDate,
        row
      );
      const endDateValue = excelService.getCellValue(
        inputSheet,
        config.excelColumns.inputSheet.endDate,
        row
      );

      // 日付を変換
      const rowStartDate = excelDateToJSDate(startDateValue);
      const rowEndDate = excelDateToJSDate(endDateValue);
      
      if (!rowStartDate || !rowEndDate) {
        console.warn(`  ⚠️  ${row}行目の日付変換に失敗しました。スキップします。`);
        continue;
      }

      companyDataList.push({
        companyId: rowCompanyId,
        companyName,
        publishingCategory,
        startDate: rowStartDate,
        endDate: rowEndDate,
        row
      });
    }

    if (companyDataList.length === 0) {
      console.error('❌ 処理可能な企業データがありません。');
      process.exit(1);
    }

    console.log(`✓ ${companyDataList.length}件の企業データを取得しました\n`);

    // ブラウザ起動とログイン（1回のみ）
    console.log('4. ブラウザを起動中...');
    await scrapingService.launchBrowser();
    console.log('✓ ブラウザを起動しました\n');

    console.log('5. ログイン中...');
    try {
      await scrapingService.login();
      console.log('✓ ログインしました\n');
    } catch (error) {
      console.error(`❌ ログインエラー: ${error.message}`);
      console.log('   環境変数（BAITORU_LOGIN_URL, BAITORU_USERNAME, BAITORU_PASSWORD）を確認してください。\n');
      await scrapingService.closeBrowser();
      process.exit(1);
    }

    console.log('6. TOPページに移動中...');
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

    // 各企業データを処理
    const results = {
      success: 0,
      failed: 0,
      total: companyDataList.length
    };

    for (let i = 0; i < companyDataList.length; i++) {
      const companyData = companyDataList[i];
      const loopIndex = i + 1;
      
      const success = await testSingleCompany(
        companyData,
        loopIndex,
        companyDataList.length,
        scrapingService,
        excelService,
        googleSheetsService,
        aiService,
        nightJobCategories,
        normalJobCategories,
        inputSheet,
        recordFailure
      );
      
      if (success) {
        results.success++;
      } else {
        results.failed++;
      }

      // 最後のループでない場合、次のループまでの待機時間を設ける
      if (i < companyDataList.length - 1) {
        console.log(`次の企業の処理まで3秒待機します...\n`);
        await new Promise(resolve => setTimeout(resolve, 3000));
      }
    }

    // ブラウザを閉じる
    await scrapingService.closeBrowser();
    console.log('ブラウザを閉じました\n');

    // 結果サマリーを表示
    console.log(`\n${'='.repeat(60)}`);
    console.log('=== テスト表処理結果サマリー ===');
    console.log(`${'='.repeat(60)}`);
    console.log(`総処理企業数: ${results.total}`);
    console.log(`成功: ${results.success}`);
    console.log(`失敗: ${results.failed}`);
    console.log(`成功率: ${((results.success / results.total) * 100).toFixed(1)}%`);
    console.log(`${'='.repeat(60)}\n`);

    // ログデータをテキストファイルに出力
    const logFilePath = join(downloadFolderPath, `execution_log_${timestamp}.txt`);
    let logContent = '='.repeat(80) + '\n';
    logContent += 'テスト表処理ログ\n';
    logContent += `実行日時: ${new Date().toLocaleString('ja-JP')}\n`;
    logContent += `企業ID: ${companyId}\n`;
    logContent += `掲載開始日: ${startDateStr}\n`;
    logContent += `掲載終了日: ${endDateStr}\n`;
    logContent += `ループ回数: ${loopCount}\n`;
    logContent += `開始行: ${targetRow}\n`;
    logContent += `総処理企業数: ${results.total}\n`;
    logContent += `成功: ${results.success}\n`;
    logContent += `失敗: ${results.failed}\n`;
    logContent += `成功率: ${((results.success / results.total) * 100).toFixed(1)}%\n`;
    logContent += '='.repeat(80) + '\n\n';

    // 処理した企業データの詳細を記録
    companyDataList.forEach((data, index) => {
      logContent += `【企業 ${index + 1}】\n`;
      logContent += `行番号: ${data.row}\n`;
      logContent += `企業ID: ${data.companyId}\n`;
      logContent += `企業名: ${data.companyName || '(空)'}\n`;
      logContent += `掲載区分: ${data.publishingCategory || '(空)'}\n`;
      logContent += `開始日: ${formatDateForInput(data.startDate)}\n`;
      logContent += `終了日: ${formatDateForInput(data.endDate)}\n`;
      logContent += '-'.repeat(80) + '\n\n';
    });

    try {
      await fs.writeFile(logFilePath, logContent, 'utf-8');
      console.log(`✓ ログデータを保存しました: ${logFilePath}\n`);
    } catch (writeError) {
      console.error(`❌ ログデータの保存エラー: ${writeError.message}`);
    }

    // 失敗ログをテキストファイルに出力
    if (failureLogs.length > 0) {
      const failureLogPath = join(downloadFolderPath, `failure_log_${timestamp}.txt`);
      let failureLogContent = '='.repeat(80) + '\n';
      failureLogContent += '取得失敗ログ\n';
      failureLogContent += `実行日時: ${new Date().toLocaleString('ja-JP')}\n`;
      failureLogContent += `総失敗数: ${failureLogs.length}件\n`;
      failureLogContent += '='.repeat(80) + '\n\n';

      failureLogs.forEach((log, index) => {
        failureLogContent += `【失敗 ${index + 1}】\n`;
        failureLogContent += `企業ID: ${log.companyId}\n`;
        failureLogContent += `企業名: ${log.companyName}\n`;
        failureLogContent += `仕事No: ${log.jobNo || '(取得できず)'}\n`;
        failureLogContent += `失敗理由: ${log.reason}\n`;
        failureLogContent += `発生時刻: ${new Date(log.timestamp).toLocaleString('ja-JP')}\n`;
        failureLogContent += '-'.repeat(80) + '\n\n';
      });

      try {
        await fs.writeFile(failureLogPath, failureLogContent, 'utf-8');
        console.log(`✓ 失敗ログを保存しました: ${failureLogPath}\n`);
      } catch (writeError) {
        console.error(`❌ 失敗ログの保存エラー: ${writeError.message}`);
      }
    } else {
      console.log('✓ 失敗はありませんでした。\n');
    }

    // 失敗があった場合は終了コード1を返す
    if (results.failed > 0) {
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
