/**
 * Excelファイルの読み込みのみをテストするスクリプト
 * スクレイピング機能は使用しません
 */

import ExcelJS from 'exceljs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';
import { readFileSync, statSync } from 'fs';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

// 設定を直接読み込み（config.jsを使わない）
const excelColumns = JSON.parse(
  readFileSync(join(__dirname, '..', 'config', 'excelColumns.json'), 'utf-8')
);

const inputFile = '【バイトル】8月実績.xlsx';
const inputFilePath = join(__dirname, '..', inputFile);

/**
 * Excelファイルの読み込みテスト
 */
async function testExcelRead() {
  try {
    console.log('=== Excelファイルの読み込みテスト ===\n');

    // 1. ファイルの存在確認
    console.log(`1. ファイルの存在確認: ${inputFile}`);
    try {
      const stats = statSync(inputFilePath);
      console.log(`   ✓ ファイルが見つかりました（サイズ: ${stats.size} bytes）\n`);
    } catch (error) {
      console.error(`   ❌ ファイルが見つかりません: ${error.message}\n`);
      return;
    }

    // 2. Excelファイルを読み込み
    console.log('2. Excelファイルを読み込み中...');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(inputFilePath);
    const inputSheet = workbook.getWorksheet(1);
    console.log(`   ✓ ファイルを読み込みました（シート名: ${inputSheet.name}）\n`);

    // 3. ヘッダー行を確認
    console.log('3. ヘッダー行を確認中...');
    const headerRow = 1;
    const companyIdHeader = inputSheet.getCell(`${excelColumns.inputSheet.companyId}${headerRow}`).value;
    const companyNameHeader = inputSheet.getCell(`${excelColumns.inputSheet.companyName}${headerRow}`).value;
    console.log(`   ${excelColumns.inputSheet.companyId}列: ${companyIdHeader}`);
    console.log(`   ${excelColumns.inputSheet.companyName}列: ${companyNameHeader}\n`);

    // 4. 2行目（最初のデータ行）のデータを取得
    console.log('4. 2行目（最初のデータ行）のデータを確認中...');
    const testRow = 2;

    const companyId = inputSheet.getCell(`${excelColumns.inputSheet.companyId}${testRow}`).value;
    const companyName = inputSheet.getCell(`${excelColumns.inputSheet.companyName}${testRow}`).value;
    const publishingCategory = inputSheet.getCell(`${excelColumns.inputSheet.publishingCategory}${testRow}`).value;
    const startDateValue = inputSheet.getCell(`${excelColumns.inputSheet.startDate}${testRow}`).value;
    const endDateValue = inputSheet.getCell(`${excelColumns.inputSheet.endDate}${testRow}`).value;

    console.log(`   企業ID (${excelColumns.inputSheet.companyId}列): ${companyId}`);
    console.log(`   企業名 (${excelColumns.inputSheet.companyName}列): ${companyName}`);
    console.log(`   掲載区分 (${excelColumns.inputSheet.publishingCategory}列): ${publishingCategory}`);
    console.log(`   開始日 (${excelColumns.inputSheet.startDate}列): ${startDateValue}`);
    console.log(`   終了日 (${excelColumns.inputSheet.endDate}列): ${endDateValue}\n`);

    if (!companyId) {
      console.log('   ⚠️  企業IDが取得できませんでした。データが存在しない可能性があります。\n');
      return;
    }

    // 5. 日付の変換テスト
    console.log('5. 日付の変換テスト...');
    try {
      const { excelDateToJSDate, formatDateForInput } = await import('./utils/dateUtils.js');
      
      if (startDateValue && endDateValue) {
        // ExcelJSはDateオブジェクト、数値、または文字列として読み込む可能性がある
        const startDate = excelDateToJSDate(startDateValue);
        const endDate = excelDateToJSDate(endDateValue);
        
        if (startDate && endDate) {
          const startDateStr = formatDateForInput(startDate);
          const endDateStr = formatDateForInput(endDate);
          
          if (startDateStr && endDateStr) {
            console.log(`   開始日（変換後）: ${startDateStr}`);
            console.log(`   終了日（変換後）: ${endDateStr}\n`);
          } else {
            console.log(`   ⚠️  日付のフォーマットに失敗しました（開始日: ${startDateValue}, 終了日: ${endDateValue}）\n`);
          }
        } else {
          console.log(`   ⚠️  日付の変換に失敗しました（開始日: ${startDateValue}, 終了日: ${endDateValue}）\n`);
        }
      } else {
        console.log('   ⚠️  日付データが取得できませんでした。\n');
      }
    } catch (error) {
      console.error(`   ❌ 日付変換エラー: ${error.message}\n`);
    }

    // 6. データ行数を確認
    console.log('6. データ行数を確認中...');
    let dataRowCount = 0;
    for (let row = 2; row <= inputSheet.rowCount; row++) {
      const cellValue = inputSheet.getCell(`${excelColumns.inputSheet.companyId}${row}`).value;
      if (cellValue) {
        dataRowCount++;
      } else {
        break;
      }
    }
    console.log(`   ✓ データ行数: ${dataRowCount}行\n`);

    console.log('=== テスト完了 ===\n');
    console.log('次のステップ:');
    console.log('1. 上記のデータが正しく読み込まれているか確認してください');
    console.log('2. 問題がなければ、実際のスクレイピングテストを実行できます');
    console.log('3. ログイン情報を.envファイルに設定してから、npm run test:single を実行してください\n');

  } catch (error) {
    console.error('❌ エラーが発生しました:', error.message);
    console.error(error.stack);
  }
}

// 実行
testExcelRead().catch(error => {
  console.error('致命的なエラー:', error);
  process.exit(1);
});
