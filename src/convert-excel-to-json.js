import ExcelService from './services/ExcelService.js';
import { writeFileSync, existsSync } from 'fs';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

/**
 * ExcelファイルからJSONファイルを生成する
 * 
 * 使用方法:
 *   npm run convert:excel-to-json
 * 
 * 変換元:
 *   - ナイト案件リスト.xlsx
 *   - 通常案件リスト.xlsx
 * 
 * 変換先:
 *   - config/jobCategoriesNight.json
 *   - config/jobCategoriesNormal.json
 */
async function convertExcelToJSON() {
  const excelService = new ExcelService();
  
  console.log('========================================');
  console.log('ExcelファイルからJSONファイルへの変換');
  console.log('========================================\n');
  
  try {
    // ナイト案件リストを変換
    const nightExcelPath = join(__dirname, '../ナイト案件リスト.xlsx');
    const nightJsonPath = join(__dirname, '../config/jobCategoriesNight.json');
    
    console.log('📋 ナイト案件リストを変換中...');
    console.log(`   変換元: ${nightExcelPath}`);
    
    if (!existsSync(nightExcelPath)) {
      console.warn(`   ⚠️  ファイルが見つかりません: ${nightExcelPath}`);
      console.warn(`   → JSONファイルは既存のものを保持します\n`);
    } else {
      const nightCategories = await excelService.loadJobCategoriesFromList(
        nightExcelPath
      );
      
      if (nightCategories.length === 0) {
        console.warn(`   ⚠️  データが0件です。JSONファイルは更新されません。\n`);
      } else {
        const nightJSON = JSON.stringify(
          nightCategories.map(cat => ({
            large: cat.large || '',
            medium: cat.medium || '',
            small: cat.small || ''
          })),
          null,
          2
        );
        
        writeFileSync(nightJsonPath, nightJSON, 'utf-8');
        console.log(`   ✓ 変換完了: ${nightCategories.length}件のデータをJSONに変換しました`);
        console.log(`   保存先: ${nightJsonPath}\n`);
      }
    }
    
    // 通常案件リストを変換
    const normalExcelPath = join(__dirname, '../通常案件リスト.xlsx');
    const normalJsonPath = join(__dirname, '../config/jobCategoriesNormal.json');
    
    console.log('📋 通常案件リストを変換中...');
    console.log(`   変換元: ${normalExcelPath}`);
    
    if (!existsSync(normalExcelPath)) {
      console.warn(`   ⚠️  ファイルが見つかりません: ${normalExcelPath}`);
      console.warn(`   → JSONファイルは既存のものを保持します\n`);
    } else {
      const normalCategories = await excelService.loadJobCategoriesFromList(
        normalExcelPath
      );
      
      if (normalCategories.length === 0) {
        console.warn(`   ⚠️  データが0件です。JSONファイルは更新されません。\n`);
      } else {
        const normalJSON = JSON.stringify(
          normalCategories.map(cat => ({
            large: cat.large || '',
            medium: cat.medium || '',
            small: cat.small || ''
          })),
          null,
          2
        );
        
        writeFileSync(normalJsonPath, normalJSON, 'utf-8');
        console.log(`   ✓ 変換完了: ${normalCategories.length}件のデータをJSONに変換しました`);
        console.log(`   保存先: ${normalJsonPath}\n`);
      }
    }
    
    console.log('========================================');
    console.log('✓ 変換処理が完了しました！');
    console.log('========================================');
    console.log('\n📝 注意事項:');
    console.log('   - JSONファイルは直接編集可能です');
    console.log('   - Excelファイルを更新した場合は、このコマンドを再実行してください');
    console.log('   - JSONファイルの形式:');
    console.log('     [');
    console.log('       {');
    console.log('         "large": "職種大",');
    console.log('         "medium": "職種中",');
    console.log('         "small": "職種小"');
    console.log('       }');
    console.log('     ]\n');
  } catch (error) {
    console.error('\n❌ 変換エラーが発生しました:');
    console.error(`   エラー内容: ${error.message}`);
    console.error(`   スタックトレース:\n${error.stack}\n`);
    process.exit(1);
  }
}

convertExcelToJSON();
