/**
 * 指定回数分のループテスト用スクリプト
 * 使用方法: npm run test:loop <回数>
 * 例: npm run test:loop 5
 * 
 * エクセルから<回数>分の企業IDを取得して、その個数分処理を繰り返します
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

/**
 * 通常案件の給与金額を処理する
 * - 「月収・日給・時給」というテキストは切り離す
 * - 「、」で区切られた場合は、最初のテキストのみを抽出
 * - 「〇〇円」と記載されている場合：数字のみを抽出
 * - 「○○万円」や「○○万円〜〇〇万円」という表記の場合：そのまま文字列として格納
 * @param {string|number} amount - 給与金額（文字列または数値）
 * @param {string} type - 給与形態（時給、日給、月給など）
 * @returns {string|number} 処理後の給与金額
 */
function processSalaryAmountForNormalCase(amount, type) {
  if (typeof amount === 'number') {
    // 数値の場合はそのまま返す
    return amount;
  }
  
  let salaryText = String(amount || '').trim();
  if (!salaryText) {
    return 0;
  }
  
  // 先頭にある「時給」「月給」「日給」までのテキストを取り除く
  salaryText = salaryText.replace(/^(時給|月給|日給|月収)[\s・、]*/i, '').trim();
  
  // 「、」で区切られた場合は、最初のテキストのみを抽出
  if (salaryText.includes('、')) {
    salaryText = salaryText.split('、')[0].trim();
  }
  
  // 「万円」が含まれている場合は、そのままテキストとして記入
  if (salaryText.includes('万円')) {
    return salaryText.trim();
  } else if (salaryText.includes('円')) {
    // 「円」を含めて取り除いて数字だけにする
    const amountMatch = salaryText.match(/([\d,]+)\s*円/);
    if (amountMatch) {
      return parseInt(amountMatch[1].replace(/,/g, ''), 10);
    } else {
      // 「円」が含まれているが数値が取得できない場合は、数値部分のみを抽出
      const numericValue = parseFloat(salaryText.replace(/[^\d.]/g, ''));
      return !isNaN(numericValue) ? numericValue : 0;
    }
  } else {
    // 「円」が含まれていない場合は、数値のみを抽出
    const numericValue = parseFloat(salaryText.replace(/[^\d.]/g, ''));
    return !isNaN(numericValue) ? numericValue : 0;
  }
}

/**
 * プラン名から「〇〇プラン」の後のテキストを除去
 * @param {string} plan - プラン名
 * @returns {string} クリーンなプラン名
 */
function cleanPlanName(plan) {
  if (!plan) return '';
  
  // 「〇〇プラン」の後のテキストを除去
  const match = plan.match(/^([A-Z]|O|PEX|EL|PL)\s*プラン/i);
  if (match) {
    return match[0].trim();
  }
  
  // プラン名のパターンに一致する部分を抽出
  const planPattern = /([A-Z]|O|PEX|EL|PL)\s*プラン/i;
  const planMatch = plan.match(planPattern);
  if (planMatch) {
    return planMatch[0].trim();
  }
  
  return plan.trim();
}

/**
 * 職種名から先頭の括弧と数字を除去
 * @param {string} jobCategory - 職種名
 * @returns {string} クリーンな職種名
 */
function cleanJobCategoryName(jobCategory) {
  if (!jobCategory) return '';
  
  // 先頭の括弧と数字を除去（例: "(1) アルバイト・パート" → "アルバイト・パート"）
  // 「[ア・パ]①」のような接頭辞も除去
  let cleaned = jobCategory
    .replace(/^[\(\[（【]\s*\d+\s*[\)\]）】]\s*/, '') // (1), [2], （3）, 【4】などを除去
    .replace(/^\d+[\.\)）]\s*/, '') // 1. や 1) などを除去
    .replace(/^\[[^\]]+\]\s*[①②③④⑤⑥⑦⑧⑨⑩]+\s*/, '') // [ア・パ]① などを除去
    .replace(/^\[[^\]]+\]\s*/, '') // [ア・パ] などを除去
    .replace(/^[①②③④⑤⑥⑦⑧⑨⑩]+\s*/, '') // ①②③ などを除去
    .trim();
  
  return cleaned;
}

/**
 * 1行のCSVデータを処理してスプレッドシートに書き込む（test-loop.js用）
 * @param {Object} csvRecord - CSVレコード
 * @param {string} companyId - 企業ID
 * @param {string} companyName - 企業名
 * @param {string} startDateStr - 掲載開始日（YYYY/MM/DD形式）
 * @param {string} endDateStr - 掲載終了日（YYYY/MM/DD形式）
 * @param {boolean} isNight - ナイト案件かどうか
 * @param {ScrapingService} scrapingService - スクレイピングサービス
 * @param {ExcelService} excelService - Excelサービス
 * @param {GoogleSheetsService} googleSheetsService - Google Sheetsサービス
 * @param {AIService} aiService - AIサービス
 * @param {Array} nightJobCategories - ナイト案件の職種カテゴリ
 * @param {Array} normalJobCategories - 通常案件の職種カテゴリ
 * @param {string} processFolderPath - 処理フォルダパス
 * @param {Array} csvRecords - CSVレコード全体（参照用）
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {Object} overrideValues - 上書きする値（PV数・応募数など）
 * @returns {Promise<boolean>} 成功したかどうか
 */
async function processSingleCSVRecordForTestLoop(
  csvRecord,
  companyId,
  companyName,
  startDateStr,
  endDateStr,
  isNight,
  scrapingService,
  excelService,
  googleSheetsService,
  aiService,
  nightJobCategories,
  normalJobCategories,
  processFolderPath,
  csvRecords,
  spreadsheetId,
  overrideValues = {},
  uniqueIdColumn = null,
  site = ''
) {
  try {
    // 空白行を見つける（スプレッドシート）
    const trendRow = await googleSheetsService.findFirstEmptyRow(
      spreadsheetId,
      config.googleSheets.sheetName,
      'A',
      2
    );
    console.log(`  書き込み開始行（スプレッドシート）: ${trendRow}`);
    
    // ナイト案件と通常案件で異なる列設定を使用
    const columnConfig = isNight 
      ? config.excelColumns.trendDatabaseNight 
      : config.excelColumns.trendDatabase;
    
    // スプレッドシートへの書き込み用ヘルパー関数
    const writeCell = async (column, value) => {
      await googleSheetsService.setCellValue(
        spreadsheetId,
        config.googleSheets.sheetName,
        column,
        trendRow,
        value
      );
    };
    
    const csvCols = config.excelColumns.downloadFile.csvColumns;
    
    // ⑦ プランを取得して転記
    let publishingPlan = excelService.getCSVValue(csvRecord, '掲載プラン') || 
                       excelService.getCSVValue(csvRecord, 'プラン') || '';
    
    // プランを判定（PEXプラン、Bプラン、ELプラン、Dプラン、Cプラン、Aプランから選択）
    let selectedPlan = null;
    if (publishingPlan) {
      // プラン名から「〇〇プラン」の後のテキストを除去
      const cleanedPlan = cleanPlanName(publishingPlan);
      
      selectedPlan = await aiService.determinePlan(cleanedPlan);
      if (selectedPlan) {
        await writeCell(columnConfig.plan, selectedPlan);
        console.log(`  プランを選択: ${selectedPlan}`);
      } else {
        // AI判定が失敗した場合はクリーンなプラン名を使用
        await writeCell(columnConfig.plan, cleanedPlan);
        console.log(`  プランを転記: ${cleanedPlan}`);
      }
    }

    // ⑧ 一覧PV数、詳細PV数、WEB応募数、TEL応募数を転記
    const listPV = overrideValues.listPV !== undefined 
      ? overrideValues.listPV 
      : (parseFloat(excelService.getCSVValue(csvRecord, csvCols.listPV) || 0));
    const detailPV = overrideValues.detailPV !== undefined 
      ? overrideValues.detailPV 
      : (parseFloat(excelService.getCSVValue(csvRecord, csvCols.detailPV) || 0));
    const webApplication = overrideValues.webApplication !== undefined 
      ? overrideValues.webApplication 
      : (parseFloat(excelService.getCSVValue(csvRecord, csvCols.webApplication) || 0));
    const telApplication = overrideValues.telApplication !== undefined 
      ? overrideValues.telApplication 
      : (parseFloat(excelService.getCSVValue(csvRecord, csvCols.normalApplication) || 0));
    
    await writeCell(columnConfig.listPV, listPV);
    await writeCell(columnConfig.detailPV, detailPV);
    await writeCell(columnConfig.webApplication, webApplication);
    await writeCell(columnConfig.telApplication, telApplication);

    // ⑨ 期間（週数）を計算して転記
    const appStartDate = excelService.getCSVValue(csvRecord, csvCols.applicationStartDate);
    const appEndDate = excelService.getCSVValue(csvRecord, csvCols.applicationEndDate);
    
    let startDateObj = null;
    let endDateObj = null;
    
    if (appStartDate && appEndDate) {
      startDateObj = excelDateToJSDate(appStartDate);
      endDateObj = excelDateToJSDate(appEndDate);
      
      if (startDateObj && endDateObj) {
        const weeks = calculateWeeks(startDateObj, endDateObj);
        await writeCell(columnConfig.period, weeks);
        console.log(`  期間（週数）を転記: ${weeks}週間`);
      }
    }

    // ⑩ 仕事Noを取得
    let jobNo = excelService.getCSVValue(csvRecord, '仕事No') || 
               excelService.getCSVValue(csvRecord, '仕事番号') || '';

    if (!jobNo) {
      console.log('  仕事Noが取得できませんでした。スキップします。');
      return false;
    }

    // ⑩⑪ 原稿検索とプレビューを開く（リトライ処理付き）
    console.log(`  仕事No: ${jobNo} で原稿を検索します`);
    let previewPage = null;
    const maxRetries = 5;
    let retryCount = 0;
    let searchSuccess = false;

    while (retryCount < maxRetries && !searchSuccess) {
      try {
        if (retryCount > 0) {
          console.log(`  リトライ ${retryCount}/${maxRetries - 1}...`);
          await scrapingService.goToTopAndReset();
        }

        // 原稿検索ページに移動
        await scrapingService.goToJobSearchPage();
        await new Promise(resolve => setTimeout(resolve, 1500));

        // 仕事Noで検索
        previewPage = await scrapingService.searchJobByNo(String(jobNo));
        await new Promise(resolve => setTimeout(resolve, 2000));

        // プレビューページが正しく読み込まれたか確認
        const previewUrl = previewPage.url();
        if (previewUrl.includes('/pv') || previewUrl.includes('preview')) {
          searchSuccess = true;
          console.log('  プレビューを開きました');
          
          // プレビューのスクリーンショットを保存
          if (processFolderPath) {
            try {
              await scrapingService.savePreviewScreenshot(
                previewPage, 
                processFolderPath, 
                String(jobNo),
                String(companyId),
                startDateStr,
                endDateStr
              );
            } catch (screenshotError) {
              console.warn(`  ⚠️  スクリーンショットの保存をスキップ: ${screenshotError.message}`);
            }
          }
          break;
        } else {
          throw new Error(`プレビューページに遷移できませんでした: ${previewUrl}`);
        }
      } catch (error) {
        retryCount++;
        if (retryCount >= maxRetries) {
          console.error(`  ❌ 原稿検索エラー（${maxRetries}回リトライ後）: ${error.message}`);
          return false;
        }
        console.warn(`  ⚠️  原稿検索エラー（リトライします）: ${error.message}`);
        await new Promise(resolve => setTimeout(resolve, 3000));
      }
    }

    if (!previewPage || !searchSuccess) {
      console.error('  ❌ プレビューページを開けませんでした');
      return false;
    }

    // プレビューページからデータを取得
    console.log('  プレビューページからデータを取得中...');
    
    // ⑫ 勤務地情報を取得
    let workLocation = null;
    for (let retry = 0; retry < 3; retry++) {
      try {
        workLocation = await scrapingService.getWorkLocation(previewPage);
        if (workLocation && (workLocation.prefecture || workLocation.city || workLocation.station)) {
          break;
        } else {
          throw new Error('勤務地情報が空です');
        }
      } catch (error) {
        if (retry < 2) {
          await new Promise(resolve => setTimeout(resolve, 2000));
          await previewPage.reload({ waitUntil: 'networkidle2' });
          await new Promise(resolve => setTimeout(resolve, 2000));
        } else {
          workLocation = { prefecture: '', city: '', station: '' };
        }
      }
    }
    
    const safeWorkLocation = workLocation || { prefecture: '', city: '', station: '' };
    
    // 都道府県はそのまま転記する（処理を削除）

    // ⑬ 年月を計算
    let year, month;
    if (startDateObj) {
      const extracted = excelService.extractYearMonth(startDateObj);
      year = extracted.year;
      month = extracted.month;
    } else if (appStartDate) {
      const convertedDate = excelDateToJSDate(appStartDate);
      if (convertedDate) {
        const extracted = excelService.extractYearMonth(convertedDate);
        year = extracted.year;
        month = extracted.month;
      }
    }
    
    if (year && month) {
      await writeCell(columnConfig.year, year);
      await writeCell(columnConfig.month, month);
      console.log(`  年月を転記: ${year}年${month}月`);
    }

    // ⑭ 地方を判定
    const regionOptions = ['北海道地方', '東北地方', '関東地方', '中部地方', '近畿地方', '中国地方', '四国地方', '九州', '沖縄地方'];
    const prefectureToRegion = {
      '北海道': '北海道地方',
      '青森県': '東北地方', '岩手県': '東北地方', '宮城県': '東北地方', '秋田県': '東北地方', '山形県': '東北地方', '福島県': '東北地方',
      '茨城県': '関東地方', '栃木県': '関東地方', '群馬県': '関東地方', '埼玉県': '関東地方', '千葉県': '関東地方', '東京都': '関東地方', '神奈川県': '関東地方',
      '新潟県': '中部地方', '富山県': '中部地方', '石川県': '中部地方', '福井県': '中部地方', '山梨県': '中部地方', '長野県': '中部地方', '岐阜県': '中部地方', '静岡県': '中部地方', '愛知県': '中部地方',
      '三重県': '近畿地方', '滋賀県': '近畿地方', '京都府': '近畿地方', '大阪府': '近畿地方', '兵庫県': '近畿地方', '奈良県': '近畿地方', '和歌山県': '近畿地方',
      '鳥取県': '中国地方', '島根県': '中国地方', '岡山県': '中国地方', '広島県': '中国地方', '山口県': '中国地方',
      '徳島県': '四国地方', '香川県': '四国地方', '愛媛県': '四国地方', '高知県': '四国地方',
      '福岡県': '九州', '佐賀県': '九州', '長崎県': '九州', '熊本県': '九州', '大分県': '九州', '宮崎県': '九州', '鹿児島県': '九州',
      '沖縄県': '沖縄地方'
    };
    
    let selectedRegion = null;
    if (safeWorkLocation.prefecture) {
      selectedRegion = prefectureToRegion[safeWorkLocation.prefecture];
      if (!selectedRegion) {
        try {
          selectedRegion = await aiService.determineRegion(safeWorkLocation.prefecture, regionOptions);
        } catch (error) {
          console.warn(`  ⚠️  地方のAI判定をスキップ: ${error.message}`);
        }
      }
      
      if (selectedRegion) {
        await writeCell(columnConfig.region, selectedRegion);
        console.log(`  地方を転記: ${selectedRegion}`);
      }
    }

    // ⑮ 都道府県、市区町村、最寄り駅を転記
    if (safeWorkLocation.prefecture) {
      await writeCell(columnConfig.prefecture, safeWorkLocation.prefecture);
      console.log(`  都道府県を転記: ${safeWorkLocation.prefecture}`);
    }
    if (safeWorkLocation.city) {
      await writeCell(columnConfig.city, safeWorkLocation.city);
      console.log(`  市区町村を転記: ${safeWorkLocation.city}`);
    }
    if (safeWorkLocation.station) {
      await writeCell(columnConfig.station, safeWorkLocation.station);
      console.log(`  最寄り駅を転記: ${safeWorkLocation.station}`);
    }

    // ⑯ 職種情報を取得
    let jobCategory = null;
    for (let retry = 0; retry < 3; retry++) {
      try {
        jobCategory = await scrapingService.getJobCategory(previewPage);
        if (jobCategory && (jobCategory.large || jobCategory.rawText)) {
          break;
        } else {
          throw new Error('職種情報が空です');
        }
      } catch (error) {
        if (retry < 2) {
          await new Promise(resolve => setTimeout(resolve, 2000));
        } else {
          jobCategory = { large: '', medium: '', small: '', rawText: '' };
        }
      }
    }
    
    // 職種大・中・小を選択
    let selectedLarge = null;
    let selectedMedium = null;
    let selectedSmall = null;
    
    const jobCategories = isNight ? nightJobCategories : normalJobCategories;
    
    if (jobCategory?.rawText && jobCategories.length > 0) {
      if (isNight) {
        // ナイト案件の場合の処理
        let cleanedRawText = jobCategory.rawText;
        cleanedRawText = cleanedRawText.replace(/^\[[^\]]+\][①②③④⑤⑥⑦⑧⑨⑩]*/, '').trim();
        
        const jobParts = cleanedRawText.split('、').map(part => part.trim()).filter(part => part);
        console.log(`  職種を分割: ${jobParts.join(', ')}`);
        
        let maxSimilarity = 0;
        let bestMatch = null;
        
        for (const jobPart of jobParts) {
          for (const category of jobCategories) {
            const combinedText = `${category.large} ${category.medium} ${category.small}`.trim();
            const similarity = aiService.calculateCosineSimilarity(jobPart, combinedText);
            
            if (similarity > maxSimilarity) {
              maxSimilarity = similarity;
              bestMatch = category;
            }
          }
        }
        
        if (maxSimilarity >= 0.3 && bestMatch) {
          selectedLarge = bestMatch.large;
          selectedMedium = bestMatch.medium;
          selectedSmall = bestMatch.small;
          console.log(`  コサイン類似度で職種を選択（類似度: ${maxSimilarity.toFixed(3)}）: 大=${selectedLarge}, 中=${selectedMedium}, 小=${selectedSmall}`);
        }
      } else {
        // 通常案件の場合の処理（既存のロジック）
        const cleanedRawText = cleanJobCategoryName(jobCategory.rawText);
        const cleanedLarge = cleanJobCategoryName(jobCategory.large);
        
        let keywordMatched = false;
        const jobTextLower = cleanedRawText.toLowerCase();
        
        for (const category of jobCategories) {
          const largeMatch = category.large && jobTextLower.includes(category.large.toLowerCase());
          const mediumMatch = category.medium && jobTextLower.includes(category.medium.toLowerCase());
          const smallMatch = category.small && jobTextLower.includes(category.small.toLowerCase());
          
          const categoryLargeLower = category.large ? category.large.toLowerCase() : '';
          const categoryMediumLower = category.medium ? category.medium.toLowerCase() : '';
          const categorySmallLower = category.small ? category.small.toLowerCase() : '';
          
          const keywords = cleanedRawText.split(/[・、,\s]+/).filter(k => k.length >= 2);
          
          const reverseLargeMatch = keywords.some(keyword => 
            categoryLargeLower.includes(keyword.toLowerCase())
          );
          const reverseMediumMatch = keywords.some(keyword => 
            categoryMediumLower.includes(keyword.toLowerCase())
          );
          const reverseSmallMatch = keywords.some(keyword => 
            categorySmallLower.includes(keyword.toLowerCase())
          );
          
          if (largeMatch || mediumMatch || smallMatch || reverseLargeMatch || reverseMediumMatch || reverseSmallMatch) {
            selectedLarge = category.large;
            selectedMedium = category.medium;
            selectedSmall = category.small;
            keywordMatched = true;
            console.log(`  キーワードマッチングで職種を選択: 大=${selectedLarge}, 中=${selectedMedium}, 小=${selectedSmall}`);
            break;
          }
        }
        
        if (!keywordMatched) {
          const matchedCategory = await aiService.determineJobCategoryByCosineSimilarity(
            cleanedRawText,
            jobCategories
          );
          
          if (matchedCategory) {
            selectedLarge = matchedCategory.large;
            selectedMedium = matchedCategory.medium;
            selectedSmall = matchedCategory.small;
            console.log(`  コサイン類似度で職種を選択: 大=${selectedLarge}, 中=${selectedMedium}, 小=${selectedSmall}`);
          } else {
            selectedLarge = cleanedLarge || jobCategory.large;
            selectedMedium = jobCategory.medium;
            selectedSmall = jobCategory.small;
          }
        }
      }
    } else {
      selectedLarge = cleanJobCategoryName(jobCategory?.large || '');
      selectedMedium = jobCategory?.medium;
      selectedSmall = jobCategory?.small;
    }
    
    // 職種を転記
    if (selectedLarge) {
      await writeCell(columnConfig.jobCategoryLarge, selectedLarge);
      console.log(`  職種大を転記: ${selectedLarge}`);
    }
    if (selectedMedium) {
      await writeCell(columnConfig.jobCategoryMedium, selectedMedium);
      console.log(`  職種中を転記: ${selectedMedium}`);
    }
    if (selectedSmall) {
      await writeCell(columnConfig.jobCategorySmall, selectedSmall);
      console.log(`  職種小を転記: ${selectedSmall}`);
    }

    // ⑰ 給与情報を取得
    let salary = null;
    for (let retry = 0; retry < 3; retry++) {
      try {
        salary = await scrapingService.getSalary(previewPage);
        if (salary && (salary.type || salary.amount > 0)) {
          break;
        } else {
          throw new Error('給与情報が空です');
        }
      } catch (error) {
        if (retry < 2) {
          await new Promise(resolve => setTimeout(resolve, 2000));
        } else {
          salary = { type: '', amount: 0 };
        }
      }
    }
    
    // 給与形態と金額を処理
    // 給与形態はそのまま転記する
    let salaryType = salary?.type || '';
    let salaryAmount = salary?.amount || 0;
    
    if (!isNight) {
      // 通常案件の場合、金額のみ処理（給与形態はそのまま）
      salaryAmount = processSalaryAmountForNormalCase(salaryAmount, salaryType);
    } else {
      // ナイト案件の場合、給与金額を処理
      if (typeof salaryAmount === 'string') {
        let salaryStr = String(salaryAmount).trim();
        
        // 「、」で区切られている場合は、1個目だけを抽出
        if (salaryStr.includes('、')) {
          salaryStr = salaryStr.split('、')[0].trim();
          console.log(`  給与（「、」で区切られているため、1個目を抽出）: ${salaryStr}`);
        }
        
        // 「時給」「月給」「日給」を抽出してL列に格納
        const salaryTypeMatch = salaryStr.match(/(時給|月給|日給)/);
        if (salaryTypeMatch) {
          salaryType = salaryTypeMatch[1];
          console.log(`  給与形態を抽出: ${salaryType}`);
        }
        
        // 「時給」「月給」「日給」という文字を除去
        salaryStr = salaryStr.replace(/時給|月給|日給/g, '').trim();
        
        // 「円」という文字を除去
        if (salaryStr.includes('円')) {
          salaryStr = salaryStr.replace(/円/g, '').trim();
        }
        
        // 「〜」「～」を除去
        salaryStr = salaryStr.replace(/[〜～]/g, '').trim();
        
        // 数値のみの場合は数値に変換
        const numericValue = parseFloat(salaryStr.replace(/[^\d.-]/g, ''));
        if (!isNaN(numericValue) && salaryStr.match(/^[\d,.\s-]+$/)) {
          // 数値のみの場合は数値として格納
          salaryAmount = numericValue;
          console.log(`  給与（数値）を格納: ${salaryAmount}`);
        } else {
          // 範囲表記（「から」など）が含まれる場合は文字列として格納
          salaryAmount = salaryStr;
          console.log(`  給与（文字列形式）を格納: ${salaryAmount}`);
        }
      }
    }
    
    if (salaryType) {
      await writeCell(columnConfig.salaryType, salaryType);
      console.log(`  給与形態を転記: ${salaryType}`);
    }
    if (salaryAmount) {
      await writeCell(columnConfig.salaryAmount, salaryAmount);
      console.log(`  給与金額を転記: ${salaryAmount}`);
    }

    // ⑱ 企業ID・企業名を転記
    if (companyId) {
      // 企業IDをCompNo列に記載
      await writeCell(columnConfig.compNo, companyId);
      console.log(`  企業IDをCompNo列に転記: ${companyId}`);
    }
    if (companyName) {
      await writeCell(columnConfig.companyName, companyName);
      console.log(`  企業名を転記: ${companyName}`);
    }

    // ⑲ 店名を転記（一旦未入力でOK）
    const storeName = '';
    if (storeName) {
      await writeCell(columnConfig.storeName, storeName);
      console.log(`  店名を転記: ${storeName}`);
    }

    // ⑳ 申込開始日・終了日を転記
    if (appStartDate && appEndDate) {
      if (startDateObj && endDateObj) {
        const formattedStartDate = formatDateForInput(startDateObj);
        const formattedEndDate = formatDateForInput(endDateObj);
        
        if (formattedStartDate && formattedEndDate) {
          await writeCell(columnConfig.applicationStartDate, formattedStartDate);
          await writeCell(columnConfig.applicationEndDate, formattedEndDate);
          console.log(`  申込開始日を転記: ${formattedStartDate}`);
          console.log(`  申込終了日を転記: ${formattedEndDate}`);
        }
      }
    }

    // 媒体を転記（通常案件のみ）
    if (!isNight) {
      const siteValue = site || '';
      await writeCell(columnConfig.media, siteValue);
      if (site) {
        console.log(`  媒体を転記: ${site}`);
      } else {
        console.warn(`  ⚠️  サイト情報がないため、媒体列は空のままです。`);
      }
    }

    // ユニークIDを転記（ナイト案件と通常案件の両方）
    if (uniqueIdColumn) {
      const uniqueIdValue = `${companyId}_${companyName}_${startDateStr}_${endDateStr}`;
      await writeCell(uniqueIdColumn, uniqueIdValue);
      console.log(`  ユニークIDを転記: ${uniqueIdValue} (列: ${uniqueIdColumn})`);
    }

    // プレビューページを閉じる
    await scrapingService.closePreviewTab(previewPage);
    console.log('  プレビューページを閉じました');
    
    return true;
  } catch (error) {
    console.error(`  ❌ 処理エラー: ${error.message}`);
    return false;
  }
}

/**
 * 1企業分のテスト処理
 * @param {Object} companyData - 企業データ（企業ID、企業名、掲載区分、開始日、終了日）
 * @param {number} loopIndex - ループのインデックス（1から開始）
 * @param {number} totalLoops - 総ループ回数
 * @param {ScrapingService} scrapingService - スクレイピングサービス（再利用）
 * @param {ExcelService} excelService - Excelサービス
 * @param {GoogleSheetsService} googleSheetsService - Google Sheetsサービス
 * @param {AIService} aiService - AIサービス
 * @param {Array} nightJobCategories - ナイト案件の職種カテゴリリスト
 * @param {Array} normalJobCategories - 通常案件の職種カテゴリリスト
 * @param {Object|null} inputSheet - 入力エクセルのシート（オプション、エクセル読み込みを行わない場合はnull）
 * @param {Function} onFailure - 失敗時に呼び出されるコールバック関数（企業ID、企業名、仕事No、失敗理由）
 * @returns {Promise<boolean>} 成功したかどうか
 */
export async function testSingleCompany(companyData, loopIndex, totalLoops, scrapingService, excelService, googleSheetsService, aiService, nightJobCategories, normalJobCategories, inputSheet, onFailure) {
  const { companyId, companyName, publishingCategory, startDate, endDate } = companyData;
  let jobNo = ''; // 失敗ログ用に仕事Noを保持

  try {
    console.log(`\n${'='.repeat(60)}`);
    console.log(`=== ループ ${loopIndex}/${totalLoops} の動作確認を開始します ===`);
    console.log(`企業ID: ${companyId}, 企業名: ${companyName}`);
    console.log(`${'='.repeat(60)}\n`);

    if (!companyId) {
      console.log('❌ 企業IDが取得できませんでした。');
      if (onFailure && typeof onFailure === 'function') {
        onFailure(companyId || '', companyName || '', '', '企業IDが取得できませんでした');
      }
      return false;
    }

    const startDateStr = formatDateForInput(startDate);
    const endDateStr = formatDateForInput(endDate);

    if (!startDateStr || !endDateStr) {
      console.log(`❌ 日付のフォーマットに失敗しました。`);
      if (onFailure && typeof onFailure === 'function') {
        onFailure(companyId, companyName || '', '', '日付のフォーマットに失敗しました');
      }
      return false;
    }

    console.log(`   開始日（変換後）: ${startDateStr}`);
    console.log(`   終了日（変換後）: ${endDateStr}\n`);

    // ナイト案件かどうかを判定
    const isNight = publishingCategory === 'ナイト';
    
    // ユニークIDを算出: 企業ID、企業名、開始日、終了日を結合したユニークな値を生成
    const uniqueId = `${companyId}_${companyName}_${startDateStr}_${endDateStr}`;
    
    // 対象の行がナイト案件か通常案件かで、参照するスプレッドシートを分岐する
    const checkSpreadsheetId = isNight ? config.googleSheets.spreadsheetIdNight : config.googleSheets.spreadsheetIdNormal;
    
    // ユニークID列を検索（既に手動で追加済みのため、追加処理は行わない）
    let uniqueIdColumn = null;
    try {
      uniqueIdColumn = await googleSheetsService.findColumnByName(
        checkSpreadsheetId,
        config.googleSheets.sheetName,
        'ユニークID'
      );
      if (uniqueIdColumn) {
        console.log(`  ✓ ユニークID列を検索しました: ${uniqueIdColumn}`);
      } else {
        console.warn(`  ⚠️  ユニークID列が見つかりませんでした`);
      }
    } catch (error) {
      console.warn(`  ⚠️  ユニークID列の検索エラー: ${error.message}`);
    }
    
    // 算出したユニークIDが対象スプレッドシートに存在するか確認する
    if (uniqueIdColumn) {
      const isDuplicate = await googleSheetsService.findValueInColumn(
        checkSpreadsheetId,
        config.googleSheets.sheetName,
        uniqueIdColumn,
        uniqueId,
        2
      );
      
      if (isDuplicate) {
        // もし存在する時は次の処理内容に移る
        console.log(`  ⚠️  重複データを検出しました。スキップします。`);
        console.log(`  ユニークID: ${uniqueId}`);
        if (onFailure && typeof onFailure === 'function') {
          onFailure(companyId, companyName || '', '', '重複データのためスキップ');
        }
        return false;
      }
    }

    // スプレッドシートIDの確認
    console.log('3. スプレッドシートを確認中...');
    const spreadsheetId = isNight ? config.googleSheets.spreadsheetIdNight : config.googleSheets.spreadsheetIdNormal;
    
    if (!spreadsheetId) {
      console.error(`❌ ${isNight ? 'ナイト' : '通常'}案件のスプレッドシートIDが設定されていません`);
      if (onFailure && typeof onFailure === 'function') {
        onFailure(companyId, companyName || '', '', `スプレッドシートIDが設定されていません`);
      }
      return false;
    }
    
    console.log(`✓ ${isNight ? 'ナイト' : '通常'}案件のスプレッドシートを使用します\n`);

    // TOP画面に遷移して入力フィールドをリセット（最初のループ以外）
    if (loopIndex > 1) {
      console.log('TOP画面に遷移して入力フィールドをリセット中...');
      try {
        await scrapingService.goToTopAndReset();
        console.log('✓ リセット完了\n');
      } catch (topError) {
        console.warn(`  ⚠️  TOPページへの移動でエラー: ${topError.message}`);
        // エラーが発生しても、直接URLで遷移を試みる
        try {
          const page = scrapingService.getPage();
          const topUrl = config.baitoru.loginUrl.includes('/top') 
            ? config.baitoru.loginUrl 
            : config.baitoru.loginUrl.replace(/\/$/, '') + '/top';
          await page.goto(topUrl, {
            waitUntil: 'networkidle2',
            timeout: 60000
          });
          await new Promise(resolve => setTimeout(resolve, 1000));
          console.log('  ✓ TOPページに直接遷移しました\n');
        } catch (directError) {
          console.warn(`  ⚠️  直接遷移も失敗: ${directError.message}\n`);
          throw directError;
        }
      }
    }

    // 企業IDで検索
    console.log(`4. 企業ID: ${companyId} で検索中...`);
    try {
      await scrapingService.searchByCompanyId(String(companyId));
      console.log('✓ 検索完了\n');
    } catch (error) {
      console.error(`❌ 検索エラー: ${error.message}`);
      console.log('   セレクタ設定（config/selectors.json）を確認してください。');
      
      if (onFailure && typeof onFailure === 'function') {
        onFailure(companyId, companyName || '', '', `企業ID検索エラー: ${error.message}`);
      }
      
      // TOPページに戻って入力フィールドをリセット
      try {
        await scrapingService.goToTopAndReset();
        console.log('  ✓ TOPページに戻りました\n');
      } catch (topError) {
        console.warn(`  ⚠️  TOPページへの移動でエラー: ${topError.message}`);
        // エラーが発生しても、直接URLで遷移を試みる
        try {
          const page = scrapingService.getPage();
          const topUrl = config.baitoru.loginUrl.includes('/top') 
            ? config.baitoru.loginUrl 
            : config.baitoru.loginUrl.replace(/\/$/, '') + '/top';
          await page.goto(topUrl, {
            waitUntil: 'networkidle2',
            timeout: 60000
          });
          await new Promise(resolve => setTimeout(resolve, 1000));
          console.log('  ✓ TOPページに直接遷移しました\n');
        } catch (directError) {
          console.warn(`  ⚠️  直接遷移も失敗: ${directError.message}\n`);
        }
      }
      
      return false;
    }

    // 選択ボタンをクリック
    console.log('5. 選択ボタンをクリック中...');
    try {
      await scrapingService.clickSelectButton();
      console.log('✓ 選択ボタンをクリックしました\n');
    } catch (error) {
      console.error(`❌ 選択ボタンクリックエラー: ${error.message}`);
      
      if (onFailure && typeof onFailure === 'function') {
        onFailure(companyId, companyName || '', '', `選択ボタンクリックエラー: ${error.message}`);
      }
      
      // TOPページに戻って入力フィールドをリセット
      try {
        await scrapingService.goToTopAndReset();
        console.log('  ✓ TOPページに戻りました\n');
      } catch (topError) {
        console.warn(`  ⚠️  TOPページへの移動でエラー: ${topError.message}`);
        // エラーが発生しても、直接URLで遷移を試みる
        try {
          const page = scrapingService.getPage();
          const topUrl = config.baitoru.loginUrl.includes('/top') 
            ? config.baitoru.loginUrl 
            : config.baitoru.loginUrl.replace(/\/$/, '') + '/top';
          await page.goto(topUrl, {
            waitUntil: 'networkidle2',
            timeout: 60000
          });
          await new Promise(resolve => setTimeout(resolve, 1000));
          console.log('  ✓ TOPページに直接遷移しました\n');
        } catch (directError) {
          console.warn(`  ⚠️  直接遷移も失敗: ${directError.message}\n`);
        }
      }
      
      return false;
    }

    // 掲載実績をダウンロード
    console.log(`6. 掲載実績をダウンロード中（${startDateStr} ～ ${endDateStr}）...`);
    let downloadFilePath = null;
    let processFolderPath = null;
    try {
      const downloadResult = await scrapingService.downloadPerformance(
        startDateStr,
        endDateStr,
        String(companyId)
      );
      downloadFilePath = downloadResult.filePath;
      processFolderPath = downloadResult.folderPath;
      console.log(`✓ ダウンロード完了: ${downloadFilePath}`);
      console.log(`✓ 処理フォルダ: ${processFolderPath}\n`);
    } catch (error) {
      console.error(`❌ ダウンロードエラー: ${error.message}`);
      
      if (onFailure && typeof onFailure === 'function') {
        onFailure(companyId, companyName || '', '', `掲載実績ダウンロードエラー: ${error.message}`);
      }
      
      // TOPページに戻って入力フィールドをリセット
      try {
        await scrapingService.goToTopAndReset();
        console.log('  ✓ TOPページに戻りました\n');
      } catch (topError) {
        console.warn(`  ⚠️  TOPページへの移動でエラー: ${topError.message}`);
        // エラーが発生しても、直接URLで遷移を試みる
        try {
          const page = scrapingService.getPage();
          const topUrl = config.baitoru.loginUrl.includes('/top') 
            ? config.baitoru.loginUrl 
            : config.baitoru.loginUrl.replace(/\/$/, '') + '/top';
          await page.goto(topUrl, {
            waitUntil: 'networkidle2',
            timeout: 60000
          });
          await new Promise(resolve => setTimeout(resolve, 1000));
          console.log('  ✓ TOPページに直接遷移しました\n');
        } catch (directError) {
          console.warn(`  ⚠️  直接遷移も失敗: ${directError.message}\n`);
        }
      }
      
      return false;
    }

    // ダウンロードファイルを読み込み（CSVまたはExcel）
    console.log('7. ダウンロードファイルを読み込み中...');
    let csvRecords = null;
    let downloadSheet = null;
    const isCSV = downloadFilePath.endsWith('.csv');

    try {
      if (isCSV) {
        csvRecords = await excelService.loadCSVFile(downloadFilePath);
        console.log(`✓ CSVファイルを読み込みました（${csvRecords.length}件のレコード）\n`);
      } else {
        const downloadWorkbook = await excelService.loadDownloadFile(downloadFilePath);
        downloadSheet = downloadWorkbook.getWorksheet(1);
        console.log('✓ Excelファイルを読み込みました\n');
      }
    } catch (fileError) {
      console.error(`❌ ファイル読み込みエラー: ${fileError.message}`);
      
      if (onFailure && typeof onFailure === 'function') {
        onFailure(companyId, companyName || '', '', `ダウンロードファイル読み込みエラー: ${fileError.message}`);
      }
      
      // TOPページに戻って入力フィールドをリセット
      try {
        await scrapingService.goToTopAndReset();
        console.log('  ✓ TOPページに戻りました\n');
      } catch (topError) {
        console.warn(`  ⚠️  TOPページへの移動でエラー: ${topError.message}`);
        // エラーが発生しても、直接URLで遷移を試みる
        try {
          const page = scrapingService.getPage();
          const topUrl = config.baitoru.loginUrl.includes('/top') 
            ? config.baitoru.loginUrl 
            : config.baitoru.loginUrl.replace(/\/$/, '') + '/top';
          await page.goto(topUrl, {
            waitUntil: 'networkidle2',
            timeout: 60000
          });
          await new Promise(resolve => setTimeout(resolve, 1000));
          console.log('  ✓ TOPページに直接遷移しました\n');
        } catch (directError) {
          console.warn(`  ⚠️  直接遷移も失敗: ${directError.message}\n`);
        }
      }
      
      return false;
    }

    // バリデーション：掲載実績期間と申込期間が同日かチェック
    console.log('8. データのバリデーション中...');
    let isValid = false;
    let aggregatedData = null;

    try {
      if (isCSV) {
        if (csvRecords.length === 0) {
          console.log('  CSVファイルにデータがありません。スキップします。');
          
          if (onFailure && typeof onFailure === 'function') {
            onFailure(companyId, companyName || '', '', 'CSVファイルにデータがありません（レコードが0件）');
          }
          
          // TOPページに戻って入力フィールドをリセット
          try {
            await scrapingService.goToTopAndReset();
            console.log('  ✓ TOPページに戻りました\n');
          } catch (topError) {
            console.warn(`  ⚠️  TOPページへの移動でエラー: ${topError.message}`);
            // エラーが発生しても、直接URLで遷移を試みる
            try {
              const page = scrapingService.getPage();
              const topUrl = config.baitoru.loginUrl.includes('/top') 
                ? config.baitoru.loginUrl 
                : config.baitoru.loginUrl.replace(/\/$/, '') + '/top';
              await page.goto(topUrl, {
                waitUntil: 'networkidle2',
                timeout: 60000
              });
              await new Promise(resolve => setTimeout(resolve, 1000));
              console.log('  ✓ TOPページに直接遷移しました\n');
            } catch (directError) {
              console.warn(`  ⚠️  直接遷移も失敗: ${directError.message}\n`);
            }
          }
          
          return false;
        }

        // CSV列名を設定から取得
        const csvCols = config.excelColumns.downloadFile.csvColumns;
        const performanceStartCol = csvCols.performanceStartDate;
        const performanceEndCol = csvCols.performanceEndDate;
        const applicationStartCol = csvCols.applicationStartDate;
        const applicationEndCol = csvCols.applicationEndDate;
        const performanceDataCols = csvCols.performanceData;

        // 各行に対して日付一致判定フラグを付ける
        const recordsWithFlags = csvRecords.map((record, index) => {
          const perfStart = excelService.getCSVValue(record, performanceStartCol);
          const perfEnd = excelService.getCSVValue(record, performanceEndCol);
          const appStart = excelService.getCSVValue(record, applicationStartCol);
          const appEnd = excelService.getCSVValue(record, applicationEndCol);

          // 開始日が同日かチェック
          const startDateMatch = excelService.isSameDate(perfStart, appStart);
          
          // 終了日が「掲載中」の場合は特別扱い（開始日が同日ならOK）
          const perfEndStr = String(perfEnd || '').trim();
          const isPublishing = perfEndStr === '掲載中';
          
          let isValid = false;
          if (isPublishing) {
            // 「掲載中」の場合は、開始日が同日であればOK
            isValid = startDateMatch;
          } else {
            // 通常の場合は、開始日と終了日が両方とも同日かチェック
            const endDateMatch = excelService.isSameDate(perfEnd, appEnd);
            isValid = startDateMatch && endDateMatch;
          }
          
          // デバッグログ：最初の3件と不一致のレコードのみ表示
          if (index < 3 || !isValid) {
            console.log(`  [DEBUG] レコード${index + 1}: 掲載実績開始日=${perfStart}, 申込開始日=${appStart}, 開始日一致=${startDateMatch}`);
            console.log(`  [DEBUG] レコード${index + 1}: 掲載実績終了日=${perfEnd}, 申込終了日=${appEnd}, 終了日一致=${!isPublishing ? excelService.isSameDate(perfEnd, appEnd) : 'N/A(掲載中)'}, 有効=${isValid}`);
          }
          
          return { record, isValid, index: index + 1 };
        });

        // Trueの行とFalseの行を分ける
        const validRecords = recordsWithFlags.filter(r => r.isValid);
        const invalidRecords = recordsWithFlags.filter(r => !r.isValid);

        console.log(`  CSVレコード数: ${csvRecords.length}件`);
        console.log(`  日付一致（True）: ${validRecords.length}件`);
        console.log(`  日付不一致（False）: ${invalidRecords.length}件\n`);

        // Trueの行がある場合、またはFalseの行のみの場合
        if (validRecords.length > 0 || invalidRecords.length > 0) {
          isValid = true;
          aggregatedData = { validRecords, invalidRecords, isMultiRow: true };
        } else {
          // レコードがない場合（通常は発生しない）
          console.log('  レコードがありません。');
          isValid = false;
          
          if (onFailure && typeof onFailure === 'function') {
            onFailure(companyId, companyName || '', '', 'CSVファイルにデータがありません');
          }
        }
      } else {
        console.log('  Excelファイルのバリデーションは未実装です');
        
        // TOPページに戻って入力フィールドをリセット
        try {
          await scrapingService.goToTopAndReset();
          console.log('  ✓ TOPページに戻りました\n');
        } catch (topError) {
          console.warn(`  ⚠️  TOPページへの移動でエラー: ${topError.message}`);
          // エラーが発生しても、直接URLで遷移を試みる
          try {
            const page = scrapingService.getPage();
            const topUrl = config.baitoru.loginUrl.includes('/top') 
              ? config.baitoru.loginUrl 
              : config.baitoru.loginUrl.replace(/\/$/, '') + '/top';
            await page.goto(topUrl, {
              waitUntil: 'networkidle2',
              timeout: 60000
            });
            await new Promise(resolve => setTimeout(resolve, 1000));
            console.log('  ✓ TOPページに直接遷移しました\n');
          } catch (directError) {
            console.warn(`  ⚠️  直接遷移も失敗: ${directError.message}\n`);
          }
        }
        
        return false;
      }
    } catch (validationError) {
      console.error(`❌ バリデーションエラー: ${validationError.message}`);
      
      if (onFailure && typeof onFailure === 'function') {
        onFailure(companyId, companyName || '', '', `バリデーションエラー: ${validationError.message}`);
      }
      
      // TOPページに戻って入力フィールドをリセット
      try {
        await scrapingService.goToTopAndReset();
        console.log('  ✓ TOPページに戻りました\n');
      } catch (topError) {
        console.warn(`  ⚠️  TOPページへの移動でエラー: ${topError.message}`);
        // エラーが発生しても、直接URLで遷移を試みる
        try {
          const page = scrapingService.getPage();
          const topUrl = config.baitoru.loginUrl.includes('/top') 
            ? config.baitoru.loginUrl 
            : config.baitoru.loginUrl.replace(/\/$/, '') + '/top';
          await page.goto(topUrl, {
            waitUntil: 'networkidle2',
            timeout: 60000
          });
          await new Promise(resolve => setTimeout(resolve, 1000));
          console.log('  ✓ TOPページに直接遷移しました\n');
        } catch (directError) {
          console.warn(`  ⚠️  直接遷移も失敗: ${directError.message}\n`);
        }
      }
      
      return false;
    }

    if (!isValid || !aggregatedData) {
      console.log('❌ データのバリデーションに失敗しました（掲載実績期間と申込期間が一致しません）。スキップします。');
      
      if (onFailure && typeof onFailure === 'function') {
        onFailure(companyId, companyName || '', '', 'データのバリデーションに失敗しました（掲載実績期間と申込期間が一致しません）');
      }
      
      // TOPページに戻って入力フィールドをリセット
      try {
        await scrapingService.goToTopAndReset();
        console.log('  ✓ TOPページに戻りました\n');
      } catch (topError) {
        console.warn(`  ⚠️  TOPページへの移動でエラー: ${topError.message}`);
        // エラーが発生しても、直接URLで遷移を試みる
        try {
          const page = scrapingService.getPage();
          const topUrl = config.baitoru.loginUrl.includes('/top') 
            ? config.baitoru.loginUrl 
            : config.baitoru.loginUrl.replace(/\/$/, '') + '/top';
          await page.goto(topUrl, {
            waitUntil: 'networkidle2',
            timeout: 60000
          });
          await new Promise(resolve => setTimeout(resolve, 1000));
          console.log('  ✓ TOPページに直接遷移しました\n');
        } catch (directError) {
          console.warn(`  ⚠️  直接遷移も失敗: ${directError.message}\n`);
        }
      }
      
      return false;
    }

    // CSVファイルが複数行の場合の処理
    if (isCSV && aggregatedData && aggregatedData.isMultiRow) {
      const { validRecords, invalidRecords } = aggregatedData;
      
      // Trueの行を1行ずつ処理
      for (let i = 0; i < validRecords.length; i++) {
        const validRecord = validRecords[i];
        console.log(`\n  [True行 ${i + 1}/${validRecords.length}] 処理を開始します`);
        
        // 仕事Noを取得
        jobNo = excelService.getCSVValue(validRecord.record, '仕事No') || 
               excelService.getCSVValue(validRecord.record, '仕事番号') || '';
        
        if (!jobNo) {
          console.log(`  ❌ 仕事Noが取得できませんでした。スキップします。`);
          continue;
        }
        
        console.log(`  ✓ 仕事No: ${jobNo}`);
        
        // サイト（媒体）を取得
        let site = '';
        if (inputSheet && companyData.row) {
          site = excelService.getCellValue(
            inputSheet,
            config.excelColumns.inputSheet.site,
            companyData.row
          ) || '';
        }
        
        // この行のデータを処理（既存の処理ロジックを再利用）
        const success = await processSingleCSVRecordForTestLoop(
          validRecord.record,
          companyId,
          companyName,
          startDateStr,
          endDateStr,
          isNight,
          scrapingService,
          excelService,
          googleSheetsService,
          aiService,
          nightJobCategories,
          normalJobCategories,
          processFolderPath,
          csvRecords,
          spreadsheetId,
          {},
          uniqueIdColumn,
          site
        );
        
        if (!success) {
          console.log(`  ❌ True行 ${i + 1} の処理に失敗しました`);
        }
      }
      
      // Falseの行が複数ある場合、合計値を算出して書き込む
      if (invalidRecords.length > 0) {
        console.log(`\n  [False行] ${invalidRecords.length}件の合計値を算出して書き込みます`);
        
        // Falseの行の合計値を算出
        const csvCols = config.excelColumns.downloadFile.csvColumns;
        let totalListPV = 0;
        let totalDetailPV = 0;
        let totalWebApplication = 0;
        let totalTelApplication = 0;
        
        for (const invalidRecord of invalidRecords) {
          const listPV = parseFloat(excelService.getCSVValue(invalidRecord.record, csvCols.listPV) || 0);
          const detailPV = parseFloat(excelService.getCSVValue(invalidRecord.record, csvCols.detailPV) || 0);
          const webApplication = parseFloat(excelService.getCSVValue(invalidRecord.record, csvCols.webApplication) || 0);
          const telApplication = parseFloat(excelService.getCSVValue(invalidRecord.record, csvCols.normalApplication) || 0);
          
          totalListPV += listPV;
          totalDetailPV += detailPV;
          totalWebApplication += webApplication;
          totalTelApplication += telApplication;
        }
        
        console.log(`  合計値: 一覧PV数=${totalListPV}, 詳細PV数=${totalDetailPV}, WEB応募数=${totalWebApplication}, TEL応募数=${totalTelApplication}`);
        
        // 最初のFalse行のデータを使用（その他の値）
        const firstInvalidRecord = invalidRecords[0].record;
        
        // 仕事Noを取得
        jobNo = excelService.getCSVValue(firstInvalidRecord, '仕事No') || 
               excelService.getCSVValue(firstInvalidRecord, '仕事番号') || '';
        
        if (!jobNo) {
          console.log(`  ❌ 仕事Noが取得できませんでした。スキップします。`);
        } else {
          console.log(`  ✓ 仕事No: ${jobNo}`);
          
          // サイト（媒体）を取得
          let site = '';
          if (inputSheet && companyData.row) {
            site = excelService.getCellValue(
              inputSheet,
              config.excelColumns.inputSheet.site,
              companyData.row
            ) || '';
          }
          
          // 既存の処理ロジックを再利用して書き込み（PV数・応募数は合計値を使用）
          const success = await processSingleCSVRecordForTestLoop(
            firstInvalidRecord,
            companyId,
            companyName,
            startDateStr,
            endDateStr,
            isNight,
            scrapingService,
            excelService,
            googleSheetsService,
            aiService,
            nightJobCategories,
            normalJobCategories,
            processFolderPath,
            csvRecords,
            spreadsheetId,
            { listPV: totalListPV, detailPV: totalDetailPV, webApplication: totalWebApplication, telApplication: totalTelApplication },
            uniqueIdColumn,
            site
          );
          
          if (!success) {
            console.log(`  ❌ False行の処理に失敗しました`);
          }
        }
      }
      
      // 複数行処理が完了したので、成功として返す
      return true;
    }

    // 単一行またはExcelファイルの場合の既存処理
    // 仕事Noを取得
    console.log('9. 仕事Noを取得中...');
    if (isCSV) {
      jobNo = excelService.getCSVValue(aggregatedData, '仕事No') || 
             excelService.getCSVValue(aggregatedData, '仕事番号') || '';
    } else {
      jobNo = excelService.getCellValue(
        downloadSheet,
        config.excelColumns.downloadFile.jobNo,
        2
      );
    }

    if (!jobNo) {
      console.log('  ❌ 仕事Noが取得できませんでした。スキップします。');
      
      if (onFailure && typeof onFailure === 'function') {
        onFailure(companyId, companyName || '', '', '仕事Noが取得できませんでした');
      }
      
      // TOPページに戻って入力フィールドをリセット
      try {
        await scrapingService.goToTopAndReset();
        console.log('  ✓ TOPページに戻りました\n');
      } catch (topError) {
        console.warn(`  ⚠️  TOPページへの移動でエラー: ${topError.message}`);
        // エラーが発生しても、直接URLで遷移を試みる
        try {
          const page = scrapingService.getPage();
          const topUrl = config.baitoru.loginUrl.includes('/top') 
            ? config.baitoru.loginUrl 
            : config.baitoru.loginUrl.replace(/\/$/, '') + '/top';
          await page.goto(topUrl, {
            waitUntil: 'networkidle2',
            timeout: 60000
          });
          await new Promise(resolve => setTimeout(resolve, 1000));
          console.log('  ✓ TOPページに直接遷移しました\n');
        } catch (directError) {
          console.warn(`  ⚠️  直接遷移も失敗: ${directError.message}\n`);
        }
      }
      
      return false;
    }
    console.log(`  ✓ 仕事No: ${jobNo}\n`);

    // 原稿検索とプレビューを開く（リトライ処理付き）
    console.log(`10. 仕事No: ${jobNo} で原稿を検索中...`);
    let previewPage = null;
    const maxRetries = 5;
    let retryCount = 0;
    let searchSuccess = false;

    while (retryCount < maxRetries && !searchSuccess) {
      try {
        if (retryCount > 0) {
          console.log(`  リトライ ${retryCount}/${maxRetries - 1}...`);
          await scrapingService.goToTopAndReset();
        }

        // 原稿検索ページに移動
        let navigationSuccess = false;
        for (let navRetry = 0; navRetry < 3; navRetry++) {
          try {
            await scrapingService.goToJobSearchPage();
            await new Promise(resolve => setTimeout(resolve, 1500));
            
            const page = scrapingService.getPage();
            const currentUrl = page.url();
            if (currentUrl.includes('/job') || currentUrl.includes('job?mode=1')) {
              navigationSuccess = true;
              break;
            }
          } catch (navError) {
            if (navRetry < 2) {
              console.warn(`  ページ移動をリトライします (${navRetry + 1}/3): ${navError.message}`);
              await new Promise(resolve => setTimeout(resolve, 1000));
            } else {
              throw navError;
            }
          }
        }

        if (!navigationSuccess) {
          throw new Error('原稿検索ページに移動できませんでした');
        }

        // 仕事Noで検索（リトライ時にフィールドをクリア）
        for (let searchRetry = 0; searchRetry < 3; searchRetry++) {
          try {
            // リトライ時は入力フィールドをクリア
            if (searchRetry > 0) {
              try {
                const page = scrapingService.getPage();
                await page.waitForSelector(config.selectors.jobSearch.jobNoInput, {
                  visible: true,
                  timeout: 5000
                });
                await page.click(config.selectors.jobSearch.jobNoInput, { clickCount: 3 });
                await page.keyboard.press('Backspace');
                await page.keyboard.press('Backspace');
                console.log(`  入力フィールドをクリアしました（リトライ ${searchRetry}）`);
              } catch (clearError) {
                console.warn(`  フィールドクリアをスキップ: ${clearError.message}`);
              }
            }
            
            // searchJobByNoはプレビューページを返す（PDF手順⑩⑪に基づく）
            previewPage = await scrapingService.searchJobByNo(String(jobNo));
            await new Promise(resolve => setTimeout(resolve, 2000));

            // プレビューページが正しく読み込まれたか確認
            const previewUrl = previewPage.url();
            if (previewUrl.includes('/pv') || previewUrl.includes('preview')) {
              searchSuccess = true;
              console.log('  ✓ プレビューを開きました（別タブ）');
              
              // プレビューのスクリーンショットを保存
              if (processFolderPath) {
                try {
                  await scrapingService.savePreviewScreenshot(
                    previewPage, 
                    processFolderPath, 
                    String(jobNo),
                    String(companyId),
                    startDateStr,
                    endDateStr
                  );
                } catch (screenshotError) {
                  console.warn(`  ⚠️  スクリーンショットの保存をスキップ: ${screenshotError.message}`);
                }
              }
              console.log('');
              break;
            } else {
              throw new Error(`プレビューページに遷移できませんでした: ${previewUrl}`);
            }
          } catch (searchError) {
            if (searchRetry < 2) {
              console.warn(`  検索をリトライします (${searchRetry + 1}/3): ${searchError.message}`);
              await new Promise(resolve => setTimeout(resolve, 1000));
            } else {
              throw searchError;
            }
          }
        }
      } catch (error) {
        retryCount++;
        if (retryCount >= maxRetries) {
          console.error(`  ❌ 原稿検索エラー（${maxRetries}回リトライ後）: ${error.message}`);
          if (onFailure && typeof onFailure === 'function') {
            onFailure(companyId, companyName || '', jobNo || '', `原稿検索エラー（${maxRetries}回リトライ後）: ${error.message}`);
          }
          return false;
        }
        console.warn(`  ⚠️  原稿検索エラー（リトライします）: ${error.message}`);
        await new Promise(resolve => setTimeout(resolve, 3000));
      }
    }

    if (!previewPage || !searchSuccess) {
      console.error('  ❌ プレビューページを開けませんでした');
      
      if (onFailure && typeof onFailure === 'function') {
        onFailure(companyId, companyName || '', jobNo || '', 'プレビューページを開けませんでした');
      }
      
      // TOPページに戻って入力フィールドをリセット
      try {
        await scrapingService.goToTopAndReset();
        console.log('  ✓ TOPページに戻りました\n');
      } catch (topError) {
        console.warn(`  ⚠️  TOPページへの移動でエラー: ${topError.message}`);
        // エラーが発生しても、直接URLで遷移を試みる
        try {
          const page = scrapingService.getPage();
          const topUrl = config.baitoru.loginUrl.includes('/top') 
            ? config.baitoru.loginUrl 
            : config.baitoru.loginUrl.replace(/\/$/, '') + '/top';
          await page.goto(topUrl, {
            waitUntil: 'networkidle2',
            timeout: 60000
          });
          await new Promise(resolve => setTimeout(resolve, 1000));
          console.log('  ✓ TOPページに直接遷移しました\n');
        } catch (directError) {
          console.warn(`  ⚠️  直接遷移も失敗: ${directError.message}\n`);
        }
      }
      
      return false;
    }

    // 勤務地情報を取得
    console.log('11. 勤務地情報を取得中...');
    let workLocation = null;
    for (let retry = 0; retry < 3; retry++) {
      try {
        workLocation = await scrapingService.getWorkLocation(previewPage);
        if (workLocation && (workLocation.prefecture || workLocation.city || workLocation.station)) {
          console.log(`  ✓ 都道府県: ${workLocation.prefecture || '取得失敗'}`);
          console.log(`  ✓ 市区町村: ${workLocation.city || '取得失敗'}`);
          console.log(`  ✓ 最寄り駅: ${workLocation.station || '取得失敗'}\n`);
          break;
        } else {
          throw new Error('勤務地情報が空です');
        }
      } catch (error) {
        if (retry < 2) {
          console.warn(`  勤務地情報の取得をリトライします (${retry + 1}/3): ${error.message}`);
          await new Promise(resolve => setTimeout(resolve, 2000));
          await previewPage.reload({ waitUntil: 'networkidle2' });
          await new Promise(resolve => setTimeout(resolve, 2000));
        } else {
          console.error(`  ❌ 勤務地情報の取得エラー: ${error.message}`);
          workLocation = { prefecture: '', city: '', station: '' };
        }
      }
    }

    // 職種情報を取得
    console.log('12. 職種情報を取得中...');
    let jobCategory = null;
    for (let retry = 0; retry < 3; retry++) {
      try {
        jobCategory = await scrapingService.getJobCategory(previewPage);
        if (jobCategory && (jobCategory.large || jobCategory.rawText)) {
          console.log(`  ✓ 職種大: ${jobCategory.large || '取得失敗'}`);
          console.log(`  ✓ 職種中: ${jobCategory.medium || '取得失敗'}`);
          console.log(`  ✓ 職種小: ${jobCategory.small || '取得失敗'}\n`);
          break;
        } else {
          throw new Error('職種情報が空です');
        }
      } catch (error) {
        if (retry < 2) {
          console.warn(`  職種情報の取得をリトライします (${retry + 1}/3): ${error.message}`);
          await new Promise(resolve => setTimeout(resolve, 2000));
        } else {
          console.error(`  ❌ 職種情報の取得エラー: ${error.message}`);
          jobCategory = { large: '', medium: '', small: '', rawText: '' };
        }
      }
    }

    // 給与情報を取得
    console.log('13. 給与情報を取得中...');
    let salary = null;
    for (let retry = 0; retry < 3; retry++) {
      try {
        salary = await scrapingService.getSalary(previewPage);
        if (salary && (salary.type || salary.amount > 0)) {
          console.log(`  ✓ 給与形態: ${salary.type || '取得失敗'}`);
          console.log(`  ✓ 給与額: ${salary.amount || '取得失敗'}\n`);
          break;
        } else {
          throw new Error('給与情報が空です');
        }
      } catch (error) {
        if (retry < 2) {
          console.warn(`  給与情報の取得をリトライします (${retry + 1}/3): ${error.message}`);
          await new Promise(resolve => setTimeout(resolve, 2000));
        } else {
          console.error(`  ❌ 給与情報の取得エラー: ${error.message}`);
          salary = { type: '', amount: 0 };
        }
      }
    }

    // 店名を取得（一旦未入力でOK）
    console.log('14. 店名を取得中...');
    const storeName = ''; // 店名は一旦未入力でOK
    console.log(`  ✓ 店名: 未入力（一旦OK）\n`);

    // プレビュータブを閉じる（データ取得後、すぐに閉じる）
    console.log('15. プレビュータブを閉じ中...');
    await scrapingService.closePreviewTab(previewPage);
    console.log('  ✓ プレビュータブを閉じました\n');

    // ===== エクセルにデータを書き込み =====
    console.log('16. エクセルにデータを書き込み中...');
    
    // 空白行を見つける（スプレッドシート）
    const trendRow = await googleSheetsService.findFirstEmptyRow(
      spreadsheetId,
      config.googleSheets.sheetName,
      'A',
      2
    );
    console.log(`  書き込み開始行（スプレッドシート）: ${trendRow}`);
    
    // ナイト案件と通常案件で異なる列設定を使用
    const columnConfig = isNight 
      ? config.excelColumns.trendDatabaseNight 
      : config.excelColumns.trendDatabase;
    
    // スプレッドシートへの書き込み用ヘルパー関数
    const writeCell = async (column, value) => {
      await googleSheetsService.setCellValue(
        spreadsheetId,
        config.googleSheets.sheetName,
        column,
        trendRow,
        value
      );
    };
    
    // ⑦ プランを取得して転記
    let publishingPlan = '';
    if (isCSV) {
      publishingPlan = excelService.getCSVValue(aggregatedData, '掲載プラン') || 
                     excelService.getCSVValue(aggregatedData, 'プラン') || '';
    } else {
      publishingPlan = excelService.getCellValue(
        downloadSheet,
        config.excelColumns.downloadFile.publishingPlan,
        2
      );
    }
    
    // プランを判定（PEXプラン、Bプラン、ELプラン、Dプラン、Cプラン、Aプランから選択）
    let selectedPlan = null;
    if (publishingPlan) {
      // プラン名から「〇〇プラン」の後のテキストを除去
      const cleanedPlan = cleanPlanName(publishingPlan);
      
      selectedPlan = await aiService.determinePlan(cleanedPlan);
      if (selectedPlan) {
        await writeCell(
          columnConfig.plan,
          selectedPlan
        );
        console.log(`  プランを選択: ${selectedPlan}`);
      } else {
        // AI判定が失敗した場合はクリーンなプラン名を使用
        await writeCell(
          columnConfig.plan,
          cleanedPlan
        );
        console.log(`  プランを転記: ${cleanedPlan}`);
      }
    }

    // ⑧ 一覧PV数、詳細PV数、WEB応募数、TEL応募数を転記
    if (isCSV) {
      const csvCols = config.excelColumns.downloadFile.csvColumns;
      const listPV = excelService.getCSVValue(aggregatedData, csvCols.listPV) || 0;
      const detailPV = excelService.getCSVValue(aggregatedData, csvCols.detailPV) || 0;
      const webApplication = excelService.getCSVValue(aggregatedData, csvCols.webApplication) || 0;
      // S列（TEL応募数）には、CSVデータのX列（通常応募数）を格納
      const telApplication = excelService.getCSVValue(aggregatedData, csvCols.normalApplication) || 0;
      
      await writeCell(columnConfig.listPV, listPV);
      await writeCell(columnConfig.detailPV, detailPV);
      await writeCell(columnConfig.webApplication, webApplication);
      await writeCell(columnConfig.telApplication, telApplication);
    } else {
      // Excelファイルの場合は、U~X列から取得（一覧PV数、詳細PV数、WEB応募数、TEL応募数の順）
      const targetCols = [
        columnConfig.listPV,
        columnConfig.detailPV,
        columnConfig.webApplication,
        columnConfig.telApplication
      ];
      for (let i = 0; i < 4; i++) {
        const sourceCol = String.fromCharCode(85 + i); // U, V, W, X
        const value = excelService.getCellValue(downloadSheet, sourceCol, 2);
        await writeCell(targetCols[i], value || 0);
      }
    }

    // ⑨ 期間（週数）を計算して転記（CSVファイルの申込開始日・終了日から計算）
    let appStartDate = null;
    let appEndDate = null;
    let startDateObj = null;
    let endDateObj = null;
    
    if (isCSV) {
      const csvCols = config.excelColumns.downloadFile.csvColumns;
      appStartDate = excelService.getCSVValue(aggregatedData, csvCols.applicationStartDate);
      appEndDate = excelService.getCSVValue(aggregatedData, csvCols.applicationEndDate);
      console.log(`  [DEBUG] CSVから取得した申込開始日: ${appStartDate}`);
      console.log(`  [DEBUG] CSVから取得した申込終了日: ${appEndDate}`);
    } else {
      appStartDate = excelService.getCellValue(
        downloadSheet,
        config.excelColumns.downloadFile.applicationStartDate,
        2
      );
      appEndDate = excelService.getCellValue(
        downloadSheet,
        config.excelColumns.downloadFile.applicationEndDate,
        2
      );
      console.log(`  [DEBUG] Excelから取得した申込開始日: ${appStartDate} (列: ${config.excelColumns.downloadFile.applicationStartDate}, 行: 2)`);
      console.log(`  [DEBUG] Excelから取得した申込終了日: ${appEndDate} (列: ${config.excelColumns.downloadFile.applicationEndDate}, 行: 2)`);
    }

    if (appStartDate && appEndDate) {
      startDateObj = excelDateToJSDate(appStartDate);
      endDateObj = excelDateToJSDate(appEndDate);
      
      console.log(`  [DEBUG] 変換後の申込開始日: ${startDateObj ? startDateObj.toISOString() : 'null'}`);
      console.log(`  [DEBUG] 変換後の申込終了日: ${endDateObj ? endDateObj.toISOString() : 'null'}`);
      
      if (startDateObj && endDateObj) {
        const weeks = calculateWeeks(startDateObj, endDateObj);
        console.log(`  [DEBUG] 計算された週数: ${weeks}週間`);
        await writeCell(
          columnConfig.period,
          weeks
        );
        console.log(`  期間（週数）を転記: ${weeks}週間`);
      } else {
        console.warn(`  ⚠️  日付の変換に失敗しました（開始日: ${appStartDate}, 終了日: ${appEndDate}）`);
      }
    } else {
      console.warn(`  ⚠️  申込開始日または終了日が取得できませんでした（開始日: ${appStartDate}, 終了日: ${appEndDate}）`);
    }

    // ===== スプレッドシートにデータを書き込み =====
    console.log('16. スプレッドシートにデータを書き込み中...');
    
    // 媒体列を転記（通常案件のみ、ナイト案件ではプランがG列に入るため媒体は不要）
    if (!isNight) {
      // inputSheetがnullの場合は、エクセルから読み込まずに空文字を使用
      let site = '';
      if (inputSheet && companyData.row) {
        site = excelService.getCellValue(
          inputSheet,
          config.excelColumns.inputSheet.site,
          companyData.row
        ) || '';
      }
      const siteValue = site || '';
      console.log(`  [DEBUG] サイト情報: ${siteValue} (案件種別: 通常, 行番号: ${trendRow})`);
      await writeCell(
        columnConfig.media,
        siteValue
      );
      if (site) {
        console.log(`  媒体を転記: ${site}`);
      } else {
        console.warn(`  ⚠️  サイト情報がないため、媒体列は空のままです。`);
      }
    }
    
    // ⑫ 勤務地情報を転記
    await writeCell(
      columnConfig.prefecture,
      workLocation?.prefecture || ''
    );
    await writeCell(
      columnConfig.city,
      workLocation?.city || ''
    );
    await writeCell(
      columnConfig.station,
      workLocation?.station || ''
    );

    // ⑬ 年月を転記（申込開始日から年と月を抽出）
    let year, month;
    if (startDateObj) {
      // 申込開始日から年と月を抽出
      const extracted = excelService.extractYearMonth(startDateObj);
      year = extracted.year;
      month = extracted.month;
      console.log(`  年月を算出（申込開始日から）: ${year}年${month}月`);
    } else if (appStartDate) {
      // 日付オブジェクトが取得できていない場合は、再度変換を試みる
      const convertedDate = excelDateToJSDate(appStartDate);
      if (convertedDate) {
        const extracted = excelService.extractYearMonth(convertedDate);
        year = extracted.year;
        month = extracted.month;
        console.log(`  年月を算出（申込開始日から、再変換後）: ${year}年${month}月`);
      } else {
        // 変換に失敗した場合は入力Excelの開始日を使用（フォールバック）
        const extracted = excelService.extractYearMonth(startDate);
        year = extracted.year;
        month = extracted.month;
        console.log(`  年月を算出（入力Excelの開始日から、フォールバック）: ${year}年${month}月`);
      }
    } else {
      // 申込開始日が取得できない場合は入力Excelの開始日を使用（フォールバック）
      const extracted = excelService.extractYearMonth(startDate);
      year = extracted.year;
      month = extracted.month;
      console.log(`  年月を算出（入力Excelの開始日から、フォールバック）: ${year}年${month}月`);
    }
    
    await writeCell(
      columnConfig.year,
      year
    );
    await writeCell(
      columnConfig.month,
      month
    );

    // ⑭ 地方を選択（都道府県から地方を判定）
    const regionOptions = ['北海道地方', '東北地方', '関東地方', '中部地方', '近畿地方', '中国地方', '四国地方', '九州', '沖縄地方'];
    const prefectureToRegion = {
      '北海道': '北海道地方',
      '青森県': '東北地方', '岩手県': '東北地方', '宮城県': '東北地方', '秋田県': '東北地方', '山形県': '東北地方', '福島県': '東北地方',
      '茨城県': '関東地方', '栃木県': '関東地方', '群馬県': '関東地方', '埼玉県': '関東地方', '千葉県': '関東地方', '東京都': '関東地方', '神奈川県': '関東地方',
      '新潟県': '中部地方', '富山県': '中部地方', '石川県': '中部地方', '福井県': '中部地方', '山梨県': '中部地方', '長野県': '中部地方', '岐阜県': '中部地方', '静岡県': '中部地方', '愛知県': '中部地方',
      '三重県': '近畿地方', '滋賀県': '近畿地方', '京都府': '近畿地方', '大阪府': '近畿地方', '兵庫県': '近畿地方', '奈良県': '近畿地方', '和歌山県': '近畿地方',
      '鳥取県': '中国地方', '島根県': '中国地方', '岡山県': '中国地方', '広島県': '中国地方', '山口県': '中国地方',
      '徳島県': '四国地方', '香川県': '四国地方', '愛媛県': '四国地方', '高知県': '四国地方',
      '福岡県': '九州', '佐賀県': '九州', '長崎県': '九州', '熊本県': '九州', '大分県': '九州', '宮崎県': '九州', '鹿児島県': '九州',
      '沖縄県': '沖縄地方'
    };
    
    // workLocationがnullやundefinedの場合のデフォルト値
    const safeWorkLocation = workLocation || { prefecture: '', city: '', station: '' };
    
    // 都道府県から地方を判定（ナイト案件では特に重要）
    let selectedRegion = null;
    if (safeWorkLocation.prefecture && safeWorkLocation.prefecture.trim() !== '') {
      // まずマッピングから判定
      selectedRegion = prefectureToRegion[safeWorkLocation.prefecture];
      if (!selectedRegion) {
        // マッピングにない場合はAI判定を使用
        try {
          selectedRegion = await aiService.determineRegion(safeWorkLocation.prefecture, regionOptions);
        } catch (error) {
          console.warn(`  ⚠️  地方のAI判定をスキップ: ${error.message}`);
        }
      }
      
      if (selectedRegion) {
        console.log(`  地方を選択: ${selectedRegion}`);
      } else {
        console.warn(`  ⚠️  都道府県「${safeWorkLocation.prefecture}」から地方を判定できませんでした`);
      }
    } else {
      console.warn(`  ⚠️  都道府県情報が取得できませんでした`);
    }

    // ⑮ 職種情報を取得（既に取得済み）
    
    // ⑯ 職種大・中・小を選択（キーワードマッチング → コサイン類似度で判定）
    let selectedLarge = null;
    let selectedMedium = null;
    let selectedSmall = null;
    
    const jobCategories = isNight ? nightJobCategories : normalJobCategories;
    
    if (jobCategory?.rawText && jobCategories.length > 0) {
      // 職種大から先頭の括弧と数字を除去
      const cleanedRawText = cleanJobCategoryName(jobCategory.rawText);
      const cleanedLarge = cleanJobCategoryName(jobCategory.large);
      
      // 「、」で分割して、それぞれの要素を処理
      const jobParts = cleanedRawText.split('、').map(part => part.trim()).filter(part => part);
      
      if (jobParts.length > 1) {
        console.log(`  職種を分割: ${jobParts.join(', ')}`);
      }
      
      // ナイト案件と通常案件の両方で、複数職種の場合は改善されたロジックを使用
      if (isNight || jobParts.length > 1) {
        // 各要素に対して最適マッチを記録
        const matches = [];
        
        for (const jobPart of jobParts) {
          let maxSimilarity = 0;
          let bestMatch = null;
          
          for (const category of jobCategories) {
            // 職種大、職種中、職種小を繋げたテキストを作成
            const combinedText = `${category.large} ${category.medium} ${category.small}`.trim();
            
            // コサイン類似度を計算
            const similarity = aiService.calculateCosineSimilarity(jobPart, combinedText);
            
            // 完全一致の場合はボーナス（職種小が完全一致）
            const bonus = (jobPart === category.small) ? 0.1 : 0;
            const adjustedSimilarity = similarity + bonus;
            
            if (adjustedSimilarity > maxSimilarity) {
              maxSimilarity = adjustedSimilarity;
              bestMatch = { category, similarity: adjustedSimilarity, jobPart };
            }
          }
          
          if (bestMatch && bestMatch.similarity >= 0.3) {
            matches.push(bestMatch);
          }
        }
        
        // より具体的な職種を優先するソート
        if (matches.length > 0) {
          matches.sort((a, b) => {
            // まず類似度で比較（0.05以上の差がある場合）
            if (Math.abs(a.similarity - b.similarity) > 0.05) {
              return b.similarity - a.similarity;
            }
            // 類似度が近い場合は、より具体的な（短い）職種を優先
            return a.jobPart.length - b.jobPart.length;
          });
          
          const best = matches[0].category;
          selectedLarge = best.large;
          selectedMedium = best.medium;
          selectedSmall = best.small;
          console.log(`  コサイン類似度で職種を選択（類似度: ${matches[0].similarity.toFixed(3)}）: 大=${selectedLarge}, 中=${selectedMedium}, 小=${selectedSmall}`);
        } else {
          console.warn(`  ⚠️  コサイン類似度が閾値未満のため、職種を選択できませんでした`);
        }
      } else {
        // 通常案件で単一職種の場合、キーワードマッチングを先に試す
        let keywordMatched = false;
        const jobTextLower = cleanedRawText.toLowerCase();
        
        for (const category of jobCategories) {
          // 職業大、職業中、職業小のいずれかにキーワードが含まれているかチェック
          const largeMatch = category.large && jobTextLower.includes(category.large.toLowerCase());
          const mediumMatch = category.medium && jobTextLower.includes(category.medium.toLowerCase());
          const smallMatch = category.small && jobTextLower.includes(category.small.toLowerCase());
          
          // 逆方向のチェック: カテゴリに職種テキストのキーワードが含まれているか
          const categoryLargeLower = category.large ? category.large.toLowerCase() : '';
          const categoryMediumLower = category.medium ? category.medium.toLowerCase() : '';
          const categorySmallLower = category.small ? category.small.toLowerCase() : '';
          
          // 職種テキストから主要なキーワードを抽出（2文字以上の単語）
          const keywords = cleanedRawText.split(/[・、,\s]+/).filter(k => k.length >= 2);
          
          const reverseLargeMatch = keywords.some(keyword => 
            categoryLargeLower.includes(keyword.toLowerCase())
          );
          const reverseMediumMatch = keywords.some(keyword => 
            categoryMediumLower.includes(keyword.toLowerCase())
          );
          const reverseSmallMatch = keywords.some(keyword => 
            categorySmallLower.includes(keyword.toLowerCase())
          );
          
          // いずれかにマッチした場合、その組み合わせを使用
          if (largeMatch || mediumMatch || smallMatch || reverseLargeMatch || reverseMediumMatch || reverseSmallMatch) {
            selectedLarge = category.large;
            selectedMedium = category.medium;
            selectedSmall = category.small;
            keywordMatched = true;
            console.log(`  キーワードマッチングで職種を選択: 大=${selectedLarge}, 中=${selectedMedium}, 小=${selectedSmall}`);
            break;
          }
        }
        
        // キーワードマッチングで見つからない場合、コサイン類似度で判定
        if (!keywordMatched) {
          const matchedCategory = await aiService.determineJobCategoryByCosineSimilarity(
            cleanedRawText,
            jobCategories
          );
          
          if (matchedCategory) {
            selectedLarge = matchedCategory.large;
            selectedMedium = matchedCategory.medium;
            selectedSmall = matchedCategory.small;
            console.log(`  コサイン類似度で職種を選択: 大=${selectedLarge}, 中=${selectedMedium}, 小=${selectedSmall}`);
          } else {
            // コサイン類似度で見つからない場合は、クリーンな値を使用
            selectedLarge = cleanedLarge || jobCategory.large;
            selectedMedium = jobCategory.medium;
            selectedSmall = jobCategory.small;
          }
        }
      }
    } else {
      // rawTextがない場合は、クリーンな値をそのまま使用
      selectedLarge = cleanJobCategoryName(jobCategory?.large || '');
      selectedMedium = jobCategory?.medium;
      selectedSmall = jobCategory?.small;
    }
    
    // 職種を転記
    console.log(`  [DEBUG] 職種書き込み: 大=${selectedLarge || '(空)'}, 中=${selectedMedium || '(空)'}, 小=${selectedSmall || '(空)'}`);
    if (selectedLarge) {
      await writeCell(
        columnConfig.jobCategoryLarge,
        selectedLarge
      );
      console.log(`  ✓ 職種大を転記: ${selectedLarge} (列: ${columnConfig.jobCategoryLarge})`);
    } else {
      console.warn(`  ⚠️  職種大が空のため、書き込みをスキップしました`);
    }
    if (selectedMedium) {
      await writeCell(
        columnConfig.jobCategoryMedium,
        selectedMedium
      );
      console.log(`  ✓ 職種中を転記: ${selectedMedium} (列: ${columnConfig.jobCategoryMedium})`);
    } else {
      console.warn(`  ⚠️  職種中が空のため、書き込みをスキップしました`);
    }
    if (selectedSmall) {
      await writeCell(
        columnConfig.jobCategorySmall,
        selectedSmall
      );
      console.log(`  ✓ 職種小を転記: ${selectedSmall} (列: ${columnConfig.jobCategorySmall})`);
    } else {
      console.warn(`  ⚠️  職種小が空のため、書き込みをスキップしました`);
    }

    // ⑰ 給与情報を転記
    // 給与情報を格納（通常案件の場合は特別な処理を適用）
    let salaryAmount = salary?.amount || 0;
    let salaryType = salary?.type || '';
    
    // 通常案件の場合、給与金額を特別に処理
    // 給与形態はそのまま転記する
    if (!isNight) {
      // 金額のみ処理（給与形態はそのまま）
      salaryAmount = processSalaryAmountForNormalCase(salaryAmount, salaryType);
    } else {
      // ナイト案件の場合、給与金額を処理
      if (typeof salaryAmount === 'string') {
        let salaryStr = String(salaryAmount).trim();
        
        // 「、」で区切られている場合は、1個目だけを抽出
        if (salaryStr.includes('、')) {
          salaryStr = salaryStr.split('、')[0].trim();
          console.log(`  給与（「、」で区切られているため、1個目を抽出）: ${salaryStr}`);
        }
        
        // 「時給」「月給」「日給」を抽出してL列に格納
        const salaryTypeMatch = salaryStr.match(/(時給|月給|日給)/);
        if (salaryTypeMatch) {
          salaryType = salaryTypeMatch[1];
          console.log(`  給与形態を抽出: ${salaryType}`);
        }
        
        // 「時給」「月給」「日給」という文字を除去
        salaryStr = salaryStr.replace(/時給|月給|日給/g, '').trim();
        
        // 「円」という文字を除去
        if (salaryStr.includes('円')) {
          salaryStr = salaryStr.replace(/円/g, '').trim();
        }
        
        // 「〜」「～」を除去
        salaryStr = salaryStr.replace(/[〜～]/g, '').trim();
        
        // 数値のみの場合は数値に変換
        const numericValue = parseFloat(salaryStr.replace(/[^\d.-]/g, ''));
        if (!isNaN(numericValue) && salaryStr.match(/^[\d,.\s-]+$/)) {
          // 数値のみの場合は数値として格納
          salaryAmount = numericValue;
          console.log(`  給与（数値）を格納: ${salaryAmount}`);
        } else {
          // 範囲表記（「から」など）が含まれる場合は文字列として格納
          salaryAmount = salaryStr;
          console.log(`  給与（文字列形式）を格納: ${salaryAmount}`);
        }
      }
    }
    
    // L列に給与形態を転記
    await writeCell(
      columnConfig.salaryType,
      salaryType
    );
    
    await writeCell(
      columnConfig.salaryAmount,
      salaryAmount
    );

    // ⑱ 企業ID・企業名を転記（ナイト案件はV列、W列、通常案件はX列、Y列）
    // 企業IDをCompNo列に記載
    await writeCell(
      columnConfig.compNo,
      companyId
    );
    await writeCell(
      columnConfig.companyName,
      companyName
    );

    // ⑲ 店名を転記
    await writeCell(
      columnConfig.storeName,
      storeName || ''
    );

    // ⑳ 申込開始日・終了日を転記（ナイト案件はY列、Z列、通常案件はAA列、AB列）
    console.log(`  [DEBUG] 申込開始日・終了日の書き込み処理開始`);
    console.log(`  [DEBUG] isNight: ${isNight}, columnConfig.applicationStartDate: ${columnConfig.applicationStartDate}, columnConfig.applicationEndDate: ${columnConfig.applicationEndDate}`);
    console.log(`  [DEBUG] appStartDate: ${appStartDate}, appEndDate: ${appEndDate}`);
    console.log(`  [DEBUG] startDateObj: ${startDateObj ? startDateObj.toISOString() : 'null'}, endDateObj: ${endDateObj ? endDateObj.toISOString() : 'null'}`);
    
    // 日付オブジェクトが取得できていない場合は、再度変換を試みる
    let finalStartDateObj = startDateObj;
    let finalEndDateObj = endDateObj;
    
    if (!finalStartDateObj || !finalEndDateObj) {
      if (appStartDate && appEndDate) {
        console.log(`  [DEBUG] 日付オブジェクトが取得できていないため、再度変換を試みます`);
        finalStartDateObj = excelDateToJSDate(appStartDate);
        finalEndDateObj = excelDateToJSDate(appEndDate);
        console.log(`  [DEBUG] 再変換後の申込開始日: ${finalStartDateObj ? finalStartDateObj.toISOString() : 'null'}`);
        console.log(`  [DEBUG] 再変換後の申込終了日: ${finalEndDateObj ? finalEndDateObj.toISOString() : 'null'}`);
      }
    }
    
    if (finalStartDateObj && finalEndDateObj) {
      // YYYY/MM/DD形式にフォーマット
      const formattedStartDate = formatDateForInput(finalStartDateObj);
      const formattedEndDate = formatDateForInput(finalEndDateObj);
      
      console.log(`  [DEBUG] フォーマット後の申込開始日: ${formattedStartDate}, 申込終了日: ${formattedEndDate}`);
      
      if (formattedStartDate && formattedEndDate) {
        console.log(`  [DEBUG] 書き込み実行: 開始日列=${columnConfig.applicationStartDate}, 終了日列=${columnConfig.applicationEndDate}`);
        try {
          await writeCell(
            columnConfig.applicationStartDate,
            formattedStartDate
          );
          console.log(`  ✓ 申込開始日を転記: ${formattedStartDate} (列: ${columnConfig.applicationStartDate})`);
        } catch (error) {
          console.error(`  ❌ 申込開始日の書き込みエラー: ${error.message}`);
          throw error;
        }
        
        try {
          await writeCell(
            columnConfig.applicationEndDate,
            formattedEndDate
          );
          console.log(`  ✓ 申込終了日を転記: ${formattedEndDate} (列: ${columnConfig.applicationEndDate})`);
        } catch (error) {
          console.error(`  ❌ 申込終了日の書き込みエラー: ${error.message}`);
          throw error;
        }
      } else {
        console.warn(`  ⚠️  日付のフォーマットに失敗しました（開始日: ${appStartDate}, 終了日: ${appEndDate}）`);
        console.warn(`  ⚠️  フォーマット結果: 開始日=${formattedStartDate}, 終了日=${formattedEndDate}`);
      }
    } else if (appStartDate && appEndDate) {
      console.warn(`  ⚠️  申込開始日・終了日の日付オブジェクトが取得できませんでした（開始日: ${appStartDate}, 終了日: ${appEndDate}）`);
    } else {
      console.warn(`  ⚠️  申込開始日または終了日が取得できませんでした`);
    }
    
    // 地方を転記（必ず転記する、ナイト案件では特に重要）
    // 都道府県がある場合は必ず地方を書き込む
    if (safeWorkLocation.prefecture && safeWorkLocation.prefecture.trim() !== '') {
      let regionToWrite = selectedRegion;
      
      // selectedRegionがnullの場合は、マッピングから判定
      if (!regionToWrite) {
        regionToWrite = prefectureToRegion[safeWorkLocation.prefecture];
      }
      
      // マッピングにもない場合はAI判定を試みる
      if (!regionToWrite) {
        try {
          regionToWrite = await aiService.determineRegion(safeWorkLocation.prefecture, regionOptions);
        } catch (error) {
          console.warn(`  ⚠️  地方のAI判定をスキップ: ${error.message}`);
        }
      }
      
      if (regionToWrite) {
        await writeCell(
          columnConfig.region,
          regionToWrite
        );
        console.log(`  地方を転記: ${regionToWrite}`);
      } else {
        console.warn(`  ⚠️  都道府県「${safeWorkLocation.prefecture}」から地方を判定できませんでした。地方列は空のままです。`);
      }
    } else {
      console.warn(`  ⚠️  都道府県情報がないため、地方列は空のままです。`);
    }
    
    // スプレッドシートへの書き込みは既に完了しているため、保存処理は不要
    console.log('  ✓ スプレッドシートへのデータ書き込みが完了しました\n');

    console.log(`=== ループ ${loopIndex}/${totalLoops} の動作確認が完了しました ===`);
    console.log('\n取得したデータ:');
    console.log(`- 企業ID: ${companyId}`);
    console.log(`- 企業名: ${companyName}`);
    console.log(`- 仕事No: ${jobNo}`);
    console.log(`- 都道府県: ${workLocation?.prefecture || '取得失敗'}`);
    console.log(`- 市区町村: ${workLocation?.city || '取得失敗'}`);
    console.log(`- 最寄り駅: ${workLocation?.station || '取得失敗'}`);
    console.log(`- 職種大: ${selectedLarge || jobCategory?.large || '取得失敗'}`);
    console.log(`- 職種中: ${selectedMedium || jobCategory?.medium || '取得失敗'}`);
    console.log(`- 職種小: ${selectedSmall || jobCategory?.small || '取得失敗'}`);
    console.log(`- 給与形態: ${salary?.type || '取得失敗'}`);
    console.log(`- 給与額: ${salary?.amount || '取得失敗'}`);
    console.log(`- 店名: ${storeName || '取得失敗'}`);
    console.log(`\n✓ スプレッドシートにデータを保存しました\n`);

    return true;
  } catch (error) {
    console.error(`❌ ループ ${loopIndex}/${totalLoops} でエラーが発生しました:`, error);
    return false;
  }
}

/**
 * メイン処理：エクセルから指定数分の企業IDを取得して処理を繰り返す
 */
async function main() {
  // コマンドライン引数から回数を取得
  const loopCount = parseInt(process.argv[2], 10);

  if (!loopCount || loopCount < 1 || isNaN(loopCount)) {
    console.error('❌ エラー: 有効な回数を指定してください。');
    console.log('使用方法: npm run test:loop <回数>');
    console.log('例: npm run test:loop 5');
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
    selectedFilePath = await fileSelector.selectFile();
    console.log(`✓ 選択されたファイル: ${selectedFilePath}\n`);
  } catch (error) {
    console.error(`❌ ファイル選択エラー: ${error.message}`);
    process.exit(1);
  }
  
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
    console.log(`=== ${loopCount}件の企業データを処理します ===`);
    console.log(`${'='.repeat(60)}\n`);

    // 入力ファイルを読み込み（選択されたファイルパスを使用）
    console.log('1. 入力ファイルを読み込み中...');
    const inputWorkbook = await excelService.loadInputFile(selectedFilePath);
    const inputSheet = inputWorkbook.getWorksheet(1);
    console.log('✓ 入力ファイルを読み込みました\n');

    // エクセルから指定数分の企業データを取得（2行目から開始）
    console.log(`2. エクセルから${loopCount}件の企業データを取得中...`);
    const companyDataList = [];
    const startRow = 2; // データ行の開始（1行目はヘッダー）
    
    for (let i = 0; i < loopCount; i++) {
      const row = startRow + i;
      
      const companyId = excelService.getCellValue(
        inputSheet,
        config.excelColumns.inputSheet.companyId,
        row
      );
      
      if (!companyId) {
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
      const startDate = excelDateToJSDate(startDateValue);
      const endDate = excelDateToJSDate(endDateValue);
      
      if (!startDate || !endDate) {
        console.warn(`  ⚠️  ${row}行目の日付変換に失敗しました。スキップします。`);
        continue;
      }

      companyDataList.push({
        companyId,
        companyName,
        publishingCategory,
        startDate,
        endDate,
        row
      });
    }

    if (companyDataList.length === 0) {
      console.error('❌ 処理可能な企業データがありません。');
      process.exit(1);
    }

    console.log(`✓ ${companyDataList.length}件の企業データを取得しました\n`);

    // ブラウザ起動とログイン（1回のみ）
    console.log('3. ブラウザを起動中...');
    await scrapingService.launchBrowser();
    console.log('✓ ブラウザを起動しました\n');

    console.log('4. ログイン中...');
    try {
      await scrapingService.login();
      console.log('✓ ログインしました\n');
    } catch (error) {
      console.error(`❌ ログインエラー: ${error.message}`);
      console.log('   環境変数（BAITORU_LOGIN_URL, BAITORU_USERNAME, BAITORU_PASSWORD）を確認してください。\n');
      await scrapingService.closeBrowser();
      process.exit(1);
    }

    console.log('5. TOPページに移動中...');
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
    console.log('=== ループテスト結果サマリー ===');
    console.log(`${'='.repeat(60)}`);
    console.log(`総処理企業数: ${results.total}`);
    console.log(`成功: ${results.success}`);
    console.log(`失敗: ${results.failed}`);
    console.log(`成功率: ${((results.success / results.total) * 100).toFixed(1)}%`);
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
