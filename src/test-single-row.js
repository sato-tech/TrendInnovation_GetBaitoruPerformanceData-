/**
 * 1行のみの動作確認用スクリプト
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
 * - 「円〜」とあったら、「円〜」を含めた以降の文字列をすべて取り除く
 * - 「円」があったら、「円」以降の文字列をすべて取り除く
 * - 「万」があったら数字に対して10000を掛け算する
 * - 「,」があっても文字列ではなく数字扱いする（カンマを除去して数値として扱う）
 * - 給与形態（月収、日給、時給、月給）を確実に除去する
 * @param {string|number} amount - 給与金額（文字列または数値）
 * @param {string} type - 給与形態（時給、日給、月給、月収など）
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
  
  // 「円〜」または「円～」とあったら、「円〜」を含めた以降の文字列をすべて取り除く
  if (salaryText.includes('円〜') || salaryText.includes('円～')) {
    const yenTildeIndex = salaryText.indexOf('円〜') !== -1 
      ? salaryText.indexOf('円〜') 
      : salaryText.indexOf('円～');
    salaryText = salaryText.substring(0, yenTildeIndex).trim();
  } else if (salaryText.includes('円')) {
    // 「円」があったら、「円」以降の文字列をすべて取り除く
    const yenIndex = salaryText.indexOf('円');
    salaryText = salaryText.substring(0, yenIndex).trim();
  }
  
  // 「万」が含まれているかチェック
  const hasMan = salaryText.includes('万');
  
  // カンマを除去して数値部分を抽出
  const numericValue = parseFloat(salaryText.replace(/[^\d.]/g, ''));
  
  if (!isNaN(numericValue)) {
    // 「万」があったら数字に対して10000を掛け算する
    if (hasMan) {
      return numericValue * 10000;
    } else {
      return numericValue;
    }
  } else {
    // 数値が取得できない場合は0
    return 0;
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
 * 1行のCSVデータを処理してスプレッドシートに書き込む（test-single-row.js用）
 */
async function processSingleCSVRecordForTestSingle(
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
  site,
  overrideValues = {},
  uniqueIdColumn = null
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
    
    // ユニークIDを算出してすぐに書き込む
    if (uniqueIdColumn) {
      const uniqueIdValue = `${companyId}_${companyName}_${startDateStr}_${endDateStr}`;
      await writeCell(uniqueIdColumn, uniqueIdValue);
      console.log(`  ユニークIDを転記: ${uniqueIdValue} (列: ${uniqueIdColumn}, 行: ${trendRow})`);
    }
    
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

    // 申込開始日・終了日を取得（期間計算と日付転記の両方で使用）
    const appStartDate = excelService.getCSVValue(csvRecord, csvCols.applicationStartDate);
    const appEndDate = excelService.getCSVValue(csvRecord, csvCols.applicationEndDate);
    
    // 日付オブジェクトを変換（期間計算と日付転記の両方で使用）
    let startDateObj = null;
    let endDateObj = null;
    
    if (appStartDate && appEndDate) {
      startDateObj = excelDateToJSDate(appStartDate);
      endDateObj = excelDateToJSDate(appEndDate);
    }
    
    // ⑨ 期間（週数）を計算して転記
    // overrideValuesに週数が指定されている場合はそれを使用（FALSE行の集約結果など）
    if (overrideValues.period !== undefined) {
      await writeCell(columnConfig.period, overrideValues.period);
      console.log(`  期間（週数）を転記: ${overrideValues.period}週間（集約した日付から計算）`);
    } else {
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
          await scrapingService.goToTop();
          await new Promise(resolve => setTimeout(resolve, 2000));
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
        
        // 「円」以降の値をすべて切り取る
        if (salaryStr.includes('円')) {
          const yenIndex = salaryStr.indexOf('円');
          salaryStr = salaryStr.substring(0, yenIndex).trim();
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
    // overrideValuesに日付が指定されている場合はそれを使用（FALSE行の集約結果など）
    let finalAppStartDate = overrideValues.applicationStartDate !== undefined 
      ? overrideValues.applicationStartDate 
      : appStartDate;
    let finalAppEndDate = overrideValues.applicationEndDate !== undefined 
      ? overrideValues.applicationEndDate 
      : appEndDate;
    
    if (finalAppStartDate && finalAppEndDate) {
      // overrideValuesから来た場合は既にフォーマット済みの文字列、そうでない場合はDateオブジェクトに変換
      let startDateObj, endDateObj;
      
      if (overrideValues.applicationStartDate !== undefined) {
        // 既にフォーマット済みの文字列の場合
        startDateObj = new Date(finalAppStartDate);
      } else {
        startDateObj = excelDateToJSDate(finalAppStartDate);
      }
      
      if (overrideValues.applicationEndDate !== undefined) {
        // 既にフォーマット済みの文字列の場合
        endDateObj = new Date(finalAppEndDate);
      } else {
        endDateObj = excelDateToJSDate(finalAppEndDate);
      }
      
      if (startDateObj && endDateObj && !isNaN(startDateObj.getTime()) && !isNaN(endDateObj.getTime())) {
        const formattedStartDate = overrideValues.applicationStartDate !== undefined 
          ? finalAppStartDate 
          : formatDateForInput(startDateObj);
        const formattedEndDate = overrideValues.applicationEndDate !== undefined 
          ? finalAppEndDate 
          : formatDateForInput(endDateObj);
        
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

    // プレビューページを閉じる
    await previewPage.close();
    console.log('  プレビューページを閉じました');
    
    return true;
  } catch (error) {
    console.error(`  ❌ 処理エラー: ${error.message}`);
    return false;
  }
}

/**
 * 1行のみのテスト処理
 */
async function testSingleRow() {
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
    console.log('=== 1行のみの動作確認を開始します ===\n');

    // 入力ファイルを読み込み（選択されたファイルパスを使用）
    console.log('1. 入力ファイルを読み込み中...');
    const inputWorkbook = await excelService.loadInputFile(selectedFilePath);
    const inputSheet = inputWorkbook.getWorksheet(1);
    console.log('✓ 入力ファイルを読み込みました\n');

    // 2行目のデータを確認
    const testRow = 2;
    console.log(`2. ${testRow}行目のデータを確認中...`);
    
    const companyId = excelService.getCellValue(
      inputSheet,
      config.excelColumns.inputSheet.companyId,
      testRow
    );
    const companyName = excelService.getCellValue(
      inputSheet,
      config.excelColumns.inputSheet.companyName,
      testRow
    );
    const publishingCategory = excelService.getCellValue(
      inputSheet,
      config.excelColumns.inputSheet.publishingCategory,
      testRow
    );
    const site = excelService.getCellValue(
      inputSheet,
      config.excelColumns.inputSheet.site,
      testRow
    );
    const startDateValue = excelService.getCellValue(
      inputSheet,
      config.excelColumns.inputSheet.startDate,
      testRow
    );
    const endDateValue = excelService.getCellValue(
      inputSheet,
      config.excelColumns.inputSheet.endDate,
      testRow
    );

    console.log(`   企業ID: ${companyId}`);
    console.log(`   企業名: ${companyName}`);
    console.log(`   掲載区分: ${publishingCategory}`);
    console.log(`   開始日: ${startDateValue}`);
    console.log(`   終了日: ${endDateValue}\n`);

    if (!companyId) {
      console.log('❌ 企業IDが取得できませんでした。');
      return;
    }

    // 日付を変換（ExcelJSはDateオブジェクト、数値、または文字列として読み込む可能性がある）
    const startDate = excelDateToJSDate(startDateValue);
    const endDate = excelDateToJSDate(endDateValue);
    
    if (!startDate || !endDate) {
      console.log(`❌ 日付の変換に失敗しました（開始日: ${startDateValue}, 終了日: ${endDateValue}）。`);
      return;
    }
    
    const startDateStr = formatDateForInput(startDate);
    const endDateStr = formatDateForInput(endDate);

    if (!startDateStr || !endDateStr) {
      console.log(`❌ 日付のフォーマットに失敗しました。`);
      return;
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
        return;
      }
    }

    // スプレッドシートIDの確認
    console.log('3. スプレッドシートを確認中...');
    const spreadsheetId = isNight ? config.googleSheets.spreadsheetIdNight : config.googleSheets.spreadsheetIdNormal;
    
    if (!spreadsheetId) {
      console.error(`❌ ${isNight ? 'ナイト' : '通常'}案件のスプレッドシートIDが設定されていません`);
      return;
    }
    
    console.log(`✓ ${isNight ? 'ナイト' : '通常'}案件のスプレッドシートを使用します\n`);
    
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
    
    const jobCategories = isNight ? nightJobCategories : normalJobCategories;

    // ブラウザ起動とログイン
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
      return;
    }

    console.log('6. TOPページに移動中...');
    try {
      await scrapingService.goToTop();
      console.log('✓ TOPページに移動しました（または既にTOPページにいます）\n');
    } catch (error) {
      console.warn(`⚠️  TOPページ移動で警告: ${error.message}`);
      console.log('（既にTOPページにいる可能性があります。続行します...）\n');
      // エラーが発生しても続行（既にTOPページにいる可能性があるため）
    }

    // 実行ごとのダウンロードフォルダを作成
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, -5); // YYYY-MM-DDTHH-MM-SS形式
    const downloadFolderName = `downloads_${timestamp}`;
    const downloadFolderPath = join(config.files.downloadDir, downloadFolderName);
    
    await fs.mkdir(downloadFolderPath, { recursive: true });
    console.log(`✓ ダウンロードフォルダを作成しました: ${downloadFolderPath}\n`);
    
    // ダウンロードフォルダをScrapingServiceに設定
    scrapingService.setDownloadFolder(downloadFolderPath);

    // 企業IDで検索
    console.log(`7. 企業ID: ${companyId} で検索中...`);
    try {
      await scrapingService.searchByCompanyId(String(companyId));
      console.log('✓ 検索完了\n');
    } catch (error) {
      console.error(`❌ 検索エラー: ${error.message}`);
      console.log('   セレクタ設定（config/selectors.json）を確認してください。\n');
      await scrapingService.closeBrowser();
      return;
    }

    // 選択ボタンをクリック
    console.log('8. 選択ボタンをクリック中...');
    try {
      await scrapingService.clickSelectButton();
      console.log('✓ 選択ボタンをクリックしました\n');
    } catch (error) {
      console.error(`❌ 選択ボタンクリックエラー: ${error.message}`);
      await scrapingService.closeBrowser();
      return;
    }

    // 掲載実績をダウンロード
    console.log(`9. 掲載実績をダウンロード中（${startDateStr} ～ ${endDateStr}）...`);
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
      await scrapingService.closeBrowser();
      return;
    }

    // ダウンロードファイルを読み込み（CSVまたはExcel）
    console.log('10. ダウンロードファイルを読み込み中...');
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
      await scrapingService.closeBrowser();
      return;
    }

    // バリデーション：掲載実績期間と申込期間が同日かチェック
    console.log('11. データのバリデーション中...');
    let isValid = false;
    let aggregatedData = null;

    try {
      if (isCSV) {
        if (csvRecords.length === 0) {
          console.log('  CSVファイルにデータがありません。スキップします。\n');
          await scrapingService.closeBrowser();
          return;
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
            isValid = startDateMatch && excelService.isSameDate(perfEnd, appEnd);
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
        }
      } else {
        // Excelの場合の処理（必要に応じて実装）
        console.log('  Excelファイルのバリデーションは未実装です\n');
        await scrapingService.closeBrowser();
        return;
      }
    } catch (validationError) {
      console.error(`❌ バリデーションエラー: ${validationError.message}`);
      await scrapingService.closeBrowser();
      return;
    }

    if (!isValid || !aggregatedData) {
      console.log('❌ データのバリデーションに失敗しました。スキップします。\n');
      await scrapingService.closeBrowser();
      return;
    }

    // CSVファイルが複数行の場合の処理
    if (isCSV && aggregatedData && aggregatedData.isMultiRow) {
      const { validRecords, invalidRecords } = aggregatedData;
      
      // Trueの行を1行ずつ処理
      for (let i = 0; i < validRecords.length; i++) {
        const validRecord = validRecords[i];
        console.log(`\n  [True行 ${i + 1}/${validRecords.length}] 処理を開始します`);
        
        // 仕事Noを取得
        let jobNo = excelService.getCSVValue(validRecord.record, '仕事No') || 
                   excelService.getCSVValue(validRecord.record, '仕事番号') || '';
        
        if (!jobNo) {
          console.log(`  ❌ 仕事Noが取得できませんでした。スキップします。`);
          continue;
        }
        
        console.log(`  ✓ 仕事No: ${jobNo}`);
        
        // この行のデータを処理（既存の処理ロジックを再利用）
        const success = await processSingleCSVRecordForTestSingle(
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
          site,
          {},
          uniqueIdColumn
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
        
        // 申込開始日・終了日を集約するための変数
        let earliestStartDate = null;
        let latestEndDate = null;
        
        for (const invalidRecord of invalidRecords) {
          const listPV = parseFloat(excelService.getCSVValue(invalidRecord.record, csvCols.listPV) || 0);
          const detailPV = parseFloat(excelService.getCSVValue(invalidRecord.record, csvCols.detailPV) || 0);
          const webApplication = parseFloat(excelService.getCSVValue(invalidRecord.record, csvCols.webApplication) || 0);
          const telApplication = parseFloat(excelService.getCSVValue(invalidRecord.record, csvCols.normalApplication) || 0);
          
          totalListPV += listPV;
          totalDetailPV += detailPV;
          totalWebApplication += webApplication;
          totalTelApplication += telApplication;
          
          // 申込開始日を取得して一番早い日付を記録
          const appStartDate = excelService.getCSVValue(invalidRecord.record, csvCols.applicationStartDate);
          if (appStartDate) {
            const startDateObj = excelDateToJSDate(appStartDate);
            if (startDateObj) {
              if (!earliestStartDate || startDateObj < earliestStartDate) {
                earliestStartDate = startDateObj;
              }
            }
          }
          
          // 申込終了日を取得して一番遅い日付を記録
          const appEndDate = excelService.getCSVValue(invalidRecord.record, csvCols.applicationEndDate);
          if (appEndDate) {
            const endDateObj = excelDateToJSDate(appEndDate);
            if (endDateObj) {
              if (!latestEndDate || endDateObj > latestEndDate) {
                latestEndDate = endDateObj;
              }
            }
          }
        }
        
        console.log(`  合計値: 一覧PV数=${totalListPV}, 詳細PV数=${totalDetailPV}, WEB応募数=${totalWebApplication}, TEL応募数=${totalTelApplication}`);
        
        // 集約した日付をフォーマット
        let aggregatedStartDateStr = startDateStr; // デフォルトは元の開始日
        let aggregatedEndDateStr = endDateStr; // デフォルトは元の終了日
        
        if (earliestStartDate) {
          aggregatedStartDateStr = formatDateForInput(earliestStartDate);
          console.log(`  集約した申込開始日（一番早い日付）: ${aggregatedStartDateStr}`);
        }
        
        if (latestEndDate) {
          aggregatedEndDateStr = formatDateForInput(latestEndDate);
          console.log(`  集約した申込終了日（一番遅い日付）: ${aggregatedEndDateStr}`);
        }
        
        // 集約した日付から週数を計算
        let aggregatedWeeks = null;
        if (earliestStartDate && latestEndDate) {
          aggregatedWeeks = calculateWeeks(earliestStartDate, latestEndDate);
          console.log(`  集約した日付から週数を計算: ${aggregatedWeeks}週間`);
        }
        
        // 最初のFalse行のデータを使用（その他の値）
        const firstInvalidRecord = invalidRecords[0].record;
        
        // 仕事Noを取得
        let jobNo = excelService.getCSVValue(firstInvalidRecord, '仕事No') || 
                   excelService.getCSVValue(firstInvalidRecord, '仕事番号') || '';
        
        if (!jobNo) {
          console.log(`  ❌ 仕事Noが取得できませんでした。スキップします。`);
        } else {
          console.log(`  ✓ 仕事No: ${jobNo}`);
          
          // 既存の処理ロジックを再利用して書き込み（PV数・応募数は合計値を使用、日付は集約した値を使用）
          const success = await processSingleCSVRecordForTestSingle(
            firstInvalidRecord,
            companyId,
            companyName,
            aggregatedStartDateStr, // 集約した開始日を使用
            aggregatedEndDateStr,   // 集約した終了日を使用
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
            site,
            { 
              listPV: totalListPV, 
              detailPV: totalDetailPV, 
              webApplication: totalWebApplication, 
              telApplication: totalTelApplication,
              applicationStartDate: earliestStartDate ? formatDateForInput(earliestStartDate) : null,
              applicationEndDate: latestEndDate ? formatDateForInput(latestEndDate) : null,
              period: aggregatedWeeks !== null ? aggregatedWeeks : undefined
            },
            uniqueIdColumn
          );
          
          if (!success) {
            console.log(`  ❌ False行の処理に失敗しました`);
          }
        }
      }
      
      // 複数行処理が完了したので、終了
      await scrapingService.closeBrowser();
      return;
    }

    // 単一行またはExcelファイルの場合の既存処理
    // 仕事Noを取得
    console.log('12. 仕事Noを取得中...');
    let jobNo = '';
    if (isCSV) {
      jobNo = excelService.getCSVValue(aggregatedData, '仕事No') || 
             excelService.getCSVValue(aggregatedData, '仕事番号') || '';
    }

    if (!jobNo) {
      console.log('  ❌ 仕事Noが取得できませんでした。スキップします。\n');
      await scrapingService.closeBrowser();
      return;
    }
    console.log(`  ✓ 仕事No: ${jobNo}\n`);

    // 原稿検索とプレビューを開く（リトライ処理付き）
    console.log(`13. 仕事No: ${jobNo} で原稿を検索中...`);
    let previewPage = null;
    const maxRetries = 5; // リトライ回数を増やす
    let retryCount = 0;
    let searchSuccess = false;

    while (retryCount < maxRetries && !searchSuccess) {
      try {
        if (retryCount > 0) {
          console.log(`  リトライ ${retryCount}/${maxRetries - 1}...`);
          // リトライ前にTOPページに戻る
          try {
            await scrapingService.goToTop();
            await new Promise(resolve => setTimeout(resolve, 2000)); // 2秒待機
          } catch (topError) {
            console.warn(`  TOPページへの移動をスキップ: ${topError.message}`);
          }
        }

        // 原稿検索ページに移動（リトライ処理付き）
        let navigationSuccess = false;
        for (let navRetry = 0; navRetry < 3; navRetry++) {
          try {
            await scrapingService.goToJobSearchPage();
            await new Promise(resolve => setTimeout(resolve, 1500)); // 1.5秒待機
            
            // ページが正しく読み込まれたか確認
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
            await new Promise(resolve => setTimeout(resolve, 2000)); // 2秒待機（プレビュー読み込み待ち）

            // プレビューページが正しく読み込まれたか確認
            const previewUrl = previewPage.url();
            if (previewUrl.includes('/pv') || previewUrl.includes('preview')) {
              searchSuccess = true;
              console.log('  ✓ プレビューを開きました');
              
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
              break; // 成功
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
          await scrapingService.closeBrowser();
          return;
        }
        console.warn(`  ⚠️  原稿検索エラー（リトライします）: ${error.message}`);
        await new Promise(resolve => setTimeout(resolve, 3000)); // 3秒待機してからリトライ
      }
    }

    if (!previewPage || !searchSuccess) {
      console.error('  ❌ プレビューページを開けませんでした');
      await scrapingService.closeBrowser();
      return;
    }

    // 勤務地情報を取得（リトライ処理付き）
    console.log('14. 勤務地情報を取得中...');
    let workLocation = null;
    for (let retry = 0; retry < 3; retry++) {
      try {
        workLocation = await scrapingService.getWorkLocation(previewPage);
        if (workLocation && (workLocation.prefecture || workLocation.city || workLocation.station)) {
          console.log(`  ✓ 都道府県: ${workLocation.prefecture || '取得失敗'}`);
          console.log(`  ✓ 市区町村: ${workLocation.city || '取得失敗'}`);
          console.log(`  ✓ 最寄り駅: ${workLocation.station || '取得失敗'}\n`);
          break; // 成功
        } else {
          throw new Error('勤務地情報が空です');
        }
      } catch (error) {
        if (retry < 2) {
          console.warn(`  勤務地情報の取得をリトライします (${retry + 1}/3): ${error.message}`);
          await new Promise(resolve => setTimeout(resolve, 2000));
          // ページを再読み込み
          await previewPage.reload({ waitUntil: 'networkidle2' });
          await new Promise(resolve => setTimeout(resolve, 2000));
        } else {
          console.error(`  ❌ 勤務地情報の取得エラー: ${error.message}`);
          workLocation = { prefecture: '', city: '', station: '' };
        }
      }
    }

    // workLocationがnullやundefinedの場合のデフォルト値
    const safeWorkLocation = workLocation || { prefecture: '', city: '', station: '' };

    // 職種情報を取得（リトライ処理付き）
    console.log('15. 職種情報を取得中...');
    let jobCategory = null;
    for (let retry = 0; retry < 3; retry++) {
      try {
        jobCategory = await scrapingService.getJobCategory(previewPage);
        if (jobCategory && (jobCategory.large || jobCategory.rawText)) {
          console.log(`  ✓ 職種大: ${jobCategory.large || '取得失敗'}`);
          console.log(`  ✓ 職種中: ${jobCategory.medium || '取得失敗'}`);
          console.log(`  ✓ 職種小: ${jobCategory.small || '取得失敗'}\n`);
          break; // 成功
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

    // 給与情報を取得（リトライ処理付き）
    console.log('16. 給与情報を取得中...');
    let salary = null;
    for (let retry = 0; retry < 3; retry++) {
      try {
        salary = await scrapingService.getSalary(previewPage);
        if (salary && (salary.type || salary.amount > 0)) {
          console.log(`  ✓ 給与形態: ${salary.type || '取得失敗'}`);
          console.log(`  ✓ 給与額: ${salary.amount || '取得失敗'}\n`);
          break; // 成功
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
    console.log('17. 店名を取得中...');
    const storeName = ''; // 店名は一旦未入力でOK
    console.log(`  ✓ 店名: 未入力（一旦OK）\n`);

    // プレビューページを閉じる（データ取得後、すぐに閉じる）
    console.log('18. プレビューページを閉じ中...');
    await previewPage.close();
    console.log('  ✓ プレビューページを閉じました\n');

    // ===== スプレッドシートにデータを書き込み =====
    console.log('19. スプレッドシートにデータを書き込み中...');
    
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
    
    // ⑦ プラン（H列）を取得して転記
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
          config.excelColumns.trendDatabase.plan,
          selectedPlan
        );
        console.log(`  プランを選択: ${selectedPlan}`);
      } else {
        // AI判定が失敗した場合はクリーンなプラン名を使用
        await writeCell(
          config.excelColumns.trendDatabase.plan,
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
      // Q列（TEL応募数）には、CSVデータのX列（通常応募数）を格納
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

    // ⑨ 期間（週数）を計算して転記
    let appStartDate = null;
    let appEndDate = null;
    let startDateObj = null;
    let endDateObj = null;
    
    if (isCSV) {
      const csvCols = config.excelColumns.downloadFile.csvColumns;
      appStartDate = excelService.getCSVValue(aggregatedData, csvCols.applicationStartDate);
      appEndDate = excelService.getCSVValue(aggregatedData, csvCols.applicationEndDate);
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
    }

    if (appStartDate && appEndDate) {
      startDateObj = excelDateToJSDate(appStartDate);
      endDateObj = excelDateToJSDate(appEndDate);
      
      if (startDateObj && endDateObj) {
        const weeks = calculateWeeks(startDateObj, endDateObj);
        await writeCell(
          columnConfig.period,
          weeks
        );
      }
    }

    // ⑫ 勤務地情報を転記
    await writeCell(
      config.excelColumns.trendDatabase.prefecture,
      safeWorkLocation.prefecture || ''
    );
    await writeCell(
      config.excelColumns.trendDatabase.city,
      safeWorkLocation.city || ''
    );
    await writeCell(
      config.excelColumns.trendDatabase.station,
      safeWorkLocation.station || ''
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
    let selectedRegion = null;
    const regionOptions = ['北海道地方', '東北地方', '関東地方', '中部地方', '近畿地方', '中国地方', '四国地方', '九州', '沖縄地方'];
    
    // workLocationがnullやundefinedの場合のデフォルト値（既に定義済みの場合は再定義しない）
    if (typeof safeWorkLocation === 'undefined') {
      const safeWorkLocation = workLocation || { prefecture: '', city: '', station: '' };
    }
    
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
    
    if (safeWorkLocation.prefecture && safeWorkLocation.prefecture.trim() !== '') {
      // 都道府県から地方を判定
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
    if (selectedLarge) {
      await writeCell(
        config.excelColumns.trendDatabase.jobCategoryLarge,
        selectedLarge
      );
    }
    if (selectedMedium) {
      await writeCell(
        config.excelColumns.trendDatabase.jobCategoryMedium,
        selectedMedium
      );
    }
    if (selectedSmall) {
      await writeCell(
        config.excelColumns.trendDatabase.jobCategorySmall,
        selectedSmall
      );
    }

    // ⑰ 給与情報を転記
    // 給与金額を処理
    let salaryAmount = salary?.amount || 0;
    let salaryType = salary?.type || '';
    
    if (!isNight) {
      // 通常案件の場合、給与金額を特別に処理
      // 給与形態はそのまま転記する
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
        
        // 「円」以降の値をすべて切り取る
        if (salaryStr.includes('円')) {
          const yenIndex = salaryStr.indexOf('円');
          salaryStr = salaryStr.substring(0, yenIndex).trim();
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

    // ⑱ 企業ID・企業名を転記
    await writeCell(
      config.excelColumns.trendDatabase.companyId,
      companyId
    );
    await writeCell(
      config.excelColumns.trendDatabase.companyName,
      companyName
    );

    // ⑲ 店名を転記
    await writeCell(
      config.excelColumns.trendDatabase.storeName,
      storeName || ''
    );

    // 媒体列を転記（入力Excelから取得）
    await writeCell(
      config.excelColumns.trendDatabase.media,
      site || ''
    );
    if (site) {
      console.log(`  媒体を転記: ${site}`);
    } else {
      console.warn(`  ⚠️  サイト情報がないため、媒体列は空のままです。`);
    }

    // ⑳ 申込開始日・終了日を転記
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
        await writeCell(
          columnConfig.applicationStartDate,
          formattedStartDate
        );
        await writeCell(
          columnConfig.applicationEndDate,
          formattedEndDate
        );
        console.log(`  ✓ 申込開始日を転記: ${formattedStartDate} (列: ${columnConfig.applicationStartDate})`);
        console.log(`  ✓ 申込終了日を転記: ${formattedEndDate} (列: ${columnConfig.applicationEndDate})`);
      } else {
        console.warn(`  ⚠️  日付のフォーマットに失敗しました（開始日: ${appStartDate}, 終了日: ${appEndDate}）`);
        console.warn(`  ⚠️  フォーマット結果: 開始日=${formattedStartDate}, 終了日=${formattedEndDate}`);
      }
    } else if (appStartDate && appEndDate) {
      console.warn(`  ⚠️  申込開始日・終了日の日付オブジェクトが取得できませんでした（開始日: ${appStartDate}, 終了日: ${appEndDate}）`);
    } else {
      console.warn(`  ⚠️  申込開始日または終了日が取得できませんでした`);
    }
    
    // 地方を転記（必ず転記する）
    if (selectedRegion) {
      await writeCell(
        config.excelColumns.trendDatabase.region,
        selectedRegion
      );
      console.log(`  地方を転記: ${selectedRegion}`);
    } else if (safeWorkLocation.prefecture && safeWorkLocation.prefecture.trim() !== '') {
      // selectedRegionがnullでも、都道府県がある場合は再度判定を試みる
      const fallbackRegion = prefectureToRegion[safeWorkLocation.prefecture];
      if (fallbackRegion) {
        await writeCell(
          config.excelColumns.trendDatabase.region,
          fallbackRegion
        );
        console.log(`  地方を転記（フォールバック）: ${fallbackRegion}`);
      } else {
        console.warn(`  ⚠️  都道府県「${safeWorkLocation.prefecture}」から地方を判定できませんでした。地方列は空のままです。`);
      }
    } else {
      console.warn(`  ⚠️  都道府県情報がないため、地方列は空のままです。`);
    }
    
    // スプレッドシートへの書き込みは既に完了しているため、保存処理は不要
    console.log('  ✓ スプレッドシートへのデータ書き込みが完了しました\n');

    console.log('=== 動作確認が完了しました ===');
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

  } catch (error) {
    console.error('❌ エラーが発生しました:', error);
    throw error;
  } finally {
    // ブラウザを閉じる
    await scrapingService.closeBrowser();
    console.log('ブラウザを閉じました');
  }
}

// 実行
testSingleRow().catch(error => {
  console.error('致命的なエラー:', error);
  process.exit(1);
});
