/**
 * メインエントリーポイント
 * 反響事例の取説に基づいた業務フローを自動化（手順①〜⑳）
 */

import ScrapingService from './services/ScrapingService.js';
import ExcelService from './services/ExcelService.js';
import GoogleSheetsService from './services/GoogleSheetsService.js';
import AIService from './services/AIService.js';
import config from '../config/config.js';
import { excelDateToJSDate, formatDateForInput, calculateWeeks } from './utils/dateUtils.js';
import { promises as fs } from 'fs';
import { join } from 'path';

/**
 * 都道府県名から先頭のラベル（勤務地、面接地など）を除去
 * @param {string} prefecture - 都道府県名（ラベル付きの可能性がある）
 * @returns {string} クリーンな都道府県名
 */
function cleanPrefectureName(prefecture) {
  if (!prefecture) return '';
  
  // 先頭のラベル（勤務地、面接地など）を除去
  const cleaned = prefecture
    .replace(/^(勤務地|面接地|所在地)[:：\s]*/i, '')
    .replace(/^[^都道府県]*([^都道府県]*[都道府県])/i, '$1')
    .trim();
  
  return cleaned;
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
 * 通常案件の給与金額を処理する
 * - 「月収・日給・時給」というテキストは切り離す
 * - 「、」で区切られた場合は、最初のテキストのみを抽出
 * - 「〇〇円」と記載されている場合：数字のみを抽出
 * - 「○○万円」や「○○万円〜〇〇万円」という表記の場合：そのまま文字列として格納
 * - 「月収○○万円」や「月収○○万円〜〇〇万円」の場合：「月収」と「○○万」や「○○万円〜〇〇万」を分離
 * - 給与形態（月収、日給、時給、月給）を確実に除去する
 * @param {string|number} amount - 給与金額（文字列または数値）
 * @param {string} type - 給与形態（時給、日給、月給、月収など）
 * @returns {string|number} 処理後の給与金額
 */
function processSalaryAmountForNormalCase(amount, type) {
  let processedAmount;
  
  if (typeof amount === 'number') {
    // 数値の場合はそのまま返す
    processedAmount = amount;
  } else {
    let salaryText = String(amount || '').trim();
    if (!salaryText) {
      processedAmount = 0;
    } else {
      // 先頭にある「時給」「月給」「日給」までのテキストを取り除く
      salaryText = salaryText.replace(/^(時給|月給|日給|月収)[\s・、]*/i, '').trim();
      
      // 「、」で区切られた場合は、最初のテキストのみを抽出
      if (salaryText.includes('、')) {
        salaryText = salaryText.split('、')[0].trim();
      }
      
      // 「万円」が含まれている場合は、そのままテキストとして記入
      if (salaryText.includes('万円')) {
        processedAmount = salaryText.trim();
      } else if (salaryText.includes('円')) {
        // 「円」を含めて取り除いて数字だけにする
        const amountMatch = salaryText.match(/([\d,]+)\s*円/);
        if (amountMatch) {
          processedAmount = parseInt(amountMatch[1].replace(/,/g, ''), 10);
        } else {
          // 「円」が含まれているが数値が取得できない場合は、数値部分のみを抽出
          const numericValue = parseFloat(salaryText.replace(/[^\d.]/g, ''));
          processedAmount = !isNaN(numericValue) ? numericValue : 0;
        }
      } else {
        // 「円」が含まれていない場合は、数値のみを抽出
        const numericValue = parseFloat(salaryText.replace(/[^\d.]/g, ''));
        processedAmount = !isNaN(numericValue) ? numericValue : 0;
      }
    }
  }
  
  return processedAmount;
}

/**
 * 1行のCSVデータを処理してスプレッドシートに書き込む
 * @param {Object} csvRecord - CSVレコード
 * @param {string} companyId - 企業ID
 * @param {string} companyName - 企業名
 * @param {string} startDateStr - 掲載開始日（YYYY/MM/DD形式）
 * @param {string} endDateStr - 掲載終了日（YYYY/MM/DD形式）
 * @param {string} site - サイト（媒体）
 * @param {boolean} isNight - ナイト案件かどうか
 * @param {Object} columnConfig - 列設定
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {Function} writeCell - セル書き込み関数
 * @param {number} trendRow - 書き込み行番号
 * @param {ScrapingService} scrapingService - スクレイピングサービス
 * @param {ExcelService} excelService - Excelサービス
 * @param {GoogleSheetsService} googleSheetsService - Google Sheetsサービス
 * @param {AIService} aiService - AIサービス
 * @param {Array} nightJobCategories - ナイト案件の職種カテゴリ
 * @param {Array} normalJobCategories - 通常案件の職種カテゴリ
 * @param {string} processFolderPath - 処理フォルダパス
 * @param {Array} csvRecords - CSVレコード全体（参照用）
 * @param {Object} overrideValues - 上書きする値（PV数・応募数など）
 */
async function processSingleCSVRecord(
  csvRecord,
  companyId,
  companyName,
  startDateStr,
  endDateStr,
  site,
  isNight,
  columnConfig,
  spreadsheetId,
  writeCell,
  trendRow,
  scrapingService,
  excelService,
  googleSheetsService,
  aiService,
  nightJobCategories,
  normalJobCategories,
  processFolderPath,
  csvRecords,
  overrideValues = {},
  uniqueIdColumn = null
) {
  const csvCols = config.excelColumns.downloadFile.csvColumns;
  
  // ⑦ プラン（H列）を取得して転記
  let publishingPlan = excelService.getCSVValue(csvRecord, '掲載プラン') || 
                     excelService.getCSVValue(csvRecord, 'プラン') || '';
  
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
  
  if (appStartDate && appEndDate) {
    const startDateObj = excelDateToJSDate(appStartDate);
    const endDateObj = excelDateToJSDate(appEndDate);
    
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
    return;
  }

  // ⑩⑪ 原稿検索とプレビューを開く
  console.log(`  仕事No: ${jobNo} で原稿を検索します`);
  const previewPage = await scrapingService.searchJobByNo(String(jobNo));
  console.log('  プレビューを開きました');

  // プレビューのスクリーンショットを保存
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

  // プレビューページからデータを取得
  console.log('  プレビューページからデータを取得中...');
  
  // ⑫ 勤務地情報を取得
  const workLocation = await scrapingService.getWorkLocation(previewPage);
  const safeWorkLocation = workLocation || { prefecture: '', city: '', station: '' };
  
  // 都道府県はそのまま転記する（処理を削除）

  // ⑬ 年月を計算
  let year, month;
  const appStartDateForYearMonth = excelService.getCSVValue(csvRecord, csvCols.applicationStartDate);
  if (appStartDateForYearMonth) {
    const perfStartDateObj = excelDateToJSDate(appStartDateForYearMonth);
    if (perfStartDateObj) {
      const extracted = excelService.extractYearMonth(perfStartDateObj);
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
  const regionOptions = ['北海道', '東北地方', '関東地方', '中部地方', '近畿地方', '中国地方', '四国地方', '九州', '沖縄地方'];
  let selectedRegion = null;
  if (safeWorkLocation.prefecture) {
    const prefectureToRegion = {
      '北海道': '北海道',
      '青森県': '東北地方', '岩手県': '東北地方', '宮城県': '東北地方', '秋田県': '東北地方', '山形県': '東北地方', '福島県': '東北地方',
      '茨城県': '関東地方', '栃木県': '関東地方', '群馬県': '関東地方', '埼玉県': '関東地方', '千葉県': '関東地方', '東京都': '関東地方', '神奈川県': '関東地方',
      '新潟県': '中部地方', '富山県': '中部地方', '石川県': '中部地方', '福井県': '中部地方', '山梨県': '中部地方', '長野県': '中部地方', '岐阜県': '中部地方', '静岡県': '中部地方', '愛知県': '中部地方',
      '三重県': '近畿地方', '滋賀県': '近畿地方', '京都府': '近畿地方', '大阪府': '近畿地方', '兵庫県': '近畿地方', '奈良県': '近畿地方', '和歌山県': '近畿地方',
      '鳥取県': '中国地方', '島根県': '中国地方', '岡山県': '中国地方', '広島県': '中国地方', '山口県': '中国地方',
      '徳島県': '四国地方', '香川県': '四国地方', '愛媛県': '四国地方', '高知県': '四国地方',
      '福岡県': '九州', '佐賀県': '九州', '長崎県': '九州', '熊本県': '九州', '大分県': '九州', '宮崎県': '九州', '鹿児島県': '九州',
      '沖縄県': '沖縄地方'
    };
    
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
  const jobCategory = await scrapingService.getJobCategory(previewPage);
  
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
  const salary = await scrapingService.getSalary(previewPage);
  
  // 給与形態と金額を処理
  // 給与形態はそのまま転記する
  let salaryType = salary.type || '';
  let salaryAmount = salary.amount || 0;
  
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

  // ユニークIDを転記（ナイト案件と通常案件の両方）
  if (uniqueIdColumn) {
    const uniqueIdValue = `${companyId}_${companyName}_${startDateStr}_${endDateStr}`;
    await writeCell(uniqueIdColumn, uniqueIdValue);
    console.log(`  ユニークIDを転記: ${uniqueIdValue} (列: ${uniqueIdColumn})`);
  }

  // ⑳ 申込開始日・終了日を転記
  if (appStartDate && appEndDate) {
    const startDateObj = excelDateToJSDate(appStartDate);
    const endDateObj = excelDateToJSDate(appEndDate);
    
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

  // プレビューページを閉じる
  await scrapingService.closePreviewTab(previewPage);
  console.log('  プレビューページを閉じました');
}

/**
 * メイン処理
 */
async function main() {
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

  // 実行ごとの出力フォルダを作成
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, -5); // YYYY-MM-DDTHH-MM-SS形式
  const outputFolderName = `run_${timestamp}`;
  const outputFolderPath = join(config.files.outputDir, outputFolderName);

  // 実行ごとのダウンロードフォルダを作成
  const downloadFolderName = `downloads_${timestamp}`;
  const downloadFolderPath = join(config.files.downloadDir, downloadFolderName);

  try {
    console.log('処理を開始します...');
    
    // ダウンロードフォルダを作成
    await fs.mkdir(downloadFolderPath, { recursive: true });
    console.log(`✓ ダウンロードフォルダを作成しました: ${downloadFolderPath}`);

    // ブラウザ起動
    await scrapingService.launchBrowser();
    
    // ダウンロードフォルダをScrapingServiceに設定
    scrapingService.setDownloadFolder(downloadFolderPath);
    console.log('ブラウザを起動しました');

    // ログイン
    await scrapingService.login();
    console.log('ログインしました');

    // TOPページに移動
    await scrapingService.goToTop();
    console.log('TOPページに移動しました');

    // 入力ファイルを読み込み
    const inputWorkbook = await excelService.loadInputFile();
    const inputSheet = inputWorkbook.getWorksheet(1); // 最初のシートを取得
    console.log('✓ 入力ファイルを読み込みました');

    // スプレッドシートIDの確認
    if (!config.googleSheets.spreadsheetIdNight || !config.googleSheets.spreadsheetIdNormal) {
      console.error('❌ スプレッドシートIDが設定されていません');
      console.error('GOOGLE_SPREADSHEET_ID_NIGHT と GOOGLE_SPREADSHEET_ID_NORMAL を.envファイルに設定してください。');
      process.exit(1);
    }
    
    console.log('✓ スプレッドシートを使用します');

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

    // 重複チェック用の列を追加（ナイト案件と通常案件の両方）
    console.log('重複チェック用の列を追加中...');
    let duplicateCheckColumnNight = null;
    let duplicateCheckColumnNormal = null;
    try {
      duplicateCheckColumnNight = await googleSheetsService.ensureDuplicateCheckColumn(
        config.googleSheets.spreadsheetIdNight,
        config.googleSheets.sheetName,
        '重複チェック'
      );
      console.log(`✓ ナイト案件の重複チェック列: ${duplicateCheckColumnNight}`);
    } catch (error) {
      console.warn(`⚠️  ナイト案件の重複チェック列の追加エラー: ${error.message}`);
    }
    
    try {
      duplicateCheckColumnNormal = await googleSheetsService.ensureDuplicateCheckColumn(
        config.googleSheets.spreadsheetIdNormal,
        config.googleSheets.sheetName,
        '重複チェック'
      );
      console.log(`✓ 通常案件の重複チェック列: ${duplicateCheckColumnNormal}`);
    } catch (error) {
      console.warn(`⚠️  通常案件の重複チェック列の追加エラー: ${error.message}`);
    }

    // ユニークID列を検索（既に手動で追加済みのため、追加処理は行わない）
    console.log('ユニークID列を検索中...');
    let uniqueIdColumnNight = null;
    let uniqueIdColumnNormal = null;
    try {
      uniqueIdColumnNight = await googleSheetsService.findColumnByName(
        config.googleSheets.spreadsheetIdNight,
        config.googleSheets.sheetName,
        'ユニークID'
      );
      if (uniqueIdColumnNight) {
        console.log(`✓ ナイト案件のユニークID列: ${uniqueIdColumnNight}`);
      } else {
        console.warn(`⚠️  ナイト案件のユニークID列が見つかりませんでした`);
      }
    } catch (error) {
      console.warn(`⚠️  ナイト案件のユニークID列の検索エラー: ${error.message}`);
    }
    
    try {
      uniqueIdColumnNormal = await googleSheetsService.findColumnByName(
        config.googleSheets.spreadsheetIdNormal,
        config.googleSheets.sheetName,
        'ユニークID'
      );
      if (uniqueIdColumnNormal) {
        console.log(`✓ 通常案件のユニークID列: ${uniqueIdColumnNormal}`);
      } else {
        console.warn(`⚠️  通常案件のユニークID列が見つかりませんでした`);
      }
    } catch (error) {
      console.warn(`⚠️  通常案件のユニークID列の検索エラー: ${error.message}`);
    }

    // 空白行を見つける（ナイト案件と通常案件の両方）
    console.log('空白行を検索中...');
    const nextTrendRowNight = await googleSheetsService.findFirstEmptyRow(
      config.googleSheets.spreadsheetIdNight,
      config.googleSheets.sheetName,
      'A',
      2
    );
    const nextTrendRowNormal = await googleSheetsService.findFirstEmptyRow(
      config.googleSheets.spreadsheetIdNormal,
      config.googleSheets.sheetName,
      'A',
      2
    );
    console.log(`✓ ナイト案件の書き込み開始行: ${nextTrendRowNight}`);
    console.log(`✓ 通常案件の書き込み開始行: ${nextTrendRowNormal}`);

    // データ行を取得（ヘッダー行を除く）
    let row = 2; // 2行目から開始（1行目はヘッダーと仮定）
    let hasData = true;
    let processedCount = 0;
    let skippedCount = 0;
    let duplicateSkippedCount = 0;

    while (hasData) {
      try {
        // ① 企業IDを取得（E列）
        const companyId = excelService.getCellValue(
          inputSheet,
          config.excelColumns.inputSheet.companyId,
          row
        );

        if (!companyId) {
          hasData = false;
          break;
        }

        console.log(`\n[${row - 1}行目] 企業ID: ${companyId} の処理を開始します`);

        // 掲載区分を確認（P列）
        const publishingCategory = excelService.getCellValue(
          inputSheet,
          config.excelColumns.inputSheet.publishingCategory,
          row
        );

        // ナイト案件かどうかを判定
        const isNight = publishingCategory === 'ナイト';
        const jobCategories = isNight ? nightJobCategories : normalJobCategories;

        // 開始日・終了日を取得（M・N列）
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
        
        // 企業名を取得（F列）
        const companyName = excelService.getCellValue(
          inputSheet,
          config.excelColumns.inputSheet.companyName,
          row
        ) || '';

        // サイト（媒体）を取得（A列）
        const site = excelService.getCellValue(
          inputSheet,
          config.excelColumns.inputSheet.site,
          row
        );
        console.log(`  [DEBUG] 入力Excelからサイト情報を取得: ${site || '(空)'} (行: ${row}, 列: A)`);

        if (!startDateValue || !endDateValue) {
          console.log('開始日または終了日が設定されていません。スキップします。');
          row++;
          skippedCount++;
          continue;
        }

        // 日付を変換（ExcelJSはDateオブジェクト、数値、または文字列として読み込む可能性がある）
        const startDate = excelDateToJSDate(startDateValue);
        const endDate = excelDateToJSDate(endDateValue);
        
        // 日付をフォーマット（YYYY/MM/DD形式）
        const checkStartDateStr = startDate ? formatDateForInput(startDate) : '';
        const checkEndDateStr = endDate ? formatDateForInput(endDate) : '';
        
        // ユニークIDを算出: 企業ID、企業名、開始日、終了日を結合したユニークな値を生成
        const uniqueId = `${companyId}_${companyName}_${checkStartDateStr}_${checkEndDateStr}`;
        
        // 対象の行がナイト案件か通常案件かで、参照するスプレッドシートを分岐する
        // （publishingCategoryとisNightは既に上で取得済み）
        const checkSpreadsheetId = isNight ? config.googleSheets.spreadsheetIdNight : config.googleSheets.spreadsheetIdNormal;
        const uniqueIdColumn = isNight ? uniqueIdColumnNight : uniqueIdColumnNormal;
        
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
            row++;
            duplicateSkippedCount++;
            continue;
          }
        }
        
        if (!startDate || !endDate) {
          console.log(`警告: 日付の変換に失敗しました（開始日: ${startDateValue}, 終了日: ${endDateValue}）。スキップします。`);
          row++;
          skippedCount++;
          continue;
        }
        
        const startDateStr = formatDateForInput(startDate);
        const endDateStr = formatDateForInput(endDate);

        if (!startDateStr || !endDateStr) {
          console.log(`警告: 日付のフォーマットに失敗しました。スキップします。`);
          row++;
          skippedCount++;
          continue;
        }

        console.log(`  掲載期間: ${startDateStr} ～ ${endDateStr}`);

        // ③ 企業IDで検索
        await scrapingService.searchByCompanyId(String(companyId));
        console.log('  企業IDで検索しました');

        // ④ 選択ボタンをクリック
        await scrapingService.clickSelectButton();
        console.log('  選択ボタンをクリックしました');

        // ⑤ 掲載実績をダウンロード（M・N列の開始日・終了日を指定）
        const downloadResult = await scrapingService.downloadPerformance(
          startDateStr,
          endDateStr,
          String(companyId)
        );
        const downloadFilePath = downloadResult.filePath;
        const processFolderPath = downloadResult.folderPath;
        console.log(`  掲載実績をダウンロードしました: ${downloadFilePath}`);
        console.log(`  処理フォルダ: ${processFolderPath}`);

        // ダウンロードファイルを読み込み（CSVまたはExcel）
        let csvRecords = null;
        let downloadSheet = null;
        const isCSV = downloadFilePath.endsWith('.csv');

        console.log(`  ダウンロードファイルを読み込みます（形式: ${isCSV ? 'CSV' : 'Excel'}）...`);
        
        try {
          if (isCSV) {
            csvRecords = await excelService.loadCSVFile(downloadFilePath);
            console.log(`  ✓ CSVファイルを読み込みました（${csvRecords.length}件のレコード）`);
          } else {
            const downloadWorkbook = await excelService.loadDownloadFile(downloadFilePath);
            downloadSheet = downloadWorkbook.getWorksheet(1);
            console.log('  ✓ Excelファイルを読み込みました');
          }
        } catch (fileError) {
          console.error(`  ファイル読み込みエラー: ${fileError.message}`);
          console.error(`  ファイルパス: ${downloadFilePath}`);
          throw new Error(`ダウンロードファイルの読み込みに失敗しました: ${fileError.message}`);
        }

        // ⑥ バリデーション：掲載実績期間と申込期間が同日かチェック
        let isValid = false;
        let aggregatedData = null;

        if (isCSV) {
          // CSVの場合：全レコードをチェック
          if (csvRecords.length === 0) {
            console.log('  CSVファイルにデータがありません（ヘッダーのみ）。スキップします。');
            
            // TOPページに戻って入力フィールドをリセット
            try {
              await scrapingService.goToTopAndReset();
              console.log('  ✓ TOPページに戻りました');
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
                console.log('  ✓ TOPページに直接遷移しました');
              } catch (directError) {
                console.warn(`  ⚠️  直接遷移も失敗: ${directError.message}`);
              }
            }
            
            row++;
            skippedCount++;
            continue;
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
          console.log(`  日付不一致（False）: ${invalidRecords.length}件`);

          // Trueの行がある場合は、1行ずつ処理する
          if (validRecords.length > 0) {
            console.log('  Trueの行を1行ずつ処理します。');
            // 後で処理するため、ここではフラグを設定
            isValid = true;
            aggregatedData = { validRecords, invalidRecords, isMultiRow: true };
          } else if (invalidRecords.length > 0) {
            // Falseの行のみの場合
            console.log('  Falseの行のみです。合計値を算出して書き込みます。');
            isValid = true;
            aggregatedData = { validRecords, invalidRecords, isMultiRow: true };
          } else {
            // レコードがない場合（通常は発生しない）
            console.log('  レコードがありません。');
            isValid = false;
          }
        } else {
          // Excelの場合：2行目（データ行）をチェック
          const perfStartDate = excelService.getCellValue(
            downloadSheet,
            config.excelColumns.downloadFile.performanceStartDate,
            2
          );
          const perfEndDate = excelService.getCellValue(
            downloadSheet,
            config.excelColumns.downloadFile.performanceEndDate,
            2
          );
          const appStartDate = excelService.getCellValue(
            downloadSheet,
            config.excelColumns.downloadFile.applicationStartDate,
            2
          );
          const appEndDate = excelService.getCellValue(
            downloadSheet,
            config.excelColumns.downloadFile.applicationEndDate,
            2
          );

          // 開始日が同日かチェック
          const startDateMatch = excelService.isSameDate(perfStartDate, appStartDate);
          
          // 終了日が「掲載中」の場合は特別扱い（開始日が同日ならOK）
          const perfEndStr = String(perfEndDate || '').trim();
          const isPublishing = perfEndStr === '掲載中';
          
          if (isPublishing) {
            // 「掲載中」の場合は、開始日が同日であればOK
            if (startDateMatch) {
              isValid = true;
              aggregatedData = { row: 2, sheet: downloadSheet };
              console.log('  掲載実績開始日と申込開始日が同日で、掲載実績終了日が「掲載中」です。');
            } else {
              console.log('  掲載実績開始日と申込開始日が異なります。');
            }
          } else {
            // 通常の場合は、開始日と終了日が両方とも同日かチェック
            if (startDateMatch && excelService.isSameDate(perfEndDate, appEndDate)) {
              isValid = true;
              aggregatedData = { row: 2, sheet: downloadSheet };
              console.log('  掲載実績日と申込日が同日です。');
            } else {
              console.log('  掲載実績日と申込日が異なります。');
              // Excelの場合、複数行の合計処理が必要な場合は追加実装
            }
          }
        }

        if (!isValid) {
          console.log('  バリデーションに失敗しました（掲載実績期間と申込期間が一致しません）。次の企業に進みます。');
          
          // TOPページに戻って入力フィールドをリセット
          try {
            await scrapingService.goToTopAndReset();
            console.log('  ✓ TOPページに戻りました');
          } catch (topError) {
            console.warn(`  ⚠️  TOPページへの移動をスキップ: ${topError.message}`);
          }
          
          row++;
          skippedCount++;
          continue;
        }

        console.log('  バリデーション成功。データを処理します。');

        // CSVファイルが複数行の場合の処理
        if (isCSV && aggregatedData && aggregatedData.isMultiRow) {
          const { validRecords, invalidRecords } = aggregatedData;
          
          // Trueの行を1行ずつ処理
          for (let i = 0; i < validRecords.length; i++) {
            const validRecord = validRecords[i];
            console.log(`\n  [True行 ${i + 1}/${validRecords.length}] 処理を開始します`);
            
            // 現在の行番号を取得（ナイト案件と通常案件で別々に管理）
            const trendRow = isNight ? nextTrendRowNight : nextTrendRowNormal;
            
            // デバッグ: 書き込み先の確認
            console.log(`  [DEBUG] 書き込み先: ${isNight ? 'ナイト案件' : '通常案件'}, 行番号: ${trendRow}`);

            // ナイト案件と通常案件で異なる列設定を使用
            const columnConfig = isNight 
              ? config.excelColumns.trendDatabaseNight 
              : config.excelColumns.trendDatabase;

            // スプレッドシートへの書き込み用ヘルパー関数
            const spreadsheetId = isNight ? config.googleSheets.spreadsheetIdNight : config.googleSheets.spreadsheetIdNormal;
            const writeCell = async (column, value) => {
              await googleSheetsService.setCellValue(
                spreadsheetId,
                config.googleSheets.sheetName,
                column,
                trendRow,
                value
              );
            };

            // 企業名を取得
            const companyName = excelService.getCellValue(
              inputSheet,
              config.excelColumns.inputSheet.companyName,
              row
            ) || '';
            
            // この行のデータを処理（既存の処理ロジックを再利用）
            const currentUniqueIdColumn = isNight ? uniqueIdColumnNight : uniqueIdColumnNormal;
            await processSingleCSVRecord(
              validRecord.record,
              companyId,
              companyName,
              startDateStr,
              endDateStr,
              site,
              isNight,
              columnConfig,
              spreadsheetId,
              writeCell,
              trendRow,
              scrapingService,
              excelService,
              googleSheetsService,
              aiService,
              nightJobCategories,
              normalJobCategories,
              processFolderPath,
              csvRecords,
              {},
              currentUniqueIdColumn
            );
            
            // 行番号を更新
            if (isNight) {
              nextTrendRowNight++;
            } else {
              nextTrendRowNormal++;
            }
          }
          
          // Falseの行が複数ある場合、合計値を算出して書き込む
          if (invalidRecords.length > 0) {
            console.log(`\n  [False行] ${invalidRecords.length}件の合計値を算出して書き込みます`);
            
            // 現在の行番号を取得
            const trendRow = isNight ? nextTrendRowNight : nextTrendRowNormal;
            
            // デバッグ: 書き込み先の確認
            console.log(`  [DEBUG] 書き込み先: ${isNight ? 'ナイト案件' : '通常案件'}, 行番号: ${trendRow}`);

            // ナイト案件と通常案件で異なる列設定を使用
            const columnConfig = isNight 
              ? config.excelColumns.trendDatabaseNight 
              : config.excelColumns.trendDatabase;

            // スプレッドシートへの書き込み用ヘルパー関数
            const spreadsheetId = isNight ? config.googleSheets.spreadsheetIdNight : config.googleSheets.spreadsheetIdNormal;
            const writeCell = async (column, value) => {
              await googleSheetsService.setCellValue(
                spreadsheetId,
                config.googleSheets.sheetName,
                column,
                trendRow,
                value
              );
            };

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
            
            // 企業名を取得
            const companyName = excelService.getCellValue(
              inputSheet,
              config.excelColumns.inputSheet.companyName,
              row
            ) || '';
            
            // 最初のFalse行のデータを使用（その他の値）
            const firstInvalidRecord = invalidRecords[0].record;
            
            // 既存の処理ロジックを再利用して書き込み（PV数・応募数は合計値を使用）
            const currentUniqueIdColumn = isNight ? uniqueIdColumnNight : uniqueIdColumnNormal;
            await processSingleCSVRecord(
              firstInvalidRecord,
              companyId,
              companyName,
              startDateStr,
              endDateStr,
              site,
              isNight,
              columnConfig,
              spreadsheetId,
              writeCell,
              trendRow,
              scrapingService,
              excelService,
              googleSheetsService,
              aiService,
              nightJobCategories,
              normalJobCategories,
              processFolderPath,
              csvRecords,
              { listPV: totalListPV, detailPV: totalDetailPV, webApplication: totalWebApplication, telApplication: totalTelApplication },
              currentUniqueIdColumn
            );
            
            // 行番号を更新
            if (isNight) {
              nextTrendRowNight++;
            } else {
              nextTrendRowNormal++;
            }
          }
          
          // 複数行処理が完了したので、次の企業に進む
          row++;
          processedCount++;
          continue;
        }

        // 単一行またはExcelファイルの場合の既存処理
        // ⑦ プラン、PV数・応募数、期間を抽出して出力エクセルに転記
        // 現在の行番号を取得（ナイト案件と通常案件で別々に管理）
        const trendRow = isNight ? nextTrendRowNight : nextTrendRowNormal;
        
        // デバッグ: 書き込み先の確認
        console.log(`  [DEBUG] 書き込み先: ${isNight ? 'ナイト案件' : '通常案件'}, 行番号: ${trendRow}`);

        // ナイト案件と通常案件で異なる列設定を使用
        const columnConfig = isNight 
          ? config.excelColumns.trendDatabaseNight 
          : config.excelColumns.trendDatabase;

        // スプレッドシートへの書き込み用ヘルパー関数
        const spreadsheetId = isNight ? config.googleSheets.spreadsheetIdNight : config.googleSheets.spreadsheetIdNormal;
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
          const csvCols = config.excelColumns.downloadFile.csvColumns;
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
        }

        if (appStartDate && appEndDate) {
          // CSVの場合は文字列、Excelの場合はDateオブジェクト、数値、または文字列の可能性がある
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

        // ⑩ 仕事Noを取得
        let jobNo = '';
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
          console.log('  仕事Noが取得できませんでした。スキップします。');
          row++;
          skippedCount++;
          continue;
        }

        // ⑩⑪ 原稿検索とプレビューを開く（PDF手順に基づく）
        // ⑩：仕事Noで検索
        // ⑪：一番上のプレビューボタンを押す
        console.log(`  仕事No: ${jobNo} で原稿を検索します`);
        const previewPage = await scrapingService.searchJobByNo(String(jobNo));
        console.log('  プレビューを開きました');

        // プレビューのスクリーンショットを保存
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

        // ===== プレビューページから必要な値をすべて変数に格納 =====
        console.log('  プレビューページからデータを取得中...');
        
        // ⑫ 勤務地情報を取得（都道府県・市区町村・最寄り駅）
        const workLocation = await scrapingService.getWorkLocation(previewPage);
        
        // workLocationがnullやundefinedの場合のデフォルト値
        const safeWorkLocation = workLocation || { prefecture: '', city: '', station: '' };
        
        // 都道府県はそのまま転記する（処理を削除）
        
        // ⑬ 年月を計算（申込開始日から年と月を抽出）
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
        
        // ⑭ 地方を選択（都道府県から地方を判定）
        let selectedRegion = null;
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
        
        if (safeWorkLocation.prefecture && safeWorkLocation.prefecture.trim() !== '') {
          const prefectureName = safeWorkLocation.prefecture.trim();
          
          // 都道府県から地方を判定（完全一致を試す）
          selectedRegion = prefectureToRegion[prefectureName];
          
          // 完全一致しない場合、部分一致で判定を試みる
          if (!selectedRegion) {
            // 都道府県名の正規化（末尾の「都」「道」「府」「県」を確認）
            let normalizedPrefecture = prefectureName;
            
            // 末尾に「都」「道」「府」「県」がない場合、追加を試みる
            if (!prefectureName.match(/[都道府県]$/)) {
              // 都道府県名の候補を生成
              const candidates = [
                prefectureName + '都',
                prefectureName + '道',
                prefectureName + '府',
                prefectureName + '県'
              ];
              
              // 各候補でマッピングを確認
              for (const candidate of candidates) {
                if (prefectureToRegion[candidate]) {
                  selectedRegion = prefectureToRegion[candidate];
                  normalizedPrefecture = candidate;
                  break;
                }
              }
            }
            
            // 部分一致で判定を試みる（例: "東京" → "東京都"）
            if (!selectedRegion) {
              for (const [pref, region] of Object.entries(prefectureToRegion)) {
                if (pref.includes(prefectureName) || prefectureName.includes(pref.replace(/[都道府県]$/, ''))) {
                  selectedRegion = region;
                  normalizedPrefecture = pref;
                  break;
                }
              }
            }
          }
          
          // マッピングにない場合はAI判定を使用
          if (!selectedRegion) {
            try {
              selectedRegion = await aiService.determineRegion(prefectureName, regionOptions);
            } catch (error) {
              console.warn(`  ⚠️  地方のAI判定をスキップ: ${error.message}`);
            }
          }
          
          if (selectedRegion) {
            console.log(`  地方を選択: ${selectedRegion} (都道府県: ${prefectureName})`);
          } else {
            console.warn(`  ⚠️  都道府県「${prefectureName}」から地方を判定できませんでした`);
          }
        } else {
          console.warn(`  ⚠️  都道府県情報が取得できませんでした`);
        }

        // ⑮ 職種情報を取得
        const jobCategory = await scrapingService.getJobCategory(previewPage);
        
        // ⑯ 職種大・中・小を選択（キーワードマッチング → コサイン類似度で判定）
        let selectedLarge = null;
        let selectedMedium = null;
        let selectedSmall = null;
        
        if (jobCategory.rawText && jobCategories.length > 0) {
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
          selectedLarge = cleanJobCategoryName(jobCategory.large);
          selectedMedium = jobCategory.medium;
          selectedSmall = jobCategory.small;
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

        // ⑰ 給与情報を取得
        const salary = await scrapingService.getSalary(previewPage);
        
        // ⑱ 企業名を取得（既に取得済みのため、再取得は不要）
        // const companyName = excelService.getCellValue(
        //   inputSheet,
        //   config.excelColumns.inputSheet.companyName,
        //   row
        // );
        
        // ⑲ 店名（応募受付先名）を取得（一旦未入力でOK）
        const storeName = ''; // 店名は一旦未入力でOK
        
        console.log('  ✓ プレビューページからのデータ取得が完了しました');
        
        // プレビューページを閉じる（データ取得後、すぐに閉じる）
        await previewPage.close();
        console.log('  ✓ プレビューページを閉じました');

        // ===== 取得したデータをスプレッドシートにまとめて書き込み =====
        console.log(`  スプレッドシートにデータを書き込み中... (行番号: ${trendRow}, 案件種別: ${isNight ? 'ナイト' : '通常'})`);
        
        // 媒体列を転記（最初に書き込む、通常案件のみ）
        if (!isNight) {
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
          safeWorkLocation.prefecture || ''
        );
        console.log(`  都道府県を転記: ${safeWorkLocation.prefecture || ''} (列: ${columnConfig.prefecture})`);
        await writeCell(
          columnConfig.city,
          safeWorkLocation.city || ''
        );
        console.log(`  市区町村を転記: ${safeWorkLocation.city || ''} (列: ${columnConfig.city})`);
        await writeCell(
          columnConfig.station,
          safeWorkLocation.station || ''
        );
        console.log(`  最寄り駅を転記: ${safeWorkLocation.station || ''} (列: ${columnConfig.station})`);

        // ⑬ 年月を転記
        await writeCell(
          columnConfig.year,
          year
        );
        await writeCell(
          columnConfig.month,
          month
        );

        // ⑭ 地方を転記（必ず転記する）
        if (selectedRegion) {
          await writeCell(
            columnConfig.region,
            selectedRegion
          );
          console.log(`  地方を転記: ${selectedRegion}`);
        } else if (safeWorkLocation.prefecture && safeWorkLocation.prefecture.trim() !== '') {
          // selectedRegionがnullでも、都道府県がある場合は再度判定を試みる
          const fallbackRegion = prefectureToRegion[safeWorkLocation.prefecture];
          if (fallbackRegion) {
            await writeCell(
              columnConfig.region,
              fallbackRegion
            );
            console.log(`  地方を転記（フォールバック）: ${fallbackRegion}`);
          } else {
            console.warn(`  ⚠️  都道府県「${safeWorkLocation.prefecture}」から地方を判定できませんでした。地方列は空のままです。`);
          }
        } else {
          console.warn(`  ⚠️  都道府県情報がないため、地方列は空のままです。`);
        }

        // ⑯ 職種大・中・小は既に上記で転記済み

        // ⑰ 給与情報を転記
        // 給与金額を処理
        // 給与形態はそのまま転記する
        let salaryAmount = salary.amount || 0;
        let salaryType = salary.type || '';
        
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
        
        // N列に給与形態を転記
        await writeCell(
          columnConfig.salaryType,
          salaryType
        );
        
        await writeCell(
          columnConfig.salaryAmount,
          salaryAmount
        );

        // ⑱ 企業ID・企業名を転記
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
          storeName
        );

        // ユニークIDを転記（ナイト案件と通常案件の両方）
        const currentUniqueIdColumn = isNight ? uniqueIdColumnNight : uniqueIdColumnNormal;
        if (currentUniqueIdColumn) {
          const uniqueIdValue = `${companyId}_${companyName}_${startDateStr}_${endDateStr}`;
          await writeCell(
            currentUniqueIdColumn,
            uniqueIdValue
          );
          console.log(`  ユニークIDを転記: ${uniqueIdValue} (列: ${currentUniqueIdColumn})`);
        }

        // ⑳ 申込開始日・終了日を転記
        if (startDateObj && endDateObj) {
          // YYYY/MM/DD形式にフォーマット
          const formattedStartDate = formatDateForInput(startDateObj);
          const formattedEndDate = formatDateForInput(endDateObj);
          
          if (formattedStartDate && formattedEndDate) {
            await writeCell(
              columnConfig.applicationStartDate,
              formattedStartDate
            );
            await writeCell(
              columnConfig.applicationEndDate,
              formattedEndDate
            );
            console.log(`  申込開始日を転記: ${formattedStartDate}`);
            console.log(`  申込終了日を転記: ${formattedEndDate}`);
          } else {
            console.warn(`  ⚠️  日付のフォーマットに失敗しました（開始日: ${appStartDate}, 終了日: ${appEndDate}）`);
          }
        } else if (appStartDate && appEndDate) {
          console.warn(`  ⚠️  申込開始日・終了日の日付オブジェクトが取得できませんでした（開始日: ${appStartDate}, 終了日: ${appEndDate}）`);
        }
        
        // 行番号を更新（ナイト案件と通常案件で別々に管理）
        if (isNight) {
          nextTrendRowNight++;
        } else {
          nextTrendRowNormal++;
        }
        
        processedCount++;
        console.log(`  ✓ スプレッドシートへのデータ書き込みが完了しました（行${trendRow}）`);
        console.log(`  企業ID: ${companyId} の処理が完了しました`);

        // 次の行へ
        row++;
      } catch (error) {
        console.error(`  行 ${row} の処理でエラーが発生しました:`, error.message);
        row++;
        skippedCount++;
        // エラーが発生しても処理を続行
      }
    }

    // スプレッドシートへの書き込みは既に完了しているため、保存処理は不要
    console.log('\n✓ すべてのデータをスプレッドシートに書き込みました');
    console.log(`\n処理完了: 処理済み ${processedCount}件、スキップ ${skippedCount}件`);
    console.log('\n全ての処理が完了しました！');
  } catch (error) {
    console.error('エラーが発生しました:', error);
    throw error;
  } finally {
    // ブラウザを閉じる
    await scrapingService.closeBrowser();
    console.log('ブラウザを閉じました');
  }
}

// エラーハンドリング
main().catch(error => {
  console.error('致命的なエラー:', error);
  process.exit(1);
});
