/**
 * 各URLのセレクターを特定するためのデバッグスクリプト
 * 使用方法: node src/debug-selectors-by-url.js <url_type>
 * url_type: top, performance, jobsearch, preview
 * 
 * 例:
 *   node src/debug-selectors-by-url.js top
 *   node src/debug-selectors-by-url.js performance
 *   node src/debug-selectors-by-url.js jobsearch
 *   node src/debug-selectors-by-url.js preview
 */

import ScrapingService from './services/ScrapingService.js';
import ExcelService from './services/ExcelService.js';
import config from '../config/config.js';
import { writeFileSync } from 'fs';
import { join } from 'path';
import { excelDateToJSDate, formatDateForInput } from './utils/dateUtils.js';

// URL定義
const URLS = {
  top: 'https://agent.baitoru.com/top',
  performance: 'https://agent.baitoru.com/publication/result?mode=1',
  jobsearch: 'https://agent.baitoru.com/job?mode=1',
  preview: 'https://agent.baitoru.com/pv?k=pNWqoYTLRjR2JFODQwFz4aci%2BZoy04l6CUC%2Fcabc%2FHs%3D&d=pc&p=a&m=w'
};

// 各ページタイプに応じたセレクター候補
const SELECTOR_CANDIDATES = {
  top: {
    companyIdInput: [
      'input[name*="company"]',
      'input[name*="企業"]',
      'input[id*="company"]',
      'input[id*="企業"]',
      'input[placeholder*="企業"]',
      'input[placeholder*="company"]',
      '#company-id',
      '#companyId',
      '#company_id',
      '.company-id',
      '.companyId'
    ],
    searchButton: [
      'button[type="submit"]',
      'input[type="submit"]',
      'input[value*="検索"]',
      'button.search',
      '.search-button',
      '.btn-search',
      '[onclick*="search"]',
      '[onclick*="検索"]'
    ],
    selectButton: [
      'button.select',
      '.select-button',
      '.btn-select',
      '[onclick*="select"]',
      '[onclick*="選択"]'
    ]
  },
  performance: {
    startDateInput: [
      'input[type="date"]',
      'input[name*="start"]',
      'input[name*="開始"]',
      'input[id*="start"]',
      'input[id*="開始"]',
      '#start-date',
      '.start-date'
    ],
    endDateInput: [
      'input[type="date"]',
      'input[name*="end"]',
      'input[name*="終了"]',
      'input[id*="end"]',
      'input[id*="終了"]',
      '#end-date',
      '.end-date'
    ],
    downloadButton: [
      'button:contains("ダウンロード")',
      'button:contains("Download")',
      'input[value*="ダウンロード"]',
      'button.download',
      '.download-button',
      '.btn-download',
      '[onclick*="download"]',
      '[onclick*="ダウンロード"]'
    ]
  },
  jobsearch: {
    jobNoInput: [
      'input[name*="job"]',
      'input[name*="仕事"]',
      'input[id*="job"]',
      'input[id*="仕事"]',
      'input[placeholder*="仕事"]',
      '#job-no',
      '#jobNo',
      '#job_no',
      '.job-no'
    ],
    searchButton: [
      'button[type="submit"]',
      'input[type="submit"]',
      'input[value*="検索"]',
      'button.search',
      '.search-button',
      '.btn-search'
    ],
    previewButton: [
      'button:contains("プレビュー")',
      'button:contains("Preview")',
      'a:contains("プレビュー")',
      'a:contains("Preview")',
      'button.preview',
      '.preview-button',
      '.btn-preview'
    ]
  },
  preview: {
    workLocation: [
      '.work-location',
      '.workLocation',
      '[class*="work-location"]',
      '[class*="勤務地"]'
    ],
    jobCategory: [
      '.job-category',
      '.jobCategory',
      '[class*="job-category"]',
      '[class*="職種"]'
    ],
    salary: [
      '.salary',
      '[class*="salary"]',
      '[class*="給与"]'
    ],
    station: [
      '.station',
      '.nearest-station',
      '[class*="station"]',
      '[class*="駅"]'
    ],
    storeName: [
      '.store-name',
      '.storeName',
      '.reception-name',
      '[class*="store"]',
      '[class*="店"]',
      '[class*="応募受付先"]'
    ]
  }
};

async function debugSelectorsByUrl(urlType) {
  const scrapingService = new ScrapingService();
  const excelService = new ExcelService();
  const url = URLS[urlType];
  const candidates = SELECTOR_CANDIDATES[urlType];

  if (!url) {
    console.error(`❌ 無効なURLタイプ: ${urlType}`);
    console.error('有効なタイプ: top, performance, jobsearch, preview');
    return;
  }

  try {
    console.log(`=== ${urlType}画面のセレクター特定デバッグ ===\n`);
    console.log(`URL: ${url}\n`);

    // 入力ファイルから1行目のデータを読み込む（performance, jobsearch, previewの場合）
    let companyId = null;
    let jobNo = null;
    let startDate = null;
    let endDate = null;

    if (urlType !== 'top') {
      console.log('0. 入力ファイルから1行目のデータを読み込み中...');
      try {
        const inputWorkbook = await excelService.loadInputFile();
        const inputSheet = inputWorkbook.getWorksheet(1);
        const testRow = 2; // 2行目（1行目はヘッダー）

        companyId = excelService.getCellValue(inputSheet, config.excelColumns.inputSheet.companyId, testRow);
        const startDateValue = excelService.getCellValue(inputSheet, config.excelColumns.inputSheet.startDate, testRow);
        const endDateValue = excelService.getCellValue(inputSheet, config.excelColumns.inputSheet.endDate, testRow);

        if (companyId) {
          console.log(`   企業ID: ${companyId}`);
          if (startDateValue && endDateValue) {
            startDate = excelDateToJSDate(startDateValue);
            endDate = excelDateToJSDate(endDateValue);
            console.log(`   開始日: ${startDate ? formatDateForInput(startDate) : 'N/A'}`);
            console.log(`   終了日: ${endDate ? formatDateForInput(endDate) : 'N/A'}`);
          }
        } else {
          console.warn('⚠️  企業IDが取得できませんでした。直接URLで遷移します。');
        }
      } catch (error) {
        console.warn(`⚠️  入力ファイルの読み込みに失敗しました: ${error.message}`);
        console.warn('   直接URLで遷移します。');
      }
      console.log('');
    }

    // ブラウザを起動
    console.log('1. ブラウザを起動中...');
    await scrapingService.launchBrowser();
    console.log('✓ ブラウザを起動しました\n');

    // ログイン（すべてのページで必要）
    console.log('2. ログイン中...');
    await scrapingService.login();
    console.log('✓ ログインしました\n');

    // TOPページに移動
    console.log('3. TOPページに移動中...');
    await scrapingService.goToTop();
    console.log('✓ TOPページに移動しました\n');

    // データを使って画面遷移（performance, jobsearch, previewの場合）
    if (urlType !== 'top' && companyId) {
      console.log(`4. 企業ID: ${companyId} で検索中...`);
      try {
        await scrapingService.searchByCompanyId(String(companyId));
        console.log('✓ 検索完了\n');

        console.log('5. 選択ボタンをクリック中...');
        await scrapingService.clickSelectButton();
        console.log('✓ 選択ボタンをクリックしました\n');

        // 掲載実績ページの場合
        if (urlType === 'performance') {
          console.log('6. 掲載実績ページに移動中...');
          await scrapingService.goToPerformancePage();
          console.log('✓ 掲載実績ページに移動しました\n');
        }
        // 原稿検索ページの場合
        else if (urlType === 'jobsearch') {
          console.log('6. 原稿検索ページに直接遷移中...');
          // 直接URLで遷移（セレクター取得のため）
          await scrapingService.page.goto(url, {
            waitUntil: 'networkidle2',
            timeout: 60000
          });
          // 安定して開くように1秒待機
          await new Promise(resolve => setTimeout(resolve, 1000));
          console.log('✓ 原稿検索ページに直接遷移しました\n');
        }
        // プレビュー画面の場合
        else if (urlType === 'preview') {
          // 仕事Noを取得するために、まず掲載実績をダウンロード
          if (startDate && endDate) {
            console.log('6. 掲載実績をダウンロード中（仕事No取得のため）...');
            try {
              const downloadPath = await scrapingService.downloadPerformance(
                formatDateForInput(startDate),
                formatDateForInput(endDate),
                String(companyId)
              );
              console.log(`✓ 掲載実績をダウンロードしました: ${downloadPath}\n`);
              
              // CSVファイルを読み込んで仕事Noを取得
              console.log('7. CSVファイルから仕事Noを取得中...');
              const csvRecords = await excelService.loadCSVFile(downloadPath);
              if (csvRecords && csvRecords.length > 0) {
                // CSVファイルの列名を確認（デバッグ用）
                const firstRecord = csvRecords[0];
                const columnNames = Object.keys(firstRecord);
                console.log(`   CSVファイルの列名: ${columnNames.join(', ')}`);
                
                // 仕事Noの列名を探す（複数の可能性を試す）
                const jobNoColumnNames = ['仕事No', '仕事番号', 'jobNo', 'job_no', '仕事Ｎｏ', 'C'];
                for (const colName of jobNoColumnNames) {
                  if (firstRecord[colName]) {
                    jobNo = firstRecord[colName];
                    console.log(`✓ 仕事Noを取得しました: ${jobNo} (列名: ${colName})\n`);
                    break;
                  }
                }
                
                if (!jobNo) {
                  // 列名が見つからない場合、最初のレコードの3番目の列（C列相当）を試す
                  const values = Object.values(firstRecord);
                  if (values.length >= 3) {
                    jobNo = values[2]; // C列は3番目（0-indexedで2）
                    console.log(`✓ 仕事Noを取得しました（推測）: ${jobNo}\n`);
                  } else {
                    console.warn('⚠️  仕事Noが取得できませんでした。\n');
                  }
                }
              } else {
                console.warn('⚠️  CSVファイルにレコードがありません。\n');
              }
            } catch (error) {
              console.warn(`⚠️  掲載実績のダウンロードに失敗しました: ${error.message}`);
              console.warn('   仕事Noなしで原稿検索ページに遷移します。\n');
            }
          }
          
          console.log('8. 原稿検索ページに移動中...');
          await scrapingService.goToJobSearchPage();
          console.log('✓ 原稿検索ページに移動しました\n');
          
          // もし仕事Noが取得できていれば、検索してプレビューを開く
          if (jobNo) {
            console.log(`9. 仕事No: ${jobNo} で検索中...`);
            await scrapingService.searchJobByNo(String(jobNo));
            console.log('✓ 検索完了\n');
            
            console.log('10. プレビューボタンをクリック中...');
            const previewPage = await scrapingService.clickPreviewButton();
            console.log('✓ プレビューページを開きました\n');
            
            // プレビューページに切り替え
            scrapingService.page = previewPage;
          } else {
            console.log('9. 仕事Noが取得できていないため、プレビューを開けません。\n');
            console.log('   直接URLで遷移します。\n');
            await scrapingService.page.goto(url, {
              waitUntil: 'networkidle2',
              timeout: 60000
            });
          }
        }
      } catch (error) {
        console.warn(`⚠️  データを使った遷移に失敗しました: ${error.message}`);
        console.warn('   直接URLで遷移します。\n');
        await scrapingService.page.goto(url, {
          waitUntil: 'networkidle2',
          timeout: 60000
        });
        console.log(`✓ ${urlType}画面に直接遷移しました\n`);
      }
    } else if (urlType === 'top') {
      // TOPページの場合はそのまま
    } else {
      // データが取得できなかった場合、直接URLで遷移
      console.log(`4. ${urlType}画面に直接遷移中...`);
      await scrapingService.page.goto(url, {
        waitUntil: 'networkidle2',
        timeout: 60000
      });
      console.log(`✓ ${urlType}画面に直接遷移しました\n`);
    }

    // 少し待機（ページが完全に読み込まれるまで）
    await new Promise(resolve => setTimeout(resolve, 2000));

    // 現在のURLを確認
    const currentUrl = scrapingService.page.url();
    console.log(`現在のURL: ${currentUrl}`);
    console.log(`期待されるURL: ${url}`);
    
    // URLが一致しているか確認
    if (urlType === 'top') {
      if (!currentUrl.includes('/top')) {
        console.warn('⚠️  現在のURLがTOPページではありません。手動でTOPページに遷移してください。\n');
      } else {
        console.log('✓ TOPページに正しく遷移しました\n');
      }
    } else {
      const expectedBaseUrl = url.split('?')[0];
      if (!currentUrl.includes(expectedBaseUrl)) {
        console.warn(`⚠️  現在のURLが期待されるURLと一致しません。\n`);
      } else {
        console.log('✓ 正しいページに遷移しました\n');
      }
    }

    // ページのHTMLを取得
    console.log('4. ページのHTMLを取得中...');
    const html = await scrapingService.page.content();
    const htmlPath = join(process.cwd(), `debug-${urlType}-page.html`);
    writeFileSync(htmlPath, html, 'utf-8');
    console.log(`✓ HTMLを保存しました: ${htmlPath}\n`);

    // スクリーンショットを取得
    const screenshotStepNumber = urlType === 'top' ? '5' : (urlType === 'performance' ? '8' : (urlType === 'jobsearch' ? '9' : '10'));
    console.log(`${screenshotStepNumber}. スクリーンショットを取得中...`);
    const screenshotPath = join(process.cwd(), `debug-${urlType}-page.png`);
    await scrapingService.page.screenshot({ path: screenshotPath, fullPage: true });
    console.log(`✓ スクリーンショットを保存しました: ${screenshotPath}\n`);

    // セレクター候補を検索
    const selectorStepNumber = urlType === 'top' ? '6' : (urlType === 'performance' ? '9' : (urlType === 'jobsearch' ? '10' : '11'));
    console.log(`${selectorStepNumber}. セレクター候補を検索中...\n`);

    for (const [selectorName, selectorList] of Object.entries(candidates)) {
      console.log(`--- ${selectorName}の候補 ---`);
      
      for (const selector of selectorList) {
        try {
          let elements = [];
          
          // :contains()はXPathを使用
          if (selector.includes(':contains(')) {
            const text = selector.match(/:contains\("([^"]+)"\)/)?.[1];
            if (text) {
              const tag = selector.split(':')[0];
              const xpath = `//${tag}[contains(text(), "${text}")] | //input[@type="submit" and contains(@value, "${text}")] | //a[contains(text(), "${text}")]`;
              elements = await scrapingService.page.$x(xpath);
            }
          } else {
            elements = await scrapingService.page.$$(selector);
          }

          if (elements.length > 0) {
            const firstElement = elements[0];
            const elementInfo = await scrapingService.page.evaluate((el) => {
              return {
                tag: el.tagName,
                id: el.id || '',
                name: el.name || '',
                placeholder: el.placeholder || '',
                className: el.className || '',
                value: el.value || '',
                text: el.textContent || ''
              };
            }, firstElement);
            
            console.log(`✓ 見つかりました: ${selector}`);
            console.log(`  タグ: ${elementInfo.tag}`);
            if (elementInfo.id) console.log(`  ID: #${elementInfo.id}`);
            if (elementInfo.name) console.log(`  name: [name="${elementInfo.name}"]`);
            if (elementInfo.placeholder) console.log(`  placeholder: "${elementInfo.placeholder}"`);
            if (elementInfo.value) console.log(`  value: "${elementInfo.value}"`);
            if (elementInfo.text) console.log(`  text: "${elementInfo.text.trim()}"`);
            if (elementInfo.className) {
              const classes = elementInfo.className.split(' ').filter(c => c);
              if (classes.length > 0) {
                console.log(`  class: .${classes.join('.')}`);
              }
            }
            console.log('');
          }
        } catch (e) {
          // セレクターが無効な場合はスキップ
        }
      }
    }

    // すべてのinput要素を取得して詳細を表示
    console.log('\n--- すべてのinput要素の詳細 ---');
    try {
      const allInputs = await scrapingService.page.evaluate(() => {
        const inputs = Array.from(document.querySelectorAll('input'));
        return inputs.map(input => ({
          tag: input.tagName,
          type: input.type,
          id: input.id || '',
          name: input.name || '',
          placeholder: input.placeholder || '',
          className: input.className || '',
          value: input.value || ''
        }));
      });

      allInputs.forEach((input, index) => {
        console.log(`\n入力フィールド ${index + 1}:`);
        console.log(`  タグ: ${input.tag}`);
        console.log(`  タイプ: ${input.type}`);
        if (input.id) console.log(`  ID: #${input.id}`);
        if (input.name) console.log(`  name: [name="${input.name}"]`);
        if (input.placeholder) console.log(`  placeholder: "${input.placeholder}"`);
        if (input.className) {
          const classes = input.className.split(' ').filter(c => c);
          if (classes.length > 0) {
            console.log(`  class: .${classes.join('.')}`);
          }
        }
      });
    } catch (e) {
      console.log('input要素の取得に失敗しました:', e.message);
    }

    // すべてのbutton要素を取得して詳細を表示
    console.log('\n--- すべてのbutton要素の詳細 ---');
    try {
      const allButtons = await scrapingService.page.evaluate(() => {
        const buttons = Array.from(document.querySelectorAll('button, input[type="submit"], input[type="button"], a[role="button"]'));
        return buttons.map(button => ({
          tag: button.tagName,
          type: button.type || '',
          id: button.id || '',
          name: button.name || '',
          className: button.className || '',
          text: (button.textContent || button.value || '').trim(),
          href: button.href || ''
        }));
      });

      allButtons.forEach((button, index) => {
        console.log(`\nボタン ${index + 1}:`);
        console.log(`  タグ: ${button.tag}`);
        if (button.type) console.log(`  タイプ: ${button.type}`);
        if (button.id) console.log(`  ID: #${button.id}`);
        if (button.name) console.log(`  name: [name="${button.name}"]`);
        if (button.text) console.log(`  テキスト: "${button.text}"`);
        if (button.href) console.log(`  href: "${button.href}"`);
        if (button.className) {
          const classes = button.className.split(' ').filter(c => c);
          if (classes.length > 0) {
            console.log(`  class: .${classes.join('.')}`);
          }
        }
      });
    } catch (e) {
      console.log('button要素の取得に失敗しました:', e.message);
    }

    console.log('\n=== デバッグ完了 ===');
    console.log('\n次のステップ:');
    console.log(`1. debug-${urlType}-page.html と debug-${urlType}-page.png を確認してください`);
    console.log(`2. 上記の情報を元に、config/selectors.json を更新してください`);
    console.log('3. ブラウザを閉じるまで待機します（30秒）...\n');

    // 30秒待機（ユーザーがブラウザを確認できるように）
    await new Promise(resolve => setTimeout(resolve, 30000));

  } catch (error) {
    console.error('❌ エラーが発生しました:', error);
  } finally {
    await scrapingService.closeBrowser();
    console.log('ブラウザを閉じました');
  }
}

// コマンドライン引数からURLタイプを取得
const urlType = process.argv[2];

if (!urlType) {
  console.error('❌ URLタイプが指定されていません');
  console.error('\n使用方法:');
  console.error('  node src/debug-selectors-by-url.js <url_type>');
  console.error('\n有効なURLタイプ:');
  console.error('  top        - TOP画面');
  console.error('  performance - 掲載実績画面');
  console.error('  jobsearch  - 原稿検索画面');
  console.error('  preview    - プレビュー画面');
  console.error('\n例:');
  console.error('  node src/debug-selectors-by-url.js top');
  process.exit(1);
}

debugSelectorsByUrl(urlType);
