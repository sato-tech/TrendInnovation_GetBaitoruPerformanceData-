import puppeteer from 'puppeteer';
import config from '../../config/config.js';
import { promises as fs } from 'fs';
import { join } from 'path';

/**
 * スクレイピング操作を担当するサービスクラス
 */
class ScrapingService {
  constructor() {
    this.browser = null;
    this.page = null;
    this.downloadFolder = null; // ダウンロードフォルダのパス
  }

  /**
   * ダウンロードフォルダを設定する
   * @param {string} folderPath - ダウンロードフォルダのパス
   */
  setDownloadFolder(folderPath) {
    this.downloadFolder = folderPath;
  }

  /**
   * ブラウザを起動する
   * @returns {Promise<void>}
   */
  async launchBrowser() {
    // 基本起動オプション
    const launchOptions = {
      headless: config.puppeteer.headless,
      args: [
        ...config.puppeteer.args,
        '--disable-blink-features=AutomationControlled',
        '--disable-features=IsolateOrigins,site-per-process',
        '--disable-web-security',
        '--disable-features=VizDisplayCompositor'
      ],
      timeout: config.puppeteer.timeout,
      ignoreHTTPSErrors: true
    };

    // macOSでの追加設定（--single-processは削除、問題を引き起こす可能性があるため）
    if (process.platform === 'darwin') {
      // macOSでは--single-processを削除し、代わりに他のオプションを使用
      launchOptions.args.push(
        '--disable-background-timer-throttling',
        '--disable-backgrounding-occluded-windows',
        '--disable-renderer-backgrounding',
        '--disable-ipc-flooding-protection'
      );
    }

    // 環境変数またはconfigでブラウザのパスが指定されている場合は使用
    if (config.puppeteer.executablePath) {
      launchOptions.executablePath = config.puppeteer.executablePath;
      console.log(`指定されたブラウザパスを使用: ${launchOptions.executablePath}`);
    }

    try {
      this.browser = await puppeteer.launch(launchOptions);
      this.page = await this.browser.newPage();
      await this.page.setDefaultTimeout(config.puppeteer.pageTimeout);
      console.log('✓ ブラウザを起動しました');
    } catch (error) {
      console.error('\n❌ ブラウザ起動エラー:', error.message);
      console.error('\n📋 対策:');
      console.error('1. Puppeteerのブラウザを再インストール:');
      console.error('   npm run install-browser');
      console.error('   または');
      console.error('   npx puppeteer browsers install chrome');
      console.error('\n2. macOSの場合、セキュリティ設定を確認してください:');
      console.error('   - システム環境設定 > セキュリティとプライバシー');
      console.error('   - Chrome/Chromiumの実行を許可');
      console.error('\n3. 手動でChrome/Chromiumのパスを指定する場合:');
      console.error('   .envファイルに以下を追加:');
      console.error('   BROWSER_EXECUTABLE_PATH=/Applications/Google Chrome.app/Contents/MacOS/Google Chrome');
      console.error('   または');
      console.error('   BROWSER_EXECUTABLE_PATH=/Applications/Chromium.app/Contents/MacOS/Chromium');
      console.error('\n4. ヘッドレスモードを無効にして試す場合:');
      console.error('   .envファイルに以下を追加:');
      console.error('   HEADLESS=false');
      
      throw error;
    }
  }

  /**
   * ブラウザを閉じる
   * @returns {Promise<void>}
   */
  async closeBrowser() {
    if (this.browser) {
      await this.browser.close();
      this.browser = null;
      this.page = null;
    }
  }

  /**
   * 現在のページオブジェクトを取得する
   * @returns {Page|null}
   */
  getPage() {
    return this.page;
  }

  /**
   * バイトル企業データにログインする
   * @returns {Promise<void>}
   */
  async login() {
    console.log(`ログイン画面に遷移中: ${config.baitoru.loginUrl}`);
    await this.page.goto(config.baitoru.loginUrl, {
      waitUntil: 'networkidle2',
      timeout: 60000
    });

    // 少し待機（ページが完全に読み込まれるまで）
    await new Promise(resolve => setTimeout(resolve, 1000));

    // ユーザー名入力
    console.log('ユーザー名入力フィールドを待機中...');
    await this.page.waitForSelector(config.selectors.login.usernameInput, {
      visible: true,
      timeout: 30000
    });
    
    // フィールドをクリアしてから入力
    await this.page.click(config.selectors.login.usernameInput, { clickCount: 3 });
    await this.page.keyboard.press('Backspace');
    await this.page.type(
      config.selectors.login.usernameInput,
      config.baitoru.username,
      { delay: 50 } // 入力速度を遅くして確実に入力
    );
    console.log('✓ ユーザー名を入力しました');

    // パスワード入力
    console.log('パスワード入力フィールドを待機中...');
    await this.page.waitForSelector(config.selectors.login.passwordInput, {
      visible: true,
      timeout: 30000
    });
    
    // フィールドをクリアしてから入力
    await this.page.click(config.selectors.login.passwordInput, { clickCount: 3 });
    await this.page.keyboard.press('Backspace');
    await this.page.type(
      config.selectors.login.passwordInput,
      config.baitoru.password,
      { delay: 50 } // 入力速度を遅くして確実に入力
    );
    console.log('✓ パスワードを入力しました');

    // ログインボタンクリック
    console.log('ログインボタンをクリック中...');
    await this.page.waitForSelector(config.selectors.login.loginButton, {
      visible: true,
      timeout: 30000
    });
    await this.page.click(config.selectors.login.loginButton);
    
    // ボタンクリック後、1秒待機
    await new Promise(resolve => setTimeout(resolve, 1000));
    
    // ログイン後のページ遷移を待機
    console.log('ログイン後のページ遷移を待機中...');
    await this.page.waitForNavigation({ 
      waitUntil: 'networkidle2',
      timeout: 60000
    });
    console.log('✓ ログインが完了しました');
  }

  /**
   * TOPページに移動する
   * ログイン後、既にTOPページにいる場合はスキップ
   * @returns {Promise<void>}
   */
  async goToTop() {
    // 現在のURLを確認
    const currentUrl = this.page.url();
    console.log(`現在のURL: ${currentUrl}`);
    
    // ログインURLと比較して、既にTOPページにいるかチェック
    // ログインURLが /top で終わる場合、ログイン後も同じURLの可能性がある
    if (currentUrl.includes('/top') || currentUrl.endsWith('/top')) {
      console.log('既にTOPページにいます。移動をスキップします。');
      return;
    }
    
    // TOPページへのリンクが存在するか確認
    try {
      // セレクターが存在するか、最大5秒待機
      const topLinkExists = await this.page.$(config.selectors.login.topPageButton).catch(() => null);
      
      if (topLinkExists) {
        console.log('TOPページへのリンクが見つかりました。クリックします...');
        await this.page.click(config.selectors.login.topPageButton);
        await this.page.waitForNavigation({ 
          waitUntil: 'networkidle2',
          timeout: 30000
        });
        console.log('✓ TOPページに移動しました');
      } else {
        // リンクが見つからない場合は、直接URLで遷移
        console.log('TOPページへのリンクが見つかりませんでした。直接URLで遷移します。');
        const topUrl = config.baitoru.loginUrl.includes('/top') 
          ? config.baitoru.loginUrl 
          : config.baitoru.loginUrl.replace(/\/$/, '') + '/top';
        
        await this.page.goto(topUrl, {
          waitUntil: 'networkidle2',
          timeout: 60000
        });
        console.log('✓ TOPページに直接遷移しました');
      }
    } catch (error) {
      // エラーが発生した場合も、直接URLで遷移を試みる
      console.log(`TOPページへの移動でエラーが発生しました: ${error.message}`);
      console.log('直接URLで遷移を試みます...');
      try {
        const topUrl = config.baitoru.loginUrl.includes('/top') 
          ? config.baitoru.loginUrl 
          : config.baitoru.loginUrl.replace(/\/$/, '') + '/top';
        
        await this.page.goto(topUrl, {
          waitUntil: 'networkidle2',
          timeout: 60000
        });
        console.log('✓ TOPページに直接遷移しました');
      } catch (directError) {
        console.log(`直接遷移も失敗しました: ${directError.message}`);
        throw directError;
      }
    }
  }

  /**
   * 企業IDで検索する
   * @param {string} companyId - 企業ID
   * @returns {Promise<void>}
   */
  async searchByCompanyId(companyId) {
    console.log(`企業IDで検索中: ${companyId}`);
    
    // 企業ID入力フィールドを待機
    await this.page.waitForSelector(config.selectors.search.companyIdInput, {
      visible: true,
      timeout: 30000
    });
    
    // フィールドをクリアしてから入力
    await this.page.click(config.selectors.search.companyIdInput, { clickCount: 3 });
    await this.page.keyboard.press('Backspace');
    await this.page.type(
      config.selectors.search.companyIdInput,
      companyId,
      { delay: 50 }
    );
    console.log('✓ 企業IDを入力しました');

    // 検索ボタンクリック
    await this.page.waitForSelector(config.selectors.search.searchButton, {
      visible: true,
      timeout: 30000
    });
    await this.page.click(config.selectors.search.searchButton);
    console.log('✓ 検索ボタンをクリックしました');
    
    // ボタンクリック後、1秒待機
    await new Promise(resolve => setTimeout(resolve, 1000));

    // 検索結果を待機（選択ボタンが表示されるまで）
    await this.page.waitForSelector(config.selectors.search.selectButton, {
      visible: true,
      timeout: 30000
    });
    console.log('✓ 検索結果が表示されました');
  }

  /**
   * 選択ボタンをクリックする
   * @returns {Promise<void>}
   */
  async clickSelectButton() {
    console.log('選択ボタンをクリック中...');
    
    // 現在のURLを記録
    const currentUrl = this.page.url();
    
    // 選択ボタンを待機（最初の選択ボタンをクリック）
    await this.page.waitForSelector(config.selectors.search.selectButton, {
      visible: true,
      timeout: 30000
    });
    
    // ページ遷移を待機するPromiseを作成（先に作成する必要がある）
    const navigationPromise = this.page.waitForNavigation({ 
      waitUntil: 'networkidle2',
      timeout: 60000
    }).catch(() => {
      // タイムアウトが発生しても続行
      console.warn('ページ遷移のタイムアウトが発生しましたが、続行します。');
      return null;
    });
    
    // 最初の選択ボタンをクリック（検索結果の最初の行）
    await this.page.click(config.selectors.search.selectButton);
    console.log('✓ 選択ボタンをクリックしました');
    
    // ボタンクリック後、1秒待機
    await new Promise(resolve => setTimeout(resolve, 1000));
    
    // ページ遷移を待機（複数の方法を試す）
    try {
      await navigationPromise;
      console.log('✓ ページ遷移が完了しました');
    } catch (error) {
      // ページ遷移が発生しない場合、URLの変化を確認
      await new Promise(resolve => setTimeout(resolve, 2000));
      const newUrl = this.page.url();
      
      if (newUrl !== currentUrl) {
        console.log('✓ URLが変化しました（ページ遷移が発生しました）');
        // ページが読み込まれるまで少し待機
        await new Promise(resolve => setTimeout(resolve, 1000));
      } else {
        console.warn('⚠️  URLが変化しませんでした。同じページに留まっている可能性があります。');
        console.warn('   続行しますが、次の処理でエラーが発生する可能性があります。');
      }
    }
    
    // エラーページに遷移していないか確認
    const finalUrl = this.page.url();
    if (finalUrl.includes('/error')) {
      throw new Error('選択ボタンをクリック後、エラーページに遷移しました。企業情報が正しく選択されていない可能性があります。');
    }
    
    // 企業照会画面に遷移したか確認（URLに企業IDが含まれているか）
    console.log(`遷移後のURL: ${finalUrl}`);
  }

  /**
   * ハンバーガーメニューをクリックしてサイドバーを開く
   * @returns {Promise<boolean>} クリックに成功したかどうか
   */
  async toggleSidebarMenu() {
    try {
      console.log('ハンバーガーメニューを探しています...');
      
      // ハンバーガーメニューのセレクター（複数のパターンを試す）
      const hamburgerSelectors = [
        '.sidebar-toggle',
        'a.sidebar-toggle',
        'button.sidebar-toggle',
        '.navbar-toggle',
        'button.navbar-toggle',
        '[data-toggle="offcanvas"]',
        '[data-toggle="collapse"]'
      ];
      
      for (const selector of hamburgerSelectors) {
        try {
          const hamburgerMenu = await this.page.$(selector);
          if (hamburgerMenu) {
            // 要素が表示されているか確認
            const isVisible = await this.page.evaluate(el => {
              const style = window.getComputedStyle(el);
              return style.display !== 'none' && style.visibility !== 'hidden' && el.offsetParent !== null;
            }, hamburgerMenu);
            
            if (isVisible) {
              console.log(`ハンバーガーメニューをクリックします（セレクター: ${selector}）...`);
              await hamburgerMenu.click();
              // メニューが開くまで待機
              await new Promise(resolve => setTimeout(resolve, 1000));
              console.log('✓ ハンバーガーメニューをクリックしました');
              return true;
            }
          }
        } catch (e) {
          // このセレクターで見つからなかった場合、次のセレクターを試す
          continue;
        }
      }
      
      console.warn('⚠️  ハンバーガーメニューが見つかりませんでした。');
      return false;
    } catch (e) {
      console.warn('ハンバーガーメニューのクリックに失敗しました:', e.message);
      return false;
    }
  }

  /**
   * メニュー要素が見つからない場合、ハンバーガーメニューをクリックしてから再試行する
   * @param {Function} findMenuFunction - メニューを探す関数（Promiseを返す）
   * @returns {Promise<ElementHandle|null>} 見つかったメニュー要素、またはnull
   */
  async findMenuWithHamburgerToggle(findMenuFunction) {
    // まず通常の方法でメニューを探す
    let menu = await findMenuFunction();
    
    // メニューが見つからない、または表示されていない場合
    if (!menu) {
      console.log('メニューが見つかりませんでした。ハンバーガーメニューをクリックして再試行します...');
      const toggleSuccess = await this.toggleSidebarMenu();
      
      if (toggleSuccess) {
        // ハンバーガーメニューをクリックした後、再度メニューを探す
        await new Promise(resolve => setTimeout(resolve, 500)); // メニューが開くまで少し待機
        menu = await findMenuFunction();
      }
    } else {
      // メニューが見つかったが、表示されていない場合もチェック
      const isVisible = await this.page.evaluate(el => {
        const style = window.getComputedStyle(el);
        return style.display !== 'none' && style.visibility !== 'hidden' && el.offsetParent !== null;
      }, menu);
      
      if (!isVisible) {
        console.log('メニューが見つかりましたが、表示されていません。ハンバーガーメニューをクリックして再試行します...');
        const toggleSuccess = await this.toggleSidebarMenu();
        
        if (toggleSuccess) {
          await new Promise(resolve => setTimeout(resolve, 500));
          menu = await findMenuFunction();
        }
      }
    }
    
    return menu;
  }

  /**
   * 掲載実績ページに直接URLで移動する
   * @returns {Promise<void>}
   */
  async goToPerformancePage() {
    console.log('掲載実績ページに移動中...');
    
    const currentUrl = this.page.url();
    
    // 既に掲載実績ページにいる場合はスキップ
    if (currentUrl.includes('publication/result') && !currentUrl.includes('/error')) {
      console.log('既に掲載実績ページにいます。');
      return;
    }

    // エラーページにいる場合は、TOPページに戻ってから再試行
    if (currentUrl.includes('/error')) {
      console.warn('⚠️  エラーページにいます。TOPページに戻ります...');
      await this.goToTop();
      await new Promise(resolve => setTimeout(resolve, 1000));
    }

    // 直接URLで遷移
    try {
      const targetUrl = 'https://agent.baitoru.com/publication/result?mode=1';
      console.log(`直接URLで遷移します: ${targetUrl}`);
      
      await this.page.goto(targetUrl, {
        waitUntil: 'networkidle2',
        timeout: 60000
      });
      
      // 安定して開くように1秒待機
      await new Promise(resolve => setTimeout(resolve, 1000));
      
      // エラーページに遷移していないか確認
      const newUrl = this.page.url();
      if (newUrl.includes('/error')) {
        throw new Error('掲載実績ページにアクセスできませんでした。エラーページに遷移しました。企業が選択されているか確認してください。');
      }
      
      console.log('✓ 掲載実績ページに直接遷移しました');
    } catch (e) {
      console.error('掲載実績ページへの遷移に失敗しました:', e.message);
      throw new Error(`掲載実績ページへの遷移に失敗しました: ${e.message}`);
    }
  }

  /**
   * 原稿検索ページに直接URLで移動する
   * @returns {Promise<void>}
   */
  async goToJobSearchPage() {
    console.log('原稿検索ページに移動中...');
    
    const currentUrl = this.page.url();
    
    // 既に原稿検索ページにいる場合はスキップ
    if (currentUrl.includes('job?mode=1') || currentUrl.includes('/job')) {
      console.log('既に原稿検索ページにいます。');
      return;
    }

    // 直接URLで遷移
    try {
      const targetUrl = 'https://agent.baitoru.com/job?mode=1';
      console.log(`直接URLで遷移します: ${targetUrl}`);
      
      await this.page.goto(targetUrl, {
        waitUntil: 'networkidle2',
        timeout: 60000
      });
      
      // 安定して開くように1秒待機
      await new Promise(resolve => setTimeout(resolve, 1000));
      
      console.log('✓ 原稿検索ページに直接遷移しました');
    } catch (e) {
      console.error('原稿検索ページへの遷移に失敗しました:', e.message);
      throw new Error(`原稿検索ページへの遷移に失敗しました: ${e.message}`);
    }
  }

  /**
   * 掲載実績をダウンロードする
   * @param {string} startDate - 開始日（YYYY-MM-DD形式）
   * @param {string} endDate - 終了日（YYYY-MM-DD形式）
   * @param {string} companyId - 企業ID（ファイル名の衝突回避用）
   * @returns {Promise<{filePath: string, folderPath: string}>} ダウンロードファイルのパスとフォルダパス
   */
  async downloadPerformance(startDate, endDate, companyId = '') {
    // ダウンロードフォルダが設定されていない場合はエラー
    if (!this.downloadFolder) {
      throw new Error('ダウンロードフォルダが設定されていません。setDownloadFolder()を呼び出してください。');
    }

    const timestamp = Date.now();
    const safeCompanyId = companyId ? String(companyId).replace(/[^a-zA-Z0-9_-]/g, '_') : 'unknown';
    const processFolderPath = this.downloadFolder;

    // 掲載実績ページに移動
    await this.goToPerformancePage();

    // ページが正しく読み込まれたか確認（エラーページでないか）
    const currentUrl = this.page.url();
    if (currentUrl.includes('/error')) {
      throw new Error('掲載実績ページにアクセスできませんでした。エラーページに遷移しました。');
    }

    // 開始日入力フィールドを待機
    console.log('開始日入力フィールドを待機中...');
    await this.page.waitForSelector(config.selectors.performance.startDateInput, {
      visible: true,
      timeout: 30000
    });
    console.log('✓ 開始日入力フィールドが見つかりました');

    // 開始日を入力（フィールドをクリアしてから入力）
    console.log(`開始日を入力します: ${startDate}`);
    await this.page.click(config.selectors.performance.startDateInput, { clickCount: 3 });
    await this.page.keyboard.press('Backspace');
    await this.page.keyboard.press('Backspace'); // 念のため2回
    await this.page.type(config.selectors.performance.startDateInput, startDate, { delay: 50 });
    console.log(`✓ 開始日を入力しました: ${startDate}`);

    // 終了日入力フィールドを待機
    console.log('終了日入力フィールドを待機中...');
    await this.page.waitForSelector(config.selectors.performance.endDateInput, {
      visible: true,
      timeout: 30000
    });
    console.log('✓ 終了日入力フィールドが見つかりました');

    // 終了日を入力（フィールドをクリアしてから入力）
    console.log(`終了日を入力します: ${endDate}`);
    await this.page.click(config.selectors.performance.endDateInput, { clickCount: 3 });
    await this.page.keyboard.press('Backspace');
    await this.page.keyboard.press('Backspace'); // 念のため2回
    await this.page.type(config.selectors.performance.endDateInput, endDate, { delay: 50 });
    console.log(`✓ 終了日を入力しました: ${endDate}`);

    // csvフォルダを作成
    const csvFolderPath = join(processFolderPath, 'csv');
    try {
      await fs.access(csvFolderPath);
    } catch {
      // フォルダが存在しない場合は作成
      await fs.mkdir(csvFolderPath, { recursive: true });
      console.log(`✓ CSVフォルダを作成しました: ${csvFolderPath}`);
    }

    // ダウンロードボタンクリック
    // ダウンロード前のファイル一覧を取得（処理フォルダ内）
    const filesBefore = await fs.readdir(processFolderPath);

    const client = await this.page.target().createCDPSession();
    await client.send('Page.setDownloadBehavior', {
      behavior: 'allow',
      downloadPath: processFolderPath
    });

    // ダウンロードボタンを待機
    console.log('ダウンロードボタンを待機中...');
    await this.page.waitForSelector(config.selectors.performance.downloadButton, {
      visible: true,
      timeout: 30000
    });
    console.log('✓ ダウンロードボタンが見つかりました');

    // ダウンロードボタンをクリック
    await this.page.click(config.selectors.performance.downloadButton);
    console.log('✓ ダウンロードボタンをクリックしました');
    
    // ボタンクリック後、1秒待機
    await new Promise(resolve => setTimeout(resolve, 1000));

    // ダウンロード完了を待つ（最大60秒）
    const maxWaitTime = 60000;
    const checkInterval = 500;
    const fileStableTime = 2000; // ファイルサイズが安定するまでの時間（2秒）
    let waitedTime = 0;
    let downloadedFile = null;
    let filePath = null;
    let lastFileSize = 0;
    let stableCount = 0;
    let fileFound = false;

    console.log('ダウンロード完了を待機中...');
    
    while (waitedTime < maxWaitTime) {
      await new Promise(resolve => setTimeout(resolve, checkInterval));
      waitedTime += checkInterval;

      try {
        const filesAfter = await fs.readdir(processFolderPath);
        const newFiles = filesAfter.filter(f => !filesBefore.includes(f));
        
        // CSVまたはExcelファイルを探す
        const csvFiles = newFiles.filter(f => f.endsWith('.csv'));
        const excelFiles = newFiles.filter(f => f.endsWith('.xlsx'));

        if (csvFiles.length > 0 && !fileFound) {
          downloadedFile = csvFiles[0];
          filePath = join(processFolderPath, downloadedFile);
          fileFound = true;
          console.log(`✓ 新しいCSVファイルを検出: ${downloadedFile}`);
        } else if (excelFiles.length > 0 && !fileFound) {
          downloadedFile = excelFiles[0];
          filePath = join(processFolderPath, downloadedFile);
          fileFound = true;
          console.log(`✓ 新しいExcelファイルを検出: ${downloadedFile}`);
        }

        // ファイルが見つかった場合、ファイルサイズが安定するまで待機
        if (filePath && fileFound) {
          try {
            const stats = await fs.stat(filePath);
            const currentSize = stats.size;
            
            // ファイルサイズが前回と同じ場合、カウントを増やす
            if (currentSize === lastFileSize && currentSize > 0) {
              stableCount++;
              // ファイルサイズが2秒間安定している場合、ダウンロード完了と判断
              if (stableCount * checkInterval >= fileStableTime) {
                console.log(`✓ ファイルサイズが安定しました: ${currentSize} bytes`);
                break;
              }
            } else {
              // ファイルサイズが変化した場合、カウントをリセット
              if (currentSize !== lastFileSize) {
                console.log(`ファイルサイズが変化: ${lastFileSize} → ${currentSize} bytes`);
                stableCount = 0;
                lastFileSize = currentSize;
              }
            }
          } catch (error) {
            // ファイルがまだ完全に書き込まれていない場合、続行
            if (error.code !== 'ENOENT') {
              console.warn(`ファイル状態確認エラー: ${error.message}`);
            }
            continue;
          }
        }
      } catch (error) {
        // ディレクトリ読み取りエラーの場合、続行
        if (waitedTime % 5000 === 0) {
          console.log(`ダウンロード待機中... (${Math.floor(waitedTime / 1000)}秒)`);
        }
        continue;
      }
    }

    if (!downloadedFile || !filePath) {
      // ダウンロードディレクトリの内容を確認
      try {
        const currentFiles = await fs.readdir(processFolderPath);
        console.error('ダウンロードディレクトリの内容:', currentFiles);
        console.error('ダウンロード前のファイル:', filesBefore);
      } catch (e) {
        console.error('ダウンロードディレクトリの確認に失敗:', e.message);
      }
      throw new Error('ファイルのダウンロードが完了しませんでした（タイムアウト）');
    }

    // ファイルが存在し、読み取り可能か確認（リトライ処理付き）
    let fileReady = false;
    for (let retry = 0; retry < 10; retry++) {
      try {
        const stats = await fs.stat(filePath);
        if (stats.size > 0) {
          fileReady = true;
          console.log(`✓ ダウンロード完了: ${downloadedFile} (${stats.size} bytes)`);
          break;
        } else {
          console.warn(`ファイルが空です。リトライ ${retry + 1}/10...`);
        }
      } catch (error) {
        if (error.code === 'ENOENT') {
          console.warn(`ファイルが見つかりません。リトライ ${retry + 1}/10...`);
        } else {
          console.warn(`ファイル状態確認エラー: ${error.message}。リトライ ${retry + 1}/10...`);
        }
        if (retry < 9) {
          await new Promise(resolve => setTimeout(resolve, 500));
        }
      }
    }

    if (!fileReady) {
      throw new Error('ダウンロードしたファイルが読み取り可能になりませんでした');
    }

    // ファイル名の衝突を避けるため、企業IDとタイムスタンプを追加
    const ext = downloadedFile.endsWith('.csv') ? '.csv' : '.xlsx';
    const newFileName = `performance${safeCompanyId ? `_${safeCompanyId}` : ''}_${timestamp}${ext}`;
    // CSVファイルはcsvフォルダに移動、Excelファイルはそのまま
    const targetFolder = ext === '.csv' ? csvFolderPath : processFolderPath;
    const newPath = join(targetFolder, newFileName);

    // ファイル名を変更（リトライ処理付き）
    let renameSuccess = false;
    for (let retry = 0; retry < 5; retry++) {
      try {
        await fs.rename(filePath, newPath);
        renameSuccess = true;
        break;
      } catch (error) {
        if (retry < 4) {
          console.warn(`ファイル名変更に失敗しました（リトライ ${retry + 1}/5）: ${error.message}`);
          await new Promise(resolve => setTimeout(resolve, 500));
        } else {
          throw new Error(`ファイル名の変更に失敗しました: ${error.message}`);
        }
      }
    }

    if (!renameSuccess) {
      throw new Error('ファイル名の変更に失敗しました');
    }

    // リネーム後のファイルが存在するか確認
    try {
      const finalStats = await fs.stat(newPath);
      if (ext === '.csv') {
        console.log(`✓ CSVファイルをcsvフォルダに移動しました: ${newFileName} (${finalStats.size} bytes)`);
      } else {
        console.log(`✓ ファイルをリネームしました: ${newFileName} (${finalStats.size} bytes)`);
      }
    } catch (error) {
      throw new Error(`リネーム後のファイルが見つかりません: ${error.message}`);
    }

    // ファイルが完全に書き込まれるまで少し待機（念のため）
    await new Promise(resolve => setTimeout(resolve, 1000));

    return {
      filePath: newPath,
      folderPath: processFolderPath
    };
  }

  /**
   * プレビューページのスクリーンショットを保存する
   * @param {Page} previewPage - プレビューページ
   * @param {string} downloadFolderPath - ダウンロードフォルダパス（1行あたりに生成されるフォルダ）
   * @param {string} jobNo - 仕事No（ファイル名に使用）
   * @param {string} companyId - 企業ID（ファイル名に使用）
   * @param {string} startDate - 掲載開始日（YYYY/MM/DD形式、ファイル名に使用）
   * @param {string} endDate - 掲載終了日（YYYY/MM/DD形式、ファイル名に使用）
   * @returns {Promise<string>} スクリーンショットファイルのパス
   */
  async savePreviewScreenshot(previewPage, downloadFolderPath, jobNo = '', companyId = '', startDate = '', endDate = '') {
    try {
      // 「スクリーンショット」フォルダのパスを作成（downloadFolderPath内に作成）
      const screenshotFolderPath = join(downloadFolderPath, 'スクリーンショット');
      
      // 「スクリーンショット」フォルダが存在しない場合は作成
      try {
        await fs.access(screenshotFolderPath);
      } catch {
        // フォルダが存在しない場合は作成
        await fs.mkdir(screenshotFolderPath, { recursive: true });
        console.log(`✓ スクリーンショットフォルダを作成しました: ${screenshotFolderPath}`);
      }
      
      // ファイル名を生成: 企業ID-仕事No-掲載開始日-掲載終了日.png
      const safeCompanyId = companyId ? String(companyId).replace(/[^a-zA-Z0-9_-]/g, '_') : 'unknown';
      const safeJobNo = jobNo ? String(jobNo).replace(/[^a-zA-Z0-9_-]/g, '_') : 'unknown';
      const safeStartDate = startDate ? String(startDate).replace(/[^0-9\/]/g, '').replace(/\//g, '') : '';
      const safeEndDate = endDate ? String(endDate).replace(/[^0-9\/]/g, '').replace(/\//g, '') : '';
      
      // ファイル名のパーツを組み立て
      const fileNameParts = [safeCompanyId, safeJobNo];
      if (safeStartDate) fileNameParts.push(safeStartDate);
      if (safeEndDate) fileNameParts.push(safeEndDate);
      
      const screenshotFileName = `${fileNameParts.join('-')}.png`;
      const screenshotPath = join(screenshotFolderPath, screenshotFileName);

      // スクリーンショットを保存
      await previewPage.screenshot({
        path: screenshotPath,
        fullPage: true, // ページ全体をスクリーンショット
        type: 'png'
      });

      console.log(`✓ プレビューのスクリーンショットを保存しました: ${screenshotPath}`);
      return screenshotPath;
    } catch (error) {
      console.error(`⚠️  スクリーンショットの保存に失敗しました: ${error.message}`);
      throw error;
    }
  }

  /**
   * 仕事Noで原稿を検索し、一番上のプレビューボタンをクリックしてプレビューページを開く
   * PDF手順⑩⑪に基づく：仕事Noで検索→一番上のプレビューボタンを押す
   * @param {string} jobNo - 仕事No
   * @returns {Promise<Page>} プレビューページのPageオブジェクト
   */
  async searchJobByNo(jobNo) {
    // 原稿検索ページに移動
    await this.goToJobSearchPage();

    // 仕事No入力フィールドを待機
    await this.page.waitForSelector(config.selectors.jobSearch.jobNoInput, {
      visible: true,
      timeout: 30000
    });
    
    // フィールドをクリアしてから入力
    await this.page.click(config.selectors.jobSearch.jobNoInput, { clickCount: 3 });
    await this.page.keyboard.press('Backspace');
    await this.page.type(
      config.selectors.jobSearch.jobNoInput,
      jobNo,
      { delay: 50 }
    );
    console.log(`✓ 仕事Noを入力しました: ${jobNo}`);

    // 検索ボタンを待機してクリック
    await this.page.waitForSelector(config.selectors.jobSearch.searchButton, {
      visible: true,
      timeout: 30000
    });
    await this.page.click(config.selectors.jobSearch.searchButton);
    console.log('✓ 検索ボタンをクリックしました');
    
    // ボタンクリック後、1秒待機
    await new Promise(resolve => setTimeout(resolve, 1000));
    
    // 検索結果を待機（プレビューボタンが表示されるまで）
    await this.page.waitForSelector(config.selectors.jobSearch.firstPreviewButton, {
      visible: true,
      timeout: 30000
    });
    console.log('✓ 検索結果が表示されました');
    
    // ⑪ 一番上のプレビューボタンをクリック（PDF手順に基づく）
    console.log('一番上のプレビューボタンをクリック中...');
    
    // 新しいページが開かれるのを待機（タイムアウト付き）
    const newPagePromise = Promise.race([
      new Promise(resolve => {
        this.browser.once('targetcreated', async target => {
          const page = await target.page();
          resolve(page);
        });
      }),
      new Promise((_, reject) => {
        setTimeout(() => reject(new Error('新しいページの作成がタイムアウトしました')), 15000);
      })
    ]);
    
    // プレビューボタンをクリック
    await this.page.click(config.selectors.jobSearch.firstPreviewButton);
    console.log('✓ 一番上のプレビューボタンをクリックしました');
    
    // 新しいページを取得
    let newPage;
    try {
      newPage = await newPagePromise;
    } catch (error) {
      console.warn(`⚠️  新しいページの取得でエラー: ${error.message}`);
      // フォールバック: ブラウザの全ページから最新のページを取得
      const pages = await this.browser.pages();
      newPage = pages[pages.length - 1];
      console.log('  フォールバック: 最新のページを使用します');
    }
    
    // ページが読み込まれるまで待機（複数の方法を試行）
    try {
      // 方法1: DOMContentLoadedを待機（より高速）
      await newPage.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: 10000 }).catch(() => {
        // タイムアウトしても続行
      });
    } catch (error) {
      // ナビゲーション待機がタイムアウトしても続行
    }
    
    // 方法2: プレビューページの特定要素が表示されるまで待機
    try {
      // URLがプレビューページか確認
      const previewUrl = newPage.url();
      if (previewUrl.includes('/pv') || previewUrl.includes('preview')) {
        console.log('✓ プレビューページのURLを確認しました');
        
        // プレビューページの主要要素（iframe）が表示されるまで待機
        await newPage.waitForSelector(config.selectors.preview.jobListPreview, {
          visible: true,
          timeout: 20000
        }).catch(() => {
          // タイムアウトしても続行
          console.warn('⚠️  プレビューページのiframe要素の待機がタイムアウトしましたが、続行します');
        });
      } else {
        // URLがまだプレビューページでない場合、少し待機して再確認
        await new Promise(resolve => setTimeout(resolve, 2000));
        const currentUrl = newPage.url();
        if (currentUrl.includes('/pv') || currentUrl.includes('preview')) {
          console.log('✓ プレビューページのURLを確認しました（再確認）');
        } else {
          console.warn(`⚠️  プレビューページのURLが予期しない形式です: ${currentUrl}`);
        }
      }
    } catch (error) {
      console.warn(`⚠️  プレビューページの確認でエラー: ${error.message}`);
    }
    
    // 少し待機してから完了
    await new Promise(resolve => setTimeout(resolve, 1000));
    console.log('✓ プレビューページを開きました');
    
    return newPage;
  }

  /**
   * プレビューボタンをクリックして新しいタブで開く
   * @returns {Promise<Page>} 新しいページオブジェクト
   */
  async clickPreviewButton() {
    console.log('プレビューボタンをクリック中...');
    const [newPage] = await Promise.all([
      new Promise(resolve => {
        this.browser.once('targetcreated', target => resolve(target.page()));
      }),
      this.page.click(config.selectors.jobSearch.previewButton)
    ]);

    // ボタンクリック後、1秒待機
    await new Promise(resolve => setTimeout(resolve, 1000));

    // 新しいページが読み込まれるまで待機
    try {
      await newPage.waitForNavigation({ waitUntil: 'networkidle2', timeout: 30000 });
    } catch (e) {
      // ナビゲーションが発生しない場合（既に読み込まれている場合など）
      await new Promise(resolve => setTimeout(resolve, 2000));
    }

    console.log('✓ プレビューページが開きました（別タブ）');
    return newPage;
  }

  /**
   * プレビュータブを閉じる
   * @param {Page} previewPage - 閉じるプレビューページ
   * @returns {Promise<void>}
   */
  async closePreviewTab(previewPage) {
    try {
      if (previewPage && !previewPage.isClosed()) {
        await previewPage.close();
        console.log('✓ プレビュータブを閉じました');
      }
    } catch (error) {
      console.warn(`⚠️  プレビュータブのクローズエラー: ${error.message}`);
    }
  }

  /**
   * TOP画面に遷移し、入力フィールドをリセットする
   * @returns {Promise<void>}
   */
  async goToTopAndReset() {
    // TOP画面に確実に遷移
    await this.goToTop();
    await new Promise(resolve => setTimeout(resolve, 1000));
    
    // 現在のURLを確認して、TOPページにいることを確認
    const currentUrl = this.page.url();
    if (!currentUrl.includes('/top')) {
      // TOPページにいない場合は、再度遷移を試みる
      console.log('TOPページにいないため、再度遷移を試みます...');
      const topUrl = config.baitoru.loginUrl.includes('/top') 
        ? config.baitoru.loginUrl 
        : config.baitoru.loginUrl.replace(/\/$/, '') + '/top';
      
      await this.page.goto(topUrl, {
        waitUntil: 'networkidle2',
        timeout: 60000
      });
      await new Promise(resolve => setTimeout(resolve, 1000));
      console.log('✓ TOPページに確実に遷移しました');
    }
    
    // 入力フィールドのリセットは不要（TOPページには入力フィールドがないため）
    // 次の企業ID検索時に、searchByCompanyIdメソッド内で入力フィールドがクリアされる
  }

  /**
   * プレビューページから勤務地情報を取得する（仕事一覧プレビューのiframe内から取得）
   * @param {Page} previewPage - プレビューページ
   * @returns {Promise<{prefecture: string, city: string, station: string}>}
   */
  async getWorkLocation(previewPage) {
    try {
      // 仕事一覧プレビューまでスクロール
      await previewPage.evaluate(() => {
        const jobListPreview = document.querySelector('#list-preview-frame, iframe[data-preview-type="list"]');
        if (jobListPreview) {
          jobListPreview.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }
      });
      await previewPage.waitForTimeout(2000);

      let prefecture = '';
      let city = '';
      let station = '';

      // 仕事一覧プレビューのiframeを取得
      const iframe = await previewPage.$('#list-preview-frame');
      if (iframe) {
        const iframeContent = await iframe.contentFrame();
        if (iframeContent) {
          // iframe内で勤務地情報を取得（複数のセレクターパターンを試行）
          try {
            // 複数のセレクターパターンを定義（順番に試行）
            const workLocationSelectors = [
              'body > div > article > div > div.bg01 > div > div.pt02 > div.pt02b > ul.ul02 > li',
              'body > div > article > div > div.bg01 > div > div.pt12 > div.pt12b > dl:nth-child(3) > dd > ul > li',
              'body > div > article > div > div.bg01 > div > div.pt12 > div.pt12b > dl > dd > ul > li',
              'body > div > article > div > div.bg01 > div > div.pt02 > div.pt02b > ul > li',
              'body > div > article > div > div.bg01 > div > div.pt12 > div.pt12b > dl > dd > ul.ul02 > li',
              // より汎用的なパターン
              'body > div > article > div > div.bg01 > div > div.pt02 > div.pt02b > ul li',
              'body > div > article > div > div.bg01 > div > div.pt12 > div.pt12b > dl dd ul li',
              // テキストに「[勤務地]」や「[勤務地・面接地]」が含まれる要素を探す
              'body > div > article > div > div.bg01 > div > div li',
              'body > div > article > div > div.bg01 > div > div dl dd ul li'
            ];

            let workLocationElements = [];
            
            // 各セレクターを順番に試行
            for (const selector of workLocationSelectors) {
              try {
                workLocationElements = await iframeContent.$$(selector);
                if (workLocationElements.length > 0) {
                  console.log(`  ✓ セレクターで要素を発見: ${selector} (${workLocationElements.length}個)`);
                  break;
                }
              } catch (selectorError) {
                // セレクターが無効な場合は次のセレクターを試行
                continue;
              }
            }

            // 要素が見つからない場合、テキストに「[勤務地]」や「[勤務地・面接地]」が含まれる要素を検索
            if (workLocationElements.length === 0) {
              try {
                const allElements = await iframeContent.$$('body li, body dd');
                for (const element of allElements) {
                  const text = await iframeContent.evaluate(el => (el.textContent || el.innerText || '').trim(), element);
                  if (text.includes('[勤務地]') || text.includes('[勤務地・面接地]') || text.includes('[勤務地･面接地]')) {
                    workLocationElements.push(element);
                    console.log(`  ✓ テキストマッチで要素を発見: ${text.substring(0, 50)}`);
                  }
                }
              } catch (searchError) {
                console.log(`  テキスト検索エラー: ${searchError.message}`);
              }
            }
            
            if (workLocationElements.length > 0) {
              // すべての要素から情報を抽出
              for (const element of workLocationElements) {
                const text = await iframeContent.evaluate(el => (el.textContent || el.innerText || '').trim(), element);
                
                // テキストに「[勤務地]」や「[勤務地・面接地]」が含まれている場合のみ処理
                if (!text.includes('[勤務地]') && !text.includes('[勤務地・面接地]') && !text.includes('[勤務地･面接地]')) {
                  continue;
                }
                
                // 都道府県を抽出（「[勤務地・面接地]」を除去して47都道府県名のみ）
                if (!prefecture && (text.includes('都') || text.includes('道') || text.includes('府') || text.includes('県'))) {
                  // 「[勤務地]」「[勤務地・面接地]」「[勤務地･面接地]」を除去
                  let cleanedText = text
                    .replace(/^\[勤務地[・･]面接地\]\s*/i, '')
                    .replace(/^\[勤務地\]\s*/i, '')
                    .trim();
                  
                  // 47都道府県名のリストから正確にマッチング
                  const prefectureList = ['北海道', '青森県', '岩手県', '宮城県', '秋田県', '山形県', '福島県', '茨城県', '栃木県', '群馬県', '埼玉県', '千葉県', '東京都', '神奈川県', '新潟県', '富山県', '石川県', '福井県', '山梨県', '長野県', '岐阜県', '静岡県', '愛知県', '三重県', '滋賀県', '京都府', '大阪府', '兵庫県', '奈良県', '和歌山県', '鳥取県', '島根県', '岡山県', '広島県', '山口県', '徳島県', '香川県', '愛媛県', '高知県', '福岡県', '佐賀県', '長崎県', '熊本県', '大分県', '宮崎県', '鹿児島県', '沖縄県'];
                  
                  for (const pref of prefectureList) {
                    if (cleanedText.includes(pref)) {
                      prefecture = pref;
                      console.log(`  都道府県を抽出: ${prefecture}`);
                      break;
                    }
                  }
                  
                  // マッチしなかった場合、正規表現で抽出を試行
                  if (!prefecture) {
                    const prefectureMatch = cleanedText.match(/(北海道|青森県|岩手県|宮城県|秋田県|山形県|福島県|茨城県|栃木県|群馬県|埼玉県|千葉県|東京都|神奈川県|新潟県|富山県|石川県|福井県|山梨県|長野県|岐阜県|静岡県|愛知県|三重県|滋賀県|京都府|大阪府|兵庫県|奈良県|和歌山県|鳥取県|島根県|岡山県|広島県|山口県|徳島県|香川県|愛媛県|高知県|福岡県|佐賀県|長崎県|熊本県|大分県|宮崎県|鹿児島県|沖縄県)/);
                    if (prefectureMatch) {
                      prefecture = prefectureMatch[1];
                      console.log(`  都道府県を抽出（正規表現）: ${prefecture}`);
                    }
                  }
                }
                
                // 市区町村を抽出
                if (!city && (text.includes('区') || text.includes('市') || text.includes('町') || text.includes('村'))) {
                  // 都道府県名を除去してから市区町村を抽出
                  let cityText = text;
                  if (prefecture) {
                    cityText = cityText.replace(prefecture, '').trim();
                  }
                  // 「[勤務地]」「[勤務地・面接地]」を除去
                  cityText = cityText
                    .replace(/^\[勤務地[・･]面接地\]\s*/i, '')
                    .replace(/^\[勤務地\]\s*/i, '')
                    .trim();
                  
                  const cityMatch = cityText.match(/([^都道府県]+?[市区町村])/);
                  if (cityMatch && !cityMatch[1].includes(prefecture)) {
                    city = cityMatch[1].trim();
                    console.log(`  市区町村を抽出: ${city}`);
                  }
                }
                
                // 駅名を抽出（「(徒歩10分)」などを除去）
                if (!station && (text.includes('駅') || text.includes('線'))) {
                  // 括弧内の情報を除去してから駅名を抽出
                  let cleanedText = text.replace(/\([^\)]+\)/g, '').replace(/（[^）]+）/g, '').trim();
                  
                  // 駅名を抽出（「駅」で終わる場合）
                  const stationMatch = cleanedText.match(/([^\s\/]+駅)/);
                  if (stationMatch) {
                    station = stationMatch[1].replace('駅', '').trim();
                    console.log(`  駅名を抽出: ${station}`);
                  } else {
                    // 駅名が「駅」で終わらない場合、駅名らしい部分を抽出
                    const stationMatch2 = cleanedText.match(/([^\s\/\(（]+駅)/);
                    if (stationMatch2) {
                      station = stationMatch2[1].replace('駅', '').trim();
                      console.log(`  駅名を抽出（フォールバック）: ${station}`);
                    }
                  }
                }
                
                // すべての情報が取得できた場合はループを抜ける
                if (prefecture && city && station) {
                  break;
                }
              }
            }
          } catch (error) {
            console.log(`  iframe内の勤務地情報取得エラー: ${error.message}`);
          }
        }
      }

      // iframe内で取得できなかった場合、メインページから取得を試みる（フォールバック）
      if (!prefecture || !city || !station) {
        try {
          const workLocationElements = await previewPage.$x(config.selectors.preview.workLocationXPath);
          if (workLocationElements.length > 0) {
            const locationText = await previewPage.evaluate(el => el.textContent.trim(), workLocationElements[0]);
            console.log(`  勤務地テキスト（フォールバック）: ${locationText}`);

            const lines = locationText.split('\n').map(line => line.trim()).filter(line => line);
            
            // 都道府県と市区町村を含む行を探す
            if (!prefecture || !city) {
              for (const line of lines) {
                // 都道府県名を抽出（完全な都道府県名を取得）
                const prefectureMatch = line.match(/(北海道|青森県|岩手県|宮城県|秋田県|山形県|福島県|茨城県|栃木県|群馬県|埼玉県|千葉県|東京都|神奈川県|新潟県|富山県|石川県|福井県|山梨県|長野県|岐阜県|静岡県|愛知県|三重県|滋賀県|京都府|大阪府|兵庫県|奈良県|和歌山県|鳥取県|島根県|岡山県|広島県|山口県|徳島県|香川県|愛媛県|高知県|福岡県|佐賀県|長崎県|熊本県|大分県|宮崎県|鹿児島県|沖縄県)/);
                if (prefectureMatch) {
                  if (!prefecture) prefecture = prefectureMatch[1];
                  // 市区町村を抽出
                  if (!city) {
                    const cityMatch = line.match(/([^都道府県]+?[市区町村])/);
                    if (cityMatch) {
                      city = cityMatch[1].trim();
                    }
                  }
                  break;
                }
                // フォールバック: 正規表現で都道府県を抽出
                const fallbackMatch = line.match(/^([^都道府県]*[都道府県])\s*(.+?[市区町村])?/);
                if (fallbackMatch) {
                  if (!prefecture) prefecture = fallbackMatch[1].trim();
                  if (!city && fallbackMatch[2]) city = fallbackMatch[2].trim();
                  break;
                }
              }
            }
            
            // 駅情報を抽出
            if (!station) {
              for (const line of lines) {
                if (line.includes('駅') || line.includes('線')) {
                  const stationMatch = line.match(/([^\s\/]+駅|[^\s\/]+線\s+[^\s\/]+)/);
                  if (stationMatch) {
                    station = stationMatch[1].replace('駅', '').trim();
                    break;
                  }
                }
              }
            }
          }
        } catch (error) {
          console.log(`  フォールバック取得エラー: ${error.message}`);
        }
      }

      // 都道府県名から先頭のラベル（[勤務地・面接地]など）を除去して47都道府県名のみに
      if (prefecture) {
        // 「[勤務地・面接地]」や「[勤務地・面接地]」を除去（全角・と半角・の両方に対応）
        prefecture = prefecture
          .replace(/^\[勤務地[・･]面接地\]\s*/i, '') // [勤務地・面接地]を除去
          .replace(/^\[勤務地[・･]面接地\]\s*/i, '') // [勤務地・面接地]を除去（念のため2回）
          .replace(/^(勤務地|面接地|所在地)[:：\s]*/i, '')
          .trim();
        
        // 47都道府県名のリストに一致するか確認し、一致しない場合は空にする
        const prefectureList = ['北海道', '青森県', '岩手県', '宮城県', '秋田県', '山形県', '福島県', '茨城県', '栃木県', '群馬県', '埼玉県', '千葉県', '東京都', '神奈川県', '新潟県', '富山県', '石川県', '福井県', '山梨県', '長野県', '岐阜県', '静岡県', '愛知県', '三重県', '滋賀県', '京都府', '大阪府', '兵庫県', '奈良県', '和歌山県', '鳥取県', '島根県', '岡山県', '広島県', '山口県', '徳島県', '香川県', '愛媛県', '高知県', '福岡県', '佐賀県', '長崎県', '熊本県', '大分県', '宮崎県', '鹿児島県', '沖縄県'];
        
        if (!prefectureList.includes(prefecture)) {
          // リストに一致しない場合、テキスト内から都道府県名を探す
          const foundPrefecture = prefectureList.find(p => prefecture.includes(p));
          if (foundPrefecture) {
            prefecture = foundPrefecture;
          } else {
            // 見つからない場合は空にする
            prefecture = '';
          }
        }
      }
      
      // 駅名から括弧内の情報（(徒歩10分)など）を除去
      if (station) {
        station = station
          .replace(/\([^\)]+\)/g, '') // (徒歩10分)などを除去
          .replace(/（[^）]+）/g, '') // （徒歩10分）などを除去
          .trim();
      }
      
      console.log(`  都道府県: ${prefecture}, 市区町村: ${city}, 駅: ${station}`);
      return { prefecture, city, station };
    } catch (error) {
      console.error('勤務地情報の取得エラー:', error.message);
      return { prefecture: '', city: '', station: '' };
    }
  }

  /**
   * プレビューページから職種情報を取得する（仕事一覧プレビューのiframe内から取得）
   * @param {Page} previewPage - プレビューページ
   * @returns {Promise<{large: string, medium: string, small: string, rawText: string}>}
   */
  async getJobCategory(previewPage) {
    try {
      // 仕事一覧プレビューまでスクロール
      await previewPage.evaluate(() => {
        const jobListPreview = document.querySelector('#list-preview-frame, iframe[data-preview-type="list"]');
        if (jobListPreview) {
          jobListPreview.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }
      });
      await previewPage.waitForTimeout(2000);

      let jobTypeText = '';

      // 仕事一覧プレビューのiframeを取得
      const iframe = await previewPage.$(config.selectors.preview.jobListPreview);
      if (iframe) {
        const iframeContent = await iframe.contentFrame();
        if (iframeContent) {
          // iframe内で職種を取得（指定されたセレクターを使用）
          try {
            // セレクターで職種を取得
            const jobTypeElement = await iframeContent.$(config.selectors.preview.jobTypeInIframeSelector);
            if (jobTypeElement) {
              jobTypeText = await iframeContent.evaluate(el => {
                return (el.textContent || el.innerText || '').trim();
              }, jobTypeElement);
            }
          } catch (error) {
            console.log(`  iframe内の職種取得エラー: ${error.message}`);
          }
        }
      }

      // iframe内で取得できなかった場合、メインページから取得を試みる（フォールバック）
      if (!jobTypeText) {
        try {
          const jobTypeElements = await previewPage.$x(config.selectors.preview.jobTypeXPath);
          if (jobTypeElements.length > 0) {
            jobTypeText = await previewPage.evaluate(el => el.textContent.trim(), jobTypeElements[0]);
          }
        } catch (error) {
          console.log(`  フォールバック取得エラー: ${error.message}`);
        }
      }

      if (!jobTypeText) {
        throw new Error('職種要素が見つかりませんでした');
      }

      console.log(`  職種テキスト: ${jobTypeText}`);

      // 職種の分類ロジック
      // 例: "アルバイト・パート 建築・建設・土木作業,建築・土木その他"
      // → large: "アルバイト・パート", medium: "建築・建設・土木作業", small: "建築・土木その他"
      // 例: "[ア・パ]①②③建築・建設・土木作業、建築・土木その他"
      // → large: "[ア・パ]①②③建築・建設・土木作業", medium: "建築・土木その他", small: ""
      // カンマ、句読点、スペースで分割
      const parts = jobTypeText.split(/[,、\s]+/).map(s => s.trim()).filter(s => s);
      
      let large = '';
      let medium = '';
      let small = '';
      
      if (parts.length >= 1) {
        large = parts[0];
      }
      if (parts.length >= 2) {
        medium = parts[1];
      }
      if (parts.length >= 3) {
        small = parts.slice(2).join(' ');
      } else if (parts.length === 2) {
        // 2つの部分しかない場合、最初の部分をlarge、2番目をmediumとする
        // smallは空のまま
        large = parts[0];
        medium = parts[1];
      }
      
      return {
        large: large || jobTypeText,
        medium: medium || '',
        small: small || '',
        rawText: jobTypeText
      };
    } catch (error) {
      console.error('職種情報の取得エラー:', error.message);
      return {
        large: '',
        medium: '',
        small: '',
        rawText: ''
      };
    }
  }

  /**
   * プレビューページから給与情報を取得する
   * @param {Page} previewPage - プレビューページ
   * @returns {Promise<{type: string, amount: number}>}
   */
  async getSalary(previewPage) {
    try {
      // PDFの取説によると「1番はじめに書いてある給与」を参照する必要がある
      // 仕事一覧プレビューのiframe内から最初の給与情報を取得
      let salaryText = '';
      
      try {
        // 仕事一覧プレビューのiframeを取得
        const iframe = await previewPage.$(config.selectors.preview.jobListPreview);
        if (iframe) {
          // iframeが読み込まれるまで待機
          await previewPage.waitForTimeout(2000);
          
          const iframeContent = await iframe.contentFrame();
          if (iframeContent) {
            // iframe内で給与情報を取得
            // 方法1: 「給与」などのテキストがある親階層のdlタグを取得
            try {
              // 「給与」「時給」「日給」「月給」を含むdt要素を探す
              const salaryDtXPath = "//dt[contains(text(), '給与') or contains(text(), '時給') or contains(text(), '日給') or contains(text(), '月給')]";
              const salaryDtElements = await iframeContent.$x(salaryDtXPath);
              
              if (salaryDtElements.length > 0) {
                // 最初のdt要素の親のdlタグを取得
                const dlElement = await iframeContent.evaluateHandle((dtEl) => {
                  let parent = dtEl.parentElement;
                  while (parent && parent.tagName !== 'DL') {
                    parent = parent.parentElement;
                  }
                  return parent;
                }, salaryDtElements[0]);
                
                if (dlElement && dlElement.asElement()) {
                  // dlタグ内のdd要素からli要素を取得
                  const salaryTextFromDl = await iframeContent.evaluate((dl) => {
                    const dd = dl.querySelector('dd');
                    if (dd) {
                      const li = dd.querySelector('li, ul > li');
                      if (li) {
                        return (li.textContent || li.innerText || '').trim();
                      }
                      // liがない場合はddのテキストを取得
                      return (dd.textContent || dd.innerText || '').trim();
                    }
                    return '';
                  }, dlElement.asElement());
                  
                  if (salaryTextFromDl) {
                    salaryText = salaryTextFromDl;
                    console.log(`  dlタグから給与情報を取得: ${salaryText}`);
                  }
                }
              }
            } catch (error) {
              console.log(`  dlタグからの給与情報取得エラー: ${error.message}`);
            }
            
            // 方法2: 指定されたセレクターを使用（フォールバック）
            if (!salaryText) {
              try {
                const salaryElement = await iframeContent.$(config.selectors.preview.salaryInIframeSelector);
                if (salaryElement) {
                  salaryText = await iframeContent.evaluate(el => {
                    return (el.textContent || el.innerText || '').trim();
                  }, salaryElement);
                  if (salaryText) {
                    console.log(`  指定セレクターから給与情報を取得: ${salaryText}`);
                  }
                }
              } catch (error) {
                console.log(`  指定セレクターからの給与情報取得エラー: ${error.message}`);
              }
            }
          }
        }
      } catch (error) {
        console.log(`  iframe内の給与情報取得エラー: ${error.message}`);
      }
      
      // iframe内で見つからない場合、メインページで給与情報を探す
      if (!salaryText) {
        try {
          // まず、金額を含むテキストを優先的に探す
          const salarySelectorsWithAmount = [
            "//*[contains(text(), '時給') and contains(text(), '円')]",
            "//*[contains(text(), '日給') and contains(text(), '円')]",
            "//*[contains(text(), '月給') and contains(text(), '円')]",
            "//td[contains(text(), '時給') and contains(text(), '円')]",
            "//td[contains(text(), '日給') and contains(text(), '円')]",
            "//td[contains(text(), '月給') and contains(text(), '円')]",
            "//span[contains(text(), '時給') and contains(text(), '円')]",
            "//span[contains(text(), '日給') and contains(text(), '円')]",
            "//span[contains(text(), '月給') and contains(text(), '円')]",
            "//div[contains(text(), '時給') and contains(text(), '円')]",
            "//div[contains(text(), '日給') and contains(text(), '円')]",
            "//div[contains(text(), '月給') and contains(text(), '円')]"
          ];
          
          // 金額を含むテキストを優先的に探す
          for (const selector of salarySelectorsWithAmount) {
            try {
              const salaryElements = await previewPage.$x(selector);
              if (salaryElements.length > 0) {
                const candidateText = await previewPage.evaluate(el => {
                  return (el.textContent || el.innerText || '').trim();
                }, salaryElements[0]);
                
                if (candidateText && (candidateText.includes('円') || /\d{3,}/.test(candidateText))) {
                  salaryText = candidateText;
                  break;
                }
              }
            } catch (error) {
              continue;
            }
          }
          
          // 金額を含むテキストが見つからない場合、通常のセレクターを試す
          if (!salaryText) {
            const salarySelectors = [
              config.selectors.preview.salaryXPath,
              "//td[contains(text(), '給与')]/following-sibling::td[1]",
              "//th[contains(text(), '給与')]/following-sibling::td[1]",
              "//td[contains(text(), '給与') or contains(text(), '時給') or contains(text(), '日給') or contains(text(), '月給')]",
              "//td[contains(text(), '円')]",
              "//div[contains(text(), '給与') or contains(text(), '時給') or contains(text(), '日給') or contains(text(), '月給')]",
              "//span[contains(text(), '時給') or contains(text(), '日給') or contains(text(), '月給')]"
            ];
            
            let tempSalaryText = ''; // 金額を含まないテキストを一時保存
            for (const selector of salarySelectors) {
              try {
                const salaryElements = await previewPage.$x(selector);
                if (salaryElements.length > 0) {
                  const candidateText = await previewPage.evaluate(el => {
                    // 要素のテキストを取得
                    let text = el.textContent || el.innerText || '';
                    // 親要素から給与情報を含む完全なテキストを取得
                    let parent = el.parentElement;
                    let depth = 0;
                    while (parent && depth < 3 && text.length < 200) {
                      const parentText = parent.textContent || parent.innerText || '';
                      if (parentText.includes('時給') || parentText.includes('日給') || parentText.includes('月給')) {
                        // 給与情報を含む行を抽出
                        const lines = parentText.split('\n').filter(line => 
                          line.includes('時給') || line.includes('日給') || line.includes('月給')
                        );
                        if (lines.length > 0) {
                          text = lines[0].trim();
                          break;
                        }
                      }
                      parent = parent.parentElement;
                      depth++;
                    }
                    return text.trim();
                  }, salaryElements[0]);
                  
                  // 金額を含むテキストを優先（「円」を含むテキストを優先）
                  if (candidateText && (candidateText.includes('時給') || candidateText.includes('日給') || candidateText.includes('月給'))) {
                    // 金額を含むテキストを優先（「円」を含む場合は採用して終了）
                    if (candidateText.includes('円') || /\d{3,}/.test(candidateText)) {
                      salaryText = candidateText;
                      break;
                    } else {
                      // 金額を含まないテキストは一時保存（他のセレクターで見つからない場合のフォールバック）
                      if (!tempSalaryText) {
                        tempSalaryText = candidateText;
                      }
                    }
                  }
                }
              } catch (error) {
                // 次のセレクターを試す
                continue;
              }
            }
            
            // 金額を含むテキストが見つからなかった場合、一時保存したテキストを使用
            if (!salaryText && tempSalaryText) {
              salaryText = tempSalaryText;
            }
          }
          
          // それでも見つからない場合、プレビューページ全体から金額を含むテキストを検索
          if (!salaryText || (!salaryText.includes('円') && !/\d{3,}/.test(salaryText))) {
            try {
              const allTextWithAmount = await previewPage.evaluate(() => {
                // ページ全体から金額を含む給与情報を探す
                const allElements = document.querySelectorAll('*');
                const candidates = [];
                
                for (const el of allElements) {
                  const text = (el.textContent || el.innerText || '').trim();
                  // 給与形態と金額の両方を含むテキストを探す
                  if ((text.includes('時給') || text.includes('日給') || text.includes('月給')) && 
                      (text.includes('円') || /\d{3,}/.test(text))) {
                    // 親要素のテキストも確認（より完全な情報を取得）
                    let fullText = text;
                    let parent = el.parentElement;
                    let depth = 0;
                    while (parent && depth < 2) {
                      const parentText = (parent.textContent || parent.innerText || '').trim();
                      if (parentText.includes('時給') || parentText.includes('日給') || parentText.includes('月給')) {
                        if (parentText.length < 200 && (parentText.includes('円') || /\d{3,}/.test(parentText))) {
                          fullText = parentText;
                          break;
                        }
                      }
                      parent = parent.parentElement;
                      depth++;
                    }
                    candidates.push(fullText);
                  }
                }
                
                // 最初に見つかった金額を含むテキストを返す
                return candidates.length > 0 ? candidates[0] : '';
              });
              
              if (allTextWithAmount) {
                salaryText = allTextWithAmount;
                console.log(`  プレビューページ全体から給与情報を取得: ${salaryText}`);
              }
            } catch (error) {
              console.log(`  プレビューページ全体の検索エラー: ${error.message}`);
            }
          }
        } catch (error) {
          console.log('  メインページで給与情報が見つかりませんでした');
        }
      }
      
      if (!salaryText) {
        // 給与情報が見つからない場合、デフォルト値を返す
        console.log('  給与情報が見つかりませんでした');
        return {
          type: '時給',
          amount: 0
        };
      }
      
      console.log(`  給与テキスト: ${salaryText}`);

      // 給与情報の解析
      // 例: "時給1,200円" → {type: "時給", amount: 1200}
      // 例: "日給10,000円" → {type: "日給", amount: 10000}
      // 例: "月給200,000円" → {type: "月給", amount: 200000}
      // 例: "月給21万円～22万円" → {type: "月給", amount: "月給21万円～22万円"} (文字列形式)
      
      let type = '時給'; // デフォルト
      let amount = 0;
      let isStringFormat = false; // 文字列形式かどうか

      // 給与形態を判定（最初に見つかったものを使用）
      // 「月収」を優先的に検出（「月給」より先にチェック）
      if (salaryText.includes('月収')) {
        type = '月収';
      } else if (salaryText.includes('時給')) {
        type = '時給';
      } else if (salaryText.includes('日給')) {
        type = '日給';
      } else if (salaryText.includes('月給')) {
        type = '月給';
      } else if (salaryText.includes('年俸')) {
        type = '年俸';
      }

      // 文字列形式の給与情報をチェック（「万円」や「～」が含まれている場合）
      // 例: "月給21万円～22万円"、"時給1,200円～1,500円"など
      if (salaryText.includes('万円') || (salaryText.includes('～') && salaryText.includes('円'))) {
        // 文字列形式の給与情報としてそのまま格納
        isStringFormat = true;
        // 給与形態を含む完全な文字列を取得
        const typeIndex = salaryText.indexOf(type);
        if (typeIndex !== -1) {
          // 給与形態以降の文字列を取得
          amount = salaryText.substring(typeIndex).trim();
        } else {
          // 給与形態が見つからない場合は、元のテキストをそのまま使用
          amount = salaryText.trim();
        }
        console.log(`  給与形態: ${type}, 金額（文字列形式）: ${amount}`);
        return { type, amount: amount }; // 文字列として返す
      }

      // 数値形式の給与情報を抽出
      // 金額を抽出（カンマと円を除去して数値に変換）
      // 給与形態の後の最初の数値を取得
      const typeIndex = salaryText.indexOf(type);
      if (typeIndex !== -1) {
        const afterType = salaryText.substring(typeIndex + type.length);
        // より正確な金額抽出：数値（カンマ含む）と「円」の前の数値を取得
        // 例: "時給1,500円～2,000円" → 1500を取得
        // 例: "日給10,000円" → 10000を取得
        // 例: "[ア・パ]時給1,500円～2,000円" → 1500を取得
        const amountMatch = afterType.match(/([\d,]+)\s*円/);
        if (amountMatch) {
          amount = parseInt(amountMatch[1].replace(/,/g, ''), 10);
        } else {
          // 「円」がない場合、最初の数値を取得
          const numberMatch = afterType.match(/([\d,]+)/);
          if (numberMatch) {
            amount = parseInt(numberMatch[1].replace(/,/g, ''), 10);
          }
        }
      } else {
        // 給与形態が見つからない場合でも、数値を探す
        const numberMatch = salaryText.match(/([\d,]+)\s*円/);
        if (numberMatch) {
          amount = parseInt(numberMatch[1].replace(/,/g, ''), 10);
        }
      }
      
      // 金額が抽出できなかった場合、テキスト全体から金額を探す
      if (amount === 0) {
        // テキスト全体から「円」を含む数値を探す
        const globalAmountMatch = salaryText.match(/([\d,]+)\s*円/);
        if (globalAmountMatch) {
          amount = parseInt(globalAmountMatch[1].replace(/,/g, ''), 10);
        } else {
          // 「円」がない場合でも、4桁以上の数値（給与の可能性が高い）を探す
          const largeNumberMatch = salaryText.match(/([\d,]{4,})/);
          if (largeNumberMatch) {
            const potentialAmount = parseInt(largeNumberMatch[1].replace(/,/g, ''), 10);
            // 1000円以上の場合のみ採用（誤検出を避ける）
            if (potentialAmount >= 1000) {
              amount = potentialAmount;
            }
          }
        }
      }

      console.log(`  給与形態: ${type}, 金額: ${amount}`);
      return { type, amount };
    } catch (error) {
      console.error('給与情報の取得エラー:', error.message);
      return {
        type: '時給',
        amount: 0
      };
    }
  }

  /**
   * プレビューページから店名（応募受付先名）を取得する
   * @param {Page} previewPage - プレビューページ
   * @returns {Promise<string>} 店名
   */
  async getStoreName(previewPage) {
    try {
      // まず、メインページから応募受付先名を取得
      const storeNameElements = await previewPage.$x(config.selectors.preview.storeNameXPath);
      if (storeNameElements.length > 0) {
        const fullText = await previewPage.evaluate(el => el.textContent.trim(), storeNameElements[0]);
        console.log(`  応募受付先名テキスト: ${fullText}`);
        
        // 「応募受付先名 : 」の後のテキストを抽出
        const match = fullText.match(/応募受付先名\s*[:：]\s*(.+)/);
        const storeName = match ? match[1].trim() : fullText.replace(/応募受付先名\s*[:：]\s*/, '').trim();
        
        if (storeName) {
          console.log(`  店名: ${storeName}`);
          return storeName;
        }
      }
      
      // メインページで見つからない場合、iframe内を探す
      const iframe = await previewPage.$(config.selectors.preview.jobListPreview);
      if (iframe) {
        const iframeContent = await iframe.contentFrame();
        if (iframeContent) {
          try {
            // iframe内で応募受付先名を探す
            const iframeStoreNameElements = await iframeContent.$x("//*[contains(text(), '応募受付先名')]");
            if (iframeStoreNameElements.length > 0) {
              const fullText = await iframeContent.evaluate(el => el.textContent.trim(), iframeStoreNameElements[0]);
              console.log(`  応募受付先名テキスト（iframe）: ${fullText}`);
              
              const match = fullText.match(/応募受付先名\s*[:：]\s*(.+)/);
              const storeName = match ? match[1].trim() : fullText.replace(/応募受付先名\s*[:：]\s*/, '').trim();
              
              if (storeName) {
                console.log(`  店名（iframe）: ${storeName}`);
                return storeName;
              }
            }
          } catch (error) {
            console.log(`  iframe内の店名取得エラー: ${error.message}`);
          }
        }
      }
      
      // 見つからない場合
      console.log('  店名: 未入力（一旦OK）');
      return '';
    } catch (error) {
      console.error('店名の取得エラー:', error.message);
      return '';
    }
  }

  /**
   * リトライ付きで関数を実行する
   * @param {Function} fn - 実行する関数
   * @param {number} retries - リトライ回数
   * @returns {Promise<any>}
   */
  async retry(fn, retries = config.retry.maxRetries) {
    for (let i = 0; i < retries; i++) {
      try {
        return await fn();
      } catch (error) {
        if (i === retries - 1) throw error;
        await new Promise(resolve => 
          setTimeout(resolve, config.retry.delay)
        );
      }
    }
  }
}

export default ScrapingService;
