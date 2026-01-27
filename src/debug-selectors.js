/**
 * ログイン画面のセレクターを特定するためのデバッグスクリプト
 * 使用方法: node src/debug-selectors.js
 */

import ScrapingService from './services/ScrapingService.js';
import config from '../config/config.js';
import { writeFileSync } from 'fs';
import { join } from 'path';

async function debugSelectors() {
  const scrapingService = new ScrapingService();

  try {
    console.log('=== ログイン画面のセレクター特定デバッグ ===\n');

    // ブラウザを起動
    console.log('1. ブラウザを起動中...');
    await scrapingService.launchBrowser();
    console.log('✓ ブラウザを起動しました\n');

    // ログイン画面に遷移
    console.log(`2. ログイン画面に遷移中: ${config.baitoru.loginUrl}`);
    await scrapingService.page.goto(config.baitoru.loginUrl, {
      waitUntil: 'networkidle2',
      timeout: 60000
    });
    console.log('✓ ログイン画面に遷移しました\n');

    // 少し待機（ページが完全に読み込まれるまで）
    await new Promise(resolve => setTimeout(resolve, 2000));

    // ページのHTMLを取得
    console.log('3. ページのHTMLを取得中...');
    const html = await scrapingService.page.content();
    const htmlPath = join(process.cwd(), 'debug-login-page.html');
    writeFileSync(htmlPath, html, 'utf-8');
    console.log(`✓ HTMLを保存しました: ${htmlPath}\n`);

    // スクリーンショットを取得
    console.log('4. スクリーンショットを取得中...');
    const screenshotPath = join(process.cwd(), 'debug-login-page.png');
    await scrapingService.page.screenshot({ path: screenshotPath, fullPage: true });
    console.log(`✓ スクリーンショットを保存しました: ${screenshotPath}\n`);

    // 入力フィールドを探す
    console.log('5. 入力フィールドを検索中...');
    
    // 一般的なユーザー名/ID入力フィールドのセレクター候補
    const usernameSelectors = [
      'input[type="text"]',
      'input[type="email"]',
      'input[name*="user"]',
      'input[name*="id"]',
      'input[name*="login"]',
      'input[id*="user"]',
      'input[id*="id"]',
      'input[id*="login"]',
      'input[placeholder*="ユーザー"]',
      'input[placeholder*="ID"]',
      'input[placeholder*="ログイン"]',
      '#username',
      '#user',
      '#loginId',
      '#login_id',
      '.username',
      '.user-id'
    ];

    // 一般的なパスワード入力フィールドのセレクター候補
    const passwordSelectors = [
      'input[type="password"]',
      'input[name*="pass"]',
      'input[id*="pass"]',
      'input[placeholder*="パスワード"]',
      '#password',
      '#pass',
      '.password'
    ];

    // 一般的なログインボタンのセレクター候補
    const buttonSelectors = [
      'button[type="submit"]',
      'input[type="submit"]',
      'button:contains("ログイン")',
      'button:contains("Login")',
      'input[value*="ログイン"]',
      'input[value*="Login"]',
      'button.login',
      '.login-button',
      'a.login',
      '[onclick*="login"]'
    ];

    console.log('\n--- ユーザー名/ID入力フィールドの候補 ---');
    for (const selector of usernameSelectors) {
      try {
        const elements = await scrapingService.page.$$(selector);
        if (elements.length > 0) {
          const firstElement = elements[0];
          const tagName = await scrapingService.page.evaluate(el => el.tagName, firstElement);
          const id = await scrapingService.page.evaluate(el => el.id, firstElement);
          const name = await scrapingService.page.evaluate(el => el.name, firstElement);
          const placeholder = await scrapingService.page.evaluate(el => el.placeholder, firstElement);
          const className = await scrapingService.page.evaluate(el => el.className, firstElement);
          
          console.log(`✓ 見つかりました: ${selector}`);
          console.log(`  タグ: ${tagName}`);
          if (id) console.log(`  ID: #${id}`);
          if (name) console.log(`  name: [name="${name}"]`);
          if (placeholder) console.log(`  placeholder: "${placeholder}"`);
          if (className) console.log(`  class: .${className.split(' ').join('.')}`);
          console.log('');
        }
      } catch (e) {
        // セレクターが無効な場合はスキップ
      }
    }

    console.log('\n--- パスワード入力フィールドの候補 ---');
    for (const selector of passwordSelectors) {
      try {
        const elements = await scrapingService.page.$$(selector);
        if (elements.length > 0) {
          const firstElement = elements[0];
          const tagName = await scrapingService.page.evaluate(el => el.tagName, firstElement);
          const id = await scrapingService.page.evaluate(el => el.id, firstElement);
          const name = await scrapingService.page.evaluate(el => el.name, firstElement);
          const placeholder = await scrapingService.page.evaluate(el => el.placeholder, firstElement);
          const className = await scrapingService.page.evaluate(el => el.className, firstElement);
          
          console.log(`✓ 見つかりました: ${selector}`);
          console.log(`  タグ: ${tagName}`);
          if (id) console.log(`  ID: #${id}`);
          if (name) console.log(`  name: [name="${name}"]`);
          if (placeholder) console.log(`  placeholder: "${placeholder}"`);
          if (className) console.log(`  class: .${className.split(' ').join('.')}`);
          console.log('');
        }
      } catch (e) {
        // セレクターが無効な場合はスキップ
      }
    }

    console.log('\n--- ログインボタンの候補 ---');
    for (const selector of buttonSelectors) {
      try {
        const elements = await scrapingService.page.$$(selector);
        if (elements.length > 0) {
          const firstElement = elements[0];
          const tagName = await scrapingService.page.evaluate(el => el.tagName, firstElement);
          const id = await scrapingService.page.evaluate(el => el.id, firstElement);
          const name = await scrapingService.page.evaluate(el => el.name, firstElement);
          const value = await scrapingService.page.evaluate(el => el.value || el.textContent, firstElement);
          const className = await scrapingService.page.evaluate(el => el.className, firstElement);
          
          console.log(`✓ 見つかりました: ${selector}`);
          console.log(`  タグ: ${tagName}`);
          if (id) console.log(`  ID: #${id}`);
          if (name) console.log(`  name: [name="${name}"]`);
          if (value) console.log(`  value/text: "${value.trim()}"`);
          if (className) console.log(`  class: .${className.split(' ').join('.')}`);
          console.log('');
        }
      } catch (e) {
        // セレクターが無効な場合はスキップ
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
        const buttons = Array.from(document.querySelectorAll('button, input[type="submit"], input[type="button"]'));
        return buttons.map(button => ({
          tag: button.tagName,
          type: button.type || '',
          id: button.id || '',
          name: button.name || '',
          className: button.className || '',
          text: (button.textContent || button.value || '').trim(),
          onclick: button.onclick ? 'あり' : 'なし'
        }));
      });

      allButtons.forEach((button, index) => {
        console.log(`\nボタン ${index + 1}:`);
        console.log(`  タグ: ${button.tag}`);
        if (button.type) console.log(`  タイプ: ${button.type}`);
        if (button.id) console.log(`  ID: #${button.id}`);
        if (button.name) console.log(`  name: [name="${button.name}"]`);
        if (button.text) console.log(`  テキスト: "${button.text}"`);
        if (button.className) {
          const classes = button.className.split(' ').filter(c => c);
          if (classes.length > 0) {
            console.log(`  class: .${classes.join('.')}`);
          }
        }
        if (button.onclick === 'あり') console.log(`  onclick: あり`);
      });
    } catch (e) {
      console.log('button要素の取得に失敗しました:', e.message);
    }

    console.log('\n=== デバッグ完了 ===');
    console.log('\n次のステップ:');
    console.log('1. debug-login-page.html と debug-login-page.png を確認してください');
    console.log('2. 上記の情報を元に、config/selectors.json を更新してください');
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

debugSelectors();
