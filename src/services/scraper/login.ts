import { Page } from 'puppeteer';
import { logger } from '../../utils/logger';
import { Config } from '../../config';

export class LoginService {
  async login(page: Page, config: Config): Promise<void> {
    try {
      logger.info(`ログインページにアクセス: ${config.scraper.loginUrl}`);
      await page.goto(config.scraper.loginUrl, {
        waitUntil: 'networkidle2',
        timeout: config.scraper.pageLoadTimeout,
      });

      // ログインフォームのセレクタは実際のサイトに合わせて調整が必要
      // ここでは一般的なパターンを想定
      logger.info('ログイン情報を入力中...');
      
      // ユーザー名入力（セレクタは実際のサイトに合わせて変更）
      await page.waitForSelector('input[type="text"], input[name="username"], input[name="email"], input[id="username"], input[id="email"]', {
        timeout: 10000,
      });
      await page.type('input[type="text"], input[name="username"], input[name="email"], input[id="username"], input[id="email"]', config.scraper.username, {
        delay: 100,
      });

      // パスワード入力
      await page.waitForSelector('input[type="password"]', { timeout: 10000 });
      await page.type('input[type="password"]', config.scraper.password, {
        delay: 100,
      });

      // ログインボタンクリック
      logger.info('ログインボタンをクリック...');
      await page.click('button[type="submit"], input[type="submit"], button:has-text("ログイン"), button:has-text("Login")');

      // ログイン完了を待機（URL変更または特定要素の出現を確認）
      await page.waitForNavigation({
        waitUntil: 'networkidle2',
        timeout: config.scraper.pageLoadTimeout,
      });

      logger.info('ログインに成功しました');
      
      // ログイン成功の確認（実際のサイトに合わせて調整）
      const currentUrl = page.url();
      if (currentUrl.includes('login') || currentUrl === config.scraper.loginUrl) {
        throw new Error('ログインに失敗した可能性があります。URLが変更されていません');
      }
    } catch (error) {
      logger.error('ログイン処理でエラーが発生しました', error);
      throw error;
    }
  }
}

export default LoginService;
