import puppeteer, { Browser } from 'puppeteer';
import { logger } from '../../utils/logger';

export class BrowserService {
  async init(headless: boolean = true): Promise<Browser> {
    const options = {
      headless: headless ? ('new' as const) : false,
      args: [
        '--no-sandbox',
        '--disable-setuid-sandbox',
        '--disable-dev-shm-usage',
        '--disable-accelerated-2d-canvas',
        '--disable-gpu',
      ],
      defaultViewport: {
        width: 1920,
        height: 1080,
      },
    };

    try {
      logger.info(`ブラウザを起動します (headless: ${headless})`);
      const browser = await puppeteer.launch(options);
      logger.info('ブラウザの起動に成功しました');
      return browser;
    } catch (error) {
      logger.error('ブラウザの起動に失敗しました', error);
      throw error;
    }
  }
}

export default BrowserService;
