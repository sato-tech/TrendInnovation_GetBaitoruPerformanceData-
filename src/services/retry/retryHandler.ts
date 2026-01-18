import { logger } from '../../utils/logger';

export interface RetryOptions {
  maxRetries?: number;
  retryDelay?: number;
  timeout?: number;
  onRetry?: (error: Error, attempt: number) => void;
}

export class RetryHandler {
  private defaultMaxRetries: number;
  private defaultRetryDelay: number;

  constructor(maxRetries: number = 3, retryDelay: number = 2000) {
    this.defaultMaxRetries = maxRetries;
    this.defaultRetryDelay = retryDelay;
  }

  async execute<T>(
    fn: () => Promise<T>,
    operationName: string,
    options: RetryOptions = {}
  ): Promise<T> {
    const maxRetries = options.maxRetries ?? this.defaultMaxRetries;
    const retryDelay = options.retryDelay ?? this.defaultRetryDelay;
    const timeout = options.timeout;

    let lastError: Error | null = null;

    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        if (timeout) {
          return await Promise.race([
            fn(),
            new Promise<T>((_, reject) =>
              setTimeout(() => reject(new Error(`タイムアウト: ${operationName}`)), timeout)
            ),
          ]);
        } else {
          return await fn();
        }
      } catch (error) {
        lastError = error instanceof Error ? error : new Error(String(error));
        
        if (options.onRetry) {
          options.onRetry(lastError, attempt);
        }

        if (attempt < maxRetries) {
          logger.warn(`${operationName} が失敗しました (試行 ${attempt}/${maxRetries})`, {
            error: lastError.message,
          });
          
          await this.sleep(retryDelay * attempt); // 指数バックオフ
          logger.info(`${operationName} をリトライします...`);
        } else {
          logger.error(`${operationName} が最大リトライ回数に達しました`, {
            error: lastError.message,
            stack: lastError.stack,
          });
        }
      }
    }

    throw lastError || new Error(`${operationName} が失敗しました`);
  }

  private sleep(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }
}

export default RetryHandler;
