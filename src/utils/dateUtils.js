/**
 * 日付ユーティリティ関数
 */

/**
 * Excelの日付値をDateオブジェクトに変換する
 * ExcelJSは日付セルをDateオブジェクト、数値（シリアル番号）、または文字列として読み込む可能性がある
 * @param {Date|number|string} value - Excelの日付値（Dateオブジェクト、シリアル番号、または文字列）
 * @returns {Date|null} 変換されたDateオブジェクト、または無効な場合はnull
 */
export function excelDateToJSDate(value) {
  // 既にDateオブジェクトの場合はそのまま返す
  if (value instanceof Date) {
    // 無効なDateオブジェクトの場合はnullを返す
    if (isNaN(value.getTime())) {
      return null;
    }
    return value;
  }

  // nullまたはundefinedの場合はnullを返す
  if (value === null || value === undefined) {
    return null;
  }

  // 数値の場合（Excelのシリアル番号）
  if (typeof value === 'number') {
    // 無効な数値の場合はnullを返す
    if (isNaN(value) || value <= 0) {
      return null;
    }
    const excelEpoch = new Date(1899, 11, 30);
    const jsDate = new Date(excelEpoch.getTime() + value * 86400000);
    // 変換結果が無効な場合はnullを返す
    if (isNaN(jsDate.getTime())) {
      return null;
    }
    return jsDate;
  }

  // 文字列の場合
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (trimmed === '') {
      return null;
    }
    
    // 日付形式の文字列（YYYY/MM/DD、YYYY-MM-DDなど）をチェック
    // スラッシュやハイフンを含む場合は日付形式として扱う
    if (trimmed.includes('/') || trimmed.includes('-')) {
      // 日付形式の文字列をパース
      // YYYY/MM/DD形式の場合、スラッシュをハイフンに置換してISO形式にする
      let dateString = trimmed;
      if (trimmed.includes('/')) {
        // YYYY/MM/DD形式をYYYY-MM-DD形式に変換
        const parts = trimmed.split('/');
        if (parts.length === 3) {
          const year = parts[0].padStart(4, '0');
          const month = parts[1].padStart(2, '0');
          const day = parts[2].padStart(2, '0');
          dateString = `${year}-${month}-${day}`;
        }
      }
      
      const parsedDate = new Date(dateString);
      if (!isNaN(parsedDate.getTime())) {
        return parsedDate;
      }
    }
    
    // 日付形式でない場合、数値文字列の可能性をチェック（Excelのシリアル番号が文字列として読み込まれた場合）
    // ただし、スラッシュやハイフンを含まない純粋な数値文字列のみ
    if (!trimmed.includes('/') && !trimmed.includes('-')) {
      const numericValue = parseFloat(trimmed);
      if (!isNaN(numericValue) && numericValue > 0 && numericValue < 1000000) {
        // Excelのシリアル番号は通常1000000未満（約2739年まで）
        const excelEpoch = new Date(1899, 11, 30);
        const jsDate = new Date(excelEpoch.getTime() + numericValue * 86400000);
        if (!isNaN(jsDate.getTime())) {
          return jsDate;
        }
      }
    }
    
    // その他の日付文字列としてパースを試みる
    const parsedDate = new Date(trimmed);
    if (!isNaN(parsedDate.getTime())) {
      return parsedDate;
    }
    
    return null;
  }

  // その他の型の場合はnullを返す
  return null;
}

/**
 * DateオブジェクトをYYYY/MM/DD形式の文字列に変換する
 * @param {Date|null} date - 日付オブジェクト
 * @returns {string} YYYY/MM/DD形式の文字列、または無効な場合は空文字列
 */
export function formatDateForInput(date) {
  // nullまたはundefinedの場合は空文字列を返す
  if (!date) {
    return '';
  }

  // Dateオブジェクトでない場合は空文字列を返す
  if (!(date instanceof Date)) {
    return '';
  }

  // 無効なDateオブジェクトの場合は空文字列を返す
  if (isNaN(date.getTime())) {
    return '';
  }

  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}/${month}/${day}`;
}

/**
 * 日付文字列をDateオブジェクトに変換する
 * @param {string} dateString - 日付文字列（YYYY-MM-DD形式）
 * @returns {Date}
 */
export function parseDate(dateString) {
  return new Date(dateString);
}

/**
 * 2つの日付の間の週数を計算する
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 * @returns {number} 週数
 */
export function calculateWeeks(startDate, endDate) {
  const diffTime = Math.abs(endDate.getTime() - startDate.getTime());
  const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
  return Math.ceil(diffDays / 7);
}
