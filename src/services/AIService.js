import OpenAI from 'openai';
import config from '../../config/config.js';

/**
 * OpenAI API連携サービス
 * ドロップダウン選択肢のマッチング処理を担当
 */
class AIService {
  constructor() {
    const apiKey = process.env.OPENAI_API_KEY;
    if (!apiKey) {
      console.warn('OPENAI_API_KEYが設定されていません。AI機能は使用できません。');
      this.client = null;
    } else {
      this.client = new OpenAI({
        apiKey: apiKey
      });
    }
  }

  /**
   * テキストから最も近い選択肢を判定する
   * @param {string} text - 判定するテキスト
   * @param {Array<string>} options - 選択肢の配列
   * @param {string} context - コンテキスト（例: "地方"、"職種"）
   * @returns {Promise<string|null>} 最も近い選択肢、見つからない場合はnull
   */
  async findBestMatch(text, options, context = '') {
    if (!this.client || !text || !options || options.length === 0) {
      return null;
    }

    try {
      const prompt = `以下のテキスト「${text}」に最も近い選択肢を、提供された選択肢リストから1つ選んでください。

選択肢:
${options.map((opt, idx) => `${idx + 1}. ${opt}`).join('\n')}

${context ? `コンテキスト: ${context}` : ''}

重要:
- 完全一致でなくても、意味的に最も近いものを選んでください
- 選択肢の番号のみを返答してください（例: "1"）
- 該当するものがなければ "なし" と返答してください`;

      const response = await this.client.chat.completions.create({
        model: process.env.OPENAI_MODEL || 'gpt-4o-mini',
        messages: [
          {
            role: 'system',
            content: 'あなたは選択肢マッチングの専門家です。テキストと選択肢リストを比較して、最も適切な選択肢を選んでください。'
          },
          {
            role: 'user',
            content: prompt
          }
        ],
        temperature: 0.3,
        max_tokens: 10
      });

      const result = response.choices[0].message.content.trim();
      
      // 番号を抽出
      const match = result.match(/\d+/);
      if (match) {
        const index = parseInt(match[0], 10) - 1;
        if (index >= 0 && index < options.length) {
          return options[index];
        }
      }

      return null;
    } catch (error) {
      console.error('AI判定エラー:', error.message);
      return null;
    }
  }

  /**
   * 地方を判定する（都道府県から地方を判定）
   * @param {string} prefecture - 都道府県名
   * @param {Array<string>} regionOptions - 地方の選択肢
   * @returns {Promise<string|null>} 判定された地方
   */
  async determineRegion(prefecture, regionOptions) {
    if (!prefecture) return null;
    
    // 都道府県から地方へのマッピング（簡易版）
    const prefectureToRegion = {
      '北海道': '北海道',
      '青森県': '東北',
      '岩手県': '東北',
      '宮城県': '東北',
      '秋田県': '東北',
      '山形県': '東北',
      '福島県': '東北',
      '茨城県': '関東',
      '栃木県': '関東',
      '群馬県': '関東',
      '埼玉県': '関東',
      '千葉県': '関東',
      '東京都': '関東',
      '神奈川県': '関東',
      '新潟県': '中部',
      '富山県': '中部',
      '石川県': '中部',
      '福井県': '中部',
      '山梨県': '中部',
      '長野県': '中部',
      '岐阜県': '中部',
      '静岡県': '中部',
      '愛知県': '中部',
      '三重県': '近畿',
      '滋賀県': '近畿',
      '京都府': '近畿',
      '大阪府': '近畿',
      '兵庫県': '近畿',
      '奈良県': '近畿',
      '和歌山県': '近畿',
      '鳥取県': '中国',
      '島根県': '中国',
      '岡山県': '中国',
      '広島県': '中国',
      '山口県': '中国',
      '徳島県': '四国',
      '香川県': '四国',
      '愛媛県': '四国',
      '高知県': '四国',
      '福岡県': '九州',
      '佐賀県': '九州',
      '長崎県': '九州',
      '熊本県': '九州',
      '大分県': '九州',
      '宮崎県': '九州',
      '鹿児島県': '九州',
      '沖縄県': '沖縄'
    };

    // まず簡易マッピングを試す
    const mappedRegion = prefectureToRegion[prefecture];
    if (mappedRegion && regionOptions.includes(mappedRegion)) {
      return mappedRegion;
    }

    // AI判定を使用
    return await this.findBestMatch(prefecture, regionOptions, '地方');
  }

  /**
   * 職種を判定する
   * @param {string} jobText - 職種テキスト
   * @param {Array<string>} categoryOptions - 職種カテゴリの選択肢
   * @returns {Promise<string|null>} 判定された職種
   */
  async determineJobCategory(jobText, categoryOptions) {
    if (!jobText) return null;
    return await this.findBestMatch(jobText, categoryOptions, '職種');
  }

  /**
   * プランを判定する（PEXプラン、Bプラン、ELプラン、Dプラン、Cプラン、Aプランから選択）
   * @param {string} planText - プランテキスト
   * @returns {Promise<string|null>} 判定されたプラン
   */
  async determinePlan(planText) {
    if (!planText) return null;
    
    const planOptions = ['PEXプラン', 'Bプラン', 'ELプラン', 'Dプラン', 'Cプラン', 'Aプラン'];
    
    // まず簡易マッチングを試す
    const upperPlanText = planText.toUpperCase();
    for (const plan of planOptions) {
      const planName = plan.replace('プラン', '').toUpperCase();
      if (upperPlanText.includes(planName) || upperPlanText.includes(plan)) {
        return plan;
      }
    }
    
    // AI判定を使用
    return await this.findBestMatch(planText, planOptions, 'プラン');
  }

  /**
   * コサイン類似度を計算する（簡易版：文字列の類似度）
   * @param {string} text1 - テキスト1
   * @param {string} text2 - テキスト2
   * @returns {number} 類似度（0-1）
   */
  calculateCosineSimilarity(text1, text2) {
    if (!text1 || !text2) return 0;
    
    // 簡易版：文字列の共通部分を計算
    const words1 = text1.toLowerCase().split(/\s+/);
    const words2 = text2.toLowerCase().split(/\s+/);
    
    const set1 = new Set(words1);
    const set2 = new Set(words2);
    
    const intersection = new Set([...set1].filter(x => set2.has(x)));
    const union = new Set([...set1, ...set2]);
    
    return union.size > 0 ? intersection.size / union.size : 0;
  }

  /**
   * 職種カテゴリをコサイン類似度で判定する
   * @param {string} jobText - 職種テキスト（3階層を繋げたもの）
   * @param {Array<Object>} jobCategoryList - 職種カテゴリリスト（{large, medium, small, combined}の配列）
   * @returns {Promise<{large: string, medium: string, small: string}|null>} 判定された職種カテゴリ
   */
  async determineJobCategoryByCosineSimilarity(jobText, jobCategoryList) {
    if (!jobText || !jobCategoryList || jobCategoryList.length === 0) {
      return null;
    }
    
    // 最も類似度の高いカテゴリを探す
    let maxSimilarity = 0;
    let bestMatch = null;
    
    for (const category of jobCategoryList) {
      const combined = category.combined || `${category.large} ${category.medium} ${category.small}`;
      const similarity = this.calculateCosineSimilarity(jobText, combined);
      
      if (similarity > maxSimilarity) {
        maxSimilarity = similarity;
        bestMatch = category;
      }
    }
    
    // 類似度が0.3以上の場合のみ返す
    if (maxSimilarity >= 0.3 && bestMatch) {
      return {
        large: bestMatch.large || '',
        medium: bestMatch.medium || '',
        small: bestMatch.small || ''
      };
    }
    
    return null;
  }
}

export default AIService;
