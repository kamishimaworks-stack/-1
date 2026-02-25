/**
 * ============================================================
 *  SimilarCompany.js - 機能5: 類似企業・ターゲット拡張
 * ============================================================
 *  新規登録企業のデータを元に、Gemini で:
 *   - 競合他社・関連企業を3〜5社推論
 *   - 営業ターゲット候補として別シートに出力
 * ============================================================
 */

const SimilarCompanyFinder = (() => {

  /**
   * 類似企業検索のシステムプロンプト
   */
  const SYSTEM_PROMPT = `あなたは日本のBtoB市場に精通したマーケットリサーチャーです。
営業チームが次のターゲットを見つけるための情報を提供してください。

【あなたの役割】
- 指定された企業と「類似したビジネスモデル」を持つ企業を提案する
- 競合他社だけでなく、同じ業種・規模・課題を持つ類似企業も含める
- 各企業について、なぜターゲット候補として適切なのか理由を付記する
- 営業優先度（高/中/低）を判定する

【判定基準】
- 高: 同じ業種・同規模で、類似の課題を抱える可能性が高い
- 中: 関連業種で、一部の課題が共通する可能性がある
- 低: 業種は異なるが、同様のビジネスモデルを持つ`;

  /**
   * 類似企業を検索・推論
   * @param {string} companyName - 基準企業名
   * @param {string} [industry]  - 業種（IndustryAnalysisの結果）
   * @return {Object} { summary, companies: [...] }
   */
  function find(companyName, industry) {
    if (!companyName) {
      return { summary: '', companies: [] };
    }

    const count = SFA_CONFIG.PARAMS.SIMILAR_COMPANY_COUNT;

    const userPrompt = `以下の企業の情報を元に、類似企業・競合他社を${count}社提案してください。

【基準企業】
- 企業名: ${companyName}
- 業種: ${industry || '不明（推定してください）'}

以下のJSON形式で回答してください:
{
  "companies": [
    {
      "name": "企業名",
      "industry": "業種",
      "reason": "類似理由（50文字以内）",
      "priority": "高" | "中" | "低",
      "estimatedUrl": "推定される公式サイトURL"
    }
  ],
  "summary": "ターゲット候補の概要（100文字以内の要約）"
}`;

    try {
      const result = GeminiService.generateJson(SYSTEM_PROMPT, userPrompt);

      return {
        summary:   result.summary || '',
        companies: Array.isArray(result.companies) ? result.companies : [],
      };

    } catch (e) {
      console.error('[SimilarCompanyFinder] 推論失敗: ' + e.message);
      return { summary: 'APIエラーにより取得できませんでした', companies: [] };
    }
  }

  /**
   * 類似企業リストを別シートに出力
   * @param {string} baseCompany - 基準企業名
   * @param {Object[]} companies  - find() で取得した companies 配列
   */
  function writeToSheet(baseCompany, companies) {
    if (!companies || companies.length === 0) return;

    const now = new Date();

    companies.forEach(comp => {
      SheetHelper.appendToSheet(SFA_CONFIG.SHEETS.SIMILAR_COMPANIES, [
        baseCompany,                    // 基準企業名
        comp.name || '',                // 類似企業名
        comp.industry || '',            // 業種
        comp.reason || '',              // 類似理由
        comp.priority || '中',          // ターゲット優先度
        comp.estimatedUrl || '',        // 推定URL
        now,                            // 生成日
      ]);
    });

    console.log(`[SimilarCompanyFinder] ${baseCompany} → ${companies.length}社の類似企業を出力`);
  }

  /**
   * 既存顧客全体に対して一括で類似企業分析を実行
   * ※ 手動実行用 (API消費量が大きいため注意)
   * 類似企業が空の企業を全件処理する
   */
  function batchAnalyze() {
    const customers = SheetHelper.getAllCustomers();

    // 類似企業カラムが空の顧客のみ対象
    const targets = customers.filter(c =>
      c.companyName && !c.similar  // 類似企業が空の顧客のみ対象
    );

    console.log(`[SimilarCompanyFinder] バッチ分析開始: ${targets.length}件`);

    let count = 0;
    targets.forEach(customer => {
      try {
        const result = find(customer.companyName, customer.industry);

        // メインシートの類似企業カラムを更新
        if (result.summary) {
          SheetHelper.updateCell(customer.rowIndex, SFA_CONFIG.COL.SIMILAR, result.summary);
        }

        // 別シートに詳細出力
        writeToSheet(customer.companyName, result.companies);

        count++;
        Utilities.sleep(2000); // API制限対策

      } catch (e) {
        console.warn(`[SimilarCompanyFinder] ${customer.companyName} 処理失敗: ${e.message}`);
      }
    });

    console.log(`[SimilarCompanyFinder] バッチ分析完了: ${count}/${targets.length}件処理`);
    return { processed: count, total: targets.length };
  }

  // ── Public API ──
  return {
    find,
    writeToSheet,
    batchAnalyze,
  };

})();


// ==========================================
// グローバル関数 (手動実行用)
// ==========================================

/**
 * 類似企業バッチ分析を手動実行するためのエントリーポイント
 * GASエディタから「batchSimilarCompanyAnalysis」を選択して実行
 */
function batchSimilarCompanyAnalysis() {
  const result = SimilarCompanyFinder.batchAnalyze();
  console.log('バッチ分析結果: ' + JSON.stringify(result));
}
