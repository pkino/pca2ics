/**
 * PCA2ICS 設定
 */

const CONFIG: Config = {
  // デフォルトのシート名（動的に検出、YYYYMM形式で最新のものを選択）
  DEFAULT_SOURCE_SHEET: '',

  // シート名
  SHEETS: {
    SOURCE_DATA: '',                     // 元データシート（実行時に設定）
    KAMOKU_MAPPING: '科目対応表',        // 科目対応表シート
    TAX_MAPPING: '税区分マッピング',      // 税区分マッピングシート
    OUTPUT: 'ICS変換結果',               // 出力シート
    ERROR_LOG: 'エラーログ'              // エラーログシート
  },

  // カラムインデックス（0始まり）
  COLUMNS: {
    SOURCE: {
      DATE: 0,              // 伝票日付
      DENPYO_NO: 1,         // 伝票番号
      KARIKATE_DEPT: 5,     // 借方部門コード
      KARIKATE_KAMOKU: 7,   // 借方科目コード
      KARIKATE_NAME: 8,     // 借方科目名
      KARIKATE_HOJO: 9,     // 借方補助コード
      KARIKATE_HOJO_NAME: 10, // 借方補助名
      KARIKATE_TAX: 11,     // 借方税区分コード
      KARIKATE_AMOUNT: 13,  // 借方金額
      KARIKATE_TAX_AMOUNT: 14, // 借方消費税額
      KASHIKATE_DEPT: 16,   // 貸方部門コード
      KASHIKATE_KAMOKU: 18, // 貸方科目コード
      KASHIKATE_NAME: 19,   // 貸方科目名
      KASHIKATE_HOJO: 20,   // 貸方補助コード
      KASHIKATE_HOJO_NAME: 21, // 貸方補助名
      KASHIKATE_TAX: 22,    // 貸方税区分コード
      KASHIKATE_AMOUNT: 24, // 貸方金額
      KASHIKATE_TAX_AMOUNT: 25, // 貸方消費税額
      TEKIYO: 26            // 摘要文
    }
  }
};

// エラーログ一時保存用（1回の実行中のみ有効）
let ERROR_LOGS: ErrorLogEntry[] = [];
