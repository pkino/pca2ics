/**
 * PCA2ICS ユーティリティ関数
 */

/**
 * 値が存在するかチェック
 */
function hasValue(value: unknown): boolean {
  return value !== null && value !== undefined && value !== '';
}

/**
 * シート名からYYYYMM形式の日付を抽出
 * @returns 日付文字列（YYYYMM）またはnull
 */
function extractDateFromSheetName(sheetName: string): string | null {
  // YYYYMM形式（6桁の数字）にマッチ
  const match = sheetName.match(/^(\d{6})$/);
  if (match) {
    return match[1];
  }
  return null;
}

/**
 * 候補シートから最新の日付シートを検出
 * @param sheetNames シート名の配列
 * @returns 最新のシート名、見つからない場合はnull
 */
function findLatestDateSheet(sheetNames: string[]): string | null {
  let latestSheet: string | null = null;
  let latestDate: string | null = null;

  for (const name of sheetNames) {
    const date = extractDateFromSheetName(name);
    if (date) {
      if (!latestDate || date > latestDate) {
        latestDate = date;
        latestSheet = name;
      }
    }
  }

  return latestSheet;
}

/**
 * 日付フォーマット変換: 20250930 → 2025/9/30
 */
function formatDate(dateValue: string | number | Date): string {
  const dateStr = String(dateValue);

  if (dateStr.length === 8) {
    const year = dateStr.substring(0, 4);
    const month = String(parseInt(dateStr.substring(4, 6), 10));
    const day = String(parseInt(dateStr.substring(6, 8), 10));
    return `${year}/${month}/${day}`;
  }

  return dateStr;
}

/**
 * 科目コード変換: PCAコード(3桁) → ICSコード(3桁)
 */
function convertKamokuCode(
  srcCode: string | number | null | undefined,
  mapping: { [key: string]: string | number },
  denpyoNo?: string | number
): string {
  if (!srcCode) return '';

  // PCAコードを文字列化（3桁）
  const srcCodeStr = String(Math.floor(Number(srcCode))).padStart(3, '0');

  // マッピングテーブルからICSコードを取得
  const destCode = mapping[String(Math.floor(Number(srcCode)))];

  if (!destCode) {
    const errorMsg = `科目コード ${srcCode} のマッピングが見つかりません`;

    ERROR_LOGS.push({
      timestamp: new Date(),
      level: 'ERROR',
      function: 'convertKamokuCode',
      sourceSheet: CONFIG.SHEETS.SOURCE_DATA,
      denpyoNo: denpyoNo || '',
      message: errorMsg,
      stack: ''
    });

    Logger.log(`エラー: ${errorMsg}${denpyoNo ? ` (伝票番号: ${denpyoNo})` : ''}`);
    return ''; // 空白を返す（勝手にマッピングしない）
  }

  // 3桁で出力（必要に応じて0埋め）
  return String(Math.floor(Number(destCode))).padStart(3, '0');
}

/**
 * 税区分変換: PCA公益 → ICS db形式
 */
function convertTaxCode(
  pcaCode: string | number | null | undefined,
  taxMapping: TaxMapping,
  denpyoNo?: string | number
): string {
  if (!pcaCode) return ''; // 値がない場合は空白

  const icsCode = taxMapping[String(pcaCode)];

  if (!icsCode) {
    const errorMsg = `税区分 ${pcaCode} のマッピングが見つかりません`;

    ERROR_LOGS.push({
      timestamp: new Date(),
      level: 'ERROR',
      function: 'convertTaxCode',
      sourceSheet: CONFIG.SHEETS.SOURCE_DATA,
      denpyoNo: denpyoNo || '',
      message: errorMsg,
      stack: ''
    });

    Logger.log(`エラー: ${errorMsg}${denpyoNo ? ` (伝票番号: ${denpyoNo})` : ''}`);
    return ''; // 空白を返す（勝手にマッピングしない）
  }

  return icsCode;
}

/**
 * 最適な税区分を選択（00以外を優先）
 */
function selectBestTaxCode(
  karikataTaxCode: string | number | null | undefined,
  kashikataTaxCode: string | number | null | undefined,
  taxMapping: TaxMapping,
  denpyoNo?: string | number
): string {
  // 両方とも値がない場合
  if (!karikataTaxCode && !kashikataTaxCode) {
    return convertTaxCode('00', taxMapping, denpyoNo);
  }

  // 片方だけの場合
  if (!karikataTaxCode) {
    return convertTaxCode(kashikataTaxCode, taxMapping, denpyoNo);
  }
  if (!kashikataTaxCode) {
    return convertTaxCode(karikataTaxCode, taxMapping, denpyoNo);
  }

  // 両方ある場合
  const karikataStr = String(karikataTaxCode);
  const kashikataStr = String(kashikataTaxCode);

  // 借方が00以外 → 借方を採用
  if (karikataStr !== '00') {
    return convertTaxCode(karikataStr, taxMapping, denpyoNo);
  }

  // 借方が00、貸方が00以外 → 貸方を採用
  if (kashikataStr !== '00') {
    return convertTaxCode(kashikataStr, taxMapping, denpyoNo);
  }

  // 両方00 → 00を採用
  return convertTaxCode('00', taxMapping, denpyoNo);
}
