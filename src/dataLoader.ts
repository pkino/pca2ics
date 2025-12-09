/**
 * PCA2ICS データ読み込み関数
 */

/**
 * 元データを読み込む
 */
function loadSourceData(ss: GoogleAppsScript.Spreadsheet.Spreadsheet): unknown[][] {
  const sheet = ss.getSheetByName(CONFIG.SHEETS.SOURCE_DATA);
  if (!sheet) {
    throw new Error(`シート "${CONFIG.SHEETS.SOURCE_DATA}" が見つかりません`);
  }

  const range = sheet.getDataRange();
  const values = range.getValues();

  // 1行目: バージョン情報
  // 2行目: ヘッダー行
  // 3行目以降: データ
  // → 最初の2行をスキップ
  return values.slice(2);
}

/**
 * 科目対応表を読み込む
 */
function loadKamokuMapping(ss: GoogleAppsScript.Spreadsheet.Spreadsheet): KamokuMapping {
  const sheet = ss.getSheetByName(CONFIG.SHEETS.KAMOKU_MAPPING);
  if (!sheet) {
    throw new Error(`シート "${CONFIG.SHEETS.KAMOKU_MAPPING}" が見つかりません`);
  }

  const range = sheet.getDataRange();
  const values = range.getValues();

  // マッピングオブジェクトを作成
  const codeMap: { [key: string]: string | number } = {};
  const nameMap: { [key: string]: string } = {};

  // ヘッダー行をスキップして処理
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const kamokuName = row[0] as string;  // 勘定科目名（列A）
    const icsCode = row[1];                // ICSコード（列B）
    const pcaCode = row[2];                // PCAコード（列C）

    if (pcaCode && icsCode) {
      // PCAコード → ICSコードのマッピング
      codeMap[String(pcaCode)] = icsCode;
    }

    if (icsCode && kamokuName) {
      // ICSコード → 科目名のマッピング
      nameMap[String(icsCode)] = kamokuName;
    }
  }

  return {
    codeMap: codeMap,
    nameMap: nameMap
  };
}

/**
 * 税区分マッピングを読み込む
 */
function loadTaxMapping(ss: GoogleAppsScript.Spreadsheet.Spreadsheet): TaxMapping {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.TAX_MAPPING);

  // シートが存在しない場合は作成
  if (!sheet) {
    Logger.log('税区分マッピングシートが見つかりません。自動作成します。');
    sheet = createTaxMappingSheet(ss);
  }

  const range = sheet.getDataRange();
  const values = range.getValues();

  // マッピングオブジェクトを作成
  const mapping: TaxMapping = {};

  // ヘッダー行をスキップして処理
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const pcaCode = row[0];     // PCA公益コード（列A）
    const icsCode = row[1];     // ICSコード（列B）

    if (pcaCode && icsCode) {
      mapping[String(pcaCode)] = String(icsCode);
    }
  }

  return mapping;
}

/**
 * 税区分マッピングシートを作成
 */
function createTaxMappingSheet(
  ss: GoogleAppsScript.Spreadsheet.Spreadsheet
): GoogleAppsScript.Spreadsheet.Sheet {
  const sheet = ss.insertSheet(CONFIG.SHEETS.TAX_MAPPING);

  // ヘッダーとデフォルトデータ
  const data: (string | number)[][] = [
    ['PCA公益コード', 'ICSコード', '説明'],
    ['00', '04', '消費税に関係ない → 不課税'],
    ['99', '04', '不明 → 不課税'],
    ['A0', '02', '非課税売上'],
    ['B5', '317', '課税売上10%'],
    ['C5', '317', '課税売上返還10%'],
    ['D5', '317', '貸倒れ10%'],
    ['E5', '317', '貸倒れ回収10%'],
    ['Q5', '317', '課税仕入10%'],
    ['R5', '317', '課税仕入返還10%'],
    ['F0', '40', '輸出免税売上'],
    ['G0', '02', '非課税売上の返還'],
    ['H0', '40', '輸出免税売上の返還'],
    ['P0', '02', '非課税仕入'],
    ['W0', '02', '非課税仕入の返還'],
    ['B1', '20', '課税売上3%'],
    ['B3', '207', '課税売上5%'],
    ['B4', '217', '課税売上8%'],
    ['C1', '20', '課税売上返還3%'],
    ['C3', '207', '課税売上返還5%'],
    ['C4', '217', '課税売上返還8%'],
    ['Q1', '20', '課税仕入3%'],
    ['Q3', '207', '課税仕入5%'],
    ['Q4', '217', '課税仕入8%'],
    ['R1', '20', '課税仕入返還3%'],
    ['R3', '207', '課税仕入返還5%'],
    ['R4', '217', '課税仕入返還8%']
  ];

  sheet.getRange(1, 1, data.length, 3).setValues(data);

  // ヘッダー行をフォーマット
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#e8f0fe');
  sheet.setFrozenRows(1);

  // 列幅を調整
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 300);

  Logger.log('税区分マッピングシートを作成しました');

  return sheet;
}
