/**
 * PCA2ICS 出力関数
 */

/**
 * エラーログシートを取得または作成
 */
function getOrCreateErrorLogSheet(
  ss: GoogleAppsScript.Spreadsheet.Spreadsheet
): GoogleAppsScript.Spreadsheet.Sheet {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.ERROR_LOG);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.ERROR_LOG);
  }

  // ヘッダーが未設定の場合のみヘッダー行を設定
  if (sheet.getLastRow() === 0) {
    const headers = [
      'タイムスタンプ',
      'レベル',
      '処理名',
      '元シート',
      '伝票番号',
      'メッセージ',
      'スタックトレース'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#f8d7da');
    sheet.setFrozenRows(1);
  }

  return sheet;
}

/**
 * エラーログをシートに追記
 */
function writeErrorLog(
  ss: GoogleAppsScript.Spreadsheet.Spreadsheet,
  logs: ErrorLogEntry[]
): void {
  if (!logs || logs.length === 0) return;

  const sheet = getOrCreateErrorLogSheet(ss);
  const startRow = sheet.getLastRow() + 1;

  const values = logs.map(log => [
    log.timestamp || new Date(),
    log.level || '',
    log.function || '',
    log.sourceSheet || '',
    log.denpyoNo || '',
    log.message || '',
    log.stack || ''
  ]);

  sheet.getRange(startRow, 1, values.length, values[0].length).setValues(values);
}

/**
 * 変換データを出力
 */
function outputData(
  ss: GoogleAppsScript.Spreadsheet.Spreadsheet,
  data: ICSOutputRow[]
): void {
  // 出力シートを取得または作成
  let outputSheet = ss.getSheetByName(CONFIG.SHEETS.OUTPUT);
  if (outputSheet) {
    outputSheet.clear();
  } else {
    outputSheet = ss.insertSheet(CONFIG.SHEETS.OUTPUT);
  }

  // 固定ヘッダー行（1-4行目）
  const fixedHeaders: (string | number)[][] = [
    ['法人'],
    ['db仕訳日記帳'],
    ['6', '株式会社　木重漆器店'],
    ['自 7年 4月 1日', '至 8年 3月31日', '月分']
  ];

  // 列ヘッダー行（5行目）
  const headers: string[] = [
    '日付', '決修', '伝票番号',
    '借方部門コード', '借方事管区分', '借方工事コード', '借方コード', '借方名称',
    '借方枝番', '借方枝番摘要', '借方枝番カナ',
    '貸方部門コード', '貸方事管区分', '貸方工事コード', '貸方コード', '貸方名称',
    '貸方枝番', '貸方枝番摘要', '貸方枝番カナ',
    '金額', '摘要', '税区分', '対価', '仕入区分', '売上業種区分',
    '仕訳区分', '特定収入区分', 'ダミー1', 'ダミー2', 'ダミー3', '内部取引',
    '税額', '証憑番号', '手形番号', '手形期日', '付箋番号', '付箋コメント',
    '免税事業者等', 'インボイス登録番号'
  ];

  // 固定ヘッダー + 列ヘッダー + データを出力
  const outputData: (string | number)[][] = [...fixedHeaders, headers, ...data];
  outputSheet.getRange(1, 1, outputData.length, Math.max(...outputData.map(row => row.length))).setValues(outputData);

  // 列ヘッダー行を固定（5行目）
  outputSheet.setFrozenRows(5);

  // 列ヘッダー行を太字に（5行目）
  outputSheet.getRange(5, 1, 1, headers.length).setFontWeight('bold');

  Logger.log(`${data.length}行を出力しました`);
}
