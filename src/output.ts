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

  // ヘッダー行を作成
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

  // ヘッダーとデータを出力
  const outputData: (string | number)[][] = [headers, ...data];
  outputSheet.getRange(1, 1, outputData.length, headers.length).setValues(outputData);

  // ヘッダー行を固定
  outputSheet.setFrozenRows(1);

  // ヘッダー行を太字に
  outputSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

  Logger.log(`${data.length}行を出力しました`);
}

/**
 * CSVデータをANSI（Shift_JIS）フォーマットでダウンロード
 */
function exportToCSV(): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.OUTPUT);

  if (!sheet) {
    SpreadsheetApp.getUi().alert(
      'エラー',
      '出力シートが見つかりません。先に変換を実行してください。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  // シートのすべてのデータを取得
  const allData = sheet.getDataRange().getValues();

  if (allData.length === 0) {
    SpreadsheetApp.getUi().alert(
      'エラー',
      '出力シートにデータがありません。先に変換を実行してください。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  // ヘッダー行を除外して2行目以降のデータのみを取得
  const data = allData.slice(1);

  if (data.length === 0) {
    SpreadsheetApp.getUi().alert(
      'エラー',
      '出力シートにデータ行がありません。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  // CSV形式に変換（ダブルクォートなし）
  const csvContent = data.map(row =>
    row.map(cell => String(cell)).join(',')
  ).join('\n');

  // Shift_JIS（ANSI）エンコーディングでBlobを作成
  const blob = Utilities.newBlob(csvContent, 'text/csv; charset=Shift_JIS', 'ICS変換結果.csv');

  // Google Driveに保存
  const file = DriveApp.createFile(blob);
  const fileUrl = file.getUrl();
  const fileName = file.getName();

  // ダウンロードリンクを表示
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'CSV エクスポート完了',
    `CSVファイル（ANSI/Shift_JIS形式）をGoogle Driveに保存しました。\n\n` +
    `ファイル名: ${fileName}\n\n` +
    `以下のリンクからダウンロードしてください:\n${fileUrl}\n\n` +
    `ダウンロード完了後、OKを押すとファイルが削除されます。`,
    ui.ButtonSet.OK_CANCEL
  );

  // ダウンロード完了後、ファイルを削除
  if (response === ui.Button.OK) {
    file.setTrashed(true);
    Logger.log('CSVファイルを削除しました: ' + fileName);
  } else {
    Logger.log('CSVファイルは保持されます: ' + fileUrl);
  }
}
