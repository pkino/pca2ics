/**
 * PCA公益法人会計V.12 → ICS財務処理db (db仕訳形式) 変換スクリプト
 *
 * 機能:
 * - 202509.csv形式のデータをICS財務処理db形式に変換
 * - 科目コード変換 (PCAコード → ICSコード)
 * - 税区分変換 (PCA公益 → ICS db形式)
 * - 複合仕訳の単純仕訳への分解
 * - 日付フォーマット変換
 */

/**
 * 元データシートを選択
 */
function selectSourceSheet(ss: GoogleAppsScript.Spreadsheet.Spreadsheet): string | null {
  const ui = SpreadsheetApp.getUi();
  const sheets = ss.getSheets();

  // 候補となるシートをリストアップ（科目対応表、税区分マッピング、出力シート、エラーログを除外）
  const excludeSheets = [
    CONFIG.SHEETS.KAMOKU_MAPPING,
    CONFIG.SHEETS.TAX_MAPPING,
    CONFIG.SHEETS.OUTPUT,
    CONFIG.SHEETS.ERROR_LOG
  ];

  const candidateSheets = sheets
    .map(sheet => sheet.getName())
    .filter(name => !excludeSheets.includes(name));

  if (candidateSheets.length === 0) {
    ui.alert('エラー', '変換可能なシートが見つかりません。', ui.ButtonSet.OK);
    return null;
  }

  // 最新の日付シートを動的に検出
  const latestDateSheet = findLatestDateSheet(candidateSheets);

  if (candidateSheets.length === 1) {
    // 候補が1つだけの場合は確認して使用
    const response = ui.alert(
      '確認',
      `元データシート「${candidateSheets[0]}」を使用しますか？`,
      ui.ButtonSet.YES_NO
    );

    return response === ui.Button.YES ? candidateSheets[0] : null;
  }

  // 複数ある場合は選択ダイアログを表示
  // 日付シートを先頭に、それ以外を後ろにソート
  const sortedSheets = [...candidateSheets].sort((a, b) => {
    const dateA = extractDateFromSheetName(a);
    const dateB = extractDateFromSheetName(b);

    // 両方日付シートの場合は降順（新しい順）
    if (dateA && dateB) {
      return dateB.localeCompare(dateA);
    }
    // 日付シートを優先
    if (dateA) return -1;
    if (dateB) return 1;
    // それ以外はアルファベット順
    return a.localeCompare(b);
  });

  let message = '変換する元データシートを選択してください:\n\n';
  sortedSheets.forEach((name, index) => {
    const isLatest = name === latestDateSheet;
    message += `${index + 1}. ${name}${isLatest ? ' (最新)' : ''}\n`;
  });
  message += '\n番号を入力してください（空欄で最新を選択）:';

  const response = ui.prompt(
    '元データシート選択',
    message,
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return null; // キャンセル
  }

  const input = response.getResponseText().trim();

  // 空欄の場合は最新シートを選択
  if (input === '' && latestDateSheet) {
    return latestDateSheet;
  }

  const selectedIndex = parseInt(input, 10) - 1;

  if (isNaN(selectedIndex) || selectedIndex < 0 || selectedIndex >= sortedSheets.length) {
    ui.alert('エラー', '無効な番号です。', ui.ButtonSet.OK);
    return null;
  }

  return sortedSheets[selectedIndex];
}

/**
 * メイン実行関数
 */
function convertPCAtoICS(): void {
  // エラーログを初期化
  ERROR_LOGS = [];

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 元データシートを選択
    const sourceSheetName = selectSourceSheet(ss);
    if (!sourceSheetName) {
      return; // キャンセルされた
    }

    CONFIG.SHEETS.SOURCE_DATA = sourceSheetName;

    Logger.log('=== PCA → ICS 変換開始 ===');
    Logger.log(`元データシート: ${sourceSheetName}`);

    // 1. データ読み込み
    Logger.log('1. データ読み込み中...');
    const sourceData = loadSourceData(ss);
    const kamokuMapping = loadKamokuMapping(ss);
    const taxMapping = loadTaxMapping(ss);

    Logger.log(`元データ: ${sourceData.length}行`);
    Logger.log(`科目マッピング: ${Object.keys(kamokuMapping.codeMap).length}件`);
    Logger.log(`税区分マッピング: ${Object.keys(taxMapping).length}件`);

    // 2. データ変換
    Logger.log('2. データ変換中...');
    const convertedData = convertData(sourceData, kamokuMapping, taxMapping);

    Logger.log(`変換後: ${convertedData.length}行`);

    // 3. 出力
    Logger.log('3. 出力中...');
    outputData(ss, convertedData);

    // エラーログ出力（必ず実行）
    writeErrorLog(ss, ERROR_LOGS);

    // すべてのシート操作を確実に完了させる
    SpreadsheetApp.flush();

    Logger.log('=== 変換完了 ===');

    // エラーの有無で完了メッセージを変える
    if (ERROR_LOGS.length > 0) {
      SpreadsheetApp.getUi().alert(
        '⚠️ 変換完了（警告あり）',
        `変換が完了しましたが、${ERROR_LOGS.length}件の警告があります。\n\n` +
        `元データ: ${sourceSheetName}\n` +
        `出力シート: ${CONFIG.SHEETS.OUTPUT}\n` +
        `変換行数: ${convertedData.length}行\n\n` +
        `⚠️ 詳細は「${CONFIG.SHEETS.ERROR_LOG}」シートを確認してください。`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } else {
      SpreadsheetApp.getUi().alert(
        '✅ 変換完了',
        `変換が正常に完了しました！\n\n` +
        `元データ: ${sourceSheetName}\n` +
        `出力シート: ${CONFIG.SHEETS.OUTPUT}\n` +
        `変換行数: ${convertedData.length}行\n\n` +
        `エラー: なし`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }

  } catch (error) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const err = error as Error;

    Logger.log('エラー: ' + err.message);
    Logger.log(err.stack || '');

    // 致命的エラーもエラーログシートへ
    ERROR_LOGS.push({
      timestamp: new Date(),
      level: 'ERROR',
      function: 'convertPCAtoICS',
      sourceSheet: CONFIG.SHEETS.SOURCE_DATA || '',
      denpyoNo: '',
      message: err.message,
      stack: err.stack || ''
    });

    writeErrorLog(ss, ERROR_LOGS);

    SpreadsheetApp.getUi().alert(
      '❌ エラー',
      `エラーが発生しました:\n${err.message}\n\n` +
      `詳細は「${CONFIG.SHEETS.ERROR_LOG}」シートを確認してください。`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * 特定のシートを指定して変換
 * 使用例: convertSpecificSheet('202510')
 */
function convertSpecificSheet(sheetName: string): void {
  // エラーログを初期化
  ERROR_LOGS = [];

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // シートの存在確認
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`シート "${sheetName}" が見つかりません`);
    }

    CONFIG.SHEETS.SOURCE_DATA = sheetName;

    Logger.log('=== PCA → ICS 変換開始 ===');
    Logger.log(`元データシート: ${sheetName}`);

    // 1. データ読み込み
    Logger.log('1. データ読み込み中...');
    const sourceData = loadSourceData(ss);
    const kamokuMapping = loadKamokuMapping(ss);
    const taxMapping = loadTaxMapping(ss);

    Logger.log(`元データ: ${sourceData.length}行`);
    Logger.log(`科目マッピング: ${Object.keys(kamokuMapping.codeMap).length}件`);
    Logger.log(`税区分マッピング: ${Object.keys(taxMapping).length}件`);

    // 2. データ変換
    Logger.log('2. データ変換中...');
    const convertedData = convertData(sourceData, kamokuMapping, taxMapping);

    Logger.log(`変換後: ${convertedData.length}行`);

    // 3. 出力
    Logger.log('3. 出力中...');
    outputData(ss, convertedData);

    // エラーログ出力（必ず実行）
    writeErrorLog(ss, ERROR_LOGS);

    // すべてのシート操作を確実に完了させる
    SpreadsheetApp.flush();

    Logger.log('=== 変換完了 ===');

    // エラーの有無で完了メッセージを変える
    if (ERROR_LOGS.length > 0) {
      SpreadsheetApp.getUi().alert(
        '⚠️ 変換完了（警告あり）',
        `変換が完了しましたが、${ERROR_LOGS.length}件の警告があります。\n\n` +
        `元データ: ${sheetName}\n` +
        `出力シート: ${CONFIG.SHEETS.OUTPUT}\n` +
        `変換行数: ${convertedData.length}行\n\n` +
        `⚠️ 詳細は「${CONFIG.SHEETS.ERROR_LOG}」シートを確認してください。`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } else {
      SpreadsheetApp.getUi().alert(
        '✅ 変換完了',
        `変換が正常に完了しました！\n\n` +
        `元データ: ${sheetName}\n` +
        `出力シート: ${CONFIG.SHEETS.OUTPUT}\n` +
        `変換行数: ${convertedData.length}行\n\n` +
        `エラー: なし`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }

  } catch (error) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const err = error as Error;

    Logger.log('エラー: ' + err.message);
    Logger.log(err.stack || '');

    // 致命的エラーもエラーログシートへ
    ERROR_LOGS.push({
      timestamp: new Date(),
      level: 'ERROR',
      function: 'convertSpecificSheet',
      sourceSheet: CONFIG.SHEETS.SOURCE_DATA || sheetName || '',
      denpyoNo: '',
      message: err.message,
      stack: err.stack || ''
    });

    writeErrorLog(ss, ERROR_LOGS);

    SpreadsheetApp.getUi().alert(
      '❌ エラー',
      `エラーが発生しました:\n${err.message}\n\n` +
      `詳細は「${CONFIG.SHEETS.ERROR_LOG}」シートを確認してください。`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * スプレッドシート起動時にメニューを追加
 */
function onOpen(): void {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('PCA→ICS変換')
    .addItem('変換実行', 'convertPCAtoICS')
    .addSeparator()
    .addItem('税区分マッピングシートを作成', 'createTaxMappingSheetManually')
    .addToUi();
}

/**
 * 税区分マッピングシートを手動作成
 */
function createTaxMappingSheetManually(): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 既存シートの確認
  const existingSheet = ss.getSheetByName(CONFIG.SHEETS.TAX_MAPPING);
  if (existingSheet) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '確認',
      '税区分マッピングシートは既に存在します。上書きしますか？',
      ui.ButtonSet.YES_NO
    );

    if (response === ui.Button.YES) {
      ss.deleteSheet(existingSheet);
    } else {
      return;
    }
  }

  createTaxMappingSheet(ss);
  SpreadsheetApp.getUi().alert('税区分マッピングシートを作成しました！');
}
