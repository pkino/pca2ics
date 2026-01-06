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

  // 列ヘッダー行の定義
  const headers: string[] = [
    '日付', '決修', '伝票番号',
    '借方部門ｺｰﾄﾞ', '借方工事ｺｰﾄﾞ', '借方ｺｰﾄﾞ', '借方名称',
    '借方枝番', '借方枝番摘要', '借方枝番ｶﾅ',
    '貸方部門ｺｰﾄﾞ', '貸方工事ｺｰﾄﾞ', '貸方ｺｰﾄﾞ', '貸方名称',
    '貸方枝番', '貸方枝番摘要', '貸方枝番ｶﾅ',
    '金額', '摘要', '税区分', '対価', '仕入区分', '売上業種区分',
    '仕訳区分', 'ﾀﾞﾐｰ1', 'ﾀﾞﾐｰ2', 'ﾀﾞﾐｰ3', '税額', 'ﾀﾞﾐｰ5',
    '手形番号', '手形期日', '付箋番号', '付箋コメント',
    '免税事業者等', 'インボイス登録番号'
  ];

  const columnCount = headers.length;

  // 固定ヘッダー行（1-4行目）を列数に合わせて作成
  const fixedHeaders: (string | number)[][] = [
    ['法人', ...Array(columnCount - 1).fill('')],
    ['db仕訳日記帳', ...Array(columnCount - 1).fill('')],
    ['6', '株式会社　木重漆器店', ...Array(columnCount - 2).fill('')],
    ['自 7年 4月 1日', '至 8年 3月31日', '月分', ...Array(columnCount - 3).fill('')]
  ];

  // 固定ヘッダー + 列ヘッダー + データを出力
  const outputData: (string | number)[][] = [...fixedHeaders, headers, ...data];
  outputSheet.getRange(1, 1, outputData.length, columnCount).setValues(outputData);

  // 列ヘッダー行を固定（5行目）
  outputSheet.setFrozenRows(5);

  // 列ヘッダー行を太字に（5行目）
  outputSheet.getRange(5, 1, 1, columnCount).setFontWeight('bold');

  Logger.log(`${data.length}行を出力しました`);
}

/**
 * CSVコンテンツを取得（サーバー側関数）
 */
function getCSVContent(): string {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.OUTPUT);

  if (!sheet) {
    throw new Error('出力シートが見つかりません。先に変換を実行してください。');
  }

  // シートのすべてのデータを取得
  const allData = sheet.getDataRange().getValues();

  if (allData.length === 0) {
    throw new Error('出力シートにデータがありません。先に変換を実行してください。');
  }

  // ヘッダー行を除外して2行目以降のデータのみを取得
  const data = allData.slice(1);

  if (data.length === 0) {
    throw new Error('出力シートにデータ行がありません。');
  }

  // CSV形式に変換（ダブルクォートなし、Windows CRLF改行）
  const csvContent = data.map(row =>
    row.map(cell => {
      // 日付は yyyy/MM/dd 形式に整形
      if (cell instanceof Date) {
        return Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy/MM/dd');
      }

      let val = String(cell);

      // データ内のカンマ(,)は 全角カンマ(，) に置換して列ズレ防止
      val = val.replace(/,/g, '，');
      // データ内の改行は スペース に置換して行ズレ防止
      val = val.replace(/[\r\n]+/g, ' ');

      return val;
    }).join(',')
  ).join('\r\n')
    .replace(/\u301C/g, '\uFF5E')  // 〜 → ～（これでCP932寄りになりやすい）
    .replace(/\u2212/g, '\uFF0D'); // −(マイナス) → －(全角ハイフン) も地雷常連


  return csvContent;
}

/**
 * CSVデータをANSI（Shift_JIS）フォーマットでダウンロード
 */
function exportToCSV(): void {
  const html = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <script src="https://cdn.jsdelivr.net/npm/encoding-japanese@2.0.0/encoding.min.js"></script>
        <style>
          body {
            font-family: Arial, sans-serif;
            padding: 20px;
            text-align: center;
          }
          button {
            background-color: #4CAF50;
            color: white;
            padding: 15px 32px;
            text-align: center;
            font-size: 16px;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            margin: 10px;
          }
          button:hover {
            background-color: #45a049;
          }
          button:disabled {
            background-color: #cccccc;
            cursor: not-allowed;
          }
          #status {
            margin-top: 20px;
            font-size: 14px;
          }
          .error {
            color: red;
          }
          .success {
            color: green;
          }
        </style>
      </head>
      <body>
        <h2>CSV エクスポート (ANSI/Shift_JIS形式)</h2>
        <p>ダウンロードボタンをクリックしてください</p>
        <button id="downloadBtn" onclick="downloadCSV()">ダウンロード</button>
        <div id="status"></div>

        <script>
          function downloadCSV() {
            const btn = document.getElementById('downloadBtn');
            const status = document.getElementById('status');

            btn.disabled = true;
            status.innerHTML = '処理中...';

            google.script.run
              .withSuccessHandler(function(csvContent) {
                try {
                  // encoding-japaneseのstringToCodeメソッドを使用
                  const unicodeArray = Encoding.stringToCode(csvContent);

                  // UnicodeからShift_JISに変換
                  const sjisArray = Encoding.convert(unicodeArray, {
                    to: 'SJIS',
                    from: 'UNICODE'
                  });

                  // Uint8Arrayに変換
                  const uint8Array = new Uint8Array(sjisArray);

                  // Blobを作成（charset明示）
                  const blob = new Blob([uint8Array], { type: 'text/csv;charset=shift_jis' });

                  // ダウンロード
                  const url = URL.createObjectURL(blob);
                  const a = document.createElement('a');
                  a.href = url;
                  a.download = 'ICS変換結果.csv';
                  document.body.appendChild(a);
                  a.click();
                  document.body.removeChild(a);
                  URL.revokeObjectURL(url);

                  status.innerHTML = '<span class="success">ダウンロード完了！このウィンドウを閉じてください。</span>';
                } catch (error) {
                  status.innerHTML = '<span class="error">エラー: ' + error.message + '</span>';
                  btn.disabled = false;
                }
              })
              .withFailureHandler(function(error) {
                status.innerHTML = '<span class="error">エラー: ' + error.message + '</span>';
                btn.disabled = false;
              })
              .getCSVContent();
          }
        </script>
      </body>
    </html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setTitle('CSV エクスポート');

  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}
