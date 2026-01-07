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

  // CSV形式に変換（ダブルクォートなし、Windows CRLF改行）
  const csvContent = allData.map(row =>
    row.map(cell => {
      // 日付は yyyy/M/d 形式に整形（1桁の月日は先頭0なし）
      if (cell instanceof Date) {
        return Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy/M/d');
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

/**
 * CSVファイルをShift_JISからUTF-8に変換してインポート
 */
function importCSV(): void {
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
          .upload-box {
            border: 2px dashed #ccc;
            border-radius: 8px;
            padding: 40px;
            margin: 20px 0;
            cursor: pointer;
            transition: all 0.3s;
          }
          .upload-box:hover {
            border-color: #4CAF50;
            background-color: #f9f9f9;
          }
          .upload-box.drag-over {
            border-color: #4CAF50;
            background-color: #e8f5e9;
          }
          input[type="file"] {
            display: none;
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
          .info {
            color: #666;
          }
          input[type="text"] {
            padding: 10px;
            font-size: 14px;
            border: 1px solid #ccc;
            border-radius: 4px;
            width: 200px;
            margin: 10px;
          }
        </style>
      </head>
      <body>
        <h2>CSV インポート (Shift_JIS → UTF-8)</h2>
        <p class="info">PCA公益法人会計からエクスポートしたCSVファイルを選択してください</p>

        <div class="upload-box" id="uploadBox" onclick="document.getElementById('fileInput').click()">
          <p id="uploadText">📂 クリックしてファイルを選択<br>またはドラッグ&ドロップ</p>
          <input type="file" id="fileInput" accept=".csv" onchange="handleFileSelect(event)">
        </div>

        <div id="fileInfo" style="display:none; margin: 20px 0;">
          <p><strong>選択されたファイル:</strong> <span id="fileName"></span></p>
          <label for="sheetName">インポート先シート名:</label>
          <input type="text" id="sheetName" placeholder="例: 202601" value="">
          <br>
          <button id="importBtn" onclick="importCSVFile()">インポート実行</button>
        </div>

        <div id="status"></div>

        <script>
          let selectedFile = null;

          // ドラッグ&ドロップ対応
          const uploadBox = document.getElementById('uploadBox');

          uploadBox.addEventListener('dragover', function(e) {
            e.preventDefault();
            uploadBox.classList.add('drag-over');
          });

          uploadBox.addEventListener('dragleave', function(e) {
            e.preventDefault();
            uploadBox.classList.remove('drag-over');
          });

          uploadBox.addEventListener('drop', function(e) {
            e.preventDefault();
            uploadBox.classList.remove('drag-over');
            const files = e.dataTransfer.files;
            if (files.length > 0) {
              handleFile(files[0]);
            }
          });

          function handleFileSelect(event) {
            const files = event.target.files;
            if (files.length > 0) {
              handleFile(files[0]);
            }
          }

          function handleFile(file) {
            if (!file.name.toLowerCase().endsWith('.csv')) {
              document.getElementById('status').innerHTML = '<span class="error">CSVファイルを選択してください</span>';
              return;
            }

            selectedFile = file;
            document.getElementById('fileName').textContent = file.name;

            // ファイル名から日付を抽出してシート名を提案（例: 202509.csv → 202509）
            const baseName = file.name.replace(/\.csv$/i, '');
            const dateMatch = baseName.match(/\d{6}/);
            if (dateMatch) {
              document.getElementById('sheetName').value = dateMatch[0];
            } else {
              document.getElementById('sheetName').value = baseName;
            }

            document.getElementById('fileInfo').style.display = 'block';
            document.getElementById('status').innerHTML = '';
          }

          function importCSVFile() {
            if (!selectedFile) {
              document.getElementById('status').innerHTML = '<span class="error">ファイルを選択してください</span>';
              return;
            }

            const sheetName = document.getElementById('sheetName').value.trim();
            if (!sheetName) {
              document.getElementById('status').innerHTML = '<span class="error">シート名を入力してください</span>';
              return;
            }

            const btn = document.getElementById('importBtn');
            const status = document.getElementById('status');

            btn.disabled = true;
            status.innerHTML = 'ファイル読み込み中...';

            const reader = new FileReader();
            reader.onload = function(e) {
              try {
                status.innerHTML = '文字コード変換中...';

                // ArrayBufferをUint8Arrayに変換
                const uint8Array = new Uint8Array(e.target.result);

                // Shift_JISからUnicodeに変換
                const unicodeArray = Encoding.convert(uint8Array, {
                  to: 'UNICODE',
                  from: 'SJIS'
                });

                // Unicodeの数値配列を文字列に変換
                const csvText = Encoding.codeToString(unicodeArray);

                status.innerHTML = 'CSV解析中...';

                // CSVを解析（改行で分割して2次元配列に変換）
                const lines = csvText.split(/\r?\n/);
                const data = lines.map(line => {
                  // 簡易CSVパーサー（カンマ区切り）
                  return line.split(',');
                });

                status.innerHTML = 'スプレッドシートに書き込み中...';

                // サーバー側にデータを送信
                google.script.run
                  .withSuccessHandler(function(result) {
                    status.innerHTML = '<span class="success">✅ インポート完了！<br>' +
                      'シート「' + sheetName + '」に ' + result.rowCount + ' 行を書き込みました。<br>' +
                      'このウィンドウを閉じてください。</span>';
                    btn.disabled = false;
                  })
                  .withFailureHandler(function(error) {
                    status.innerHTML = '<span class="error">❌ エラー: ' + error.message + '</span>';
                    btn.disabled = false;
                  })
                  .writeCSVToSheet(sheetName, data);

              } catch (error) {
                status.innerHTML = '<span class="error">❌ エラー: ' + error.message + '</span>';
                btn.disabled = false;
              }
            };

            reader.onerror = function() {
              status.innerHTML = '<span class="error">❌ ファイル読み込みエラー</span>';
              btn.disabled = false;
            };

            reader.readAsArrayBuffer(selectedFile);
          }
        </script>
      </body>
    </html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(600)
    .setTitle('CSV インポート');

  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

/**
 * CSVデータをシートに書き込む（サーバー側関数）
 */
function writeCSVToSheet(sheetName: string, data: string[][]): { rowCount: number } {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // データが空の場合はエラー
  if (!data || data.length === 0) {
    throw new Error('CSVデータが空です');
  }

  // 既存シートを確認
  let sheet = ss.getSheetByName(sheetName);

  if (sheet) {
    // 既存シートがある場合は確認（UIから呼ばれるので直接上書き）
    sheet.clear();
  } else {
    // 新規シート作成
    sheet = ss.insertSheet(sheetName);
  }

  // データを書き込み
  const rowCount = data.length;
  const colCount = Math.max(...data.map(row => row.length));

  // 行ごとに列数が違う場合があるので、空文字で埋める
  const normalizedData = data.map(row => {
    const newRow = [...row];
    while (newRow.length < colCount) {
      newRow.push('');
    }
    return newRow;
  });

  sheet.getRange(1, 1, rowCount, colCount).setValues(normalizedData);

  // 1行目をフリーズ（ヘッダー行として）
  if (rowCount >= 2) {
    sheet.setFrozenRows(2); // PCA形式は1行目がバージョン、2行目がヘッダー
  }

  Logger.log(`シート「${sheetName}」に ${rowCount} 行を書き込みました`);

  return { rowCount: rowCount };
}
