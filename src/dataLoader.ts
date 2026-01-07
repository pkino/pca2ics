/**
 * PCA2ICS データ読み込み関数
 *
 * - スプレッドシートシートからのデータ読み込み
 * - Shift_JIS CSV ファイルのインポート（UTF-8変換）
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
    const pcaCode = row[0];     // PCAコード（列A）
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
    ['PCAコード', 'ICSコード', '説明'],
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
        <p class="info">PCA商魂商管からエクスポートしたCSVファイルを選択してください</p>

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
            try {
              const files = event.target.files;
              if (files.length > 0) {
                handleFile(files[0]);
              }
            } catch (error) {
              document.getElementById('status').innerHTML = '<span class="error">エラー: ' + error.message + '</span>';
            }
          }

          function handleFile(file) {
            try {
              if (!file.name.toLowerCase().endsWith('.csv')) {
                document.getElementById('status').innerHTML = '<span class="error">CSVファイルを選択してください</span>';
                return;
              }

              selectedFile = file;
              document.getElementById('fileName').textContent = file.name;

              // ファイル名から日付を抽出してシート名を提案（例: 202509.csv → 202509）
              const baseName = file.name.replace(/\\.csv$/i, '');
              const dateMatch = baseName.match(/\\d{6}/);
              if (dateMatch) {
                document.getElementById('sheetName').value = dateMatch[0];
              } else {
                document.getElementById('sheetName').value = baseName;
              }

              document.getElementById('fileInfo').style.display = 'block';
              document.getElementById('status').innerHTML = '';
            } catch (error) {
              document.getElementById('status').innerHTML = '<span class="error">エラー: ' + error.message + '</span>';
            }
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
                // ArrayBufferをUint8Arrayに変換
                const uint8Array = new Uint8Array(e.target.result);

                // Unicodeの数値配列を文字列に変換
                const csvText = new TextDecoder('shift_jis').decode(uint8Array);

                // 文字化けチェック（?が含まれている場合は変換に失敗している可能性が高い）
                if (csvText.includes('?') || csvText.includes('\ufffd')) {
                  throw new Error('文字コード変換に失敗しました。ファイルがShift_JIS形式でない可能性があります。');
                }

                status.innerHTML = 'CSV解析中...';

                // CSVを解析（改行で分割して2次元配列に変換）
                const lines = csvText.split(/\\r?\\n/);
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
    // 既存シートがある場合はエラー
    throw new Error('シート「' + sheetName + '」は既に存在します。別のシート名を指定してください。');
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

  // データを書き込み
  const range = sheet.getRange(1, 1, rowCount, colCount);
  range.setValues(normalizedData);

  // すべてのセルを文字列として扱う（数字や日付を変換しない、000なども保持）
  range.setNumberFormat('@');

  // 1行目をフリーズ（ヘッダー行として）
  if (rowCount >= 2) {
    sheet.setFrozenRows(2); // PCA形式は1行目がバージョン、2行目がヘッダー
  }

  Logger.log(`シート「${sheetName}」に ${rowCount} 行を書き込みました`);

  return { rowCount: rowCount };
}
