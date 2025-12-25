/**
 * PCA2ICS データ変換関数
 */

/**
 * 伝票番号でグループ化
 */
function groupByDenpyo(data: unknown[][]): DenpyoGroups {
  const groups: DenpyoGroups = {};

  for (const row of data) {
    const denpyoNo = String(row[CONFIG.COLUMNS.SOURCE.DENPYO_NO]);

    if (!groups[denpyoNo]) {
      groups[denpyoNo] = [];
    }

    groups[denpyoNo].push(row);
  }

  return groups;
}

/**
 * 複合仕訳を単純仕訳に分解（方式D: 金額分割方式）
 */
function simplifyCompoundJournal(rows: unknown[][]): SimpleJournal[] {
  const simplified: CompoundJournal[] = [];
  const COL = CONFIG.COLUMNS.SOURCE;

  const karikatePending: JournalItem[] = [];
  const kashikatePending: JournalItem[] = [];

  for (const row of rows) {
    const hasKarikata = hasValue(row[COL.KARIKATE_KAMOKU]) &&
                        Number(row[COL.KARIKATE_AMOUNT]) > 0;
    const hasKashikata = hasValue(row[COL.KASHIKATE_KAMOKU]) &&
                         Number(row[COL.KASHIKATE_AMOUNT]) > 0;

    if (hasKarikata && hasKashikata) {
      // 両方ある場合 → そのまま追加
      karikatePending.push({
        row: row,
        kamoku: row[COL.KARIKATE_KAMOKU] as string | number,
        name: row[COL.KARIKATE_NAME] as string,
        hojo: row[COL.KARIKATE_HOJO] as string | number,
        hojoName: row[COL.KARIKATE_HOJO_NAME] as string,
        amount: Number(row[COL.KARIKATE_AMOUNT]),
        taxAmount: Number(row[COL.KARIKATE_TAX_AMOUNT]) || 0,
        taxCode: row[COL.KARIKATE_TAX] as string | number,
        dept: row[COL.KARIKATE_DEPT] as string | number
      });
      kashikatePending.push({
        row: row,
        kamoku: row[COL.KASHIKATE_KAMOKU] as string | number,
        name: row[COL.KASHIKATE_NAME] as string,
        hojo: row[COL.KASHIKATE_HOJO] as string | number,
        hojoName: row[COL.KASHIKATE_HOJO_NAME] as string,
        amount: Number(row[COL.KASHIKATE_AMOUNT]),
        taxAmount: Number(row[COL.KASHIKATE_TAX_AMOUNT]) || 0,
        taxCode: row[COL.KASHIKATE_TAX] as string | number,
        dept: row[COL.KASHIKATE_DEPT] as string | number
      });

    } else if (hasKarikata && !hasKashikata) {
      // 借方のみ
      karikatePending.push({
        row: row,
        kamoku: row[COL.KARIKATE_KAMOKU] as string | number,
        name: row[COL.KARIKATE_NAME] as string,
        hojo: row[COL.KARIKATE_HOJO] as string | number,
        hojoName: row[COL.KARIKATE_HOJO_NAME] as string,
        amount: Number(row[COL.KARIKATE_AMOUNT]),
        taxAmount: Number(row[COL.KARIKATE_TAX_AMOUNT]) || 0,
        taxCode: row[COL.KARIKATE_TAX] as string | number,
        dept: row[COL.KARIKATE_DEPT] as string | number
      });

    } else if (!hasKarikata && hasKashikata) {
      // 貸方のみ
      kashikatePending.push({
        row: row,
        kamoku: row[COL.KASHIKATE_KAMOKU] as string | number,
        name: row[COL.KASHIKATE_NAME] as string,
        hojo: row[COL.KASHIKATE_HOJO] as string | number,
        hojoName: row[COL.KASHIKATE_HOJO_NAME] as string,
        amount: Number(row[COL.KASHIKATE_AMOUNT]),
        taxAmount: Number(row[COL.KASHIKATE_TAX_AMOUNT]) || 0,
        taxCode: row[COL.KASHIKATE_TAX] as string | number,
        dept: row[COL.KASHIKATE_DEPT] as string | number
      });
    }
  }

  // 借方・貸方のリストが両方ある場合のみ処理
  if (karikatePending.length > 0 && kashikatePending.length > 0) {
    const baseRow = rows[0];
    simplified.push({
      baseRow: baseRow,
      karikataItems: karikatePending,
      kashikataItems: kashikatePending
    });
  } else {
    // どちらか片方しかない場合はエラーログに記録してthrow
    const denpyoNo = rows[0][CONFIG.COLUMNS.SOURCE.DENPYO_NO];
    const errorMsg = `借方・貸方のバランスが取れていません（借方:${karikatePending.length}項目, 貸方:${kashikatePending.length}項目）`;

    ERROR_LOGS.push({
      timestamp: new Date(),
      level: 'WARN',
      function: 'simplifyCompoundJournal',
      sourceSheet: CONFIG.SHEETS.SOURCE_DATA,
      denpyoNo: denpyoNo as string | number,
      message: errorMsg,
      stack: ''
    });

    Logger.log(`警告: 伝票${denpyoNo} - ${errorMsg}`);
    throw new Error(errorMsg);
  }

  // 複合仕訳内に335または191の科目コードが含まれているかチェック
  const has335or191 = checkFor335or191(karikatePending, kashikatePending);

  // 金額分割処理
  const splitJournals: SimpleJournal[] = [];
  for (const journal of simplified) {
    const split = splitJournalByAmount(journal, has335or191);
    splitJournals.push(...split);
  }

  return splitJournals;
}

/**
 * 複合仕訳内に335または191の科目コードが含まれているかチェック
 */
function checkFor335or191(karikataItems: JournalItem[], kashikataItems: JournalItem[]): boolean {
  // 借方と貸方の全項目をチェック
  const allItems = [...karikataItems, ...kashikataItems];

  for (const item of allItems) {
    const kamokuStr = String(item.kamoku);
    if (kamokuStr === '335' || kamokuStr === '191') {
      return true;
    }
  }

  return false;
}

/**
 * 金額分割処理（方式D）
 * 借方・貸方の金額が一致するように分割
 */
function splitJournalByAmount(journal: CompoundJournal, has335or191: boolean): SimpleJournal[] {
  const result: SimpleJournal[] = [];

  const karikataItems = [...journal.karikataItems];
  const kashikataItems = [...journal.kashikataItems];

  while (karikataItems.length > 0 && kashikataItems.length > 0) {
    const karikataItem = karikataItems[0];
    const kashikataItem = kashikataItems[0];

    const karikataAmount = karikataItem.amount;
    const kashikataAmount = kashikataItem.amount;

    if (karikataAmount === kashikataAmount) {
      // 金額が一致 → そのまま1仕訳として出力
      result.push({
        baseRow: journal.baseRow,
        karikataItem: karikataItem,
        kashikataItem: kashikataItem,
        amount: karikataAmount,
        has335or191InGroup: has335or191
      });

      karikataItems.shift();
      kashikataItems.shift();

    } else if (karikataAmount < kashikataAmount) {
      // 借方が小さい → 借方金額に合わせて貸方を分割
      result.push({
        baseRow: journal.baseRow,
        karikataItem: karikataItem,
        kashikataItem: {
          ...kashikataItem,
          amount: karikataAmount
        },
        amount: karikataAmount,
        has335or191InGroup: has335or191
      });

      karikataItems.shift();
      kashikataItems[0].amount -= karikataAmount;

      // 貸方が0になったら削除
      if (kashikataItems[0].amount === 0) {
        kashikataItems.shift();
      }

    } else {
      // 貸方が小さい → 貸方金額に合わせて借方を分割
      result.push({
        baseRow: journal.baseRow,
        karikataItem: {
          ...karikataItem,
          amount: kashikataAmount
        },
        kashikataItem: kashikataItem,
        amount: kashikataAmount,
        has335or191InGroup: has335or191
      });

      kashikataItems.shift();
      karikataItems[0].amount -= kashikataAmount;

      // 借方が0になったら削除
      if (karikataItems[0].amount === 0) {
        karikataItems.shift();
      }
    }
  }

  return result;
}

/**
 * 1仕訳を変換
 */
function convertJournal(
  journal: SimpleJournal,
  kamokuMapping: KamokuMapping,
  taxMapping: TaxMapping
): ICSOutputRow {
  const baseRow = journal.baseRow;
  const COL = CONFIG.COLUMNS.SOURCE;
  const denpyoNo = baseRow[COL.DENPYO_NO] as string | number;

  // 日付変換
  const date = formatDate(baseRow[COL.DATE] as string | number);

  // 科目コード変換
  const karikataCode = convertKamokuCode(journal.karikataItem.kamoku, kamokuMapping.codeMap, denpyoNo);
  const kashikataCode = convertKamokuCode(journal.kashikataItem.kamoku, kamokuMapping.codeMap, denpyoNo);

  // 科目名を科目対応表から取得（なければ元データの名称を使用）
  const karikataName = kamokuMapping.nameMap[String(karikataCode)] || journal.karikataItem.name || '';
  const kashikataName = kamokuMapping.nameMap[String(kashikataCode)] || journal.kashikataItem.name || '';

  // 税区分変換
  let taxCode: string;

  // 借方消費税額または貸方消費税額が0より大きい場合は315にする
  if (journal.karikataItem.taxAmount > 0 || journal.kashikataItem.taxAmount > 0) {
    taxCode = '315';
  } else if (journal.has335or191InGroup) {
    // 複合仕訳内に335/191科目が含まれる場合
    const karikataKamokuStr = String(journal.karikataItem.kamoku);
    const kashikataKamokuStr = String(journal.kashikataItem.kamoku);

    // 現在の仕訳が335/191科目の場合は、通常通りの変換
    if (karikataKamokuStr === '335' || karikataKamokuStr === '191' ||
        kashikataKamokuStr === '335' || kashikataKamokuStr === '191') {
      taxCode = selectBestTaxCode(
        journal.karikataItem.taxCode,
        journal.kashikataItem.taxCode,
        taxMapping,
        denpyoNo
      );
    } else {
      // 335/191科目以外の場合は、変換後の税区分コードを311にする
      taxCode = '311';
    }
  } else {
    // 通常の税区分変換（00以外を優先）
    taxCode = selectBestTaxCode(
      journal.karikataItem.taxCode,
      journal.kashikataItem.taxCode,
      taxMapping,
      denpyoNo
    );
  }

  // 税額（借方消費税額を使用）
  const taxAmount = journal.karikataItem.taxAmount || 0;

  // ICS形式の行を作成
  return [
    date,                                    // 日付
    '',                                      // 決修（空白）
    baseRow[COL.DENPYO_NO] as string | number, // 伝票番号
    journal.karikataItem.dept || '',         // 借方部門コード
    '',                                      // 借方事管区分（空白）
    '',                                      // 借方工事コード（空白）
    karikataCode,                            // 借方コード
    karikataName,                            // 借方名称（科目対応表から）
    journal.karikataItem.hojo || '',         // 借方枝番
    journal.karikataItem.hojoName || '',     // 借方枝番摘要
    '',                                      // 借方枝番カナ（空白）
    journal.kashikataItem.dept || '',        // 貸方部門コード
    '',                                      // 貸方事管区分（空白）
    '',                                      // 貸方工事コード（空白）
    kashikataCode,                           // 貸方コード
    kashikataName,                           // 貸方名称（科目対応表から）
    journal.kashikataItem.hojo || '',        // 貸方枝番
    journal.kashikataItem.hojoName || '',    // 貸方枝番摘要
    '',                                      // 貸方枝番カナ（空白）
    journal.amount,                          // 金額
    (baseRow[COL.TEKIYO] as string) || '',   // 摘要
    taxCode,                                 // 税区分
    '',                                      // 対価（空白）
    '',                                      // 仕入区分（空白）
    '',                                      // 売上業種区分（空白）
    '',                                      // 仕訳区分（空白）
    '',                                      // 特定収入区分（空白）
    '',                                      // ダミー1
    '',                                      // ダミー2
    '',                                      // ダミー3
    '',                                      // 内部取引（空白）
    taxAmount,                               // 税額
    '',                                      // 証憑番号（空白）
    '',                                      // 手形番号（空白）
    '',                                      // 手形期日（空白）
    '',                                      // 付箋番号（空白）
    '',                                      // 付箋コメント（空白）
    '',                                      // 免税事業者等（空白）
    ''                                       // インボイス登録番号（空白）
  ];
}

/**
 * データ全体を変換
 */
function convertData(
  sourceData: unknown[][],
  kamokuMapping: KamokuMapping,
  taxMapping: TaxMapping
): ICSOutputRow[] {
  // 伝票番号でグループ化
  const denpyoGroups = groupByDenpyo(sourceData);

  const convertedRows: ICSOutputRow[] = [];

  // 伝票ごとに処理
  for (const denpyoNo in denpyoGroups) {
    const rows = denpyoGroups[denpyoNo];

    try {
      // 複合仕訳を単純仕訳に分解
      const simplifiedJournals = simplifyCompoundJournal(rows);

      // 各仕訳を変換
      for (const journal of simplifiedJournals) {
        const converted = convertJournal(journal, kamokuMapping, taxMapping);
        convertedRows.push(converted);
      }

    } catch (error) {
      const err = error as Error;
      ERROR_LOGS.push({
        timestamp: new Date(),
        level: 'WARN',
        function: 'convertData',
        sourceSheet: CONFIG.SHEETS.SOURCE_DATA,
        denpyoNo: denpyoNo,
        message: err.message,
        stack: err.stack || ''
      });
      Logger.log(`伝票${denpyoNo}の変換エラー: ${err.message}`);
    }
  }

  if (ERROR_LOGS.length > 0) {
    Logger.log(`警告: ${ERROR_LOGS.length}件のエラーが発生しました`);
  }

  return convertedRows;
}
