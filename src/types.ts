/**
 * PCA2ICS 型定義
 */

/** 設定オブジェクトの型 */
interface Config {
  DEFAULT_SOURCE_SHEET: string;
  SHEETS: {
    SOURCE_DATA: string;
    KAMOKU_MAPPING: string;
    TAX_MAPPING: string;
    OUTPUT: string;
    ERROR_LOG: string;
  };
  COLUMNS: {
    SOURCE: SourceColumns;
  };
}

/** 元データのカラムインデックス */
interface SourceColumns {
  DATE: number;
  DENPYO_NO: number;
  KARIKATE_DEPT: number;
  KARIKATE_KAMOKU: number;
  KARIKATE_NAME: number;
  KARIKATE_HOJO: number;
  KARIKATE_HOJO_NAME: number;
  KARIKATE_TAX: number;
  KARIKATE_AMOUNT: number;
  KARIKATE_TAX_AMOUNT: number;
  KASHIKATE_DEPT: number;
  KASHIKATE_KAMOKU: number;
  KASHIKATE_NAME: number;
  KASHIKATE_HOJO: number;
  KASHIKATE_HOJO_NAME: number;
  KASHIKATE_TAX: number;
  KASHIKATE_AMOUNT: number;
  KASHIKATE_TAX_AMOUNT: number;
  TEKIYO: number;
}

/** 科目マッピングの型 */
interface KamokuMapping {
  codeMap: { [key: string]: string | number };
  nameMap: { [key: string]: string };
}

/** 税区分マッピングの型 */
interface TaxMapping {
  [key: string]: string;
}

/** エラーログエントリの型 */
interface ErrorLogEntry {
  timestamp: Date;
  level: 'ERROR' | 'WARN' | 'INFO';
  function: string;
  sourceSheet: string;
  denpyoNo: string | number;
  message: string;
  stack: string;
}

/** 仕訳項目（借方/貸方）の型 */
interface JournalItem {
  row: unknown[];
  kamoku: string | number;
  name: string;
  hojo: string | number;
  hojoName: string;
  amount: number;
  taxAmount: number;
  taxCode: string | number;
  dept: string | number;
}

/** 複合仕訳の型 */
interface CompoundJournal {
  baseRow: unknown[];
  karikataItems: JournalItem[];
  kashikataItems: JournalItem[];
}

/** 分割後の単純仕訳の型 */
interface SimpleJournal {
  baseRow: unknown[];
  karikataItem: JournalItem;
  kashikataItem: JournalItem;
  amount: number;
  has335or191InGroup: boolean; // 複合仕訳内に335/191科目を含むか
}

/** 伝票グループの型 */
interface DenpyoGroups {
  [denpyoNo: string]: unknown[][];
}

/** ICS出力行の型 */
type ICSOutputRow = (string | number)[];
