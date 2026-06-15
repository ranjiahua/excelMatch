/** Parsed representation of an Excel workbook. */
export interface WorkbookData {
  /** File name. */
  fileName: string;
  /** All sheet names in the workbook. */
  sheetNames: string[];
  /** Parsed sheet data keyed by sheet name. */
  sheets: Record<string, SheetData>;
  /** Original file buffer — used for format-preserving export. */
  originalBuffer: ArrayBuffer;
}

/** Parsed representation of a single Excel sheet. */
export interface SheetData {
  /** Sheet name. */
  name: string;
  /** Column letters: ["A", "B", "C", ...]. Index = column index. */
  columns: string[];
  /**
   * All rows including title rows, header rows, data rows, and footer rows.
   * null = empty cell. Fully empty rows are excluded.
   */
  rows: (string | null)[][];
  /** Total number of rows (excluding fully empty rows). */
  rowCount: number;
  colCount: number;
  /**
   * The 0-based row index within `rows` that contains column headers.
   * This is the row used for column labels in dropdowns.
   * Default 0. User-adjustable in the UI.
   */
  headerRowIndex: number;
  /**
   * The 0-based row index within `rows` where actual data starts.
   * Default = headerRowIndex + 1. User-adjustable in the UI.
   */
  dataStartRow: number;
  /**
   * Maps each row in `rows` (0-based) back to its 1-based row number
   * in the original Excel sheet. Used for format-preserving export.
   */
  originalRowIndex: number[];
}

/** Match strategy when looking up values in the source. */
export type MatchStrategy = "exact" | "ignore-case" | "trim";

/** Behaviour when the same key appears multiple times in the source. */
export type DuplicateStrategy = "first" | "last";

/** A single column-to-column mapping under one lookup rule. */
export interface ColumnMapping {
  /** Source column letter, e.g. "C". */
  sourceColumn: string;
  /** Target column letter, e.g. "D". */
  targetColumn: string;
}

/** One mapping rule: match on a key column, then copy N columns. */
export interface MappingRule {
  /** Unique identifier. */
  id: string;
  /** Target column letter used as the lookup key. */
  targetKeyColumn: string;
  /** Source column letter to match against. */
  sourceKeyColumn: string;
  /** How values are compared. */
  matchStrategy: MatchStrategy;
  /** What to do on duplicate source keys. */
  duplicateStrategy: DuplicateStrategy;
  /** Column mappings applied when a match is found. */
  mappings: ColumnMapping[];
}

/** Persisted mapping configuration (exported/imported as JSON). */
export interface MappingConfig {
  version: 1;
  targetSheet: string;
  sourceSheet: string;
  rules: MappingRule[];
}

/** Result after applying mappings to target data. */
export interface MappingResult {
  /** Column letters (inherited from target). */
  columns: string[];
  /** Data rows after applying all mapping rules. */
  rows: (string | null)[][];
  /** Per-row match status. */
  matchStatus: ("matched" | "unmatched")[];
  /** Statistics. */
  stats: {
    totalRows: number;
    matchedRows: number;
    unmatchedRows: number;
  };
}
