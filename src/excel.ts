import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
import type {
  WorkbookData,
  SheetData,
  MappingRule,
  MappingResult,
  MatchStrategy,
} from "./types";

// ──── Column reference helpers ────

/** Convert 0-based column index to letter(s): 0 → "A", 25 → "Z", 26 → "AA". */
export function colIndexToLetter(index: number): string {
  let result = "";
  let n = index;
  while (n >= 0) {
    result = String.fromCharCode(65 + (n % 26)) + result;
    n = Math.floor(n / 26) - 1;
  }
  return result;
}

/** Convert column letter(s) to 0-based index: "A" → 0, "Z" → 25, "AA" → 26. */
export function colLetterToIndex(letter: string): number {
  let result = 0;
  for (let i = 0; i < letter.length; i++) {
    result = result * 26 + (letter.charCodeAt(i) - 64);
  }
  return result - 1;
}

// ──── Match helpers ────

/** Normalize a cell value according to the match strategy. */
function normalizeValue(val: string | null, strategy: MatchStrategy): string {
  if (val === null || val === undefined) return "";
  let s = String(val);
  if (strategy === "ignore-case" || strategy === "trim") {
    s = s.trim();
  }
  if (strategy === "ignore-case") {
    s = s.toLowerCase();
  }
  return s;
}

/** Check if a row is fully empty (all cells null or empty string). */
function isRowEmpty(row: (string | null)[]): boolean {
  return row.every((cell) => cell === null || cell === "");
}

// ──── Header detection ────

/**
 * Auto-detect the header row index (0-based) within the rows array.
 * Strategy: find the first row with at least 3 non-empty cells.
 */
export function detectHeaderRow(rows: (string | null)[][]): number {
  for (let i = 0; i < rows.length; i++) {
    const filledCount = rows[i].filter(
      (c) => c !== null && c !== "" && c !== undefined
    ).length;
    if (filledCount >= 3) {
      return i;
    }
  }
  return 0;
}

/**
 * Auto-detect the data start row (0-based) within the rows array.
 * Strategy: after the header row, find the first row where column A is an integer.
 */
export function detectDataStartRow(
  rows: (string | null)[][],
  headerRow: number
): number {
  for (let i = headerRow + 1; i < rows.length; i++) {
    const firstCell = rows[i]?.[0];
    if (firstCell !== null && firstCell !== "" && firstCell !== undefined) {
      if (/^\d+$/.test(String(firstCell).trim())) {
        return i;
      }
    }
  }
  return headerRow + 1;
}

// ──── Excel file parsing ────

/** Parse an uploaded Excel File into WorkbookData (all sheets). */
export function parseExcelFile(file: File): Promise<WorkbookData> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const buffer = e.target!.result as ArrayBuffer;
        const data = new Uint8Array(buffer);
        const wb = XLSX.read(data, { type: "array" });
        const sheetNames = wb.SheetNames;
        const sheets: Record<string, SheetData> = {};

        for (const name of sheetNames) {
          const sheet = wb.Sheets[name];
          const raw: (string | null)[][] = XLSX.utils.sheet_to_json(sheet, {
            header: 1,
            defval: null,
          });

          if (raw.length === 0) {
            sheets[name] = {
              name,
              columns: [],
              rows: [],
              rowCount: 0,
              colCount: 0,
              headerRowIndex: 0,
              dataStartRow: 0,
              originalRowIndex: [],
            };
            continue;
          }

          const maxCols = Math.max(
            ...raw.map((row) => (Array.isArray(row) ? row.length : 0))
          );

          // Normalize rows, filter out fully empty ones.
          // Track original 1-based row number for each surviving row.
          const allRows: (string | null)[][] = [];
          const originalRowIndex: number[] = [];

          for (let r = 0; r < raw.length; r++) {
            const row = raw[r];
            const arr = Array.isArray(row) ? row : [row];
            const normalized = Array.from({ length: maxCols }, (_, i) => {
              const cell = arr[i];
              if (cell === null || cell === undefined) return null;
              return String(cell);
            });
            if (!isRowEmpty(normalized)) {
              allRows.push(normalized);
              originalRowIndex.push(r + 1); // 1-based Excel row number
            }
          }

          const columns = Array.from({ length: maxCols }, (_, i) =>
            colIndexToLetter(i)
          );

          const headerRowIndex = detectHeaderRow(allRows);
          const dataStartRow = detectDataStartRow(allRows, headerRowIndex);

          sheets[name] = {
            name,
            columns,
            rows: allRows,
            rowCount: allRows.length,
            colCount: maxCols,
            headerRowIndex,
            dataStartRow,
            originalRowIndex,
          };
        }

        resolve({ fileName: file.name, sheetNames, sheets, originalBuffer: buffer });
      } catch (err) {
        reject(
          err instanceof Error ? err : new Error("Failed to parse Excel file.")
        );
      }
    };
    reader.onerror = () => reject(new Error("Failed to read file."));
    reader.readAsArrayBuffer(file);
  });
}

// ──── Mapping engine ────

/**
 * Build a lookup index for the source sheet.
 * Only indexes rows from dataStartRow onwards.
 */
function buildLookupIndex(
  sourceRows: (string | null)[][],
  sourceKeyCol: number,
  strategy: MatchStrategy,
  dataStartRow: number
): Map<string, number[]> {
  const index = new Map<string, number[]>();
  for (let i = dataStartRow; i < sourceRows.length; i++) {
    const rawVal = sourceRows[i]?.[sourceKeyCol] ?? null;
    const key = normalizeValue(rawVal, strategy);
    if (!index.has(key)) {
      index.set(key, []);
    }
    index.get(key)!.push(i);
  }
  return index;
}

/** Apply mapping rules to the target sheet using the source sheet. */
export function applyMappings(
  targetSheet: SheetData,
  sourceSheet: SheetData,
  rules: MappingRule[]
): MappingResult {
  // Deep copy target rows
  const rows: (string | null)[][] = targetSheet.rows.map((row) => [...row]);
  const matchStatus: ("matched" | "unmatched")[] = new Array(rows.length).fill(
    "unmatched"
  );

  // Mark all rows before dataStartRow as "matched" (headers, titles etc. are not data)
  const dataStart = targetSheet.dataStartRow;
  for (let i = 0; i < dataStart && i < rows.length; i++) {
    matchStatus[i] = "matched";
  }

  for (const rule of rules) {
    const tKeyCol = colLetterToIndex(rule.targetKeyColumn);
    const sKeyCol = colLetterToIndex(rule.sourceKeyColumn);

    if (
      tKeyCol >= targetSheet.colCount ||
      sKeyCol >= sourceSheet.colCount
    ) {
      console.warn(
        `Skipping rule ${rule.id}: key column out of range ` +
          `(target=${rule.targetKeyColumn}, source=${rule.sourceKeyColumn})`
      );
      continue;
    }

    const lookup = buildLookupIndex(
      sourceSheet.rows,
      sKeyCol,
      rule.matchStrategy,
      sourceSheet.dataStartRow
    );

    const mappingIndices = rule.mappings
      .map((m) => ({
        sCol: colLetterToIndex(m.sourceColumn),
        tCol: colLetterToIndex(m.targetColumn),
      }))
      .filter(({ sCol, tCol }) => {
        const valid =
          sCol < sourceSheet.colCount && tCol < targetSheet.colCount;
        if (!valid) {
          console.warn(
            `Skipping mapping in rule ${rule.id}: column out of range`
          );
        }
        return valid;
      });

    for (let tRow = dataStart; tRow < rows.length; tRow++) {
      const lookupValue = rows[tRow]?.[tKeyCol] ?? null;
      const normalized = normalizeValue(lookupValue, rule.matchStrategy);

      const matchRows = lookup.get(normalized);
      if (!matchRows || matchRows.length === 0) continue;

      const sRow =
        rule.duplicateStrategy === "last"
          ? matchRows[matchRows.length - 1]
          : matchRows[0];

      for (const { sCol, tCol } of mappingIndices) {
        const srcVal =
          sCol < sourceSheet.colCount && sRow < sourceSheet.rows.length
            ? sourceSheet.rows[sRow]?.[sCol] ?? null
            : null;
        rows[tRow][tCol] = srcVal;
      }

      matchStatus[tRow] = "matched";
    }
  }

  const matchedRows = matchStatus.filter((s) => s === "matched").length;
  const dataRows = Math.max(0, rows.length - dataStart);
  const dataMatched = Math.max(0, matchedRows - dataStart);

  return {
    columns: targetSheet.columns,
    rows,
    matchStatus,
    stats: {
      totalRows: dataRows,
      matchedRows: dataMatched,
      unmatchedRows: dataRows - dataMatched,
    },
  };
}

// ──── Format-preserving Excel export ────

/**
 * Generate an Excel file from the original target workbook, applying only
 * the changed cell values from the mapping result.
 *
 * Uses exceljs (not xlsx) to load the original buffer → modify cell values →
 * write back. exceljs preserves ALL formatting: fonts, colors, borders,
 * merged cells, row heights, column widths, number formats, alignments, etc.
 *
 * @param targetWb  The parsed target workbook (contains originalBuffer).
 * @param sheetName The sheet name within the workbook.
 * @param result    The mapping result from applyMappings().
 * @param targetSheet The original target SheetData (for row-index mapping).
 */
export async function generateExcelBlob(
  originalBuffer: ArrayBuffer,
  sheetName: string,
  result: MappingResult,
  targetSheet: SheetData,
  allSheetNames: string[]
): Promise<Uint8Array> {
  // Strategy: use xlsx (SheetJS) for round-trip — it already reads the file
  // successfully during parsing. Read original buffer → modify cell values
  // in-place on the parsed sheet → write back. This preserves the original
  // cell objects (which carry partial formatting info) better than creating
  // a new workbook from scratch.
  //
  // Also try exceljs first (better formatting fidelity); fall back to xlsx
  // if exceljs can't find the sheet.

  // --- Attempt 1: exceljs (best formatting preservation) ---
  try {
    const eb = new ExcelJS.Workbook();
    await eb.xlsx.load(originalBuffer);

    // Find sheet by name (exceljs's own names) or by index
    const excelSheetNames = eb.worksheets.map((ws) => ws.name);
    const eWs =
      eb.getWorksheet(sheetName) ||
      eb.worksheets[allSheetNames.indexOf(sheetName)] ||
      eb.worksheets[0];

    if (eWs) {
      const originalRowMap = targetSheet.originalRowIndex;
      for (let r = 0; r < result.rows.length; r++) {
        const originalRow = originalRowMap[r];
        if (originalRow === undefined) continue;
        const excelRow = eWs.getRow(originalRow);
        for (let c = 0; c < result.rows[r].length; c++) {
          const newVal = result.rows[r][c];
          if (newVal !== null && newVal !== "") {
            excelRow.getCell(c + 1).value = newVal;
          }
        }
      }
      const outBuffer = await eb.xlsx.writeBuffer();
      return new Uint8Array(outBuffer);
    }
    console.warn(
      `exceljs: sheet "${sheetName}" not found (available: ${excelSheetNames.join(", ")}). Falling back to xlsx.`
    );
  } catch (exceljsErr) {
    console.warn("exceljs load failed, falling back to xlsx:", exceljsErr);
  }

  // --- Attempt 2: xlsx (SheetJS) fallback ---
  const data = new Uint8Array(originalBuffer);
  const sWb = XLSX.read(data, { type: "array" });
  // Use the actual sheet name from the xlsx workbook, not our stored name
  const actualSheetName = sWb.SheetNames[0] || sheetName;
  const sheet = sWb.Sheets[actualSheetName];
  if (!sheet) {
    throw new Error(
      `Sheet not found. Tried "${sheetName}" and "${actualSheetName}". ` +
      `Available: ${sWb.SheetNames.join(", ")}`
    );
  }

  const originalRowMap = targetSheet.originalRowIndex;
  for (let r = 0; r < result.rows.length; r++) {
    const originalRow = originalRowMap[r];
    if (originalRow === undefined) continue;
    for (let c = 0; c < result.rows[r].length; c++) {
      const newVal = result.rows[r][c];
      if (newVal !== null && newVal !== "") {
        const cellAddr = XLSX.utils.encode_cell({ r: originalRow - 1, c });
        const oldCell = sheet[cellAddr];
        if (oldCell && typeof oldCell === "object") {
          oldCell.v = newVal;
        } else if (oldCell !== undefined) {
          sheet[cellAddr] = { v: newVal };
        } else {
          sheet[cellAddr] = { t: "s", v: newVal };
        }
      }
    }
  }

  return XLSX.write(sWb, { type: "array", bookType: "xlsx" });
}
