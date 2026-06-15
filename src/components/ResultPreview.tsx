import { writeFile } from "@tauri-apps/plugin-fs";
import { save } from "@tauri-apps/plugin-dialog";
import { useState } from "react";
import type { WorkbookData, SheetData, MappingResult } from "../types";
import { generateExcelBlob } from "../excel";
import SheetPreview from "./SheetPreview";

interface ResultPreviewProps {
  result: MappingResult | null;
  targetWorkbook: WorkbookData | null;
  targetSheet: SheetData | null;
  targetSheetName: string;
  targetFile: File | null;
}

/** Read a File as ArrayBuffer (returns a fresh buffer each time). */
function readFileAsBuffer(file: File): Promise<ArrayBuffer> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as ArrayBuffer);
    reader.onerror = () => reject(new Error("Failed to read file"));
    reader.readAsArrayBuffer(file);
  });
}

export default function ResultPreview({
  result,
  targetWorkbook,
  targetSheet,
  targetSheetName,
  targetFile,
}: ResultPreviewProps) {
  const [saving, setSaving] = useState(false);
  const [successMsg, setSuccessMsg] = useState("");
  const [errorMsg, setErrorMsg] = useState("");

  async function handleSave() {
    if (!result || !targetWorkbook || !targetSheet || !targetFile) return;

    setSaving(true);
    setSuccessMsg("");
    setErrorMsg("");

    try {
      const filePath = await save({
        filters: [{ name: "Excel 文件", extensions: ["xlsx"] }],
        defaultPath: "结果.xlsx",
      });

      if (!filePath) {
        setSaving(false);
        return;
      }

      // Re-read the file fresh — guarantees a valid, uncorrupted ArrayBuffer
      const freshBuffer = await readFileAsBuffer(targetFile);
      const blob = await generateExcelBlob(
        freshBuffer,
        targetSheetName,
        result,
        targetSheet,
        targetWorkbook.sheetNames
      );
      await writeFile(filePath, blob);
      setSuccessMsg(`已保存至：${filePath}`);
    } catch (err) {
      setErrorMsg(err instanceof Error ? err.message : String(err));
    } finally {
      setSaving(false);
    }
  }

  if (!result) return null;

  const previewData = {
    name: "结果",
    columns: result.columns,
    rows: result.rows,
    rowCount: result.rows.length,
    colCount: result.columns.length,
    headerRowIndex: targetSheet?.headerRowIndex ?? 0,
    dataStartRow: targetSheet?.dataStartRow ?? 0,
    originalRowIndex: targetSheet?.originalRowIndex ?? [],
  };

  return (
    <div className="result-section">
      <h2>结果</h2>

      <div className="result-stats">
        <span className="stat-item">
          共 <strong>{result.stats.totalRows}</strong> 行
        </span>
        <span className="stat-item stat-matched">
          已匹配：<strong>{result.stats.matchedRows}</strong>
        </span>
        {result.stats.unmatchedRows > 0 && (
          <span className="stat-item stat-unmatched">
            未匹配：<strong>{result.stats.unmatchedRows}</strong>
          </span>
        )}
      </div>

      <SheetPreview
        data={previewData}
        title=""
        highlight={{ matchStatus: result.matchStatus }}
      />

      <div className="save-area">
        <button className="primary" onClick={handleSave} disabled={saving}>
          {saving ? "保存中..." : "保存为 Excel"}
        </button>
        {successMsg && <p className="success-msg">{successMsg}</p>}
        {errorMsg && <p className="error-msg">{errorMsg}</p>}
      </div>
    </div>
  );
}
