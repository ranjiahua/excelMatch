import { useState, useMemo, useRef } from "react";
import type {
  WorkbookData,
  SheetData,
  MappingRule,
  MappingResult,
} from "./types";
import { applyMappings } from "./excel";
import FileUploader from "./components/FileUploader";
import SheetPreview from "./components/SheetPreview";
import MappingTable from "./components/MappingTable";
import ResultPreview from "./components/ResultPreview";
import "./App.css";

function App() {
  const [targetWorkbook, setTargetWorkbook] = useState<WorkbookData | null>(
    null
  );
  const [sourceWorkbook, setSourceWorkbook] = useState<WorkbookData | null>(
    null
  );
  const [targetSheetName, setTargetSheetName] = useState("");
  const [sourceSheetName, setSourceSheetName] = useState("");
  const [rules, setRules] = useState<MappingRule[]>([]);
  const [result, setResult] = useState<MappingResult | null>(null);
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  // Store original File object — re-read when saving for a fresh ArrayBuffer
  const targetFileRef = useRef<File | null>(null);

  const targetSheet: SheetData | null = useMemo(() => {
    if (!targetWorkbook) return null;
    const name = targetSheetName || targetWorkbook.sheetNames[0] || "";
    return targetWorkbook.sheets[name] ?? null;
  }, [targetWorkbook, targetSheetName]);

  const sourceSheet: SheetData | null = useMemo(() => {
    if (!sourceWorkbook) return null;
    const name = sourceSheetName || sourceWorkbook.sheetNames[0] || "";
    return sourceWorkbook.sheets[name] ?? null;
  }, [sourceWorkbook, sourceSheetName]);

  const bothLoaded = targetWorkbook !== null && sourceWorkbook !== null;

  function handleGenerate() {
    if (!targetSheet || !sourceSheet) return;
    setError("");
    try {
      const res = applyMappings(targetSheet, sourceSheet, rules);
      setResult(res);
    } catch (err) {
      setError(
        err instanceof Error ? err.message : "生成结果失败"
      );
    }
  }

  function handleClearAll() {
    if (!confirm("确认清除所有数据？\n将清除已上传的文件、映射规则和生成结果。")) {
      return;
    }
    setTargetWorkbook(null);
    setSourceWorkbook(null);
    targetFileRef.current = null;
    setTargetSheetName("");
    setSourceSheetName("");
    setRules([]);
    setResult(null);
    setError("");
    setLoading(false);
  }

  /** Update headerRowIndex / dataStartRow for a sheet in a workbook. */
  function updateSheetConfig(
    wb: WorkbookData | null,
    sheetName: string,
    field: "headerRowIndex" | "dataStartRow",
    value: number,
    setWb: (wb: WorkbookData | null) => void
  ) {
    if (!wb) return;
    const sheet = wb.sheets[sheetName];
    if (!sheet) return;
    const updated = { ...sheet, [field]: value };
    setWb({
      ...wb,
      sheets: { ...wb.sheets, [sheetName]: updated },
    });
    setResult(null);
  }

  return (
    <main className="container">
      <div className="header-row">
        <h1>Excel 映射工具</h1>
        <button className="secondary" onClick={handleClearAll}>
          一键清除
        </button>
      </div>

      {/* Error banner */}
      {error && (
        <div className="error-banner" onClick={() => setError("")}>
          {error}
          <span className="close-btn">&times;</span>
        </div>
      )}

      {/* Section 1: File Upload */}
      <section className="section">
        <h2>1. 上传文件</h2>
        <div className="upload-row">
          <FileUploader
            label="目标文件（待填充）"
            onParsed={(wb, file) => {
              setTargetWorkbook(wb);
              targetFileRef.current = file;
              setTargetSheetName(wb.sheetNames[0] ?? "");
              setResult(null);
            }}
            onError={setError}
            onLoading={setLoading}
          />
          <FileUploader
            label="源文件（数据来源）"
            onParsed={(wb) => {
              setSourceWorkbook(wb);
              setSourceSheetName(wb.sheetNames[0] ?? "");
              setResult(null);
            }}
            onError={setError}
            onLoading={setLoading}
          />
        </div>
        {loading && <p className="loading-msg">正在解析文件...</p>}

        {targetWorkbook && (
          <div className="file-info-block">
            <p className="file-info">
              目标文件：<strong>{targetWorkbook.fileName}</strong>
            </p>
            <div className="sheet-config-row">
              <div className="sheet-selector">
                <label>工作表：</label>
                <select
                  value={targetSheetName}
                  onChange={(e) => {
                    setTargetSheetName(e.target.value);
                    setResult(null);
                  }}
                >
                  {targetWorkbook.sheetNames.map((n) => (
                    <option key={n} value={n}>
                      {n}
                    </option>
                  ))}
                </select>
                {targetSheet && (
                  <span className="sheet-meta">
                    （{targetSheet.rowCount} 行 × {targetSheet.colCount} 列）
                  </span>
                )}
              </div>
              {targetSheet && (
                <div className="row-config">
                  <label>表头行：</label>
                  <input
                    type="number"
                    min={1}
                    max={targetSheet.rowCount}
                    value={targetSheet.headerRowIndex + 1}
                    onChange={(e) =>
                      updateSheetConfig(
                        targetWorkbook,
                        targetSheetName,
                        "headerRowIndex",
                        Math.max(0, parseInt(e.target.value) - 1 || 0),
                        setTargetWorkbook
                      )
                    }
                    className="row-num-input"
                  />
                  <label>数据起始行：</label>
                  <input
                    type="number"
                    min={1}
                    max={targetSheet.rowCount}
                    value={targetSheet.dataStartRow + 1}
                    onChange={(e) =>
                      updateSheetConfig(
                        targetWorkbook,
                        targetSheetName,
                        "dataStartRow",
                        Math.max(0, parseInt(e.target.value) - 1 || 0),
                        setTargetWorkbook
                      )
                    }
                    className="row-num-input"
                  />
                </div>
              )}
            </div>
          </div>
        )}
        {sourceWorkbook && (
          <div className="file-info-block">
            <p className="file-info">
              源文件：<strong>{sourceWorkbook.fileName}</strong>
            </p>
            <div className="sheet-config-row">
              <div className="sheet-selector">
                <label>工作表：</label>
                <select
                  value={sourceSheetName}
                  onChange={(e) => {
                    setSourceSheetName(e.target.value);
                    setResult(null);
                  }}
                >
                  {sourceWorkbook.sheetNames.map((n) => (
                    <option key={n} value={n}>
                      {n}
                    </option>
                  ))}
                </select>
                {sourceSheet && (
                  <span className="sheet-meta">
                    （{sourceSheet.rowCount} 行 × {sourceSheet.colCount} 列）
                  </span>
                )}
              </div>
              {sourceSheet && (
                <div className="row-config">
                  <label>表头行：</label>
                  <input
                    type="number"
                    min={1}
                    max={sourceSheet.rowCount}
                    value={sourceSheet.headerRowIndex + 1}
                    onChange={(e) =>
                      updateSheetConfig(
                        sourceWorkbook,
                        sourceSheetName,
                        "headerRowIndex",
                        Math.max(0, parseInt(e.target.value) - 1 || 0),
                        setSourceWorkbook
                      )
                    }
                    className="row-num-input"
                  />
                  <label>数据起始行：</label>
                  <input
                    type="number"
                    min={1}
                    max={sourceSheet.rowCount}
                    value={sourceSheet.dataStartRow + 1}
                    onChange={(e) =>
                      updateSheetConfig(
                        sourceWorkbook,
                        sourceSheetName,
                        "dataStartRow",
                        Math.max(0, parseInt(e.target.value) - 1 || 0),
                        setSourceWorkbook
                      )
                    }
                    className="row-num-input"
                  />
                </div>
              )}
            </div>
          </div>
        )}
      </section>

      {/* Section 2: Preview */}
      {bothLoaded && targetSheet && sourceSheet && (
        <section className="section">
          <h2>2. 数据预览</h2>
          <div className="preview-row">
            <SheetPreview
              data={targetSheet}
              title={`目标：${targetSheet.name}`}
            />
            <SheetPreview
              data={sourceSheet}
              title={`源：${sourceSheet.name}`}
            />
          </div>
        </section>
      )}

      {/* Section 3: Mapping Rules */}
      {bothLoaded && (
        <section className="section">
          <h2>3. 映射规则</h2>
          <MappingTable
            rules={rules}
            onChange={setRules}
            targetSheet={targetSheet}
            sourceSheet={sourceSheet}
          />
          <div className="generate-area">
            <button
              className="primary"
              onClick={handleGenerate}
              disabled={rules.length === 0}
            >
              生成结果
            </button>
          </div>
        </section>
      )}

      {/* Section 4: Result */}
      {result && (
        <section className="section">
          <ResultPreview
            result={result}
            targetWorkbook={targetWorkbook}
            targetSheet={targetSheet}
            targetSheetName={targetSheetName}
            targetFile={targetFileRef.current}
          />
        </section>
      )}
    </main>
  );
}

export default App;
