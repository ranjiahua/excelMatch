import { useState, useRef } from "react";
import type {
  MappingRule,
  MappingConfig,
  ColumnMapping,
  SheetData,
  MatchStrategy,
  DuplicateStrategy,
  MappingResult,
} from "../types";
import { applyMappings, colLetterToIndex } from "../excel";
import SheetPreview from "./SheetPreview";

interface MappingTableProps {
  rules: MappingRule[];
  onChange: (rules: MappingRule[]) => void;
  targetSheet: SheetData | null;
  sourceSheet: SheetData | null;
}

function columnLabel(col: string, sheet: SheetData | null): string {
  if (!sheet) return col;
  const idx = colLetterToIndex(col);
  const headerRow = sheet.headerRowIndex;
  const header =
    sheet.rows.length > headerRow ? sheet.rows[headerRow]?.[idx] : null;
  if (header && header.trim()) {
    return `${col}（${header.trim()}）`;
  }
  return col;
}

const MATCH_STRATEGY_LABELS: Record<MatchStrategy, string> = {
  exact: "精确匹配",
  "ignore-case": "忽略大小写",
  trim: "去除空格",
};

const DUPLICATE_STRATEGY_LABELS: Record<DuplicateStrategy, string> = {
  first: "取第一个",
  last: "取最后一个",
};

export default function MappingTable({
  rules,
  onChange,
  targetSheet,
  sourceSheet,
}: MappingTableProps) {
  const [matchPreview, setMatchPreview] = useState<MappingResult | null>(null);
  const [showPreview, setShowPreview] = useState(false);
  const importRef = useRef<HTMLInputElement>(null);

  const targetCols = targetSheet?.columns ?? [];
  const sourceCols = sourceSheet?.columns ?? [];

  function addRule() {
    const newRule: MappingRule = {
      id: crypto.randomUUID(),
      targetKeyColumn: targetCols[0] ?? "",
      sourceKeyColumn: sourceCols[0] ?? "",
      matchStrategy: "exact",
      duplicateStrategy: "first",
      mappings: [],
    };
    onChange([...rules, newRule]);
  }

  function updateRule(
    id: string,
    field: keyof MappingRule,
    value: string
  ) {
    const updated = rules.map((r) => {
      if (r.id !== id) return r;
      if (field === "matchStrategy") {
        return { ...r, matchStrategy: value as MatchStrategy };
      }
      if (field === "duplicateStrategy") {
        return { ...r, duplicateStrategy: value as DuplicateStrategy };
      }
      return { ...r, [field]: value };
    });
    onChange(updated);
    setMatchPreview(null);
    setShowPreview(false);
  }

  function deleteRule(id: string) {
    onChange(rules.filter((r) => r.id !== id));
    setMatchPreview(null);
    setShowPreview(false);
  }

  function addMapping(ruleId: string) {
    const updated = rules.map((r) => {
      if (r.id !== ruleId) return r;
      return {
        ...r,
        mappings: [
          ...r.mappings,
          { sourceColumn: sourceCols[0] ?? "", targetColumn: targetCols[0] ?? "" },
        ],
      };
    });
    onChange(updated);
    setMatchPreview(null);
    setShowPreview(false);
  }

  function updateMapping(
    ruleId: string,
    index: number,
    field: keyof ColumnMapping,
    value: string
  ) {
    const updated = rules.map((r) => {
      if (r.id !== ruleId) return r;
      const newMappings = r.mappings.map((m, i) =>
        i === index ? { ...m, [field]: value } : m
      );
      return { ...r, mappings: newMappings };
    });
    onChange(updated);
    setMatchPreview(null);
    setShowPreview(false);
  }

  function deleteMapping(ruleId: string, index: number) {
    const updated = rules.map((r) => {
      if (r.id !== ruleId) return r;
      return { ...r, mappings: r.mappings.filter((_, i) => i !== index) };
    });
    onChange(updated);
    setMatchPreview(null);
    setShowPreview(false);
  }

  function handlePreview() {
    if (!targetSheet || !sourceSheet || rules.length === 0) return;
    try {
      const result = applyMappings(targetSheet, sourceSheet, rules);
      setMatchPreview(result);
      setShowPreview(true);
    } catch {
      // ignore
    }
  }

  function exportJSON() {
    const config: MappingConfig = {
      version: 1,
      targetSheet: targetSheet?.name ?? "",
      sourceSheet: sourceSheet?.name ?? "",
      rules,
    };
    const blob = new Blob([JSON.stringify(config, null, 2)], {
      type: "application/json",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "mapping-config.json";
    a.click();
    URL.revokeObjectURL(url);
  }

  function handleImport(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const json = JSON.parse(reader.result as string);
        if (json.version !== 1 || !Array.isArray(json.rules)) {
          alert("无效的映射配置文件：需要 version 1 且包含 rules 数组");
          return;
        }
        const imported = json.rules as MappingRule[];
        for (const r of imported) {
          if (
            !r.id ||
            !r.targetKeyColumn ||
            !r.sourceKeyColumn ||
            !Array.isArray(r.mappings)
          ) {
            alert(
              "无效的映射配置：每条规则必须包含 id、targetKeyColumn、sourceKeyColumn 和 mappings"
            );
            return;
          }
        }
        onChange(imported);
        setMatchPreview(null);
        setShowPreview(false);
      } catch {
        alert("无法解析 JSON 文件，请检查文件格式");
      }
    };
    reader.readAsText(file);
    e.target.value = "";
  }

  return (
    <div className="mapping-table">
      <div className="mapping-toolbar">
        <button onClick={addRule}>+ 添加规则</button>
        <button
          className="secondary"
          onClick={exportJSON}
          disabled={rules.length === 0}
        >
          导出 JSON
        </button>
        <button
          className="secondary"
          onClick={() => importRef.current?.click()}
        >
          导入 JSON
        </button>
        <input
          ref={importRef}
          type="file"
          accept=".json"
          style={{ display: "none" }}
          onChange={handleImport}
        />
      </div>

      {rules.length === 0 ? (
        <p className="empty-state">
          暂无映射规则，点击"+ 添加规则"创建
        </p>
      ) : (
        <>
          {rules.map((rule, i) => (
            <div key={rule.id} className="rule-card">
              <div className="rule-card-header">
                <span className="rule-title">规则 #{i + 1}</span>
                <button
                  className="danger"
                  onClick={() => deleteRule(rule.id)}
                >
                  删除
                </button>
              </div>

              {/* Lookup key config */}
              <div className="rule-key-config">
                <div className="key-config-label">关联键</div>
                <div className="key-config-row">
                  <div className="key-config-item">
                    <label>目标列：</label>
                    <select
                      value={rule.targetKeyColumn}
                      onChange={(e) =>
                        updateRule(rule.id, "targetKeyColumn", e.target.value)
                      }
                    >
                      {targetCols.map((col) => (
                        <option key={col} value={col}>
                          {columnLabel(col, targetSheet)}
                        </option>
                      ))}
                    </select>
                  </div>
                  <span className="key-equals">=</span>
                  <div className="key-config-item">
                    <label>源列：</label>
                    <select
                      value={rule.sourceKeyColumn}
                      onChange={(e) =>
                        updateRule(rule.id, "sourceKeyColumn", e.target.value)
                      }
                    >
                      {sourceCols.map((col) => (
                        <option key={col} value={col}>
                          {columnLabel(col, sourceSheet)}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>
              </div>

              {/* Match options */}
              <div className="rule-options">
                <div className="option-item">
                  <label>匹配方式：</label>
                  <select
                    value={rule.matchStrategy}
                    onChange={(e) =>
                      updateRule(rule.id, "matchStrategy", e.target.value)
                    }
                  >
                    {(
                      Object.entries(MATCH_STRATEGY_LABELS) as [
                        MatchStrategy,
                        string
                      ][]
                    ).map(([value, label]) => (
                      <option key={value} value={value}>
                        {label}
                      </option>
                    ))}
                  </select>
                </div>
                <div className="option-item">
                  <label>重复处理：</label>
                  <select
                    value={rule.duplicateStrategy}
                    onChange={(e) =>
                      updateRule(rule.id, "duplicateStrategy", e.target.value)
                    }
                  >
                    {(
                      Object.entries(DUPLICATE_STRATEGY_LABELS) as [
                        DuplicateStrategy,
                        string
                      ][]
                    ).map(([value, label]) => (
                      <option key={value} value={value}>
                        {label}
                      </option>
                    ))}
                  </select>
                </div>
              </div>

              {/* Column mappings */}
              <div className="rule-mappings">
                <div className="mappings-label">列映射</div>
                {rule.mappings.length === 0 && (
                  <p className="empty-state">
                    暂无列映射，点击下方"+ 添加映射"
                  </p>
                )}
                {rule.mappings.map((m, mi) => (
                  <div key={mi} className="mapping-row">
                    <span className="mapping-arrow-label">源列</span>
                    <select
                      value={m.sourceColumn}
                      onChange={(e) =>
                        updateMapping(
                          rule.id,
                          mi,
                          "sourceColumn",
                          e.target.value
                        )
                      }
                    >
                      {sourceCols.map((col) => (
                        <option key={col} value={col}>
                          {columnLabel(col, sourceSheet)}
                        </option>
                      ))}
                    </select>
                    <span className="mapping-arrow">→</span>
                    <span className="mapping-arrow-label">目标列</span>
                    <select
                      value={m.targetColumn}
                      onChange={(e) =>
                        updateMapping(
                          rule.id,
                          mi,
                          "targetColumn",
                          e.target.value
                        )
                      }
                    >
                      {targetCols.map((col) => (
                        <option key={col} value={col}>
                          {columnLabel(col, targetSheet)}
                        </option>
                      ))}
                    </select>
                    <button
                      className="danger small"
                      onClick={() => deleteMapping(rule.id, mi)}
                    >
                      ✕
                    </button>
                  </div>
                ))}
                <button
                  className="secondary small"
                  onClick={() => addMapping(rule.id)}
                >
                  + 添加映射
                </button>
              </div>
            </div>
          ))}

          <div className="generate-area">
            <button
              className="primary"
              onClick={handlePreview}
              disabled={rules.length === 0}
            >
              预览匹配结果
            </button>
          </div>
        </>
      )}

      {/* Match preview */}
      {showPreview && matchPreview && (
        <div className="match-preview-section">
          <h3>匹配预览</h3>
          <div className="result-stats">
            <span className="stat-item">
              共 <strong>{matchPreview.stats.totalRows}</strong> 行
            </span>
            <span className="stat-item stat-matched">
              已匹配：<strong>{matchPreview.stats.matchedRows}</strong>
            </span>
            {matchPreview.stats.unmatchedRows > 0 && (
              <span className="stat-item stat-unmatched">
                未匹配：<strong>{matchPreview.stats.unmatchedRows}</strong>
              </span>
            )}
            <span className="stat-item">
              匹配率：
              <strong>
                {matchPreview.stats.totalRows > 0
                  ? Math.round(
                      (matchPreview.stats.matchedRows /
                        matchPreview.stats.totalRows) *
                        100
                    )
                  : 0}
                %
              </strong>
            </span>
          </div>
          <SheetPreview
            data={{
              name: "Preview",
              columns: matchPreview.columns,
              rows: matchPreview.rows,
              rowCount: matchPreview.rows.length,
              colCount: matchPreview.columns.length,
              headerRowIndex: targetSheet?.headerRowIndex ?? 0,
              dataStartRow: targetSheet?.dataStartRow ?? 0,
              originalRowIndex: [],
            }}
            title=""
            highlight={{ matchStatus: matchPreview.matchStatus }}
          />
        </div>
      )}
    </div>
  );
}
