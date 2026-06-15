import type { SheetData } from "../types";

interface SheetPreviewProps {
  data: SheetData;
  title: string;
  /** Optional: highlight matched/unmatched rows (from MappingResult). */
  highlight?: {
    matchStatus: ("matched" | "unmatched")[];
  };
}

const MAX_PREVIEW_ROWS = 100;

export default function SheetPreview({
  data,
  title,
  highlight,
}: SheetPreviewProps) {
  if (data.rows.length === 0) {
    return (
      <div className="sheet-preview">
        {title && <h3>{title}</h3>}
        <p className="empty-state">工作表中无数据</p>
      </div>
    );
  }

  const truncated = data.rowCount > MAX_PREVIEW_ROWS;
  const displayRows = data.rows.slice(0, MAX_PREVIEW_ROWS);

  return (
    <div className="sheet-preview">
      {title && <h3>{title}</h3>}
      <div className="table-container">
        <table>
          <thead>
            <tr>
              <th className="row-header">#</th>
              {data.columns.map((col) => (
                <th key={col}>{col}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {displayRows.map((row, i) => {
              const status = highlight?.matchStatus?.[i];
              const isHeader = i === data.headerRowIndex;
              const isDataStart = i === data.dataStartRow;

              let rowClass = "";
              if (status === "unmatched") rowClass = "row-unmatched";
              if (isHeader) rowClass += " row-header-highlight";

              return (
                <tr key={i} className={rowClass}>
                  <td className="row-header">
                    {status === "unmatched" ? (
                      <span title="未匹配行">⚠️ {i + 1}</span>
                    ) : isHeader ? (
                      <span title="表头行">📋 {i + 1}</span>
                    ) : isDataStart ? (
                      <span title="数据起始行">📊 {i + 1}</span>
                    ) : (
                      i + 1
                    )}
                  </td>
                  {row.map((cell, j) => (
                    <td key={j}>{cell ?? ""}</td>
                  ))}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      {truncated && (
        <p className="truncation-note">
          仅显示前 {MAX_PREVIEW_ROWS} 行，共 {data.rowCount} 行
        </p>
      )}
      <p className="truncation-note">
        表头行：第 {data.headerRowIndex + 1} 行，数据起始：第{" "}
        {data.dataStartRow + 1} 行
      </p>
    </div>
  );
}
