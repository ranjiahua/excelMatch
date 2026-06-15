# CLAUDE.md

## Project Overview

**Excel 数据映射工具** — Tauri v2 桌面应用。通过关联键匹配（类似 VLOOKUP），将源 Excel 文件的列数据填充到目标 Excel 文件中。

- 前端: React 19 + TypeScript + Vite 7
- 后端: Rust (Tauri v2)
- 包管理: **pnpm**
- Excel 解析: `xlsx` (SheetJS) 用于解析，`exceljs` 用于格式保留导出

## Commands

```bash
pnpm dev          # Vite dev server (port 1420)
pnpm build        # TypeScript check + Vite production build
pnpm tauri dev    # Tauri dev mode (auto-starts Vite + native window)
pnpm tauri build  # Build native desktop app bundle
```

## Architecture

```
src/
├── main.tsx           # React entry point
├── App.tsx            # Root component: file upload → sheet select → preview → rules → result
├── App.css            # Modern design system (CSS variables, dark mode, indigo theme)
├── types.ts           # WorkbookData, SheetData, MappingRule, MappingResult, etc.
├── excel.ts           # Core logic: parseExcelFile, detectHeaderRow, detectDataStartRow,
│                      #   applyMappings (VLOOKUP engine), generateExcelBlob (format-preserving)
└── components/
    ├── FileUploader.tsx    # File input → parseExcelFile → callback with WorkbookData + File
    ├── SheetPreview.tsx    # Table preview with header row highlight (📋) and data start marker (📊)
    ├── MappingTable.tsx    # Rule cards: key config, match options, column mappings, match preview
    └── ResultPreview.tsx   # Result stats + preview + save (re-reads File for fresh ArrayBuffer)

src-tauri/
├── Cargo.toml             # excel-mapper crate: tauri, tauri-plugin-opener/dialog/fs
├── tauri.conf.json        # productName: "Excel数据映射工具", window 1200x800
├── capabilities/default.json  # core:default, opener:default, dialog:default, fs:default, fs:allow-write-file
└── src/
    ├── main.rs            # Binary entry: excel_mapper_lib::run()
    └── lib.rs             # Tauri builder with plugins (opener, dialog, fs) + greet command
```

## Key Design Decisions

### VLOOKUP-style Mapping Engine (`excel.ts`)
- `applyMappings(targetSheet, sourceSheet, rules)` — deep-copies target rows, applies rules
- Each rule has: `targetKeyColumn` + `sourceKeyColumn` (the lookup key pair), `matchStrategy` (exact/ignore-case/trim), `duplicateStrategy` (first/last), and `mappings[]` (sourceColumn → targetColumn pairs)
- Builds a `Map<normalizedKey, rowIndex[]>` lookup index from source for O(1) matching
- **Skips rows before `dataStartRow`** — headers/titles are never modified
- **Table header row is never overwritten** — mapping only applies to data rows

### Auto Header Detection
- `detectHeaderRow()`: finds first row with ≥3 non-empty cells (skips 1-2 cell title/abstract rows)
- `detectDataStartRow()`: after header, finds first row where col A is a pure integer (skips "总计"/"合计")
- Both can be manually adjusted via number inputs in the UI
- Empty rows are filtered out during parsing

### Format-Preserving Export
- `generateExcelBlob()` uses **exceljs** to load original file → modify cell values → write back
- This preserves ALL formatting: merged cells, fonts, colors, borders, row heights, column widths
- **Critical**: ArrayBuffer must NOT be stored in React state (gets corrupted). Instead:
  - Store the raw `File` object in a `useRef`
  - Re-read the file with `FileReader` at save time to get a fresh ArrayBuffer
  - Pass the fresh buffer to exceljs
- Falls back to xlsx (SheetJS) if exceljs fails

### State Management
- `targetFileRef` / `sourceFileRef` are `useRef<File>` — never store ArrayBuffer in state
- Sheet data lives in React state as `WorkbookData`
- `headerRowIndex` and `dataStartRow` are mutable on the SheetData object (user can adjust)
- `originalRowIndex[]` maps each filtered row back to its 1-based Excel row number

## Tauri Plugins & Permissions
- `tauri-plugin-dialog` — `save()` for native file save dialog
- `tauri-plugin-fs` — `writeFile()` for direct file writing (avoids IPC serialization issues with large binary data)
- `tauri-plugin-opener` — default opener plugin
- Permission `fs:allow-write-file` is required (fs:default only grants read)

## Troubleshooting History

1. **Empty export file**: IPC `invoke("save_file", data)` serialized large arrays incorrectly → switched to `writeFile()` from tauri-plugin-fs
2. **"Sheet not found" error**: ArrayBuffer corrupted in React state → switched to storing `File` object in ref, re-reading at save time
3. **Header row being overwritten**: Mapping loop started from row 0 → changed to start from `dataStartRow`
4. **Format lost on export**: xlsx (SheetJS) doesn't preserve formatting → switched to exceljs for export
5. **Permissions error**: `fs:default` only grants read → added `fs:allow-write-file`
6. **Build cache stale after rename**: Renamed project directory → needed `rm -rf src-tauri/target` to clear cached paths
