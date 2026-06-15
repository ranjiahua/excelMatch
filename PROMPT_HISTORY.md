# Excel 数据映射工具 — 开发对话记录

## 项目概述
Tauri v2 桌面应用，通过关联键匹配（类似 VLOOKUP）将源 Excel 文件的列数据填充到目标 Excel 文件中。
前端: React 19 + TypeScript + Vite 7 | 后端: Rust (Tauri v2) | 包管理: pnpm

## 对话时间线

### 1. 需求分析与设计
需求：两个 Excel，通过表头关联键匹配，类似 VLOOKUP，需要可配置界面。
设计：4 步向导式 UI(上传→预览→规则→结果)，MappingRule(关联键+列映射列表)，精确/忽略大小写/去空格匹配

### 2. 核心功能实现
重写 types.ts, excel.ts, App.tsx, 4 个组件。VLOOKUP 式映射引擎，多 Sheet 选择，规则 JSON 导入/导出

### 3. 界面中文化 + 一键清除
全部英文→中文，标题栏"一键清除"按钮(confirm 确认)

### 4. 导出空文件问题
问题: IPC invoke("save_file", data) 传递大数组导致数据丢失
解决: 改用 @tauri-apps/plugin-fs 的 writeFile() 直接写文件

### 5. 文件系统权限问题
问题: fs.write_file not allowed
解决: capabilities/default.json 添加 "fs:allow-write-file"

### 6. 表头被覆盖问题
问题: 源文件表头值覆盖目标文件表头
解决: applyMappings 循环从 tRow=1 开始, buildLookupIndex 从 i=1 开始

### 7. 复杂 Excel 文件支持（表头不在第一行）
真实文件(蜀山大队工资表): Row0标题, Row1摘要, Row2-4三层合并表头, Row5汇总, Row6+数据, 60248行多为空
解决: SheetData增加headerRowIndex/dataStartRow/originalRowIndex, 自动检测+手动调整, 过滤空行, 预览高亮

### 8. 导出保留目标文件格式
问题: xlsx(SheetJS) round-trip 丢失格式
解决: exceljs.load(buffer)→逐格改值→writeBuffer(), 完整保留字体/边框/行高/列宽/合并单元格

### 9. "Sheet not found" 错误（多次）
根因: ArrayBuffer 存在 React state 中被破坏
最终方案: useRef<File> 缓存原始文件, 保存时 FileReader 重新读取全新 ArrayBuffer

### 10. UI 美化
Indigo 紫色主色调+CSS变量, 系统字体栈(苹方/微软雅黑), 卡片阴影/按钮发光/输入框光环, 斑马纹表格, 完整dark mode

### 11. 项目配置与打包
目录重命名 template→excelMatch, 应用名"Excel数据映射工具", GitHub Actions 三平台构建(macOS/Windows/Linux), 更新 CLAUDE.md

## 关键设计决策
- ArrayBuffer 绝不放 React state → 用 useRef<File>
- 映射循环从 dataStartRow 开始，不覆盖表头
- 解析用 xlsx(SheetJS)，导出用 exceljs(双层兜底)

## 常见问题速查
| 问题 | 解决 |
|------|------|
| 导出空文件 | writeFile() 替代 IPC invoke |
| fs.write_file not allowed | 加 fs:allow-write-file |
| Sheet not found | File ref, 保存时重新读取 |
| 表头被覆盖 | 从 dataStartRow 开始循环 |
| 格式丢失 | exceljs 导出 |
| 构建缓存错误 | rm -rf src-tauri/target |
