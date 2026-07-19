# etzhayyim-project-xlsx — xlsx.etzhayyim.com

**Excel editor** — XLSX upload, HTML DOM grid editing, 131-function formula engine, kagami graph persistence, XLSX/CSV export.

## Architecture

| Item | Value |
|---|---|
| Domain | `xlsx.etzhayyim.com` |
| Runtime | **Single Worker** (TS Native) |
| nanoid | `il0ndq6a` |
| performerType | `service` (default sensitivity: `internal`) |
| uiType | `appview` |
| Rendering | **HTML DOM** (virtualized table, Google Sheets 方式) |
| Tests | **5 suites, 505 BDD scenarios** |

## Data Pipeline

```
Upload (.xlsx ZIP / .csv)
  → fflate zip decompress + SpreadsheetML XML parse (ooxml-parser.ts)
  → kagami graph: SQL nodes (Workbook/Sheet/Cell/Style/SharedString/MergedRegion/Chart/Table/DefinedName)
  → kagami graph write (P10v2 per-label Iceberg tables)
  → Client: Hyperdrive direct query → RisingWave
  → HTML DOM: virtualized table grid (30 rows × 20 cols viewport)
  → Formula engine: 131 functions, topological recalc, circular detection
  → Export: graph → SpreadsheetML XML → zip → .xlsx download
  → Export: CSV (RFC 4180)
```

## Rendering Architecture

**HTML DOM + Svelte 5 reactivity** (Canvas 2D から移行済み, 2026-04-05)

| 要素 | 実装 |
|---|---|
| Grid | `<table>` with virtualized rows/cols (scrollRow/scrollCol offset) |
| Cell input | Native `<input>` in-cell (ダブルクリックで edit mode) |
| Selection | CSS class `.selected` `.active` (ネイティブ DOM) |
| Scroll | wheel event → scrollRow/scrollCol 更新 → Svelte re-render |
| Style | CSS inline style (font-weight, color, background, text-align) |
| Reactivity | `renderVersion` counter で `Map.set()` 変更を強制 track |

**重要**: `Map.set()` / `Map.delete()` は Svelte 5 `$state` の reactive tracking 対象外。`getCellAt(ref)` 内で `void renderVersion` を読み、`forceRender()` (`renderVersion++`) で全セルの再評価を強制する。

## Formula Engine (131 functions)

| Category | Count | Functions |
|---|---|---|
| Math & Trig | 27 | SUM, SUMPRODUCT, AVERAGE, MIN, MAX, COUNT, COUNTA, COUNTBLANK, ROUND, ROUNDUP, ROUNDDOWN, CEILING, FLOOR, INT, SIGN, ABS, MOD, SQRT, POWER, LOG, LOG10, LN, EXP, PI, RAND, RANDBETWEEN, PRODUCT, SUBTOTAL |
| Text | 22 | CONCATENATE, CONCAT, TEXTJOIN, LEFT, RIGHT, MID, LEN, FIND, SEARCH, SUBSTITUTE, REPLACE, REPT, EXACT, TRIM, CLEAN, UPPER, LOWER, PROPER, TEXT, VALUE, CHAR, CODE, T |
| Logical | 11 | IF, IFS, SWITCH, AND, OR, XOR, NOT, IFERROR, IFNA, TRUE, FALSE |
| Lookup | 11 | VLOOKUP, HLOOKUP, INDEX, MATCH, XLOOKUP, CHOOSE, ROW, COLUMN, ROWS, COLUMNS, INDIRECT |
| Date & Time | 18 | NOW, TODAY, DATE, TIME, YEAR, MONTH, DAY, HOUR, MINUTE, SECOND, WEEKDAY, DATEVALUE, DATEDIF, EDATE, EOMONTH, NETWORKDAYS, WORKDAY, ISOWEEKNUM |
| Statistical | 16 | COUNTIF, COUNTIFS, SUMIF, SUMIFS, AVERAGEIF, MEDIAN, MODE, STDEV/S/P, VAR/S/P, LARGE, SMALL, RANK, PERCENTILE, MAXIFS, MINIFS |
| Information | 11 | ISNUMBER, ISTEXT, ISBLANK, ISERROR, ISNA, ISLOGICAL, ISFORMULA, TYPE, N, NA, ERROR.TYPE |
| Financial | 8 | PMT, FV, PV, NPV, IRR, RATE, NPER, SLN |

## Graph Model (kagami SQL)

| XLSX Structure | SQL Node | Edge |
|---|---|---|
| Workbook | `(:Workbook {did, title, activeSheet})` | |
| Sheet | `(:Sheet {name, order, hidden, frozenRow, frozenCol})` | `(:Workbook)-[:HAS_SHEET]->(:Sheet)` |
| Cell | `(:Cell {ref, row, col, type, value, formula, styleId})` | `(:Sheet)-[:CONTAINS]->(:Cell)` |
| Style | `(:Style {numFmt, font, fill, border, alignment})` | `(:Cell)-[:HAS_STYLE]->(:Style)` |
| SharedString | `(:SharedString {index, text})` | `(:Cell)-[:USES_STRING]->(:SharedString)` |
| MergedRegion | `(:MergedRegion {startRef, endRef})` | `(:Sheet)-[:HAS_MERGE]->(:MergedRegion)` |
| DefinedName | `(:DefinedName {name, scope, ref})` | `(:Workbook)-[:HAS_NAME]->(:DefinedName)` |
| Table | `(:Table {name, ref, headerRow})` | `(:Sheet)-[:HAS_TABLE]->(:Table)` |
| Chart | `(:Chart {type, title, dataRange})` | `(:Sheet)-[:HAS_CHART]->(:Chart)` |
| ConditionalFormat | `(:ConditionalFormat {type, priority, formula})` | `(:Sheet)-[:HAS_COND_FMT]->(:ConditionalFormat)` |
| DataValidation | `(:DataValidation {type, ref, formula1})` | `(:Sheet)-[:HAS_VALIDATION]->(:DataValidation)` |
| Formula dependency | — | `(:Cell)-[:DEPENDS_ON]->(:Cell)` |

## Multi-DID Architecture `[DESIGN]`

| DID | Purpose |
|---|---|
| `did:web:xlsx.etzhayyim.com` | Controller (app) |
| `did:web:xlsx.etzhayyim.com:workbook:{nanoid}` | Individual workbook |
| `did:web:xlsx.etzhayyim.com:template:{nanoid}` | Reusable workbook template |

## Design E 3-Tier Write

| Tier | Purpose | Function | Collection NSID |
|---|---|---|---|
| **1 Social** | Share workbook | `AppBskyFeedPost(did, text, {embed})` | `app.bsky.feed.post` |
| **2 Domain** | workbook/sheet/cell/style/chart | `ComAtprotoRepoCreateRecord(kind, payload)` | `com.etzhayyim.apps.xlsx.*` |
| **3 State** | Editor preferences | `Preferences()` | server-side |

## Domain Record Types (Tier 2, camelCase) `[DESIGN]`

| Kind | NSID | Content |
|---|---|---|
| `workbook` | `com.etzhayyim.apps.xlsx.workbook` | Workbook metadata (title, activeSheet) |
| `sheet` | `com.etzhayyim.apps.xlsx.sheet` | Sheet definition (name, order, frozen panes) |
| `cell` | `com.etzhayyim.apps.xlsx.cell` | Cell value/formula/type/style reference |
| `style` | `com.etzhayyim.apps.xlsx.style` | Number format, font, fill, border, alignment |
| `chart` | `com.etzhayyim.apps.xlsx.chart` | Chart definition (type, data range, series) |
| `table` | `com.etzhayyim.apps.xlsx.table` | Structured table (ListObject) |
| `definedName` | `com.etzhayyim.apps.xlsx.definedName` | Named range |
| `workbookTemplate` | `com.etzhayyim.apps.xlsx.workbookTemplate` | Reusable template |

## File Structure

```
60-apps/etzhayyim-project-xlsx/
├── CLAUDE.md                            # This file
├── COMPATIBILITY.md                     # Excel API / Google Sheets API coverage matrix
├── COVERAGE.md                          # Test coverage report
├── wit/xlsx/package.wit                  # Domain WIT capability (3 interfaces)
└── wasm/etzhayyim-wasm-xlsx-il0ndq6a/
    ├── src/app.ts                       # TS Native — Design E, 20 XRPC commands
    ├── kotodama.jsonld                  # App metadata, DID, triggers
    ├── wrangler.jsonc                   # CF Worker config
    ├── wit/world.wit                    # Component WIT (contract + capability export)
    └── svelte/
        ├── src/App.svelte               # Main UI: toolbar, formula bar, HTML grid, sheet tabs, find/replace
        └── src/lib/
            ├── ooxml-parser.ts          # XLSX ZIP → XlsxWorkbook (SST, styles, cells, merges, freeze)
            ├── xlsx-exporter.ts         # XlsxWorkbook → SpreadsheetML → ZIP → .xlsx
            ├── formula-engine.ts        # 131 functions, tokenizer, recursive descent, topological recalc
            ├── editor-state.svelte.ts   # Svelte 5 rune state (43 exports: undo/redo, CRUD, formatting)
            ├── grid-renderer.ts         # Canvas 2D renderer (retained for KAMI chart bridge)
            ├── cell-selection.ts        # Range utils, autofill, keyboard nav, jump-to-edge
            ├── clipboard-handler.ts     # TSV + HTML copy/paste (Excel/Sheets interop)
            ├── csv-handler.ts           # CSV import/export (RFC 4180)
            ├── find-replace.ts          # Find/Replace engine
            ├── comments.ts              # CommentStore (threaded comments, JSON serialize)
            ├── kami-bridge.ts           # WebGPU/KAMI 3D chart (optional)
            └── __tests__/
                ├── e2e-bdd.test.ts           # 193 scenarios (formula, roundtrip, CSV, find, selection)
                ├── editor-render-bdd.test.ts  # 52 scenarios (grid sizing, viewport, hit test, OOXML)
                ├── coverage-expansion.test.ts # 140 scenarios (edge cases, all remaining formulas)
                ├── deep-coverage.test.ts      # 69 scenarios (integration, style, kami-bridge, stress)
                └── editor-logic-bdd.test.ts   # 46 scenarios (undo/redo, sheet ops, formatting, workflow)
```

## Test Coverage

```
5 Test Suites, 505 BDD Scenarios, 0 Failures
Run: cd svelte && npx tsx src/lib/__tests__/<suite>.test.ts
```

| Metric | Value |
|---|---|
| Formula functions tested | 131/131 = **100%** |
| Editor-state exports tested | 42/43 = **98%** (autoResizeColumn = Canvas dep) |
| Module coverage | 11/11 = **100%** |
| Browser-only untestable | 7 functions (renderGrid, WebGPU, clipboard, downloadBlob) |
| Feature coverage vs Excel | **~40%** |
| Feature coverage vs Google Sheets | **~56%** |

## CRITICAL: Map Reactivity

→ `etzhayyim dodaf tv1 query --id xlsx-map-reactivity` / MCP `etzhayyim.dodaf.tv1.query`
