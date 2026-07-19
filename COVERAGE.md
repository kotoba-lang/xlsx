# xlsx.etzhayyim.com — API Coverage & Test Coverage Report

## 1. Formula Function Coverage vs Excel / Google Sheets

### Implementation: 133 unique functions

| Category | xlsx.etzhayyim.com | Excel (500+) | Google Sheets (~400) | Coverage vs Excel | Coverage vs Sheets |
|---|---|---|---|---|---|
| **Math & Trig** | 27 (SUM, SUMPRODUCT, AVERAGE, MIN, MAX, COUNT, COUNTA, COUNTBLANK, ROUND, ROUNDUP, ROUNDDOWN, CEILING, FLOOR, INT, SIGN, ABS, MOD, SQRT, POWER, LOG, LOG10, LN, EXP, PI, RAND, RANDBETWEEN, PRODUCT, SUBTOTAL) | 60+ | 40+ | 45% | 68% |
| **Text** | 22 (CONCATENATE, CONCAT, TEXTJOIN, LEFT, RIGHT, MID, LEN, FIND, SEARCH, SUBSTITUTE, REPLACE, REPT, EXACT, TRIM, CLEAN, UPPER, LOWER, PROPER, TEXT, VALUE, CHAR, CODE, T) | 30+ | 25+ | 73% | 88% |
| **Logical** | 11 (IF, IFS, SWITCH, AND, OR, XOR, NOT, IFERROR, IFNA, TRUE, FALSE) | 10 | 8 | **110%** | **138%** |
| **Lookup & Reference** | 11 (VLOOKUP, HLOOKUP, INDEX, MATCH, XLOOKUP, CHOOSE, ROW, COLUMN, ROWS, COLUMNS, INDIRECT) | 20+ (+ XMATCH, FILTER, SORT, UNIQUE, OFFSET) | 15+ | 55% | 73% |
| **Date & Time** | 18 (NOW, TODAY, DATE, TIME, YEAR, MONTH, DAY, HOUR, MINUTE, SECOND, WEEKDAY, DATEVALUE, DATEDIF, EDATE, EOMONTH, NETWORKDAYS, WORKDAY, ISOWEEKNUM) | 25+ | 20+ | 72% | 90% |
| **Statistical** | 16 (COUNTIF, COUNTIFS, SUMIF, SUMIFS, AVERAGEIF, MEDIAN, MODE, STDEV/S/P, VAR/S/P, LARGE, SMALL, RANK, PERCENTILE, MAXIFS, MINIFS) | 50+ | 25+ | 32% | 64% |
| **Information** | 11 (ISNUMBER, ISTEXT, ISBLANK, ISERROR, ISNA, ISLOGICAL, ISFORMULA, TYPE, N, NA, ERROR.TYPE) | 15+ | 12+ | 73% | 92% |
| **Financial** | 8 (PMT, FV, PV, NPV, IRR, RATE, NPER, SLN) | 50+ | 15+ | 16% | 53% |
| **Engineering** | 0 | 40+ | 0 | 0% | N/A |
| **Database** | 0 | 10 | 0 | 0% | N/A |
| **Web** | 0 | 3 | 3 | 0% | 0% |
| **Dynamic Arrays** | 0 (XLOOKUP only) | 20+ (FILTER, SORT, UNIQUE, SEQUENCE, etc.) | 0 | 0% | N/A |
| **Cube** | 0 | 7 | 0 | 0% | N/A |
| **Total** | **133** | **~500** | **~400** | **26.6%** | **33.3%** |

### Missing High-Priority Functions (used in >50% of real spreadsheets)

| Priority | Functions | Impact |
|---|---|---|
| **Critical** | SUMIFS (multi-criteria), COUNTIFS (multi-criteria), AVERAGEIFS | Most-used aggregation — current impl is simplified (single criteria pair) |
| **High** | XMATCH, FILTER, SORT, UNIQUE, SEQUENCE | Dynamic array functions (Excel 365) |
| **High** | OFFSET, ADDRESS, HYPERLINK | Reference manipulation |
| **High** | TRIM (full: multi-space → single), NUMBERVALUE, FIXED | Text cleanup |
| **Medium** | FREQUENCY, CORREL, LINEST, TREND, FORECAST | Advanced statistics |
| **Medium** | IPMT, PPMT, CUMIPMT, DB, DDB, SYD | Financial (depreciation, amortization) |
| **Medium** | WEEKNUM, DAYS, DAYS360, YEARFRAC | Date functions |
| **Low** | BIN2DEC, DEC2BIN, HEX2DEC, CONVERT | Engineering |
| **Low** | DAVERAGE, DSUM, DCOUNT | Database functions |
| **Low** | CUBEVALUE, CUBEMEMBER | Cube functions |

---

## 2. Feature Coverage vs Excel API / Google Sheets API

### Spreadsheet Operations

| Feature Category | Sub-Feature | xlsx.etzhayyim.com | Excel API | Google Sheets API | BDD Tested |
|---|---|---|---|---|---|
| **Workbook** | Create | YES | YES | YES | YES (roundtrip) |
| | Open (XLSX parse) | YES | YES | YES | YES (3 scenarios) |
| | Save (XLSX export) | YES | YES | YES | YES (3 scenarios) |
| | CSV Import | YES | YES | YES | YES (6 scenarios) |
| | CSV Export | YES | YES | YES | YES (included in CSV) |
| | Recalculation (topological) | YES | YES | YES | YES (3 scenarios) |
| | Protection | NO | YES | NO | — |
| | Properties | title only | YES (full) | YES | — |
| **Sheet** | Add / Delete | YES | YES | YES | — |
| | Rename | YES | YES | YES | — |
| | Reorder | YES | YES | YES | — |
| | Hidden | parse only | YES (+veryHidden) | YES | — |
| | Freeze Panes | YES (parse+apply) | YES | YES | YES (roundtrip) |
| | Tab Color | NO | YES | YES | — |
| | Protection | NO | YES (granular) | NO | — |
| | Page Setup / Print | NO | YES | NO | — |
| **Cell / Range** | Get / Set Value | YES | YES | YES | YES (clipboard) |
| | Formula (133 functions) | YES | YES (500+) | YES (400+) | YES (42 scenarios) |
| | Number types (string/number/bool/date/error) | YES | YES | YES | YES (roundtrip) |
| | Insert / Delete Rows | YES | YES | YES | — |
| | Insert / Delete Columns | YES | YES | YES | — |
| | Merge / Unmerge | YES | YES | YES | YES (roundtrip) |
| | Sort Range | backend cmd | YES (multi-key) | YES | — |
| | Filter / AutoFilter | backend cmd | YES (advanced) | YES | — |
| | Find / Replace | YES | YES | YES | YES (4 scenarios) |
| | Autofill (pattern detection) | YES | YES | NO (API) | YES (4 scenarios) |
| | Copy / Paste (TSV + HTML) | YES | YES | YES | YES (2 scenarios) |
| | Hyperlink | parse only | YES | YES | — |
| | Rich Text (in-cell) | parse only | YES | YES | — |
| | Comments / Notes | YES (CommentStore) | YES (threaded) | YES (notes) | YES (4 scenarios) |
| | Auto-resize columns | YES | YES | YES | — |
| | Row / Column grouping | NO | YES | YES | — |
| **Formatting** | Bold / Italic / Underline | YES (apply) | YES | YES | — |
| | Font Size / Color | YES (apply) | YES | YES | — |
| | Fill / Background Color | YES (apply) | YES | YES | — |
| | Borders (4 edges) | YES (render + apply) | YES (8 positions, diagonal) | YES | — |
| | Number Format | YES (apply) | YES (50+ built-in) | YES | — |
| | Alignment (h/v) | YES (apply) | YES (+shrinkToFit, indent) | YES | — |
| | Wrap Text | parse only | YES | YES | — |
| | Text Rotation | parse only | YES | YES | — |
| | Conditional Formatting | data bar render | YES (15+ rule types) | YES (boolean+gradient) | — |
| | Cell Styles / Named | NO | YES (60+ table styles) | NO | — |
| **Charts** | Bar / Column / Line / Pie / Scatter / Area / Radar | YES (KAMI 3D) | YES (40+ types) | YES (16 types) | — |
| | Data Labels / Trendlines | NO | YES | YES | — |
| | Chart Export (image) | NO | YES | NO | — |
| **Tables** | Create (ListObject) | parse only | YES | NO (native) | — |
| | Headers / Totals Row | parse only | YES | NO | — |
| | Structured References | NO | YES | NO | — |
| **Pivot Tables** | Create / Refresh | NO | YES | YES | — |
| | Slicers / Timelines | NO | YES | NO | — |
| **Named Ranges** | Create / Use in formulas | parse only | YES | YES | — |
| **Data Validation** | List / Number / Date | parse only | YES | YES | — |
| **Shapes / Drawing** | Shapes / Images / TextBoxes | NO | YES | NO | — |
| **Collaboration** | Real-time Co-Editing | NO | YES | YES | — |

---

## 3. BDD Test Coverage Matrix

### 83 Scenarios across 18 Features

| Feature | Scenarios | Module(s) Tested | Coverage Level |
|---|---|---|---|
| Cell Reference Helpers | 7 | ooxml-parser.ts | **Complete** — all 5 exported functions |
| XLSX Save / Open Roundtrip | 3 | ooxml-parser.ts, xlsx-exporter.ts | **Good** — string/number/formula/merge/freeze/colWidth/multisheet |
| Formula — Basic Math | 6 | formula-engine.ts | **Partial** — 6 of 27 math functions |
| Formula — Text | 9 | formula-engine.ts | **Good** — 9 of 22 text functions |
| Formula — Logical | 7 | formula-engine.ts | **Good** — 7 of 11 logical functions |
| Formula — Statistical | 5 | formula-engine.ts | **Partial** — 5 of 16 statistical functions |
| Formula — Lookup | 3 | formula-engine.ts | **Partial** — 3 of 11 lookup functions |
| Formula — Financial | 4 | formula-engine.ts | **Partial** — 4 of 8 financial functions |
| Formula — Information | 5 | formula-engine.ts | **Partial** — 5 of 11 information functions |
| Formula — Dependency & Recalc | 3 | formula-engine.ts | **Complete** — recalc order, circular detect, dep extraction |
| CSV Import / Export | 6 | csv-handler.ts | **Complete** — parse, quoted fields, newlines, workbook conv, export, roundtrip |
| Find & Replace | 4 | find-replace.ts | **Complete** — find, matchEntireCell, matchCase, replaceAll |
| Cell Selection & Range | 7 | cell-selection.ts | **Complete** — normalize, toString, parse, containment, tab nav |
| Autofill Pattern Detection | 4 | cell-selection.ts | **Complete** — increment, step detect, copy fallback, fill generation |
| Clipboard (TSV Paste) | 2 | clipboard-handler.ts | **Good** — parse, apply with type detection |
| Comments / Notes | 4 | comments.ts | **Complete** — CRUD, reply, serialize/deserialize |
| Column / Row Selection | 2 | cell-selection.ts | **Complete** |
| XML Entity Roundtrip | 1 | ooxml-parser.ts, xlsx-exporter.ts | **Complete** — formula with >, " chars |

### Formula BDD Test Coverage (Tested / Implemented)

| Category | Implemented | BDD Tested | Test Rate |
|---|---|---|---|
| Math & Trig | 27 | 6 | 22% |
| Text | 22 | 9 | 41% |
| Logical | 11 | 7 | 64% |
| Statistical | 16 | 5 | 31% |
| Lookup | 11 | 3 | 27% |
| Financial | 8 | 4 | 50% |
| Information | 11 | 5 | 45% |
| Date & Time | 18 | 0 | 0% |
| **Total** | **133** | **39** | **29%** |

---

## 4. Visual / Rendering Test Coverage

### Canvas 2D Rendering (grid-renderer.ts) — NO automated visual tests

| Render Feature | Implemented | Visually Tested (manual) | Automated Test |
|---|---|---|---|
| Cell text rendering | YES | YES | NO |
| Number right-alignment | YES | YES | NO |
| Font style (bold/italic) | YES | — | NO |
| Cell background fill | YES | YES | NO |
| Border rendering (4 edges) | YES | — | NO |
| Gridlines | YES | YES | NO |
| Selection highlight | YES | YES | NO |
| Fill handle (autofill corner) | YES | YES | NO |
| Frozen pane dividers | YES | — | NO |
| Row headers (1, 2, 3...) | YES | YES | NO |
| Column headers (A, B, C...) | YES | YES | NO |
| Virtual scrolling (100K+ rows) | YES | — | NO |
| Zoom scaling | YES | — | NO |
| Conditional format data bar | YES | — | NO |
| Hit testing (cell click) | YES | — | NO |
| Hit testing (header click) | YES | — | NO |
| **Visual test rate** | **16 features** | **~6 manual** | **0% automated** |

### Recommended Visual Testing Approach

1. **Playwright pixel-diff** — headless browser → Canvas screenshot → compare baseline
2. **Storybook-like isolation** — render individual cells/ranges → snapshot
3. **Hit test unit tests** — mock canvas coordinates → verify cell ref output

---

## 5. Overall Coverage Summary (2026-04-05 updated)

| Metric | Value |
|---|---|
| **Formula functions implemented** | 131 / 500 (Excel) = 26.2%, 131 / 400 (Sheets) = 32.8% |
| **Feature operations** | 50 / 120 (Excel) = 41.7%, 50 / 80 (Sheets) = 62.5% |
| **BDD scenarios** | **505** |
| **BDD pass rate** | **505 / 505 = 100%** |
| **Test suites** | 5 files, 3,887 lines |
| **Module coverage** | 11/11 = **100%** |
| **Formula test coverage** | **131 / 131 = 100%** |
| **Editor-state export coverage** | **42 / 43 = 98%** (autoResizeColumn = Canvas dep) |
| **Browser-only untestable** | 7 functions (renderGrid, WebGPU×5, clipboard, downloadBlob) |
| **Rendering** | HTML DOM (virtualized table, Google Sheets 方式) — Canvas 2D から移行済み |
| **Total code** | ~8,000 lines (app + lib + tests) |

### Remaining Gaps

| Gap | Impact | Effort | Status |
|---|---|---|---|
| ~~Formula BDD coverage~~ | ~~29%~~ | — | **DONE (100%)** |
| ~~Date/Time tests~~ | ~~0%~~ | — | **DONE (15 scenarios)** |
| ~~Formatting apply tests~~ | ~~0%~~ | — | **DONE (6 scenarios)** |
| ~~Editor state tests~~ | ~~0%~~ | — | **DONE (46 scenarios)** |
| Dynamic array formulas (FILTER, SORT, UNIQUE) | Low | High | P3 — not in Google Sheets API |
| Pivot tables | High | Very High | P4 — major Excel/Sheets feature |
| Real-time collaboration (CRDT) | High | Very High | P4 — kagami CRDT integration |
| Playwright visual tests | Medium | Medium | P3 — HTML DOM approach reduces need |
| Named range resolution in formulas | Medium | Medium | P3 |
| Data validation dropdown UI | Medium | Medium | P3 |
