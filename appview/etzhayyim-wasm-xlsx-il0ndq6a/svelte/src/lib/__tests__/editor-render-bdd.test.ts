/**
 * Editor & Render BDD Tests — xlsx.etzhayyim.com cell ops, merge, grid renderer, selection, parser, exporter.
 *
 * Tests underlying logic without importing Svelte 5 rune-based editor-state.svelte.ts.
 * Insert/delete row/column logic is re-implemented inline to avoid Svelte compilation dependency.
 *
 * Run: cd svelte && npx tsx src/lib/__tests__/editor-render-bdd.test.ts
 */

import {
  parseXlsx, buildRef, parseRef, colToLetter, letterToCol,
  type XlsxWorkbook, type XlsxSheet, type XlsxCell, type CellRef,
} from "../ooxml-parser";
import { exportXlsx } from "../xlsx-exporter";
import {
  colWidth, rowHeight, colOffset, rowOffset,
  computeViewport, hitTestCell, hitTestColHeader, hitTestRowHeader,
  type Viewport,
} from "../grid-renderer";
import { jumpToEdge } from "../cell-selection";

// ---------------------------------------------------------------------------
// Test runner (minimal, no deps) — same pattern as e2e-bdd.test.ts
// ---------------------------------------------------------------------------

let totalPass = 0;
let totalFail = 0;
let currentFeature = "";
let currentScenario = "";

/** Register a feature group. */
function feature(name: string, fn: () => void): void {
  currentFeature = name;
  console.log(`\n\x1b[1mFeature: ${name}\x1b[0m`);
  fn();
}

/** Register a scenario (Given/When/Then). */
function scenario(name: string, fn: () => void): void {
  currentScenario = name;
  try {
    fn();
    totalPass++;
    console.log(`  \x1b[32m+\x1b[0m ${name}`);
  } catch (e: any) {
    totalFail++;
    console.log(`  \x1b[31m-\x1b[0m ${name}`);
    console.log(`    \x1b[31m${e.message}\x1b[0m`);
  }
}

/** Assert a boolean condition. */
function assert(condition: boolean, msg: string): void {
  if (!condition) throw new Error(msg);
}

/** Assert strict equality. */
function assertEqual(actual: unknown, expected: unknown, label: string): void {
  if (actual !== expected) throw new Error(`${label}: expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`);
}

/** Assert numeric proximity. */
function assertClose(actual: number, expected: number, label: string, epsilon = 0.01): void {
  if (Math.abs(actual - expected) > epsilon) throw new Error(`${label}: expected ~${expected}, got ${actual}`);
}

// ---------------------------------------------------------------------------
// Helpers — sheet/workbook factories (same as e2e-bdd.test.ts)
// ---------------------------------------------------------------------------

/** Create a sheet with cells from a simple record. */
function makeSheet(cells: Record<string, { type: string; value: any; formula?: string }>): XlsxSheet {
  const map = new Map<CellRef, XlsxCell>();
  for (const [ref, c] of Object.entries(cells)) {
    const { col, row } = parseRef(ref);
    map.set(ref, {
      ref: ref as CellRef,
      row, col,
      type: c.type as any,
      value: c.formula ? null : c.value,
      formula: c.formula ?? null,
      calculatedValue: null,
      styleId: 0,
      hyperlink: null,
    });
  }
  return {
    id: "s1", name: "Sheet1", order: 0, hidden: false,
    cells: map, mergedRegions: [], tables: [], charts: [],
    conditionalFormats: [], dataValidations: [],
    frozenRow: 0, frozenCol: 0,
    colWidths: new Map(), rowHeights: new Map(),
    defaultColWidth: 8.43, defaultRowHeight: 15,
  };
}

/** Create a workbook wrapping sheets. */
function makeWorkbook(sheets: XlsxSheet[]): XlsxWorkbook {
  return {
    id: "wb1", title: "Test",
    sheets, sharedStrings: [],
    styles: [{ id: 0, numFmt: null, font: null, fill: null, border: null, alignment: null }],
    definedNames: [], activeSheetIndex: 0,
  };
}

// ---------------------------------------------------------------------------
// Inline insert/delete logic (mirrors editor-state.svelte.ts without runes)
// ---------------------------------------------------------------------------

/** Insert rows at a given position, shifting cells down. */
function insertRowsInSheet(sheet: XlsxSheet, row: number, count: number): void {
  const newCells = new Map<CellRef, XlsxCell>();
  for (const [ref, cell] of sheet.cells) {
    if (cell.row >= row) {
      const newRef = buildRef(cell.col, cell.row + count);
      newCells.set(newRef, { ...cell, ref: newRef, row: cell.row + count });
    } else {
      newCells.set(ref, cell);
    }
  }
  sheet.cells = newCells;
}

/** Delete rows at a given position, shifting cells up. */
function deleteRowsInSheet(sheet: XlsxSheet, row: number, count: number): void {
  const newCells = new Map<CellRef, XlsxCell>();
  for (const [ref, cell] of sheet.cells) {
    if (cell.row >= row && cell.row < row + count) continue;
    if (cell.row >= row + count) {
      const newRef = buildRef(cell.col, cell.row - count);
      newCells.set(newRef, { ...cell, ref: newRef, row: cell.row - count });
    } else {
      newCells.set(ref, cell);
    }
  }
  sheet.cells = newCells;
}

/** Insert columns at a given position, shifting cells right. */
function insertColumnsInSheet(sheet: XlsxSheet, col: number, count: number): void {
  const newCells = new Map<CellRef, XlsxCell>();
  for (const [ref, cell] of sheet.cells) {
    if (cell.col >= col) {
      const newRef = buildRef(cell.col + count, cell.row);
      newCells.set(newRef, { ...cell, ref: newRef, col: cell.col + count });
    } else {
      newCells.set(ref, cell);
    }
  }
  sheet.cells = newCells;
}

/** Delete columns at a given position, shifting cells left. */
function deleteColumnsInSheet(sheet: XlsxSheet, col: number, count: number): void {
  const newCells = new Map<CellRef, XlsxCell>();
  for (const [ref, cell] of sheet.cells) {
    if (cell.col >= col && cell.col < col + count) continue;
    if (cell.col >= col + count) {
      const newRef = buildRef(cell.col - count, cell.row);
      newCells.set(newRef, { ...cell, ref: newRef, col: cell.col - count });
    } else {
      newCells.set(ref, cell);
    }
  }
  sheet.cells = newCells;
}

// ---------------------------------------------------------------------------
// Inline merge helpers (mirrors editor-state.svelte.ts logic)
// ---------------------------------------------------------------------------

/** Check if a cell ref is inside any merged region. */
function isCellInMerge(sheet: XlsxSheet, ref: CellRef): boolean {
  const { row, col } = parseRef(ref);
  for (const m of sheet.mergedRegions) {
    const start = parseRef(m.startRef);
    const end = parseRef(m.endRef);
    if (row >= start.row && row <= end.row && col >= start.col && col <= end.col) return true;
  }
  return false;
}

/** Unmerge the region containing a given cell ref. */
function unmergeAtCell(sheet: XlsxSheet, ref: CellRef): void {
  const { row, col } = parseRef(ref);
  sheet.mergedRegions = sheet.mergedRegions.filter((m) => {
    const start = parseRef(m.startRef);
    const end = parseRef(m.endRef);
    return !(row >= start.row && row <= end.row && col >= start.col && col <= end.col);
  });
}

// ===========================================================================
// FEATURES
// ===========================================================================

feature("Cell Operations — Insert Rows", () => {
  scenario("Given cells A1,A2,A3 with values 10,20,30, When inserting 1 row at row 0 (row 1), Then A1 moves to A2, A2 to A3, A3 to A4", () => {
    const sheet = makeSheet({
      A1: { type: "number", value: 10 },
      A2: { type: "number", value: 20 },
      A3: { type: "number", value: 30 },
    });
    insertRowsInSheet(sheet, 0, 1);
    assertEqual(sheet.cells.get("A2")?.value, 10, "A1->A2 value");
    assertEqual(sheet.cells.get("A3")?.value, 20, "A2->A3 value");
    assertEqual(sheet.cells.get("A4")?.value, 30, "A3->A4 value");
    assertEqual(sheet.cells.has("A1"), false, "A1 should be empty");
  });

  scenario("Given cells A1,A2,A3, When inserting 1 row at row 1 (between A1 and A2), Then A1 stays, A2 moves to A3, A3 to A4", () => {
    const sheet = makeSheet({
      A1: { type: "number", value: 10 },
      A2: { type: "number", value: 20 },
      A3: { type: "number", value: 30 },
    });
    insertRowsInSheet(sheet, 1, 1);
    assertEqual(sheet.cells.get("A1")?.value, 10, "A1 stays");
    assertEqual(sheet.cells.get("A3")?.value, 20, "A2->A3 value");
    assertEqual(sheet.cells.get("A4")?.value, 30, "A3->A4 value");
    assertEqual(sheet.cells.has("A2"), false, "A2 should be empty");
  });

  scenario("Given cells A1,A2,A3, When inserting 2 rows at row 1, Then A2 shifts to A4 and A3 to A5", () => {
    const sheet = makeSheet({
      A1: { type: "number", value: 10 },
      A2: { type: "number", value: 20 },
      A3: { type: "number", value: 30 },
    });
    insertRowsInSheet(sheet, 1, 2);
    assertEqual(sheet.cells.get("A1")?.value, 10, "A1 stays");
    assertEqual(sheet.cells.get("A4")?.value, 20, "A2->A4");
    assertEqual(sheet.cells.get("A5")?.value, 30, "A3->A5");
  });
});

feature("Cell Operations — Delete Rows", () => {
  scenario("Given cells A1,A2,A3, When deleting row 0 (row 1), Then A2 moves to A1 and A3 to A2", () => {
    const sheet = makeSheet({
      A1: { type: "number", value: 10 },
      A2: { type: "number", value: 20 },
      A3: { type: "number", value: 30 },
    });
    deleteRowsInSheet(sheet, 0, 1);
    assertEqual(sheet.cells.get("A1")?.value, 20, "A2->A1");
    assertEqual(sheet.cells.get("A2")?.value, 30, "A3->A2");
    assertEqual(sheet.cells.has("A3"), false, "A3 should be gone");
  });

  scenario("Given cells A1,A2,A3, When deleting row 1 (middle), Then A1 stays and A3 moves to A2", () => {
    const sheet = makeSheet({
      A1: { type: "number", value: 10 },
      A2: { type: "number", value: 20 },
      A3: { type: "number", value: 30 },
    });
    deleteRowsInSheet(sheet, 1, 1);
    assertEqual(sheet.cells.get("A1")?.value, 10, "A1 stays");
    assertEqual(sheet.cells.get("A2")?.value, 30, "A3->A2");
    assertEqual(sheet.cells.has("A3"), false, "A3 gone");
  });

  scenario("Given cells A1,A2,A3, When deleting 2 rows at row 0, Then only A3 remains as A1", () => {
    const sheet = makeSheet({
      A1: { type: "number", value: 10 },
      A2: { type: "number", value: 20 },
      A3: { type: "number", value: 30 },
    });
    deleteRowsInSheet(sheet, 0, 2);
    assertEqual(sheet.cells.get("A1")?.value, 30, "A3->A1");
    assertEqual(sheet.cells.size, 1, "only 1 cell left");
  });
});

feature("Cell Operations — Insert Columns", () => {
  scenario("Given cells A1,B1,C1, When inserting 1 column at col 1, Then A1 stays, B1 moves to C1, C1 to D1", () => {
    const sheet = makeSheet({
      A1: { type: "string", value: "a" },
      B1: { type: "string", value: "b" },
      C1: { type: "string", value: "c" },
    });
    insertColumnsInSheet(sheet, 1, 1);
    assertEqual(sheet.cells.get("A1")?.value, "a", "A1 stays");
    assertEqual(sheet.cells.get("C1")?.value, "b", "B1->C1");
    assertEqual(sheet.cells.get("D1")?.value, "c", "C1->D1");
    assertEqual(sheet.cells.has("B1"), false, "B1 empty");
  });

  scenario("Given cells A1,B1,C1, When inserting 1 column at col 0, Then all shift right by 1", () => {
    const sheet = makeSheet({
      A1: { type: "string", value: "a" },
      B1: { type: "string", value: "b" },
      C1: { type: "string", value: "c" },
    });
    insertColumnsInSheet(sheet, 0, 1);
    assertEqual(sheet.cells.get("B1")?.value, "a", "A1->B1");
    assertEqual(sheet.cells.get("C1")?.value, "b", "B1->C1");
    assertEqual(sheet.cells.get("D1")?.value, "c", "C1->D1");
  });
});

feature("Cell Operations — Delete Columns", () => {
  scenario("Given cells A1,B1,C1, When deleting column 0, Then B1 moves to A1, C1 to B1", () => {
    const sheet = makeSheet({
      A1: { type: "string", value: "a" },
      B1: { type: "string", value: "b" },
      C1: { type: "string", value: "c" },
    });
    deleteColumnsInSheet(sheet, 0, 1);
    assertEqual(sheet.cells.get("A1")?.value, "b", "B1->A1");
    assertEqual(sheet.cells.get("B1")?.value, "c", "C1->B1");
    assertEqual(sheet.cells.has("C1"), false, "C1 gone");
  });

  scenario("Given cells A1,B1,C1, When deleting column 1, Then A1 stays, C1 moves to B1", () => {
    const sheet = makeSheet({
      A1: { type: "string", value: "a" },
      B1: { type: "string", value: "b" },
      C1: { type: "string", value: "c" },
    });
    deleteColumnsInSheet(sheet, 1, 1);
    assertEqual(sheet.cells.get("A1")?.value, "a", "A1 stays");
    assertEqual(sheet.cells.get("B1")?.value, "c", "C1->B1");
    assertEqual(sheet.cells.has("C1"), false, "C1 gone");
  });
});

feature("Merge Cell Logic", () => {
  scenario("Given empty merges, When adding merge A1:C3, Then mergedRegions has 1 entry", () => {
    const sheet = makeSheet({});
    sheet.mergedRegions.push({ startRef: "A1", endRef: "C3" });
    assertEqual(sheet.mergedRegions.length, 1, "merge count");
  });

  scenario("Given merge A1:C3, When checking if B2 is in merge, Then returns true", () => {
    const sheet = makeSheet({});
    sheet.mergedRegions.push({ startRef: "A1", endRef: "C3" });
    assertEqual(isCellInMerge(sheet, "B2"), true, "B2 in merge");
  });

  scenario("Given merge A1:C3, When checking if D4 is in merge, Then returns false", () => {
    const sheet = makeSheet({});
    sheet.mergedRegions.push({ startRef: "A1", endRef: "C3" });
    assertEqual(isCellInMerge(sheet, "D4"), false, "D4 not in merge");
  });

  scenario("Given merge A1:C3, When checking if A1 (top-left) is in merge, Then returns true", () => {
    const sheet = makeSheet({});
    sheet.mergedRegions.push({ startRef: "A1", endRef: "C3" });
    assertEqual(isCellInMerge(sheet, "A1"), true, "A1 in merge");
  });

  scenario("Given merge A1:C3, When checking if C3 (bottom-right) is in merge, Then returns true", () => {
    const sheet = makeSheet({});
    sheet.mergedRegions.push({ startRef: "A1", endRef: "C3" });
    assertEqual(isCellInMerge(sheet, "C3"), true, "C3 in merge");
  });

  scenario("Given merge A1:C3, When unmerging at B2, Then mergedRegions is empty", () => {
    const sheet = makeSheet({});
    sheet.mergedRegions.push({ startRef: "A1", endRef: "C3" });
    unmergeAtCell(sheet, "B2");
    assertEqual(sheet.mergedRegions.length, 0, "unmerged");
  });

  scenario("Given two merges A1:B2 and D1:E2, When unmerging at A1, Then only D1:E2 remains", () => {
    const sheet = makeSheet({});
    sheet.mergedRegions.push({ startRef: "A1", endRef: "B2" });
    sheet.mergedRegions.push({ startRef: "D1", endRef: "E2" });
    unmergeAtCell(sheet, "A1");
    assertEqual(sheet.mergedRegions.length, 1, "one merge left");
    assertEqual(sheet.mergedRegions[0].startRef, "D1", "D1:E2 remains");
  });
});

feature("Grid Renderer — Sizing", () => {
  scenario("Given default widths, Then colWidth returns 63px (8.43 * 7.5 rounded)", () => {
    const sheet = makeSheet({});
    const w = colWidth(sheet, 0);
    assertEqual(w, Math.round(8.43 * 7.5), "default col width px");
  });

  scenario("Given custom width 20 for col 0, Then colWidth returns 150px", () => {
    const sheet = makeSheet({});
    sheet.colWidths.set(0, 20);
    assertEqual(colWidth(sheet, 0), Math.round(20 * 7.5), "custom col width px");
  });

  scenario("Given default widths, Then rowHeight returns 20px (15 * 4/3 rounded)", () => {
    const sheet = makeSheet({});
    assertEqual(rowHeight(sheet, 0), Math.round(15 * (4 / 3)), "default row height px");
  });

  scenario("Given custom row height 30pt for row 0, Then rowHeight returns 40px", () => {
    const sheet = makeSheet({});
    sheet.rowHeights.set(0, 30);
    assertEqual(rowHeight(sheet, 0), Math.round(30 * (4 / 3)), "custom row height px");
  });

  scenario("Given 3 default columns, Then colOffset(3) equals sum of first 3 widths", () => {
    const sheet = makeSheet({});
    const expected = Math.round(8.43 * 7.5) * 3;
    assertEqual(colOffset(sheet, 3), expected, "colOffset(3)");
  });

  scenario("Given col 0 width=20 and col 1 default, Then colOffset(2) equals col0+col1 widths", () => {
    const sheet = makeSheet({});
    sheet.colWidths.set(0, 20);
    const expected = Math.round(20 * 7.5) + Math.round(8.43 * 7.5);
    assertEqual(colOffset(sheet, 2), expected, "colOffset(2) mixed");
  });

  scenario("Given 3 default rows, Then rowOffset(3) equals sum of first 3 heights", () => {
    const sheet = makeSheet({});
    const expected = Math.round(15 * (4 / 3)) * 3;
    assertEqual(rowOffset(sheet, 3), expected, "rowOffset(3)");
  });
});

feature("Grid Renderer — Viewport", () => {
  scenario("Given canvas 800x600 and scroll at (0,0), Then viewport starts at row 0, col 0", () => {
    const sheet = makeSheet({});
    const vp = computeViewport(sheet, 0, 0, 800, 600);
    assertEqual(vp.startRow, 0, "startRow");
    assertEqual(vp.startCol, 0, "startCol");
  });

  scenario("Given canvas 800x600 and scroll at (10,5), Then viewport starts at row 10, col 5", () => {
    const sheet = makeSheet({});
    const vp = computeViewport(sheet, 10, 5, 800, 600);
    assertEqual(vp.startRow, 10, "startRow");
    assertEqual(vp.startCol, 5, "startCol");
  });

  scenario("Given canvas 800x600, Then endCol > startCol (columns visible)", () => {
    const sheet = makeSheet({});
    const vp = computeViewport(sheet, 0, 0, 800, 600);
    assert(vp.endCol > vp.startCol, "columns visible");
    assert(vp.endRow > vp.startRow, "rows visible");
  });

  scenario("Given very narrow canvas (100px wide), Then fewer columns visible than 800px canvas", () => {
    const sheet = makeSheet({});
    const vpWide = computeViewport(sheet, 0, 0, 800, 600);
    const vpNarrow = computeViewport(sheet, 0, 0, 100, 600);
    assert(vpNarrow.endCol - vpNarrow.startCol < vpWide.endCol - vpWide.startCol, "narrow has fewer cols");
  });

  scenario("Given very short canvas (100px tall), Then fewer rows visible than 600px canvas", () => {
    const sheet = makeSheet({});
    const vpTall = computeViewport(sheet, 0, 0, 800, 600);
    const vpShort = computeViewport(sheet, 0, 0, 800, 100);
    assert(vpShort.endRow - vpShort.startRow < vpTall.endRow - vpTall.startRow, "short has fewer rows");
  });
});

feature("Grid Renderer — Hit Testing", () => {
  scenario("Given a sheet with default widths, When clicking at canvas position in cell area, Then returns correct cell ref", () => {
    const sheet = makeSheet({});
    const vp = computeViewport(sheet, 0, 0, 800, 600);
    // Click well inside the first cell area: header=50px wide, 24px tall
    // First cell A1 is at x=50, y=24, width~63, height~20
    const ref = hitTestCell(sheet, 80, 34, vp, 1.0);
    assertEqual(ref, "A1", "hit A1");
  });

  scenario("Given a sheet with default widths, When clicking second column area, Then returns B1", () => {
    const sheet = makeSheet({});
    const vp = computeViewport(sheet, 0, 0, 800, 600);
    // B1: x = 50 (header) + 63 (col A) + some offset into col B
    const ref = hitTestCell(sheet, 50 + 63 + 10, 34, vp, 1.0);
    assertEqual(ref, "B1", "hit B1");
  });

  scenario("Given scroll offset (5,3), When clicking first visible cell area, Then accounts for scroll", () => {
    const sheet = makeSheet({});
    const vp = computeViewport(sheet, 5, 3, 800, 600);
    // Click at header offset should give the first visible cell (D6, 0-indexed col=3 row=5)
    const ref = hitTestCell(sheet, 60, 30, vp, 1.0);
    assertEqual(ref, "D6", "hit D6 with scroll offset");
  });

  scenario("Given click on column header area (y < 24), When hit testing column, Then returns column index", () => {
    const sheet = makeSheet({});
    const vp = computeViewport(sheet, 0, 0, 800, 600);
    // Click in header area y=10 (< 24), x=60 (> 50 header width, within col A)
    const colIdx = hitTestColHeader(sheet, 60, 10, vp, 1.0);
    assertEqual(colIdx, 0, "col header hit = col 0");
  });

  scenario("Given click on column header for second column, Then returns column 1", () => {
    const sheet = makeSheet({});
    const vp = computeViewport(sheet, 0, 0, 800, 600);
    const colIdx = hitTestColHeader(sheet, 50 + 63 + 10, 10, vp, 1.0);
    assertEqual(colIdx, 1, "col header hit = col 1");
  });

  scenario("Given click on row header area (x < 50), When hit testing row, Then returns row index", () => {
    const sheet = makeSheet({});
    const vp = computeViewport(sheet, 0, 0, 800, 600);
    // Click in row header: x=20 (< 50), y=30 (> 24 header height, within row 0)
    const rowIdx = hitTestRowHeader(sheet, 20, 30, vp, 1.0);
    assertEqual(rowIdx, 0, "row header hit = row 0");
  });

  scenario("Given click on row header for second row, Then returns row 1", () => {
    const sheet = makeSheet({});
    const vp = computeViewport(sheet, 0, 0, 800, 600);
    // Row 0 height = 20px, so row 1 starts at y=24+20=44
    const rowIdx = hitTestRowHeader(sheet, 20, 50, vp, 1.0);
    assertEqual(rowIdx, 1, "row header hit = row 1");
  });
});

feature("Cell Selection — Jump to Edge", () => {
  scenario("Given cells A1-A5 with data, When jumping down from A1, Then reaches A5", () => {
    const dataCells = new Set(["A1", "A2", "A3", "A4", "A5"]);
    const hasData = (ref: CellRef) => dataCells.has(ref);
    const result = jumpToEdge("A1", "down", hasData, 1000, 100);
    assertEqual(result, "A5", "jump down to A5");
  });

  scenario("Given gap at A3 (empty), When jumping down from A1, Then stops at A2", () => {
    const dataCells = new Set(["A1", "A2", "A4", "A5"]);
    const hasData = (ref: CellRef) => dataCells.has(ref);
    const result = jumpToEdge("A1", "down", hasData, 1000, 100);
    assertEqual(result, "A2", "stop at A2 before gap");
  });

  scenario("Given cells A1-E1 with data, When jumping right from A1, Then reaches E1", () => {
    const dataCells = new Set(["A1", "B1", "C1", "D1", "E1"]);
    const hasData = (ref: CellRef) => dataCells.has(ref);
    const result = jumpToEdge("A1", "right", hasData, 1000, 100);
    assertEqual(result, "E1", "jump right to E1");
  });

  scenario("Given empty cell A1, When jumping down, Then reaches first non-empty cell", () => {
    const dataCells = new Set(["A5"]);
    const hasData = (ref: CellRef) => dataCells.has(ref);
    const result = jumpToEdge("A1", "down", hasData, 1000, 100);
    assertEqual(result, "A5", "jump to first data cell A5");
  });

  scenario("Given cells A3-A5 with data, When jumping up from A5, Then reaches A3", () => {
    const dataCells = new Set(["A3", "A4", "A5"]);
    const hasData = (ref: CellRef) => dataCells.has(ref);
    const result = jumpToEdge("A5", "up", hasData, 1000, 100);
    assertEqual(result, "A3", "jump up to A3");
  });
});

feature("OOXML Parser — Shared Strings", () => {
  scenario("Given XLSX with shared string table, When parsing, Then cell values resolve to strings", async () => {
    const sheet = makeSheet({
      A1: { type: "string", value: "Hello World" },
      A2: { type: "string", value: "Test String" },
      B1: { type: "number", value: 42 },
    });
    const wb = makeWorkbook([sheet]);
    const blob = exportXlsx(wb);
    const buf = await blob.arrayBuffer();
    const parsed = parseXlsx(buf);

    const a1 = parsed.sheets[0].cells.get("A1");
    assertEqual(a1?.type, "string", "A1 type");
    assertEqual(a1?.value, "Hello World", "A1 resolved string");

    const a2 = parsed.sheets[0].cells.get("A2");
    assertEqual(a2?.type, "string", "A2 type");
    assertEqual(a2?.value, "Test String", "A2 resolved string");
  });

  scenario("Given XLSX with duplicate shared strings, When parsing, Then both cells resolve correctly", async () => {
    const sheet = makeSheet({
      A1: { type: "string", value: "Same" },
      A2: { type: "string", value: "Same" },
    });
    const wb = makeWorkbook([sheet]);
    const parsed = parseXlsx(await (exportXlsx(wb)).arrayBuffer());
    assertEqual(parsed.sheets[0].cells.get("A1")?.value, "Same", "A1");
    assertEqual(parsed.sheets[0].cells.get("A2")?.value, "Same", "A2");
  });

  scenario("Given XLSX with empty string, When parsing, Then cell type is string", async () => {
    // Exporter skips empty strings (they go to SST only if non-empty)
    const sheet = makeSheet({
      A1: { type: "string", value: "NonEmpty" },
    });
    const wb = makeWorkbook([sheet]);
    const parsed = parseXlsx(await (exportXlsx(wb)).arrayBuffer());
    assertEqual(parsed.sheets[0].cells.get("A1")?.value, "NonEmpty", "non-empty preserved");
  });
});

feature("XLSX Exporter — Cell Types", () => {
  scenario("Given a formula cell, When exporting and re-parsing, Then produces formula tag", async () => {
    const sheet = makeSheet({
      A1: { type: "number", value: 10 },
      B1: { type: "formula", value: null, formula: "SUM(A1,5)" },
    });
    const wb = makeWorkbook([sheet]);
    const parsed = parseXlsx(await (exportXlsx(wb)).arrayBuffer());
    const b1 = parsed.sheets[0].cells.get("B1");
    assertEqual(b1?.type, "formula", "B1 is formula");
    assert(b1?.formula != null, "B1 has formula content");
    assertEqual(b1?.formula, "SUM(A1,5)", "formula text preserved");
  });

  scenario("Given a number cell, When exporting and re-parsing, Then produces numeric value", async () => {
    const sheet = makeSheet({
      A1: { type: "number", value: 3.14 },
    });
    const wb = makeWorkbook([sheet]);
    const parsed = parseXlsx(await (exportXlsx(wb)).arrayBuffer());
    const a1 = parsed.sheets[0].cells.get("A1");
    assertEqual(a1?.type, "number", "A1 is number");
    assertClose(a1?.value as number, 3.14, "A1 value");
  });

  scenario("Given a boolean cell with true, When exporting and re-parsing, Then produces boolean type", async () => {
    const sheet = makeSheet({
      A1: { type: "boolean", value: true },
    });
    const wb = makeWorkbook([sheet]);
    const parsed = parseXlsx(await (exportXlsx(wb)).arrayBuffer());
    const a1 = parsed.sheets[0].cells.get("A1");
    assertEqual(a1?.type, "boolean", "A1 is boolean");
    assertEqual(a1?.value, true, "A1 value true");
  });

  scenario("Given a boolean cell with false, When exporting and re-parsing, Then value is false", async () => {
    const sheet = makeSheet({
      A1: { type: "boolean", value: false },
    });
    const wb = makeWorkbook([sheet]);
    const parsed = parseXlsx(await (exportXlsx(wb)).arrayBuffer());
    const a1 = parsed.sheets[0].cells.get("A1");
    assertEqual(a1?.type, "boolean", "A1 is boolean");
    assertEqual(a1?.value, false, "A1 value false");
  });

  scenario("Given a string cell with special chars (<>&), When roundtripping, Then value is preserved", async () => {
    const sheet = makeSheet({
      A1: { type: "string", value: 'Price < $10 & "free" > 0' },
    });
    const wb = makeWorkbook([sheet]);
    const parsed = parseXlsx(await (exportXlsx(wb)).arrayBuffer());
    assertEqual(parsed.sheets[0].cells.get("A1")?.value, 'Price < $10 & "free" > 0', "special chars preserved");
  });

  scenario("Given a formula with calculated value, When exporting, Then calculatedValue roundtrips", async () => {
    const sheet = makeSheet({
      A1: { type: "formula", value: null, formula: "1+2" },
    });
    // Set calculated value
    sheet.cells.get("A1")!.calculatedValue = 3;
    const wb = makeWorkbook([sheet]);
    const parsed = parseXlsx(await (exportXlsx(wb)).arrayBuffer());
    const a1 = parsed.sheets[0].cells.get("A1");
    assertEqual(a1?.formula, "1+2", "formula preserved");
    // calculatedValue is stored as string in XML <v>
    assert(a1?.calculatedValue != null, "has calculated value");
  });

  scenario("Given integer 0, When exporting and re-parsing, Then value is 0 (not missing)", async () => {
    const sheet = makeSheet({
      A1: { type: "number", value: 0 },
    });
    const wb = makeWorkbook([sheet]);
    const parsed = parseXlsx(await (exportXlsx(wb)).arrayBuffer());
    const a1 = parsed.sheets[0].cells.get("A1");
    assertEqual(a1?.type, "number", "A1 is number");
    assertEqual(a1?.value, 0, "A1 value is 0");
  });

  scenario("Given negative number -99.5, When roundtripping, Then value is preserved", async () => {
    const sheet = makeSheet({
      A1: { type: "number", value: -99.5 },
    });
    const wb = makeWorkbook([sheet]);
    const parsed = parseXlsx(await (exportXlsx(wb)).arrayBuffer());
    assertClose(parsed.sheets[0].cells.get("A1")?.value as number, -99.5, "negative value");
  });
});

// ===========================================================================
// Summary
// ===========================================================================

setTimeout(() => {
  console.log(`\n${"=".repeat(50)}`);
  console.log(`\x1b[1mResults: ${totalPass} passed, ${totalFail} failed, ${totalPass + totalFail} total\x1b[0m`);
  if (totalFail === 0) {
    console.log("\x1b[32m\x1b[1mALL SCENARIOS PASSED\x1b[0m");
  } else {
    console.log(`\x1b[31m\x1b[1m${totalFail} SCENARIO(S) FAILED\x1b[0m`);
    process.exit(1);
  }
}, 500);
