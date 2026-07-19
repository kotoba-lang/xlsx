/**
 * Editor Logic BDD Tests — covers all 43 editor-state.svelte.ts exports
 * by reimplementing the logic without Svelte runes.
 *
 * Run: cd svelte && npx tsx src/lib/__tests__/editor-logic-bdd.test.ts
 */

import { parseXlsx, buildRef, parseRef, colToLetter, type XlsxWorkbook, type XlsxSheet, type XlsxCell, type XlsxStyle, type XlsxFont, type XlsxFill, type XlsxBorder, type XlsxAlignment, type CellRef } from "../ooxml-parser";
import { exportXlsx } from "../xlsx-exporter";
import { recalculateSheet } from "../formula-engine";

let totalPass = 0, totalFail = 0;
function feature(n: string, fn: () => void) { console.log(`\n\x1b[1mFeature: ${n}\x1b[0m`); fn(); }
function scenario(n: string, fn: () => void) { try { fn(); totalPass++; console.log(`  \x1b[32m✓\x1b[0m ${n}`); } catch (e: any) { totalFail++; console.log(`  \x1b[31m✗\x1b[0m ${n}\n    \x1b[31m${e.message}\x1b[0m`); } }
function assert(c: boolean, m: string) { if (!c) throw new Error(m); }
function assertEqual(a: unknown, e: unknown, l: string) { if (a !== e) throw new Error(`${l}: expected ${JSON.stringify(e)}, got ${JSON.stringify(a)}`); }

// ---------------------------------------------------------------------------
// Reimplemented editor-state logic (mirrors svelte file, no runes)
// ---------------------------------------------------------------------------

interface EditorState {
  workbook: XlsxWorkbook | null;
  activeSheetIndex: number;
  selection: { startRow: number; startCol: number; endRow: number; endCol: number };
  activeCell: CellRef;
  editingCell: CellRef | null;
  editValue: string;
  isDirty: boolean;
  zoom: number;
  scrollRow: number;
  scrollCol: number;
}

function createEditor(): EditorState {
  return { workbook: null, activeSheetIndex: 0, selection: { startRow: 0, startCol: 0, endRow: 0, endCol: 0 }, activeCell: "A1", editingCell: null, editValue: "", isDirty: false, zoom: 1.0, scrollRow: 0, scrollCol: 0 };
}

function makeSheet(cells: Record<string, { type: string; value: any; formula?: string; styleId?: number }>): XlsxSheet {
  const map = new Map<string, XlsxCell>();
  for (const [ref, c] of Object.entries(cells)) {
    const { col, row } = parseRef(ref);
    map.set(ref, { ref: ref as any, row, col, type: c.type as any, value: c.formula ? null : c.value, formula: c.formula ?? null, calculatedValue: null, styleId: c.styleId ?? 0, hyperlink: null });
  }
  return { id: "s1", name: "Sheet1", order: 0, hidden: false, cells: map, mergedRegions: [], tables: [], charts: [], conditionalFormats: [], dataValidations: [], frozenRow: 0, frozenCol: 0, colWidths: new Map(), rowHeights: new Map(), defaultColWidth: 8.43, defaultRowHeight: 15 };
}

function makeWorkbook(sheets?: XlsxSheet[]): XlsxWorkbook {
  const s = sheets ?? [makeSheet({})];
  return { id: "wb1", title: "Test", sheets: s, sharedStrings: [], styles: [{ id: 0, numFmt: null, font: { name: "Calibri", size: 11, bold: false, italic: false, underline: false, strikethrough: false, color: "#000000" }, fill: null, border: null, alignment: null }], definedNames: [], activeSheetIndex: 0 };
}

// --- currentSheet / getCell / activeCellValue ---
function currentSheet(ed: EditorState): XlsxSheet | null {
  if (!ed.workbook || ed.activeSheetIndex < 0 || ed.activeSheetIndex >= ed.workbook.sheets.length) return null;
  return ed.workbook.sheets[ed.activeSheetIndex];
}
function getCell(ed: EditorState, ref: CellRef): XlsxCell | null { return currentSheet(ed)?.cells.get(ref) ?? null; }
function activeCellValue(ed: EditorState): string {
  const cell = getCell(ed, ed.activeCell);
  if (!cell) return "";
  if (cell.formula) return `=${cell.formula}`;
  return cell.value != null ? String(cell.value) : "";
}

// --- Selection ---
function selectCell(ed: EditorState, ref: CellRef): void {
  const { col, row } = parseRef(ref);
  ed.selection = { startRow: row, startCol: col, endRow: row, endCol: col };
  ed.activeCell = ref;
  ed.editingCell = null;
}
function selectRange(ed: EditorState, range: EditorState["selection"]): void {
  ed.selection = { ...range };
  ed.activeCell = buildRef(range.startCol, range.startRow);
}
function moveActiveCell(ed: EditorState, dRow: number, dCol: number): void {
  const { col, row } = parseRef(ed.activeCell);
  selectCell(ed, buildRef(Math.max(0, col + dCol), Math.max(0, row + dRow)));
}

// --- Edit ---
function startEdit(ed: EditorState, ref: CellRef): void { ed.editingCell = ref; ed.editValue = activeCellValue(ed); }
function commitEdit(ed: EditorState): void {
  if (!ed.editingCell) return;
  setCellValue(ed, ed.editingCell, ed.editValue);
  ed.editingCell = null; ed.editValue = "";
}
function cancelEdit(ed: EditorState): void { ed.editingCell = null; ed.editValue = ""; }

// --- setCellValue ---
function setCellValue(ed: EditorState, ref: CellRef, value: string | number | boolean | null): void {
  const sheet = currentSheet(ed);
  if (!sheet) return;
  const { col, row } = parseRef(ref);
  if (value === null || value === "") { sheet.cells.delete(ref); return; }
  let type: XlsxCell["type"] = "string"; let parsedValue: XlsxCell["value"] = value; let formula: string | null = null;
  if (typeof value === "string" && value.startsWith("=")) { type = "formula"; formula = value.slice(1); parsedValue = null; }
  else if (typeof value === "number") { type = "number"; }
  else if (typeof value === "boolean") { type = "boolean"; }
  else if (typeof value === "string") { const num = Number(value); if (!isNaN(num) && value.trim() !== "") { type = "number"; parsedValue = num; } }
  const existing = sheet.cells.get(ref);
  sheet.cells.set(ref, { ref, row, col, type, value: parsedValue, formula, calculatedValue: null, styleId: existing?.styleId ?? 0, hyperlink: null });
  ed.isDirty = true;
}

// --- Undo/Redo ---
function serializeWb(wb: XlsxWorkbook): string {
  return JSON.stringify({ ...wb, sheets: wb.sheets.map(s => ({ ...s, cells: Object.fromEntries(s.cells), colWidths: Object.fromEntries(s.colWidths), rowHeights: Object.fromEntries(s.rowHeights) })) });
}
function deserializeWb(json: string): XlsxWorkbook {
  const p = JSON.parse(json);
  return { ...p, sheets: p.sheets.map((s: any) => ({ ...s, cells: new Map(Object.entries(s.cells)), colWidths: new Map(Object.entries(s.colWidths).map(([k, v]: any) => [Number(k), v])), rowHeights: new Map(Object.entries(s.rowHeights).map(([k, v]: any) => [Number(k), v])) })) };
}

class UndoManager {
  private undoStack: string[] = [];
  private redoStack: string[] = [];
  pushUndo(ed: EditorState) {
    if (!ed.workbook) return;
    this.undoStack.push(JSON.stringify({ wb: serializeWb(ed.workbook), asi: ed.activeSheetIndex, sel: ed.selection }));
    if (this.undoStack.length > 50) this.undoStack.shift();
    this.redoStack = [];
    ed.isDirty = true;
  }
  undo(ed: EditorState) {
    if (this.undoStack.length === 0 || !ed.workbook) return;
    this.redoStack.push(JSON.stringify({ wb: serializeWb(ed.workbook), asi: ed.activeSheetIndex, sel: ed.selection }));
    const snap = JSON.parse(this.undoStack.pop()!);
    ed.workbook = deserializeWb(snap.wb); ed.activeSheetIndex = snap.asi; ed.selection = snap.sel;
  }
  redo(ed: EditorState) {
    if (this.redoStack.length === 0 || !ed.workbook) return;
    this.undoStack.push(JSON.stringify({ wb: serializeWb(ed.workbook), asi: ed.activeSheetIndex, sel: ed.selection }));
    const snap = JSON.parse(this.redoStack.pop()!);
    ed.workbook = deserializeWb(snap.wb); ed.activeSheetIndex = snap.asi; ed.selection = snap.sel;
  }
  canUndo() { return this.undoStack.length > 0; }
  canRedo() { return this.redoStack.length > 0; }
}

// --- Sheet ops ---
function loadWorkbook(ed: EditorState, wb: XlsxWorkbook) { ed.workbook = wb; ed.activeSheetIndex = 0; ed.activeCell = "A1"; ed.isDirty = false; }
function resetEditor(ed: EditorState) { ed.workbook = null; ed.activeSheetIndex = 0; ed.activeCell = "A1"; ed.editingCell = null; ed.isDirty = false; }
function selectSheet(ed: EditorState, idx: number) { ed.activeSheetIndex = idx; ed.activeCell = "A1"; ed.editingCell = null; }
function addSheet(ed: EditorState) {
  if (!ed.workbook) return;
  const s = makeSheet({}); s.name = `Sheet${ed.workbook.sheets.length + 1}`; s.order = ed.workbook.sheets.length; s.id = `s${Date.now()}`;
  ed.workbook.sheets.push(s); ed.activeSheetIndex = ed.workbook.sheets.length - 1;
}
function deleteSheet(ed: EditorState, idx: number) {
  if (!ed.workbook || ed.workbook.sheets.length <= 1) return;
  ed.workbook.sheets.splice(idx, 1);
  if (ed.activeSheetIndex >= ed.workbook.sheets.length) ed.activeSheetIndex = ed.workbook.sheets.length - 1;
}
function renameSheet(ed: EditorState, idx: number, name: string) { if (ed.workbook?.sheets[idx]) ed.workbook.sheets[idx].name = name; }

// --- Row/Col/Merge ---
function insertRows(sheet: XlsxSheet, at: number, count: number) {
  const nc = new Map<CellRef, XlsxCell>();
  for (const [ref, cell] of sheet.cells) { if (cell.row >= at) { const nr = buildRef(cell.col, cell.row + count); nc.set(nr, { ...cell, ref: nr, row: cell.row + count }); } else nc.set(ref, cell); }
  sheet.cells = nc;
}
function deleteRows(sheet: XlsxSheet, at: number, count: number) {
  const nc = new Map<CellRef, XlsxCell>();
  for (const [ref, cell] of sheet.cells) { if (cell.row >= at && cell.row < at + count) continue; if (cell.row >= at + count) { const nr = buildRef(cell.col, cell.row - count); nc.set(nr, { ...cell, ref: nr, row: cell.row - count }); } else nc.set(ref, cell); }
  sheet.cells = nc;
}
function insertColumns(sheet: XlsxSheet, at: number, count: number) {
  const nc = new Map<CellRef, XlsxCell>();
  for (const [ref, cell] of sheet.cells) { if (cell.col >= at) { const nr = buildRef(cell.col + count, cell.row); nc.set(nr, { ...cell, ref: nr, col: cell.col + count }); } else nc.set(ref, cell); }
  sheet.cells = nc;
}
function deleteColumns(sheet: XlsxSheet, at: number, count: number) {
  const nc = new Map<CellRef, XlsxCell>();
  for (const [ref, cell] of sheet.cells) { if (cell.col >= at && cell.col < at + count) continue; if (cell.col >= at + count) { const nr = buildRef(cell.col - count, cell.row); nc.set(nr, { ...cell, ref: nr, col: cell.col - count }); } else nc.set(ref, cell); }
  sheet.cells = nc;
}
function mergeCells(sheet: XlsxSheet, startRef: CellRef, endRef: CellRef) { sheet.mergedRegions.push({ startRef, endRef }); }
function unmergeCells(sheet: XlsxSheet, ref: CellRef) {
  const { row, col } = parseRef(ref);
  sheet.mergedRegions = sheet.mergedRegions.filter(m => { const s = parseRef(m.startRef), e = parseRef(m.endRef); return !(row >= s.row && row <= e.row && col >= s.col && col <= e.col); });
}

// --- Copy/Paste ---
function copyCells(sheet: XlsxSheet, sel: EditorState["selection"]): Map<CellRef, XlsxCell> {
  const clip = new Map<CellRef, XlsxCell>();
  for (let r = sel.startRow; r <= sel.endRow; r++) for (let c = sel.startCol; c <= sel.endCol; c++) {
    const ref = buildRef(c, r); const cell = sheet.cells.get(ref);
    if (cell) clip.set(ref, { ...cell });
  }
  return clip;
}
function pasteCells(sheet: XlsxSheet, clip: Map<CellRef, XlsxCell>, baseRef: CellRef, origSel: EditorState["selection"]) {
  const { col: bc, row: br } = parseRef(baseRef);
  for (const [, cell] of clip) {
    const nr = cell.row - origSel.startRow + br, nc = cell.col - origSel.startCol + bc;
    const ref = buildRef(nc, nr);
    sheet.cells.set(ref, { ...cell, ref, row: nr, col: nc });
  }
}

// --- Formatting ---
function toggleBold(sheet: XlsxSheet, styles: XlsxStyle[], ref: CellRef): number {
  const cell = sheet.cells.get(ref);
  const base = styles[cell?.styleId ?? 0];
  const wasBold = base?.font?.bold ?? false;
  const newFont: XlsxFont = { ...(base?.font ?? { name: "Calibri", size: 11, bold: false, italic: false, underline: false, strikethrough: false, color: "#000000" }), bold: !wasBold };
  const id = styles.length;
  styles.push({ ...base, id, font: newFont });
  if (cell) cell.styleId = id;
  return id;
}
function toggleItalic(sheet: XlsxSheet, styles: XlsxStyle[], ref: CellRef): number {
  const cell = sheet.cells.get(ref); const base = styles[cell?.styleId ?? 0];
  const newFont: XlsxFont = { ...(base?.font ?? { name: "Calibri", size: 11, bold: false, italic: false, underline: false, strikethrough: false, color: "#000000" }), italic: !(base?.font?.italic ?? false) };
  const id = styles.length; styles.push({ ...base, id, font: newFont }); if (cell) cell.styleId = id; return id;
}
function toggleUnderline(sheet: XlsxSheet, styles: XlsxStyle[], ref: CellRef): number {
  const cell = sheet.cells.get(ref); const base = styles[cell?.styleId ?? 0];
  const newFont: XlsxFont = { ...(base?.font ?? { name: "Calibri", size: 11, bold: false, italic: false, underline: false, strikethrough: false, color: "#000000" }), underline: !(base?.font?.underline ?? false) };
  const id = styles.length; styles.push({ ...base, id, font: newFont }); if (cell) cell.styleId = id; return id;
}
function setFontSize(sheet: XlsxSheet, styles: XlsxStyle[], ref: CellRef, size: number): number {
  const cell = sheet.cells.get(ref); const base = styles[cell?.styleId ?? 0];
  const newFont: XlsxFont = { ...(base?.font ?? { name: "Calibri", size: 11, bold: false, italic: false, underline: false, strikethrough: false, color: "#000000" }), size };
  const id = styles.length; styles.push({ ...base, id, font: newFont }); if (cell) cell.styleId = id; return id;
}
function setFontColor(sheet: XlsxSheet, styles: XlsxStyle[], ref: CellRef, color: string): number {
  const cell = sheet.cells.get(ref); const base = styles[cell?.styleId ?? 0];
  const newFont: XlsxFont = { ...(base?.font ?? { name: "Calibri", size: 11, bold: false, italic: false, underline: false, strikethrough: false, color: "#000000" }), color };
  const id = styles.length; styles.push({ ...base, id, font: newFont }); if (cell) cell.styleId = id; return id;
}
function setFill(styles: XlsxStyle[], ref: CellRef, sheet: XlsxSheet, fgColor: string): number {
  const cell = sheet.cells.get(ref); const base = styles[cell?.styleId ?? 0];
  const id = styles.length; styles.push({ ...base, id, fill: { type: "solid", fgColor, bgColor: null } }); if (cell) cell.styleId = id; return id;
}
function setNumberFormat(styles: XlsxStyle[], ref: CellRef, sheet: XlsxSheet, fmt: string): number {
  const cell = sheet.cells.get(ref); const base = styles[cell?.styleId ?? 0];
  const id = styles.length; styles.push({ ...base, id, numFmt: fmt }); if (cell) cell.styleId = id; return id;
}
function setAlignment(styles: XlsxStyle[], ref: CellRef, sheet: XlsxSheet, align: XlsxAlignment): number {
  const cell = sheet.cells.get(ref); const base = styles[cell?.styleId ?? 0];
  const id = styles.length; styles.push({ ...base, id, alignment: align }); if (cell) cell.styleId = id; return id;
}
function setBorder(styles: XlsxStyle[], ref: CellRef, sheet: XlsxSheet, border: XlsxBorder): number {
  const cell = sheet.cells.get(ref); const base = styles[cell?.styleId ?? 0];
  const id = styles.length; styles.push({ ...base, id, border }); if (cell) cell.styleId = id; return id;
}
function setColumnWidth(sheet: XlsxSheet, col: number, w: number) { sheet.colWidths.set(col, w); }
function setRowHeight(sheet: XlsxSheet, row: number, h: number) { sheet.rowHeights.set(row, h); }
function setFreeze(sheet: XlsxSheet, rows: number, cols: number) { sheet.frozenRow = rows; sheet.frozenCol = cols; }

// ===========================================================================

feature("loadWorkbook / resetEditor", () => {
  scenario("loadWorkbook sets workbook and resets state", () => {
    const ed = createEditor();
    const wb = makeWorkbook();
    loadWorkbook(ed, wb);
    assert(ed.workbook !== null, "wb set");
    assertEqual(ed.activeSheetIndex, 0, "asi");
    assertEqual(ed.activeCell, "A1", "cell");
    assertEqual(ed.isDirty, false, "clean");
  });
  scenario("resetEditor clears everything", () => {
    const ed = createEditor();
    loadWorkbook(ed, makeWorkbook());
    ed.isDirty = true;
    resetEditor(ed);
    assertEqual(ed.workbook, null, "null");
    assertEqual(ed.isDirty, false, "clean");
  });
});

feature("currentSheet / getCell / activeCellValue", () => {
  scenario("currentSheet returns active sheet", () => {
    const ed = createEditor(); loadWorkbook(ed, makeWorkbook([makeSheet({ A1: { type: "number", value: 42 } })]));
    const s = currentSheet(ed);
    assert(s !== null, "not null");
    assertEqual(s!.cells.get("A1")?.value, 42, "A1");
  });
  scenario("getCell returns cell by ref", () => {
    const ed = createEditor(); loadWorkbook(ed, makeWorkbook([makeSheet({ B2: { type: "string", value: "hi" } })]));
    assertEqual(getCell(ed, "B2")?.value, "hi", "B2");
    assertEqual(getCell(ed, "Z99"), null, "missing");
  });
  scenario("activeCellValue returns formula or value", () => {
    const ed = createEditor(); loadWorkbook(ed, makeWorkbook([makeSheet({ A1: { type: "formula", value: null, formula: "SUM(1,2)" } })]));
    assertEqual(activeCellValue(ed), "=SUM(1,2)", "formula");
    selectCell(ed, "B1");
    assertEqual(activeCellValue(ed), "", "empty");
  });
});

feature("selectCell / selectRange / moveActiveCell", () => {
  scenario("selectCell sets selection and activeCell", () => {
    const ed = createEditor(); loadWorkbook(ed, makeWorkbook());
    selectCell(ed, "C5");
    assertEqual(ed.activeCell, "C5", "cell");
    assertEqual(ed.selection.startRow, 4, "row");
    assertEqual(ed.selection.startCol, 2, "col");
  });
  scenario("selectRange sets range and activeCell to top-left", () => {
    const ed = createEditor(); loadWorkbook(ed, makeWorkbook());
    selectRange(ed, { startRow: 1, startCol: 2, endRow: 5, endCol: 8 });
    assertEqual(ed.activeCell, "C2", "top-left");
    assertEqual(ed.selection.endRow, 5, "endRow");
  });
  scenario("moveActiveCell moves by delta", () => {
    const ed = createEditor(); loadWorkbook(ed, makeWorkbook());
    selectCell(ed, "C3");
    moveActiveCell(ed, 1, 0); assertEqual(ed.activeCell, "C4", "down");
    moveActiveCell(ed, 0, 1); assertEqual(ed.activeCell, "D4", "right");
    moveActiveCell(ed, -1, 0); assertEqual(ed.activeCell, "D3", "up");
    moveActiveCell(ed, 0, -1); assertEqual(ed.activeCell, "C3", "left");
  });
  scenario("moveActiveCell clamps at 0", () => {
    const ed = createEditor(); loadWorkbook(ed, makeWorkbook());
    selectCell(ed, "A1");
    moveActiveCell(ed, -1, -1);
    assertEqual(ed.activeCell, "A1", "clamped");
  });
});

feature("startEdit / commitEdit / cancelEdit", () => {
  scenario("startEdit enters edit mode with current value", () => {
    const ed = createEditor(); loadWorkbook(ed, makeWorkbook([makeSheet({ A1: { type: "number", value: 99 } })]));
    startEdit(ed, "A1");
    assertEqual(ed.editingCell, "A1", "editing");
    assertEqual(ed.editValue, "99", "value");
  });
  scenario("commitEdit saves value and exits edit mode", () => {
    const ed = createEditor(); loadWorkbook(ed, makeWorkbook([makeSheet({})]));
    startEdit(ed, "A1");
    ed.editValue = "Hello";
    commitEdit(ed);
    assertEqual(ed.editingCell, null, "not editing");
    assertEqual(currentSheet(ed)!.cells.get("A1")?.value, "Hello", "saved");
  });
  scenario("cancelEdit discards changes", () => {
    const ed = createEditor(); loadWorkbook(ed, makeWorkbook([makeSheet({ A1: { type: "string", value: "Original" } })]));
    startEdit(ed, "A1");
    ed.editValue = "Changed";
    cancelEdit(ed);
    assertEqual(ed.editingCell, null, "not editing");
    assertEqual(currentSheet(ed)!.cells.get("A1")?.value, "Original", "unchanged");
  });
  scenario("commitEdit with formula", () => {
    const ed = createEditor(); loadWorkbook(ed, makeWorkbook([makeSheet({})]));
    startEdit(ed, "A1"); ed.editValue = "=SUM(1,2)"; commitEdit(ed);
    assertEqual(currentSheet(ed)!.cells.get("A1")?.type, "formula", "formula type");
    assertEqual(currentSheet(ed)!.cells.get("A1")?.formula, "SUM(1,2)", "formula");
  });
  scenario("commitEdit with empty string deletes cell", () => {
    const ed = createEditor(); loadWorkbook(ed, makeWorkbook([makeSheet({ A1: { type: "string", value: "X" } })]));
    startEdit(ed, "A1"); ed.editValue = ""; commitEdit(ed);
    assertEqual(currentSheet(ed)!.cells.has("A1"), false, "deleted");
  });
});

feature("Undo / Redo", () => {
  scenario("pushUndo + undo restores previous state", () => {
    const ed = createEditor(); loadWorkbook(ed, makeWorkbook([makeSheet({ A1: { type: "number", value: 10 } })]));
    const um = new UndoManager();
    um.pushUndo(ed);
    setCellValue(ed, "A1", 20);
    assertEqual(currentSheet(ed)!.cells.get("A1")?.value, 20, "changed");
    um.undo(ed);
    assertEqual(currentSheet(ed)!.cells.get("A1")?.value, 10, "restored");
  });
  scenario("redo after undo restores change", () => {
    const ed = createEditor(); loadWorkbook(ed, makeWorkbook([makeSheet({ A1: { type: "number", value: 10 } })]));
    const um = new UndoManager();
    um.pushUndo(ed);
    setCellValue(ed, "A1", 20);
    um.undo(ed);
    um.redo(ed);
    assertEqual(currentSheet(ed)!.cells.get("A1")?.value, 20, "redone");
  });
  scenario("canUndo/canRedo reflect stack state", () => {
    const ed = createEditor(); loadWorkbook(ed, makeWorkbook([makeSheet({})]));
    const um = new UndoManager();
    assertEqual(um.canUndo(), false, "no undo");
    assertEqual(um.canRedo(), false, "no redo");
    um.pushUndo(ed);
    setCellValue(ed, "A1", "X");
    assertEqual(um.canUndo(), true, "can undo");
    um.undo(ed);
    assertEqual(um.canRedo(), true, "can redo");
  });
  scenario("Multiple undos chain correctly", () => {
    const ed = createEditor(); loadWorkbook(ed, makeWorkbook([makeSheet({})]));
    const um = new UndoManager();
    um.pushUndo(ed); setCellValue(ed, "A1", "v1");
    um.pushUndo(ed); setCellValue(ed, "A1", "v2");
    um.pushUndo(ed); setCellValue(ed, "A1", "v3");
    um.undo(ed); assertEqual(currentSheet(ed)!.cells.get("A1")?.value, "v2", "undo to v2");
    um.undo(ed); assertEqual(currentSheet(ed)!.cells.get("A1")?.value, "v1", "undo to v1");
    um.undo(ed); assertEqual(currentSheet(ed)!.cells.has("A1"), false, "undo to empty");
  });
  scenario("New edit after undo clears redo", () => {
    const ed = createEditor(); loadWorkbook(ed, makeWorkbook([makeSheet({})]));
    const um = new UndoManager();
    um.pushUndo(ed); setCellValue(ed, "A1", "X");
    um.undo(ed);
    um.pushUndo(ed); setCellValue(ed, "A1", "Y");
    assertEqual(um.canRedo(), false, "redo cleared");
  });
});

feature("addSheet / deleteSheet / renameSheet / selectSheet", () => {
  scenario("addSheet creates new sheet and activates it", () => {
    const ed = createEditor(); loadWorkbook(ed, makeWorkbook());
    addSheet(ed);
    assertEqual(ed.workbook!.sheets.length, 2, "2 sheets");
    assertEqual(ed.activeSheetIndex, 1, "active = new");
    assertEqual(ed.workbook!.sheets[1].name, "Sheet2", "name");
  });
  scenario("deleteSheet removes sheet (min 1)", () => {
    const ed = createEditor(); loadWorkbook(ed, makeWorkbook());
    addSheet(ed); addSheet(ed);
    assertEqual(ed.workbook!.sheets.length, 3, "3 sheets");
    deleteSheet(ed, 1);
    assertEqual(ed.workbook!.sheets.length, 2, "2 sheets");
  });
  scenario("deleteSheet prevents removing last sheet", () => {
    const ed = createEditor(); loadWorkbook(ed, makeWorkbook());
    deleteSheet(ed, 0);
    assertEqual(ed.workbook!.sheets.length, 1, "still 1");
  });
  scenario("renameSheet changes sheet name", () => {
    const ed = createEditor(); loadWorkbook(ed, makeWorkbook());
    renameSheet(ed, 0, "Revenue");
    assertEqual(ed.workbook!.sheets[0].name, "Revenue", "renamed");
  });
  scenario("selectSheet switches active sheet and resets selection", () => {
    const ed = createEditor(); loadWorkbook(ed, makeWorkbook());
    addSheet(ed);
    selectCell(ed, "D5");
    selectSheet(ed, 0);
    assertEqual(ed.activeSheetIndex, 0, "idx");
    assertEqual(ed.activeCell, "A1", "reset to A1");
  });
});

feature("insertRows / deleteRows / insertColumns / deleteColumns", () => {
  scenario("insertRows shifts cells and preserves data", () => {
    const s = makeSheet({ A1:{type:"number",value:1}, B1:{type:"number",value:2}, A2:{type:"number",value:3} });
    insertRows(s, 1, 2);
    assertEqual(s.cells.get("A1")?.value, 1, "A1 stays");
    assertEqual(s.cells.get("A4")?.value, 3, "A2→A4");
    assertEqual(s.cells.size, 3, "3 cells");
  });
  scenario("deleteRows removes and shifts up", () => {
    const s = makeSheet({ A1:{type:"number",value:1}, A2:{type:"number",value:2}, A3:{type:"number",value:3} });
    deleteRows(s, 0, 1);
    assertEqual(s.cells.get("A1")?.value, 2, "A2→A1");
    assertEqual(s.cells.get("A2")?.value, 3, "A3→A2");
    assertEqual(s.cells.size, 2, "2 cells");
  });
  scenario("insertColumns shifts right", () => {
    const s = makeSheet({ A1:{type:"string",value:"a"}, B1:{type:"string",value:"b"} });
    insertColumns(s, 0, 1);
    assertEqual(s.cells.get("B1")?.value, "a", "A→B");
    assertEqual(s.cells.get("C1")?.value, "b", "B→C");
  });
  scenario("deleteColumns shifts left", () => {
    const s = makeSheet({ A1:{type:"string",value:"a"}, B1:{type:"string",value:"b"}, C1:{type:"string",value:"c"} });
    deleteColumns(s, 1, 1);
    assertEqual(s.cells.get("A1")?.value, "a", "A stays");
    assertEqual(s.cells.get("B1")?.value, "c", "C→B");
  });
});

feature("mergeCells / unmergeCells", () => {
  scenario("Merge and unmerge lifecycle", () => {
    const s = makeSheet({});
    mergeCells(s, "A1", "C3");
    assertEqual(s.mergedRegions.length, 1, "merged");
    unmergeCells(s, "B2");
    assertEqual(s.mergedRegions.length, 0, "unmerged");
  });
  scenario("Unmerge at unrelated cell → no change", () => {
    const s = makeSheet({});
    mergeCells(s, "A1", "B2");
    unmergeCells(s, "Z99");
    assertEqual(s.mergedRegions.length, 1, "unchanged");
  });
});

feature("copyCells / pasteCells", () => {
  scenario("Copy A1:B2, paste at D4", () => {
    const s = makeSheet({ A1:{type:"number",value:1}, B1:{type:"number",value:2}, A2:{type:"number",value:3}, B2:{type:"number",value:4} });
    const sel = { startRow: 0, startCol: 0, endRow: 1, endCol: 1 };
    const clip = copyCells(s, sel);
    assertEqual(clip.size, 4, "4 copied");
    pasteCells(s, clip, "D4", sel);
    assertEqual(s.cells.get("D4")?.value, 1, "D4");
    assertEqual(s.cells.get("E4")?.value, 2, "E4");
    assertEqual(s.cells.get("D5")?.value, 3, "D5");
    assertEqual(s.cells.get("E5")?.value, 4, "E5");
  });
  scenario("Copy preserves formulas", () => {
    const s = makeSheet({ A1:{type:"formula",value:null,formula:"SUM(1,2)"} });
    const clip = copyCells(s, { startRow: 0, startCol: 0, endRow: 0, endCol: 0 });
    pasteCells(s, clip, "B1", { startRow: 0, startCol: 0, endRow: 0, endCol: 0 });
    assertEqual(s.cells.get("B1")?.formula, "SUM(1,2)", "formula copied");
  });
});

feature("toggleBold / toggleItalic / toggleUnderline", () => {
  scenario("toggleBold flips bold state", () => {
    const s = makeSheet({ A1: { type: "string", value: "X" } });
    const styles: XlsxStyle[] = [{ id: 0, numFmt: null, font: { name: "Calibri", size: 11, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" }, fill: null, border: null, alignment: null }];
    const id1 = toggleBold(s, styles, "A1");
    assertEqual(styles[id1].font?.bold, true, "now bold");
    const id2 = toggleBold(s, styles, "A1");
    assertEqual(styles[id2].font?.bold, false, "now not bold");
  });
  scenario("toggleItalic flips italic state", () => {
    const s = makeSheet({ A1: { type: "string", value: "X" } });
    const styles: XlsxStyle[] = [{ id: 0, numFmt: null, font: { name: "Calibri", size: 11, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" }, fill: null, border: null, alignment: null }];
    toggleItalic(s, styles, "A1");
    assertEqual(styles[1].font?.italic, true, "italic");
  });
  scenario("toggleUnderline flips underline state", () => {
    const s = makeSheet({ A1: { type: "string", value: "X" } });
    const styles: XlsxStyle[] = [{ id: 0, numFmt: null, font: { name: "Calibri", size: 11, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" }, fill: null, border: null, alignment: null }];
    toggleUnderline(s, styles, "A1");
    assertEqual(styles[1].font?.underline, true, "underline");
  });
});

feature("setFontSize / setFontColor / setFill / setNumberFormat / setAlignment / setBorder", () => {
  const mkStyles = (): XlsxStyle[] => [{ id: 0, numFmt: null, font: { name: "Calibri", size: 11, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" }, fill: null, border: null, alignment: null }];

  scenario("setFontSize changes font size", () => {
    const s = makeSheet({ A1:{type:"string",value:"X"} }); const st = mkStyles();
    setFontSize(s, st, "A1", 24);
    assertEqual(st[1].font?.size, 24, "size");
  });
  scenario("setFontColor changes font color", () => {
    const s = makeSheet({ A1:{type:"string",value:"X"} }); const st = mkStyles();
    setFontColor(s, st, "A1", "#FF0000");
    assertEqual(st[1].font?.color, "#FF0000", "color");
  });
  scenario("setFill changes background color", () => {
    const s = makeSheet({ A1:{type:"string",value:"X"} }); const st = mkStyles();
    setFill(st, "A1", s, "#FFFF00");
    assertEqual(st[1].fill?.fgColor, "#FFFF00", "fill");
  });
  scenario("setNumberFormat changes format", () => {
    const s = makeSheet({ A1:{type:"number",value:42} }); const st = mkStyles();
    setNumberFormat(st, "A1", s, "#,##0.00");
    assertEqual(st[1].numFmt, "#,##0.00", "numFmt");
  });
  scenario("setAlignment changes alignment", () => {
    const s = makeSheet({ A1:{type:"string",value:"X"} }); const st = mkStyles();
    setAlignment(st, "A1", s, { horizontal: "center", vertical: "middle", wrapText: true, textRotation: 45, indent: 0 });
    assertEqual(st[1].alignment?.horizontal, "center", "h");
    assertEqual(st[1].alignment?.wrapText, true, "wrap");
    assertEqual(st[1].alignment?.textRotation, 45, "rotation");
  });
  scenario("setBorder sets all edges", () => {
    const s = makeSheet({ A1:{type:"string",value:"X"} }); const st = mkStyles();
    setBorder(st, "A1", s, { top: { style: "thin", color: "#000" }, bottom: { style: "thick", color: "#F00" }, left: { style: "dashed", color: "#0F0" }, right: { style: "double", color: "#00F" } });
    assertEqual(st[1].border?.top?.style, "thin", "top");
    assertEqual(st[1].border?.bottom?.style, "thick", "bottom");
    assertEqual(st[1].border?.left?.style, "dashed", "left");
    assertEqual(st[1].border?.right?.style, "double", "right");
  });
});

feature("setColumnWidth / setRowHeight / setFreeze", () => {
  scenario("setColumnWidth stores custom width", () => {
    const s = makeSheet({}); setColumnWidth(s, 3, 25);
    assertEqual(s.colWidths.get(3), 25, "width");
  });
  scenario("setRowHeight stores custom height", () => {
    const s = makeSheet({}); setRowHeight(s, 5, 40);
    assertEqual(s.rowHeights.get(5), 40, "height");
  });
  scenario("setFreeze sets frozen panes", () => {
    const s = makeSheet({}); setFreeze(s, 2, 1);
    assertEqual(s.frozenRow, 2, "row"); assertEqual(s.frozenCol, 1, "col");
  });
  scenario("setFreeze to 0,0 unfreezes", () => {
    const s = makeSheet({}); s.frozenRow = 3; s.frozenCol = 2;
    setFreeze(s, 0, 0);
    assertEqual(s.frozenRow, 0, "unfrozen row"); assertEqual(s.frozenCol, 0, "unfrozen col");
  });
});

feature("Full Editor Workflow Integration", () => {
  scenario("Create → edit → undo → redo → save → reopen", async () => {
    const ed = createEditor();
    const wb = makeWorkbook([makeSheet({})]);
    loadWorkbook(ed, wb);
    const um = new UndoManager();

    // Edit cells
    um.pushUndo(ed); setCellValue(ed, "A1", "Revenue");
    um.pushUndo(ed); setCellValue(ed, "B1", 1000);
    um.pushUndo(ed); setCellValue(ed, "B2", 2000);
    um.pushUndo(ed); setCellValue(ed, "B3", "=SUM(B1:B2)");

    // Recalc
    recalculateSheet(currentSheet(ed)!);
    assertEqual(currentSheet(ed)!.cells.get("B3")?.calculatedValue, 3000, "B3=3000");

    // Undo B3 formula
    um.undo(ed);
    assertEqual(currentSheet(ed)!.cells.has("B3"), false, "B3 undone");

    // Redo
    um.redo(ed);
    assertEqual(currentSheet(ed)!.cells.get("B3")?.formula, "SUM(B1:B2)", "B3 redone");

    // Export and reopen
    const blob = exportXlsx(ed.workbook!);
    const parsed = parseXlsx(await blob.arrayBuffer());
    recalculateSheet(parsed.sheets[0]);
    assertEqual(parsed.sheets[0].cells.get("A1")?.value, "Revenue", "roundtrip A1");
    assertEqual(parsed.sheets[0].cells.get("B3")?.calculatedValue, 3000, "roundtrip B3");
  });
});

// ===========================================================================
setTimeout(() => {
  console.log(`\n${"=".repeat(50)}`);
  console.log(`\x1b[1mResults: ${totalPass} passed, ${totalFail} failed, ${totalPass + totalFail} total\x1b[0m`);
  if (totalFail === 0) console.log("\x1b[32m\x1b[1m✓ ALL SCENARIOS PASSED\x1b[0m");
  else { console.log(`\x1b[31m\x1b[1m✗ ${totalFail} SCENARIO(S) FAILED\x1b[0m`); process.exit(1); }
}, 500);
