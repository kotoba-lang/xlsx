/**
 * Editor State — Svelte 5 rune-based reactive state for the XLSX editor.
 *
 * Uses a single exported `$state` object so properties can be mutated from imports.
 * Supports cell selection, range operations, undo/redo, and sheet tab management.
 */

import type {
  XlsxWorkbook,
  XlsxSheet,
  XlsxCell,
  XlsxStyle,
  XlsxFont,
  XlsxFill,
  XlsxBorder,
  XlsxBorderEdge,
  XlsxAlignment,
  CellRef,
} from "./ooxml-parser";
import { buildRef, parseRef, colToLetter } from "./ooxml-parser";

export type EditorTool = "select" | "text" | "fill" | "merge" | "chart";

export interface CellRange {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

export interface EditorSnapshot {
  workbook: XlsxWorkbook;
  activeSheetIndex: number;
  selection: CellRange;
}

/** Module-level clipboard for copy/paste operations. */
let clipboard: { cells: Map<CellRef, XlsxCell>; range: CellRange } | null = null;

/** Reactive editor state container. */
export const editor = $state({
  workbook: null as XlsxWorkbook | null,
  activeSheetIndex: 0,
  /** Current cell selection range (0-based row/col). */
  selection: {
    startRow: 0,
    startCol: 0,
    endRow: 0,
    endCol: 0,
  } as CellRange,
  /** Active cell within selection (where typing goes). */
  activeCell: "A1" as CellRef,
  /** Whether the user is editing a cell (formula bar or in-cell). */
  editingCell: null as CellRef | null,
  /** Current edit value in formula bar / in-cell editor. */
  editValue: "",
  isDirty: false,
  statusMessage: "",
  zoom: 1.0,
  scrollRow: 0,
  scrollCol: 0,
  showFormulaBar: true,
  showGridlines: true,
  activeTool: "select" as EditorTool,
  viewMode: "normal" as "normal" | "formulaView" | "pageBreak",
});

// --- Undo / Redo ---

const MAX_HISTORY = 50;
let undoStack: string[] = [];
let redoStack: string[] = [];

/** Serialise current workbook for undo snapshot. */
function serializeWorkbook(wb: XlsxWorkbook): string {
  const plain = {
    ...wb,
    sheets: wb.sheets.map((s) => ({
      ...s,
      cells: Object.fromEntries(s.cells),
      colWidths: Object.fromEntries(s.colWidths),
      rowHeights: Object.fromEntries(s.rowHeights),
    })),
  };
  return JSON.stringify(plain);
}

/** Deserialise a workbook from undo snapshot. */
function deserializeWorkbook(json: string): XlsxWorkbook {
  const plain = JSON.parse(json);
  return {
    ...plain,
    sheets: plain.sheets.map((s: Record<string, unknown>) => ({
      ...s,
      cells: new Map(Object.entries(s.cells as Record<string, XlsxCell>)),
      colWidths: new Map(Object.entries(s.colWidths as Record<string, number>).map(([k, v]) => [Number(k), v])),
      rowHeights: new Map(Object.entries(s.rowHeights as Record<string, number>).map(([k, v]) => [Number(k), v])),
    })),
  };
}

/** Take a snapshot for undo. */
export function pushUndo(): void {
  if (!editor.workbook) return;
  const snap = {
    workbook: serializeWorkbook($state.snapshot(editor.workbook) as XlsxWorkbook),
    activeSheetIndex: editor.activeSheetIndex,
    selection: { ...editor.selection },
  };
  undoStack.push(JSON.stringify(snap));
  if (undoStack.length > MAX_HISTORY) undoStack.shift();
  redoStack = [];
  editor.isDirty = true;
}

/** Undo the last action. */
export function undo(): void {
  if (undoStack.length === 0 || !editor.workbook) return;
  const current = {
    workbook: serializeWorkbook($state.snapshot(editor.workbook) as XlsxWorkbook),
    activeSheetIndex: editor.activeSheetIndex,
    selection: { ...editor.selection },
  };
  redoStack.push(JSON.stringify(current));
  const snap = JSON.parse(undoStack.pop()!);
  editor.workbook = deserializeWorkbook(snap.workbook);
  editor.activeSheetIndex = snap.activeSheetIndex;
  editor.selection = snap.selection;
  editor.statusMessage = "Undo";
}

/** Redo the last undone action. */
export function redo(): void {
  if (redoStack.length === 0 || !editor.workbook) return;
  const current = {
    workbook: serializeWorkbook($state.snapshot(editor.workbook) as XlsxWorkbook),
    activeSheetIndex: editor.activeSheetIndex,
    selection: { ...editor.selection },
  };
  undoStack.push(JSON.stringify(current));
  const snap = JSON.parse(redoStack.pop()!);
  editor.workbook = deserializeWorkbook(snap.workbook);
  editor.activeSheetIndex = snap.activeSheetIndex;
  editor.selection = snap.selection;
  editor.statusMessage = "Redo";
}

export function canUndo(): boolean { return undoStack.length > 0; }
export function canRedo(): boolean { return redoStack.length > 0; }

// --- Derived ---

/** Currently active sheet. */
export function currentSheet(): XlsxSheet | null {
  const wb = editor.workbook;
  if (!wb || editor.activeSheetIndex < 0 || editor.activeSheetIndex >= wb.sheets.length) return null;
  return wb.sheets[editor.activeSheetIndex];
}

/** Get cell at the given reference in the active sheet. */
export function getCell(ref: CellRef): XlsxCell | null {
  const sheet = currentSheet();
  return sheet?.cells.get(ref) ?? null;
}

/** Get the active cell value. */
export function activeCellValue(): string {
  const cell = getCell(editor.activeCell);
  if (!cell) return "";
  if (cell.formula) return `=${cell.formula}`;
  return cell.value != null ? String(cell.value) : "";
}

// --- Selection ---

/** Move selection to a specific cell. */
export function selectCell(ref: CellRef): void {
  const { col, row } = parseRef(ref);
  editor.selection = { startRow: row, startCol: col, endRow: row, endCol: col };
  editor.activeCell = ref;
  editor.editingCell = null;
}

/** Extend selection to a range. */
export function selectRange(range: CellRange): void {
  editor.selection = { ...range };
  editor.activeCell = buildRef(range.startCol, range.startRow);
}

/** Move active cell by delta (arrow keys). */
export function moveActiveCell(dRow: number, dCol: number): void {
  const { col, row } = parseRef(editor.activeCell);
  const newRow = Math.max(0, row + dRow);
  const newCol = Math.max(0, col + dCol);
  const ref = buildRef(newCol, newRow);
  selectCell(ref);
}

// --- Mutations ---

/** Load a new workbook. */
export function loadWorkbook(wb: XlsxWorkbook): void {
  undoStack = [];
  redoStack = [];
  editor.workbook = wb;
  editor.activeSheetIndex = wb.activeSheetIndex;
  editor.selection = { startRow: 0, startCol: 0, endRow: 0, endCol: 0 };
  editor.activeCell = "A1";
  editor.isDirty = false;
  editor.statusMessage = `Loaded: ${wb.title} (${wb.sheets.length} sheets)`;
}

/** Select a sheet tab. */
export function selectSheet(index: number): void {
  editor.activeSheetIndex = index;
  editor.selection = { startRow: 0, startCol: 0, endRow: 0, endCol: 0 };
  editor.activeCell = "A1";
  editor.editingCell = null;
}

/** Set cell value at a reference in the active sheet. */
export function setCellValue(ref: CellRef, value: string | number | boolean | null): void {
  const sheet = currentSheet();
  if (!sheet) return;
  pushUndo();

  const { col, row } = parseRef(ref);
  const existing = sheet.cells.get(ref);

  if (value === null || value === "") {
    sheet.cells.delete(ref);
    return;
  }

  let type: XlsxCell["type"] = "string";
  let parsedValue: XlsxCell["value"] = value;
  let formula: string | null = null;

  if (typeof value === "string" && value.startsWith("=")) {
    type = "formula";
    formula = value.slice(1);
    parsedValue = null;
  } else if (typeof value === "number") {
    type = "number";
  } else if (typeof value === "boolean") {
    type = "boolean";
  } else if (typeof value === "string") {
    const num = Number(value);
    if (!isNaN(num) && value.trim() !== "") {
      type = "number";
      parsedValue = num;
    }
  }

  sheet.cells.set(ref, {
    ref,
    row,
    col,
    type,
    value: parsedValue,
    formula,
    calculatedValue: null,
    styleId: existing?.styleId ?? 0,
    hyperlink: existing?.hyperlink ?? null,
  });
}

/** Start editing a cell. */
export function startEdit(ref: CellRef): void {
  editor.editingCell = ref;
  editor.editValue = activeCellValue();
}

/** Commit the current edit. */
export function commitEdit(): void {
  if (!editor.editingCell) return;
  setCellValue(editor.editingCell, editor.editValue);
  editor.editingCell = null;
  editor.editValue = "";
}

/** Cancel the current edit. */
export function cancelEdit(): void {
  editor.editingCell = null;
  editor.editValue = "";
}

/** Add a new sheet. */
export function addSheet(): void {
  const wb = editor.workbook;
  if (!wb) return;
  pushUndo();

  const name = `Sheet${wb.sheets.length + 1}`;
  const newSheet: XlsxSheet = {
    id: `sheet_${Date.now()}`,
    name,
    order: wb.sheets.length,
    hidden: false,
    cells: new Map(),
    mergedRegions: [],
    tables: [],
    charts: [],
    conditionalFormats: [],
    dataValidations: [],
    frozenRow: 0,
    frozenCol: 0,
    colWidths: new Map(),
    rowHeights: new Map(),
    defaultColWidth: 8.43,
    defaultRowHeight: 15,
  };
  wb.sheets.push(newSheet);
  editor.activeSheetIndex = wb.sheets.length - 1;
  editor.statusMessage = `Added: ${name}`;
}

/** Delete a sheet by index. */
export function deleteSheet(index: number): void {
  const wb = editor.workbook;
  if (!wb || wb.sheets.length <= 1) return;
  pushUndo();

  wb.sheets.splice(index, 1);
  for (let i = 0; i < wb.sheets.length; i++) wb.sheets[i].order = i;
  if (editor.activeSheetIndex >= wb.sheets.length) {
    editor.activeSheetIndex = wb.sheets.length - 1;
  }
}

/** Rename a sheet. */
export function renameSheet(index: number, name: string): void {
  const wb = editor.workbook;
  if (!wb || !wb.sheets[index]) return;
  pushUndo();
  wb.sheets[index].name = name;
}

/** Insert rows at a given position in the active sheet. */
export function insertRows(row: number, count: number): void {
  const sheet = currentSheet();
  if (!sheet) return;
  pushUndo();

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

/** Delete rows at a given position in the active sheet. */
export function deleteRows(row: number, count: number): void {
  const sheet = currentSheet();
  if (!sheet) return;
  pushUndo();

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

/** Insert columns at a given position in the active sheet. */
export function insertColumns(col: number, count: number): void {
  const sheet = currentSheet();
  if (!sheet) return;
  pushUndo();

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

/** Delete columns at a given position in the active sheet. */
export function deleteColumns(col: number, count: number): void {
  const sheet = currentSheet();
  if (!sheet) return;
  pushUndo();

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

/** Merge cells in the current selection. */
export function mergeCells(): void {
  const sheet = currentSheet();
  if (!sheet) return;
  const { startRow, startCol, endRow, endCol } = editor.selection;
  if (startRow === endRow && startCol === endCol) return;
  pushUndo();

  const startRef = buildRef(startCol, startRow);
  const endRef = buildRef(endCol, endRow);
  sheet.mergedRegions.push({ startRef, endRef });
}

/** Unmerge cells at the active cell position. */
export function unmergeCells(): void {
  const sheet = currentSheet();
  if (!sheet) return;
  pushUndo();

  const { row, col } = parseRef(editor.activeCell);
  sheet.mergedRegions = sheet.mergedRegions.filter((m) => {
    const start = parseRef(m.startRef);
    const end = parseRef(m.endRef);
    return !(row >= start.row && row <= end.row && col >= start.col && col <= end.col);
  });
}

/** Copy selected cells to clipboard. */
export function copyCells(): void {
  const sheet = currentSheet();
  if (!sheet) return;

  const { startRow, startCol, endRow, endCol } = editor.selection;
  const cells = new Map<CellRef, XlsxCell>();
  for (let r = startRow; r <= endRow; r++) {
    for (let c = startCol; c <= endCol; c++) {
      const ref = buildRef(c, r);
      const cell = sheet.cells.get(ref);
      if (cell) cells.set(ref, structuredClone($state.snapshot(cell)));
    }
  }
  clipboard = { cells, range: { ...editor.selection } };
  editor.statusMessage = "Copied";
}

/** Paste clipboard cells at the active cell position. */
export function pasteCells(): void {
  const sheet = currentSheet();
  if (!sheet || !clipboard) return;
  pushUndo();

  const { row: baseRow, col: baseCol } = parseRef(editor.activeCell);
  const { startRow, startCol } = clipboard.range;

  for (const [, cell] of clipboard.cells) {
    const newRow = baseRow + (cell.row - startRow);
    const newCol = baseCol + (cell.col - startCol);
    const newRef = buildRef(newCol, newRow);
    sheet.cells.set(newRef, { ...cell, ref: newRef, row: newRow, col: newCol });
  }
  editor.statusMessage = "Pasted";
}

// ---------------------------------------------------------------------------
// Formatting — Style application helpers
// ---------------------------------------------------------------------------

/** Normalise selection so start <= end. */
function normalizedSelection(): CellRange {
  const s = editor.selection;
  return {
    startRow: Math.min(s.startRow, s.endRow),
    startCol: Math.min(s.startCol, s.endCol),
    endRow: Math.max(s.startRow, s.endRow),
    endCol: Math.max(s.startCol, s.endCol),
  };
}

/** Default style values used when creating a fresh style. */
const DEFAULT_FONT: XlsxFont = {
  name: "Calibri",
  size: 11,
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  color: "#000000",
};

const DEFAULT_STYLE: XlsxStyle = {
  id: 0,
  numFmt: null,
  font: { ...DEFAULT_FONT },
  fill: null,
  border: null,
  alignment: null,
};

/**
 * Retrieve the resolved style for a cell, falling back to DEFAULT_STYLE.
 * Returns a deep clone so mutations do not affect the original.
 */
function resolveStyle(cell: XlsxCell | null): XlsxStyle {
  const styles = editor.workbook?.styles ?? [];
  const base = (cell ? styles[cell.styleId] : null) ?? DEFAULT_STYLE;
  return structuredClone(base);
}

/**
 * Register a new style in the workbook styles array and return its id.
 * Assigns a unique sequential id.
 */
function registerStyle(style: XlsxStyle): number {
  const wb = editor.workbook;
  if (!wb) return 0;
  const id = wb.styles.length;
  style.id = id;
  wb.styles.push(style);
  return id;
}

/**
 * Ensure a cell exists at the given ref in the sheet, creating an empty one if needed.
 */
function ensureCell(sheet: XlsxSheet, ref: CellRef, row: number, col: number): XlsxCell {
  let cell = sheet.cells.get(ref);
  if (!cell) {
    cell = {
      ref,
      row,
      col,
      type: "empty",
      value: null,
      formula: null,
      calculatedValue: null,
      styleId: 0,
      hyperlink: null,
    };
    sheet.cells.set(ref, cell);
  }
  return cell;
}

/**
 * Iterate over the normalised selection, calling `fn` for each coordinate.
 * Pushes undo once before iteration.
 */
function forEachSelected(fn: (sheet: XlsxSheet, ref: CellRef, row: number, col: number) => void): void {
  const sheet = currentSheet();
  if (!sheet) return;
  pushUndo();
  const { startRow, startCol, endRow, endCol } = normalizedSelection();
  for (let r = startRow; r <= endRow; r++) {
    for (let c = startCol; c <= endCol; c++) {
      fn(sheet, buildRef(c, r), r, c);
    }
  }
}

/** Apply number format to selection range. */
export function setNumberFormat(numFmt: string): void {
  forEachSelected((sheet, ref, r, c) => {
    const cell = ensureCell(sheet, ref, r, c);
    const style = resolveStyle(cell);
    style.numFmt = numFmt;
    cell.styleId = registerStyle(style);
  });
}

/** Apply font properties to selection range. */
export function setFont(props: Partial<XlsxFont>): void {
  forEachSelected((sheet, ref, r, c) => {
    const cell = ensureCell(sheet, ref, r, c);
    const style = resolveStyle(cell);
    style.font = { ...(style.font ?? DEFAULT_FONT), ...props };
    cell.styleId = registerStyle(style);
  });
}

/** Apply fill color to selection range. */
export function setFill(fgColor: string | null): void {
  forEachSelected((sheet, ref, r, c) => {
    const cell = ensureCell(sheet, ref, r, c);
    const style = resolveStyle(cell);
    style.fill = fgColor ? { type: "solid", fgColor, bgColor: null } : null;
    cell.styleId = registerStyle(style);
  });
}

/** Apply border to selection range. */
export function setBorder(
  edge: "top" | "bottom" | "left" | "right" | "all" | "none",
  style: string,
  color: string,
): void {
  forEachSelected((sheet, ref, r, c) => {
    const cell = ensureCell(sheet, ref, r, c);
    const resolved = resolveStyle(cell);
    if (edge === "none") {
      resolved.border = null;
    } else {
      const edgeDef: XlsxBorderEdge = {
        style: style as XlsxBorderEdge["style"],
        color,
      };
      const existing: XlsxBorder = resolved.border ?? {
        top: null,
        bottom: null,
        left: null,
        right: null,
      };
      if (edge === "all") {
        existing.top = edgeDef;
        existing.bottom = edgeDef;
        existing.left = edgeDef;
        existing.right = edgeDef;
      } else {
        existing[edge] = edgeDef;
      }
      resolved.border = existing;
    }
    cell.styleId = registerStyle(resolved);
  });
}

/** Apply alignment to selection range. */
export function setAlignment(props: Partial<XlsxAlignment>): void {
  forEachSelected((sheet, ref, r, c) => {
    const cell = ensureCell(sheet, ref, r, c);
    const style = resolveStyle(cell);
    const base: XlsxAlignment = style.alignment ?? {
      horizontal: "left",
      vertical: "bottom",
      wrapText: false,
      textRotation: 0,
      indent: 0,
    };
    style.alignment = { ...base, ...props };
    cell.styleId = registerStyle(style);
  });
}

/** Toggle bold on selection. */
export function toggleBold(): void {
  const anchor = getCell(editor.activeCell);
  const anchorStyle = resolveStyle(anchor);
  const newBold = !(anchorStyle.font?.bold ?? false);
  setFont({ bold: newBold });
}

/** Toggle italic on selection. */
export function toggleItalic(): void {
  const anchor = getCell(editor.activeCell);
  const anchorStyle = resolveStyle(anchor);
  const newItalic = !(anchorStyle.font?.italic ?? false);
  setFont({ italic: newItalic });
}

/** Toggle underline on selection. */
export function toggleUnderline(): void {
  const anchor = getCell(editor.activeCell);
  const anchorStyle = resolveStyle(anchor);
  const newUnderline = !(anchorStyle.font?.underline ?? false);
  setFont({ underline: newUnderline });
}

/** Set font size on selection. */
export function setFontSize(size: number): void {
  setFont({ size });
}

/** Set font color on selection. */
export function setFontColor(color: string): void {
  setFont({ color });
}

/** Set column width for a given column (in character widths). */
export function setColumnWidth(col: number, width: number): void {
  const sheet = currentSheet();
  if (!sheet) return;
  pushUndo();
  sheet.colWidths.set(col, width);
}

/** Set row height for a given row (in points). */
export function setRowHeight(row: number, height: number): void {
  const sheet = currentSheet();
  if (!sheet) return;
  pushUndo();
  sheet.rowHeights.set(row, height);
}

/** Auto-resize column to fit the widest cell content. */
export function autoResizeColumn(col: number): void {
  const sheet = currentSheet();
  if (!sheet) return;
  pushUndo();

  const canvas = document.createElement("canvas");
  const ctx = canvas.getContext("2d");
  if (!ctx) return;

  let maxWidth = 8.43; // minimum = default char width
  for (const [, cell] of sheet.cells) {
    if (cell.col !== col) continue;
    const text = cell.value != null ? String(cell.value) : "";
    if (!text) continue;

    const styles = editor.workbook?.styles ?? [];
    const style = styles[cell.styleId];
    const fontSize = style?.font?.size ?? 11;
    const fontName = style?.font?.name ?? "Calibri";
    ctx.font = `${fontSize}px ${fontName}, sans-serif`;
    const measured = ctx.measureText(text).width;
    // Convert px to char-width units (~7.5px per char) with 2-char padding
    const charWidth = measured / 7.5 + 2;
    if (charWidth > maxWidth) maxWidth = charWidth;
  }
  sheet.colWidths.set(col, maxWidth);
}

/** Set frozen rows/columns (freeze panes). */
export function setFreeze(rows: number, cols: number): void {
  const sheet = currentSheet();
  if (!sheet) return;
  pushUndo();
  sheet.frozenRow = rows;
  sheet.frozenCol = cols;
}

/** Reset editor to blank state. */
export function resetEditor(): void {
  editor.workbook = null;
  editor.activeSheetIndex = 0;
  editor.selection = { startRow: 0, startCol: 0, endRow: 0, endCol: 0 };
  editor.activeCell = "A1";
  editor.editingCell = null;
  editor.editValue = "";
  editor.isDirty = false;
  undoStack = [];
  redoStack = [];
}
