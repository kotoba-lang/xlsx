/**
 * Grid Renderer — Canvas 2D spreadsheet grid rendering pipeline.
 *
 * Renders only the visible viewport (virtual scrolling) for smooth performance
 * with 100K+ rows. Handles column/row headers, frozen panes, selection highlight,
 * merged cells, and cell content.
 */

import type { XlsxSheet, XlsxCell, XlsxStyle, XlsxBorderEdge, XlsxConditionalFormat, CellRef } from "./ooxml-parser";
import { buildRef, colToLetter, parseRef } from "./ooxml-parser";
import type { CellRange } from "./editor-state.svelte";

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

/** Default column width in pixels (8.43 char widths × ~7.5px per char). */
const DEFAULT_COL_WIDTH_PX = 64;
/** Default row height in pixels (15pt ≈ 20px). */
const DEFAULT_ROW_HEIGHT_PX = 20;
/** Row/col header size in pixels. */
const HEADER_WIDTH = 50;
const HEADER_HEIGHT = 24;
/** Grid line colour. */
const GRID_COLOR = "#d0d0d0";
/** Header background. */
const HEADER_BG = "#f3f3f3";
/** Header text colour. */
const HEADER_TEXT = "#555555";
/** Selection fill (translucent blue). */
const SELECTION_FILL = "rgba(66, 133, 244, 0.12)";
/** Selection border colour. */
const SELECTION_BORDER = "#1a73e8";
/** Active cell border colour. */
const ACTIVE_CELL_BORDER = "#1a73e8";
/** Frozen pane divider colour. */
const FROZEN_DIVIDER = "#888888";

// ---------------------------------------------------------------------------
// Sizing helpers
// ---------------------------------------------------------------------------

/** Get pixel width for a column index. */
export function colWidth(sheet: XlsxSheet, col: number): number {
  const charWidth = sheet.colWidths.get(col) ?? sheet.defaultColWidth;
  return Math.round(charWidth * 7.5);
}

/** Get pixel height for a row index. */
export function rowHeight(sheet: XlsxSheet, row: number): number {
  const ptHeight = sheet.rowHeights.get(row) ?? sheet.defaultRowHeight;
  return Math.round(ptHeight * (4 / 3));
}

/** Compute X offset for a column (sum of preceding widths). */
export function colOffset(sheet: XlsxSheet, col: number): number {
  let x = 0;
  for (let c = 0; c < col; c++) x += colWidth(sheet, c);
  return x;
}

/** Compute Y offset for a row (sum of preceding heights). */
export function rowOffset(sheet: XlsxSheet, row: number): number {
  let y = 0;
  for (let r = 0; r < row; r++) y += rowHeight(sheet, r);
  return y;
}

// ---------------------------------------------------------------------------
// Viewport calculation
// ---------------------------------------------------------------------------

export interface Viewport {
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
  canvasWidth: number;
  canvasHeight: number;
}

/** Determine visible row/col range from scroll position + canvas size. */
export function computeViewport(
  sheet: XlsxSheet,
  scrollRow: number,
  scrollCol: number,
  canvasWidth: number,
  canvasHeight: number,
): Viewport {
  const availW = canvasWidth - HEADER_WIDTH;
  const availH = canvasHeight - HEADER_HEIGHT;

  let endCol = scrollCol;
  let w = 0;
  while (w < availW && endCol < 16384) {
    w += colWidth(sheet, endCol);
    endCol++;
  }

  let endRow = scrollRow;
  let h = 0;
  while (h < availH && endRow < 1048576) {
    h += rowHeight(sheet, endRow);
    endRow++;
  }

  return { startRow: scrollRow, endRow, startCol: scrollCol, endCol, canvasWidth, canvasHeight };
}

// ---------------------------------------------------------------------------
// Main render
// ---------------------------------------------------------------------------

/** Render the spreadsheet grid onto a Canvas 2D context. */
export function renderGrid(
  ctx: CanvasRenderingContext2D,
  sheet: XlsxSheet,
  styles: XlsxStyle[],
  viewport: Viewport,
  selection: CellRange,
  activeCell: CellRef,
  zoom: number,
  showGridlines: boolean,
  editingCell?: CellRef | null,
  editValue?: string,
): void {
  const { startRow, endRow, startCol, endCol, canvasWidth, canvasHeight } = viewport;
  const dpr = window.devicePixelRatio || 1;

  ctx.save();
  ctx.scale(dpr * zoom, dpr * zoom);

  const logicalW = canvasWidth / zoom;
  const logicalH = canvasHeight / zoom;

  // Clear
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, logicalW, logicalH);

  // --- Cell content + background ---
  let y = HEADER_HEIGHT;
  for (let r = startRow; r < endRow; r++) {
    const rh = rowHeight(sheet, r);
    let x = HEADER_WIDTH;
    for (let c = startCol; c < endCol; c++) {
      const cw = colWidth(sheet, c);
      const ref = buildRef(c, r);
      const cell = sheet.cells.get(ref);

      // Conditional format data bar (rendered behind cell content)
      const cfDataBar = findDataBarFormat(sheet, c, r);
      if (cfDataBar) {
        const barRatio = computeDataBarRatio(sheet, c, r, cfDataBar);
        if (barRatio > 0) {
          const barColor = cfDataBar.style?.fill?.fgColor ?? "4472C4";
          ctx.fillStyle = `#${barColor.replace(/^FF/i, "")}33`; // 20% opacity
          ctx.fillRect(x + 1, y + 1, (cw - 2) * barRatio, rh - 2);
        }
      }

      // Determine display text: live edit preview or committed value
      const isEditing = editingCell === ref && editValue != null;
      let displayText = "";

      if (cell) {
        // Background fill
        const style = styles[cell.styleId];
        if (style?.fill?.fgColor) {
          ctx.fillStyle = `#${style.fill.fgColor.replace(/^FF/i, "")}`;
          ctx.fillRect(x, y, cw, rh);
        }
        displayText = isEditing ? editValue! : cellDisplayText(cell);
      } else if (isEditing) {
        displayText = editValue!;
      }

      if (displayText) {
        const style = cell ? styles[cell.styleId] : undefined;
        ctx.fillStyle = style?.font?.color ? `#${style.font.color.replace(/^FF/i, "")}` : "#000000";
        ctx.font = buildFontString(style);
        ctx.textBaseline = "middle";
        ctx.textAlign = "left";
        const textX = x + 3;
        const textY = y + rh / 2;
        ctx.save();
        ctx.beginPath();
        ctx.rect(x, y, cw, rh);
        ctx.clip();
        if (!isEditing && cell && (cell.type === "number" || (cell.type === "formula" && typeof cell.calculatedValue === "number"))) {
          ctx.textAlign = "right";
          ctx.fillText(displayText, x + cw - 3, textY);
        } else {
          ctx.fillText(displayText, textX, textY);
        }
        ctx.restore();
      }

      x += cw;
    }
    y += rh;
  }

  // --- Cell borders ---
  renderCellBorders(ctx, sheet, styles, viewport);

  // --- Gridlines ---
  if (showGridlines) {
    ctx.strokeStyle = GRID_COLOR;
    ctx.lineWidth = 0.5;
    ctx.beginPath();

    // Horizontal lines
    y = HEADER_HEIGHT;
    for (let r = startRow; r <= endRow; r++) {
      ctx.moveTo(HEADER_WIDTH, y);
      ctx.lineTo(logicalW, y);
      y += rowHeight(sheet, r);
    }

    // Vertical lines
    let x = HEADER_WIDTH;
    for (let c = startCol; c <= endCol; c++) {
      ctx.moveTo(x, HEADER_HEIGHT);
      ctx.lineTo(x, logicalH);
      x += colWidth(sheet, c);
    }
    ctx.stroke();
  }

  // --- Selection highlight ---
  renderSelection(ctx, sheet, viewport, selection, activeCell);

  // --- Frozen pane dividers ---
  if (sheet.frozenRow > 0) {
    const frozenY = HEADER_HEIGHT + rowOffset(sheet, sheet.frozenRow) - rowOffset(sheet, startRow);
    ctx.strokeStyle = FROZEN_DIVIDER;
    ctx.lineWidth = 2;
    ctx.beginPath();
    ctx.moveTo(0, frozenY);
    ctx.lineTo(logicalW, frozenY);
    ctx.stroke();
  }
  if (sheet.frozenCol > 0) {
    const frozenX = HEADER_WIDTH + colOffset(sheet, sheet.frozenCol) - colOffset(sheet, startCol);
    ctx.strokeStyle = FROZEN_DIVIDER;
    ctx.lineWidth = 2;
    ctx.beginPath();
    ctx.moveTo(frozenX, 0);
    ctx.lineTo(frozenX, logicalH);
    ctx.stroke();
  }

  // --- Row headers ---
  ctx.fillStyle = HEADER_BG;
  ctx.fillRect(0, HEADER_HEIGHT, HEADER_WIDTH, logicalH - HEADER_HEIGHT);
  ctx.fillStyle = HEADER_TEXT;
  ctx.font = "11px -apple-system, BlinkMacSystemFont, sans-serif";
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  y = HEADER_HEIGHT;
  for (let r = startRow; r < endRow; r++) {
    const rh = rowHeight(sheet, r);
    ctx.fillText(`${r + 1}`, HEADER_WIDTH / 2, y + rh / 2);
    y += rh;
  }

  // --- Column headers ---
  ctx.fillStyle = HEADER_BG;
  ctx.fillRect(HEADER_WIDTH, 0, logicalW - HEADER_WIDTH, HEADER_HEIGHT);
  ctx.fillStyle = HEADER_TEXT;
  let x = HEADER_WIDTH;
  for (let c = startCol; c < endCol; c++) {
    const cw = colWidth(sheet, c);
    ctx.fillText(colToLetter(c), x + cw / 2, HEADER_HEIGHT / 2);
    x += cw;
  }

  // --- Top-left corner ---
  ctx.fillStyle = HEADER_BG;
  ctx.fillRect(0, 0, HEADER_WIDTH, HEADER_HEIGHT);

  // --- Header borders ---
  ctx.strokeStyle = "#bbbbbb";
  ctx.lineWidth = 1;
  ctx.strokeRect(0, 0, HEADER_WIDTH, HEADER_HEIGHT);
  ctx.beginPath();
  ctx.moveTo(HEADER_WIDTH, 0);
  ctx.lineTo(HEADER_WIDTH, logicalH);
  ctx.moveTo(0, HEADER_HEIGHT);
  ctx.lineTo(logicalW, HEADER_HEIGHT);
  ctx.stroke();

  ctx.restore();
}

// ---------------------------------------------------------------------------
// Selection rendering
// ---------------------------------------------------------------------------

/** Render selection highlight + active cell border. */
function renderSelection(
  ctx: CanvasRenderingContext2D,
  sheet: XlsxSheet,
  viewport: Viewport,
  selection: CellRange,
  activeCell: CellRef,
): void {
  const { startRow: vStartRow, startCol: vStartCol } = viewport;

  const selMinRow = Math.min(selection.startRow, selection.endRow);
  const selMaxRow = Math.max(selection.startRow, selection.endRow);
  const selMinCol = Math.min(selection.startCol, selection.endCol);
  const selMaxCol = Math.max(selection.startCol, selection.endCol);

  // Selection rectangle
  const x1 = HEADER_WIDTH + colOffset(sheet, selMinCol) - colOffset(sheet, vStartCol);
  const y1 = HEADER_HEIGHT + rowOffset(sheet, selMinRow) - rowOffset(sheet, vStartRow);
  let w = 0;
  for (let c = selMinCol; c <= selMaxCol; c++) w += colWidth(sheet, c);
  let h = 0;
  for (let r = selMinRow; r <= selMaxRow; r++) h += rowHeight(sheet, r);

  // Fill
  ctx.fillStyle = SELECTION_FILL;
  ctx.fillRect(x1, y1, w, h);

  // Border
  ctx.strokeStyle = SELECTION_BORDER;
  ctx.lineWidth = 2;
  ctx.strokeRect(x1, y1, w, h);

  // Fill handle (bottom-right corner dot)
  ctx.fillStyle = SELECTION_BORDER;
  ctx.fillRect(x1 + w - 3, y1 + h - 3, 6, 6);
}

// ---------------------------------------------------------------------------
// Border rendering
// ---------------------------------------------------------------------------

/** Map border style names to canvas line widths. */
function borderLineWidth(edgeStyle: string): number {
  switch (edgeStyle) {
    case "thin": return 1;
    case "medium": return 2;
    case "thick": return 3;
    case "double": return 2;
    case "dashed": return 1;
    case "dotted": return 1;
    default: return 1;
  }
}

/** Configure canvas dash pattern for border style. */
function applyBorderDash(ctx: CanvasRenderingContext2D, edgeStyle: string): void {
  switch (edgeStyle) {
    case "dashed": ctx.setLineDash([4, 2]); break;
    case "dotted": ctx.setLineDash([1, 2]); break;
    default: ctx.setLineDash([]); break;
  }
}

/** Draw a single border edge between two points. */
function drawBorderEdge(
  ctx: CanvasRenderingContext2D,
  edge: XlsxBorderEdge,
  x1: number,
  y1: number,
  x2: number,
  y2: number,
): void {
  ctx.strokeStyle = `#${edge.color.replace(/^FF/i, "").replace(/^#/, "")}`;
  ctx.lineWidth = borderLineWidth(edge.style);
  applyBorderDash(ctx, edge.style);
  ctx.beginPath();
  ctx.moveTo(x1, y1);
  ctx.lineTo(x2, y2);
  ctx.stroke();
  if (edge.style === "double") {
    const offset = y1 === y2 ? 2 : 0;
    const offsetX = x1 === x2 ? 2 : 0;
    ctx.beginPath();
    ctx.moveTo(x1 + offsetX, y1 + offset);
    ctx.lineTo(x2 + offsetX, y2 + offset);
    ctx.stroke();
  }
  ctx.setLineDash([]);
}

/** Render cell borders for all visible cells that have border styles. */
function renderCellBorders(
  ctx: CanvasRenderingContext2D,
  sheet: XlsxSheet,
  styles: XlsxStyle[],
  viewport: Viewport,
): void {
  const { startRow, endRow, startCol, endCol, startRow: vStartRow, startCol: vStartCol } = viewport;

  let y = HEADER_HEIGHT;
  for (let r = startRow; r < endRow; r++) {
    const rh = rowHeight(sheet, r);
    let x = HEADER_WIDTH;
    for (let c = startCol; c < endCol; c++) {
      const cw = colWidth(sheet, c);
      const ref = buildRef(c, r);
      const cell = sheet.cells.get(ref);
      if (cell) {
        const style = styles[cell.styleId];
        const border = style?.border;
        if (border) {
          if (border.top) drawBorderEdge(ctx, border.top, x, y, x + cw, y);
          if (border.bottom) drawBorderEdge(ctx, border.bottom, x, y + rh, x + cw, y + rh);
          if (border.left) drawBorderEdge(ctx, border.left, x, y, x, y + rh);
          if (border.right) drawBorderEdge(ctx, border.right, x + cw, y, x + cw, y + rh);
        }
      }
      x += cw;
    }
    y += rh;
  }
}

// ---------------------------------------------------------------------------
// Conditional formatting — Data bars
// ---------------------------------------------------------------------------

/**
 * Find the first dataBar conditional format that covers the given cell.
 * Returns null if no applicable data bar rule exists.
 */
function findDataBarFormat(sheet: XlsxSheet, col: number, row: number): XlsxConditionalFormat | null {
  for (const cf of sheet.conditionalFormats) {
    if (cf.type !== "dataBar") continue;
    if (rangeContains(cf.ref, col, row)) return cf;
  }
  return null;
}

/**
 * Check if a cell range string (e.g. "A1:D10") contains the given col/row.
 * Supports single-ref ("A1") and range ("A1:D10") formats.
 */
function rangeContains(rangeStr: string, col: number, row: number): boolean {
  const parts = rangeStr.split(":");
  if (parts.length === 1) {
    const p = parseRef(parts[0]);
    return p.col === col && p.row === row;
  }
  const start = parseRef(parts[0]);
  const end = parseRef(parts[1]);
  return (
    col >= Math.min(start.col, end.col) &&
    col <= Math.max(start.col, end.col) &&
    row >= Math.min(start.row, end.row) &&
    row <= Math.max(start.row, end.row)
  );
}

/**
 * Compute the data bar fill ratio (0..1) for a numeric cell relative to
 * the min/max values in the conditional format range.
 */
function computeDataBarRatio(
  sheet: XlsxSheet,
  col: number,
  row: number,
  cf: XlsxConditionalFormat,
): number {
  const ref = buildRef(col, row);
  const cell = sheet.cells.get(ref);
  if (!cell) return 0;
  const val = typeof cell.value === "number" ? cell.value : Number(cell.value);
  if (isNaN(val)) return 0;

  // Scan the CF range to find min/max
  const parts = cf.ref.split(":");
  const start = parseRef(parts[0]);
  const end = parts.length > 1 ? parseRef(parts[1]) : start;
  let min = Infinity;
  let max = -Infinity;

  for (let r = Math.min(start.row, end.row); r <= Math.max(start.row, end.row); r++) {
    for (let c = Math.min(start.col, end.col); c <= Math.max(start.col, end.col); c++) {
      const cellRef = buildRef(c, r);
      const rangeCell = sheet.cells.get(cellRef);
      if (!rangeCell) continue;
      const v = typeof rangeCell.value === "number" ? rangeCell.value : Number(rangeCell.value);
      if (isNaN(v)) continue;
      if (v < min) min = v;
      if (v > max) max = v;
    }
  }

  if (min === max) return val > 0 ? 1 : 0;
  if (min >= 0) return (val - min) / (max - min);
  // Handle negative ranges: ratio from 0..max relative to min..max
  return Math.max(0, val / max);
}

// ---------------------------------------------------------------------------
// Cell display helpers
// ---------------------------------------------------------------------------

/** Get display text for a cell. */
function cellDisplayText(cell: XlsxCell): string {
  if (cell.type === "formula") {
    return cell.calculatedValue != null ? String(cell.calculatedValue) : "";
  }
  if (cell.value == null) return "";
  return String(cell.value);
}

/** Build CSS font string from style. */
function buildFontString(style?: XlsxStyle | null): string {
  if (!style?.font) return "11px Calibri, sans-serif";
  const f = style.font;
  const parts: string[] = [];
  if (f.italic) parts.push("italic");
  if (f.bold) parts.push("bold");
  parts.push(`${f.size}px`);
  parts.push(`${f.name}, sans-serif`);
  return parts.join(" ");
}

// ---------------------------------------------------------------------------
// Hit testing
// ---------------------------------------------------------------------------

/** Determine which cell ref was clicked given canvas coordinates. */
export function hitTestCell(
  sheet: XlsxSheet,
  canvasX: number,
  canvasY: number,
  viewport: Viewport,
  zoom: number,
): CellRef | null {
  const x = canvasX / zoom - HEADER_WIDTH;
  const y = canvasY / zoom - HEADER_HEIGHT;
  if (x < 0 || y < 0) return null;

  let col = viewport.startCol;
  let accX = 0;
  while (col < viewport.endCol && accX + colWidth(sheet, col) < x) {
    accX += colWidth(sheet, col);
    col++;
  }

  let row = viewport.startRow;
  let accY = 0;
  while (row < viewport.endRow && accY + rowHeight(sheet, row) < y) {
    accY += rowHeight(sheet, row);
    row++;
  }

  return buildRef(col, row);
}

/** Check if click is on a column header. Returns column index or -1. */
export function hitTestColHeader(
  sheet: XlsxSheet,
  canvasX: number,
  canvasY: number,
  viewport: Viewport,
  zoom: number,
): number {
  const x = canvasX / zoom - HEADER_WIDTH;
  const y = canvasY / zoom;
  if (y >= HEADER_HEIGHT || x < 0) return -1;

  let col = viewport.startCol;
  let accX = 0;
  while (col < viewport.endCol) {
    const w = colWidth(sheet, col);
    if (accX + w > x) return col;
    accX += w;
    col++;
  }
  return -1;
}

/** Check if click is on a row header. Returns row index or -1. */
export function hitTestRowHeader(
  sheet: XlsxSheet,
  canvasX: number,
  canvasY: number,
  viewport: Viewport,
  zoom: number,
): number {
  const x = canvasX / zoom;
  const y = canvasY / zoom - HEADER_HEIGHT;
  if (x >= HEADER_WIDTH || y < 0) return -1;

  let row = viewport.startRow;
  let accY = 0;
  while (row < viewport.endRow) {
    const h = rowHeight(sheet, row);
    if (accY + h > y) return row;
    accY += h;
    row++;
  }
  return -1;
}
