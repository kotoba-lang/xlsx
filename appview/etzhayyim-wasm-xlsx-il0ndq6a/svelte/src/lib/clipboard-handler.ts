/**
 * Clipboard Handler — copy/paste with TSV and HTML table support for
 * interoperability with Excel, Google Sheets, and other spreadsheet applications.
 */

import type { XlsxSheet, XlsxCell, CellRef } from "./ooxml-parser";
import { buildRef, parseRef } from "./ooxml-parser";
import type { CellRange } from "./editor-state.svelte";
import { normalizeRange } from "./cell-selection";

// ---------------------------------------------------------------------------
// Copy to clipboard (TSV + HTML)
// ---------------------------------------------------------------------------

/** Copy a cell range to the system clipboard as TSV and HTML table. */
export async function copyRangeToClipboard(sheet: XlsxSheet, range: CellRange): Promise<void> {
  const n = normalizeRange(range);
  const tsv = buildTSV(sheet, n);
  const html = buildHTMLTable(sheet, n);

  try {
    const items = [
      new ClipboardItem({
        "text/plain": new Blob([tsv], { type: "text/plain" }),
        "text/html": new Blob([html], { type: "text/html" }),
      }),
    ];
    await navigator.clipboard.write(items);
  } catch {
    // Fallback: plain text only
    await navigator.clipboard.writeText(tsv);
  }
}

/** Build TSV (Tab-Separated Values) string from a range. */
function buildTSV(sheet: XlsxSheet, range: CellRange): string {
  const rows: string[] = [];
  for (let r = range.startRow; r <= range.endRow; r++) {
    const cols: string[] = [];
    for (let c = range.startCol; c <= range.endCol; c++) {
      const ref = buildRef(c, r);
      const cell = sheet.cells.get(ref);
      cols.push(cellText(cell));
    }
    rows.push(cols.join("\t"));
  }
  return rows.join("\n");
}

/** Build HTML table string from a range (for rich paste into Excel/Sheets). */
function buildHTMLTable(sheet: XlsxSheet, range: CellRange): string {
  let html = "<table>";
  for (let r = range.startRow; r <= range.endRow; r++) {
    html += "<tr>";
    for (let c = range.startCol; c <= range.endCol; c++) {
      const ref = buildRef(c, r);
      const cell = sheet.cells.get(ref);
      html += `<td>${escHtml(cellText(cell))}</td>`;
    }
    html += "</tr>";
  }
  html += "</table>";
  return html;
}

// ---------------------------------------------------------------------------
// Paste from clipboard
// ---------------------------------------------------------------------------

/** Parse pasted text (TSV or plain) into a 2D array of values. */
export function parsePastedText(text: string): string[][] {
  const lines = text.split(/\r?\n/).filter((line, i, arr) => !(i === arr.length - 1 && line === ""));
  return lines.map((line) => line.split("\t"));
}

/** Apply pasted values to a sheet starting at a given cell. */
export function applyPastedValues(
  sheet: XlsxSheet,
  startRef: CellRef,
  values: string[][],
): void {
  const { col: startCol, row: startRow } = parseRef(startRef);

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const ref = buildRef(startCol + c, startRow + r);
      const text = values[r][c];

      if (text === "") {
        sheet.cells.delete(ref);
        continue;
      }

      const num = Number(text);
      const isNum = !isNaN(num) && text.trim() !== "";

      sheet.cells.set(ref, {
        ref,
        row: startRow + r,
        col: startCol + c,
        type: text.startsWith("=") ? "formula" : isNum ? "number" : "string",
        value: text.startsWith("=") ? null : isNum ? num : text,
        formula: text.startsWith("=") ? text.slice(1) : null,
        calculatedValue: null,
        styleId: 0,
        hyperlink: null,
      });
    }
  }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/** Get display text for a cell. */
function cellText(cell: XlsxCell | undefined): string {
  if (!cell) return "";
  if (cell.type === "formula") return cell.calculatedValue != null ? String(cell.calculatedValue) : "";
  return cell.value != null ? String(cell.value) : "";
}

/** Escape HTML entities. */
function escHtml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
