/**
 * Find & Replace — search cells by value/formula, navigate matches, replace.
 *
 * Operates on the in-memory {@link XlsxSheet} cell map. Supports case-sensitive
 * matching, entire-cell matching, and formula-level search.
 */
import type { XlsxSheet, XlsxCell, CellRef } from "./ooxml-parser";
import { buildRef } from "./ooxml-parser";

/** Options controlling how the find operation matches cells. */
export interface FindOptions {
  /** The search query string. */
  query: string;
  /** Whether comparison is case-sensitive. Default false. */
  matchCase: boolean;
  /** Whether the query must match the entire cell value. Default false. */
  matchEntireCell: boolean;
  /** Whether to search formula text instead of display values. Default false. */
  searchFormulas: boolean;
  /** Scope of the search — currently only "sheet" is used by per-sheet helpers. */
  searchScope: "sheet" | "workbook";
}

/** A single find result pointing to a matched cell. */
export interface FindResult {
  /** The cell reference (e.g. "A1"). */
  ref: CellRef;
  /** Index of the sheet within the workbook. */
  sheetIndex: number;
  /** Display name of the sheet. */
  sheetName: string;
  /** The matched display value or formula text. */
  value: string;
}

/**
 * Extract the searchable text from a cell based on search options.
 *
 * When {@link FindOptions.searchFormulas} is true and the cell has a formula,
 * the formula string (prefixed with `=`) is returned. Otherwise the display
 * value is coerced to a string.
 */
function cellText(cell: XlsxCell, searchFormulas: boolean): string {
  if (searchFormulas && cell.formula) {
    return `=${cell.formula}`;
  }
  if (cell.value != null) return String(cell.value);
  if (cell.calculatedValue != null) return String(cell.calculatedValue);
  return "";
}

/**
 * Test whether `text` matches `query` given the match options.
 *
 * Case folding is applied when `matchCase` is false. Entire-cell matching
 * requires the full text to equal the query; otherwise a substring check
 * is performed.
 */
function matches(text: string, query: string, matchCase: boolean, matchEntireCell: boolean): boolean {
  const t = matchCase ? text : text.toLowerCase();
  const q = matchCase ? query : query.toLowerCase();
  return matchEntireCell ? t === q : t.includes(q);
}

/**
 * Find all matching cells in a single sheet.
 *
 * Iterates every cell in the sheet's sparse cell map and returns results
 * sorted by row then column.
 *
 * @param sheet - The sheet to search.
 * @param sheetIndex - The sheet's position in the workbook (for result metadata).
 * @param options - Search parameters.
 * @returns Array of {@link FindResult} ordered top-left to bottom-right.
 */
export function findInSheet(sheet: XlsxSheet, sheetIndex: number, options: FindOptions): FindResult[] {
  const results: FindResult[] = [];
  if (!options.query) return results;

  for (const [ref, cell] of sheet.cells) {
    const text = cellText(cell, options.searchFormulas);
    if (text && matches(text, options.query, options.matchCase, options.matchEntireCell)) {
      results.push({
        ref,
        sheetIndex,
        sheetName: sheet.name,
        value: text,
      });
    }
  }

  // Sort by row, then column for predictable navigation order.
  results.sort((a, b) => {
    const cellA = sheet.cells.get(a.ref)!;
    const cellB = sheet.cells.get(b.ref)!;
    return cellA.row !== cellB.row ? cellA.row - cellB.row : cellA.col - cellB.col;
  });

  return results;
}

/**
 * Replace the find string in a single cell's value.
 *
 * Only cells whose type is `"string"` or whose formula text matches are
 * modified. Number, boolean, and error cells are left unchanged.
 *
 * @param cell - The cell to mutate.
 * @param find - The substring to find.
 * @param replacement - The replacement string.
 * @param matchCase - Whether the replacement is case-sensitive.
 * @returns `true` if the cell was modified.
 */
export function replaceInCell(cell: XlsxCell, find: string, replacement: string, matchCase: boolean): boolean {
  if (cell.formula) {
    const pattern = matchCase
      ? new RegExp(escapeRegExp(find), "g")
      : new RegExp(escapeRegExp(find), "gi");
    const original = cell.formula;
    const replaced = original.replace(pattern, replacement);
    if (replaced !== original) {
      cell.formula = replaced;
      return true;
    }
    return false;
  }

  if (cell.type === "string" && typeof cell.value === "string") {
    const pattern = matchCase
      ? new RegExp(escapeRegExp(find), "g")
      : new RegExp(escapeRegExp(find), "gi");
    const original = cell.value;
    const replaced = original.replace(pattern, replacement);
    if (replaced !== original) {
      cell.value = replaced;
      return true;
    }
  }

  return false;
}

/**
 * Replace all occurrences of `find` in every cell of a sheet.
 *
 * @param sheet - The sheet to mutate.
 * @param find - The substring to find.
 * @param replacement - The replacement string.
 * @param matchCase - Whether the replacement is case-sensitive.
 * @returns The number of cells that were modified.
 */
export function replaceAllInSheet(
  sheet: XlsxSheet,
  find: string,
  replacement: string,
  matchCase: boolean,
): number {
  let count = 0;
  for (const [, cell] of sheet.cells) {
    if (replaceInCell(cell, find, replacement, matchCase)) {
      count++;
    }
  }
  return count;
}

/**
 * Escape special regex characters so a literal string can be used in a RegExp.
 */
function escapeRegExp(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
