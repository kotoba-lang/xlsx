/**
 * Cell Selection — selection model for spreadsheet range selection, multi-select,
 * autofill handle, and keyboard navigation.
 */

import type { CellRef } from "./ooxml-parser";
import { parseRef, buildRef, colToLetter } from "./ooxml-parser";
import type { CellRange } from "./editor-state.svelte";

// ---------------------------------------------------------------------------
// Range utilities
// ---------------------------------------------------------------------------

/** Normalise a range so start <= end. */
export function normalizeRange(range: CellRange): CellRange {
  return {
    startRow: Math.min(range.startRow, range.endRow),
    startCol: Math.min(range.startCol, range.endCol),
    endRow: Math.max(range.startRow, range.endRow),
    endCol: Math.max(range.startCol, range.endCol),
  };
}

/** Check if a cell is within a range. */
export function isCellInRange(ref: CellRef, range: CellRange): boolean {
  const { col, row } = parseRef(ref);
  const n = normalizeRange(range);
  return row >= n.startRow && row <= n.endRow && col >= n.startCol && col <= n.endCol;
}

/** Count cells in a range. */
export function rangeSize(range: CellRange): number {
  const n = normalizeRange(range);
  return (n.endRow - n.startRow + 1) * (n.endCol - n.startCol + 1);
}

/** Convert a range to "A1:B3" notation. */
export function rangeToString(range: CellRange): string {
  const n = normalizeRange(range);
  const start = buildRef(n.startCol, n.startRow);
  const end = buildRef(n.endCol, n.endRow);
  if (start === end) return start;
  return `${start}:${end}`;
}

/** Parse "A1:B3" or "A1" to CellRange. */
export function parseRangeString(rangeStr: string): CellRange {
  const parts = rangeStr.split(":");
  const start = parseRef(parts[0]);
  const end = parts[1] ? parseRef(parts[1]) : start;
  return {
    startRow: start.row,
    startCol: start.col,
    endRow: end.row,
    endCol: end.col,
  };
}

// ---------------------------------------------------------------------------
// Autofill
// ---------------------------------------------------------------------------

/** Detect number pattern in a series of values for autofill. */
export function detectFillPattern(values: (string | number | boolean | null)[]): { type: "increment"; step: number } | { type: "copy" } | null {
  if (values.length === 0) return null;

  const nums = values.filter((v): v is number => typeof v === "number");
  if (nums.length >= 2) {
    const step = nums[1] - nums[0];
    const isLinear = nums.every((v, i) => i === 0 || Math.abs(v - nums[i - 1] - step) < 1e-10);
    if (isLinear) return { type: "increment", step };
  }

  return { type: "copy" };
}

/** Generate autofill values based on source pattern. */
export function generateFillValues(
  sourceValues: (string | number | boolean | null)[],
  count: number,
): (string | number | boolean | null)[] {
  if (sourceValues.length === 0) return Array(count).fill(null);

  const pattern = detectFillPattern(sourceValues);
  const result: (string | number | boolean | null)[] = [];

  for (let i = 0; i < count; i++) {
    if (pattern?.type === "increment" && typeof sourceValues[sourceValues.length - 1] === "number") {
      const last = sourceValues[sourceValues.length - 1] as number;
      result.push(last + pattern.step * (i + 1));
    } else {
      result.push(sourceValues[i % sourceValues.length]);
    }
  }

  return result;
}

// ---------------------------------------------------------------------------
// Keyboard navigation helpers
// ---------------------------------------------------------------------------

/** Compute next cell ref after Tab key (move right, wrap to next row). */
export function nextCellTab(ref: CellRef, maxCol: number): CellRef {
  const { col, row } = parseRef(ref);
  if (col + 1 <= maxCol) return buildRef(col + 1, row);
  return buildRef(0, row + 1);
}

/** Compute next cell ref after Shift+Tab (move left, wrap to previous row). */
export function prevCellTab(ref: CellRef): CellRef {
  const { col, row } = parseRef(ref);
  if (col > 0) return buildRef(col - 1, row);
  if (row > 0) return buildRef(25, row - 1);
  return ref;
}

/** Compute cell ref after Ctrl+Arrow (jump to edge of data). */
export function jumpToEdge(
  ref: CellRef,
  direction: "up" | "down" | "left" | "right",
  hasData: (ref: CellRef) => boolean,
  maxRow: number,
  maxCol: number,
): CellRef {
  const { col, row } = parseRef(ref);
  let r = row;
  let c = col;

  const dr = direction === "up" ? -1 : direction === "down" ? 1 : 0;
  const dc = direction === "left" ? -1 : direction === "right" ? 1 : 0;

  const currentHasData = hasData(ref);

  if (currentHasData) {
    // Jump to last cell with data in direction
    while (true) {
      const nr = r + dr;
      const nc = c + dc;
      if (nr < 0 || nr > maxRow || nc < 0 || nc > maxCol) break;
      if (!hasData(buildRef(nc, nr))) break;
      r = nr;
      c = nc;
    }
  } else {
    // Jump to first cell with data in direction
    while (true) {
      r += dr;
      c += dc;
      if (r < 0 || r > maxRow || c < 0 || c > maxCol) { r -= dr; c -= dc; break; }
      if (hasData(buildRef(c, r))) break;
    }
  }

  return buildRef(Math.max(0, c), Math.max(0, r));
}

// ---------------------------------------------------------------------------
// Column/Row selection
// ---------------------------------------------------------------------------

/** Create a range selecting an entire column. */
export function selectEntireColumn(col: number, maxRow: number): CellRange {
  return { startRow: 0, startCol: col, endRow: maxRow, endCol: col };
}

/** Create a range selecting an entire row. */
export function selectEntireRow(row: number, maxCol: number): CellRange {
  return { startRow: row, startCol: 0, endRow: row, endCol: maxCol };
}
