/**
 * xlsx kotoba — spreadsheet document-tree record types (workbook → sheet →
 * cell, + merged regions).
 *
 * Per ADR-2606011400 (Consensys pattern) + ADR-2605172400 (3-axis OR-test).
 *
 * AXIS NOTE: (a) document-editor product (the pptx / editor / bim / cad cluster).
 * A workbook is the USER'S OWN document held in their OWN repo — first-party
 * content, so the AT PDS record IS the canonical store (no third-party PII
 * custody, no settlement, no fulfillment liability). The 131-function formula
 * engine, OOXML (SpreadsheetML) parse, and XLSX/CSV export all run CLIENT-SIDE
 * in the WASM appview — not server compute. This package replaces the RW/kagami
 * persistence with AT PDS records; the editor UI is unchanged.
 *
 * AT-Lexicon: no float. Cell values are stored as STRINGS with a `dataType`
 * discriminator (numbers as decimal strings — preserves precision and dodges the
 * float ban); counts / indices are integers. Formulas are strings.
 *
 * Identity hierarchy:
 *   did:web:xlsx.etzhayyim.com                          — controller
 *   did:web:xlsx.etzhayyim.com:wb:{workbookId}          — a workbook
 *   did:web:xlsx.etzhayyim.com:sheet:{sheetId}          — a sheet
 *   did:web:xlsx.etzhayyim.com:cell:{cellId}            — a cell
 *   did:web:xlsx.etzhayyim.com:merge:{mergeId}          — a merged region
 */

export const XLSX_DID_PREFIX = "did:web:xlsx.etzhayyim.com:" as const;

export const WORKBOOK_COLLECTION = "com.etzhayyim.apps.xlsx.workbook";
export const SHEET_COLLECTION = "com.etzhayyim.apps.xlsx.sheet";
export const CELL_COLLECTION = "com.etzhayyim.apps.xlsx.cell";
export const MERGED_REGION_COLLECTION = "com.etzhayyim.apps.xlsx.mergedRegion";

// ─── Cell data type ─────────────────────────────────────────────────

export type CellType = "number" | "string" | "boolean" | "formula" | "date" | "error" | "empty";

export const CELL_TYPES: ReadonlySet<string> = new Set([
  "number",
  "string",
  "boolean",
  "formula",
  "date",
  "error",
  "empty",
]);

// ─── Workbook ───────────────────────────────────────────────────────

export interface WorkbookRecord {
  did: string;
  workbookId: string;
  title: string;
  sheetCount: number;
  activeSheet: number;
  createdAt: string;
}
export interface WorkbookView extends WorkbookRecord {
  workbookUri: string;
}
export interface CreateWorkbookInput {
  workbookId: string;
  title: string;
  sheetCount?: number;
  activeSheet?: number;
}
export interface CreateWorkbookOutput {
  status: "created" | "alreadyExists" | "rejected";
  workbookUri?: string;
  did?: string;
  workbookId?: string;
  error?: string;
}
export interface GetWorkbookInput {
  workbookId: string;
}
export interface GetWorkbookOutput {
  workbook?: WorkbookView;
  error?: string;
}
export interface ListWorkbooksInput {
  /** App-layer substring search over title. */
  q?: string;
  limit?: number;
  cursor?: string;
}
export interface ListWorkbooksOutput {
  items: WorkbookView[];
  cursor?: string;
  total: number;
}

// ─── Sheet ──────────────────────────────────────────────────────────

export interface SheetRecord {
  did: string;
  sheetId: string;
  /** FK → workbook. */
  workbookId: string;
  name: string;
  index: number;
  rowCount?: number;
  colCount?: number;
  createdAt: string;
}
export interface SheetView extends SheetRecord {
  sheetUri: string;
}
export interface AddSheetInput {
  sheetId: string;
  workbookId: string;
  name: string;
  index: number;
  rowCount?: number;
  colCount?: number;
}
export interface AddSheetOutput {
  status: "added" | "alreadyExists" | "rejected" | "workbookNotFound";
  sheetUri?: string;
  did?: string;
  sheetId?: string;
  error?: string;
}
export interface ListSheetsInput {
  workbookId?: string;
  limit?: number;
  cursor?: string;
}
export interface ListSheetsOutput {
  items: SheetView[];
  cursor?: string;
  total: number;
}

// ─── Cell ───────────────────────────────────────────────────────────

export interface CellRecord {
  did: string;
  cellId: string;
  /** FK → sheet. */
  sheetId: string;
  /** A1 reference, e.g. "B7". */
  ref: string;
  dataType: CellType;
  /** Literal / computed value as a string (numbers as decimal strings). */
  value?: string;
  /** Formula source, e.g. "=SUM(A1:A5)". */
  formula?: string;
  styleRef?: string;
  createdAt: string;
}
export interface CellView extends CellRecord {
  cellUri: string;
}
export interface SetCellInput {
  cellId: string;
  sheetId: string;
  ref: string;
  dataType: CellType;
  value?: string;
  formula?: string;
  styleRef?: string;
}
export interface SetCellOutput {
  status: "set" | "alreadyExists" | "rejected" | "sheetNotFound";
  cellUri?: string;
  did?: string;
  cellId?: string;
  error?: string;
}
export interface ListCellsInput {
  sheetId?: string;
  dataType?: CellType;
  limit?: number;
  cursor?: string;
}
export interface ListCellsOutput {
  items: CellView[];
  cursor?: string;
  total: number;
}

// ─── Merged region ──────────────────────────────────────────────────

export interface MergedRegionRecord {
  did: string;
  mergeId: string;
  /** FK → sheet. */
  sheetId: string;
  /** A1 range, e.g. "A1:C3". */
  range: string;
  createdAt: string;
}
export interface MergedRegionView extends MergedRegionRecord {
  mergeUri: string;
}
export interface AddMergedRegionInput {
  mergeId: string;
  sheetId: string;
  range: string;
}
export interface AddMergedRegionOutput {
  status: "added" | "alreadyExists" | "rejected" | "sheetNotFound";
  mergeUri?: string;
  did?: string;
  mergeId?: string;
  error?: string;
}
export interface ListMergedRegionsInput {
  sheetId?: string;
  limit?: number;
  cursor?: string;
}
export interface ListMergedRegionsOutput {
  items: MergedRegionView[];
  cursor?: string;
  total: number;
}

// ─── Coverage ───────────────────────────────────────────────────────

export interface CoverageInput {
  maxScan?: number;
}
export interface CoverageOutput {
  workbookCount?: number;
  sheetCount?: number;
  cellCount?: number;
  mergedRegionCount?: number;
  cellsByType?: Record<string, number>;
  truncated?: boolean;
  error?: string;
}

// ─── Validation + helpers ───────────────────────────────────────────

export function isUint(n: unknown): n is number {
  return typeof n === "number" && Number.isInteger(n) && n >= 0;
}
/** A1 cell reference, e.g. "A1", "BC123". */
export function isA1Ref(s: string): boolean {
  return /^[A-Z]{1,3}[1-9]\d*$/.test(s);
}
/** A1 range, e.g. "A1:C3". */
export function isA1Range(s: string): boolean {
  return /^[A-Z]{1,3}[1-9]\d*:[A-Z]{1,3}[1-9]\d*$/.test(s);
}

export function workbookDidFor(id: string): string {
  return `${XLSX_DID_PREFIX}wb:${id.toLowerCase()}`;
}
export function workbookRkey(id: string): string {
  return `wb-${id.toLowerCase()}`;
}
export function sheetDidFor(id: string): string {
  return `${XLSX_DID_PREFIX}sheet:${id.toLowerCase()}`;
}
export function sheetRkey(id: string): string {
  return `sheet-${id.toLowerCase()}`;
}
export function cellDidFor(id: string): string {
  return `${XLSX_DID_PREFIX}cell:${id.toLowerCase()}`;
}
export function cellRkey(id: string): string {
  return `cell-${id.toLowerCase()}`;
}
export function mergeDidFor(id: string): string {
  return `${XLSX_DID_PREFIX}merge:${id.toLowerCase()}`;
}
export function mergeRkey(id: string): string {
  return `merge-${id.toLowerCase()}`;
}
