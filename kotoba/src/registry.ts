/**
 * xlsx kotoba — workbook → sheet → cell + mergedRegion document tree + coverage.
 * AT PDS records (no RW). Sheets FK→workbook; cells & merges FK→sheet. First-
 * party user documents; the formula engine + OOXML parse/export stay client-side.
 */

import type { Etzhayyim } from "@etzhayyim/sdk";
import {
  CELL_COLLECTION,
  CELL_TYPES,
  MERGED_REGION_COLLECTION,
  SHEET_COLLECTION,
  WORKBOOK_COLLECTION,
  cellDidFor,
  cellRkey,
  isA1Range,
  isA1Ref,
  isUint,
  mergeDidFor,
  mergeRkey,
  sheetDidFor,
  sheetRkey,
  workbookDidFor,
  workbookRkey,
  type AddMergedRegionInput,
  type AddMergedRegionOutput,
  type AddSheetInput,
  type AddSheetOutput,
  type CellRecord,
  type CellView,
  type CoverageInput,
  type CoverageOutput,
  type CreateWorkbookInput,
  type CreateWorkbookOutput,
  type GetWorkbookInput,
  type GetWorkbookOutput,
  type ListCellsInput,
  type ListCellsOutput,
  type ListMergedRegionsInput,
  type ListMergedRegionsOutput,
  type ListSheetsInput,
  type ListSheetsOutput,
  type ListWorkbooksInput,
  type ListWorkbooksOutput,
  type MergedRegionRecord,
  type MergedRegionView,
  type SetCellInput,
  type SetCellOutput,
  type SheetRecord,
  type SheetView,
  type WorkbookRecord,
  type WorkbookView,
} from "./types.js";

const PAGE_LIMIT = 100;
const DEFAULT_MAX_SCAN = 10_000;

async function exists(e: Etzhayyim, collection: string, rkey: string): Promise<boolean> {
  const resp = await e.read({ collection, rkey }).catch(() => ({ records: [] }));
  return Boolean(resp.records[0]?.value);
}

async function scanAll<T>(e: Etzhayyim, collection: string, maxScan: number, onRow: (v: T) => void): Promise<number> {
  let cursor: string | undefined;
  let scanned = 0;
  while (scanned < maxScan) {
    const page = await e.read<T>({ collection, cursor, limit: PAGE_LIMIT });
    for (const r of page.records) {
      if (scanned >= maxScan) break;
      onRow(r.value);
      scanned += 1;
    }
    if (scanned >= maxScan || !page.cursor || page.records.length < PAGE_LIMIT) break;
    cursor = page.cursor;
  }
  return scanned;
}

// ─── Workbook ───────────────────────────────────────────────────────

export async function createWorkbook(e: Etzhayyim, input: CreateWorkbookInput): Promise<CreateWorkbookOutput> {
  if (!input.workbookId || !input.title) return { status: "rejected", error: "missingRequiredFields" };
  if (input.sheetCount != null && !isUint(input.sheetCount)) return { status: "rejected", error: "sheetCountMustBeUint" };
  if (input.activeSheet != null && !isUint(input.activeSheet)) return { status: "rejected", error: "activeSheetMustBeUint" };
  const rkey = workbookRkey(input.workbookId);
  const existing = await e.read<WorkbookRecord>({ collection: WORKBOOK_COLLECTION, rkey }).catch(() => ({ records: [] }));
  if (existing.records[0]?.value) {
    return { status: "alreadyExists", workbookUri: existing.records[0].uri, did: existing.records[0].value.did, workbookId: input.workbookId };
  }
  const did = workbookDidFor(input.workbookId);
  const record: WorkbookRecord = {
    did,
    workbookId: input.workbookId,
    title: input.title,
    sheetCount: input.sheetCount ?? 1,
    activeSheet: input.activeSheet ?? 0,
    createdAt: new Date().toISOString(),
  };
  const receipt = await e.write({ collection: WORKBOOK_COLLECTION, record: record as unknown as Record<string, unknown>, rkey });
  return { status: "created", workbookUri: receipt.uri, did, workbookId: input.workbookId };
}

export async function getWorkbook(e: Etzhayyim, input: GetWorkbookInput): Promise<GetWorkbookOutput> {
  if (!input.workbookId) return { error: "invalidWorkbookId" };
  const resp = await e.read<WorkbookRecord>({ collection: WORKBOOK_COLLECTION, rkey: workbookRkey(input.workbookId) }).catch(() => ({ records: [] }));
  const r = resp.records[0];
  if (!r) return { error: "notFound" };
  return { workbook: { ...r.value, workbookUri: r.uri } };
}

export async function listWorkbooks(e: Etzhayyim, input: ListWorkbooksInput = {}): Promise<ListWorkbooksOutput> {
  const limit = Math.min(input.limit ?? 50, 200);
  const resp = await e.read<WorkbookRecord>({ collection: WORKBOOK_COLLECTION, cursor: input.cursor, limit });
  const q = input.q?.toLowerCase();
  const items: WorkbookView[] = resp.records
    .filter((r) => !q || r.value.title.toLowerCase().includes(q))
    .map((r) => ({ ...r.value, workbookUri: r.uri }));
  return { items, cursor: resp.cursor, total: items.length };
}

// ─── Sheet ──────────────────────────────────────────────────────────

export async function addSheet(e: Etzhayyim, input: AddSheetInput): Promise<AddSheetOutput> {
  if (!input.sheetId || !input.workbookId || !input.name) return { status: "rejected", error: "missingRequiredFields" };
  if (!isUint(input.index)) return { status: "rejected", error: "indexMustBeUint" };
  if (input.rowCount != null && !isUint(input.rowCount)) return { status: "rejected", error: "rowCountMustBeUint" };
  if (input.colCount != null && !isUint(input.colCount)) return { status: "rejected", error: "colCountMustBeUint" };
  if (!(await exists(e, WORKBOOK_COLLECTION, workbookRkey(input.workbookId)))) {
    return { status: "workbookNotFound", error: `workbookNotFound:${input.workbookId}` };
  }
  const rkey = sheetRkey(input.sheetId);
  const existing = await e.read<SheetRecord>({ collection: SHEET_COLLECTION, rkey }).catch(() => ({ records: [] }));
  if (existing.records[0]?.value) {
    return { status: "alreadyExists", sheetUri: existing.records[0].uri, did: existing.records[0].value.did, sheetId: input.sheetId };
  }
  const did = sheetDidFor(input.sheetId);
  const record: SheetRecord = {
    did,
    sheetId: input.sheetId,
    workbookId: input.workbookId,
    name: input.name,
    index: input.index,
    rowCount: input.rowCount,
    colCount: input.colCount,
    createdAt: new Date().toISOString(),
  };
  const receipt = await e.write({ collection: SHEET_COLLECTION, record: record as unknown as Record<string, unknown>, rkey });
  return { status: "added", sheetUri: receipt.uri, did, sheetId: input.sheetId };
}

export async function listSheets(e: Etzhayyim, input: ListSheetsInput = {}): Promise<ListSheetsOutput> {
  const limit = Math.min(input.limit ?? 50, 200);
  const resp = await e.read<SheetRecord>({ collection: SHEET_COLLECTION, cursor: input.cursor, limit });
  const items: SheetView[] = resp.records
    .filter((r) => !input.workbookId || r.value.workbookId === input.workbookId)
    .map((r) => ({ ...r.value, sheetUri: r.uri }))
    .sort((a, b) => a.index - b.index);
  return { items, cursor: resp.cursor, total: items.length };
}

// ─── Cell ───────────────────────────────────────────────────────────

export async function setCell(e: Etzhayyim, input: SetCellInput): Promise<SetCellOutput> {
  if (!input.cellId || !input.sheetId || !input.ref) return { status: "rejected", error: "missingRequiredFields" };
  if (!CELL_TYPES.has(input.dataType)) return { status: "rejected", error: "invalidDataType" };
  if (!isA1Ref(input.ref.toUpperCase())) return { status: "rejected", error: "invalidA1Ref" };
  if (!(await exists(e, SHEET_COLLECTION, sheetRkey(input.sheetId)))) {
    return { status: "sheetNotFound", error: `sheetNotFound:${input.sheetId}` };
  }
  const rkey = cellRkey(input.cellId);
  const existing = await e.read<CellRecord>({ collection: CELL_COLLECTION, rkey }).catch(() => ({ records: [] }));
  if (existing.records[0]?.value) {
    return { status: "alreadyExists", cellUri: existing.records[0].uri, did: existing.records[0].value.did, cellId: input.cellId };
  }
  const did = cellDidFor(input.cellId);
  const record: CellRecord = {
    did,
    cellId: input.cellId,
    sheetId: input.sheetId,
    ref: input.ref.toUpperCase(),
    dataType: input.dataType,
    value: input.value,
    formula: input.formula,
    styleRef: input.styleRef,
    createdAt: new Date().toISOString(),
  };
  const receipt = await e.write({ collection: CELL_COLLECTION, record: record as unknown as Record<string, unknown>, rkey });
  return { status: "set", cellUri: receipt.uri, did, cellId: input.cellId };
}

export async function listCells(e: Etzhayyim, input: ListCellsInput = {}): Promise<ListCellsOutput> {
  const limit = Math.min(input.limit ?? 100, 500);
  const resp = await e.read<CellRecord>({ collection: CELL_COLLECTION, cursor: input.cursor, limit });
  const items: CellView[] = resp.records
    .filter((r) => {
      const v = r.value;
      if (input.sheetId && v.sheetId !== input.sheetId) return false;
      if (input.dataType && v.dataType !== input.dataType) return false;
      return true;
    })
    .map((r) => ({ ...r.value, cellUri: r.uri }));
  return { items, cursor: resp.cursor, total: items.length };
}

// ─── Merged region ──────────────────────────────────────────────────

export async function addMergedRegion(e: Etzhayyim, input: AddMergedRegionInput): Promise<AddMergedRegionOutput> {
  if (!input.mergeId || !input.sheetId || !input.range) return { status: "rejected", error: "missingRequiredFields" };
  if (!isA1Range(input.range.toUpperCase())) return { status: "rejected", error: "invalidA1Range" };
  if (!(await exists(e, SHEET_COLLECTION, sheetRkey(input.sheetId)))) {
    return { status: "sheetNotFound", error: `sheetNotFound:${input.sheetId}` };
  }
  const rkey = mergeRkey(input.mergeId);
  const existing = await e.read<MergedRegionRecord>({ collection: MERGED_REGION_COLLECTION, rkey }).catch(() => ({ records: [] }));
  if (existing.records[0]?.value) {
    return { status: "alreadyExists", mergeUri: existing.records[0].uri, did: existing.records[0].value.did, mergeId: input.mergeId };
  }
  const did = mergeDidFor(input.mergeId);
  const record: MergedRegionRecord = {
    did,
    mergeId: input.mergeId,
    sheetId: input.sheetId,
    range: input.range.toUpperCase(),
    createdAt: new Date().toISOString(),
  };
  const receipt = await e.write({ collection: MERGED_REGION_COLLECTION, record: record as unknown as Record<string, unknown>, rkey });
  return { status: "added", mergeUri: receipt.uri, did, mergeId: input.mergeId };
}

export async function listMergedRegions(e: Etzhayyim, input: ListMergedRegionsInput = {}): Promise<ListMergedRegionsOutput> {
  const limit = Math.min(input.limit ?? 50, 200);
  const resp = await e.read<MergedRegionRecord>({ collection: MERGED_REGION_COLLECTION, cursor: input.cursor, limit });
  const items: MergedRegionView[] = resp.records
    .filter((r) => !input.sheetId || r.value.sheetId === input.sheetId)
    .map((r) => ({ ...r.value, mergeUri: r.uri }));
  return { items, cursor: resp.cursor, total: items.length };
}

// ─── Coverage ───────────────────────────────────────────────────────

export async function coverage(e: Etzhayyim, input: CoverageInput = {}): Promise<CoverageOutput> {
  const maxScan = Math.min(input.maxScan ?? DEFAULT_MAX_SCAN, DEFAULT_MAX_SCAN);
  const cellsByType: Record<string, number> = {};
  const workbookCount = await scanAll<WorkbookRecord>(e, WORKBOOK_COLLECTION, maxScan, () => {});
  const sheetCount = await scanAll<SheetRecord>(e, SHEET_COLLECTION, maxScan, () => {});
  const cellCount = await scanAll<CellRecord>(e, CELL_COLLECTION, maxScan, (v) => {
    cellsByType[v.dataType] = (cellsByType[v.dataType] ?? 0) + 1;
  });
  const mergedRegionCount = await scanAll<MergedRegionRecord>(e, MERGED_REGION_COLLECTION, maxScan, () => {});
  return {
    workbookCount,
    sheetCount,
    cellCount,
    mergedRegionCount,
    cellsByType,
    truncated: workbookCount >= maxScan || sheetCount >= maxScan || cellCount >= maxScan || mergedRegionCount >= maxScan,
  };
}
