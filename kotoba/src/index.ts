/**
 * xlsx kotoba — barrel.
 *
 * Per ADR-2606011400. Spreadsheet document-tree (workbook → sheet → cell +
 * merged regions) on the etzhayyim substrate (AT PDS records; no RW).
 *
 *   workbook     : createWorkbook / getWorkbook / listWorkbooks (title search)
 *   sheet        : addSheet (FK→workbook) / listSheets (index-ordered)
 *   cell         : setCell (FK→sheet, A1 ref, value-as-string + dataType) / listCells
 *   mergedRegion : addMergedRegion (FK→sheet, A1 range) / listMergedRegions
 *   coverage
 *
 * (a) document-editor product (pptx/editor/bim/cad cluster). First-party user
 * documents; the 131-function formula engine + OOXML parse/export stay client-
 * side in the WASM appview.
 */

export * from "./types.js";
export {
  createWorkbook,
  getWorkbook,
  listWorkbooks,
  addSheet,
  listSheets,
  setCell,
  listCells,
  addMergedRegion,
  listMergedRegions,
  coverage,
} from "./registry.js";
