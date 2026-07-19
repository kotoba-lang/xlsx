import { describe, it, expect, beforeEach } from "vitest";
import { MockEtzhayyim } from "@etzhayyim/sdk-mock";
import {
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
} from "../src/index.js";

describe("xlsx kotoba", () => {
  let e: any;
  beforeEach(() => {
    e = new MockEtzhayyim({ did: "did:web:xlsx.etzhayyim.com" });
  });

  describe("workbook + sheet", () => {
    it("creates workbooks (uint counts), reads, lists, searches; adds sheets FK→workbook", async () => {
      expect((await createWorkbook(e, { workbookId: "WB-1", title: "2026 Budget", sheetCount: 3, activeSheet: 0 })).status).toBe("created");
      expect((await getWorkbook(e, { workbookId: "WB-1" })).workbook?.sheetCount).toBe(3);
      expect((await createWorkbook(e, { workbookId: "WB-X", title: "x", sheetCount: -1 })).status).toBe("rejected"); // uint
      await createWorkbook(e, { workbookId: "WB-2", title: "Sales Forecast" });
      expect((await listWorkbooks(e, { q: "budget" })).total).toBe(1);
      expect((await addSheet(e, { sheetId: "S-2", workbookId: "WB-1", name: "Q2", index: 1 })).status).toBe("added");
      expect((await addSheet(e, { sheetId: "S-1", workbookId: "WB-1", name: "Q1", index: 0 })).status).toBe("added");
      expect((await addSheet(e, { sheetId: "S-G", workbookId: "GHOST", name: "x", index: 0 })).status).toBe("workbookNotFound");
      const sheets = await listSheets(e, { workbookId: "WB-1" });
      expect(sheets.total).toBe(2);
      expect(sheets.items[0].name).toBe("Q1"); // index-ordered
    });
  });

  describe("cells + merged regions FK→sheet", () => {
    beforeEach(async () => {
      await createWorkbook(e, { workbookId: "WB-1", title: "WB" });
      await addSheet(e, { sheetId: "S-1", workbookId: "WB-1", name: "Sheet1", index: 0 });
    });
    it("sets cells (A1 ref + dataType validated, value-as-string), rejects bad ref/sheet", async () => {
      expect((await setCell(e, { cellId: "C-1", sheetId: "S-1", ref: "b7", dataType: "number", value: "1234.56" })).status).toBe("set");
      expect((await setCell(e, { cellId: "C-2", sheetId: "S-1", ref: "A1", dataType: "formula", formula: "=SUM(B1:B7)", value: "1234.56" })).status).toBe("set");
      expect((await setCell(e, { cellId: "C-X", sheetId: "S-1", ref: "7B", dataType: "string", value: "x" })).status).toBe("rejected"); // A1 ref
      expect((await setCell(e, { cellId: "C-Y", sheetId: "S-1", ref: "C3", dataType: "matrix" as any })).status).toBe("rejected"); // dataType
      expect((await setCell(e, { cellId: "C-G", sheetId: "GHOST", ref: "A1", dataType: "string" })).status).toBe("sheetNotFound");
      expect((await listCells(e, { sheetId: "S-1", dataType: "formula" })).total).toBe(1);
    });
    it("adds merged regions (A1 range validated)", async () => {
      expect((await addMergedRegion(e, { mergeId: "M-1", sheetId: "S-1", range: "a1:c3" })).status).toBe("added");
      expect((await addMergedRegion(e, { mergeId: "M-X", sheetId: "S-1", range: "A1" })).status).toBe("rejected"); // not a range
      expect((await listMergedRegions(e, { sheetId: "S-1" })).total).toBe(1);
    });
    it("coverage rolls up the document tree by cell type", async () => {
      await setCell(e, { cellId: "C-1", sheetId: "S-1", ref: "A1", dataType: "number", value: "10" });
      await setCell(e, { cellId: "C-2", sheetId: "S-1", ref: "A2", dataType: "string", value: "hello" });
      await addMergedRegion(e, { mergeId: "M-1", sheetId: "S-1", range: "B1:B5" });
      const cov = await coverage(e);
      expect(cov.workbookCount).toBe(1);
      expect(cov.sheetCount).toBe(1);
      expect(cov.cellCount).toBe(2);
      expect(cov.mergedRegionCount).toBe(1);
      expect(cov.cellsByType?.number).toBe(1);
      expect(cov.cellsByType?.string).toBe(1);
    });
  });
});
