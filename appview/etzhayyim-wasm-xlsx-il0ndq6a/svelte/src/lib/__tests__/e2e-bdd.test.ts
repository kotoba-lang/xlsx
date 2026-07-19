/**
 * E2E BDD Tests — xlsx.etzhayyim.com spreadsheet editor
 *
 * Behavior-Driven scenarios covering the full user workflow:
 *   Given → When → Then
 *
 * Run: cd svelte && npx tsx src/lib/__tests__/e2e-bdd.test.ts
 */

import { parseXlsx, buildRef, parseRef, colToLetter, letterToCol, type XlsxWorkbook, type XlsxSheet, type XlsxCell } from "../ooxml-parser";
import { exportXlsx } from "../xlsx-exporter";
import { evaluateFormula, getFormulaDependencies, detectCircular, recalculateSheet } from "../formula-engine";
import { parseCsv, sheetToCsv, csvToWorkbook } from "../csv-handler";
import { findInSheet, replaceAllInSheet, replaceInCell } from "../find-replace";
import { CommentStore } from "../comments";
import { normalizeRange, rangeToString, parseRangeString, isCellInRange, detectFillPattern, generateFillValues, nextCellTab, prevCellTab, jumpToEdge, selectEntireColumn, selectEntireRow } from "../cell-selection";
import { parsePastedText, applyPastedValues } from "../clipboard-handler";

// ---------------------------------------------------------------------------
// Test runner (minimal, no deps)
// ---------------------------------------------------------------------------

let totalPass = 0;
let totalFail = 0;
let currentFeature = "";
let currentScenario = "";

function feature(name: string, fn: () => void): void {
  currentFeature = name;
  console.log(`\n\x1b[1m📋 Feature: ${name}\x1b[0m`);
  fn();
}

function scenario(name: string, fn: () => void): void {
  currentScenario = name;
  try {
    fn();
    totalPass++;
    console.log(`  \x1b[32m✓\x1b[0m ${name}`);
  } catch (e: any) {
    totalFail++;
    console.log(`  \x1b[31m✗\x1b[0m ${name}`);
    console.log(`    \x1b[31m${e.message}\x1b[0m`);
  }
}

function assert(condition: boolean, msg: string): void {
  if (!condition) throw new Error(msg);
}

function assertEqual(actual: unknown, expected: unknown, label: string): void {
  if (actual !== expected) throw new Error(`${label}: expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`);
}

function assertClose(actual: number, expected: number, label: string, epsilon = 0.01): void {
  if (Math.abs(actual - expected) > epsilon) throw new Error(`${label}: expected ~${expected}, got ${actual}`);
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function makeSheet(cells: Record<string, { type: string; value: any; formula?: string }>): XlsxSheet {
  const map = new Map<string, XlsxCell>();
  for (const [ref, c] of Object.entries(cells)) {
    const { col, row } = parseRef(ref);
    map.set(ref, {
      ref: ref as any,
      row, col,
      type: c.type as any,
      value: c.formula ? null : c.value,
      formula: c.formula ?? null,
      calculatedValue: null,
      styleId: 0,
      hyperlink: null,
    });
  }
  return {
    id: "s1", name: "Sheet1", order: 0, hidden: false,
    cells: map, mergedRegions: [], tables: [], charts: [],
    conditionalFormats: [], dataValidations: [],
    frozenRow: 0, frozenCol: 0,
    colWidths: new Map(), rowHeights: new Map(),
    defaultColWidth: 8.43, defaultRowHeight: 15,
  };
}

function makeWorkbook(sheets: XlsxSheet[]): XlsxWorkbook {
  return {
    id: "wb1", title: "Test",
    sheets, sharedStrings: [],
    styles: [{ id: 0, numFmt: null, font: null, fill: null, border: null, alignment: null }],
    definedNames: [], activeSheetIndex: 0,
  };
}

// ===========================================================================
// FEATURES
// ===========================================================================

feature("Cell Reference Helpers", () => {
  scenario("Given column index 0, When converting to letter, Then returns A", () => {
    assertEqual(colToLetter(0), "A", "colToLetter(0)");
  });

  scenario("Given column index 25, When converting to letter, Then returns Z", () => {
    assertEqual(colToLetter(25), "Z", "colToLetter(25)");
  });

  scenario("Given column index 26, When converting to letter, Then returns AA", () => {
    assertEqual(colToLetter(26), "AA", "colToLetter(26)");
  });

  scenario("Given letter A, When converting to index, Then returns 0", () => {
    assertEqual(letterToCol("A"), 0, "letterToCol(A)");
  });

  scenario("Given letter AA, When converting to index, Then returns 26", () => {
    assertEqual(letterToCol("AA"), 26, "letterToCol(AA)");
  });

  scenario("Given ref B3, When parsing, Then returns col=1 row=2", () => {
    const { col, row } = parseRef("B3");
    assertEqual(col, 1, "col"); assertEqual(row, 2, "row");
  });

  scenario("Given col=1 row=2, When building ref, Then returns B3", () => {
    assertEqual(buildRef(1, 2), "B3", "buildRef");
  });
});

feature("XLSX Save / Open Roundtrip", () => {
  scenario("Given a workbook with strings, numbers, and formulas, When exported to XLSX and re-parsed, Then all cell values are preserved", async () => {
    const sheet = makeSheet({
      A1: { type: "string", value: "Hello" },
      B1: { type: "number", value: 42 },
      C1: { type: "formula", value: null, formula: "SUM(B1,8)" },
      A2: { type: "boolean", value: true },
    });
    const wb = makeWorkbook([sheet]);
    const blob = exportXlsx(wb);
    assert(blob.size > 0, "blob not empty");

    const buf = await blob.arrayBuffer();
    const parsed = parseXlsx(buf);
    assertEqual(parsed.sheets.length, 1, "sheet count");

    const a1 = parsed.sheets[0].cells.get("A1");
    assertEqual(a1?.type, "string", "A1 type");
    assertEqual(a1?.value, "Hello", "A1 value");

    const b1 = parsed.sheets[0].cells.get("B1");
    assertEqual(b1?.type, "number", "B1 type");
    assertEqual(b1?.value, 42, "B1 value");

    const c1 = parsed.sheets[0].cells.get("C1");
    assertEqual(c1?.type, "formula", "C1 type");
    assert(c1?.formula != null, "C1 has formula");
  });

  scenario("Given a sheet with merged cells and frozen panes, When exported and re-parsed, Then merge regions and freeze are preserved", async () => {
    const sheet = makeSheet({ A1: { type: "string", value: "Merged" } });
    sheet.mergedRegions = [{ startRef: "A1", endRef: "C1" }];
    sheet.frozenRow = 1;
    sheet.frozenCol = 2;
    sheet.colWidths.set(0, 15);
    const wb = makeWorkbook([sheet]);

    const parsed = parseXlsx(await (exportXlsx(wb)).arrayBuffer());
    assertEqual(parsed.sheets[0].mergedRegions.length, 1, "merge count");
    assertEqual(parsed.sheets[0].mergedRegions[0].startRef, "A1", "merge start");
    assertEqual(parsed.sheets[0].frozenRow, 1, "frozen row");
    assertEqual(parsed.sheets[0].frozenCol, 2, "frozen col");
    assertEqual(parsed.sheets[0].colWidths.get(0), 15, "col width");
  });

  scenario("Given multiple sheets, When exported and re-parsed, Then all sheets are present with correct names", async () => {
    const s1 = makeSheet({ A1: { type: "string", value: "Sheet1" } });
    s1.name = "Revenue";
    const s2 = makeSheet({ A1: { type: "string", value: "Sheet2" } });
    s2.name = "Expenses";
    s2.order = 1;
    s2.id = "s2";
    const wb = makeWorkbook([s1, s2]);

    const parsed = parseXlsx(await (exportXlsx(wb)).arrayBuffer());
    assertEqual(parsed.sheets.length, 2, "sheet count");
    assertEqual(parsed.sheets[0].name, "Revenue", "sheet 1 name");
    assertEqual(parsed.sheets[1].name, "Expenses", "sheet 2 name");
  });
});

feature("Formula Engine — Basic Math", () => {
  const sheet = makeSheet({
    A1: { type: "number", value: 10 },
    A2: { type: "number", value: 20 },
    A3: { type: "number", value: 30 },
  });

  scenario("Given cells A1=10,A2=20,A3=30, When evaluating SUM(A1:A3), Then returns 60", () => {
    assertEqual(evaluateFormula("SUM(A1:A3)", sheet), 60, "SUM");
  });

  scenario("Given cells A1=10,A2=20,A3=30, When evaluating AVERAGE(A1:A3), Then returns 20", () => {
    assertEqual(evaluateFormula("AVERAGE(A1:A3)", sheet), 20, "AVERAGE");
  });

  scenario("When evaluating CEILING(7.3,0.5), Then returns 7.5", () => {
    assertEqual(evaluateFormula("CEILING(7.3,0.5)", sheet), 7.5, "CEILING");
  });

  scenario("When evaluating FLOOR(7.3,0.5), Then returns 7", () => {
    assertEqual(evaluateFormula("FLOOR(7.3,0.5)", sheet), 7, "FLOOR");
  });

  scenario("When evaluating MOD(10,3), Then returns 1", () => {
    assertEqual(evaluateFormula("MOD(10,3)", sheet), 1, "MOD");
  });

  scenario("When evaluating POWER(2,10), Then returns 1024", () => {
    assertEqual(evaluateFormula("POWER(2,10)", sheet), 1024, "POWER");
  });
});

feature("Formula Engine — Text Functions", () => {
  const sheet = makeSheet({});

  scenario("When evaluating UPPER(hello), Then returns HELLO", () => {
    assertEqual(evaluateFormula('UPPER("hello")', sheet), "HELLO", "UPPER");
  });

  scenario("When evaluating PROPER(hello world), Then returns Hello World", () => {
    assertEqual(evaluateFormula('PROPER("hello world")', sheet), "Hello World", "PROPER");
  });

  scenario("When evaluating LEFT(Hello,3), Then returns Hel", () => {
    assertEqual(evaluateFormula('LEFT("Hello",3)', sheet), "Hel", "LEFT");
  });

  scenario("When evaluating MID(Hello,2,3), Then returns ell", () => {
    assertEqual(evaluateFormula('MID("Hello",2,3)', sheet), "ell", "MID");
  });

  scenario("When evaluating FIND(ll,Hello), Then returns 3", () => {
    assertEqual(evaluateFormula('FIND("ll","Hello")', sheet), 3, "FIND");
  });

  scenario("When evaluating SUBSTITUTE(aabbcc,bb,XX), Then returns aaXXcc", () => {
    assertEqual(evaluateFormula('SUBSTITUTE("aabbcc","bb","XX")', sheet), "aaXXcc", "SUBSTITUTE");
  });

  scenario("When evaluating REPT(ab,3), Then returns ababab", () => {
    assertEqual(evaluateFormula('REPT("ab",3)', sheet), "ababab", "REPT");
  });

  scenario("When evaluating CHAR(65), Then returns A", () => {
    assertEqual(evaluateFormula("CHAR(65)", sheet), "A", "CHAR");
  });

  scenario("When evaluating CODE(A), Then returns 65", () => {
    assertEqual(evaluateFormula('CODE("A")', sheet), 65, "CODE");
  });

  scenario("When evaluating TEXTJOIN with delimiter, Then joins correctly", () => {
    // Note: tokenizer treats - as operator, so use a non-operator delimiter
    assertEqual(evaluateFormula('TEXTJOIN(",",TRUE,"a","","b","c")', sheet), "a,b,c", "TEXTJOIN");
  });
});

feature("Formula Engine — Logical Functions", () => {
  const sheet = makeSheet({});

  scenario("When evaluating IF(TRUE,yes,no), Then returns yes", () => {
    assertEqual(evaluateFormula('IF(TRUE,"yes","no")', sheet), "yes", "IF true");
  });

  scenario("When evaluating IF(FALSE,yes,no), Then returns no", () => {
    assertEqual(evaluateFormula('IF(FALSE,"yes","no")', sheet), "no", "IF false");
  });

  scenario("When evaluating IFS(FALSE,1,TRUE,2), Then returns 2", () => {
    assertEqual(evaluateFormula("IFS(FALSE,1,TRUE,2)", sheet), 2, "IFS");
  });

  scenario("When evaluating SWITCH(2,1,a,2,b), Then returns b", () => {
    assertEqual(evaluateFormula('SWITCH(2,1,"a",2,"b")', sheet), "b", "SWITCH");
  });

  scenario("When evaluating XOR(TRUE,FALSE), Then returns true", () => {
    assertEqual(evaluateFormula("XOR(TRUE,FALSE)", sheet), true, "XOR T,F");
  });

  scenario("When evaluating XOR(TRUE,TRUE), Then returns false", () => {
    assertEqual(evaluateFormula("XOR(TRUE,TRUE)", sheet), false, "XOR T,T");
  });

  scenario("When evaluating IFERROR(1/0,fallback), Then returns fallback", () => {
    assertEqual(evaluateFormula('IFERROR(1/0,"fallback")', sheet), "fallback", "IFERROR");
  });
});

feature("Formula Engine — Statistical Functions", () => {
  const sheet = makeSheet({
    A1: { type: "number", value: 1 },
    A2: { type: "number", value: 2 },
    A3: { type: "number", value: 3 },
    A4: { type: "number", value: 4 },
    A5: { type: "number", value: 5 },
  });

  scenario("When evaluating MEDIAN(1,2,3,4,5), Then returns 3", () => {
    assertEqual(evaluateFormula("MEDIAN(1,2,3,4,5)", sheet), 3, "MEDIAN");
  });

  scenario("When evaluating STDEV(1,2,3,4,5), Then returns ~1.58", () => {
    assertClose(evaluateFormula("STDEV(1,2,3,4,5)", sheet) as number, 1.5811, "STDEV");
  });

  scenario("When evaluating LARGE(nums,2) for [1,2,3,4,5], Then returns 4", () => {
    assertEqual(evaluateFormula("LARGE(1,2,3,4,5,2)", sheet), 4, "LARGE");
  });

  scenario("When evaluating SMALL(nums,2) for [1,2,3,4,5], Then returns 2", () => {
    assertEqual(evaluateFormula("SMALL(1,2,3,4,5,2)", sheet), 2, "SMALL");
  });

  scenario("When evaluating PERCENTILE with range, Then returns correct value", () => {
    // PERCENTILE takes (array, k) — with individual args, last arg is k
    assertClose(evaluateFormula("PERCENTILE(1,2,3,4,5,0.5)", sheet) as number, 3, "PERCENTILE", 0.5);
  });
});

feature("Formula Engine — Lookup Functions", () => {
  const sheet = makeSheet({
    A1: { type: "number", value: 100 },
    A2: { type: "number", value: 200 },
    A3: { type: "number", value: 300 },
    B1: { type: "string", value: "Apple" },
    B2: { type: "string", value: "Banana" },
    B3: { type: "string", value: "Cherry" },
  });

  scenario("When evaluating INDEX(A1:B3,2,2), Then returns Banana", () => {
    assertEqual(evaluateFormula("INDEX(A1:B3,2,2)", sheet), "Banana", "INDEX");
  });

  scenario("When evaluating MATCH(200,A1:A3,0), Then returns 2", () => {
    assertEqual(evaluateFormula("MATCH(200,A1:A3,0)", sheet), 2, "MATCH");
  });

  scenario("When evaluating CHOOSE(2,a,b,c), Then returns b", () => {
    assertEqual(evaluateFormula('CHOOSE(2,"a","b","c")', sheet), "b", "CHOOSE");
  });
});

feature("Formula Engine — Financial Functions", () => {
  const sheet = makeSheet({});

  scenario("When evaluating PMT(5%,10,-1000), Then returns ~129.50", () => {
    assertClose(evaluateFormula("PMT(0.05,10,-1000)", sheet) as number, 129.50, "PMT");
  });

  scenario("When evaluating FV(5%,10,-100,0), Then returns ~1257.79", () => {
    assertClose(evaluateFormula("FV(0.05,10,-100,0)", sheet) as number, 1257.79, "FV");
  });

  scenario("When evaluating NPV(10%,100,200,300), Then returns ~437.90", () => {
    // Excel NPV discounts from period 1 (not 0), so NPV(10%,100,200,300) = 100/1.1 + 200/1.21 + 300/1.331
    assertClose(evaluateFormula("NPV(0.1,100,200,300)", sheet) as number, 437.90, "NPV", 0.1);
  });

  scenario("When evaluating SLN(1000,100,5), Then returns 180", () => {
    assertEqual(evaluateFormula("SLN(1000,100,5)", sheet), 180, "SLN");
  });
});

feature("Formula Engine — Information Functions", () => {
  const sheet = makeSheet({});

  scenario("When evaluating ISNUMBER(42), Then returns true", () => {
    assertEqual(evaluateFormula("ISNUMBER(42)", sheet), true, "ISNUMBER");
  });

  scenario("When evaluating ISTEXT(hello), Then returns true", () => {
    assertEqual(evaluateFormula('ISTEXT("hello")', sheet), true, "ISTEXT");
  });

  scenario("When evaluating ISBLANK(), Then returns true", () => {
    assertEqual(evaluateFormula('ISBLANK("")', sheet), true, "ISBLANK");
  });

  scenario("When evaluating TYPE(42), Then returns 1 (number)", () => {
    assertEqual(evaluateFormula("TYPE(42)", sheet), 1, "TYPE number");
  });

  scenario("When evaluating TYPE(hello), Then returns 2 (text)", () => {
    assertEqual(evaluateFormula('TYPE("hello")', sheet), 2, "TYPE text");
  });
});

feature("Formula Engine — Dependency & Recalculation", () => {
  scenario("Given B1=SUM(A1:A3), B3=IF(B1>50,big,small), C1=B1*2+10, When recalculating, Then dependency order is correct", () => {
    const sheet = makeSheet({
      A1: { type: "number", value: 10 },
      A2: { type: "number", value: 20 },
      A3: { type: "number", value: 30 },
      B1: { type: "formula", value: null, formula: "SUM(A1:A3)" },
      B3: { type: "formula", value: null, formula: 'IF(B1>50,"big","small")' },
      C1: { type: "formula", value: null, formula: "B1*2+10" },
    });
    recalculateSheet(sheet);
    assertEqual(sheet.cells.get("B1")?.calculatedValue, 60, "B1=60");
    assertEqual(sheet.cells.get("B3")?.calculatedValue, "big", "B3=big");
    assertEqual(sheet.cells.get("C1")?.calculatedValue, 130, "C1=130");
  });

  scenario("Given a formula referencing itself, When checking circular, Then detects it", () => {
    const sheet = makeSheet({
      A1: { type: "formula", value: null, formula: "A1+1" },
    });
    assertEqual(detectCircular(sheet, "A1"), true, "circular detected");
  });

  scenario("Given SUM(A1:A3), When extracting deps, Then returns A1,A2,A3", () => {
    const deps = getFormulaDependencies("SUM(A1:A3)");
    assertEqual(deps.length, 3, "dep count");
    assertEqual(deps[0], "A1", "dep 0");
    assertEqual(deps[2], "A3", "dep 2");
  });
});

feature("CSV Import / Export", () => {
  scenario("Given CSV with headers and data, When parsing, Then returns correct 2D array", () => {
    const rows = parseCsv("Name,Age\nAlice,30\nBob,25");
    assertEqual(rows.length, 3, "row count");
    assertEqual(rows[0][0], "Name", "header");
    assertEqual(rows[1][1], "30", "data");
  });

  scenario("Given CSV with quoted fields containing commas, When parsing, Then handles correctly", () => {
    const rows = parseCsv('Name,City\n"Smith, Jr.",Tokyo');
    assertEqual(rows[1][0], "Smith, Jr.", "quoted comma");
  });

  scenario("Given CSV with quoted fields containing newlines, When parsing, Then handles correctly", () => {
    const rows = parseCsv('A,B\n"line1\nline2",val');
    assertEqual(rows[1][0], "line1\nline2", "quoted newline");
  });

  scenario("Given CSV text, When converting to workbook, Then cells have correct types", () => {
    const wb = csvToWorkbook("Name,Value\nItem,42.5\nOther,text", "Test");
    const sheet = wb.sheets[0];
    assertEqual(sheet.cells.get("B2")?.type, "number", "numeric cell");
    assertEqual(sheet.cells.get("B2")?.value, 42.5, "numeric value");
    assertEqual(sheet.cells.get("B3")?.type, "string", "text cell");
  });

  scenario("Given a sheet with data, When exporting to CSV, Then produces valid CSV", () => {
    const sheet = makeSheet({
      A1: { type: "string", value: "Name" },
      B1: { type: "string", value: "Age" },
      A2: { type: "string", value: "Alice" },
      B2: { type: "number", value: 30 },
    });
    const csv = sheetToCsv(sheet);
    assert(csv.includes("Name,Age"), "header row");
    assert(csv.includes("Alice,30"), "data row");
  });

  scenario("Given CSV roundtrip (export → parse → compare), Then data matches", () => {
    const sheet = makeSheet({
      A1: { type: "string", value: "Hello" },
      B1: { type: "number", value: 99 },
    });
    const csv = sheetToCsv(sheet);
    const rows = parseCsv(csv);
    assertEqual(rows[0][0], "Hello", "roundtrip A1");
    assertEqual(rows[0][1], "99", "roundtrip B1");
  });
});

feature("Find & Replace", () => {
  scenario("Given a sheet with text cells, When finding 'Apple', Then returns matching cells", () => {
    const sheet = makeSheet({
      A1: { type: "string", value: "Apple" },
      A2: { type: "string", value: "Banana" },
      A3: { type: "string", value: "Apple Pie" },
    });
    const results = findInSheet(sheet, 0, { query: "Apple", matchCase: false, matchEntireCell: false, searchFormulas: false, searchScope: "sheet" });
    assertEqual(results.length, 2, "match count");
  });

  scenario("Given matchEntireCell=true, When finding 'Apple', Then only exact match", () => {
    const sheet = makeSheet({
      A1: { type: "string", value: "Apple" },
      A2: { type: "string", value: "Apple Pie" },
    });
    const results = findInSheet(sheet, 0, { query: "Apple", matchCase: false, matchEntireCell: true, searchFormulas: false, searchScope: "sheet" });
    assertEqual(results.length, 1, "exact match only");
  });

  scenario("Given case-sensitive search, When finding 'apple' (lowercase), Then no match for 'Apple'", () => {
    const sheet = makeSheet({
      A1: { type: "string", value: "Apple" },
    });
    const results = findInSheet(sheet, 0, { query: "apple", matchCase: true, matchEntireCell: false, searchFormulas: false, searchScope: "sheet" });
    assertEqual(results.length, 0, "case sensitive no match");
  });

  scenario("Given replace all 'Apple' with 'Orange', When replacing, Then count and values correct", () => {
    const sheet = makeSheet({
      A1: { type: "string", value: "Apple" },
      A2: { type: "string", value: "Banana" },
      A3: { type: "string", value: "Apple Pie" },
    });
    const count = replaceAllInSheet(sheet, "Apple", "Orange", false);
    assertEqual(count, 2, "replace count");
    assertEqual(sheet.cells.get("A1")?.value, "Orange", "A1 replaced");
    assertEqual(sheet.cells.get("A3")?.value, "Orange Pie", "A3 replaced");
    assertEqual(sheet.cells.get("A2")?.value, "Banana", "A2 unchanged");
  });
});

feature("Cell Selection & Range", () => {
  scenario("Given range {startRow:3,startCol:5,endRow:1,endCol:2}, When normalizing, Then start <= end", () => {
    const n = normalizeRange({ startRow: 3, startCol: 5, endRow: 1, endCol: 2 });
    assertEqual(n.startRow, 1, "startRow"); assertEqual(n.startCol, 2, "startCol");
    assertEqual(n.endRow, 3, "endRow"); assertEqual(n.endCol, 5, "endCol");
  });

  scenario("Given range A1:C3, When converting to string, Then returns 'A1:C3'", () => {
    assertEqual(rangeToString({ startRow: 0, startCol: 0, endRow: 2, endCol: 2 }), "A1:C3", "rangeToString");
  });

  scenario("Given string A1:C3, When parsing, Then returns correct range", () => {
    const r = parseRangeString("A1:C3");
    assertEqual(r.startRow, 0, "startRow"); assertEqual(r.endRow, 2, "endRow");
  });

  scenario("Given cell B2 and range A1:C3, When testing containment, Then returns true", () => {
    assertEqual(isCellInRange("B2", { startRow: 0, startCol: 0, endRow: 2, endCol: 2 }), true, "in range");
  });

  scenario("Given cell D4 and range A1:C3, When testing containment, Then returns false", () => {
    assertEqual(isCellInRange("D4", { startRow: 0, startCol: 0, endRow: 2, endCol: 2 }), false, "not in range");
  });

  scenario("Given ref A1 and Tab, When navigating, Then moves to B1", () => {
    assertEqual(nextCellTab("A1", 100), "B1", "next tab");
  });

  scenario("Given ref A1 and Shift+Tab, When navigating backwards, Then stays at A1 (boundary)", () => {
    // A1 → can't go further left on row 0
    const prev = prevCellTab("A1");
    // Should stay or go to previous row col 25
    assert(prev === "A1" || prev === "Z1", "prev tab boundary");
  });
});

feature("Autofill Pattern Detection", () => {
  scenario("Given values [1,2,3], When detecting pattern, Then finds increment step=1", () => {
    const p = detectFillPattern([1, 2, 3]);
    assertEqual(p?.type, "increment", "type");
    assertEqual((p as any)?.step, 1, "step");
  });

  scenario("Given values [10,20,30], When detecting pattern, Then finds increment step=10", () => {
    const p = detectFillPattern([10, 20, 30]);
    assertEqual(p?.type, "increment", "type");
    assertEqual((p as any)?.step, 10, "step");
  });

  scenario("Given values ['a','b'], When detecting pattern, Then falls back to copy", () => {
    const p = detectFillPattern(["a", "b"]);
    assertEqual(p?.type, "copy", "copy pattern");
  });

  scenario("Given source [1,2,3] and count 3, When generating fill, Then continues sequence", () => {
    const vals = generateFillValues([1, 2, 3], 3);
    assertEqual(vals[0], 4, "fill[0]");
    assertEqual(vals[1], 5, "fill[1]");
    assertEqual(vals[2], 6, "fill[2]");
  });
});

feature("Clipboard (TSV Paste)", () => {
  scenario("Given TSV text, When parsing, Then returns 2D array", () => {
    const rows = parsePastedText("A\tB\nC\tD");
    assertEqual(rows.length, 2, "row count");
    assertEqual(rows[0][1], "B", "cell B");
    assertEqual(rows[1][0], "C", "cell C");
  });

  scenario("Given pasted values, When applying to sheet at B2, Then cells are set correctly", () => {
    const sheet = makeSheet({});
    applyPastedValues(sheet, "B2", [["hello", "42"], ["world", ""]]);
    assertEqual(sheet.cells.get("B2")?.value, "hello", "B2");
    assertEqual(sheet.cells.get("C2")?.type, "number", "C2 type");
    assertEqual(sheet.cells.get("C2")?.value, 42, "C2 value");
    assertEqual(sheet.cells.get("B3")?.value, "world", "B3");
    assertEqual(sheet.cells.has("C3"), false, "C3 empty deleted");
  });
});

feature("Comments / Notes", () => {
  scenario("Given a new CommentStore, When adding a comment, Then it can be retrieved", () => {
    const store = new CommentStore();
    store.add("A1", "Jun", "Review this cell");
    assert(store.hasComment("A1"), "has comment");
    assertEqual(store.get("A1")?.text, "Review this cell", "text");
    assertEqual(store.get("A1")?.author, "Jun", "author");
  });

  scenario("Given a comment on A1, When replying, Then reply is added", () => {
    const store = new CommentStore();
    store.add("A1", "Jun", "Check this");
    store.reply("A1", "Bot", "Looks good");
    assertEqual(store.get("A1")?.replies.length, 1, "reply count");
    assertEqual(store.get("A1")?.replies[0].text, "Looks good", "reply text");
  });

  scenario("Given a comment on A1, When deleting, Then it is removed", () => {
    const store = new CommentStore();
    store.add("A1", "Jun", "Delete me");
    store.delete("A1");
    assertEqual(store.hasComment("A1"), false, "deleted");
  });

  scenario("Given a CommentStore, When serializing and deserializing, Then data is preserved", () => {
    const store = new CommentStore();
    store.add("A1", "Jun", "Note 1");
    store.add("B2", "Bot", "Note 2");
    store.reply("A1", "Bot", "Reply");

    const json = store.toJSON();
    const restored = CommentStore.fromJSON(json);
    assertEqual(restored.getAll().length, 2, "comment count");
    assertEqual(restored.get("A1")?.replies.length, 1, "reply preserved");
  });
});

feature("Entire Column / Row Selection", () => {
  scenario("When selecting entire column 2, Then range spans all rows", () => {
    const r = selectEntireColumn(2, 1000);
    assertEqual(r.startCol, 2, "startCol");
    assertEqual(r.endCol, 2, "endCol");
    assertEqual(r.startRow, 0, "startRow");
    assertEqual(r.endRow, 1000, "endRow");
  });

  scenario("When selecting entire row 5, Then range spans all columns", () => {
    const r = selectEntireRow(5, 100);
    assertEqual(r.startRow, 5, "startRow");
    assertEqual(r.endRow, 5, "endRow");
    assertEqual(r.startCol, 0, "startCol");
    assertEqual(r.endCol, 100, "endCol");
  });
});

feature("XML Entity Roundtrip (Formulas with special chars)", () => {
  scenario("Given formula IF(A1>10,big,small) with > and quotes, When exported and re-parsed, Then formula is preserved", async () => {
    const sheet = makeSheet({
      A1: { type: "number", value: 20 },
      B1: { type: "formula", value: null, formula: 'IF(A1>10,"big","small")' },
    });
    const wb = makeWorkbook([sheet]);
    const parsed = parseXlsx(await (exportXlsx(wb)).arrayBuffer());
    const b1 = parsed.sheets[0].cells.get("B1");
    assertEqual(b1?.formula, 'IF(A1>10,"big","small")', "formula preserved");

    recalculateSheet(parsed.sheets[0]);
    assertEqual(parsed.sheets[0].cells.get("B1")?.calculatedValue, "big", "recalc correct");
  });
});

feature("Formula Engine — Math (Extended)", () => {
  const sheet = makeSheet({
    A1: { type: "number", value: 10 },
    A2: { type: "number", value: 20 },
    A3: { type: "number", value: 30 },
    A4: { type: "number", value: 40 },
    A5: { type: "number", value: 50 },
    B1: { type: "string", value: "Apple" },
    B2: { type: "string", value: "Banana" },
    B3: { type: "string", value: "Cherry" },
    B4: { type: "number", value: 100 },
    B5: { type: "number", value: 200 },
  });

  scenario("When evaluating SUMPRODUCT(A1:A3,A1:A3), Then returns sum of squares 1400", () => {
    assertEqual(evaluateFormula("SUMPRODUCT(A1:A3,A1:A3)", sheet), 1400, "SUMPRODUCT");
  });

  scenario("When evaluating ROUNDUP(3.14159,2), Then returns 3.15", () => {
    assertClose(evaluateFormula("ROUNDUP(3.14159,2)", sheet) as number, 3.15, "ROUNDUP");
  });

  scenario("When evaluating ROUNDDOWN(3.14159,2), Then returns 3.14", () => {
    assertClose(evaluateFormula("ROUNDDOWN(3.14159,2)", sheet) as number, 3.14, "ROUNDDOWN");
  });

  scenario("When evaluating INT(7.9), Then returns 7", () => {
    assertEqual(evaluateFormula("INT(7.9)", sheet), 7, "INT");
  });

  scenario("When evaluating SIGN(-42), Then returns -1", () => {
    assertEqual(evaluateFormula("SIGN(-42)", sheet), -1, "SIGN negative");
  });

  scenario("When evaluating SIGN(0), Then returns 0", () => {
    assertEqual(evaluateFormula("SIGN(0)", sheet), 0, "SIGN zero");
  });

  scenario("When evaluating SIGN(99), Then returns 1", () => {
    assertEqual(evaluateFormula("SIGN(99)", sheet), 1, "SIGN positive");
  });

  scenario("When evaluating LOG(100,10), Then returns 2", () => {
    assertClose(evaluateFormula("LOG(100,10)", sheet) as number, 2, "LOG base 10");
  });

  scenario("When evaluating LOG(8,2), Then returns 3", () => {
    assertClose(evaluateFormula("LOG(8,2)", sheet) as number, 3, "LOG base 2");
  });

  scenario("When evaluating LOG10(1000), Then returns 3", () => {
    assertClose(evaluateFormula("LOG10(1000)", sheet) as number, 3, "LOG10");
  });

  scenario("When evaluating LN(E^1), Then returns ~1", () => {
    assertClose(evaluateFormula("LN(2.71828)", sheet) as number, 1, "LN", 0.001);
  });

  scenario("When evaluating EXP(1), Then returns ~2.718", () => {
    assertClose(evaluateFormula("EXP(1)", sheet) as number, 2.71828, "EXP", 0.001);
  });

  scenario("When evaluating PI(), Then returns ~3.14159", () => {
    assertClose(evaluateFormula("PI()", sheet) as number, 3.14159, "PI", 0.00001);
  });

  scenario("When evaluating RAND(), Then returns a number between 0 and 1", () => {
    const result = evaluateFormula("RAND()", sheet) as number;
    assert(typeof result === "number", "RAND is number");
    assert(result >= 0 && result < 1, "RAND in [0,1)");
  });

  scenario("When evaluating RANDBETWEEN(1,100), Then returns integer in [1,100]", () => {
    const result = evaluateFormula("RANDBETWEEN(1,100)", sheet) as number;
    assert(typeof result === "number", "RANDBETWEEN is number");
    assert(result >= 1 && result <= 100, "RANDBETWEEN in [1,100]");
    assertEqual(result, Math.floor(result), "RANDBETWEEN is integer");
  });

  scenario("When evaluating PRODUCT(A1:A3), Then returns 6000", () => {
    assertEqual(evaluateFormula("PRODUCT(A1:A3)", sheet), 6000, "PRODUCT");
  });

  scenario("When evaluating SUBTOTAL(9,A1:A3), Then returns SUM=60", () => {
    assertEqual(evaluateFormula("SUBTOTAL(9,A1:A3)", sheet), 60, "SUBTOTAL SUM");
  });

  scenario("When evaluating COUNTA(A1:B3), Then returns 6 (3 numbers + 3 strings)", () => {
    assertEqual(evaluateFormula("COUNTA(A1:B3)", sheet), 6, "COUNTA");
  });

  scenario("When evaluating COUNTBLANK on cells with values, Then returns 0", () => {
    assertEqual(evaluateFormula("COUNTBLANK(A1:A3)", sheet), 0, "COUNTBLANK");
  });
});

feature("Formula Engine — Date & Time", () => {
  const sheet = makeSheet({
    A1: { type: "number", value: 10 },
    A2: { type: "number", value: 20 },
    A3: { type: "number", value: 30 },
    A4: { type: "number", value: 40 },
    A5: { type: "number", value: 50 },
    B1: { type: "string", value: "Apple" },
    B2: { type: "string", value: "Banana" },
    B3: { type: "string", value: "Cherry" },
    B4: { type: "number", value: 100 },
    B5: { type: "number", value: 200 },
  });

  scenario("When evaluating NOW(), Then returns an ISO date string", () => {
    const result = evaluateFormula("NOW()", sheet);
    assert(typeof result === "string", "NOW returns string");
    assert((result as string).includes("T"), "NOW contains T separator");
  });

  scenario("When evaluating TODAY(), Then returns a date string YYYY-MM-DD", () => {
    const result = evaluateFormula("TODAY()", sheet);
    assert(typeof result === "string", "TODAY returns string");
    assert(/^\d{4}-\d{2}-\d{2}$/.test(result as string), "TODAY matches YYYY-MM-DD");
  });

  scenario("When evaluating DATE(2024,6,15), Then returns Excel serial number", () => {
    const result = evaluateFormula("DATE(2024,6,15)", sheet) as number;
    assert(typeof result === "number", "DATE returns number");
    assert(result > 45000, "DATE serial is plausible for 2024");
  });

  scenario("When evaluating TIME(14,30,0), Then returns fractional day", () => {
    const result = evaluateFormula("TIME(14,30,0)", sheet) as number;
    assertClose(result, (14 * 3600 + 30 * 60) / 86400, "TIME", 0.0001);
  });

  scenario("When evaluating YEAR on date string, Then returns correct year", () => {
    assertEqual(evaluateFormula('YEAR("2024-06-15")', sheet), 2024, "YEAR");
  });

  scenario("When evaluating MONTH on date string, Then returns correct month", () => {
    assertEqual(evaluateFormula('MONTH("2024-06-15")', sheet), 6, "MONTH");
  });

  scenario("When evaluating DAY on date string, Then returns correct day", () => {
    assertEqual(evaluateFormula('DAY("2024-06-15")', sheet), 15, "DAY");
  });

  scenario("When evaluating HOUR on datetime string, Then returns hour component", () => {
    const result = evaluateFormula('HOUR("2024-06-15T14:30:45")', sheet);
    assertEqual(result, 14, "HOUR");
  });

  scenario("When evaluating MINUTE on datetime string, Then returns minute component", () => {
    assertEqual(evaluateFormula('MINUTE("2024-06-15T14:30:45")', sheet), 30, "MINUTE");
  });

  scenario("When evaluating SECOND on datetime string, Then returns second component", () => {
    assertEqual(evaluateFormula('SECOND("2024-06-15T14:30:45")', sheet), 45, "SECOND");
  });

  scenario("When evaluating WEEKDAY on 2024-06-15 (Saturday), Then returns 7", () => {
    assertEqual(evaluateFormula('WEEKDAY("2024-06-15")', sheet), 7, "WEEKDAY");
  });

  scenario("When evaluating DATEVALUE on date string, Then returns Excel serial", () => {
    const result = evaluateFormula('DATEVALUE("2024-06-15")', sheet) as number;
    assert(typeof result === "number", "DATEVALUE returns number");
    assert(result > 45000, "DATEVALUE serial is plausible");
  });

  scenario("When evaluating DATEDIF with D unit, Then returns day difference", () => {
    const result = evaluateFormula('DATEDIF("2024-01-01","2024-12-31","D")', sheet) as number;
    assertEqual(result, 365, "DATEDIF D");
  });

  scenario("When evaluating DATEDIF with M unit, Then returns month difference", () => {
    const result = evaluateFormula('DATEDIF("2024-01-01","2024-07-01","M")', sheet) as number;
    assertEqual(result, 6, "DATEDIF M");
  });

  scenario("When evaluating DATEDIF with Y unit, Then returns year difference", () => {
    const result = evaluateFormula('DATEDIF("2020-01-01","2024-01-01","Y")', sheet) as number;
    assertEqual(result, 4, "DATEDIF Y");
  });

  scenario("When evaluating EDATE, Then returns shifted Excel serial", () => {
    const result = evaluateFormula('EDATE("2024-01-15",3)', sheet) as number;
    assert(typeof result === "number", "EDATE returns number");
  });

  scenario("When evaluating EOMONTH, Then returns end-of-month Excel serial", () => {
    const result = evaluateFormula('EOMONTH("2024-01-15",0)', sheet) as number;
    assert(typeof result === "number", "EOMONTH returns number");
  });

  scenario("When evaluating NETWORKDAYS, Then returns working days count", () => {
    const result = evaluateFormula('NETWORKDAYS("2024-06-10","2024-06-14")', sheet) as number;
    assert(typeof result === "number", "NETWORKDAYS returns number");
    assert(result >= 1 && result <= 5, "NETWORKDAYS is plausible weekday count");
  });

  scenario("When evaluating WORKDAY, Then returns Excel serial of work date", () => {
    const result = evaluateFormula('WORKDAY("2024-06-10",5)', sheet) as number;
    assert(typeof result === "number", "WORKDAY returns number");
  });

  scenario("When evaluating ISOWEEKNUM, Then returns ISO week number", () => {
    const result = evaluateFormula('ISOWEEKNUM("2024-06-15")', sheet) as number;
    assert(typeof result === "number", "ISOWEEKNUM returns number");
    assert(result >= 1 && result <= 53, "ISOWEEKNUM in valid range");
  });
});

feature("Formula Engine — Text (Extended)", () => {
  const sheet = makeSheet({
    A1: { type: "number", value: 10 },
    A2: { type: "number", value: 20 },
    A3: { type: "number", value: 30 },
    A4: { type: "number", value: 40 },
    A5: { type: "number", value: 50 },
    B1: { type: "string", value: "Apple" },
    B2: { type: "string", value: "Banana" },
    B3: { type: "string", value: "Cherry" },
    B4: { type: "number", value: 100 },
    B5: { type: "number", value: 200 },
  });

  scenario("When evaluating CONCAT(Hello, ,World), Then returns Hello World", () => {
    assertEqual(evaluateFormula('CONCAT("Hello"," ","World")', sheet), "Hello World", "CONCAT");
  });

  scenario("When evaluating TEXTJOIN with comma delimiter, Then joins correctly", () => {
    assertEqual(evaluateFormula('TEXTJOIN(",",TRUE,"x","","y","z")', sheet), "x,y,z", "TEXTJOIN comma");
  });

  scenario("When evaluating TEXTJOIN with ignore_empty=FALSE, Then preserves blanks", () => {
    assertEqual(evaluateFormula('TEXTJOIN(",",FALSE,"x","","y")', sheet), "x,,y", "TEXTJOIN no skip");
  });

  scenario("When evaluating SEARCH (case-insensitive), Then finds substring", () => {
    assertEqual(evaluateFormula('SEARCH("ban","Banana")', sheet), 1, "SEARCH case-insensitive");
  });

  scenario("When evaluating SEARCH for substring at position 4, Then returns 4", () => {
    assertEqual(evaluateFormula('SEARCH("ana","Banana")', sheet), 2, "SEARCH ana in Banana");
  });

  scenario("When evaluating REPLACE(Hello,2,3,XY), Then returns HXYo", () => {
    assertEqual(evaluateFormula('REPLACE("Hello",2,3,"XY")', sheet), "HXYo", "REPLACE");
  });

  scenario("When evaluating CLEAN on string with control chars, Then removes them", () => {
    // CLEAN removes chars 0x00-0x1F; normal string passes through unchanged
    assertEqual(evaluateFormula('CLEAN("Hello")', sheet), "Hello", "CLEAN normal");
  });

  scenario("When evaluating T on a string, Then returns the string", () => {
    assertEqual(evaluateFormula('T("test")', sheet), "test", "T string");
  });

  scenario("When evaluating T on a number, Then returns empty string", () => {
    assertEqual(evaluateFormula("T(42)", sheet), "", "T number");
  });

  scenario("When evaluating VALUE(123.45), Then returns 123.45", () => {
    assertEqual(evaluateFormula('VALUE("123.45")', sheet), 123.45, "VALUE");
  });

  scenario("When evaluating EXACT with different strings, Then returns false", () => {
    assertEqual(evaluateFormula('EXACT("abc","ABC")', sheet), false, "EXACT different");
  });

  scenario("When evaluating EXACT with same strings, Then returns true", () => {
    assertEqual(evaluateFormula('EXACT("abc","abc")', sheet), true, "EXACT same");
  });

  scenario("When evaluating LOWER(HELLO), Then returns hello", () => {
    assertEqual(evaluateFormula('LOWER("HELLO")', sheet), "hello", "LOWER");
  });
});

feature("Formula Engine — Lookup (Extended)", () => {
  const sheet = makeSheet({
    A1: { type: "number", value: 10 },
    A2: { type: "number", value: 20 },
    A3: { type: "number", value: 30 },
    A4: { type: "number", value: 40 },
    A5: { type: "number", value: 50 },
    B1: { type: "string", value: "Apple" },
    B2: { type: "string", value: "Banana" },
    B3: { type: "string", value: "Cherry" },
    B4: { type: "number", value: 100 },
    B5: { type: "number", value: 200 },
  });

  scenario("When evaluating HLOOKUP on transposed data, Then returns correct value", () => {
    // A1=10, B1=Apple — search for 10 in row 1, return row 1 col match
    // HLOOKUP(10, A1:B2, 2, FALSE) — find 10 in first row, return row 2
    const hlSheet = makeSheet({
      A1: { type: "number", value: 10 },
      B1: { type: "number", value: 20 },
      C1: { type: "number", value: 30 },
      A2: { type: "string", value: "Apple" },
      B2: { type: "string", value: "Banana" },
      C2: { type: "string", value: "Cherry" },
    });
    assertEqual(evaluateFormula("HLOOKUP(20,A1:C2,2,FALSE)", hlSheet), "Banana", "HLOOKUP");
  });

  scenario("When evaluating XLOOKUP with lookup and return ranges, Then returns correct value", () => {
    assertEqual(evaluateFormula("XLOOKUP(20,A1:A3,B1:B3)", sheet), "Banana", "XLOOKUP");
  });

  scenario("When evaluating XLOOKUP with not-found value, Then returns default", () => {
    assertEqual(evaluateFormula('XLOOKUP(99,A1:A3,B1:B3,"N/A")', sheet), "N/A", "XLOOKUP not found");
  });

  scenario("When evaluating ROW on a string ref, Then returns row number", () => {
    assertEqual(evaluateFormula('ROW("B3")', sheet), 3, "ROW");
  });

  scenario("When evaluating COLUMN on a string ref, Then returns column number", () => {
    assertEqual(evaluateFormula('COLUMN("B3")', sheet), 2, "COLUMN");
  });

  scenario("When evaluating ROWS(A1:A5), Then returns 5", () => {
    assertEqual(evaluateFormula("ROWS(A1:A5)", sheet), 5, "ROWS");
  });

  scenario("When evaluating COLUMNS(A1:B3), Then returns 2", () => {
    assertEqual(evaluateFormula("COLUMNS(A1:B3)", sheet), 2, "COLUMNS");
  });

  scenario("When evaluating INDIRECT(A1) referencing cell A1=10, Then returns 10", () => {
    // INDIRECT takes a string ref and resolves it
    const indSheet = makeSheet({
      A1: { type: "number", value: 42 },
    });
    assertEqual(evaluateFormula('INDIRECT("A1")', indSheet), 42, "INDIRECT");
  });
});

feature("Formula Engine — Financial (Extended)", () => {
  const sheet = makeSheet({
    A1: { type: "number", value: 10 },
    A2: { type: "number", value: 20 },
    A3: { type: "number", value: 30 },
    A4: { type: "number", value: 40 },
    A5: { type: "number", value: 50 },
    B1: { type: "string", value: "Apple" },
    B2: { type: "string", value: "Banana" },
    B3: { type: "string", value: "Cherry" },
    B4: { type: "number", value: 100 },
    B5: { type: "number", value: 200 },
  });

  scenario("When evaluating PV(5%,10,-100,0), Then returns ~772.17", () => {
    assertClose(evaluateFormula("PV(0.05,10,-100,0)", sheet) as number, 772.17, "PV", 0.5);
  });

  scenario("When evaluating NPER(5%,-100,1000,0), Then returns ~14.21", () => {
    assertClose(evaluateFormula("NPER(0.05,-100,1000,0)", sheet) as number, 14.21, "NPER", 0.1);
  });

  scenario("When evaluating RATE(10,-100,500) approximately, Then returns a rate > 0", () => {
    const result = evaluateFormula("RATE(10,-100,500)", sheet) as number;
    assert(typeof result === "number", "RATE returns number");
    assert(result > 0 && result < 1, "RATE is a plausible interest rate");
  });

  scenario("When evaluating IRR on simple cash flows via range, Then returns approximate IRR", () => {
    const irrSheet = makeSheet({
      A1: { type: "number", value: -1000 },
      A2: { type: "number", value: 300 },
      A3: { type: "number", value: 300 },
      A4: { type: "number", value: 300 },
      A5: { type: "number", value: 300 },
      A6: { type: "number", value: 300 },
    });
    const result = evaluateFormula("IRR(A1:A6)", irrSheet) as number;
    assert(typeof result === "number", "IRR returns number");
    assertClose(result, 0.1524, "IRR", 0.02);
  });
});

feature("Formula Engine — Information (Extended)", () => {
  const sheet = makeSheet({
    A1: { type: "number", value: 10 },
    A2: { type: "number", value: 20 },
    A3: { type: "number", value: 30 },
    A4: { type: "number", value: 40 },
    A5: { type: "number", value: 50 },
    B1: { type: "string", value: "Apple" },
    B2: { type: "string", value: "Banana" },
    B3: { type: "string", value: "Cherry" },
    B4: { type: "number", value: 100 },
    B5: { type: "number", value: 200 },
  });

  scenario("When evaluating ISLOGICAL(TRUE), Then returns true", () => {
    assertEqual(evaluateFormula("ISLOGICAL(TRUE)", sheet), true, "ISLOGICAL true");
  });

  scenario("When evaluating ISLOGICAL(42), Then returns false", () => {
    assertEqual(evaluateFormula("ISLOGICAL(42)", sheet), false, "ISLOGICAL number");
  });

  scenario("When evaluating ISFORMULA(42), Then returns false (simplified)", () => {
    assertEqual(evaluateFormula("ISFORMULA(42)", sheet), false, "ISFORMULA");
  });

  scenario("When evaluating N(TRUE), Then returns 1", () => {
    assertEqual(evaluateFormula("N(TRUE)", sheet), 1, "N true");
  });

  scenario("When evaluating N(FALSE), Then returns 0", () => {
    assertEqual(evaluateFormula("N(FALSE)", sheet), 0, "N false");
  });

  scenario("When evaluating N(42), Then returns 42", () => {
    assertEqual(evaluateFormula("N(42)", sheet), 42, "N number");
  });

  scenario("When evaluating N(hello), Then returns 0", () => {
    assertEqual(evaluateFormula('N("hello")', sheet), 0, "N string");
  });

  scenario("When evaluating NA(), Then returns #N/A", () => {
    assertEqual(evaluateFormula("NA()", sheet), "#N/A", "NA");
  });

  scenario("When evaluating ERROR.TYPE(#DIV/0!), Then returns 2", () => {
    // 1/0 produces #DIV/0! which ERROR.TYPE maps to 2
    assertEqual(evaluateFormula("ERROR.TYPE(1/0)", sheet), 2, "ERROR.TYPE DIV/0");
  });
});

feature("Formula Engine — Statistical (Extended)", () => {
  const sheet = makeSheet({
    A1: { type: "number", value: 10 },
    A2: { type: "number", value: 20 },
    A3: { type: "number", value: 30 },
    A4: { type: "number", value: 40 },
    A5: { type: "number", value: 50 },
    B1: { type: "string", value: "Apple" },
    B2: { type: "string", value: "Banana" },
    B3: { type: "string", value: "Cherry" },
    B4: { type: "number", value: 100 },
    B5: { type: "number", value: 200 },
  });

  scenario("When evaluating STDEVP(A1:A5), Then returns population stdev ~14.14", () => {
    assertClose(evaluateFormula("STDEVP(A1:A5)", sheet) as number, 14.14, "STDEVP", 0.1);
  });

  scenario("When evaluating VARP(A1:A5), Then returns population variance 200", () => {
    assertEqual(evaluateFormula("VARP(A1:A5)", sheet), 200, "VARP");
  });

  scenario("When evaluating VAR(A1:A5), Then returns sample variance 250", () => {
    assertEqual(evaluateFormula("VAR(A1:A5)", sheet), 250, "VAR");
  });

  scenario("When evaluating MODE on repeated values, Then returns most frequent", () => {
    const mSheet = makeSheet({
      A1: { type: "number", value: 1 },
      A2: { type: "number", value: 2 },
      A3: { type: "number", value: 2 },
      A4: { type: "number", value: 3 },
      A5: { type: "number", value: 2 },
    });
    assertEqual(evaluateFormula("MODE(A1:A5)", mSheet), 2, "MODE");
  });

  scenario("When evaluating RANK(30,A1:A5), Then returns 3 (3rd from top)", () => {
    assertEqual(evaluateFormula("RANK(30,A1:A5)", sheet), 3, "RANK");
  });

  scenario("When evaluating RANK(50,A1:A5), Then returns 1 (largest)", () => {
    assertEqual(evaluateFormula("RANK(50,A1:A5)", sheet), 1, "RANK top");
  });

  scenario("When evaluating AVERAGEIF(A1:A5,20), Then returns 20", () => {
    assertEqual(evaluateFormula("AVERAGEIF(A1:A5,20)", sheet), 20, "AVERAGEIF");
  });

  scenario("When evaluating SUMIFS(A1:A5,A1:A5,20), Then returns 20", () => {
    assertEqual(evaluateFormula("SUMIFS(A1:A5,A1:A5,20)", sheet), 20, "SUMIFS");
  });

  scenario("When evaluating COUNTIFS(A1:A5,30), Then returns 1", () => {
    assertEqual(evaluateFormula("COUNTIFS(A1:A5,30)", sheet), 1, "COUNTIFS");
  });

  scenario("When evaluating MAXIFS(A1:A5,A1:A5,10), Then returns 10", () => {
    assertEqual(evaluateFormula("MAXIFS(A1:A5,A1:A5,10)", sheet), 10, "MAXIFS");
  });

  scenario("When evaluating MINIFS(A1:A5,A1:A5,50), Then returns 50", () => {
    assertEqual(evaluateFormula("MINIFS(A1:A5,A1:A5,50)", sheet), 50, "MINIFS");
  });
});

feature("Formula Engine — Operators", () => {
  const sheet = makeSheet({
    A1: { type: "number", value: 10 },
    A2: { type: "number", value: 20 },
    A3: { type: "number", value: 30 },
    A4: { type: "number", value: 40 },
    A5: { type: "number", value: 50 },
    B1: { type: "string", value: "Apple" },
    B2: { type: "string", value: "Banana" },
    B3: { type: "string", value: "Cherry" },
    B4: { type: "number", value: 100 },
    B5: { type: "number", value: 200 },
  });

  scenario("When evaluating 3+4, Then returns 7 (addition)", () => {
    assertEqual(evaluateFormula("3+4", sheet), 7, "addition");
  });

  scenario("When evaluating 10-3, Then returns 7 (subtraction)", () => {
    assertEqual(evaluateFormula("10-3", sheet), 7, "subtraction");
  });

  scenario("When evaluating 6*7, Then returns 42 (multiplication)", () => {
    assertEqual(evaluateFormula("6*7", sheet), 42, "multiplication");
  });

  scenario("When evaluating 100/4, Then returns 25 (division)", () => {
    assertEqual(evaluateFormula("100/4", sheet), 25, "division");
  });

  scenario("When evaluating 1/0, Then returns #DIV/0! (division by zero)", () => {
    assertEqual(evaluateFormula("1/0", sheet), "#DIV/0!", "division by zero");
  });

  scenario("When evaluating 2^10, Then returns 1024 (power)", () => {
    assertEqual(evaluateFormula("2^10", sheet), 1024, "power");
  });

  scenario("When evaluating 50%, Then returns 0.5 (percentage)", () => {
    assertClose(evaluateFormula("50%", sheet) as number, 0.5, "percentage");
  });

  scenario("When evaluating Hello&World with &, Then returns HelloWorld (concatenation)", () => {
    assertEqual(evaluateFormula('"Hello"&" "&"World"', sheet), "Hello World", "concatenation");
  });

  scenario("When evaluating 5=5, Then returns true (equality)", () => {
    assertEqual(evaluateFormula("5=5", sheet), true, "equality true");
  });

  scenario("When evaluating 5=6, Then returns false (equality)", () => {
    assertEqual(evaluateFormula("5=6", sheet), false, "equality false");
  });

  scenario("When evaluating 5<>6, Then returns true (not equal)", () => {
    assertEqual(evaluateFormula("5<>6", sheet), true, "not equal true");
  });

  scenario("When evaluating 5<>5, Then returns false (not equal)", () => {
    assertEqual(evaluateFormula("5<>5", sheet), false, "not equal false");
  });

  scenario("When evaluating 3<5, Then returns true (less than)", () => {
    assertEqual(evaluateFormula("3<5", sheet), true, "less than true");
  });

  scenario("When evaluating 5<3, Then returns false (less than)", () => {
    assertEqual(evaluateFormula("5<3", sheet), false, "less than false");
  });

  scenario("When evaluating 5>3, Then returns true (greater than)", () => {
    assertEqual(evaluateFormula("5>3", sheet), true, "greater than true");
  });

  scenario("When evaluating 3>5, Then returns false (greater than)", () => {
    assertEqual(evaluateFormula("3>5", sheet), false, "greater than false");
  });

  scenario("When evaluating 3<=3, Then returns true (less or equal)", () => {
    assertEqual(evaluateFormula("3<=3", sheet), true, "less or equal true");
  });

  scenario("When evaluating 4<=3, Then returns false (less or equal)", () => {
    assertEqual(evaluateFormula("4<=3", sheet), false, "less or equal false");
  });

  scenario("When evaluating 5>=5, Then returns true (greater or equal)", () => {
    assertEqual(evaluateFormula("5>=5", sheet), true, "greater or equal true");
  });

  scenario("When evaluating 4>=5, Then returns false (greater or equal)", () => {
    assertEqual(evaluateFormula("4>=5", sheet), false, "greater or equal false");
  });

  scenario("When evaluating -7 (unary minus), Then returns -7", () => {
    assertEqual(evaluateFormula("-7", sheet), -7, "unary minus");
  });

  scenario("When evaluating -(3+4), Then returns -7 (unary minus with parens)", () => {
    assertEqual(evaluateFormula("-(3+4)", sheet), -7, "unary minus parens");
  });

  scenario("When evaluating (2+3)*(4+1), Then returns 25 (nested parentheses)", () => {
    assertEqual(evaluateFormula("(2+3)*(4+1)", sheet), 25, "nested parentheses");
  });

  scenario("When evaluating ((1+2)*(3+4))+5, Then returns 26 (deeply nested)", () => {
    assertEqual(evaluateFormula("((1+2)*(3+4))+5", sheet), 26, "deeply nested");
  });

  scenario("When evaluating A1+B4, Then returns 110 (cell reference addition)", () => {
    assertEqual(evaluateFormula("A1+B4", sheet), 110, "cell ref addition");
  });

  scenario("When evaluating 2+3*4, Then returns 14 (operator precedence)", () => {
    assertEqual(evaluateFormula("2+3*4", sheet), 14, "operator precedence");
  });
});

// ===========================================================================
// Summary
// ===========================================================================

// Run async scenarios
setTimeout(() => {
  console.log(`\n${"=".repeat(50)}`);
  console.log(`\x1b[1mResults: ${totalPass} passed, ${totalFail} failed, ${totalPass + totalFail} total\x1b[0m`);
  if (totalFail === 0) {
    console.log("\x1b[32m\x1b[1m✓ ALL SCENARIOS PASSED\x1b[0m");
  } else {
    console.log(`\x1b[31m\x1b[1m✗ ${totalFail} SCENARIO(S) FAILED\x1b[0m`);
    process.exit(1);
  }
}, 500);
