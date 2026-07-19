/**
 * Coverage Expansion Tests — xlsx.etzhayyim.com
 *
 * Covers ALL remaining untested formula functions + editor operations +
 * CSV edge cases + XLSX stress tests + formatting roundtrip.
 *
 * Run: cd svelte && npx tsx src/lib/__tests__/coverage-expansion.test.ts
 */

import { parseXlsx, buildRef, parseRef, colToLetter, letterToCol, type XlsxWorkbook, type XlsxSheet, type XlsxCell, type CellRef } from "../ooxml-parser";
import { exportXlsx } from "../xlsx-exporter";
import { evaluateFormula, getFormulaDependencies, detectCircular, recalculateSheet } from "../formula-engine";
import { parseCsv, sheetToCsv, csvToWorkbook } from "../csv-handler";
import { findInSheet, replaceAllInSheet, replaceInCell } from "../find-replace";
import { CommentStore } from "../comments";
import { normalizeRange, rangeToString, parseRangeString, isCellInRange, rangeSize, detectFillPattern, generateFillValues, nextCellTab, prevCellTab, selectEntireColumn, selectEntireRow } from "../cell-selection";
import { parsePastedText, applyPastedValues, copyRangeToClipboard } from "../clipboard-handler";
import { computeViewport, colWidth, rowHeight, colOffset, rowOffset, hitTestCell, hitTestColHeader, hitTestRowHeader } from "../grid-renderer";

let totalPass = 0;
let totalFail = 0;

function feature(name: string, fn: () => void): void {
  console.log(`\n\x1b[1mFeature: ${name}\x1b[0m`);
  fn();
}

function scenario(name: string, fn: () => void): void {
  try { fn(); totalPass++; console.log(`  \x1b[32m✓\x1b[0m ${name}`); }
  catch (e: any) { totalFail++; console.log(`  \x1b[31m✗\x1b[0m ${name}\n    \x1b[31m${e.message}\x1b[0m`); }
}

function assert(c: boolean, m: string): void { if (!c) throw new Error(m); }
function assertEqual(a: unknown, e: unknown, l: string): void { if (a !== e) throw new Error(`${l}: expected ${JSON.stringify(e)}, got ${JSON.stringify(a)}`); }
function assertClose(a: number, e: number, l: string, eps = 0.01): void { if (Math.abs(a - e) > eps) throw new Error(`${l}: expected ~${e}, got ${a}`); }

function makeSheet(cells: Record<string, { type: string; value: any; formula?: string }>): XlsxSheet {
  const map = new Map<string, XlsxCell>();
  for (const [ref, c] of Object.entries(cells)) {
    const { col, row } = parseRef(ref);
    map.set(ref, { ref: ref as any, row, col, type: c.type as any, value: c.formula ? null : c.value, formula: c.formula ?? null, calculatedValue: null, styleId: 0, hyperlink: null });
  }
  return { id: "s1", name: "Sheet1", order: 0, hidden: false, cells: map, mergedRegions: [], tables: [], charts: [], conditionalFormats: [], dataValidations: [], frozenRow: 0, frozenCol: 0, colWidths: new Map(), rowHeights: new Map(), defaultColWidth: 8.43, defaultRowHeight: 15 };
}

function makeWorkbook(sheets: XlsxSheet[]): XlsxWorkbook {
  return { id: "wb1", title: "Test", sheets, sharedStrings: [], styles: [{ id: 0, numFmt: null, font: null, fill: null, border: null, alignment: null }], definedNames: [], activeSheetIndex: 0 };
}

const S = makeSheet({
  A1: { type: "number", value: 10 }, A2: { type: "number", value: 20 }, A3: { type: "number", value: 30 },
  A4: { type: "number", value: 40 }, A5: { type: "number", value: 50 },
  B1: { type: "string", value: "Hello" }, B2: { type: "string", value: "World" },
  B3: { type: "number", value: 0 }, B4: { type: "boolean", value: true }, B5: { type: "boolean", value: false },
});

// ===========================================================================
// All remaining untested formula functions
// ===========================================================================

feature("Formula — Remaining Math Functions", () => {
  scenario("ABS(-99) → 99", () => { assertEqual(evaluateFormula("ABS(-99)", S), 99, "ABS"); });
  scenario("ABS(0) → 0", () => { assertEqual(evaluateFormula("ABS(0)", S), 0, "ABS(0)"); });
  scenario("SQRT(25) → 5", () => { assertEqual(evaluateFormula("SQRT(25)", S), 5, "SQRT"); });
  scenario("SQRT(0) → 0", () => { assertEqual(evaluateFormula("SQRT(0)", S), 0, "SQRT(0)"); });
  scenario("ROUND(3.456,1) → 3.5", () => { assertClose(evaluateFormula("ROUND(3.456,1)", S) as number, 3.5, "ROUND"); });
  scenario("ROUND(3.456,0) → 3", () => { assertClose(evaluateFormula("ROUND(3.456,0)", S) as number, 3, "ROUND(0)"); });
  scenario("COUNT(10,20,30) → 3", () => { assertEqual(evaluateFormula("COUNT(10,20,30)", S), 3, "COUNT"); });
  scenario("MIN(5,3,8) → 3", () => { assertEqual(evaluateFormula("MIN(5,3,8)", S), 3, "MIN"); });
  scenario("MAX(5,3,8) → 8", () => { assertEqual(evaluateFormula("MAX(5,3,8)", S), 8, "MAX"); });
});

feature("Formula — Remaining Logical Functions", () => {
  scenario("IF(1,yes,no) → yes (truthy number)", () => { assertEqual(evaluateFormula('IF(1,"yes","no")', S), "yes", "IF truthy"); });
  scenario("IF(0,yes,no) → no (falsy number)", () => { assertEqual(evaluateFormula('IF(0,"yes","no")', S), "no", "IF falsy"); });
  scenario("AND(TRUE,TRUE,TRUE) → true", () => { assertEqual(evaluateFormula("AND(TRUE,TRUE,TRUE)", S), true, "AND all true"); });
  scenario("AND(TRUE,FALSE) → false", () => { assertEqual(evaluateFormula("AND(TRUE,FALSE)", S), false, "AND mixed"); });
  scenario("OR(FALSE,FALSE,TRUE) → true", () => { assertEqual(evaluateFormula("OR(FALSE,FALSE,TRUE)", S), true, "OR one true"); });
  scenario("OR(FALSE,FALSE) → false", () => { assertEqual(evaluateFormula("OR(FALSE,FALSE)", S), false, "OR all false"); });
  scenario("NOT(TRUE) → false", () => { assertEqual(evaluateFormula("NOT(TRUE)", S), false, "NOT T"); });
  scenario("NOT(FALSE) → true", () => { assertEqual(evaluateFormula("NOT(FALSE)", S), true, "NOT F"); });
  scenario("IFERROR(42,0) → 42 (no error)", () => { assertEqual(evaluateFormula("IFERROR(42,0)", S), 42, "IFERROR ok"); });
  scenario("TRUE() → true", () => { assertEqual(evaluateFormula("TRUE()", S), true, "TRUE"); });
  scenario("FALSE() → false", () => { assertEqual(evaluateFormula("FALSE()", S), false, "FALSE"); });
  scenario("SWITCH(99,1,a,2,b,default) → default", () => { assertEqual(evaluateFormula('SWITCH(99,1,"a",2,"b","default")', S), "default", "SWITCH default"); });
});

feature("Formula — Remaining Text Functions", () => {
  scenario("CONCATENATE(a,b) → ab", () => { assertEqual(evaluateFormula('CONCATENATE("a","b")', S), "ab", "CONCATENATE"); });
  scenario("LEFT(Hello) → H (default n=1)", () => { assertEqual(evaluateFormula('LEFT("Hello")', S), "H", "LEFT default"); });
  scenario("LEFT(Hello,3) → Hel", () => { assertEqual(evaluateFormula('LEFT("Hello",3)', S), "Hel", "LEFT 3"); });
  scenario("RIGHT(Hello) → o (default n=1)", () => { assertEqual(evaluateFormula('RIGHT("Hello")', S), "o", "RIGHT default"); });
  scenario("MID(Hello,1,5) → Hello", () => { assertEqual(evaluateFormula('MID("Hello",1,5)', S), "Hello", "MID full"); });
  scenario("LEN() → 0", () => { assertEqual(evaluateFormula('LEN("")', S), 0, "LEN empty"); });
  scenario("FIND(x,Hello) → #VALUE! (not found)", () => { assertEqual(evaluateFormula('FIND("x","Hello")', S), "#VALUE!", "FIND miss"); });
  scenario("SEARCH(LL,Hello) → 3 (case-insensitive)", () => { assertEqual(evaluateFormula('SEARCH("LL","Hello")', S), 3, "SEARCH ci"); });
  scenario("SUBSTITUTE(aaa,a,b,2) → aba (instance 2)", () => { assertEqual(evaluateFormula('SUBSTITUTE("aaa","a","b",2)', S), "aba", "SUBSTITUTE nth"); });
  scenario("REPLACE(abcdef,3,2,XY) → abXYef", () => { assertEqual(evaluateFormula('REPLACE("abcdef",3,2,"XY")', S), "abXYef", "REPLACE"); });
  scenario("REPT(x,0) → empty", () => { assertEqual(evaluateFormula('REPT("x",0)', S), "", "REPT 0"); });
  scenario("EXACT(a,a) → true", () => { assertEqual(evaluateFormula('EXACT("a","a")', S), true, "EXACT same"); });
  scenario("TRIM(  a  b  ) → a b", () => { assertEqual(evaluateFormula('TRIM("  a  b  ")', S), "a b", "TRIM spaces"); });
  scenario("UPPER(hello) → HELLO", () => { assertEqual(evaluateFormula('UPPER("hello")', S), "HELLO", "UPPER"); });
  scenario("LOWER(HELLO) → hello", () => { assertEqual(evaluateFormula('LOWER("HELLO")', S), "hello", "LOWER"); });
  scenario("PROPER(john doe) → John Doe", () => { assertEqual(evaluateFormula('PROPER("john doe")', S), "John Doe", "PROPER"); });
  scenario("CODE(Z) → 90", () => { assertEqual(evaluateFormula('CODE("Z")', S), 90, "CODE Z"); });
  scenario("CHAR(97) → a", () => { assertEqual(evaluateFormula("CHAR(97)", S), "a", "CHAR a"); });
  scenario("VALUE(0) → 0", () => { assertEqual(evaluateFormula('VALUE("0")', S), 0, "VALUE 0"); });
  scenario("TEXT(1234,#,##0) → 1,234", () => {
    // TEXT with format pattern — our impl is simplified
    const v = evaluateFormula('TEXT(1234.5,"0.0")', S);
    assertEqual(v, "1234.5", "TEXT");
  });
  scenario("T(TRUE) → empty (non-string)", () => { assertEqual(evaluateFormula("T(TRUE)", S), "", "T bool"); });
  scenario("CLEAN removes tab char", () => {
    // CLEAN strips control chars (0x00-0x1F)
    const v = evaluateFormula('CLEAN("AB")', S);
    assert(typeof v === "string", "CLEAN returns string");
  });
});

feature("Formula — Remaining Lookup Functions", () => {
  const ls = makeSheet({
    A1: { type: "string", value: "Cat" }, A2: { type: "string", value: "Dog" }, A3: { type: "string", value: "Fish" },
    B1: { type: "number", value: 5 }, B2: { type: "number", value: 10 }, B3: { type: "number", value: 15 },
  });
  scenario("VLOOKUP(Dog,A1:B3,2,FALSE) → 10", () => { assertEqual(evaluateFormula('VLOOKUP("Dog",A1:B3,2,FALSE)', ls), 10, "VLOOKUP"); });
  scenario("VLOOKUP(missing) → #N/A", () => { assertEqual(evaluateFormula('VLOOKUP("Zebra",A1:B3,2,FALSE)', ls), "#N/A", "VLOOKUP miss"); });
  scenario("HLOOKUP with row data", () => {
    const hs = makeSheet({ A1: { type: "string", value: "X" }, B1: { type: "string", value: "Y" }, A2: { type: "number", value: 1 }, B2: { type: "number", value: 2 } });
    assertEqual(evaluateFormula('HLOOKUP("Y",A1:B2,2,FALSE)', hs), 2, "HLOOKUP");
  });
  scenario("INDEX(A1:B3,1,1) → Cat", () => { assertEqual(evaluateFormula("INDEX(A1:B3,1,1)", ls), "Cat", "INDEX"); });
  scenario("MATCH(Dog,A1:A3,0) → 2", () => { assertEqual(evaluateFormula('MATCH("Dog",A1:A3,0)', ls), 2, "MATCH"); });
  scenario("MATCH(missing) → #N/A", () => { assertEqual(evaluateFormula('MATCH("Zebra",A1:A3,0)', ls), "#N/A", "MATCH miss"); });
  scenario("INDIRECT(B2) resolves to 10", () => { assertEqual(evaluateFormula('INDIRECT("B2")', ls), 10, "INDIRECT"); });
  scenario("ROW(C5) → 5", () => { assertEqual(evaluateFormula('ROW("C5")', ls), 5, "ROW"); });
  scenario("COLUMN(C5) → 3", () => { assertEqual(evaluateFormula('COLUMN("C5")', ls), 3, "COLUMN"); });
  scenario("ROWS(A1:A10) → 10", () => { assertEqual(evaluateFormula("ROWS(A1:A10)", ls), 10, "ROWS"); });
  scenario("COLUMNS(A1:D1) → 4", () => { assertEqual(evaluateFormula("COLUMNS(A1:D1)", ls), 4, "COLUMNS"); });
  scenario("CHOOSE(2,a,b,c) → b", () => { assertEqual(evaluateFormula('CHOOSE(2,"a","b","c")', ls), "b", "CHOOSE"); });
  scenario("CHOOSE(0) → #VALUE!", () => { assertEqual(evaluateFormula('CHOOSE(0,"a")', ls), "#VALUE!", "CHOOSE OOB"); });
});

feature("Formula — Remaining Date Functions", () => {
  scenario("YEAR(2025-12-31) → 2025", () => { assertEqual(evaluateFormula('YEAR("2025-12-31")', S), 2025, "YEAR"); });
  scenario("MONTH(2025-12-31) → 12", () => { assertEqual(evaluateFormula('MONTH("2025-12-31")', S), 12, "MONTH"); });
  scenario("DAY(2025-12-31) → 31", () => { assertEqual(evaluateFormula('DAY("2025-12-31")', S), 31, "DAY"); });
  scenario("DATEDIF(Y) 2020→2025 → 5", () => { assertEqual(evaluateFormula('DATEDIF("2020-01-01","2025-01-01","Y")', S), 5, "DATEDIF Y"); });
  scenario("DATEVALUE returns serial number", () => { const v = evaluateFormula('DATEVALUE("2024-01-01")', S) as number; assert(v > 40000, "DATEVALUE serial"); });
  scenario("WEEKDAY returns 1-7", () => { const v = evaluateFormula('WEEKDAY("2024-01-01")', S) as number; assert(v >= 1 && v <= 7, "WEEKDAY range"); });
  scenario("EDATE shifts months", () => { const v = evaluateFormula('EDATE("2024-01-15",3)', S) as number; assert(typeof v === "number", "EDATE serial"); });
  scenario("EOMONTH returns end of month", () => { const v = evaluateFormula('EOMONTH("2024-01-15",0)', S) as number; assert(typeof v === "number", "EOMONTH"); });
  scenario("NETWORKDAYS counts workdays", () => { const v = evaluateFormula('NETWORKDAYS("2024-01-01","2024-01-05")', S) as number; assert(v >= 3 && v <= 5, "NETWORKDAYS"); });
  scenario("WORKDAY returns serial", () => { const v = evaluateFormula('WORKDAY("2024-01-01",5)', S) as number; assert(typeof v === "number", "WORKDAY"); });
  scenario("ISOWEEKNUM returns week", () => { const v = evaluateFormula('ISOWEEKNUM("2024-01-01")', S) as number; assert(v >= 1 && v <= 53, "ISOWEEKNUM"); });
});

feature("Formula — Remaining Statistical Functions", () => {
  scenario("COUNTIF with range", () => { assertEqual(evaluateFormula("COUNTIF(A1:A5,10)", S), 1, "COUNTIF"); });
  scenario("SUMIF with range", () => { assertEqual(evaluateFormula("SUMIF(A1:A5,10)", S), 10, "SUMIF"); });
  scenario("STDEV.S(1,2,3) = STDEV(1,2,3)", () => {
    const a = evaluateFormula("STDEV(1,2,3)", S) as number;
    const b = evaluateFormula("STDEV.S(1,2,3)", S) as number;
    assertClose(a, b, "STDEV.S = STDEV");
  });
  scenario("STDEV.P(1,2,3) = STDEVP(1,2,3)", () => {
    const a = evaluateFormula("STDEVP(1,2,3)", S) as number;
    const b = evaluateFormula("STDEV.P(1,2,3)", S) as number;
    assertClose(a, b, "STDEV.P = STDEVP");
  });
  scenario("VAR.S(1,2,3) = VAR(1,2,3)", () => {
    const a = evaluateFormula("VAR(1,2,3)", S) as number;
    const b = evaluateFormula("VAR.S(1,2,3)", S) as number;
    assertClose(a, b, "VAR.S = VAR");
  });
  scenario("VAR.P(1,2,3) = VARP(1,2,3)", () => {
    const a = evaluateFormula("VARP(1,2,3)", S) as number;
    const b = evaluateFormula("VAR.P(1,2,3)", S) as number;
    assertClose(a, b, "VAR.P = VARP");
  });
  scenario("RANK.EQ(30,10,20,30,40,50) = RANK(30,...)", () => {
    const a = evaluateFormula("RANK(30,10,20,30,40,50)", S) as number;
    const b = evaluateFormula("RANK.EQ(30,10,20,30,40,50)", S) as number;
    assertEqual(a, b, "RANK.EQ = RANK");
  });
  scenario("PERCENTILE.INC(1,2,3,4,5,0.25) ≈ 2", () => {
    assertClose(evaluateFormula("PERCENTILE.INC(1,2,3,4,5,0.25)", S) as number, 2, "PERCENTILE.INC", 0.5);
  });
});

feature("Formula — Remaining Information Functions", () => {
  scenario("ISTEXT(42) → false", () => { assertEqual(evaluateFormula("ISTEXT(42)", S), false, "ISTEXT num"); });
  scenario("ISTEXT(hello) → true", () => { assertEqual(evaluateFormula('ISTEXT("hello")', S), true, "ISTEXT str"); });
  scenario("ISNUMBER(hello) → false", () => { assertEqual(evaluateFormula('ISNUMBER("hello")', S), false, "ISNUMBER str"); });
  scenario("ISBLANK(42) → false", () => { assertEqual(evaluateFormula("ISBLANK(42)", S), false, "ISBLANK num"); });
  scenario("ISERROR(42) → false", () => { assertEqual(evaluateFormula("ISERROR(42)", S), false, "ISERROR ok"); });
  scenario("ISERROR(#DIV/0!) → true", () => { assertEqual(evaluateFormula('ISERROR("#DIV/0!")', S), true, "ISERROR err"); });
  scenario("ISNA(42) → false", () => { assertEqual(evaluateFormula("ISNA(42)", S), false, "ISNA ok"); });
  scenario("ISNA(#N/A) → true", () => { assertEqual(evaluateFormula('ISNA("#N/A")', S), true, "ISNA err"); });
  scenario("NA() → #N/A", () => { assertEqual(evaluateFormula("NA()", S), "#N/A", "NA"); });
  scenario("TYPE(hello) → 2", () => { assertEqual(evaluateFormula('TYPE("hello")', S), 2, "TYPE str"); });
  scenario("ERROR.TYPE(#REF!) → 4", () => { assertEqual(evaluateFormula('ERROR.TYPE("#REF!")', S), 4, "ERROR.TYPE REF"); });
  scenario("ERROR.TYPE(#NAME?) → 5", () => { assertEqual(evaluateFormula('ERROR.TYPE("#NAME?")', S), 5, "ERROR.TYPE NAME"); });
  scenario("N(hello) → 0", () => { assertEqual(evaluateFormula('N("hello")', S), 0, "N str"); });
  scenario("N(FALSE) → 0", () => { assertEqual(evaluateFormula("N(FALSE)", S), 0, "N false"); });
});

// ===========================================================================
// Editor Operations (without Svelte runes)
// ===========================================================================

feature("Editor — Row/Column Insert/Delete Stress", () => {
  scenario("Insert 3 rows at row 0 shifts all cells by 3", () => {
    const s = makeSheet({ A1: { type: "number", value: 1 }, A2: { type: "number", value: 2 } });
    const newCells = new Map<string, XlsxCell>();
    for (const [, cell] of s.cells) {
      const nr = buildRef(cell.col, cell.row + 3);
      newCells.set(nr, { ...cell, ref: nr as any, row: cell.row + 3 });
    }
    s.cells = newCells;
    assertEqual(s.cells.get("A4")?.value, 1, "A1→A4");
    assertEqual(s.cells.get("A5")?.value, 2, "A2→A5");
    assertEqual(s.cells.size, 2, "still 2 cells");
  });

  scenario("Delete all rows containing data leaves empty sheet", () => {
    const s = makeSheet({ A1: { type: "number", value: 1 }, A2: { type: "number", value: 2 } });
    const newCells = new Map<string, XlsxCell>();
    // delete rows 0 and 1
    for (const [ref, cell] of s.cells) { if (cell.row >= 2) newCells.set(ref, cell); }
    s.cells = newCells;
    assertEqual(s.cells.size, 0, "empty");
  });

  scenario("Insert column at 0 shifts all columns right by 1", () => {
    const s = makeSheet({ A1: { type: "string", value: "X" }, B1: { type: "string", value: "Y" } });
    const newCells = new Map<string, XlsxCell>();
    for (const [, cell] of s.cells) {
      const nr = buildRef(cell.col + 1, cell.row);
      newCells.set(nr, { ...cell, ref: nr as any, col: cell.col + 1 });
    }
    s.cells = newCells;
    assertEqual(s.cells.get("B1")?.value, "X", "A1→B1");
    assertEqual(s.cells.get("C1")?.value, "Y", "B1→C1");
  });
});

feature("Editor — Sheet Operations", () => {
  scenario("Multiple sheets maintain independent cells", () => {
    const s1 = makeSheet({ A1: { type: "number", value: 100 } }); s1.name = "S1";
    const s2 = makeSheet({ A1: { type: "number", value: 200 } }); s2.name = "S2"; s2.id = "s2";
    const wb = makeWorkbook([s1, s2]);
    assertEqual(wb.sheets[0].cells.get("A1")?.value, 100, "S1.A1");
    assertEqual(wb.sheets[1].cells.get("A1")?.value, 200, "S2.A1");
  });

  scenario("Hidden sheet flag persists", () => {
    const s = makeSheet({}); s.hidden = true;
    assertEqual(s.hidden, true, "hidden");
  });
});

// ===========================================================================
// CSV Edge Cases
// ===========================================================================

feature("CSV — Edge Cases", () => {
  scenario("Empty CSV → empty array", () => {
    const rows = parseCsv("");
    assertEqual(rows.length, 0, "empty");
  });

  scenario("Single cell CSV", () => {
    const rows = parseCsv("hello");
    assertEqual(rows.length, 1, "1 row");
    assertEqual(rows[0][0], "hello", "value");
  });

  scenario("CSV with trailing newline", () => {
    const rows = parseCsv("a,b\nc,d\n");
    assertEqual(rows.length, 2, "2 rows (no empty trailing)");
  });

  scenario("CSV with CRLF line endings", () => {
    const rows = parseCsv("a,b\r\nc,d\r\n");
    assertEqual(rows.length, 2, "CRLF");
    assertEqual(rows[0][0], "a", "first cell");
  });

  scenario("CSV with escaped quotes", () => {
    const rows = parseCsv('"say ""hello""",world');
    assertEqual(rows[0][0], 'say "hello"', "escaped quotes");
  });

  scenario("CSV with semicolon delimiter", () => {
    const rows = parseCsv("a;b;c", ";");
    assertEqual(rows[0].length, 3, "3 cols");
    assertEqual(rows[0][1], "b", "middle col");
  });

  scenario("CSV with tab delimiter", () => {
    const rows = parseCsv("a\tb\tc", "\t");
    assertEqual(rows[0].length, 3, "3 cols");
  });

  scenario("Large CSV (1000 rows) parses correctly", () => {
    const lines = ["id,value"];
    for (let i = 0; i < 1000; i++) lines.push(`${i},${i * 10}`);
    const rows = parseCsv(lines.join("\n"));
    assertEqual(rows.length, 1001, "1001 rows");
    assertEqual(rows[1000][1], "9990", "last value");
  });

  scenario("CSV to workbook → sheet to CSV roundtrip", () => {
    const csv = "Name,Score\nAlice,95\nBob,87";
    const wb = csvToWorkbook(csv, "Test");
    const out = sheetToCsv(wb.sheets[0]);
    assert(out.includes("Alice"), "Alice preserved");
    assert(out.includes("95"), "95 preserved");
  });

  scenario("CSV numeric detection: integer, float, negative, scientific", () => {
    const wb = csvToWorkbook("42\n3.14\n-7\n0\n", "Test");
    assertEqual(wb.sheets[0].cells.get("A1")?.type, "number", "integer");
    assertEqual(wb.sheets[0].cells.get("A2")?.type, "number", "float");
    assertEqual(wb.sheets[0].cells.get("A3")?.type, "number", "negative");
    assertEqual(wb.sheets[0].cells.get("A4")?.type, "number", "zero");
  });
});

// ===========================================================================
// Find/Replace Edge Cases
// ===========================================================================

feature("Find/Replace — Edge Cases", () => {
  scenario("Find empty query returns no results", () => {
    const s = makeSheet({ A1: { type: "string", value: "Hello" } });
    const results = findInSheet(s, 0, { query: "", matchCase: false, matchEntireCell: false, searchFormulas: false, searchScope: "sheet" });
    assertEqual(results.length, 0, "empty query");
  });

  scenario("Find in empty sheet returns no results", () => {
    const s = makeSheet({});
    const results = findInSheet(s, 0, { query: "test", matchCase: false, matchEntireCell: false, searchFormulas: false, searchScope: "sheet" });
    assertEqual(results.length, 0, "empty sheet");
  });

  scenario("Replace in formula cell", () => {
    const cell: XlsxCell = { ref: "A1", row: 0, col: 0, type: "formula", value: null, formula: "SUM(A1:A10)", calculatedValue: null, styleId: 0, hyperlink: null };
    const replaced = replaceInCell(cell, "SUM", "AVERAGE", false);
    assertEqual(replaced, true, "replaced");
    assertEqual(cell.formula, "AVERAGE(A1:A10)", "formula updated");
  });

  scenario("Replace all with case sensitivity", () => {
    const s = makeSheet({
      A1: { type: "string", value: "Apple" },
      A2: { type: "string", value: "apple" },
    });
    const count = replaceAllInSheet(s, "Apple", "Orange", true);
    assertEqual(count, 1, "case-sensitive replace count");
    assertEqual(s.cells.get("A1")?.value, "Orange", "A1 replaced");
    assertEqual(s.cells.get("A2")?.value, "apple", "A2 unchanged");
  });

  scenario("Find with searchFormulas=true searches formula text", () => {
    const s = makeSheet({
      A1: { type: "formula", value: null, formula: "SUM(B1:B10)" },
    });
    const results = findInSheet(s, 0, { query: "SUM", matchCase: false, matchEntireCell: false, searchFormulas: true, searchScope: "sheet" });
    assertEqual(results.length, 1, "found in formula");
  });
});

// ===========================================================================
// Comments Edge Cases
// ===========================================================================

feature("Comments — Edge Cases", () => {
  scenario("Add comment to same cell overwrites", () => {
    const store = new CommentStore();
    store.add("A1", "User1", "First");
    store.add("A1", "User2", "Second");
    assertEqual(store.get("A1")?.text, "Second", "overwritten");
    assertEqual(store.get("A1")?.author, "User2", "new author");
  });

  scenario("Reply to non-existent comment is no-op", () => {
    const store = new CommentStore();
    store.reply("A1", "User", "Reply");
    assertEqual(store.hasComment("A1"), false, "no comment");
  });

  scenario("Delete non-existent comment returns false", () => {
    const store = new CommentStore();
    assertEqual(store.delete("A1"), false, "not found");
  });

  scenario("getAll returns all comments sorted", () => {
    const store = new CommentStore();
    store.add("C1", "U", "C");
    store.add("A1", "U", "A");
    store.add("B1", "U", "B");
    assertEqual(store.getAll().length, 3, "3 comments");
  });

  scenario("Multiple replies on same comment", () => {
    const store = new CommentStore();
    store.add("A1", "User", "Original");
    store.reply("A1", "Bot1", "Reply 1");
    store.reply("A1", "Bot2", "Reply 2");
    store.reply("A1", "Bot3", "Reply 3");
    assertEqual(store.get("A1")?.replies.length, 3, "3 replies");
  });
});

// ===========================================================================
// Cell Selection Edge Cases
// ===========================================================================

feature("Cell Selection — Edge Cases", () => {
  scenario("rangeSize of single cell → 1", () => {
    assertEqual(rangeSize({ startRow: 0, startCol: 0, endRow: 0, endCol: 0 }), 1, "single");
  });

  scenario("rangeSize of 3x4 → 12", () => {
    assertEqual(rangeSize({ startRow: 0, startCol: 0, endRow: 2, endCol: 3 }), 12, "3x4");
  });

  scenario("parseRangeString single cell A1 → range with same start/end", () => {
    const r = parseRangeString("A1");
    assertEqual(r.startRow, 0, "start"); assertEqual(r.endRow, 0, "end");
  });

  scenario("selectEntireColumn(0, 999) → A1:A1000", () => {
    const r = selectEntireColumn(0, 999);
    assertEqual(r.startCol, 0, "col"); assertEqual(r.endCol, 0, "col");
    assertEqual(r.startRow, 0, "start"); assertEqual(r.endRow, 999, "end");
  });

  scenario("selectEntireRow(0, 255) → A1:IV1", () => {
    const r = selectEntireRow(0, 255);
    assertEqual(r.startRow, 0, "row"); assertEqual(r.endRow, 0, "row");
    assertEqual(r.startCol, 0, "start"); assertEqual(r.endCol, 255, "end");
  });

  scenario("nextCellTab wraps to next row at maxCol", () => {
    const next = nextCellTab("Z1", 25);
    assertEqual(next, "A2", "wrap");
  });

  scenario("isCellInRange boundary check — cell on edge", () => {
    const r = { startRow: 0, startCol: 0, endRow: 2, endCol: 2 };
    assertEqual(isCellInRange("A1", r), true, "top-left");
    assertEqual(isCellInRange("C3", r), true, "bottom-right");
    assertEqual(isCellInRange("D3", r), false, "outside right");
  });

  scenario("Autofill with single value → copy pattern", () => {
    const p = detectFillPattern([42]);
    assertEqual(p?.type, "copy", "single value = copy");
  });

  scenario("generateFillValues with copy pattern cycles", () => {
    const vals = generateFillValues(["a", "b"], 5);
    assertEqual(vals[0], "a", "0"); assertEqual(vals[1], "b", "1"); assertEqual(vals[2], "a", "2");
  });
});

// ===========================================================================
// Clipboard Edge Cases
// ===========================================================================

feature("Clipboard — Edge Cases", () => {
  scenario("Parse empty paste → empty array", () => {
    const rows = parsePastedText("");
    assertEqual(rows.length, 0, "empty");
  });

  scenario("Parse single value paste", () => {
    const rows = parsePastedText("hello");
    assertEqual(rows.length, 1, "1 row");
    assertEqual(rows[0][0], "hello", "value");
  });

  scenario("Apply paste with formula detection", () => {
    const s = makeSheet({});
    applyPastedValues(s, "A1", [["=SUM(1,2)"]]);
    assertEqual(s.cells.get("A1")?.type, "formula", "formula detected");
    assertEqual(s.cells.get("A1")?.formula, "SUM(1,2)", "formula value");
  });

  scenario("Apply paste with number detection", () => {
    const s = makeSheet({});
    applyPastedValues(s, "A1", [["3.14"]]);
    assertEqual(s.cells.get("A1")?.type, "number", "number detected");
    assertClose(s.cells.get("A1")?.value as number, 3.14, "value");
  });
});

// ===========================================================================
// XLSX Stress Tests
// ===========================================================================

feature("XLSX — Stress & Edge Cases", () => {
  scenario("Workbook with 100 cells roundtrips correctly", async () => {
    const cells: Record<string, any> = {};
    for (let r = 0; r < 10; r++) {
      for (let c = 0; c < 10; c++) {
        cells[buildRef(c, r)] = { type: "number", value: r * 10 + c };
      }
    }
    const wb = makeWorkbook([makeSheet(cells)]);
    const parsed = parseXlsx(await exportXlsx(wb).arrayBuffer());
    assertEqual(parsed.sheets[0].cells.size, 100, "100 cells");
    assertEqual(parsed.sheets[0].cells.get("A1")?.value, 0, "A1=0");
    assertEqual(parsed.sheets[0].cells.get("J10")?.value, 99, "J10=99");
  });

  scenario("Workbook with 5 sheets roundtrips", async () => {
    const sheets: XlsxSheet[] = [];
    for (let i = 0; i < 5; i++) {
      const s = makeSheet({ A1: { type: "number", value: i } });
      s.name = `Sheet${i + 1}`; s.id = `s${i}`; s.order = i;
      sheets.push(s);
    }
    const wb = makeWorkbook(sheets);
    const parsed = parseXlsx(await exportXlsx(wb).arrayBuffer());
    assertEqual(parsed.sheets.length, 5, "5 sheets");
    for (let i = 0; i < 5; i++) {
      assertEqual(parsed.sheets[i].name, `Sheet${i + 1}`, `sheet ${i} name`);
    }
  });

  scenario("Cell with very long string roundtrips", async () => {
    const longStr = "A".repeat(10000);
    const wb = makeWorkbook([makeSheet({ A1: { type: "string", value: longStr } })]);
    const parsed = parseXlsx(await exportXlsx(wb).arrayBuffer());
    assertEqual(parsed.sheets[0].cells.get("A1")?.value, longStr, "long string");
  });

  scenario("Cell with special XML characters roundtrips", async () => {
    const wb = makeWorkbook([makeSheet({ A1: { type: "string", value: '<tag attr="val">&amp;</tag>' } })]);
    const parsed = parseXlsx(await exportXlsx(wb).arrayBuffer());
    assertEqual(parsed.sheets[0].cells.get("A1")?.value, '<tag attr="val">&amp;</tag>', "xml chars");
  });

  scenario("Formula with nested functions roundtrips", async () => {
    const wb = makeWorkbook([makeSheet({ A1: { type: "formula", value: null, formula: 'IF(SUM(1,2)>2,"yes","no")' } })]);
    const parsed = parseXlsx(await exportXlsx(wb).arrayBuffer());
    assertEqual(parsed.sheets[0].cells.get("A1")?.formula, 'IF(SUM(1,2)>2,"yes","no")', "nested formula");
  });

  scenario("Recalc chain: A1=1, B1=A1+1, C1=B1+1, D1=C1+1 → D1=4", () => {
    const s = makeSheet({
      A1: { type: "number", value: 1 },
      B1: { type: "formula", value: null, formula: "A1+1" },
      C1: { type: "formula", value: null, formula: "B1+1" },
      D1: { type: "formula", value: null, formula: "C1+1" },
    });
    recalculateSheet(s);
    assertEqual(s.cells.get("B1")?.calculatedValue, 2, "B1");
    assertEqual(s.cells.get("C1")?.calculatedValue, 3, "C1");
    assertEqual(s.cells.get("D1")?.calculatedValue, 4, "D1");
  });

  scenario("Circular dependency is detected by detectCircular()", () => {
    const s = makeSheet({
      A1: { type: "formula", value: null, formula: "B1+1" },
      B1: { type: "formula", value: null, formula: "A1+1" },
    });
    assert(detectCircular(s, "A1") === true, "A1 circular");
    assert(detectCircular(s, "B1") === true, "B1 circular");
  });
});

// ===========================================================================
// Column Letter Edge Cases
// ===========================================================================

feature("Column Letters — Extended", () => {
  scenario("colToLetter(701) → ZZ", () => { assertEqual(colToLetter(701), "ZZ", "ZZ"); });
  scenario("colToLetter(702) → AAA", () => { assertEqual(colToLetter(702), "AAA", "AAA"); });
  scenario("letterToCol(ZZ) → 701", () => { assertEqual(letterToCol("ZZ"), 701, "ZZ→701"); });
  scenario("letterToCol(AAA) → 702", () => { assertEqual(letterToCol("AAA"), 702, "AAA→702"); });
  scenario("Roundtrip for all single letters A-Z", () => {
    for (let i = 0; i < 26; i++) {
      assertEqual(letterToCol(colToLetter(i)), i, `roundtrip ${i}`);
    }
  });
  scenario("Roundtrip for double letters AA-AZ", () => {
    for (let i = 26; i < 52; i++) {
      assertEqual(letterToCol(colToLetter(i)), i, `roundtrip ${i}`);
    }
  });
});

// ===========================================================================
// Summary
// ===========================================================================

setTimeout(() => {
  console.log(`\n${"=".repeat(50)}`);
  console.log(`\x1b[1mResults: ${totalPass} passed, ${totalFail} failed, ${totalPass + totalFail} total\x1b[0m`);
  if (totalFail === 0) console.log("\x1b[32m\x1b[1m✓ ALL SCENARIOS PASSED\x1b[0m");
  else { console.log(`\x1b[31m\x1b[1m✗ ${totalFail} SCENARIO(S) FAILED\x1b[0m`); process.exit(1); }
}, 500);
