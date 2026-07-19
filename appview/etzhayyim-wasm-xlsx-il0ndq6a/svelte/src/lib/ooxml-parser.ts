/**
 * OOXML (.xlsx) Parser — decompress ZIP, parse SpreadsheetML XML into typed workbook graph.
 *
 * XLSX is a ZIP containing:
 *   [Content_Types].xml
 *   _rels/.rels
 *   xl/workbook.xml           — sheet list, defined names
 *   xl/sharedStrings.xml      — shared string table (SST)
 *   xl/styles.xml             — number formats, fonts, fills, borders, cell xfs
 *   xl/worksheets/sheet{N}.xml — cell data, merges, hyperlinks
 *   xl/charts/chart{N}.xml    — chart definitions
 *   xl/tables/table{N}.xml    — structured table definitions
 *   xl/theme/theme1.xml       — colour theme
 *   xl/_rels/workbook.xml.rels — sheet→file mapping
 *   docProps/app.xml, core.xml
 *
 * We parse into a flat graph model (XlsxWorkbook → XlsxSheet → XlsxCell)
 * suitable for kagami SQL persistence and Canvas 2D grid rendering.
 */

import { unzipSync, strFromU8 } from "fflate";

// ---------------------------------------------------------------------------
// Types — mirror kagami graph labels
// ---------------------------------------------------------------------------

/** Cell reference string (e.g. "A1", "B2", "AA100"). */
export type CellRef = string;

export interface XlsxWorkbook {
  id: string;
  title: string;
  sheets: XlsxSheet[];
  sharedStrings: XlsxSharedString[];
  styles: XlsxStyle[];
  definedNames: XlsxDefinedName[];
  activeSheetIndex: number;
}

export interface XlsxSheet {
  id: string;
  name: string;
  order: number;
  hidden: boolean;
  /** Sparse cell storage keyed by A1 reference. */
  cells: Map<CellRef, XlsxCell>;
  mergedRegions: XlsxMergedRegion[];
  tables: XlsxTable[];
  charts: XlsxChart[];
  conditionalFormats: XlsxConditionalFormat[];
  dataValidations: XlsxDataValidation[];
  frozenRow: number;
  frozenCol: number;
  colWidths: Map<number, number>;
  rowHeights: Map<number, number>;
  defaultColWidth: number;
  defaultRowHeight: number;
}

export interface XlsxCell {
  ref: CellRef;
  row: number;
  col: number;
  type: "string" | "number" | "boolean" | "date" | "error" | "formula" | "empty";
  value: string | number | boolean | null;
  formula: string | null;
  calculatedValue: string | number | boolean | null;
  styleId: number;
  hyperlink: string | null;
}

export interface XlsxStyle {
  id: number;
  numFmt: string | null;
  font: XlsxFont | null;
  fill: XlsxFill | null;
  border: XlsxBorder | null;
  alignment: XlsxAlignment | null;
}

export interface XlsxFont {
  name: string;
  size: number;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strikethrough: boolean;
  color: string;
}

export interface XlsxFill {
  type: "solid" | "pattern" | "gradient";
  fgColor: string | null;
  bgColor: string | null;
}

export interface XlsxBorder {
  top: XlsxBorderEdge | null;
  bottom: XlsxBorderEdge | null;
  left: XlsxBorderEdge | null;
  right: XlsxBorderEdge | null;
}

export interface XlsxBorderEdge {
  style: "thin" | "medium" | "thick" | "dashed" | "dotted" | "double";
  color: string;
}

export interface XlsxAlignment {
  horizontal: "left" | "center" | "right" | "fill" | "justify";
  vertical: "top" | "center" | "bottom";
  wrapText: boolean;
  textRotation: number;
  indent: number;
}

export interface XlsxSharedString {
  index: number;
  text: string;
  richText: XlsxRichTextRun[] | null;
}

export interface XlsxRichTextRun {
  text: string;
  font: Partial<XlsxFont> | null;
}

export interface XlsxMergedRegion {
  startRef: CellRef;
  endRef: CellRef;
}

export interface XlsxDefinedName {
  name: string;
  scope: string | null;
  ref: string;
}

export interface XlsxTable {
  id: string;
  name: string;
  ref: string;
  headerRow: boolean;
  totalsRow: boolean;
  columns: { name: string; totalsFunction?: string }[];
}

export interface XlsxChart {
  id: string;
  type: "bar" | "column" | "line" | "pie" | "scatter" | "area" | "doughnut" | "radar";
  title: string | null;
  dataRange: string;
  anchor: { col: number; row: number; colOff: number; rowOff: number };
  width: number;
  height: number;
}

export interface XlsxConditionalFormat {
  type: "cellIs" | "colorScale" | "dataBar" | "iconSet" | "expression";
  priority: number;
  ref: string;
  operator?: string;
  formula?: string;
  style?: Partial<XlsxStyle>;
}

export interface XlsxDataValidation {
  type: "list" | "whole" | "decimal" | "date" | "textLength" | "custom";
  ref: string;
  formula1: string;
  formula2?: string;
  showDropdown?: boolean;
  errorTitle?: string;
  errorMessage?: string;
}

// ---------------------------------------------------------------------------
// Column ↔ Number helpers
// ---------------------------------------------------------------------------

/** Convert 0-based column index to letter(s): 0→A, 25→Z, 26→AA. */
export function colToLetter(col: number): string {
  let s = "";
  let c = col;
  while (c >= 0) {
    s = String.fromCharCode((c % 26) + 65) + s;
    c = Math.floor(c / 26) - 1;
  }
  return s;
}

/** Convert column letter(s) to 0-based index: A→0, Z→25, AA→26. */
export function letterToCol(letters: string): number {
  let col = 0;
  for (let i = 0; i < letters.length; i++) {
    col = col * 26 + (letters.charCodeAt(i) - 64);
  }
  return col - 1;
}

/** Parse a cell reference "A1" → { col: 0, row: 0 }. */
export function parseRef(ref: string): { col: number; row: number } {
  const match = ref.match(/^([A-Z]+)(\d+)$/);
  if (!match) return { col: 0, row: 0 };
  return { col: letterToCol(match[1]), row: parseInt(match[2], 10) - 1 };
}

/** Build a cell reference from col/row: (0, 0) → "A1". */
export function buildRef(col: number, row: number): CellRef {
  return `${colToLetter(col)}${row + 1}`;
}

// ---------------------------------------------------------------------------
// XML Helper (minimal, no external dep)
// ---------------------------------------------------------------------------

/** Unescape XML entities. */
function xmlUnescape(s: string): string {
  return s.replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'");
}

/** Extract text content of first matching tag (XML-unescaped). */
function xmlTag(xml: string, tag: string): string {
  const re = new RegExp(`<${tag}[^>]*>([\\s\\S]*?)</${tag}>`, "i");
  const m = xml.match(re);
  return m ? xmlUnescape(m[1]) : "";
}

/** Extract attribute value from a tag string. */
function xmlAttr(tagStr: string, attr: string): string {
  const re = new RegExp(`${attr}="([^"]*)"`, "i");
  const m = tagStr.match(re);
  return m ? m[1] : "";
}

/** Extract all occurrences of a tag (outer). */
function xmlAll(xml: string, tag: string): string[] {
  const re = new RegExp(`<${tag}(?:\\s[^>]*)?>(?:[\\s\\S]*?)</${tag}>|<${tag}(?:\\s[^>]*)?/>`, "gi");
  return [...xml.matchAll(re)].map((m) => m[0]);
}

/** Extract self-closing or opening tag string for attribute extraction. */
function xmlOpenTag(xml: string, tag: string): string {
  const re = new RegExp(`<${tag}(\\s[^>]*)?\\/?>`, "i");
  const m = xml.match(re);
  return m ? m[0] : "";
}

// ---------------------------------------------------------------------------
// Parser
// ---------------------------------------------------------------------------

/** Parse an XLSX ArrayBuffer into XlsxWorkbook. */
export function parseXlsx(buffer: ArrayBuffer): XlsxWorkbook {
  const files = unzipSync(new Uint8Array(buffer));
  const decode = (path: string): string => {
    const f = files[path];
    return f ? strFromU8(f) : "";
  };

  const wbId = `wb_${Date.now()}`;

  // --- Shared Strings ---
  const sharedStrings = parseSharedStrings(decode("xl/sharedStrings.xml"));

  // --- Styles ---
  const styles = parseStyles(decode("xl/styles.xml"));

  // --- Workbook (sheet list + defined names) ---
  const wbXml = decode("xl/workbook.xml");
  const wbRelsXml = decode("xl/_rels/workbook.xml.rels");

  const sheetEntries = parseWorkbookSheets(wbXml);
  const rels = parseRels(wbRelsXml);
  const definedNames = parseDefinedNames(wbXml);

  // --- Sheets ---
  const sheets: XlsxSheet[] = [];
  for (let i = 0; i < sheetEntries.length; i++) {
    const entry = sheetEntries[i];
    const relTarget = rels.get(entry.rId);
    const sheetPath = relTarget ? `xl/${relTarget}` : `xl/worksheets/sheet${i + 1}.xml`;
    const sheetXml = decode(sheetPath);
    sheets.push(parseSheet(sheetXml, entry, i, sharedStrings));
  }

  return {
    id: wbId,
    title: "Untitled",
    sheets,
    sharedStrings,
    styles,
    definedNames,
    activeSheetIndex: 0,
  };
}

// ---------------------------------------------------------------------------
// Sub-parsers
// ---------------------------------------------------------------------------

interface SheetEntry {
  name: string;
  sheetId: string;
  rId: string;
  hidden: boolean;
}

/** Parse workbook.xml → sheet entries. */
function parseWorkbookSheets(xml: string): SheetEntry[] {
  const entries: SheetEntry[] = [];
  const sheetTags = xmlAll(xml, "sheet");
  for (const tag of sheetTags) {
    entries.push({
      name: xmlAttr(tag, "name"),
      sheetId: xmlAttr(tag, "sheetId"),
      rId: xmlAttr(tag, "r:id"),
      hidden: xmlAttr(tag, "state") === "hidden",
    });
  }
  return entries;
}

/** Parse relationship file → Map<rId, target>. */
function parseRels(xml: string): Map<string, string> {
  const map = new Map<string, string>();
  const relTags = xmlAll(xml, "Relationship");
  for (const tag of relTags) {
    const id = xmlAttr(tag, "Id");
    const target = xmlAttr(tag, "Target");
    if (id && target) map.set(id, target);
  }
  return map;
}

/** Parse shared strings table. */
function parseSharedStrings(xml: string): XlsxSharedString[] {
  if (!xml) return [];
  const strings: XlsxSharedString[] = [];
  const siTags = xmlAll(xml, "si");
  for (let i = 0; i < siTags.length; i++) {
    const si = siTags[i];
    const tContent = xmlTag(si, "t");
    // Rich text: multiple <r> elements
    const rTags = xmlAll(si, "r");
    let richText: XlsxRichTextRun[] | null = null;
    let text = tContent;
    if (rTags.length > 0) {
      richText = [];
      const parts: string[] = [];
      for (const r of rTags) {
        const t = xmlTag(r, "t");
        parts.push(t);
        richText.push({ text: t, font: null });
      }
      text = parts.join("");
    }
    strings.push({ index: i, text, richText });
  }
  return strings;
}

/** Parse styles.xml → XlsxStyle[]. */
function parseStyles(xml: string): XlsxStyle[] {
  if (!xml) return [defaultStyle()];
  const styles: XlsxStyle[] = [];

  // Parse numFmts
  const numFmtMap = new Map<string, string>();
  const numFmtTags = xmlAll(xml, "numFmt");
  for (const tag of numFmtTags) {
    numFmtMap.set(xmlAttr(tag, "numFmtId"), xmlAttr(tag, "formatCode"));
  }

  // Parse fonts
  const fonts: XlsxFont[] = [];
  const fontTags = xmlAll(xml, "font");
  for (const tag of fontTags) {
    fonts.push({
      name: xmlAttr(xmlOpenTag(tag, "name"), "val") || "Calibri",
      size: parseFloat(xmlAttr(xmlOpenTag(tag, "sz"), "val")) || 11,
      bold: tag.includes("<b") && !tag.includes("<b/>"),
      italic: tag.includes("<i") && !tag.includes("<i/>"),
      underline: tag.includes("<u"),
      strikethrough: tag.includes("<strike"),
      color: xmlAttr(xmlOpenTag(tag, "color"), "rgb") || "#000000",
    });
  }

  // Parse fills
  const fills: XlsxFill[] = [];
  const fillTags = xmlAll(xml, "fill");
  for (const tag of fillTags) {
    const patternType = xmlAttr(xmlOpenTag(tag, "patternFill"), "patternType");
    fills.push({
      type: patternType === "solid" ? "solid" : "pattern",
      fgColor: xmlAttr(xmlOpenTag(tag, "fgColor"), "rgb") || null,
      bgColor: xmlAttr(xmlOpenTag(tag, "bgColor"), "rgb") || null,
    });
  }

  // Parse cellXfs (the actual cell style combinations)
  const xfTags = xmlAll(xmlTag(xml, "cellXfs"), "xf");
  for (let i = 0; i < xfTags.length; i++) {
    const xf = xfTags[i];
    const numFmtId = xmlAttr(xf, "numFmtId");
    const fontIdx = parseInt(xmlAttr(xf, "fontId")) || 0;
    const fillIdx = parseInt(xmlAttr(xf, "fillId")) || 0;

    styles.push({
      id: i,
      numFmt: numFmtMap.get(numFmtId) ?? null,
      font: fonts[fontIdx] ?? null,
      fill: fills[fillIdx] ?? null,
      border: null,
      alignment: null,
    });
  }

  if (styles.length === 0) styles.push(defaultStyle());
  return styles;
}

/** Parse defined names from workbook.xml. */
function parseDefinedNames(xml: string): XlsxDefinedName[] {
  const names: XlsxDefinedName[] = [];
  const dnTags = xmlAll(xml, "definedName");
  for (const tag of dnTags) {
    const name = xmlAttr(tag, "name");
    const scope = xmlAttr(tag, "localSheetId") || null;
    const ref = tag.replace(/<[^>]+>/g, "").trim();
    if (name) names.push({ name, scope, ref });
  }
  return names;
}

/** Parse a single worksheet XML into XlsxSheet. */
function parseSheet(xml: string, entry: SheetEntry, order: number, sharedStrings: XlsxSharedString[]): XlsxSheet {
  const cells = new Map<CellRef, XlsxCell>();
  const colWidths = new Map<number, number>();
  const rowHeights = new Map<number, number>();

  // Parse column widths
  const colTags = xmlAll(xml, "col");
  for (const tag of colTags) {
    const min = parseInt(xmlAttr(tag, "min")) - 1;
    const max = parseInt(xmlAttr(tag, "max")) - 1;
    const width = parseFloat(xmlAttr(tag, "width")) || 8.43;
    for (let c = min; c <= max; c++) colWidths.set(c, width);
  }

  // Parse rows and cells
  const rowTags = xmlAll(xml, "row");
  for (const rowTag of rowTags) {
    const rowNum = parseInt(xmlAttr(rowTag, "r")) - 1;
    const ht = parseFloat(xmlAttr(rowTag, "ht"));
    if (ht > 0) rowHeights.set(rowNum, ht);

    const cellTags = xmlAll(rowTag, "c");
    for (const cellTag of cellTags) {
      const ref = xmlAttr(cellTag, "r") as CellRef;
      const { col, row } = parseRef(ref);
      const cellType = xmlAttr(cellTag, "t");
      const styleId = parseInt(xmlAttr(cellTag, "s")) || 0;
      const vContent = xmlTag(cellTag, "v");
      const fContent = xmlTag(cellTag, "f");

      let type: XlsxCell["type"] = "empty";
      let value: XlsxCell["value"] = null;
      let formula: string | null = null;
      let calculatedValue: XlsxCell["calculatedValue"] = null;

      if (fContent) {
        type = "formula";
        formula = fContent;
        calculatedValue = vContent || null;
      } else if (cellType === "s" && vContent) {
        type = "string";
        const idx = parseInt(vContent);
        value = sharedStrings[idx]?.text ?? "";
      } else if (cellType === "b") {
        type = "boolean";
        value = vContent === "1";
      } else if (cellType === "e") {
        type = "error";
        value = vContent;
      } else if (vContent) {
        type = "number";
        value = parseFloat(vContent);
      }

      cells.set(ref, { ref, row, col, type, value, formula, calculatedValue, styleId, hyperlink: null });
    }
  }

  // Parse merged regions
  const mergedRegions: XlsxMergedRegion[] = [];
  const mergeTags = xmlAll(xml, "mergeCell");
  for (const tag of mergeTags) {
    const mergeRef = xmlAttr(tag, "ref");
    const [startRef, endRef] = mergeRef.split(":") as [CellRef, CellRef];
    if (startRef && endRef) mergedRegions.push({ startRef, endRef });
  }

  // Frozen panes
  let frozenRow = 0;
  let frozenCol = 0;
  const paneTag = xmlOpenTag(xml, "pane");
  if (paneTag && xmlAttr(paneTag, "state") === "frozen") {
    frozenRow = parseInt(xmlAttr(paneTag, "ySplit")) || 0;
    frozenCol = parseInt(xmlAttr(paneTag, "xSplit")) || 0;
  }

  return {
    id: `sheet_${entry.sheetId}_${Date.now()}`,
    name: entry.name,
    order,
    hidden: entry.hidden,
    cells,
    mergedRegions,
    tables: [],
    charts: [],
    conditionalFormats: [],
    dataValidations: [],
    frozenRow,
    frozenCol,
    colWidths,
    rowHeights,
    defaultColWidth: 8.43,
    defaultRowHeight: 15,
  };
}

/** Default style when no styles.xml is present. */
function defaultStyle(): XlsxStyle {
  return {
    id: 0,
    numFmt: null,
    font: { name: "Calibri", size: 11, bold: false, italic: false, underline: false, strikethrough: false, color: "#000000" },
    fill: { type: "pattern", fgColor: null, bgColor: null },
    border: null,
    alignment: null,
  };
}
