/**
 * XLSX Exporter — rebuild SpreadsheetML XML from workbook graph, ZIP, and download.
 *
 * Generates a valid .xlsx (Office Open XML) from the in-memory XlsxWorkbook.
 * Uses fflate for ZIP compression.
 */

import { zipSync, strToU8 } from "fflate";
import type { XlsxWorkbook, XlsxSheet, XlsxCell, XlsxStyle, XlsxSharedString, CellRef } from "./ooxml-parser";
import { colToLetter } from "./ooxml-parser";

/** Export an XlsxWorkbook to a .xlsx Blob for download. */
export function exportXlsx(wb: XlsxWorkbook): Blob {
  const files: Record<string, Uint8Array> = {};

  // Build shared string table from all cells
  const sstMap = new Map<string, number>();
  const sstList: string[] = [];
  for (const sheet of wb.sheets) {
    for (const cell of sheet.cells.values()) {
      if (cell.type === "string" && typeof cell.value === "string" && !sstMap.has(cell.value)) {
        sstMap.set(cell.value, sstList.length);
        sstList.push(cell.value);
      }
    }
  }

  files["[Content_Types].xml"] = strToU8(buildContentTypes(wb));
  files["_rels/.rels"] = strToU8(buildRootRels());
  files["xl/workbook.xml"] = strToU8(buildWorkbookXml(wb));
  files["xl/_rels/workbook.xml.rels"] = strToU8(buildWorkbookRels(wb));
  files["xl/sharedStrings.xml"] = strToU8(buildSharedStringsXml(sstList));
  files["xl/styles.xml"] = strToU8(buildStylesXml(wb.styles));
  files["xl/theme/theme1.xml"] = strToU8(buildThemeXml());
  files["docProps/app.xml"] = strToU8(buildAppXml());
  files["docProps/core.xml"] = strToU8(buildCoreXml(wb.title));

  for (let i = 0; i < wb.sheets.length; i++) {
    const sheet = wb.sheets[i];
    files[`xl/worksheets/sheet${i + 1}.xml`] = strToU8(buildSheetXml(sheet, sstMap));
  }

  const zipped = zipSync(files, { level: 6 });
  const exactBuffer = (zipped.buffer as ArrayBuffer).slice(zipped.byteOffset, zipped.byteOffset + zipped.byteLength);
  return new Blob([exactBuffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

/** Trigger browser download of a Blob. */
export function downloadBlob(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// ---------------------------------------------------------------------------
// XML Builders
// ---------------------------------------------------------------------------

function esc(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function buildContentTypes(wb: XlsxWorkbook): string {
  let sheets = "";
  for (let i = 0; i < wb.sheets.length; i++) {
    sheets += `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  ${sheets}
</Types>`;
}

function buildRootRels(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
</Relationships>`;
}

function buildWorkbookXml(wb: XlsxWorkbook): string {
  let sheets = "";
  for (let i = 0; i < wb.sheets.length; i++) {
    const s = wb.sheets[i];
    const state = s.hidden ? ` state="hidden"` : "";
    sheets += `<sheet name="${esc(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"${state}/>`;
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>${sheets}</sheets>
</workbook>`;
}

function buildWorkbookRels(wb: XlsxWorkbook): string {
  let rels = "";
  for (let i = 0; i < wb.sheets.length; i++) {
    rels += `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`;
  }
  const nextId = wb.sheets.length + 1;
  rels += `<Relationship Id="rId${nextId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>`;
  rels += `<Relationship Id="rId${nextId + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`;
  rels += `<Relationship Id="rId${nextId + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`;
}

function buildSharedStringsXml(strings: string[]): string {
  let items = "";
  for (const s of strings) {
    items += `<si><t>${esc(s)}</t></si>`;
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${strings.length}" uniqueCount="${strings.length}">${items}</sst>`;
}

function buildStylesXml(styles: XlsxStyle[]): string {
  // Minimal styles.xml — enough for Excel to open without complaint
  let fonts = '<font><sz val="11"/><name val="Calibri"/></font>';
  let fills = '<fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>';
  let borders = '<border><left/><right/><top/><bottom/><diagonal/></border>';
  let xfs = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>';

  for (let i = 1; i < styles.length; i++) {
    const st = styles[i];
    if (st.font) {
      fonts += `<font>`;
      if (st.font.bold) fonts += `<b/>`;
      if (st.font.italic) fonts += `<i/>`;
      if (st.font.underline) fonts += `<u/>`;
      fonts += `<sz val="${st.font.size}"/>`;
      fonts += `<name val="${esc(st.font.name)}"/>`;
      fonts += `</font>`;
    }
    xfs += `<xf numFmtId="0" fontId="${st.font ? i : 0}" fillId="0" borderId="0"/>`;
  }

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="${styles.length}">${fonts}</fonts>
  <fills count="2">${fills}</fills>
  <borders count="1">${borders}</borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="${styles.length}">${xfs}</cellXfs>
</styleSheet>`;
}

function buildSheetXml(sheet: XlsxSheet, sstMap: Map<string, number>): string {
  // Collect cells sorted by row then col
  const sorted = [...sheet.cells.values()].sort((a, b) => a.row !== b.row ? a.row - b.row : a.col - b.col);

  // Group by row
  const rowMap = new Map<number, XlsxCell[]>();
  for (const cell of sorted) {
    if (!rowMap.has(cell.row)) rowMap.set(cell.row, []);
    rowMap.get(cell.row)!.push(cell);
  }

  let sheetData = "";
  const rowNums = [...rowMap.keys()].sort((a, b) => a - b);
  for (const rowNum of rowNums) {
    const cells = rowMap.get(rowNum)!;
    const ht = sheet.rowHeights.get(rowNum);
    const htAttr = ht ? ` ht="${ht}" customHeight="1"` : "";
    let rowXml = `<row r="${rowNum + 1}"${htAttr}>`;
    for (const cell of cells) {
      rowXml += buildCellXml(cell, sstMap);
    }
    rowXml += `</row>`;
    sheetData += rowXml;
  }

  // Merged cells
  let merges = "";
  if (sheet.mergedRegions.length > 0) {
    merges = `<mergeCells count="${sheet.mergedRegions.length}">`;
    for (const m of sheet.mergedRegions) {
      merges += `<mergeCell ref="${m.startRef}:${m.endRef}"/>`;
    }
    merges += `</mergeCells>`;
  }

  // Column widths
  let cols = "";
  if (sheet.colWidths.size > 0) {
    cols = "<cols>";
    for (const [colIdx, width] of sheet.colWidths) {
      cols += `<col min="${colIdx + 1}" max="${colIdx + 1}" width="${width}" customWidth="1"/>`;
    }
    cols += "</cols>";
  }

  // Frozen panes
  let pane = "";
  if (sheet.frozenRow > 0 || sheet.frozenCol > 0) {
    const topLeft = `${colToLetter(sheet.frozenCol)}${sheet.frozenRow + 1}`;
    pane = `<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="${sheet.frozenRow}" xSplit="${sheet.frozenCol}" topLeftCell="${topLeft}" state="frozen"/></sheetView></sheetViews>`;
  }

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  ${pane}
  ${cols}
  <sheetData>${sheetData}</sheetData>
  ${merges}
</worksheet>`;
}

function buildCellXml(cell: XlsxCell, sstMap: Map<string, number>): string {
  const sAttr = cell.styleId > 0 ? ` s="${cell.styleId}"` : "";

  if (cell.formula) {
    const cv = cell.calculatedValue != null ? `<v>${esc(String(cell.calculatedValue))}</v>` : "";
    return `<c r="${cell.ref}"${sAttr}><f>${esc(cell.formula)}</f>${cv}</c>`;
  }

  if (cell.type === "string" && typeof cell.value === "string") {
    const idx = sstMap.get(cell.value);
    if (idx !== undefined) {
      return `<c r="${cell.ref}" t="s"${sAttr}><v>${idx}</v></c>`;
    }
    return `<c r="${cell.ref}" t="inlineStr"${sAttr}><is><t>${esc(cell.value)}</t></is></c>`;
  }

  if (cell.type === "number" && typeof cell.value === "number") {
    return `<c r="${cell.ref}"${sAttr}><v>${cell.value}</v></c>`;
  }

  if (cell.type === "boolean") {
    return `<c r="${cell.ref}" t="b"${sAttr}><v>${cell.value ? "1" : "0"}</v></c>`;
  }

  if (cell.type === "error" && typeof cell.value === "string") {
    return `<c r="${cell.ref}" t="e"${sAttr}><v>${esc(cell.value)}</v></c>`;
  }

  return `<c r="${cell.ref}"${sAttr}/>`;
}

function buildThemeXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme>
    <a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst></a:fmtScheme>
  </a:themeElements>
</a:theme>`;
}

function buildAppXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>xlsx.etzhayyim.com</Application>
</Properties>`;
}

function buildCoreXml(title: string): string {
  const now = new Date().toISOString();
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>${esc(title)}</dc:title>
  <dc:creator>xlsx.etzhayyim.com</dc:creator>
  <dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>
</cp:coreProperties>`;
}
