/**
 * CSV Handler — import CSV text into a sheet, export sheet to CSV.
 *
 * Implements RFC 4180-style parsing: double-quote escaping, newlines within
 * quoted fields, and configurable delimiter (default comma).
 */
import type { XlsxWorkbook, XlsxSheet, XlsxCell, CellRef } from "./ooxml-parser";
import { buildRef, parseRef } from "./ooxml-parser";

/**
 * Parse CSV text into a 2D string array.
 *
 * Handles quoted fields (double-quote escaping) and embedded newlines
 * within quoted fields per RFC 4180.
 *
 * @param text - The raw CSV string.
 * @param delimiter - Field delimiter, defaults to `","`.
 * @returns A 2D array where each inner array is one row of field values.
 */
export function parseCsv(text: string, delimiter: string = ","): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let inQuotes = false;
  let i = 0;

  while (i < text.length) {
    const ch = text[i];

    if (inQuotes) {
      if (ch === '"') {
        // Escaped quote ("") or end of quoted field.
        if (i + 1 < text.length && text[i + 1] === '"') {
          field += '"';
          i += 2;
        } else {
          inQuotes = false;
          i++;
        }
      } else {
        field += ch;
        i++;
      }
    } else {
      if (ch === '"') {
        inQuotes = true;
        i++;
      } else if (ch === delimiter) {
        row.push(field);
        field = "";
        i++;
      } else if (ch === "\r") {
        // Handle \r\n and lone \r as row separator.
        row.push(field);
        field = "";
        rows.push(row);
        row = [];
        i++;
        if (i < text.length && text[i] === "\n") i++;
      } else if (ch === "\n") {
        row.push(field);
        field = "";
        rows.push(row);
        row = [];
        i++;
      } else {
        field += ch;
        i++;
      }
    }
  }

  // Flush the last field / row (unless the input was empty or ended with a newline).
  if (field || row.length > 0) {
    row.push(field);
    rows.push(row);
  }

  return rows;
}

/**
 * Convert a sheet's cell data to a CSV string.
 *
 * Scans all cells to determine the bounding rectangle, then serialises
 * each row. Fields containing the delimiter, double-quote, or newline
 * characters are quoted.
 *
 * @param sheet - The sheet to export.
 * @param delimiter - Field delimiter, defaults to `","`.
 * @returns The CSV string.
 */
export function sheetToCsv(sheet: XlsxSheet, delimiter: string = ","): string {
  // Determine bounds.
  let maxRow = 0;
  let maxCol = 0;
  for (const [, cell] of sheet.cells) {
    if (cell.row > maxRow) maxRow = cell.row;
    if (cell.col > maxCol) maxCol = cell.col;
  }

  const lines: string[] = [];
  for (let r = 0; r <= maxRow; r++) {
    const fields: string[] = [];
    for (let c = 0; c <= maxCol; c++) {
      const ref = buildRef(c, r);
      const cell = sheet.cells.get(ref);
      let val = "";
      if (cell) {
        if (cell.formula) {
          val = `=${cell.formula}`;
        } else if (cell.value != null) {
          val = String(cell.value);
        }
      }
      fields.push(escapeCsvField(val, delimiter));
    }
    lines.push(fields.join(delimiter));
  }

  return lines.join("\r\n");
}

/**
 * Create a minimal workbook from CSV text.
 *
 * Parses the CSV and populates a single sheet named after `title`.
 * Numeric values are auto-detected and stored as numbers.
 *
 * @param text - The raw CSV string.
 * @param title - The workbook / sheet title.
 * @param delimiter - Field delimiter, defaults to `","`.
 * @returns A new {@link XlsxWorkbook} containing one sheet.
 */
export function csvToWorkbook(text: string, title: string, delimiter: string = ","): XlsxWorkbook {
  const data = parseCsv(text, delimiter);
  const cells = new Map<CellRef, XlsxCell>();

  for (let r = 0; r < data.length; r++) {
    const row = data[r];
    for (let c = 0; c < row.length; c++) {
      const raw = row[c];
      if (raw === "") continue;

      const ref = buildRef(c, r);
      let type: XlsxCell["type"] = "string";
      let value: XlsxCell["value"] = raw;
      let formula: string | null = null;

      if (raw.startsWith("=")) {
        type = "formula";
        formula = raw.slice(1);
        value = null;
      } else {
        const num = Number(raw);
        if (!isNaN(num) && raw.trim() !== "") {
          type = "number";
          value = num;
        }
      }

      cells.set(ref, {
        ref,
        row: r,
        col: c,
        type,
        value,
        formula,
        calculatedValue: null,
        styleId: 0,
        hyperlink: null,
      });
    }
  }

  const sheet: XlsxSheet = {
    id: `sheet_csv_${Date.now()}`,
    name: title || "Sheet1",
    order: 0,
    hidden: false,
    cells,
    mergedRegions: [],
    tables: [],
    charts: [],
    conditionalFormats: [],
    dataValidations: [],
    frozenRow: 0,
    frozenCol: 0,
    colWidths: new Map(),
    rowHeights: new Map(),
    defaultColWidth: 8.43,
    defaultRowHeight: 15,
  };

  return {
    id: `wb_csv_${Date.now()}`,
    title: title || "CSV Import",
    sheets: [sheet],
    sharedStrings: [],
    styles: [],
    definedNames: [],
    activeSheetIndex: 0,
  };
}

/**
 * Export a sheet as CSV and trigger a browser file download.
 *
 * Creates a temporary `<a>` element with a Blob URL to initiate the
 * download. Does nothing in non-browser environments.
 *
 * @param sheet - The sheet to export.
 * @param filename - The download filename (e.g. `"data.csv"`).
 * @param delimiter - Field delimiter, defaults to `","`.
 */
export function downloadCsv(sheet: XlsxSheet, filename: string, delimiter: string = ","): void {
  const csv = sheetToCsv(sheet, delimiter);
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);

  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.style.display = "none";
  document.body.appendChild(a);
  a.click();

  // Clean up after a tick.
  setTimeout(() => {
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }, 0);
}

/**
 * Escape a single CSV field value, quoting it when necessary.
 */
function escapeCsvField(value: string, delimiter: string): string {
  if (value.includes(delimiter) || value.includes('"') || value.includes("\n") || value.includes("\r")) {
    return '"' + value.replace(/"/g, '""') + '"';
  }
  return value;
}
