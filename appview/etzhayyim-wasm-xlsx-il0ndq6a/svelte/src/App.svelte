<script lang="ts">
  import { parseXlsx, type XlsxWorkbook, type XlsxSheet, type XlsxCell } from "./lib/ooxml-parser";
  import { exportXlsx, downloadBlob } from "./lib/xlsx-exporter";
  import { recalculateSheet } from "./lib/formula-engine";
  import {
    normalizeRange,
    rangeToString,
    nextCellTab,
    prevCellTab,
  } from "./lib/cell-selection";
  import {
    copyRangeToClipboard,
    parsePastedText,
    applyPastedValues,
  } from "./lib/clipboard-handler";
  import {
    editor,
    loadWorkbook,
    selectCell,
    selectRange,
    moveActiveCell,
    setCellValue,
    startEdit,
    commitEdit,
    cancelEdit,
    addSheet,
    deleteSheet,
    renameSheet,
    selectSheet,
    insertRows,
    deleteRows,
    insertColumns,
    deleteColumns,
    mergeCells,
    unmergeCells,
    copyCells,
    pasteCells,
    pushUndo,
    undo,
    redo,
    canUndo,
    canRedo,
    currentSheet,
    resetEditor,
    type CellRange,
  } from "./lib/editor-state.svelte";
  import {
    toggleBold,
    toggleItalic,
    toggleUnderline,
    setFontSize,
    setFontColor,
    setFill,
    setBorder,
    setAlignment,
    setNumberFormat,
    setFreeze,
    autoResizeColumn,
  } from "./lib/editor-state.svelte";
  import { csvToWorkbook } from "./lib/csv-handler";
  import { buildRef, parseRef, colToLetter, type CellRef } from "./lib/ooxml-parser";

  let fileInput: HTMLInputElement;
  let formulaInput: HTMLInputElement;

  // --- (canvas drag state removed — HTML DOM handles selection directly) ---

  // --- Sheet tab rename ---
  let renamingSheetIndex = $state<number | null>(null);
  let renameValue = $state("");

  // --- Formatting toolbar state ---
  let fontSizeValue = $state("11");
  let fontColorValue = $state("#000000");
  let fillColorValue = $state("#ffffff");
  let borderValue = $state("none");
  let alignmentValue = $state("left");
  let numberFormatValue = $state("general");

  // --- Find/Replace state ---
  let showFindReplace = $state(false);
  let findQuery = $state("");
  let replaceQuery = $state("");
  let findMatchCase = $state(false);

  let webgpuAvailable = $state(false);
  let gpuInfo = $state<string | null>(null);

  $effect(() => {
    checkWebGPU().then((ok: boolean) => { webgpuAvailable = ok; });
    getGPUInfo().then((info: string | null) => { gpuInfo = info; });
    loadKamiEngine();
  });

  // --- File handling ---
  /** Handle file upload for both .xlsx and .csv files. */
  async function handleFileUpload(e: Event): Promise<void> {
    const input = e.target as HTMLInputElement;
    const file = input.files?.[0];
    if (!file) return;

    let wb: XlsxWorkbook;
    if (file.name.toLowerCase().endsWith(".csv")) {
      const text = await file.text();
      wb = csvToWorkbook(text, file.name.replace(/\.csv$/i, ""));
    } else {
      const buffer = await file.arrayBuffer();
      wb = parseXlsx(buffer);
      wb.title = file.name.replace(/\.xlsx$/i, "");
    }

    recalculateAllSheets(wb);
    loadWorkbook(wb);
    requestRender();
  }

  function handleNew(): void {
    const wb: XlsxWorkbook = {
      id: `wb_${Date.now()}`,
      title: "New Workbook",
      sheets: [{
        id: `sheet_${Date.now()}`,
        name: "Sheet1",
        order: 0,
        hidden: false,
        cells: new Map(),
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
      }],
      sharedStrings: [],
      styles: [{ id: 0, numFmt: null, font: { name: "Calibri", size: 11, bold: false, italic: false, underline: false, strikethrough: false, color: "#000000" }, fill: { type: "pattern", fgColor: null, bgColor: null }, border: null, alignment: null }],
      definedNames: [],
      activeSheetIndex: 0,
    };
    loadWorkbook(wb);
    requestRender();
  }

  function handleExport(): void {
    if (!editor.workbook) return;
    const snap = $state.snapshot(editor.workbook) as XlsxWorkbook;
    // Restore Maps from snapshot (snapshot converts to plain objects)
    const restored: XlsxWorkbook = {
      ...snap,
      sheets: snap.sheets.map((s: any) => ({
        ...s,
        cells: s.cells instanceof Map ? s.cells : new Map(Object.entries(s.cells)),
        colWidths: s.colWidths instanceof Map ? s.colWidths : new Map(Object.entries(s.colWidths).map(([k, v]: [string, any]) => [Number(k), v])),
        rowHeights: s.rowHeights instanceof Map ? s.rowHeights : new Map(Object.entries(s.rowHeights).map(([k, v]: [string, any]) => [Number(k), v])),
      })),
    };
    const blob = exportXlsx(restored);
    downloadBlob(blob, `${editor.workbook.title || "workbook"}.xlsx`);
  }

  function recalculateAllSheets(wb: XlsxWorkbook): void {
    for (const sheet of wb.sheets) recalculateSheet(sheet);
  }

  // --- Rendering (HTML DOM — no canvas) ---
  // Svelte 5 reactivity handles re-rendering automatically via $state.
  // No manual requestRender/forceRender needed for DOM-based grid.

  /** Force reactivity trigger for Map mutations. */
  let renderVersion = $state(0);
  function forceRender(): void { renderVersion++; }
  function requestRender(): void { forceRender(); }

  /**
   * Get cell at ref — reads renderVersion to force Svelte re-evaluation
   * when Map contents change (Map.set/delete are invisible to $state).
   */
  function getCellAt(ref: CellRef): XlsxCell | undefined {
    void renderVersion; // dependency: forces re-read after forceRender()
    return currentSheet()?.cells.get(ref);
  }

  // (Canvas interaction handlers removed — replaced by HTML DOM event handlers above)

  // --- Keyboard ---
  function handleKeyDown(e: KeyboardEvent): void {
    if (!editor.workbook) return;

    // Ctrl shortcuts
    if (e.ctrlKey || e.metaKey) {
      switch (e.key.toLowerCase()) {
        case "z": e.preventDefault(); if (e.shiftKey) redo(); else undo(); requestRender(); return;
        case "y": e.preventDefault(); redo(); requestRender(); return;
        case "c": e.preventDefault(); handleCopy(); return;
        case "v": e.preventDefault(); handlePaste(); return;
        case "s": e.preventDefault(); handleExport(); return;
        case "n": e.preventDefault(); handleNew(); return;
        case "o": e.preventDefault(); fileInput?.click(); return;
        case "f": e.preventDefault(); showFindReplace = true; return;
      }
    }

    // Close Find/Replace on Escape (outside editing mode)
    if (e.key === "Escape" && showFindReplace) {
      showFindReplace = false;
      return;
    }

    // Editing mode
    if (editor.editingCell) {
      switch (e.key) {
        case "Enter": e.preventDefault(); commitEdit(); recalcCurrent(); moveActiveCell(1, 0); forceRender(); return;
        case "Tab": e.preventDefault(); commitEdit(); recalcCurrent(); { const ref = e.shiftKey ? prevCellTab(editor.activeCell) : nextCellTab(editor.activeCell, 16383); selectCell(ref); } forceRender(); return;
        case "Escape": e.preventDefault(); cancelEdit(); forceRender(); return;
        case "ArrowUp": e.preventDefault(); commitEdit(); recalcCurrent(); moveActiveCell(-1, 0); forceRender(); return;
        case "ArrowDown": e.preventDefault(); commitEdit(); recalcCurrent(); moveActiveCell(1, 0); forceRender(); return;
        case "ArrowLeft": e.preventDefault(); commitEdit(); recalcCurrent(); moveActiveCell(0, -1); forceRender(); return;
        case "ArrowRight": e.preventDefault(); commitEdit(); recalcCurrent(); moveActiveCell(0, 1); forceRender(); return;
      }
      return;
    }

    // Navigation mode
    switch (e.key) {
      case "ArrowUp": e.preventDefault(); moveActiveCell(-1, 0); forceRender(); return;
      case "ArrowDown": e.preventDefault(); moveActiveCell(1, 0); forceRender(); return;
      case "ArrowLeft": e.preventDefault(); moveActiveCell(0, -1); forceRender(); return;
      case "ArrowRight": e.preventDefault(); moveActiveCell(0, 1); forceRender(); return;
      case "Tab": e.preventDefault(); { const ref = e.shiftKey ? prevCellTab(editor.activeCell) : nextCellTab(editor.activeCell, 16383); selectCell(ref); } forceRender(); return;
      case "Enter": e.preventDefault(); startEdit(editor.activeCell); if (formulaInput) formulaInput.focus(); return;
      case "Delete": case "Backspace": e.preventDefault(); pushUndo(); { const sheet = currentSheet(); if (sheet) { const n = normalizeRange(editor.selection); for (let r = n.startRow; r <= n.endRow; r++) { for (let c = n.startCol; c <= n.endCol; c++) { sheet.cells.delete(buildRef(c, r)); } } } } forceRender(); return;
      case "F2": e.preventDefault(); startEdit(editor.activeCell); if (formulaInput) formulaInput.focus(); return;
    }

    // Start typing → enter edit mode
    if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
      startEdit(editor.activeCell);
      editor.editValue = e.key;
      if (formulaInput) { formulaInput.focus(); formulaInput.value = e.key; }
    }
  }

  async function handleCopy(): Promise<void> {
    const sheet = currentSheet();
    if (!sheet) return;
    copyCells();
    await copyRangeToClipboard(sheet, editor.selection);
  }

  async function handlePaste(): Promise<void> {
    const sheet = currentSheet();
    if (!sheet) return;
    try {
      const text = await navigator.clipboard.readText();
      if (text) {
        pushUndo();
        const values = parsePastedText(text);
        applyPastedValues(sheet, editor.activeCell, values);
        recalcCurrent();
        requestRender();
        return;
      }
    } catch {
      // fallback to internal clipboard
    }
    pasteCells();
    recalcCurrent();
    requestRender();
  }

  function recalcCurrent(): void {
    const sheet = currentSheet();
    if (sheet) recalculateSheet(sheet);
  }

  // --- Formula bar ---
  function handleFormulaInput(e: Event): void {
    editor.editValue = (e.target as HTMLInputElement).value;
    requestRender(); // live preview in cell
  }

  function handleFormulaKeyDown(e: KeyboardEvent): void {
    if (e.key === "Enter") {
      e.preventDefault();
      commitEdit();
      recalcCurrent();
      moveActiveCell(1, 0);
      forceRender();
      gridScrollEl?.focus();
    } else if (e.key === "Escape") {
      e.preventDefault();
      cancelEdit();
      forceRender();
      gridScrollEl?.focus();
    } else if (e.key === "Tab") {
      e.preventDefault();
      commitEdit();
      recalcCurrent();
      const ref = e.shiftKey ? prevCellTab(editor.activeCell) : nextCellTab(editor.activeCell, 16383);
      selectCell(ref);
      forceRender();
    }
  }

  // --- Sheet tabs ---
  function handleSheetTabClick(index: number): void {
    if (editor.editingCell) commitEdit();
    selectSheet(index);
    requestRender();
  }

  function handleSheetTabDblClick(index: number): void {
    renamingSheetIndex = index;
    renameValue = editor.workbook?.sheets[index]?.name ?? "";
  }

  function handleRenameCommit(): void {
    if (renamingSheetIndex !== null && renameValue.trim()) {
      renameSheet(renamingSheetIndex, renameValue.trim());
    }
    renamingSheetIndex = null;
  }

  function handleAddSheet(): void {
    addSheet();
    requestRender();
  }

  function handleDeleteSheet(index: number, e: MouseEvent): void {
    e.stopPropagation();
    deleteSheet(index);
    requestRender();
  }

  // --- HTML Grid Helpers ---
  let gridScrollEl: HTMLDivElement;
  const VISIBLE_ROWS = 30;
  const VISIBLE_COLS = 20;

  /** Visible row indices based on scroll. */
  function visibleRows(): number[] {
    const rows: number[] = [];
    for (let i = 0; i < VISIBLE_ROWS; i++) rows.push(editor.scrollRow + i);
    return rows;
  }

  /** Visible column indices based on scroll. */
  function visibleCols(): number[] {
    const cols: number[] = [];
    for (let i = 0; i < VISIBLE_COLS; i++) cols.push(editor.scrollCol + i);
    return cols;
  }

  /** Get pixel width for column. */
  function colWidthPx(col: number): number {
    const sheet = currentSheet();
    if (!sheet) return 64;
    const cw = sheet.colWidths.get(col) ?? sheet.defaultColWidth;
    return Math.round(cw * 7.5);
  }

  /** Check if cell is in current selection. */
  function isCellSelected(row: number, col: number): boolean {
    const s = editor.selection;
    const minR = Math.min(s.startRow, s.endRow), maxR = Math.max(s.startRow, s.endRow);
    const minC = Math.min(s.startCol, s.endCol), maxC = Math.max(s.startCol, s.endCol);
    return row >= minR && row <= maxR && col >= minC && col <= maxC;
  }

  /** Get cell display text. Reads renderVersion for reactivity. */
  function cellDisplay(cell: XlsxCell | undefined): string {
    void renderVersion;
    if (!cell) return "";
    if (cell.type === "formula") return cell.calculatedValue != null ? String(cell.calculatedValue) : "";
    return cell.value != null ? String(cell.value) : "";
  }

  /** Build inline style for cell from style table. */
  function cellStyle(cell: XlsxCell | undefined): string {
    if (!cell || !editor.workbook) return "";
    const style = editor.workbook.styles[cell.styleId];
    if (!style) return "";
    const parts: string[] = [];
    if (style.font?.bold) parts.push("font-weight:bold");
    if (style.font?.italic) parts.push("font-style:italic");
    if (style.font?.underline) parts.push("text-decoration:underline");
    if (style.font?.size) parts.push(`font-size:${style.font.size}px`);
    if (style.font?.color) parts.push(`color:#${style.font.color.replace(/^#/, "").replace(/^FF/i, "")}`);
    if (style.font?.name) parts.push(`font-family:${style.font.name},sans-serif`);
    if (style.fill?.fgColor) parts.push(`background:#${style.fill.fgColor.replace(/^#/, "").replace(/^FF/i, "")}`);
    if (style.alignment?.horizontal) parts.push(`text-align:${style.alignment.horizontal}`);
    if (cell.type === "number" || (cell.type === "formula" && typeof cell.calculatedValue === "number")) parts.push("text-align:right");
    return parts.join(";");
  }

  /** Handle cell click. */
  function handleCellClick(ref: CellRef, e: MouseEvent): void {
    if (editor.editingCell && editor.editingCell !== ref) {
      commitEdit();
      recalcCurrent();
      forceRender();
    }
    const { col, row } = parseRef(ref);
    if (e.shiftKey) {
      editor.selection = { ...editor.selection, endRow: row, endCol: col };
    } else {
      selectCell(ref);
    }
  }

  /** Handle cell double-click → enter edit mode. */
  function handleCellDblClick(ref: CellRef): void {
    selectCell(ref);
    startEdit(ref);
  }

  /** Handle keydown in cell editor input. */
  function handleCellEditorKeyDown(e: KeyboardEvent): void {
    if (e.key === "Enter") {
      e.preventDefault();
      commitEdit();
      recalcCurrent();
      moveActiveCell(1, 0);
      forceRender();
    } else if (e.key === "Tab") {
      e.preventDefault();
      commitEdit();
      recalcCurrent();
      const ref = e.shiftKey ? prevCellTab(editor.activeCell) : nextCellTab(editor.activeCell, 16383);
      selectCell(ref);
      forceRender();
    } else if (e.key === "Escape") {
      e.preventDefault();
      cancelEdit();
      forceRender();
    } else if (e.key === "ArrowUp" || e.key === "ArrowDown") {
      e.preventDefault();
      commitEdit();
      recalcCurrent();
      moveActiveCell(e.key === "ArrowUp" ? -1 : 1, 0);
      forceRender();
    }
  }

  /** Handle grid wheel for scrolling. */
  function handleGridWheel(e: WheelEvent): void {
    e.preventDefault();
    if (e.ctrlKey || e.metaKey) {
      editor.zoom = Math.max(0.25, Math.min(4, editor.zoom + (e.deltaY > 0 ? -0.1 : 0.1)));
    } else if (e.shiftKey) {
      editor.scrollCol = Math.max(0, editor.scrollCol + Math.sign(e.deltaY) * 3);
    } else {
      editor.scrollRow = Math.max(0, editor.scrollRow + Math.sign(e.deltaY) * 3);
    }
  }

  /** Handle mousedown on the grid table for drag selection. */
  function handleGridMouseDown(e: MouseEvent): void {
    // Drag selection is handled per-cell via handleCellClick
  }

  /** Svelte action: auto-focus the cell editor input when it mounts. */
  function autoFocus(node: HTMLInputElement) {
    node.focus();
    // Place cursor at end
    const len = node.value.length;
    node.setSelectionRange(len, len);
  }

  // --- Selection info ---
  function selectionInfo(): string {
    const sheet = currentSheet();
    if (!sheet) return "";
    const n = normalizeRange(editor.selection);
    const count = (n.endRow - n.startRow + 1) * (n.endCol - n.startCol + 1);
    if (count <= 1) return editor.activeCell;
    return `${rangeToString(editor.selection)} (${count} cells)`;
  }

  /** Compute SUM, AVERAGE, COUNT for numeric cells in current selection. */
  function selectionStats(): { sum: number; avg: number; count: number } | null {
    const sheet = currentSheet();
    if (!sheet) return null;
    const n = normalizeRange(editor.selection);
    const cellCount = (n.endRow - n.startRow + 1) * (n.endCol - n.startCol + 1);
    if (cellCount <= 1) return null;

    let sum = 0;
    let count = 0;
    for (let r = n.startRow; r <= n.endRow; r++) {
      for (let c = n.startCol; c <= n.endCol; c++) {
        const ref = buildRef(c, r);
        const cell = sheet.cells.get(ref);
        if (!cell) continue;
        const v = cell.computed !== undefined ? cell.computed : cell.value;
        const num = typeof v === "number" ? v : parseFloat(String(v));
        if (!isNaN(num)) {
          sum += num;
          count++;
        }
      }
    }
    if (count === 0) return null;
    return { sum, avg: sum / count, count };
  }

  /** Find next matching cell from current position. */
  function findNext(): void {
    if (!findQuery) return;
    const sheet = currentSheet();
    if (!sheet) return;
    const { col: startCol, row: startRow } = parseRef(editor.activeCell);
    const query = findMatchCase ? findQuery : findQuery.toLowerCase();

    // Scan from current position, wrapping around
    for (const [ref, cell] of sheet.cells) {
      const { col, row } = parseRef(ref);
      if (row < startRow || (row === startRow && col <= startCol)) continue;
      const val = String(cell.computed !== undefined ? cell.computed : cell.value ?? "");
      const compare = findMatchCase ? val : val.toLowerCase();
      if (compare.includes(query)) {
        selectCell(ref);
        requestRender();
        return;
      }
    }
    // Wrap from beginning
    for (const [ref, cell] of sheet.cells) {
      const val = String(cell.computed !== undefined ? cell.computed : cell.value ?? "");
      const compare = findMatchCase ? val : val.toLowerCase();
      if (compare.includes(query)) {
        selectCell(ref);
        requestRender();
        return;
      }
    }
  }

  /** Replace value in active cell if it matches find query. */
  function replaceCurrent(): void {
    if (!findQuery) return;
    const sheet = currentSheet();
    if (!sheet) return;
    const cell = sheet.cells.get(editor.activeCell);
    if (!cell) return;
    const val = String(cell.computed !== undefined ? cell.computed : cell.value ?? "");
    const compare = findMatchCase ? val : val.toLowerCase();
    const query = findMatchCase ? findQuery : findQuery.toLowerCase();
    if (compare.includes(query)) {
      pushUndo();
      const newVal = findMatchCase
        ? val.replace(findQuery, replaceQuery)
        : val.replace(new RegExp(findQuery.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "i"), replaceQuery);
      setCellValue(editor.activeCell, newVal);
      recalcCurrent();
      requestRender();
    }
    findNext();
  }

  /** Replace all matching cells in the current sheet. */
  function replaceAll(): void {
    if (!findQuery) return;
    const sheet = currentSheet();
    if (!sheet) return;
    pushUndo();
    const query = findMatchCase ? findQuery : findQuery.toLowerCase();
    const regex = new RegExp(findQuery.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), findMatchCase ? "g" : "gi");
    for (const [ref, cell] of sheet.cells) {
      const val = String(cell.computed !== undefined ? cell.computed : cell.value ?? "");
      const compare = findMatchCase ? val : val.toLowerCase();
      if (compare.includes(query)) {
        setCellValue(ref, val.replace(regex, replaceQuery));
      }
    }
    recalcCurrent();
    requestRender();
  }
</script>

<svelte:window onkeydown={handleKeyDown} />

<div class="editor">
  {#if !editor.workbook}
    <!-- Landing -->
    <div class="landing">
      <div class="landing-card">
        <h1>XLSX Editor</h1>
        <p>xlsx.etzhayyim.com — Excel spreadsheet editor with formula evaluation</p>
        <div class="landing-actions">
          <button class="btn btn-primary" onclick={handleNew}>New Workbook</button>
          <button class="btn" onclick={() => fileInput?.click()}>Open .xlsx / .csv</button>
        </div>
      </div>
    </div>
  {:else}
    <!-- Toolbar -->
    <div class="toolbar">
      <div class="toolbar-group">
        <button class="btn btn-sm" onclick={handleNew} title="New (Ctrl+N)">New</button>
        <button class="btn btn-sm" onclick={() => fileInput?.click()} title="Open (Ctrl+O)">Open</button>
        <button class="btn btn-sm" onclick={handleExport} title="Save (Ctrl+S)">Save</button>
      </div>
      <div class="toolbar-sep"></div>
      <div class="toolbar-group">
        <button class="btn btn-sm" onclick={() => { undo(); requestRender(); }} disabled={!canUndo()} title="Undo (Ctrl+Z)">Undo</button>
        <button class="btn btn-sm" onclick={() => { redo(); requestRender(); }} disabled={!canRedo()} title="Redo (Ctrl+Y)">Redo</button>
      </div>
      <div class="toolbar-sep"></div>
      <div class="toolbar-group">
        <button class="btn btn-sm" onclick={() => { insertRows(editor.selection.startRow, 1); requestRender(); }}>+Row</button>
        <button class="btn btn-sm" onclick={() => { deleteRows(editor.selection.startRow, 1); requestRender(); }}>-Row</button>
        <button class="btn btn-sm" onclick={() => { insertColumns(editor.selection.startCol, 1); requestRender(); }}>+Col</button>
        <button class="btn btn-sm" onclick={() => { deleteColumns(editor.selection.startCol, 1); requestRender(); }}>-Col</button>
      </div>
      <div class="toolbar-sep"></div>
      <div class="toolbar-group">
        <button class="btn btn-sm" onclick={() => { mergeCells(); requestRender(); }}>Merge</button>
        <button class="btn btn-sm" onclick={() => { unmergeCells(); requestRender(); }}>Unmerge</button>
      </div>
      <div class="toolbar-spacer"></div>
      <div class="toolbar-group">
        <span class="status-text">{editor.statusMessage}</span>
      </div>
    </div>

    <!-- Formatting toolbar -->
    <div class="toolbar formatting-toolbar">
      <div class="toolbar-group">
        <button class="btn btn-sm fmt-btn" title="Bold (Ctrl+B)" onclick={() => { toggleBold(); requestRender(); }}><strong>B</strong></button>
        <button class="btn btn-sm fmt-btn" title="Italic (Ctrl+I)" onclick={() => { toggleItalic(); requestRender(); }}><em>I</em></button>
        <button class="btn btn-sm fmt-btn" title="Underline (Ctrl+U)" onclick={() => { toggleUnderline(); requestRender(); }}><u>U</u></button>
      </div>
      <div class="toolbar-sep"></div>
      <div class="toolbar-group">
        <select class="fmt-select" title="Font Size" value={fontSizeValue} onchange={(e: Event) => { fontSizeValue = (e.target as HTMLSelectElement).value; setFontSize(Number(fontSizeValue)); requestRender(); }}>
          {#each [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72] as size}
            <option value={String(size)}>{size}</option>
          {/each}
        </select>
      </div>
      <div class="toolbar-sep"></div>
      <div class="toolbar-group">
        <label class="color-picker-label" title="Font Color">
          <span class="color-icon" style="border-bottom: 3px solid {fontColorValue}">A</span>
          <input type="color" class="color-input" value={fontColorValue} oninput={(e: Event) => { fontColorValue = (e.target as HTMLInputElement).value; setFontColor(fontColorValue); requestRender(); }} />
        </label>
        <label class="color-picker-label" title="Fill Color">
          <span class="color-icon fill-icon" style="background: {fillColorValue}"></span>
          <input type="color" class="color-input" value={fillColorValue} oninput={(e: Event) => { fillColorValue = (e.target as HTMLInputElement).value; setFill(fillColorValue); requestRender(); }} />
        </label>
      </div>
      <div class="toolbar-sep"></div>
      <div class="toolbar-group">
        <select class="fmt-select" title="Borders" value={borderValue} onchange={(e: Event) => { borderValue = (e.target as HTMLSelectElement).value; setBorder(borderValue); requestRender(); }}>
          <option value="none">No borders</option>
          <option value="all">All borders</option>
          <option value="top">Top</option>
          <option value="bottom">Bottom</option>
          <option value="left">Left</option>
          <option value="right">Right</option>
        </select>
      </div>
      <div class="toolbar-sep"></div>
      <div class="toolbar-group">
        <button class="btn btn-sm fmt-btn" title="Align Left" onclick={() => { alignmentValue = "left"; setAlignment("left"); requestRender(); }}>&#9776;</button>
        <button class="btn btn-sm fmt-btn" title="Align Center" onclick={() => { alignmentValue = "center"; setAlignment("center"); requestRender(); }}>&#9776;</button>
        <button class="btn btn-sm fmt-btn" title="Align Right" onclick={() => { alignmentValue = "right"; setAlignment("right"); requestRender(); }}>&#9776;</button>
      </div>
      <div class="toolbar-sep"></div>
      <div class="toolbar-group">
        <select class="fmt-select" title="Number Format" value={numberFormatValue} onchange={(e: Event) => { numberFormatValue = (e.target as HTMLSelectElement).value; setNumberFormat(numberFormatValue); requestRender(); }}>
          <option value="general">General</option>
          <option value="number">Number</option>
          <option value="currency">Currency</option>
          <option value="percent">Percent</option>
          <option value="date">Date</option>
          <option value="text">Text</option>
        </select>
      </div>
      <div class="toolbar-sep"></div>
      <div class="toolbar-group">
        <button class="btn btn-sm" title="Freeze Panes" onclick={() => { setFreeze(); requestRender(); }}>Freeze</button>
        <button class="btn btn-sm" title="Auto-Resize Column" onclick={() => { autoResizeColumn(); requestRender(); }}>AutoFit</button>
      </div>
    </div>

    <!-- Formula bar -->
    {#if editor.showFormulaBar}
      <div class="formula-bar">
        <div class="name-box">{selectionInfo()}</div>
        <div class="formula-sep"></div>
        <span class="fx-label">fx</span>
        <input
          bind:this={formulaInput}
          class="formula-input"
          value={editor.editingCell ? editor.editValue : (() => { const sheet = currentSheet(); const cell = sheet?.cells.get(editor.activeCell); if (!cell) return ""; return cell.formula ? `=${cell.formula}` : String(cell.value ?? ""); })()}
          oninput={handleFormulaInput}
          onkeydown={handleFormulaKeyDown}
          onfocus={() => { if (!editor.editingCell) startEdit(editor.activeCell); }}
        />
      </div>
    {/if}

    <!-- Grid (HTML DOM with virtualization) -->
    <div class="grid-container" onwheel={handleGridWheel}>
      <div class="grid-scroll" bind:this={gridScrollEl}>
        <table class="grid-table" onmousedown={handleGridMouseDown}>
          <!-- Column headers -->
          <thead>
            <tr>
              <th class="corner-header"></th>
              {#each visibleCols() as col}
                <th class="col-header" style="width:{colWidthPx(col)}px;min-width:{colWidthPx(col)}px">{colToLetter(col)}</th>
              {/each}
            </tr>
          </thead>
          <tbody>
            {#each visibleRows() as row}
              <tr>
                <td class="row-header">{row + 1}</td>
                {#each visibleCols() as col}
                  {@const ref = buildRef(col, row)}
                  {@const cell = getCellAt(ref)}
                  {@const isActive = editor.activeCell === ref}
                  {@const isSelected = isCellSelected(row, col)}
                  {@const isEditing = editor.editingCell === ref}
                  <td
                    class="grid-cell"
                    class:selected={isSelected}
                    class:active={isActive}
                    class:editing={isEditing}
                    style="{cellStyle(cell)}"
                    onmousedown={(e) => handleCellClick(ref, e)}
                    ondblclick={() => handleCellDblClick(ref)}
                  >
                    {#if isEditing}
                      <input
                        class="cell-editor"
                        value={editor.editValue}
                        oninput={(e) => { editor.editValue = (e.target as HTMLInputElement).value; }}
                        onkeydown={handleCellEditorKeyDown}
                        onfocus={() => {}}
                        use:autoFocus
                      />
                    {:else}
                      {cellDisplay(cell)}
                    {/if}
                  </td>
                {/each}
              </tr>
            {/each}
          </tbody>
        </table>
      </div>
    </div>

    <!-- Find/Replace dialog -->
    {#if showFindReplace}
      <div class="find-replace-overlay">
        <div class="find-replace-dialog">
          <div class="find-replace-header">
            <span class="find-replace-title">Find and Replace</span>
            <button class="btn btn-sm find-replace-close" onclick={() => { showFindReplace = false; }}>&times;</button>
          </div>
          <div class="find-replace-row">
            <label class="find-replace-label">Find:</label>
            <input
              class="find-replace-input"
              value={findQuery}
              oninput={(e: Event) => { findQuery = (e.target as HTMLInputElement).value; }}
              onkeydown={(e: KeyboardEvent) => { if (e.key === "Enter") findNext(); if (e.key === "Escape") { showFindReplace = false; } }}
              placeholder="Search..."
            />
          </div>
          <div class="find-replace-row">
            <label class="find-replace-label">Replace:</label>
            <input
              class="find-replace-input"
              value={replaceQuery}
              oninput={(e: Event) => { replaceQuery = (e.target as HTMLInputElement).value; }}
              onkeydown={(e: KeyboardEvent) => { if (e.key === "Enter") replaceCurrent(); if (e.key === "Escape") { showFindReplace = false; } }}
              placeholder="Replace with..."
            />
          </div>
          <div class="find-replace-row">
            <label class="find-replace-checkbox">
              <input type="checkbox" checked={findMatchCase} onchange={(e: Event) => { findMatchCase = (e.target as HTMLInputElement).checked; }} />
              Match case
            </label>
          </div>
          <div class="find-replace-actions">
            <button class="btn btn-sm" onclick={findNext}>Find Next</button>
            <button class="btn btn-sm" onclick={replaceCurrent}>Replace</button>
            <button class="btn btn-sm" onclick={replaceAll}>Replace All</button>
          </div>
        </div>
      </div>
    {/if}

    <!-- Sheet tabs -->
    <div class="sheet-tabs">
      {#each editor.workbook.sheets as sheet, i}
        <button
          class="sheet-tab"
          class:active={i === editor.activeSheetIndex}
          onclick={() => handleSheetTabClick(i)}
          ondblclick={() => handleSheetTabDblClick(i)}
          type="button"
        >
          {#if renamingSheetIndex === i}
            <input
              class="sheet-rename-input"
              value={renameValue}
              oninput={(e: Event) => { renameValue = (e.target as HTMLInputElement).value; }}
              onblur={handleRenameCommit}
              onkeydown={(e: KeyboardEvent) => { if (e.key === "Enter") handleRenameCommit(); if (e.key === "Escape") { renamingSheetIndex = null; } }}
            />
          {:else}
            <span>{sheet.name}</span>
            {#if editor.workbook.sheets.length > 1}
              <span class="sheet-close" role="button" tabindex="-1" onclick={(e: MouseEvent) => handleDeleteSheet(i, e)} onkeydown={(e: KeyboardEvent) => { if (e.key === "Enter") handleDeleteSheet(i, e as unknown as MouseEvent); }}>&times;</span>
            {/if}
          {/if}
        </button>
      {/each}
      <button class="sheet-add" onclick={handleAddSheet}>+</button>
      <div class="sheet-tabs-spacer"></div>
      {#if selectionStats()}
        {@const stats = selectionStats()!}
        <span class="status-bar-stats">
          SUM: {stats.sum.toLocaleString(undefined, { maximumFractionDigits: 4 })}
          &nbsp; AVG: {stats.avg.toLocaleString(undefined, { maximumFractionDigits: 4 })}
          &nbsp; COUNT: {stats.count}
        </span>
      {/if}
      <div class="zoom-info">{Math.round(editor.zoom * 100)}%</div>
    </div>
  {/if}

  <input bind:this={fileInput} type="file" accept=".xlsx,.csv" onchange={handleFileUpload} style="display:none" />
</div>

<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  .editor { display: flex; flex-direction: column; width: 100vw; height: 100vh; background: #ffffff; color: #333; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; overflow: hidden; }

  /* Landing */
  .landing { display: flex; align-items: center; justify-content: center; flex: 1; background: #f5f5f5; }
  .landing-card { text-align: center; padding: 48px; background: #fff; border-radius: 12px; box-shadow: 0 2px 12px rgba(0,0,0,0.08); }
  .landing-card h1 { font-size: 28px; font-weight: 700; margin-bottom: 8px; color: #1a73e8; }
  .landing-card p { color: #666; margin-bottom: 24px; }
  .landing-actions { display: flex; gap: 12px; justify-content: center; }

  /* Buttons */
  .btn { padding: 8px 16px; border: 1px solid #dadce0; border-radius: 4px; background: #fff; color: #333; cursor: pointer; font-size: 13px; }
  .btn:hover { background: #f1f3f4; }
  .btn-primary { background: #1a73e8; color: #fff; border-color: #1a73e8; }
  .btn-primary:hover { background: #1557b0; }
  .btn-sm { padding: 4px 10px; font-size: 12px; }
  .btn:disabled { opacity: 0.4; cursor: default; }

  /* Toolbar */
  .toolbar { display: flex; align-items: center; gap: 4px; padding: 4px 8px; background: #f8f9fa; border-bottom: 1px solid #dadce0; min-height: 36px; flex-shrink: 0; }
  .toolbar-group { display: flex; gap: 2px; align-items: center; }
  .toolbar-sep { width: 1px; height: 20px; background: #dadce0; margin: 0 4px; }
  .toolbar-spacer { flex: 1; }
  .status-text { font-size: 11px; color: #888; }

  /* Formula bar */
  .formula-bar { display: flex; align-items: center; border-bottom: 1px solid #dadce0; height: 28px; flex-shrink: 0; }
  .name-box { width: 100px; padding: 0 8px; font-size: 12px; border-right: 1px solid #dadce0; height: 100%; display: flex; align-items: center; background: #f8f9fa; white-space: nowrap; overflow: hidden; }
  .formula-sep { width: 0; }
  .fx-label { padding: 0 6px; font-size: 12px; font-style: italic; color: #888; }
  .formula-input { flex: 1; border: none; outline: none; padding: 0 4px; font-size: 13px; font-family: 'Consolas', 'Monaco', monospace; height: 100%; }
  .formula-input:focus { background: #e8f0fe; }

  /* Grid (HTML DOM) */
  .grid-container { flex: 1; position: relative; overflow: auto; }
  .grid-scroll { min-width: 100%; min-height: 100%; }
  .grid-table { border-collapse: collapse; table-layout: fixed; user-select: none; font-size: 13px; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; }
  .corner-header { width: 50px; min-width: 50px; height: 24px; background: #f3f3f3; border: 1px solid #d0d0d0; position: sticky; top: 0; left: 0; z-index: 3; }
  .col-header { height: 24px; background: #f3f3f3; border: 1px solid #d0d0d0; text-align: center; font-size: 11px; color: #555; font-weight: normal; position: sticky; top: 0; z-index: 2; }
  .row-header { width: 50px; min-width: 50px; background: #f3f3f3; border: 1px solid #d0d0d0; text-align: center; font-size: 11px; color: #555; position: sticky; left: 0; z-index: 1; }
  .grid-cell { height: 20px; border: 1px solid #e0e0e0; padding: 0 3px; white-space: nowrap; overflow: hidden; cursor: cell; position: relative; }
  .grid-cell.selected { background: rgba(66, 133, 244, 0.08); }
  .grid-cell.active { outline: 2px solid #1a73e8; outline-offset: -1px; z-index: 1; }
  .grid-cell.editing { padding: 0; overflow: visible; }
  .cell-editor { width: 100%; height: 100%; border: none; outline: none; padding: 0 3px; font-size: 13px; font-family: inherit; background: #fff; box-shadow: 0 0 0 2px #1a73e8; position: relative; z-index: 2; }

  /* Sheet tabs */
  .sheet-tabs { display: flex; align-items: center; height: 28px; background: #f8f9fa; border-top: 1px solid #dadce0; padding: 0 4px; flex-shrink: 0; gap: 1px; overflow-x: auto; }
  .sheet-tab { display: flex; align-items: center; gap: 4px; padding: 4px 12px; font-size: 12px; cursor: pointer; border: 1px solid transparent; border-radius: 4px 4px 0 0; white-space: nowrap; user-select: none; }
  .sheet-tab:hover { background: #e8eaed; }
  .sheet-tab.active { background: #fff; border-color: #dadce0; border-bottom-color: #fff; font-weight: 600; }
  .sheet-close { background: none; border: none; cursor: pointer; font-size: 14px; color: #888; padding: 0 2px; line-height: 1; }
  .sheet-close:hover { color: #d93025; }
  .sheet-add { background: none; border: 1px solid #dadce0; border-radius: 4px; cursor: pointer; font-size: 16px; padding: 2px 8px; color: #555; }
  .sheet-add:hover { background: #e8eaed; }
  .sheet-rename-input { border: 1px solid #1a73e8; border-radius: 2px; font-size: 12px; padding: 1px 4px; width: 80px; outline: none; }
  .sheet-tabs-spacer { flex: 1; }
  .zoom-info { font-size: 11px; color: #888; padding: 0 8px; }

  /* Formatting toolbar */
  .formatting-toolbar { gap: 3px; padding: 2px 8px; min-height: 32px; }
  .fmt-btn { min-width: 28px; padding: 3px 6px; font-size: 13px; text-align: center; }
  .fmt-btn strong { font-weight: 700; }
  .fmt-btn em { font-style: italic; }
  .fmt-btn u { text-decoration: underline; }
  .fmt-select { padding: 2px 4px; font-size: 12px; border: 1px solid #dadce0; border-radius: 3px; background: #fff; cursor: pointer; height: 26px; }
  .fmt-select:hover { background: #f1f3f4; }
  .color-picker-label { position: relative; display: inline-flex; align-items: center; cursor: pointer; }
  .color-icon { font-size: 14px; font-weight: 700; padding: 2px 6px; display: inline-block; line-height: 1.2; }
  .fill-icon { width: 16px; height: 16px; border: 1px solid #dadce0; border-radius: 2px; display: inline-block; }
  .color-input { position: absolute; opacity: 0; width: 0; height: 0; pointer-events: none; }
  .color-picker-label:hover .color-icon { background: #e8eaed; border-radius: 3px; }

  /* Find/Replace */
  .find-replace-overlay { position: absolute; top: 80px; right: 16px; z-index: 100; }
  .find-replace-dialog { background: #fff; border: 1px solid #dadce0; border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12); padding: 16px; min-width: 340px; }
  .find-replace-header { display: flex; align-items: center; justify-content: space-between; margin-bottom: 12px; }
  .find-replace-title { font-size: 14px; font-weight: 600; color: #333; }
  .find-replace-close { border: none; background: none; font-size: 18px; color: #888; cursor: pointer; padding: 0 4px; }
  .find-replace-close:hover { color: #d93025; }
  .find-replace-row { display: flex; align-items: center; gap: 8px; margin-bottom: 8px; }
  .find-replace-label { font-size: 12px; color: #555; min-width: 52px; }
  .find-replace-input { flex: 1; padding: 4px 8px; border: 1px solid #dadce0; border-radius: 4px; font-size: 13px; outline: none; }
  .find-replace-input:focus { border-color: #1a73e8; box-shadow: 0 0 0 2px rgba(26,115,232,0.15); }
  .find-replace-checkbox { font-size: 12px; color: #555; display: flex; align-items: center; gap: 4px; margin-left: 60px; }
  .find-replace-actions { display: flex; gap: 6px; margin-top: 12px; justify-content: flex-end; }

  /* Status bar stats */
  .status-bar-stats { font-size: 11px; color: #555; padding: 0 12px; white-space: nowrap; }
</style>
