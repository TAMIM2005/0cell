/**
 * 0cell Modern Grid Engine
 * Full MS Excel Interaction Model:
 * 1. Instant Typing on Cell Click / Double-Click / F2 with floating in-cell editor
 * 2. Enter moves down, Tab moves right, Escape cancels
 * 3. Cut Rows/Columns (Ctrl+X) with pulsating marching-ants dashed border
 * 4. Insert Cut Cells / Insert Copied Cells (Shift Down) in context menu & shortcuts
 * 5. Full Merged Cells (colspan/rowspan), Custom Widths, Custom Heights, and Undo/Redo.
 */

class WorkbookManager {
  constructor() {
    this.name = 'Workbook1';
    this.sheets = [];
    this.activeSheetIndex = 0;
    this.undoStack = [];
    this.redoStack = [];
    this.maxUndo = 50;

    this.addSheet('Sheet1');
  }

  addSheet(name) {
    const sheetName = name || `Sheet${this.sheets.length + 1}`;
    const newSheet = new Worksheet(sheetName);
    this.sheets.push(newSheet);
    this.activeSheetIndex = this.sheets.length - 1;
    return newSheet;
  }

  getActiveSheet() {
    return this.sheets[this.activeSheetIndex] || this.sheets[0];
  }

  getSheet(name) {
    return this.sheets.find(s => s.name.toLowerCase() === (name || '').toLowerCase());
  }

  removeSheet(index) {
    if (this.sheets.length <= 1) return false;
    this.sheets.splice(index, 1);
    if (this.activeSheetIndex >= this.sheets.length) {
      this.activeSheetIndex = this.sheets.length - 1;
    }
    return true;
  }

  pushUndoState(actionName = 'Edit') {
    const snapshot = JSON.stringify(this.serialize());
    this.undoStack.push({ name: actionName, data: snapshot });
    if (this.undoStack.length > this.maxUndo) this.undoStack.shift();
    this.redoStack = [];
  }

  undo() {
    if (this.undoStack.length === 0) return false;
    const currentState = JSON.stringify(this.serialize());
    this.redoStack.push({ name: 'Redo', data: currentState });

    const prevState = this.undoStack.pop();
    this.deserialize(JSON.parse(prevState.data));
    return true;
  }

  redo() {
    if (this.redoStack.length === 0) return false;
    const currentState = JSON.stringify(this.serialize());
    this.undoStack.push({ name: 'Undo', data: currentState });

    const nextState = this.redoStack.pop();
    this.deserialize(JSON.parse(nextState.data));
    return true;
  }

  serialize() {
    return {
      name: this.name,
      activeSheetIndex: this.activeSheetIndex,
      sheets: this.sheets.map(s => s.serialize())
    };
  }

  deserialize(data) {
    this.name = data.name || 'Workbook1';
    this.sheets = (data.sheets || []).map(sData => {
      const s = new Worksheet(sData.name);
      s.deserialize(sData);
      return s;
    });
    if (this.sheets.length === 0) this.addSheet('Sheet1');
    this.activeSheetIndex = Math.min(data.activeSheetIndex || 0, this.sheets.length - 1);
  }
}

class Worksheet {
  constructor(name) {
    this.name = name;
    this.cells = {}; // "col,row" -> { value, formula, style }
    this.merges = []; // [ { startCol, startRow, endCol, endRow } ]
    this.colWidths = {}; // colIdx -> width in px
    this.rowHeights = {}; // rowIdx -> height in px
    this.numCols = 26;
    this.numRows = 50;
    this.selection = { startCol: 0, startRow: 0, endCol: 0, endRow: 0 };
    this.activeCell = { col: 0, row: 0 };
    this.autoFilterRange = null;
    this.hiddenRows = new Set();
    this.charts = [];
    this.images = [];
  }

  getCell(col, row) {
    return this.cells[`${col},${row}`] || null;
  }

  getOrCreateCell(col, row) {
    const key = `${col},${row}`;
    if (!this.cells[key]) {
      this.cells[key] = { value: '', formula: '', style: {} };
    }
    return this.cells[key];
  }

  setCellValue(col, row, val) {
    const cell = this.getOrCreateCell(col, row);
    if (typeof val === 'string' && val.startsWith('=')) {
      cell.formula = val;
      cell.value = val;
    } else {
      cell.formula = '';
      if (!isNaN(val) && val !== '' && typeof val !== 'boolean') {
        cell.value = Number(val);
      } else {
        cell.value = val;
      }
    }
    if (col >= this.numCols) this.numCols = col + 5;
    if (row >= this.numRows) this.numRows = row + 10;
  }

  deleteCell(col, row) {
    delete this.cells[`${col},${row}`];
  }

  // --- Merges ---
  getMergeOrigin(col, row) {
    if (!this.merges) return null;
    for (const m of this.merges) {
      if (m.startCol === col && m.startRow === row) {
        return m;
      }
    }
    return null;
  }

  isMergeCovered(col, row) {
    if (!this.merges) return false;
    for (const m of this.merges) {
      if (col >= m.startCol && col <= m.endCol && row >= m.startRow && row <= m.endRow) {
        if (col !== m.startCol || row !== m.startRow) {
          return true;
        }
      }
    }
    return false;
  }

  mergeSelection() {
    const sel = this.selection;
    if (!sel || (sel.startCol === sel.endCol && sel.startRow === sel.endRow)) return;

    const minC = Math.min(sel.startCol, sel.endCol);
    const maxC = Math.max(sel.startCol, sel.endCol);
    const minR = Math.min(sel.startRow, sel.endRow);
    const maxR = Math.max(sel.startRow, sel.endRow);

    this.unmergeSelection();
    this.merges.push({ startCol: minC, startRow: minR, endCol: maxC, endRow: maxR });
  }

  unmergeSelection() {
    const sel = this.selection;
    if (!sel || !this.merges) return;

    const minC = Math.min(sel.startCol, sel.endCol);
    const maxC = Math.max(sel.startCol, sel.endCol);
    const minR = Math.min(sel.startRow, sel.endRow);
    const maxR = Math.max(sel.startRow, sel.endRow);

    this.merges = this.merges.filter(m => {
      const overlap = !(maxC < m.startCol || minC > m.endCol || maxR < m.startRow || minR > m.endRow);
      return !overlap;
    });
  }

  // --- Row/Column Operations ---
  insertRow(atRow) {
    const newCells = {};
    Object.keys(this.cells).forEach(key => {
      const [c, r] = key.split(',').map(Number);
      if (r < atRow) {
        newCells[`${c},${r}`] = this.cells[key];
      } else {
        newCells[`${c},${r + 1}`] = this.cells[key];
      }
    });
    this.cells = newCells;

    this.merges.forEach(m => {
      if (m.startRow >= atRow) m.startRow++;
      if (m.endRow >= atRow) m.endRow++;
    });

    this.numRows++;
  }

  deleteRow(atRow) {
    if (this.numRows <= 1) return;
    const newCells = {};
    Object.keys(this.cells).forEach(key => {
      const [c, r] = key.split(',').map(Number);
      if (r < atRow) {
        newCells[`${c},${r}`] = this.cells[key];
      } else if (r > atRow) {
        newCells[`${c},${r - 1}`] = this.cells[key];
      }
    });
    this.cells = newCells;

    this.merges.forEach(m => {
      if (m.startRow > atRow) m.startRow--;
      if (m.endRow >= atRow) m.endRow = Math.max(m.startRow, m.endRow - 1);
    });

    this.numRows = Math.max(1, this.numRows - 1);
  }

  insertColumn(atCol) {
    const newCells = {};
    Object.keys(this.cells).forEach(key => {
      const [c, r] = key.split(',').map(Number);
      if (c < atCol) {
        newCells[`${c},${r}`] = this.cells[key];
      } else {
        newCells[`${c + 1},${r}`] = this.cells[key];
      }
    });
    this.cells = newCells;

    this.merges.forEach(m => {
      if (m.startCol >= atCol) m.startCol++;
      if (m.endCol >= atCol) m.endCol++;
    });

    this.numCols++;
  }

  deleteColumn(atCol) {
    if (this.numCols <= 1) return;
    const newCells = {};
    Object.keys(this.cells).forEach(key => {
      const [c, r] = key.split(',').map(Number);
      if (c < atCol) {
        newCells[`${c},${r}`] = this.cells[key];
      } else if (c > atCol) {
        newCells[`${c - 1},${r}`] = this.cells[key];
      }
    });
    this.cells = newCells;

    this.merges.forEach(m => {
      if (m.startCol > atCol) m.startCol--;
      if (m.endCol >= atCol) m.endCol = Math.max(m.startCol, m.endCol - 1);
    });

    this.numCols = Math.max(1, this.numCols - 1);
  }

  // --- MS Excel Move Row Range (Zero Data Loss) ---
  moveRowRange(srcStartRow, srcEndRow, targetRow) {
    if (targetRow >= srcStartRow && targetRow <= srcEndRow + 1) {
      return { destStartRow: srcStartRow, destEndRow: srcEndRow };
    }

    const count = srcEndRow - srcStartRow + 1;

    // 1. Snapshot rows to move
    const savedRows = [];
    for (let r = srcStartRow; r <= srcEndRow; r++) {
      const rowCells = {};
      for (let c = 0; c < this.numCols; c++) {
        const cell = this.getCell(c, r);
        if (cell) rowCells[c] = JSON.parse(JSON.stringify(cell));
      }
      savedRows.push({
        cells: rowCells,
        height: this.rowHeights[r] || null
      });
    }

    // 2. Delete source rows
    for (let i = 0; i < count; i++) {
      this.deleteRow(srcStartRow);
    }

    // 3. Compute destination index after deletion
    let destRow = targetRow;
    if (targetRow > srcEndRow) {
      destRow = targetRow - count;
    }

    // 4. Insert new rows at destination
    for (let i = 0; i < count; i++) {
      this.insertRow(destRow);
    }

    // 5. Populate saved data
    savedRows.forEach((rData, offset) => {
      const r = destRow + offset;
      if (rData.height) this.rowHeights[r] = rData.height;
      Object.keys(rData.cells).forEach(col => {
        this.cells[`${col},${r}`] = rData.cells[col];
      });
    });

    return { destStartRow: destRow, destEndRow: destRow + count - 1 };
  }

  serialize() {
    return {
      name: this.name,
      cells: this.cells,
      merges: this.merges,
      colWidths: this.colWidths,
      rowHeights: this.rowHeights,
      numCols: this.numCols,
      numRows: this.numRows,
      autoFilterRange: this.autoFilterRange,
      charts: this.charts,
      images: this.images
    };
  }

  deserialize(data) {
    this.name = data.name || this.name;
    this.cells = data.cells || {};
    this.merges = data.merges || [];
    this.colWidths = data.colWidths || {};
    this.rowHeights = data.rowHeights || {};
    this.numCols = data.numCols || 26;
    this.numRows = data.numRows || 50;
    this.autoFilterRange = data.autoFilterRange || null;
    this.charts = data.charts || [];
    this.images = data.images || [];
  }
}

class GridEngine {
  constructor(app, containerEl) {
    this.app = app;
    this.container = containerEl;
    this.isSelecting = false;
    this.internalClipboard = null;
    this.cutSource = null;
    this.isEditing = false;
    this.inCellEditor = null;

    this.createInCellEditor();
    this.initEvents();
  }

  createInCellEditor() {
    this.inCellEditor = document.createElement('input');
    this.inCellEditor.type = 'text';
    this.inCellEditor.className = 'in-cell-editor';
    this.inCellEditor.id = 'in-cell-editor';
    document.body.appendChild(this.inCellEditor);

    this.inCellEditor.addEventListener('input', () => {
      const sheet = this.app.workbook.getActiveSheet();
      const active = sheet.activeCell || { col: 0, row: 0 };
      const formulaInput = document.getElementById('formula-input');
      if (formulaInput) formulaInput.value = this.inCellEditor.value;
      this.app.autoSaveEngine?.triggerSave();
    });

    this.inCellEditor.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        this.commitCellEdit(e.shiftKey ? -1 : 1, 0);
        e.preventDefault();
      } else if (e.key === 'Tab') {
        this.commitCellEdit(0, e.shiftKey ? -1 : 1);
        e.preventDefault();
      } else if (e.key === 'Escape') {
        this.cancelCellEdit();
        e.preventDefault();
      }
    });

    this.inCellEditor.addEventListener('blur', () => {
      if (this.isEditing) {
        this.commitCellEdit(0, 0);
      }
    });
  }

  initEvents() {
    window.addEventListener('keydown', (e) => this.handleKeyDown(e));
    window.addEventListener('mouseup', () => { this.isSelecting = false; });
    
    window.addEventListener('copy', (e) => this.handleCopyEvent(e));
    window.addEventListener('cut', (e) => this.handleCutEvent(e));
    window.addEventListener('paste', (e) => this.handlePasteEvent(e));

    document.addEventListener('click', (e) => {
      if (!e.target.closest('.context-menu')) {
        document.querySelectorAll('.context-menu').forEach(m => m.classList.remove('active'));
      }
    });

    // Column & Row Drag Resize Event Listeners
    let resizingCol = null;
    let resizingRow = null;
    let startX = 0;
    let startY = 0;
    let startWidth = 0;
    let startHeight = 0;

    document.addEventListener('mousedown', (e) => {
      if (e.target.classList.contains('col-resize-handle')) {
        e.preventDefault();
        e.stopPropagation();
        resizingCol = parseInt(e.target.dataset.col, 10);
        startX = e.clientX;

        const sheet = this.app.workbook.getActiveSheet();
        startWidth = sheet.colWidths && sheet.colWidths[resizingCol] ? sheet.colWidths[resizingCol] : 110;
        document.body.style.cursor = 'col-resize';
      } else if (e.target.classList.contains('row-resize-handle')) {
        e.preventDefault();
        e.stopPropagation();
        resizingRow = parseInt(e.target.dataset.row, 10);
        startY = e.clientY;

        const sheet = this.app.workbook.getActiveSheet();
        startHeight = sheet.rowHeights && sheet.rowHeights[resizingRow] ? sheet.rowHeights[resizingRow] : 24;
        document.body.style.cursor = 'row-resize';
      }
    });

    document.addEventListener('mousemove', (e) => {
      if (resizingCol !== null) {
        const diffX = e.clientX - startX;
        const newW = Math.max(35, startWidth + diffX);
        const sheet = this.app.workbook.getActiveSheet();
        if (!sheet.colWidths) sheet.colWidths = {};
        sheet.colWidths[resizingCol] = newW;

        // Live DOM Column Width Update
        const colHeader = document.querySelector(`.grid-col-header[data-col="${resizingCol}"]`);
        if (colHeader) {
          colHeader.style.width = `${newW}px`;
          colHeader.style.minWidth = `${newW}px`;
        }
        document.querySelectorAll(`.grid-cell[data-col="${resizingCol}"]`).forEach(cell => {
          cell.style.width = `${newW}px`;
          cell.style.minWidth = `${newW}px`;
        });
      } else if (resizingRow !== null) {
        const diffY = e.clientY - startY;
        const newH = Math.max(18, startHeight + diffY);
        const sheet = this.app.workbook.getActiveSheet();
        if (!sheet.rowHeights) sheet.rowHeights = {};
        sheet.rowHeights[resizingRow] = newH;

        // Live DOM Row Height Update
        const tr = document.querySelector(`tr[data-row="${resizingRow}"]`);
        if (tr) {
          tr.style.height = `${newH}px`;
        }
      }
    });

    document.addEventListener('mouseup', () => {
      if (resizingCol !== null || resizingRow !== null) {
        if (resizingCol !== null) {
          this.app.workbook.pushUndoState('Resize Column');
        } else if (resizingRow !== null) {
          this.app.workbook.pushUndoState('Resize Row');
        }
        resizingCol = null;
        resizingRow = null;
        document.body.style.cursor = '';
        this.render();
        this.app.autoSaveEngine?.triggerSave();
      }
    });
  }

  // --- In-Cell Editing Model (MS Excel Style) ---
  startCellEdit(initialChar = null) {
    const sheet = this.app.workbook.getActiveSheet();
    const active = sheet.activeCell || { col: 0, row: 0 };
    const cellEl = document.querySelector(`.grid-cell[data-col="${active.col}"][data-row="${active.row}"]`);
    if (!cellEl) return;

    this.isEditing = true;
    const rect = cellEl.getBoundingClientRect();
    const cell = sheet.getCell(active.col, active.row);
    const existingVal = cell ? (cell.formula || (cell.value !== undefined ? String(cell.value) : '')) : '';

    this.inCellEditor.style.display = 'block';
    this.inCellEditor.style.left = `${rect.left + window.scrollX}px`;
    this.inCellEditor.style.top = `${rect.top + window.scrollY}px`;
    this.inCellEditor.style.width = `${Math.max(rect.width, 100)}px`;
    this.inCellEditor.style.height = `${rect.height}px`;
    this.inCellEditor.style.fontSize = cell?.style?.fontSize ? `${cell.style.fontSize}px` : '12px';
    this.inCellEditor.style.fontWeight = cell?.style?.bold ? 'bold' : 'normal';
    this.inCellEditor.style.textAlign = cell?.style?.align || 'left';

    if (initialChar !== null) {
      this.inCellEditor.value = initialChar;
    } else {
      this.inCellEditor.value = existingVal;
    }

    this.inCellEditor.focus();
    if (initialChar === null) {
      this.inCellEditor.select();
    }
  }

  commitCellEdit(rowDelta = 0, colDelta = 0) {
    if (!this.isEditing) return;
    this.isEditing = false;
    this.inCellEditor.style.display = 'none';

    const sheet = this.app.workbook.getActiveSheet();
    const active = sheet.activeCell || { col: 0, row: 0 };
    const newVal = this.inCellEditor.value;

    this.app.workbook.pushUndoState('Edit Cell');
    sheet.setCellValue(active.col, active.row, newVal);

    // Move active cell
    if (rowDelta !== 0 || colDelta !== 0) {
      const nextR = Math.min(sheet.numRows - 1, Math.max(0, active.row + rowDelta));
      const nextC = Math.min(sheet.numCols - 1, Math.max(0, active.col + colDelta));
      sheet.activeCell = { col: nextC, row: nextR };
      sheet.selection = { startCol: nextC, startRow: nextR, endCol: nextC, endRow: nextR };
    }

    this.render();
    this.updateFormulaBar();
    this.app.autoSaveEngine?.triggerSave();
  }

  cancelCellEdit() {
    this.isEditing = false;
    this.inCellEditor.style.display = 'none';
    this.updateFormulaBar();
  }

  // --- Render DOM Table Grid ---
  render() {
    const sheet = this.app.workbook.getActiveSheet();
    this.container.innerHTML = '';

    const tableWrap = document.createElement('div');
    tableWrap.className = 'grid-table-container';

    const table = document.createElement('table');
    table.className = 'sheet-grid-table';

    // 1. Column Headers (A, B, C...)
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');

    const cornerTh = document.createElement('th');
    cornerTh.className = 'grid-corner-header';
    cornerTh.innerHTML = '⚏';
    cornerTh.onclick = () => {
      sheet.selection = { startCol: 0, startRow: 0, endCol: sheet.numCols - 1, endRow: sheet.numRows - 1 };
      this.updateSelectionVisuals();
      this.updateFormulaBar();
    };
    headerRow.appendChild(cornerTh);

    for (let c = 0; c < sheet.numCols; c++) {
      const th = document.createElement('th');
      th.className = 'grid-col-header';
      th.dataset.col = c;
      th.style.position = 'relative';

      const customW = sheet.colWidths && sheet.colWidths[c];
      th.style.width = customW ? `${customW}px` : '110px';
      th.style.minWidth = customW ? `${customW}px` : '110px';

      const colLetter = FormulaEngine.colToLetter(c);
      th.innerHTML = `
        <div style="display:flex;align-items:center;justify-content:center;position:relative;height:100%;">
          <span>${colLetter}</span>
          ${sheet.autoFilterRange && c >= sheet.autoFilterRange.startCol && c <= sheet.autoFilterRange.endCol ? `
            <span class="col-filter-btn" title="Filter" onclick="event.stopPropagation(); window.zeroCellApp.dataTools.openFilterDropdown(${c}, event.clientX, event.clientY)">▾</span>
          ` : ''}
        </div>
        <div class="col-resize-handle" data-col="${c}"></div>
      `;

      th.onclick = (e) => {
        if (e.target.classList.contains('col-resize-handle')) return;
        sheet.selection = { startCol: c, startRow: 0, endCol: c, endRow: sheet.numRows - 1 };
        sheet.activeCell = { col: c, row: 0 };
        this.updateSelectionVisuals();
        this.updateFormulaBar();
      };

      th.oncontextmenu = (e) => {
        e.preventDefault();
        this.openColContextMenu(c, e.clientX, e.clientY);
      };

      headerRow.appendChild(th);
    }

    const addColTh = document.createElement('th');
    addColTh.className = 'btn-add-grid-col';
    addColTh.title = 'Add Column at End';
    addColTh.innerHTML = '+';
    addColTh.onclick = () => this.addColumn();
    headerRow.appendChild(addColTh);

    thead.appendChild(headerRow);
    table.appendChild(thead);

    // 2. Data Rows
    const tbody = document.createElement('tbody');

    for (let r = 0; r < sheet.numRows; r++) {
      if (sheet.hiddenRows && sheet.hiddenRows.has(r)) continue;

      const tr = document.createElement('tr');
      tr.dataset.row = r;

      if (sheet.rowHeights && sheet.rowHeights[r]) {
        tr.style.height = `${sheet.rowHeights[r]}px`;
      }

      const rowTh = document.createElement('th');
      rowTh.className = 'grid-row-header';
      rowTh.dataset.row = r;
      rowTh.style.position = 'relative';
      rowTh.innerHTML = `
        <span>${r + 1}</span>
        <div class="row-resize-handle" data-row="${r}"></div>
      `;

      rowTh.onclick = (e) => {
        if (e.target.classList.contains('row-resize-handle')) return;
        sheet.selection = { startCol: 0, startRow: r, endCol: sheet.numCols - 1, endRow: r };
        sheet.activeCell = { col: 0, row: r };
        this.updateSelectionVisuals();
        this.updateFormulaBar();
      };

      rowTh.oncontextmenu = (e) => {
        e.preventDefault();
        this.openRowContextMenu(r, e.clientX, e.clientY);
      };

      tr.appendChild(rowTh);

      // Cells
      for (let c = 0; c < sheet.numCols; c++) {
        if (sheet.isMergeCovered(c, r)) {
          continue;
        }

        const td = document.createElement('td');
        td.className = 'grid-cell';
        td.dataset.col = c;
        td.dataset.row = r;

        const merge = sheet.getMergeOrigin(c, r);
        if (merge) {
          td.colSpan = merge.endCol - merge.startCol + 1;
          td.rowSpan = merge.endRow - merge.startRow + 1;
        }

        const cell = sheet.getCell(c, r);
        if (cell) {
          let displayVal = cell.formula ? this.app.formulaEngine.evaluate(cell.formula, sheet.name) : cell.value;
          if (displayVal === null || displayVal === undefined) displayVal = '';

          if (cell.style?.format === 'currency' && !isNaN(displayVal) && displayVal !== '') {
            displayVal = `$${Number(displayVal).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
          } else if (cell.style?.format === 'percent' && !isNaN(displayVal) && displayVal !== '') {
            displayVal = `${(Number(displayVal) * 100).toFixed(2)}%`;
          }

          td.textContent = displayVal;

          if (cell.style) {
            if (cell.style.bold) td.style.fontWeight = 'bold';
            if (cell.style.italic) td.style.fontStyle = 'italic';
            if (cell.style.underline) td.style.textDecoration = 'underline';
            if (cell.style.fontSize) td.style.fontSize = `${cell.style.fontSize}px`;
            if (cell.style.bgColor) td.style.backgroundColor = cell.style.bgColor;
            if (cell.style.color) td.style.color = cell.style.color;
            if (cell.style.align) td.style.textAlign = cell.style.align;
            if (cell.style.wrapText) td.style.whiteSpace = 'normal';
          }
        }

        // Mouse Events
        td.onmousedown = (e) => {
          if (this.isEditing) this.commitCellEdit(0, 0);
          this.isSelecting = true;
          sheet.activeCell = { col: c, row: r };
          if (e.shiftKey && sheet.selection) {
            sheet.selection.endCol = c;
            sheet.selection.endRow = r;
          } else {
            sheet.selection = { startCol: c, startRow: r, endCol: c, endRow: r };
          }
          this.updateSelectionVisuals();
          this.updateFormulaBar();
        };

        td.ondblclick = () => {
          this.startCellEdit();
        };

        td.onmouseenter = () => {
          if (this.isSelecting) {
            sheet.selection.endCol = c;
            sheet.selection.endRow = r;
            this.updateSelectionVisuals();
            this.updateFormulaBar();
          }
        };

        td.oncontextmenu = (e) => {
          e.preventDefault();
          sheet.activeCell = { col: c, row: r };
          this.updateSelectionVisuals();
          this.openCellContextMenu(c, r, e.clientX, e.clientY);
        };

        tr.appendChild(td);
      }

      const dummyTd = document.createElement('td');
      dummyTd.style.border = 'none';
      dummyTd.style.background = 'transparent';
      tr.appendChild(dummyTd);

      tbody.appendChild(tr);
    }

    table.appendChild(tbody);
    tableWrap.appendChild(table);

    // Bottom + Add Row
    const addRowBar = document.createElement('div');
    addRowBar.className = 'add-row-bottom-bar';
    addRowBar.innerHTML = `
      <button class="btn-add-grid-row" id="btn-quick-add-row">
        <span>+</span> Add Row
      </button>
      <button class="btn-add-grid-row" id="btn-quick-add-5-rows" style="font-weight: normal; color: var(--text-secondary);">
        +5 Rows
      </button>
      <button class="btn-add-grid-row" id="btn-quick-add-20-rows" style="font-weight: normal; color: var(--text-secondary);">
        +20 Rows
      </button>
    `;
    tableWrap.appendChild(addRowBar);

    this.container.appendChild(tableWrap);

    document.getElementById('btn-quick-add-row').onclick = () => this.addRow(1);
    document.getElementById('btn-quick-add-5-rows').onclick = () => this.addRow(5);
    document.getElementById('btn-quick-add-20-rows').onclick = () => this.addRow(20);

    this.updateSelectionVisuals();
    this.app.chartEngine.renderCharts();
    this.renderFloatingImages();
  }

  // Update selection highlight & cut marching-ants
  updateSelectionVisuals() {
    const sheet = this.app.workbook.getActiveSheet();
    const sel = sheet.selection;
    if (!sel) return;

    const minC = Math.min(sel.startCol, sel.endCol);
    const maxC = Math.max(sel.startCol, sel.endCol);
    const minR = Math.min(sel.startRow, sel.endRow);
    const maxR = Math.max(sel.startRow, sel.endRow);

    document.querySelectorAll('.grid-cell').forEach(td => {
      const c = parseInt(td.dataset.col, 10);
      const r = parseInt(td.dataset.row, 10);

      const inSel = c >= minC && c <= maxC && r >= minR && r <= maxR;
      const isActive = c === sheet.activeCell.col && r === sheet.activeCell.row;

      let inCut = false;
      if (this.cutSource && this.cutSource.sheetName === sheet.name) {
        inCut = c >= this.cutSource.minC && c <= this.cutSource.maxC && r >= this.cutSource.minR && r <= this.cutSource.maxR;
      }

      td.classList.toggle('selected', inSel);
      td.classList.toggle('active-cell', isActive);
      td.classList.toggle('cut-active', inCut);
    });

    document.querySelectorAll('.grid-col-header').forEach(th => {
      const c = parseInt(th.dataset.col, 10);
      th.classList.toggle('selected', c >= minC && c <= maxC);
    });

    document.querySelectorAll('.grid-row-header').forEach(th => {
      const r = parseInt(th.dataset.row, 10);
      th.classList.toggle('selected', r >= minR && r <= maxR);
    });

    this.updateStatusBarStats(sheet);
  }

  updateFormulaBar() {
    const sheet = this.app.workbook.getActiveSheet();
    const active = sheet.activeCell || { col: 0, row: 0 };

    const nameBox = document.getElementById('name-box');
    const formulaInput = document.getElementById('formula-input');

    if (nameBox) nameBox.value = `${FormulaEngine.colToLetter(active.col)}${active.row + 1}`;

    const cell = sheet.getCell(active.col, active.row);
    if (formulaInput && !this.isEditing) {
      formulaInput.value = cell ? (cell.formula || (cell.value !== undefined ? String(cell.value) : '')) : '';
    }
  }

  // --- CLIPBOARD ACTIONS: Ctrl+C, Ctrl+V, Ctrl+X (MS Excel Style) ---
  copySelection() {
    this.cancelCut();
    const sheet = this.app.workbook.getActiveSheet();
    const sel = sheet.selection;
    if (!sel) return;

    const minC = Math.min(sel.startCol, sel.endCol);
    const maxC = Math.max(sel.startCol, sel.endCol);
    const minR = Math.min(sel.startRow, sel.endRow);
    const maxR = Math.max(sel.startRow, sel.endRow);

    const rowsText = [];
    const richCells = [];

    for (let r = minR; r <= maxR; r++) {
      const rowArr = [];
      const richRow = [];
      for (let c = minC; c <= maxC; c++) {
        const cell = sheet.getCell(c, r);
        const val = cell ? (cell.formula || (cell.value !== undefined ? cell.value : '')) : '';
        rowArr.push(val);
        richRow.push(cell ? JSON.parse(JSON.stringify(cell)) : null);
      }
      rowsText.push(rowArr.join('\t'));
      richCells.push(richRow);
    }

    const tsvString = rowsText.join('\r\n');
    this.internalClipboard = {
      tsv: tsvString,
      rich: richCells,
      isCut: false,
      sourceSheet: sheet.name,
      minC, minR, maxC, maxR
    };

    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(tsvString).catch(() => {});
    }

    this.app.showToast(`Copied ${((maxC - minC + 1) * (maxR - minR + 1))} cell(s) (Ctrl+C)`, 'info');
  }

  startCutSelection() {
    const sheet = this.app.workbook.getActiveSheet();
    const sel = sheet.selection;
    if (!sel) return;

    const minC = Math.min(sel.startCol, sel.endCol);
    const maxC = Math.max(sel.startCol, sel.endCol);
    const minR = Math.min(sel.startRow, sel.endRow);
    const maxR = Math.max(sel.startRow, sel.endRow);

    const isFullRow = (minC === 0 && maxC >= sheet.numCols - 1);
    const isFullCol = (minR === 0 && maxR >= sheet.numRows - 1);

    const rowsText = [];
    const richCells = [];

    for (let r = minR; r <= maxR; r++) {
      const rowArr = [];
      const richRow = [];
      for (let c = minC; c <= maxC; c++) {
        const cell = sheet.getCell(c, r);
        const val = cell ? (cell.formula || (cell.value !== undefined ? cell.value : '')) : '';
        rowArr.push(val);
        richRow.push(cell ? JSON.parse(JSON.stringify(cell)) : null);
      }
      rowsText.push(rowArr.join('\t'));
      richCells.push(richRow);
    }

    const tsvString = rowsText.join('\r\n');
    this.internalClipboard = {
      tsv: tsvString,
      rich: richCells,
      isCut: true,
      isFullRow,
      isFullCol,
      sourceSheet: sheet.name,
      minC, minR, maxC, maxR
    };

    this.cutSource = {
      sheetName: sheet.name,
      minC, minR, maxC, maxR,
      isFullRow,
      isFullCol,
      richCells,
      tsvString
    };

    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(tsvString).catch(() => {});
    }

    this.updateSelectionVisuals();
    const label = isFullRow ? `Row(s) ${minR + 1}${maxR > minR ? `-${maxR + 1}` : ''}` : `${((maxC - minC + 1) * (maxR - minR + 1))} cell(s)`;
    this.app.showToast(`✂️ Cut ${label}. Select destination & press Ctrl+V, Enter, or Right-Click > Insert Cut Rows`, 'info');
  }

  cancelCut() {
    if (this.cutSource) {
      this.cutSource = null;
      this.updateSelectionVisuals();
    }
  }

  // --- INSERT CUT CELLS / INSERT COPIED CELLS (Shift Down) ---
  insertCutCells(insertMode = true) {
    if (!this.cutSource && !this.internalClipboard) {
      this.app.showToast('No cut or copied cells in clipboard', 'warning');
      return;
    }

    const targetSheet = this.app.workbook.getActiveSheet();
    const active = targetSheet.activeCell || { col: 0, row: 0 };
    const targetRow = active.row;
    const targetCol = active.col;

    // 1. MS Excel ROW MOVE (e.g. Cut Row 65 and Insert at Row 50 without any data loss)
    if (this.cutSource && this.cutSource.isFullRow && this.cutSource.sheetName === targetSheet.name) {
      this.app.workbook.pushUndoState('Move Cut Row(s)');
      const res = targetSheet.moveRowRange(this.cutSource.minR, this.cutSource.maxR, targetRow);
      this.cutSource = null;
      
      if (res) {
        targetSheet.selection = {
          startCol: 0,
          startRow: res.destStartRow,
          endCol: targetSheet.numCols - 1,
          endRow: res.destEndRow
        };
        targetSheet.activeCell = { col: 0, row: res.destStartRow };
      }
      this.render();
      this.updateFormulaBar();
      this.app.autoSaveEngine?.triggerSave();
      this.app.showToast(`Moved Row to Row ${targetRow + 1} with zero data loss!`, 'success');
      return;
    }

    // 2. Regular Cut / Copied cells insertion
    const sourceData = this.cutSource || this.internalClipboard;
    const numRows = sourceData.rich.length;
    const numCols = sourceData.rich[0] ? sourceData.rich[0].length : 1;

    this.app.workbook.pushUndoState('Insert Cut Cells');

    if (insertMode) {
      for (let i = 0; i < numRows; i++) {
        targetSheet.insertRow(targetRow);
      }
    }

    sourceData.rich.forEach((rowCells, rOffset) => {
      rowCells.forEach((cellData, cOffset) => {
        const destC = targetCol + cOffset;
        const destR = targetRow + rOffset;
        if (cellData) {
          const targetCell = targetSheet.getOrCreateCell(destC, destR);
          targetCell.value = cellData.value;
          targetCell.formula = cellData.formula;
          if (cellData.style) targetCell.style = JSON.parse(JSON.stringify(cellData.style));
        }
      });
    });

    if (this.cutSource) {
      const sourceSheet = this.app.workbook.getSheet(this.cutSource.sheetName) || targetSheet;
      let origMinR = this.cutSource.minR;
      if (sourceSheet === targetSheet && origMinR >= targetRow) {
        origMinR += numRows;
      }
      for (let i = 0; i < numRows; i++) {
        sourceSheet.deleteRow(origMinR);
      }
      this.cutSource = null;
    }

    targetSheet.selection = {
      startCol: targetCol,
      startRow: targetRow,
      endCol: targetCol + numCols - 1,
      endRow: targetRow + numRows - 1
    };

    this.render();
    this.updateFormulaBar();
    this.app.autoSaveEngine?.triggerSave();
    this.app.showToast(`Inserted ${numRows} row(s) successfully!`, 'success');
  }

  async pasteToSelection(clipboardText = null) {
    try {
      const targetSheet = this.app.workbook.getActiveSheet();
      const active = targetSheet.activeCell || { col: 0, row: 0 };
      const startCol = active.col;
      const startRow = active.row;

      // 1. MS Excel-style CUT & MOVE execution
      if (this.cutSource) {
        // If it was a full row cut and target is row header or cell, execute flawless row move
        if (this.cutSource.isFullRow && this.cutSource.sheetName === targetSheet.name) {
          this.insertCutCells(true);
          return;
        }

        const cut = this.cutSource;
        const sourceSheet = this.app.workbook.getSheet(cut.sheetName) || targetSheet;

        this.app.workbook.pushUndoState('Move Cells (Ctrl+X)');

        for (let r = cut.minR; r <= cut.maxR; r++) {
          for (let c = cut.minC; c <= cut.maxC; c++) {
            sourceSheet.deleteCell(c, r);
          }
        }

        const numRows = cut.richCells.length;
        let numCols = 0;

        cut.richCells.forEach((rowCells, rOffset) => {
          if (rowCells.length > numCols) numCols = rowCells.length;
          rowCells.forEach((cellData, cOffset) => {
            const targetC = startCol + cOffset;
            const targetR = startRow + rOffset;
            if (cellData) {
              const targetCell = targetSheet.getOrCreateCell(targetC, targetR);
              targetCell.value = cellData.value;
              targetCell.formula = cellData.formula;
              if (cellData.style) targetCell.style = JSON.parse(JSON.stringify(cellData.style));
            }
          });
        });

        this.cutSource = null;

        targetSheet.selection = {
          startCol,
          startRow,
          endCol: startCol + Math.max(0, numCols - 1),
          endRow: startRow + Math.max(0, numRows - 1)
        };

        this.render();
        this.updateFormulaBar();
        this.app.showToast(`Moved ${numRows * numCols} cell(s) successfully!`, 'success');
        return;
      }

      // 2. Regular Clipboard Paste
      let textToPaste = clipboardText;
      if (!textToPaste && navigator.clipboard && navigator.clipboard.readText) {
        try {
          textToPaste = await navigator.clipboard.readText();
        } catch (e) {
          textToPaste = this.internalClipboard ? this.internalClipboard.tsv : '';
        }
      } else if (!textToPaste && this.internalClipboard) {
        textToPaste = this.internalClipboard.tsv;
      }

      if (!textToPaste) {
        this.app.showToast('Clipboard is empty', 'warning');
        return;
      }

      this.app.workbook.pushUndoState('Paste (Ctrl+V)');

      const lines = textToPaste.split(/\r\n|\n|\r/);
      let pastedRows = 0;
      let pastedCols = 0;

      lines.forEach((line, rIdx) => {
        if (!line && rIdx === lines.length - 1) return;
        pastedRows++;
        const cols = line.split('\t');
        if (cols.length > pastedCols) pastedCols = cols.length;

        cols.forEach((colVal, cIdx) => {
          const targetC = startCol + cIdx;
          const targetR = startRow + rIdx;

          if (this.internalClipboard && this.internalClipboard.rich && this.internalClipboard.rich[rIdx] && this.internalClipboard.rich[rIdx][cIdx]) {
            const richCell = this.internalClipboard.rich[rIdx][cIdx];
            const targetCell = targetSheet.getOrCreateCell(targetC, targetR);
            targetCell.value = richCell.value;
            targetCell.formula = richCell.formula;
            if (richCell.style) targetCell.style = JSON.parse(JSON.stringify(richCell.style));
          } else {
            targetSheet.setCellValue(targetC, targetR, colVal);
          }
        });
      });

      targetSheet.selection = {
        startCol,
        startRow,
        endCol: startCol + Math.max(0, pastedCols - 1),
        endRow: startRow + Math.max(0, pastedRows - 1)
      };

      this.render();
      this.updateFormulaBar();
      this.app.autoSaveEngine?.triggerSave();
      this.app.showToast(`Pasted ${pastedRows * pastedCols} cell(s) (Ctrl+V)`, 'success');
    } catch (err) {
      console.warn('Paste error:', err);
      this.app.showToast('Paste error: ' + err.message, 'warning');
    }
  }

  handleCopyEvent(e) {
    if (this.isEditing || (document.activeElement && document.activeElement.tagName === 'INPUT')) return;
    this.copySelection();
    if (e.clipboardData && this.internalClipboard) {
      e.clipboardData.setData('text/plain', this.internalClipboard.tsv);
      e.preventDefault();
    }
  }

  handleCutEvent(e) {
    if (this.isEditing || (document.activeElement && document.activeElement.tagName === 'INPUT')) return;
    this.startCutSelection();
    if (e.clipboardData && this.internalClipboard) {
      e.clipboardData.setData('text/plain', this.internalClipboard.tsv);
      e.preventDefault();
    }
  }

  handlePasteEvent(e) {
    if (this.isEditing || (document.activeElement && document.activeElement.tagName === 'INPUT')) return;
    if (e.clipboardData) {
      const items = e.clipboardData.items;
      if (items) {
        for (let i = 0; i < items.length; i++) {
          if (items[i].type.indexOf('image') !== -1) {
            const blob = items[i].getAsFile();
            const reader = new FileReader();
            reader.onload = (evt) => {
              this.insertFloatingImage(evt.target.result);
            };
            reader.readAsDataURL(blob);
            e.preventDefault();
            return;
          }
        }
      }

      const text = e.clipboardData.getData('text/plain');
      if (text) {
        e.preventDefault();
        this.pasteToSelection(text);
      }
    }
  }

  // --- FLOATING IMAGES SUPPORT (OpenXML / Clipboard / Local) ---
  insertFloatingImage(src, options = {}) {
    const sheet = this.app.workbook.getActiveSheet();
    const active = sheet.activeCell || { col: 0, row: 0 };
    const imgId = `img_${Date.now()}_${Math.floor(Math.random() * 1000)}`;

    const imgObj = {
      id: imgId,
      src: src,
      col: options.col !== undefined ? options.col : active.col,
      row: options.row !== undefined ? options.row : active.row,
      x: options.x !== undefined ? options.x : 60 + ((sheet.images ? sheet.images.length : 0) * 30),
      y: options.y !== undefined ? options.y : 60 + ((sheet.images ? sheet.images.length : 0) * 30),
      width: options.width || 320,
      height: options.height || 200
    };

    sheet.images = sheet.images || [];
    sheet.images.push(imgObj);
    this.renderFloatingImages();
    this.app.autoSaveEngine?.triggerSave();
  }

  renderFloatingImages() {
    const container = document.getElementById('charts-layer');
    if (!container) return;

    const existingImages = container.querySelectorAll('.excel-floating-image, .floating-sheet-image');
    existingImages.forEach(el => el.remove());

    const sheet = this.app.workbook.getActiveSheet();
    if (!sheet.images || sheet.images.length === 0) return;

    sheet.images.forEach(img => {
      const imgWrap = document.createElement('div');
      imgWrap.className = 'excel-floating-image';
      imgWrap.id = img.id;
      imgWrap.style.left = `${img.x}px`;
      imgWrap.style.top = `${img.y}px`;
      imgWrap.style.width = `${img.width}px`;
      imgWrap.style.height = `${img.height}px`;

      imgWrap.innerHTML = `
        <img src="${img.src}" alt="Sheet Image" />
        <div class="excel-img-handle nw" data-dir="nw"></div>
        <div class="excel-img-handle ne" data-dir="ne"></div>
        <div class="excel-img-handle se" data-dir="se"></div>
        <div class="excel-img-handle sw" data-dir="sw"></div>
        <div class="excel-img-handle n" data-dir="n"></div>
        <div class="excel-img-handle s" data-dir="s"></div>
        <div class="excel-img-handle w" data-dir="w"></div>
        <div class="excel-img-handle e" data-dir="e"></div>
      `;

      // Select and Drag
      imgWrap.addEventListener('mousedown', (e) => {
        if (e.target.classList.contains('excel-img-handle')) return;

        e.stopPropagation();
        container.querySelectorAll('.excel-floating-image.selected').forEach(el => el.classList.remove('selected'));
        imgWrap.classList.add('selected');

        const startX = e.clientX - imgWrap.offsetLeft;
        const startY = e.clientY - imgWrap.offsetTop;

        const onMouseMove = (moveEvent) => {
          const newX = Math.max(0, moveEvent.clientX - startX);
          const newY = Math.max(0, moveEvent.clientY - startY);
          imgWrap.style.left = `${newX}px`;
          imgWrap.style.top = `${newY}px`;
          img.x = newX;
          img.y = newY;
        };

        const onMouseUp = () => {
          document.removeEventListener('mousemove', onMouseMove);
          document.removeEventListener('mouseup', onMouseUp);
          this.app.autoSaveEngine?.triggerSave();
        };

        document.addEventListener('mousemove', onMouseMove);
        document.addEventListener('mouseup', onMouseUp);
      });

      // 8 Resize Handles
      imgWrap.querySelectorAll('.excel-img-handle').forEach(handle => {
        handle.addEventListener('mousedown', (e) => {
          e.stopPropagation();
          e.preventDefault();
          const dir = handle.dataset.dir;
          const startX = e.clientX;
          const startY = e.clientY;
          const startW = imgWrap.offsetWidth;
          const startH = imgWrap.offsetHeight;
          const startL = imgWrap.offsetLeft;
          const startT = imgWrap.offsetTop;

          const onResizeMove = (moveEvt) => {
            const dx = moveEvt.clientX - startX;
            const dy = moveEvt.clientY - startY;

            if (dir.includes('e')) {
              const newW = Math.max(20, startW + dx);
              imgWrap.style.width = `${newW}px`;
              img.width = newW;
            }
            if (dir.includes('s')) {
              const newH = Math.max(20, startH + dy);
              imgWrap.style.height = `${newH}px`;
              img.height = newH;
            }
            if (dir.includes('w')) {
              const newW = Math.max(20, startW - dx);
              const newL = startL + (startW - newW);
              imgWrap.style.width = `${newW}px`;
              imgWrap.style.left = `${newL}px`;
              img.width = newW;
              img.x = newL;
            }
            if (dir.includes('n')) {
              const newH = Math.max(20, startH - dy);
              const newT = startT + (startH - newH);
              imgWrap.style.height = `${newH}px`;
              imgWrap.style.top = `${newT}px`;
              img.height = newH;
              img.y = newT;
            }
          };

          const onResizeUp = () => {
            document.removeEventListener('mousemove', onResizeMove);
            document.removeEventListener('mouseup', onResizeUp);
            this.app.autoSaveEngine?.triggerSave();
          };

          document.addEventListener('mousemove', onResizeMove);
          document.addEventListener('mouseup', onResizeUp);
        });
      });

      container.appendChild(imgWrap);
    });

    // Deselect on click outside
    if (!this._imageGlobalListenersBound) {
      this._imageGlobalListenersBound = true;
      document.addEventListener('mousedown', (e) => {
        if (!e.target.closest('.excel-floating-image')) {
          document.querySelectorAll('.excel-floating-image.selected').forEach(el => el.classList.remove('selected'));
        }
      });

      // Delete key removes selected image
      document.addEventListener('keydown', (e) => {
        if (e.key === 'Delete' || e.key === 'Backspace') {
          const selectedEl = document.querySelector('.excel-floating-image.selected');
          if (selectedEl && !this.isEditing && !(document.activeElement && document.activeElement.tagName === 'INPUT')) {
            const imgId = selectedEl.id;
            const curSheet = this.app.workbook.getActiveSheet();
            curSheet.images = curSheet.images.filter(i => i.id !== imgId);
            selectedEl.remove();
            this.app.autoSaveEngine?.triggerSave();
            e.preventDefault();
          }
        }
      });
    }
  }

  // --- KEYBOARD NAVIGATION & INSTANT TYPING ---
  handleKeyDown(e) {
    if (document.getElementById('modal-overlay').classList.contains('active')) return;
    
    // Formula input box handling
    if (e.target === document.getElementById('formula-input')) {
      if (e.key === 'Enter') {
        const sheet = this.app.workbook.getActiveSheet();
        const active = sheet.activeCell || { col: 0, row: 0 };
        sheet.setCellValue(active.col, active.row, e.target.value);
        this.render();
        this.app.autoSaveEngine?.triggerSave();
      }
      return;
    }

    if (this.isEditing) return; // In-cell editor has focus

    // Escape cancels Cut mode
    if (e.key === 'Escape') {
      if (this.cutSource) {
        this.cancelCut();
        this.app.showToast('Cut cancelled', 'info');
        e.preventDefault();
        return;
      }
    }

    // Enter confirms Cut & Move or starts editing
    if (e.key === 'Enter') {
      if (this.cutSource) {
        this.pasteToSelection();
      } else {
        this.startCellEdit();
      }
      e.preventDefault();
      return;
    }

    // F2 enters edit mode
    if (e.key === 'F2') {
      this.startCellEdit();
      e.preventDefault();
      return;
    }

    // Ctrl / Cmd Shortcuts
    if (e.ctrlKey || e.metaKey) {
      const key = e.key.toLowerCase();
      const active = this.app.workbook.getActiveSheet().activeCell || { col: 0, row: 0 };
      const sheet = this.app.workbook.getActiveSheet();

      // Ctrl + Minus: Delete Row or Column (Ctrl -)
      if (key === '-' || key === '_' || e.code === 'NumpadSubtract' || e.code === 'Minus') {
        const sel = sheet.selection;
        if (sel && sel.startCol === 0 && sel.endCol >= sheet.numCols - 1) {
          // Delete selected rows
          const minR = Math.min(sel.startRow, sel.endRow);
          const maxR = Math.max(sel.startRow, sel.endRow);
          const count = maxR - minR + 1;
          this.app.workbook.pushUndoState('Delete Row(s)');
          for (let i = 0; i < count; i++) {
            sheet.deleteRow(minR);
          }
          sheet.activeCell = { col: 0, row: Math.min(sheet.numRows - 1, minR) };
          sheet.selection = { startCol: 0, startRow: sheet.activeCell.row, endCol: sheet.numCols - 1, endRow: sheet.activeCell.row };
          this.render();
          this.app.autoSaveEngine?.triggerSave();
          this.app.showToast(`🗑️ Deleted ${count} Row(s) (Ctrl-)`, 'info');
        } else if (sel && sel.startRow === 0 && sel.endRow >= sheet.numRows - 1) {
          // Delete selected columns
          const minC = Math.min(sel.startCol, sel.endCol);
          const maxC = Math.max(sel.startCol, sel.endCol);
          const count = maxC - minC + 1;
          this.app.workbook.pushUndoState('Delete Column(s)');
          for (let i = 0; i < count; i++) {
            sheet.deleteColumn(minC);
          }
          sheet.activeCell = { col: Math.min(sheet.numCols - 1, minC), row: 0 };
          sheet.selection = { startCol: sheet.activeCell.col, startRow: 0, endCol: sheet.activeCell.col, endRow: sheet.numRows - 1 };
          this.render();
          this.app.autoSaveEngine?.triggerSave();
          this.app.showToast(`🗑️ Deleted ${count} Column(s) (Ctrl-)`, 'info');
        } else {
          this.deleteRow(active.row);
        }
        e.preventDefault();
        return;
      }

      // Ctrl + Plus / Equal: Insert Cut Cells or Insert Row (Ctrl +)
      if (key === '+' || key === '=' || e.code === 'NumpadAdd' || e.code === 'Equal') {
        if (this.cutSource) {
          this.insertCutCells(true);
        } else {
          const sel = sheet.selection;
          if (sel && sel.startRow === 0 && sel.endRow >= sheet.numRows - 1) {
            this.addColumn(1, active.col);
          } else {
            this.addRow(1, active.row);
          }
        }
        e.preventDefault();
        return;
      }

      if (key === 'c') {
        this.copySelection();
        e.preventDefault();
        return;
      }
      if (key === 'x') {
        this.startCutSelection();
        e.preventDefault();
        return;
      }
      if (key === 'v') {
        this.pasteToSelection();
        e.preventDefault();
        return;
      }
      if (key === 'z') {
        this.app.workbook.undo();
        this.render();
        this.updateFormulaBar();
        this.app.autoSaveEngine?.triggerSave();
        e.preventDefault();
        return;
      }
      if (key === 'y') {
        this.app.workbook.redo();
        this.render();
        this.updateFormulaBar();
        this.app.autoSaveEngine?.triggerSave();
        e.preventDefault();
        return;
      }
      if (key === 'b') {
        this.app.toggleStyleProperty('bold');
        e.preventDefault();
        return;
      }
      if (key === 'i') {
        this.app.toggleStyleProperty('italic');
        e.preventDefault();
        return;
      }
      if (key === 'u') {
        this.app.toggleStyleProperty('underline');
        e.preventDefault();
        return;
      }
      if (key === 's') {
        this.app.autoSaveEngine?.saveInstant();
        this.app.showToast('Spreadsheet saved instantly!', 'success');
        e.preventDefault();
        return;
      }
      if (key === 'f') {
        this.app.dataTools.openFindReplaceDialog('find');
        e.preventDefault();
        return;
      }
      if (key === 'a') {
        const sheet = this.app.workbook.getActiveSheet();
        sheet.selection = { startCol: 0, startRow: 0, endCol: sheet.numCols - 1, endRow: sheet.numRows - 1 };
        this.updateSelectionVisuals();
        e.preventDefault();
        return;
      }
    }

    // Direct Instant Typing: Any printable key starts typing immediately!
    if (!e.ctrlKey && !e.metaKey && !e.altKey && e.key.length === 1) {
      this.startCellEdit(e.key);
      e.preventDefault();
      return;
    }

    // Navigation & Editing keys
    const sheet = this.app.workbook.getActiveSheet();
    const active = sheet.activeCell || { col: 0, row: 0 };

    if (e.key === 'Delete' || e.key === 'Backspace') {
      this.clearSelectionValues();
      e.preventDefault();
    } else if (e.key === 'Tab') {
      const nextC = Math.min(sheet.numCols - 1, Math.max(0, active.col + (e.shiftKey ? -1 : 1)));
      sheet.activeCell = { col: nextC, row: active.row };
      sheet.selection = { startCol: nextC, startRow: active.row, endCol: nextC, endRow: active.row };
      this.updateSelectionVisuals();
      this.updateFormulaBar();
      e.preventDefault();
    } else if (e.key === 'ArrowUp') {
      sheet.activeCell = { col: active.col, row: Math.max(0, active.row - 1) };
      sheet.selection = { ...sheet.activeCell, endCol: sheet.activeCell.col, endRow: sheet.activeCell.row };
      this.updateSelectionVisuals();
      this.updateFormulaBar();
      e.preventDefault();
    } else if (e.key === 'ArrowDown') {
      sheet.activeCell = { col: active.col, row: Math.min(sheet.numRows - 1, active.row + 1) };
      sheet.selection = { ...sheet.activeCell, endCol: sheet.activeCell.col, endRow: sheet.activeCell.row };
      this.updateSelectionVisuals();
      this.updateFormulaBar();
      e.preventDefault();
    } else if (e.key === 'ArrowLeft') {
      sheet.activeCell = { col: Math.max(0, active.col - 1), row: active.row };
      sheet.selection = { ...sheet.activeCell, endCol: sheet.activeCell.col, endRow: sheet.activeCell.row };
      this.updateSelectionVisuals();
      this.updateFormulaBar();
      e.preventDefault();
    } else if (e.key === 'ArrowRight') {
      sheet.activeCell = { col: Math.min(sheet.numCols - 1, active.col + 1), row: active.row };
      sheet.selection = { ...sheet.activeCell, endCol: sheet.activeCell.col, endRow: sheet.activeCell.row };
      this.updateSelectionVisuals();
      this.updateFormulaBar();
      e.preventDefault();
    }
  }

  // --- ROW & COLUMN OPERATIONS (+ and -) ---
  addRow(count = 1, atIndex = null) {
    const sheet = this.app.workbook.getActiveSheet();
    this.app.workbook.pushUndoState('Add Row');

    if (atIndex !== null) {
      for (let i = 0; i < count; i++) {
        sheet.insertRow(atIndex);
      }
      this.app.showToast(`Inserted ${count} row(s) at row ${atIndex + 1}`, 'success');
    } else {
      sheet.numRows += count;
      this.app.showToast(`Added ${count} row(s) at bottom`, 'success');
    }
    this.render();
    this.app.autoSaveEngine?.triggerSave();
  }

  deleteRow(atIndex = null) {
    const sheet = this.app.workbook.getActiveSheet();
    const targetRow = atIndex !== null ? atIndex : (sheet.activeCell ? sheet.activeCell.row : 0);

    if (sheet.numRows <= 1) {
      this.app.showToast('Cannot delete all rows', 'warning');
      return;
    }

    this.app.workbook.pushUndoState('Delete Row');
    sheet.deleteRow(targetRow);
    this.app.showToast(`Deleted Row ${targetRow + 1}`, 'info');
    this.render();
    this.app.autoSaveEngine?.triggerSave();
  }

  addColumn(count = 1, atIndex = null) {
    const sheet = this.app.workbook.getActiveSheet();
    this.app.workbook.pushUndoState('Add Column');

    if (atIndex !== null) {
      for (let i = 0; i < count; i++) {
        sheet.insertColumn(atIndex);
      }
      this.app.showToast(`Inserted Column at ${FormulaEngine.colToLetter(atIndex)}`, 'success');
    } else {
      sheet.numCols += count;
      this.app.showToast(`Added column at right (Col ${FormulaEngine.colToLetter(sheet.numCols - 1)})`, 'success');
    }
    this.render();
    this.app.autoSaveEngine?.triggerSave();
  }

  deleteColumn(atIndex = null) {
    const sheet = this.app.workbook.getActiveSheet();
    const targetCol = atIndex !== null ? atIndex : (sheet.activeCell ? sheet.activeCell.col : 0);

    if (sheet.numCols <= 1) {
      this.app.showToast('Cannot delete all columns', 'warning');
      return;
    }

    this.app.workbook.pushUndoState('Delete Column');
    sheet.deleteColumn(targetCol);
    this.app.showToast(`Deleted Column ${FormulaEngine.colToLetter(targetCol)}`, 'info');
    this.render();
    this.app.autoSaveEngine?.triggerSave();
  }

  // --- Merge & Center ---
  toggleMerge() {
    const sheet = this.app.workbook.getActiveSheet();
    const sel = sheet.selection;
    if (!sel || (sel.startCol === sel.endCol && sel.startRow === sel.endRow)) {
      this.app.showToast('Select 2 or more cells to merge', 'warning');
      return;
    }

    this.app.workbook.pushUndoState('Merge Cells');
    const isAlreadyMerged = sheet.getMergeOrigin(Math.min(sel.startCol, sel.endCol), Math.min(sel.startRow, sel.endRow));

    if (isAlreadyMerged) {
      sheet.unmergeSelection();
      this.app.showToast('Unmerged cells', 'info');
    } else {
      sheet.mergeSelection();
      const originCell = sheet.getOrCreateCell(Math.min(sel.startCol, sel.endCol), Math.min(sel.startRow, sel.endRow));
      originCell.style = originCell.style || {};
      originCell.style.align = 'center';
      this.app.showToast('Merged & Centered selected cells', 'success');
    }
    this.render();
    this.app.autoSaveEngine?.triggerSave();
  }

  // --- Context Menus ---
  openCellContextMenu(col, row, x, y) {
    const menu = document.getElementById('grid-context-menu');
    const hasCutOrCopy = !!(this.cutSource || this.internalClipboard);

    menu.innerHTML = `
      <div class="menu-item" onclick="window.zeroCellApp.gridEngine.copySelection()">
        📋 Copy (Ctrl+C)
      </div>
      <div class="menu-item" onclick="window.zeroCellApp.gridEngine.startCutSelection()">
        ✂️ Cut (Ctrl+X)
      </div>
      <div class="menu-item" onclick="window.zeroCellApp.gridEngine.pasteToSelection()">
        📥 Paste / Move (Ctrl+V)
      </div>
      ${hasCutOrCopy ? `
        <div class="menu-item" style="color:var(--brand-primary);font-weight:600;" onclick="window.zeroCellApp.gridEngine.insertCutCells(true)">
          📥 Insert Cut / Copied Cells (Shift Down)
        </div>
      ` : ''}
      <div class="menu-divider"></div>
      <div class="menu-item" onclick="window.zeroCellApp.gridEngine.addRow(1, ${row})">
        ➕ Insert Row Above
      </div>
      <div class="menu-item" onclick="window.zeroCellApp.gridEngine.addRow(1, ${row + 1})">
        ➕ Insert Row Below
      </div>
      <div class="menu-item danger" onclick="window.zeroCellApp.gridEngine.deleteRow(${row})">
        🗑️ Delete Row (${row + 1})
      </div>
      <div class="menu-divider"></div>
      <div class="menu-item" onclick="window.zeroCellApp.gridEngine.addColumn(1, ${col})">
        ➕ Insert Column Left
      </div>
      <div class="menu-item" onclick="window.zeroCellApp.gridEngine.addColumn(1, ${col + 1})">
        ➕ Insert Column Right
      </div>
      <div class="menu-item danger" onclick="window.zeroCellApp.gridEngine.deleteColumn(${col})">
        🗑️ Delete Column (${FormulaEngine.colToLetter(col)})
      </div>
      <div class="menu-divider"></div>
      <div class="menu-item" onclick="window.zeroCellApp.gridEngine.toggleMerge()">
        ⊞ Merge & Center / Unmerge
      </div>
      <div class="menu-item" onclick="window.zeroCellApp.duplicateTools.highlightDuplicates()">
        ✨ Highlight Duplicates
      </div>
      <div class="menu-item" onclick="window.zeroCellApp.duplicateTools.openRemoveDuplicatesDialog()">
        🧹 Remove Duplicates...
      </div>
      <div class="menu-divider"></div>
      <div class="menu-item" onclick="window.zeroCellApp.gridEngine.clearSelectionValues()">
        Clear Contents (Del)
      </div>
    `;
    this.positionMenu(menu, x, y);
  }

  openRowContextMenu(row, x, y) {
    const menu = document.getElementById('grid-context-menu');
    const hasCutOrCopy = !!(this.cutSource || this.internalClipboard);

    menu.innerHTML = `
      <div class="menu-item" onclick="window.zeroCellApp.gridEngine.startCutSelection()">
        ✂️ Cut Row (${row + 1})
      </div>
      <div class="menu-item" onclick="window.zeroCellApp.gridEngine.copySelection()">
        📋 Copy Row (${row + 1})
      </div>
      ${hasCutOrCopy ? `
        <div class="menu-item" style="color:var(--brand-primary);font-weight:600;" onclick="window.zeroCellApp.gridEngine.insertCutCells(true)">
          📥 Insert Cut Rows Here (Shift Down)
        </div>
      ` : ''}
      <div class="menu-divider"></div>
      <div class="menu-item" onclick="window.zeroCellApp.gridEngine.addRow(1, ${row})">
        ➕ Insert Row Above
      </div>
      <div class="menu-item" onclick="window.zeroCellApp.gridEngine.addRow(1, ${row + 1})">
        ➕ Insert Row Below
      </div>
      <div class="menu-item danger" onclick="window.zeroCellApp.gridEngine.deleteRow(${row})">
        🗑️ Delete Row (${row + 1})
      </div>
    `;
    this.positionMenu(menu, x, y);
  }

  openColContextMenu(col, x, y) {
    const menu = document.getElementById('grid-context-menu');
    menu.innerHTML = `
      <div class="menu-item" onclick="window.zeroCellApp.gridEngine.addColumn(1, ${col})">
        ➕ Insert Column Left
      </div>
      <div class="menu-item" onclick="window.zeroCellApp.gridEngine.addColumn(1, ${col + 1})">
        ➕ Insert Column Right
      </div>
      <div class="menu-item danger" onclick="window.zeroCellApp.gridEngine.deleteColumn(${col})">
        🗑️ Delete Column (${FormulaEngine.colToLetter(col)})
      </div>
      <div class="menu-divider"></div>
      <div class="menu-item" onclick="window.zeroCellApp.dataTools.sortColumn(${col}, true)">
        Sort A to Z
      </div>
      <div class="menu-item" onclick="window.zeroCellApp.dataTools.sortColumn(${col}, false)">
        Sort Z to A
      </div>
    `;
    this.positionMenu(menu, x, y);
  }

  positionMenu(menu, x, y) {
    menu.style.left = `${Math.min(window.innerWidth - 240, x)}px`;
    menu.style.top = `${Math.min(window.innerHeight - 280, y)}px`;
    menu.classList.add('active');
  }

  clearSelectionValues() {
    const sheet = this.app.workbook.getActiveSheet();
    const sel = sheet.selection;
    if (!sel) return;

    this.app.workbook.pushUndoState('Clear Cells');
    const minC = Math.min(sel.startCol, sel.endCol);
    const maxC = Math.max(sel.startCol, sel.endCol);
    const minR = Math.min(sel.startRow, sel.endRow);
    const maxR = Math.max(sel.startRow, sel.endRow);

    for (let r = minR; r <= maxR; r++) {
      for (let c = minC; c <= maxC; c++) {
        sheet.deleteCell(c, r);
      }
    }
    this.render();
    this.updateFormulaBar();
  }

  updateStatusBarStats(sheet) {
    const sel = sheet.selection;
    if (!sel) return;

    const minC = Math.min(sel.startCol, sel.endCol);
    const maxC = Math.max(sel.startCol, sel.endCol);
    const minR = Math.min(sel.startRow, sel.endRow);
    const maxR = Math.max(sel.startRow, sel.endRow);

    let count = 0, numCount = 0, sum = 0, min = Infinity, max = -Infinity;

    for (let r = minR; r <= maxR; r++) {
      for (let c = minC; c <= maxC; c++) {
        const cell = sheet.getCell(c, r);
        if (cell && cell.value !== undefined && cell.value !== '') {
          count++;
          const val = this.app.formulaEngine.evaluate(cell.formula || cell.value, sheet.name);
          if (!isNaN(val) && typeof val === 'number') {
            numCount++;
            sum += val;
            if (val < min) min = val;
            if (val > max) max = val;
          }
        }
      }
    }

    const avg = numCount > 0 ? (sum / numCount).toFixed(2) : 0;
    const statsContainer = document.getElementById('status-stats');
    if (statsContainer) {
      if (count <= 1 && numCount === 0) {
        statsContainer.innerHTML = `<span class="stat-pill">Ready</span>`;
      } else {
        statsContainer.innerHTML = `
          ${numCount > 0 ? `<span class="stat-pill">Avg: <strong>${avg}</strong></span>` : ''}
          <span class="stat-pill">Count: <strong>${count}</strong></span>
          ${numCount > 0 ? `<span class="stat-pill">Min: <strong>${min}</strong></span>` : ''}
          ${numCount > 0 ? `<span class="stat-pill">Max: <strong>${max}</strong></span>` : ''}
          ${numCount > 0 ? `<span class="stat-pill">Sum: <strong>${sum.toLocaleString()}</strong></span>` : ''}
        `;
      }
    }
  }
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = { WorkbookManager, Worksheet, GridEngine };
}
