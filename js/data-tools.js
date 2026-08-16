/**
 * 0cell Advanced Data Tools
 * AutoFilter, Multi-column Sort, Find & Replace, Text-to-Columns, Data Validation
 */

class DataTools {
  constructor(app) {
    this.app = app;
    this.activeFilterMenuCol = null;
    this.filterCriteria = new Map(); // col -> Set of allowed values
  }

  // Toggle AutoFilter on selection
  toggleAutoFilter() {
    const sheet = this.app.workbook.getActiveSheet();
    let sel = sheet.selection;
    if (!sel) return;

    if (sheet.autoFilterRange) {
      sheet.autoFilterRange = null;
      sheet.hiddenRows.clear();
      this.filterCriteria.clear();
      this.app.showToast('AutoFilter turned off', 'info');
    } else {
      if (sel.startCol === sel.endCol && sel.startRow === sel.endRow) {
        sel = this.app.duplicateTools.detectDataRange(sheet, sel.startCol, sel.startRow);
      }
      sheet.autoFilterRange = { ...sel };
      this.app.showToast('AutoFilter applied', 'success');
    }
    this.app.gridEngine.render();
  }

  // Open the Column Filter Dropdown menu
  openFilterDropdown(col, screenX, screenY) {
    this.activeFilterMenuCol = col;
    const sheet = this.app.workbook.getActiveSheet();
    const filterRange = sheet.autoFilterRange;
    if (!filterRange) return;

    const startRow = filterRange.startRow + 1; // data starts below header
    const endRow = filterRange.endRow;

    // Collect distinct values
    const distinctVals = new Map();
    for (let r = startRow; r <= endRow; r++) {
      const cell = sheet.getCell(col, r);
      const strVal = cell && cell.value !== undefined ? String(cell.value) : '(Blanks)';
      distinctVals.set(strVal, (distinctVals.get(strVal) || 0) + 1);
    }

    const currentFilterSet = this.filterCriteria.get(col);

    let itemsHtml = '';
    distinctVals.forEach((cnt, val) => {
      const isChecked = !currentFilterSet || currentFilterSet.has(val);
      itemsHtml += `
        <label class="checkbox-item" style="font-size:11px;">
          <input type="checkbox" class="filter-val-check" data-val="${encodeURIComponent(val)}" ${isChecked ? 'checked' : ''}>
          <span>${val} <span style="color:var(--excel-text-muted);">(${cnt})</span></span>
        </label>
      `;
    });

    const menu = document.getElementById('filter-dropdown-menu');
    menu.innerHTML = `
      <div style="padding:6px;width:240px;background:var(--excel-bg);border-radius:6px;box-shadow:var(--excel-shadow-lg);border:1px solid var(--excel-border);">
        <div class="menu-item" id="filter-sort-asc">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 6h18M3 12h12M3 18h6"/></svg>
          Sort A to Z (Smallest to Largest)
        </div>
        <div class="menu-item" id="filter-sort-desc">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 18h18M3 12h12M3 6h6"/></svg>
          Sort Z to A (Largest to Smallest)
        </div>
        <div class="menu-divider"></div>
        <div class="menu-item" id="filter-clear-col">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><path d="M15 9l-6 6M9 9l6 6"/></svg>
          Clear Filter From This Column
        </div>
        <div class="menu-divider"></div>
        <input type="text" id="filter-search-input" placeholder="Search..." style="width:100%;padding:4px 6px;margin-bottom:6px;border:1px solid var(--excel-border);border-radius:3px;font-size:11px;outline:none;">
        <label class="checkbox-item" style="font-size:11px;font-weight:600;margin-bottom:4px;">
          <input type="checkbox" id="filter-select-all" checked>
          <span>(Select All)</span>
        </label>
        <div class="checkbox-group" id="filter-values-list" style="max-height:140px;">
          ${itemsHtml}
        </div>
        <div style="display:flex;justify-content:flex-end;gap:6px;margin-top:8px;">
          <button class="btn-secondary" id="filter-btn-cancel" style="padding:2px 8px;font-size:11px;">Cancel</button>
          <button class="btn-primary" id="filter-btn-ok" style="padding:2px 10px;font-size:11px;">OK</button>
        </div>
      </div>
    `;

    menu.style.left = `${Math.min(window.innerWidth - 260, screenX)}px`;
    menu.style.top = `${Math.min(window.innerHeight - 300, screenY)}px`;
    menu.classList.add('active');

    // Event listeners
    document.getElementById('filter-sort-asc').onclick = () => {
      this.sortColumn(col, true);
      menu.classList.remove('active');
    };
    document.getElementById('filter-sort-desc').onclick = () => {
      this.sortColumn(col, false);
      menu.classList.remove('active');
    };
    document.getElementById('filter-clear-col').onclick = () => {
      this.filterCriteria.delete(col);
      this.recalculateFilters(sheet);
      menu.classList.remove('active');
    };
    document.getElementById('filter-select-all').onchange = (e) => {
      document.querySelectorAll('.filter-val-check').forEach(cb => cb.checked = e.target.checked);
    };
    document.getElementById('filter-search-input').oninput = (e) => {
      const q = e.target.value.toLowerCase();
      document.querySelectorAll('#filter-values-list .checkbox-item').forEach(item => {
        const text = item.querySelector('span').textContent.toLowerCase();
        item.style.display = text.includes(q) ? 'flex' : 'none';
      });
    };
    document.getElementById('filter-btn-cancel').onclick = () => menu.classList.remove('active');
    document.getElementById('filter-btn-ok').onclick = () => {
      const selectedSet = new Set();
      document.querySelectorAll('.filter-val-check:checked').forEach(cb => {
        selectedSet.add(decodeURIComponent(cb.dataset.val));
      });
      this.filterCriteria.set(col, selectedSet);
      this.recalculateFilters(sheet);
      menu.classList.remove('active');
    };
  }

  // Recalculate row visibility based on active filter criteria
  recalculateFilters(sheet) {
    sheet.hiddenRows.clear();
    const filterRange = sheet.autoFilterRange;
    if (!filterRange) {
      this.app.gridEngine.render();
      return;
    }

    const startRow = filterRange.startRow + 1;
    const endRow = filterRange.endRow;

    for (let r = startRow; r <= endRow; r++) {
      let isVisible = true;
      this.filterCriteria.forEach((allowedVals, col) => {
        const cell = sheet.getCell(col, r);
        const strVal = cell && cell.value !== undefined ? String(cell.value) : '(Blanks)';
        if (!allowedVals.has(strVal)) {
          isVisible = false;
        }
      });
      if (!isVisible) {
        sheet.hiddenRows.add(r);
      }
    }

    this.app.gridEngine.render();
  }

  // Sort single column in active selection or table
  sortColumn(col, ascending = true) {
    const sheet = this.app.workbook.getActiveSheet();
    let sel = sheet.selection || (sheet.autoFilterRange ? sheet.autoFilterRange : null);
    if (!sel) return;

    const hasHeader = sheet.autoFilterRange ? true : (sel.startRow < sel.endRow);
    const firstDataRow = hasHeader ? sel.startRow + 1 : sel.startRow;
    const endRow = sel.endRow;

    if (firstDataRow >= endRow) return;

    this.app.workbook.pushUndoState(`Sort ${ascending ? 'A to Z' : 'Z to A'}`);

    const rowsData = [];
    for (let r = firstDataRow; r <= endRow; r++) {
      const rowCells = [];
      for (let c = sel.startCol; c <= sel.endCol; c++) {
        rowCells.push(sheet.getCell(c, r));
      }
      rowsData.push({ rowIdx: r, cells: rowCells });
    }

    const sortColOffset = col - sel.startCol;
    rowsData.sort((a, b) => {
      const cellA = a.cells[sortColOffset];
      const cellB = b.cells[sortColOffset];
      const valA = cellA && cellA.value !== undefined ? cellA.value : '';
      const valB = cellB && cellB.value !== undefined ? cellB.value : '';

      if (valA === '' && valB !== '') return 1;
      if (valA !== '' && valB === '') return -1;

      if (!isNaN(valA) && !isNaN(valB) && valA !== '' && valB !== '') {
        return ascending ? Number(valA) - Number(valB) : Number(valB) - Number(valA);
      }
      const cmp = String(valA).localeCompare(String(valB));
      return ascending ? cmp : -cmp;
    });

    // Write back sorted rows
    for (let i = 0; i < rowsData.length; i++) {
      const targetRow = firstDataRow + i;
      const rowObj = rowsData[i];
      for (let j = 0; j < rowObj.cells.length; j++) {
        const sourceCell = rowObj.cells[j];
        const targetCol = sel.startCol + j;
        if (sourceCell) {
          const newCell = sheet.getOrCreateCell(targetCol, targetRow);
          newCell.value = sourceCell.value;
          newCell.formula = sourceCell.formula;
          if (sourceCell.style) newCell.style = JSON.parse(JSON.stringify(sourceCell.style));
        } else {
          sheet.deleteCell(targetCol, targetRow);
        }
      }
    }

    this.app.gridEngine.render();
    this.app.showToast(`Sorted data by Column ${FormulaEngine.colToLetter(col)}`, 'success');
  }

  // Open Multi-column Custom Sort Dialog
  openCustomSortDialog() {
    const sheet = this.app.workbook.getActiveSheet();
    let sel = sheet.selection;
    if (!sel) return;

    const modalOverlay = document.getElementById('modal-overlay');
    const modalContainer = document.getElementById('modal-container');

    let colOptions = '';
    for (let c = sel.startCol; c <= sel.endCol; c++) {
      const topCell = sheet.getCell(c, sel.startRow);
      const label = topCell && topCell.value ? String(topCell.value) : `Column ${FormulaEngine.colToLetter(c)}`;
      colOptions += `<option value="${c}">${label} (Col ${FormulaEngine.colToLetter(c)})</option>`;
    }

    modalContainer.innerHTML = `
      <div class="modal-dialog">
        <div class="modal-header">
          <span>Sort</span>
          <button class="modal-close-btn" onclick="document.getElementById('modal-overlay').classList.remove('active')">&times;</button>
        </div>
        <div class="modal-body">
          <div style="display:flex;justify-content:space-between;align-items:center;">
            <label class="checkbox-item" style="font-weight:600;">
              <input type="checkbox" id="sort-has-headers" checked>
              <span>My data has headers</span>
            </label>
          </div>
          <div class="form-row">
            <label>Sort by</label>
            <div style="display:flex;gap:8px;">
              <select id="sort-col-primary" class="form-input" style="flex:1;">
                ${colOptions}
              </select>
              <select id="sort-order-primary" class="form-input" style="width:140px;">
                <option value="asc">A to Z / Smallest to Largest</option>
                <option value="desc">Z to A / Largest to Smallest</option>
              </select>
            </div>
          </div>
        </div>
        <div class="modal-footer">
          <button class="btn-secondary" onclick="document.getElementById('modal-overlay').classList.remove('active')">Cancel</button>
          <button class="btn-primary" id="btn-execute-custom-sort">Sort</button>
        </div>
      </div>
    `;

    modalOverlay.classList.add('active');

    document.getElementById('btn-execute-custom-sort').onclick = () => {
      const col = parseInt(document.getElementById('sort-col-primary').value, 10);
      const ascending = document.getElementById('sort-order-primary').value === 'asc';
      modalOverlay.classList.remove('active');
      this.sortColumn(col, ascending);
    };
  }

  // Open Find & Replace Dialog (Ctrl+F / Ctrl+H)
  openFindReplaceDialog(initialTab = 'find') {
    const modalOverlay = document.getElementById('modal-overlay');
    const modalContainer = document.getElementById('modal-container');

    modalContainer.innerHTML = `
      <div class="modal-dialog" style="width:420px;">
        <div class="modal-header">
          <span>Find and Replace</span>
          <button class="modal-close-btn" onclick="document.getElementById('modal-overlay').classList.remove('active')">&times;</button>
        </div>
        <div class="modal-body">
          <div style="display:flex;gap:12px;border-bottom:1px solid var(--excel-border);padding-bottom:6px;margin-bottom:6px;">
            <button id="tab-find-btn" class="ribbon-tab ${initialTab === 'find' ? 'active' : ''}" style="color:var(--excel-text);">Find</button>
            <button id="tab-replace-btn" class="ribbon-tab ${initialTab === 'replace' ? 'active' : ''}" style="color:var(--excel-text);">Replace</button>
          </div>
          <div class="form-row">
            <label>Find what:</label>
            <input type="text" id="find-input" class="form-input" placeholder="Type text to find...">
          </div>
          <div class="form-row" id="replace-row" style="display:${initialTab === 'replace' ? 'flex' : 'none'};">
            <label>Replace with:</label>
            <input type="text" id="replace-input" class="form-input" placeholder="Replacement text...">
          </div>
          <div style="display:flex;gap:12px;margin-top:4px;">
            <label class="checkbox-item">
              <input type="checkbox" id="chk-match-case">
              <span>Match case</span>
            </label>
            <label class="checkbox-item">
              <input type="checkbox" id="chk-match-entire">
              <span>Match entire cell contents</span>
            </label>
          </div>
        </div>
        <div class="modal-footer">
          <button class="btn-secondary" id="btn-find-next">Find Next</button>
          <button class="btn-secondary" id="btn-replace" style="display:${initialTab === 'replace' ? 'inline-block' : 'none'};">Replace</button>
          <button class="btn-primary" id="btn-replace-all" style="display:${initialTab === 'replace' ? 'inline-block' : 'none'};">Replace All</button>
        </div>
      </div>
    `;

    modalOverlay.classList.add('active');
    setTimeout(() => document.getElementById('find-input').focus(), 50);

    const tabFind = document.getElementById('tab-find-btn');
    const tabReplace = document.getElementById('tab-replace-btn');
    const replaceRow = document.getElementById('replace-row');
    const btnReplace = document.getElementById('btn-replace');
    const btnReplaceAll = document.getElementById('btn-replace-all');

    tabFind.onclick = () => {
      tabFind.classList.add('active'); tabReplace.classList.remove('active');
      replaceRow.style.display = 'none'; btnReplace.style.display = 'none'; btnReplaceAll.style.display = 'none';
    };
    tabReplace.onclick = () => {
      tabReplace.classList.add('active'); tabFind.classList.remove('active');
      replaceRow.style.display = 'flex'; btnReplace.style.display = 'inline-block'; btnReplaceAll.style.display = 'inline-block';
    };

    document.getElementById('btn-find-next').onclick = () => this.findNext();
    btnReplace.onclick = () => this.replaceCurrent();
    btnReplaceAll.onclick = () => this.replaceAll();
  }

  findNext() {
    const term = document.getElementById('find-input').value;
    if (!term) return;
    const matchCase = document.getElementById('chk-match-case').checked;
    const matchEntire = document.getElementById('chk-match-entire').checked;

    const sheet = this.app.workbook.getActiveSheet();
    const cells = Object.keys(sheet.cells);

    let startCol = 0, startRow = 0;
    if (sheet.selection) {
      startCol = sheet.selection.startCol;
      startRow = sheet.selection.startRow + 1;
    }

    for (let r = 0; r < 200; r++) {
      for (let c = 0; c < 50; c++) {
        const cell = sheet.getCell(c, r);
        if (!cell || cell.value === undefined) continue;

        let strVal = String(cell.value);
        let findStr = term;
        if (!matchCase) {
          strVal = strVal.toLowerCase();
          findStr = findStr.toLowerCase();
        }

        let isMatch = matchEntire ? strVal === findStr : strVal.includes(findStr);
        if (isMatch) {
          sheet.selection = { startCol: c, startRow: r, endCol: c, endRow: r };
          this.app.gridEngine.render();
          this.app.showToast(`Found match at ${FormulaEngine.colToLetter(c)}${r + 1}`, 'info');
          return;
        }
      }
    }
    this.app.showToast('No matching cells found', 'warning');
  }

  replaceAll() {
    const term = document.getElementById('find-input').value;
    const repl = document.getElementById('replace-input').value;
    if (!term) return;

    const matchCase = document.getElementById('chk-match-case').checked;
    const matchEntire = document.getElementById('chk-match-entire').checked;

    const sheet = this.app.workbook.getActiveSheet();
    let replaceCount = 0;
    this.app.workbook.pushUndoState('Replace All');

    Object.keys(sheet.cells).forEach(key => {
      const cell = sheet.cells[key];
      if (!cell || cell.value === undefined) return;

      let strVal = String(cell.value);
      let findStr = term;
      let matched = false;

      if (matchEntire) {
        if (matchCase ? strVal === findStr : strVal.toLowerCase() === findStr.toLowerCase()) {
          cell.value = repl;
          matched = true;
        }
      } else {
        if (matchCase) {
          if (strVal.includes(findStr)) {
            cell.value = strVal.split(findStr).join(repl);
            matched = true;
          }
        } else {
          const regex = new RegExp(findStr.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
          if (regex.test(strVal)) {
            cell.value = strVal.replace(regex, repl);
            matched = true;
          }
        }
      }

      if (matched) replaceCount++;
    });

    this.app.gridEngine.render();
    document.getElementById('modal-overlay').classList.remove('active');
    this.app.showToast(`Excel completed its search and made ${replaceCount} replacement(s).`, 'success');
  }

  // Open Text to Columns Wizard
  openTextToColumnsWizard() {
    const sheet = this.app.workbook.getActiveSheet();
    const sel = sheet.selection;
    if (!sel) return;

    const modalOverlay = document.getElementById('modal-overlay');
    const modalContainer = document.getElementById('modal-container');

    modalContainer.innerHTML = `
      <div class="modal-dialog" style="width:480px;">
        <div class="modal-header">
          <span>Convert Text to Columns Wizard</span>
          <button class="modal-close-btn" onclick="document.getElementById('modal-overlay').classList.remove('active')">&times;</button>
        </div>
        <div class="modal-body">
          <p style="font-size:12px;color:var(--excel-text-secondary);">Select the delimiters that separate your data:</p>
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:6px;margin:6px 0;">
            <label class="checkbox-item"><input type="checkbox" id="del-comma" checked><span>Comma (,)</span></label>
            <label class="checkbox-item"><input type="checkbox" id="del-tab"><span>Tab</span></label>
            <label class="checkbox-item"><input type="checkbox" id="del-semicolon"><span>Semicolon (;)</span></label>
            <label class="checkbox-item"><input type="checkbox" id="del-space"><span>Space</span></label>
          </div>
        </div>
        <div class="modal-footer">
          <button class="btn-secondary" onclick="document.getElementById('modal-overlay').classList.remove('active')">Cancel</button>
          <button class="btn-primary" id="btn-execute-text-to-columns">Finish</button>
        </div>
      </div>
    `;

    modalOverlay.classList.add('active');

    document.getElementById('btn-execute-text-to-columns').onclick = () => {
      const useComma = document.getElementById('del-comma').checked;
      const useTab = document.getElementById('del-tab').checked;
      const useSemi = document.getElementById('del-semicolon').checked;
      const useSpace = document.getElementById('del-space').checked;

      const delimiters = [];
      if (useComma) delimiters.push(',');
      if (useTab) delimiters.push('\t');
      if (useSemi) delimiters.push(';');
      if (useSpace) delimiters.push(' ');

      if (delimiters.length === 0) delimiters.push(',');

      modalOverlay.classList.remove('active');
      this.executeTextToColumns(sheet, sel, delimiters);
    };
  }

  executeTextToColumns(sheet, sel, delimiters) {
    const delimRegex = new RegExp(`[${delimiters.map(d => '\\' + d).join('')}]`);
    this.app.workbook.pushUndoState('Text to Columns');

    for (let r = sel.startRow; r <= sel.endRow; r++) {
      const cell = sheet.getCell(sel.startCol, r);
      if (cell && cell.value !== undefined) {
        const parts = String(cell.value).split(delimRegex);
        for (let i = 0; i < parts.length; i++) {
          const targetCol = sel.startCol + i;
          sheet.setCellValue(targetCol, r, parts[i].trim());
        }
      }
    }

    this.app.gridEngine.render();
    this.app.showToast('Text converted to columns successfully', 'success');
  }

  // Open Data Validation Dialog
  openDataValidationDialog() {
    const sheet = this.app.workbook.getActiveSheet();
    const sel = sheet.selection;
    if (!sel) return;

    const modalOverlay = document.getElementById('modal-overlay');
    const modalContainer = document.getElementById('modal-container');

    modalContainer.innerHTML = `
      <div class="modal-dialog" style="width:420px;">
        <div class="modal-header">
          <span>Data Validation</span>
          <button class="modal-close-btn" onclick="document.getElementById('modal-overlay').classList.remove('active')">&times;</button>
        </div>
        <div class="modal-body">
          <div class="form-row">
            <label>Allow:</label>
            <select id="val-allow-type" class="form-input">
              <option value="list">List</option>
              <option value="number">Whole Number</option>
              <option value="decimal">Decimal</option>
              <option value="textlength">Text Length</option>
            </select>
          </div>
          <div class="form-row" id="val-source-row">
            <label id="val-source-label">Source (comma-separated items or range e.g. Pass, Fail, Pending):</label>
            <input type="text" id="val-source-input" class="form-input" placeholder="Pass, Fail, In Review">
          </div>
        </div>
        <div class="modal-footer">
          <button class="btn-secondary" id="btn-val-clear">Clear All</button>
          <button class="btn-primary" id="btn-val-apply">OK</button>
        </div>
      </div>
    `;

    modalOverlay.classList.add('active');

    document.getElementById('btn-val-apply').onclick = () => {
      const type = document.getElementById('val-allow-type').value;
      const source = document.getElementById('val-source-input').value;

      this.app.workbook.pushUndoState('Apply Data Validation');
      for (let r = sel.startRow; r <= sel.endRow; r++) {
        for (let c = sel.startCol; c <= sel.endCol; c++) {
          const cell = sheet.getOrCreateCell(c, r);
          cell.validation = { type, source };
        }
      }

      modalOverlay.classList.remove('active');
      this.app.gridEngine.render();
      this.app.showToast('Data validation rule applied', 'success');
    };

    document.getElementById('btn-val-clear').onclick = () => {
      this.app.workbook.pushUndoState('Clear Data Validation');
      for (let r = sel.startRow; r <= sel.endRow; r++) {
        for (let c = sel.startCol; c <= sel.endCol; c++) {
          const cell = sheet.getCell(c, r);
          if (cell && cell.validation) delete cell.validation;
        }
      }
      modalOverlay.classList.remove('active');
      this.app.gridEngine.render();
      this.app.showToast('Data validation cleared', 'info');
    };
  }
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = DataTools;
}
