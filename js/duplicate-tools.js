/**
 * 0cell Duplicate Management Suite
 * Full Excel-compatible Duplicate finder, highlighter, remover, and unique extractor
 */

class DuplicateTools {
  constructor(app) {
    this.app = app;
  }

  // Highlight duplicate values in selected cell range (1-click toggle)
  highlightDuplicates(customStyle = { bg: '#ffc7ce', color: '#9c0006' }) {
    const sheet = this.app.workbook.getActiveSheet();
    const sel = sheet.selection;
    if (!sel) {
      this.app.showToast('Please select a range of cells first', 'warning');
      return;
    }

    const { startCol, startRow, endCol, endRow } = sel;

    // Check if selection already has duplicate highlights - if so, toggle/clear them
    let hasExistingHighlights = false;
    for (let r = startRow; r <= endRow; r++) {
      for (let c = startCol; c <= endCol; c++) {
        const cell = sheet.getCell(c, r);
        if (cell && cell.style && cell.style.bgColor === customStyle.bg) {
          hasExistingHighlights = true;
          break;
        }
      }
      if (hasExistingHighlights) break;
    }

    if (hasExistingHighlights) {
      this.clearHighlights();
      this.app.showToast('Cleared duplicate highlights', 'info');
      return;
    }

    const valueMap = new Map();

    // 1st pass: count occurrences
    for (let r = startRow; r <= endRow; r++) {
      for (let c = startCol; c <= endCol; c++) {
        const cell = sheet.getCell(c, r);
        const val = cell ? (cell.value !== undefined ? String(cell.value).trim() : '') : '';
        if (val !== '') {
          valueMap.set(val, (valueMap.get(val) || 0) + 1);
        }
      }
    }

    // 2nd pass: apply highlight style if count > 1
    let duplicateCount = 0;
    this.app.workbook.pushUndoState('Highlight Duplicates');

    for (let r = startRow; r <= endRow; r++) {
      for (let c = startCol; c <= endCol; c++) {
        const cell = sheet.getOrCreateCell(c, r);
        const val = cell.value !== undefined ? String(cell.value).trim() : '';
        if (val !== '' && (valueMap.get(val) || 0) > 1) {
          cell.style = cell.style || {};
          cell.style.bgColor = customStyle.bg;
          cell.style.color = customStyle.color;
          duplicateCount++;
        }
      }
    }

    this.app.gridEngine.render();
    this.app.autoSaveEngine?.triggerSave();
    if (duplicateCount > 0) {
      this.app.showToast(`Found and highlighted ${duplicateCount} duplicate cell(s)`, 'success');
    } else {
      this.app.showToast('No duplicate values found in the selected range', 'info');
    }
  }

  // Clear conditional duplicate formatting in selection
  clearHighlights() {
    const sheet = this.app.workbook.getActiveSheet();
    const sel = sheet.selection;
    if (!sel) return;

    this.app.workbook.pushUndoState('Clear Duplicate Highlights');
    for (let r = sel.startRow; r <= sel.endRow; r++) {
      for (let c = sel.startCol; c <= sel.endCol; c++) {
        const cell = sheet.getCell(c, r);
        if (cell && cell.style) {
          delete cell.style.bgColor;
          delete cell.style.color;
        }
      }
    }
    this.app.gridEngine.render();
    this.app.showToast('Cleared formatting from selection', 'info');
  }

  // Open the Excel-style "Remove Duplicates" modal
  openRemoveDuplicatesDialog() {
    const sheet = this.app.workbook.getActiveSheet();
    let sel = sheet.selection;

    // If only single cell selected, auto-expand to contiguous data table range
    if (!sel || (sel.startCol === sel.endCol && sel.startRow === sel.endRow)) {
      sel = this.detectDataRange(sheet, sel ? sel.startCol : 0, sel ? sel.startRow : 0);
    }

    const { startCol, startRow, endCol, endRow } = sel;
    const numRows = endRow - startRow + 1;
    const numCols = endCol - startCol + 1;

    if (numRows <= 1) {
      this.app.showToast('Select a range with at least 2 rows to remove duplicates', 'warning');
      return;
    }

    // Build modal HTML
    const modalOverlay = document.getElementById('modal-overlay');
    const modalContainer = document.getElementById('modal-container');

    let columnCheckboxes = '';
    for (let c = startCol; c <= endCol; c++) {
      const topCell = sheet.getCell(c, startRow);
      const headerLabel = topCell && topCell.value ? String(topCell.value) : `Column ${FormulaEngine.colToLetter(c)}`;
      columnCheckboxes += `
        <label class="checkbox-item">
          <input type="checkbox" class="col-dup-check" data-col="${c}" checked>
          <span>${headerLabel} (Col ${FormulaEngine.colToLetter(c)})</span>
        </label>
      `;
    }

    modalContainer.innerHTML = `
      <div class="modal-dialog">
        <div class="modal-header">
          <span>Remove Duplicates</span>
          <button class="modal-close-btn" onclick="document.getElementById('modal-overlay').classList.remove('active')">&times;</button>
        </div>
        <div class="modal-body">
          <p style="font-size:12px;color:var(--excel-text-secondary);">
            To delete duplicate values, select one or more columns that contain duplicates.
          </p>
          <div style="display:flex;gap:8px;margin-bottom:6px;">
            <button class="btn-secondary" id="btn-dup-select-all" style="padding:3px 10px;font-size:11px;">Select All</button>
            <button class="btn-secondary" id="btn-dup-unselect-all" style="padding:3px 10px;font-size:11px;">Unselect All</button>
          </div>
          <label class="checkbox-item" style="margin-bottom:8px;font-weight:600;">
            <input type="checkbox" id="chk-has-headers" checked>
            <span>My data has headers</span>
          </label>
          <div class="checkbox-group" id="dup-cols-list">
            ${columnCheckboxes}
          </div>
        </div>
        <div class="modal-footer">
          <button class="btn-secondary" onclick="document.getElementById('modal-overlay').classList.remove('active')">Cancel</button>
          <button class="btn-primary" id="btn-execute-remove-duplicates">OK</button>
        </div>
      </div>
    `;

    modalOverlay.classList.add('active');

    // Wire events
    document.getElementById('btn-dup-select-all').onclick = () => {
      document.querySelectorAll('.col-dup-check').forEach(cb => cb.checked = true);
    };
    document.getElementById('btn-dup-unselect-all').onclick = () => {
      document.querySelectorAll('.col-dup-check').forEach(cb => cb.checked = false);
    };

    document.getElementById('btn-execute-remove-duplicates').onclick = () => {
      const hasHeaders = document.getElementById('chk-has-headers').checked;
      const selectedCols = Array.from(document.querySelectorAll('.col-dup-check:checked')).map(cb => parseInt(cb.dataset.col, 10));

      if (selectedCols.length === 0) {
        alert('Please select at least one column.');
        return;
      }

      modalOverlay.classList.remove('active');
      this.executeRemoveDuplicates(sheet, sel, selectedCols, hasHeaders);
    };
  }

  // Execute the removal of duplicate rows
  executeRemoveDuplicates(sheet, range, compareCols, hasHeaders) {
    const { startCol, startRow, endCol, endRow } = range;
    const rowOffset = hasHeaders ? 1 : 0;
    const firstDataRow = startRow + rowOffset;

    if (firstDataRow > endRow) {
      this.app.showToast('No data rows found below header', 'warning');
      return;
    }

    const seenKeySet = new Set();
    const rowsToKeep = [];
    let duplicatesRemoved = 0;

    // Collect header row if present
    if (hasHeaders) {
      const headerRowData = [];
      for (let c = startCol; c <= endCol; c++) {
        headerRowData.push(sheet.getCell(c, startRow));
      }
      rowsToKeep.push({ rowIdx: startRow, cells: headerRowData });
    }

    // Inspect data rows
    for (let r = firstDataRow; r <= endRow; r++) {
      // Build unique key from chosen columns
      const keyParts = compareCols.map(c => {
        const cell = sheet.getCell(c, r);
        return cell && cell.value !== undefined ? String(cell.value).trim().toLowerCase() : '';
      });
      const rowKey = keyParts.join('|||');

      if (seenKeySet.has(rowKey)) {
        duplicatesRemoved++;
      } else {
        seenKeySet.add(rowKey);
        const rowData = [];
        for (let c = startCol; c <= endCol; c++) {
          rowData.push(sheet.getCell(c, r));
        }
        rowsToKeep.push({ rowIdx: r, cells: rowData });
      }
    }

    const uniqueValuesCount = seenKeySet.size;

    // Perform atomic update with Undo support
    this.app.workbook.pushUndoState('Remove Duplicates');

    // Clear the original range
    for (let r = startRow; r <= endRow; r++) {
      for (let c = startCol; c <= endCol; c++) {
        sheet.deleteCell(c, r);
      }
    }

    // Re-populate with distinct rows
    for (let i = 0; i < rowsToKeep.length; i++) {
      const targetRow = startRow + i;
      const rowObj = rowsToKeep[i];
      for (let j = 0; j < rowObj.cells.length; j++) {
        const sourceCell = rowObj.cells[j];
        if (sourceCell) {
          const targetCol = startCol + j;
          const newCell = sheet.getOrCreateCell(targetCol, targetRow);
          newCell.value = sourceCell.value;
          newCell.formula = sourceCell.formula;
          if (sourceCell.style) newCell.style = JSON.parse(JSON.stringify(sourceCell.style));
        }
      }
    }

    // Update selection to match new size
    const newEndRow = startRow + rowsToKeep.length - 1;
    sheet.selection = { startCol, startRow, endCol, endRow: newEndRow };

    this.app.gridEngine.render();

    // Excel identical summary message dialog
    this.showDuplicateSummaryModal(duplicatesRemoved, uniqueValuesCount);
  }

  showDuplicateSummaryModal(removed, remaining) {
    const modalOverlay = document.getElementById('modal-overlay');
    const modalContainer = document.getElementById('modal-container');

    modalContainer.innerHTML = `
      <div class="modal-dialog" style="width:380px;">
        <div class="modal-header">
          <span>Microsoft Excel / 0cell</span>
          <button class="modal-close-btn" onclick="document.getElementById('modal-overlay').classList.remove('active')">&times;</button>
        </div>
        <div class="modal-body" style="gap:8px;">
          <div style="display:flex;align-items:center;gap:12px;">
            <div style="width:36px;height:36px;border-radius:50%;background:#e8f5ec;display:flex;align-items:center;justify-content:center;color:var(--excel-green);font-size:20px;font-weight:bold;">
              ✓
            </div>
            <div>
              <p style="font-weight:600;font-size:13px;margin-bottom:4px;">
                ${removed} duplicate value(s) found and removed.
              </p>
              <p style="font-size:12px;color:var(--excel-text-secondary);">
                ${remaining} unique value(s) remain.
              </p>
            </div>
          </div>
        </div>
        <div class="modal-footer">
          <button class="btn-primary" onclick="document.getElementById('modal-overlay').classList.remove('active')">OK</button>
        </div>
      </div>
    `;
    modalOverlay.classList.add('active');
  }

  // Extract unique/distinct values from selection to next column or new sheet
  extractUniqueValues() {
    const sheet = this.app.workbook.getActiveSheet();
    const sel = sheet.selection;
    if (!sel) {
      this.app.showToast('Please select a range of cells first', 'warning');
      return;
    }

    const { startCol, startRow, endCol, endRow } = sel;
    const uniqueValues = [];
    const seen = new Set();

    for (let r = startRow; r <= endRow; r++) {
      for (let c = startCol; c <= endCol; c++) {
        const cell = sheet.getCell(c, r);
        if (cell && cell.value !== undefined && cell.value !== '') {
          const sVal = String(cell.value).trim();
          if (!seen.has(sVal.toLowerCase())) {
            seen.add(sVal.toLowerCase());
            uniqueValues.push(cell.value);
          }
        }
      }
    }

    if (uniqueValues.length === 0) {
      this.app.showToast('No values found in selection', 'info');
      return;
    }

    // Paste unique values adjacent to the selection
    const targetCol = endCol + 1;
    this.app.workbook.pushUndoState('Extract Unique Values');

    sheet.setCellValue(targetCol, startRow, 'Unique Values');
    const headerCell = sheet.getOrCreateCell(targetCol, startRow);
    headerCell.style = { bold: true, bgColor: '#f3f2f1' };

    for (let i = 0; i < uniqueValues.length; i++) {
      sheet.setCellValue(targetCol, startRow + 1 + i, uniqueValues[i]);
    }

    this.app.gridEngine.render();
    this.app.showToast(`Extracted ${uniqueValues.length} unique value(s) to Column ${FormulaEngine.colToLetter(targetCol)}`, 'success');
  }

  // Auto-detect contiguous data table block around a cell
  detectDataRange(sheet, startCol, startRow) {
    let minC = startCol, maxC = startCol;
    let minR = startRow, maxR = startRow;

    // Scan bounding box of populated cells around the point
    const cells = Object.keys(sheet.cells);
    if (cells.length === 0) return { startCol: 0, startRow: 0, endCol: 5, endRow: 10 };

    let foundAny = false;
    cells.forEach(key => {
      const [c, r] = key.split(',').map(Number);
      if (Math.abs(c - startCol) <= 10 && Math.abs(r - startRow) <= 50) {
        if (!foundAny) {
          minC = maxC = c;
          minR = maxR = r;
          foundAny = true;
        } else {
          minC = Math.min(minC, c);
          maxC = Math.max(maxC, c);
          minR = Math.min(minR, r);
          maxR = Math.max(maxR, r);
        }
      }
    });

    return { startCol: minC, startRow: minR, endCol: maxC, endRow: maxR };
  }
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = DuplicateTools;
}
