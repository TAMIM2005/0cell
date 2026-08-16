/**
 * 0cell Unlimited High-Performance IndexedDB Storage
 * Supports unlimited project sizes, thousands of sheets/rows, and embedded images.
 */
class ZeroStorage {
  static dbName = 'ZeroCellDatabase_v3';
  static version = 1;
  static _dbPromise = null;

  static getDB() {
    if (!ZeroStorage._dbPromise) {
      ZeroStorage._dbPromise = new Promise((resolve, reject) => {
        const req = indexedDB.open(ZeroStorage.dbName, ZeroStorage.version);
        req.onupgradeneeded = (e) => {
          const db = e.target.result;
          if (!db.objectStoreNames.contains('projects')) {
            db.createObjectStore('projects', { keyPath: 'id' });
          }
          if (!db.objectStoreNames.contains('meta')) {
            db.createObjectStore('meta', { keyPath: 'key' });
          }
        };
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
      });
    }
    return ZeroStorage._dbPromise;
  }

  static async saveProject(projectData) {
    try {
      const db = await ZeroStorage.getDB();
      return new Promise((resolve) => {
        const tx = db.transaction(['projects', 'meta'], 'readwrite');
        tx.objectStore('projects').put(projectData);
        tx.objectStore('meta').put({ key: 'lastActiveDocId', value: projectData.id });
        tx.objectStore('meta').put({ key: 'lastActiveProjectData', value: projectData });
        tx.oncomplete = () => resolve(true);
        tx.onerror = () => resolve(false);
      });
    } catch (e) {
      console.warn('ZeroStorage save error:', e);
      return false;
    }
  }

  static async getLastActiveProject() {
    try {
      const db = await ZeroStorage.getDB();
      return new Promise((resolve) => {
        const tx = db.transaction(['meta'], 'readonly');
        const req = tx.objectStore('meta').get('lastActiveProjectData');
        req.onsuccess = () => resolve(req.result ? req.result.value : null);
        req.onerror = () => resolve(null);
      });
    } catch (e) {
      return null;
    }
  }

  static async getProject(id) {
    try {
      const db = await ZeroStorage.getDB();
      return new Promise((resolve) => {
        const tx = db.transaction(['projects'], 'readonly');
        const req = tx.objectStore('projects').get(id);
        req.onsuccess = () => resolve(req.result || null);
        req.onerror = () => resolve(null);
      });
    } catch (e) {
      return null;
    }
  }

  static async getAllRecentProjects() {
    try {
      const db = await ZeroStorage.getDB();
      return new Promise((resolve) => {
        const tx = db.transaction(['projects'], 'readonly');
        const req = tx.objectStore('projects').getAll();
        req.onsuccess = () => {
          const list = req.result || [];
          list.sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0));
          resolve(list);
        };
        req.onerror = () => resolve([]);
      });
    } catch (e) {
      return [];
    }
  }

  static async deleteProject(id) {
    try {
      const db = await ZeroStorage.getDB();
      return new Promise((resolve) => {
        const tx = db.transaction(['projects'], 'readwrite');
        tx.objectStore('projects').delete(id);
        tx.oncomplete = () => resolve(true);
        tx.onerror = () => resolve(false);
      });
    } catch (e) {
      return false;
    }
  }

  static async clearAllProjects() {
    try {
      const db = await ZeroStorage.getDB();
      return new Promise((resolve) => {
        const tx = db.transaction(['projects', 'meta'], 'readwrite');
        tx.objectStore('projects').clear();
        tx.objectStore('meta').clear();
        tx.oncomplete = () => resolve(true);
        tx.onerror = () => resolve(false);
      });
    } catch (e) {
      return false;
    }
  }
}

/**
 * 0cell Multi-Project AutoSave & Workspace Memory Engine
 */
class AutoSaveEngine {
  constructor(app) {
    this.app = app;
    this.currentDocId = 'proj_default';
    this.statusDot = null;
    this.statusLabel = null;
    this.saveTimer = null;
  }

  init() {
    this.statusDot = document.getElementById('autosave-dot');
    this.statusLabel = document.getElementById('autosave-label');

    const docNameInput = document.getElementById('doc-name-input');
    docNameInput?.addEventListener('input', () => {
      const val = docNameInput.value.trim();
      if (val) {
        this.app.workbook.name = val;
        this.currentDocId = 'proj_' + val.toLowerCase().replace(/[^a-z0-9_]/g, '_');
      }
      this.triggerSave();
    });

    window.addEventListener('beforeunload', () => this.saveInstant());
  }

  setCurrentProject(docName, explicitId = null) {
    const cleanName = (docName || 'Untitled Spreadsheet').trim();
    this.currentDocId = explicitId || ('proj_' + cleanName.toLowerCase().replace(/[^a-z0-9_]/g, '_'));
    this.app.workbook.name = cleanName;

    const docNameInput = document.getElementById('doc-name-input');
    if (docNameInput) docNameInput.value = cleanName;

    this.saveInstant();
  }

  saveInstant() {
    try {
      if (this.statusDot) this.statusDot.className = 'doc-status-dot saving';
      if (this.statusLabel) this.statusLabel.textContent = 'Saving...';

      const docName = document.getElementById('doc-name-input')?.value || this.app.workbook.name || 'Untitled Spreadsheet';

      const projectData = {
        id: this.currentDocId,
        docName: docName,
        workbook: this.app.workbook.serialize(),
        activeSheetIndex: this.app.workbook.activeSheetIndex || 0,
        timestamp: Date.now(),
        timeStr: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
      };

      // 1. Unlimited IndexedDB Save
      ZeroStorage.saveProject(projectData);

      // 2. Fast Session Memory & Live Excel Binary Cache
      try {
        sessionStorage.setItem('0cell_current_active_proj', JSON.stringify(projectData));
        
        // Auto-generate live native Excel (.xlsx) payload in background
        if (this.app.xlsxIO && typeof XLSX !== 'undefined') {
          setTimeout(() => {
            try {
              const wb = XLSX.utils.book_new();
              this.app.workbook.sheets.forEach(sheet => {
                const wsData = [];
                let maxC = 0, maxR = 0;
                Object.keys(sheet.cells).forEach(key => {
                  const [c, r] = key.split(',').map(Number);
                  if (c > maxC) maxC = c;
                  if (r > maxR) maxR = r;
                });
                for (let r = 0; r <= maxR; r++) {
                  const rowData = [];
                  for (let c = 0; c <= maxC; c++) {
                    const cell = sheet.getCell(c, r);
                    if (cell) {
                      if (cell.formula) rowData.push({ f: cell.formula.startsWith('=') ? cell.formula.substring(1) : cell.formula });
                      else if (cell.value !== undefined) rowData.push(cell.value);
                      else rowData.push('');
                    } else rowData.push('');
                  }
                  wsData.push(rowData);
                }
                const ws = XLSX.utils.aoa_to_sheet(wsData);
                if (sheet.merges && sheet.merges.length > 0) {
                  ws['!merges'] = sheet.merges.map(m => ({ s: { r: m.startRow, c: m.startCol }, e: { r: m.endRow, c: m.endCol } }));
                }
                XLSX.utils.book_append_sheet(wb, ws, sheet.name.substring(0, 31));
              });
              this.lastExcelWorkbook = wb;
            } catch (err) {}
          }, 50);
        }
      } catch (e) {}

      if (this.statusDot) this.statusDot.className = 'doc-status-dot saved';
      if (this.statusLabel) {
        this.statusLabel.textContent = 'Saved';
        this.statusLabel.title = `Auto-saved in real-time (${projectData.timeStr})`;
      }
    } catch (e) {
      console.warn('AutoSave error:', e);
    }
  }

  triggerSave() {
    if (this.saveTimer) cancelAnimationFrame(this.saveTimer);
    this.saveTimer = requestAnimationFrame(() => {
      this.saveInstant();
    });
  }

  async restoreSession() {
    try {
      // 1. Try to load the latest active project from IndexedDB
      const lastProject = await ZeroStorage.getLastActiveProject();
      if (lastProject && lastProject.workbook) {
        this.applyProjectPayload(lastProject);
        return true;
      }

      // 2. Fallback to session memory
      const sessStr = sessionStorage.getItem('0cell_current_active_proj');
      if (sessStr) {
        const payload = JSON.parse(sessStr);
        if (payload && payload.workbook) {
          this.applyProjectPayload(payload);
          return true;
        }
      }

      return false;
    } catch (err) {
      console.warn('Restore session error:', err);
      return false;
    }
  }

  applyProjectPayload(payload) {
    this.currentDocId = payload.id || 'proj_default';
    this.app.workbook = new WorkbookManager();
    this.app.formulaEngine.workbook = this.app.workbook;
    this.app.workbook.deserialize(payload.workbook);

    if (payload.activeSheetIndex !== undefined && this.app.workbook.sheets[payload.activeSheetIndex]) {
      this.app.workbook.activeSheetIndex = payload.activeSheetIndex;
    } else {
      this.app.workbook.activeSheetIndex = 0;
    }

    const docNameInput = document.getElementById('doc-name-input');
    const docTitle = payload.docName || this.app.workbook.name || 'Untitled Spreadsheet';
    if (docNameInput) docNameInput.value = docTitle;
    this.app.workbook.name = docTitle;

    this.app.updateSheetsTabBar();
    this.app.gridEngine.render();

    if (this.statusLabel) {
      this.statusLabel.textContent = 'Restored';
      this.statusLabel.title = `Restored "${docTitle}"`;
    }
  }

  async loadProject(docId) {
    try {
      const project = await ZeroStorage.getProject(docId);
      if (!project || !project.workbook) return false;

      // Save current project first
      this.saveInstant();

      this.applyProjectPayload(project);
      this.saveInstant();
      return true;
    } catch (err) {
      console.warn('Load project error:', err);
      return false;
    }
  }

  async renderRecentManagerUI(searchQuery = '') {
    const list = document.getElementById('recent-manager-list');
    if (!list) return;

    try {
      let recents = await ZeroStorage.getAllRecentProjects();
      if (searchQuery && searchQuery.trim()) {
        const q = searchQuery.trim().toLowerCase();
        recents = recents.filter(p => (p.docName || p.name || '').toLowerCase().includes(q));
      }

      if (!recents || recents.length === 0) {
        list.innerHTML = `
          <div style="padding:40px 20px; text-align:center; color:var(--text-muted);">
            <div style="font-size:32px; margin-bottom:8px;">📁</div>
            <div style="font-weight:600; font-size:13px;">No recent workbooks found</div>
            <div style="font-size:11px; margin-top:4px;">Open or create workbooks and they will appear here automatically.</div>
          </div>
        `;
        return;
      }

      list.innerHTML = '';
      recents.forEach(p => {
        const isCurrent = p.id === this.currentDocId;
        const card = document.createElement('div');
        card.style.display = 'flex';
        card.style.alignItems = 'center';
        card.style.justifyContent = 'space-between';
        card.style.padding = '10px 14px';
        card.style.background = isCurrent ? '#f0f6ff' : 'var(--bg-surface)';
        card.style.border = `1px solid ${isCurrent ? '#93c5fd' : 'var(--border-default)'}`;
        card.style.borderRadius = 'var(--radius-md)';
        card.style.transition = 'all 0.15s ease';
        card.style.gap = '12px';

        const dateStr = p.timestamp ? new Date(p.timestamp).toLocaleDateString([], { month: 'short', day: 'numeric', year: 'numeric' }) : '';
        const timeStr = p.timeStr || (p.timestamp ? new Date(p.timestamp).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }) : '');

        card.innerHTML = `
          <div style="display:flex; align-items:center; gap:12px; min-width:0; flex:1;">
            <div style="width:34px; height:34px; border-radius:6px; background:${isCurrent ? '#dbeafe' : '#f1f5f9'}; display:flex; align-items:center; justify-content:center; font-size:16px; flex-shrink:0;">
              📊
            </div>
            <div style="min-width:0; flex:1;">
              <div style="font-weight:${isCurrent ? '700' : '600'}; font-size:13px; color:${isCurrent ? '#1d4ed8' : 'var(--text-primary)'}; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;" title="${p.docName || p.name || 'Spreadsheet'}">
                ${p.docName || p.name || 'Spreadsheet'} ${isCurrent ? '<span style="font-size:10px; background:#2563eb; color:#fff; padding:1px 6px; border-radius:10px; margin-left:6px;">Current</span>' : ''}
              </div>
              <div style="font-size:11px; color:var(--text-muted); margin-top:2px; display:flex; gap:8px;">
                <span>📑 ${p.workbook?.sheets?.length || 1} sheet(s)</span>
                <span>•</span>
                <span>🕒 ${dateStr} ${timeStr}</span>
              </div>
            </div>
          </div>
          <div style="display:flex; align-items:center; gap:6px; flex-shrink:0;">
            <button class="btn-primary btn-open-recent" style="padding:6px 12px; font-size:11px; cursor:pointer;">
              Open
            </button>
            <button class="btn-secondary btn-del-recent" title="Remove from recent list" style="padding:6px 8px; font-size:11px; color:#ef4444; border-color:#fee2e2; cursor:pointer;">
              🗑️
            </button>
          </div>
        `;

        card.querySelector('.btn-open-recent')?.addEventListener('click', async () => {
          await this.loadProject(p.id);
          document.getElementById('recent-manager-modal')?.classList.remove('active');
        });

        card.querySelector('.btn-del-recent')?.addEventListener('click', async (e) => {
          e.stopPropagation();
          if (confirm(`Remove "${p.docName || p.name}" from recent list?`)) {
            await ZeroStorage.deleteProject(p.id);
            this.renderRecentManagerUI(searchQuery);
          }
        });

        list.appendChild(card);
      });
    } catch (err) {
      console.warn('Render recent manager error:', err);
    }
  }
}

class ZeroCellApp {
  constructor() {
    this.workbook = new WorkbookManager();
    this.formulaEngine = new FormulaEngine(this.workbook);
    this.duplicateTools = new DuplicateTools(this);
    this.dataTools = new DataTools(this);
    this.chartEngine = new ChartEngine(this);
    this.xlsxIO = new XlsxIO(this);
    this.autoSaveEngine = new AutoSaveEngine(this);
    this.gridEngine = null;
  }

  async init() {
    const gridContainer = document.getElementById('grid-container');
    this.gridEngine = new GridEngine(this, gridContainer);

    this.bindToolbarEvents();
    this.bindSheetTabs();
    this.bindFileMenu();
    this.bindRecentModal();
    this.bindDragAndDrop();
    this.autoSaveEngine.init();

    // Fast session restore or clean blank sheet
    const restored = await this.autoSaveEngine.restoreSession();
    if (!restored) {
      this.initBlankSheet();
      this.updateSheetsTabBar();
    }

    console.log('0cell Multi-Project Spreadsheet ready with Instant Real-Time Auto-Save.');
  }

  // Global Drag & Drop of XLSX, 0Cell, CSV files with Live File Handle Binding
  bindDragAndDrop() {
    window.addEventListener('dragover', (e) => {
      e.preventDefault();
      e.dataTransfer.dropEffect = 'copy';
    });

    window.addEventListener('drop', (e) => {
      e.preventDefault();
      const files = e.dataTransfer.files;
      if (files && files.length > 0) {
        const file = files[0];
        this.activeFileHandle = file;
        const fileNameLower = file.name.toLowerCase();
        if (fileNameLower.endsWith('.xlsx') || fileNameLower.endsWith('.xls') || fileNameLower.endsWith('.ods')) {
          const reader = new FileReader();
          reader.onload = (evt) => {
            this.xlsxIO.importXLSX(evt.target.result, file.name);
            this.setupLiveFileWatcher(file);
          };
          reader.readAsArrayBuffer(file);
        } else if (fileNameLower.endsWith('.0cell') || fileNameLower.endsWith('.json')) {
          const reader = new FileReader();
          reader.onload = (evt) => {
            this.xlsxIO.importFrom0Cell(evt.target.result);
          };
          reader.readAsText(file);
        } else if (fileNameLower.endsWith('.csv') || fileNameLower.endsWith('.tsv') || fileNameLower.endsWith('.txt')) {
          const reader = new FileReader();
          reader.onload = (evt) => {
            this.xlsxIO.importCSV(evt.target.result);
            this.setupLiveFileWatcher(file);
          };
          reader.readAsText(file);
        }
      }
    });
  }

  // Setup Live File Watcher: Polling for local Excel changes to sync instantly into 0cell
  setupLiveFileWatcher(file) {
    if (!file || !file.lastModified) return;
    if (this._fileWatcherInterval) clearInterval(this._fileWatcherInterval);

    let lastKnownTime = file.lastModified;
    this._fileWatcherInterval = setInterval(async () => {
      if (this.activeFileHandle && this.activeFileHandle.lastModified > lastKnownTime) {
        lastKnownTime = this.activeFileHandle.lastModified;
        const reader = new FileReader();
        reader.onload = (evt) => {
          this.xlsxIO.importXLSX(evt.target.result, file.name);
          this.showToast(`🔄 Live Sync: Excel updates imported into 0cell!`, 'info');
        };
        reader.readAsArrayBuffer(this.activeFileHandle);
      }
    }, 2000);
  }

  // Setup Native File Handle Watcher (File System Access API): Live 2-Way Sync with Excel File
  setupNativeFileHandleWatcher(handle) {
    if (!handle) return;
    if (this._fileWatcherInterval) clearInterval(this._fileWatcherInterval);

    let lastModTime = 0;
    handle.getFile().then(f => { lastModTime = f.lastModified; }).catch(() => {});

    this._fileWatcherInterval = setInterval(async () => {
      try {
        const freshFile = await handle.getFile();
        if (freshFile && freshFile.lastModified > lastModTime) {
          lastModTime = freshFile.lastModified;
          const reader = new FileReader();
          reader.onload = (evt) => {
            this.xlsxIO.importXLSX(evt.target.result, freshFile.name);
            this.showToast(`🔄 Live Sync: MS Excel changes imported into 0cell!`, 'info');
          };
          reader.readAsArrayBuffer(freshFile);
        }
      } catch (err) {}
    }, 1500);
  }

  // Initialize a completely clean, pristine, empty spreadsheet
  initBlankSheet() {
    const sheet = this.workbook.getActiveSheet();
    sheet.name = 'Sheet1';
    sheet.numCols = 26; // Columns A to Z
    sheet.numRows = 50; // Rows 1 to 50
    sheet.cells = {};
    sheet.merges = [];
    sheet.colWidths = {};
    sheet.rowHeights = {};
    sheet.selection = { startCol: 0, startRow: 0, endCol: 0, endRow: 0 };
    sheet.activeCell = { col: 0, row: 0 };

    const docNameInput = document.getElementById('doc-name-input');
    if (docNameInput) docNameInput.value = 'Untitled Spreadsheet';

    this.gridEngine.render();
  }

  // Bind all modern action toolbar buttons
  bindToolbarEvents() {
    // --- CLIPBOARD (Copy, Cut, Paste) ---
    document.getElementById('btn-tool-copy')?.addEventListener('click', () => {
      this.gridEngine.copySelection(false);
    });
    document.getElementById('btn-tool-cut')?.addEventListener('click', () => {
      this.gridEngine.copySelection(true);
    });
    document.getElementById('btn-tool-paste')?.addEventListener('click', () => {
      this.gridEngine.pasteToSelection();
    });

    // --- ROW & COLUMN OPERATIONS (+ and -) ---
    document.getElementById('btn-add-row-tool')?.addEventListener('click', () => {
      this.gridEngine.addRow(1);
    });
    document.getElementById('btn-del-row-tool')?.addEventListener('click', () => {
      this.gridEngine.deleteRow();
    });
    document.getElementById('btn-add-col-tool')?.addEventListener('click', () => {
      this.gridEngine.addColumn(1);
    });
    document.getElementById('btn-del-col-tool')?.addEventListener('click', () => {
      this.gridEngine.deleteColumn();
    });

    // --- DUPLICATE TOOLS (Featured) ---
    document.getElementById('btn-tool-highlight-dup')?.addEventListener('click', () => {
      this.duplicateTools.highlightDuplicates();
    });
    document.getElementById('btn-tool-clear-dup')?.addEventListener('click', () => {
      this.duplicateTools.clearHighlights();
    });
    document.getElementById('btn-tool-remove-dup')?.addEventListener('click', () => {
      this.duplicateTools.openRemoveDuplicatesDialog();
    });
    document.getElementById('btn-tool-extract-unique')?.addEventListener('click', () => {
      this.duplicateTools.extractUniqueValues();
    });

    // --- FORMULAS & AUTOSUM ---
    document.getElementById('btn-tool-autosum')?.addEventListener('click', () => {
      this.insertAutoSum('SUM');
    });
    document.getElementById('btn-tool-autoavg')?.addEventListener('click', () => {
      this.insertAutoSum('AVERAGE');
    });
    document.getElementById('btn-tool-autocount')?.addEventListener('click', () => {
      this.insertAutoSum('COUNT');
    });
    document.getElementById('btn-tool-automax')?.addEventListener('click', () => {
      this.insertAutoSum('MAX');
    });
    document.getElementById('btn-tool-automin')?.addEventListener('click', () => {
      this.insertAutoSum('MIN');
    });

    // --- FORMATTING & MERGE ---
    document.getElementById('btn-fmt-merge')?.addEventListener('click', () => this.gridEngine.toggleMerge());
    document.getElementById('btn-fmt-bold')?.addEventListener('click', () => this.toggleStyleProperty('bold'));
    document.getElementById('btn-fmt-italic')?.addEventListener('click', () => this.toggleStyleProperty('italic'));
    document.getElementById('btn-fmt-underline')?.addEventListener('click', () => this.toggleStyleProperty('underline'));

    // Google Sheets Fill Color Picker
    this.initGSheetFillColorPicker();

    const colorPickerFont = document.getElementById('color-picker-font');
    document.getElementById('btn-fmt-color')?.addEventListener('click', () => colorPickerFont.click());
    colorPickerFont?.addEventListener('input', (e) => this.setStyleProperty('color', e.target.value));

    document.getElementById('btn-fmt-currency')?.addEventListener('click', () => this.setStyleProperty('format', 'currency'));
    document.getElementById('btn-fmt-percent')?.addEventListener('click', () => this.setStyleProperty('format', 'percent'));

    document.getElementById('btn-align-left')?.addEventListener('click', () => this.setStyleProperty('align', 'left'));
    document.getElementById('btn-align-center')?.addEventListener('click', () => this.setStyleProperty('align', 'center'));
    document.getElementById('btn-align-right')?.addEventListener('click', () => this.setStyleProperty('align', 'right'));

    // --- DATA & SORT ---
    document.getElementById('btn-tool-autofilter')?.addEventListener('click', () => this.dataTools.toggleAutoFilter());
    document.getElementById('btn-tool-sort-asc')?.addEventListener('click', () => {
      const sheet = this.workbook.getActiveSheet();
      if (sheet.activeCell) this.dataTools.sortColumn(sheet.activeCell.col, true);
    });
    document.getElementById('btn-tool-sort-desc')?.addEventListener('click', () => {
      const sheet = this.workbook.getActiveSheet();
      if (sheet.activeCell) this.dataTools.sortColumn(sheet.activeCell.col, false);
    });
    document.getElementById('btn-tool-find')?.addEventListener('click', () => this.dataTools.openFindReplaceDialog('find'));

    // --- CHARTS ---
    document.getElementById('btn-insert-chart')?.addEventListener('click', () => this.chartEngine.insertChart('column'));

    // --- UNDO / REDO / THEME ---
    document.getElementById('btn-undo')?.addEventListener('click', () => {
      this.workbook.undo();
      this.gridEngine.render();
      this.autoSaveEngine?.triggerSave();
    });
    // --- WINDOW CAPTION CONTROLS (Minimize, Maximize, Close) ---
    document.getElementById('btn-win-close')?.addEventListener('click', () => {
      this.autoSaveEngine?.saveInstant();
      window.close();
    });

    document.getElementById('btn-win-max')?.addEventListener('click', () => {
      if (!document.fullscreenElement) {
        if (document.documentElement.requestFullscreen) {
          document.documentElement.requestFullscreen().catch(() => {});
        }
      } else {
        if (document.exitFullscreen) {
          document.exitFullscreen().catch(() => {});
        }
      }
    });

    document.getElementById('btn-win-min')?.addEventListener('click', () => {
      if (document.fullscreenElement) {
        document.exitFullscreen().catch(() => {});
      }
      window.blur();
    });
  }

  // Google Sheets Style Fill Color Picker & Reset Controller
  initGSheetFillColorPicker() {
    const wrap = document.getElementById('fill-color-wrap');
    const btn = document.getElementById('btn-fmt-fill');
    const popover = document.getElementById('fill-color-popover');
    const paletteGrid = document.getElementById('gsheet-fill-palette');
    const resetBtn = document.getElementById('btn-fill-reset');
    const customBtn = document.getElementById('btn-fill-custom');
    const colorPickerBg = document.getElementById('color-picker-bg');
    const stripe = document.getElementById('fill-color-stripe');

    if (!btn || !popover || !paletteGrid) return;

    // Google Sheets Standard 10x7 Palette Matrix
    const GSHEET_COLORS = [
      '#000000', '#434343', '#666666', '#999999', '#b7b7b7', '#cccccc', '#d9d9d9', '#efefef', '#f3f3f3', '#ffffff',
      '#980000', '#ff0000', '#ff9900', '#ffff00', '#00ff00', '#00ffff', '#4a86e8', '#0000ff', '#9900ff', '#ff00ff',
      '#e6b8af', '#f4cccc', '#fce5cd', '#fff2cc', '#d9ead3', '#d0e0e3', '#c9daf8', '#cfe2f3', '#d9d2e9', '#ead1dc',
      '#dd7e6b', '#ea9999', '#f9cb9c', '#ffe599', '#b6d7a8', '#a2c4c9', '#a4c2f4', '#9fc5e8', '#b4a7d6', '#d5a6bd',
      '#cc4125', '#e06666', '#f6b26b', '#ffd966', '#93c47d', '#76a5af', '#6d9eeb', '#6fa8dc', '#8e7cc3', '#c27ba0',
      '#a61c1c', '#cc0000', '#e69138', '#f1c232', '#6aa84f', '#45818e', '#3c78d8', '#3d85c6', '#674ea7', '#a64d79',
      '#5b0f00', '#660000', '#783f04', '#7f6000', '#274e13', '#0c343d', '#1155cc', '#0b5394', '#351c75', '#4c1130'
    ];

    paletteGrid.innerHTML = '';
    GSHEET_COLORS.forEach(hex => {
      const swatch = document.createElement('div');
      swatch.className = 'gsheet-swatch';
      swatch.style.backgroundColor = hex;
      swatch.title = hex;
      swatch.addEventListener('click', (e) => {
        e.stopPropagation();
        e.preventDefault();
        this.setStyleProperty('bgColor', hex);
        if (stripe) stripe.style.backgroundColor = hex;
        popover.classList.remove('active');
      });
      paletteGrid.appendChild(swatch);
    });

    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      e.preventDefault();
      popover.classList.toggle('active');
    });

    resetBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      e.preventDefault();
      this.resetFillColor();
      if (stripe) stripe.style.backgroundColor = 'transparent';
      popover.classList.remove('active');
    });

    customBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      e.preventDefault();
      colorPickerBg?.click();
    });

    colorPickerBg?.addEventListener('input', (e) => {
      const color = e.target.value;
      this.setStyleProperty('bgColor', color);
      if (stripe) stripe.style.backgroundColor = color;
      popover.classList.remove('active');
    });

    document.addEventListener('click', (e) => {
      if (wrap && !wrap.contains(e.target)) {
        popover.classList.remove('active');
      }
    });
  }

  resetFillColor() {
    const sheet = this.workbook.getActiveSheet();
    let sel = sheet.selection;
    if (!sel) {
      const active = sheet.activeCell || { col: 0, row: 0 };
      sel = sheet.selection = { startCol: active.col, startRow: active.row, endCol: active.col, endRow: active.row };
    }

    this.workbook.pushUndoState('Reset Fill Color');
    const minC = Math.min(sel.startCol, sel.endCol);
    const maxC = Math.max(sel.startCol, sel.endCol);
    const minR = Math.min(sel.startRow, sel.endRow);
    const maxR = Math.max(sel.startRow, sel.endRow);

    for (let r = minR; r <= maxR; r++) {
      for (let c = minC; c <= maxC; c++) {
        const cell = sheet.getCell(c, r);
        if (cell && cell.style) {
          delete cell.style.bgColor;
        }
      }
    }

    this.gridEngine.render();
    this.autoSaveEngine?.triggerSave();
    this.showToast('Fill color reset to none', 'info');
  }

  insertAutoSum(fnName) {
    const sheet = this.workbook.getActiveSheet();
    const active = sheet.activeCell || { col: 0, row: 0 };
    this.workbook.pushUndoState(`Insert ${fnName}`);

    const colLetter = FormulaEngine.colToLetter(active.col);
    const startRow = 1;
    const endRow = Math.max(1, active.row);
    const formula = `=${fnName}(${colLetter}${startRow}:${colLetter}${endRow})`;

    sheet.setCellValue(active.col, active.row, formula);
    this.gridEngine.render();
    this.showToast(`Inserted ${fnName} formula`, 'success');
  }

  toggleStyleProperty(prop) {
    const sheet = this.workbook.getActiveSheet();
    let sel = sheet.selection;
    if (!sel) {
      const active = sheet.activeCell || { col: 0, row: 0 };
      sel = sheet.selection = { startCol: active.col, startRow: active.row, endCol: active.col, endRow: active.row };
    }

    this.workbook.pushUndoState(`Toggle ${prop}`);
    const minC = Math.min(sel.startCol, sel.endCol);
    const maxC = Math.max(sel.startCol, sel.endCol);
    const minR = Math.min(sel.startRow, sel.endRow);
    const maxR = Math.max(sel.startRow, sel.endRow);

    const firstCell = sheet.getCell(minC, minR);
    const newVal = !(firstCell?.style?.[prop] || false);

    for (let r = minR; r <= maxR; r++) {
      for (let c = minC; c <= maxC; c++) {
        const cell = sheet.getOrCreateCell(c, r);
        cell.style = cell.style || {};
        cell.style[prop] = newVal;
      }
    }
    this.gridEngine.render();
    this.autoSaveEngine?.triggerSave();
  }

  setStyleProperty(prop, value) {
    const sheet = this.workbook.getActiveSheet();
    let sel = sheet.selection;
    if (!sel) {
      const active = sheet.activeCell || { col: 0, row: 0 };
      sel = sheet.selection = { startCol: active.col, startRow: active.row, endCol: active.col, endRow: active.row };
    }

    this.workbook.pushUndoState(`Set ${prop}`);
    const minC = Math.min(sel.startCol, sel.endCol);
    const maxC = Math.max(sel.startCol, sel.endCol);
    const minR = Math.min(sel.startRow, sel.endRow);
    const maxR = Math.max(sel.startRow, sel.endRow);

    for (let r = minR; r <= maxR; r++) {
      for (let c = minC; c <= maxC; c++) {
        const cell = sheet.getOrCreateCell(c, r);
        cell.style = cell.style || {};
        cell.style[prop] = value;
      }
    }
    this.gridEngine.render();
    this.autoSaveEngine?.triggerSave();
  }

  toggleTheme() {
    const isDark = document.body.getAttribute('data-theme') === 'dark';
    document.body.setAttribute('data-theme', isDark ? 'light' : 'dark');
    document.getElementById('btn-theme-toggle').textContent = isDark ? '🌙 Dark' : '☀️ Light';
    this.gridEngine.render();
  }

  bindSheetTabs() {
    // Add Sheet (+)
    document.getElementById('btn-add-sheet')?.addEventListener('click', () => {
      const newSheet = this.workbook.addSheet();
      this.workbook.activeSheetIndex = this.workbook.sheets.length - 1;
      this.updateSheetsTabBar();
      this.gridEngine.render();
      this.autoSaveEngine?.triggerSave();
    });

    // All Sheets (☰) Popup Navigator
    document.getElementById('btn-all-sheets')?.addEventListener('click', (e) => {
      e.stopPropagation();
      this.showAllSheetsMenu(e.currentTarget);
    });

    // Scroll Tabs (‹ and ›)
    const viewport = document.getElementById('sheet-tabs-viewport');
    document.getElementById('btn-scroll-sheets-left')?.addEventListener('click', () => {
      if (viewport) viewport.scrollBy({ left: -200, behavior: 'smooth' });
    });
    document.getElementById('btn-scroll-sheets-right')?.addEventListener('click', () => {
      if (viewport) viewport.scrollBy({ left: 200, behavior: 'smooth' });
    });
  }

  showAllSheetsMenu(targetBtn) {
    const existing = document.getElementById('all-sheets-popup');
    if (existing) { existing.remove(); return; }

    const popup = document.createElement('div');
    popup.id = 'all-sheets-popup';
    popup.className = 'dropdown-menu active';
    popup.style.position = 'fixed';
    popup.style.zIndex = '999999';

    const rect = targetBtn.getBoundingClientRect();
    popup.style.left = `${rect.left}px`;
    popup.style.bottom = `${window.innerHeight - rect.top + 6}px`;

    let html = `<div class="menu-header" style="padding:6px 12px;font-size:11px;font-weight:700;color:var(--text-muted);text-transform:uppercase;">All Sheets (${this.workbook.sheets.length})</div>`;
    this.workbook.sheets.forEach((sheet, idx) => {
      const isActive = idx === this.workbook.activeSheetIndex;
      html += `
        <div class="menu-item sheet-nav-item" data-idx="${idx}" style="display:flex;align-items:center;justify-content:space-between;padding:7px 12px;cursor:pointer;font-size:12px;font-weight:${isActive ? '600' : '400'};color:${isActive ? '#1a73e8' : 'var(--text-primary)'};background:${isActive ? '#e8f0fe' : 'transparent'};">
          <span>${sheet.name}</span>
          ${isActive ? '<span style="color:#1a73e8;">✓</span>' : ''}
        </div>
      `;
    });
    popup.innerHTML = html;

    popup.querySelectorAll('.sheet-nav-item').forEach(item => {
      item.addEventListener('click', () => {
        const idx = parseInt(item.dataset.idx, 10);
        this.workbook.activeSheetIndex = idx;
        this.updateSheetsTabBar();
        this.gridEngine.render();
        this.autoSaveEngine?.saveInstant();
        popup.remove();
      });
    });

    const closeHandler = (e) => {
      if (!popup.contains(e.target) && e.target !== targetBtn) {
        popup.remove();
        document.removeEventListener('mousedown', closeHandler);
      }
    };
    setTimeout(() => document.addEventListener('mousedown', closeHandler), 10);
    document.body.appendChild(popup);
  }

  showSheetTabContextMenu(e, sheetIdx) {
    e.preventDefault();
    e.stopPropagation();

    const existing = document.getElementById('sheet-tab-context-menu');
    if (existing) existing.remove();

    const sheet = this.workbook.sheets[sheetIdx];
    if (!sheet) return;

    const menu = document.createElement('div');
    menu.id = 'sheet-tab-context-menu';
    menu.className = 'dropdown-menu active';
    menu.style.position = 'fixed';
    menu.style.zIndex = '999999';

    const clickX = e.clientX || 100;
    menu.style.left = `${Math.max(10, Math.min(window.innerWidth - 180, clickX - 20))}px`;
    menu.style.bottom = `42px`; // Above the bottom sheet bar

    menu.innerHTML = `
      <div class="menu-item" id="menu-rename-sheet">
        <span>✏️</span>
        <span>Rename</span>
      </div>
      <div class="menu-item" id="menu-dup-sheet">
        <span>📋</span>
        <span>Duplicate</span>
      </div>
      ${this.workbook.sheets.length > 1 ? `
        <div class="menu-divider"></div>
        <div class="menu-item danger" id="menu-del-sheet">
          <span>🗑️</span>
          <span>Delete</span>
        </div>
      ` : ''}
    `;

    menu.querySelector('#menu-rename-sheet')?.addEventListener('click', () => {
      menu.remove();
      const newName = prompt('Rename sheet:', sheet.name);
      if (newName && newName.trim()) {
        sheet.name = newName.trim();
        this.updateSheetsTabBar();
        this.autoSaveEngine?.triggerSave();
      }
    });

    menu.querySelector('#menu-dup-sheet')?.addEventListener('click', () => {
      menu.remove();
      const dupData = sheet.serialize();
      const newSheet = this.workbook.addSheet(`${sheet.name} (Copy)`);
      newSheet.deserialize(dupData);
      newSheet.name = `${sheet.name} (Copy)`;
      this.workbook.activeSheetIndex = this.workbook.sheets.length - 1;
      this.updateSheetsTabBar();
      this.gridEngine.render();
      this.autoSaveEngine?.triggerSave();
    });

    menu.querySelector('#menu-del-sheet')?.addEventListener('click', () => {
      menu.remove();
      if (confirm(`Delete sheet '${sheet.name}'?`)) {
        this.workbook.removeSheet(sheetIdx);
        this.updateSheetsTabBar();
        this.gridEngine.render();
        this.autoSaveEngine?.triggerSave();
      }
    });

    const closeHandler = (evt) => {
      if (!menu.contains(evt.target)) {
        menu.remove();
        document.removeEventListener('mousedown', closeHandler);
      }
    };
    setTimeout(() => document.addEventListener('mousedown', closeHandler), 10);
    document.body.appendChild(menu);
  }

  updateSheetsTabBar() {
    const container = document.getElementById('sheets-tabs-container');
    if (!container) return;
    container.innerHTML = '';

    this.workbook.sheets.forEach((sheet, idx) => {
      const isActive = idx === this.workbook.activeSheetIndex;
      const tab = document.createElement('div');
      tab.className = `sheet-tab ${isActive ? 'active' : ''}`;
      tab.title = `${sheet.name} (Click to switch, double click to rename)`;

      tab.innerHTML = `
        <span class="sheet-tab-name">${sheet.name}</span>
        <span class="sheet-tab-arrow" title="Sheet Options">▾</span>
      `;

      // Click to activate
      tab.addEventListener('click', (e) => {
        if (e.target.classList.contains('sheet-tab-arrow')) {
          this.showSheetTabContextMenu(e, idx);
          return;
        }
        this.workbook.activeSheetIndex = idx;
        this.updateSheetsTabBar();
        this.gridEngine.render();
        this.autoSaveEngine?.saveInstant();
      });

      // Right Click Context Menu
      tab.addEventListener('contextmenu', (e) => {
        this.showSheetTabContextMenu(e, idx);
      });

      // Double Click to Rename
      tab.addEventListener('dblclick', () => {
        const newName = prompt('Rename sheet:', sheet.name);
        if (newName && newName.trim()) {
          sheet.name = newName.trim();
          this.updateSheetsTabBar();
          this.autoSaveEngine?.triggerSave();
        }
      });

      container.appendChild(tab);
    });
  }

  bindFileMenu() {
    const fileModal = document.getElementById('file-menu-modal');
    document.getElementById('btn-file-menu')?.addEventListener('click', () => {
      fileModal.classList.add('active');
    });
    document.getElementById('file-close-modal')?.addEventListener('click', () => fileModal.classList.remove('active'));

    document.getElementById('file-new')?.addEventListener('click', () => {
      this.autoSaveEngine.saveInstant();
      this.workbook = new WorkbookManager();
      this.formulaEngine.workbook = this.workbook;
      const newName = `Untitled Spreadsheet ${new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}`;
      this.autoSaveEngine.setCurrentProject('Untitled Spreadsheet', 'proj_' + Date.now());
      this.initBlankSheet();
      this.updateSheetsTabBar();
      this.autoSaveEngine.saveInstant();
      fileModal.classList.remove('active');
    });

    const fileInput = document.getElementById('file-input-open');
    document.getElementById('file-open')?.addEventListener('click', async () => {
      fileModal.classList.remove('active');

      // Native W3C File System Access API for Live Real-Time Excel Synchronization
      if (window.showOpenFilePicker) {
        try {
          const [handle] = await window.showOpenFilePicker({
            types: [{
              description: 'Spreadsheet Files',
              accept: {
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
                'application/vnd.ms-excel': ['.xls'],
                'text/csv': ['.csv']
              }
            }],
            multiple: false
          });
          if (handle) {
            this.activeFileSystemHandle = handle;
            const file = await handle.getFile();
            this.activeFileHandle = file;
            const reader = new FileReader();
            reader.onload = (evt) => {
              this.xlsxIO.importXLSX(evt.target.result, file.name);
              this.setupNativeFileHandleWatcher(handle);
            };
            reader.readAsArrayBuffer(file);
            return;
          }
        } catch (e) {
          // User cancelled picker or fallback
        }
      }
      fileInput.click();
    });

    fileInput?.addEventListener('change', (e) => {
      const file = e.target.files[0];
      if (!file) return;

      this.activeFileHandle = file;
      const fileNameLower = file.name.toLowerCase();

      if (fileNameLower.endsWith('.xlsx') || fileNameLower.endsWith('.xls') || fileNameLower.endsWith('.ods')) {
        const reader = new FileReader();
        reader.onload = (evt) => {
          this.xlsxIO.importXLSX(evt.target.result, file.name);
          this.setupLiveFileWatcher(file);
        };
        reader.readAsArrayBuffer(file);
      } else if (fileNameLower.endsWith('.0cell') || fileNameLower.endsWith('.json')) {
        const reader = new FileReader();
        reader.onload = (evt) => {
          this.xlsxIO.importFrom0Cell(evt.target.result);
        };
        reader.readAsText(file);
      } else {
        const reader = new FileReader();
        reader.onload = (evt) => {
          this.xlsxIO.importCSV(evt.target.result);
          this.setupLiveFileWatcher(file);
        };
        reader.readAsText(file);
      }

      e.target.value = '';
      fileModal.classList.remove('active');
    });

    document.getElementById('file-save-0cell')?.addEventListener('click', () => {
      this.xlsxIO.exportTo0Cell();
      fileModal.classList.remove('active');
    });

    document.getElementById('file-export-excel')?.addEventListener('click', () => {
      this.xlsxIO.exportToXLSX();
      fileModal.classList.remove('active');
    });

    document.getElementById('file-export-csv')?.addEventListener('click', () => {
      this.xlsxIO.exportToCSV();
      fileModal.classList.remove('active');
    });

    document.getElementById('file-print')?.addEventListener('click', () => {
      this.xlsxIO.printSheet();
      fileModal.classList.remove('active');
    });
  }

  openRecentManager() {
    const modal = document.getElementById('recent-manager-modal');
    if (!modal) return;
    const searchInput = document.getElementById('recent-search-input');
    if (searchInput) searchInput.value = '';
    this.autoSaveEngine.renderRecentManagerUI();
    modal.classList.add('active');
  }

  bindRecentModal() {
    const modal = document.getElementById('recent-manager-modal');
    document.getElementById('btn-recent-workbooks')?.addEventListener('click', () => {
      this.openRecentManager();
    });

    document.getElementById('recent-close-modal')?.addEventListener('click', () => {
      modal?.classList.remove('active');
    });

    document.getElementById('recent-footer-close')?.addEventListener('click', () => {
      modal?.classList.remove('active');
    });

    const searchInput = document.getElementById('recent-search-input');
    searchInput?.addEventListener('input', (e) => {
      this.autoSaveEngine.renderRecentManagerUI(e.target.value);
    });

    document.getElementById('btn-clear-recent-history')?.addEventListener('click', async () => {
      if (confirm('Are you sure you want to clear all recent workbooks history?')) {
        await ZeroStorage.clearAllProjects();
        this.autoSaveEngine.renderRecentManagerUI();
      }
    });
  }

  showToast(message, type = 'info') {
    return;
  }
}

window.addEventListener('DOMContentLoaded', () => {
  window.zeroCellApp = new ZeroCellApp();
  window.zeroCellApp.init();
});
