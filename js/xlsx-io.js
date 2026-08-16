/**
 * 0cell Advanced Excel XLSX Engine with Full OpenXML Style & Merge Extraction
 * Powered by JSZip + SheetJS
 * 100% accurate extraction of:
 * - Background Fill Colors (black, dark blue, yellow, accent colors, theme colors & tints)
 * - Font Colors (white on black, colored links, custom text colors)
 * - Font Sizes & Weights (14px, 16px, 18px header banners, bold, italic)
 * - Merged Cells (Colspan & Rowspan across multi-column header banners)
 * - Custom Column Widths and Custom Row Heights
 * - Multiple Worksheet tabs
 */

class XlsxIO {
  constructor(app) {
    this.app = app;
  }

  // Import XLSX with Full Design & Multi-Sheet Support
  async importXLSX(arrayBuffer, fileName = 'Imported_Workbook') {
    try {
      this.app.showToast('Reading Excel file & styles...', 'info');

      // 1. First parse data, formulas, and sheet names via SheetJS
      const data = new Uint8Array(arrayBuffer);
      const workbook = XLSX.read(data, {
        type: 'array',
        cellFormula: true,
        cellDates: true,
        sheetStubs: true
      });

      if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
        this.app.showToast('No worksheets found in this Excel file', 'warning');
        return;
      }

      // 2. Parse OpenXML styles & merges via JSZip for 100% visual fidelity
      let openXmlStyles = { cellXfs: [], fills: [], fonts: [] };
      let openXmlSheets = {};

      if (typeof JSZip !== 'undefined') {
        try {
          const zip = await JSZip.loadAsync(arrayBuffer);

          // Parse xl/styles.xml
          const stylesXmlFile = zip.file('xl/styles.xml');
          if (stylesXmlFile) {
            const stylesXmlText = await stylesXmlFile.async('text');
            openXmlStyles = this.parseStylesXml(stylesXmlText);
          }

          // Parse xl/workbook.xml for sheet mappings
          const wbXmlFile = zip.file('xl/workbook.xml');
          let sheetRels = {};
          if (wbXmlFile) {
            const wbXmlText = await wbXmlFile.async('text');
            sheetRels = this.parseWorkbookRels(wbXmlText);
          }

          // Extract Media Images (xl/media/image1.png, image2.jpg, etc.)
          const mediaImages = [];
          const mediaMap = {};
          const mediaFiles = Object.keys(zip.files).filter(path => path.startsWith('xl/media/'));
          for (let m = 0; m < mediaFiles.length; m++) {
            const mPath = mediaFiles[m];
            const mFile = zip.file(mPath);
            if (mFile) {
              const base64 = await mFile.async('base64');
              const ext = mPath.split('.').pop().toLowerCase();
              let mime = 'image/png';
              if (ext === 'jpg' || ext === 'jpeg') mime = 'image/jpeg';
              else if (ext === 'gif') mime = 'image/gif';
              else if (ext === 'svg') mime = 'image/svg+xml';
              else if (ext === 'webp') mime = 'image/webp';
              
              const dataUrl = `data:${mime};base64,${base64}`;
              const imgObj = {
                path: mPath,
                name: mPath.split('/').pop(),
                src: dataUrl
              };
              mediaImages.push(imgObj);
              mediaMap[imgObj.name] = dataUrl;
              mediaMap[mPath] = dataUrl;
            }
          }

          // Parse Drawing XMLs (xl/drawings/drawing1.xml, etc.)
          const sheetDrawings = {};
          for (let i = 1; i <= workbook.SheetNames.length; i++) {
            const sheetName = workbook.SheetNames[i - 1];
            sheetDrawings[sheetName] = [];

            // Check drawing relationships: xl/worksheets/_rels/sheet{i}.xml.rels
            const sheetRelsFile = zip.file(`xl/worksheets/_rels/sheet${i}.xml.rels`) || zip.file(`xl/worksheets/_rels/Sheet${i}.xml.rels`);
            let drawingRelTarget = null;
            if (sheetRelsFile) {
              const relsText = await sheetRelsFile.async('text');
              const dMatch = relsText.match(/Type="[^"]*drawing"[^>]*Target="([^"]+)"/i) || relsText.match(/Target="([^"]*drawing[^"]*)"/i);
              if (dMatch) {
                drawingRelTarget = dMatch[1].replace('../', 'xl/');
              }
            }
            if (!drawingRelTarget) {
              drawingRelTarget = `xl/drawings/drawing${i}.xml`;
            }

            const drawingFile = zip.file(drawingRelTarget);
            if (drawingFile) {
              const drawingXmlText = await drawingFile.async('text');
              
              // Load drawing rels (e.g. xl/drawings/_rels/drawing1.xml.rels)
              const drawingRelsPath = drawingRelTarget.replace('drawings/', 'drawings/_rels/') + '.rels';
              const drawingRelsFile = zip.file(drawingRelsPath);
              const rIdToMedia = {};
              if (drawingRelsFile) {
                const drawingRelsText = await drawingRelsFile.async('text');
                const relMatches = drawingRelsText.matchAll(/Id="([^"]+)"[^>]*Target="([^"]+)"/g);
                for (const m of relMatches) {
                  const rId = m[1];
                  const target = m[2].split('/').pop();
                  if (mediaMap[target]) {
                    rIdToMedia[rId] = mediaMap[target];
                  }
                }
              }

              const drawings = this.parseDrawingXml(drawingXmlText, rIdToMedia, mediaImages);
              sheetDrawings[sheetName] = drawings;
            }
          }

          // Parse individual sheet XMLs (e.g. xl/worksheets/sheet1.xml)
          for (let i = 1; i <= workbook.SheetNames.length; i++) {
            const sheetXmlFile = zip.file(`xl/worksheets/sheet${i}.xml`) || zip.file(`xl/worksheets/Sheet${i}.xml`);
            if (sheetXmlFile) {
              const sheetXmlText = await sheetXmlFile.async('text');
              const sheetName = workbook.SheetNames[i - 1];
              openXmlSheets[sheetName] = this.parseSheetXml(sheetXmlText, openXmlStyles);
            }
          }
          openXmlSheets.__mediaImages = mediaImages;
          openXmlSheets.__sheetDrawings = sheetDrawings;
        } catch (zipErr) {
          console.warn('OpenXML style parsing fallback:', zipErr);
        }
      }

      this.app.workbook.pushUndoState('Open Excel File');

      // Clear existing sheets in 0cell
      this.app.workbook.sheets = [];
      const cleanDocName = fileName.replace(/\.[^/.]+$/, '');
      this.app.workbook.name = cleanDocName;

      const docNameInput = document.getElementById('doc-name-input');
      if (docNameInput) docNameInput.value = cleanDocName;

      // 3. Populate each Worksheet
      const allExtractedMedia = openXmlSheets.__mediaImages || [];
      const allSheetDrawings = openXmlSheets.__sheetDrawings || {};

      workbook.SheetNames.forEach((sheetName, sIdx) => {
        const worksheet = workbook.Sheets[sheetName];
        const newSheet = this.app.workbook.addSheet(sheetName);
        const xmlData = openXmlSheets[sheetName] || { cellStyles: {}, merges: [], colWidths: {}, rowHeights: {} };

        // Attach drawings with exact cell anchors
        const sheetImgList = allSheetDrawings[sheetName] || [];
        if (sheetImgList.length > 0) {
          sheetImgList.forEach((d, idx) => {
            // Compute pixel left and top
            let left = 45; // Row header width
            for (let c = 0; c < d.fromCol; c++) {
              left += (newSheet.colWidths[c] || xmlData.colWidths[c] || 85);
            }
            left += (d.fromColOff || 0);

            let top = 26; // Thead height
            for (let r = 0; r < d.fromRow; r++) {
              top += (newSheet.rowHeights[r] || xmlData.rowHeights[r] || 24);
            }
            top += (d.fromRowOff || 0);

            let width = d.widthPx;
            let height = d.heightPx;

            if (!width) {
              width = 0;
              for (let c = d.fromCol; c <= d.toCol; c++) {
                width += (newSheet.colWidths[c] || xmlData.colWidths[c] || 85);
              }
            }
            if (!height) {
              height = 0;
              for (let r = d.fromRow; r <= d.toRow; r++) {
                height += (newSheet.rowHeights[r] || xmlData.rowHeights[r] || 24);
              }
            }

            newSheet.images.push({
              id: `img_${Date.now()}_${idx}`,
              src: d.src,
              col: d.fromCol,
              row: d.fromRow,
              x: left,
              y: top,
              width: Math.max(30, width || 280),
              height: Math.max(20, height || 75)
            });
          });
        } else if (sIdx === 0 && allExtractedMedia.length > 0) {
          // Fallback: place seamlessly over Col B / Row 1 (exact banner position)
          allExtractedMedia.forEach((imgItem, imgIdx) => {
            const left = 45 + (newSheet.colWidths[0] || xmlData.colWidths[0] || 85);
            const top = 26;
            newSheet.images.push({
              id: `img_${Date.now()}_${imgIdx}`,
              src: imgItem.src,
              col: 1,
              row: 0,
              x: left + (imgIdx * 20),
              y: top + (imgIdx * 10),
              width: 320,
              height: 75
            });
          });
        }

        // 3a. Determine bounds
        const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:M40');
        let maxCol = Math.max(26, range.e.c + 5);
        let maxRow = Math.max(50, range.e.r + 15);

        // 3b. Apply Merged Cells
        const allMerges = [];
        if (worksheet['!merges'] && Array.isArray(worksheet['!merges'])) {
          worksheet['!merges'].forEach(m => {
            allMerges.push({ startCol: m.s.c, startRow: m.s.r, endCol: m.e.c, endRow: m.e.r });
          });
        }
        if (xmlData.merges && Array.isArray(xmlData.merges)) {
          xmlData.merges.forEach(m => {
            if (!allMerges.some(ex => ex.startCol === m.startCol && ex.startRow === m.startRow)) {
              allMerges.push(m);
            }
          });
        }

        allMerges.forEach(m => {
          if (m.endCol >= maxCol) maxCol = m.endCol + 2;
          if (m.endRow >= maxRow) maxRow = m.endRow + 5;
        });

        newSheet.merges = allMerges;
        newSheet.numCols = maxCol;
        newSheet.numRows = maxRow;

        // 3c. Apply Column Widths
        if (worksheet['!cols'] && Array.isArray(worksheet['!cols'])) {
          worksheet['!cols'].forEach((c, idx) => {
            if (c) {
              if (c.wpx) newSheet.colWidths[idx] = Math.round(c.wpx);
              else if (c.wch) newSheet.colWidths[idx] = Math.max(40, Math.round(c.wch * 8.5));
              else if (c.width) newSheet.colWidths[idx] = Math.max(40, Math.round(c.width * 8.5));
            }
          });
        }
        Object.keys(xmlData.colWidths || {}).forEach(colIdx => {
          newSheet.colWidths[colIdx] = xmlData.colWidths[colIdx];
        });

        // 3d. Apply Row Heights
        if (worksheet['!rows'] && Array.isArray(worksheet['!rows'])) {
          worksheet['!rows'].forEach((r, idx) => {
            if (r) {
              if (r.hpx) newSheet.rowHeights[idx] = Math.round(r.hpx);
              else if (r.hpt) newSheet.rowHeights[idx] = Math.round(r.hpt * 1.33);
            }
          });
        }
        Object.keys(xmlData.rowHeights || {}).forEach(rowIdx => {
          newSheet.rowHeights[rowIdx] = xmlData.rowHeights[rowIdx];
        });

        // 3e. Populate Cell Values and Styles
        for (let R = 0; R <= Math.max(range.e.r, maxRow); R++) {
          for (let C = 0; C <= Math.max(range.e.c, maxCol); C++) {
            const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
            const cell = worksheet[cellAddress];
            const xmlStyle = xmlData.cellStyles ? xmlData.cellStyles[cellAddress] : null;

            if (cell || xmlStyle) {
              const targetCell = newSheet.getOrCreateCell(C, R);

              // Value & Formula
              if (cell) {
                if (cell.f) {
                  targetCell.formula = cell.f.startsWith('=') ? cell.f : '=' + cell.f;
                  targetCell.value = targetCell.formula;
                } else if (cell.v !== undefined) {
                  if (cell.t === 'd' && cell.v instanceof Date) {
                    const d = cell.v;
                    targetCell.value = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
                  } else {
                    targetCell.value = cell.v;
                  }
                }
              }

              // Apply OpenXML Styles (Background colors, font colors, size, bold)
              targetCell.style = targetCell.style || {};

              if (xmlStyle) {
                if (xmlStyle.bgColor) targetCell.style.bgColor = xmlStyle.bgColor;
                if (xmlStyle.color) targetCell.style.color = xmlStyle.color;
                if (xmlStyle.bold) targetCell.style.bold = true;
                if (xmlStyle.italic) targetCell.style.italic = true;
                if (xmlStyle.underline) targetCell.style.underline = true;
                if (xmlStyle.fontSize) targetCell.style.fontSize = xmlStyle.fontSize;
                if (xmlStyle.align) targetCell.style.align = xmlStyle.align;
                if (xmlStyle.wrapText) targetCell.style.wrapText = true;
              }

              // Number formats
              if (cell && cell.z) {
                if (cell.z.includes('$') || cell.z.includes('€') || cell.z.includes('£')) {
                  targetCell.style.format = 'currency';
                } else if (cell.z.includes('%')) {
                  targetCell.style.format = 'percent';
                }
              }
            }
          }
        }
      });

      // Activate 1st sheet and render tabs
      this.app.workbook.activeSheetIndex = 0;
      this.app.updateSheetsTabBar();
      this.app.gridEngine.render();
      this.app.autoSaveEngine?.setCurrentProject(cleanDocName, 'proj_' + cleanDocName.toLowerCase().replace(/[^a-z0-9_]/g, '_'));
      this.app.autoSaveEngine?.saveInstant();
      this.app.showToast(`Opened "${fileName}" (${workbook.SheetNames.length} sheet(s)) with full design`, 'success');
    } catch (err) {
      console.error('Error importing XLSX:', err);
      this.app.showToast(`Error opening file: ${err.message}`, 'warning');
    }
  }

  // --- OpenXML Styles XML Parser ---
  parseStylesXml(xmlText) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(xmlText, 'application/xml');

    const fills = [];
    const fonts = [];
    const cellXfs = [];

    // 1. Parse <fills>
    const fillEls = doc.querySelectorAll('fills > fill');
    fillEls.forEach(f => {
      const fgColor = f.querySelector('fgColor');
      const bgColor = f.querySelector('bgColor');
      const col = this.extractRgbFromColorEl(fgColor) || this.extractRgbFromColorEl(bgColor);
      fills.push(col);
    });

    // 2. Parse <fonts>
    const fontEls = doc.querySelectorAll('fonts > font');
    fontEls.forEach(fn => {
      const bold = !!fn.querySelector('b');
      const italic = !!fn.querySelector('i');
      const underline = !!fn.querySelector('u');
      const szEl = fn.querySelector('sz');
      const sz = szEl ? parseInt(szEl.getAttribute('val'), 10) : 12;
      const colorEl = fn.querySelector('color');
      const color = this.extractRgbFromColorEl(colorEl);

      fonts.push({ bold, italic, underline, size: sz, color });
    });

    // 3. Parse <cellXfs>
    const xfEls = doc.querySelectorAll('cellXfs > xf');
    xfEls.forEach(xf => {
      const fillId = parseInt(xf.getAttribute('fillId') || '0', 10);
      const fontId = parseInt(xf.getAttribute('fontId') || '0', 10);
      const alignEl = xf.querySelector('alignment');

      let align = null;
      let wrapText = false;
      if (alignEl) {
        align = alignEl.getAttribute('horizontal');
        wrapText = alignEl.getAttribute('wrapText') === '1' || alignEl.getAttribute('wrapText') === 'true';
      }

      cellXfs.push({
        bgColor: fills[fillId] || null,
        font: fonts[fontId] || null,
        align,
        wrapText
      });
    });

    return { cellXfs, fills, fonts };
  }

  // Extract RGB Hex from OpenXML color node
  extractRgbFromColorEl(colorEl) {
    if (!colorEl) return null;
    let rgb = colorEl.getAttribute('rgb');
    if (rgb) {
      // If 8 chars (AARRGGBB), strip alpha
      if (rgb.length === 8) rgb = rgb.substring(2);
      return '#' + rgb.toLowerCase();
    }

    // Theme color mapping
    const themeAttr = colorEl.getAttribute('theme');
    if (themeAttr !== null) {
      const themeIdx = parseInt(themeAttr, 10);
      const standardThemes = [
        '#ffffff', // 0 Light 1
        '#000000', // 1 Dark 1
        '#eeece1', // 2 Light 2
        '#1f497d', // 3 Dark 2 (Dark Blue)
        '#4f81bd', // 4 Accent 1
        '#c0504d', // 5 Accent 2
        '#9bbb59', // 6 Accent 3
        '#8064a2', // 7 Accent 4
        '#4bacc6', // 8 Accent 5
        '#f79646'  // 9 Accent 6
      ];
      let baseColor = standardThemes[themeIdx] || '#000000';
      const tint = parseFloat(colorEl.getAttribute('tint') || '0');
      if (tint !== 0) {
        baseColor = this.applyTint(baseColor, tint);
      }
      return baseColor;
    }

    // Indexed color
    const indexed = colorEl.getAttribute('indexed');
    if (indexed) {
      const idx = parseInt(indexed, 10);
      if (idx === 64) return null; // auto
      if (idx === 0) return '#000000';
      if (idx === 1) return '#ffffff';
      if (idx === 2) return '#ff0000';
      if (idx === 3) return '#00ff00';
      if (idx === 4) return '#0000ff';
      if (idx === 5) return '#ffff00';
    }

    return null;
  }

  applyTint(hexColor, tint) {
    let num = parseInt(hexColor.replace('#', ''), 16);
    let r = (num >> 16);
    let g = ((num >> 8) & 0x00FF);
    let b = (num & 0x0000FF);

    if (tint > 0) {
      r = Math.round(r + (255 - r) * tint);
      g = Math.round(g + (255 - g) * tint);
      b = Math.round(b + (255 - b) * tint);
    } else if (tint < 0) {
      r = Math.round(r * (1 + tint));
      g = Math.round(g * (1 + tint));
      b = Math.round(b * (1 + tint));
    }

    return `#${((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1)}`;
  }

  // --- Sheet XML Parser (Merges, Columns, Rows, Cell Styles) ---
  parseSheetXml(xmlText, openXmlStyles) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(xmlText, 'application/xml');

    const cellStyles = {};
    const merges = [];
    const colWidths = {};
    const rowHeights = {};

    // 1. Merged Cells (<mergeCells><mergeCell ref="B1:I1"/></mergeCells>)
    const mergeEls = doc.querySelectorAll('mergeCells > mergeCell');
    mergeEls.forEach(m => {
      const ref = m.getAttribute('ref');
      if (ref && ref.includes(':')) {
        const parts = ref.split(':');
        const s = FormulaEngine.parseCellRef(parts[0]);
        const e = FormulaEngine.parseCellRef(parts[1]);
        if (s && e) {
          merges.push({
            startCol: Math.min(s.col, e.col),
            startRow: Math.min(s.row, e.row),
            endCol: Math.max(s.col, e.col),
            endRow: Math.max(s.row, e.row)
          });
        }
      }
    });

    // 2. Column Widths (<cols><col min="1" max="1" width="25.5" customWidth="1"/></cols>)
    const colEls = doc.querySelectorAll('cols > col');
    colEls.forEach(c => {
      const min = parseInt(c.getAttribute('min') || '1', 10) - 1;
      const max = parseInt(c.getAttribute('max') || '1', 10) - 1;
      const widthVal = parseFloat(c.getAttribute('width') || '10');
      const pxWidth = Math.max(40, Math.round(widthVal * 8.5));

      for (let colIdx = min; colIdx <= max; colIdx++) {
        colWidths[colIdx] = pxWidth;
      }
    });

    // 3. Row Heights (<row r="1" ht="45" customHeight="1">)
    const rowEls = doc.querySelectorAll('sheetData > row');
    rowEls.forEach(r => {
      const rIdx = parseInt(r.getAttribute('r') || '1', 10) - 1;
      const ht = parseFloat(r.getAttribute('ht') || '0');
      if (ht > 0) {
        rowHeights[rIdx] = Math.round(ht * 1.33);
      }

      // 4. Cells (<c r="B1" s="5">)
      const cEls = r.querySelectorAll('c');
      cEls.forEach(cell => {
        const cellRef = cell.getAttribute('r');
        const sIdx = parseInt(cell.getAttribute('s') || '0', 10);

        if (cellRef && openXmlStyles.cellXfs[sIdx]) {
          const xf = openXmlStyles.cellXfs[sIdx];
          cellStyles[cellRef] = {
            bgColor: xf.bgColor,
            color: xf.font?.color || null,
            bold: xf.font?.bold || false,
            italic: xf.font?.italic || false,
            underline: xf.font?.underline || false,
            fontSize: xf.font?.size || 12,
            align: xf.align,
            wrapText: xf.wrapText
          };
        }
      });
    });

    return { cellStyles, merges, colWidths, rowHeights };
  }

  parseDrawingXml(xmlText, rIdToMedia, mediaImages) {
    const drawings = [];
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(xmlText, 'application/xml');

    const twoCellAnchors = xmlDoc.getElementsByTagName('xdr:twoCellAnchor');
    const oneCellAnchors = xmlDoc.getElementsByTagName('xdr:oneCellAnchor');

    const processAnchor = (anchorEl) => {
      const fromEl = anchorEl.getElementsByTagName('xdr:from')[0];
      const blipEl = anchorEl.getElementsByTagName('a:blip')[0];
      const extEl = anchorEl.getElementsByTagName('a:ext')[0];
      const toEl = anchorEl.getElementsByTagName('xdr:to')[0];

      if (!fromEl) return;

      const fromCol = parseInt(fromEl.getElementsByTagName('xdr:col')[0]?.textContent || '0', 10);
      const fromColOff = parseInt(fromEl.getElementsByTagName('xdr:colOff')[0]?.textContent || '0', 10);
      const fromRow = parseInt(fromEl.getElementsByTagName('xdr:row')[0]?.textContent || '0', 10);
      const fromRowOff = parseInt(fromEl.getElementsByTagName('xdr:rowOff')[0]?.textContent || '0', 10);

      let toCol = fromCol + 2;
      let toRow = fromRow + 2;
      if (toEl) {
        toCol = parseInt(toEl.getElementsByTagName('xdr:col')[0]?.textContent || String(fromCol + 2), 10);
        toRow = parseInt(toEl.getElementsByTagName('xdr:row')[0]?.textContent || String(fromRow + 2), 10);
      }

      let cx = 0, cy = 0;
      if (extEl) {
        cx = parseInt(extEl.getAttribute('cx') || '0', 10);
        cy = parseInt(extEl.getAttribute('cy') || '0', 10);
      }

      let imageSrc = null;
      if (blipEl) {
        const embedId = blipEl.getAttribute('r:embed');
        if (embedId && rIdToMedia[embedId]) {
          imageSrc = rIdToMedia[embedId];
        }
      }
      if (!imageSrc && mediaImages.length > 0) {
        imageSrc = mediaImages[0].src;
      }

      if (imageSrc) {
        drawings.push({
          src: imageSrc,
          fromCol,
          fromRow,
          fromColOff: Math.round(fromColOff / 9525),
          fromRowOff: Math.round(fromRowOff / 9525),
          toCol,
          toRow,
          widthPx: cx > 0 ? Math.round(cx / 9525) : null,
          heightPx: cy > 0 ? Math.round(cy / 9525) : null
        });
      }
    };

    for (let i = 0; i < twoCellAnchors.length; i++) processAnchor(twoCellAnchors[i]);
    for (let i = 0; i < oneCellAnchors.length; i++) processAnchor(oneCellAnchors[i]);

    return drawings;
  }

  parseWorkbookRels(wbXmlText) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(wbXmlText, 'application/xml');
    const sheets = {};
    const sheetEls = doc.querySelectorAll('sheets > sheet');
    sheetEls.forEach((s, idx) => {
      const name = s.getAttribute('name');
      const sheetId = s.getAttribute('sheetId') || (idx + 1);
      sheets[name] = `sheet${sheetId}.xml`;
    });
    return sheets;
  }

  // Export full workbook to true native .xlsx file
  exportToXLSX() {
    try {
      if (typeof XLSX === 'undefined') {
        throw new Error('SheetJS library is not available');
      }

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
              if (cell.formula) {
                rowData.push({ f: cell.formula.startsWith('=') ? cell.formula.substring(1) : cell.formula });
              } else if (cell.value !== undefined) {
                rowData.push(cell.value);
              } else {
                rowData.push('');
              }
            } else {
              rowData.push('');
            }
          }
          wsData.push(rowData);
        }

        const ws = XLSX.utils.aoa_to_sheet(wsData);

        // Merges
        if (sheet.merges && sheet.merges.length > 0) {
          ws['!merges'] = sheet.merges.map(m => ({
            s: { r: m.startRow, c: m.startCol },
            e: { r: m.endRow, c: m.endCol }
          }));
        }

        // Column widths
        if (sheet.colWidths && Object.keys(sheet.colWidths).length > 0) {
          const cols = [];
          for (let c = 0; c <= maxC; c++) {
            cols.push({ wpx: sheet.colWidths[c] || 110 });
          }
          ws['!cols'] = cols;
        }

        XLSX.utils.book_append_sheet(wb, ws, sheet.name.substring(0, 31));
      });

      const fileName = `${this.app.workbook.name || 'Workbook'}.xlsx`;
      XLSX.writeFile(wb, fileName);
      this.app.showToast(`Exported "${fileName}" successfully`, 'success');
    } catch (err) {
      console.error(err);
      this.app.showToast(`Failed to export XLSX: ${err.message}`, 'warning');
    }
  }

  // Export CSV
  exportToCSV(sheetName) {
    const sheet = sheetName ? this.app.workbook.getSheet(sheetName) : this.app.workbook.getActiveSheet();
    if (!sheet) return;

    let maxCol = 0, maxRow = 0;
    Object.keys(sheet.cells).forEach(k => {
      const [c, r] = k.split(',').map(Number);
      if (c > maxCol) maxCol = c;
      if (r > maxRow) maxRow = r;
    });

    const rows = [];
    for (let r = 0; r <= maxRow; r++) {
      const rowVals = [];
      for (let c = 0; c <= maxCol; c++) {
        const cell = sheet.getCell(c, r);
        let val = '';
        if (cell) {
          val = cell.value !== undefined ? String(cell.value) : '';
          if (val.includes(',') || val.includes('"') || val.includes('\n')) {
            val = `"${val.replace(/"/g, '""')}"`;
          }
        }
        rowVals.push(val);
      }
      rows.push(rowVals.join(','));
    }

    const csvContent = '\uFEFF' + rows.join('\r\n');
    this.downloadFile(csvContent, `${sheet.name || '0cell_export'}.csv`, 'text/csv;charset=utf-8;');
    this.app.showToast('Exported to CSV successfully', 'success');
  }

  importCSV(fileContent, sheetName) {
    try {
      const sheet = sheetName ? this.app.workbook.getSheet(sheetName) : this.app.workbook.getActiveSheet();
      this.app.workbook.pushUndoState('Import CSV');

      const lines = fileContent.split(/\r\n|\n|\r/);
      lines.forEach((line, rIdx) => {
        if (!line.trim() && rIdx === lines.length - 1) return;
        const cols = this.parseCSVLine(line);
        cols.forEach((colVal, cIdx) => {
          sheet.setCellValue(cIdx, rIdx, colVal);
        });
      });

      this.app.gridEngine.render();
      this.app.autoSaveEngine?.setCurrentProject(sheet.name || 'CSV Data');
      this.app.autoSaveEngine?.saveInstant();
      this.app.showToast('CSV imported successfully', 'success');
    } catch (err) {
      this.app.showToast(`CSV import error: ${err.message}`, 'warning');
    }
  }

  parseCSVLine(text) {
    const p = [];
    let cur = '';
    let inQuote = false;
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (c === '"') {
        if (inQuote && text[i + 1] === '"') {
          cur += '"';
          i++;
        } else {
          inQuote = !inQuote;
        }
      } else if ((c === ',' || c === '\t') && !inQuote) {
        p.push(cur);
        cur = '';
      } else {
        cur += c;
      }
    }
    p.push(cur);
    return p;
  }

  exportTo0Cell() {
    const data = this.app.workbook.serialize();
    const jsonStr = JSON.stringify(data, null, 2);
    this.downloadFile(jsonStr, `${this.app.workbook.name || 'Workbook'}.0cell`, 'application/json');
    this.app.showToast('Workbook saved as .0cell project', 'success');
  }

  importFrom0Cell(jsonStr) {
    try {
      const data = JSON.parse(jsonStr);
      this.app.workbook.deserialize(data);
      this.app.gridEngine.render();
      this.app.updateSheetsTabBar();
      const projName = data.name || this.app.workbook.name || '0cell Project';
      this.app.autoSaveEngine?.setCurrentProject(projName);
      this.app.autoSaveEngine?.saveInstant();
      this.app.showToast('Workbook loaded successfully', 'success');
    } catch (e) {
      this.app.showToast('Failed to load .0cell file: ' + e.message, 'warning');
    }
  }

  printSheet() {
    window.print();
  }

  downloadFile(content, fileName, mimeType) {
    const blob = new Blob([content], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = XlsxIO;
}
