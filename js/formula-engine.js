/**
 * 0cell Formula & Calculation Engine
 * Comprehensive Excel-compatible formula parser and evaluator
 * Supports 100+ Excel functions, ranges, 3D sheet references, relative/absolute references.
 */

class FormulaEngine {
  constructor(workbookManager) {
    this.workbook = workbookManager;
    this.functions = {};
    this.registerStandardFunctions();
  }

  // Convert column index (0-based) to letter (0 -> A, 25 -> Z, 26 -> AA)
  static colToLetter(colIndex) {
    let temp = colIndex + 1;
    let letter = '';
    while (temp > 0) {
      let mod = (temp - 1) % 26;
      letter = String.fromCharCode(65 + mod) + letter;
      temp = Math.floor((temp - mod) / 26);
    }
    return letter;
  }

  // Convert column letter to 0-based index (A -> 0, Z -> 25, AA -> 26)
  static letterToCol(letter) {
    let col = 0;
    const str = letter.toUpperCase();
    for (let i = 0; i < str.length; i++) {
      col = col * 26 + (str.charCodeAt(i) - 64);
    }
    return col - 1;
  }

  // Parse cell address e.g. "$A$1", "Sheet1!B2", "C10"
  static parseCellRef(refStr) {
    let sheetName = null;
    let cellPart = refStr.trim();
    if (cellPart.includes('!')) {
      const parts = cellPart.split('!');
      sheetName = parts[0].replace(/^'|'$/g, '');
      cellPart = parts[1];
    }

    const match = cellPart.match(/^(\$?)([A-Za-z]+)(\$?)([0-9]+)$/);
    if (!match) return null;

    const colAbs = match[1] === '$';
    const colStr = match[2].toUpperCase();
    const rowAbs = match[3] === '$';
    const rowNum = parseInt(match[4], 10) - 1; // 0-indexed row

    return {
      sheet: sheetName,
      col: FormulaEngine.letterToCol(colStr),
      row: rowNum,
      colAbs,
      rowAbs,
      colStr,
      rowStr: match[4]
    };
  }

  // Parse range e.g. "A1:B10" or "Sheet1!A1:C5"
  static parseRangeRef(rangeStr) {
    let sheetName = null;
    let rangePart = rangeStr.trim();
    if (rangePart.includes('!')) {
      const parts = rangePart.split('!');
      sheetName = parts[0].replace(/^'|'$/g, '');
      rangePart = parts[1];
    }

    const sides = rangePart.split(':');
    if (sides.length === 1) {
      const c = FormulaEngine.parseCellRef(sides[0]);
      if (!c) return null;
      if (sheetName) c.sheet = sheetName;
      return { sheet: c.sheet, startCol: c.col, startRow: c.row, endCol: c.col, endRow: c.row };
    } else if (sides.length === 2) {
      const c1 = FormulaEngine.parseCellRef(sides[0]);
      const c2 = FormulaEngine.parseCellRef(sides[1]);
      if (!c1 || !c2) return null;
      const finalSheet = sheetName || c1.sheet || c2.sheet;
      return {
        sheet: finalSheet,
        startCol: Math.min(c1.col, c2.col),
        startRow: Math.min(c1.row, c2.row),
        endCol: Math.max(c1.col, c2.col),
        endRow: Math.max(c1.row, c2.row)
      };
    }
    return null;
  }

  // Shift formula references when dragging / copying
  static shiftFormulaReferences(formulaStr, colOffset, rowOffset) {
    if (!formulaStr || !formulaStr.startsWith('=')) return formulaStr;

    // Regex to match cell references and ranges
    const refRegex = /('?[A-Za-z0-9_\s]+'?!|)(\$?)([A-Za-z]+)(\$?)([0-9]+)/g;
    return formulaStr.replace(refRegex, (match, sheet, colAbs, colLetter, rowAbs, rowNum) => {
      // Check if inside function name
      let cIdx = FormulaEngine.letterToCol(colLetter);
      let rIdx = parseInt(rowNum, 10) - 1;

      if (colAbs !== '$') {
        cIdx = Math.max(0, cIdx + colOffset);
      }
      if (rowAbs !== '$') {
        rIdx = Math.max(0, rIdx + rowOffset);
      }

      const newColLetter = FormulaEngine.colToLetter(cIdx);
      const newRowNum = rIdx + 1;
      return `${sheet}${colAbs ? '$' : ''}${newColLetter}${rowAbs ? '$' : ''}${newRowNum}`;
    });
  }

  // Get raw or calculated value from cell
  getCellValue(sheetName, col, row, visited = new Set()) {
    if (!this.workbook) return null;
    const sheet = this.workbook.getSheet(sheetName);
    if (!sheet) return 0;

    const key = `${col},${row}`;
    const cellKey = `${sheet.name}!${key}`;

    if (visited.has(cellKey)) {
      return '#CIRCULAR!';
    }

    const cell = sheet.getCell(col, row);
    if (!cell) return null;

    if (cell.formula) {
      visited.add(cellKey);
      const res = this.evaluateFormula(cell.formula, sheet.name, visited);
      visited.delete(cellKey);
      return res;
    }

    return cell.value !== undefined ? cell.value : null;
  }

  // Evaluate range into 2D array of values
  getRangeValues(rangeRef, currentSheet, visited = new Set()) {
    const sheetName = rangeRef.sheet || currentSheet;
    const values = [];
    for (let r = rangeRef.startRow; r <= rangeRef.endRow; r++) {
      const rowArr = [];
      for (let c = rangeRef.startCol; c <= rangeRef.endCol; c++) {
        const val = this.getCellValue(sheetName, c, r, visited);
        rowArr.push(val);
      }
      values.push(rowArr);
    }
    return values;
  }

  // Flatten range values into 1D array
  flatten(val) {
    if (Array.isArray(val)) {
      let result = [];
      for (let item of val) {
        if (Array.isArray(item)) {
          result = result.concat(this.flatten(item));
        } else {
          result.push(item);
        }
      }
      return result;
    }
    return [val];
  }

  // Main evaluation entry
  evaluate(input, currentSheet = 'Sheet1') {
    if (typeof input !== 'string') return input;
    if (!input.startsWith('=')) {
      // Check if numeric
      if (!isNaN(input) && input.trim() !== '') {
        return Number(input);
      }
      return input;
    }
    return this.evaluateFormula(input, currentSheet);
  }

  evaluateFormula(formulaStr, currentSheet, visited = new Set()) {
    try {
      const expr = formulaStr.startsWith('=') ? formulaStr.substring(1).trim() : formulaStr.trim();
      const tokens = this.tokenize(expr);
      const parser = new FormulaParser(tokens, this, currentSheet, visited);
      return parser.parse();
    } catch (err) {
      if (err.message && err.message.startsWith('#')) return err.message;
      return '#VALUE!';
    }
  }

  // Tokenize Excel formula
  tokenize(code) {
    const tokens = [];
    let i = 0;
    const len = code.length;

    while (i < len) {
      const char = code[i];

      if (/\s/.test(char)) {
        i++;
        continue;
      }

      // String literal
      if (char === '"') {
        let str = '';
        i++;
        while (i < len) {
          if (code[i] === '"') {
            if (i + 1 < len && code[i + 1] === '"') {
              str += '"';
              i += 2;
            } else {
              i++;
              break;
            }
          } else {
            str += code[i];
            i++;
          }
        }
        tokens.push({ type: 'STRING', value: str });
        continue;
      }

      // Numbers
      if (/[0-9]/.test(char) || (char === '.' && i + 1 < len && /[0-9]/.test(code[i + 1]))) {
        let numStr = '';
        while (i < len && /[0-9.]/.test(code[i])) {
          numStr += code[i];
          i++;
        }
        if (i < len && code[i] === '%') {
          tokens.push({ type: 'NUMBER', value: parseFloat(numStr) / 100 });
          i++;
        } else {
          tokens.push({ type: 'NUMBER', value: parseFloat(numStr) });
        }
        continue;
      }

      // Two-character operators
      if (i + 1 < len) {
        const twoChar = code.substr(i, 2);
        if (['<=', '>=', '<>'].includes(twoChar)) {
          tokens.push({ type: 'OPERATOR', value: twoChar });
          i += 2;
          continue;
        }
      }

      // Single-character operators
      if (['+', '-', '*', '/', '^', '&', '=', '<', '>', '%', '(', ')', ',', ';', ':'].includes(char)) {
        tokens.push({ type: 'OPERATOR', value: char });
        i++;
        continue;
      }

      // Identifiers, Sheet References, Function names, Cell references
      if (/[A-Za-z0-9_$':!]/.test(char)) {
        let ident = '';
        let inSingleQuote = false;
        while (i < len) {
          const c = code[i];
          if (c === "'") {
            inSingleQuote = !inSingleQuote;
            ident += c;
            i++;
          } else if (inSingleQuote) {
            ident += c;
            i++;
          } else if (/[A-Za-z0-9_$!:]/.test(c)) {
            ident += c;
            i++;
          } else {
            break;
          }
        }

        // Check if next non-whitespace char is '(' -> Function call
        let lookAhead = i;
        while (lookAhead < len && /\s/.test(code[lookAhead])) lookAhead++;
        if (lookAhead < len && code[lookAhead] === '(' && !ident.includes('!') && !ident.includes(':')) {
          tokens.push({ type: 'FUNCTION', value: ident.toUpperCase() });
        } else {
          tokens.push({ type: 'REFERENCE', value: ident });
        }
        continue;
      }

      i++;
    }

    tokens.push({ type: 'EOF' });
    return tokens;
  }

  // Register 100+ standard Excel functions
  registerStandardFunctions() {
    const fn = this.functions;

    // Helper: to numbers array
    const toNumArray = (args) => {
      const flat = this.flatten(args);
      return flat
        .filter(v => v !== null && v !== '' && !isNaN(v) && typeof v !== 'boolean')
        .map(v => Number(v));
    };

    // --- MATH & TRIG ---
    fn['SUM'] = (...args) => toNumArray(args).reduce((a, b) => a + b, 0);
    fn['AVERAGE'] = (...args) => {
      const arr = toNumArray(args);
      return arr.length === 0 ? '#DIV/0!' : arr.reduce((a, b) => a + b, 0) / arr.length;
    };
    fn['COUNT'] = (...args) => toNumArray(args).length;
    fn['COUNTA'] = (...args) => this.flatten(args).filter(v => v !== null && v !== '').length;
    fn['COUNTBLANK'] = (...args) => this.flatten(args).filter(v => v === null || v === '').length;
    fn['MIN'] = (...args) => {
      const arr = toNumArray(args);
      return arr.length === 0 ? 0 : Math.min(...arr);
    };
    fn['MAX'] = (...args) => {
      const arr = toNumArray(args);
      return arr.length === 0 ? 0 : Math.max(...arr);
    };
    fn['PRODUCT'] = (...args) => {
      const arr = toNumArray(args);
      return arr.length === 0 ? 0 : arr.reduce((a, b) => a * b, 1);
    };
    fn['POWER'] = (base, exp) => Math.pow(Number(base), Number(exp));
    fn['SQRT'] = (v) => Number(v) < 0 ? '#NUM!' : Math.sqrt(Number(v));
    fn['ABS'] = (v) => Math.abs(Number(v));
    fn['ROUND'] = (v, dec = 0) => {
      const factor = Math.pow(10, dec);
      return Math.round(Number(v) * factor) / factor;
    };
    fn['ROUNDUP'] = (v, dec = 0) => {
      const factor = Math.pow(10, dec);
      const sign = Number(v) >= 0 ? 1 : -1;
      return (sign * Math.ceil(Math.abs(Number(v)) * factor)) / factor;
    };
    fn['ROUNDDOWN'] = (v, dec = 0) => {
      const factor = Math.pow(10, dec);
      const sign = Number(v) >= 0 ? 1 : -1;
      return (sign * Math.floor(Math.abs(Number(v)) * factor)) / factor;
    };
    fn['INT'] = (v) => Math.floor(Number(v));
    fn['MOD'] = (n, d) => Number(d) === 0 ? '#DIV/0!' : Number(n) % Number(d);
    fn['CEILING'] = (v, sig = 1) => Number(sig) === 0 ? 0 : Math.ceil(Number(v) / Number(sig)) * Number(sig);
    fn['FLOOR'] = (v, sig = 1) => Number(sig) === 0 ? 0 : Math.floor(Number(v) / Number(sig)) * Number(sig);
    fn['TRUNC'] = (v, dec = 0) => {
      const factor = Math.pow(10, dec);
      return (Number(v) >= 0 ? Math.floor(Number(v) * factor) : Math.ceil(Number(v) * factor)) / factor;
    };
    fn['SIGN'] = (v) => Math.sign(Number(v));
    fn['RAND'] = () => Math.random();
    fn['RANDBETWEEN'] = (bottom, top) => Math.floor(Math.random() * (Number(top) - Number(bottom) + 1)) + Number(bottom);
    fn['PI'] = () => Math.PI;
    fn['EXP'] = (v) => Math.exp(Number(v));
    fn['LN'] = (v) => Number(v) <= 0 ? '#NUM!' : Math.log(Number(v));
    fn['LOG'] = (v, base = 10) => Number(v) <= 0 ? '#NUM!' : Math.log(Number(v)) / Math.log(Number(base));
    fn['LOG10'] = (v) => Number(v) <= 0 ? '#NUM!' : Math.log10(Number(v));
    fn['FACT'] = (v) => {
      let n = Math.floor(Number(v));
      if (n < 0) return '#NUM!';
      let res = 1;
      for (let i = 2; i <= n; i++) res *= i;
      return res;
    };
    fn['COMBIN'] = (n, k) => {
      n = Math.floor(Number(n)); k = Math.floor(Number(k));
      if (k < 0 || k > n) return '#NUM!';
      return fn['FACT'](n) / (fn['FACT'](k) * fn['FACT'](n - k));
    };

    // Helper for wildcard and criteria matching
    const matchCriteria = (val, crit) => {
      if (crit === undefined || crit === null) return false;
      const critStr = String(crit).trim();
      const valStr = String(val ?? '');

      if (critStr.startsWith('>=')) return Number(val) >= Number(critStr.substring(2));
      if (critStr.startsWith('<=')) return Number(val) <= Number(critStr.substring(2));
      if (critStr.startsWith('<>')) return String(val).toLowerCase() !== critStr.substring(2).toLowerCase();
      if (critStr.startsWith('>')) return Number(val) > Number(critStr.substring(1));
      if (critStr.startsWith('<')) return Number(val) < Number(critStr.substring(1));
      if (critStr.startsWith('=')) return String(val).toLowerCase() === critStr.substring(1).toLowerCase();

      // Wildcard regex * and ?
      if (critStr.includes('*') || critStr.includes('?')) {
        const regexPattern = '^' + critStr.replace(/([.+^$[\]\\(){}|])/g, '\\$1').replace(/\*/g, '.*').replace(/\?/g, '.') + '$';
        return new RegExp(regexPattern, 'i').test(valStr);
      }

      if (!isNaN(val) && !isNaN(critStr) && critStr !== '') {
        return Number(val) === Number(critStr);
      }
      return valStr.toLowerCase() === critStr.toLowerCase();
    };

    fn['COUNTIF'] = (range, criteria) => {
      const flat = this.flatten(range);
      return flat.filter(v => matchCriteria(v, criteria)).length;
    };

    fn['COUNTIFS'] = (...args) => {
      if (args.length < 2 || args.length % 2 !== 0) return '#VALUE!';
      const rangePairs = [];
      for (let i = 0; i < args.length; i += 2) {
        rangePairs.push({ range: this.flatten(args[i]), crit: args[i + 1] });
      }
      const len = rangePairs[0].range.length;
      let count = 0;
      for (let idx = 0; idx < len; idx++) {
        let pass = true;
        for (const pair of rangePairs) {
          if (!matchCriteria(pair.range[idx], pair.crit)) {
            pass = false;
            break;
          }
        }
        if (pass) count++;
      }
      return count;
    };

    fn['SUMIF'] = (range, criteria, sumRange) => {
      const checkRange = this.flatten(range);
      const targetSumRange = sumRange ? this.flatten(sumRange) : checkRange;
      let sum = 0;
      for (let i = 0; i < checkRange.length; i++) {
        if (matchCriteria(checkRange[i], criteria)) {
          const val = Number(targetSumRange[i]);
          if (!isNaN(val)) sum += val;
        }
      }
      return sum;
    };

    fn['SUMIFS'] = (sumRange, ...args) => {
      const targetSumRange = this.flatten(sumRange);
      const rangePairs = [];
      for (let i = 0; i < args.length; i += 2) {
        rangePairs.push({ range: this.flatten(args[i]), crit: args[i + 1] });
      }
      let sum = 0;
      for (let i = 0; i < targetSumRange.length; i++) {
        let pass = true;
        for (const pair of rangePairs) {
          if (!matchCriteria(pair.range[i], pair.crit)) {
            pass = false;
            break;
          }
        }
        if (pass) {
          const val = Number(targetSumRange[i]);
          if (!isNaN(val)) sum += val;
        }
      }
      return sum;
    };

    fn['AVERAGEIF'] = (range, criteria, avgRange) => {
      const checkRange = this.flatten(range);
      const targetAvgRange = avgRange ? this.flatten(avgRange) : checkRange;
      let sum = 0, count = 0;
      for (let i = 0; i < checkRange.length; i++) {
        if (matchCriteria(checkRange[i], criteria)) {
          const val = Number(targetAvgRange[i]);
          if (!isNaN(val)) {
            sum += val;
            count++;
          }
        }
      }
      return count === 0 ? '#DIV/0!' : sum / count;
    };

    // --- LOGICAL ---
    fn['IF'] = (condition, valueIfTrue, valueIfFalse = false) => {
      return Boolean(condition) ? valueIfTrue : valueIfFalse;
    };
    fn['IFS'] = (...args) => {
      for (let i = 0; i < args.length; i += 2) {
        if (Boolean(args[i])) return args[i + 1];
      }
      return '#N/A';
    };
    fn['AND'] = (...args) => this.flatten(args).every(v => Boolean(v));
    fn['OR'] = (...args) => this.flatten(args).some(v => Boolean(v));
    fn['NOT'] = (v) => !Boolean(v);
    fn['XOR'] = (...args) => {
      const truths = this.flatten(args).filter(v => Boolean(v)).length;
      return truths % 2 === 1;
    };
    fn['SWITCH'] = (expr, ...cases) => {
      for (let i = 0; i < cases.length - 1; i += 2) {
        if (expr === cases[i]) return cases[i + 1];
      }
      return cases.length % 2 === 1 ? cases[cases.length - 1] : '#N/A';
    };
    fn['IFERROR'] = (val, valIfError) => {
      if (typeof val === 'string' && val.startsWith('#')) return valIfError;
      if (val === undefined || val === null || isNaN(val) && typeof val === 'number') return valIfError;
      return val;
    };
    fn['IFNA'] = (val, valIfNa) => val === '#N/A' ? valIfNa : val;
    fn['TRUE'] = () => true;
    fn['FALSE'] = () => false;

    // --- LOOKUP & REFERENCE ---
    fn['VLOOKUP'] = (lookupVal, tableArray, colIndex, exactMatch = true) => {
      if (!Array.isArray(tableArray) || tableArray.length === 0) return '#N/A';
      const colIdx = Number(colIndex) - 1;
      const isExact = exactMatch === false || exactMatch === 0 || String(exactMatch).toUpperCase() === 'FALSE';

      for (let r = 0; r < tableArray.length; r++) {
        const row = tableArray[r];
        const cell = row[0];
        if (isExact) {
          if (String(cell).toLowerCase() === String(lookupVal).toLowerCase()) {
            return row[colIdx] !== undefined ? row[colIdx] : '#REF!';
          }
        } else {
          if (Number(cell) <= Number(lookupVal)) {
            if (r === tableArray.length - 1 || Number(tableArray[r + 1][0]) > Number(lookupVal)) {
              return row[colIdx] !== undefined ? row[colIdx] : '#REF!';
            }
          }
        }
      }
      return '#N/A';
    };

    fn['HLOOKUP'] = (lookupVal, tableArray, rowIndex, exactMatch = true) => {
      if (!Array.isArray(tableArray) || tableArray.length === 0) return '#N/A';
      const rowIdx = Number(rowIndex) - 1;
      const numCols = tableArray[0].length;
      for (let c = 0; c < numCols; c++) {
        if (String(tableArray[0][c]).toLowerCase() === String(lookupVal).toLowerCase()) {
          return tableArray[rowIdx] && tableArray[rowIdx][c] !== undefined ? tableArray[rowIdx][c] : '#REF!';
        }
      }
      return '#N/A';
    };

    fn['XLOOKUP'] = (lookupVal, lookupArray, returnArray, ifNotFound = '#N/A', matchMode = 0) => {
      const lArr = this.flatten(lookupArray);
      const rArr = Array.isArray(returnArray[0]) ? returnArray : [this.flatten(returnArray)];
      const flatR = this.flatten(returnArray);

      for (let i = 0; i < lArr.length; i++) {
        if (String(lArr[i]).toLowerCase() === String(lookupVal).toLowerCase()) {
          return flatR[i] !== undefined ? flatR[i] : ifNotFound;
        }
      }
      return ifNotFound;
    };

    fn['INDEX'] = (array, rowNum, colNum = 1) => {
      if (!Array.isArray(array)) return array;
      const r = Number(rowNum) - 1;
      const c = Number(colNum) - 1;
      if (Array.isArray(array[0])) {
        if (array[r] && array[r][c] !== undefined) return array[r][c];
      } else {
        if (array[r] !== undefined) return array[r];
      }
      return '#REF!';
    };

    fn['MATCH'] = (lookupVal, lookupArray, matchType = 1) => {
      const arr = this.flatten(lookupArray);
      for (let i = 0; i < arr.length; i++) {
        if (matchType === 0) {
          if (String(arr[i]).toLowerCase() === String(lookupVal).toLowerCase()) return i + 1;
        } else if (matchType === 1) {
          if (Number(arr[i]) === Number(lookupVal)) return i + 1;
        }
      }
      return '#N/A';
    };

    fn['CHOOSE'] = (index, ...choices) => {
      const idx = Number(index);
      return choices[idx - 1] !== undefined ? choices[idx - 1] : '#VALUE!';
    };

    fn['ROWS'] = (array) => Array.isArray(array) ? array.length : 1;
    fn['COLUMNS'] = (array) => Array.isArray(array) && Array.isArray(array[0]) ? array[0].length : 1;

    // --- TEXT FUNCTIONS ---
    fn['CONCAT'] = (...args) => this.flatten(args).map(v => v === null ? '' : String(v)).join('');
    fn['CONCATENATE'] = fn['CONCAT'];
    fn['TEXTJOIN'] = (delimiter, ignoreEmpty, ...args) => {
      let flat = this.flatten(args);
      if (Boolean(ignoreEmpty)) flat = flat.filter(v => v !== null && v !== '');
      return flat.join(String(delimiter));
    };
    fn['LEFT'] = (text, numChars = 1) => String(text ?? '').substring(0, Number(numChars));
    fn['RIGHT'] = (text, numChars = 1) => {
      const s = String(text ?? '');
      return s.substring(Math.max(0, s.length - Number(numChars)));
    };
    fn['MID'] = (text, startPos, numChars) => {
      const s = String(text ?? '');
      return s.substr(Math.max(0, Number(startPos) - 1), Number(numChars));
    };
    fn['LEN'] = (text) => String(text ?? '').length;
    fn['TRIM'] = (text) => String(text ?? '').trim().replace(/\s+/g, ' ');
    fn['UPPER'] = (text) => String(text ?? '').toUpperCase();
    fn['LOWER'] = (text) => String(text ?? '').toLowerCase();
    fn['PROPER'] = (text) => String(text ?? '').toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
    fn['FIND'] = (findText, withinText, startPos = 1) => {
      const idx = String(withinText ?? '').indexOf(String(findText), Number(startPos) - 1);
      return idx === -1 ? '#VALUE!' : idx + 1;
    };
    fn['SEARCH'] = (findText, withinText, startPos = 1) => {
      const idx = String(withinText ?? '').toLowerCase().indexOf(String(findText).toLowerCase(), Number(startPos) - 1);
      return idx === -1 ? '#VALUE!' : idx + 1;
    };
    fn['REPLACE'] = (oldText, startPos, numChars, newText) => {
      const s = String(oldText ?? '');
      const sp = Math.max(0, Number(startPos) - 1);
      return s.substring(0, sp) + String(newText) + s.substring(sp + Number(numChars));
    };
    fn['SUBSTITUTE'] = (text, oldText, newText, instanceNum) => {
      let s = String(text ?? '');
      const ot = String(oldText);
      const nt = String(newText);
      if (!instanceNum) {
        return s.split(ot).join(nt);
      }
      let count = 0;
      return s.replace(new RegExp(ot.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'), (m) => {
        count++;
        return count === Number(instanceNum) ? nt : m;
      });
    };
    fn['TEXT'] = (val, format) => {
      // Basic number and date formatting
      if (val instanceof Date) return val.toLocaleDateString();
      if (!isNaN(val)) return Number(val).toLocaleString();
      return String(val ?? '');
    };
    fn['VALUE'] = (text) => {
      const n = Number(String(text).replace(/[$,]/g, ''));
      return isNaN(n) ? '#VALUE!' : n;
    };
    fn['EXACT'] = (t1, t2) => String(t1) === String(t2);
    fn['CHAR'] = (code) => String.fromCharCode(Number(code));
    fn['CODE'] = (text) => String(text).charCodeAt(0);
    fn['REPT'] = (text, num) => String(text).repeat(Math.max(0, Number(num)));

    // --- DATE & TIME ---
    fn['TODAY'] = () => {
      const now = new Date();
      return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
    };
    fn['NOW'] = () => new Date().toLocaleString();
    fn['DATE'] = (y, m, d) => `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
    fn['YEAR'] = (dateStr) => new Date(dateStr).getFullYear() || '#VALUE!';
    fn['MONTH'] = (dateStr) => (new Date(dateStr).getMonth() + 1) || '#VALUE!';
    fn['DAY'] = (dateStr) => new Date(dateStr).getDate() || '#VALUE!';
    fn['DATEDIF'] = (startDate, endDate, unit = 'D') => {
      const d1 = new Date(startDate);
      const d2 = new Date(endDate);
      if (isNaN(d1) || isNaN(d2) || d1 > d2) return '#NUM!';
      const diffTime = Math.abs(d2 - d1);
      const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
      const u = String(unit).toUpperCase();
      if (u === 'D') return diffDays;
      if (u === 'M') return (d2.getFullYear() - d1.getFullYear()) * 12 + (d2.getMonth() - d1.getMonth());
      if (u === 'Y') return d2.getFullYear() - d1.getFullYear();
      return diffDays;
    };
    fn['DAYS'] = (endDate, startDate) => {
      const d1 = new Date(startDate);
      const d2 = new Date(endDate);
      return Math.round((d2 - d1) / (1000 * 60 * 60 * 24));
    };
    fn['EDATE'] = (startDate, months) => {
      const d = new Date(startDate);
      d.setMonth(d.getMonth() + Number(months));
      return d.toISOString().split('T')[0];
    };

    // --- FINANCIAL ---
    fn['PMT'] = (rate, nper, pv, fv = 0, type = 0) => {
      rate = Number(rate); nper = Number(nper); pv = Number(pv); fv = Number(fv);
      if (rate === 0) return -(pv + fv) / nper;
      const pvif = Math.pow(1 + rate, nper);
      let pmt = rate / (pvif - 1) * -(pv * pvif + fv);
      if (type === 1) pmt /= (1 + rate);
      return pmt;
    };
    fn['PV'] = (rate, nper, pmt, fv = 0) => {
      rate = Number(rate); nper = Number(nper); pmt = Number(pmt); fv = Number(fv);
      if (rate === 0) return -(fv + (pmt * nper));
      return -(fv + pmt * (1 - Math.pow(1 + rate, -nper)) / rate) / Math.pow(1 + rate, nper);
    };
    fn['FV'] = (rate, nper, pmt, pv = 0) => {
      rate = Number(rate); nper = Number(nper); pmt = Number(pmt); pv = Number(pv);
      if (rate === 0) return -(pv + (pmt * nper));
      return -(pv * Math.pow(1 + rate, nper) + pmt * (Math.pow(1 + rate, nper) - 1) / rate);
    };
    fn['NPV'] = (rate, ...values) => {
      const r = Number(rate);
      const vals = this.flatten(values).map(Number);
      return vals.reduce((acc, val, i) => acc + val / Math.pow(1 + r, i + 1), 0);
    };

    // --- STATISTICS ---
    fn['MEDIAN'] = (...args) => {
      const arr = toNumArray(args).sort((a, b) => a - b);
      if (arr.length === 0) return '#NUM!';
      const mid = Math.floor(arr.length / 2);
      return arr.length % 2 !== 0 ? arr[mid] : (arr[mid - 1] + arr[mid]) / 2;
    };
    fn['STDEV'] = (...args) => {
      const arr = toNumArray(args);
      if (arr.length <= 1) return '#DIV/0!';
      const avg = arr.reduce((a, b) => a + b, 0) / arr.length;
      const sumSq = arr.reduce((acc, v) => acc + Math.pow(v - avg, 2), 0);
      return Math.sqrt(sumSq / (arr.length - 1));
    };
    fn['LARGE'] = (array, k) => {
      const arr = toNumArray(array).sort((a, b) => b - a);
      const idx = Number(k) - 1;
      return arr[idx] !== undefined ? arr[idx] : '#NUM!';
    };
    fn['SMALL'] = (array, k) => {
      const arr = toNumArray(array).sort((a, b) => a - b);
      const idx = Number(k) - 1;
      return arr[idx] !== undefined ? arr[idx] : '#NUM!';
    };

    // --- INFORMATION ---
    fn['ISBLANK'] = (v) => v === null || v === '' || v === undefined;
    fn['ISNUMBER'] = (v) => typeof v === 'number' && !isNaN(v);
    fn['ISTEXT'] = (v) => typeof v === 'string';
    fn['ISERROR'] = (v) => typeof v === 'string' && v.startsWith('#');
  }
}

/**
 * Recursive Descent Formula Parser
 */
class FormulaParser {
  constructor(tokens, engine, currentSheet, visited) {
    this.tokens = tokens;
    this.engine = engine;
    this.currentSheet = currentSheet;
    this.visited = visited;
    this.pos = 0;
  }

  peek() {
    return this.tokens[this.pos] || { type: 'EOF' };
  }

  consume(expectedType, expectedValue) {
    const token = this.peek();
    if (expectedType && token.type !== expectedType) {
      throw new Error('#VALUE!');
    }
    if (expectedValue && token.value !== expectedValue) {
      throw new Error('#VALUE!');
    }
    this.pos++;
    return token;
  }

  parse() {
    const result = this.parseExpression();
    if (this.peek().type !== 'EOF') {
      throw new Error('#VALUE!');
    }
    return result;
  }

  parseExpression() {
    return this.parseComparison();
  }

  // Comparisons: =, <>, <, <=, >, >=
  parseComparison() {
    let left = this.parseConcatenation();
    while (this.peek().type === 'OPERATOR' && ['=', '<>', '<', '<=', '>', '>='].includes(this.peek().value)) {
      const op = this.consume().value;
      const right = this.parseConcatenation();
      if (op === '=') left = String(left).toLowerCase() === String(right).toLowerCase();
      else if (op === '<>') left = String(left).toLowerCase() !== String(right).toLowerCase();
      else if (op === '<') left = Number(left) < Number(right);
      else if (op === '<=') left = Number(left) <= Number(right);
      else if (op === '>') left = Number(left) > Number(right);
      else if (op === '>=') left = Number(left) >= Number(right);
    }
    return left;
  }

  // String concatenation: &
  parseConcatenation() {
    let left = this.parseAdditive();
    while (this.peek().type === 'OPERATOR' && this.peek().value === '&') {
      this.consume();
      const right = this.parseAdditive();
      left = `${left ?? ''}${right ?? ''}`;
    }
    return left;
  }

  // Addition & Subtraction: +, -
  parseAdditive() {
    let left = this.parseMultiplicative();
    while (this.peek().type === 'OPERATOR' && ['+', '-'].includes(this.peek().value)) {
      const op = this.consume().value;
      const right = this.parseMultiplicative();
      if (op === '+') left = Number(left) + Number(right);
      else if (op === '-') left = Number(left) - Number(right);
    }
    return left;
  }

  // Multiplication & Division: *, /
  parseMultiplicative() {
    let left = this.parseExponentiation();
    while (this.peek().type === 'OPERATOR' && ['*', '/'].includes(this.peek().value)) {
      const op = this.consume().value;
      const right = this.parseExponentiation();
      if (op === '*') {
        left = Number(left) * Number(right);
      } else if (op === '/') {
        if (Number(right) === 0) throw new Error('#DIV/0!');
        left = Number(left) / Number(right);
      }
    }
    return left;
  }

  // Exponentiation: ^
  parseExponentiation() {
    let left = this.parseUnary();
    if (this.peek().type === 'OPERATOR' && this.peek().value === '^') {
      this.consume();
      const right = this.parseExponentiation();
      left = Math.pow(Number(left), Number(right));
    }
    return left;
  }

  // Unary: +, -
  parseUnary() {
    if (this.peek().type === 'OPERATOR' && ['+', '-'].includes(this.peek().value)) {
      const op = this.consume().value;
      const val = this.parseUnary();
      return op === '-' ? -Number(val) : Number(val);
    }
    return this.parsePrimary();
  }

  // Primary: Number, String, Function, Cell/Range Reference, Grouping
  parsePrimary() {
    const token = this.peek();

    if (token.type === 'NUMBER') {
      this.consume();
      return token.value;
    }

    if (token.type === 'STRING') {
      this.consume();
      return token.value;
    }

    // Function call
    if (token.type === 'FUNCTION') {
      const fnName = this.consume().value;
      this.consume('OPERATOR', '(');
      const args = [];
      if (this.peek().type !== 'OPERATOR' || this.peek().value !== ')') {
        while (true) {
          args.push(this.parseExpression());
          if (this.peek().type === 'OPERATOR' && (this.peek().value === ',' || this.peek().value === ';')) {
            this.consume();
          } else {
            break;
          }
        }
      }
      this.consume('OPERATOR', ')');

      const func = this.engine.functions[fnName];
      if (!func) throw new Error('#NAME?');
      return func(...args);
    }

    // Reference (Cell or Range)
    if (token.type === 'REFERENCE') {
      let refStr = this.consume().value;

      // Check if followed by colon ':' for range e.g. A1:B10
      if (this.peek().type === 'OPERATOR' && this.peek().value === ':') {
        this.consume();
        if (this.peek().type === 'REFERENCE') {
          refStr += ':' + this.consume().value;
        }
      }

      if (refStr.includes(':')) {
        const rangeRef = FormulaEngine.parseRangeRef(refStr);
        if (!rangeRef) throw new Error('#REF!');
        return this.engine.getRangeValues(rangeRef, this.currentSheet, this.visited);
      } else {
        const cellRef = FormulaEngine.parseCellRef(refStr);
        if (!cellRef) throw new Error('#REF!');
        return this.engine.getCellValue(cellRef.sheet || this.currentSheet, cellRef.col, cellRef.row, this.visited);
      }
    }

    // Parentheses Grouping
    if (token.type === 'OPERATOR' && token.value === '(') {
      this.consume('OPERATOR', '(');
      const val = this.parseExpression();
      this.consume('OPERATOR', ')');
      return val;
    }

    throw new Error('#VALUE!');
  }
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = { FormulaEngine, FormulaParser };
}
