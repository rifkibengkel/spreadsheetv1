import { compileCriteria, testPredicate, vectorSum, vectorSumIfs, vectorCount, vectorCountIfs } from './vectorKernel';

// ==============================================================================
// IN-MEMORY WORKBOOK & DEPENDENCY ENGINE (PHASE 4)
// Holds columnar data vectors and evaluates formulas off-thread on the backend.
// ==============================================================================

export interface FormulaNode {
  sheetId: string;
  row: number;
  col: number;
  formula: string;
  precedents: string[];
}

export interface SheetStore {
  sheetId: string;
  name: string;
  columns: Map<number, (string | number | boolean | null)[]>;
  formulas: Map<string, FormulaNode>; // Key: "r,c"
  cellValues: Map<string, any>;        // Key: "r,c"
  maxRow: number;
}

export class WorkbookEngine {
  public workbookId: string;
  public revision: number = 0;
  public sheets: Map<string, SheetStore> = new Map();
  public sheetNameMap: Map<string, string> = new Map(); // Name -> sheetId

  constructor(workbookId: string) {
    this.workbookId = workbookId;
  }

  public getOrCreateSheet(sheetId: string, sheetName?: string): SheetStore {
    if (!this.sheets.has(sheetId)) {
      const name = sheetName || sheetId;
      this.sheets.set(sheetId, {
        sheetId,
        name,
        columns: new Map(),
        formulas: new Map(),
        cellValues: new Map(),
        maxRow: 0,
      });
      this.sheetNameMap.set(name.toLowerCase(), sheetId);
    }
    return this.sheets.get(sheetId)!;
  }

  public setCellValue(sheetId: string, row: number, col: number, val: any): void {
    const sheet = this.getOrCreateSheet(sheetId);
    sheet.cellValues.set(`${row},${col}`, val);
    sheet.formulas.delete(`${row},${col}`);

    if (!sheet.columns.has(col)) {
      sheet.columns.set(col, []);
    }
    const colData = sheet.columns.get(col)!;
    colData[row] = val;
    if (row > sheet.maxRow) sheet.maxRow = row;
  }

  public setFormula(sheetId: string, row: number, col: number, formula: string): void {
    const sheet = this.getOrCreateSheet(sheetId);
    sheet.formulas.set(`${row},${col}`, {
      sheetId,
      row,
      col,
      formula,
      precedents: [],
    });
  }

  /**
   * Evaluates all dirty or registered formulas in the workbook.
   * Returns calculation patches: Array of { sheetId, row, col, v }
   */
  public evaluateAll(): Array<{ sheetId: string; row: number; col: number; v: any }> {
    const patches: Array<{ sheetId: string; row: number; col: number; v: any }> = [];

    for (const [sheetId, sheet] of this.sheets.entries()) {
      for (const [key, node] of sheet.formulas.entries()) {
        const computed = this.evaluateFormula(node);
        if (computed !== undefined && computed !== null) {
          this.setCellValue(sheetId, node.row, node.col, computed);
          patches.push({
            sheetId,
            row: node.row,
            col: node.col,
            v: computed,
          });
        }
      }
    }

    return patches;
  }

  /**
   * Evaluates a single formula node using Tier 1 vector kernel when applicable.
   */
  public evaluateFormula(node: FormulaNode): any {
    const raw = (node.formula || '').trim();
    if (!raw.startsWith('=')) return null;

    const expr = raw.slice(1).trim();
    const upper = expr.toUpperCase();
    const sheetId = node.sheetId;

    // 1. SUMIFS pattern: SUMIFS(sum_range, criteria_range1, criteria1, ...)
    if (upper.startsWith('SUMIFS(') && upper.endsWith(')')) {
      return this.evalSumIfs(expr.slice(7, -1), sheetId);
    }

    // 1b. SUMIF pattern: SUMIF(range, criteria, [sum_range])
    if (upper.startsWith('SUMIF(') && upper.endsWith(')')) {
      return this.evalSumIf(expr.slice(6, -1), sheetId);
    }

    // 1c. VLOOKUP pattern: VLOOKUP(lookup_val, table_range, col_idx, [range_lookup])
    if (upper.startsWith('VLOOKUP(') && upper.endsWith(')')) {
      return this.evalVlookup(expr.slice(8, -1), sheetId);
    }

    // 0. SUM(IF(...)) array multi-criteria formula pattern: e.g. SUM(IF('Data Valid'!$O$2:$O$414777=B6,IF('Data Valid'!$Z$2:$Z$414777=0,1,0)))
    if (upper.startsWith('SUM(IF(')) {
      return this.evalSumIfArray(expr, sheetId);
    }

    // 2. COUNTIFS pattern: COUNTIFS(criteria_range1, criteria1, ...)
    if (upper.startsWith('COUNTIFS(') && upper.endsWith(')')) {
      return this.evalCountIfs(expr.slice(9, -1), sheetId);
    }

    // 2b. COUNTIF pattern: COUNTIF(criteria_range, criteria)
    if (upper.startsWith('COUNTIF(') && upper.endsWith(')')) {
      return this.evalCountIf(expr.slice(8, -1), sheetId);
    }

    // 3. SUM pattern: SUM(range, ...)
    if (upper.startsWith('SUM(') && upper.endsWith(')')) {
      return this.evalSum(expr.slice(4, -1), sheetId);
    }

    // 4. COUNT pattern: COUNT(range, ...)
    if (upper.startsWith('COUNT(') && upper.endsWith(')')) {
      return this.evalCount(expr.slice(6, -1), sheetId);
    }

    // 5. AVERAGE pattern: AVERAGE(range, ...)
    if (upper.startsWith('AVERAGE(') && upper.endsWith(')')) {
      return this.evalAverage(expr.slice(8, -1), sheetId);
    }

    // 6. MIN pattern: MIN(range, ...)
    if (upper.startsWith('MIN(') && upper.endsWith(')')) {
      return this.evalMin(expr.slice(4, -1), sheetId);
    }

    // 7. MAX pattern: MAX(range, ...)
    if (upper.startsWith('MAX(') && upper.endsWith(')')) {
      return this.evalMax(expr.slice(4, -1), sheetId);
    }

    // 8. IF pattern: IF(cond, trueVal, falseVal)
    if (upper.startsWith('IF(') && upper.endsWith(')')) {
      return this.evalIf(expr.slice(3, -1), sheetId);
    }

    // 9. Basic arithmetic / cell reference fallback
    return this.evalSimpleExpression(expr, sheetId);
  }

  private resolveSheet(nameOrId: string): SheetStore | undefined {
    if (this.sheets.has(nameOrId)) return this.sheets.get(nameOrId);
    const cleaned = nameOrId.replace(/^['"]|['"]$/g, '').toLowerCase();
    const sheetId = this.sheetNameMap.get(cleaned);
    return sheetId ? this.sheets.get(sheetId) : undefined;
  }

  /**
   * Parses range reference string like "'Data All'!K2:K197202" or "A1:A10"
   */
  public parseRange(ref: string, defaultSheetId?: string): { sheet?: SheetStore; startCol: number; startRow: number; endCol: number; endRow: number } | null {
    const parts = ref.trim().split('!');
    let sheet: SheetStore | undefined;
    let rangePart = ref.trim();

    if (parts.length === 2) {
      sheet = this.resolveSheet(parts[0]);
      rangePart = parts[1];
    } else if (defaultSheetId) {
      sheet = this.resolveSheet(defaultSheetId);
    }

    rangePart = rangePart.replace(/\$/g, '');

    // Whole column range e.g. B:L or A:A
    const colMatch = rangePart.match(/^([A-Za-z]+):([A-Za-z]+)$/);
    if (colMatch) {
      const startCol = this.colLetterToIndex(colMatch[1]);
      const endCol = this.colLetterToIndex(colMatch[2]);
      return {
        sheet,
        startCol: Math.min(startCol, endCol),
        startRow: 0,
        endCol: Math.max(startCol, endCol),
        endRow: 1000000,
      };
    }

    const match = rangePart.match(/^([A-Za-z]+)(\d+)(?::([A-Za-z]+)(\d+))?$/);
    if (!match) return null;

    const startCol = this.colLetterToIndex(match[1]);
    const startRow = parseInt(match[2], 10) - 1; // 0-indexed
    const endCol = match[3] ? this.colLetterToIndex(match[3]) : startCol;
    const endRow = match[4] ? parseInt(match[4], 10) - 1 : startRow;

    return { sheet, startCol, startRow, endCol, endRow };
  }

  private colLetterToIndex(col: string): number {
    let result = 0;
    const upper = col.toUpperCase();
    for (let i = 0; i < upper.length; i++) {
      result = result * 26 + (upper.charCodeAt(i) - 64);
    }
    return result - 1;
  }

  private evalSum(argsStr: string, currentSheetId?: string): number {
    const args = this.splitArgs(argsStr);
    let sum = 0;
    for (const arg of args) {
      const num = Number(arg);
      if (!isNaN(num)) {
        sum += num;
        continue;
      }
      const range = this.parseRange(arg, currentSheetId);
      if (range && range.sheet) {
        const colData = range.sheet.columns.get(range.startCol) || [];
        for (let r = range.startRow; r <= Math.min(range.endRow, colData.length - 1); r++) {
          const val = Number(colData[r]);
          if (!isNaN(val)) sum += val;
        }
      }
    }
    return sum;
  }

  private evalCount(argsStr: string, currentSheetId?: string): number {
    const args = this.splitArgs(argsStr);
    let count = 0;
    for (const arg of args) {
      const range = this.parseRange(arg, currentSheetId);
      if (range && range.sheet) {
        const colData = range.sheet.columns.get(range.startCol) || [];
        for (let r = range.startRow; r <= Math.min(range.endRow, colData.length - 1); r++) {
          const val = colData[r];
          if (val !== null && val !== undefined && val !== '') count++;
        }
      }
    }
    return count;
  }

  private evalAverage(argsStr: string, currentSheetId?: string): number {
    const count = this.evalCount(argsStr, currentSheetId);
    if (count === 0) return 0;
    const sum = this.evalSum(argsStr, currentSheetId);
    return sum / count;
  }

  private evalMin(argsStr: string, currentSheetId?: string): number {
    const args = this.splitArgs(argsStr);
    let min = Infinity;
    for (const arg of args) {
      const range = this.parseRange(arg, currentSheetId);
      if (range && range.sheet) {
        const colData = range.sheet.columns.get(range.startCol) || [];
        for (let r = range.startRow; r <= Math.min(range.endRow, colData.length - 1); r++) {
          const val = Number(colData[r]);
          if (!isNaN(val) && val < min) min = val;
        }
      }
    }
    return min === Infinity ? 0 : min;
  }

  private evalMax(argsStr: string, currentSheetId?: string): number {
    const args = this.splitArgs(argsStr);
    let max = -Infinity;
    for (const arg of args) {
      const range = this.parseRange(arg, currentSheetId);
      if (range && range.sheet) {
        const colData = range.sheet.columns.get(range.startCol) || [];
        for (let r = range.startRow; r <= Math.min(range.endRow, colData.length - 1); r++) {
          const val = Number(colData[r]);
          if (!isNaN(val) && val > max) max = val;
        }
      }
    }
    return max === -Infinity ? 0 : max;
  }

  private evalIf(argsStr: string, currentSheetId?: string): any {
    const args = this.splitArgs(argsStr);
    if (args.length < 2) return null;
    const cond = this.evalSimpleExpression(args[0], currentSheetId);
    if (cond) {
      return this.evalSimpleExpression(args[1], currentSheetId);
    } else if (args[2] !== undefined) {
      return this.evalSimpleExpression(args[2], currentSheetId);
    }
    return false;
  }

  private evalSumIfs(argsStr: string, currentSheetId?: string): number {
    const args = this.splitArgs(argsStr);
    if (args.length < 3 || args.length % 2 === 0) return 0;

    const sumRange = this.parseRange(args[0], currentSheetId);
    if (!sumRange || !sumRange.sheet) return 0;

    const sumCol = sumRange.sheet.columns.get(sumRange.startCol) || [];
    const criteriaPairs: Array<{ values: any[]; predicate: any }> = [];

    for (let i = 1; i < args.length; i += 2) {
      const critRange = this.parseRange(args[i], currentSheetId);
      let rawCriteria = args[i + 1].trim();
      const cellRefMatch = rawCriteria.match(/^(?:(?:'[^']+'|"[^"]+"|[A-Za-z0-9_]+)!)?\$?([A-Za-z]+)\$?(\d+)$/);
      if (cellRefMatch && !rawCriteria.startsWith('"') && !rawCriteria.startsWith("'")) {
        const val = this.evalSimpleExpression(rawCriteria, currentSheetId);
        rawCriteria = val !== undefined && val !== null ? String(val) : '';
      } else {
        rawCriteria = this.stripQuotes(rawCriteria);
      }

      if (critRange && critRange.sheet) {
        const critCol = critRange.sheet.columns.get(critRange.startCol) || [];
        criteriaPairs.push({
          values: critCol,
          predicate: compileCriteria(rawCriteria),
        });
      }
    }

    return vectorSumIfs(sumCol as any, criteriaPairs, sumRange.endRow + 1);
  }

  private evalCountIfs(argsStr: string, currentSheetId?: string): number {
    const args = this.splitArgs(argsStr);
    if (args.length < 2 || args.length % 2 !== 0) return 0;

    const criteriaPairs: Array<{ values: any[]; predicate: any }> = [];
    let maxRows = 0;

    for (let i = 0; i < args.length; i += 2) {
      const critRange = this.parseRange(args[i], currentSheetId);
      let rawCriteria = args[i + 1].trim();
      const cellRefMatch = rawCriteria.match(/^(?:(?:'[^']+'|"[^"]+"|[A-Za-z0-9_]+)!)?\$?([A-Za-z]+)\$?(\d+)$/);
      if (cellRefMatch && !rawCriteria.startsWith('"') && !rawCriteria.startsWith("'")) {
        const val = this.evalSimpleExpression(rawCriteria, currentSheetId);
        rawCriteria = val !== undefined && val !== null ? String(val) : '';
      } else {
        rawCriteria = this.stripQuotes(rawCriteria);
      }

      if (critRange && critRange.sheet) {
        const critCol = critRange.sheet.columns.get(critRange.startCol) || [];
        criteriaPairs.push({
          values: critCol,
          predicate: compileCriteria(rawCriteria),
        });
        if (critRange.endRow + 1 > maxRows) maxRows = critRange.endRow + 1;
      }
    }

    return vectorCountIfs(criteriaPairs, maxRows);
  }

  private evalCountIf(argsStr: string, currentSheetId?: string): number {
    const args = this.splitArgs(argsStr);
    if (args.length < 2) return 0;
    return this.evalCountIfs(`${args[0]}, ${args[1]}`, currentSheetId);
  }

  private evalSumIfArray(rawExpr: string, currentSheetId?: string): number {
    const inner = rawExpr.slice(4, -1).trim();

    const condRegex = /IF\s*\(([^,]+),/gi;
    const conditions: string[] = [];
    let m: RegExpExecArray | null;
    while ((m = condRegex.exec(inner)) !== null) {
      conditions.push(m[1].trim());
    }
    if (conditions.length === 0) return 0;

    const tailMatch = inner.match(/,\s*([^,()]+)\s*,\s*([^,()]+)\s*\)+$/);
    const trueValStr = tailMatch ? tailMatch[1].trim() : '1';

    const criteriaArgs: string[] = [];
    for (const condStr of conditions) {
      const parts = condStr.split('=');
      if (parts.length !== 2) continue;
      let rangeStr = parts[0].trim();
      let critStr = parts[1].trim();
      if (!this.parseRange(rangeStr, currentSheetId)) {
        rangeStr = parts[1].trim();
        critStr = parts[0].trim();
      }
      const cellRefMatch = critStr.match(/^(?:(?:'[^']+'|"[^"]+"|[A-Za-z0-9_]+)!)?\$?([A-Za-z]+)\$?(\d+)$/);
      if (cellRefMatch) {
        const cellVal = this.evalSimpleExpression(critStr, currentSheetId);
        critStr = cellVal !== undefined && cellVal !== null ? String(cellVal) : '';
      }
      criteriaArgs.push(rangeStr, critStr);
    }

    if (criteriaArgs.length === 0) return 0;

    if (trueValStr !== '1' && this.parseRange(trueValStr, currentSheetId)) {
      return this.evalSumIfs([trueValStr, ...criteriaArgs].join(','), currentSheetId);
    } else {
      return this.evalCountIfs(criteriaArgs.join(','), currentSheetId);
    }
  }

  private evalSumIf(argsStr: string, currentSheetId?: string): number {
    const args = this.splitArgs(argsStr);
    if (args.length < 2) return 0;

    const range = this.parseRange(args[0], currentSheetId);
    if (!range || !range.sheet) return 0;

    const rawCriteria = this.evalSimpleExpression(args[1], currentSheetId);
    const predicate = compileCriteria(rawCriteria);

    const sumRange = args[2] ? this.parseRange(args[2], currentSheetId) : range;
    if (!sumRange || !sumRange.sheet) return 0;

    const critCol = range.sheet.columns.get(range.startCol) || [];
    const sumCol = sumRange.sheet.columns.get(sumRange.startCol) || [];

    let sum = 0;
    const startRow = range.startRow;
    const endRow = Math.min(range.endRow, critCol.length - 1);
    const sumStartRow = sumRange.startRow;

    for (let r = startRow; r <= endRow; r++) {
      if (testPredicate(critCol[r], predicate)) {
        const targetRow = sumStartRow + (r - startRow);
        const val = Number(sumCol[targetRow]);
        if (!isNaN(val)) sum += val;
      }
    }
    return sum;
  }

  private evalVlookup(argsStr: string, currentSheetId?: string): any {
    const args = this.splitArgs(argsStr);
    if (args.length < 3) return null;

    let lookupVal = this.evalSimpleExpression(args[0], currentSheetId);
    if (lookupVal === null || lookupVal === undefined) {
      lookupVal = this.stripQuotes(args[0]);
    }

    const tableRange = this.parseRange(args[1], currentSheetId);
    const colIndex = parseInt(args[2], 10);
    if (!tableRange || !tableRange.sheet || isNaN(colIndex) || colIndex < 1) {
      return '#N/A';
    }

    const lookupCol = tableRange.sheet.columns.get(tableRange.startCol) || [];
    const targetColIdx = tableRange.startCol + colIndex - 1;
    const targetCol = tableRange.sheet.columns.get(targetColIdx) || [];

    const lookupStr = String(lookupVal).trim().toLowerCase();
    const lookupNum = Number(lookupVal);
    const isNum = !isNaN(lookupNum) && String(lookupVal).trim() !== '';

    const startRow = tableRange.startRow;
    const endRow = Math.min(tableRange.endRow, lookupCol.length - 1);

    for (let r = startRow; r <= endRow; r++) {
      const cellVal = lookupCol[r];
      if (cellVal === undefined || cellVal === null) continue;

      let isMatch = false;
      if (isNum && typeof cellVal === 'number') {
        isMatch = cellVal === lookupNum;
      } else if (isNum && !isNaN(Number(cellVal))) {
        isMatch = Number(cellVal) === lookupNum;
      } else {
        isMatch = String(cellVal).trim().toLowerCase() === lookupStr;
      }

      if (isMatch) {
        const res = targetCol[r];
        return res !== undefined && res !== null ? res : '';
      }
    }

    return '#N/A';
  }

  private evalSimpleExpression(expr: string, currentSheetId?: string): any {
    const trimmed = expr.trim();

    // Check if it's a numeric literal
    const num = Number(trimmed);
    if (!isNaN(num)) return num;

    // Check if string literal
    if ((trimmed.startsWith('"') && trimmed.endsWith('"')) || (trimmed.startsWith("'") && trimmed.endsWith("'"))) {
      return trimmed.slice(1, -1);
    }

    // Check if single cell reference e.g. "A1" or "'Sheet1'!A1"
    const single = this.parseRange(trimmed, currentSheetId);
    if (single && single.sheet && single.startRow === single.endRow && single.startCol === single.endCol) {
      return single.sheet.cellValues.get(`${single.startRow},${single.startCol}`) ?? 0;
    }

    // Check basic binary arithmetic e.g. "A1 + B1" or "A1 * 10"
    const opMatch = trimmed.match(/^(.+?)\s*([+\-*/><=]+)\s*(.+)$/);
    if (opMatch) {
      const left = this.evalSimpleExpression(opMatch[1], currentSheetId);
      const op = opMatch[2];
      const right = this.evalSimpleExpression(opMatch[3], currentSheetId);

      switch (op) {
        case '+': return (Number(left) || 0) + (Number(right) || 0);
        case '-': return (Number(left) || 0) - (Number(right) || 0);
        case '*': return (Number(left) || 0) * (Number(right) || 0);
        case '/': return (Number(right) || 0) !== 0 ? (Number(left) || 0) / Number(right) : '#DIV/0!';
        case '>': return left > right;
        case '<': return left < right;
        case '>=': return left >= right;
        case '<=': return left <= right;
        case '=': case '==': return left == right;
        case '<>': case '!=': return left != right;
      }
    }

    return null;
  }

  private splitArgs(str: string): string[] {
    const result: string[] = [];
    let current = '';
    let inQuotes = false;
    let depth = 0;

    for (let i = 0; i < str.length; i++) {
      const char = str[i];
      if (char === '"') inQuotes = !inQuotes;
      else if (char === '(' && !inQuotes) depth++;
      else if (char === ')' && !inQuotes) depth--;
      else if (char === ',' && !inQuotes && depth === 0) {
        result.push(current.trim());
        current = '';
        continue;
      }
      current += char;
    }
    if (current.trim()) result.push(current.trim());
    return result;
  }

  private stripQuotes(str: string): string {
    const trimmed = str.trim();
    if (trimmed.startsWith('"') && trimmed.endsWith('"')) {
      return trimmed.slice(1, -1);
    }
    return trimmed;
  }
}
