/**
 * MicroFormulaEvaluator
 * 
 * Provides instantaneous (<0.01ms) formula evaluation directly from active worksheet
 * memory, bypassing Univer's 45-second synchronous topological graph calculation storm.
 */

export function colLetterToIndex(colStr: string): number {
  let result = 0;
  const upper = colStr.toUpperCase();
  for (let i = 0; i < upper.length; i++) {
    result = result * 26 + (upper.charCodeAt(i) - 64);
  }
  return result - 1; // 0-indexed
}

export function indexToColLetter(index: number): string {
  let col = '';
  let temp = index;
  while (temp >= 0) {
    col = String.fromCharCode((temp % 26) + 65) + col;
    temp = Math.floor(temp / 26) - 1;
  }
  return col;
}

export interface CellCoord {
  sheetName?: string;
  row: number; // 0-indexed
  col: number; // 0-indexed
}

export interface RangeCoord {
  sheetName?: string;
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

export function parseCellRef(ref: string): CellCoord | null {
  const parts = ref.trim().split('!');
  let sheetName: string | undefined;
  let cellPart = ref.trim();

  if (parts.length === 2) {
    sheetName = parts[0].replace(/^['"]|['"]$/g, '');
    cellPart = parts[1];
  }

  // Strip $ if present, e.g. $C$4, $C4, C$4 -> C4
  cellPart = cellPart.replace(/\$/g, '');

  const match = cellPart.match(/^([A-Za-z]+)(\d+)$/);
  if (!match) return null;

  return {
    sheetName,
    col: colLetterToIndex(match[1]),
    row: parseInt(match[2], 10) - 1,
  };
}

export function parseRangeRef(ref: string, maxRows = 1000000): RangeCoord | null {
  const parts = ref.trim().split('!');
  let sheetName: string | undefined;
  let rangePart = ref.trim();

  if (parts.length === 2) {
    sheetName = parts[0].replace(/^['"]|['"]$/g, '');
    rangePart = parts[1];
  }

  rangePart = rangePart.replace(/\$/g, '');

  // 1. Check whole column format e.g. B:L or A:A
  const colMatch = rangePart.match(/^([A-Za-z]+):([A-Za-z]+)$/);
  if (colMatch) {
    const startCol = colLetterToIndex(colMatch[1]);
    const endCol = colLetterToIndex(colMatch[2]);
    return {
      sheetName,
      startRow: 0,
      startCol: Math.min(startCol, endCol),
      endRow: maxRows - 1,
      endCol: Math.max(startCol, endCol),
    };
  }

  // 2. Standard format e.g. C4:C662
  const match = rangePart.match(/^([A-Za-z]+)(\d+):([A-Za-z]+)(\d+)$/);
  if (!match) return null;

  const startCol = colLetterToIndex(match[1]);
  const startRow = parseInt(match[2], 10) - 1;
  const endCol = colLetterToIndex(match[3]);
  const endRow = parseInt(match[4], 10) - 1;

  return {
    sheetName,
    startRow: Math.min(startRow, endRow),
    startCol: Math.min(startCol, endCol),
    endRow: Math.max(startRow, endRow),
    endCol: Math.max(startCol, endCol),
  };
}

export function splitArguments(argsStr: string): string[] {
  const args: string[] = [];
  let current = '';
  let inQuotes = false;
  let quoteChar = '';
  let parenDepth = 0;

  for (let i = 0; i < argsStr.length; i++) {
    const char = argsStr[i];
    if ((char === '"' || char === "'") && !inQuotes) {
      inQuotes = true;
      quoteChar = char;
      current += char;
    } else if (char === quoteChar && inQuotes) {
      inQuotes = false;
      quoteChar = '';
      current += char;
    } else if (char === '(' && !inQuotes) {
      parenDepth++;
      current += char;
    } else if (char === ')' && !inQuotes) {
      parenDepth--;
      current += char;
    } else if (char === ',' && !inQuotes && parenDepth === 0) {
      args.push(current.trim());
      current = '';
    } else {
      current += char;
    }
  }
  if (current.trim().length > 0 || args.length > 0) {
    args.push(current.trim());
  }
  return args;
}

export function extractReferencedCells(formula: string): CellCoord[] {
  if (!formula || !formula.startsWith('=')) return [];
  const coords: CellCoord[] = [];
  const visited = new Set<string>();

  // 1. Match ranges: e.g. A1:A10, $C$4:$C$662, or 'Sheet 1'!B2:B5
  const rangeRegex = /(?:(?:'[^']+'|"[^"]+"|[A-Za-z0-9_]+)!)?\$?([A-Za-z]+)\$?(\d+)?:\$?([A-Za-z]+)\$?(\d+)?/g;
  let rMatch: RegExpExecArray | null;
  while ((rMatch = rangeRegex.exec(formula)) !== null) {
    const range = parseRangeRef(rMatch[0]);
    if (range) {
      const rowCount = range.endRow - range.startRow + 1;
      const colCount = range.endCol - range.startCol + 1;
      // Cap at 2000 cells to prevent huge memory spikes for whole column references
      if (rowCount * colCount <= 2000) {
        for (let r = range.startRow; r <= range.endRow; r++) {
          for (let c = range.startCol; c <= range.endCol; c++) {
            const key = `${range.sheetName || ''}_${r}_${c}`;
            if (!visited.has(key)) {
              visited.add(key);
              coords.push({ sheetName: range.sheetName, row: r, col: c });
            }
          }
        }
      }
    }
  }

  // 2. Match single cells: e.g. D10, $E$10, H5 (not followed by :)
  const singleRegex = /(?:(?:'[^']+'|"[^"]+"|[A-Za-z0-9_]+)!)?\$?([A-Za-z]+)\$?(\d+)(?![:\d])/g;
  let sMatch: RegExpExecArray | null;
  while ((sMatch = singleRegex.exec(formula)) !== null) {
    if (sMatch.index > 0 && formula[sMatch.index - 1] === ':') continue;
    const cell = parseCellRef(sMatch[0]);
    if (cell) {
      const key = `${cell.sheetName || ''}_${cell.row}_${cell.col}`;
      if (!visited.has(key)) {
        visited.add(key);
        coords.push(cell);
      }
    }
  }

  return coords;
}

export type CellValueLookup = (sheetName: string | undefined, row: number, col: number) => any;

function compileCriteriaMatcher(rawCrit: any, lookup: CellValueLookup) {
  let crit = rawCrit;
  if (typeof crit === 'string') {
    const trimmed = crit.trim();
    const cellRef = parseCellRef(trimmed);
    if (cellRef) {
      crit = lookup(cellRef.sheetName, cellRef.row, cellRef.col);
    } else if ((trimmed.startsWith('"') && trimmed.endsWith('"')) || (trimmed.startsWith("'") && trimmed.endsWith("'"))) {
      crit = trimmed.slice(1, -1);
    }
  }

  if (crit === undefined || crit === null) {
    return (val: any) => val === undefined || val === null || val === '';
  }

  if (typeof crit === 'string') {
    const trimmedCrit = crit.trim();
    if (trimmedCrit.startsWith('>=')) {
      const num = Number(trimmedCrit.slice(2).trim());
      return (val: any) => Number(val) >= num;
    }
    if (trimmedCrit.startsWith('<=')) {
      const num = Number(trimmedCrit.slice(2).trim());
      return (val: any) => Number(val) <= num;
    }
    if (trimmedCrit.startsWith('<>')) {
      const target = trimmedCrit.slice(2).trim().toLowerCase();
      return (val: any) => String(val ?? '').toLowerCase() !== target;
    }
    if (trimmedCrit.startsWith('>')) {
      const num = Number(trimmedCrit.slice(1).trim());
      return (val: any) => Number(val) > num;
    }
    if (trimmedCrit.startsWith('<')) {
      const num = Number(trimmedCrit.slice(1).trim());
      return (val: any) => Number(val) < num;
    }
    if (trimmedCrit.startsWith('=')) {
      const target = trimmedCrit.slice(1).trim().toLowerCase();
      return (val: any) => String(val ?? '').toLowerCase() === target;
    }

    const lower = trimmedCrit.toLowerCase();
    const numCrit = Number(trimmedCrit);
    const isNum = !isNaN(numCrit) && trimmedCrit !== '';

    return (val: any) => {
      if (val === undefined || val === null) return false;
      if (isNum && Number(val) === numCrit) return true;
      return String(val).toLowerCase() === lower;
    };
  }

  if (typeof crit === 'number') {
    return (val: any) => Number(val) === crit;
  }

  return (val: any) => val == crit;
}

/**
 * Evaluates array formulas structured as SUM(IF(cond1, [IF(cond2, ...)], trueVal, falseVal))
 * e.g. =SUM(IF('Data Valid'!$O$2:$O$414777=B6,IF('Data Valid'!$Z$2:$Z$414777=0,1,0)))
 */
export function evaluateSumIfArrayFormula(rawFormula: string, lookup: CellValueLookup): number | null {
  const upper = rawFormula.toUpperCase();
  if (!upper.startsWith('SUM(IF(')) return null;

  const inner = rawFormula.slice(4, -1).trim();

  // Find all IF conditions: e.g. IF(cond1, IF(cond2, ... , trueVal, falseVal)...)
  const condRegex = /IF\s*\(([^,]+),/gi;
  const conditions: string[] = [];
  let m: RegExpExecArray | null;
  while ((m = condRegex.exec(inner)) !== null) {
    conditions.push(m[1].trim());
  }

  if (conditions.length === 0) return null;

  // Find the return value after all IFs: e.g. ",1,0)" or ", 1, 0)"
  const tailMatch = inner.match(/,\s*([^,()]+)\s*,\s*([^,()]+)\s*\)+$/);
  const trueValStr = tailMatch ? tailMatch[1].trim() : '1';

  // Parse conditions
  const parsedConds: Array<{ range: RangeCoord; matcher: (val: any) => boolean }> = [];
  for (const condStr of conditions) {
    const parts = condStr.split('=');
    if (parts.length !== 2) continue;
    let range = parseRangeRef(parts[0]);
    let critStr = parts[1].trim();
    if (!range) {
      range = parseRangeRef(parts[1]);
      critStr = parts[0].trim();
    }
    if (!range) continue;
    const matcher = compileCriteriaMatcher(critStr, lookup);
    parsedConds.push({ range, matcher });
  }

  if (parsedConds.length === 0) return null;

  // Determine row range
  const startRow = Math.max(...parsedConds.map((c) => c.range.startRow));
  const endRow = Math.min(...parsedConds.map((c) => c.range.endRow));
  const sheetName = parsedConds[0].range.sheetName;

  // Check if trueVal is 1 (COUNTIFS) or a range (SUMIFS)
  const sumRange = trueValStr !== '1' ? parseRangeRef(trueValStr) : null;
  const sumSheet = sumRange ? sumRange.sheetName : sheetName;
  const sumCol = sumRange ? sumRange.startCol : 0;
  const sumStartRow = sumRange ? sumRange.startRow : 0;

  let sum = 0;
  for (let r = startRow; r <= endRow; r++) {
    let allMatch = true;
    for (let ci = 0; ci < parsedConds.length; ci++) {
      const c = parsedConds[ci];
      const val = lookup(c.range.sheetName, r, c.range.startCol);
      if (!c.matcher(val)) {
        allMatch = false;
        break;
      }
    }
    if (allMatch) {
      if (sumRange) {
        const val = Number(lookup(sumSheet, sumStartRow + (r - startRow), sumCol));
        if (!isNaN(val)) sum += val;
      } else {
        sum += 1;
      }
    }
  }

  return sum;
}

/**
 * Fast synchronous formula evaluator for instant client UI display.
 */
export function evaluateMicroFormula(
  formula: string,
  lookup: CellValueLookup
): any {
  if (!formula || typeof formula !== 'string' || !formula.startsWith('=')) {
    return null;
  }

  const raw = formula.slice(1).trim();
  const upper = raw.toUpperCase();

  try {
    // 0. SUM(IF(...)) Array Multi-Criteria Multi-Condition Formula: e.g. =SUM(IF('Data Valid'!$O$2:$O$414777=B6,IF('Data Valid'!$Z$2:$Z$414777=0,1,0)))
    if (upper.startsWith('SUM(IF(')) {
      const arrayRes = evaluateSumIfArrayFormula(raw, lookup);
      if (arrayRes !== null) return arrayRes;
    }

    // 1. COUNTIF function: =COUNTIF(range, criteria)
    if (upper.startsWith('COUNTIF(') && upper.endsWith(')')) {
      const inner = raw.slice(8, -1).trim();
      const args = splitArguments(inner);
      if (args.length >= 2) {
        const range = parseRangeRef(args[0]);
        if (range) {
          const matcher = compileCriteriaMatcher(args[1], lookup);
          let count = 0;
          for (let r = range.startRow; r <= range.endRow; r++) {
            for (let c = range.startCol; c <= range.endCol; c++) {
              const val = lookup(range.sheetName, r, c);
              if (matcher(val)) {
                count++;
              }
            }
          }
          return count;
        }
      }
    }

    // 1b. SUMIF function: =SUMIF(range, criteria, [sum_range])
    if (upper.startsWith('SUMIF(') && upper.endsWith(')')) {
      const inner = raw.slice(6, -1).trim();
      const args = splitArguments(inner);
      if (args.length >= 2) {
        const range = parseRangeRef(args[0]);
        if (range) {
          const matcher = compileCriteriaMatcher(args[1], lookup);
          const sumRange = args[2] ? parseRangeRef(args[2]) : range;
          const sumStartRow = sumRange ? sumRange.startRow : range.startRow;
          const sumStartCol = sumRange ? sumRange.startCol : range.startCol;
          const sumSheet = sumRange?.sheetName || range.sheetName;

          let sum = 0;
          for (let r = range.startRow; r <= range.endRow; r++) {
            const critVal = lookup(range.sheetName, r, range.startCol);
            if (matcher(critVal)) {
              const targetRow = sumStartRow + (r - range.startRow);
              const val = Number(lookup(sumSheet, targetRow, sumStartCol));
              if (!isNaN(val)) sum += val;
            }
          }
          return sum;
        }
      }
    }

    // 2. VLOOKUP function: =VLOOKUP(lookup_value, table_array, col_index, [range_lookup])
    if (upper.startsWith('VLOOKUP(') && upper.endsWith(')')) {
      const inner = raw.slice(8, -1).trim();
      const args = splitArguments(inner);
      if (args.length >= 3) {
        let lookupVal: any;
        const cellRef = parseCellRef(args[0]);
        if (cellRef) {
          lookupVal = lookup(cellRef.sheetName, cellRef.row, cellRef.col);
        } else {
          const trimmed = args[0].trim();
          if ((trimmed.startsWith('"') && trimmed.endsWith('"')) || (trimmed.startsWith("'") && trimmed.endsWith("'"))) {
            lookupVal = trimmed.slice(1, -1);
          } else {
            const num = Number(trimmed);
            lookupVal = !isNaN(num) ? num : trimmed;
          }
        }

        const tableRange = parseRangeRef(args[1]);
        const colIndex = parseInt(args[2], 10);

        if (tableRange && !isNaN(colIndex) && colIndex >= 1) {
          const targetCol = tableRange.startCol + colIndex - 1;
          const lookupStr = lookupVal !== undefined && lookupVal !== null ? String(lookupVal).trim().toLowerCase() : '';
          const lookupNum = Number(lookupVal);
          const isNum = lookupVal !== undefined && lookupVal !== null && !isNaN(lookupNum) && String(lookupVal).trim() !== '';

          let consecutiveEmpty = 0;
          for (let r = tableRange.startRow; r <= tableRange.endRow; r++) {
            const cellVal = lookup(tableRange.sheetName, r, tableRange.startCol);

            if (cellVal === undefined || cellVal === null || cellVal === '') {
              consecutiveEmpty++;
              if (r > 1000 && consecutiveEmpty > 2000) break;
              continue;
            }
            consecutiveEmpty = 0;

            let isMatch = false;
            if (isNum && typeof cellVal === 'number') {
              isMatch = cellVal === lookupNum;
            } else if (isNum && !isNaN(Number(cellVal))) {
              isMatch = Number(cellVal) === lookupNum;
            } else {
              isMatch = String(cellVal).trim().toLowerCase() === lookupStr;
            }

            if (isMatch) {
              const res = lookup(tableRange.sheetName, r, targetCol);
              return res !== undefined && res !== null ? res : '';
            }
          }
          return '#N/A';
        }
      }
    }

    // 3. SUM function: =SUM(A1:A5) or =SUM(A1, B1, C1)
    if (upper.startsWith('SUM(') && upper.endsWith(')')) {
      const inner = raw.slice(4, -1).trim();
      const args = splitArguments(inner);
      let sum = 0;
      for (const arg of args) {
        const range = parseRangeRef(arg);
        if (range) {
          for (let r = range.startRow; r <= range.endRow; r++) {
            for (let c = range.startCol; c <= range.endCol; c++) {
              const val = Number(lookup(range.sheetName, r, c));
              if (!isNaN(val)) sum += val;
            }
          }
          continue;
        }
        const cell = parseCellRef(arg);
        if (cell) {
          const val = Number(lookup(cell.sheetName, cell.row, cell.col));
          if (!isNaN(val)) sum += val;
          continue;
        }
        const num = Number(arg);
        if (!isNaN(num)) sum += num;
      }
      return sum;
    }

    // 4. AVERAGE function: =AVERAGE(A1:A5)
    if (upper.startsWith('AVERAGE(') && upper.endsWith(')')) {
      const inner = raw.slice(8, -1).trim();
      const args = splitArguments(inner);
      let sum = 0;
      let count = 0;
      for (const arg of args) {
        const range = parseRangeRef(arg);
        if (range) {
          for (let r = range.startRow; r <= range.endRow; r++) {
            for (let c = range.startCol; c <= range.endCol; c++) {
              const val = Number(lookup(range.sheetName, r, c));
              if (!isNaN(val)) {
                sum += val;
                count++;
              }
            }
          }
          continue;
        }
        const cell = parseCellRef(arg);
        if (cell) {
          const val = Number(lookup(cell.sheetName, cell.row, cell.col));
          if (!isNaN(val)) {
            sum += val;
            count++;
          }
          continue;
        }
        const num = Number(arg);
        if (!isNaN(num)) {
          sum += num;
          count++;
        }
      }
      return count > 0 ? sum / count : 0;
    }

    // 5. COUNT function: =COUNT(A1:A5)
    if (upper.startsWith('COUNT(') && upper.endsWith(')')) {
      const inner = raw.slice(6, -1).trim();
      const args = splitArguments(inner);
      let count = 0;
      for (const arg of args) {
        const range = parseRangeRef(arg);
        if (range) {
          for (let r = range.startRow; r <= range.endRow; r++) {
            for (let c = range.startCol; c <= range.endCol; c++) {
              const val = lookup(range.sheetName, r, c);
              if (val !== null && val !== undefined && val !== '' && !isNaN(Number(val))) {
                count++;
              }
            }
          }
          continue;
        }
        const cell = parseCellRef(arg);
        if (cell) {
          const val = lookup(cell.sheetName, cell.row, cell.col);
          if (val !== null && val !== undefined && val !== '' && !isNaN(Number(val))) {
            count++;
          }
          continue;
        }
        const num = Number(arg);
        if (!isNaN(num)) count++;
      }
      return count;
    }

    // 6. Arithmetic and single cell references: e.g. =D10+E10, =A1*2, =A1+B1-C1
    let expr = raw;
    expr = expr.replace(/(?:(?:'[^']+'|"[^"]+"|[A-Za-z0-9_]+)!)?\$?([A-Za-z]+)\$?(\d+)/g, (match) => {
      const coord = parseCellRef(match);
      if (!coord) return '0';
      const val = lookup(coord.sheetName, coord.row, coord.col);
      const num = Number(val);
      return !isNaN(num) ? String(num) : '0';
    });

    if (/^[0-9\s.+\-*/()]+$/.test(expr)) {
      const evalFunc = new Function(`return (${expr});`);
      const res = evalFunc();
      return isFinite(res) ? res : '#DIV/0!';
    }
  } catch (err) {
    console.warn('[MicroFormulaEvaluator] Evaluation error:', err);
  }

  return null;
}
