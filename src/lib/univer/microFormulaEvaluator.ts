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

  const match = cellPart.match(/^([A-Za-z]+)(\d+)$/);
  if (!match) return null;

  return {
    sheetName,
    col: colLetterToIndex(match[1]),
    row: parseInt(match[2], 10) - 1,
  };
}

export function parseRangeRef(ref: string): RangeCoord | null {
  const parts = ref.trim().split('!');
  let sheetName: string | undefined;
  let rangePart = ref.trim();

  if (parts.length === 2) {
    sheetName = parts[0].replace(/^['"]|['"]$/g, '');
    rangePart = parts[1];
  }

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

export function extractReferencedCells(formula: string): CellCoord[] {
  if (!formula || !formula.startsWith('=')) return [];
  const coords: CellCoord[] = [];
  const visited = new Set<string>();

  // 1. Match ranges: e.g. A1:A10 or 'Sheet 1'!B2:B5
  const rangeRegex = /(?:(?:'[^']+'|"[^"]+"|[A-Za-z0-9_]+)!)?([A-Za-z]+)(\d+):([A-Za-z]+)(\d+)/g;
  let rMatch: RegExpExecArray | null;
  while ((rMatch = rangeRegex.exec(formula)) !== null) {
    const range = parseRangeRef(rMatch[0]);
    if (range) {
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

  // 2. Match single cells: e.g. D10, E10 (not followed by :)
  const singleRegex = /(?:(?:'[^']+'|"[^"]+"|[A-Za-z0-9_]+)!)?([A-Za-z]+)(\d+)(?!:)/g;
  let sMatch: RegExpExecArray | null;
  while ((sMatch = singleRegex.exec(formula)) !== null) {
    // Check if preceded by :
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
    // 1. SUM function: =SUM(A1:A5) or =SUM(A1, B1, C1)
    if (upper.startsWith('SUM(') && upper.endsWith(')')) {
      const inner = raw.slice(4, -1).trim();
      const args = inner.split(',').map((s) => s.trim());
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

    // 2. AVERAGE function: =AVERAGE(A1:A5)
    if (upper.startsWith('AVERAGE(') && upper.endsWith(')')) {
      const inner = raw.slice(8, -1).trim();
      const args = inner.split(',').map((s) => s.trim());
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

    // 3. COUNT function: =COUNT(A1:A5)
    if (upper.startsWith('COUNT(') && upper.endsWith(')')) {
      const inner = raw.slice(6, -1).trim();
      const args = inner.split(',').map((s) => s.trim());
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

    // 4. Arithmetic and single cell references: e.g. =D10+E10, =A1*2, =A1+B1-C1
    // Substitute all cell references with their numeric values
    let expr = raw;
    expr = expr.replace(/(?:(?:'[^']+'|"[^"]+"|[A-Za-z0-9_]+)!)?([A-Za-z]+)(\d+)/g, (match) => {
      const coord = parseCellRef(match);
      if (!coord) return '0';
      const val = lookup(coord.sheetName, coord.row, coord.col);
      const num = Number(val);
      return !isNaN(num) ? String(num) : '0';
    });

    // Safely evaluate simple mathematical expressions (+, -, *, /, parentheses)
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
