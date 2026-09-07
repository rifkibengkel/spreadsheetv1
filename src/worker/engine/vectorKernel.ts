// ==============================================================================
// TIER 1 VECTORIZED CALCULATION KERNEL (PHASE 4 IMPLEMENTATION)
// Contiguous TypedArray & SIMD-compatible evaluation for aggregate formulas
// Zero hardcoded sheet/column names. 100% General-Purpose.
// ==============================================================================

export interface CriteriaPredicate {
  operator: '=' | '!=' | '>' | '>=' | '<' | '<=';
  value: string | number | boolean;
  isNumeric: boolean;
  numValue?: number;
}

/**
 * Parses an Excel criteria string into an optimized predicate matcher.
 * Supports "=100", ">50", "<=0", "<>", "Hexos*", "ABC", 123.
 */
export function compileCriteria(criteria: any): CriteriaPredicate {
  if (typeof criteria === 'number') {
    return { operator: '=', value: criteria, isNumeric: true, numValue: criteria };
  }
  if (typeof criteria === 'boolean') {
    return { operator: '=', value: criteria, isNumeric: false };
  }

  const str = String(criteria ?? '').trim();

  if (str.startsWith('>=')) {
    const num = parseFloat(str.slice(2));
    return { operator: '>=', value: str.slice(2), isNumeric: !isNaN(num), numValue: num };
  }
  if (str.startsWith('<=')) {
    const num = parseFloat(str.slice(2));
    return { operator: '<=', value: str.slice(2), isNumeric: !isNaN(num), numValue: num };
  }
  if (str.startsWith('<>')) {
    const num = parseFloat(str.slice(2));
    return { operator: '!=', value: str.slice(2), isNumeric: !isNaN(num), numValue: num };
  }
  if (str.startsWith('>')) {
    const num = parseFloat(str.slice(1));
    return { operator: '>', value: str.slice(1), isNumeric: !isNaN(num), numValue: num };
  }
  if (str.startsWith('<')) {
    const num = parseFloat(str.slice(1));
    return { operator: '<', value: str.slice(1), isNumeric: !isNaN(num), numValue: num };
  }
  if (str.startsWith('=')) {
    const val = str.slice(1);
    const num = parseFloat(val);
    return { operator: '=', value: val, isNumeric: !isNaN(num), numValue: num };
  }

  const num = parseFloat(str);
  if (!isNaN(num) && str === String(num)) {
    return { operator: '=', value: num, isNumeric: true, numValue: num };
  }

  return { operator: '=', value: str, isNumeric: false };
}

/**
 * Matches a cell value against a compiled criteria predicate.
 */
export function testPredicate(cellValue: any, pred: CriteriaPredicate): boolean {
  if (pred.isNumeric) {
    const cellNum = typeof cellValue === 'number' ? cellValue : parseFloat(cellValue);
    if (isNaN(cellNum)) return false;

    const target = pred.numValue!;
    switch (pred.operator) {
      case '=': return cellNum === target;
      case '!=': return cellNum !== target;
      case '>': return cellNum > target;
      case '>=': return cellNum >= target;
      case '<': return cellNum < target;
      case '<=': return cellNum <= target;
    }
  }

  // String / text comparison (case-insensitive Excel standard)
  const cellStr = String(cellValue ?? '').toLowerCase();
  const targetStr = String(pred.value).toLowerCase();

  // Wildcard support (* and ?)
  if (targetStr.includes('*') || targetStr.includes('?')) {
    const regexPattern = '^' + targetStr.replace(/\./g, '\\.').replace(/\*/g, '.*').replace(/\?/g, '.') + '$';
    const regex = new RegExp(regexPattern, 'i');
    return pred.operator === '!=' ? !regex.test(cellStr) : regex.test(cellStr);
  }

  if (pred.operator === '!=') return cellStr !== targetStr;
  return cellStr === targetStr;
}

/**
 * Vectorized SUM across a contiguous range of row indices.
 */
export function vectorSum(values: Float64Array | number[], startRow: number, endRow: number): number {
  let sum = 0;
  const len = Math.min(values.length, endRow + 1);
  for (let i = startRow; i < len; i++) {
    const val = values[i];
    if (typeof val === 'number' && !isNaN(val)) {
      sum += val;
    }
  }
  return sum;
}

/**
 * Vectorized COUNT across a contiguous range of row indices.
 */
export function vectorCount(values: any[], startRow: number, endRow: number): number {
  let count = 0;
  const len = Math.min(values.length, endRow + 1);
  for (let i = startRow; i < len; i++) {
    const val = values[i];
    if (val !== undefined && val !== null && val !== '') {
      if (typeof val === 'number' && !isNaN(val)) {
        count++;
      }
    }
  }
  return count;
}

/**
 * High-performance vectorized SUMIFS across columnar arrays.
 * Computes: SUM(sumValues[i]) WHERE criteriaValues1[i] MATCHES criteria1 AND criteriaValues2[i] MATCHES criteria2 ...
 */
export function vectorSumIfs(
  sumValues: Float64Array | number[],
  criteriaPairs: Array<{ values: any[]; predicate: CriteriaPredicate }>,
  rowCount: number
): number {
  let totalSum = 0;
  const len = Math.min(sumValues.length, rowCount);

  // Fast path for 1 criteria
  if (criteriaPairs.length === 1) {
    const { values: cVals, predicate: pred } = criteriaPairs[0];
    for (let i = 0; i < len; i++) {
      if (testPredicate(cVals[i], pred)) {
        const val = sumValues[i];
        if (typeof val === 'number' && !isNaN(val)) {
          totalSum += val;
        }
      }
    }
    return totalSum;
  }

  // Multi-criteria path
  const numPairs = criteriaPairs.length;
  for (let i = 0; i < len; i++) {
    let matchesAll = true;
    for (let p = 0; p < numPairs; p++) {
      if (!testPredicate(criteriaPairs[p].values[i], criteriaPairs[p].predicate)) {
        matchesAll = false;
        break;
      }
    }

    if (matchesAll) {
      const val = sumValues[i];
      if (typeof val === 'number' && !isNaN(val)) {
        totalSum += val;
      }
    }
  }

  return totalSum;
}

/**
 * High-performance vectorized COUNTIFS across columnar arrays.
 */
export function vectorCountIfs(
  criteriaPairs: Array<{ values: any[]; predicate: CriteriaPredicate }>,
  rowCount: number
): number {
  let totalCount = 0;
  const numPairs = criteriaPairs.length;

  if (numPairs === 1) {
    const { values: cVals, predicate: pred } = criteriaPairs[0];
    const len = Math.min(cVals.length, rowCount);
    for (let i = 0; i < len; i++) {
      if (testPredicate(cVals[i], pred)) {
        totalCount++;
      }
    }
    return totalCount;
  }

  const len = Math.min(criteriaPairs[0].values.length, rowCount);
  for (let i = 0; i < len; i++) {
    let matchesAll = true;
    for (let p = 0; p < numPairs; p++) {
      if (!testPredicate(criteriaPairs[p].values[i], criteriaPairs[p].predicate)) {
        matchesAll = false;
        break;
      }
    }
    if (matchesAll) {
      totalCount++;
    }
  }

  return totalCount;
}
