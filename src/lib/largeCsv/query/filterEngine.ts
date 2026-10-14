import fs from 'fs';
import { parseCsvLine } from '../parser/csvLineParser';
import {
  FilterPredicate,
  FilterDefinition,
  DatasetQueryError,
} from './types';

export interface CompiledPredicate {
  colIndex: number;
  predicate: FilterPredicate;
  valueSet?: Set<string>;
  lowerValueSet?: Set<string>;
}

export class FilterEngine {
  /**
   * Compiles filter predicates into fast columnar index lookups.
   */
  public static compileFilter(
    filter: FilterDefinition | undefined,
    headers: string[]
  ): { compiled: CompiledPredicate[]; conjunction: 'AND' | 'OR' } {
    if (!filter || !filter.predicates || filter.predicates.length === 0) {
      return { compiled: [], conjunction: 'AND' };
    }

    const headerMap = new Map<string, number>();
    for (let i = 0; i < headers.length; i++) {
      headerMap.set(headers[i], i);
    }

    const compiled: CompiledPredicate[] = [];

    for (const pred of filter.predicates) {
      if (!pred.column) {
        throw new DatasetQueryError('Filter predicate missing column name', 400, 'INVALID_PREDICATE');
      }

      const colIndex = headerMap.get(pred.column);
      if (colIndex === undefined) {
        throw new DatasetQueryError(
          `Column "${pred.column}" not found in dataset headers`,
          400,
          'COLUMN_NOT_FOUND'
        );
      }

      let valueSet: Set<string> | undefined;
      let lowerValueSet: Set<string> | undefined;

      if (pred.operator === 'IN') {
        const values = Array.isArray(pred.value) ? pred.value : [pred.value || ''];
        valueSet = new Set(values);
        lowerValueSet = new Set(values.map((v) => v.toLowerCase()));
      }

      compiled.push({
        colIndex,
        predicate: pred,
        valueSet,
        lowerValueSet,
      });
    }

    return {
      compiled,
      conjunction: filter.conjunction || 'AND',
    };
  }

  /**
   * Evaluates a single predicate against an exact cell string.
   * NEVER mutates, truncates, or implicitly coerces values unless explicitly requested.
   */
  public static evaluatePredicate(cellValue: string, comp: CompiledPredicate): boolean {
    const { predicate: pred, valueSet, lowerValueSet } = comp;
    const rawValue = cellValue !== undefined && cellValue !== null ? cellValue : '';

    switch (pred.operator) {
      case 'IS_EMPTY':
        return rawValue === '';

      case 'IS_NOT_EMPTY':
        return rawValue !== '';

      case 'EQUALS': {
        const target = pred.value !== undefined ? String(pred.value) : '';
        if (pred.numeric) {
          const numCell = Number(rawValue);
          const numTarget = Number(target);
          if (isNaN(numCell) || isNaN(numTarget)) return false;
          return numCell === numTarget;
        }
        return pred.caseSensitive ? rawValue === target : rawValue.toLowerCase() === target.toLowerCase();
      }

      case 'NOT_EQUALS': {
        const target = pred.value !== undefined ? String(pred.value) : '';
        if (pred.numeric) {
          const numCell = Number(rawValue);
          const numTarget = Number(target);
          if (isNaN(numCell) || isNaN(numTarget)) return false;
          return numCell !== numTarget;
        }
        return pred.caseSensitive ? rawValue !== target : rawValue.toLowerCase() !== target.toLowerCase();
      }

      case 'CONTAINS': {
        const target = pred.value !== undefined ? String(pred.value) : '';
        if (target === '') return true;
        return pred.caseSensitive
          ? rawValue.includes(target)
          : rawValue.toLowerCase().includes(target.toLowerCase());
      }

      case 'NOT_CONTAINS': {
        const target = pred.value !== undefined ? String(pred.value) : '';
        if (target === '') return false;
        return pred.caseSensitive
          ? !rawValue.includes(target)
          : !rawValue.toLowerCase().includes(target.toLowerCase());
      }

      case 'STARTS_WITH': {
        const target = pred.value !== undefined ? String(pred.value) : '';
        return pred.caseSensitive
          ? rawValue.startsWith(target)
          : rawValue.toLowerCase().startsWith(target.toLowerCase());
      }

      case 'ENDS_WITH': {
        const target = pred.value !== undefined ? String(pred.value) : '';
        return pred.caseSensitive
          ? rawValue.endsWith(target)
          : rawValue.toLowerCase().endsWith(target.toLowerCase());
      }

      case 'GREATER_THAN': {
        const target = pred.value !== undefined ? String(pred.value) : '';
        if (pred.numeric) {
          const numCell = Number(rawValue);
          const numTarget = Number(target);
          if (isNaN(numCell) || isNaN(numTarget)) return false;
          return numCell > numTarget;
        }
        return rawValue > target;
      }

      case 'GREATER_THAN_OR_EQUAL': {
        const target = pred.value !== undefined ? String(pred.value) : '';
        if (pred.numeric) {
          const numCell = Number(rawValue);
          const numTarget = Number(target);
          if (isNaN(numCell) || isNaN(numTarget)) return false;
          return numCell >= numTarget;
        }
        return rawValue >= target;
      }

      case 'LESS_THAN': {
        const target = pred.value !== undefined ? String(pred.value) : '';
        if (pred.numeric) {
          const numCell = Number(rawValue);
          const numTarget = Number(target);
          if (isNaN(numCell) || isNaN(numTarget)) return false;
          return numCell < numTarget;
        }
        return rawValue < target;
      }

      case 'LESS_THAN_OR_EQUAL': {
        const target = pred.value !== undefined ? String(pred.value) : '';
        if (pred.numeric) {
          const numCell = Number(rawValue);
          const numTarget = Number(target);
          if (isNaN(numCell) || isNaN(numTarget)) return false;
          return numCell <= numTarget;
        }
        return rawValue <= target;
      }

      case 'IN': {
        if (pred.caseSensitive) {
          return valueSet ? valueSet.has(rawValue) : false;
        }
        return lowerValueSet ? lowerValueSet.has(rawValue.toLowerCase()) : false;
      }

      default:
        return false;
    }
  }

  /**
   * Evaluates compiled predicates across cells of a row.
   */
  public static evaluateRow(
    cells: string[],
    compiled: CompiledPredicate[],
    conjunction: 'AND' | 'OR'
  ): boolean {
    if (compiled.length === 0) return true;

    if (conjunction === 'AND') {
      for (let i = 0; i < compiled.length; i++) {
        const comp = compiled[i];
        const cellValue = comp.colIndex < cells.length ? cells[comp.colIndex] : '';
        if (!this.evaluatePredicate(cellValue, comp)) {
          return false;
        }
      }
      return true;
    } else {
      for (let i = 0; i < compiled.length; i++) {
        const comp = compiled[i];
        const cellValue = comp.colIndex < cells.length ? cells[comp.colIndex] : '';
        if (this.evaluatePredicate(cellValue, comp)) {
          return true;
        }
      }
      return false;
    }
  }

  /**
   * Asynchronously streams through a partition in bounded row chunks (default: 25,000 rows),
   * parsing lines and applying compiled filter predicates.
   * Guarantees zero descriptor leaks and bounded memory.
   */
  public static async scanPartition(
    part: {
      partitionNumber: number;
      stableRowStart: bigint | number;
      rowCount: number;
      dataFilePath: string;
      indexFilePath: string;
      byteSize: bigint | number;
    },
    compiled: CompiledPredicate[],
    conjunction: 'AND' | 'OR',
    onMatch: (stableRowId: number, cells: string[]) => Promise<void> | void,
    options: {
      chunkSize?: number;
      abortSignal?: AbortSignal;
    } = {}
  ): Promise<{ scannedRows: number; matchedRows: number }> {
    const chunkSize = options.chunkSize || 25_000;
    const stableRowStart = Number(part.stableRowStart);
    const totalPartitionRows = part.rowCount;

    let scannedRows = 0;
    let matchedRows = 0;

    let idxHandle: fs.promises.FileHandle | null = null;
    let binHandle: fs.promises.FileHandle | null = null;

    try {
      idxHandle = await fs.promises.open(part.indexFilePath, 'r');
      binHandle = await fs.promises.open(part.dataFilePath, 'r');

      let rowOffset = 0;
      while (rowOffset < totalPartitionRows) {
        if (options.abortSignal?.aborted) {
          throw new DatasetQueryError('Query execution was aborted by client', 499, 'QUERY_ABORTED');
        }

        const subCount = Math.min(chunkSize, totalPartitionRows - rowOffset);
        const idxBytesToRead = (subCount + 1) * 8;
        const idxBuf = Buffer.allocUnsafe(idxBytesToRead);

        const { bytesRead: idxRead } = await idxHandle.read(idxBuf, 0, idxBytesToRead, rowOffset * 8);
        if (idxRead !== idxBytesToRead) {
          throw new DatasetQueryError(
            `Short index read in partition ${part.partitionNumber}: expected ${idxBytesToRead}, got ${idxRead}`,
            500,
            'CORRUPT_INDEX_FILE'
          );
        }

        const startByte = idxBuf.readBigUInt64LE(0);
        const endByte = idxBuf.readBigUInt64LE(subCount * 8);
        const dataBytesToRead = Number(endByte - startByte);

        const dataBuf = Buffer.allocUnsafe(dataBytesToRead);
        if (dataBytesToRead > 0) {
          const { bytesRead: binRead } = await binHandle.read(
            dataBuf,
            0,
            dataBytesToRead,
            Number(startByte)
          );
          if (binRead !== dataBytesToRead) {
            throw new DatasetQueryError(
              `Short data read in partition ${part.partitionNumber}: expected ${dataBytesToRead}, got ${binRead}`,
              500,
              'CORRUPT_DATA_FILE'
            );
          }
        }

        // Process rows in chunk
        for (let i = 0; i < subCount; i++) {
          const lineStart = Number(idxBuf.readBigUInt64LE(i * 8) - startByte);
          const lineEnd = Number(idxBuf.readBigUInt64LE((i + 1) * 8) - startByte);

          let rawLine = dataBuf.subarray(lineStart, lineEnd).toString('utf8');
          if (rawLine.endsWith('\n')) {
            rawLine = rawLine.slice(0, -1);
            if (rawLine.endsWith('\r')) {
              rawLine = rawLine.slice(0, -1);
            }
          }

          const cells = parseCsvLine(rawLine);
          const stableRowId = stableRowStart + rowOffset + i;
          scannedRows++;

          if (this.evaluateRow(cells, compiled, conjunction)) {
            matchedRows++;
            await onMatch(stableRowId, cells);
          }
        }

        rowOffset += subCount;
      }

      return { scannedRows, matchedRows };
    } finally {
      if (idxHandle) {
        try {
          await idxHandle.close();
        } catch {}
      }
      if (binHandle) {
        try {
          await binHandle.close();
        } catch {}
      }
    }
  }
}
