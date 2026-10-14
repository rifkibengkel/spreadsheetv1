import fs from 'fs';
import path from 'path';
import { SortDefinition, SortColumn, MAX_SORT_CHUNK_SIZE } from './types';

export interface ExternalSortItem {
  stableRowId: number;
  sortKeys: (string | number)[];
}

/**
 * Min-Heap for K-Way Merge of sorted runs.
 */
class MinHeap<T> {
  private heap: T[] = [];

  constructor(private readonly compare: (a: T, b: T) => number) {}

  public get size(): number {
    return this.heap.length;
  }

  public push(item: T): void {
    this.heap.push(item);
    this.bubbleUp(this.heap.length - 1);
  }

  public pop(): T | undefined {
    if (this.heap.length === 0) return undefined;
    const top = this.heap[0];
    const bottom = this.heap.pop()!;
    if (this.heap.length > 0) {
      this.heap[0] = bottom;
      this.sinkDown(0);
    }
    return top;
  }

  public peek(): T | undefined {
    return this.heap[0];
  }

  private bubbleUp(index: number): void {
    const item = this.heap[index];
    while (index > 0) {
      const parentIdx = (index - 1) >> 1;
      const parent = this.heap[parentIdx];
      if (this.compare(item, parent) >= 0) break;
      this.heap[index] = parent;
      index = parentIdx;
    }
    this.heap[index] = item;
  }

  private sinkDown(index: number): void {
    const length = this.heap.length;
    const item = this.heap[index];
    const halfLength = length >> 1;

    while (index < halfLength) {
      let left = (index << 1) + 1;
      let right = left + 1;
      let bestIdx = left;
      let bestItem = this.heap[left];

      if (right < length && this.compare(this.heap[right], bestItem) < 0) {
        bestIdx = right;
        bestItem = this.heap[right];
      }

      if (this.compare(item, bestItem) <= 0) break;
      this.heap[index] = bestItem;
      index = bestIdx;
    }
    this.heap[index] = item;
  }
}

export class ExternalSortEngine {
  private currentChunk: ExternalSortItem[] = [];
  private runFilePaths: string[] = [];
  private totalItems = 0;
  private readonly comparator: (a: ExternalSortItem, b: ExternalSortItem) => number;

  constructor(
    private readonly sortDef: SortDefinition,
    private readonly tempDir: string,
    private readonly maxChunkSize: number = MAX_SORT_CHUNK_SIZE
  ) {
    this.comparator = ExternalSortEngine.createComparator(sortDef);
  }

  /**
   * Deterministic comparator across multiple sort columns with stableRowId tie-breaking.
   */
  public static createComparator(
    sortDef: SortDefinition
  ): (a: ExternalSortItem, b: ExternalSortItem) => number {
    const cols = sortDef.columns;

    return (a: ExternalSortItem, b: ExternalSortItem): number => {
      for (let i = 0; i < cols.length; i++) {
        const col = cols[i];
        const valA = a.sortKeys[i];
        const valB = b.sortKeys[i];

        let diff = 0;
        if (col.type === 'number' || (typeof valA === 'number' && typeof valB === 'number')) {
          const numA = Number(valA);
          const numB = Number(valB);
          const aIsNaN = isNaN(numA);
          const bIsNaN = isNaN(numB);

          if (aIsNaN && bIsNaN) {
            diff = 0;
          } else if (aIsNaN) {
            diff = 1; // NaNs sorted to end
          } else if (bIsNaN) {
            diff = -1;
          } else {
            diff = numA < numB ? -1 : numA > numB ? 1 : 0;
          }
        } else {
          // String comparison
          const strA = valA !== undefined && valA !== null ? String(valA) : '';
          const strB = valB !== undefined && valB !== null ? String(valB) : '';
          if (strA < strB) diff = -1;
          else if (strA > strB) diff = 1;
          else diff = 0;
        }

        if (diff !== 0) {
          return col.direction === 'DESC' ? -diff : diff;
        }
      }

      // Final deterministic tie-breaker: stableRowId ASC
      return a.stableRowId - b.stableRowId;
    };
  }

  /**
   * Adds an item to the sorter. Automatically flushes chunk to disk when maxChunkSize reached.
   */
  public async addItem(item: ExternalSortItem): Promise<void> {
    this.currentChunk.push(item);
    this.totalItems++;

    if (this.currentChunk.length >= this.maxChunkSize) {
      await this.flushRun();
    }
  }

  /**
   * Flushes currently buffered items as a sorted run on disk.
   */
  private async flushRun(): Promise<void> {
    if (this.currentChunk.length === 0) return;

    // Sort chunk in memory
    this.currentChunk.sort(this.comparator);

    await fs.promises.mkdir(this.tempDir, { recursive: true });
    const runPath = path.join(
      this.tempDir,
      `sort_run_${String(this.runFilePaths.length).padStart(4, '0')}.tmp`
    );

    // Serialize run: JSON lines or packed records
    // Using fast line-based format: stableRowId \t key0 \t key1 ... \n
    const lines: string[] = new Array(this.currentChunk.length);
    for (let i = 0; i < this.currentChunk.length; i++) {
      const item = this.currentChunk[i];
      lines[i] = JSON.stringify({
        id: item.stableRowId,
        k: item.sortKeys,
      });
    }

    await fs.promises.writeFile(runPath, lines.join('\n') + '\n', 'utf8');
    this.runFilePaths.push(runPath);
    this.currentChunk = [];
  }

  /**
   * Finalizes the sort and streams sorted stableRowIds to the output writer.
   * Cleans up all temporary run files in a finally block.
   */
  public async streamSortedIds(
    onSortedId: (stableRowId: number) => Promise<void> | void
  ): Promise<number> {
    try {
      // If there are remaining items in currentChunk:
      if (this.runFilePaths.length === 0) {
        // Fast path: all items fit in memory! Zero disk I/O!
        this.currentChunk.sort(this.comparator);
        for (let i = 0; i < this.currentChunk.length; i++) {
          await onSortedId(this.currentChunk[i].stableRowId);
        }
        return this.totalItems;
      }

      // Flush final chunk if any
      if (this.currentChunk.length > 0) {
        await this.flushRun();
      }

      // Multi-way merge across run files
      await this.kWayMerge(onSortedId);
      return this.totalItems;
    } finally {
      await this.cleanup();
    }
  }

  /**
   * Executes K-way merge using MinHeap.
   */
  private async kWayMerge(
    onSortedId: (stableRowId: number) => Promise<void> | void
  ): Promise<void> {
    interface RunReader {
      filePath: string;
      lines: string[];
      lineIndex: number;
    }

    // Load each run file (in chunks or full lines)
    const readers: RunReader[] = [];
    for (const runPath of this.runFilePaths) {
      const content = await fs.promises.readFile(runPath, 'utf8');
      const lines = content.split('\n').filter((l) => l.trim().length > 0);
      readers.push({
        filePath: runPath,
        lines,
        lineIndex: 0,
      });
    }

    const parseLine = (line: string): ExternalSortItem => {
      const parsed = JSON.parse(line);
      return {
        stableRowId: parsed.id,
        sortKeys: parsed.k,
      };
    };

    interface HeapEntry {
      item: ExternalSortItem;
      readerIdx: number;
    }

    const heap = new MinHeap<HeapEntry>((a, b) => this.comparator(a.item, b.item));

    // Seed heap with first item from each run
    for (let r = 0; r < readers.length; r++) {
      const reader = readers[r];
      if (reader.lines.length > 0) {
        const item = parseLine(reader.lines[0]);
        reader.lineIndex = 1;
        heap.push({ item, readerIdx: r });
      }
    }

    // Pop min, emit stableRowId, advance corresponding reader
    while (heap.size > 0) {
      const top = heap.pop()!;
      await onSortedId(top.item.stableRowId);

      const reader = readers[top.readerIdx];
      if (reader.lineIndex < reader.lines.length) {
        const nextItem = parseLine(reader.lines[reader.lineIndex]);
        reader.lineIndex++;
        heap.push({ item: nextItem, readerIdx: top.readerIdx });
      }
    }
  }

  /**
   * Deletes all temporary run files.
   */
  public async cleanup(): Promise<void> {
    for (const p of this.runFilePaths) {
      try {
        await fs.promises.unlink(p);
      } catch {}
    }
    this.runFilePaths = [];
    this.currentChunk = [];
  }
}
