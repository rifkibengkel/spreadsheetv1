import { DatasetRow } from '../largeCsv/reader/datasetRangeReader';

export interface ViewportPage {
  pageIndex: number;
  startRow: number;
  rowCount: number;
  rows: DatasetRow[];
  cachedAt: number;
}

export interface CacheStats {
  size: number;
  maxPages: number;
  pageSize: number;
  hits: number;
  misses: number;
  evictions: number;
}

export class BoundedLruCache {
  public readonly maxPages: number;
  public readonly pageSize: number;
  private cache: Map<string, ViewportPage>;
  private hits: number = 0;
  private misses: number = 0;
  private evictions: number = 0;

  constructor(maxPages = 25, pageSize = 200) {
    this.maxPages = maxPages;
    this.pageSize = pageSize;
    this.cache = new Map<string, ViewportPage>();
  }

  private buildKey(datasetId: string, pageIndex: number): string {
    return `${datasetId}:${pageIndex}`;
  }

  public get(datasetId: string, pageIndex: number): ViewportPage | undefined {
    const key = this.buildKey(datasetId, pageIndex);
    const entry = this.cache.get(key);
    if (!entry) {
      this.misses++;
      return undefined;
    }

    // Refresh LRU position by re-inserting to end of Map
    this.cache.delete(key);
    this.cache.set(key, entry);
    this.hits++;
    return entry;
  }

  public set(datasetId: string, pageIndex: number, page: ViewportPage): void {
    const key = this.buildKey(datasetId, pageIndex);

    // If key already exists, delete to update position
    if (this.cache.has(key)) {
      this.cache.delete(key);
    } else if (this.cache.size >= this.maxPages) {
      // Evict oldest (first key in insertion order)
      const oldestKey = this.cache.keys().next().value;
      if (oldestKey !== undefined) {
        this.cache.delete(oldestKey);
        this.evictions++;
      }
    }

    this.cache.set(key, page);
  }

  public has(datasetId: string, pageIndex: number): boolean {
    return this.cache.has(this.buildKey(datasetId, pageIndex));
  }

  public invalidate(datasetId: string): void {
    const prefix = `${datasetId}:`;
    for (const key of this.cache.keys()) {
      if (key.startsWith(prefix)) {
        this.cache.delete(key);
      }
    }
  }

  public clear(): void {
    this.cache.clear();
    this.hits = 0;
    this.misses = 0;
    this.evictions = 0;
  }

  public getStats(): CacheStats {
    return {
      size: this.cache.size,
      maxPages: this.maxPages,
      pageSize: this.pageSize,
      hits: this.hits,
      misses: this.misses,
      evictions: this.evictions,
    };
  }
}
