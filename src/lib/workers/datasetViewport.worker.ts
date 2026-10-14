/**
 * Dedicated Browser Worker for Range Query & Viewport Window Hydration
 * Architecture Phase 1B:
 * - Maintains bounded LRU cache (max 25 pages x 200 rows = max 5,000 rows retained)
 * - Uses generationToken + requestId for strict race correctness
 * - Distinguishes foreground requests from speculative prefetch requests
 * - Performs all response validation, caching, and payload transformation off the UI thread
 */

import { BoundedLruCache, ViewportPage } from '../viewport/boundedLruCache';
import { DatasetRow } from '../largeCsv/reader/datasetRangeReader';

const PAGE_SIZE = 200;
const MAX_CACHED_PAGES = 25;

const cache = new BoundedLruCache(MAX_CACHED_PAGES, PAGE_SIZE);

let currentForegroundGenerationToken = 0;
let activeForegroundAbortController: AbortController | null = null;
const activePrefetchControllers = new Map<string, AbortController>();

export interface FetchViewportRangeMsg {
  type: 'FETCH_VIEWPORT_RANGE';
  requestId?: string;
  datasetId: string;
  queryId?: string;
  startRow: number;
  rowCount: number;
  generationToken: number;
  isPrefetch?: boolean;
  apiBaseUrl?: string;
}

export interface InvalidateCacheMsg {
  type: 'INVALIDATE_CACHE';
  datasetId?: string;
}

export interface GetStatsMsg {
  type: 'GET_CACHE_STATS';
}

type WorkerIncomingMsg = FetchViewportRangeMsg | InvalidateCacheMsg | GetStatsMsg;

self.onmessage = async (e: MessageEvent<WorkerIncomingMsg>) => {
  const data = e.data;
  if (!data) return;

  if (data.type === 'INVALIDATE_CACHE') {
    if (data.datasetId) {
      cache.invalidate(data.datasetId);
    } else {
      cache.clear();
    }
    self.postMessage({ type: 'CACHE_INVALIDATED' });
    return;
  }

  if (data.type === 'GET_CACHE_STATS') {
    self.postMessage({ type: 'CACHE_STATS', stats: cache.getStats() });
    return;
  }

  if (data.type === 'FETCH_VIEWPORT_RANGE') {
    const t6 = performance.now();
    const {
      requestId = `${data.generationToken}_${data.startRow}`,
      datasetId,
      queryId,
      startRow,
      rowCount,
      generationToken,
      isPrefetch = false,
      apiBaseUrl = '',
    } = data;

    const cacheScope = queryId ? `${datasetId}#${queryId}` : datasetId;

    // For foreground requests: Reject stale tokens immediately and abort prior foreground fetch
    if (!isPrefetch) {
      if (generationToken < currentForegroundGenerationToken) {
        return;
      }
      if (activeForegroundAbortController) {
        activeForegroundAbortController.abort();
        activeForegroundAbortController = null;
      }
      currentForegroundGenerationToken = generationToken;
    }

    const abortController = new AbortController();
    if (!isPrefetch) {
      activeForegroundAbortController = abortController;
    } else {
      activePrefetchControllers.set(requestId, abortController);
    }

    const endRow = startRow + rowCount - 1;
    const startPageIndex = Math.floor(startRow / PAGE_SIZE);
    const endPageIndex = Math.floor(endRow / PAGE_SIZE);

    // 1. Check which pages are cached
    const missingPageIndices: number[] = [];
    const cachedPages = new Map<number, ViewportPage>();

    for (let p = startPageIndex; p <= endPageIndex; p++) {
      const page = cache.get(cacheScope, p);
      if (page) {
        cachedPages.set(p, page);
      } else {
        missingPageIndices.push(p);
      }
    }

    const isFullCacheHit = missingPageIndices.length === 0;
    let serverMetrics: any = null;
    let totalRows = 0;
    let columns: string[] = [];
    let columnCount = 0;

    try {
      // 2. Fetch missing pages if needed
      if (!isFullCacheHit) {
        const fetchStartRow = missingPageIndices[0] * PAGE_SIZE;
        const lastMissingPage = missingPageIndices[missingPageIndices.length - 1];
        const fetchEndRow = (lastMissingPage + 1) * PAGE_SIZE - 1;
        const fetchRowCount = fetchEndRow - fetchStartRow + 1;

        const queryParam = queryId ? `&queryId=${encodeURIComponent(queryId)}` : '';
        const url = `${apiBaseUrl}/api/v1/datasets/${datasetId}/range?startRow=${fetchStartRow}&rowCount=${fetchRowCount}${queryParam}`;
        const response = await fetch(url, {
          signal: abortController.signal,
          headers: {
            Accept: 'application/json',
          },
        });

        if (!response.ok) {
          let errMsg = `HTTP ${response.status} ${response.statusText}`;
          try {
            const errBody = await response.json();
            if (errBody.error) errMsg = errBody.error;
          } catch {}
          throw new Error(errMsg);
        }

        const serverData = await response.json();
        serverMetrics = serverData.metrics;
        totalRows = serverData.totalRows;
        columns = serverData.columns;
        columnCount = serverData.columnCount;

        // Partition fetched rows into pages and populate LRU cache
        const fetchedRows: DatasetRow[] = serverData.rows;
        let rowIdx = 0;
        for (const p of missingPageIndices) {
          const pageStart = p * PAGE_SIZE;
          const pageRows: DatasetRow[] = [];
          while (
            rowIdx < fetchedRows.length &&
            fetchedRows[rowIdx].stableRowId < pageStart + PAGE_SIZE
          ) {
            pageRows.push(fetchedRows[rowIdx]);
            rowIdx++;
          }

          const newPage: ViewportPage = {
            pageIndex: p,
            startRow: pageStart,
            rowCount: pageRows.length,
            rows: pageRows,
            cachedAt: Date.now(),
          };

          cache.set(cacheScope, p, newPage);
          cachedPages.set(p, newPage);
        }
      }

      // Check if foreground request was superseded while waiting for fetch
      if (!isPrefetch && generationToken !== currentForegroundGenerationToken) {
        return;
      }

      // 3. Assemble exact requested rows [startRow .. endRow] from pages
      const resultRows: DatasetRow[] = [];
      for (let p = startPageIndex; p <= endPageIndex; p++) {
        const page = cachedPages.get(p);
        if (!page) continue;
        for (const row of page.rows) {
          if (row.stableRowId >= startRow && row.stableRowId <= endRow) {
            resultRows.push(row);
          }
        }
      }

      const t7 = performance.now();

      self.postMessage({
        type: 'VIEWPORT_DATA',
        requestId,
        datasetId,
        generationToken,
        startRow,
        rowCount: resultRows.length,
        totalRows,
        columns,
        columnCount,
        rows: resultRows,
        isPrefetch,
        isCacheHit: isFullCacheHit,
        metrics: {
          ...(serverMetrics || {}),
          workerDurationMs: t7 - t6,
        },
      });
    } catch (err: any) {
      if (err.name === 'AbortError') {
        // Request intentionally aborted because superseded; discard
        return;
      }

      if (!isPrefetch && generationToken === currentForegroundGenerationToken) {
        self.postMessage({
          type: 'ERROR',
          requestId,
          datasetId,
          generationToken,
          error: err.message || 'Worker range query failed',
        });
      }
    } finally {
      if (!isPrefetch && activeForegroundAbortController === abortController) {
        activeForegroundAbortController = null;
      }
      if (isPrefetch) {
        activePrefetchControllers.delete(requestId);
      }
    }
  }
};
