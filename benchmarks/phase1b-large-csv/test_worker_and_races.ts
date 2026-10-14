import { BoundedLruCache, ViewportPage } from '../../src/lib/viewport/boundedLruCache';

let passed = 0;
let failed = 0;

function assert(condition: boolean, name: string, detail?: any) {
  if (condition) {
    console.log(`  [PASS] ${name}`);
    passed++;
  } else {
    console.error(`  [FAIL] ${name}`, detail !== undefined ? detail : '');
    failed++;
  }
}

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function runWorkerAndRaceTests() {
  console.log('='.repeat(80));
  console.log('   PHASE 1B — STEP 7: WORKER CACHE & RACE PROTECTION TESTS');
  console.log('='.repeat(80));

  // ---------------------------------------------------------------------------
  // 1. Bounded LRU Cache Verification
  // ---------------------------------------------------------------------------
  console.log('\n[1. Bounded LRU Cache Tests]');
  const cache = new BoundedLruCache(5, 100); // max 5 pages, 100 rows per page
  assert(cache.maxPages === 5, 'Cache maxPages configured to 5');

  // Insert 5 pages
  for (let i = 0; i < 5; i++) {
    cache.set('dataset-1', i, {
      pageIndex: i,
      startRow: i * 100,
      rowCount: 100,
      rows: [{ stableRowId: i * 100, cells: ['test'] }],
      cachedAt: Date.now(),
    });
  }

  let stats = cache.getStats();
  assert(stats.size === 5, `Cache size is 5 (got ${stats.size})`);
  assert(stats.evictions === 0, `Evictions count is 0 (got ${stats.evictions})`);

  // Hit page 0 to refresh its LRU position (making page 1 the oldest)
  const page0 = cache.get('dataset-1', 0);
  assert(page0 !== undefined && page0.pageIndex === 0, 'Cache hit for page 0');
  stats = cache.getStats();
  assert(stats.hits === 1, `Hits count is 1 (got ${stats.hits})`);

  // Add 6th page (page 5). Oldest should be page 1, since page 0 was accessed!
  cache.set('dataset-1', 5, {
    pageIndex: 5,
    startRow: 500,
    rowCount: 100,
    rows: [{ stableRowId: 500, cells: ['test5'] }],
    cachedAt: Date.now(),
  });

  stats = cache.getStats();
  assert(stats.size === 5, `Cache size remains bounded at 5 (got ${stats.size})`);
  assert(stats.evictions === 1, `Evictions count is 1 (got ${stats.evictions})`);
  assert(!cache.has('dataset-1', 1), 'Page 1 was evicted (least recently used)');
  assert(cache.has('dataset-1', 0), 'Page 0 was retained (recently accessed)');
  assert(cache.has('dataset-1', 5), 'Page 5 was added');

  // Miss check
  const missing = cache.get('dataset-1', 99);
  assert(missing === undefined, 'Non-existent page returns undefined');
  stats = cache.getStats();
  assert(stats.misses === 1, `Misses count is 1 (got ${stats.misses})`);

  // Invalidate by datasetId
  cache.set('dataset-2', 0, {
    pageIndex: 0,
    startRow: 0,
    rowCount: 100,
    rows: [],
    cachedAt: Date.now(),
  });
  assert(cache.has('dataset-2', 0), 'Dataset-2 page present');
  cache.invalidate('dataset-1');
  assert(!cache.has('dataset-1', 0), 'Dataset-1 pages removed on invalidate');
  assert(cache.has('dataset-2', 0), 'Dataset-2 page intact after Dataset-1 invalidation');

  // ---------------------------------------------------------------------------
  // 2. Race Condition / Out-Of-Order Response Protection Simulation
  // ---------------------------------------------------------------------------
  console.log('\n[2. Out-of-Order / Stale Request Protection Tests]');

  // Viewport State Machine simulation
  class ViewportClientSimulator {
    public currentGenerationToken = 0;
    public hydratedStartRow: number | null = null;
    public hydrationCount = 0;
    public droppedCount = 0;

    public requestWindow(startRow: number, rowCount: number, delayMs: number) {
      const token = ++this.currentGenerationToken;

      // Simulate async worker fetch with delayMs
      return (async () => {
        await delay(delayMs);

        // Verification check: Is this response still current?
        if (token !== this.currentGenerationToken) {
          this.droppedCount++;
          return { dropped: true, token };
        }

        this.hydratedStartRow = startRow;
        this.hydrationCount++;
        return { dropped: false, token, startRow };
      })();
    }
  }

  const client = new ViewportClientSimulator();

  // SCENARIO 1: Request A starts with 60ms delay, Request B starts with 15ms delay
  // B completes BEFORE A. When A completes, it must NOT overwrite B.
  console.log('  Triggering Request A (startRow=1000, delay=60ms)...');
  const promiseA = client.requestWindow(1000, 100, 60);

  await delay(5); // Slight tick

  console.log('  Triggering Request B (startRow=5000, delay=15ms)...');
  const promiseB = client.requestWindow(5000, 100, 15);

  const [resA, resB] = await Promise.all([promiseA, promiseB]);

  assert(resB.dropped === false, 'Request B (newer) was accepted');
  assert(resB.startRow === 5000, 'Request B set viewport to startRow=5000');
  assert(resA.dropped === true, 'Request A (older, delayed) was dropped');
  assert(client.hydratedStartRow === 5000, 'Client state maintained Request B (5000), not overwritten by A (1000)');
  assert(client.droppedCount === 1, 'Dropped count is exactly 1');

  // SCENARIO 2: Rapid scrolling (10 consecutive requests in 10ms intervals with varying network delays)
  console.log('  Testing rapid scroll burst (10 concurrent requests)...');
  client.droppedCount = 0;
  client.hydrationCount = 0;

  const promises = [];
  for (let i = 1; i <= 10; i++) {
    // Delays vary: earlier requests have longer delays (worst-case network jitter)
    const simulatedDelay = (11 - i) * 10;
    promises.push(client.requestWindow(i * 1000, 100, simulatedDelay));
  }

  await Promise.all(promises);

  assert(
    client.hydratedStartRow === 10_000,
    `Final client state matches exactly the 10th request (startRow=10000, got ${client.hydratedStartRow})`
  );
  assert(client.droppedCount === 9, `9 obsolete requests dropped (got ${client.droppedCount})`);
  assert(client.hydrationCount === 1, `Only the latest request hydrated (got ${client.hydrationCount})`);

  // ---------------------------------------------------------------------------
  // 3. AbortController Network Cancellation Simulation
  // ---------------------------------------------------------------------------
  console.log('\n[3. AbortController Cancellation Tests]');
  let priorAbortController: AbortController | null = null;
  let abortedCount = 0;

  function simulateNetworkFetch(token: number): Promise<string> {
    if (priorAbortController) {
      priorAbortController.abort();
      abortedCount++;
    }
    const ac = new AbortController();
    priorAbortController = ac;

    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        resolve(`Success for token ${token}`);
      }, 50);

      ac.signal.addEventListener('abort', () => {
        clearTimeout(timer);
        reject(new Error('AbortError'));
      });
    });
  }

  const p1 = simulateNetworkFetch(1).catch((e) => e.message);
  const p2 = simulateNetworkFetch(2).catch((e) => e.message);

  const [resP1, resP2] = await Promise.all([p1, p2]);
  assert(resP1 === 'AbortError', 'In-flight fetch 1 was aborted');
  assert(resP2 === 'Success for token 2', 'Latest fetch 2 completed successfully');
  assert(abortedCount === 1, 'Abort signal fired exactly once');

  console.log('\n' + '='.repeat(80));
  console.log(`STEP 7 TEST RESULTS: ${passed} PASSED, ${failed} FAILED`);
  console.log('='.repeat(80));

  if (failed > 0) {
    process.exit(1);
  }
}

runWorkerAndRaceTests().catch((err) => {
  console.error('Fatal test error:', err);
  process.exit(1);
}).finally(() => process.exit(0));
