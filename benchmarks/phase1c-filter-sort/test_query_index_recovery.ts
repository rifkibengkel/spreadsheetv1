import path from 'path';
import fs from 'fs';
import { QueryIndexManager } from '../../src/lib/largeCsv/query/queryIndexManager';
import { QueryDefinition, QUERY_INDEX_MAGIC, QUERY_INDEX_HEADER_SIZE } from '../../src/lib/largeCsv/query/types';

const DATASET_ID = '80af4daa-3c4a-42eb-b656-dd855dc6f866';

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

async function runTests() {
  console.log('='.repeat(80));
  console.log('   PHASE 1C — TEST 4: QUERY INDEX ATOMIC PROMOTION & RECOVERY TESTS');
  console.log('='.repeat(80));

  const queriesDir = QueryIndexManager.getQueriesDir(DATASET_ID);
  await fs.promises.mkdir(queriesDir, { recursive: true });

  // ---------------------------------------------------------------------------
  // TEST 1: Two-stage Atomic Promotion (Real Query Execution)
  // ---------------------------------------------------------------------------
  console.log('\n[Test 1] Two-stage Atomic Promotion');

  // Let's create a targeted filter query: Registration_State == "CT" (Connecticut plates)
  const queryDef1: QueryDefinition = {
    filter: {
      conjunction: 'AND',
      predicates: [{ column: 'Registration_State', operator: 'EQUALS', value: 'CT' }],
    },
  };

  const queryId1 = QueryIndexManager.generateQueryId(DATASET_ID, queryDef1);
  console.log(`  Generated deterministic queryId: ${queryId1}`);

  // Clean up any pre-existing query files for clean test
  await QueryIndexManager.deleteQuery(DATASET_ID, queryId1);

  const meta1 = await QueryIndexManager.executeQuery(DATASET_ID, queryDef1);
  assert(meta1.status === 'READY', 'Query status is READY');
  assert(meta1.matchedRows > 0, `Matched CT plates in dataset: ${meta1.matchedRows}`);
  assert(meta1.totalDatasetRows === 11_809_233, 'Total dataset rows recorded as 11,809,233');

  const finalQidxPath = path.join(queriesDir, `query_${queryId1}.qidx`);
  const finalMetaPath = path.join(queriesDir, `query_${queryId1}.meta.json`);

  assert(fs.existsSync(finalQidxPath), 'Final .qidx file exists on disk');
  assert(fs.existsSync(finalMetaPath), 'Final .meta.json file exists on disk');

  // Verify file size: exactly 2048 + matchedRows * 4
  const stats = await fs.promises.stat(finalQidxPath);
  const expectedSize = QUERY_INDEX_HEADER_SIZE + meta1.matchedRows * 4;
  assert(stats.size === expectedSize, `File size matches exact contract: ${stats.size} == ${expectedSize}`);

  // ---------------------------------------------------------------------------
  // TEST 2: Cache Hit on Identical Query
  // ---------------------------------------------------------------------------
  console.log('\n[Test 2] Cache Hit on Identical Query');
  const t0 = Date.now();
  const cachedMeta = await QueryIndexManager.executeQuery(DATASET_ID, queryDef1);
  const t1 = Date.now();
  assert(cachedMeta.queryId === queryId1, 'Returned same query ID');
  assert(cachedMeta.status === 'READY', 'Cached query is READY');
  assert(t1 - t0 < 50, `Cache hit returned instantly in ${t1 - t0}ms`);

  // ---------------------------------------------------------------------------
  // TEST 3: Bounded Random Access Read
  // ---------------------------------------------------------------------------
  console.log('\n[Test 3] Bounded Random-Access Slice Read');
  const slice1 = await QueryIndexManager.readStableRowIds(DATASET_ID, queryId1, 0, 100);
  assert(slice1.stableRowIds.length === 100, 'Read first 100 stableRowIds');
  assert(slice1.totalMatchedRows === meta1.matchedRows, 'totalMatchedRows matches metadata');

  const slice2 = await QueryIndexManager.readStableRowIds(DATASET_ID, queryId1, 50, 50);
  assert(slice2.stableRowIds.length === 50, 'Read middle 50 stableRowIds');
  // Check overlap with slice1
  let overlapMatch = true;
  for (let i = 0; i < 50; i++) {
    if (slice2.stableRowIds[i] !== slice1.stableRowIds[50 + i]) {
      overlapMatch = false;
      break;
    }
  }
  assert(overlapMatch, 'Slice [50..99] matches exact row IDs from slice [0..99]');

  // ---------------------------------------------------------------------------
  // TEST 4: Corrupted Index Rejection
  // ---------------------------------------------------------------------------
  console.log('\n[Test 4] Corrupted Index Rejection');
  const corruptPath = path.join(queriesDir, 'corrupt_test.qidx');
  const corruptBuf = Buffer.alloc(2048 + 400); // 100 rows
  corruptBuf.write('INVALID_MAGIC!!', 0, 16, 'ascii');
  await fs.promises.writeFile(corruptPath, corruptBuf);

  const isValidCorrupt = await QueryIndexManager.validateIndexFile(corruptPath, 100);
  assert(!isValidCorrupt, 'Corrupted magic correctly rejected by validateIndexFile');
  await fs.promises.unlink(corruptPath);

  // Truncated size rejection
  const truncatedPath = path.join(queriesDir, 'truncated_test.qidx');
  const truncBuf = Buffer.alloc(2048 + 200); // expected 400, but only 200
  truncBuf.write(QUERY_INDEX_MAGIC, 0, 16, 'ascii');
  truncBuf.writeUInt32LE(1, 16);
  truncBuf.writeUInt32LE(1, 20);
  await fs.promises.writeFile(truncatedPath, truncBuf);

  const isValidTruncated = await QueryIndexManager.validateIndexFile(truncatedPath, 100);
  assert(!isValidTruncated, 'Truncated file size correctly rejected by validateIndexFile');
  await fs.promises.unlink(truncatedPath);

  // ---------------------------------------------------------------------------
  // TEST 5: Interrupted Generation Cleanup
  // ---------------------------------------------------------------------------
  console.log('\n[Test 5] Interrupted Generation / Abort Handling');
  const abortController = new AbortController();
  // Abort immediately
  abortController.abort();

  const queryDefAbort: QueryDefinition = {
    filter: {
      conjunction: 'AND',
      predicates: [{ column: 'Plate_ID', operator: 'CONTAINS', value: 'XYZ999' }],
    },
  };

  let caughtAbort = false;
  try {
    await QueryIndexManager.executeQuery(DATASET_ID, queryDefAbort, {
      abortSignal: abortController.signal,
    });
  } catch (err: any) {
    caughtAbort = true;
    assert(err.code === 'QUERY_ABORTED', `Caught aborted error: ${err.code}`);
  }
  assert(caughtAbort, 'Aborted query threw expected error');

  // Verify no orphaned temporary file exists
  const files = await fs.promises.readdir(queriesDir);
  const orphanedTmp = files.filter((f) => f.startsWith('tmp_query_'));
  assert(orphanedTmp.length === 0, `No orphaned temporary files remain after abort (found ${orphanedTmp.length})`);

  console.log('\n' + '='.repeat(80));
  console.log(`   QUERY INDEX & RECOVERY RESULTS: ${passed} PASSED, ${failed} FAILED`);
  console.log('='.repeat(80));

  if (failed > 0) {
    process.exit(1);
  }
}

runTests().catch((err) => {
  console.error('Test execution fatal error:', err);
  process.exit(1);
}).finally(() => {
  process.exit(0);
});
