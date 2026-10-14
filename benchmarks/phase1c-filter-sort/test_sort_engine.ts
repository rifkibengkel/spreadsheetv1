import path from 'path';
import fs from 'fs';
import { ExternalSortEngine } from '../../src/lib/largeCsv/query/externalSortEngine';
import { SortDefinition } from '../../src/lib/largeCsv/query/types';

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
  console.log('   PHASE 1C — TEST 2: EXTERNAL SORT ENGINE TESTS');
  console.log('='.repeat(80));

  const tempDir = path.join(process.cwd(), 'storage', 'test_sort_scratch');
  await fs.promises.mkdir(tempDir, { recursive: true });

  try {
    // ---------------------------------------------------------------------------
    // TEST 1: In-memory Fast Path (Single Chunk)
    // ---------------------------------------------------------------------------
    console.log('\n[Test 1] In-Memory Fast Path (Single Chunk)');
    const sortDef1: SortDefinition = {
      columns: [{ column: 'Amount', direction: 'ASC', type: 'number' }],
    };
    const sorter1 = new ExternalSortEngine(sortDef1, tempDir, 1000);

    const items1 = [
      { stableRowId: 10, sortKeys: [100] },
      { stableRowId: 5, sortKeys: [25] },
      { stableRowId: 2, sortKeys: [50] },
      { stableRowId: 8, sortKeys: [50] }, // duplicate key!
    ];

    for (const item of items1) {
      await sorter1.addItem(item);
    }

    const sortedIds1: number[] = [];
    await sorter1.streamSortedIds((id) => {
      sortedIds1.push(id);
    });

    assert(sortedIds1.length === 4, 'Emitted all 4 items');
    assert(sortedIds1[0] === 5, 'First item is amount 25 (stableRowId 5)');
    assert(sortedIds1[1] === 2, 'Second item is amount 50 tie-break 1 (stableRowId 2)');
    assert(sortedIds1[2] === 8, 'Third item is amount 50 tie-break 2 (stableRowId 8)');
    assert(sortedIds1[3] === 10, 'Fourth item is amount 100 (stableRowId 10)');

    // ---------------------------------------------------------------------------
    // TEST 2: Descending Sort & String Keys
    // ---------------------------------------------------------------------------
    console.log('\n[Test 2] Descending String Sort');
    const sortDef2: SortDefinition = {
      columns: [{ column: 'State', direction: 'DESC', type: 'string' }],
    };
    const sorter2 = new ExternalSortEngine(sortDef2, tempDir, 1000);

    await sorter2.addItem({ stableRowId: 1, sortKeys: ['CA'] });
    await sorter2.addItem({ stableRowId: 2, sortKeys: ['NY'] });
    await sorter2.addItem({ stableRowId: 3, sortKeys: ['TX'] });
    await sorter2.addItem({ stableRowId: 4, sortKeys: ['NY'] }); // tie

    const sortedIds2: number[] = [];
    await sorter2.streamSortedIds((id) => {
      sortedIds2.push(id);
    });

    assert(sortedIds2[0] === 3, 'First item TX (DESC)');
    assert(sortedIds2[1] === 2, 'Second item NY (tie-break stableRowId 2)');
    assert(sortedIds2[2] === 4, 'Third item NY (tie-break stableRowId 4)');
    assert(sortedIds2[3] === 1, 'Fourth item CA');

    // ---------------------------------------------------------------------------
    // TEST 3: Multi-Run External Merge Sort (Force Small Chunks)
    // ---------------------------------------------------------------------------
    console.log('\n[Test 3] Multi-Run External Merge Sort (10,000 items, chunk size = 500 -> 20 runs)');
    const sortDef3: SortDefinition = {
      columns: [{ column: 'Val', direction: 'ASC', type: 'number' }],
    };
    const CHUNK_SIZE = 500;
    const TOTAL_ITEMS = 10_000;
    const sorter3 = new ExternalSortEngine(sortDef3, tempDir, CHUNK_SIZE);

    // Generate pseudo-random numbers
    let seed = 123456789;
    const rng = () => {
      seed = (seed * 1103515245 + 12345) & 0x7fffffff;
      return seed % 1000; // lots of duplicates to test tie-breaker!
    };

    const expectedPairs: { id: number; val: number }[] = [];
    for (let i = 0; i < TOTAL_ITEMS; i++) {
      const val = rng();
      expectedPairs.push({ id: i, val });
      await sorter3.addItem({ stableRowId: i, sortKeys: [val] });
    }

    // Sort expected using JS
    expectedPairs.sort((a, b) => {
      if (a.val !== b.val) return a.val - b.val;
      return a.id - b.id;
    });

    const sortedIds3: number[] = [];
    await sorter3.streamSortedIds((id) => {
      sortedIds3.push(id);
    });

    assert(sortedIds3.length === TOTAL_ITEMS, `Emitted exactly ${TOTAL_ITEMS} items`);

    let allMatch = true;
    for (let i = 0; i < TOTAL_ITEMS; i++) {
      if (sortedIds3[i] !== expectedPairs[i].id) {
        allMatch = false;
        console.error(`Mismatch at ${i}: expected id ${expectedPairs[i].id}, got ${sortedIds3[i]}`);
        break;
      }
    }
    assert(allMatch, 'External K-Way Merge matches exact expected sorted order with deterministic tie-breaking');

    // Verify temp files were cleaned up
    const remainingFiles = await fs.promises.readdir(tempDir);
    const runFiles = remainingFiles.filter((f) => f.startsWith('sort_run_'));
    assert(runFiles.length === 0, `All temporary run files safely cleaned up (found ${runFiles.length})`);

  } finally {
    try {
      await fs.promises.rm(tempDir, { recursive: true, force: true });
    } catch {}
  }

  console.log('\n' + '='.repeat(80));
  console.log(`   SORT ENGINE RESULTS: ${passed} PASSED, ${failed} FAILED`);
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
