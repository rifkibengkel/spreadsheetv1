import { DatasetRowResolver } from '../../src/lib/largeCsv/query/datasetRowResolver';
import { DatasetRangeReader } from '../../src/lib/largeCsv/reader/datasetRangeReader';

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
  console.log('   PHASE 1C — TEST 3: DATASET ROW RESOLVER TESTS');
  console.log('='.repeat(80));

  // 1. Empty input
  console.log('\n[Test 1] Empty input');
  const emptyRes = await DatasetRowResolver.resolveRows(DATASET_ID, []);
  assert(Array.isArray(emptyRes) && emptyRes.length === 0, 'Empty input returns empty array');

  // 2. Contiguous rows in partition 0
  console.log('\n[Test 2] Contiguous rows in partition 0 (rows 100..104)');
  const contigIds = [100, 101, 102, 103, 104];
  const contigRes = await DatasetRowResolver.resolveRows(DATASET_ID, contigIds);
  assert(contigRes.length === 5, 'Returned 5 rows');
  assert(contigRes.every((r, idx) => r.stableRowId === contigIds[idx]), 'Returned exact requested stableRowIds in order');
  assert(contigRes.every((r) => r.cells.length === 54), 'All rows have 54 columns');

  // Verify against DatasetRangeReader
  const rangeRef = await DatasetRangeReader.readRange(DATASET_ID, 100, 5);
  let fidelityMatch = true;
  for (let i = 0; i < 5; i++) {
    for (let c = 0; c < 54; c++) {
      if (contigRes[i].cells[c] !== rangeRef.rows[i].cells[c]) {
        fidelityMatch = false;
        console.error(`Mismatch at row ${i} col ${c}: "${contigRes[i].cells[c]}" vs "${rangeRef.rows[i].cells[c]}"`);
      }
    }
  }
  assert(fidelityMatch, '100% cell fidelity match with DatasetRangeReader');

  // 3. Scattered and Reverse-ordered rows
  console.log('\n[Test 3] Scattered and reverse-ordered rows');
  const revIds = [5000, 100, 3000, 50];
  const revRes = await DatasetRowResolver.resolveRows(DATASET_ID, revIds);
  assert(revRes.length === 4, 'Returned 4 rows');
  assert(revRes[0].stableRowId === 5000, 'Index 0 is stableRowId 5000');
  assert(revRes[1].stableRowId === 100, 'Index 1 is stableRowId 100');
  assert(revRes[2].stableRowId === 3000, 'Index 2 is stableRowId 3000');
  assert(revRes[3].stableRowId === 50, 'Index 3 is stableRowId 50');

  // 4. Cross-partition rows (all 6 partitions)
  console.log('\n[Test 4] Cross-partition rows spanning all 6 partitions');
  const crossIds = [
    0,            // Part 0 first row
    1_999_999,    // Part 0 last row
    2_000_000,    // Part 1 first row
    4_000_000,    // Part 2 first row
    6_000_000,    // Part 3 first row
    8_000_000,    // Part 4 first row
    10_000_000,   // Part 5 first row
    11_809_232,   // Part 5 last row (terminal row of full 11.8M dataset!)
  ];

  const crossRes = await DatasetRowResolver.resolveRows(DATASET_ID, crossIds);
  assert(crossRes.length === crossIds.length, `Returned exact ${crossIds.length} rows`);
  assert(crossRes.every((r, idx) => r.stableRowId === crossIds[idx]), 'All cross-partition row IDs match requested order');
  assert(crossRes[crossRes.length - 1].stableRowId === 11_809_232, 'Successfully resolved final row 11,809,232');
  assert(crossRes.every((r) => r.cells.length === 54), 'All cross-partition rows have 54 columns');

  console.log('\n' + '='.repeat(80));
  console.log(`   ROW RESOLVER RESULTS: ${passed} PASSED, ${failed} FAILED`);
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
