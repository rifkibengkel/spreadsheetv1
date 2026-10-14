import prisma from '../../src/lib/db/prisma';
import { FilterEngine } from '../../src/lib/largeCsv/query/filterEngine';
import { FilterDefinition } from '../../src/lib/largeCsv/query/types';

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
  console.log('   PHASE 1C — TEST 1: FILTER ENGINE UNIT & PARTITION SCAN TESTS');
  console.log('='.repeat(80));

  // 1. Pure Predicate Tests
  console.log('\n[Section 1] Pure Predicate Evaluation');

  const headers = ['ID', 'State', 'Code', 'Amount', 'Date', 'Notes'];
  
  // EQUALS string case insensitive & sensitive
  const eqFilter = FilterEngine.compileFilter(
    { conjunction: 'AND', predicates: [{ column: 'State', operator: 'EQUALS', value: 'ny' }] },
    headers
  );
  assert(FilterEngine.evaluateRow(['1', 'NY', '10', '100', '2023-01-01', ''], eqFilter.compiled, 'AND'), 'EQUALS case-insensitive match (NY == ny)');
  assert(!FilterEngine.evaluateRow(['1', 'NJ', '10', '100', '2023-01-01', ''], eqFilter.compiled, 'AND'), 'EQUALS negative match (NJ != ny)');

  // Leading zeros preservation
  const leadingZeroFilter = FilterEngine.compileFilter(
    { conjunction: 'AND', predicates: [{ column: 'Code', operator: 'EQUALS', value: '05' }] },
    headers
  );
  assert(FilterEngine.evaluateRow(['1', 'NY', '05', '100', '2023-01-01', ''], leadingZeroFilter.compiled, 'AND'), 'Exact string matches with leading zero ("05" == "05")');
  assert(!FilterEngine.evaluateRow(['1', 'NY', '5', '100', '2023-01-01', ''], leadingZeroFilter.compiled, 'AND'), 'Exact string does NOT match stripped zero ("5" != "05")');

  // Numeric comparison
  const numFilter = FilterEngine.compileFilter(
    { conjunction: 'AND', predicates: [{ column: 'Amount', operator: 'GREATER_THAN', value: '50', numeric: true }] },
    headers
  );
  assert(FilterEngine.evaluateRow(['1', 'NY', '05', '100', '2023-01-01', ''], numFilter.compiled, 'AND'), 'Numeric 100 > 50 match');
  assert(!FilterEngine.evaluateRow(['1', 'NY', '05', '25', '2023-01-01', ''], numFilter.compiled, 'AND'), 'Numeric 25 > 50 reject');

  // IS_EMPTY and IS_NOT_EMPTY
  const emptyFilter = FilterEngine.compileFilter(
    { conjunction: 'AND', predicates: [{ column: 'Notes', operator: 'IS_EMPTY' }] },
    headers
  );
  assert(FilterEngine.evaluateRow(['1', 'NY', '05', '100', '2023-01-01', ''], emptyFilter.compiled, 'AND'), 'IS_EMPTY matches empty string ""');
  assert(!FilterEngine.evaluateRow(['1', 'NY', '05', '100', '2023-01-01', 'some note'], emptyFilter.compiled, 'AND'), 'IS_EMPTY rejects non-empty string');

  // IN operator
  const inFilter = FilterEngine.compileFilter(
    { conjunction: 'AND', predicates: [{ column: 'State', operator: 'IN', value: ['NY', 'NJ', 'CT'] }] },
    headers
  );
  assert(FilterEngine.evaluateRow(['1', 'NJ', '05', '100', '2023-01-01', ''], inFilter.compiled, 'AND'), 'IN operator matches present item');
  assert(!FilterEngine.evaluateRow(['1', 'PA', '05', '100', '2023-01-01', ''], inFilter.compiled, 'AND'), 'IN operator rejects missing item');

  // Conjunctions AND vs OR
  const orFilter = FilterEngine.compileFilter(
    {
      conjunction: 'OR',
      predicates: [
        { column: 'State', operator: 'EQUALS', value: 'CA' },
        { column: 'Amount', operator: 'GREATER_THAN', value: '500', numeric: true }
      ]
    },
    headers
  );
  assert(FilterEngine.evaluateRow(['1', 'CA', '05', '10', '2023-01-01', ''], orFilter.compiled, 'OR'), 'OR matches first predicate');
  assert(FilterEngine.evaluateRow(['1', 'NY', '05', '600', '2023-01-01', ''], orFilter.compiled, 'OR'), 'OR matches second predicate');
  assert(!FilterEngine.evaluateRow(['1', 'NY', '05', '100', '2023-01-01', ''], orFilter.compiled, 'OR'), 'OR rejects when neither matches');

  // 2. Real Dataset Partition Scan
  console.log('\n[Section 2] Real Dataset Partition Scan');
  const dataset = await prisma.dataset.findUnique({
    where: { id: DATASET_ID },
    include: { partitions: { where: { partitionNumber: 0 } } }
  });

  if (!dataset || dataset.partitions.length === 0) {
    throw new Error(`Dataset ${DATASET_ID} partition 0 not found`);
  }

  const part0 = dataset.partitions[0];
  const realHeaders = dataset.headerRow as string[];

  // Test scan on first 25,000 rows of partition 0 for Plate_ID starting with 'A'
  const realFilter = FilterEngine.compileFilter(
    {
      conjunction: 'AND',
      predicates: [
        { column: 'Registration_State', operator: 'EQUALS', value: 'NY' }
      ]
    },
    realHeaders
  );

  let matchCount = 0;
  const firstMatches: { id: number; cells: string[] }[] = [];

  const tScan0 = Date.now();
  const scanResult = await FilterEngine.scanPartition(
    part0,
    realFilter.compiled,
    realFilter.conjunction,
    (stableRowId, cells) => {
      matchCount++;
      if (firstMatches.length < 5) {
        firstMatches.push({ id: stableRowId, cells });
      }
    },
    { chunkSize: 25_000 }
  );
  const tScan1 = Date.now();

  assert(scanResult.scannedRows === 2_000_000, `Scanned all 2,000,000 rows in Partition 0 (got ${scanResult.scannedRows})`);
  assert(scanResult.matchedRows > 1_000_000, `Matched NY plates in Partition 0 (got ${scanResult.matchedRows} matches, ~${((scanResult.matchedRows/2_000_000)*100).toFixed(1)}%)`);
  assert(firstMatches.length === 5, 'Retrieved first 5 matches');
  assert(firstMatches.every(m => m.cells[3] === 'NY'), 'Every matched row has Registration_State == "NY"');
  console.log(`  -> Partition 0 (2,000,000 rows, 537 MB) streamed & filtered in ${(tScan1 - tScan0) / 1000}s`);

  console.log('\n' + '='.repeat(80));
  console.log(`   FILTER ENGINE RESULTS: ${passed} PASSED, ${failed} FAILED`);
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
