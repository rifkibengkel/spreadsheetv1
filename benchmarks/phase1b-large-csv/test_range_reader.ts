import fs from 'fs';
import path from 'path';
import { DatasetRangeReader, DatasetRangeError } from '../../src/lib/largeCsv/reader/datasetRangeReader';
import { CsvImportWorker } from '../../src/lib/largeCsv/worker/csvImportWorker';

const DATASET_ID = '80af4daa-3c4a-42eb-b656-dd855dc6f866';
const SOURCE_CSV = 'C:\\Users\\BKKI-1\\Downloads\\archive\\parking_violation.csv';
const EXPECTED_TOTAL_ROWS = 11_809_233;
const EXPECTED_COLUMNS = 54;

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
  console.log('   PHASE 1B — STEP 3: UNIT TESTS FOR DATASET RANGE READER');
  console.log(`   Dataset ID: ${DATASET_ID}`);
  console.log('='.repeat(80));

  // ---------------------------------------------------------------------------
  // TEST 1: First rows (0..100)
  // ---------------------------------------------------------------------------
  console.log('\n[Test 1] First rows of dataset (startRow=0, count=100)');
  const res1 = await DatasetRangeReader.readRange(DATASET_ID, 0, 100);
  assert(res1.rowCount === 100, `Returned exact rowCount=100 (got ${res1.rowCount})`);
  assert(res1.startRow === 0, `startRow is 0 (got ${res1.startRow})`);
  assert(res1.totalRows === EXPECTED_TOTAL_ROWS, `totalRows is ${EXPECTED_TOTAL_ROWS} (got ${res1.totalRows})`);
  assert(res1.columnCount === EXPECTED_COLUMNS, `columnCount is 54 (got ${res1.columnCount})`);
  assert(res1.columns.length === EXPECTED_COLUMNS, `columns.length is 54 (got ${res1.columns.length})`);
  assert(res1.rows[0].stableRowId === 0, `Row 0 stableRowId is 0 (got ${res1.rows[0].stableRowId})`);
  assert(res1.rows[99].stableRowId === 99, `Row 99 stableRowId is 99 (got ${res1.rows[99].stableRowId})`);
  assert(res1.nextRow === 100, `nextRow is 100 (got ${res1.nextRow})`);
  assert(res1.rows.every((r) => r.cells.length === EXPECTED_COLUMNS), 'All 100 rows have exactly 54 columns');

  // ---------------------------------------------------------------------------
  // TEST 2: Normal viewport (startRow=100,000, count=500)
  // ---------------------------------------------------------------------------
  console.log('\n[Test 2] Normal viewport (startRow=100000, count=500)');
  const res2 = await DatasetRangeReader.readRange(DATASET_ID, 100_000, 500);
  assert(res2.rowCount === 500, `Returned exact rowCount=500 (got ${res2.rowCount})`);
  assert(res2.startRow === 100_000, `startRow is 100,000`);
  assert(res2.rows[0].stableRowId === 100_000, `Row 0 stableRowId is 100,000`);
  assert(res2.rows[499].stableRowId === 100_499, `Row 499 stableRowId is 100,499`);
  assert(res2.nextRow === 100_500, `nextRow is 100,500`);
  assert(res2.rows.every((r) => r.cells.length === EXPECTED_COLUMNS), 'All 500 rows have exactly 54 columns');

  // ---------------------------------------------------------------------------
  // TEST 3: Partition Boundary 0 -> 1 (startRow=1,999,500, count=1,000)
  // Spans row 1,999,500..1,999,999 (Part 0) and 2,000,000..2,000,499 (Part 1)
  // ---------------------------------------------------------------------------
  console.log('\n[Test 3] Cross-Partition Boundary 0 -> 1 (startRow=1,999,500, count=1,000)');
  const res3 = await DatasetRangeReader.readRange(DATASET_ID, 1_999_500, 1_000);
  assert(res3.rowCount === 1_000, `Returned exact rowCount=1,000 (got ${res3.rowCount})`);
  assert(res3.rows[0].stableRowId === 1_999_500, `Row 0 stableRowId is 1,999,500`);
  assert(res3.rows[499].stableRowId === 1_999_999, `Boundary row 499 is 1,999,999 (Part 0 end)`);
  assert(res3.rows[500].stableRowId === 2_000_000, `Boundary row 500 is 2,000,000 (Part 1 start)`);
  assert(res3.rows[999].stableRowId === 2_000_499, `Row 999 is 2,000,499`);
  assert(res3.rows.every((r) => r.cells.length === EXPECTED_COLUMNS), 'All 1,000 boundary rows have exactly 54 columns');
  // Validate contiguous stableRowId sequence across boundary
  let contiguous3 = true;
  for (let i = 0; i < res3.rows.length; i++) {
    if (res3.rows[i].stableRowId !== 1_999_500 + i) contiguous3 = false;
  }
  assert(contiguous3, 'All 1,000 rows have strictly contiguous, monotonically increasing stableRowId');

  // ---------------------------------------------------------------------------
  // TEST 4: Partition Boundary 1 -> 2 (startRow=3,999,500, count=1,000)
  // ---------------------------------------------------------------------------
  console.log('\n[Test 4] Cross-Partition Boundary 1 -> 2 (startRow=3,999,500, count=1,000)');
  const res4 = await DatasetRangeReader.readRange(DATASET_ID, 3_999_500, 1_000);
  assert(res4.rowCount === 1_000, `Returned rowCount=1,000`);
  assert(res4.rows[499].stableRowId === 3_999_999, `Boundary row 499 is 3,999,999 (Part 1 end)`);
  assert(res4.rows[500].stableRowId === 4_000_000, `Boundary row 500 is 4,000,000 (Part 2 start)`);
  assert(res4.rows.every((r) => r.cells.length === EXPECTED_COLUMNS), 'All rows have 54 columns');

  // ---------------------------------------------------------------------------
  // TEST 5: Partition Boundary 2 -> 3 (startRow=5,999,500, count=1,000)
  // ---------------------------------------------------------------------------
  console.log('\n[Test 5] Cross-Partition Boundary 2 -> 3 (startRow=5,999,500, count=1,000)');
  const res5 = await DatasetRangeReader.readRange(DATASET_ID, 5_999_500, 1_000);
  assert(res5.rowCount === 1_000, `Returned rowCount=1,000`);
  assert(res5.rows[499].stableRowId === 5_999_999, `Boundary row 499 is 5,999,999 (Part 2 end)`);
  assert(res5.rows[500].stableRowId === 6_000_000, `Boundary row 500 is 6,000,000 (Part 3 start)`);
  assert(res5.rows.every((r) => r.cells.length === EXPECTED_COLUMNS), 'All rows have 54 columns');

  // ---------------------------------------------------------------------------
  // TEST 6: Partition Boundary 3 -> 4 (startRow=7,999,500, count=1,000)
  // ---------------------------------------------------------------------------
  console.log('\n[Test 6] Cross-Partition Boundary 3 -> 4 (startRow=7,999,500, count=1,000)');
  const res6 = await DatasetRangeReader.readRange(DATASET_ID, 7_999_500, 1_000);
  assert(res6.rowCount === 1_000, `Returned rowCount=1,000`);
  assert(res6.rows[499].stableRowId === 7_999_999, `Boundary row 499 is 7,999,999 (Part 3 end)`);
  assert(res6.rows[500].stableRowId === 8_000_000, `Boundary row 500 is 8,000,000 (Part 4 start)`);
  assert(res6.rows.every((r) => r.cells.length === EXPECTED_COLUMNS), 'All rows have 54 columns');

  // ---------------------------------------------------------------------------
  // TEST 7: Partition Boundary 4 -> 5 (startRow=9,999,500, count=1,000)
  // ---------------------------------------------------------------------------
  console.log('\n[Test 7] Cross-Partition Boundary 4 -> 5 (startRow=9,999,500, count=1,000)');
  const res7 = await DatasetRangeReader.readRange(DATASET_ID, 9_999_500, 1_000);
  assert(res7.rowCount === 1_000, `Returned rowCount=1,000`);
  assert(res7.rows[499].stableRowId === 9_999_999, `Boundary row 499 is 9,999,999 (Part 4 end)`);
  assert(res7.rows[500].stableRowId === 10_000_000, `Boundary row 500 is 10,000,000 (Part 5 start)`);
  assert(res7.rows.every((r) => r.cells.length === EXPECTED_COLUMNS), 'All rows have 54 columns');

  // ---------------------------------------------------------------------------
  // TEST 8: Final 100 rows (startRow=11,809,133, count=100)
  // ---------------------------------------------------------------------------
  console.log('\n[Test 8] Final 100 rows of dataset (startRow=11,809,133, count=100)');
  const res8 = await DatasetRangeReader.readRange(DATASET_ID, 11_809_133, 100);
  assert(res8.rowCount === 100, `Returned exact rowCount=100`);
  assert(res8.rows[0].stableRowId === 11_809_133, `First row is 11,809,133`);
  assert(res8.rows[99].stableRowId === 11_809_232, `Last row is exactly 11,809,232 (final row)`);
  assert(res8.nextRow === null, `nextRow is null (end of dataset reached)`);
  assert(res8.rows.every((r) => r.cells.length === EXPECTED_COLUMNS), 'All final 100 rows have 54 columns');

  // ---------------------------------------------------------------------------
  // TEST 9: Final single row (startRow=11,809,232, count=1)
  // ---------------------------------------------------------------------------
  console.log('\n[Test 9] Final single row (startRow=11,809,232, count=1)');
  const res9 = await DatasetRangeReader.readRange(DATASET_ID, 11_809_232, 1);
  assert(res9.rowCount === 1, `Returned exact rowCount=1`);
  assert(res9.rows[0].stableRowId === 11_809_232, `stableRowId is 11,809,232`);
  assert(res9.nextRow === null, `nextRow is null`);
  assert(res9.rows[0].cells.length === EXPECTED_COLUMNS, `Row has 54 columns`);

  // ---------------------------------------------------------------------------
  // TEST 10: Clamped row count near end of dataset
  // Requesting 50 rows starting at 11,809,200 (only 33 rows remain: 200..232)
  // ---------------------------------------------------------------------------
  console.log('\n[Test 10] Bounded clamp at dataset end');
  const res10 = await DatasetRangeReader.readRange(DATASET_ID, 11_809_200, 50);
  assert(res10.rowCount === 33, `Clamped rowCount is 33 (got ${res10.rowCount})`);
  assert(res10.rows[32].stableRowId === 11_809_232, `Final returned row is 11,809,232`);
  assert(res10.nextRow === null, `nextRow is null`);

  // ---------------------------------------------------------------------------
  // TEST 11: Textual Data Fidelity against Source CSV
  // ---------------------------------------------------------------------------
  console.log('\n[Test 11] Data Fidelity Verification against Source CSV');
  if (fs.existsSync(SOURCE_CSV)) {
    // Check row 0 against source CSV row 1 (since line 0 of CSV is header)
    const sourceFd = fs.openSync(SOURCE_CSV, 'r');
    const headerAndFirstRowBuf = Buffer.alloc(16384);
    fs.readSync(sourceFd, headerAndFirstRowBuf, 0, 16384, 0);
    fs.closeSync(sourceFd);

    const rawLines = headerAndFirstRowBuf.toString('utf8').split('\n');
    const sourceHeaderLine = rawLines[0].replace(/\r$/, '');
    const sourceFirstDataLine = rawLines[1].replace(/\r$/, '');

    const expectedHeaderCells = CsvImportWorker.parseCsvLine(sourceHeaderLine);
    const expectedFirstRowCells = CsvImportWorker.parseCsvLine(sourceFirstDataLine);

    assert(
      JSON.stringify(res1.columns) === JSON.stringify(expectedHeaderCells),
      'Columns match source CSV header exactly'
    );
    assert(
      JSON.stringify(res1.rows[0].cells) === JSON.stringify(expectedFirstRowCells),
      'Row 0 cells match source CSV data row 0 exactly'
    );

    // Verify critical values on Row 0
    assert(res1.rows[0].cells[0] === expectedFirstRowCells[0], `Row 0 cell[0] fidelity (${res1.rows[0].cells[0]})`);
    assert(res1.rows[0].cells[1] === expectedFirstRowCells[1], `Row 0 cell[1] summons (${res1.rows[0].cells[1]})`);
    assert(res1.rows[0].cells[2] === expectedFirstRowCells[2], `Row 0 cell[2] plate (${res1.rows[0].cells[2]})`);
  } else {
    console.warn(`  [SKIP] Source CSV not accessible at ${SOURCE_CSV}`);
  }

  // ---------------------------------------------------------------------------
  // TEST 12: Bounded Timings Decomposition (T1..T5)
  // ---------------------------------------------------------------------------
  console.log('\n[Test 12] Timing Observability Decomposition');
  console.log(`  Partition Resolution: ${res3.metrics.partitionResolutionMs.toFixed(3)} ms`);
  console.log(`  Index Lookup:         ${res3.metrics.indexLookupMs.toFixed(3)} ms`);
  console.log(`  File Read:            ${res3.metrics.fileReadMs.toFixed(3)} ms`);
  console.log(`  Decode (1,000 rows):  ${res3.metrics.decodeMs.toFixed(3)} ms`);
  console.log(`  Server Total:         ${res3.metrics.serverTotalMs.toFixed(3)} ms`);
  assert(res3.metrics.serverTotalMs < 50, `1,000 cross-partition rows read in < 50ms (got ${res3.metrics.serverTotalMs.toFixed(2)}ms)`);

  // ---------------------------------------------------------------------------
  // TEST 13: Failure & Edge Case Handling
  // ---------------------------------------------------------------------------
  console.log('\n[Test 13] Deterministic Error Handling');

  // 13.1 startRow < 0
  try {
    await DatasetRangeReader.readRange(DATASET_ID, -1, 100);
    assert(false, 'Expected negative startRow to fail');
  } catch (err: any) {
    assert(err instanceof DatasetRangeError && err.code === 'INVALID_START_ROW', 'Negative startRow rejected with INVALID_START_ROW');
  }

  // 13.2 rowCount <= 0
  try {
    await DatasetRangeReader.readRange(DATASET_ID, 0, 0);
    assert(false, 'Expected rowCount=0 to fail');
  } catch (err: any) {
    assert(err instanceof DatasetRangeError && err.code === 'INVALID_ROW_COUNT', 'Zero rowCount rejected with INVALID_ROW_COUNT');
  }

  // 13.3 rowCount > 5,000
  try {
    await DatasetRangeReader.readRange(DATASET_ID, 0, 5001);
    assert(false, 'Expected rowCount > 5000 to fail');
  } catch (err: any) {
    assert(err instanceof DatasetRangeError && err.code === 'ROW_COUNT_EXCEEDS_LIMIT', 'Excessive rowCount rejected with ROW_COUNT_EXCEEDS_LIMIT');
  }

  // 13.4 startRow >= totalRows
  try {
    await DatasetRangeReader.readRange(DATASET_ID, EXPECTED_TOTAL_ROWS, 100);
    assert(false, 'Expected startRow >= totalRows to fail');
  } catch (err: any) {
    assert(err instanceof DatasetRangeError && err.code === 'START_ROW_OUT_OF_BOUNDS', 'Out-of-bounds startRow rejected with START_ROW_OUT_OF_BOUNDS');
  }

  // 13.5 Non-existent dataset
  try {
    await DatasetRangeReader.readRange('00000000-0000-0000-0000-000000000000', 0, 100);
    assert(false, 'Expected missing dataset to fail');
  } catch (err: any) {
    assert(err instanceof DatasetRangeError && err.code === 'DATASET_NOT_FOUND', 'Non-existent dataset rejected with DATASET_NOT_FOUND');
  }

  console.log('\n' + '='.repeat(80));
  console.log(`STEP 3 TEST RESULTS: ${passed} PASSED, ${failed} FAILED`);
  console.log('='.repeat(80));

  if (failed > 0) {
    process.exit(1);
  }
}

runTests().catch((err) => {
  console.error('Fatal test error:', err);
  process.exit(1);
}).finally(() => process.exit(0));
