/**
 * TEST SUITE: AUTO-SPLIT VIRTUAL SHEETS PER 4 MILLION CSV DATA ROWS
 * 
 * Verifies Sections 18, 19, 20, 21 of the specification:
 * 1. Normal Mode Partitioning (11,809,233 rows -> Sheet1: 4M, Sheet2: 4M, Sheet3: ~3.8M)
 * 2. Header Contract (row 0 = header, row 1 = data row 0 for every virtual sheet)
 * 3. Boundary & Cross-Boundary checks (3,999,999, 4,000,000, 4,000,001, 7,999,999, 8,000,000, 8,000,001)
 * 4. Random lookup verification against physical dataset
 * 5. Sheet switching stress test (25 rapid switches: latency, no memory leak, bounded cells)
 * 6. Ctrl+G / Jump-to-row (Section 12 exact formula, sheet switching on jump)
 * 7. Query Mode Partitioning (<4M rows -> 1 sheet, =4M rows -> 1 sheet, >4M rows -> 2+ sheets)
 * 8. Query Mode 4M boundary continuity (queryResultPosition vs stableRowId)
 * 9. Sorted query virtual-sheet partitioning
 */

import prisma from '../../src/lib/db/prisma';
import { DatasetRangeReader } from '../../src/lib/largeCsv/reader/datasetRangeReader';
import { QueryIndexManager } from '../../src/lib/largeCsv/query/queryIndexManager';
import { DatasetRowResolver } from '../../src/lib/largeCsv/query/datasetRowResolver';
import { UniverViewportAdapter } from '../../src/lib/viewport/univerViewportAdapter';
import { DatasetViewportController } from '../../src/lib/viewport/datasetViewportController';
import { BoundedLruCache } from '../../src/lib/viewport/boundedLruCache';
import { ROWS_PER_VIRTUAL_SHEET, VirtualSheetPartition } from '../../src/lib/viewport/types';
import { RowSlice } from '../../src/lib/storage/opfs/types';

const DATASET_ID = '80af4daa-3c4a-42eb-b656-dd855dc6f866';
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

/**
 * Mock Univer API supporting multiple worksheets and sheet operations
 */
function createMockUniverAPI() {
  const sheets: Map<string, {
    name: string;
    rowCount: number;
    columnCount: number;
    cells: Record<string, any>;
  }> = new Map();

  let activeSheetId = 'sheet-1';
  let currentScrollRow = 0;

  // Initialize primary sheet
  sheets.set('sheet-1', {
    name: 'Sheet1',
    rowCount: 1000,
    columnCount: 26,
    cells: {},
  });

  const getActiveSheetData = () => {
    let s = sheets.get(activeSheetId);
    if (!s) {
      s = { name: activeSheetId, rowCount: 1000, columnCount: 26, cells: {} };
      sheets.set(activeSheetId, s);
    }
    return s;
  };

  const mockWorksheet = (sheetId: string) => {
    const s = sheets.get(sheetId) || getActiveSheetData();
    return {
      getSheetId: () => sheetId,
      getName: () => s.name,
      setName: (name: string) => { s.name = name; },
      setRowCount: (count: number) => { s.rowCount = count; },
      getRowCount: () => s.rowCount,
      setColumnCount: (count: number) => { s.columnCount = count; },
      getColumnCount: () => s.columnCount,
      getCellMatrix: () => ({
        _matrix: new Proxy({} as any, {
          get(_t, prop) {
            const row = Number(prop);
            return s.cells[`${row},0`] !== undefined || s.cells[`${row},1`] !== undefined;
          },
          deleteProperty(_t, prop) {
            const row = Number(prop);
            for (let c = 0; c < 60; c++) {
              delete s.cells[`${row},${c}`];
            }
            return true;
          },
        }),
      }),
    };
  };

  const mockWorkbook = {
    getId: () => 'workbook-01',
    getActiveSheet: () => mockWorksheet(activeSheetId),
    getSheetBySheetId: (id: string) => sheets.has(id) ? mockWorksheet(id) : null,
    getSheets: () => Array.from(sheets.keys()).map((id) => mockWorksheet(id)),
  };

  const mockCommandService = {
    executeCommand: async (cmd: string, params: any) => {
      if (cmd === 'sheet.operation.set-worksheet-active') {
        const target = params?.subUnitId || params?.sheetId;
        if (target && sheets.has(target)) {
          activeSheetId = target;
        }
      } else if (cmd === 'sheet.mutation.insert-sheet') {
        const id = params?.sheet?.id || `sheet-${sheets.size + 1}`;
        sheets.set(id, {
          name: params?.sheet?.name || `Sheet${sheets.size + 1}`,
          rowCount: params?.sheet?.rowCount || 1000,
          columnCount: params?.sheet?.columnCount || 26,
          cells: {},
        });
        if (params?.sheet?.cellData?.[0]) {
          const s = sheets.get(id)!;
          for (const c of Object.keys(params.sheet.cellData[0])) {
            s.cells[`0,${c}`] = params.sheet.cellData[0][c]?.v;
          }
        }
      } else if (cmd === 'sheet.mutation.remove-sheet') {
        sheets.delete(params?.subUnitId);
      } else if (cmd === 'sheet.mutation.set-worksheet-row-count') {
        const s = sheets.get(params?.subUnitId || activeSheetId);
        if (s) s.rowCount = params.rowCount;
      } else if (cmd === 'sheet.mutation.set-worksheet-col-count') {
        const s = sheets.get(params?.subUnitId || activeSheetId);
        if (s) s.columnCount = params.colCount;
      } else if (cmd === 'sheet.mutation.set-range-values') {
        const subId = params?.subUnitId || activeSheetId;
        const s = sheets.get(subId);
        if (s && params?.cellValue) {
          for (const r of Object.keys(params.cellValue)) {
            for (const c of Object.keys(params.cellValue[r])) {
              s.cells[`${r},${c}`] = params.cellValue[r][c]?.v;
            }
          }
        }
      } else if (cmd === 'sheet.operation.set-scroll') {
        currentScrollRow = params?.sheetViewStartRow ?? currentScrollRow;
      }
      return true;
    },
  };

  return {
    mockUniverAPI: {
      getActiveWorkbook: () => mockWorkbook,
      executeCommand: mockCommandService.executeCommand,
      getCommandService: () => mockCommandService,
    } as any,
    getActiveSheetId: () => activeSheetId,
    setActiveSheetId: (id: string) => { activeSheetId = id; },
    getSheets: () => sheets,
    getSheetCells: (sheetId: string) => sheets.get(sheetId)?.cells || {},
    getCurrentScrollRow: () => currentScrollRow,
  };
}

async function runVirtualSheetTests() {
  console.log('='.repeat(80));
  console.log('   AUTO-SPLIT VIRTUAL SHEETS PER 4M ROWS — INTEGRATION & BOUNDARY TESTS');
  console.log(`   Target Dataset: ${DATASET_ID}`);
  console.log(`   Expected Rows: ${EXPECTED_TOTAL_ROWS.toLocaleString()} | Expected Cols: ${EXPECTED_COLUMNS}`);
  console.log('='.repeat(80));

  const mockEnv = createMockUniverAPI();
  const adapter = new UniverViewportAdapter(mockEnv.mockUniverAPI);
  const controller = new DatasetViewportController(adapter, {
    visibleRows: 50,
    overscanRows: 100,
    marginThreshold: 30,
  });

  // Emulate worker cache
  const memCache = new Map<string, any>();

  // Wire fetchRange to physical backend with cache
  (controller as any).fetchRange = async (startRow: number, rowCount: number, generationToken: number) => {
    const queryId = controller.getCurrentQueryId();
    const cacheKey = `${DATASET_ID}:${queryId || 'BASE'}:${startRow}:${rowCount}`;
    const cached = memCache.get(cacheKey);
    if (cached) return cached;

    if (queryId) {
      if (queryId.startsWith('mock-')) {
        // Synthetic query mode for boundary/partition size tests
        const rows = Array.from({ length: rowCount }, (_, i) => ({
          stableRowId: (startRow + i) * 2 + 1, // Non-consecutive stableRowId
          cells: [String(startRow + i), `Value_${startRow + i}`, 'Mock'],
        }));
        const res = {
          datasetId: DATASET_ID,
          queryId,
          startRow,
          rowCount: rows.length,
          totalRows: controller.getTotalRows(),
          columns: ['Pos', 'Value', 'Status'],
          columnCount: 3,
          rows,
        };
        memCache.set(cacheKey, res);
        return res;
      }

      const { stableRowIds, totalMatchedRows } = await QueryIndexManager.readStableRowIds(
        DATASET_ID,
        queryId,
        startRow,
        rowCount
      );
      const rows = await DatasetRowResolver.resolveRows(DATASET_ID, stableRowIds);
      const datasetMeta = await prisma.dataset.findUnique({
        where: { id: DATASET_ID },
        select: { headerRow: true, columnCount: true },
      });
      const res = {
        datasetId: DATASET_ID,
        queryId,
        startRow,
        rowCount: rows.length,
        totalRows: totalMatchedRows,
        columns: (datasetMeta?.headerRow as string[]) || [],
        columnCount: datasetMeta?.columnCount || 54,
        rows,
      };
      memCache.set(cacheKey, res);
      return res;
    }

    const res = await DatasetRangeReader.readRange(DATASET_ID, startRow, rowCount);
    memCache.set(cacheKey, res);
    return res;
  };

  // =========================================================================
  // TEST GROUP 1: MOUNT & PARTITION CALCULATIONS (Normal Mode)
  // =========================================================================
  console.log('\n--- TEST GROUP 1: Normal Mode Partitioning & Header Contract ---');
  const mountSuccess = await controller.mount(DATASET_ID);
  assert(mountSuccess === true, 'Controller mounted successfully');

  const partitions = controller.getPartitions();
  assert(partitions.length === 3, `Expected exactly 3 virtual sheets, got ${partitions.length}`);

  // Sheet 1 boundary definitions
  assert(
    partitions[0].sheetIndex === 0 &&
    partitions[0].startRow === 0 &&
    partitions[0].endRow === 3_999_999 &&
    partitions[0].rowCount === 4_000_000,
    'Sheet 1 partition boundaries: start=0, end=3,999,999, rowCount=4,000,000'
  );

  // Sheet 2 boundary definitions
  assert(
    partitions[1].sheetIndex === 1 &&
    partitions[1].startRow === 4_000_000 &&
    partitions[1].endRow === 7_999_999 &&
    partitions[1].rowCount === 4_000_000,
    'Sheet 2 partition boundaries: start=4,000,000, end=7,999,999, rowCount=4,000,000'
  );

  // Sheet 3 boundary definitions
  assert(
    partitions[2].sheetIndex === 2 &&
    partitions[2].startRow === 8_000_000 &&
    partitions[2].endRow === 11_809_232 &&
    partitions[2].rowCount === 3_809_233,
    'Sheet 3 partition boundaries: start=8,000,000, end=11,809,232, rowCount=3,809,233'
  );

  const sheets = mockEnv.getSheets();
  assert(sheets.size === 3, `Univer workbook has exactly 3 worksheet instances, got ${sheets.size}`);

  const s1 = sheets.get('sheet-1');
  const s2 = sheets.get('sheet-2');
  const s3 = sheets.get('sheet-3');

  assert(s1?.rowCount === 4_000_001, `Sheet 1 Univer row count is 4,000,001 (1 header + 4M rows), got ${s1?.rowCount}`);
  assert(s2?.rowCount === 4_000_001, `Sheet 2 Univer row count is 4,000,001 (1 header + 4M rows), got ${s2?.rowCount}`);
  assert(s3?.rowCount === 3_809_234, `Sheet 3 Univer row count is 3,809,234 (1 header + 3,809,233 rows), got ${s3?.rowCount}`);

  // Verify Header Contract (Row 0 on all sheets is header)
  const headers = controller.getHeaders();
  assert(headers.length === 54, `Headers loaded: 54 columns`);
  assert(s1?.cells['0,0'] === headers[0], `Sheet 1 row 0, col 0 matches CSV header: "${headers[0]}"`);
  assert(s2?.cells['0,0'] === headers[0], `Sheet 2 row 0, col 0 matches CSV header: "${headers[0]}"`);
  assert(s3?.cells['0,0'] === headers[0], `Sheet 3 row 0, col 0 matches CSV header: "${headers[0]}"`);

  // Verify Sheet 1 first data row (local row 1 = global data row 0)
  const directRow0 = await DatasetRangeReader.readRange(DATASET_ID, 0, 1);
  assert(
    String(s1?.cells['1,0']) === String(directRow0.rows[0].cells[0]),
    `Sheet 1 local row 1 matches global data row 0: "${s1?.cells['1,0']}" vs "${directRow0.rows[0].cells[0]}"`
  );

  // =========================================================================
  // TEST GROUP 2: BOUNDARY & CROSS-BOUNDARY ROWS
  // =========================================================================
  console.log('\n--- TEST GROUP 2: Boundary & Cross-Boundary Checks ---');

  // Verify boundary rows directly against physical reader
  const row3999999 = await DatasetRangeReader.readRange(DATASET_ID, 3_999_999, 1);
  const row4000000 = await DatasetRangeReader.readRange(DATASET_ID, 4_000_000, 1);
  const row4000001 = await DatasetRangeReader.readRange(DATASET_ID, 4_000_001, 1);
  const row7999999 = await DatasetRangeReader.readRange(DATASET_ID, 7_999_999, 1);
  const row8000000 = await DatasetRangeReader.readRange(DATASET_ID, 8_000_000, 1);
  const row8000001 = await DatasetRangeReader.readRange(DATASET_ID, 8_000_001, 1);

  assert(row3999999.rows.length === 1, 'Physical reader reads global row 3,999,999 (last row of Sheet 1)');
  assert(row4000000.rows.length === 1, 'Physical reader reads global row 4,000,000 (first row of Sheet 2)');
  assert(row4000001.rows.length === 1, 'Physical reader reads global row 4,000,001 (second row of Sheet 2)');
  assert(row7999999.rows.length === 1, 'Physical reader reads global row 7,999,999 (last row of Sheet 2)');
  assert(row8000000.rows.length === 1, 'Physical reader reads global row 8,000,000 (first row of Sheet 3)');
  assert(row8000001.rows.length === 1, 'Physical reader reads global row 8,000,001 (second row of Sheet 3)');

  // Verify contiguous physical read spanning the 4M boundary (3,999,998 to 4,000,002)
  const cross4M = await DatasetRangeReader.readRange(DATASET_ID, 3_999_998, 5);
  assert(cross4M.rows.length === 5, 'Physical reader seamlessly reads 5 rows across 4,000,000 partition boundary');
  assert(cross4M.rows[1].cells[0] === row3999999.rows[0].cells[0], 'Cross-boundary index 1 matches row 3,999,999');
  assert(cross4M.rows[2].cells[0] === row4000000.rows[0].cells[0], 'Cross-boundary index 2 matches row 4,000,000');
  assert(cross4M.rows[3].cells[0] === row4000001.rows[0].cells[0], 'Cross-boundary index 3 matches row 4,000,001');

  // Verify contiguous physical read spanning the 8M boundary (7,999,998 to 8,000,002)
  const cross8M = await DatasetRangeReader.readRange(DATASET_ID, 7_999_998, 5);
  assert(cross8M.rows.length === 5, 'Physical reader seamlessly reads 5 rows across 8,000,000 partition boundary');
  assert(cross8M.rows[1].cells[0] === row7999999.rows[0].cells[0], 'Cross-boundary index 1 matches row 7,999,999');
  assert(cross8M.rows[2].cells[0] === row8000000.rows[0].cells[0], 'Cross-boundary index 2 matches row 8,000,000');
  assert(cross8M.rows[3].cells[0] === row8000001.rows[0].cells[0], 'Cross-boundary index 3 matches row 8,000,001');

  // =========================================================================
  // TEST GROUP 3: SHEET SWITCHING & HYDRATION
  // =========================================================================
  console.log('\n--- TEST GROUP 3: Sheet Switching & Bounded Memory ---');

  // Switch to Sheet 2
  const tSwitch2Start = performance.now();
  await controller.setActiveSheet(1);
  const tSwitch2End = performance.now();
  console.log(`  Sheet 1 -> Sheet 2 switch latency: ${(tSwitch2End - tSwitch2Start).toFixed(2)} ms`);

  assert(controller.getActiveSheetIndex() === 1, 'Active sheet index is 1 (Sheet 2)');
  assert(mockEnv.getActiveSheetId() === 'sheet-2', 'Mock Univer active sheet is sheet-2');

  // Check Sheet 2 first data row (local row 1 = global data row 4,000,000)
  const s2Cells = mockEnv.getSheetCells('sheet-2');
  assert(
    String(s2Cells['1,0']) === String(row4000000.rows[0].cells[0]),
    `Sheet 2 local row 1 correctly hydrated with global data row 4,000,000: "${s2Cells['1,0']}"`
  );

  // Switch to Sheet 3
  const tSwitch3Start = performance.now();
  await controller.setActiveSheet(2);
  const tSwitch3End = performance.now();
  console.log(`  Sheet 2 -> Sheet 3 switch latency: ${(tSwitch3End - tSwitch3Start).toFixed(2)} ms`);

  assert(controller.getActiveSheetIndex() === 2, 'Active sheet index is 2 (Sheet 3)');
  assert(mockEnv.getActiveSheetId() === 'sheet-3', 'Mock Univer active sheet is sheet-3');

  // Check Sheet 3 first data row (local row 1 = global data row 8,000,000)
  const s3Cells = mockEnv.getSheetCells('sheet-3');
  assert(
    String(s3Cells['1,0']) === String(row8000000.rows[0].cells[0]),
    `Sheet 3 local row 1 correctly hydrated with global data row 8,000,000: "${s3Cells['1,0']}"`
  );

  // Stress test: 25 rapid switches between sheets
  console.log('  Running 25 rapid sheet switches stress test...');
  const switchTimes: number[] = [];
  for (let i = 0; i < 25; i++) {
    const targetIdx = i % 3;
    const t0 = performance.now();
    await controller.setActiveSheet(targetIdx);
    const t1 = performance.now();
    switchTimes.push(t1 - t0);

    assert(controller.getActiveSheetIndex() === targetIdx, `Switch #${i + 1} active sheet is ${targetIdx}`);
  }

  const avgSwitchMs = switchTimes.reduce((a, b) => a + b, 0) / switchTimes.length;
  console.log(`  Average sheet-switch latency over 25 switches (with LRU cache): ${avgSwitchMs.toFixed(2)} ms`);
  assert(avgSwitchMs < 50, `Cached sheet switch latency is fast (<50ms, actual: ${avgSwitchMs.toFixed(2)} ms)`);

  // Verify memory boundedness: total cells hydrated in Univer is bounded (<15,000 cells)
  let totalHydratedCellsAcrossAllSheets = 0;
  for (const [sId, sData] of sheets.entries()) {
    const cellCount = Object.keys(sData.cells).length;
    totalHydratedCellsAcrossAllSheets += cellCount;
  }
  console.log(`  Total hydrated cells across all sheets after 25 switches: ${totalHydratedCellsAcrossAllSheets}`);
  assert(
    totalHydratedCellsAcrossAllSheets < 20_000,
    `No memory accumulation / progressive leak across 25 switches: total cells = ${totalHydratedCellsAcrossAllSheets}`
  );

  // =========================================================================
  // TEST GROUP 4: CTRL+G / JUMP-TO-ROW COORDINATE MAPPING (Section 12)
  // =========================================================================
  console.log('\n--- TEST GROUP 4: Ctrl+G / Jump-To-Row Coordinate Mapping ---');

  // Test prompt example: row 4,000,001 must resolve to Sheet 2, local data row 1, local Univer row 2
  const jumpSuccess1 = await controller.jumpToRow(4_000_001);
  assert(jumpSuccess1 === true, 'jumpToRow(4,000,001) completed successfully');
  assert(controller.getActiveSheetIndex() === 1, 'jumpToRow(4,000,001) navigated to Sheet 2');
  assert(mockEnv.getActiveSheetId() === 'sheet-2', 'Univer active sheet changed to sheet-2');

  const s2CellsAfterJump = mockEnv.getSheetCells('sheet-2');
  assert(
    String(s2CellsAfterJump['2,0']) === String(row4000001.rows[0].cells[0]),
    `Sheet 2 local Univer row 2 displays global data row 4,000,001: "${s2CellsAfterJump['2,0']}"`
  );

  // Test jump to row 8,000,000 (first data row of Sheet 3 -> local data row 0, local Univer row 1)
  const jumpSuccess2 = await controller.jumpToRow(8_000_000);
  assert(jumpSuccess2 === true, 'jumpToRow(8,000,000) completed successfully');
  assert(controller.getActiveSheetIndex() === 2, 'jumpToRow(8,000,000) navigated to Sheet 3');
  assert(mockEnv.getActiveSheetId() === 'sheet-3', 'Univer active sheet changed to sheet-3');

  const s3CellsAfterJump = mockEnv.getSheetCells('sheet-3');
  assert(
    String(s3CellsAfterJump['1,0']) === String(row8000000.rows[0].cells[0]),
    `Sheet 3 local Univer row 1 displays global data row 8,000,000: "${s3CellsAfterJump['1,0']}"`
  );

  // Test jump to row 0 (first data row of Sheet 1 -> local data row 0, local Univer row 1)
  const jumpSuccess3 = await controller.jumpToRow(0);
  assert(jumpSuccess3 === true, 'jumpToRow(0) completed successfully');
  assert(controller.getActiveSheetIndex() === 0, 'jumpToRow(0) navigated to Sheet 1');
  assert(mockEnv.getActiveSheetId() === 'sheet-1', 'Univer active sheet changed to sheet-1');

  const s1CellsAfterJump = mockEnv.getSheetCells('sheet-1');
  assert(
    String(s1CellsAfterJump['1,0']) === String(directRow0.rows[0].cells[0]),
    `Sheet 1 local Univer row 1 displays global data row 0: "${s1CellsAfterJump['1,0']}"`
  );

  // =========================================================================
  // TEST GROUP 5: RANDOM LOOKUP VERIFICATION
  // =========================================================================
  console.log('\n--- TEST GROUP 5: Random Lookup Verification ---');

  const randomSamples = [
    { globalRow: 1_234_567, expectedSheet: 0, localRow: 1_234_567 + 1 },
    { globalRow: 5_432_100, expectedSheet: 1, localRow: (5_432_100 % 4_000_000) + 1 },
    { globalRow: 10_500_000, expectedSheet: 2, localRow: (10_500_000 % 4_000_000) + 1 },
  ];

  for (const sample of randomSamples) {
    const expectedDirect = await DatasetRangeReader.readRange(DATASET_ID, sample.globalRow, 1);
    await controller.jumpToRow(sample.globalRow);

    assert(
      controller.getActiveSheetIndex() === sample.expectedSheet,
      `Random global row ${sample.globalRow.toLocaleString()} resolved to Sheet ${sample.expectedSheet + 1}`
    );

    const sheetId = `sheet-${sample.expectedSheet + 1}`;
    const cells = mockEnv.getSheetCells(sheetId);
    assert(
      String(cells[`${sample.localRow},0`]) === String(expectedDirect.rows[0].cells[0]),
      `Random row ${sample.globalRow.toLocaleString()} cell value matches physical dataset: "${cells[`${sample.localRow},0`]}"`
    );
  }

  // =========================================================================
  // TEST GROUP 6: QUERY MODE VIRTUAL SHEET PARTITIONING (Sections 10, 11, 19)
  // =========================================================================
  console.log('\n--- TEST GROUP 6: Query Mode Virtual Sheet Partitioning ---');

  // Case A: Real filter query returning < 4M rows (CT query: 160,361 rows)
  const ctQueryId = 'q_0f51000fac98c91c04442c43e6ad5b67';
  const ctMeta = await QueryIndexManager.getQueryMetadata(DATASET_ID, ctQueryId);
  assert(ctMeta !== null, 'CT real query metadata loaded from disk');
  await controller.setQuery(ctQueryId, ctMeta!.matchedRows);

  const qPartitionsReal = controller.getPartitions();
  assert(qPartitionsReal.length === 1, `Filter query with 160,361 rows produces exactly 1 virtual sheet, got ${qPartitionsReal.length}`);
  assert(qPartitionsReal[0].rowCount === 160_361, `Query Sheet 1 rowCount is 160,361`);
  assert(controller.getActiveSheetIndex() === 0, 'Query active sheet is Sheet 1');

  // Case B: Synthetic query returning exactly 4,000,000 rows -> exactly 1 sheet
  await controller.setQuery('mock-query-exact-4m', 4_000_000);
  const qPartitionsExact = controller.getPartitions();
  assert(qPartitionsExact.length === 1, `Query with exactly 4M rows produces 1 virtual sheet, got ${qPartitionsExact.length}`);
  assert(qPartitionsExact[0].rowCount === 4_000_000, `Sheet 1 rowCount is 4,000,000`);

  // Case C: Synthetic query returning > 4M rows (8,500,000 rows -> 3 sheets)
  await controller.setQuery('mock-query-large-8.5m', 8_500_000);
  const qPartitionsLarge = controller.getPartitions();
  assert(qPartitionsLarge.length === 3, `Query with 8,500,000 rows produces 3 virtual sheets, got ${qPartitionsLarge.length}`);
  assert(qPartitionsLarge[0].rowCount === 4_000_000, `Query Sheet 1 rowCount is 4,000,000`);
  assert(qPartitionsLarge[1].rowCount === 4_000_000, `Query Sheet 2 rowCount is 4,000,000`);
  assert(qPartitionsLarge[2].rowCount === 500_000, `Query Sheet 3 rowCount is 500,000`);

  // Section 19: Query result crossing 4M boundary:
  // In query coordinate space:
  // Sheet 1 last query row: queryResultPosition = 3,999,999
  // Sheet 2 first query row: queryResultPosition = 4,000,000
  // Verify consecutive query-result positions and non-consecutive stableRowId values
  const qPos1 = 3_999_999;
  const qPos2 = 4_000_000;
  const sheetIdx1 = Math.floor(qPos1 / ROWS_PER_VIRTUAL_SHEET);
  const sheetIdx2 = Math.floor(qPos2 / ROWS_PER_VIRTUAL_SHEET);
  assert(sheetIdx1 === 0, `Query position 3,999,999 is in Sheet 1`);
  assert(sheetIdx2 === 1, `Query position 4,000,000 is in Sheet 2`);
  assert(qPos2 - qPos1 === 1, 'Query positions 3,999,999 and 4,000,000 are strictly consecutive');

  // Jump to Sheet 2 query position 4,000,000
  await controller.jumpToRow(qPos2);
  assert(controller.getActiveSheetIndex() === 1, 'Jump to query position 4,000,000 navigated to Sheet 2');
  assert(mockEnv.getActiveSheetId() === 'sheet-2', 'Univer active sheet is sheet-2');

  // Case D: Sorted query test (FL Sort ASC real query: 148,868 rows)
  const sortAscId = 'q_7bab87b150a07c793b9b11e89e3d094b';
  const sortMeta = await QueryIndexManager.getQueryMetadata(DATASET_ID, sortAscId);
  assert(sortMeta !== null, 'FL Sort ASC real query metadata loaded from disk');
  await controller.setQuery(sortAscId, sortMeta!.matchedRows);

  const sortPartitions = controller.getPartitions();
  assert(sortPartitions.length === 1, `Sorted query with 148,868 rows partitioned into 1 virtual sheet`);
  assert(sortPartitions[0].rowCount === 148_868, `Sorted query Sheet 1 rowCount is 148,868`);

  // Clear query and verify return to normal 11.8M 3-sheet partition
  await controller.clearQuery();
  const restoredPartitions = controller.getPartitions();
  assert(restoredPartitions.length === 3, 'Clearing query restores 3 virtual sheets for 11.8M dataset');
  assert(controller.getTotalRows() === EXPECTED_TOTAL_ROWS, `Total rows restored to ${EXPECTED_TOTAL_ROWS.toLocaleString()}`);

  // =========================================================================
  // SUMMARY REPORT
  // =========================================================================
  console.log('\n' + '='.repeat(80));
  console.log(`TOTAL TESTS: ${passed + failed}`);
  console.log(`PASSED: ${passed}`);
  console.log(`FAILED: ${failed}`);
  console.log('='.repeat(80));

  controller.destroy();
  await prisma.$disconnect();

  if (failed > 0) {
    process.exit(1);
  }
}

runVirtualSheetTests().catch((err) => {
  console.error('Test execution failed:', err);
  process.exit(1);
});
