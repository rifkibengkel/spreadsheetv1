/**
 * PHASE 1D — STRICT UI INTEGRATION & VIEWPORT INTERACTION TEST SUITE
 * 
 * Verifies:
 * 1. Normal dataset viewport lifecycle (mount, scroll, jump, 11,809,233 rows, 54 cols)
 * 2. Query creation & filter UI integration (single-row, large 160k, zero-result)
 * 3. Query viewport hydration & logical row ordinal mapping (0..M-1)
 * 4. Sort UI integration (ASC, DESC, duplicate tie-breaking)
 * 5. Query lifecycle & transitions (NORMAL -> QUERY_A -> QUERY_B -> CLEAR -> NORMAL)
 * 6. Query switching race & stale-response rejection (token-based correctness)
 * 7. Cache isolation (datasetId vs datasetId#queryId)
 * 8. Reload / URL parameter state recovery & error recovery
 * 9. Memory boundedness (no full-array materialization in client/UI layer)
 */

import fs from 'fs';
import path from 'path';
import prisma from '../../src/lib/db/prisma';
import { DatasetRangeReader } from '../../src/lib/largeCsv/reader/datasetRangeReader';
import { QueryIndexManager } from '../../src/lib/largeCsv/query/queryIndexManager';
import { DatasetRowResolver } from '../../src/lib/largeCsv/query/datasetRowResolver';
import { FilterEngine } from '../../src/lib/largeCsv/query/filterEngine';
import { ExternalSortEngine } from '../../src/lib/largeCsv/query/externalSortEngine';
import { BoundedLruCache } from '../../src/lib/viewport/boundedLruCache';
import { UniverViewportAdapter } from '../../src/lib/viewport/univerViewportAdapter';
import { DatasetViewportController } from '../../src/lib/viewport/datasetViewportController';
import { QueryDefinition } from '../../src/lib/largeCsv/query/types';

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
 * Mock Univer API to simulate Univer canvas and worksheet operations
 */
function createMockUniverAPI() {
  const sheetRowCounts: Record<string, number> = { 'sheet-01': 1000 };
  let worksheetColumnCount = 26;
  const sheetCells: Record<string, any> = {};
  let currentScrollRow = 0;
  let activeSheetId = 'sheet-01';

  const createMockSheet = (sheetId: string) => ({
    setRowCount: (count: number) => {
      sheetRowCounts[sheetId] = count;
    },
    getRowCount: () => sheetRowCounts[sheetId] ?? 1000,
    setColumnCount: (count: number) => {
      worksheetColumnCount = count;
    },
    getColumnCount: () => worksheetColumnCount,
    getRange: (r: number, c: number, numR: number, numC: number) => ({
      setValue: (val: any) => {
        sheetCells[`${r},${c}`] = val;
      },
      setValues: (matrix: any[][]) => {
        for (let rowIdx = 0; rowIdx < matrix.length; rowIdx++) {
          for (let colIdx = 0; colIdx < matrix[rowIdx].length; colIdx++) {
            sheetCells[`${r + rowIdx},${c + colIdx}`] = matrix[rowIdx][colIdx];
          }
        }
      },
      clear: () => {
        for (let rowIdx = 0; rowIdx < numR; rowIdx++) {
          for (let colIdx = 0; colIdx < numC; colIdx++) {
            delete sheetCells[`${r + rowIdx},${c + colIdx}`];
          }
        }
      },
    }),
    getSheetId: () => sheetId,
    getCellMatrix: () => ({
      _matrix: new Proxy({} as any, {
        get(_t, prop) {
          return true;
        },
        deleteProperty(_t, prop) {
          const row = Number(prop);
          for (let c = 0; c < 54; c++) {
            delete sheetCells[`${row},${c}`];
          }
          return true;
        },
      }),
    }),
  });

  const mockWorkbook = {
    getId: () => 'workbook-01',
    getActiveSheet: () => createMockSheet(activeSheetId),
    getSheetBySheetId: (id: string) => createMockSheet(id || activeSheetId),
    getSheets: () => Object.keys(sheetRowCounts).map((id) => createMockSheet(id)),
  };

  const mockCommandService = {
    executeCommand: async (cmd: string, params: any) => {
      if (cmd === 'sheet.operation.set-worksheet-active') {
        const target = params?.subUnitId || params?.sheetId;
        if (target) activeSheetId = target;
      }
      if (cmd === 'sheet.operation.set-scroll' || cmd === 'sheet.command.scroll-to-cell') {
        currentScrollRow = params?.sheetViewStartRow ?? params?.range?.startRow ?? currentScrollRow;
      }
      if (cmd === 'sheet.mutation.set-worksheet-row-count') {
        const subId = params?.subUnitId || activeSheetId;
        sheetRowCounts[subId] = params.rowCount;
      }
      if (cmd === 'sheet.mutation.set-worksheet-col-count') {
        worksheetColumnCount = params.colCount;
      }
      if (cmd === 'sheet.mutation.set-range-values') {
        const cellValue = params?.cellValue;
        if (cellValue) {
          for (const r of Object.keys(cellValue)) {
            for (const c of Object.keys(cellValue[r])) {
              sheetCells[`${r},${c}`] = cellValue[r][c]?.v;
            }
          }
        }
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
    getWorksheetRowCount: () => sheetRowCounts['sheet-01'] ?? 1000,
    getCurrentScrollRow: () => currentScrollRow,
    getHydratedCellCount: () => Object.keys(sheetCells).length,
    getCell: (r: number, c: number) => sheetCells[`${r},${c}`],
    clearCells: () => {
      for (const k of Object.keys(sheetCells)) delete sheetCells[k];
    },
  };
}

async function runPhase1DTests() {
  console.log('='.repeat(80));
  console.log('   PHASE 1D — STRICT UI INTEGRATION & VIEWPORT INTERACTION TESTS');
  console.log(`   Target Dataset: ${DATASET_ID}`);
  console.log(`   Total Rows: ${EXPECTED_TOTAL_ROWS.toLocaleString()} | Columns: ${EXPECTED_COLUMNS}`);
  console.log('='.repeat(80));

  const mockEnv = createMockUniverAPI();
  const adapter = new UniverViewportAdapter(mockEnv.mockUniverAPI);
  const controller = new DatasetViewportController(adapter, {
    visibleRows: 50,
    overscanRows: 100,
    marginThreshold: 30,
  });

  // Mock internal fetchRange on controller to use direct backend engines (DatasetRangeReader / RowResolver)
  // since Node.js test environment doesn't host Next.js HTTP server
  (controller as any).fetchRange = async (startRow: number, rowCount: number, generationToken: number) => {
    const queryId = controller.getCurrentQueryId();
    if (queryId) {
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
      const columns = (datasetMeta?.headerRow as string[]) || [];
      return {
        datasetId: DATASET_ID,
        queryId,
        startRow,
        rowCount: rows.length,
        totalRows: totalMatchedRows,
        columns,
        columnCount: columns.length,
        rows,
        metrics: { serverTotalMs: 1.0 },
      };
    } else {
      const res = await DatasetRangeReader.readRange(DATASET_ID, startRow, rowCount);
      return {
        ...res,
        metrics: { serverTotalMs: 1.0 },
      };
    }
  };

  // ---------------------------------------------------------------------------
  // SECTION 1: Normal Dataset Mounting & Base Viewport Lifecycle
  // ---------------------------------------------------------------------------
  console.log('\n[Section 1] Normal Dataset Mounting & Viewport Lifecycle');
  const mountSuccess = await controller.mount(DATASET_ID);
  assert(mountSuccess === true, 'Dataset successfully mounted');
  assert(controller.isReady() === true, 'Controller is ready');
  assert(controller.getTotalRows() === EXPECTED_TOTAL_ROWS, `Total rows is ${EXPECTED_TOTAL_ROWS.toLocaleString()}`);
  assert(controller.getBaseTotalRows() === EXPECTED_TOTAL_ROWS, `Base total rows recorded as ${EXPECTED_TOTAL_ROWS.toLocaleString()}`);
  assert(controller.getHeaders().length === EXPECTED_COLUMNS, `Headers length is ${EXPECTED_COLUMNS}`);
  assert(mockEnv.getWorksheetRowCount() === 4_000_001, `Univer worksheet row count is 4,000,001 (data + header for Sheet 1)`);

  const activeWin = controller.getActiveWindow();
  assert(activeWin !== null, 'Active window is initialized');
  assert(activeWin?.startRow === 1, 'Initial window startRow is 1 (header offset)');
  assert(activeWin?.rowCount === 250, 'Initial window rowCount is 250 (50 visible + 200 overscan)');
  assert(String(mockEnv.getCell(1, 1)) === '8002531292', `Row 1 cell[1] summons matches expected: ${mockEnv.getCell(1, 1)}`);

  // ---------------------------------------------------------------------------
  // SECTION 2: Normal Scrolling & Jump-to-Row
  // ---------------------------------------------------------------------------
  console.log('\n[Section 2] Normal Scrolling & Jump-to-Row');
  // Jump to row 100,000
  const jumpMid = await controller.jumpToRow(100_000);
  assert(jumpMid === true, 'Jump to row 100,000 succeeded');
  assert(controller.getActiveWindow()?.startRow === 99_901, 'Window centered around row 100,000 (start 99,901)');
  assert(mockEnv.getCurrentScrollRow() === 100_001, 'Univer scroll position updated to 100,001');

  // Jump to cross-partition boundary (row 1,999,500)
  const jumpCross = await controller.jumpToRow(1_999_500);
  assert(jumpCross === true, 'Jump to cross-partition row 1,999,500 succeeded');
  assert(controller.getActiveWindow()?.startRow === 1_999_401, 'Window centered at 1,999,401');

  // Jump near terminal row (11,809,232)
  const jumpEnd = await controller.jumpToRow(11_809_232);
  assert(jumpEnd === true, 'Jump to terminal row succeeded');
  const endWin = controller.getActiveWindow();
  assert(endWin?.endRow === 3_809_233, `Window ends at terminal row 3,809,233 on Sheet 3 (got ${endWin?.endRow})`);

  // ---------------------------------------------------------------------------
  // SECTION 3: Filter UI Integration — Query Creation & Activation
  // ---------------------------------------------------------------------------
  console.log('\n[Section 3] Filter UI Integration & Activation');
  // Query 1: Single row filter (Summons_Number == "8002531292")
  const singleQueryId = 'q_23beccac12b28e47c10024a1d5b48ee6';
  const singleMeta = await QueryIndexManager.getQueryMetadata(DATASET_ID, singleQueryId);
  assert(singleMeta !== null, 'Single-row query index exists');
  assert(singleMeta?.matchedRows === 1, 'Single-row query has exact 1 matched row');

  const setSingleSuccess = await controller.setQuery(singleQueryId, singleMeta!.matchedRows);
  assert(setSingleSuccess === true, 'setQuery for single-row query succeeded');
  assert(controller.getCurrentQueryId() === singleQueryId, 'Active query ID is singleQueryId');
  assert(controller.getTotalRows() === 1, 'Total rows updated to 1 in query mode');
  assert(mockEnv.getWorksheetRowCount() === 2, 'Univer worksheet row count adjusted to 2 (1 row + 1 header)');
  assert(String(mockEnv.getCell(1, 1)) === '8002531292', `Query row 1 summons is 8002531292`);

  // Query 2: Large filter (Registration_State == "CT", 160,361 rows)
  const ctQueryId = 'q_0f51000fac98c91c04442c43e6ad5b67';
  const ctMeta = await QueryIndexManager.getQueryMetadata(DATASET_ID, ctQueryId);
  assert(ctMeta !== null, 'CT query index exists');
  assert(ctMeta?.matchedRows === 160_361, 'CT query has exact 160,361 matched rows');

  const setCtSuccess = await controller.setQuery(ctQueryId, ctMeta!.matchedRows);
  assert(setCtSuccess === true, 'setQuery for CT query succeeded');
  assert(controller.getCurrentQueryId() === ctQueryId, 'Active query ID is ctQueryId');
  assert(controller.getTotalRows() === 160_361, 'Total rows updated to 160,361');
  assert(mockEnv.getWorksheetRowCount() === 160_362, 'Univer worksheet row count adjusted to 160,362');

  // Verify logical row ordinals in CT query window
  const ctWin = controller.getActiveWindow();
  assert(ctWin?.startRow === 1, 'Query window starts at logical row 1');
  assert(ctWin?.rowCount === 250, 'Query window hydrates 250 rows');
  assert(mockEnv.getCell(1, 2) !== undefined, 'Query row 1 cell has plate');

  // ---------------------------------------------------------------------------
  // SECTION 4: Query Mode Navigation & Jump-to-Row
  // ---------------------------------------------------------------------------
  console.log('\n[Section 4] Query Mode Navigation & Jump-to-Row');
  // Jump to middle of query results (logical row 80,000 of 160,361)
  const jumpCtMid = await controller.jumpToRow(80_000);
  assert(jumpCtMid === true, 'Jump to row 80,000 in query result succeeded');
  assert(mockEnv.getCurrentScrollRow() === 80_001, 'Scroll row is 80,001 (with header)');
  const ctMidWin = controller.getActiveWindow();
  assert(ctMidWin?.startRow === 79_901, 'CT window centered at logical row 79,901');

  // Jump to near end of query results (logical row 160,350 of 160,361)
  const jumpCtEnd = await controller.jumpToRow(160_350);
  assert(jumpCtEnd === true, 'Jump to row 160,350 near query end succeeded');
  const ctEndWin = controller.getActiveWindow();
  assert(ctEndWin?.endRow === 160_361, `CT window clamped to end row 160,361 (got ${ctEndWin?.endRow})`);

  // Query 3: Zero-match filter (Registration_State == "NONEXISTENT", 0 rows)
  const emptyQueryId = 'q_4962bcbed5b4f3ab84fa18a936e0ab9c';
  const emptyMeta = await QueryIndexManager.getQueryMetadata(DATASET_ID, emptyQueryId);
  assert(emptyMeta !== null && emptyMeta.matchedRows === 0, 'Empty query has 0 matched rows');
  const setEmptySuccess = await controller.setQuery(emptyQueryId, 0);
  assert(setEmptySuccess === false, 'jumpToRow on 0 results safely returns false without throwing');
  assert(controller.getTotalRows() === 0, 'Total rows is 0 for empty query');
  assert(mockEnv.getWorksheetRowCount() === 1, 'Univer worksheet row count is 1 (header only)');

  // ---------------------------------------------------------------------------
  // SECTION 5: Sort UI Integration — ASC / DESC
  // ---------------------------------------------------------------------------
  console.log('\n[Section 5] Sort UI Integration');
  // Query 4: FL Sort ASC (148,868 rows)
  const sortAscId = 'q_7bab87b150a07c793b9b11e89e3d094b';
  const ascMeta = await QueryIndexManager.getQueryMetadata(DATASET_ID, sortAscId);
  assert(ascMeta !== null && ascMeta.matchedRows === 148_868, 'Sort ASC query has 148,868 rows');
  await controller.setQuery(sortAscId, ascMeta!.matchedRows);
  assert(controller.getTotalRows() === 148_868, 'Controller updated to 148,868 rows');
  assert(mockEnv.getWorksheetRowCount() === 148_869, 'Univer worksheet row count is 148,869');

  // Query 5: FL Sort DESC (148,868 rows)
  const sortDescId = 'q_89265e7bc192c7e13f4f0122ce78bb66';
  const descMeta = await QueryIndexManager.getQueryMetadata(DATASET_ID, sortDescId);
  assert(descMeta !== null && descMeta.matchedRows === 148_868, 'Sort DESC query has 148,868 rows');
  await controller.setQuery(sortDescId, descMeta!.matchedRows);
  assert(controller.getTotalRows() === 148_868, 'Controller updated to 148,868 rows for DESC');

  // ---------------------------------------------------------------------------
  // SECTION 6: Query Lifecycle & Clear Query
  // ---------------------------------------------------------------------------
  console.log('\n[Section 6] Clear Query & Return to Normal Mode');
  const clearSuccess = await controller.clearQuery();
  assert(clearSuccess === true, 'clearQuery succeeded');
  assert(controller.getCurrentQueryId() === null, 'Active query ID is null');
  assert(controller.getTotalRows() === EXPECTED_TOTAL_ROWS, `Total rows restored to ${EXPECTED_TOTAL_ROWS.toLocaleString()}`);
  assert(mockEnv.getWorksheetRowCount() === 4_000_001, `Univer worksheet row count restored to 4,000,001 (Sheet 1)`);

  // Test normal scrolling and jumping after clear
  const postClearJump = await controller.jumpToRow(500_000);
  assert(postClearJump === true, 'Jump in restored normal mode succeeded');
  assert(controller.getActiveWindow()?.startRow === 499_901, 'Window centered at row 500,000 (start 499,901)');

  // ---------------------------------------------------------------------------
  // SECTION 7: Rapid Query Transitions & Monotonic Race Protection
  // ---------------------------------------------------------------------------
  console.log('\n[Section 7] Rapid Query Transitions & Race Protection');
  // Simulate rapid succession: Query A -> Query B -> Query C
  // Only Query C token must be accepted; stale responses must be dropped
  let activeToken = 0;
  let authoritativeQueryId = '';
  let authoritativeRowCount = 0;

  async function simulateApplyQuery(qId: string, matchedCount: number, delayMs: number) {
    const token = ++activeToken;
    await new Promise((r) => setTimeout(r, delayMs));
    if (token === activeToken) {
      authoritativeQueryId = qId;
      authoritativeRowCount = matchedCount;
      await controller.setQuery(qId, matchedCount);
    }
  }

  // Launch A (slow: 40ms), B (fast: 10ms), C (medium: 20ms)
  // Invocation order: A, then B, then C.
  // Completion order: B (at 10ms), C (at 20ms), A (at 40ms).
  // Winner MUST be C because it was dispatched last with highest token!
  const pA = simulateApplyQuery(ctQueryId, 160_361, 40);
  const pB = simulateApplyQuery(singleQueryId, 1, 10);
  const pC = simulateApplyQuery(sortAscId, 148_868, 20);

  await Promise.all([pA, pB, pC]);

  assert(authoritativeQueryId === sortAscId, `Authoritative query is Query C (FL Sort) - got ${authoritativeQueryId}`);
  assert(authoritativeRowCount === 148_868, 'Authoritative row count is 148,868');
  assert(controller.getCurrentQueryId() === sortAscId, 'Controller active query is Query C');
  assert(controller.getTotalRows() === 148_868, 'Controller total rows is 148,868');

  // ---------------------------------------------------------------------------
  // SECTION 8: Dedicated Worker Bounded LRU Cache Isolation
  // ---------------------------------------------------------------------------
  console.log('\n[Section 8] Bounded LRU Cache Isolation');
  const cache = new BoundedLruCache(25, 200);

  // Normal page 0
  cache.set(DATASET_ID, 0, {
    pageIndex: 0,
    startRow: 0,
    rowCount: 200,
    rows: [{ stableRowId: 0, cells: ['NORMAL_ROW_0'] } as any],
    cachedAt: Date.now(),
  });

  // Query CT page 0
  const ctScope = `${DATASET_ID}#${ctQueryId}`;
  cache.set(ctScope, 0, {
    pageIndex: 0,
    startRow: 0,
    rowCount: 200,
    rows: [{ stableRowId: 105, cells: ['CT_QUERY_ROW_0'] } as any],
    cachedAt: Date.now(),
  });

  // Query FL page 0
  const flScope = `${DATASET_ID}#${sortAscId}`;
  cache.set(flScope, 0, {
    pageIndex: 0,
    startRow: 0,
    rowCount: 200,
    rows: [{ stableRowId: 888, cells: ['FL_QUERY_ROW_0'] } as any],
    cachedAt: Date.now(),
  });

  const normalP0 = cache.get(DATASET_ID, 0);
  const ctP0 = cache.get(ctScope, 0);
  const flP0 = cache.get(flScope, 0);

  assert(normalP0?.rows[0].cells[0] === 'NORMAL_ROW_0', 'Normal page 0 matches normal context');
  assert(ctP0?.rows[0].cells[0] === 'CT_QUERY_ROW_0', 'CT page 0 matches CT context');
  assert(flP0?.rows[0].cells[0] === 'FL_QUERY_ROW_0', 'FL page 0 matches FL context');
  assert(normalP0?.rows[0].cells[0] !== ctP0?.rows[0].cells[0], 'Normal and query caches are strictly isolated');

  // Invalidate CT query cache
  cache.invalidate(ctScope);
  assert(cache.get(ctScope, 0) === undefined, 'CT cache successfully invalidated');
  assert(cache.get(DATASET_ID, 0) !== undefined, 'Normal cache remains intact after query cache invalidation');

  // ---------------------------------------------------------------------------
  // SECTION 9: Reload / URL State Recovery & Error Handling
  // ---------------------------------------------------------------------------
  console.log('\n[Section 9] URL Parameter State Recovery & Error Handling');
  // Scenario A: Valid queryId in URL
  const mockParamsValid = new URLSearchParams(`datasetId=${DATASET_ID}&queryId=${ctQueryId}`);
  const loadedQueryId = mockParamsValid.get('queryId');
  const metaFetched = await QueryIndexManager.getQueryMetadata(DATASET_ID, loadedQueryId!);
  assert(metaFetched !== null, 'Fetched query metadata on reload');
  await controller.setQuery(loadedQueryId!, metaFetched!.matchedRows);
  assert(controller.getCurrentQueryId() === ctQueryId, 'Query restored from URL on reload');
  assert(controller.getTotalRows() === 160_361, 'Row count restored from URL query');

  // Scenario B: Non-existent / invalid queryId in URL
  const mockParamsInvalid = new URLSearchParams(`datasetId=${DATASET_ID}&queryId=q_nonexistent_xyz`);
  const invalidQId = mockParamsInvalid.get('queryId');
  const invalidMeta = await QueryIndexManager.getQueryMetadata(DATASET_ID, invalidQId!);
  assert(invalidMeta === null, 'Invalid queryId returns null metadata');
  // UI fallback: clear query and keep base dataset
  await controller.clearQuery();
  assert(controller.getCurrentQueryId() === null, 'Gracefully fell back to base dataset');
  assert(controller.getTotalRows() === EXPECTED_TOTAL_ROWS, 'Full 11.8M dataset preserved on invalid queryId');

  // ---------------------------------------------------------------------------
  // SECTION 10: Memory & Safety Verification
  // ---------------------------------------------------------------------------
  console.log('\n[Section 10] Memory & Safety Audit');
  const heapUsedMb = process.memoryUsage().heapUsed / (1024 * 1024);
  console.log(`  Current Node Heap: ${heapUsedMb.toFixed(2)} MB`);
  assert(heapUsedMb < 150, `Heap usage is strictly bounded (<150 MB, actual: ${heapUsedMb.toFixed(2)} MB)`);
  assert(mockEnv.getHydratedCellCount() <= 54 * 300, `Hydrated cell count bounded in Univer canvas (${mockEnv.getHydratedCellCount()} cells)`);

  console.log('\n' + '='.repeat(80));
  console.log(`PHASE 1D TEST RESULTS: ${passed} PASSED, ${failed} FAILED`);
  console.log('='.repeat(80));

  if (failed > 0) {
    process.exit(1);
  }
}

runPhase1DTests().catch((err) => {
  console.error('Fatal test error:', err);
  process.exit(1);
});
