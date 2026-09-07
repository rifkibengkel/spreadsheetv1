'use client';

import React, { useEffect, useRef, useState, useCallback } from 'react';
import { createUniver, LocaleType, mergeLocales } from '@univerjs/presets';
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core';
import UniverPresetSheetsCoreEnUS from '@univerjs/preset-sheets-core/locales/en-US';
import { UniverSheetsSortPreset } from '@univerjs/preset-sheets-sort';
import { UniverSheetsFilterPreset } from '@univerjs/preset-sheets-filter';
import UniverPresetSheetsSortEnUS from '@univerjs/preset-sheets-sort/locales/en-US';
import UniverPresetSheetsFilterEnUS from '@univerjs/preset-sheets-filter/locales/en-US';
import { SheetInterceptorService, INTERCEPTOR_POINT, SetRangeValuesMutation } from '@univerjs/sheets';
import { InterceptorEffectEnum, ICommandService } from '@univerjs/core';
import { IActiveDirtyManagerService } from '@univerjs/engine-formula';
import '@univerjs/presets/lib/styles/preset-sheets-core.css';
import '@univerjs/preset-sheets-core/lib/index.css';
import '@univerjs/preset-sheets-sort/lib/index.css';
import '@univerjs/preset-sheets-filter/lib/index.css';

import { initialWorkbookData } from '@/lib/univer/univerConfig';
import { workbookSession } from '@/lib/session/workbookSession';
import { CalculationScheduler } from '@/lib/workers/pool/calculationScheduler';
import { installUniverCalculationPassivation, applyCalculationPatchSilent } from '@/lib/univer/univerPassivation';
import { evaluateMicroFormula, extractReferencedCells } from '@/lib/univer/microFormulaEvaluator';
import SpreadsheetToolbar from './SpreadsheetToolbar';

export default function Spreadsheet() {
  const containerRef = useRef<HTMLDivElement>(null);
  const univerRef = useRef<any>(null);
  const schedulerRef = useRef<CalculationScheduler | null>(null);
  const formulaRegistryRef = useRef<Map<string, { sheetId: string; row: number; col: number; formula: string }>>(new Map());
  const [univerAPI, setUniverAPI] = useState<any>(null);

  const [activeFileName, setActiveFileName] = useState('workbook-initial.xlsx');
  const [activeFileFormat, setActiveFileFormat] = useState<'xlsx' | 'csv'>('xlsx');

  const handleFileImported = useCallback((filename: string, format: 'xlsx' | 'csv') => {
    setActiveFileName(filename);
    setActiveFileFormat(format);
  }, []);

  useEffect(() => {
    if (!containerRef.current) return;

    if (univerRef.current?.univer) {
      setUniverAPI(univerRef.current.univerAPI);
      return;
    }

    let currentRefObj: any = null;
    const pendingFormulaDisplayMap = new Map<string, { previousValue: any }>();
    let interceptorDisposable: { dispose: () => void } | null = null;

    try {
      if (containerRef.current && containerRef.current.children.length > 0) {
        containerRef.current.innerHTML = '';
      }

      // Configure in-process formula computation (Zero IPC / No structured clone transfer of 9.6M cells)
      const corePreset = UniverSheetsCorePreset({
        container: containerRef.current,
        formula: {
          initialFormulaComputing: 2, // CalculationMode.NO_CALCULATION on initial load
        },
      } as any);

      const sortPreset = UniverSheetsSortPreset();
      const filterPreset = UniverSheetsFilterPreset();

      const { univer, univerAPI: api } = createUniver({
        locale: LocaleType.EN_US,
        locales: {
          [LocaleType.EN_US]: mergeLocales(
            UniverPresetSheetsCoreEnUS,
            UniverPresetSheetsSortEnUS,
            UniverPresetSheetsFilterEnUS
          ),
        },
        presets: [corePreset, sortPreset, filterPreset],
      });

      currentRefObj = { univer, univerAPI: api };
      univerRef.current = currentRefObj;
      (window as any).univerAPI = api;
      (window as any).univer = univer;

      // Register SheetInterceptorService and Enterprise Calculation Passivation BEFORE sheet creation
      try {
        const injector = (api as any)?._injector || (univer as any)?.__injector || (univer as any)?._injector;
        if (injector) {
          // Enterprise Calculation Passivation: Completely suppresses browser-side formula recalculation storms
          try {
            installUniverCalculationPassivation(injector);
          } catch (passivationErr) {
            console.warn('[UNIVER_PASSIVATION] Error installing calculation passivation:', passivationErr);
          }

          const sheetInterceptorService = injector.get(SheetInterceptorService);
          if (sheetInterceptorService && INTERCEPTOR_POINT?.CELL_CONTENT) {
            interceptorDisposable = sheetInterceptorService.intercept(INTERCEPTOR_POINT.CELL_CONTENT, {
              priority: 100,
              effect: InterceptorEffectEnum.Value,
              handler(cell: any, context: any, next: any) {
                if (cell && (cell.f || context?.rawData?.f) && (cell.v === null || cell.v === undefined)) {
                  const key = `${context?.unitId}_${context?.subUnitId}_${context?.row}_${context?.col}`;
                  if (pendingFormulaDisplayMap.has(key)) {
                    const fallback = pendingFormulaDisplayMap.get(key);
                    return next({
                      ...cell,
                      v: fallback?.previousValue,
                    });
                  }
                }
                return next(cell);
              },
            });
          }
        }
      } catch (interceptErr) {
        console.warn('SheetInterceptorService registration warning:', interceptErr);
      }

      // Initialize the workbook snapshot directly
      api.createUniverSheet(initialWorkbookData);
      setUniverAPI(api);

      // Initialize 4-worker Calculation Pool and Scheduler
      const scheduler = new CalculationScheduler(
        {
          onCalculationCompleted(result) {
            try {
              const workbook = api.getActiveWorkbook?.();
              const worksheet = result.subUnitId
                ? workbook?.getSheetBySheetId?.(result.subUnitId)
                : workbook?.getActiveSheet?.();
              if (worksheet) {
                const range = worksheet.getRange?.(result.targetCell.row, result.targetCell.col);
                if (range) {
                  range.setValue(result.value);
                }
              }
            } catch (applyErr) {
              console.warn('[CalculationScheduler] Failed to apply result:', applyErr);
            }
          },
        },
        4 // 4 Dedicated Workers
      );
      schedulerRef.current = scheduler;
      (window as any).calculationScheduler = scheduler;

      // Temporary diagnostic instrumentation for forensic investigation (Section 10)
      const diagState = {
        commandCounts: {} as Record<string, number>,
        sortLifecycle: [] as any[],
        filterLifecycle: [] as any[],
        calculationRequests: 0,
        workerScheduledCount: 0,
      };
      (window as any).__diagState = diagState;

      // Track active sort/reorder operations to distinguish row reorder shifting from genuine user formula edits
      let isSortReorderActive = false;
      try {
        const cmdSvc = (api as any)?._injector?.get(ICommandService) || (api as any)?.getCommandService?.();
        if (cmdSvc?.beforeCommandExecuted) {
          cmdSvc.beforeCommandExecuted((command: any) => {
            if (command?.id === 'sheet.command.sort-range' || command?.id === 'sheet.command.reorder-range') {
              isSortReorderActive = true;
            }
          });
        }
      } catch (e) {
        console.warn('Command service listener setup notice:', e);
      }

      // Register mutation listener for copy-through dirty tracking and formula visual lifecycle
      try {
        if (api.onCommandExecuted) {
          api.onCommandExecuted((commandInfo: any, options: any) => {
            const commandId = commandInfo?.id || 'unknown';
            const params = commandInfo?.params;
            const subUnitId = params?.subUnitId || params?.sheetId;
            const unitId = params?.unitId || api.getActiveWorkbook?.()?.getId?.() || '';

            diagState.commandCounts[commandId] = (diagState.commandCounts[commandId] || 0) + 1;

            if (commandId === 'sheet.command.sort-range' || commandId === 'sheet.command.reorder-range') {
              isSortReorderActive = false;
            }

            if (
              commandId.includes('sort') ||
              commandId.includes('filter') ||
              commandId.includes('reorder') ||
              commandId.includes('calc') ||
              commandId.includes('formula') ||
              commandId.includes('dirty')
            ) {
              console.log(`[DIAG_COMMAND] ${commandId}`, { params, options, count: diagState.commandCounts[commandId] });
            }

            if (commandId === 'formula.mutation.set-formula-calculation-start') {
              diagState.calculationRequests++;
              console.warn(`[DIAG_CALCULATION_STORM] SetFormulaCalculationStartMutation FIRED! (count=${diagState.calculationRequests})`, params);
            }

            // Rapid Sheet Switching Protection: Update active sheet & cancel obsolete background jobs
            if (commandId === 'sheet.operation.set-worksheet-active') {
              const activeSub = params?.subUnitId || params?.sheetId;
              if (activeSub && schedulerRef.current) {
                schedulerRef.current.onSheetActivated(activeSub);
              }
            }

            // Dirty sheet tracking for export: ONLY on actual cell mutations / structural changes
            if (
              (commandId === 'sheet.mutation.set-range-values' ||
                commandId === 'sheet.mutation.remove-rows' ||
                commandId === 'sheet.mutation.insert-row' ||
                commandId === 'sheet.mutation.remove-col' ||
                commandId === 'sheet.mutation.insert-col' ||
                commandId === 'sheet.mutation.reorder-range') &&
              !options?.fromFormula &&
              !options?.applyFormulaCalculationResult
            ) {
              if (subUnitId) {
                workbookSession.markSheetDirty(subUnitId);
              } else {
                const activeSheet = api.getActiveWorkbook?.()?.getActiveSheet?.();
                if (activeSheet) {
                  const sName = activeSheet.getSheetName?.() || activeSheet.getSheetId?.();
                  if (sName) workbookSession.markSheetDirty(sName);
                }
              }
            }

            // Capture cell deltas and dispatch to backend calculation engine (ONLY for genuine edits, NOT sort row reorders or silent patches)
            if (
              commandId === 'sheet.mutation.set-range-values' &&
              !options?.fromFormula &&
              !options?.applyFormulaCalculationResult &&
              !options?.fromCalculationPatch &&
              !options?.onlyLocal &&
              !isSortReorderActive &&
              !(window as any).__isSortActive
            ) {
              const cellValue = params?.cellValue;
              if (cellValue && typeof cellValue === 'object') {
                const workbook = api.getActiveWorkbook?.();
                const worksheet =
                  (workbook?.getSheets ? workbook.getSheets().find((s: any) => s.getSheetId?.() === subUnitId) : null) ||
                  (workbook as any)?.getWorkbook?.()?.getSheetBySheetId?.(subUnitId) ||
                  workbook?.getActiveSheet?.();
                const targetSheetId = subUnitId || worksheet?.getSheetId?.() || 'sheet-1';
                const matrix = (worksheet as any)?.getSheet?.()?.getCellMatrix?.() || (worksheet as any)?.getCellMatrix?.();
                const deltas: Array<{ sheetId: string; row: number; col: number; v?: any; f?: string }> = [];

                const lookupVal = (sheetName: string | undefined, rIdx: number, cIdx: number) => {
                  let targetWs = worksheet;
                  if (sheetName) {
                    targetWs =
                      (workbook?.getSheets
                        ? workbook.getSheets().find((s: any) => {
                            const sName = (s.getSheetName?.() || s.getName?.() || s.getSheetId?.() || '').toLowerCase();
                            return sName === sheetName.toLowerCase();
                          })
                        : null) || worksheet;
                  }
                  const mat = (targetWs as any)?.getSheet?.()?.getCellMatrix?.() || (targetWs as any)?.getCellMatrix?.();
                  const cell = mat?.getValue?.(rIdx, cIdx) || (targetWs as any)?.getCellRaw?.(rIdx, cIdx);
                  return cell?.v;
                };

                const immediatePatches: Array<{ sheetId: string; row: number; col: number; v: any; f?: string; t?: number }> = [];

                for (const rStr of Object.keys(cellValue)) {
                  const r = parseInt(rStr, 10);
                  const rowObj = cellValue[r];
                  if (rowObj && typeof rowObj === 'object') {
                    for (const cStr of Object.keys(rowObj)) {
                      const c = parseInt(cStr, 10);
                      const cellItem = rowObj[c];
                      if (!cellItem) continue;

                      // Handle formula input: evaluate immediately with zero UI freeze
                      if (cellItem?.f && typeof cellItem.f === 'string' && cellItem.f.startsWith('=')) {
                        formulaRegistryRef.current.set(`${targetSheetId}_${r}_${c}`, {
                          sheetId: targetSheetId,
                          row: r,
                          col: c,
                          formula: cellItem.f,
                        });

                        const computed = evaluateMicroFormula(cellItem.f, lookupVal);
                        if (computed !== null && computed !== undefined) {
                          immediatePatches.push({
                            sheetId: targetSheetId,
                            row: r,
                            col: c,
                            f: cellItem.f,
                            v: computed,
                            t: 2,
                          });
                        }

                        // Extract precedent cells so backend WorkbookEngine and BullMQ receive their values
                        const refs = extractReferencedCells(cellItem.f);
                        for (const ref of refs) {
                          const refVal = lookupVal(ref.sheetName, ref.row, ref.col);
                          if (refVal !== undefined && refVal !== null) {
                            deltas.push({
                              sheetId: targetSheetId,
                              row: ref.row,
                              col: ref.col,
                              v: refVal,
                            });
                          }
                        }
                      } else if (cellItem?.v !== undefined) {
                        // User modified a value: check if any registered formulas depend on this cell
                        for (const [, entry] of formulaRegistryRef.current.entries()) {
                          if (entry.sheetId === targetSheetId && !(entry.row === r && entry.col === c)) {
                            const refs = extractReferencedCells(entry.formula);
                            const isDep = refs.some((rf) => rf.row === r && rf.col === c);
                            if (isDep) {
                              const recomputed = evaluateMicroFormula(entry.formula, (s, ro, co) => {
                                if (ro === r && co === c) return cellItem.v;
                                return lookupVal(s, ro, co);
                              });
                              if (recomputed !== null && recomputed !== undefined) {
                                immediatePatches.push({
                                  sheetId: entry.sheetId,
                                  row: entry.row,
                                  col: entry.col,
                                  f: entry.formula,
                                  v: recomputed,
                                  t: 2,
                                });
                                deltas.push({
                                  sheetId: entry.sheetId,
                                  row: entry.row,
                                  col: entry.col,
                                  f: entry.formula,
                                  v: recomputed,
                                });
                              }
                            }
                          }
                        }
                      }

                      deltas.push({
                        sheetId: targetSheetId,
                        row: r,
                        col: c,
                        v: cellItem.v,
                        f: cellItem.f,
                      });
                    }
                  }
                }

                // Immediately apply instant patches to Univer with zero UI freeze
                if (immediatePatches.length > 0) {
                  applyCalculationPatchSilent(api, unitId, immediatePatches);
                }

                if (deltas.length > 0) {
                  diagState.workerScheduledCount++;
                  console.log(`[BACKEND_CALCULATION] Dispatching ${deltas.length} delta(s) to backend API (batch #${diagState.workerScheduledCount})`);

                  // Asynchronously send to Next.js PostgreSQL WAL & Backend Calculation Engine
                  fetch(`/api/v1/workbooks/${unitId}/mutations`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ deltas }),
                  })
                    .then((res) => res.json())
                    .then((resData) => {
                      if (resData?.patches && Array.isArray(resData.patches) && resData.patches.length > 0) {
                        const validPatches = resData.patches.filter((p: any) => p.v !== null && p.v !== undefined && p.v !== '' && !String(p.v).includes('#'));
                        if (validPatches.length > 0) {
                          console.log(`[BACKEND_CALCULATION] Applying ${validPatches.length} patch(es) from backend engine (exec=${resData.executionMs}ms)`);
                          applyCalculationPatchSilent(api, unitId, validPatches);
                        }
                      }
                    })
                    .catch((err) => {
                      console.warn('[BACKEND_CALCULATION] Mutation sync notice:', err);
                    });
                }
              }
            }

            // Clear pending visual entry when authoritative calculation result arrives
            if (
              commandId === 'sheet.mutation.set-range-values' &&
              (options?.applyFormulaCalculationResult || options?.fromFormula || options?.fromCalculationPatch)
            ) {
              const cellValue = params?.cellValue;
              if (cellValue && typeof cellValue === 'object') {
                for (const rStr of Object.keys(cellValue)) {
                  const r = parseInt(rStr, 10);
                  const rowObj = cellValue[r];
                  if (rowObj && typeof rowObj === 'object') {
                    for (const cStr of Object.keys(rowObj)) {
                      const c = parseInt(cStr, 10);
                      const key = `${unitId}_${subUnitId}_${r}_${c}`;
                      pendingFormulaDisplayMap.delete(key);
                    }
                  }
                }
              }
            }
          });
        }
      } catch (cmdErr) {
        console.warn('Univer command tracking setup warning:', cmdErr);
      }



      setTimeout(() => {
        if (typeof window !== 'undefined') {
          window.dispatchEvent(new Event('resize'));
        }
      }, 100);
    } catch (e) {
      console.error('Failed to initialize Univer:', e);
    }

    return () => {
      const activeObj = currentRefObj;
      pendingFormulaDisplayMap.clear();

      if (schedulerRef.current) {
        schedulerRef.current.dispose();
        schedulerRef.current = null;
      }

      if (interceptorDisposable?.dispose) {
        try {
          interceptorDisposable.dispose();
        } catch {}
      }

      setTimeout(() => {
        if (univerRef.current !== activeObj && activeObj?.univer) {
          try {
            activeObj.univer.dispose();
          } catch (e) {
            console.warn('Dispose warning:', e);
          }
        }
      }, 300);
    };
  }, []);

  return (
    <div className="univer-wrapper">
      <div className="univer-header">
        <p>Edit, Import, & Save Excel / CSV Files directly | Glory-Glory Masferr.AI</p>
      </div>
      <SpreadsheetToolbar
        univerAPI={univerAPI}
        activeFileName={activeFileName}
        activeFileFormat={activeFileFormat}
        onFileImported={handleFileImported}
      />
      <div className="univer-container-outer">
        <div ref={containerRef} className="univer-container-inner" />
      </div>
    </div>
  );
}