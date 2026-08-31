'use client';

import React, { useEffect, useRef, useState, useCallback } from 'react';
import { createUniver, LocaleType, mergeLocales } from '@univerjs/presets';
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core';
import UniverPresetSheetsCoreEnUS from '@univerjs/preset-sheets-core/locales/en-US';
import { SheetInterceptorService, INTERCEPTOR_POINT } from '@univerjs/sheets';
import { InterceptorEffectEnum } from '@univerjs/core';
import '@univerjs/presets/lib/styles/preset-sheets-core.css';
import '@univerjs/preset-sheets-core/lib/index.css';

import { initialWorkbookData } from '@/lib/univer/univerConfig';
import { workbookSession } from '@/lib/session/workbookSession';
import SpreadsheetToolbar from './SpreadsheetToolbar';

export default function Spreadsheet() {
  const containerRef = useRef<HTMLDivElement>(null);
  const univerRef = useRef<any>(null);
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

      const { univer, univerAPI: api } = createUniver({
        locale: LocaleType.EN_US,
        locales: {
          [LocaleType.EN_US]: mergeLocales(UniverPresetSheetsCoreEnUS),
        },
        presets: [corePreset],
      });

      currentRefObj = { univer, univerAPI: api };
      univerRef.current = currentRefObj;
      (window as any).univerAPI = api;
      (window as any).univer = univer;

      // Set initial formula calculation mode to NO_CALCULATION before sheet creation
      try {
        const formulaEngine = api.getFormula ? api.getFormula() : null;
        if (formulaEngine && formulaEngine.setInitialFormulaComputing) {
          formulaEngine.setInitialFormulaComputing(2); // CalculationMode.NO_CALCULATION
        }
      } catch (err) {
        console.warn('Initial formula computing setup warning:', err);
      }

      // Initialize the workbook snapshot directly
      api.createUniverSheet(initialWorkbookData);
      setUniverAPI(api);

      // Register Official SheetInterceptorService for seamless formula visual continuity
      try {
        const injector = (univer as any)?.__injector;
        if (injector) {
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

      // Register mutation listener for copy-through dirty tracking and formula visual lifecycle
      try {
        if (api.onCommandExecuted) {
          api.onCommandExecuted((commandInfo: any, options: any) => {
            const commandId = commandInfo?.id;
            const params = commandInfo?.params;
            const subUnitId = params?.subUnitId || params?.sheetId;
            const unitId = params?.unitId || api.getActiveWorkbook?.()?.getId?.() || '';

            // Dirty sheet tracking for export: ONLY on actual cell mutations / structural changes
            if (
              (commandId === 'sheet.mutation.set-range-values' ||
                commandId === 'sheet.mutation.remove-rows' ||
                commandId === 'sheet.mutation.insert-row' ||
                commandId === 'sheet.mutation.remove-col' ||
                commandId === 'sheet.mutation.insert-col') &&
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

            // Capture previous value before formula invalidation
            if (
              commandId === 'sheet.mutation.set-range-values' &&
              !options?.fromFormula &&
              !options?.applyFormulaCalculationResult
            ) {
              const cellValue = params?.cellValue;
              if (cellValue && typeof cellValue === 'object') {
                const workbook = api.getActiveWorkbook?.();
                const worksheet = subUnitId ? workbook?.getSheetBySheetId?.(subUnitId) : workbook?.getActiveSheet?.();
                const matrix = (worksheet as any)?.getSheet?.()?.getCellMatrix?.() || (worksheet as any)?.getCellMatrix?.();

                for (const rStr of Object.keys(cellValue)) {
                  const r = parseInt(rStr, 10);
                  const rowObj = cellValue[r];
                  if (rowObj && typeof rowObj === 'object') {
                    for (const cStr of Object.keys(rowObj)) {
                      const c = parseInt(cStr, 10);
                      const cellItem = rowObj[c];
                      if (cellItem?.f && typeof cellItem.f === 'string' && cellItem.f.startsWith('=')) {
                        const existingCell = matrix?.getValue?.(r, c) || (worksheet as any)?.getCellRaw?.(r, c);
                        const prevVal = existingCell?.v;
                        if (prevVal !== undefined && prevVal !== null) {
                          const key = `${unitId}_${subUnitId || worksheet?.getSheetId?.()}_${r}_${c}`;
                          pendingFormulaDisplayMap.set(key, { previousValue: prevVal });
                        }
                      }
                    }
                  }
                }
              }
            }

            // Clear pending visual entry when authoritative calculation result arrives
            if (
              commandId === 'sheet.mutation.set-range-values' &&
              (options?.applyFormulaCalculationResult || options?.fromFormula)
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