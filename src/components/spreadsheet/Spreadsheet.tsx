"use client";

import React, { useEffect, useRef, useState, useCallback } from "react";
import { createUniver, LocaleType, mergeLocales } from "@univerjs/presets";
import { UniverSheetsCorePreset } from "@univerjs/preset-sheets-core";
import UniverPresetSheetsCoreEnUS from "@univerjs/preset-sheets-core/locales/en-US";
import { UniverSheetsSortPreset } from "@univerjs/preset-sheets-sort";
import { UniverSheetsFilterPreset } from "@univerjs/preset-sheets-filter";
import UniverPresetSheetsSortEnUS from "@univerjs/preset-sheets-sort/locales/en-US";
import UniverPresetSheetsFilterEnUS from "@univerjs/preset-sheets-filter/locales/en-US";
import {
  SheetInterceptorService,
  INTERCEPTOR_POINT,
  SetRangeValuesMutation,
} from "@univerjs/sheets";
import { InterceptorEffectEnum, ICommandService } from "@univerjs/core";
import { IActiveDirtyManagerService } from "@univerjs/engine-formula";
import "@univerjs/presets/lib/styles/preset-sheets-core.css";
import "@univerjs/preset-sheets-core/lib/index.css";
import "@univerjs/preset-sheets-sort/lib/index.css";
import "@univerjs/preset-sheets-filter/lib/index.css";

import { initialWorkbookData } from "@/lib/univer/univerConfig";
import { workbookSession } from "@/lib/session/workbookSession";
import { CalculationScheduler } from "@/lib/workers/pool/calculationScheduler";
import {
  installUniverCalculationPassivation,
  applyCalculationPatchSilent,
} from "@/lib/univer/univerPassivation";
import {
  evaluateMicroFormula,
  extractReferencedCells,
} from "@/lib/univer/microFormulaEvaluator";
import { DatasetReader } from "@/lib/storage/opfs/datasetReader";
import { ViewportController } from "@/lib/viewport/viewportController";
import { DatasetViewportController } from "@/lib/viewport/datasetViewportController";
import { UniverViewportAdapter } from "@/lib/viewport/univerViewportAdapter";
import SpreadsheetToolbar from "./SpreadsheetToolbar";
import JumpToRowModal from "./JumpToRowModal";
import DatasetQueryModal from "./DatasetQueryModal";
import { QueryDefinition } from "@/lib/largeCsv/query/types";
import { getCommittedDatasetStatus } from "@/lib/import/csvLargeDataset";
import { MetadataStore } from "@/lib/storage/opfs/metadataStore";
import { replaceUniverWorkbook } from "@/lib/import/excel";

export default function Spreadsheet() {
  const containerRef = useRef<HTMLDivElement>(null);
  const univerRef = useRef<any>(null);
  const schedulerRef = useRef<CalculationScheduler | null>(null);
  const formulaRegistryRef = useRef<
    Map<string, { sheetId: string; row: number; col: number; formula: string }>
  >(new Map());
  const viewportControllerRef = useRef<ViewportController | null>(null);
  const datasetViewportControllerRef = useRef<DatasetViewportController | null>(
    null,
  );
  const datasetReaderRef = useRef<DatasetReader | null>(null);
  const [univerAPI, setUniverAPI] = useState<any>(null);

  const [activeFileName, setActiveFileName] = useState("workbook-initial.xlsx");
  const [activeFileFormat, setActiveFileFormat] = useState<"xlsx" | "csv">(
    "xlsx",
  );
  const [isJumpModalOpen, setIsJumpModalOpen] = useState(false);
  const [datasetRowCount, setDatasetRowCount] = useState<number>(0);
  const [activeDatasetId, setActiveDatasetId] = useState<string | null>(null);
  const [activeQuery, setActiveQuery] = useState<{
    queryId: string;
    matchedRows: number;
    queryDef: QueryDefinition;
  } | null>(null);
  const [isQueryModalOpen, setIsQueryModalOpen] = useState(false);
  const [isQueryLoading, setIsQueryLoading] = useState(false);
  const queryTokenRef = useRef(0);
  const activeQueryAbortRef = useRef<AbortController | null>(null);

  const mountOPFSDataset = useCallback(async (meta?: any) => {
    const api = univerRef.current?.univerAPI;
    if (!api) return;

    let reader = datasetReaderRef.current;
    if (!reader) {
      reader = new DatasetReader();
      datasetReaderRef.current = reader;
    }

    const ready = await reader.init();
    if (!ready) {
      console.warn(
        "[Spreadsheet] OPFS dataset init failed:",
        reader.getInitError(),
      );
      return;
    }

    const metadata = meta || reader.getMetadata();
    if (!metadata || metadata.status !== "COMMITTED") return;

    const adapter = new UniverViewportAdapter(api);
    let controller = viewportControllerRef.current;
    if (!controller) {
      controller = new ViewportController(reader, adapter);
      viewportControllerRef.current = controller;
    } else {
      controller.setAdapter(adapter);
    }

    const success = await controller.mount(metadata);
    if (success) {
      setDatasetRowCount(metadata.rowCount);
      setActiveFileName(metadata.fileName);
      setActiveFileFormat("csv");
      (window as any).__viewportController = controller;
      (window as any).__datasetReader = reader;
      console.log(
        `[Spreadsheet] Successfully mounted OPFS dataset: ${metadata.fileName} (${metadata.rowCount} rows)`,
      );
    } else {
      console.warn(
        "[Spreadsheet] controller.mount returned false for:",
        metadata.fileName,
      );
    }
  }, []);

  const handleFileImported = useCallback(
    async (filename: string, format: "xlsx" | "csv", metadata?: any) => {
      setActiveFileName(filename);
      setActiveFileFormat(format);
      if (format === "csv") {
        await mountOPFSDataset(metadata);
      } else {
        setDatasetRowCount(0);
        if (viewportControllerRef.current) {
          viewportControllerRef.current.clearUserEdits();
        }
      }
    },
    [mountOPFSDataset],
  );

  const handleNewSheet = useCallback(async () => {
    const api = univerRef.current?.univerAPI;
    if (!api) return;

    try {
      await MetadataStore.clearStorage();
    } catch (e) {
      console.warn("[Spreadsheet] Clear storage notice:", e);
    }

    if (viewportControllerRef.current) {
      viewportControllerRef.current.clearUserEdits();
      viewportControllerRef.current = null;
    }
    if (datasetViewportControllerRef.current) {
      datasetViewportControllerRef.current = null;
    }
    (window as any).__viewportController = null;
    (window as any).__datasetViewportController = null;
    (window as any).__datasetReader = null;

    setActiveFileName("workbook-initial.xlsx");
    setActiveFileFormat("xlsx");
    setDatasetRowCount(0);
    setActiveDatasetId(null);
    setActiveQuery(null);

    try {
      replaceUniverWorkbook(api, initialWorkbookData);
    } catch (err) {
      console.warn("[Spreadsheet] Replace workbook notice:", err);
    }

    if (typeof window !== "undefined") {
      const url = new URL(window.location.href);
      url.searchParams.delete("datasetId");
      url.searchParams.delete("queryId");
      window.history.replaceState(null, "", url.pathname);
    }
  }, []);

  const mountServerDataset = useCallback(
    async (datasetId: string, initialQueryId?: string) => {
      const api = univerRef.current?.univerAPI;
      if (!api) return;

      const adapter = new UniverViewportAdapter(api);
      let controller = datasetViewportControllerRef.current;
      if (!controller) {
        controller = new DatasetViewportController(adapter);
        datasetViewportControllerRef.current = controller;
      } else {
        controller.setAdapter(adapter);
      }

      const success = await controller.mount(datasetId);
      if (success) {
        setActiveDatasetId(datasetId);
        let effectiveRows = controller.getTotalRows();
        setActiveFileName(`dataset-${datasetId}.csv`);
        setActiveFileFormat("csv");
        (window as any).__datasetViewportController = controller;

        if (initialQueryId) {
          try {
            const qRes = await fetch(
              `/api/v1/datasets/${datasetId}/query/${initialQueryId}`,
            );
            if (qRes.ok) {
              const qMeta = await qRes.json();
              await controller.setQuery(initialQueryId, qMeta.matchedRows);
              effectiveRows = qMeta.matchedRows;
              setActiveQuery({
                queryId: initialQueryId,
                matchedRows: qMeta.matchedRows,
                queryDef: qMeta.queryDef,
              });
            } else {
              const url = new URL(window.location.href);
              url.searchParams.delete("queryId");
              window.history.replaceState(null, "", url.pathname + url.search);
            }
          } catch {
            const url = new URL(window.location.href);
            url.searchParams.delete("queryId");
            window.history.replaceState(null, "", url.pathname + url.search);
          }
        } else {
          setActiveQuery(null);
        }

        setDatasetRowCount(effectiveRows);
        console.log(
          `[Spreadsheet] Successfully mounted server dataset: ${datasetId} (${effectiveRows} rows)`,
        );
      }
    },
    [],
  );

  const handleApplyServerQuery = useCallback(
    async (queryDef: QueryDefinition) => {
      if (!activeDatasetId) return;

      const token = ++queryTokenRef.current;
      if (activeQueryAbortRef.current) {
        activeQueryAbortRef.current.abort();
      }
      const abortController = new AbortController();
      activeQueryAbortRef.current = abortController;

      setIsQueryLoading(true);

      try {
        const res = await fetch(`/api/v1/datasets/${activeDatasetId}/query`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(queryDef),
          signal: abortController.signal,
        });

        if (!res.ok) {
          const errData = await res.json().catch(() => ({}));
          throw new Error(errData.error || `Query failed: HTTP ${res.status}`);
        }

        const meta = await res.json();
        if (token !== queryTokenRef.current) {
          // Stale query response dropped
          return;
        }

        const controller = datasetViewportControllerRef.current;
        if (controller) {
          await controller.setQuery(meta.queryId, meta.matchedRows);
          setDatasetRowCount(meta.matchedRows);
        }

        setActiveQuery({
          queryId: meta.queryId,
          matchedRows: meta.matchedRows,
          queryDef,
        });

        // Update URL queryId without full reload
        if (typeof window !== "undefined") {
          const url = new URL(window.location.href);
          url.searchParams.set("datasetId", activeDatasetId);
          url.searchParams.set("queryId", meta.queryId);
          window.history.replaceState(null, "", url.pathname + url.search);
        }
      } finally {
        if (token === queryTokenRef.current) {
          setIsQueryLoading(false);
          activeQueryAbortRef.current = null;
        }
      }
    },
    [activeDatasetId],
  );

  const handleClearServerQuery = useCallback(async () => {
    const token = ++queryTokenRef.current;
    if (activeQueryAbortRef.current) {
      activeQueryAbortRef.current.abort();
      activeQueryAbortRef.current = null;
    }

    setIsQueryLoading(false);

    const controller = datasetViewportControllerRef.current;
    if (controller) {
      await controller.clearQuery();
      const baseRows = controller.getBaseTotalRows();
      setDatasetRowCount(baseRows);
    }

    setActiveQuery(null);

    if (activeDatasetId && typeof window !== "undefined") {
      const url = new URL(window.location.href);
      url.searchParams.set("datasetId", activeDatasetId);
      url.searchParams.delete("queryId");
      window.history.replaceState(null, "", url.pathname + url.search);
    }
  }, [activeDatasetId]);

  const getAvailableHeaders = useCallback((): string[] => {
    if (datasetViewportControllerRef.current) {
      const h = datasetViewportControllerRef.current.getHeaders();
      if (h && h.length > 0) return h;
    }
    if (datasetReaderRef.current) {
      const h = datasetReaderRef.current.getHeaders();
      if (h && h.length > 0) return h;
    }
    if (viewportControllerRef.current) {
      const h = viewportControllerRef.current.getMetadata()?.headers;
      if (h && h.length > 0) return h;
    }
    return [];
  }, []);

  const handleJumpToRow = useCallback(async (rowNum: number) => {
    if (datasetViewportControllerRef.current?.isReady()) {
      await datasetViewportControllerRef.current.jumpToRow(rowNum);
    } else if (viewportControllerRef.current?.isReady?.()) {
      await viewportControllerRef.current.jumpToRow(rowNum);
    } else if (univerRef.current?.univerAPI) {
      try {
        const sheet = univerRef.current.univerAPI
          .getActiveWorkbook()
          ?.getActiveSheet();
        if (sheet?.scrollToCell) {
          sheet.scrollToCell(rowNum - 1, 0);
        } else if (sheet?.getRange) {
          const range = sheet.getRange(rowNum - 1, 0);
          sheet.setActiveRange?.(range);
        }
      } catch (err) {
        console.warn("[handleJumpToRow] Local sheet scroll notice:", err);
      }
    }
  }, []);

  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "g") {
        e.preventDefault();
        setIsJumpModalOpen(true);
      }
    };
    window.addEventListener("keydown", handleKeyDown);
    return () => window.removeEventListener("keydown", handleKeyDown);
  }, []);

  useEffect(() => {
    if (typeof window !== "undefined" && univerAPI) {
      (window as any).__mountOPFSDataset = mountOPFSDataset;
      (window as any).__handleNewSheet = handleNewSheet;
      const params = new URLSearchParams(window.location.search);
      const datasetIdParam = params.get("datasetId");
      const queryIdParam = params.get("queryId") || undefined;
      if (datasetIdParam) {
        mountServerDataset(datasetIdParam, queryIdParam);
      } else {
        // Default clean state on initial launch: clear residual OPFS storage so spreadsheet is empty
        MetadataStore.clearStorage().catch((err) => {
          console.warn("[Spreadsheet] OPFS initial clear notice:", err);
        });
      }
    }
  }, [univerAPI, mountServerDataset, mountOPFSDataset, handleNewSheet]);

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
        containerRef.current.innerHTML = "";
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
            UniverPresetSheetsFilterEnUS,
          ),
        },
        presets: [corePreset, sortPreset, filterPreset],
      });

      currentRefObj = { univer, univerAPI: api };
      univerRef.current = currentRefObj;
      (window as any).univerAPI = api;
      (window as any).univer = univer;
      (window as any).mountServerDataset = mountServerDataset;

      // Register SheetInterceptorService and Enterprise Calculation Passivation BEFORE sheet creation
      try {
        const injector =
          (api as any)?._injector ||
          (univer as any)?.__injector ||
          (univer as any)?._injector;
        if (injector) {
          // Enterprise Calculation Passivation: Completely suppresses browser-side formula recalculation storms
          try {
            installUniverCalculationPassivation(injector);
          } catch (passivationErr) {
            console.warn(
              "[UNIVER_PASSIVATION] Error installing calculation passivation:",
              passivationErr,
            );
          }

          const sheetInterceptorService = injector.get(SheetInterceptorService);
          if (sheetInterceptorService && INTERCEPTOR_POINT?.CELL_CONTENT) {
            interceptorDisposable = sheetInterceptorService.intercept(
              INTERCEPTOR_POINT.CELL_CONTENT,
              {
                priority: 100,
                effect: InterceptorEffectEnum.Value,
                handler(cell: any, context: any, next: any) {
                  if (
                    cell &&
                    (cell.f || context?.rawData?.f) &&
                    (cell.v === null || cell.v === undefined)
                  ) {
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
              },
            );
          }
        }
      } catch (interceptErr) {
        console.warn(
          "SheetInterceptorService registration warning:",
          interceptErr,
        );
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
                const range = worksheet.getRange?.(
                  result.targetCell.row,
                  result.targetCell.col,
                );
                if (range) {
                  range.setValue(result.value);
                }
              }
            } catch (applyErr) {
              console.warn(
                "[CalculationScheduler] Failed to apply result:",
                applyErr,
              );
            }
          },
        },
        4, // 4 Dedicated Workers
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
        const cmdSvc =
          (api as any)?._injector?.get(ICommandService) ||
          (api as any)?.getCommandService?.();
        if (cmdSvc?.beforeCommandExecuted) {
          cmdSvc.beforeCommandExecuted((command: any) => {
            if (
              command?.id === "sheet.command.sort-range" ||
              command?.id === "sheet.command.reorder-range"
            ) {
              isSortReorderActive = true;
            }
          });
        }
      } catch (e) {
        console.warn("Command service listener setup notice:", e);
      }

      // Register mutation listener for copy-through dirty tracking and formula visual lifecycle
      try {
        if (api.onCommandExecuted) {
          api.onCommandExecuted((commandInfo: any, options: any) => {
            const commandId = commandInfo?.id || "unknown";
            const params = commandInfo?.params;
            const subUnitId = params?.subUnitId || params?.sheetId;
            const unitId =
              params?.unitId || api.getActiveWorkbook?.()?.getId?.() || "";

            diagState.commandCounts[commandId] =
              (diagState.commandCounts[commandId] || 0) + 1;

            if (
              commandId === "sheet.command.sort-range" ||
              commandId === "sheet.command.reorder-range"
            ) {
              isSortReorderActive = false;
            }

            if (
              commandId.includes("sort") ||
              commandId.includes("filter") ||
              commandId.includes("reorder") ||
              commandId.includes("calc") ||
              commandId.includes("formula") ||
              commandId.includes("dirty")
            ) {
              console.log(`[DIAG_COMMAND] ${commandId}`, {
                params,
                options,
                count: diagState.commandCounts[commandId],
              });
            }

            if (
              commandId === "formula.mutation.set-formula-calculation-start"
            ) {
              diagState.calculationRequests++;
              console.warn(
                `[DIAG_CALCULATION_STORM] SetFormulaCalculationStartMutation FIRED! (count=${diagState.calculationRequests})`,
                params,
              );
            }

            // Rapid Sheet Switching Protection: Update active sheet & cancel obsolete background jobs
            if (commandId === "sheet.operation.set-worksheet-active") {
              const activeSub = params?.subUnitId || params?.sheetId;
              if (activeSub) {
                if (datasetViewportControllerRef.current?.isReady()) {
                  if (
                    datasetViewportControllerRef.current.getActiveSheetPartition()
                      ?.sheetId !== activeSub
                  ) {
                    datasetViewportControllerRef.current.setActiveSheet(
                      activeSub,
                    );
                  }
                } else if (viewportControllerRef.current?.isReady?.()) {
                  if (
                    viewportControllerRef.current.getActiveSheetPartition()
                      ?.sheetId !== activeSub
                  ) {
                    viewportControllerRef.current.setActiveSheet(activeSub);
                  }
                }
                if (schedulerRef.current) {
                  schedulerRef.current.onSheetActivated(activeSub);
                }
              }
            }

            // Architecture D: Viewport sliding on user scrolling
            if (commandId === "sheet.operation.set-scroll") {
              const startRow = params?.sheetViewStartRow;
              if (typeof startRow === "number") {
                if (datasetViewportControllerRef.current?.isReady()) {
                  datasetViewportControllerRef.current.handleScroll(startRow);
                } else if (viewportControllerRef.current?.isReady?.()) {
                  viewportControllerRef.current.handleScroll(startRow);
                }
              }
            }

            // Dirty sheet tracking for export: ONLY on actual cell mutations / structural changes
            if (
              (commandId === "sheet.mutation.set-range-values" ||
                commandId === "sheet.mutation.remove-rows" ||
                commandId === "sheet.mutation.insert-row" ||
                commandId === "sheet.mutation.remove-col" ||
                commandId === "sheet.mutation.insert-col" ||
                commandId === "sheet.mutation.reorder-range") &&
              !options?.fromFormula &&
              !options?.applyFormulaCalculationResult
            ) {
              if (subUnitId) {
                workbookSession.markSheetDirty(subUnitId);
                const activeWorkbook = api.getActiveWorkbook?.();
                const fWs = activeWorkbook?.getSheetBySheetId?.(subUnitId);
                const sName = fWs?.getSheetName?.() || (fWs as any)?.getName?.();
                if (sName) workbookSession.markSheetDirty(sName);
              } else {
                const activeSheet = api
                  .getActiveWorkbook?.()
                  ?.getActiveSheet?.();
                if (activeSheet) {
                  const sId = activeSheet.getSheetId?.();
                  const sName =
                    activeSheet.getSheetName?.() || activeSheet.getSheetId?.();
                  if (sId) workbookSession.markSheetDirty(sId);
                  if (sName) workbookSession.markSheetDirty(sName);
                }
              }
            }

            // Capture cell deltas and dispatch to backend calculation engine (ONLY for genuine edits, NOT sort row reorders or silent patches)
            if (
              commandId === "sheet.mutation.set-range-values" &&
              !options?.fromFormula &&
              !options?.applyFormulaCalculationResult &&
              !options?.fromCalculationPatch &&
              !options?.onlyLocal &&
              !isSortReorderActive &&
              !(window as any).__isSortActive
            ) {
              const cellValue = params?.cellValue;
              if (cellValue && typeof cellValue === "object") {
                const workbook = api.getActiveWorkbook?.();
                const worksheet =
                  (workbook?.getSheets
                    ? workbook
                        .getSheets()
                        .find((s: any) => s.getSheetId?.() === subUnitId)
                    : null) ||
                  (workbook as any)
                    ?.getWorkbook?.()
                    ?.getSheetBySheetId?.(subUnitId) ||
                  workbook?.getActiveSheet?.();
                const targetSheetId =
                  subUnitId || worksheet?.getSheetId?.() || "sheet-1";
                const matrix =
                  (worksheet as any)?.getSheet?.()?.getCellMatrix?.() ||
                  (worksheet as any)?.getCellMatrix?.();
                const deltas: Array<{
                  sheetId: string;
                  row: number;
                  col: number;
                  v?: any;
                  f?: string;
                }> = [];

                const lookupVal = (
                  sheetName: string | undefined,
                  rIdx: number,
                  cIdx: number,
                ) => {
                  let targetWs = worksheet;
                  if (sheetName) {
                    const cleanName = sheetName
                      .replace(/^['"]|['"]$/g, "")
                      .trim()
                      .toLowerCase();
                    targetWs =
                      (workbook?.getSheets
                        ? workbook.getSheets().find((s: any) => {
                            const sName = (
                              s.getSheetName?.() ||
                              s.getName?.() ||
                              s.getSheetId?.() ||
                              ""
                            ).toLowerCase();
                            return sName === cleanName;
                          })
                        : null) || worksheet;
                  }
                  const mat =
                    (targetWs as any)?.getSheet?.()?.getCellMatrix?.() ||
                    (targetWs as any)?.getCellMatrix?.();
                  const cell =
                    mat?.getValue?.(rIdx, cIdx) ||
                    (targetWs as any)?.getCellRaw?.(rIdx, cIdx);
                  return cell?.v;
                };

                const immediatePatches: Array<{
                  sheetId: string;
                  row: number;
                  col: number;
                  v: any;
                  f?: string;
                  t?: number;
                }> = [];

                const recordEdit = (
                  sId: string,
                  rowIdx: number,
                  colIdx: number,
                  val: any,
                  form?: string,
                ) => {
                  if (datasetViewportControllerRef.current?.isReady?.()) {
                    datasetViewportControllerRef.current.recordUserEdit(
                      sId,
                      rowIdx,
                      colIdx,
                      val,
                      form,
                    );
                  } else if (viewportControllerRef.current) {
                    viewportControllerRef.current.recordUserEdit(
                      sId,
                      rowIdx,
                      colIdx,
                      val,
                      form,
                    );
                  }
                };

                const changedCells = new Map<string, any>();

                for (const rStr of Object.keys(cellValue)) {
                  const r = parseInt(rStr, 10);
                  const rowObj = cellValue[r];
                  if (rowObj && typeof rowObj === "object") {
                    for (const cStr of Object.keys(rowObj)) {
                      const c = parseInt(cStr, 10);
                      const cellItem = rowObj[c];
                      if (cellItem === undefined) continue;

                      // Extract value and formula while preserving strict semantic distinction:
                      // undefined = no edit
                      // null = explicitly cleared cell
                      // "" = explicitly empty string
                      // 0, false, strings, numbers = valid edits
                      let editValue: any;
                      let editFormula: string | undefined = undefined;

                      if (cellItem === null) {
                        editValue = null;
                      } else if (typeof cellItem === "object") {
                        if (cellItem.v === null) {
                          editValue = null;
                        } else if (cellItem.v !== undefined) {
                          editValue = cellItem.v;
                        } else if (cellItem.f !== undefined) {
                          editValue = null;
                        } else {
                          editValue = null;
                        }
                        editFormula = cellItem.f;
                      } else {
                        editValue = cellItem;
                      }

                      // Handle formula input: evaluate immediately with zero UI freeze
                      if (
                        editFormula &&
                        typeof editFormula === "string" &&
                        editFormula.startsWith("=")
                      ) {
                        formulaRegistryRef.current.set(
                          `${targetSheetId}_${r}_${c}`,
                          {
                            sheetId: targetSheetId,
                            row: r,
                            col: c,
                            formula: editFormula,
                          },
                        );

                        const computed = evaluateMicroFormula(
                          editFormula,
                          lookupVal,
                        );
                        if (computed !== null && computed !== undefined) {
                          editValue = computed;
                          immediatePatches.push({
                            sheetId: targetSheetId,
                            row: r,
                            col: c,
                            f: editFormula,
                            v: computed,
                            t: 2,
                          });
                        }

                        // Record user edit with computed formula result so it survives eviction cycles
                        recordEdit(targetSheetId, r, c, editValue, editFormula);
                        changedCells.set(`${targetSheetId}_${r}_${c}`, editValue);

                        // Extract precedent cells so backend WorkbookEngine and BullMQ receive their values
                        const refs = extractReferencedCells(editFormula);
                        for (const ref of refs) {
                          const refVal = lookupVal(
                            ref.sheetName,
                            ref.row,
                            ref.col,
                          );
                          if (refVal !== undefined && refVal !== null) {
                            deltas.push({
                              sheetId: targetSheetId,
                              row: ref.row,
                              col: ref.col,
                              v: refVal,
                            });
                          }
                        }
                      } else {
                        // Cell was edited without a formula, or explicitly cleared:
                        // Clean up existing formula registration if any
                        formulaRegistryRef.current.delete(
                          `${targetSheetId}_${r}_${c}`,
                        );

                        // Record user edit in ViewportController (null for clear, or raw value)
                        recordEdit(targetSheetId, r, c, editValue, undefined);
                        if (editValue !== undefined) {
                          changedCells.set(
                            `${targetSheetId}_${r}_${c}`,
                            editValue,
                          );
                        }
                      }

                      deltas.push({
                        sheetId: targetSheetId,
                        row: r,
                        col: c,
                        v: editValue,
                        f: editFormula,
                      });
                    }
                  }
                }

                // Check cascading dependent formulas for all changed cells in this batch
                if (
                  changedCells.size > 0 &&
                  formulaRegistryRef.current.size > 0
                ) {
                  let hasChanges = true;
                  let iterations = 0;
                  const localValLookup = (
                    s: string | undefined,
                    ro: number,
                    co: number,
                  ) => {
                    let sheetKey = targetSheetId;
                    if (s) {
                      const cleanName = s
                        .replace(/^['"]|['"]$/g, "")
                        .trim()
                        .toLowerCase();
                      const targetWs =
                        (workbook?.getSheets
                          ? workbook.getSheets().find((ws: any) => {
                              const sName = (
                                ws.getSheetName?.() ||
                                ws.getName?.() ||
                                ws.getSheetId?.() ||
                                ""
                              ).toLowerCase();
                              return sName === cleanName;
                            })
                          : null) || worksheet;
                      sheetKey =
                        (targetWs as any)?.getSheetId?.() ||
                        targetWs?.id ||
                        targetSheetId;
                    }
                    const k = `${sheetKey}_${ro}_${co}`;
                    if (changedCells.has(k)) {
                      return changedCells.get(k);
                    }
                    return lookupVal(s, ro, co);
                  };

                  while (hasChanges && iterations < 10) {
                    hasChanges = false;
                    iterations++;

                    for (const [
                      ,
                      entry,
                    ] of formulaRegistryRef.current.entries()) {
                      if (entry.sheetId === targetSheetId) {
                        const entryKey = `${entry.sheetId}_${entry.row}_${entry.col}`;
                        const refs = extractReferencedCells(entry.formula);
                        const isDep = refs.some((rf) => {
                          const rfSheet = rf.sheetName || entry.sheetId;
                          return changedCells.has(
                            `${rfSheet}_${rf.row}_${rf.col}`,
                          );
                        });

                        if (isDep) {
                          const recomputed = evaluateMicroFormula(
                            entry.formula,
                            localValLookup,
                          );
                          if (
                            recomputed !== null &&
                            recomputed !== undefined
                          ) {
                            const prevVal = changedCells.get(entryKey);
                            if (prevVal !== recomputed) {
                              changedCells.set(entryKey, recomputed);
                              hasChanges = true;

                              // Record in ViewportController so CSV & eviction overlay have the latest calculated value!
                              recordEdit(
                                entry.sheetId,
                                entry.row,
                                entry.col,
                                recomputed,
                                entry.formula,
                              );

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
                  }
                }

                // Immediately apply instant patches to Univer with zero UI freeze
                if (immediatePatches.length > 0) {
                  applyCalculationPatchSilent(api, unitId, immediatePatches);
                }

                if (deltas.length > 0) {
                  diagState.workerScheduledCount++;
                  console.log(
                    `[BACKEND_CALCULATION] Dispatching ${deltas.length} delta(s) to backend API (batch #${diagState.workerScheduledCount})`,
                  );

                  // Asynchronously send to Next.js PostgreSQL WAL & Backend Calculation Engine
                  fetch(`/api/v1/workbooks/${unitId}/mutations`, {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ deltas }),
                  })
                    .then((res) => res.json())
                    .then((resData) => {
                      if (
                        resData?.patches &&
                        Array.isArray(resData.patches) &&
                        resData.patches.length > 0
                      ) {
                        const validPatches = resData.patches.filter(
                          (p: any) =>
                            p.v !== null &&
                            p.v !== undefined &&
                            p.v !== "" &&
                            !String(p.v).includes("#"),
                        );
                        if (validPatches.length > 0) {
                          console.log(
                            `[BACKEND_CALCULATION] Applying ${validPatches.length} patch(es) from backend engine (exec=${resData.executionMs}ms)`,
                          );
                          applyCalculationPatchSilent(
                            api,
                            unitId,
                            validPatches,
                          );
                        }
                      }
                    })
                    .catch((err) => {
                      console.warn(
                        "[BACKEND_CALCULATION] Mutation sync notice:",
                        err,
                      );
                    });
                }
              }
            }

            // Clear pending visual entry when authoritative calculation result arrives
            if (
              commandId === "sheet.mutation.set-range-values" &&
              (options?.applyFormulaCalculationResult ||
                options?.fromFormula ||
                options?.fromCalculationPatch)
            ) {
              const cellValue = params?.cellValue;
              if (cellValue && typeof cellValue === "object") {
                for (const rStr of Object.keys(cellValue)) {
                  const r = parseInt(rStr, 10);
                  const rowObj = cellValue[r];
                  if (rowObj && typeof rowObj === "object") {
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
        console.warn("Univer command tracking setup warning:", cmdErr);
      }

      setTimeout(() => {
        if (typeof window !== "undefined") {
          window.dispatchEvent(new Event("resize"));
        }
      }, 100);
    } catch (e) {
      console.error("Failed to initialize Univer:", e);
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
            console.warn("Dispose warning:", e);
          }
        }
      }, 300);
    };
  }, []);

  return (
    <div className="univer-wrapper">
      <SpreadsheetToolbar
        univerAPI={univerAPI}
        activeFileName={activeFileName}
        activeFileFormat={activeFileFormat}
        onFileImported={handleFileImported}
        onNewSheet={handleNewSheet}
        onOpenJumpToRow={() => setIsJumpModalOpen(true)}
        datasetRowCount={datasetRowCount}
        datasetId={activeDatasetId}
        onOpenQueryModal={() => setIsQueryModalOpen(true)}
        activeQuery={activeQuery}
        onClearQuery={handleClearServerQuery}
        isQueryLoading={isQueryLoading}
      />
      <JumpToRowModal
        isOpen={isJumpModalOpen}
        onClose={() => setIsJumpModalOpen(false)}
        maxRow={datasetRowCount > 0 ? datasetRowCount : 1000}
        onJump={handleJumpToRow}
      />
      <DatasetQueryModal
        isOpen={isQueryModalOpen}
        onClose={() => setIsQueryModalOpen(false)}
        headers={getAvailableHeaders()}
        datasetId={activeDatasetId || ""}
        datasetRowCount={datasetRowCount}
        activeQueryDef={activeQuery?.queryDef}
        onApplyQuery={handleApplyServerQuery}
        isLoading={isQueryLoading}
      />
      <div className="univer-container-outer">
        <div ref={containerRef} className="univer-container-inner" />
      </div>
    </div>
  );
}
