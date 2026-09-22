import {
  SetRangeValuesMutation,
  SheetsSelectionsService,
  DeltaRowHeightCommand,
  SetRowHeightCommand,
  COMMAND_LISTENER_SKELETON_CHANGE,
  getSheetCommandTarget,
  SheetRangeThemeModel,
} from "@univerjs/sheets";
import { IActiveDirtyManagerService } from "@univerjs/engine-formula";
import {
  SheetSkeleton,
  ICommandService,
  IUndoRedoService,
  IUniverInstanceService,
  CommandType,
  RANGE_TYPE,
} from "@univerjs/core";
import { SheetsFilterService } from "@univerjs/sheets-filter";

// ==============================================================================
// UNIVER PASSIVATION & CALCULATION SUPPRESSION (PHASE 4)
// Neutralizes browser-side full-workbook recalculation storm on cell edits
// while preserving editing, rendering, formatting, selection, and formula bar display.
// ==============================================================================

export interface PassivationController {
  isPassivated: boolean;
  restore: () => void;
}

/**
 * Passivates Univer's client-side calculation trigger on the browser UI thread.
 * Guarantees that user cell edits (SetRangeValuesMutation) NEVER trigger
 * SetFormulaCalculationStartMutation or full-sheet synchronous formula walks.
 */
export function installUniverCalculationPassivation(
  injector: any,
): PassivationController {
  if (!injector) {
    console.warn(
      "[UNIVER_PASSIVATION] Injector not provided, skipping passivation",
    );
    return { isPassivated: false, restore: () => {} };
  }

  try {
    const activeDirtyManager = injector.get(IActiveDirtyManagerService);
    if (!activeDirtyManager) {
      console.warn("[UNIVER_PASSIVATION] IActiveDirtyManagerService not found");
      return { isPassivated: false, restore: () => {} };
    }

    // Save original SetRangeValuesMutation dirty configuration if present
    const originalSetRange = activeDirtyManager.get(SetRangeValuesMutation.id);

    const SUPPRESSED_MUTATIONS = [
      SetRangeValuesMutation.id,
      "sheet.mutation.set-filter-range",
      "sheet.mutation.set-filter-criteria",
      "sheet.mutation.remove-filter",
      "sheet.mutation.re-calc-filter",
    ];

    // Remove suppressed mutations from active dirty manager
    SUPPRESSED_MUTATIONS.forEach((cmdId) => activeDirtyManager.remove(cmdId));

    // Wrap get so suppressed mutations always return undefined to TriggerCalculationController
    const origGet = activeDirtyManager.get.bind(activeDirtyManager);
    activeDirtyManager.get = function (commandId: string) {
      if (SUPPRESSED_MUTATIONS.includes(commandId)) {
        return undefined;
      }
      return origGet(commandId);
    };

    // Wrap register so no subsequent plugin can re-register suppressed mutations
    const origRegister = activeDirtyManager.register.bind(activeDirtyManager);
    activeDirtyManager.register = function (
      commandId: string,
      dirtyConversion: any,
    ) {
      if (SUPPRESSED_MUTATIONS.includes(commandId)) {
        return; // Suppress registration completely
      }
      return origRegister(commandId, dirtyConversion);
    };

    console.log(
      "[UNIVER_PASSIVATION] Successfully passivated browser-side formula calculation engine (TriggerCalculationController disabled for edits and filters)",
    );

    return {
      isPassivated: true,
      restore: () => {
        if (originalSetRange) {
          activeDirtyManager.register(
            SetRangeValuesMutation.id,
            originalSetRange,
          );
        }
      },
    };
  } catch (err) {
    console.warn(
      "[UNIVER_PASSIVATION] Failed to install calculation passivation:",
      err,
    );
    return { isPassivated: false, restore: () => {} };
  }
}

/**
 * Injects authoritative calculation patches from backend into Univer CellMatrix
 * strictly using options: { onlyLocal: true } to guarantee zero recursive calculation storms.
 */
export function applyCalculationPatchSilent(
  univerAPI: any,
  unitId: string,
  patches: Array<{
    sheetId: string;
    row: number;
    col: number;
    v: any;
    f?: string;
    t?: number;
  }>,
): void {
  if (!univerAPI || !patches || patches.length === 0) return;

  try {
    const workbook = univerAPI.getActiveWorkbook?.();
    if (!workbook) return;

    // Group patches by worksheet for batched application
    const sheetPatchMap = new Map<
      string,
      Record<number, Record<number, any>>
    >();

    for (const p of patches) {
      if (!sheetPatchMap.has(p.sheetId)) {
        sheetPatchMap.set(p.sheetId, {});
      }
      const rowMap = sheetPatchMap.get(p.sheetId)!;
      if (!rowMap[p.row]) {
        rowMap[p.row] = {};
      }
      rowMap[p.row][p.col] = {
        v: p.v,
        ...(p.f ? { f: p.f } : {}),
        ...(p.t !== undefined ? { t: p.t } : {}),
      };
    }

    const cmdSvc =
      (univerAPI as any)?._commandService ||
      (univerAPI as any)?.getCommandService?.() ||
      (workbook as any)?._commandService;

    for (const [sheetId, cellValueMap] of sheetPatchMap.entries()) {
      const targetUnitId = workbook.getId?.() || unitId;
      const mutationParams = {
        unitId: targetUnitId,
        subUnitId: sheetId,
        cellValue: cellValueMap,
      };
      const mutationOptions = {
        onlyLocal: true,
        fromCalculationPatch: true,
      };

      if (cmdSvc?.syncExecuteCommand) {
        cmdSvc.syncExecuteCommand(
          SetRangeValuesMutation.id,
          mutationParams,
          mutationOptions,
        );
      } else if (cmdSvc?.executeCommand) {
        cmdSvc.executeCommand(
          SetRangeValuesMutation.id,
          mutationParams,
          mutationOptions,
        );
      } else if (univerAPI.executeCommand) {
        univerAPI.executeCommand(
          SetRangeValuesMutation.id,
          mutationParams,
          mutationOptions,
        );
      }
    }
  } catch (err) {
    console.warn(
      "[UNIVER_PASSIVATION] Failed to apply silent calculation patch:",
      err,
    );
  }
}

/**
 * Mutation identifier for updating worksheet default row height without 4M row object materialization.
 */
export const SetWorksheetDefaultRowHeightMutationId =
  "sheet.mutation.set-worksheet-default-row-height";

export interface ISetWorksheetDefaultRowHeightMutationParams {
  unitId: string;
  subUnitId: string;
  defaultRowHeight: number;
  clearedRowEntries?: Array<{ row: number; prevH?: number; prevIa?: number }>;
  isUndo?: boolean;
}

let isSkeletonOptimizationInstalled = false;
let isCommandInterceptionInstalled = false;

/**
 * Installs command interception for Select-All row resizing (DeltaRowHeightCommand & SetRowHeightCommand).
 * Eliminates 4,000,000 range object materialization and sparse rowData pollution.
 * Correctly preserves unrelated row metadata (hidden state hd, row styles s, and custom metadata).
 */
export function installRowHeightCommandInterception(injector: any): boolean {
  if (isCommandInterceptionInstalled || !injector) return true;

  try {
    let commandService: any = null;
    let univerInstanceService: any = null;
    let undoRedoService: any = null;

    try {
      commandService = injector.get(ICommandService);
      univerInstanceService = injector.get(IUniverInstanceService);
      undoRedoService = injector.get(IUndoRedoService);
    } catch {
      // Fallback service resolution
      const resolved =
        injector.resolvedDependencyCollection?.resolvedDependencies;
      if (resolved) {
        for (const [k] of resolved.entries()) {
          const name = k?.name || "";
          if (name.includes("CommandService") && !commandService)
            commandService = injector.get(k);
          if (name.includes("UniverInstanceService") && !univerInstanceService)
            univerInstanceService = injector.get(k);
          if (name.includes("UndoRedoService") && !undoRedoService)
            undoRedoService = injector.get(k);
        }
      }
    }

    if (!commandService || !univerInstanceService) return false;

    // 1. Register SetWorksheetDefaultRowHeightMutation
    if (!commandService.hasCommand(SetWorksheetDefaultRowHeightMutationId)) {
      commandService.registerCommand({
        id: SetWorksheetDefaultRowHeightMutationId,
        type: CommandType.MUTATION,
        handler: (
          accessor: any,
          params: ISetWorksheetDefaultRowHeightMutationParams,
        ) => {
          const target = getSheetCommandTarget(
            accessor.get(IUniverInstanceService),
            params,
          );
          if (!target) return false;
          const { worksheet } = target;
          const { defaultRowHeight, clearedRowEntries, isUndo } = params;

          // Update defaultRowHeight in worksheet configuration
          const config = worksheet.getConfig();
          config.defaultRowHeight = defaultRowHeight;

          // Preserve unrelated row metadata (hd, s, custom), cleanly manage rowData overrides
          const rowManager = worksheet.getRowManager();
          const rowData = rowManager.getRowData();
          if (rowData) {
            if (isUndo) {
              if (clearedRowEntries && clearedRowEntries.length > 0) {
                for (let i = 0; i < clearedRowEntries.length; i++) {
                  const entry = clearedRowEntries[i];
                  const row = rowManager.getRowOrCreate(entry.row);
                  if (entry.prevH !== undefined) row.h = entry.prevH;
                  if (entry.prevIa !== undefined) row.ia = entry.prevIa;
                }
              }
            } else {
              if (clearedRowEntries && clearedRowEntries.length > 0) {
                for (let i = 0; i < clearedRowEntries.length; i++) {
                  const entry = clearedRowEntries[i];
                  const row = rowData[entry.row];
                  if (row) {
                    delete row.h;
                    // Keep rowData sparse: delete key only if no other metadata (hd, s, custom, ah) exists
                    if (
                      row.hd === undefined &&
                      row.s === undefined &&
                      row.custom === undefined &&
                      row.ah === undefined
                    ) {
                      delete rowData[entry.row];
                    }
                  }
                }
              }
            }
          }

          return true;
        },
      });
    }

    // 2. Register mutation in COMMAND_LISTENER_SKELETON_CHANGE so Univer dirties SheetSkeleton automatically
    if (
      Array.isArray(COMMAND_LISTENER_SKELETON_CHANGE) &&
      !COMMAND_LISTENER_SKELETON_CHANGE.includes(
        SetWorksheetDefaultRowHeightMutationId,
      )
    ) {
      COMMAND_LISTENER_SKELETON_CHANGE.push(
        SetWorksheetDefaultRowHeightMutationId,
      );
    }

    let deltaWrapped = false;
    let setWrapped = false;

    // 3. Wrap DeltaRowHeightCommand
    if (commandService.hasCommand(DeltaRowHeightCommand.id)) {
      const origDeltaCmd = (commandService as any)._commandRegistry?.getCommand(
        DeltaRowHeightCommand.id,
      )?.[0];
      if (origDeltaCmd && !(origDeltaCmd as any).__isOptimizedWrapper) {
        commandService.unregisterCommand(DeltaRowHeightCommand.id);
        const wrappedDeltaCmd = {
          ...origDeltaCmd,
          id: DeltaRowHeightCommand.id,
          type: CommandType.COMMAND,
          __isOptimizedWrapper: true,
          handler: async (accessor: any, params: any) => {
            const selections = accessor
              .get(SheetsSelectionsService)
              ?.getCurrentSelections();
            if (!selections?.length) return false;

            const RANGE_ALL = (RANGE_TYPE as any)?.ALL ?? 3;
            const RANGE_ROW_TYPE = (RANGE_TYPE as any)?.ROW ?? 1;

            const isAllSheetRange =
              selections.length === 1 &&
              selections[0].range?.rangeType === RANGE_ALL;
            const rowSelections = selections.filter(
              (s: any) => s.range?.rangeType === RANGE_ROW_TYPE,
            );
            const { anchorRow, deltaY } = params;
            const rangeType = isAllSheetRange
              ? RANGE_ALL
              : rowSelections.some(
                    ({ range }: any) =>
                      range.startRow <= anchorRow && anchorRow <= range.endRow,
                  )
                ? RANGE_ROW_TYPE
                : ((RANGE_TYPE as any)?.NORMAL ?? 0);

            if (rangeType === RANGE_ALL) {
              const target = getSheetCommandTarget(
                accessor.get(IUniverInstanceService),
              );
              if (!target) return false;
              const { worksheet, unitId, subUnitId } = target;

              const prevDefaultRowHeight =
                worksheet.getConfig().defaultRowHeight;
              const currentHeight = worksheet.getRowHeight(anchorRow);
              const destRowHeight = Math.max(
                1,
                Math.min(2000, currentHeight + deltaY),
              );

              // Find any existing sparse rows with explicit h or ia
              const rowData = worksheet.getRowManager().getRowData() || {};
              const keys = Object.keys(rowData);
              const clearedRowEntries: Array<{
                row: number;
                prevH?: number;
                prevIa?: number;
              }> = [];
              for (let i = 0; i < keys.length; i++) {
                const r = Number(keys[i]);
                const item = rowData[r];
                if (item && (item.h !== undefined || item.ia !== undefined)) {
                  clearedRowEntries.push({
                    row: r,
                    prevH: item.h,
                    prevIa: item.ia,
                  });
                }
              }

              const redoMutationParams: ISetWorksheetDefaultRowHeightMutationParams =
                {
                  unitId,
                  subUnitId,
                  defaultRowHeight: destRowHeight,
                  clearedRowEntries,
                  isUndo: false,
                };
              const undoMutationParams: ISetWorksheetDefaultRowHeightMutationParams =
                {
                  unitId,
                  subUnitId,
                  defaultRowHeight: prevDefaultRowHeight,
                  clearedRowEntries,
                  isUndo: true,
                };

              const cmdSvc = accessor.get(ICommandService);
              const undoRedoSvc = accessor.get(IUndoRedoService);

              const res = cmdSvc.syncExecuteCommand(
                SetWorksheetDefaultRowHeightMutationId,
                redoMutationParams,
              );
              if (res) {
                undoRedoSvc.pushUndoRedo({
                  unitID: unitId,
                  undoMutations: [
                    {
                      id: SetWorksheetDefaultRowHeightMutationId,
                      params: undoMutationParams,
                    },
                  ],
                  redoMutations: [
                    {
                      id: SetWorksheetDefaultRowHeightMutationId,
                      params: redoMutationParams,
                    },
                  ],
                });
                return true;
              }
              return false;
            }

            // Normal or partial-row resize delegates cleanly to original handler
            return origDeltaCmd.handler(accessor, params);
          },
        };
        commandService.registerCommand(wrappedDeltaCmd);
        deltaWrapped = true;
      } else if (origDeltaCmd?.__isOptimizedWrapper) {
        deltaWrapped = true;
      }
    }

    // 4. Wrap SetRowHeightCommand
    if (commandService.hasCommand(SetRowHeightCommand.id)) {
      const origSetCmd = (commandService as any)._commandRegistry?.getCommand(
        SetRowHeightCommand.id,
      )?.[0];
      if (origSetCmd && !(origSetCmd as any).__isOptimizedWrapper) {
        commandService.unregisterCommand(SetRowHeightCommand.id);
        const wrappedSetCmd = {
          ...origSetCmd,
          id: SetRowHeightCommand.id,
          type: CommandType.COMMAND,
          __isOptimizedWrapper: true,
          handler: (accessor: any, params: any) => {
            const selectionManagerService = accessor.get(
              SheetsSelectionsService,
            );
            const selections = params?.ranges?.length
              ? params.ranges
              : selectionManagerService
                  ?.getCurrentSelections()
                  ?.map((s: any) => s.range);
            if (!selections?.length) return false;

            const target = getSheetCommandTarget(
              accessor.get(IUniverInstanceService),
              params,
            );
            if (!target) return false;
            const { worksheet, unitId, subUnitId } = target;
            const rowCount = worksheet.getRowCount();

            const RANGE_ALL = (RANGE_TYPE as any)?.ALL ?? 3;
            const isAll =
              selections.length === 1 &&
              (selections[0].rangeType === RANGE_ALL ||
                (selections[0].startRow === 0 &&
                  selections[0].endRow >= rowCount - 1));

            if (isAll) {
              const prevDefaultRowHeight =
                worksheet.getConfig().defaultRowHeight;
              const destRowHeight = Math.max(
                1,
                Math.min(2000, Number(params.value)),
              );

              const rowData = worksheet.getRowManager().getRowData() || {};
              const keys = Object.keys(rowData);
              const clearedRowEntries: Array<{
                row: number;
                prevH?: number;
                prevIa?: number;
              }> = [];
              for (let i = 0; i < keys.length; i++) {
                const r = Number(keys[i]);
                const item = rowData[r];
                if (item && (item.h !== undefined || item.ia !== undefined)) {
                  clearedRowEntries.push({
                    row: r,
                    prevH: item.h,
                    prevIa: item.ia,
                  });
                }
              }

              const redoMutationParams: ISetWorksheetDefaultRowHeightMutationParams =
                {
                  unitId,
                  subUnitId,
                  defaultRowHeight: destRowHeight,
                  clearedRowEntries,
                  isUndo: false,
                };
              const undoMutationParams: ISetWorksheetDefaultRowHeightMutationParams =
                {
                  unitId,
                  subUnitId,
                  defaultRowHeight: prevDefaultRowHeight,
                  clearedRowEntries,
                  isUndo: true,
                };

              const cmdSvc = accessor.get(ICommandService);
              const undoRedoSvc = accessor.get(IUndoRedoService);

              const res = cmdSvc.syncExecuteCommand(
                SetWorksheetDefaultRowHeightMutationId,
                redoMutationParams,
              );
              if (res) {
                undoRedoSvc.pushUndoRedo({
                  unitID: unitId,
                  undoMutations: [
                    {
                      id: SetWorksheetDefaultRowHeightMutationId,
                      params: undoMutationParams,
                    },
                  ],
                  redoMutations: [
                    {
                      id: SetWorksheetDefaultRowHeightMutationId,
                      params: redoMutationParams,
                    },
                  ],
                });
                return true;
              }
              return false;
            }

            return origSetCmd.handler(accessor, params);
          },
        };
        commandService.registerCommand(wrappedSetCmd);
        setWrapped = true;
      } else if (origSetCmd?.__isOptimizedWrapper) {
        setWrapped = true;
      }
    }

    if (deltaWrapped && setWrapped) {
      isCommandInterceptionInstalled = true;
      console.log(
        "[SKELETON_OPT] Successfully installed DeltaRowHeightCommand & SetRowHeightCommand enterprise Select-All optimization",
      );
      return true;
    }

    return false;
  } catch (cmdErr) {
    console.warn(
      "[SKELETON_OPT] Failed to install command interception:",
      cmdErr,
    );
    return false;
  }
}

/**
 * Installs enterprise SheetSkeleton layout caching and accumulation optimization.
 * Neutralizes main-thread blocking during rapid row/column resizing on large (4M+ rows) sheets.
 * Guarantees 100% semantic identity and strictly isolates cache per SheetSkeleton instance.
 */
export function installSheetSkeletonOptimization(injector?: any): void {
  if (injector && !isCommandInterceptionInstalled) {
    installRowHeightCommandInterception(injector);
  }

  if (isSkeletonOptimizationInstalled) return;

  try {
    // Enterprise Passivation: SheetRangeThemeModel.refreshSheetRowVisibleFuncSet
    // Eliminates 4,000,000-iteration synchronous loop on SetWorksheetRowHeightMutation
    // by leveraging sparse rowData dictionary instead of full-table scan.
    if (
      SheetRangeThemeModel?.prototype &&
      !(SheetRangeThemeModel.prototype as any).__passivated
    ) {
      (SheetRangeThemeModel.prototype as any).__passivated = true;
      SheetRangeThemeModel.prototype.refreshSheetRowVisibleFuncSet = function (
        unitId: string,
        subUnitId: string,
      ) {
        const set = (this as any)._getSheetRowVisibleFuncSet(unitId, subUnitId);
        set.clear();
        const workbook = (this as any)._univerInstanceService?.getUnit(unitId);
        if (!workbook) return;
        const sheet = workbook.getSheetBySheetId(subUnitId);
        if (!sheet) return;
        const rowCount = sheet.getRowCount();
        const rowManager = sheet.getRowManager();
        const defaultHeight = sheet.getConfig?.()?.defaultRowHeight ?? 24;
        if (defaultHeight === 0) {
          for (let i = 1; i <= rowCount; i++) set.add(i);
          return;
        }
        const rowData = rowManager.getRowData();
        if (rowData) {
          for (const rowKey in rowData) {
            const row = Number(rowKey);
            if (row >= 1 && row <= rowCount) {
              if (
                !sheet.getRowVisible(row) ||
                rowManager.getRowHeight(row) === 0
              ) {
                set.add(row);
              }
            }
          }
        }
      };
    }

    const proto: any = SheetSkeleton.prototype;
    if (!proto || typeof proto._updateLayout !== "function") {
      console.warn(
        "[SKELETON_OPT] SheetSkeleton._updateLayout not found, skipping optimization",
      );
      return;
    }

    // Optimized _generateRowMatrixCache with segment building for sparse rows/filters
    proto._generateRowMatrixCache = function (
      rowCount: number,
      rowData: any,
      defaultRowHeight: number,
    ) {
      if (rowCount <= 0) {
        return { rowTotalHeight: 0, rowHeightAccumulation: [] };
      }

      const data = rowData || {};
      const dataKeys = Object.keys(data);
      const hasCustomRows = dataKeys.length > 0;

      const gapConfig = this._gapConfig || {};
      const rowGaps = gapConfig.rowGaps;
      const hasRowGaps = !!(rowGaps && Object.keys(rowGaps).length > 0);

      let filterSet: Set<number> | null = null;
      let hasFilterInterceptor = false;

      try {
        const unitId = this.worksheet?.getUnitId?.();
        const subUnitId = this.worksheet?.getSheetId?.();
        const inj = this._injector || injector;
        if (inj && unitId && subUnitId) {
          let filterService: any = null;
          try {
            filterService = inj.get(SheetsFilterService);
          } catch {
            const resolved =
              inj.resolvedDependencyCollection?.resolvedDependencies;
            if (resolved) {
              for (const [k] of resolved.entries()) {
                if ((k?.name || "").includes("SheetsFilterService")) {
                  filterService = inj.get(k);
                  break;
                }
              }
            }
          }
          const filterModel = filterService?.getFilterModel(unitId, subUnitId);
          if (filterModel) {
            if (filterModel._alreadyFilteredOutRows instanceof Set) {
              filterSet = filterModel._alreadyFilteredOutRows;
            } else if (typeof filterModel.isRowFiltered === "function") {
              hasFilterInterceptor = true;
            }
          }
        }
      } catch {}

      const hasActiveFilterSet = !!(filterSet && filterSet.size > 0);

      // Fast Path A: Completely uniform rows (no custom rows, no gaps, no active filter)
      if (
        !hasCustomRows &&
        !hasRowGaps &&
        !hasActiveFilterSet &&
        !hasFilterInterceptor
      ) {
        let rowTotalHeight = 0;
        const rowHeightAccumulation = new Array(rowCount);
        for (let r = 0; r < rowCount; r++) {
          rowTotalHeight += defaultRowHeight;
          rowHeightAccumulation[r] = rowTotalHeight;
        }
        return { rowTotalHeight, rowHeightAccumulation };
      }

      // Fast Path B: Sparse special rows (< 5000 special rows)
      // Uses segment filling: uniform spans are filled with tight loops without hash lookups
      const specialRows = new Map<number, number>(); // row -> actualHeight (including gap)
      if (hasRowGaps) {
        for (const k of Object.keys(rowGaps)) {
          const r = Number(k);
          if (r < rowCount) {
            specialRows.set(r, defaultRowHeight + (rowGaps[r]?.size || 0));
          }
        }
      }
      if (hasCustomRows) {
        for (let i = 0; i < dataKeys.length; i++) {
          const r = Number(dataKeys[i]);
          if (r < rowCount) {
            const item = data[r];
            let h = defaultRowHeight;
            if (item) {
              const { h: customH = defaultRowHeight, ah, ia } = item;
              if (
                (ia == null || ia === 1) &&
                typeof ah === "number" &&
                ah > 0
              ) {
                h = ah;
              } else {
                h = customH;
              }
              if (item.hd === 1) h = 0;
            }
            const gap = hasRowGaps && rowGaps[r]?.size ? rowGaps[r].size : 0;
            specialRows.set(r, h + gap);
          }
        }
      }
      if (hasActiveFilterSet && filterSet && filterSet.size < 5000) {
        for (const r of filterSet) {
          if (r < rowCount) {
            const gap = hasRowGaps && rowGaps[r]?.size ? rowGaps[r].size : 0;
            specialRows.set(r, 0 + gap);
          }
        }
      }

      // If special rows count is bounded and no function-based filter interceptor:
      if (
        specialRows.size < 5000 &&
        !hasFilterInterceptor &&
        (!hasActiveFilterSet || (filterSet && filterSet.size < 5000))
      ) {
        let rowTotalHeight = 0;
        const rowHeightAccumulation = new Array(rowCount);
        const sortedRows = Array.from(specialRows.keys()).sort((a, b) => a - b);
        let curr = 0;

        for (let i = 0; i < sortedRows.length; i++) {
          const targetRow = sortedRows[i];
          if (targetRow >= rowCount) break;
          // Uniform segment before targetRow
          while (curr < targetRow) {
            rowTotalHeight += defaultRowHeight;
            rowHeightAccumulation[curr] = rowTotalHeight;
            curr++;
          }
          // Special row
          rowTotalHeight += specialRows.get(targetRow)!;
          rowHeightAccumulation[curr] = rowTotalHeight;
          curr++;
        }

        // Remaining uniform segment
        while (curr < rowCount) {
          rowTotalHeight += defaultRowHeight;
          rowHeightAccumulation[curr] = rowTotalHeight;
          curr++;
        }

        return { rowTotalHeight, rowHeightAccumulation };
      }

      // Fallback Path C: Dense special rows / large active filter set
      let rowTotalHeight = 0;
      const rowHeightAccumulation = new Array(rowCount);

      for (let r = 0; r < rowCount; r++) {
        let rowHeight = defaultRowHeight;

        if (hasActiveFilterSet && filterSet!.has(r)) {
          rowHeight = 0;
        } else if (hasFilterInterceptor && this.worksheet.getRowFiltered(r)) {
          rowHeight = 0;
        } else if (hasCustomRows && data[r] != null) {
          const rowDataItem = data[r];
          if (rowDataItem) {
            const { h = defaultRowHeight, ah, ia } = rowDataItem;
            if ((ia == null || ia === 1) && typeof ah === "number" && ah > 0) {
              rowHeight = ah;
            } else {
              rowHeight = h;
            }
            if (rowDataItem.hd === 1) rowHeight = 0;
          }
        }

        const gapSize = hasRowGaps && rowGaps[r]?.size ? rowGaps[r].size : 0;
        rowTotalHeight += gapSize + rowHeight;
        rowHeightAccumulation[r] = rowTotalHeight;
      }

      return { rowTotalHeight, rowHeightAccumulation };
    };

    // Optimized _updateLayout with independent row and column caching & single-row incremental shift
    proto._updateLayout = function () {
      if (!this.dirty) return;

      if (this._injector && !isCommandInterceptionInstalled) {
        installRowHeightCommandInterception(this._injector);
      }

      const {
        rowData,
        columnData,
        defaultRowHeight,
        defaultColumnWidth,
        rowCount,
        columnCount,
        rowHeader,
        columnHeader,
      } = this._worksheetData;

      const gapConfig = this._gapConfig || {};
      const rowGaps = gapConfig.rowGaps;
      const rowGapsKeys = rowGaps ? Object.keys(rowGaps) : [];
      const rowGapsKey =
        rowGapsKeys.length > 0
          ? rowGapsKeys.map((k) => `${k}:${rowGaps[k]?.size}`).join(";")
          : "empty";

      let filterSetRef: Set<number> | null = null;
      try {
        const unitId = this.worksheet?.getUnitId?.();
        const subUnitId = this.worksheet?.getSheetId?.();
        const inj = this._injector || injector;
        if (inj && unitId && subUnitId) {
          let filterService: any = null;
          try {
            filterService = inj.get(SheetsFilterService);
          } catch {
            const resolved =
              inj.resolvedDependencyCollection?.resolvedDependencies;
            if (resolved) {
              for (const [k] of resolved.entries()) {
                if ((k?.name || "").includes("SheetsFilterService")) {
                  filterService = inj.get(k);
                  break;
                }
              }
            }
          }
          const filterModel = filterService?.getFilterModel(unitId, subUnitId);
          if (
            filterModel &&
            filterModel._alreadyFilteredOutRows instanceof Set
          ) {
            filterSetRef = filterModel._alreadyFilteredOutRows;
          }
        }
      } catch {}

      const cachedRow = this._cachedRowLayout;
      let rowTotalHeight: number;
      let rowHeightAccumulation: number[];

      const canUseIncremental =
        cachedRow &&
        cachedRow.rowCount === rowCount &&
        cachedRow.defaultRowHeight === defaultRowHeight &&
        cachedRow.rowGapsKey === rowGapsKey &&
        cachedRow.filterSetRef === filterSetRef &&
        cachedRow.filterSize === (filterSetRef?.size || 0) &&
        Array.isArray(cachedRow.rowHeightAccumulation) &&
        cachedRow.rowHeightAccumulation.length === rowCount;

      if (canUseIncremental) {
        const dataKeys = rowData ? Object.keys(rowData) : [];
        let hasStructuralChange = false;
        const changedRows: Array<{ row: number; oldH: number; newH: number }> =
          [];

        // Compare current rowData against cached overrides
        for (let i = 0; i < dataKeys.length; i++) {
          const k = dataKeys[i];
          const r = Number(k);
          if (isNaN(r) || r < 0 || r >= rowCount) continue;
          const item = rowData[k];
          if (!item) continue;

          let effH = defaultRowHeight;
          const { h = defaultRowHeight, ah, ia, hd } = item;
          if ((ia == null || ia === 1) && typeof ah === "number" && ah > 0) {
            effH = ah;
          } else {
            effH = h;
          }
          if (hd === 1) effH = 0;

          const wasHidden = cachedRow.hiddenRowsSet?.has(r) ?? false;
          const isHidden = hd === 1;
          if (wasHidden !== isHidden) {
            hasStructuralChange = true;
            break;
          }

          const oldH = cachedRow.customHeightsMap?.get(r) ?? defaultRowHeight;
          if (effH !== oldH) {
            changedRows.push({ row: r, oldH, newH: effH });
            if (changedRows.length > 1) {
              hasStructuralChange = true;
              break;
            }
          }
        }

        // Also check if any row in customHeightsMap was removed from rowData
        if (!hasStructuralChange && cachedRow.customHeightsMap) {
          for (const [oldR, oldH] of cachedRow.customHeightsMap.entries()) {
            if (!rowData || rowData[oldR] == null) {
              const effH = defaultRowHeight;
              if (effH !== oldH) {
                changedRows.push({ row: oldR, oldH, newH: effH });
                if (changedRows.length > 1) {
                  hasStructuralChange = true;
                  break;
                }
              }
            }
          }
        }

        if (!hasStructuralChange && changedRows.length === 0) {
          // Path 1: Zero row changes (e.g. Column resize!)
          // 100% row layout reuse in O(1) time (0.00ms)
          rowTotalHeight = cachedRow.rowTotalHeight;
          rowHeightAccumulation = cachedRow.rowHeightAccumulation;
        } else if (!hasStructuralChange && changedRows.length === 1) {
          // Path 2: Single-row resize! Incremental in-place suffix shift in native array
          const { row: R, oldH, newH } = changedRows[0];
          let deltaH = newH - oldH;

          // If row is filtered out, rendered delta is 0
          if (filterSetRef && filterSetRef.has(R)) {
            deltaH = 0;
          }

          if (deltaH !== 0) {
            const arr = cachedRow.rowHeightAccumulation;
            for (let i = R; i < rowCount; i++) {
              arr[i] += deltaH;
            }
            cachedRow.rowTotalHeight += deltaH;
          }

          if (newH === defaultRowHeight) {
            cachedRow.customHeightsMap.delete(R);
          } else {
            cachedRow.customHeightsMap.set(R, newH);
          }

          rowTotalHeight = cachedRow.rowTotalHeight;
          rowHeightAccumulation = cachedRow.rowHeightAccumulation;
        } else {
          // Multi-row change: recalculate
          const res = this._generateRowMatrixCache(
            rowCount,
            rowData,
            defaultRowHeight,
          );
          rowTotalHeight = res.rowTotalHeight;
          rowHeightAccumulation = res.rowHeightAccumulation;

          const newCustomHeightsMap = new Map<number, number>();
          const newHiddenRowsSet = new Set<number>();
          if (rowData) {
            for (let i = 0; i < dataKeys.length; i++) {
              const k = dataKeys[i];
              const r = Number(k);
              const item = rowData[k];
              if (item) {
                if (item.hd === 1) newHiddenRowsSet.add(r);
                let effH = defaultRowHeight;
                const { h = defaultRowHeight, ah, ia } = item;
                if (
                  (ia == null || ia === 1) &&
                  typeof ah === "number" &&
                  ah > 0
                ) {
                  effH = ah;
                } else {
                  effH = h;
                }
                if (item.hd === 1) effH = 0;
                if (effH !== defaultRowHeight) {
                  newCustomHeightsMap.set(r, effH);
                }
              }
            }
          }

          this._cachedRowLayout = {
            rowCount,
            defaultRowHeight,
            customHeightsMap: newCustomHeightsMap,
            hiddenRowsSet: newHiddenRowsSet,
            rowGapsKey,
            filterSetRef,
            filterSize: filterSetRef?.size || 0,
            rowTotalHeight,
            rowHeightAccumulation,
          };
        }
      } else {
        // Initial layout / structural invalidation: full build
        const res = this._generateRowMatrixCache(
          rowCount,
          rowData,
          defaultRowHeight,
        );
        rowTotalHeight = res.rowTotalHeight;
        rowHeightAccumulation = res.rowHeightAccumulation;

        const dataKeys = rowData ? Object.keys(rowData) : [];
        const newCustomHeightsMap = new Map<number, number>();
        const newHiddenRowsSet = new Set<number>();
        if (rowData) {
          for (let i = 0; i < dataKeys.length; i++) {
            const k = dataKeys[i];
            const r = Number(k);
            const item = rowData[k];
            if (item) {
              if (item.hd === 1) newHiddenRowsSet.add(r);
              let effH = defaultRowHeight;
              const { h = defaultRowHeight, ah, ia } = item;
              if (
                (ia == null || ia === 1) &&
                typeof ah === "number" &&
                ah > 0
              ) {
                effH = ah;
              } else {
                effH = h;
              }
              if (item.hd === 1) effH = 0;
              if (effH !== defaultRowHeight) {
                newCustomHeightsMap.set(r, effH);
              }
            }
          }
        }

        this._cachedRowLayout = {
          rowCount,
          defaultRowHeight,
          customHeightsMap: newCustomHeightsMap,
          hiddenRowsSet: newHiddenRowsSet,
          rowGapsKey,
          filterSetRef,
          filterSize: filterSetRef?.size || 0,
          rowTotalHeight,
          rowHeightAccumulation,
        };
      }

      // 2. Build column layout cache key
      const colKeys = columnData ? Object.keys(columnData) : [];
      let colDataKey = "empty";
      if (colKeys.length > 0) {
        colDataKey = colKeys
          .map((k) => {
            const item = columnData[k];
            return `${k}:${item?.w}:${item?.hd}`;
          })
          .join(";");
      }

      const colGaps = gapConfig.colGaps;
      const colGapsKeys = colGaps ? Object.keys(colGaps) : [];
      let colGapsKey = "empty";
      if (colGapsKeys.length > 0) {
        colGapsKey = colGapsKeys
          .map((k) => `${k}:${colGaps[k]?.size}`)
          .join(";");
      }

      const cachedCol = this._cachedColLayout;
      let columnTotalWidth: number;
      let columnWidthAccumulation: number[];

      if (
        cachedCol &&
        cachedCol.columnCount === columnCount &&
        cachedCol.defaultColumnWidth === defaultColumnWidth &&
        cachedCol.colDataKey === colDataKey &&
        cachedCol.colGapsKey === colGapsKey &&
        cachedCol.columnWidthAccumulation?.length === columnCount
      ) {
        columnTotalWidth = cachedCol.columnTotalWidth;
        columnWidthAccumulation = cachedCol.columnWidthAccumulation;
      } else {
        const res = this._generateColumnMatrixCache(
          columnCount,
          columnData,
          defaultColumnWidth,
        );
        columnTotalWidth = res.columnTotalWidth;
        columnWidthAccumulation = res.columnWidthAccumulation;
        this._cachedColLayout = {
          columnCount,
          defaultColumnWidth,
          colDataKey,
          colGapsKey,
          columnTotalWidth,
          columnWidthAccumulation,
        };
      }

      this._rowHeaderWidth =
        rowHeader.hidden !== 1
          ? this._dynamicallyUpdateRowHeaderWidth(rowHeader)
          : 0;
      this._columnHeaderHeight =
        columnHeader.hidden !== 1 ? columnHeader.height : 0;
      this._rowTotalHeight = rowTotalHeight;
      this._rowHeightAccumulation = rowHeightAccumulation;
      this._columnTotalWidth = columnTotalWidth;
      this._columnWidthAccumulation = columnWidthAccumulation;
      this.makeDirty(false);
    };

    isSkeletonOptimizationInstalled = true;
    console.log(
      "[SKELETON_OPT] Successfully installed SheetSkeleton layout caching & incremental resize optimization",
    );
  } catch (err) {
    console.warn(
      "[SKELETON_OPT] Failed to install SheetSkeleton optimization:",
      err,
    );
  }
}
