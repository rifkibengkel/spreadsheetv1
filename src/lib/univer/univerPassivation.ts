import { SetRangeValuesMutation } from '@univerjs/sheets';
import { IActiveDirtyManagerService } from '@univerjs/engine-formula';

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
export function installUniverCalculationPassivation(injector: any): PassivationController {
  if (!injector) {
    console.warn('[UNIVER_PASSIVATION] Injector not provided, skipping passivation');
    return { isPassivated: false, restore: () => {} };
  }

  try {
    const activeDirtyManager = injector.get(IActiveDirtyManagerService);
    if (!activeDirtyManager) {
      console.warn('[UNIVER_PASSIVATION] IActiveDirtyManagerService not found');
      return { isPassivated: false, restore: () => {} };
    }

    // Save original SetRangeValuesMutation dirty configuration if present
    const originalSetRange = activeDirtyManager.get(SetRangeValuesMutation.id);

    const SUPPRESSED_MUTATIONS = [
      SetRangeValuesMutation.id,
      'sheet.mutation.set-filter-range',
      'sheet.mutation.set-filter-criteria',
      'sheet.mutation.remove-filter',
      'sheet.mutation.re-calc-filter',
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
    activeDirtyManager.register = function (commandId: string, dirtyConversion: any) {
      if (SUPPRESSED_MUTATIONS.includes(commandId)) {
        return; // Suppress registration completely
      }
      return origRegister(commandId, dirtyConversion);
    };

    console.log('[UNIVER_PASSIVATION] Successfully passivated browser-side formula calculation engine (TriggerCalculationController disabled for edits and filters)');

    return {
      isPassivated: true,
      restore: () => {
        if (originalSetRange) {
          activeDirtyManager.register(SetRangeValuesMutation.id, originalSetRange);
        }
      },
    };
  } catch (err) {
    console.warn('[UNIVER_PASSIVATION] Failed to install calculation passivation:', err);
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
  patches: Array<{ sheetId: string; row: number; col: number; v: any; f?: string; t?: number }>
): void {
  if (!univerAPI || !patches || patches.length === 0) return;

  try {
    const workbook = univerAPI.getActiveWorkbook?.();
    if (!workbook) return;

    // Group patches by worksheet for batched application
    const sheetPatchMap = new Map<string, Record<number, Record<number, any>>>();

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
        cmdSvc.syncExecuteCommand(SetRangeValuesMutation.id, mutationParams, mutationOptions);
      } else if (cmdSvc?.executeCommand) {
        cmdSvc.executeCommand(SetRangeValuesMutation.id, mutationParams, mutationOptions);
      } else if (univerAPI.executeCommand) {
        univerAPI.executeCommand(SetRangeValuesMutation.id, mutationParams, mutationOptions);
      }
    }
  } catch (err) {
    console.warn('[UNIVER_PASSIVATION] Failed to apply silent calculation patch:', err);
  }
}
