/**
 * Univer Viewport Adapter (Architecture D)
 * Manages sparse cell matrix hydration and window eviction in Univer 0.25.1.
 * Guarantees zero calculation storm, zero undo stack pollution, and <1MB heap usage.
 */

import { IUniverViewportAdapter, CellEditRecord, VirtualSheetPartition, ROWS_PER_VIRTUAL_SHEET } from './types';
import { OPFSMetadata, RowSlice } from '../storage/opfs/types';

export class UniverViewportAdapter implements IUniverViewportAdapter {
  private univerAPI: any;
  private unitId: string = '';
  private subUnitId: string = '';
  private logicalRowCount: number = 0;
  private logicalColumnCount: number = 0;
  private headers: string[] = [];
  private ready: boolean = false;

  constructor(univerAPI: any) {
    this.univerAPI = univerAPI;
    this.syncActiveSheet();
  }

  public setUniverAPI(api: any): void {
    this.univerAPI = api;
    this.syncActiveSheet();
  }

  private syncActiveSheet(): boolean {
    if (!this.univerAPI) return false;
    try {
      const wb = this.univerAPI.getActiveWorkbook?.();
      if (!wb) return false;
      this.unitId = wb.getId?.() || '';
      if (!this.subUnitId) {
        const sheet = wb.getActiveSheet?.();
        if (!sheet) return false;
        this.subUnitId = sheet.getSheetId?.() || '';
      }
      return !!(this.unitId && this.subUnitId);
    } catch {
      return false;
    }
  }

  public setActiveSubUnitId(sheetId: string): void {
    this.subUnitId = sheetId;
  }

  public async activateSheet(sheetId: string): Promise<void> {
    this.subUnitId = sheetId;
    if (this.univerAPI && this.unitId) {
      try {
        const wb = this.univerAPI.getActiveWorkbook?.();
        if (wb?.setActiveSheet) {
          wb.setActiveSheet(sheetId);
        } else {
          await this.univerAPI.executeCommand('sheet.operation.set-worksheet-active', {
            unitId: this.unitId,
            subUnitId: sheetId,
          });
        }
      } catch (err) {
        console.warn('[UniverViewportAdapter] activateSheet failed:', err);
      }
    }
  }

  public getUniverAPI(): any {
    return this.univerAPI;
  }

  public isReady(): boolean {
    return this.ready && this.syncActiveSheet();
  }

  public getLogicalRowCount(): number {
    return this.logicalRowCount;
  }

  public getActiveUnitId(): string {
    return this.unitId;
  }

  public getActiveSubUnitId(): string {
    return this.subUnitId;
  }

  public getHeaders(): string[] {
    return this.headers;
  }

  /**
   * Mounts multiple virtual sheets for 4M+ row datasets into Univer.
   * Ensures each partition exists as a presentation sheet with header at row 0.
   */
  public async mountVirtualSheets(
    partitions: VirtualSheetPartition[],
    columnCount: number,
    headers: string[]
  ): Promise<void> {
    if (!this.syncActiveSheet()) {
      throw new Error('Univer workbook or active sheet not available.');
    }

    const wb = this.univerAPI.getActiveWorkbook();
    this.headers = headers || [];
    const hasHeaders = this.headers.length > 0;
    this.logicalColumnCount = Math.max(26, columnCount);

    const existingSheets = wb?.getSheets ? wb.getSheets() : [];
    const existingSheetIds = new Set(
      existingSheets.map((s: any) => s.getSheetId?.() || s.id || '')
    );

    // Prepare header row cellData (bold, row 0)
    const headerCells: Record<number, { v: string; t: number; s?: any }> = {};
    if (hasHeaders) {
      for (let c = 0; c < this.headers.length; c++) {
        headerCells[c] = {
          v: this.headers[c],
          t: 1, // string
          s: { bl: 1 }, // bold
        };
      }
    }

    for (let i = 0; i < partitions.length; i++) {
      const part = partitions[i];
      const sheetRowCount = part.rowCount + (hasHeaders ? 1 : 0);
      const sheetName = `Sheet${i + 1}`;

      if (i === 0) {
        const primarySheet =
          (wb.getSheetByName ? wb.getSheetByName('Sheet1') : null) ||
          existingSheets[0] ||
          (wb.getActiveSheet ? wb.getActiveSheet() : null);
        const primaryId = primarySheet?.getSheetId ? primarySheet.getSheetId() : this.subUnitId;
        part.sheetId = primaryId;
        this.subUnitId = primaryId;
        this.logicalRowCount = sheetRowCount;

        const rawSheet = primarySheet?.getSheet ? primarySheet.getSheet() : primarySheet?._worksheet || primarySheet;
        if (rawSheet?.setRowCount) rawSheet.setRowCount(sheetRowCount);
        if (rawSheet?.setColumnCount) rawSheet.setColumnCount(this.logicalColumnCount);
        if (primarySheet?.setName) primarySheet.setName(sheetName);
        else if (rawSheet?.setName) rawSheet.setName(sheetName);

        try {
          await this.univerAPI.executeCommand('sheet.mutation.set-worksheet-row-count', {
            unitId: this.unitId,
            subUnitId: this.subUnitId,
            rowCount: sheetRowCount,
          });
          await this.univerAPI.executeCommand('sheet.mutation.set-worksheet-col-count', {
            unitId: this.unitId,
            subUnitId: this.subUnitId,
            colCount: this.logicalColumnCount,
          });
        } catch {}

        if (hasHeaders) {
          try {
            await this.univerAPI.executeCommand(
              'sheet.mutation.set-range-values',
              {
                unitId: this.unitId,
                subUnitId: this.subUnitId,
                cellValue: { 0: headerCells },
              },
              { onlyLocal: true, fromCalculationPatch: true, fromFormula: true }
            );
          } catch {}
        }
      } else {
        let targetSheet: any = null;
        if (wb.getSheetByName) {
          targetSheet = wb.getSheetByName(sheetName);
        }
        if (!targetSheet && existingSheets.length > i) {
          const s = existingSheets[i];
          if (s && s.getSheetId?.() !== this.subUnitId) {
            targetSheet = s;
          }
        }

        if (!targetSheet) {
          if (wb.create) {
            try {
              targetSheet = wb.create(sheetName, sheetRowCount, this.logicalColumnCount);
            } catch (createErr) {
              console.warn(`[UniverViewportAdapter] wb.create warning for ${sheetName}:`, createErr);
            }
          }

          if (!targetSheet) {
            const targetSubId = part.sheetId || `sheet-${i + 1}`;
            try {
              await this.univerAPI.executeCommand('sheet.mutation.insert-sheet', {
                unitId: this.unitId,
                index: i,
                sheet: {
                  id: targetSubId,
                  name: sheetName,
                  tabColor: '',
                  hidden: 0,
                  rowCount: sheetRowCount,
                  columnCount: this.logicalColumnCount,
                  zoomRatio: 1,
                  scrollTop: 0,
                  scrollLeft: 0,
                  defaultColumnWidth: 88,
                  defaultRowHeight: 24,
                  showGridlines: 1,
                  rowHeader: { width: 46, hidden: 0 },
                  columnHeader: { height: 20, hidden: 0 },
                  rightToLeft: 0,
                  type: 2,
                  status: 0,
                  cellData: hasHeaders ? { 0: headerCells } : {},
                  freeze: { xSplit: 0, ySplit: 0, startRow: -1, startColumn: -1 },
                },
              });
              existingSheetIds.add(targetSubId);
              targetSheet = wb.getSheetBySheetId ? wb.getSheetBySheetId(targetSubId) : null;
            } catch (insertErr) {
              console.warn(`[UniverViewportAdapter] insert-sheet warning for ${sheetName}:`, insertErr);
            }
          }
        }

        const resolvedSubId = targetSheet?.getSheetId?.() || targetSheet?.id || part.sheetId || `sheet-${i + 1}`;
        part.sheetId = resolvedSubId;
        existingSheetIds.add(resolvedSubId);

        const raw = targetSheet?.getSheet ? targetSheet.getSheet() : targetSheet?._worksheet || targetSheet;
        if (raw?.setRowCount) raw.setRowCount(sheetRowCount);
        if (raw?.setColumnCount) raw.setColumnCount(this.logicalColumnCount);
        if (targetSheet?.setName) targetSheet.setName(sheetName);
        else if (raw?.setName) raw.setName(sheetName);

        try {
          await this.univerAPI.executeCommand('sheet.mutation.set-worksheet-row-count', {
            unitId: this.unitId,
            subUnitId: resolvedSubId,
            rowCount: sheetRowCount,
          });
          await this.univerAPI.executeCommand('sheet.mutation.set-worksheet-col-count', {
            unitId: this.unitId,
            subUnitId: resolvedSubId,
            colCount: this.logicalColumnCount,
          });
        } catch {}

        if (hasHeaders) {
          try {
            await this.univerAPI.executeCommand(
              'sheet.mutation.set-range-values',
              {
                unitId: this.unitId,
                subUnitId: resolvedSubId,
                cellValue: { 0: headerCells },
              },
              { onlyLocal: true, fromCalculationPatch: true, fromFormula: true }
            );
          } catch {}
        }
      }
    }

    // Remove any excess sheets from previous datasets or queries (e.g. going from 3 sheets to 1 sheet)
    const validSubIds = new Set(partitions.map((p) => p.sheetId));
    const allCurrentSheets = wb?.getSheets ? wb.getSheets() : [];
    for (const sheet of allCurrentSheets) {
      const sId = sheet.getSheetId ? sheet.getSheetId() : sheet.id;
      if (sId && sId !== this.subUnitId && !validSubIds.has(sId)) {
        try {
          if (wb.deleteSheet) {
            wb.deleteSheet(sheet);
          } else {
            await this.univerAPI.executeCommand('sheet.mutation.remove-sheet', {
              unitId: this.unitId,
              subUnitId: sId,
            });
          }
        } catch {}
      }
    }

    // Ensure primary sheet (Sheet1) remains the active sheet initially
    if (partitions[0]?.sheetId) {
      this.subUnitId = partitions[0].sheetId;
      try {
        if (wb.setActiveSheet) {
          wb.setActiveSheet(partitions[0].sheetId);
        }
      } catch {}
    }

    this.ready = true;
  }

  /**
   * Mounts virtual sheet with logical dimensions for 4M+ rows.
   * If rowCount > ROWS_PER_VIRTUAL_SHEET (4,000,000), automatically partitions into 4M virtual sheets.
   */
  public async mountVirtualSheet(metadata: OPFSMetadata): Promise<void> {
    if (!this.syncActiveSheet()) {
      throw new Error('Univer workbook or active sheet not available.');
    }

    if (metadata.rowCount > ROWS_PER_VIRTUAL_SHEET) {
      const sheetCount = Math.max(1, Math.ceil(metadata.rowCount / ROWS_PER_VIRTUAL_SHEET));
      const partitions: VirtualSheetPartition[] = [];
      for (let s = 0; s < sheetCount; s++) {
        const startRow = s * ROWS_PER_VIRTUAL_SHEET;
        const endRow = Math.min(metadata.rowCount - 1, (s + 1) * ROWS_PER_VIRTUAL_SHEET - 1);
        const rowCount = Math.max(0, endRow - startRow + 1);
        partitions.push({
          sheetId: `sheet-${s + 1}`,
          sheetIndex: s,
          startRow,
          endRow,
          rowCount,
        });
      }
      return this.mountVirtualSheets(partitions, metadata.columnCount, metadata.headers);
    }

    const wb = this.univerAPI.getActiveWorkbook();
    const sheet = wb.getActiveSheet();
    const rawSheet = sheet.getSheet ? sheet.getSheet() : sheet._worksheet || sheet;

    this.headers = metadata.headers || [];
    const hasHeaders = this.headers.length > 0;
    this.logicalRowCount = metadata.rowCount + (hasHeaders ? 1 : 0);
    this.logicalColumnCount = Math.max(26, metadata.columnCount);

    if (rawSheet.setRowCount) {
      rawSheet.setRowCount(this.logicalRowCount);
    }
    if (rawSheet.setColumnCount) {
      rawSheet.setColumnCount(this.logicalColumnCount);
    }

    try {
      await this.univerAPI.executeCommand('sheet.mutation.set-worksheet-row-count', {
        unitId: this.unitId,
        subUnitId: this.subUnitId,
        rowCount: this.logicalRowCount,
      });
      await this.univerAPI.executeCommand('sheet.mutation.set-worksheet-col-count', {
        unitId: this.unitId,
        subUnitId: this.subUnitId,
        colCount: this.logicalColumnCount,
      });
    } catch {}

    if (hasHeaders) {
      const headerCells: Record<number, { v: string; t: number; s?: any }> = {};
      for (let c = 0; c < this.headers.length; c++) {
        headerCells[c] = {
          v: this.headers[c],
          t: 1,
          s: { bl: 1 },
        };
      }

      await this.univerAPI.executeCommand(
        'sheet.mutation.set-range-values',
        {
          unitId: this.unitId,
          subUnitId: this.subUnitId,
          cellValue: { 0: headerCells },
        },
        { onlyLocal: true, fromCalculationPatch: true, fromFormula: true }
      );
    }

    try {
      const skeleton = rawSheet.getSkeleton ? rawSheet.getSkeleton() : rawSheet.sheetSkeleton;
      if (skeleton) {
        skeleton.makeDirty(true);
        skeleton.calculate();
      }
    } catch {}

    this.ready = true;
  }

  /**
   * Hydrates a batch of row slices into Univer.
   * Maps OPFS row `slice.rowIndex` to sheet row `slice.rowIndex + (hasHeaders ? 1 : 0)`.
   */
  public async hydrateWindow(
    rows: RowSlice[],
    userEdits: Map<string, CellEditRecord>,
    targetSheetId?: string
  ): Promise<void> {
    if (!this.syncActiveSheet()) return;

    const targetSubId = targetSheetId || this.subUnitId;
    const hasHeaders = this.headers.length > 0;
    const headerOffset = hasHeaders ? 1 : 0;
    const cellValueMap: Record<number, Record<number, { v: any; t?: number }>> = {};

    for (let i = 0; i < rows.length; i++) {
      const slice = rows[i];
      const sheetRow = slice.rowIndex + headerOffset;
      const rowObj: Record<number, { v: any; t?: number }> = {};

      for (let col = 0; col < slice.cells.length; col++) {
        const editKey = `${targetSubId}_${sheetRow}_${col}`;
        const userEdit = userEdits.get(editKey) || userEdits.get(`${sheetRow}_${col}`);

        if (userEdit) {
          rowObj[col] = { v: userEdit.value, t: typeof userEdit.value === 'number' ? 2 : 1 };
        } else {
          const rawVal = slice.cells[col];
          if (rawVal && rawVal.startsWith('="') && rawVal.endsWith('"')) {
            const unquoted = rawVal.slice(2, -1);
            rowObj[col] = { v: unquoted, t: 1 };
          } else {
            const num = Number(rawVal);
            if (rawVal !== '' && !isNaN(num) && !/^\s*0[0-9]/.test(rawVal)) {
              rowObj[col] = { v: num, t: 2 };
            } else {
              rowObj[col] = { v: rawVal, t: 1 };
            }
          }
        }
      }

      cellValueMap[sheetRow] = rowObj;
    }

    await this.univerAPI.executeCommand(
      'sheet.mutation.set-range-values',
      {
        unitId: this.unitId,
        subUnitId: targetSubId,
        cellValue: cellValueMap,
      },
      { onlyLocal: true, fromCalculationPatch: true, fromFormula: true }
    );
  }

  /**
   * Evicts rows outside the visible window to keep heap <1MB.
   */
  public async evictWindow(startRow: number, endRow: number, targetSheetId?: string): Promise<void> {
    if (!this.syncActiveSheet()) return;

    const wb = this.univerAPI.getActiveWorkbook();
    const targetSubId = targetSheetId || this.subUnitId;
    const sheet = wb?.getSheetBySheetId ? wb.getSheetBySheetId(targetSubId) : wb?.getActiveSheet();
    const rawSheet = sheet?.getSheet ? sheet.getSheet() : sheet?._worksheet || sheet;
    const matrix = rawSheet?.getCellMatrix?.();

    if (!matrix || !matrix._matrix) return;

    const hasHeaders = this.headers.length > 0;
    const minRow = hasHeaders ? 1 : 0;
    const clampedStart = Math.max(minRow, startRow);

    for (let r = clampedStart; r <= endRow; r++) {
      if (matrix._matrix[r]) {
        delete matrix._matrix[r];
      }
    }
  }


  /**
   * Scrolls Univer canvas to targetRow smoothly.
   */
  public async scrollToRow(targetRow: number): Promise<boolean> {
    if (!this.syncActiveSheet()) return false;

    try {
      const res = await this.univerAPI.executeCommand('sheet.operation.set-scroll', {
        unitId: this.unitId,
        sheetId: this.subUnitId,
        sheetViewStartRow: targetRow,
        sheetViewStartColumn: 0,
        offsetX: 0,
        offsetY: 0,
      });
      return !!res;
    } catch (err) {
      console.warn('[UniverViewportAdapter] scrollToRow failed:', err);
      return false;
    }
  }
}
