/**
 * Workbook Session Manager
 * Retains the original imported XLSX File/Buffer and tracks sheet modifications (dirty state).
 */

class WorkbookSessionManager {
  private originalFileBuffer: Uint8Array | null = null;
  private originalFileName: string | null = null;
  private dirtySheetIds = new Set<string>();
  private sheetNameToIndexMap = new Map<string, number>();

  public setOriginalFile(file: File | Blob, arrayBuffer?: ArrayBuffer, sheetNames?: string[]) {
    this.originalFileName = (file as any).name || 'workbook.xlsx';
    this.dirtySheetIds.clear();
    this.sheetNameToIndexMap.clear();

    if (sheetNames && Array.isArray(sheetNames)) {
      sheetNames.forEach((name, idx) => {
        this.sheetNameToIndexMap.set(name, idx + 1);
      });
    }

    if (arrayBuffer) {
      this.originalFileBuffer = new Uint8Array(arrayBuffer);
    } else if (file instanceof File || file instanceof Blob) {
      file.arrayBuffer().then((buf) => {
        this.originalFileBuffer = new Uint8Array(buf);
      });
    }
  }

  public setOriginalBuffer(buffer: Uint8Array, fileName: string, sheetNames?: string[]) {
    this.originalFileBuffer = buffer;
    this.originalFileName = fileName;
    this.dirtySheetIds.clear();
    this.sheetNameToIndexMap.clear();

    if (sheetNames && Array.isArray(sheetNames)) {
      sheetNames.forEach((name, idx) => {
        this.sheetNameToIndexMap.set(name, idx + 1);
      });
    }
  }

  public getOriginalBuffer(): Uint8Array | null {
    return this.originalFileBuffer;
  }

  public getOriginalFileName(): string | null {
    return this.originalFileName;
  }

  public hasOriginalWorkbook(): boolean {
    return this.originalFileBuffer !== null && this.originalFileBuffer.length > 0;
  }

  public markSheetDirty(sheetIdOrName: string) {
    this.dirtySheetIds.add(sheetIdOrName);
  }

  public isSheetDirty(sheetIdOrName: string): boolean {
    return this.dirtySheetIds.has(sheetIdOrName);
  }

  public getDirtySheetIds(): Set<string> {
    return new Set(this.dirtySheetIds);
  }

  public markAllClean() {
    this.dirtySheetIds.clear();
  }

  public clear() {
    this.originalFileBuffer = null;
    this.originalFileName = null;
    this.dirtySheetIds.clear();
    this.sheetNameToIndexMap.clear();
  }
}

const GLOBAL_SESSION_KEY = '__WORKBOOK_SESSION_INSTANCE__';

export const workbookSession: WorkbookSessionManager =
  typeof window !== 'undefined' && (window as any)[GLOBAL_SESSION_KEY]
    ? (window as any)[GLOBAL_SESSION_KEY]
    : new WorkbookSessionManager();

if (typeof window !== 'undefined') {
  (window as any)[GLOBAL_SESSION_KEY] = workbookSession;
  (window as any).workbookSession = workbookSession;
}
