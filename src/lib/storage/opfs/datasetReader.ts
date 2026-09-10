/**
 * OPFS Dataset Reader
 * Invariant 6 & Section 10: Non-exclusive runtime reader using getFile().slice().arrayBuffer().
 * Enables concurrent random reads without SyncAccessHandle contention.
 */

import { MetadataStore } from './metadataStore';
import { RowIndex } from './rowIndex';
import { parseCsvLine } from './csvTokenizer';
import { OPFS_DATA_FILE, OPFS_INDEX_FILE, OPFSMetadata, RowSlice } from './types';

export class DatasetReader {
  private dataFile: File | null = null;
  private indexArray: BigUint64Array | null = null;
  private metadata: OPFSMetadata | null = null;
  private decoder = new TextDecoder('utf-8');
  private isReady = false;
  private initError: string | null = null;

  /**
   * Initializes the reader by loading metadata and cached index.
   * Enforces B-3 physical file validation on startup.
   */
  public async init(root?: FileSystemDirectoryHandle): Promise<boolean> {
    this.initError = null;
    this.isReady = false;
    this.dataFile = null;
    this.indexArray = null;

    const dir = root || (await MetadataStore.getRootDir());

    this.metadata = await MetadataStore.read(dir);
    if (!this.metadata) {
      this.initError = 'Metadata file table_meta.json not found or unreadable.';
      return false;
    }

    if (
      !this.metadata ||
      this.metadata.status !== 'COMMITTED' ||
      typeof this.metadata.rowCount !== 'number' ||
      typeof this.metadata.dataSizeBytes !== 'number' ||
      typeof this.metadata.indexSizeBytes !== 'number' ||
      isNaN(this.metadata.rowCount) ||
      isNaN(this.metadata.dataSizeBytes) ||
      isNaN(this.metadata.indexSizeBytes) ||
      this.metadata.rowCount < 0 ||
      this.metadata.dataSizeBytes < 0 ||
      this.metadata.indexSizeBytes < 0
    ) {
      this.initError = `Metadata is structurally invalid or uncommitted: status=${this.metadata?.status}, rowCount=${this.metadata?.rowCount}, dataSize=${this.metadata?.dataSizeBytes}, indexSize=${this.metadata?.indexSizeBytes}.`;
      return false;
    }

    // B-3 physical validation 1: Verify data file exists and size matches
    let dataHandle: any;
    try {
      dataHandle = await dir.getFileHandle(OPFS_DATA_FILE);
    } catch (e: any) {
      this.initError = `Physical data file ${OPFS_DATA_FILE} not found: ${e.message}`;
      return false;
    }

    const dataFile = await dataHandle.getFile();
    if (dataFile.size !== this.metadata.dataSizeBytes) {
      this.initError = `Physical data file size mismatch: expected ${this.metadata.dataSizeBytes} bytes, found ${dataFile.size} bytes.`;
      return false;
    }
    this.dataFile = dataFile;

    // B-3 physical validation 2: Verify index file exists and size matches
    let indexHandle: any;
    try {
      indexHandle = await dir.getFileHandle(OPFS_INDEX_FILE);
    } catch (e: any) {
      this.initError = `Physical index file ${OPFS_INDEX_FILE} not found: ${e.message}`;
      this.dataFile = null;
      return false;
    }

    const indexFile = await indexHandle.getFile();
    if (indexFile.size !== this.metadata.indexSizeBytes) {
      this.initError = `Physical index file size mismatch: expected ${this.metadata.indexSizeBytes} bytes, found ${indexFile.size} bytes.`;
      this.dataFile = null;
      return false;
    }

    // B-3 physical validation 3: Verify index size equals (rowCount + 1) * 8
    const expectedIndexSize = (this.metadata.rowCount + 1) * 8;
    if (this.metadata.indexSizeBytes !== expectedIndexSize || indexFile.size !== expectedIndexSize) {
      this.initError = `Mathematically invalid index size: rowCount ${this.metadata.rowCount} requires ${expectedIndexSize} bytes, found metadata=${this.metadata.indexSizeBytes}, file=${indexFile.size}.`;
      this.dataFile = null;
      return false;
    }

    // B-3 physical validation 4: Parse index array and verify length and bounds
    const indexBuffer = await indexFile.arrayBuffer();
    try {
      this.indexArray = RowIndex.fromBuffer(indexBuffer);
    } catch (err: any) {
      this.initError = `Failed to parse index buffer: ${err.message}`;
      this.dataFile = null;
      return false;
    }

    // B-3 physical validation 5: Full monotonicity, bounds, and range sweep
    const indexValidation = RowIndex.validate(
      this.indexArray,
      this.metadata.rowCount,
      this.metadata.dataSizeBytes
    );
    if (!indexValidation.valid) {
      this.initError = `Index validation failed: ${indexValidation.error}`;
      this.dataFile = null;
      this.indexArray = null;
      return false;
    }

    this.isReady = true;
    return true;
  }

  public getInitError(): string | null {
    return this.initError;
  }

  public getSessionId(): string | null {
    return this.metadata?.sessionId || null;
  }

  public getMetadata(): OPFSMetadata | null {
    return this.metadata;
  }

  public getRowCount(): number {
    return this.metadata ? this.metadata.rowCount : 0;
  }

  public getColumnCount(): number {
    return this.metadata ? this.metadata.columnCount : 0;
  }

  public getHeaders(): string[] {
    return this.metadata ? this.metadata.headers : [];
  }

  /**
   * Reads a single row by 0-based data row index.
   */
  public async readRow(rowIndex: number): Promise<RowSlice> {
    if (!this.isReady || !this.dataFile || !this.indexArray) {
      throw new Error('DatasetReader is not initialized with a committed dataset.');
    }

    const { start, end } = RowIndex.getRowBounds(this.indexArray, rowIndex);
    const sliceBlob = this.dataFile.slice(start, end);
    const sliceBuf = await sliceBlob.arrayBuffer();
    const rawBytes = new Uint8Array(sliceBuf);
    const rawText = this.decoder.decode(rawBytes).replace(/\r?\n$/, '');
    const cells = parseCsvLine(rawText);

    return {
      rowIndex,
      rawBytes,
      rawText,
      cells,
    };
  }

  /**
   * Reads a contiguous range of rows in a single batch I/O read.
   * startRow and endRow are inclusive.
   */
  public async readRowRange(startRow: number, endRow: number): Promise<RowSlice[]> {
    if (!this.isReady || !this.dataFile || !this.indexArray || !this.metadata) {
      throw new Error('DatasetReader is not initialized with a committed dataset.');
    }

    const maxRow = this.metadata.rowCount - 1;
    const clampedStart = Math.max(0, Math.min(startRow, maxRow));
    const clampedEnd = Math.max(clampedStart, Math.min(endRow, maxRow));
    const rowCount = clampedEnd - clampedStart + 1;

    if (rowCount <= 0) return [];

    // Contiguous byte range: from start of clampedStart to end of clampedEnd
    const startByte = Number(this.indexArray[clampedStart]);
    const endByte = Number(this.indexArray[clampedEnd + 1]);

    const batchBlob = this.dataFile.slice(startByte, endByte);
    const batchBuf = await batchBlob.arrayBuffer();
    const batchBytes = new Uint8Array(batchBuf);

    const results: RowSlice[] = new Array(rowCount);

    for (let r = clampedStart; r <= clampedEnd; r++) {
      const rowStartInBatch = Number(this.indexArray[r]) - startByte;
      const rowEndInBatch = Number(this.indexArray[r + 1]) - startByte;
      const rowBytes = batchBytes.subarray(rowStartInBatch, rowEndInBatch);
      const rawText = this.decoder.decode(rowBytes).replace(/\r?\n$/, '');
      const cells = parseCsvLine(rawText);

      results[r - clampedStart] = {
        rowIndex: r,
        rawBytes: rowBytes,
        rawText,
        cells,
      };
    }

    return results;
  }
}
