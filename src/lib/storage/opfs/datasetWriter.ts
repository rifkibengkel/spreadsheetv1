/**
 * OPFS Dataset Writer
 * Invariant 1, 2, 6 & Section 10: Import-time writer managing exclusive FileSystemSyncAccessHandle.
 */

import { MetadataStore } from './metadataStore';
import { OPFS_DATA_FILE, OPFS_INDEX_FILE } from './types';

export class DatasetWriter {
  private dataHandle: any = null;
  private indexHandle: any = null;
  private dataBytesWritten = 0;
  private isOpen = false;

  /**
   * Initializes handles for table_data.bin and table_index.idx.
   */
  public async open(root?: FileSystemDirectoryHandle): Promise<void> {
    const dir = root || (await MetadataStore.getRootDir());

    const dataFileHandle = await dir.getFileHandle(OPFS_DATA_FILE, { create: true });
    const indexFileHandle = await dir.getFileHandle(OPFS_INDEX_FILE, { create: true });

    if ((dataFileHandle as any).createSyncAccessHandle) {
      this.dataHandle = await (dataFileHandle as any).createSyncAccessHandle();
      this.dataHandle.truncate(0);
    } else if ((dataFileHandle as any).createWritable) {
      this.dataHandle = await (dataFileHandle as any).createWritable();
    } else {
      throw new Error('No writable handle interface found on OPFS data file.');
    }

    if ((indexFileHandle as any).createSyncAccessHandle) {
      this.indexHandle = await (indexFileHandle as any).createSyncAccessHandle();
      this.indexHandle.truncate(0);
    } else if ((indexFileHandle as any).createWritable) {
      this.indexHandle = await (indexFileHandle as any).createWritable();
    } else {
      throw new Error('No writable handle interface found on OPFS index file.');
    }

    this.dataBytesWritten = 0;
    this.isOpen = true;
  }

  /**
   * Writes a chunk of raw data row bytes to table_data.bin.
   */
  public async writeDataChunk(chunk: Uint8Array): Promise<number> {
    if (!this.isOpen) throw new Error('DatasetWriter is not open.');

    if (this.dataHandle.write && typeof this.dataHandle.write === 'function') {
      const written = this.dataHandle.write(chunk, { at: this.dataBytesWritten });
      const bytes = typeof written === 'number' ? written : chunk.byteLength;
      this.dataBytesWritten += bytes;
      return bytes;
    }

    throw new Error('Unable to write to data handle.');
  }

  /**
   * Writes the complete N+1 uint64 index buffer to table_index.idx.
   */
  public async writeIndexBuffer(indexBuffer: Uint8Array): Promise<void> {
    if (!this.isOpen) throw new Error('DatasetWriter is not open.');

    if (this.indexHandle.write && typeof this.indexHandle.write === 'function') {
      this.indexHandle.write(indexBuffer, { at: 0 });
      return;
    }

    throw new Error('Unable to write to index handle.');
  }

  public getDataBytesWritten(): number {
    return this.dataBytesWritten;
  }

  /**
   * Flushes and closes all exclusive handles.
   * MUST be executed before commit and before runtime readers attempt access.
   */
  public async close(): Promise<void> {
    if (!this.isOpen) return;

    try {
      if (this.dataHandle) {
        if (this.dataHandle.flush) this.dataHandle.flush();
        if (this.dataHandle.close) this.dataHandle.close();
      }
    } catch (e) {
      console.warn('[DatasetWriter] Data handle close notice:', e);
    }

    try {
      if (this.indexHandle) {
        if (this.indexHandle.flush) this.indexHandle.flush();
        if (this.indexHandle.close) this.indexHandle.close();
      }
    } catch (e) {
      console.warn('[DatasetWriter] Index handle close notice:', e);
    }

    this.dataHandle = null;
    this.indexHandle = null;
    this.isOpen = false;
  }

  /**
   * Emergency cleanup on worker crash or abort.
   */
  public async abort(): Promise<void> {
    await this.close();
  }
}
