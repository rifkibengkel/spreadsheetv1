/**
 * Commit Manager & State Transition Controller
 * Invariant 31 & Sections 7, 8, 11: Guarantees atomic commit lifecycle and corruption defense.
 */

import { MetadataStore } from './metadataStore';
import { RowIndex } from './rowIndex';
import { OPFS_DATA_FILE, OPFS_INDEX_FILE, OPFSMetadata } from './types';

export class CommitManager {
  /**
   * Begins the import process and records the IMPORTING state in metadata.
   */
  public static async beginImport(
    fileName: string,
    root?: FileSystemDirectoryHandle
  ): Promise<OPFSMetadata> {
    const dir = root || (await MetadataStore.getRootDir());

    const sessionId =
      typeof crypto !== 'undefined' && crypto.randomUUID
        ? crypto.randomUUID()
        : `sess_${Date.now()}_${Math.random().toString(36).slice(2, 9)}`;

    const meta: OPFSMetadata = {
      magic: 'SHEET_OPFS_V1',
      version: 1,
      sessionId,
      fileName,
      rowCount: 0,
      columnCount: 0,
      headers: [],
      dataSizeBytes: 0,
      indexSizeBytes: 0,
      status: 'IMPORTING',
      createdAt: new Date().toISOString(),
    };

    await MetadataStore.write(meta, dir);
    return meta;
  }

  /**
   * Validates dataset files, index integrity, and monotonicity.
   * Commits metadata to COMMITTED only if all checks strictly pass.
   */
  public static async validateAndCommit(
    meta: OPFSMetadata,
    root?: FileSystemDirectoryHandle
  ): Promise<{ success: boolean; error?: string }> {
    const dir = root || (await MetadataStore.getRootDir());

    try {
      // 1. Verify table_data.bin exists and size matches
      const dataHandle = await dir.getFileHandle(OPFS_DATA_FILE);
      const dataFile = await dataHandle.getFile();
      if (dataFile.size !== meta.dataSizeBytes) {
        throw new Error(
          `Data file size mismatch: expected ${meta.dataSizeBytes} bytes, actual ${dataFile.size} bytes.`
        );
      }

      // 2. Verify table_index.idx exists and size matches (rowCount + 1) * 8
      const expectedIndexSize = (meta.rowCount + 1) * 8;
      const indexHandle = await dir.getFileHandle(OPFS_INDEX_FILE);
      const indexFile = await indexHandle.getFile();
      if (indexFile.size !== expectedIndexSize) {
        throw new Error(
          `Index file size mismatch: expected ${expectedIndexSize} bytes, actual ${indexFile.size} bytes.`
        );
      }

      // 3. Read index buffer and validate invariants
      const indexBuffer = await indexFile.arrayBuffer();
      const indexArray = RowIndex.fromBuffer(indexBuffer);
      const validation = RowIndex.validate(indexArray, meta.rowCount, meta.dataSizeBytes);

      if (!validation.valid) {
        throw new Error(`Index validation failed: ${validation.error}`);
      }

      // 4. Mark COMMITTED
      meta.status = 'COMMITTED';
      meta.committedAt = new Date().toISOString();
      meta.indexSizeBytes = expectedIndexSize;
      await MetadataStore.write(meta, dir);

      console.log(
        `[CommitManager] ✅ Dataset COMMITTED successfully: ${meta.rowCount} rows, ${meta.dataSizeBytes} bytes, index ${meta.indexSizeBytes} bytes.`
      );
      return { success: true };
    } catch (err: any) {
      const errorMsg = err.message || String(err);
      console.error(`[CommitManager] ❌ Commit validation FAILED:`, errorMsg);

      meta.status = 'FAILED';
      meta.errorMessage = errorMsg;
      try {
        await MetadataStore.write(meta, dir);
      } catch (writeErr) {
        console.warn('[CommitManager] Failed to write FAILED metadata status:', writeErr);
      }

      return { success: false, error: errorMsg };
    }
  }

  /**
   * Validates dataset state on session reopen or startup.
   */
  public static async checkDatasetStatus(
    root?: FileSystemDirectoryHandle
  ): Promise<{ ready: boolean; metadata: OPFSMetadata | null; reason?: string }> {
    const meta = await MetadataStore.read(root);
    if (!meta) {
      return { ready: false, metadata: null, reason: 'No OPFS dataset found.' };
    }

    if (meta.status !== 'COMMITTED') {
      return {
        ready: false,
        metadata: meta,
        reason: `Dataset is in incomplete state: ${meta.status}. (Error: ${meta.errorMessage || 'Interrupted import'})`,
      };
    }

    return { ready: true, metadata: meta };
  }
}
