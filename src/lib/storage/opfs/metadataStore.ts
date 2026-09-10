/**
 * OPFS Metadata Store
 * Invariant 1 & 31: Enforces commit lifecycle and manages table_meta.json
 */

import { OPFS_META_FILE, OPFSMetadata, OPFS_MAGIC_V1 } from './types';

export class MetadataStore {
  /**
   * Retrieves the OPFS root directory handle.
   */
  public static async getRootDir(): Promise<FileSystemDirectoryHandle> {
    if (typeof navigator === 'undefined' || !navigator.storage?.getDirectory) {
      throw new Error('OPFS (navigator.storage.getDirectory) is not supported in this environment.');
    }
    return await navigator.storage.getDirectory();
  }

  /**
   * Reads and parses table_meta.json from OPFS.
   * Returns null if the file does not exist or cannot be parsed.
   */
  public static async read(root?: FileSystemDirectoryHandle): Promise<OPFSMetadata | null> {
    try {
      const dir = root || (await this.getRootDir());
      const fileHandle = await dir.getFileHandle(OPFS_META_FILE);
      const file = await fileHandle.getFile();
      const text = await file.text();
      const parsed = JSON.parse(text) as OPFSMetadata;

      if (parsed.magic !== OPFS_MAGIC_V1) {
        console.warn(`[MetadataStore] Invalid magic number: expected ${OPFS_MAGIC_V1}, got ${parsed.magic}`);
        return null;
      }

      return parsed;
    } catch (e: any) {
      // File does not exist or reading failed
      return null;
    }
  }

  /**
   * Atomically writes table_meta.json to OPFS.
   */
  public static async write(
    metadata: OPFSMetadata,
    root?: FileSystemDirectoryHandle
  ): Promise<void> {
    const dir = root || (await this.getRootDir());
    const fileHandle = await dir.getFileHandle(OPFS_META_FILE, { create: true });
    
    // In Worker with SyncAccessHandle if supported, or standard writable stream
    if ((fileHandle as any).createWritable) {
      const writable = await (fileHandle as any).createWritable();
      await writable.write(JSON.stringify(metadata, null, 2));
      await writable.close();
    } else if ((fileHandle as any).createSyncAccessHandle) {
      const syncHandle = await (fileHandle as any).createSyncAccessHandle();
      const encoded = new TextEncoder().encode(JSON.stringify(metadata, null, 2));
      syncHandle.truncate(0);
      syncHandle.write(encoded);
      syncHandle.flush();
      syncHandle.close();
    } else {
      throw new Error('Neither createWritable nor createSyncAccessHandle is available on file handle.');
    }
  }

  /**
   * Checks whether the metadata represents a valid COMMITTED dataset.
   */
  public static isCommitted(meta: OPFSMetadata | null): boolean {
    if (!meta) return false;
    return meta.status === 'COMMITTED' && meta.rowCount > 0 && meta.dataSizeBytes > 0;
  }

  /**
   * Removes all OPFS files for table storage.
   */
  public static async clearStorage(root?: FileSystemDirectoryHandle): Promise<void> {
    try {
      const dir = root || (await this.getRootDir());
      for (const name of [OPFS_META_FILE, 'table_data.bin', 'table_index.idx']) {
        try {
          await dir.removeEntry(name);
        } catch {}
      }
    } catch (e) {
      console.warn('[MetadataStore] clearStorage warning:', e);
    }
  }
}
