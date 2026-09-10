/**
 * IndexedDB Fallback for OPFS (Architecture D)
 * 
 * Provides transparent FileSystemDirectoryHandle / FileSystemFileHandle
 * emulation via IndexedDB when native OPFS (navigator.storage.getDirectory)
 * is unavailable (e.g. non-secure contexts, HTTP over LAN IP, private mode, webviews).
 */

const DB_NAME = 'SheetLargeDataset_FallbackFS';
const DB_VERSION = 1;
const STORE_NAME = 'files';

function getIDBFactory(): IDBFactory | null {
  if (typeof self !== 'undefined' && self.indexedDB) return self.indexedDB;
  if (typeof window !== 'undefined' && window.indexedDB) return window.indexedDB;
  if (typeof globalThis !== 'undefined' && (globalThis as any).indexedDB) {
    return (globalThis as any).indexedDB;
  }
  return null;
}

let dbInstancePromise: Promise<IDBDatabase> | null = null;

function openFallbackDB(): Promise<IDBDatabase> {
  if (dbInstancePromise) return dbInstancePromise;

  dbInstancePromise = new Promise((resolve, reject) => {
    const idb = getIDBFactory();
    if (!idb) {
      reject(new Error('Neither OPFS nor IndexedDB is available in this environment.'));
      return;
    }

    const req = idb.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = (e: any) => {
      const db = e.target.result as IDBDatabase;
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        db.createObjectStore(STORE_NAME);
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error || new Error('Failed to open IndexedDB fallback storage.'));
  });

  return dbInstancePromise;
}

export async function idbGetFile(key: string): Promise<any> {
  const db = await openFallbackDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readonly');
    const req = tx.objectStore(STORE_NAME).get(key);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

export async function idbPutFile(key: string, value: any): Promise<void> {
  const db = await openFallbackDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readwrite');
    tx.objectStore(STORE_NAME).put(value, key);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

export async function idbDeleteFile(key: string): Promise<void> {
  const db = await openFallbackDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readwrite');
    tx.objectStore(STORE_NAME).delete(key);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

export class IDBVirtualFileHandle {
  public name: string;

  constructor(name: string) {
    this.name = name;
  }

  public async getFile(): Promise<File | Blob> {
    const data = await idbGetFile(this.name);
    if (data === undefined || data === null) {
      throw new Error(`NotFoundError: File ${this.name} does not exist in fallback storage.`);
    }

    if (data instanceof File) return data;
    if (data instanceof Blob) return data;

    if (data instanceof ArrayBuffer) {
      return new Blob([data], { type: 'application/octet-stream' });
    }

    if (typeof data === 'string') {
      return new File([data], this.name, { type: 'application/json' });
    }

    return new Blob([data]);
  }

  public async createSyncAccessHandle(): Promise<any> {
    return this.createWritableHandle();
  }

  public async createWritable(): Promise<any> {
    return this.createWritableHandle();
  }

  private createWritableHandle(): any {
    const fileName = this.name;
    let chunks: Uint8Array[] = [];
    let textContent: string | null = null;
    let bytesWritten = 0;

    return {
      write(chunk: any, options?: { at?: number }) {
        if (typeof chunk === 'string') {
          textContent = (textContent || '') + chunk;
          bytesWritten += chunk.length;
          return chunk.length;
        }

        const buf = chunk instanceof Uint8Array ? chunk : new Uint8Array(chunk);
        chunks.push(buf);
        bytesWritten += buf.byteLength;
        return buf.byteLength;
      },
      truncate(size: number) {
        chunks = [];
        textContent = null;
        bytesWritten = 0;
      },
      flush() {},
      async close() {
        if (textContent !== null) {
          await idbPutFile(fileName, textContent);
        } else if (fileName.endsWith('.json')) {
          const decoder = new TextDecoder('utf-8');
          let str = '';
          for (const c of chunks) str += decoder.decode(c, { stream: true });
          str += decoder.decode();
          await idbPutFile(fileName, str);
        } else if (fileName === 'table_data.bin') {
          const blob = new Blob(chunks as any, { type: 'application/octet-stream' });
          await idbPutFile(fileName, blob);
        } else if (fileName === 'table_index.idx') {
          const totalLen = chunks.reduce((acc, c) => acc + c.byteLength, 0);
          const merged = new Uint8Array(totalLen);
          let offset = 0;
          for (const c of chunks) {
            merged.set(c, offset);
            offset += c.byteLength;
          }
          await idbPutFile(fileName, merged.buffer);
        } else {
          const blob = new Blob(chunks as any);
          await idbPutFile(fileName, blob);
        }
        chunks = [];
      },
    };
  }
}

export class IDBVirtualDirectoryHandle {
  public async getFileHandle(
    name: string,
    options?: { create?: boolean }
  ): Promise<IDBVirtualFileHandle> {
    if (!options?.create) {
      const exists = (await idbGetFile(name)) !== undefined;
      if (!exists) {
        throw new Error(`NotFoundError: File ${name} not found in fallback storage.`);
      }
    }
    return new IDBVirtualFileHandle(name);
  }

  public async removeEntry(name: string): Promise<void> {
    await idbDeleteFile(name);
  }
}

let fallbackRootDirInstance: IDBVirtualDirectoryHandle | null = null;

export async function getIDBDirectoryHandle(): Promise<IDBVirtualDirectoryHandle> {
  if (!fallbackRootDirInstance) {
    fallbackRootDirInstance = new IDBVirtualDirectoryHandle();
  }
  return fallbackRootDirInstance;
}
