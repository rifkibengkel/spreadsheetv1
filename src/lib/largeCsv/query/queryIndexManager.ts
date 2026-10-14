import fs from 'fs';
import path from 'path';
import crypto from 'crypto';
import prisma from '@/lib/db/prisma';
import {
  QUERY_INDEX_MAGIC,
  QUERY_INDEX_FORMAT_VERSION,
  QUERY_INDEX_HEADER_SIZE,
  MAX_CONCURRENT_QUERY_JOBS,
  QueryDefinition,
  QueryMetadata,
  DatasetQueryError,
} from './types';
import { FilterEngine } from './filterEngine';
import { ExternalSortEngine } from './externalSortEngine';

class AsyncQuerySemaphore {
  private active = 0;
  private queue: Array<() => void> = [];

  constructor(private readonly maxConcurrent: number, private readonly maxQueue: number) {}

  public async acquire(): Promise<() => void> {
    if (this.active < this.maxConcurrent) {
      this.active++;
      return () => this.release();
    }

    if (this.queue.length >= this.maxQueue) {
      throw new DatasetQueryError(
        'Server is experiencing high query compilation load. Please retry shortly.',
        429,
        'QUERY_CONCURRENCY_LIMIT'
      );
    }

    return new Promise<() => void>((resolve) => {
      this.queue.push(() => {
        this.active++;
        resolve(() => this.release());
      });
    });
  }

  private release(): void {
    this.active--;
    if (this.queue.length > 0) {
      const next = this.queue.shift();
      if (next) next();
    }
  }
}

const queryJobSemaphore = new AsyncQuerySemaphore(MAX_CONCURRENT_QUERY_JOBS, 16);

export class QueryIndexManager {
  /**
   * Resolves storage directory for dataset query indexes.
   */
  public static getQueriesDir(datasetId: string): string {
    return path.join(process.cwd(), 'storage', 'datasets', 'queries', datasetId);
  }

  /**
   * Deterministically generates a query ID from dataset ID and query definition.
   */
  public static generateQueryId(datasetId: string, queryDef: QueryDefinition): string {
    const normalized = JSON.stringify({
      datasetId,
      filter: queryDef.filter || null,
      sort: queryDef.sort || null,
    });
    const hash = crypto.createHash('sha256').update(normalized).digest('hex');
    return `q_${hash.slice(0, 32)}`;
  }

  /**
   * Returns metadata for an existing query index if it exists and is valid.
   */
  public static async getQueryMetadata(
    datasetId: string,
    queryId: string
  ): Promise<QueryMetadata | null> {
    const dir = this.getQueriesDir(datasetId);
    const metaPath = path.join(dir, `query_${queryId}.meta.json`);
    const qidxPath = path.join(dir, `query_${queryId}.qidx`);

    try {
      await fs.promises.access(metaPath, fs.constants.R_OK);
      await fs.promises.access(qidxPath, fs.constants.R_OK);

      const content = await fs.promises.readFile(metaPath, 'utf8');
      const meta: QueryMetadata = JSON.parse(content);

      if (meta.status !== 'READY') return null;

      // Validate .qidx file on disk
      const isValid = await this.validateIndexFile(qidxPath, meta.matchedRows);
      if (!isValid) {
        return null;
      }

      return meta;
    } catch {
      return null;
    }
  }

  /**
   * Validates binary index file integrity.
   */
  public static async validateIndexFile(
    qidxPath: string,
    expectedMatchedRows: number
  ): Promise<boolean> {
    try {
      const stats = await fs.promises.stat(qidxPath);
      const expectedSize = QUERY_INDEX_HEADER_SIZE + expectedMatchedRows * 4;
      if (stats.size !== expectedSize) {
        return false;
      }

      const handle = await fs.promises.open(qidxPath, 'r');
      try {
        const headerBuf = Buffer.allocUnsafe(32);
        await handle.read(headerBuf, 0, 32, 0);

        const magic = headerBuf.subarray(0, 16).toString('ascii').replace(/\0+$/, '');
        if (magic !== QUERY_INDEX_MAGIC) {
          return false;
        }

        const formatVersion = headerBuf.readUInt32LE(16);
        if (formatVersion !== QUERY_INDEX_FORMAT_VERSION) {
          return false;
        }

        const status = headerBuf.readUInt32LE(20);
        if (status !== 1) {
          // 1 = READY
          return false;
        }

        return true;
      } finally {
        await handle.close();
      }
    } catch {
      return false;
    }
  }

  /**
   * Executes or retrieves a query index.
   * Atomic promotion: tmp_query_<id>_<timestamp>.qidx -> validate -> rename to query_<id>.qidx.
   */
  public static async executeQuery(
    datasetId: string,
    queryDef: QueryDefinition,
    options: { abortSignal?: AbortSignal } = {}
  ): Promise<QueryMetadata> {
    const queryId = this.generateQueryId(datasetId, queryDef);

    // 1. Check if valid query index already exists (cache hit)
    const existing = await this.getQueryMetadata(datasetId, queryId);
    if (existing) {
      return existing;
    }

    // 2. Concurrency limiting for index compilation
    const releaseConcurrency = await queryJobSemaphore.acquire();

    const tStart = Date.now();
    const queriesDir = this.getQueriesDir(datasetId);
    await fs.promises.mkdir(queriesDir, { recursive: true });

    const timestamp = Date.now();
    const tmpQidxPath = path.join(queriesDir, `tmp_query_${queryId}_${timestamp}.qidx`);
    const tmpMetaPath = path.join(queriesDir, `tmp_query_${queryId}_${timestamp}.meta.json`);
    const finalQidxPath = path.join(queriesDir, `query_${queryId}.qidx`);
    const finalMetaPath = path.join(queriesDir, `query_${queryId}.meta.json`);
    const tempSortDir = path.join(queriesDir, `tmp_sort_${queryId}_${timestamp}`);

    let qidxHandle: fs.promises.FileHandle | null = null;

    try {
      // 3. Fetch dataset and partitions from database
      const dataset = await prisma.dataset.findUnique({
        where: { id: datasetId },
        include: {
          partitions: {
            where: { status: 'READY' },
            orderBy: { partitionNumber: 'asc' },
          },
        },
      });

      if (!dataset) {
        throw new DatasetQueryError(`Dataset ${datasetId} not found`, 404, 'DATASET_NOT_FOUND');
      }
      if (dataset.status !== 'READY') {
        throw new DatasetQueryError(
          `Dataset ${datasetId} is not in READY status`,
          409,
          'DATASET_NOT_READY'
        );
      }

      const totalDatasetRows = Number(dataset.rowCount);
      const headers = (dataset.headerRow as string[]) || [];

      // Compile filter
      const compiledFilter = FilterEngine.compileFilter(queryDef.filter, headers);

      // Compile sort columns
      const hasSort = queryDef.sort && queryDef.sort.columns && queryDef.sort.columns.length > 0;
      const sortColumnIndices: number[] = [];
      if (hasSort && queryDef.sort) {
        const headerMap = new Map<string, number>();
        for (let i = 0; i < headers.length; i++) headerMap.set(headers[i], i);

        for (const col of queryDef.sort.columns) {
          const idx = headerMap.get(col.column);
          if (idx === undefined) {
            throw new DatasetQueryError(
              `Sort column "${col.column}" not found in dataset`,
              400,
              'SORT_COLUMN_NOT_FOUND'
            );
          }
          sortColumnIndices.push(idx);
        }
      }

      // 4. Initialize temporary .qidx file with blank 2048-byte header
      qidxHandle = await fs.promises.open(tmpQidxPath, 'w+');
      const headerBuf = Buffer.alloc(QUERY_INDEX_HEADER_SIZE, 0);

      // Write preliminary header
      headerBuf.write(QUERY_INDEX_MAGIC, 0, 16, 'ascii');
      headerBuf.writeUInt32LE(QUERY_INDEX_FORMAT_VERSION, 16);
      headerBuf.writeUInt32LE(0, 20); // 0 = GENERATING
      headerBuf.write(datasetId.slice(0, 36), 24, 36, 'ascii');
      headerBuf.writeBigUInt64LE(BigInt(totalDatasetRows), 60);
      headerBuf.writeBigUInt64LE(BigInt(0), 68); // matchedRows (preliminary)
      headerBuf.writeBigInt64LE(BigInt(timestamp), 76); // createdAt
      headerBuf.writeBigInt64LE(BigInt(0), 84); // completedAt

      const queryDefJson = JSON.stringify(queryDef);
      const jsonBytes = Buffer.from(queryDefJson, 'utf8');
      headerBuf.writeUInt32LE(jsonBytes.length, 92);
      jsonBytes.copy(headerBuf, 96, 0, Math.min(jsonBytes.length, QUERY_INDEX_HEADER_SIZE - 96));

      await qidxHandle.write(headerBuf, 0, QUERY_INDEX_HEADER_SIZE, 0);

      let matchedRows = 0;

      // 5. Execution
      const BUFFER_CAPACITY = 16384; // 16,384 uint32s = 64 KB buffer
      const writeBuffer = Buffer.allocUnsafe(BUFFER_CAPACITY * 4);
      let bufferOffset = 0;
      let currentFileOffset = QUERY_INDEX_HEADER_SIZE;

      const flushWriteBuffer = async () => {
        if (bufferOffset > 0 && qidxHandle) {
          const bytesToWrite = bufferOffset * 4;
          await qidxHandle.write(writeBuffer, 0, bytesToWrite, currentFileOffset);
          currentFileOffset += bytesToWrite;
          bufferOffset = 0;
        }
      };

      const writeStableRowId = async (id: number) => {
        writeBuffer.writeUInt32LE(id, bufferOffset * 4);
        bufferOffset++;
        matchedRows++;
        if (bufferOffset >= BUFFER_CAPACITY) {
          await flushWriteBuffer();
        }
      };

      if (!hasSort) {
        // Stream directly to index
        for (const part of dataset.partitions) {
          await FilterEngine.scanPartition(
            part,
            compiledFilter.compiled,
            compiledFilter.conjunction,
            async (stableRowId) => {
              await writeStableRowId(stableRowId);
            },
            { abortSignal: options.abortSignal }
          );
        }
        await flushWriteBuffer();
      } else {
        // Sort requested: stream matches into ExternalSortEngine
        const sorter = new ExternalSortEngine(queryDef.sort!, tempSortDir);

        for (const part of dataset.partitions) {
          await FilterEngine.scanPartition(
            part,
            compiledFilter.compiled,
            compiledFilter.conjunction,
            async (stableRowId, cells) => {
              const sortKeys: (string | number)[] = new Array(sortColumnIndices.length);
              for (let k = 0; k < sortColumnIndices.length; k++) {
                const cIdx = sortColumnIndices[k];
                sortKeys[k] = cIdx < cells.length ? cells[cIdx] : '';
              }
              await sorter.addItem({ stableRowId, sortKeys });
            },
            { abortSignal: options.abortSignal }
          );
        }

        // Stream sorted IDs to index
        await sorter.streamSortedIds(async (sortedId) => {
          await writeStableRowId(sortedId);
        });
        await flushWriteBuffer();
      }

      // 6. Finalize header with READY status & matchedRows
      const tCompleted = Date.now();
      headerBuf.writeUInt32LE(1, 20); // 1 = READY
      headerBuf.writeBigUInt64LE(BigInt(matchedRows), 68);
      headerBuf.writeBigInt64LE(BigInt(tCompleted), 84); // completedAt
      await qidxHandle.write(headerBuf, 0, QUERY_INDEX_HEADER_SIZE, 0);

      await qidxHandle.close();
      qidxHandle = null;

      // 7. Validate generated file
      const isValid = await this.validateIndexFile(tmpQidxPath, matchedRows);
      if (!isValid) {
        throw new DatasetQueryError(
          'Generated query index failed post-generation validation check',
          500,
          'QUERY_INDEX_VALIDATION_FAILED'
        );
      }

      // 8. Atomic Promotion
      await fs.promises.rename(tmpQidxPath, finalQidxPath);

      const meta: QueryMetadata = {
        queryId,
        datasetId,
        status: 'READY',
        totalDatasetRows,
        matchedRows,
        indexFilePath: finalQidxPath,
        metadataFilePath: finalMetaPath,
        createdAt: timestamp,
        completedAt: tCompleted,
        executionTimeMs: tCompleted - tStart,
        queryDef,
      };

      await fs.promises.writeFile(tmpMetaPath, JSON.stringify(meta, null, 2), 'utf8');
      await fs.promises.rename(tmpMetaPath, finalMetaPath);

      return meta;
    } catch (err: any) {
      // Clean up temporary files on error
      if (qidxHandle) {
        try {
          await qidxHandle.close();
        } catch {}
      }
      try {
        await fs.promises.unlink(tmpQidxPath);
      } catch {}
      try {
        await fs.promises.unlink(tmpMetaPath);
      } catch {}
      try {
        await fs.promises.rm(tempSortDir, { recursive: true, force: true });
      } catch {}

      throw err;
    } finally {
      releaseConcurrency();
    }
  }

  /**
   * Bounded asynchronous read of stableRowIds from query index file.
   * Offset: 2048 + startRow * 4, Length: rowCount * 4.
   */
  public static async readStableRowIds(
    datasetId: string,
    queryId: string,
    startRow: number,
    requestedRowCount: number
  ): Promise<{ stableRowIds: number[]; totalMatchedRows: number }> {
    const meta = await this.getQueryMetadata(datasetId, queryId);
    if (!meta) {
      throw new DatasetQueryError(
        `Query ${queryId} for dataset ${datasetId} not found or not in READY status`,
        404,
        'QUERY_NOT_FOUND'
      );
    }

    const totalMatchedRows = meta.matchedRows;
    if (totalMatchedRows === 0 || startRow >= totalMatchedRows) {
      return { stableRowIds: [], totalMatchedRows };
    }

    const effectiveRowCount = Math.min(requestedRowCount, totalMatchedRows - startRow);
    const bytesToRead = effectiveRowCount * 4;
    const fileOffset = QUERY_INDEX_HEADER_SIZE + startRow * 4;

    const buf = Buffer.allocUnsafe(bytesToRead);
    const handle = await fs.promises.open(meta.indexFilePath, 'r');
    try {
      const { bytesRead } = await handle.read(buf, 0, bytesToRead, fileOffset);
      if (bytesRead !== bytesToRead) {
        throw new DatasetQueryError(
          `Short read from query index ${queryId}: expected ${bytesToRead} bytes, got ${bytesRead}`,
          500,
          'CORRUPT_QUERY_INDEX'
        );
      }

      const stableRowIds: number[] = new Array(effectiveRowCount);
      for (let i = 0; i < effectiveRowCount; i++) {
        stableRowIds[i] = buf.readUInt32LE(i * 4);
      }

      return { stableRowIds, totalMatchedRows };
    } finally {
      await handle.close();
    }
  }

  /**
   * Safely deletes query index files.
   */
  public static async deleteQuery(datasetId: string, queryId: string): Promise<boolean> {
    const dir = this.getQueriesDir(datasetId);
    const qidxPath = path.join(dir, `query_${queryId}.qidx`);
    const metaPath = path.join(dir, `query_${queryId}.meta.json`);

    try {
      await fs.promises.unlink(qidxPath);
    } catch {}
    try {
      await fs.promises.unlink(metaPath);
    } catch {}
    return true;
  }
}
