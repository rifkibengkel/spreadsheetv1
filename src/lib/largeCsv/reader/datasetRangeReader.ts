import fs from 'fs';
import prisma from '@/lib/db/prisma';
import { parseCsvLine } from '../parser/csvLineParser';

export const MAX_RANGE_ROW_LIMIT = 5000;
export const MAX_RANGE_RESPONSE_BYTES = 10 * 1024 * 1024; // 10 MB limit
export const MAX_CONCURRENT_RANGE_QUERIES = 16;
export const MAX_QUEUE_DEPTH = 32;

/**
 * In-memory asynchronous semaphore to bound interactive range query concurrency
 * and prevent event loop/descriptor starvation under high concurrent load.
 */
class AsyncSemaphore {
  private currentActive = 0;
  private queue: Array<() => void> = [];

  constructor(
    private readonly maxConcurrent: number,
    private readonly maxQueue: number
  ) {}

  public async acquire(): Promise<() => void> {
    if (this.currentActive < this.maxConcurrent) {
      this.currentActive++;
      return () => this.release();
    }

    if (this.queue.length >= this.maxQueue) {
      throw new DatasetRangeError(
        'Server is experiencing high query load. Please retry in a moment.',
        429,
        'RANGE_QUERY_CONCURRENCY_LIMIT'
      );
    }

    return new Promise<() => void>((resolve) => {
      this.queue.push(() => {
        this.currentActive++;
        resolve(() => this.release());
      });
    });
  }

  private release(): void {
    this.currentActive--;
    if (this.queue.length > 0) {
      const next = this.queue.shift();
      if (next) next();
    }
  }

  public getActiveCount(): number {
    return this.currentActive;
  }

  public getQueueLength(): number {
    return this.queue.length;
  }
}

export const rangeQuerySemaphore = new AsyncSemaphore(
  MAX_CONCURRENT_RANGE_QUERIES,
  MAX_QUEUE_DEPTH
);

export interface DatasetRow {
  /** Stable 0-based data row index, guaranteed JSON-safe JavaScript number */
  stableRowId: number;
  cells: string[];
}

export interface DatasetRangeResult {
  datasetId: string;
  startRow: number;
  rowCount: number;
  totalRows: number;
  columns: string[];
  columnCount: number;
  rows: DatasetRow[];
  nextRow: number | null;
  metrics: {
    partitionResolutionMs: number;
    indexLookupMs: number;
    fileReadMs: number;
    decodeMs: number;
    serverTotalMs: number;
  };
}

export class DatasetRangeError extends Error {
  public statusCode: number;
  public code: string;

  constructor(message: string, statusCode = 400, code = 'INVALID_RANGE_REQUEST') {
    super(message);
    this.name = 'DatasetRangeError';
    this.statusCode = statusCode;
    this.code = code;
  }
}

export class DatasetRangeReader {
  /**
   * Reads a bounded row range across physical partitions with zero full-file loading
   * and non-blocking asynchronous filesystem I/O.
   */
  static async readRange(
    datasetId: string,
    startRow: number,
    requestedRowCount: number
  ): Promise<DatasetRangeResult> {
    const t0 = process.hrtime.bigint();

    // 1. Validation of parameters
    if (!datasetId || typeof datasetId !== 'string') {
      throw new DatasetRangeError('Invalid datasetId', 400, 'INVALID_DATASET_ID');
    }
    if (typeof startRow !== 'number' || isNaN(startRow) || !Number.isInteger(startRow) || startRow < 0) {
      throw new DatasetRangeError('startRow must be a non-negative integer', 400, 'INVALID_START_ROW');
    }
    if (
      typeof requestedRowCount !== 'number' ||
      isNaN(requestedRowCount) ||
      !Number.isInteger(requestedRowCount) ||
      requestedRowCount <= 0
    ) {
      throw new DatasetRangeError('rowCount must be a positive integer', 400, 'INVALID_ROW_COUNT');
    }
    if (requestedRowCount > MAX_RANGE_ROW_LIMIT) {
      throw new DatasetRangeError(
        `rowCount (${requestedRowCount}) exceeds maximum allowed limit of ${MAX_RANGE_ROW_LIMIT}`,
        400,
        'ROW_COUNT_EXCEEDS_LIMIT'
      );
    }

    // 2. Concurrency limiting via asynchronous semaphore
    const releaseConcurrencySlot = await rangeQuerySemaphore.acquire();

    try {
      // 3. T1: Partition resolution & metadata lookup
      const tPartition0 = process.hrtime.bigint();
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
        throw new DatasetRangeError(`Dataset ${datasetId} not found`, 404, 'DATASET_NOT_FOUND');
      }

      if (dataset.status !== 'READY') {
        throw new DatasetRangeError(
          `Dataset ${datasetId} is not in READY status (current status: ${dataset.status})`,
          409,
          'DATASET_NOT_READY'
        );
      }

      const totalRows = Number(dataset.rowCount);
      if (startRow >= totalRows) {
        throw new DatasetRangeError(
          `startRow (${startRow}) is beyond total dataset rows (${totalRows})`,
          400,
          'START_ROW_OUT_OF_BOUNDS'
        );
      }

      const effectiveRowCount = Math.min(requestedRowCount, totalRows - startRow);
      const endRow = startRow + effectiveRowCount - 1;

      const partitions = dataset.partitions;
      if (!partitions || partitions.length === 0) {
        throw new DatasetRangeError(
          `Dataset ${datasetId} has no committed READY partitions`,
          500,
          'NO_READY_PARTITIONS'
        );
      }

      // Find all overlapping partitions
      const overlappingPartitions = partitions.filter((p) => {
        const pStart = Number(p.stableRowStart);
        const pEnd = Number(p.stableRowEnd);
        return pEnd >= startRow && pStart <= endRow;
      });

      if (overlappingPartitions.length === 0) {
        throw new DatasetRangeError(
          `No partitions cover requested row range [${startRow}..${endRow}]`,
          500,
          'PARTITION_RESOLUTION_FAILED'
        );
      }

      const tPartition1 = process.hrtime.bigint();

      let totalIndexLookupNs = BigInt(0);
      let totalFileReadNs = BigInt(0);
      let totalDecodeNs = BigInt(0);

      // 4. Pre-check index byte ranges across partitions to guarantee MAX_RANGE_RESPONSE_BYTES
      interface PartitionSlicePlan {
        part: (typeof overlappingPartitions)[0];
        localStart: number;
        subStart: number;
        subEnd: number;
        subCount: number;
        idxBuf: Buffer;
        startByte: bigint;
        endByte: bigint;
        dataBytesToRead: number;
      }

      const slicePlans: PartitionSlicePlan[] = [];
      let cumulativeEstimatedBytes = 0;

      for (const part of overlappingPartitions) {
        const pStart = Number(part.stableRowStart);
        const pEnd = Number(part.stableRowEnd);

        const subStart = Math.max(startRow, pStart);
        const subEnd = Math.min(endRow, pEnd);
        const subCount = subEnd - subStart + 1;
        const localStart = subStart - pStart;

        // Verify physical file existence asynchronously
        try {
          await fs.promises.access(part.indexFilePath, fs.constants.R_OK);
        } catch {
          throw new DatasetRangeError(
            `Index file missing or unreadable on disk: ${part.indexFilePath}`,
            500,
            'INDEX_FILE_MISSING'
          );
        }
        try {
          await fs.promises.access(part.dataFilePath, fs.constants.R_OK);
        } catch {
          throw new DatasetRangeError(
            `Binary data file missing or unreadable on disk: ${part.dataFilePath}`,
            500,
            'DATA_FILE_MISSING'
          );
        }

        // T2: Asynchronous Index Lookup
        const tIndex0 = process.hrtime.bigint();
        const idxBytesToRead = (subCount + 1) * 8;
        const idxBuf = Buffer.allocUnsafe(idxBytesToRead);

        const idxHandle = await fs.promises.open(part.indexFilePath, 'r');
        try {
          const { bytesRead } = await idxHandle.read(idxBuf, 0, idxBytesToRead, localStart * 8);
          if (bytesRead !== idxBytesToRead) {
            throw new DatasetRangeError(
              `Corrupt index file: expected ${idxBytesToRead} bytes at offset ${localStart * 8}, read ${bytesRead}`,
              500,
              'CORRUPT_INDEX_FILE'
            );
          }
        } finally {
          await idxHandle.close();
        }
        const tIndex1 = process.hrtime.bigint();
        totalIndexLookupNs += tIndex1 - tIndex0;

        const startByte = idxBuf.readBigUInt64LE(0);
        const endByte = idxBuf.readBigUInt64LE(subCount * 8);

        // Sanity & corruption checks
        if (startByte > endByte) {
          throw new DatasetRangeError(
            `Corrupt index offset sequence: startByte (${startByte}) > endByte (${endByte}) in partition ${part.partitionNumber}`,
            500,
            'CORRUPT_INDEX_OFFSETS'
          );
        }
        if (endByte > part.byteSize) {
          throw new DatasetRangeError(
            `Index offset (${endByte}) exceeds partition file size (${part.byteSize}) in partition ${part.partitionNumber}`,
            500,
            'INDEX_OFFSET_OUT_OF_BOUNDS'
          );
        }

        const dataBytesToRead = Number(endByte - startByte);
        cumulativeEstimatedBytes += dataBytesToRead;

        // Hard response byte limit enforcement
        if (cumulativeEstimatedBytes > MAX_RANGE_RESPONSE_BYTES) {
          throw new DatasetRangeError(
            `Requested range estimated raw size (${cumulativeEstimatedBytes} bytes) exceeds maximum limit of ${MAX_RANGE_RESPONSE_BYTES} bytes. Please request a smaller rowCount.`,
            400,
            'RANGE_RESPONSE_BYTES_EXCEEDED'
          );
        }

        slicePlans.push({
          part,
          localStart,
          subStart,
          subEnd,
          subCount,
          idxBuf,
          startByte,
          endByte,
          dataBytesToRead,
        });
      }

      // 5. Asynchronous Bounded File Reads and Row Decoding
      const rows: DatasetRow[] = [];

      for (const plan of slicePlans) {
        const { part, subStart, subCount, idxBuf, startByte, dataBytesToRead } = plan;

        // T3: Asynchronous Bounded File Read
        const tRead0 = process.hrtime.bigint();
        const dataBuf = Buffer.allocUnsafe(dataBytesToRead);

        if (dataBytesToRead > 0) {
          const binHandle = await fs.promises.open(part.dataFilePath, 'r');
          try {
            const { bytesRead } = await binHandle.read(
              dataBuf,
              0,
              dataBytesToRead,
              Number(startByte)
            );
            if (bytesRead !== dataBytesToRead) {
              throw new DatasetRangeError(
                `Short file read: expected ${dataBytesToRead} bytes at offset ${startByte}, got ${bytesRead}`,
                500,
                'FILE_READ_FAILURE'
              );
            }
          } finally {
            await binHandle.close();
          }
        }
        const tRead1 = process.hrtime.bigint();
        totalFileReadNs += tRead1 - tRead0;

        // T4: Pure RFC 4180 Row Decode
        const tDecode0 = process.hrtime.bigint();
        for (let i = 0; i < subCount; i++) {
          const lineStart = Number(idxBuf.readBigUInt64LE(i * 8) - startByte);
          const lineEnd = Number(idxBuf.readBigUInt64LE((i + 1) * 8) - startByte);

          let rawLine = dataBuf.subarray(lineStart, lineEnd).toString('utf8');
          if (rawLine.endsWith('\n')) {
            rawLine = rawLine.slice(0, -1);
            if (rawLine.endsWith('\r')) {
              rawLine = rawLine.slice(0, -1);
            }
          }

          const cells = parseCsvLine(rawLine);
          rows.push({
            stableRowId: Number(subStart + i), // JSON-safe JavaScript number
            cells,
          });
        }
        const tDecode1 = process.hrtime.bigint();
        totalDecodeNs += tDecode1 - tDecode0;
      }

      const tEnd = process.hrtime.bigint();

      const columns = (dataset.headerRow as string[]) || [];
      const nextRow = endRow + 1 < totalRows ? endRow + 1 : null;

      return {
        datasetId,
        startRow,
        rowCount: rows.length,
        totalRows,
        columns,
        columnCount: dataset.columnCount || columns.length,
        rows,
        nextRow,
        metrics: {
          partitionResolutionMs: Number(tPartition1 - tPartition0) / 1_000_000,
          indexLookupMs: Number(totalIndexLookupNs) / 1_000_000,
          fileReadMs: Number(totalFileReadNs) / 1_000_000,
          decodeMs: Number(totalDecodeNs) / 1_000_000,
          serverTotalMs: Number(tEnd - t0) / 1_000_000,
        },
      };
    } finally {
      releaseConcurrencySlot();
    }
  }
}
