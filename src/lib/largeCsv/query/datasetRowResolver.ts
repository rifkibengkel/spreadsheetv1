import fs from 'fs';
import prisma from '@/lib/db/prisma';
import { parseCsvLine } from '../parser/csvLineParser';
import { DatasetRow, DatasetRangeError } from '../reader/datasetRangeReader';
import { DatasetQueryError } from './types';

interface PartitionItem {
  reqIdx: number;
  stableRowId: number;
  localIndex: number;
}

interface RowSpan {
  localStart: number;
  count: number;
  items: PartitionItem[];
}

export class DatasetRowResolver {
  /**
   * Resolves arbitrary stableRowId sequences back to physical CSV cells
   * using bounded index lookups and span clustering.
   * Guarantees asynchronous, descriptor-safe I/O and exact ordering preservation.
   */
  public static async resolveRows(
    datasetId: string,
    stableRowIds: number[]
  ): Promise<DatasetRow[]> {
    if (!stableRowIds || stableRowIds.length === 0) {
      return [];
    }

    // 1. Fetch partition metadata
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

    const partitions = dataset.partitions;
    if (!partitions || partitions.length === 0) {
      throw new DatasetQueryError(
        `Dataset ${datasetId} has no committed partitions`,
        500,
        'NO_READY_PARTITIONS'
      );
    }

    // 2. Map requested IDs to partitions
    const partitionMap = new Map<number, PartitionItem[]>();
    for (let i = 0; i < stableRowIds.length; i++) {
      const rowId = stableRowIds[i];
      let foundPart: (typeof partitions)[0] | null = null;

      for (const p of partitions) {
        if (rowId >= Number(p.stableRowStart) && rowId <= Number(p.stableRowEnd)) {
          foundPart = p;
          break;
        }
      }

      if (!foundPart) {
        throw new DatasetQueryError(
          `stableRowId ${rowId} does not belong to any partition in dataset ${datasetId}`,
          400,
          'ROW_ID_OUT_OF_BOUNDS'
        );
      }

      const pNum = foundPart.partitionNumber;
      if (!partitionMap.has(pNum)) {
        partitionMap.set(pNum, []);
      }

      partitionMap.get(pNum)!.push({
        reqIdx: i,
        stableRowId: rowId,
        localIndex: rowId - Number(foundPart.stableRowStart),
      });
    }

    const resultRows: DatasetRow[] = new Array(stableRowIds.length);

    // 3. Process each partition with span clustering
    for (const part of partitions) {
      const items = partitionMap.get(part.partitionNumber);
      if (!items || items.length === 0) continue;

      // Sort items by localIndex to enable span clustering and sequential disk access
      items.sort((a, b) => a.localIndex - b.localIndex);

      // Cluster adjacent items into contiguous spans
      const spans: RowSpan[] = [];
      let currentSpan: RowSpan | null = null;

      for (const item of items) {
        if (!currentSpan) {
          currentSpan = { localStart: item.localIndex, count: 1, items: [item] };
        } else if (item.localIndex === currentSpan.localStart + currentSpan.count) {
          // Contiguous extension!
          currentSpan.count++;
          currentSpan.items.push(item);
        } else {
          spans.push(currentSpan);
          currentSpan = { localStart: item.localIndex, count: 1, items: [item] };
        }
      }
      if (currentSpan) {
        spans.push(currentSpan);
      }

      // Read spans from partition files
      let idxHandle: fs.promises.FileHandle | null = null;
      let binHandle: fs.promises.FileHandle | null = null;

      try {
        idxHandle = await fs.promises.open(part.indexFilePath, 'r');
        binHandle = await fs.promises.open(part.dataFilePath, 'r');

        for (const span of spans) {
          const idxBytesToRead = (span.count + 1) * 8;
          const idxBuf = Buffer.allocUnsafe(idxBytesToRead);

          const { bytesRead: idxRead } = await idxHandle.read(
            idxBuf,
            0,
            idxBytesToRead,
            span.localStart * 8
          );

          if (idxRead !== idxBytesToRead) {
            throw new DatasetQueryError(
              `Index read failure in partition ${part.partitionNumber}`,
              500,
              'CORRUPT_INDEX_FILE'
            );
          }

          const startByte = idxBuf.readBigUInt64LE(0);
          const endByte = idxBuf.readBigUInt64LE(span.count * 8);
          const dataBytesToRead = Number(endByte - startByte);

          const dataBuf = Buffer.allocUnsafe(dataBytesToRead);
          if (dataBytesToRead > 0) {
            const { bytesRead: binRead } = await binHandle.read(
              dataBuf,
              0,
              dataBytesToRead,
              Number(startByte)
            );

            if (binRead !== dataBytesToRead) {
              throw new DatasetQueryError(
                `Data read failure in partition ${part.partitionNumber}`,
                500,
                'CORRUPT_DATA_FILE'
              );
            }
          }

          // Decode each row in the span
          for (let j = 0; j < span.count; j++) {
            const lineStart = Number(idxBuf.readBigUInt64LE(j * 8) - startByte);
            const lineEnd = Number(idxBuf.readBigUInt64LE((j + 1) * 8) - startByte);

            let rawLine = dataBuf.subarray(lineStart, lineEnd).toString('utf8');
            if (rawLine.endsWith('\n')) {
              rawLine = rawLine.slice(0, -1);
              if (rawLine.endsWith('\r')) {
                rawLine = rawLine.slice(0, -1);
              }
            }

            const cells = parseCsvLine(rawLine);
            const targetItem = span.items[j];
            resultRows[targetItem.reqIdx] = {
              stableRowId: targetItem.stableRowId,
              cells,
            };
          }
        }
      } finally {
        if (idxHandle) {
          try {
            await idxHandle.close();
          } catch {}
        }
        if (binHandle) {
          try {
            await binHandle.close();
          } catch {}
        }
      }
    }

    return resultRows;
  }
}
