import fs from 'fs';
import path from 'path';
import crypto from 'crypto';
import { StringDecoder } from 'string_decoder';
import prisma from '@/lib/db/prisma';
import {
  getDatasetSpoolFilePath,
  getPartitionBinPath,
  getPartitionIdxPath,
  getDatasetPartitionDir,
} from '../storage/storagePaths';

export const PARTITION_ROW_LIMIT = 2_000_000;

export interface IngestionProgressCallback {
  (progress: {
    datasetId: string;
    rowsProcessed: number;
    currentPartition: number;
    bytesProcessed: number;
    bytesTotal: number;
    progress: number;
  }): void;
}

/**
 * High-performance bounded buffered writer for high-throughput sequential binary writes.
 * Keeps memory strictly bounded (e.g. 256KB or 64KB) while reducing millions of syscalls to thousands.
 */
class BoundedBufferedWriter {
  private fd: number;
  private buffer: Buffer;
  private offset: number = 0;

  constructor(fd: number, bufferSize: number = 256 * 1024) {
    this.fd = fd;
    this.buffer = Buffer.allocUnsafe(bufferSize);
  }

  write(data: Buffer) {
    if (data.length >= this.buffer.length) {
      this.flush();
      fs.writeSync(this.fd, data, 0, data.length);
      return;
    }
    if (this.offset + data.length > this.buffer.length) {
      this.flush();
    }
    data.copy(this.buffer, this.offset);
    this.offset += data.length;
  }

  flush() {
    if (this.offset > 0) {
      fs.writeSync(this.fd, this.buffer, 0, this.offset);
      this.offset = 0;
    }
  }

  close() {
    this.flush();
    fs.closeSync(this.fd);
  }
}

export class CsvImportWorker {
  /**
   * Process and partition a finalized CSV dataset in streaming fashion.
   * Zero full-file memory loading: uses strictly bounded 64KB read buffers.
   */
  static async processDataset(datasetId: string, onProgress?: IngestionProgressCallback) {
    const dataset = await prisma.dataset.findUnique({
      where: { id: datasetId },
    });

    if (!dataset) {
      throw new Error(`Dataset ${datasetId} not found`);
    }

    const sourceFilePath = dataset.spoolPath || getDatasetSpoolFilePath(datasetId, true);
    if (!fs.existsSync(sourceFilePath)) {
      throw new Error(`Source CSV file not found at ${sourceFilePath}`);
    }

    const totalFileBytes = fs.statSync(sourceFilePath).size;

    // Transition state: QUEUED -> PARSING
    await prisma.dataset.update({
      where: { id: datasetId },
      data: {
        status: 'PARSING',
        currentStage: 'STREAMING_CSV_PARTITIONS',
        progress: 0.0,
      },
    });

    const partitionDir = getDatasetPartitionDir(datasetId);

    // Clean up any stale incomplete .tmp files in partition directory
    const existingFiles = fs.readdirSync(partitionDir);
    for (const f of existingFiles) {
      if (f.endsWith('.tmp')) {
        try {
          fs.unlinkSync(path.join(partitionDir, f));
        } catch {}
      }
    }

    // Check existing committed partitions for crash recovery
    const committedPartitions = await prisma.datasetPartition.findMany({
      where: { datasetId, status: 'READY' },
      orderBy: { partitionNumber: 'asc' },
    });

    let skipUntilRow = 0;
    let nextPartitionNumber = 0;

    // Verify contiguous committed partitions from partition 0
    for (const cp of committedPartitions) {
      if (cp.partitionNumber === nextPartitionNumber) {
        const binPath = getPartitionBinPath(datasetId, cp.partitionNumber, false);
        const idxPath = getPartitionIdxPath(datasetId, cp.partitionNumber, false);
        if (fs.existsSync(binPath) && fs.existsSync(idxPath)) {
          const binStat = fs.statSync(binPath);
          const idxStat = fs.statSync(idxPath);
          const expectedIdxSize = (cp.rowCount + 1) * 8;
          if (BigInt(binStat.size) === cp.byteSize && idxStat.size === expectedIdxSize) {
            skipUntilRow += cp.rowCount;
            nextPartitionNumber++;
            continue;
          }
        }
      }
      break; // stop at first broken or non-contiguous partition
    }

    if (skipUntilRow > 0) {
      console.log(`[CsvImportWorker] Crash Recovery: Detected ${nextPartitionNumber} committed partitions (${skipUntilRow} rows). Fast-forwarding...`);
    }

    // Setup streaming reader with 64KB bounded buffer
    const CHUNK_SIZE = 64 * 1024;
    const sourceFd = fs.openSync(sourceFilePath, 'r');
    const readBuffer = Buffer.alloc(CHUNK_SIZE);
    const decoder = new StringDecoder('utf8');

    let bytesReadTotal = 0;
    let totalDataRows = 0;
    let headerParsed = false;
    let headerRowCells: string[] = (dataset.headerRow as string[]) || [];
    let leftoverLine = '';

    // Active partition writer state
    let activePartNum = nextPartitionNumber;
    let activeBinWriter: BoundedBufferedWriter | null = null;
    let activeIdxWriter: BoundedBufferedWriter | null = null;
    let activeBinBytes = BigInt(0);
    let activePartRows = 0;
    let activePartStartRow = BigInt(skipUntilRow);
    let activeBinHash: crypto.Hash | null = null;

    function openPartitionFiles(partNum: number, startRow: bigint) {
      activePartNum = partNum;
      activePartStartRow = startRow;
      activePartRows = 0;
      activeBinBytes = BigInt(0);
      activeBinHash = crypto.createHash('sha256');

      const binTmp = getPartitionBinPath(datasetId, partNum, true);
      const idxTmp = getPartitionIdxPath(datasetId, partNum, true);

      if (fs.existsSync(binTmp)) fs.unlinkSync(binTmp);
      if (fs.existsSync(idxTmp)) fs.unlinkSync(idxTmp);

      const binFd = fs.openSync(binTmp, 'w');
      const idxFd = fs.openSync(idxTmp, 'w');

      activeBinWriter = new BoundedBufferedWriter(binFd, 256 * 1024);
      activeIdxWriter = new BoundedBufferedWriter(idxFd, 64 * 1024);
    }

    async function commitActivePartition() {
      if (activeBinWriter === null || activeIdxWriter === null || activePartRows === 0) {
        return;
      }

      // Write terminal offset: total file size as uint64 at end of .idx (index[rowCount] = fileSize)
      const finalOffsetBuf = Buffer.alloc(8);
      finalOffsetBuf.writeBigUInt64LE(activeBinBytes, 0);
      activeIdxWriter.write(finalOffsetBuf);

      activeBinWriter.close();
      activeIdxWriter.close();
      activeBinWriter = null;
      activeIdxWriter = null;

      const checksum = activeBinHash ? activeBinHash.digest('hex').toLowerCase() : '';
      const binTmp = getPartitionBinPath(datasetId, activePartNum, true);
      const idxTmp = getPartitionIdxPath(datasetId, activePartNum, true);
      const binFinal = getPartitionBinPath(datasetId, activePartNum, false);
      const idxFinal = getPartitionIdxPath(datasetId, activePartNum, false);

      // Verify sizes before promotion
      const binStat = fs.statSync(binTmp);
      const idxStat = fs.statSync(idxTmp);
      if (BigInt(binStat.size) !== activeBinBytes) {
        throw new Error(`Partition ${activePartNum} tmp .bin size mismatch before commit: expected ${activeBinBytes}, got ${binStat.size}`);
      }
      const expectedIdxSize = (activePartRows + 1) * 8;
      if (idxStat.size !== expectedIdxSize) {
        throw new Error(`Partition ${activePartNum} tmp .idx size mismatch before commit: expected ${expectedIdxSize}, got ${idxStat.size}`);
      }

      if (fs.existsSync(binFinal)) fs.unlinkSync(binFinal);
      if (fs.existsSync(idxFinal)) fs.unlinkSync(idxFinal);

      // Atomic rename from .tmp to final
      fs.renameSync(binTmp, binFinal);
      fs.renameSync(idxTmp, idxFinal);

      // Commit partition metadata to PostgreSQL
      await prisma.datasetPartition.upsert({
        where: {
          datasetId_partitionNumber: {
            datasetId,
            partitionNumber: activePartNum,
          },
        },
        create: {
          datasetId,
          partitionNumber: activePartNum,
          stableRowStart: activePartStartRow,
          stableRowEnd: activePartStartRow + BigInt(activePartRows - 1),
          rowCount: activePartRows,
          dataFilePath: binFinal,
          indexFilePath: idxFinal,
          byteSize: activeBinBytes,
          checksum,
          status: 'READY',
          completedAt: new Date(),
        },
        update: {
          stableRowStart: activePartStartRow,
          stableRowEnd: activePartStartRow + BigInt(activePartRows - 1),
          rowCount: activePartRows,
          dataFilePath: binFinal,
          indexFilePath: idxFinal,
          byteSize: activeBinBytes,
          checksum,
          status: 'READY',
          completedAt: new Date(),
        },
      });

      console.log(`[CsvImportWorker] Partition ${activePartNum} committed: rows [${activePartStartRow}..${activePartStartRow + BigInt(activePartRows - 1)}], bytes ${activeBinBytes}`);
    }

    // Initialize active partition
    openPartitionFiles(activePartNum, activePartStartRow);

    let lastProgressUpdate = Date.now();
    const startTime = Date.now();

    try {
      let bytesRead: number;
      while ((bytesRead = fs.readSync(sourceFd, readBuffer, 0, CHUNK_SIZE, null)) > 0) {
        bytesReadTotal += bytesRead;
        const chunkStr = leftoverLine + decoder.write(readBuffer.subarray(0, bytesRead));
        const lines = chunkStr.split('\n');
        // The last element is potentially an incomplete line
        leftoverLine = lines.pop() || '';

        for (let i = 0; i < lines.length; i++) {
          const line = lines[i].replace(/\r$/, '');
          if (line.length === 0) continue;

          // 1. Header parsing (Row 0 of CSV)
          if (!headerParsed) {
            headerRowCells = CsvImportWorker.parseCsvLine(line);
            headerParsed = true;
            await prisma.dataset.update({
              where: { id: datasetId },
              data: {
                hasHeader: true,
                columnCount: headerRowCells.length,
                headerRow: headerRowCells,
              },
            });
            continue;
          }

          // 2. Data row processing
          // If recovering, skip rows that belong to already committed partitions
          if (totalDataRows < skipUntilRow) {
            totalDataRows++;
            continue;
          }

          const currentStableRow = totalDataRows;
          const targetPartition = Math.floor(currentStableRow / PARTITION_ROW_LIMIT);

          // If boundary reached (e.g. at 2,000,000 / 4,000,000), commit and rotate
          if (targetPartition !== activePartNum) {
            await commitActivePartition();
            openPartitionFiles(targetPartition, BigInt(currentStableRow));
          }

          // Write 8-byte LE offset to .idx
          const offsetBuf = Buffer.alloc(8);
          offsetBuf.writeBigUInt64LE(activeBinBytes, 0);
          activeIdxWriter!.write(offsetBuf);

          // Write row bytes + LF to .bin
          const rowBytes = Buffer.from(line + '\n', 'utf8');
          activeBinWriter!.write(rowBytes);
          activeBinHash!.update(rowBytes);

          activeBinBytes += BigInt(rowBytes.length);
          activePartRows++;
          totalDataRows++;

          // Throttled progress update (every 100,000 rows or 1.5s)
          const now = Date.now();
          if (totalDataRows % 100_000 === 0 || now - lastProgressUpdate > 1500) {
            lastProgressUpdate = now;
            const progress = Math.min(0.95, bytesReadTotal / totalFileBytes);
            await prisma.dataset.update({
              where: { id: datasetId },
              data: {
                rowCount: BigInt(totalDataRows),
                progress,
              },
            });
            if (onProgress) {
              onProgress({
                datasetId,
                rowsProcessed: totalDataRows,
                currentPartition: activePartNum,
                bytesProcessed: bytesReadTotal,
                bytesTotal: totalFileBytes,
                progress,
              });
            }
          }
        }
      }

      // Handle final leftover line if any
      leftoverLine += decoder.end();
      if (leftoverLine && leftoverLine.trim().length > 0) {
        const line = leftoverLine.replace(/\r$/, '');
        if (!headerParsed) {
          headerRowCells = CsvImportWorker.parseCsvLine(line);
          headerParsed = true;
        } else if (totalDataRows >= skipUntilRow) {
          const currentStableRow = totalDataRows;
          const targetPartition = Math.floor(currentStableRow / PARTITION_ROW_LIMIT);

          if (targetPartition !== activePartNum) {
            await commitActivePartition();
            openPartitionFiles(targetPartition, BigInt(currentStableRow));
          }

          const offsetBuf = Buffer.alloc(8);
          offsetBuf.writeBigUInt64LE(activeBinBytes, 0);
          activeIdxWriter!.write(offsetBuf);

          const rowBytes = Buffer.from(line + '\n', 'utf8');
          activeBinWriter!.write(rowBytes);
          activeBinHash!.update(rowBytes);

          activeBinBytes += BigInt(rowBytes.length);
          activePartRows++;
          totalDataRows++;
        }
      }

      // Commit final active partition
      if (activePartRows > 0) {
        await commitActivePartition();
      }

      // -----------------------------------------------------------------------
      // FINAL VERIFICATION PHASE
      // -----------------------------------------------------------------------
      await prisma.dataset.update({
        where: { id: datasetId },
        data: {
          status: 'VERIFYING',
          currentStage: 'VERIFYING_PARTITION_INTEGRITY',
        },
      });

      const verification = await CsvImportWorker.verifyDatasetPartitions(datasetId, totalDataRows);

      const totalDurationSec = (Date.now() - startTime) / 1000;
      console.log(`[CsvImportWorker] Verification passed: ${totalDataRows} rows across ${verification.partitionsCount} partitions in ${totalDurationSec.toFixed(2)}s`);

      // Mark dataset READY!
      const readyDataset = await prisma.dataset.update({
        where: { id: datasetId },
        data: {
          status: 'READY',
          currentStage: 'READY',
          progress: 1.0,
          rowCount: BigInt(totalDataRows),
          completedAt: new Date(),
        },
      });

      return {
        datasetId,
        status: readyDataset.status,
        totalDataRows,
        partitionsCount: verification.partitionsCount,
        columnCount: headerRowCells.length,
        durationSeconds: totalDurationSec,
        rowsPerSecond: Math.round(totalDataRows / Math.max(0.001, totalDurationSec)),
      };
    } catch (err: any) {
      if (activeBinWriter !== null) {
        try { activeBinWriter.close(); } catch {}
      }
      if (activeIdxWriter !== null) {
        try { activeIdxWriter.close(); } catch {}
      }

      await prisma.dataset.update({
        where: { id: datasetId },
        data: {
          status: 'FAILED',
          currentStage: 'FAILED',
          errorMessage: err.message || 'Ingestion failed',
        },
      });
      throw err;
    } finally {
      fs.closeSync(sourceFd);
    }
  }

  /**
   * RFC 4180 CSV line parser with quotes and commas
   */
  static parseCsvLine(line: string): string[] {
    if (!line || line.length === 0) return [''];
    const cells: string[] = [];
    const len = line.length;
    let inQuotes = false;
    let currentCell = '';
    let i = 0;

    while (i < len) {
      const char = line[i];
      if (char === '"') {
        if (inQuotes) {
          if (i + 1 < len && line[i + 1] === '"') {
            currentCell += '"';
            i += 2;
            continue;
          } else {
            inQuotes = false;
            i++;
            continue;
          }
        } else {
          inQuotes = true;
          i++;
          continue;
        }
      }
      if (char === ',' && !inQuotes) {
        cells.push(currentCell);
        currentCell = '';
        i++;
        continue;
      }
      currentCell += char;
      i++;
    }
    cells.push(currentCell);
    return cells;
  }

  /**
   * Verify all partitions of a dataset on disk against database metadata.
   */
  static async verifyDatasetPartitions(
    datasetId: string,
    expectedRowCount?: number
  ): Promise<{ verifiedRows: number; partitionsCount: number }> {
    const allPartitions = await prisma.datasetPartition.findMany({
      where: { datasetId, status: 'READY' },
      orderBy: { partitionNumber: 'asc' },
    });

    if (allPartitions.length === 0) {
      throw new Error(`No READY partitions found for dataset ${datasetId}`);
    }

    let verifiedRowsSum = 0;
    let expectedNextRow = BigInt(0);

    for (const p of allPartitions) {
      if (!fs.existsSync(p.dataFilePath) || !fs.existsSync(p.indexFilePath)) {
        throw new Error(`Partition ${p.partitionNumber} files missing on disk during verification`);
      }
      const binStat = fs.statSync(p.dataFilePath);
      const idxStat = fs.statSync(p.indexFilePath);

      if (BigInt(binStat.size) !== p.byteSize) {
        throw new Error(`Partition ${p.partitionNumber} byte size mismatch: DB ${p.byteSize}, disk ${binStat.size}`);
      }
      // idx size must be exactly (p.rowCount + 1) * 8
      const expectedIdxBytes = (p.rowCount + 1) * 8;
      if (idxStat.size !== expectedIdxBytes) {
        throw new Error(`Partition ${p.partitionNumber} index byte size mismatch: expected ${expectedIdxBytes}, got ${idxStat.size}`);
      }

      if (p.stableRowStart !== expectedNextRow) {
        throw new Error(`Partition ${p.partitionNumber} stableRowStart gap: expected ${expectedNextRow}, found ${p.stableRowStart}`);
      }
      expectedNextRow = p.stableRowEnd + BigInt(1);
      verifiedRowsSum += p.rowCount;
    }

    if (expectedRowCount !== undefined && verifiedRowsSum !== expectedRowCount) {
      throw new Error(`Row count reconciliation failed: sum of partitions (${verifiedRowsSum}) != expected rows (${expectedRowCount})`);
    }

    return { verifiedRows: verifiedRowsSum, partitionsCount: allPartitions.length };
  }
}
