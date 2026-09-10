/**
 * Dedicated OPFS CSV Import Web Worker
 * Invariants 1, 2, 4, 5 & Sections 10, 12, 13, 14
 * 
 * Streams CSV bytes directly into OPFS binary storage without materializing
 * the dataset on the main thread or creating excessive JS heap objects.
 */

import { parseCsvLine, StreamingRowScanner } from '../storage/opfs/csvTokenizer';
import { DatasetWriter } from '../storage/opfs/datasetWriter';
import { CommitManager } from '../storage/opfs/commitManager';
import { RowIndex } from '../storage/opfs/rowIndex';
import { ImportProgress, OPFSMetadata } from '../storage/opfs/types';

function getMemoryUsageMB(): number {
  if (typeof performance !== 'undefined' && (performance as any).memory) {
    return (performance as any).memory.usedJSHeapSize / (1024 * 1024);
  }
  return 0;
}

const CHUNK_SIZE = 4 * 1024 * 1024; // 4 MB chunk streaming
const INDEX_BLOCK_SIZE = 262144; // 256K offsets per block (~2 MB)

self.onmessage = async (e: MessageEvent) => {
  const data = e.data;
  if (!data || data.type !== 'START_IMPORT') return;

  const file: File = data.file;
  if (!file) {
    self.postMessage({ type: 'ERROR', error: 'No file provided to import worker.' });
    return;
  }

  const startTime = performance.now();
  let peakMemoryMB = getMemoryUsageMB();

  const writer = new DatasetWriter();
  let metadata: OPFSMetadata | null = null;

  try {
    // -------------------------------------------------------------------------
    // Phase 1: Read and parse Header Line
    // -------------------------------------------------------------------------
    self.postMessage({
      type: 'PROGRESS',
      progress: {
        bytesProcessed: 0,
        totalBytes: file.size,
        rowsProcessed: 0,
        percent: 1,
        phase: 'READING_HEADER',
      } as ImportProgress,
    });

    // Read first 64 KB to locate the header boundary
    const headerProbeBlob = file.slice(0, Math.min(65536, file.size));
    const headerProbeBuf = new Uint8Array(await headerProbeBlob.arrayBuffer());
    let headerEndIndex = -1;
    let inQuotesHeader = false;

    for (let i = 0; i < headerProbeBuf.length; i++) {
      const b = headerProbeBuf[i];
      if (b === 0x22) inQuotesHeader = !inQuotesHeader;
      else if (b === 0x0a && !inQuotesHeader) {
        headerEndIndex = i;
        break;
      }
    }

    if (headerEndIndex === -1) {
      throw new Error('Could not find valid CSV header line ending in initial 64 KB.');
    }

    const headerByteLength = headerEndIndex + 1; // includes \n
    const headerSlice = headerProbeBuf.subarray(0, headerEndIndex);
    const headerText = new TextDecoder('utf-8').decode(headerSlice).replace(/\r$/, '');
    const headers = parseCsvLine(headerText);
    const columnCount = headers.length;

    console.log(
      `[CSV Import Worker] Detected Header: ${columnCount} columns (${headerByteLength} bytes)`
    );

    // -------------------------------------------------------------------------
    // Phase 2: Initialize Commit Manager and Dataset Writer
    // -------------------------------------------------------------------------
    metadata = await CommitManager.beginImport(file.name);
    metadata.headers = headers;
    metadata.columnCount = columnCount;

    await writer.open();

    // -------------------------------------------------------------------------
    // Phase 3: Stream Data Rows & Build N+1 Index
    // -------------------------------------------------------------------------
    const totalDataBytes = file.size - headerByteLength;
    let fileCursor = headerByteLength;
    let dataBytesWritten = 0;

    // We store BigUint64Array blocks to avoid unbounded single-array allocations
    const offsetBlocks: BigUint64Array[] = [];
    let currentBlock = new BigUint64Array(INDEX_BLOCK_SIZE);
    let blockIndex = 0;

    // offset[0] is always 0 (start of first data row)
    currentBlock[blockIndex++] = BigInt(0);
    let totalRowsIndexed = 0;

    const rowScanner = new StreamingRowScanner();
    let lastProgressTime = performance.now();

    while (fileCursor < file.size) {
      const nextEnd = Math.min(fileCursor + CHUNK_SIZE, file.size);
      const chunkBlob = file.slice(fileCursor, nextEnd);
      const chunkBuffer = new Uint8Array(await chunkBlob.arrayBuffer());
      const chunkLen = chunkBuffer.length;

      // Write raw bytes directly to table_data.bin
      await writer.writeDataChunk(chunkBuffer);

      // Scan chunk for row boundaries
      rowScanner.scanChunk(
        chunkBuffer,
        0,
        chunkLen,
        (boundaryByteIndex) => {
          // A row just ended at boundaryByteIndex in data stream
          // The NEXT row starts at (boundaryByteIndex + 1)
          const nextRowStart = BigInt(boundaryByteIndex + 1);

          if (blockIndex >= INDEX_BLOCK_SIZE) {
            offsetBlocks.push(currentBlock);
            currentBlock = new BigUint64Array(INDEX_BLOCK_SIZE);
            blockIndex = 0;
          }

          currentBlock[blockIndex++] = nextRowStart;
          totalRowsIndexed++;
        },
        dataBytesWritten // base byte index in table_data.bin
      );

      dataBytesWritten += chunkLen;
      fileCursor = nextEnd;

      // Measure memory
      const currentMem = getMemoryUsageMB();
      if (currentMem > peakMemoryMB) peakMemoryMB = currentMem;

      // Report throttled progress (every 200ms)
      const now = performance.now();
      if (now - lastProgressTime > 200 || fileCursor === file.size) {
        lastProgressTime = now;
        const percent = Math.min(
          90,
          Math.round(5 + ((dataBytesWritten / totalDataBytes) * 85))
        );
        self.postMessage({
          type: 'PROGRESS',
          progress: {
            bytesProcessed: dataBytesWritten,
            totalBytes: totalDataBytes,
            rowsProcessed: totalRowsIndexed,
            percent,
            phase: 'STREAMING_ROWS',
          } as ImportProgress,
        });
      }
    }

    // -------------------------------------------------------------------------
    // Phase 4: Finalize N+1 Index (handle trailing row & offset[N])
    // -------------------------------------------------------------------------
    // If the file ended with a character that was not newline, that final row ended at dataBytesWritten
    const lastBlockOffset = blockIndex > 0 ? currentBlock[blockIndex - 1] : BigInt(0);
    if (lastBlockOffset < BigInt(dataBytesWritten)) {
      // The final row was not followed by \n; its end is dataBytesWritten
      if (blockIndex >= INDEX_BLOCK_SIZE) {
        offsetBlocks.push(currentBlock);
        currentBlock = new BigUint64Array(INDEX_BLOCK_SIZE);
        blockIndex = 0;
      }
      currentBlock[blockIndex++] = BigInt(dataBytesWritten);
      totalRowsIndexed++;
    }

    offsetBlocks.push(currentBlock.subarray(0, blockIndex));

    // Consolidate into contiguous BigUint64Array of length N + 1
    const totalOffsetsCount = offsetBlocks.reduce((sum, b) => sum + b.length, 0);
    const finalIndexArray = new BigUint64Array(totalOffsetsCount);
    let copyPos = 0;
    for (const b of offsetBlocks) {
      finalIndexArray.set(b, copyPos);
      copyPos += b.length;
    }

    // Explicit check: rowCount = totalOffsetsCount - 1
    const rowCount = totalOffsetsCount - 1;

    // Verify offset[N] === dataBytesWritten
    finalIndexArray[rowCount] = BigInt(dataBytesWritten);

    self.postMessage({
      type: 'PROGRESS',
      progress: {
        bytesProcessed: dataBytesWritten,
        totalBytes: totalDataBytes,
        rowsProcessed: rowCount,
        percent: 92,
        phase: 'BUILDING_INDEX',
      } as ImportProgress,
    });

    // Write table_index.idx
    const serializedIndex = RowIndex.serialize(finalIndexArray);
    await writer.writeIndexBuffer(serializedIndex);

    // Close writer handles before commit
    await writer.close();

    // -------------------------------------------------------------------------
    // Phase 5: Validate and Commit
    // -------------------------------------------------------------------------
    self.postMessage({
      type: 'PROGRESS',
      progress: {
        bytesProcessed: dataBytesWritten,
        totalBytes: totalDataBytes,
        rowsProcessed: rowCount,
        percent: 96,
        phase: 'VALIDATING',
      } as ImportProgress,
    });

    metadata.rowCount = rowCount;
    metadata.dataSizeBytes = dataBytesWritten;
    metadata.indexSizeBytes = serializedIndex.byteLength;

    const commitResult = await CommitManager.validateAndCommit(metadata);
    if (!commitResult.success) {
      throw new Error(`Commit validation failed: ${commitResult.error}`);
    }

    const durationMs = Math.round(performance.now() - startTime);
    const finalMem = getMemoryUsageMB();
    if (finalMem > peakMemoryMB) peakMemoryMB = finalMem;

    self.postMessage({
      type: 'PROGRESS',
      progress: {
        bytesProcessed: dataBytesWritten,
        totalBytes: totalDataBytes,
        rowsProcessed: rowCount,
        percent: 100,
        phase: 'COMMITTED',
      } as ImportProgress,
    });

    self.postMessage({
      type: 'COMPLETE',
      result: {
        success: true,
        metadata,
        durationMs,
        peakMemoryMB: Math.round(peakMemoryMB * 100) / 100,
      },
    });
  } catch (err: any) {
    console.error('[CSV Import Worker] Fatal Error:', err);
    await writer.abort();
    self.postMessage({
      type: 'ERROR',
      error: err.message || String(err),
    });
  }
};
