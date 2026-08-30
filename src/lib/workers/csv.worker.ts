/**
 * Dedicated Low-Memory Streaming CSV Web Worker
 * Processes worksheet rows in bounded incremental chunks without holding full JS strings in memory.
 */

const textEncoder = new TextEncoder();

function formatCsvRow(row: string[]): string {
  return row
    .map((cell) => {
      if (cell === '' || cell === undefined || cell === null) return '';
      const str = String(cell);
      if (
        str.includes('"') ||
        str.includes(',') ||
        str.includes('\n') ||
        str.includes('\r')
      ) {
        return `"${str.replace(/"/g, '""')}"`;
      }
      return str;
    })
    .join(',');
}

let outputChunks: Uint8Array[] = [];
let totalRowsExpected = 0;
let processedRowsCount = 0;
let sheetName = '';

self.onmessage = async (e: MessageEvent) => {
  try {
    const data = e.data;
    if (!data) return;

    const { type, chunkIndex, totalChunks, chunkRows, totalRows } = data;

    if (type === 'INIT_CSV') {
      outputChunks = [];
      totalRowsExpected = totalRows || 0;
      processedRowsCount = 0;
      sheetName = data.sheetName || 'Sheet';

      self.postMessage({ type: 'ACK' });
      return;
    }

    if (type === 'APPEND_CSV_CHUNK') {
      if (!chunkRows) {
        self.postMessage({ type: 'CHUNK_ACK', chunkIndex });
        return;
      }

      const chunkRowKeys = Object.keys(chunkRows).map(Number).sort((a, b) => a - b);
      const csvLines: string[] = [];

      for (let i = 0; i < chunkRowKeys.length; i++) {
        const rIndex = chunkRowKeys[i];
        const cols = chunkRows[rIndex];
        if (!cols) continue;

        const colIndices = Object.keys(cols).map(Number).sort((a, b) => a - b);
        if (colIndices.length === 0) continue;
        const maxCol = colIndices[colIndices.length - 1];

        const rowArr: string[] = new Array(maxCol + 1).fill('');
        for (let j = 0; j < colIndices.length; j++) {
          const cIndex = colIndices[j];
          const cell = cols[cIndex];
          if (cell) {
            let strVal = '';
            if (cell.v !== undefined && cell.v !== null) {
              strVal = String(cell.v);
            } else if (cell.p && cell.p.body && typeof cell.p.body.dataStream === 'string') {
              strVal = cell.p.body.dataStream.replace(/[\r\n]+$/, '');
            }
            rowArr[cIndex] = strVal;
          }
        }

        csvLines.push(formatCsvRow(rowArr));
      }

      if (csvLines.length > 0) {
        const chunkCsvString = csvLines.join('\r\n') + '\r\n';
        const chunkUint8 = textEncoder.encode(chunkCsvString);
        outputChunks.push(chunkUint8);
      }

      processedRowsCount += chunkRowKeys.length;

      // Send ACK back to release backpressure on main thread
      self.postMessage({ type: 'CHUNK_ACK', chunkIndex });
      return;
    }

    if (type === 'FINALIZE_CSV') {
      self.postMessage({ type: 'PROGRESS', percent: 95, message: 'Mengemas file CSV...' });

      const totalLength = outputChunks.reduce((sum, c) => sum + c.length, 0);
      const result = new Uint8Array(totalLength);
      let offset = 0;
      for (const chunk of outputChunks) {
        result.set(chunk, offset);
        offset += chunk.length;
      }

      outputChunks = [];

      self.postMessage({ type: 'PROGRESS', percent: 100, message: 'File CSV selesai dibuat!' });
      (self as any).postMessage(
        { type: 'COMPLETE', buffer: result.buffer },
        [result.buffer]
      );
      return;
    }
  } catch (err: any) {
    console.error('[CSV WORKER] Error:', err);
    outputChunks = [];
    self.postMessage({
      type: 'ERROR',
      error: err?.message || 'Gagal mengekspor CSV di Background Worker',
    });
  }
};
