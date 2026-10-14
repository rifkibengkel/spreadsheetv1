import { NextRequest, NextResponse } from 'next/server';
import path from 'path';
import os from 'os';
import fs from 'fs';
import { Readable } from 'stream';
import prisma from '@/lib/db/prisma';
import { authenticateRequest, unauthorizedResponse, forbiddenResponse } from '@/lib/auth/rbacGuard';
import { QueryIndexManager } from '@/lib/largeCsv/query/queryIndexManager';
import { DatasetRowResolver } from '@/lib/largeCsv/query/datasetRowResolver';

export const dynamic = 'force-dynamic';

const ROWS_PER_LOGICAL_SHEET = 4_000_000;
const WINDOWS_RESERVED_NAMES = /^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])$/i;

function formatCsvRow(cells: (string | number | null | undefined)[]): string {
  return cells
    .map((cell) => {
      if (cell === '' || cell === undefined || cell === null) return '';
      const str = String(cell);
      if (str.includes('"') || str.includes(',') || str.includes('\n') || str.includes('\r')) {
        return `"${str.replace(/"/g, '""')}"`;
      }
      return str;
    })
    .join(',') + '\n';
}

function sanitizeExportFilename(rawInput: string, defaultName: string): string {
  let clean = (rawInput || defaultName).trim();
  clean = clean.replace(/[<>:"/\\|?*\x00-\x1F]/g, '_').trim();
  clean = clean.replace(/[. ]+$/, '');
  clean = clean.replace(/\.csv$/i, '');
  if (WINDOWS_RESERVED_NAMES.test(clean) || !clean) {
    clean = defaultName.replace(/\.csv$/i, '');
  }
  return `${clean}.csv`;
}

interface ExportPlan {
  dataset: any;
  startRow: number;
  endRow: number;
  totalRowsToExport: number;
  safeFilename: string;
  headerLine: string;
  partitionSlices: Array<{
    partNumber: number;
    filePath: string;
    startByte: number;
    endByte: number; // inclusive for fs.createReadStream
    rowCount: number;
  }>;
  queryId?: string | null;
  totalMatchedRows?: number;
}

async function prepareExportPlan(
  datasetId: string,
  scope: string,
  sheetIndex: number,
  queryId?: string | null,
  customFilename?: string
): Promise<ExportPlan> {
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
    throw new Error(`Dataset ${datasetId} not found`);
  }
  if (dataset.status !== 'READY') {
    throw new Error(`Dataset ${datasetId} is not in READY status (current: ${dataset.status})`);
  }

  const totalDatasetRows = Number(dataset.rowCount);
  const rawDatasetName = dataset.originalFilename || dataset.name || `dataset-${datasetId}`;
  const baseName = path.basename(rawDatasetName, path.extname(rawDatasetName));

  // Handle Query filter/sort mode
  if (queryId) {
    const qMeta = await QueryIndexManager.getQueryMetadata(datasetId, queryId);
    if (!qMeta) {
      throw new Error(`Query ${queryId} not found or not in READY status`);
    }
    const matchedRows = qMeta.matchedRows;
    let startRow = 0;
    let endRow = Math.max(0, matchedRows - 1);

    if (scope === 'sheet') {
      startRow = sheetIndex * ROWS_PER_LOGICAL_SHEET;
      endRow = Math.min(matchedRows - 1, (sheetIndex + 1) * ROWS_PER_LOGICAL_SHEET - 1);
      if (matchedRows > 0 && startRow >= matchedRows) {
        throw new Error(`sheetIndex ${sheetIndex} is out of bounds for query result (${matchedRows} matched rows)`);
      }
    }

    const rowCount = matchedRows === 0 ? 0 : Math.max(0, endRow - startRow + 1);
    const defaultName = scope === 'sheet'
      ? `${baseName}-query-sheet-${sheetIndex + 1}.csv`
      : `${baseName}-query-all.csv`;

    const safeFilename = sanitizeExportFilename(customFilename || '', defaultName);
    const headerCells = (dataset.headerRow as string[]) || [];
    const headerLine = headerCells.length > 0 ? formatCsvRow(headerCells) : '';

    return {
      dataset,
      startRow,
      endRow,
      totalRowsToExport: rowCount,
      safeFilename,
      headerLine,
      partitionSlices: [],
      queryId,
      totalMatchedRows: matchedRows,
    };
  }

  // Direct physical partition streaming mode
  let startRow = 0;
  let endRow = Math.max(0, totalDatasetRows - 1);

  if (scope === 'sheet') {
    startRow = sheetIndex * ROWS_PER_LOGICAL_SHEET;
    endRow = Math.min(totalDatasetRows - 1, (sheetIndex + 1) * ROWS_PER_LOGICAL_SHEET - 1);
    if (totalDatasetRows > 0 && startRow >= totalDatasetRows) {
      throw new Error(`sheetIndex ${sheetIndex} is out of bounds for dataset (${totalDatasetRows} rows)`);
    }
  }

  const totalRowsToExport = totalDatasetRows === 0 ? 0 : Math.max(0, endRow - startRow + 1);
  const defaultName = scope === 'sheet'
    ? `${baseName}-sheet-${sheetIndex + 1}.csv`
    : `${baseName}-all.csv`;

  const safeFilename = sanitizeExportFilename(customFilename || '', defaultName);
  const headerCells = (dataset.headerRow as string[]) || [];
  const headerLine = headerCells.length > 0 ? formatCsvRow(headerCells) : '';

  const partitions = dataset.partitions;
  if (!partitions || partitions.length === 0) {
    throw new Error(`Dataset ${datasetId} has no committed READY partitions`);
  }

  const overlapping = partitions.filter((p) => {
    const pStart = Number(p.stableRowStart);
    const pEnd = Number(p.stableRowEnd);
    return pEnd >= startRow && pStart <= endRow;
  });

  const partitionSlices: ExportPlan['partitionSlices'] = [];

  for (const part of overlapping) {
    const pStart = Number(part.stableRowStart);
    const pEnd = Number(part.stableRowEnd);
    const localStart = Math.max(0, startRow - pStart);
    const localEnd = Math.min(part.rowCount - 1, endRow - pStart);
    const subCount = localEnd - localStart + 1;
    if (subCount <= 0) continue;

    let startByte = 0;
    let endByte = Number(part.byteSize) - 1;

    // If slice does not span the full partition, read exact byte offsets from .idx
    if (localStart > 0 || localEnd < part.rowCount - 1) {
      const idxFd = await fs.promises.open(part.indexFilePath, 'r');
      try {
        const offsetBuf = Buffer.alloc(16);
        await idxFd.read(offsetBuf, 0, 8, localStart * 8);
        await idxFd.read(offsetBuf, 8, 8, (localEnd + 1) * 8);
        startByte = Number(offsetBuf.readBigUInt64LE(0));
        const endByteExclusive = Number(offsetBuf.readBigUInt64LE(8));
        endByte = endByteExclusive - 1; // inclusive for fs.createReadStream
      } finally {
        await idxFd.close();
      }
    }

    partitionSlices.push({
      partNumber: part.partitionNumber,
      filePath: part.dataFilePath,
      startByte,
      endByte,
      rowCount: subCount,
    });
  }

  return {
    dataset,
    startRow,
    endRow,
    totalRowsToExport,
    safeFilename,
    headerLine,
    partitionSlices,
    queryId: null,
  };
}

export async function GET(
  req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const { id } = await params;
    if (!id) {
      return NextResponse.json({ error: 'Dataset ID is required' }, { status: 400 });
    }

    const { searchParams } = req.nextUrl;
    const scope = (searchParams.get('scope') || 'sheet').toLowerCase();
    const sheetIndex = parseInt(searchParams.get('sheetIndex') || '0', 10);
    const queryId = searchParams.get('queryId') || null;
    const customFilename = searchParams.get('filename') || undefined;
    const saveToDesktop = searchParams.get('saveToDesktop') === 'true';

    // Authorization check
    const dsCheck = await prisma.dataset.findUnique({
      where: { id },
      select: { ownerId: true },
    });
    if (!dsCheck) {
      return NextResponse.json({ error: `Dataset ${id} not found` }, { status: 404 });
    }
    if (dsCheck.ownerId) {
      const user = await authenticateRequest(req);
      if (!user) {
        return unauthorizedResponse('Authentication required to access this dataset.');
      }
      const isOwner = user.id === dsCheck.ownerId;
      const isRoot = user.isSystemRoot || user.roles.includes('ROOT_ADMIN');
      if (!isOwner && !isRoot) {
        return forbiddenResponse('You do not have permission to access this dataset.');
      }
    }

    const plan = await prepareExportPlan(id, scope, sheetIndex, queryId, customFilename);

    // Save to Desktop direct execution
    if (saveToDesktop) {
      return await executeSaveToDesktop(plan);
    }

    // Direct HTTP chunked streaming download to browser
    const stream = createCsvExportStream(plan);

    return new Response(stream, {
      status: 200,
      headers: {
        'Content-Type': 'text/csv; charset=utf-8',
        'Content-Disposition': `attachment; filename="${encodeURIComponent(plan.safeFilename)}"`,
        'Cache-Control': 'no-cache, no-store, must-revalidate',
        'Transfer-Encoding': 'chunked',
        'X-Total-Rows': String(plan.totalRowsToExport),
      },
    });
  } catch (err: any) {
    console.error('[CSV EXPORT API] Error:', err);
    return NextResponse.json(
      { error: err.message || 'CSV Export failed' },
      { status: 500 }
    );
  }
}

export async function POST(
  req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const { id } = await params;
    if (!id) {
      return NextResponse.json({ error: 'Dataset ID is required' }, { status: 400 });
    }

    let body: any = {};
    try {
      body = await req.json();
    } catch {}

    const scope = (body.scope || 'sheet').toLowerCase();
    const sheetIndex = typeof body.sheetIndex === 'number' ? body.sheetIndex : 0;
    const queryId = body.queryId || null;
    const customFilename = body.filename || undefined;
    const saveToDesktop = body.saveToDesktop === true;

    // Authorization check
    const dsCheck = await prisma.dataset.findUnique({
      where: { id },
      select: { ownerId: true },
    });
    if (!dsCheck) {
      return NextResponse.json({ error: `Dataset ${id} not found` }, { status: 404 });
    }
    if (dsCheck.ownerId) {
      const user = await authenticateRequest(req);
      if (!user) {
        return unauthorizedResponse('Authentication required to access this dataset.');
      }
      const isOwner = user.id === dsCheck.ownerId;
      const isRoot = user.isSystemRoot || user.roles.includes('ROOT_ADMIN');
      if (!isOwner && !isRoot) {
        return forbiddenResponse('You do not have permission to access this dataset.');
      }
    }

    const plan = await prepareExportPlan(id, scope, sheetIndex, queryId, customFilename);

    if (saveToDesktop) {
      return await executeSaveToDesktop(plan);
    }

    const stream = createCsvExportStream(plan);

    return new Response(stream, {
      status: 200,
      headers: {
        'Content-Type': 'text/csv; charset=utf-8',
        'Content-Disposition': `attachment; filename="${encodeURIComponent(plan.safeFilename)}"`,
        'Cache-Control': 'no-cache, no-store, must-revalidate',
        'Transfer-Encoding': 'chunked',
        'X-Total-Rows': String(plan.totalRowsToExport),
      },
    });
  } catch (err: any) {
    console.error('[CSV EXPORT API POST] Error:', err);
    return NextResponse.json(
      { error: err.message || 'CSV Export failed' },
      { status: 500 }
    );
  }
}

/**
 * Creates a Web ReadableStream that streams CSV data without intermediate full-file buffering.
 */
function createCsvExportStream(plan: ExportPlan): ReadableStream<Uint8Array> {
  let activeFileStream: fs.ReadStream | null = null;
  let isCancelled = false;

  return new ReadableStream<Uint8Array>({
    async start(controller) {
      try {
        // 1. Output CSV Header Row
        if (plan.headerLine) {
          controller.enqueue(Buffer.from(plan.headerLine, 'utf8'));
        }

        // 2. Stream Rows: Query mode or Direct Partition mode
        if (plan.queryId) {
          const CHUNK_SIZE = 5000;
          let currentOffset = plan.startRow;
          const endOffset = plan.endRow;

          while (currentOffset <= endOffset && !isCancelled) {
            const batchCount = Math.min(CHUNK_SIZE, endOffset - currentOffset + 1);
            const { stableRowIds } = await QueryIndexManager.readStableRowIds(
              plan.dataset.id,
              plan.queryId,
              currentOffset,
              batchCount
            );

            if (stableRowIds.length === 0) break;

            const resolvedRows = await DatasetRowResolver.resolveRows(plan.dataset.id, stableRowIds);
            let chunkStr = '';
            for (let i = 0; i < resolvedRows.length; i++) {
              chunkStr += formatCsvRow(resolvedRows[i].cells);
            }

            if (chunkStr.length > 0) {
              controller.enqueue(Buffer.from(chunkStr, 'utf8'));
            }

            currentOffset += stableRowIds.length;
          }
        } else {
          // Direct physical partition stream
          for (const slice of plan.partitionSlices) {
            if (isCancelled) break;
            if (slice.endByte < slice.startByte) continue;

            await new Promise<void>((resolve, reject) => {
              const fileStream = fs.createReadStream(slice.filePath, {
                start: slice.startByte,
                end: slice.endByte,
                highWaterMark: 256 * 1024,
              });

              activeFileStream = fileStream;

              fileStream.on('data', (chunk: Buffer | string) => {
                if (isCancelled) {
                  fileStream.destroy();
                  return;
                }
                controller.enqueue(typeof chunk === 'string' ? Buffer.from(chunk, 'utf8') : chunk);
              });

              fileStream.on('end', () => {
                activeFileStream = null;
                resolve();
              });

              fileStream.on('error', (err) => {
                activeFileStream = null;
                reject(err);
              });
            });
          }
        }

        controller.close();
      } catch (streamErr) {
        console.error('[CSV STREAM] Pipeline error:', streamErr);
        controller.error(streamErr);
      }
    },
    cancel() {
      isCancelled = true;
      if (activeFileStream) {
        try {
          activeFileStream.destroy();
        } catch {}
      }
    },
  });
}

/**
 * Saves exported CSV directly to user's Desktop on filesystem.
 * Bypasses HTTP payload buffering and 50MB request size limits.
 */
async function executeSaveToDesktop(plan: ExportPlan) {
  const desktopDir = path.resolve(os.homedir(), 'Desktop');
  await fs.promises.mkdir(desktopDir, { recursive: true });

  const targetFilePath = path.resolve(desktopDir, plan.safeFilename);

  // Security check: Must reside directly within Desktop
  const relative = path.relative(desktopDir, targetFilePath);
  if (relative.startsWith('..') || path.isAbsolute(relative) || relative.includes(path.sep)) {
    return NextResponse.json({ error: 'Invalid destination path' }, { status: 400 });
  }

  const writeStream = fs.createWriteStream(targetFilePath, {
    flags: 'w',
    highWaterMark: 512 * 1024,
  });

  const tStart = performance.now();

  try {
    // 1. Write Header
    if (plan.headerLine) {
      writeStream.write(Buffer.from(plan.headerLine, 'utf8'));
    }

    // 2. Stream Data
    if (plan.queryId) {
      const CHUNK_SIZE = 5000;
      let currentOffset = plan.startRow;
      const endOffset = plan.endRow;

      while (currentOffset <= endOffset) {
        const batchCount = Math.min(CHUNK_SIZE, endOffset - currentOffset + 1);
        const { stableRowIds } = await QueryIndexManager.readStableRowIds(
          plan.dataset.id,
          plan.queryId,
          currentOffset,
          batchCount
        );
        if (stableRowIds.length === 0) break;

        const resolvedRows = await DatasetRowResolver.resolveRows(plan.dataset.id, stableRowIds);
        let chunkStr = '';
        for (let i = 0; i < resolvedRows.length; i++) {
          chunkStr += formatCsvRow(resolvedRows[i].cells);
        }
        if (chunkStr.length > 0) {
          const canContinue = writeStream.write(Buffer.from(chunkStr, 'utf8'));
          if (!canContinue) {
            await new Promise<void>((res) => writeStream.once('drain', () => res()));
          }
        }
        currentOffset += stableRowIds.length;
      }
    } else {
      for (const slice of plan.partitionSlices) {
        if (slice.endByte < slice.startByte) continue;

        await new Promise<void>((resolve, reject) => {
          const fileStream = fs.createReadStream(slice.filePath, {
            start: slice.startByte,
            end: slice.endByte,
            highWaterMark: 256 * 1024,
          });

          fileStream.on('data', (chunk) => {
            const canContinue = writeStream.write(chunk);
            if (!canContinue) {
              fileStream.pause();
              writeStream.once('drain', () => fileStream.resume());
            }
          });

          fileStream.on('end', () => resolve());
          fileStream.on('error', (err) => reject(err));
        });
      }
    }

    await new Promise<void>((resolve, reject) => {
      writeStream.end(() => resolve());
      writeStream.on('error', reject);
    });

    const stat = await fs.promises.stat(targetFilePath);
    const durationMs = Math.round(performance.now() - tStart);

    return NextResponse.json({
      success: true,
      message: 'File berhasil diekspor langsung ke Desktop.',
      filename: plan.safeFilename,
      filePath: targetFilePath,
      sizeBytes: stat.size,
      rowCount: plan.totalRowsToExport,
      durationMs,
    });
  } catch (saveErr: any) {
    writeStream.destroy();
    try {
      await fs.promises.unlink(targetFilePath);
    } catch {}
    throw saveErr;
  }
}
