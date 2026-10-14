import { NextRequest, NextResponse } from 'next/server';
import prisma from '@/lib/db/prisma';
import { authenticateRequest, unauthorizedResponse, forbiddenResponse } from '@/lib/auth/rbacGuard';
import { QueryIndexManager } from '@/lib/largeCsv/query/queryIndexManager';
import { DatasetRowResolver } from '@/lib/largeCsv/query/datasetRowResolver';
import { MAX_RANGE_ROW_LIMIT } from '@/lib/largeCsv/reader/datasetRangeReader';
import { DatasetQueryError } from '@/lib/largeCsv/query/types';

export async function GET(
  req: NextRequest,
  { params }: { params: Promise<{ id: string; queryId: string }> }
) {
  const t0 = process.hrtime.bigint();
  try {
    const { id, queryId } = await params;

    const dataset = await prisma.dataset.findUnique({
      where: { id },
      select: { id: true, ownerId: true, status: true, headerRow: true, columnCount: true },
    });

    if (!dataset) {
      return NextResponse.json(
        { error: `Dataset ${id} not found`, code: 'DATASET_NOT_FOUND' },
        { status: 404 }
      );
    }

    if (dataset.ownerId) {
      const user = await authenticateRequest(req);
      if (!user) {
        return unauthorizedResponse('Authentication required.');
      }
      const isOwner = user.id === dataset.ownerId;
      const isRoot = user.isSystemRoot || user.roles.includes('ROOT_ADMIN');
      if (!isOwner && !isRoot) {
        return forbiddenResponse('Permission denied.');
      }
    }

    if (dataset.status !== 'READY') {
      return NextResponse.json(
        { error: `Dataset ${id} is not in READY status`, code: 'DATASET_NOT_READY' },
        { status: 409 }
      );
    }

    const { searchParams } = req.nextUrl;
    const rawStartRow = searchParams.get('startRow');
    const rawRowCount = searchParams.get('rowCount');

    if (rawStartRow === null || rawStartRow === '') {
      return NextResponse.json(
        { error: 'Query parameter "startRow" is required', code: 'MISSING_START_ROW' },
        { status: 400 }
      );
    }

    const startRow = parseInt(rawStartRow, 10);
    if (isNaN(startRow) || startRow < 0) {
      return NextResponse.json(
        { error: '"startRow" must be a non-negative integer', code: 'INVALID_START_ROW' },
        { status: 400 }
      );
    }

    const rowCount = rawRowCount !== null && rawRowCount !== '' ? parseInt(rawRowCount, 10) : 100;
    if (isNaN(rowCount) || rowCount <= 0) {
      return NextResponse.json(
        { error: '"rowCount" must be a positive integer', code: 'INVALID_ROW_COUNT' },
        { status: 400 }
      );
    }

    if (rowCount > MAX_RANGE_ROW_LIMIT) {
      return NextResponse.json(
        {
          error: `"rowCount" exceeds maximum allowed limit of ${MAX_RANGE_ROW_LIMIT}`,
          code: 'ROW_COUNT_EXCEEDS_LIMIT',
        },
        { status: 400 }
      );
    }

    // 1. Read bounded stableRowIds from query index
    const tIndex0 = process.hrtime.bigint();
    const { stableRowIds, totalMatchedRows } = await QueryIndexManager.readStableRowIds(
      id,
      queryId,
      startRow,
      rowCount
    );
    const tIndex1 = process.hrtime.bigint();

    // 2. Resolve row contents
    const tResolve0 = process.hrtime.bigint();
    const rows = await DatasetRowResolver.resolveRows(id, stableRowIds);
    const tResolve1 = process.hrtime.bigint();

    const tEnd = process.hrtime.bigint();

    const columns = (dataset.headerRow as string[]) || [];
    const nextRow = startRow + rows.length < totalMatchedRows ? startRow + rows.length : null;

    return NextResponse.json(
      {
        datasetId: id,
        queryId,
        startRow,
        rowCount: rows.length,
        totalRows: totalMatchedRows,
        columns,
        columnCount: dataset.columnCount || columns.length,
        rows,
        nextRow,
        metrics: {
          indexLookupMs: Number(tIndex1 - tIndex0) / 1_000_000,
          rowResolveMs: Number(tResolve1 - tResolve0) / 1_000_000,
          serverTotalMs: Number(tEnd - t0) / 1_000_000,
        },
      },
      { status: 200 }
    );
  } catch (err: any) {
    if (err instanceof DatasetQueryError) {
      return NextResponse.json(
        { error: err.message, code: err.code },
        { status: err.statusCode }
      );
    }

    return NextResponse.json(
      { error: err.message || 'Internal server error', code: 'INTERNAL_SERVER_ERROR' },
      { status: 500 }
    );
  }
}
