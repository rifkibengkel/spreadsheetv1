import { NextRequest, NextResponse } from 'next/server';
import prisma from '@/lib/db/prisma';
import { authenticateRequest, unauthorizedResponse, forbiddenResponse } from '@/lib/auth/rbacGuard';
import {
  DatasetRangeReader,
  DatasetRangeError,
  MAX_RANGE_ROW_LIMIT,
} from '@/lib/largeCsv/reader/datasetRangeReader';
import { QueryIndexManager } from '@/lib/largeCsv/query/queryIndexManager';
import { DatasetRowResolver } from '@/lib/largeCsv/query/datasetRowResolver';
import { DatasetQueryError } from '@/lib/largeCsv/query/types';

export async function GET(
  req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const { id } = await params;

    if (!id || typeof id !== 'string') {
      return NextResponse.json(
        { error: 'Dataset ID is required', code: 'INVALID_DATASET_ID' },
        { status: 400 }
      );
    }

    // 1. Fetch dataset for authorization & existence check
    const dataset = await prisma.dataset.findUnique({
      where: { id },
      select: { id: true, ownerId: true, status: true },
    });

    if (!dataset) {
      return NextResponse.json(
        { error: `Dataset ${id} not found`, code: 'DATASET_NOT_FOUND' },
        { status: 404 }
      );
    }

    // 2. Authorization check
    if (dataset.ownerId) {
      const user = await authenticateRequest(req);
      if (!user) {
        return unauthorizedResponse('Authentication required to access this dataset.');
      }
      const isOwner = user.id === dataset.ownerId;
      const isRoot = user.isSystemRoot || user.roles.includes('ROOT_ADMIN');
      if (!isOwner && !isRoot) {
        return forbiddenResponse('You do not have permission to access this dataset.');
      }
    }

    // 3. Status check
    if (dataset.status !== 'READY') {
      return NextResponse.json(
        {
          error: `Dataset ${id} is not in READY status (current status: ${dataset.status})`,
          code: 'DATASET_NOT_READY',
        },
        { status: 409 }
      );
    }

    // 4. Parse and validate query parameters
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

    const queryId = searchParams.get('queryId');

    // 5. Read bounded range: either via Query Index + Row Resolver, or default DatasetRangeReader
    if (queryId) {
      const { stableRowIds, totalMatchedRows } = await QueryIndexManager.readStableRowIds(
        id,
        queryId,
        startRow,
        rowCount
      );

      const rows = await DatasetRowResolver.resolveRows(id, stableRowIds);
      const datasetMeta = await prisma.dataset.findUnique({
        where: { id },
        select: { headerRow: true, columnCount: true },
      });

      const columns = (datasetMeta?.headerRow as string[]) || [];
      const nextRow = startRow + rows.length < totalMatchedRows ? startRow + rows.length : null;

      return NextResponse.json(
        {
          datasetId: id,
          queryId,
          startRow,
          rowCount: rows.length,
          totalRows: totalMatchedRows,
          columns,
          columnCount: datasetMeta?.columnCount || columns.length,
          rows,
          nextRow,
          metrics: {
            partitionResolutionMs: 0,
            indexLookupMs: 0,
            fileReadMs: 0,
            decodeMs: 0,
            serverTotalMs: 0,
          },
        },
        { status: 200 }
      );
    }

    const result = await DatasetRangeReader.readRange(id, startRow, rowCount);

    return NextResponse.json(result, { status: 200 });
  } catch (err: any) {
    if (err instanceof DatasetRangeError || err instanceof DatasetQueryError) {
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
