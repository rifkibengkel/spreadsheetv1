import { NextRequest, NextResponse } from 'next/server';
import fs from 'fs';
import path from 'path';
import prisma from '@/lib/db/prisma';
import { authenticateRequest, unauthorizedResponse, forbiddenResponse } from '@/lib/auth/rbacGuard';
import { QueryIndexManager } from '@/lib/largeCsv/query/queryIndexManager';
import { QueryDefinition, DatasetQueryError } from '@/lib/largeCsv/query/types';

export async function POST(
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
        return unauthorizedResponse('Authentication required to query this dataset.');
      }
      const isOwner = user.id === dataset.ownerId;
      const isRoot = user.isSystemRoot || user.roles.includes('ROOT_ADMIN');
      if (!isOwner && !isRoot) {
        return forbiddenResponse('You do not have permission to query this dataset.');
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

    // 4. Parse body
    let body: any;
    try {
      body = await req.json();
    } catch {
      return NextResponse.json(
        { error: 'Invalid JSON request body', code: 'INVALID_JSON' },
        { status: 400 }
      );
    }

    const queryDef: QueryDefinition = {
      filter: body.filter,
      sort: body.sort,
    };

    // 5. Execute or fetch query index
    const meta = await QueryIndexManager.executeQuery(id, queryDef);

    return NextResponse.json(meta, { status: 200 });
  } catch (err: any) {
    if (err instanceof DatasetQueryError) {
      return NextResponse.json(
        { error: err.message, code: err.code },
        { status: err.statusCode }
      );
    }

    return NextResponse.json(
      { error: err.message || 'Internal server error during query execution', code: 'INTERNAL_SERVER_ERROR' },
      { status: 500 }
    );
  }
}

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

    const dir = QueryIndexManager.getQueriesDir(id);
    if (!fs.existsSync(dir)) {
      return NextResponse.json({ queries: [] }, { status: 200 });
    }

    const files = await fs.promises.readdir(dir);
    const metaFiles = files.filter((f) => f.startsWith('query_') && f.endsWith('.meta.json'));

    const queries: any[] = [];
    for (const f of metaFiles) {
      try {
        const raw = await fs.promises.readFile(path.join(dir, f), 'utf8');
        queries.push(JSON.parse(raw));
      } catch {}
    }

    return NextResponse.json({ queries }, { status: 200 });
  } catch (err: any) {
    return NextResponse.json(
      { error: err.message || 'Internal server error', code: 'INTERNAL_SERVER_ERROR' },
      { status: 500 }
    );
  }
}
