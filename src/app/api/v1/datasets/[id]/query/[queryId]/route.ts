import { NextRequest, NextResponse } from 'next/server';
import prisma from '@/lib/db/prisma';
import { authenticateRequest, unauthorizedResponse, forbiddenResponse } from '@/lib/auth/rbacGuard';
import { QueryIndexManager } from '@/lib/largeCsv/query/queryIndexManager';

export async function GET(
  req: NextRequest,
  { params }: { params: Promise<{ id: string; queryId: string }> }
) {
  try {
    const { id, queryId } = await params;

    const dataset = await prisma.dataset.findUnique({
      where: { id },
      select: { id: true, ownerId: true },
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

    const meta = await QueryIndexManager.getQueryMetadata(id, queryId);
    if (!meta) {
      return NextResponse.json(
        { error: `Query ${queryId} not found or not in READY status`, code: 'QUERY_NOT_FOUND' },
        { status: 404 }
      );
    }

    return NextResponse.json(meta, { status: 200 });
  } catch (err: any) {
    return NextResponse.json(
      { error: err.message || 'Internal server error', code: 'INTERNAL_SERVER_ERROR' },
      { status: 500 }
    );
  }
}

export async function DELETE(
  req: NextRequest,
  { params }: { params: Promise<{ id: string; queryId: string }> }
) {
  try {
    const { id, queryId } = await params;

    const dataset = await prisma.dataset.findUnique({
      where: { id },
      select: { id: true, ownerId: true },
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

    await QueryIndexManager.deleteQuery(id, queryId);

    return NextResponse.json({ success: true, queryId }, { status: 200 });
  } catch (err: any) {
    return NextResponse.json(
      { error: err.message || 'Internal server error', code: 'INTERNAL_SERVER_ERROR' },
      { status: 500 }
    );
  }
}
