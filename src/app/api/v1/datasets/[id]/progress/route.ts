import { NextRequest, NextResponse } from 'next/server';
import prisma from '@/lib/db/prisma';

export async function GET(
  req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const { id } = await params;
    const dataset = await prisma.dataset.findUnique({
      where: { id },
      select: {
        id: true,
        status: true,
        progress: true,
        currentStage: true,
        rowCount: true,
        columnCount: true,
        errorMessage: true,
        createdAt: true,
        updatedAt: true,
        completedAt: true,
        _count: {
          select: {
            partitions: true,
            uploadChunks: true,
          },
        },
      },
    });

    if (!dataset) {
      return NextResponse.json({ error: 'Dataset not found' }, { status: 404 });
    }

    return NextResponse.json({
      datasetId: dataset.id,
      status: dataset.status,
      progress: dataset.progress,
      currentStage: dataset.currentStage,
      rowCount: dataset.rowCount.toString(),
      columnCount: dataset.columnCount,
      partitionsCount: dataset._count.partitions,
      chunksUploaded: dataset._count.uploadChunks,
      errorMessage: dataset.errorMessage,
      updatedAt: dataset.updatedAt.toISOString(),
      completedAt: dataset.completedAt?.toISOString() || null,
    });
  } catch (err: any) {
    return NextResponse.json(
      { error: err.message || 'Failed to fetch progress' },
      { status: 500 }
    );
  }
}
