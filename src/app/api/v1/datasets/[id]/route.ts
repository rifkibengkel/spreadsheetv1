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
      include: {
        partitions: {
          orderBy: { partitionNumber: 'asc' },
        },
      },
    });

    if (!dataset) {
      return NextResponse.json({ error: 'Dataset not found' }, { status: 404 });
    }

    return NextResponse.json({
      id: dataset.id,
      name: dataset.name,
      originalFilename: dataset.originalFilename,
      byteSize: dataset.byteSize.toString(),
      actualChecksum: dataset.actualChecksum,
      rowCount: dataset.rowCount.toString(),
      columnCount: dataset.columnCount,
      hasHeader: dataset.hasHeader,
      headerRow: dataset.headerRow,
      status: dataset.status,
      progress: dataset.progress,
      currentStage: dataset.currentStage,
      errorMessage: dataset.errorMessage,
      createdAt: dataset.createdAt.toISOString(),
      completedAt: dataset.completedAt?.toISOString() || null,
      partitions: dataset.partitions.map((p) => ({
        id: p.id,
        partitionNumber: p.partitionNumber,
        stableRowStart: p.stableRowStart.toString(),
        stableRowEnd: p.stableRowEnd.toString(),
        rowCount: p.rowCount,
        byteSize: p.byteSize.toString(),
        checksum: p.checksum,
        status: p.status,
        createdAt: p.createdAt.toISOString(),
        completedAt: p.completedAt?.toISOString() || null,
      })),
    });
  } catch (err: any) {
    return NextResponse.json(
      { error: err.message || 'Failed to fetch dataset' },
      { status: 500 }
    );
  }
}
