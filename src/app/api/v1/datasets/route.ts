import { NextRequest, NextResponse } from 'next/server';
import prisma from '@/lib/db/prisma';
import { DatasetUploadService } from '@/lib/largeCsv/datasetUploadService';

export async function POST(req: NextRequest) {
  try {
    const body = await req.json();
    const { name, originalFilename, byteSize, expectedChecksum, delimiter } = body;

    if (!name || !originalFilename || !byteSize) {
      return NextResponse.json(
        { error: 'name, originalFilename, and byteSize are required' },
        { status: 400 }
      );
    }

    const session = await DatasetUploadService.createUploadSession({
      name,
      originalFilename,
      byteSize,
      expectedChecksum,
      delimiter,
    });

    return NextResponse.json(session, { status: 201 });
  } catch (err: any) {
    return NextResponse.json(
      { error: err.message || 'Failed to create upload session' },
      { status: 500 }
    );
  }
}

export async function GET() {
  try {
    const datasets = await prisma.dataset.findMany({
      orderBy: { createdAt: 'desc' },
      take: 50,
      include: {
        _count: {
          select: { partitions: true, uploadChunks: true },
        },
      },
    });

    const serialized = datasets.map((d) => ({
      id: d.id,
      name: d.name,
      originalFilename: d.originalFilename,
      byteSize: d.byteSize.toString(),
      rowCount: d.rowCount.toString(),
      columnCount: d.columnCount,
      status: d.status,
      progress: d.progress,
      currentStage: d.currentStage,
      partitionsCount: d._count.partitions,
      createdAt: d.createdAt.toISOString(),
      completedAt: d.completedAt?.toISOString() || null,
    }));

    return NextResponse.json({ datasets: serialized });
  } catch (err: any) {
    return NextResponse.json(
      { error: err.message || 'Failed to list datasets' },
      { status: 500 }
    );
  }
}
