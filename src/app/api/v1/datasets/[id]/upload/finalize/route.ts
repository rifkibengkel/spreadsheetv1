import { NextRequest, NextResponse } from 'next/server';
import { DatasetUploadService } from '@/lib/largeCsv/datasetUploadService';

export async function POST(
  req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const { id: datasetId } = await params;

    let expectedChecksum: string | undefined = undefined;
    try {
      const body = await req.json();
      expectedChecksum = body.expectedChecksum;
    } catch {
      // Body may be empty
    }

    const result = await DatasetUploadService.finalizeUpload({
      datasetId,
      expectedChecksum,
    });

    return NextResponse.json(result, { status: 200 });
  } catch (err: any) {
    const status = err.message?.includes('not found')
      ? 404
      : err.message?.includes('mismatch') || err.message?.includes('Gap')
      ? 422
      : 400;

    return NextResponse.json(
      { error: err.message || 'Finalization failed' },
      { status }
    );
  }
}
