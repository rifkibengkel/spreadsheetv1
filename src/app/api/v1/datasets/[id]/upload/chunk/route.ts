import { NextRequest, NextResponse } from "next/server";
import { DatasetUploadService } from "@/lib/largeCsv/datasetUploadService";

export async function POST(
  req: NextRequest,
  { params }: { params: Promise<{ id: string }> },
) {
  try {
    const { id: datasetId } = await params;

    const chunkIndexHeader = req.headers.get("x-chunk-index");
    const byteOffsetHeader = req.headers.get("x-byte-offset");
    const chunkChecksum = req.headers.get("x-chunk-checksum") || undefined;

    if (!chunkIndexHeader || !byteOffsetHeader) {
      return NextResponse.json(
        { error: "x-chunk-index and x-byte-offset headers are required" },
        { status: 400 },
      );
    }

    const chunkIndex = parseInt(chunkIndexHeader, 10);
    const byteOffset = BigInt(byteOffsetHeader);

    if (isNaN(chunkIndex) || byteOffset < 0) {
      return NextResponse.json(
        { error: "Invalid chunkIndex or byteOffset" },
        { status: 400 },
      );
    }

    const arrayBuffer = await req.arrayBuffer();
    const chunkBuffer = Buffer.from(arrayBuffer);

    if (chunkBuffer.length === 0) {
      return NextResponse.json(
        { error: "Empty chunk body received" },
        { status: 400 },
      );
    }

    const result = await DatasetUploadService.uploadChunk({
      datasetId,
      chunkIndex,
      byteOffset,
      chunkBuffer,
      chunkChecksum,
    });

    return NextResponse.json(result, { status: 200 });
  } catch (err: any) {
    const status = err.message?.includes("not found")
      ? 404
      : err.message?.includes("Conflicting")
        ? 409
        : 400;

    return NextResponse.json(
      { error: err.message || "Chunk upload failed" },
      { status },
    );
  }
}
