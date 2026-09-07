import { NextRequest, NextResponse } from 'next/server';
import { prisma } from '@/lib/db/prisma';
import { normalizeWorkbookUuid } from '@/lib/uuid/normalizeUuid';

export async function GET(
  request: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const { id: rawWorkbookId } = await params;
    const workbookId = normalizeWorkbookUuid(rawWorkbookId);
    const searchParams = request.nextUrl.searchParams;
    const since = searchParams.get('since');

    const sinceRev = since ? BigInt(since) : BigInt(0);

    const results = await prisma.calculationResult.findMany({
      where: {
        workbookId,
        calculatedRev: {
          gt: sinceRev,
        },
      },
      orderBy: {
        calculatedRev: 'asc',
      },
    });

    const allPatches: any[] = [];
    for (const res of results) {
      if (Array.isArray(res.cellPatches)) {
        allPatches.push(...res.cellPatches);
      }
    }

    return NextResponse.json({
      workbookId,
      patches: allPatches,
      latestRev: results.length > 0 ? Number(results[results.length - 1].calculatedRev) : Number(sinceRev),
    });
  } catch (err: any) {
    console.error('[PATCHES_API] Error fetching patches:', err);
    return NextResponse.json({ error: err.message }, { status: 500 });
  }
}
