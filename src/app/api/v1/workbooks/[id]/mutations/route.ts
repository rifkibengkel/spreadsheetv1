import { NextRequest, NextResponse } from "next/server";
import { prisma } from "@/lib/db/prisma";
import { calculationWorker } from "@/worker/calculationWorker";
import { normalizeWorkbookUuid } from "@/lib/uuid/normalizeUuid";
import { enqueueCalculationJob } from "@/lib/queue/calculationQueue";

export async function POST(
  request: NextRequest,
  { params }: { params: Promise<{ id: string }> },
) {
  try {
    const { id: rawWorkbookId } = await params;
    const workbookId = normalizeWorkbookUuid(rawWorkbookId);
    let body: any;
    try {
      const rawText = await request.text();
      body = rawText ? JSON.parse(rawText) : {};
    } catch (parseErr: any) {
      return NextResponse.json(
        { error: "Invalid JSON body: " + parseErr.message },
        { status: 400 },
      );
    }
    const { deltas, idempotencyKey } = body;

    if (!Array.isArray(deltas) || deltas.length === 0) {
      return NextResponse.json(
        { error: "deltas must be a non-empty array" },
        { status: 400 },
      );
    }

    const key =
      idempotencyKey ||
      `idem_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;

    // 1. Transactional mutation persistence in PostgreSQL WAL
    const { newRev, batch, user } = await prisma.$transaction(
      async (tx: any) => {
        // Find system user or first user for foreign key
        let user = await tx.user.findFirst();
        if (!user) {
          user = await tx.user.create({
            data: {
              username: "system",
              email: "system@sheet.local",
              passwordHash: "dummy",
              fullName: "System Automation",
            },
          });
        }

        // Check existing workbook row
        let wb = await tx.workbook.findUnique({ where: { id: workbookId } });
        if (!wb) {
          wb = await tx.workbook.create({
            data: {
              id: workbookId,
              name: "Active Workbook",
              originalFilename: "active.xlsx",
              sizeBytes: BigInt(0),
              storageKey: "local",
              ownerId: user.id,
            },
          });
        }

        // Upsert workbook state and lock row
        let state = await tx.workbookState.findUnique({
          where: { workbookId },
        });

        if (!state) {
          state = await tx.workbookState.create({
            data: {
              workbookId,
              latestPersistedRev: BigInt(0),
              latestCalculatedRev: BigInt(0),
              fencingToken: BigInt(1),
            },
          });
        }

        // Atomically increment latestPersistedRev with PostgreSQL row lock
        const updatedState = await tx.workbookState.update({
          where: { workbookId },
          data: {
            latestPersistedRev: { increment: 1 },
            updatedAt: new Date(),
          },
        });
        const newRev = Number(updatedState.latestPersistedRev);

        // Append to Mutation WAL
        const batch = await tx.mutationBatch.create({
          data: {
            workbookId,
            revisionId: BigInt(newRev),
            userId: user.id,
            idempotencyKey: key,
            deltas,
          },
        });

        return { newRev, batch, user };
      },
    );

    // 2. Offload to BullMQ calculation queue backed by Redis
    try {
      await enqueueCalculationJob({
        workbookId,
        revisionId: newRev,
        enqueuedAt: Date.now(),
        actorUserId: user.id,
        deltaSummary: {
          sheetCount: 1,
          cellCount: deltas.length,
        },
      });
    } catch (queueErr: any) {
      console.warn("[MUTATION_API] BullMQ enqueue notice:", queueErr.message);
    }

    // 3. Process workbook revision in backend engine and publish to Redis
    const calcResult = await calculationWorker.processWorkbookRevision(
      workbookId,
      newRev,
    );

    return NextResponse.json({
      status: "COMMITTED",
      workbookId,
      revisionId: newRev,
      patches: calcResult?.patches || [],
      executionMs: calcResult?.executionMs || 0,
    });
  } catch (err: any) {
    console.error("[MUTATION_API] Error processing mutation:", err);
    return NextResponse.json({ error: err.message }, { status: 500 });
  }
}
