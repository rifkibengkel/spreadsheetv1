import fs from 'fs';
import path from 'path';
import crypto from 'crypto';
import prisma from '@/lib/db/prisma';
import {
  getDatasetSpoolDir,
  getDatasetSpoolFilePath,
} from './storage/storagePaths';
import { enqueueDatasetImportJob } from './queue/datasetImportQueue';

export interface CreateSessionParams {
  name?: string;
  originalFilename: string;
  byteSize: number | bigint;
  expectedChecksum?: string;
  ownerId?: string;
  delimiter?: string;
}

export interface UploadChunkParams {
  datasetId: string;
  chunkIndex: number;
  byteOffset: number | bigint;
  chunkBuffer: Buffer;
  chunkChecksum?: string;
}

export interface FinalizeUploadParams {
  datasetId: string;
  expectedChecksum?: string;
}

export class DatasetUploadService {
  /**
   * Create an upload session for a new large CSV dataset.
   */
  static async createUploadSession(params: CreateSessionParams) {
    const byteSizeBigInt = BigInt(params.byteSize);

    if (byteSizeBigInt <= BigInt(0)) {
      throw new Error('byteSize must be greater than 0');
    }

    // Maximum upload limit guard (e.g. 50 GB)
    const MAX_UPLOAD_BYTES = BigInt(50) * BigInt(1024) * BigInt(1024) * BigInt(1024);
    if (byteSizeBigInt > MAX_UPLOAD_BYTES) {
      throw new Error(`File exceeds maximum permitted upload limit of 50 GB`);
    }

    const dataset = await prisma.dataset.create({
      data: {
        name: params.name || path.basename(params.originalFilename, path.extname(params.originalFilename)) || 'dataset',
        originalFilename: params.originalFilename,
        byteSize: byteSizeBigInt,
        expectedChecksum: params.expectedChecksum ? params.expectedChecksum.toLowerCase() : null,
        ownerId: params.ownerId || null,
        delimiter: params.delimiter || ',',
        status: 'UPLOADING',
        currentStage: 'UPLOADING',
        progress: 0.0,
      },
    });

    // Initialize spool directory and empty spool file
    const spoolDir = getDatasetSpoolDir(dataset.id);
    const spoolFilePath = getDatasetSpoolFilePath(dataset.id, false);
    if (!fs.existsSync(spoolFilePath)) {
      fs.writeFileSync(spoolFilePath, Buffer.alloc(0));
    }

    await prisma.dataset.update({
      where: { id: dataset.id },
      data: { spoolPath: spoolFilePath },
    });

    return {
      datasetId: dataset.id,
      name: dataset.name,
      byteSize: dataset.byteSize.toString(),
      status: dataset.status,
      spoolDir,
    };
  }

  /**
   * Positional chunk upload with strict idempotency and boundary validation.
   */
  static async uploadChunk(params: UploadChunkParams) {
    const { datasetId, chunkIndex, chunkBuffer, chunkChecksum } = params;
    const byteOffsetBigInt = BigInt(params.byteOffset);
    const byteLength = chunkBuffer.length;

    if (byteOffsetBigInt < BigInt(0)) {
      throw new Error(`Invalid negative byteOffset: ${params.byteOffset}`);
    }
    if (byteLength <= 0) {
      throw new Error(`Chunk byteLength must be greater than 0`);
    }

    const dataset = await prisma.dataset.findUnique({
      where: { id: datasetId },
    });

    if (!dataset) {
      throw new Error(`Dataset ${datasetId} not found`);
    }
    if (dataset.status !== 'UPLOADING') {
      throw new Error(`Cannot upload chunk to dataset in status ${dataset.status}`);
    }

    if (byteOffsetBigInt + BigInt(byteLength) > dataset.byteSize) {
      throw new Error(`Chunk range (${byteOffsetBigInt}..${byteOffsetBigInt + BigInt(byteLength)}) exceeds total dataset byte size (${dataset.byteSize})`);
    }

    // Compute actual chunk checksum
    const computedChunkSha256 = crypto.createHash('sha256').update(chunkBuffer).digest('hex').toLowerCase();
    if (chunkChecksum && chunkChecksum.toLowerCase() !== computedChunkSha256) {
      throw new Error(`Chunk checksum mismatch: expected ${chunkChecksum}, computed ${computedChunkSha256}`);
    }

    // Check existing chunk for idempotency
    const existingChunk = await prisma.uploadChunk.findUnique({
      where: {
        datasetId_chunkIndex: {
          datasetId,
          chunkIndex,
        },
      },
    });

    if (existingChunk) {
      if (
        existingChunk.byteOffset === byteOffsetBigInt &&
        existingChunk.byteLength === byteLength &&
        existingChunk.checksumSha256 === computedChunkSha256
      ) {
        // Idempotent duplicate: safe no-op
        return {
          chunkIndex,
          byteOffset: byteOffsetBigInt.toString(),
          byteLength,
          checksumSha256: computedChunkSha256,
          isDuplicate: true,
        };
      } else {
        throw new Error(`Conflicting chunk data for chunkIndex ${chunkIndex}`);
      }
    }

    // Write chunk directly to positional offset in spool file
    const spoolFilePath = getDatasetSpoolFilePath(datasetId, false);
    const fd = fs.openSync(spoolFilePath, 'r+');
    try {
      fs.writeSync(fd, chunkBuffer, 0, byteLength, Number(byteOffsetBigInt));
    } finally {
      fs.closeSync(fd);
    }

    // Record chunk in DB
    await prisma.uploadChunk.create({
      data: {
        datasetId,
        chunkIndex,
        byteOffset: byteOffsetBigInt,
        byteLength,
        checksumSha256: computedChunkSha256,
        isCommitted: true,
      },
    });

    // Update dataset progress
    const uploadedAgg = await prisma.uploadChunk.aggregate({
      where: { datasetId },
      _sum: { byteLength: true },
    });
    const totalUploadedBytes = BigInt(uploadedAgg._sum.byteLength || 0);
    const progress = Math.min(1.0, Number(totalUploadedBytes) / Number(dataset.byteSize));

    await prisma.dataset.update({
      where: { id: datasetId },
      data: { progress },
    });

    return {
      chunkIndex,
      byteOffset: byteOffsetBigInt.toString(),
      byteLength,
      checksumSha256: computedChunkSha256,
      progress,
      isDuplicate: false,
    };
  }

  /**
   * Finalize upload session: validates full contiguous coverage and whole-file checksum,
   * atomically promotes spool file to source.csv, and enqueues BullMQ import job.
   */
  static async finalizeUpload(params: FinalizeUploadParams) {
    const { datasetId, expectedChecksum } = params;

    const dataset = await prisma.dataset.findUnique({
      where: { id: datasetId },
      include: {
        uploadChunks: {
          orderBy: { byteOffset: 'asc' },
        },
      },
    });

    if (!dataset) {
      throw new Error(`Dataset ${datasetId} not found`);
    }
    if (dataset.status !== 'UPLOADING') {
      throw new Error(`Cannot finalize dataset in status ${dataset.status}`);
    }

    const chunks = dataset.uploadChunks;
    if (chunks.length === 0) {
      throw new Error(`No upload chunks found for dataset ${datasetId}`);
    }

    // Verify contiguous byte range coverage
    let currentExpectedOffset = BigInt(0);
    for (const chunk of chunks) {
      if (chunk.byteOffset !== currentExpectedOffset) {
        throw new Error(
          `Gap detected in upload chunks: expected offset ${currentExpectedOffset}, found ${chunk.byteOffset} at chunk ${chunk.chunkIndex}`
        );
      }
      currentExpectedOffset += BigInt(chunk.byteLength);
    }

    if (currentExpectedOffset !== dataset.byteSize) {
      throw new Error(
        `Total chunks coverage (${currentExpectedOffset} bytes) does not match expected dataset size (${dataset.byteSize} bytes)`
      );
    }

    // Verify physical file size on disk
    const spoolPath = getDatasetSpoolFilePath(datasetId, false);
    const stat = fs.statSync(spoolPath);
    if (BigInt(stat.size) !== dataset.byteSize) {
      throw new Error(
        `Physical spool file size (${stat.size} bytes) does not match expected size (${dataset.byteSize} bytes)`
      );
    }

    // Stream spool file with 64KB bounded buffer to compute SHA-256
    const computedFileSha256 = await new Promise<string>((resolve, reject) => {
      const hash = crypto.createHash('sha256');
      const stream = fs.createReadStream(spoolPath, { highWaterMark: 64 * 1024 });
      stream.on('data', (chunk) => hash.update(chunk));
      stream.on('end', () => resolve(hash.digest('hex').toLowerCase()));
      stream.on('error', reject);
    });

    const targetChecksum = (expectedChecksum || dataset.expectedChecksum || '').toLowerCase();
    if (targetChecksum && targetChecksum !== computedFileSha256) {
      throw new Error(
        `Whole-file checksum mismatch: expected ${targetChecksum}, computed ${computedFileSha256}`
      );
    }

    // Atomically rename source.csv.spool -> source.csv
    const finalSpoolPath = getDatasetSpoolFilePath(datasetId, true);
    if (fs.existsSync(finalSpoolPath)) {
      fs.unlinkSync(finalSpoolPath);
    }
    fs.renameSync(spoolPath, finalSpoolPath);

    // Update status to UPLOADED, then enqueue BullMQ job
    await prisma.dataset.update({
      where: { id: datasetId },
      data: {
        status: 'UPLOADED',
        currentStage: 'UPLOADED',
        actualChecksum: computedFileSha256,
        spoolPath: finalSpoolPath,
        progress: 1.0,
      },
    });

    // Enqueue BullMQ job
    const jobId = await enqueueDatasetImportJob(datasetId);

    const queuedDataset = await prisma.dataset.update({
      where: { id: datasetId },
      data: {
        status: 'QUEUED',
        currentStage: 'QUEUED',
      },
    });

    return {
      datasetId: queuedDataset.id,
      status: queuedDataset.status,
      byteSize: queuedDataset.byteSize.toString(),
      actualChecksum: computedFileSha256,
      jobId,
      spoolPath: finalSpoolPath,
    };
  }
}
