/**
 * Phase 1A Automated Failure & Edge Case Test Suite
 * Tests upload validation, idempotency, checksum enforcement, gap rejection, and error recovery.
 */

if (typeof process.loadEnvFile === 'function') {
  try { process.loadEnvFile('.env'); } catch {}
}
import { PrismaClient } from '@prisma/client';
import crypto from 'crypto';
import fs from 'fs';
import path from 'path';
import { DatasetUploadService } from '../../src/lib/largeCsv/datasetUploadService';

const prisma = new PrismaClient();

async function runFailureTests() {
  console.log('================================================================================');
  console.log('   PHASE 1A AUTOMATED FAILURE & VALIDATION TEST SUITE');
  console.log('================================================================================\n');

  let passed = 0;
  let failed = 0;

  function assert(condition: boolean, name: string) {
    if (condition) {
      console.log(`  [PASS] ${name}`);
      passed++;
    } else {
      console.error(`  [FAIL] ${name}`);
      failed++;
    }
  }

  // Test 1: Create session with zero byte size
  try {
    await DatasetUploadService.createUploadSession({
      name: 'Zero Size Test',
      originalFilename: 'zero.csv',
      byteSize: 0,
    });
    assert(false, 'Zero byteSize rejected');
  } catch (err: any) {
    assert(err.message.includes('greater than 0'), 'Zero byteSize rejected');
  }

  // Test 2: Create valid session
  const testSize = 1000;
  const session = await DatasetUploadService.createUploadSession({
    name: 'Failure Test Dataset',
    originalFilename: 'failure_test.csv',
    byteSize: testSize,
  });
  assert(!!session.datasetId && session.status === 'UPLOADING', 'Valid session created with UPLOADING status');

  const chunk0Data = Buffer.from('a'.repeat(500));
  const chunk0Hash = crypto.createHash('sha256').update(chunk0Data).digest('hex');

  // Test 3: Upload valid chunk 0
  const c0 = await DatasetUploadService.uploadChunk({
    datasetId: session.datasetId,
    chunkIndex: 0,
    byteOffset: 0,
    chunkBuffer: chunk0Data,
    chunkChecksum: chunk0Hash,
  });
  assert(c0.isDuplicate === false && c0.chunkIndex === 0, 'Chunk 0 uploaded successfully');

  // Test 4: Idempotent duplicate retry of chunk 0
  const c0Dup = await DatasetUploadService.uploadChunk({
    datasetId: session.datasetId,
    chunkIndex: 0,
    byteOffset: 0,
    chunkBuffer: chunk0Data,
    chunkChecksum: chunk0Hash,
  });
  assert(c0Dup.isDuplicate === true, 'Duplicate chunk 0 accepted idempotently without duplication');

  // Test 5: Conflicting chunk 0 (different content/hash)
  try {
    const fakeData = Buffer.from('b'.repeat(500));
    await DatasetUploadService.uploadChunk({
      datasetId: session.datasetId,
      chunkIndex: 0,
      byteOffset: 0,
      chunkBuffer: fakeData,
    });
    assert(false, 'Conflicting chunk 0 rejected');
  } catch (err: any) {
    assert(err.message.includes('Conflicting'), 'Conflicting chunk 0 rejected');
  }

  // Test 6: Checksum mismatch on new chunk
  try {
    const chunk1Data = Buffer.from('c'.repeat(500));
    await DatasetUploadService.uploadChunk({
      datasetId: session.datasetId,
      chunkIndex: 1,
      byteOffset: 500,
      chunkBuffer: chunk1Data,
      chunkChecksum: '0000000000000000000000000000000000000000000000000000000000000000',
    });
    assert(false, 'Checksum mismatch rejected');
  } catch (err: any) {
    assert(err.message.includes('checksum mismatch'), 'Checksum mismatch rejected');
  }

  // Test 7: Negative offset rejected
  try {
    await DatasetUploadService.uploadChunk({
      datasetId: session.datasetId,
      chunkIndex: 1,
      byteOffset: -10,
      chunkBuffer: Buffer.from('test'),
    });
    assert(false, 'Negative offset rejected');
  } catch (err: any) {
    assert(err.message.includes('negative'), 'Negative offset rejected');
  }

  // Test 8: Offset exceeding total file size rejected
  try {
    await DatasetUploadService.uploadChunk({
      datasetId: session.datasetId,
      chunkIndex: 1,
      byteOffset: 950,
      chunkBuffer: Buffer.from('x'.repeat(100)), // 950 + 100 = 1050 > 1000
    });
    assert(false, 'Exceeding total dataset size rejected');
  } catch (err: any) {
    assert(err.message.includes('exceeds total dataset byte size'), 'Exceeding total dataset size rejected');
  }

  // Test 9: Gap detection during finalization (chunk 0 uploaded [0..500], chunk 1 missing, finalize called)
  try {
    await DatasetUploadService.finalizeUpload({
      datasetId: session.datasetId,
    });
    assert(false, 'Finalization with missing chunks rejected');
  } catch (err: any) {
    assert(err.message.includes('coverage') || err.message.includes('Gap'), 'Finalization with missing chunks rejected');
  }

  // Test 10: Complete chunk 1 and finalize
  console.log('  [Step] Uploading valid chunk 1...');
  const chunk1Valid = Buffer.from('d'.repeat(500));
  await DatasetUploadService.uploadChunk({
    datasetId: session.datasetId,
    chunkIndex: 1,
    byteOffset: 500,
    chunkBuffer: chunk1Valid,
  });
  console.log('  [Step] Chunk 1 uploaded. Calling finalizeUpload...');

  const fin = await DatasetUploadService.finalizeUpload({
    datasetId: session.datasetId,
  });
  console.log('  [Step] finalizeUpload completed:', fin.status);
  assert(fin.status === 'QUEUED' && !!fin.actualChecksum, 'Finalization completed with QUEUED status and valid checksum');

  // Clean up failure test dataset
  await prisma.uploadChunk.deleteMany({ where: { datasetId: session.datasetId } });
  await prisma.dataset.delete({ where: { id: session.datasetId } });
  try {
    fs.rmSync(path.join(process.cwd(), 'storage', 'datasets', 'spool', session.datasetId), { recursive: true, force: true });
  } catch {}

  console.log(`\nFailure & Edge Case Tests Complete: ${passed} passed, ${failed} failed.\n`);
  const { closeDatasetImportQueue } = require('../../src/lib/largeCsv/queue/datasetImportQueue');
  await closeDatasetImportQueue();
  await prisma.$disconnect();

  process.exit(failed > 0 ? 1 : 0);
}

runFailureTests().catch(err => {
  console.error('Fatal test runner error:', err);
  process.exit(1);
});
