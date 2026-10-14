import { NextRequest } from 'next/server';
import prisma from '../../src/lib/db/prisma';
import { GET } from '../../src/app/api/v1/datasets/[id]/range/route';
import { DatasetRangeReader, DatasetRangeError } from '../../src/lib/largeCsv/reader/datasetRangeReader';
import { createSession } from '../../src/lib/auth/session';
import fs from 'fs';
import path from 'path';

const REAL_DATASET_ID = '80af4daa-3c4a-42eb-b656-dd855dc6f866';

let passed = 0;
let failed = 0;

function assert(condition: boolean, name: string, detail?: any) {
  if (condition) {
    console.log(`  [PASS] ${name}`);
    passed++;
  } else {
    console.error(`  [FAIL] ${name}`, detail !== undefined ? detail : '');
    failed++;
  }
}

function makeRequest(urlStr: string, cookie?: string): NextRequest {
  const headers: Record<string, string> = {};
  if (cookie) {
    headers['cookie'] = cookie;
  }
  return new NextRequest(new URL(urlStr), { headers });
}

async function runApiTests() {
  console.log('='.repeat(80));
  console.log('   PHASE 1B — STEP 5: API INTEGRATION & FAILURE TESTS');
  console.log(`   Target Dataset: ${REAL_DATASET_ID}`);
  console.log('='.repeat(80));

  // ---------------------------------------------------------------------------
  // 1. Valid API Requests
  // ---------------------------------------------------------------------------
  console.log('\n[1. Valid API Requests]');
  {
    const req = makeRequest(`http://localhost:3000/api/v1/datasets/${REAL_DATASET_ID}/range?startRow=0&rowCount=50`);
    const res = await GET(req, { params: Promise.resolve({ id: REAL_DATASET_ID }) });
    assert(res.status === 200, 'Valid request returns 200 OK');
    const data = await res.json();
    assert(data.rowCount === 50, `Returns 50 rows (got ${data.rowCount})`);
    assert(data.datasetId === REAL_DATASET_ID, 'Matches datasetId');
    assert(data.columnCount === 54, 'columnCount is 54');
    assert(data.nextRow === 50, 'nextRow is 50');
    assert(data.metrics && typeof data.metrics.serverTotalMs === 'number', 'Contains structured metrics');
  }

  // Cross-partition API request
  {
    const req = makeRequest(`http://localhost:3000/api/v1/datasets/${REAL_DATASET_ID}/range?startRow=1999950&rowCount=100`);
    const res = await GET(req, { params: Promise.resolve({ id: REAL_DATASET_ID }) });
    assert(res.status === 200, 'Cross-partition request returns 200 OK');
    const data = await res.json();
    assert(data.rowCount === 100, `Returns 100 rows across boundary`);
    assert(data.rows[0].stableRowId === 1999950, 'First row is 1,999,950');
    assert(data.rows[99].stableRowId === 2000049, 'Last row is 2,000,049');
    assert(data.nextRow === 2000050, 'nextRow is 2,000,050');
  }

  // ---------------------------------------------------------------------------
  // 2. Query Parameter Validation Failures (400)
  // ---------------------------------------------------------------------------
  console.log('\n[2. Parameter Validation Failures]');
  {
    // Missing startRow
    const req = makeRequest(`http://localhost:3000/api/v1/datasets/${REAL_DATASET_ID}/range?rowCount=100`);
    const res = await GET(req, { params: Promise.resolve({ id: REAL_DATASET_ID }) });
    assert(res.status === 400, 'Missing startRow returns 400');
    const data = await res.json();
    assert(data.code === 'MISSING_START_ROW', 'Returns code MISSING_START_ROW');
  }

  {
    // Negative startRow
    const req = makeRequest(`http://localhost:3000/api/v1/datasets/${REAL_DATASET_ID}/range?startRow=-5&rowCount=100`);
    const res = await GET(req, { params: Promise.resolve({ id: REAL_DATASET_ID }) });
    assert(res.status === 400, 'Negative startRow returns 400');
    const data = await res.json();
    assert(data.code === 'INVALID_START_ROW', 'Returns code INVALID_START_ROW');
  }

  {
    // Zero rowCount
    const req = makeRequest(`http://localhost:3000/api/v1/datasets/${REAL_DATASET_ID}/range?startRow=0&rowCount=0`);
    const res = await GET(req, { params: Promise.resolve({ id: REAL_DATASET_ID }) });
    assert(res.status === 400, 'Zero rowCount returns 400');
    const data = await res.json();
    assert(data.code === 'INVALID_ROW_COUNT', 'Returns code INVALID_ROW_COUNT');
  }

  {
    // Negative rowCount
    const req = makeRequest(`http://localhost:3000/api/v1/datasets/${REAL_DATASET_ID}/range?startRow=0&rowCount=-10`);
    const res = await GET(req, { params: Promise.resolve({ id: REAL_DATASET_ID }) });
    assert(res.status === 400, 'Negative rowCount returns 400');
    const data = await res.json();
    assert(data.code === 'INVALID_ROW_COUNT', 'Returns code INVALID_ROW_COUNT');
  }

  {
    // Exceeding MAX_LIMIT (5001)
    const req = makeRequest(`http://localhost:3000/api/v1/datasets/${REAL_DATASET_ID}/range?startRow=0&rowCount=5001`);
    const res = await GET(req, { params: Promise.resolve({ id: REAL_DATASET_ID }) });
    assert(res.status === 400, 'rowCount > 5000 returns 400');
    const data = await res.json();
    assert(data.code === 'ROW_COUNT_EXCEEDS_LIMIT', 'Returns code ROW_COUNT_EXCEEDS_LIMIT');
  }

  {
    // startRow >= totalRows
    const req = makeRequest(`http://localhost:3000/api/v1/datasets/${REAL_DATASET_ID}/range?startRow=12000000&rowCount=100`);
    const res = await GET(req, { params: Promise.resolve({ id: REAL_DATASET_ID }) });
    assert(res.status === 400, 'startRow out of bounds returns 400');
    const data = await res.json();
    assert(data.code === 'START_ROW_OUT_OF_BOUNDS', 'Returns code START_ROW_OUT_OF_BOUNDS');
  }

  // ---------------------------------------------------------------------------
  // 3. Dataset Existence & Status Failures (404 / 409)
  // ---------------------------------------------------------------------------
  console.log('\n[3. Dataset Existence & Status Failures]');
  {
    // Non-existent dataset
    const fakeId = 'ffffffff-ffff-ffff-ffff-ffffffffffff';
    const req = makeRequest(`http://localhost:3000/api/v1/datasets/${fakeId}/range?startRow=0&rowCount=100`);
    const res = await GET(req, { params: Promise.resolve({ id: fakeId }) });
    assert(res.status === 404, 'Non-existent dataset returns 404');
    const data = await res.json();
    assert(data.code === 'DATASET_NOT_FOUND', 'Returns code DATASET_NOT_FOUND');
  }

  {
    // Dataset not in READY status (simulate with temp dataset)
    const tempDataset = await prisma.dataset.create({
      data: {
        name: 'Unready Test Dataset',
        originalFilename: 'test.csv',
        byteSize: BigInt(1000),
        status: 'PARSING',
        rowCount: BigInt(100),
      },
    });

    try {
      const req = makeRequest(`http://localhost:3000/api/v1/datasets/${tempDataset.id}/range?startRow=0&rowCount=10`);
      const res = await GET(req, { params: Promise.resolve({ id: tempDataset.id }) });
      assert(res.status === 409, 'Non-READY dataset returns 409 Conflict');
      const data = await res.json();
      assert(data.code === 'DATASET_NOT_READY', 'Returns code DATASET_NOT_READY');
    } finally {
      await prisma.dataset.delete({ where: { id: tempDataset.id } });
    }
  }

  // ---------------------------------------------------------------------------
  // 4. Authentication & Authorization Failures (401 / 403)
  // ---------------------------------------------------------------------------
  console.log('\n[4. Authentication & Authorization Failures]');
  {
    // Find or create test users
    let ownerUser = await prisma.user.findFirst({ where: { username: 'test_dataset_owner' } });
    if (!ownerUser) {
      ownerUser = await prisma.user.create({
        data: {
          username: 'test_dataset_owner',
          email: 'owner@test.local',
          passwordHash: 'dummyhash',
          fullName: 'Test Dataset Owner',
          status: 'ACTIVE',
        },
      });
    }

    let otherUser = await prisma.user.findFirst({ where: { username: 'test_other_user' } });
    if (!otherUser) {
      otherUser = await prisma.user.create({
        data: {
          username: 'test_other_user',
          email: 'other@test.local',
          passwordHash: 'dummyhash',
          fullName: 'Other User',
          status: 'ACTIVE',
        },
      });
    }

    // Create an owned dataset
    const ownedDataset = await prisma.dataset.create({
      data: {
        ownerId: ownerUser.id,
        name: 'Owned Private Dataset',
        originalFilename: 'private.csv',
        byteSize: BigInt(1000),
        status: 'READY',
        rowCount: BigInt(100),
      },
    });

    try {
      // 4.1 Unauthenticated access to owned dataset -> 401
      const reqUnauth = makeRequest(`http://localhost:3000/api/v1/datasets/${ownedDataset.id}/range?startRow=0&rowCount=10`);
      const resUnauth = await GET(reqUnauth, { params: Promise.resolve({ id: ownedDataset.id }) });
      assert(resUnauth.status === 401, 'Unauthenticated access to owned dataset returns 401');

      // 4.2 Non-owner access -> 403
      const { sessionToken: otherToken } = await createSession(otherUser.id);
      const reqOther = makeRequest(
        `http://localhost:3000/api/v1/datasets/${ownedDataset.id}/range?startRow=0&rowCount=10`,
        `__session=${otherToken}`
      );
      const resOther = await GET(reqOther, { params: Promise.resolve({ id: ownedDataset.id }) });
      assert(resOther.status === 403, 'Unauthorized user access to owned dataset returns 403');

      // 4.3 Owner access -> reaches reader (fails with NO_READY_PARTITIONS because no partitions attached, but passed auth!)
      const { sessionToken: ownerToken } = await createSession(ownerUser.id);
      const reqOwner = makeRequest(
        `http://localhost:3000/api/v1/datasets/${ownedDataset.id}/range?startRow=0&rowCount=10`,
        `__session=${ownerToken}`
      );
      const resOwner = await GET(reqOwner, { params: Promise.resolve({ id: ownedDataset.id }) });
      assert(resOwner.status !== 401 && resOwner.status !== 403, 'Owner bypasses auth/authz checks successfully');
    } finally {
      await prisma.dataset.delete({ where: { id: ownedDataset.id } });
      await prisma.session.deleteMany({ where: { userId: { in: [ownerUser.id, otherUser.id] } } });
    }
  }

  // ---------------------------------------------------------------------------
  // 5. Storage Failure Modes (Missing .bin, Missing .idx, Corrupt .idx)
  // ---------------------------------------------------------------------------
  console.log('\n[5. Storage Failure Modes]');
  {
    // Create a temporary partition with non-existent data file
    const ghostDataset = await prisma.dataset.create({
      data: {
        name: 'Ghost Files Dataset',
        originalFilename: 'ghost.csv',
        byteSize: BigInt(1000),
        status: 'READY',
        rowCount: BigInt(100),
      },
    });

    const ghostPart = await prisma.datasetPartition.create({
      data: {
        datasetId: ghostDataset.id,
        partitionNumber: 0,
        stableRowStart: BigInt(0),
        stableRowEnd: BigInt(99),
        rowCount: 100,
        dataFilePath: 'C:\\non_existent_path\\ghost.bin',
        indexFilePath: 'C:\\non_existent_path\\ghost.idx',
        byteSize: BigInt(1000),
        status: 'READY',
      },
    });

    try {
      // 5.1 Missing files return 500 with deterministic code
      const req = makeRequest(`http://localhost:3000/api/v1/datasets/${ghostDataset.id}/range?startRow=0&rowCount=10`);
      const res = await GET(req, { params: Promise.resolve({ id: ghostDataset.id }) });
      assert(res.status === 500, 'Missing partition index file returns 500');
      const data = await res.json();
      assert(data.code === 'INDEX_FILE_MISSING', `Returns deterministic code INDEX_FILE_MISSING (got ${data.code})`);
    } finally {
      await prisma.datasetPartition.delete({ where: { id: ghostPart.id } });
      await prisma.dataset.delete({ where: { id: ghostDataset.id } });
    }
  }

  console.log('\n' + '='.repeat(80));
  console.log(`STEP 5 TEST RESULTS: ${passed} PASSED, ${failed} FAILED`);
  console.log('='.repeat(80));

  if (failed > 0) {
    process.exit(1);
  }
}

runApiTests().catch((err) => {
  console.error('Fatal API test error:', err);
  process.exit(1);
}).finally(() => process.exit(0));
