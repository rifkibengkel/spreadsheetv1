import { NextResponse } from 'next/server';
import crypto from 'crypto';
import path from 'path';
import os from 'os';
import fs from 'fs';

interface DownloadItem {
  buffer: Buffer;
  filename: string;
  contentType: string;
  expiresAt: number;
}

// In-memory store for ephemeral export downloads (isolated to Node process)
const globalStore = globalThis as unknown as {
  __exportDownloadStore__?: Map<string, DownloadItem>;
};

if (!globalStore.__exportDownloadStore__) {
  globalStore.__exportDownloadStore__ = new Map();
}

const downloadStore = globalStore.__exportDownloadStore__;
const TEMP_DIR = path.join(os.tmpdir(), 'redbox_export_temp');

try {
  if (!fs.existsSync(TEMP_DIR)) {
    fs.mkdirSync(TEMP_DIR, { recursive: true });
  }
} catch (e) {
  console.warn('[API EXPORT DOWNLOAD] Failed to create temp dir:', e);
}

function cleanupExpired() {
  const now = Date.now();
  // Clean memory store
  for (const [token, item] of downloadStore.entries()) {
    if (item.expiresAt < now) {
      downloadStore.delete(token);
    }
  }

  // Clean disk temp files older than 10 minutes
  try {
    if (fs.existsSync(TEMP_DIR)) {
      const files = fs.readdirSync(TEMP_DIR);
      for (const file of files) {
        const filePath = path.join(TEMP_DIR, file);
        try {
          const stat = fs.statSync(filePath);
          if (now - stat.mtimeMs > 10 * 60 * 1000) {
            fs.unlinkSync(filePath);
          }
        } catch (_) {}
      }
    }
  } catch (_) {}
}

export async function POST(req: Request) {
  try {
    cleanupExpired();

    const url = new URL(req.url);
    let rawFilename = url.searchParams.get('filename') || url.searchParams.get('name') || req.headers.get('x-filename') || 'hexos.xlsx';
    rawFilename = decodeURIComponent(rawFilename).trim();

    // Sanitize filename for Windows & OS safety
    let cleanFilename = rawFilename.replace(/[<>:"/\\|?*\x00-\x1F]/g, '_').trim();
    cleanFilename = cleanFilename.replace(/[. ]+$/, '');
    if (!cleanFilename.toLowerCase().endsWith('.xlsx') && !cleanFilename.toLowerCase().endsWith('.csv')) {
      cleanFilename += '.xlsx';
    }

    const contentType = cleanFilename.toLowerCase().endsWith('.csv')
      ? 'text/csv; charset=utf-8'
      : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

    const arrayBuffer = await req.arrayBuffer();
    if (!arrayBuffer || arrayBuffer.byteLength === 0) {
      return NextResponse.json({ success: false, error: 'Buffer payload kosong.' }, { status: 400 });
    }

    const buffer = Buffer.from(arrayBuffer);
    const token = crypto.randomBytes(16).toString('hex');
    const expiresAt = Date.now() + 10 * 60 * 1000; // 10 minutes

    // 1. Save in memory
    downloadStore.set(token, {
      buffer,
      filename: cleanFilename,
      contentType,
      expiresAt,
    });

    // 2. Persist to disk as backup against module reloads/worker boundaries
    try {
      const metaPath = path.join(TEMP_DIR, `${token}.meta.json`);
      const dataPath = path.join(TEMP_DIR, `${token}.bin`);
      fs.writeFileSync(metaPath, JSON.stringify({ filename: cleanFilename, contentType, expiresAt }));
      fs.writeFileSync(dataPath, buffer);
    } catch (diskErr) {
      console.warn('[API EXPORT DOWNLOAD] Disk backup warning:', diskErr);
    }

    const downloadUrl = `/api/export/download?token=${token}&name=${encodeURIComponent(cleanFilename)}`;
    console.log(`[API EXPORT DOWNLOAD] POST success: token=${token}, filename=${cleanFilename}, size=${buffer.length} bytes`);

    return NextResponse.json({
      success: true,
      downloadUrl,
      filename: cleanFilename,
      sizeBytes: buffer.length,
    });
  } catch (err: any) {
    console.error('[API EXPORT DOWNLOAD] POST error:', err);
    return NextResponse.json(
      { success: false, error: err?.message || 'Gagal menyiapkan download' },
      { status: 500 }
    );
  }
}

export async function GET(req: Request) {
  try {
    cleanupExpired();

    const url = new URL(req.url);
    const token = url.searchParams.get('token');

    let item: DownloadItem | null = null;

    if (token && downloadStore.has(token)) {
      item = downloadStore.get(token)!;
    } else if (token) {
      // Check disk backup
      try {
        const metaPath = path.join(TEMP_DIR, `${token}.meta.json`);
        const dataPath = path.join(TEMP_DIR, `${token}.bin`);
        if (fs.existsSync(metaPath) && fs.existsSync(dataPath)) {
          const meta = JSON.parse(fs.readFileSync(metaPath, 'utf8'));
          const buffer = fs.readFileSync(dataPath);
          item = {
            buffer,
            filename: meta.filename,
            contentType: meta.contentType,
            expiresAt: meta.expiresAt,
          };
        }
      } catch (diskReadErr) {
        console.warn('[API EXPORT DOWNLOAD] Disk backup read warning:', diskReadErr);
      }
    }

    if (!item) {
      return new Response('File export kadaluarsa atau tidak ditemukan. Silakan lakukan ekspor ulang.', {
        status: 404,
        headers: { 'Content-Type': 'text/plain; charset=utf-8' },
      });
    }

    // Standard RFC 5987 / RFC 6266 Content-Disposition header
    const cleanAsciiName = item.filename.replace(/[^\x20-\x7E]/g, '_').replace(/"/g, '');
    const encodedFilename = encodeURIComponent(item.filename);
    const contentDisposition = `attachment; filename="${cleanAsciiName}"; filename*=UTF-8''${encodedFilename}`;

    console.log(`[API EXPORT DOWNLOAD] GET streaming: filename="${item.filename}", length=${item.buffer.length}`);

    return new Response(new Uint8Array(item.buffer), {
      status: 200,
      headers: {
        'Content-Type': item.contentType,
        'Content-Disposition': contentDisposition,
        'Content-Length': String(item.buffer.length),
        'Cache-Control': 'no-store, no-cache, must-revalidate, max-age=0',
        'Pragma': 'no-cache',
      },
    });
  } catch (err: any) {
    console.error('[API EXPORT DOWNLOAD] GET error:', err);
    return new Response('Gagal mengunduh file export.', {
      status: 500,
      headers: { 'Content-Type': 'text/plain; charset=utf-8' },
    });
  }
}
