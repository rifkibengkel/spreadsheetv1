import { NextResponse } from 'next/server';
import path from 'path';
import os from 'os';
import fs from 'fs/promises';
import { requireAuth, verifyCsrfOrigin, AuthError } from '@/lib/auth/rbacGuard';

const MAX_FILE_SIZE_BYTES = 50 * 1024 * 1024; // 50 MB
const ALLOWED_EXTENSIONS = new Set(['.xlsx', '.csv']);
const WINDOWS_RESERVED_NAMES = /^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])$/i;

export async function POST(req: Request) {
  try {
    // 1. Disabled in production
    if (process.env.NODE_ENV === 'production') {
      return NextResponse.json(
        { success: false, error: 'Fitur Desktop Save dinonaktifkan pada lingkungan produksi.' },
        { status: 403 }
      );
    }

    // 2. CSRF Origin Verification
    if (!verifyCsrfOrigin(req)) {
      return NextResponse.json(
        { success: false, error: 'Akses ditolak: Verifikasi CSRF origin gagal.' },
        { status: 403 }
      );
    }

    // 3. Strict Authentication Requirement
    const user = await requireAuth(req);

    // 4. Parse multipart form data
    const formData = await req.formData();
    const file = formData.get('file') as File | null;
    let rawFilename = (formData.get('filename') as string) || '';

    if (!file) {
      return NextResponse.json({ success: false, error: 'File tidak ditemukan dalam request.' }, { status: 400 });
    }

    // 5. Size Limit Check
    if (file.size > MAX_FILE_SIZE_BYTES) {
      return NextResponse.json(
        { success: false, error: 'Ukuran file melebihi batas maksimum 50MB.' },
        { status: 413 }
      );
    }

    if (!rawFilename) {
      rawFilename = file.name || `spreadsheet-${Date.now()}.xlsx`;
    }

    // 6. Extension Allowlist Check
    const ext = path.extname(rawFilename).toLowerCase();
    if (!ALLOWED_EXTENSIONS.has(ext)) {
      return NextResponse.json(
        {
          success: false,
          error: `Ekstensi file '${ext}' tidak diizinkan. Hanya file .xlsx dan .csv yang diperbolehkan.`,
        },
        { status: 400 }
      );
    }

    // 7. Reserved Device Name Protection (Windows CON, PRN, AUX, NUL, COM1-9, LPT1-9)
    let baseName = path.basename(rawFilename, ext).trim();
    if (WINDOWS_RESERVED_NAMES.test(baseName)) {
      return NextResponse.json(
        {
          success: false,
          error: `Nama file '${baseName}' menggunakan reserved device name sistem operasi.`,
        },
        { status: 400 }
      );
    }

    // 8. Sanitize Filename (strip dangerous characters, slashes, null bytes)
    baseName = baseName.replace(/[/\\?%*:|"<> \x00-\x1F]/g, '_').trim();
    if (!baseName) {
      baseName = `spreadsheet-${Date.now()}`;
    }
    const safeFilename = `${baseName}${ext}`;

    // 9. Strict Path Validation & Directory Confinement
    const userHome = os.homedir();
    const desktopDir = path.resolve(userHome, 'Desktop');
    await fs.mkdir(desktopDir, { recursive: true });

    const targetFilePath = path.resolve(desktopDir, safeFilename);

    // Confinement assertion: must be directly within desktopDir
    const relative = path.relative(desktopDir, targetFilePath);
    if (relative.startsWith('..') || path.isAbsolute(relative) || relative.includes(path.sep)) {
      return NextResponse.json(
        { success: false, error: 'Path traversal terdeteksi. Akses ditolak.' },
        { status: 400 }
      );
    }

    const arrayBuffer = await file.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);

    let finalFilePath = targetFilePath;
    let finalFilename = safeFilename;

    try {
      await fs.writeFile(finalFilePath, buffer);
    } catch (writeErr: any) {
      // Handle Windows EBUSY or EPERM if file is locked by Excel or another process
      if (writeErr.code === 'EBUSY' || writeErr.code === 'EPERM') {
        const timestampStr = new Date().toISOString().replace(/[:.]/g, '-').slice(11, 19);
        finalFilename = `${baseName}_${timestampStr}${ext}`;
        finalFilePath = path.resolve(desktopDir, finalFilename);
        await fs.writeFile(finalFilePath, buffer);
      } else {
        throw writeErr;
      }
    }

    return NextResponse.json({
      success: true,
      message: 'File berhasil disimpan ke Desktop.',
      filename: finalFilename,
      filePath: finalFilePath,
      sizeBytes: buffer.length,
      savedBy: user.username,
    });
  } catch (error: any) {
    if (error instanceof AuthError) {
      return NextResponse.json(
        { success: false, error: error.message },
        { status: error.statusCode }
      );
    }
    console.error('Error saving to Desktop:', error);
    return NextResponse.json(
      { success: false, error: error.message || 'Gagal menyimpan file ke Desktop.' },
      { status: 500 }
    );
  }
}

