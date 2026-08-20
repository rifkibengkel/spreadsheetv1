import { NextResponse } from 'next/server';
import path from 'path';
import os from 'os';
import fs from 'fs/promises';

export async function POST(req: Request) {
  try {
    const formData = await req.formData();
    const file = formData.get('file') as File | null;
    let filename = (formData.get('filename') as string) || '';

    if (!file) {
      return NextResponse.json({ error: 'No file provided in request.' }, { status: 400 });
    }

    if (!filename) {
      filename = file.name || `spreadsheet-${Date.now()}.xlsx`;
    }

    // Ensure valid filename (strip invalid path characters)
    filename = filename.replace(/[/\\?%*:|"<>]/g, '_');

    const arrayBuffer = await file.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);

    // Resolve Desktop folder
    const userHome = os.homedir();
    const desktopDir = path.join(userHome, 'Desktop');

    await fs.mkdir(desktopDir, { recursive: true });

    let filePath = path.join(desktopDir, filename);

    try {
      await fs.writeFile(filePath, buffer);
    } catch (writeErr: any) {
      // Handle Windows EBUSY or EPERM if file is locked by Excel or another process
      if (writeErr.code === 'EBUSY' || writeErr.code === 'EPERM') {
        const ext = path.extname(filename);
        const nameWithoutExt = path.basename(filename, ext);
        const timestampStr = new Date().toISOString().replace(/[:.]/g, '-').slice(11, 19);
        filename = `${nameWithoutExt}_${timestampStr}${ext}`;
        filePath = path.join(desktopDir, filename);
        await fs.writeFile(filePath, buffer);
      } else {
        throw writeErr;
      }
    }

    return NextResponse.json({
      success: true,
      message: `File saved successfully to Desktop`,
      filename,
      filePath,
      sizeBytes: buffer.length,
    });
  } catch (error: any) {
    console.error('Error saving to Desktop:', error);
    return NextResponse.json(
      { error: error.message || 'Failed to save file to Desktop' },
      { status: 500 }
    );
  }
}
