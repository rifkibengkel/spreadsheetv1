import ExcelJS from 'exceljs';
import JSZip from 'jszip';

export interface RawSheetData {
  sheetId: string;
  name: string;
  rows: {
    number: number;
    cells: { colNumber: number; v?: any; f?: string }[];
  }[];
  rowCount: number;
  columnCount: number;
}

function unescapeXml(str: string): string {
  if (!str) return '';
  return str
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

function colLetterToIndex(colStr: string): number {
  let index = 0;
  for (let i = 0; i < colStr.length; i++) {
    index = index * 26 + (colStr.charCodeAt(i) - 64);
  }
  return index - 1;
}

function parseCellRef(ref: string): { row: number; col: number } {
  let colStr = '';
  let rowStr = '';
  for (let i = 0; i < ref.length; i++) {
    const code = ref.charCodeAt(i);
    if ((code >= 65 && code <= 90) || (code >= 97 && code <= 122)) {
      colStr += ref[i].toUpperCase();
    } else if (code >= 48 && code <= 57) {
      rowStr += ref[i];
    }
  }
  const row = parseInt(rowStr, 10) - 1;
  const col = colLetterToIndex(colStr);
  return { row, col };
}

/**
 * Fallback incremental streaming parser for giant worksheets (>300 MB XML).
 * Avoids monolithic string allocations in V8 heap.
 */
async function parseStreamingWorkbookFallback(
  fileArrayBuffer: ArrayBuffer,
  fileName: string
): Promise<any> {
  self.postMessage({
    type: 'PROGRESS',
    percent: 20,
    message: 'Menggunakan Fallback Streaming Parser untuk lembar kerja raksasa...',
  });

  const zip = await JSZip.loadAsync(fileArrayBuffer);

  // 1. Parse workbook relationships (xl/_rels/workbook.xml.rels)
  const relMap: Record<string, string> = {};
  const relsFile = zip.files['xl/_rels/workbook.xml.rels'];
  if (relsFile) {
    const relsXml = await relsFile.async('string');
    const relRegex = /<Relationship [^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g;
    let rMatch;
    while ((rMatch = relRegex.exec(relsXml)) !== null) {
      relMap[rMatch[1]] = rMatch[2];
    }
  }

  // 2. Parse workbook structure (xl/workbook.xml)
  const wbFile = zip.files['xl/workbook.xml'];
  if (!wbFile) {
    throw new Error('File tidak memiliki xl/workbook.xml yang valid.');
  }

  const wbXml = await wbFile.async('string');
  const sheetEntries: { name: string; sheetId: string; target: string }[] = [];
  const sheetRegex = /<sheet [^>]*name="([^"]+)"[^>]*sheetId="([^"]+)"[^>]*r:id="([^"]+)"/g;
  let sMatch;
  while ((sMatch = sheetRegex.exec(wbXml)) !== null) {
    const name = unescapeXml(sMatch[1]);
    const rId = sMatch[3];
    let target = relMap[rId] || '';
    if (target.startsWith('worksheets/')) target = `xl/${target}`;
    else if (!target.startsWith('xl/')) target = `xl/worksheets/${target}`;
    sheetEntries.push({ name, sheetId: sMatch[2], target });
  }

  // 3. Incremental stream parse for shared strings (xl/sharedStrings.xml)
  self.postMessage({
    type: 'PROGRESS',
    percent: 30,
    message: 'Membaca tabel shared strings...',
  });

  const ssFile = zip.files['xl/sharedStrings.xml'];
  const sharedStrings: string[] = [];
  if (ssFile) {
    const decoder = new TextDecoder('utf-8');
    let leftover = '';

    await new Promise<void>((resolve, reject) => {
      const stream = (ssFile as any).internalStream('uint8array');
      stream.on('data', (chunk: Uint8Array) => {
        const text = leftover + decoder.decode(chunk, { stream: true });
        const parts = text.split('</si>');
        leftover = parts.pop() || '';
        for (let i = 0; i < parts.length; i++) {
          const part = parts[i];
          const siIdx = part.indexOf('<si>');
          if (siIdx !== -1) {
            const siContent = part.substring(siIdx + 4);
            let str = '';
            const tRegex = /<t(?: [^>]*)?>([^<]*)<\/t>/g;
            let m;
            while ((m = tRegex.exec(siContent)) !== null) {
              str += m[1];
            }
            sharedStrings.push(unescapeXml(str));
          }
        }
      });
      stream.on('end', () => {
        if (leftover && leftover.includes('<si>')) {
          const siIdx = leftover.indexOf('<si>');
          const siContent = leftover.substring(siIdx + 4);
          let str = '';
          const tRegex = /<t(?: [^>]*)?>([^<]*)<\/t>/g;
          let m;
          while ((m = tRegex.exec(siContent)) !== null) {
            str += m[1];
          }
          sharedStrings.push(unescapeXml(str));
        }
        resolve();
      });
      stream.on('error', reject);
      stream.resume();
    });
  }

  // 4. Stream each worksheet entry
  const sheetOrder: string[] = [];
  const sheetsData: any = {};
  const totalSheets = sheetEntries.length;

  for (let sIdx = 0; sIdx < totalSheets; sIdx++) {
    const entry = sheetEntries[sIdx];
    const sheetId = entry.name;
    sheetOrder.push(sheetId);

    const sheetXmlFile = zip.files[entry.target];
    const cellData: Record<number, Record<number, any>> = {};
    let maxRow = 0;
    let maxCol = 0;

    if (sheetXmlFile) {
      const decoder = new TextDecoder('utf-8');
      let leftover = '';

      await new Promise<void>((resolve, reject) => {
        const stream = (sheetXmlFile as any).internalStream('uint8array');
        stream.on('data', (chunk: Uint8Array) => {
          const text = leftover + decoder.decode(chunk, { stream: true });
          const rows = text.split('</row>');
          leftover = rows.pop() || '';

          for (let i = 0; i < rows.length; i++) {
            const rowStr = rows[i];
            const rOpenIdx = rowStr.indexOf('<row ');
            if (rOpenIdx === -1) continue;

            const rowTag = rowStr.substring(rOpenIdx);
            const rMatch = rowTag.match(/r="(\d+)"/);
            if (!rMatch) continue;
            const rowIndex = parseInt(rMatch[1], 10) - 1;
            if (rowIndex > maxRow) maxRow = rowIndex;

            const rowCellMap: Record<number, any> = {};

            const cRegex = /<c ([^>]+)(?:\/>|>([\s\S]*?)<\/c>)/g;
            let cMatch;
            while ((cMatch = cRegex.exec(rowTag)) !== null) {
              const attrs = cMatch[1];
              const body = cMatch[2] || '';

              const refMatch = attrs.match(/r="([A-Za-z0-9]+)"/);
              if (!refMatch) continue;
              const { col } = parseCellRef(refMatch[1]);
              if (col > maxCol) maxCol = col;

              const typeMatch = attrs.match(/t="([^"]+)"/);
              const cellType = typeMatch ? typeMatch[1] : '';

              let f: string | undefined = undefined;
              let v: any = undefined;
              let fType: string | undefined = undefined;
              let ref: string | undefined = undefined;

              // Extract formula <f>...</f>
              const fMatch = body.match(/<f([^>]*)>([^<]+)<\/f>/);
              if (fMatch) {
                const fAttrs = fMatch[1];
                let rawF = unescapeXml(fMatch[2]).trim();
                rawF = rawF.replace(/'\[[^'\]]+\]#REF'/g, '');
                f = rawF.startsWith('=') ? rawF : `=${rawF}`;
                if (fAttrs.includes('t="array"')) {
                  fType = 'array';
                  const refM = fAttrs.match(/ref="([^"]+)"/);
                  if (refM) ref = refM[1];
                }
              }

              // Extract value <v>...</v>
              const vMatch = body.match(/<v[^>]*>([^<]+)<\/v>/);
              if (vMatch) {
                const rawV = unescapeXml(vMatch[1]).trim();
                if (cellType === 's') {
                  const sIndex = parseInt(rawV, 10);
                  v = sharedStrings[sIndex] !== undefined ? sharedStrings[sIndex] : rawV;
                } else if (cellType === 'b') {
                  v = rawV === '1' || rawV.toLowerCase() === 'true';
                } else {
                  const num = Number(rawV);
                  v = !isNaN(num) && rawV !== '' ? num : rawV;
                }
              } else if (cellType === 'inlineStr') {
                const isMatch = body.match(/<t[^>]*>([^<]*)<\/t>/);
                if (isMatch) {
                  v = unescapeXml(isMatch[1]);
                }
              }

              const cellInfo: any = {};
              if (f) cellInfo.f = f;
              if (fType) cellInfo.fType = fType;
              if (ref) cellInfo.ref = ref;
              if (v !== undefined && v !== null && v !== '') {
                cellInfo.v = v;
                if (typeof v === 'number') cellInfo.t = 2; // NUMBER
                else if (typeof v === 'boolean') cellInfo.t = 3; // BOOLEAN
                else cellInfo.t = 1; // STRING
              }

              if (Object.keys(cellInfo).length > 0) {
                rowCellMap[col] = cellInfo;
              }
            }

            if (Object.keys(rowCellMap).length > 0) {
              cellData[rowIndex] = rowCellMap;
            }
          }
        });

        stream.on('end', () => {
          resolve();
        });
        stream.on('error', reject);
        stream.resume();
      });
    }

    sheetsData[sheetId] = {
      id: sheetId,
      name: entry.name,
      type: 2, // SheetType.GRID - CRITICAL FOR FORMULA ENGINE REGISTRATION
      status: sIdx === 0 ? 1 : 0,
      hidden: 0,
      rowCount: Math.max(maxRow + 50, 100),
      columnCount: Math.max(maxCol + 10, 30),
      zoomRatio: 1,
      scrollTop: 0,
      scrollLeft: 0,
      defaultColumnWidth: 88,
      defaultRowHeight: 24,
      showGridlines: 1,
      rightToLeft: 0,
      rowHeader: { width: 46 },
      columnHeader: { height: 20 },
      cellData,
      columnData: {},
      rowData: {},
    };

    const percent = 35 + Math.round(((sIdx + 1) / totalSheets) * 60);
    self.postMessage({
      type: 'PROGRESS',
      percent,
      message: `Memproses data sheet ${entry.name} (${sIdx + 1}/${totalSheets})...`,
    });
  }

  return {
    id: `workbook-imported-${Date.now()}`,
    name: fileName,
    sheetOrder,
    appVersion: '3.0.0-alpha',
    sheets: sheetsData,
  };
}

self.onmessage = async (e: MessageEvent) => {
  const { fileArrayBuffer, fileName } = e.data;

  try {
    self.postMessage({ type: 'PROGRESS', percent: 5, message: 'Membaca file di Background Worker (5%)...' });

    let workbookData: any = null;

    // 1. Check if workbook has a giant worksheet XML (>300 MB uncompressed)
    let needsStreamingFallback = false;
    try {
      const zipInspect = await JSZip.loadAsync(fileArrayBuffer.slice(0));
      for (const fName of Object.keys(zipInspect.files)) {
        if (fName.startsWith('xl/worksheets/sheet') && fName.endsWith('.xml')) {
          const entry = zipInspect.files[fName];
          const uncompressedSize = (entry as any)?._data?.uncompressedSize || 0;
          if (uncompressedSize > 300 * 1024 * 1024) {
            needsStreamingFallback = true;
            console.log(`[Excel Worker] Giant worksheet detected (${fName}: ${(uncompressedSize / (1024 * 1024)).toFixed(1)} MB). Using streaming fallback.`);
            break;
          }
        }
      }
    } catch {
      // If zip inspection fails, proceed with standard ExcelJS
    }

    if (needsStreamingFallback) {
      workbookData = await parseStreamingWorkbookFallback(fileArrayBuffer, fileName);
    } else {
      try {
        let workbook: ExcelJS.Workbook | null = new ExcelJS.Workbook();
        self.postMessage({ type: 'PROGRESS', percent: 15, message: 'Mengurai struktur XML Excel di Worker (15%)...' });

        await workbook.xlsx.load(fileArrayBuffer);

        self.postMessage({ type: 'PROGRESS', percent: 25, message: 'Menyiapkan metadata dan mengurai semua sheet...' });

        const sheetOrder: string[] = [];
        const sheetsData: any = {};
        const totalSheets = workbook.worksheets.length;

        workbook.worksheets.forEach((ws, sIdx) => {
          const sheetId = ws.name;
          sheetOrder.push(sheetId);

          const cellData: Record<number, Record<number, any>> = {};
          let maxRow = 0;
          let maxCol = 0;

          ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
            const rowIndex = rowNumber - 1;
            if (rowIndex > maxRow) maxRow = rowIndex;

            const rowCellMap: Record<number, any> = {};

            row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
              const colIndex = colNumber - 1;
              if (colIndex > maxCol) maxCol = colIndex;

              let v: any = undefined;
              let f: string | undefined = undefined;

              // 1. Extract formula if present
              let rawFormula: string | undefined = undefined;
              if (cell.type === ExcelJS.ValueType.Formula) {
                rawFormula = cell.formula || (cell as any).model?.formula || (typeof cell.value === 'object' && cell.value !== null ? (cell.value as any).formula : undefined);
              } else if (typeof cell.value === 'object' && cell.value !== null && (cell.value as any).formula) {
                rawFormula = (cell.value as any).formula;
              } else if ((cell as any).model?.formula) {
                rawFormula = (cell as any).model.formula;
              } else if (cell.formula) {
                rawFormula = cell.formula;
              }

              if (rawFormula && typeof rawFormula === 'string' && rawFormula.trim().length > 0) {
                const trimmed = rawFormula.trim();
                f = trimmed.startsWith('=') ? trimmed : `=${trimmed}`;
              }

              // 2. Extract value / cached result
              if (cell.type === ExcelJS.ValueType.Formula) {
                const res = cell.result;
                if (res !== undefined && res !== null) {
                  if (typeof res === 'object') {
                    if ((res as any).result !== undefined && (res as any).result !== null && (res as any).result !== '') {
                      v = (res as any).result;
                    } else if ((res as any).error !== undefined) {
                      v = (res as any).error;
                    }
                  } else if (res !== '') {
                    v = res;
                  }
                }
              } else if (cell.value !== undefined && cell.value !== null) {
                const val = cell.value;
                if (typeof val === 'object') {
                  if (val instanceof Date) {
                    v = val.toISOString();
                  } else if (Array.isArray((val as any).richText)) {
                    v = (val as any).richText.map((t: any) => t.text || '').join('');
                  } else if ((val as any).text !== undefined) {
                    v = (val as any).text;
                  } else if ((val as any).result !== undefined) {
                    v = (val as any).result;
                  } else {
                    v = String(val);
                  }
                } else {
                  v = val;
                }
              }

              const cellInfo: any = {};

              if (f !== undefined && f !== null && f !== '') {
                let cleanedF = f.trim();
                if (!cleanedF.startsWith('=')) cleanedF = `=${cleanedF}`;
                cleanedF = cleanedF.replace(/'\[[^'\]]+\]#REF'/g, '');
                cellInfo.f = cleanedF;

                const model = (cell as any).model;
                if (model?.shareType === 'array' || (cell.value as any)?.shareType === 'array') {
                  cellInfo.fType = 'array';
                  if (model?.ref) cellInfo.ref = model.ref;
                }
              }

              if (v !== undefined && v !== null && v !== '') {
                if (typeof v === 'number') {
                  cellInfo.v = v;
                  cellInfo.t = 2; // CellValueType.NUMBER
                } else if (typeof v === 'boolean') {
                  cellInfo.v = v;
                  cellInfo.t = 3; // CellValueType.BOOLEAN
                } else if (typeof v === 'string') {
                  const trimmedV = v.trim();
                  const numV = Number(trimmedV);
                  if (!isNaN(numV) && trimmedV !== '') {
                    cellInfo.v = numV;
                    cellInfo.t = 2; // CellValueType.NUMBER
                  } else {
                    cellInfo.v = v;
                    cellInfo.t = 1; // CellValueType.STRING
                  }
                } else {
                  cellInfo.v = String(v);
                  cellInfo.t = 1; // CellValueType.STRING
                }
              }

              if (Object.keys(cellInfo).length > 0) {
                rowCellMap[colIndex] = cellInfo;
              }
            });

            if (Object.keys(rowCellMap).length > 0) {
              cellData[rowIndex] = rowCellMap;
            }
          });

          sheetsData[sheetId] = {
            id: sheetId,
            name: ws.name,
            type: 2, // SheetType.GRID - CRITICAL FOR FORMULA ENGINE REGISTRATION
            status: sIdx === 0 ? 1 : 0, // Sheet 0 active by default
            hidden: 0,
            rowCount: Math.max(maxRow + 50, 100),
            columnCount: Math.max(maxCol + 10, 30),
            zoomRatio: 1,
            scrollTop: 0,
            scrollLeft: 0,
            defaultColumnWidth: 88,
            defaultRowHeight: 24,
            showGridlines: 1,
            rightToLeft: 0,
            rowHeader: { width: 46 },
            columnHeader: { height: 20 },
            cellData,
            columnData: {},
            rowData: {},
          };

          const percent = 25 + Math.round(((sIdx + 1) / totalSheets) * 70);
          self.postMessage({
            type: 'PROGRESS',
            percent,
            message: `Memproses data sheet ${ws.name} (${sIdx + 1}/${totalSheets})...`,
          });
        });

        workbook = null;

        workbookData = {
          id: `workbook-imported-${Date.now()}`,
          name: fileName,
          sheetOrder,
          appVersion: '3.0.0-alpha',
          sheets: sheetsData,
        };
      } catch (excelJsErr: any) {
        console.warn('[Excel Worker] ExcelJS failed, attempting streaming fallback:', excelJsErr?.message);
        workbookData = await parseStreamingWorkbookFallback(fileArrayBuffer, fileName);
      }
    }

    self.postMessage({ type: 'PROGRESS', percent: 98, message: 'Menyinkronkan formula...' });

    // Pre-resolve scalar cross-sheet & same-sheet formulas missing 'v'
    const sheetsData = workbookData.sheets;
    for (let pass = 0; pass < 5; pass++) {
      let resolvedAny = false;
      Object.keys(sheetsData).forEach((sName) => {
        const cData = sheetsData[sName].cellData;
        Object.keys(cData).forEach((rIdxStr) => {
          const rIdx = parseInt(rIdxStr, 10);
          Object.keys(cData[rIdx]).forEach((cIdxStr) => {
            const cIdx = parseInt(cIdxStr, 10);
            const cell = cData[rIdx][cIdx];
            if (cell.f && (cell.v === undefined || cell.v === null || cell.v === '')) {
              let match = cell.f.match(/^=('([^']+)'|([A-Za-z0-9_\s]+))!([A-Za-z]+)([0-9]+)$/i);
              let targetSheetName = match ? (match[2] || match[3]) : null;
              let colStr = match ? match[4].toUpperCase() : null;
              let rowNum = match ? parseInt(match[5], 10) : null;

              if (!match) {
                const sameMatch = cell.f.match(/^=([A-Za-z]+)([0-9]+)$/i);
                if (sameMatch) {
                  targetSheetName = sName;
                  colStr = sameMatch[1].toUpperCase();
                  rowNum = parseInt(sameMatch[2], 10);
                }
              }

              if (targetSheetName && colStr && rowNum) {
                let targetCol = 0;
                for (let i = 0; i < colStr.length; i++) {
                  targetCol = targetCol * 26 + (colStr.charCodeAt(i) - 64);
                }
                targetCol -= 1;
                const targetRow = rowNum - 1;

                const targetSheet = sheetsData[targetSheetName];
                if (targetSheet && targetSheet.cellData[targetRow] && targetSheet.cellData[targetRow][targetCol]) {
                  const targetCell = targetSheet.cellData[targetRow][targetCol];
                  if (targetCell.v !== undefined && targetCell.v !== null && targetCell.v !== '') {
                    cell.v = targetCell.v;
                    cell.t = targetCell.t !== undefined ? targetCell.t : (typeof targetCell.v === 'number' ? 2 : 1);
                    resolvedAny = true;
                  }
                }
              }
            }
          });
        });
      });
      if (!resolvedAny) break;
    }

    self.postMessage({ type: 'PROGRESS', percent: 100, message: 'Impor selesai! (100%)' });

    self.postMessage({
      type: 'COMPLETE',
      workbookData,
    });
  } catch (err: any) {
    self.postMessage({
      type: 'ERROR',
      error: err?.message || 'Error processing Excel in Worker',
    });
  }
};
