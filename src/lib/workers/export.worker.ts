import { deflateSync, inflateSync } from 'fflate';
import Papa from 'papaparse';

function colToLetter(colIndex: number): string {
  let temp = colIndex + 1;
  let letter = '';
  while (temp > 0) {
    const mod = (temp - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    temp = Math.floor((temp - mod) / 26);
  }
  return letter;
}

function escapeXml(str: any): string {
  if (str === undefined || str === null) return '';
  const s = String(str);
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function escapeFormulaXml(str: any): string {
  if (str === undefined || str === null) return '';
  const s = String(str);
  // In OpenXML formula text (<f>), single quotes (') and double quotes (") MUST NOT be escaped.
  // Single quotes are Excel syntax for sheet references (e.g. 'Sheet Name'!A1)
  // Double quotes are Excel syntax for string literals (e.g. "valid")
  // Only &, <, and > are XML special characters in element content.
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

const CRC32_TABLE = new Int32Array(256);
for (let i = 0; i < 256; i++) {
  let c = i;
  for (let k = 0; k < 8; k++) {
    c = ((c & 1) !== 0) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
  }
  CRC32_TABLE[i] = c;
}

function crc32Update(crc: number, buf: Uint8Array): number {
  let c = crc ^ (-1);
  for (let i = 0; i < buf.length; i++) {
    c = CRC32_TABLE[(c ^ buf[i]) & 0xFF] ^ (c >>> 8);
  }
  return (c ^ (-1)) >>> 0;
}

const textEncoder = new TextEncoder();
const textDecoder = new TextDecoder();

interface ZipEntryMeta {
  filename: string;
  filenameBytes: Uint8Array;
  compMethod: number;
  modTime: number;
  modDate: number;
  crc32: number;
  compressedSize: number;
  uncompressedSize: number;
  localHeaderOffset: number;
  dataOffset: number;
}

class FastZipReader {
  public buffer: Uint8Array;
  public view: DataView;
  public entries = new Map<string, ZipEntryMeta>();
  public sheetIndexToEntry = new Map<number, ZipEntryMeta>();

  constructor(buffer: Uint8Array) {
    this.buffer = buffer;
    this.view = new DataView(buffer.buffer, buffer.byteOffset, buffer.byteLength);
    this._parse();
  }

  private _parse() {
    const len = this.buffer.length;
    let eocdOffset = -1;
    for (let i = len - 22; i >= Math.max(0, len - 65557); i--) {
      if (this.view.getUint32(i, true) === 0x06054b50) {
        eocdOffset = i;
        break;
      }
    }
    if (eocdOffset === -1) throw new Error('Invalid ZIP archive: EOCD not found');

    const totalEntries = this.view.getUint16(eocdOffset + 10, true);
    const cdOffset = this.view.getUint32(eocdOffset + 16, true);

    let offset = cdOffset;
    for (let i = 0; i < totalEntries; i++) {
      if (this.view.getUint32(offset, true) !== 0x02014b50) {
        throw new Error(`Invalid Central Directory header at ${offset}`);
      }
      const compMethod = this.view.getUint16(offset + 10, true);
      const modTime = this.view.getUint16(offset + 12, true);
      const modDate = this.view.getUint16(offset + 14, true);
      const crc32 = this.view.getUint32(offset + 16, true);
      const compressedSize = this.view.getUint32(offset + 20, true);
      const uncompressedSize = this.view.getUint32(offset + 24, true);
      const filenameLen = this.view.getUint16(offset + 28, true);
      const extraLen = this.view.getUint16(offset + 30, true);
      const commentLen = this.view.getUint16(offset + 32, true);
      const localHeaderOffset = this.view.getUint32(offset + 42, true);

      const filenameBytes = this.buffer.subarray(offset + 46, offset + 46 + filenameLen);
      const filename = textDecoder.decode(filenameBytes);

      const localFilenameLen = this.view.getUint16(localHeaderOffset + 26, true);
      const localExtraLen = this.view.getUint16(localHeaderOffset + 28, true);
      const dataOffset = localHeaderOffset + 30 + localFilenameLen + localExtraLen;

      const entry: ZipEntryMeta = {
        filename,
        filenameBytes,
        compMethod,
        modTime,
        modDate,
        crc32,
        compressedSize,
        uncompressedSize,
        localHeaderOffset,
        dataOffset,
      };

      this.entries.set(filename, entry);

      const sheetMatch = filename.match(/^xl\/worksheets\/sheet([0-9]+)\.xml$/);
      if (sheetMatch) {
        const idx = parseInt(sheetMatch[1], 10);
        this.sheetIndexToEntry.set(idx, entry);
      }

      offset += 46 + filenameLen + extraLen + commentLen;
    }
  }

  public getRawCompressedData(entry: ZipEntryMeta): Uint8Array {
    return this.buffer.subarray(entry.dataOffset, entry.dataOffset + entry.compressedSize);
  }
}

interface OutEntryMeta {
  filename: string;
  filenameBytes: Uint8Array;
  crc: number;
  compMethod: number;
  compressedSize: number;
  uncompressedSize: number;
  localHeaderOffset: number;
}

class CopyThroughZipWriter {
  private entries: OutEntryMeta[] = [];
  private outputChunks: Uint8Array[] = [];
  private currentOffset = 0;

  public copyRawEntry(entry: ZipEntryMeta, rawCompressedBytes: Uint8Array): void {
    const localHeaderOffset = this.currentOffset;

    const header = new Uint8Array(30 + entry.filenameBytes.length);
    const view = new DataView(header.buffer, header.byteOffset, header.byteLength);
    view.setUint32(0, 0x04034b50, true);
    view.setUint16(4, 20, true);
    view.setUint16(6, 0x0800, true);
    view.setUint16(8, entry.compMethod, true);
    view.setUint16(10, entry.modTime, true);
    view.setUint16(12, entry.modDate, true);
    view.setUint32(14, entry.crc32, true);
    view.setUint32(18, entry.compressedSize, true);
    view.setUint32(22, entry.uncompressedSize, true);
    view.setUint16(26, entry.filenameBytes.length, true);
    view.setUint16(28, 0, true);
    header.set(entry.filenameBytes, 30);

    this.outputChunks.push(header);
    this.outputChunks.push(rawCompressedBytes);
    this.currentOffset += header.length + rawCompressedBytes.length;

    this.entries.push({
      filename: entry.filename,
      filenameBytes: entry.filenameBytes,
      crc: entry.crc32,
      compMethod: entry.compMethod,
      compressedSize: entry.compressedSize,
      uncompressedSize: entry.uncompressedSize,
      localHeaderOffset,
    });
  }

  public addNewEntry(filename: string, uncompressedData: string | Uint8Array, compressionLevel = 1): void {
    const buf = typeof uncompressedData === 'string' ? textEncoder.encode(uncompressedData) : uncompressedData;
    const filenameBytes = textEncoder.encode(filename);
    const crc = crc32Update(0, buf);
    const uncompressedSize = buf.length;

    let compressedBuf: Uint8Array;
    let compMethod = 8;
    if (compressionLevel === 0) {
      compressedBuf = buf;
      compMethod = 0;
    } else {
      compressedBuf = deflateSync(buf, { level: compressionLevel as any });
    }
    const compressedSize = compressedBuf.length;
    const localHeaderOffset = this.currentOffset;

    const header = new Uint8Array(30 + filenameBytes.length);
    const view = new DataView(header.buffer, header.byteOffset, header.byteLength);
    view.setUint32(0, 0x04034b50, true);
    view.setUint16(4, 20, true);
    view.setUint16(6, 0x0800, true);
    view.setUint16(8, compMethod, true);
    view.setUint16(10, 0, true);
    view.setUint16(12, 0, true);
    view.setUint32(14, crc, true);
    view.setUint32(18, compressedSize, true);
    view.setUint32(22, uncompressedSize, true);
    view.setUint16(26, filenameBytes.length, true);
    view.setUint16(28, 0, true);
    header.set(filenameBytes, 30);

    this.outputChunks.push(header);
    this.outputChunks.push(compressedBuf);
    this.currentOffset += header.length + compressedBuf.length;

    this.entries.push({
      filename,
      filenameBytes,
      crc,
      compMethod,
      compressedSize,
      uncompressedSize,
      localHeaderOffset,
    });
  }

  public finalize(): Uint8Array {
    const centralDirOffset = this.currentOffset;
    let centralDirSize = 0;

    for (const e of this.entries) {
      const cdHeader = new Uint8Array(46 + e.filenameBytes.length);
      const view = new DataView(cdHeader.buffer, cdHeader.byteOffset, cdHeader.byteLength);
      view.setUint32(0, 0x02014b50, true);
      view.setUint16(4, 20, true);
      view.setUint16(6, 20, true);
      view.setUint16(8, 0x0800, true);
      view.setUint16(10, e.compMethod, true);
      view.setUint16(12, 0, true);
      view.setUint16(14, 0, true);
      view.setUint32(16, e.crc, true);
      view.setUint32(20, e.compressedSize, true);
      view.setUint32(24, e.uncompressedSize, true);
      view.setUint16(28, e.filenameBytes.length, true);
      view.setUint16(30, 0, true);
      view.setUint16(32, 0, true);
      view.setUint16(34, 0, true);
      view.setUint16(36, 0, true);
      view.setUint32(38, 0, true);
      view.setUint32(42, e.localHeaderOffset, true);
      cdHeader.set(e.filenameBytes, 46);

      this.outputChunks.push(cdHeader);
      this.currentOffset += cdHeader.length;
      centralDirSize += cdHeader.length;
    }

    const eocd = new Uint8Array(22);
    const eocdView = new DataView(eocd.buffer, eocd.byteOffset, eocd.byteLength);
    eocdView.setUint32(0, 0x06054b50, true);
    eocdView.setUint16(4, 0, true);
    eocdView.setUint16(6, 0, true);
    eocdView.setUint16(8, this.entries.length, true);
    eocdView.setUint16(10, this.entries.length, true);
    eocdView.setUint32(12, centralDirSize, true);
    eocdView.setUint32(16, centralDirOffset, true);
    eocdView.setUint16(20, 0, true);

    this.outputChunks.push(eocd);
    this.currentOffset += eocd.length;

    const totalLength = this.outputChunks.reduce((sum, chunk) => sum + chunk.length, 0);
    const result = new Uint8Array(totalLength);
    let offset = 0;
    for (const chunk of this.outputChunks) {
      result.set(chunk, offset);
      offset += chunk.length;
    }

    this.outputChunks = [];
    this.entries = [];

    return result;
  }
}

// Session state inside Worker
let exportMode: 'COPY_THROUGH' | 'GREENFIELD' = 'GREENFIELD';
let originalZipReader: FastZipReader | null = null;
let dirtySheetIndexSet = new Set<number>();
const dirtySheetXmlMap = new Map<string, Uint8Array>();

let copyThroughWriter: CopyThroughZipWriter | null = null;
let sheetsMetadata: Array<{ id: string; index: number; name: string }> = [];
let availableSheetNames: string[] = [];
let currentSheetChunks: Uint8Array[] = [];
let currentSheetTail = '';
const sharedFormulaMap = new Map<string, string>();

function sanitizeSheetNames(formula: string): string {
  if (!formula) return formula;
  let result = formula;
  for (let i = 0; i < availableSheetNames.length; i++) {
    const name = availableSheetNames[i];
    if (/[\s&\-\.\/\\]/.test(name)) {
      const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      const regex = new RegExp(`(?<!')${escaped}(?!')!`, 'g');
      result = result.replace(regex, `'${name}'!`);
    }
  }
  return result;
}

self.onmessage = async (e: MessageEvent) => {
  const {
    type,
    mode,
    originalBuffer,
    dirtySheetIndices,
    format,
    availableSheetNames: names,
    sheetsMetadata: meta,
    sheetId,
    sheetIndex,
    sheetName,
    columnData,
    chunkIndex,
    totalChunks: _totalChunks,
    chunkRows,
    snapshot,
    activeSheetId,
  } = e.data;

  try {
    if (type === 'INIT_WORKBOOK') {
      console.log(`[WORKER] Received INIT_WORKBOOK (mode: ${mode || 'GREENFIELD'})`);
      exportMode = mode === 'COPY_THROUGH' && originalBuffer ? 'COPY_THROUGH' : 'GREENFIELD';
      availableSheetNames = names || [];
      sheetsMetadata = meta || [];
      currentSheetChunks = [];
      sharedFormulaMap.clear();
      dirtySheetXmlMap.clear();
      dirtySheetIndexSet = new Set(dirtySheetIndices || []);

      if (exportMode === 'COPY_THROUGH') {
        const u8 = originalBuffer instanceof Uint8Array ? originalBuffer : new Uint8Array(originalBuffer);
        originalZipReader = new FastZipReader(u8);
        copyThroughWriter = new CopyThroughZipWriter();
        console.log(`[WORKER] Copy-Through Mode initialized with ${originalZipReader.entries.size} original entries. Dirty sheets: ${dirtySheetIndexSet.size}`);
      } else {
        copyThroughWriter = new CopyThroughZipWriter();
        originalZipReader = null;

        // Write OpenXML container metadata files for Greenfield mode
        let contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>`;
        sheetsMetadata.forEach((s) => {
          contentTypesXml += `\n  <Override PartName="/xl/worksheets/sheet${s.index}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
        });
        contentTypesXml += '\n</Types>';
        copyThroughWriter.addNewEntry('[Content_Types].xml', contentTypesXml, 1);

        const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;
        copyThroughWriter.addNewEntry('_rels/.rels', relsXml, 1);

        let workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <calcPr fullCalcOnLoad="1" forceFullCalc="1"/>
  <sheets>`;
        sheetsMetadata.forEach((s) => {
          workbookXml += `\n    <sheet name="${escapeXml(s.name)}" sheetId="${s.index}" r:id="rId${s.index}"/>`;
        });
        workbookXml += '\n  </sheets>\n</workbook>';
        copyThroughWriter.addNewEntry('xl/workbook.xml', workbookXml, 1);

        let workbookRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`;
        sheetsMetadata.forEach((s) => {
          workbookRelsXml += `\n  <Relationship Id="rId${s.index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${s.index}.xml"/>`;
        });
        workbookRelsXml += `\n  <Relationship Id="rIdStyle" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;
        copyThroughWriter.addNewEntry('xl/_rels/workbook.xml.rels', workbookRelsXml, 1);

        const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
</styleSheet>`;
        copyThroughWriter.addNewEntry('xl/styles.xml', stylesXml, 1);
      }

      self.postMessage({ type: 'ACK' });
      return;
    }

    if (type === 'INIT_SHEET') {
      console.log(`[WORKER] Received INIT_SHEET: ${sheetName} (sheet index: ${sheetIndex})`);
      currentSheetChunks = [];
      sharedFormulaMap.clear();
      currentSheetTail = '';

      let colsXml = '';
      if (columnData && typeof columnData === 'object') {
        const colIndices = Object.keys(columnData).map(Number).sort((a, b) => a - b);
        if (colIndices.length > 0) {
          colsXml = '  <cols>\n';
          for (let i = 0; i < colIndices.length; i++) {
            const cIdx = colIndices[i];
            const col = columnData[cIdx];
            if (col) {
              const colNum = cIdx + 1;
              const w = col.w !== undefined ? col.w : 10;
              const hd = col.hd === 1 ? ' hidden="1"' : '';
              const customW = col.w !== undefined ? ' customWidth="1"' : '';
              colsXml += `    <col min="${colNum}" max="${colNum}" width="${w}"${hd}${customW}/>\n`;
            }
          }
          colsXml += '  </cols>\n';
        }
      }

      let headerXml = '';
      if (exportMode === 'COPY_THROUGH' && originalZipReader) {
        const origEntry = originalZipReader.entries.get(`xl/worksheets/sheet${sheetIndex}.xml`);
        if (origEntry) {
          const raw = originalZipReader.getRawCompressedData(origEntry);
          const uncompressed = inflateSync(raw);
          const origXml = textDecoder.decode(uncompressed);
          const sheetDataStart = origXml.indexOf('<sheetData');
          const sheetDataEnd = origXml.indexOf('</sheetData>');
          if (sheetDataStart !== -1 && sheetDataEnd !== -1) {
            let origHead = origXml.substring(0, sheetDataStart);
            currentSheetTail = origXml.substring(sheetDataEnd + '</sheetData>'.length);

            if (colsXml) {
              if (origHead.includes('<cols>')) {
                origHead = origHead.replace(/<cols>[\s\S]*?<\/cols>/, colsXml.trim());
              } else if (origHead.includes('</sheetViews>')) {
                const svEnd = origHead.indexOf('</sheetViews>') + '</sheetViews>'.length;
                origHead = origHead.substring(0, svEnd) + '\n' + colsXml + origHead.substring(svEnd);
              } else {
                origHead = origHead + colsXml;
              }
            }
            headerXml = origHead + '<sheetData>\n';
          }
        }
      }

      if (!headerXml) {
        headerXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n${colsXml}  <sheetData>\n`;
      }

      currentSheetChunks.push(textEncoder.encode(headerXml));
      self.postMessage({ type: 'ACK' });
      return;
    }

    if (type === 'APPEND_CHUNK') {
      const cellDataRaw = chunkRows || {};
      const rowKeys = Object.keys(cellDataRaw).map(Number).sort((a, b) => a - b);

      // Register shared master formulas in this chunk
      for (let rIdx = 0; rIdx < rowKeys.length; rIdx++) {
        const rowIndex = rowKeys[rIdx];
        const cols = cellDataRaw[rowIndex];
        if (cols) {
          const colKeys = Object.keys(cols).map(Number);
          for (let cIdx = 0; cIdx < colKeys.length; cIdx++) {
            const cellPay = cols[colKeys[cIdx]];
            if (cellPay && cellPay.f && cellPay.si !== undefined && cellPay.si !== null) {
              sharedFormulaMap.set(String(cellPay.si), cellPay.f);
            }
          }
        }
      }

      let chunkXml = '';

      // Convert cells directly to OpenXML tags
      for (let rIdx = 0; rIdx < rowKeys.length; rIdx++) {
        const rowIndex = rowKeys[rIdx];
        const cols = cellDataRaw[rowIndex];
        if (!cols) continue;

        const rowNumber = rowIndex + 1;
        const colKeys = Object.keys(cols).map(Number).sort((a, b) => a - b);
        if (colKeys.length === 0) continue;

        chunkXml += `    <row r="${rowNumber}">`;

        for (let cIdx = 0; cIdx < colKeys.length; cIdx++) {
          const colIndex = colKeys[cIdx];
          const cellPay = cols[colIndex];
          if (!cellPay) continue;

          const cellRef = `${colToLetter(colIndex)}${rowNumber}`;

          let formulaText = cellPay.f;
          if (!formulaText && cellPay.si !== undefined && cellPay.si !== null) {
            formulaText = sharedFormulaMap.get(String(cellPay.si));
          }

          if (formulaText) {
            let parsedFormula = formulaText.startsWith('=') ? formulaText.substring(1) : formulaText;
            parsedFormula = sanitizeSheetNames(parsedFormula);
            if (parsedFormula.startsWith('{') && parsedFormula.endsWith('}')) {
              parsedFormula = parsedFormula.substring(1, parsedFormula.length - 1);
            }

            const formulaRes =
              cellPay.v !== undefined && cellPay.v !== null
                ? cellPay.v
                : cellPay.p?.body?.dataStream
                ? cellPay.p.body.dataStream.replace(/[\r\n]+$/, '')
                : undefined;

            const isArrayFormula = cellPay.fType === 'array' || cellPay.shareType === 'array';
            const refAttr = isArrayFormula ? (cellPay.ref ? ` t="array" ref="${escapeXml(cellPay.ref)}"` : ` t="array" ref="${cellRef}"`) : '';

            if (formulaRes !== undefined) {
              if (cellPay.t === 2 || typeof formulaRes === 'number') {
                chunkXml += `<c r="${cellRef}"><f${refAttr}>${escapeFormulaXml(parsedFormula)}</f><v>${formulaRes}</v></c>`;
              } else if (cellPay.t === 3 || typeof formulaRes === 'boolean') {
                chunkXml += `<c r="${cellRef}" t="b"><f${refAttr}>${escapeFormulaXml(parsedFormula)}</f><v>${formulaRes ? 1 : 0}</v></c>`;
              } else if (typeof formulaRes === 'string' && formulaRes.startsWith('#')) {
                chunkXml += `<c r="${cellRef}" t="e"><f${refAttr}>${escapeFormulaXml(parsedFormula)}</f><v>${escapeXml(formulaRes)}</v></c>`;
              } else {
                chunkXml += `<c r="${cellRef}" t="str"><f${refAttr}>${escapeFormulaXml(parsedFormula)}</f><v>${escapeXml(String(formulaRes))}</v></c>`;
              }
            } else {
              chunkXml += `<c r="${cellRef}"><f${refAttr}>${escapeFormulaXml(parsedFormula)}</f></c>`;
            }
          } else if (cellPay.v !== undefined && cellPay.v !== null) {
            if (cellPay.t === 2 || typeof cellPay.v === 'number') {
              chunkXml += `<c r="${cellRef}"><v>${cellPay.v}</v></c>`;
            } else if (cellPay.t === 3 || typeof cellPay.v === 'boolean') {
              chunkXml += `<c r="${cellRef}" t="b"><v>${cellPay.v ? 1 : 0}</v></c>`;
            } else {
              chunkXml += `<c r="${cellRef}" t="inlineStr"><is><t>${escapeXml(cellPay.v)}</t></is></c>`;
            }
          } else if (cellPay.p && cellPay.p.body && typeof cellPay.p.body.dataStream === 'string') {
            const textVal = cellPay.p.body.dataStream.replace(/[\r\n]+$/, '');
            if (cellPay.t === 2) {
              chunkXml += `<c r="${cellRef}"><v>${escapeXml(textVal)}</v></c>`;
            } else if (cellPay.t === 3) {
              chunkXml += `<c r="${cellRef}" t="b"><v>${textVal === 'true' || textVal === '1' ? 1 : 0}</v></c>`;
            } else {
              chunkXml += `<c r="${cellRef}" t="inlineStr"><is><t>${escapeXml(textVal)}</t></is></c>`;
            }
          }
        }

        chunkXml += '</row>\n';
      }

      if (chunkXml.length > 0) {
        currentSheetChunks.push(textEncoder.encode(chunkXml));
      }

      self.postMessage({ type: 'CHUNK_ACK', sheetId, chunkIndex });
      return;
    }

    if (type === 'END_SHEET') {
      console.log(`[WORKER] sheet XML finalized for sheet index ${sheetIndex}`);
      if (currentSheetTail) {
        currentSheetChunks.push(textEncoder.encode('  </sheetData>' + currentSheetTail));
        currentSheetTail = '';
      } else {
        currentSheetChunks.push(textEncoder.encode('  </sheetData>\n</worksheet>'));
      }

      const totalLength = currentSheetChunks.reduce((sum, b) => sum + b.length, 0);
      const fullSheetBuf = new Uint8Array(totalLength);
      let offset = 0;
      for (let i = 0; i < currentSheetChunks.length; i++) {
        const b = currentSheetChunks[i];
        fullSheetBuf.set(b, offset);
        offset += b.length;
      }

      currentSheetChunks = [];
      sharedFormulaMap.clear();

      if (exportMode === 'COPY_THROUGH') {
        dirtySheetXmlMap.set(`xl/worksheets/sheet${sheetIndex}.xml`, fullSheetBuf);
      } else if (copyThroughWriter) {
        copyThroughWriter.addNewEntry(`xl/worksheets/sheet${sheetIndex}.xml`, fullSheetBuf, 1);
      }

      self.postMessage({ type: 'SHEET_ACK', sheetId });
      return;
    }

    if (type === 'FINALIZE_WORKBOOK') {
      console.log(`[WORKER] XLSX packaging start (Mode: ${exportMode})`);
      if (!copyThroughWriter) throw new Error('ZIP generator is uninitialized.');

      if (exportMode === 'COPY_THROUGH' && originalZipReader) {
        const hasDirtySheets = dirtySheetXmlMap.size > 0;

        // Stream all untouched entries directly in raw compressed form!
        for (const [filename, entry] of originalZipReader.entries) {
          if (hasDirtySheets && filename === 'xl/calcChain.xml') {
            // Discard stale calcChain when workbook has modified sheets!
            console.log(`[WORKER] Discarding stale ${filename} because workbook has modified sheets.`);
            continue;
          }

          if (dirtySheetXmlMap.has(filename)) {
            // Replace with newly modified sheet XML buffer
            const modifiedXmlBuf = dirtySheetXmlMap.get(filename)!;
            copyThroughWriter.addNewEntry(filename, modifiedXmlBuf, 1);
            console.log(`[WORKER] Replaced modified sheet in stream: ${filename}`);
          } else if (hasDirtySheets && filename === 'xl/_rels/workbook.xml.rels') {
            // Clean calcChain relationship from workbook.xml.rels
            const rawCompressedData = originalZipReader.getRawCompressedData(entry);
            const uncompressed = inflateSync(rawCompressedData);
            const relsXml = textDecoder.decode(uncompressed);
            const cleanedRelsXml = relsXml.replace(/<Relationship [^>]*Target="calcChain\.xml"[^>]*\/>/g, '');
            copyThroughWriter.addNewEntry(filename, cleanedRelsXml, 1);
            console.log(`[WORKER] Cleaned calcChain relationship from ${filename}`);
          } else if (hasDirtySheets && filename === '[Content_Types].xml') {
            // Clean calcChain override from [Content_Types].xml
            const rawCompressedData = originalZipReader.getRawCompressedData(entry);
            const uncompressed = inflateSync(rawCompressedData);
            const ctXml = textDecoder.decode(uncompressed);
            const cleanedCtXml = ctXml.replace(/<Override [^>]*PartName="\/xl\/calcChain\.xml"[^>]*\/>/g, '');
            copyThroughWriter.addNewEntry(filename, cleanedCtXml, 1);
            console.log(`[WORKER] Cleaned calcChain override from ${filename}`);
          } else if (hasDirtySheets && filename === 'xl/workbook.xml') {
            // When calcChain is discarded, ensure workbook.xml has fullCalcOnLoad="1" forceFullCalc="1"
            const rawCompressedData = originalZipReader.getRawCompressedData(entry);
            const uncompressed = inflateSync(rawCompressedData);
            let wbXml = textDecoder.decode(uncompressed);
            if (wbXml.includes('<calcPr')) {
              wbXml = wbXml.replace(/<calcPr\b([^>]*?)\/?>/g, (_match, attrs) => {
                const cleanAttrs = attrs.replace(/\s*(fullCalcOnLoad|forceFullCalc)="[^"]*"/g, '');
                return `<calcPr${cleanAttrs} fullCalcOnLoad="1" forceFullCalc="1"/>`;
              });
            } else {
              wbXml = wbXml.replace('</workbook>', '  <calcPr fullCalcOnLoad="1" forceFullCalc="1"/>\n</workbook>');
            }
            copyThroughWriter.addNewEntry(filename, wbXml, 1);
            console.log(`[WORKER] Ensured fullCalcOnLoad and forceFullCalc in ${filename}`);
          } else {
            // Direct zero-copy from original binary slice!
            const rawCompressedData = originalZipReader.getRawCompressedData(entry);
            copyThroughWriter.copyRawEntry(entry, rawCompressedData);
          }
        }
      }

      const uint8Array = copyThroughWriter.finalize();
      console.log(`[WORKER] XLSX packaging complete. Output size: ${(uint8Array.byteLength / 1024 / 1024).toFixed(2)} MB`);

      copyThroughWriter = null;
      originalZipReader = null;
      dirtySheetXmlMap.clear();
      dirtySheetIndexSet.clear();

      (self as any).postMessage({ type: 'COMPLETE', buffer: uint8Array.buffer }, [uint8Array.buffer]);
      return;
    }

    // CSV export handler
    if (format === 'csv' && snapshot) {
      self.postMessage({ type: 'PROGRESS', percent: 10, message: 'Menyiapkan data CSV...' });
      const targetSheetId = activeSheetId || (snapshot.sheetOrder ? snapshot.sheetOrder[0] : Object.keys(snapshot.sheets)[0]);
      const sheetData = snapshot.sheets[targetSheetId];

      if (!sheetData || !sheetData.cellData) {
        throw new Error('Sheet data is empty or invalid.');
      }

      const rows: string[][] = [];
      const cellDataRaw = sheetData.cellData;
      const rowKeys = Object.keys(cellDataRaw).map(Number).sort((a, b) => a - b);

      for (let i = 0; i < rowKeys.length; i++) {
        const rIdx = rowKeys[i];
        const cols = cellDataRaw[rIdx];
        if (!cols) continue;

        const maxCol = Math.max(...Object.keys(cols).map(Number));
        const rowArr: string[] = new Array(maxCol + 1).fill('');

        Object.keys(cols).forEach((cStr) => {
          const cIdx = parseInt(cStr, 10);
          const cell = cols[cIdx];
          if (cell) {
            let strVal = '';
            if (cell.v !== undefined && cell.v !== null) {
              strVal = String(cell.v);
            } else if (cell.p && cell.p.body && typeof cell.p.body.dataStream === 'string') {
              strVal = cell.p.body.dataStream.replace(/[\r\n]+$/, '');
            }
            rowArr[cIdx] = strVal;
          }
        });

        rows.push(rowArr);

        if (i % 2000 === 0 || i === rowKeys.length - 1) {
          const percent = Math.min(85, Math.round(10 + (i / rowKeys.length) * 75));
          self.postMessage({ type: 'PROGRESS', percent, message: `Mengonversi baris CSV (${i + 1}/${rowKeys.length})...` });
          await new Promise((resolve) => setTimeout(resolve, 0));
        }
      }

      self.postMessage({ type: 'PROGRESS', percent: 90, message: 'Meng-encode teks CSV...' });
      const csvText = Papa.unparse(rows);
      const encoder = new TextEncoder();
      const uint8 = encoder.encode(csvText);
      const buffer = uint8.buffer;

      (self as any).postMessage({ type: 'PROGRESS', percent: 100, message: 'File CSV selesai dibuat!' });
      (self as any).postMessage({ type: 'COMPLETE', buffer }, [buffer]);
      return;
    }
  } catch (err: any) {
    console.error('[WORKER] Export error:', err);
    (self as any).postMessage({
      type: 'ERROR',
      error: err?.message || 'Gagal mengekspor file di Background Worker',
    });
  }
};
