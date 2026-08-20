import ExcelJS from 'exceljs';

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

self.onmessage = async (e: MessageEvent) => {
  const { fileArrayBuffer, fileName } = e.data;

  try {
    self.postMessage({ type: 'PROGRESS', percent: 5, message: 'Membaca file di Background Worker (5%)...' });

    let workbook: ExcelJS.Workbook | null = new ExcelJS.Workbook();
    self.postMessage({ type: 'PROGRESS', percent: 15, message: 'Mengurai struktur XML Excel di Worker (15%)...' });

    await workbook.xlsx.load(fileArrayBuffer);

    self.postMessage({ type: 'PROGRESS', percent: 25, message: 'Menyiapkan metadata dan mengurai semua sheet...' });

    const sheetOrder: string[] = [];
    const sheetsData: any = {};
    const totalSheets = workbook.worksheets.length;

    workbook.worksheets.forEach((ws, sIdx) => {
      const sheetId = ws.name; // Preserve exact sheet name as sheet ID for cross-sheet formula lookup!
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
            // Clean malformed WPS sheet references if present
            cleanedF = cleanedF.replace(/'\[[^'\]]+\]#REF'/g, '');
            cellInfo.f = cleanedF;
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

    // Release ExcelJS workbook memory immediately after extracting raw sheet data
    workbook = null;

    self.postMessage({ type: 'PROGRESS', percent: 100, message: 'Impor selesai! (100%)' });

    // Pre-resolve scalar cross-sheet & same-sheet formulas missing 'v'
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
              // 1. Cross-sheet scalar reference: =Sheet!A1 or ='Sheet Name'!A1
              let match = cell.f.match(/^=('([^']+)'|([A-Za-z0-9_\s]+))!([A-Za-z]+)([0-9]+)$/i);
              let targetSheetName = match ? (match[2] || match[3]) : null;
              let colStr = match ? match[4].toUpperCase() : null;
              let rowNum = match ? parseInt(match[5], 10) : null;

              // 2. Same-sheet scalar reference: =A1
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
                    console.log(`[Formula Worker] Pre-resolved ${sName}!R${rIdx}C${cIdx} (${cell.f}) -> ${targetCell.v}`);
                  }
                }
              }
            }
          });
        });
      });
      if (!resolvedAny) break;
    }

    const workbookData = {
      id: `workbook-imported-${Date.now()}`,
      name: fileName,
      sheetOrder,
      appVersion: '3.0.0-alpha',
      sheets: sheetsData,
    };

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

