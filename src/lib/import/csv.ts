import Papa from 'papaparse';

export function importCSVBatched(
  file: File,
  univerAPI: any,
  onProgress: (percent: number) => void,
  onComplete: () => void,
  onError: (error: any) => void
) {
  const api = univerAPI || (typeof window !== 'undefined' ? (window as any).univerAPI : null);
  if (!api) {
    onError(new Error("Univer API engine is not ready."));
    return;
  }

  const activeWorkbook = api.getActiveWorkbook ? api.getActiveWorkbook() : null;
  if (!activeWorkbook) {
    onError(new Error("No active workbook found in Univer."));
    return;
  }

  const sheetName = file.name.substring(0, 31);
  const sheetId = `csv-sheet-${Date.now()}`;
  
  // Aggregate data in memory to avoid locking Univer Command History (Undo/Redo) with huge inserts
  const cellData: any = {};
  let rowIndex = 0;
  let maxCol = 10; // Default min width

  Papa.parse(file, {
    worker: true,
    chunk: (results) => {
      const rows = results.data as any[][];
      if (rows.length === 0) return;

      // Construct native Univer cellData directly
      for (let i = 0; i < rows.length; i++) {
          const rowArr = rows[i];
          const rTarget = rowIndex + i;
          cellData[rTarget] = {};
          
          if (rowArr.length > maxCol) {
              maxCol = rowArr.length;
          }

          for (let j = 0; j < rowArr.length; j++) {
            const rawVal = rowArr[j];
            if (rawVal !== undefined && rawVal !== null && rawVal !== '') {
              const numVal = Number(rawVal);
              if (!isNaN(numVal) && String(rawVal).trim() !== '') {
                cellData[rTarget][j] = { v: numVal, t: 2 }; // CellValueType.NUMBER
              } else {
                cellData[rTarget][j] = { v: String(rawVal), t: 1 }; // CellValueType.STRING
              }
            }
          }
      }

      rowIndex += rows.length;

      // Report lazyload progress
      const progressPercent = Math.round(((results.meta.cursor || 0) / file.size) * 100);
      onProgress(progressPercent);
    },
    complete: () => {
      onProgress(100);
      
      try {
          const snapshot = activeWorkbook.save();
          
          snapshot.sheets[sheetId] = {
              id: sheetId,
              name: sheetName,
              type: 2, // SheetType.GRID - CRITICAL FOR FORMULA ENGINE REGISTRATION
              rowCount: rowIndex + 100, // Allocate buffer
              columnCount: maxCol + 5,
              cellData: cellData
          };
          
          if (!snapshot.sheetOrder) snapshot.sheetOrder = [];
          snapshot.sheetOrder.push(sheetId);
          
          // Reinitialize workbook entirely with the new dataset
          const activeId = activeWorkbook.getId();
          if (api.disposeUnit) {
              api.disposeUnit(activeId);
          }
          api.createUniverSheet(snapshot);
          
          onComplete();
      } catch (err) {
          onError(err);
      }
    },
    error: (error) => {
      onError(error);
    }
  });
}
