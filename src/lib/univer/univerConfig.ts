import type { IWorkbookData } from '@univerjs/core';

export const initialWorkbookData: Partial<IWorkbookData> = {
  id: 'workbook-initial',
  name: 'Workbook',
  sheetOrder: ['sheet-1'],
  appVersion: '3.0.0-alpha',
  sheets: {
    'sheet-1': {
      id: 'sheet-1',
      name: 'Sheet1',
      type: 2,
      status: 1,
      rowCount: 100,
      columnCount: 30,
      zoomRatio: 1,
      columnData: {},
      rowData: {},
      cellData: {},
    } as any,
  },
};
