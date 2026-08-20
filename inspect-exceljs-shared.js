const fs = require('fs');
const ExcelJS = require('exceljs');

async function inspectSharedFormulas() {
  const filePath = 'C:\\Users\\BKKI-1\\Desktop\\report barakat.xlsx';
  console.log('Loading report barakat.xlsx...');

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const ws = workbook.getWorksheet('Summary');
  if (!ws) {
    console.error('Worksheet Summary not found!');
    return;
  }

  const out = { summaryCells: [], summaryHadiahCells: [] };

  for (let r = 4; r <= 15; r++) {
    const cell = ws.getCell(`I${r}`);
    out.summaryCells.push({
      address: `I${r}`,
      type: cell.type,
      formula: cell.formula,
      result: cell.result,
      value: cell.value,
      model: cell.model,
      sharedFormula: cell.model ? cell.model.sharedFormula : undefined,
      master: cell.model ? cell.model.master : undefined
    });
  }

  const wsHadiah = workbook.getWorksheet('Summary Hadiah');
  if (wsHadiah) {
    for (let r = 4; r <= 10; r++) {
      const cell = wsHadiah.getCell(`C${r}`);
      out.summaryHadiahCells.push({
        address: `C${r}`,
        type: cell.type,
        formula: cell.formula,
        result: cell.result,
        value: cell.value,
        model: cell.model,
        sharedFormula: cell.model ? cell.model.sharedFormula : undefined,
        master: cell.model ? cell.model.master : undefined
      });
    }
  }

  fs.writeFileSync('shared-inspection-out.json', JSON.stringify(out, null, 2));
  console.log('Saved shared-inspection-out.json');
}

inspectSharedFormulas().catch(console.error);
