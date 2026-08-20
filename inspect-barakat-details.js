const fs = require('fs');
const ExcelJS = require('exceljs');

async function inspectBarakatDetails() {
  const filePath = 'C:\\Users\\BKKI-1\\Desktop\\report barakat.xlsx';
  const buf = fs.readFileSync(filePath);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buf);

  const sheets = workbook.worksheets.map((ws, idx) => ({
    index: idx,
    id: ws.id,
    name: ws.name
  }));

  const sampleFormulas = [];
  workbook.worksheets.forEach((ws) => {
    let count = 0;
    ws.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        let rawFormula = null;
        if (cell.type === ExcelJS.ValueType.Formula) {
          rawFormula = cell.formula || (cell.model && cell.model.formula) || (typeof cell.value === 'object' && cell.value !== null ? cell.value.formula : undefined);
        } else if (typeof cell.value === 'object' && cell.value !== null && cell.value.formula) {
          rawFormula = cell.value.formula;
        }

        if (rawFormula && count < 10) {
          count++;
          sampleFormulas.push({
            sheet: ws.name,
            cell: cell.address,
            formula: rawFormula,
            result: cell.result
          });
        }
      });
    });
  });

  const report = {
    sheets,
    sampleFormulas
  };

  fs.writeFileSync('barakat-inspection.json', JSON.stringify(report, null, 2));
  console.log('Saved barakat-inspection.json');
}

inspectBarakatDetails().catch(console.error);
