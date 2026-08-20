const ExcelJS = require('exceljs');
const path = require('path');

async function createTestExcel() {
  const workbook = new ExcelJS.Workbook();
  
  // Sheet 1
  const sheet1 = workbook.addWorksheet('Sheet1');
  sheet1.getCell('A1').value = 100;
  sheet1.getCell('A2').value = 200;
  sheet1.getCell('A3').value = { formula: 'A1+A2', result: 300 };

  // Sheet 2
  const sheet2 = workbook.addWorksheet('Sheet2');
  sheet2.getCell('A1').value = { formula: 'Sheet1!A1', result: 100 };
  sheet2.getCell('A2').value = { formula: 'Sheet1!A2', result: 200 };
  sheet2.getCell('A3').value = { formula: 'Sheet1!A1+Sheet1!A2', result: 300 };

  const filePath = path.join(__dirname, 'test-compare.xlsx');
  await workbook.xlsx.writeFile(filePath);
  console.log('Created test-compare.xlsx successfully at:', filePath);
}

createTestExcel().catch(console.error);
