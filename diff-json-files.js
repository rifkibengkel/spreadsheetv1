const fs = require('fs');

const manual = JSON.parse(fs.readFileSync('manualWbData.json', 'utf8'));
const barakat = JSON.parse(fs.readFileSync('barakatWbData.json', 'utf8'));

console.log('--- Summary Sheet Differences ---');
console.log('Manual Summary keys:', Object.keys(manual.sheets.Summary));
console.log('Barakat Summary keys:', Object.keys(barakat.sheets.Summary));

console.log('Manual Summary row 0 col 0:', JSON.stringify(manual.sheets.Summary.cellData[0][0]));
console.log('Barakat Summary row 0 col 0:', JSON.stringify(barakat.sheets.Summary.cellData[0][0]));

console.log('\n--- Summary Hadiah Sheet Differences ---');
console.log('Manual Summary Hadiah row 0 col 0:', JSON.stringify(manual.sheets['Summary Hadiah'].cellData[0][0]));
console.log('Barakat Summary Hadiah row 0 col 0:', JSON.stringify(barakat.sheets['Summary Hadiah'].cellData[0][0]));

console.log('\n--- Comparing Cell Count & Cell Data Structure ---');
console.log('Manual Summary row count of cellData:', Object.keys(manual.sheets.Summary.cellData).length);
console.log('Barakat Summary row count of cellData:', Object.keys(barakat.sheets.Summary.cellData).length);

console.log('Manual Summary Hadiah row count of cellData:', Object.keys(manual.sheets['Summary Hadiah'].cellData).length);
console.log('Barakat Summary Hadiah row count of cellData:', Object.keys(barakat.sheets['Summary Hadiah'].cellData).length);

// Dump all cells in Barakat Summary sheet to see what else is in row 0 or cellData
fs.writeFileSync('barakat-summary-celldata.json', JSON.stringify(barakat.sheets.Summary.cellData, null, 2));
fs.writeFileSync('barakat-hadiah-celldata.json', JSON.stringify(barakat.sheets['Summary Hadiah'].cellData, null, 2));

console.log('Dumped barakat-summary-celldata.json and barakat-hadiah-celldata.json');
