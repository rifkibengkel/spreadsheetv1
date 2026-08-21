const fs = require('fs');
const unzipper = require('unzipper');

(async () => {
  console.log('=== Inspecting Row 4 of xl/worksheets/sheet1.xml in report barakat.xlsx ===');

  const filePath = 'C:\\Users\\BKKI-1\\Documents\\report barakat.xlsx';
  const directory = await unzipper.Open.file(filePath);

  const sheet1File = directory.files.find(f => f.path === 'xl/worksheets/sheet1.xml');
  const xmlStr = (await sheet1File.buffer()).toString('utf8');

  // Read Shared Strings if present
  let sharedStrings = [];
  const sstFile = directory.files.find(f => f.path === 'xl/sharedStrings.xml');
  if (sstFile) {
    const sstXml = (await sstFile.buffer()).toString('utf8');
    sharedStrings = [...sstXml.matchAll(/<t[^>]*>([^<]*)<\/t>/g)].map(m => m[1]);
    console.log(`Shared Strings loaded: ${sharedStrings.length} strings.`);
  }

  const row4Regex = /<c r="([A-Z]+4)"([^>]*)>(?:<f([^>]*)>([^<]*)<\/f>)?(?:<v>([^<]*)<\/v>)?/g;
  let m;

  console.log('Row 4 Cells in Source Sheet 1:');
  while ((m = row4Regex.exec(xmlStr)) !== null) {
    const cellRef = m[1];
    const attrs = m[2];
    const formula = m[4];
    const rawVal = m[5];

    let actualVal = rawVal;
    if (attrs.includes('t="s"') && rawVal !== undefined) {
      const idx = Number(rawVal);
      actualVal = sharedStrings[idx] !== undefined ? `${sharedStrings[idx]} (SharedString #${idx})` : `SharedString #${idx}`;
    }

    console.log(`Cell ${cellRef}: attrs="${attrs}", formula="${formula || ''}", rawVal="${rawVal || ''}", actualVal="${actualVal || ''}"`);
  }
})();
