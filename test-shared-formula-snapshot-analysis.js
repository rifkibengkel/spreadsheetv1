const fs = require('fs');
const unzipper = require('unzipper');

(async () => {
  console.log('=== Forensic Analysis of Shared Formulas in updatehexos.xlsx XML ===');

  const filePath = 'C:\\Users\\BKKI-1\\Desktop\\updatehexos.xlsx';
  const directory = await unzipper.Open.file(filePath);

  const sheet8File = directory.files.find(f => f.path.endsWith('sheet8.xml'));
  if (!sheet8File) {
    console.log('sheet8.xml not found.');
    return;
  }

  const content = await sheet8File.buffer();
  const xmlStr = content.toString('utf8');

  // Look for shared formula master <f t="shared" ...> and slave <f t="shared" si="..."/>
  const masterMatches = xmlStr.match(/<f[^>]*t="shared"[^>]*ref="[^"]*"[^>]*>[^<]+<\/f>/g) || [];
  const slaveMatches = xmlStr.match(/<f[^>]*t="shared"[^>]*\/>/g) || [];

  console.log(`Original sheet8.xml:`);
  console.log(`- Master Shared Formulas: ${masterMatches.length}`);
  console.log(`- Slave Shared Formulas: ${slaveMatches.length}`);

  if (masterMatches.length > 0) {
    console.log('Master Sample:', masterMatches[0]);
  }
  if (slaveMatches.length > 0) {
    console.log('Slave Sample:', slaveMatches[0]);
  }
})();
