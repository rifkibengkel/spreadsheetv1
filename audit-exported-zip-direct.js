const fs = require('fs');
const unzipper = require('unzipper');

(async () => {
  console.log('=== Forensic Audit of Exported Zip XML files ===');

  const zipFiles = ['wps-export-fixed-test.xlsx', 'test-exported.xlsx', 'final-verified-wps-export.xlsx'];

  for (const zipPath of zipFiles) {
    if (!fs.existsSync(zipPath)) continue;

    console.log(`\n========================================`);
    console.log(`FILE: ${zipPath}`);
    console.log(`========================================`);

    const directory = await unzipper.Open.file(zipPath);

    for (const file of directory.files) {
      if (file.path.startsWith('xl/worksheets/sheet') && file.path.endsWith('.xml')) {
        const content = await file.buffer();
        const xmlStr = content.toString('utf8');

        const formulaCount = (xmlStr.match(/<f[^>]*>/g) || []).length;
        const arrayCount = (xmlStr.match(/<f[^>]*t="array"[^>]*>/g) || []).length;
        const sharedCount = (xmlStr.match(/<f[^>]*t="shared"[^>]*>/g) || []).length;
        const cachedCount = (xmlStr.match(/<v>[^<]*<\/v>/g) || []).length;

        console.log(`Sheet: ${file.path}`);
        console.log(`  - Total <f> formula tags: ${formulaCount}`);
        console.log(`  - Array <f t="array"> tags: ${arrayCount}`);
        console.log(`  - Shared <f t="shared"> tags: ${sharedCount}`);
        console.log(`  - Cached <v> tags: ${cachedCount}`);
      }
    }
  }
})();
