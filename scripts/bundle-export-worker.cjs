const esbuild = require('esbuild');
const fs = require('fs');
const path = require('path');

const projectRoot = path.resolve(__dirname, '..');
const workerEntry = path.join(projectRoot, 'src/lib/workers/export.worker.ts');
const targetBundleTs = path.join(projectRoot, 'src/lib/workers/exportWorkerBundle.ts');

console.log('[BUNDLE] Bundling export.worker.ts with esbuild...');
const result = esbuild.buildSync({
  entryPoints: [workerEntry],
  bundle: true,
  format: 'iife',
  minify: true,
  write: false,
});

if (!result.outputFiles || result.outputFiles.length === 0) {
  throw new Error('esbuild produced no output files.');
}

const bundledCode = result.outputFiles[0].text;
const escapedJson = JSON.stringify(bundledCode);
const content = `// Auto-generated bundled worker to bypass Chrome HTTP/1.1 socket exhaustion (6 connections per host limit)
// Generated: ${new Date().toISOString()}
export const EXPORT_WORKER_CODE: string = ${escapedJson};
`;

fs.writeFileSync(targetBundleTs, content, 'utf8');
console.log(`[BUNDLE] Successfully created exportWorkerBundle.ts (${(bundledCode.length / 1024).toFixed(2)} KB)`);
