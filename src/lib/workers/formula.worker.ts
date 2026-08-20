import { createUniver, LocaleType } from '@univerjs/presets';
import { UniverSheetsCoreWorkerPreset } from '@univerjs/preset-sheets-core/worker';

console.log('[Formula Worker] 🚀 Script starting...');

try {
  createUniver({
    locale: LocaleType.EN_US,
    presets: [
      UniverSheetsCoreWorkerPreset(),
    ],
  });
  console.log('[Formula Worker] ✅ createUniver initialized successfully!');
} catch (e: any) {
  console.error('[Formula Worker] ❌ Failed to initialize:', e.message, e.stack);
}

self.addEventListener('message', (e) => {
  // console.log('[Formula Worker] 📩 Received message:', e.data); // Too noisy
});
