const fs = require('fs');
const puppeteer = require('puppeteer');

(async () => {
  console.log('=== Testing Cell Editor Save vs Cancel Behavior ===');

  const browser = await puppeteer.launch({
    headless: 'new',
    args: ['--max-old-space-size=4096']
  });

  const page = await browser.newPage();
  await page.goto('http://localhost:3000');
  await page.waitForFunction('window.univerAPI !== undefined');

  const result = await page.evaluate(async () => {
    const api = window.univerAPI;
    const commandService = api.getCommandService ? api.getCommandService() : null;

    // Test command parameters for set-activate-cell-edit
    let testWithSaveTrue = null;
    let testWithoutSaveTrue = null;

    // Try executing set-activate-cell-edit with { active: false, save: true }
    if (commandService) {
      try {
        testWithSaveTrue = await commandService.executeCommand('sheet.command.set-activate-cell-edit', { active: false, save: true });
      } catch (e) {
        testWithSaveTrue = e.message;
      }
    }

    return {
      testWithSaveTrue,
      hasEndEdit: !!api.endEdit
    };
  });

  console.log('Command execution result:', JSON.stringify(result, null, 2));
  await browser.close();
})();
