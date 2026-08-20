const puppeteer = require('puppeteer');

async function runTest() {
  console.log("🚀 Starting Puppeteer Test...");
  const browser = await puppeteer.launch({ headless: 'new' });
  const page = await browser.newPage();

  page.on('console', msg => {
    const text = msg.text();
    if (!text.includes('DOM element is missing') && !text.includes('React DevTools')) {
      console.log(`[Browser] ${text}`);
    }
  });

  console.log("🌐 Navigating to http://localhost:3000 ...");
  await page.goto('http://localhost:3000', { waitUntil: 'networkidle2' });

  // Wait 10 seconds for the auto-test to run
  await new Promise(r => setTimeout(r, 10000));

  await browser.close();
  process.exit(0);
}

runTest().catch(console.error);
