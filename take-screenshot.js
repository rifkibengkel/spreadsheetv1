const puppeteer = require('puppeteer');
const path = require('path');

(async () => {
  const browser = await puppeteer.launch({ headless: 'new', args: ['--no-sandbox', '--disable-setuid-sandbox'] });
  const page = await browser.newPage();
  await page.setViewport({ width: 1280, height: 800 });
  
  await page.goto('http://localhost:3000', { waitUntil: 'networkidle2', timeout: 60000 });
  
  // Wait for auto-test to complete
  await new Promise(r => setTimeout(r, 8000));
  
  const screenshotPath = path.resolve(__dirname, 'poc-screenshot.png');
  await page.screenshot({ path: screenshotPath, fullPage: false });
  console.log('Screenshot saved to:', screenshotPath);
  
  await browser.close();
  process.exit(0);
})();
