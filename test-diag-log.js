const puppeteer = require('puppeteer');

(async () => {
  const browser = await puppeteer.launch({ headless: 'new', args: ['--no-sandbox', '--disable-setuid-sandbox'] });
  const page = await browser.newPage();
  page.on('console', msg => {
    if (msg.text().includes('🔥')) console.log(msg.text());
  });
  page.on('pageerror', err => {
    console.log('🔥 [ERROR]', err.toString());
  });
  await page.goto('http://localhost:3000', { waitUntil: 'domcontentloaded', timeout: 30000 });
  await new Promise(r => setTimeout(r, 12000));
  await browser.close();
  process.exit(0);
})();
