
    const puppeteer = require('puppeteer');
    (async () => {
      const browser = await puppeteer.launch({ headless: 'new' });
      const page = await browser.newPage();
      page.on('console', msg => {
        const text = msg.text();
        if (text.includes('[TEST]') || text.includes('[DEBUG] Cell Render Model:')) {
          console.log(text);
        }
      });
      await page.goto('http://localhost:3000');
      await new Promise(r => setTimeout(r, 5000));
      await browser.close();
    })();
  