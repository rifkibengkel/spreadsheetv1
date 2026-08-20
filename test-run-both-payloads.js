const fs = require('fs');
const puppeteer = require('puppeteer');

async function runBothPayloads() {
  const manualWbData = JSON.parse(fs.readFileSync('manualWbData.json', 'utf8'));
  const barakatWbData = JSON.parse(fs.readFileSync('barakatWbData.json', 'utf8'));

  const browser = await puppeteer.launch({ headless: 'new', args: ['--no-sandbox'] });
  const page = await browser.newPage();
  page.on('console', msg => console.log('[BROWSER]', msg.text()));

  await page.goto('http://localhost:3000', { waitUntil: 'domcontentloaded' });
  await page.waitForFunction('window.univerAPI !== undefined');

  const resManual = await page.evaluate(async (data) => {
    const api = window.univerAPI;
    const oldWbs = api.getAllWorkbooks ? api.getAllWorkbooks() : [];
    oldWbs.forEach(w => { try { api.disposeUnit(w.getId()); } catch(e){} });

    const wb = api.createUniverSheet(data);
    await new Promise(r => setTimeout(r, 2000));

    const sHadiah = wb.getSheetByName('Summary Hadiah');
    return {
      val: sHadiah ? sHadiah.getRange(0, 0, 1, 1).getValue() : null,
      cell: wb.save().sheets['Summary Hadiah']?.cellData?.[0]?.[0]
    };
  }, manualWbData);

  const resBarakat = await page.evaluate(async (data) => {
    const api = window.univerAPI;
    const oldWbs = api.getAllWorkbooks ? api.getAllWorkbooks() : [];
    oldWbs.forEach(w => { try { api.disposeUnit(w.getId()); } catch(e){} });

    const wb = api.createUniverSheet(data);
    await new Promise(r => setTimeout(r, 2000));

    const sHadiah = wb.getSheetByName('Summary Hadiah');
    return {
      val: sHadiah ? sHadiah.getRange(0, 0, 1, 1).getValue() : null,
      cell: wb.save().sheets['Summary Hadiah']?.cellData?.[0]?.[0]
    };
  }, barakatWbData);

  const report = { manual: resManual, barakat: resBarakat };
  fs.writeFileSync('payload-comparison.json', JSON.stringify(report, null, 2));
  console.log('Saved payload-comparison.json');

  await browser.close();
}

runBothPayloads().catch(console.error);
