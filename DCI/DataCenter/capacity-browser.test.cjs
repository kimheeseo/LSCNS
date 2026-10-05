'use strict';
const assert = require('node:assert/strict');
const fs = require('node:fs');
const http = require('node:http');
const path = require('node:path');
const {chromium} = require('playwright');
async function main() {
  const root = __dirname;
  const mime = {'.html': 'text/html', '.js': 'text/javascript', '.json': 'application/json', '.css': 'text/css', '.woff2': 'font/woff2'};
  const server = http.createServer((req, res) => {
    const requestPath = decodeURIComponent(req.url.split('?')[0]);
    const fontRoot = process.env.CAPACITY_QA_FONT_DIR;
    const dir = requestPath.startsWith('/qa-font/') && fontRoot ? fontRoot : root;
    const file = path.resolve(dir, '.' + (dir === fontRoot ? requestPath.replace('/qa-font', '') : requestPath), requestPath.endsWith('/') ? 'index.html' : '');
    if (!file.startsWith(dir + path.sep)) {res.writeHead(403); res.end(); return;}
    fs.readFile(file, (error, data) => {res.writeHead(error ? 404 : 200, {'Content-Type': mime[path.extname(file)] || 'application/octet-stream'}); res.end(error ? 'Not found' : data);});
  });
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  const browser = await chromium.launch({headless: true, executablePath: process.env.CAPACITY_CHROMIUM_BIN || undefined, args: ['--no-sandbox', '--disable-dev-shm-usage', '--use-gl=angle', '--use-angle=swiftshader', '--disable-gpu'], env: {...process.env}});
  const context = await browser.newContext({viewport: {width: 1440, height: 1000}, acceptDownloads: true});
  const page = await context.newPage(), errors = [];
  page.on('pageerror', e => errors.push(e.message));
  const url = 'http://127.0.0.1:' + server.address().port + '/';
  const qaDir = process.env.CAPACITY_QA_OUTPUT || '/tmp/dc-capacity-qa';
  try {
    await page.goto(url);
    await page.waitForSelector('#capacityButton');
    let fontCss;
    if (process.env.CAPACITY_QA_FONT_DIR) {
      const css = fs.readFileSync(path.join(process.env.CAPACITY_QA_FONT_DIR, '400.css'), 'utf8').replaceAll('./files/', url + 'qa-font/files/');
      fontCss = css + 'body,input,select,button{font-family:"Noto Sans KR",sans-serif!important}';
      await page.addStyleTag({content: fontCss}); await page.evaluate(() => document.fonts.ready);
    }
    await page.click('#capacityButton');
    await page.waitForFunction(() => window.DCCapacityDesign?.summary.gpus === 10000);
    assert.equal(await page.locator('#view-result .cap-content').isVisible(), true);
    assert.equal(await page.locator('#view-result .resultPanel').isVisible(), false);
    await page.screenshot({path: path.join(qaDir, 'capacity-desktop.png'), fullPage: true});
    await page.fill('#cap-scale', '랙 2,000개, 랙당 10kW');
    await page.selectOption('#cap-workload', 'colo');
    await page.waitForFunction(() => window.DCCapacityDesign?.summary.facilityKw === 28000);
    await page.click('#tab-bom');
    const csvPromise = page.waitForEvent('download'); await page.click('#cap-csv'); const csv = await csvPromise;
    const csvPath = path.join(qaDir, 'capacity-test.csv'); await csv.saveAs(csvPath);
    assert.match(fs.readFileSync(csvPath, 'utf8'), /사용 가정/);
    const xlsxPromise = page.waitForEvent('download'); await page.click('#exportWorkbook'); const xlsx = await xlsxPromise;
    const xlsxPath = path.join(qaDir, 'capacity-test.xlsx'); await xlsx.saveAs(xlsxPath);
    const sheets = await page.evaluate(async () => {const buf = await new ExcelJS.Workbook().xlsx.writeBuffer(); return !!buf;});
    assert.ok(sheets); assert.ok(fs.statSync(xlsxPath).size > 10000);
    await page.click('#cap-detail');
    await page.selectOption('#cap-dual-compute', '2');
    await page.waitForFunction(() => window.DCCapacityDesign?.summary.endpointPorts === 192000);
    await page.click('#cap-add-phase');
    await page.fill('[data-phase-field="start"]', '33'); await page.fill('[data-phase-field="end"]', '63');
    await page.waitForFunction(() => window.DCCapacityDesign?.phaseRows.length === 2);
    await page.fill('#cap-scale', '100MW'); await page.selectOption('#cap-workload', 'training');
    await page.selectOption('#cap-dual-compute', '1');
    await page.click('[data-delete-phase]'); await page.fill('#cap-buildings', '1');
    await page.waitForFunction(() => window.DCCapacityDesign?.summary.racks === 1667);
    await page.fill('#cap-scale', '130MW'); await page.selectOption('#cap-basis', 'facility');
    await page.waitForFunction(() => window.DCCapacityDesign?.summary.itKw === 100000);
    await page.fill('#cap-scale', '15GW'); await page.selectOption('#cap-basis', 'it'); await page.fill('#cap-buildings', '');
    await page.waitForFunction(() => window.DCCapacityDesign?.summary.racks === 250000);
    assert.equal(await page.evaluate(() => window.DCCapacityDesign.summary.blocks), 7813);
    await page.selectOption('#cap-coefficient-group', 'selected');
    await page.fill('[data-coefficient="equipmentPresets.generic8.computePorts"]', '2');
    await page.waitForFunction(() => window.DCCapacityDesign?.summary.endpointPorts === 19000000);
    await page.fill('[data-coefficient="equipmentPresets.generic8.computePorts"]', '1');
    await page.fill('#cap-scale', 'GB200 NVL72 랙 500개');
    await page.waitForFunction(() => window.DCCapacityDesign?.summary.gpus === 36000);
    assert.equal(await page.inputValue('#cap-equipment'), 'nvl72');
    await page.fill('#cap-preset-name', 'OEM 사용자 랙'); await page.click('#cap-add-preset');
    assert.ok((await page.inputValue('#cap-equipment')).startsWith('custom-'));
    const customId = await page.inputValue('#cap-equipment');
    await page.fill('[data-coefficient="equipmentPresets.' + customId + '.computePorts"]', '2');
    await page.waitForFunction(() => window.DCCapacityDesign?.summary.endpointPorts === 99000);
    await page.check('#cap-cpo'); await page.waitForFunction(() => window.DCCapacityDesign?.alternatives.length === 4);
    await page.reload(); await page.waitForFunction(() => window.DCCapacityDesign?.alternatives.length === 4);
    if (fontCss) {await page.addStyleTag({content: fontCss}); await page.evaluate(() => document.fonts.ready);}
    assert.ok((await page.inputValue('#cap-equipment')).startsWith('custom-'));
    await page.fill('#cap-pue', '0.8'); await page.waitForFunction(() => !window.DCCapacityDesign);
    assert.match(await page.locator('#cap-error').innerText(), /PUE/);
    await page.fill('#cap-pue', '1.3'); await page.waitForFunction(() => !!window.DCCapacityDesign);
    await page.click('#capacityLegacy'); assert.equal(await page.locator('#capacity-input').isVisible(), false);
    assert.equal(await page.locator('#view-result .resultPanel').isVisible(), true);
    await page.click('#capacityButton'); await page.click('#cap-detail'); await page.fill('#cap-scale', 'GPU 10,000개'); await page.selectOption('#cap-equipment', 'generic8'); await page.fill('#cap-buildings', '');
    await page.waitForFunction(() => window.DCCapacityDesign?.summary.gpus === 10000);
    await page.setViewportSize({width: 390, height: 844}); await page.click('#cap-quick'); await page.click('#tab-result');
    await page.screenshot({path: path.join(qaDir, 'capacity-mobile.png'), fullPage: true});
    const overflow = await page.evaluate(() => ({width: window.innerWidth, scrollWidth: document.documentElement.scrollWidth, elements: [...document.querySelectorAll('body *')].filter(x => x.getBoundingClientRect().right > window.innerWidth + 1 && getComputedStyle(x).position !== 'absolute').slice(0, 10).map(x => x.id || x.className || x.tagName)}));
    assert.ok(overflow.scrollWidth <= overflow.width + 1, 'Mobile horizontal overflow: ' + JSON.stringify(overflow));
    assert.deepEqual(errors, []);
    console.log('Browser checks passed: modes, A/B/C, 15GW, assumptions, redundancy, phases, custom presets, reload, CSV/Excel, invalid input, legacy and mobile layout.');
  } finally {await browser.close(); await new Promise(resolve => server.close(resolve));}
}
main().catch(error => {console.error(error); process.exitCode = 1;});
