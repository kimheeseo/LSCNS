/**
 * Browser M1 smoke + old solver UI non-regression (requires Playwright).
 * TEST_BASE_URL can override http://127.0.0.1:8765/COMSOL/universal_silica_v01/web/
 */
import assert from "node:assert/strict";
import {createRequire} from "node:module";
const {chromium}=createRequire(import.meta.url)("playwright");
const browser=await chromium.launch({headless:true,args:["--no-sandbox","--disable-dev-shm-usage"]});
try{
 const page=await browser.newPage({viewport:{width:1550,height:1000},acceptDownloads:true});
 const errors=[];page.on("pageerror",e=>errors.push(e.message));
 await page.goto(process.env.TEST_BASE_URL||"http://127.0.0.1:8765/COMSOL/universal_silica_v01/web/",{waitUntil:"load",timeout:90000});
 // Startup wizard can cover the workbench; it must not be modified by M1.
 if(await page.locator('[data-wizard="cancel"]').isVisible())await page.locator('[data-wizard="cancel"]').click();
 const tab=page.locator("#ls-m1-validation-tab");
 assert.ok(await tab.isVisible(),"M1 nav shortcut available by default");
 assert.ok(await page.locator("#ls-m1-validation").isHidden(),"Validation section starts hidden");
 const oldNeff=await page.locator("#neff").textContent();
 const oldPreset=await page.locator("#preset").inputValue();
 const oldRadius=await page.locator("#core_radius_um").inputValue();
 await tab.click();
 assert.ok(await page.locator("#ls-m1-validation").isVisible());
 await page.locator("#ls-m1-run").click();
 await page.waitForFunction(()=>window.LSWaveM1Validation?.latest()?.totalCount===28,undefined,{timeout:90000});
 assert.equal(await page.locator("#ls-m1-results tr").count(),28);
 const report=await page.evaluate(()=>window.LSWaveM1Validation.latest());
 assert.equal(report.passed,true);
 assert.equal(report.passedCount,28);
 assert.equal(await page.locator("#preset").inputValue(),oldPreset,"M1 must not switch presets");
 assert.equal(await page.locator("#core_radius_um").inputValue(),oldRadius,"M1 must not change active design");
 assert.equal(await page.locator("#neff").textContent(),oldNeff,"M1 must not change original FEM cards");
 const [download]=await Promise.all([
  page.waitForEvent("download"),
  page.locator("#ls-m1-export").click()
 ]);
 assert.equal(download.suggestedFilename(),"ls_wave_optics_m1_validation.json");
 await page.locator("#ls-m1-hide").click();
 assert.equal(await page.locator("#ls-m1-validation").isHidden(),true);
 // Legacy controls remain available.
 assert.ok(await page.locator("#calculateBtn").isEnabled());
 assert.ok(await page.locator("#meshBtn").isEnabled());
 assert.ok(await page.locator("#sweepBtn").isEnabled());
 assert.deepEqual(errors,[]);
 console.log("PASS M1 browser: 28 rows, report download, all original controls and model values unchanged");
}finally{await browser.close();}
