/**
 * Optional Playwright smoke test: strict Yee M3 shortcut, M2 provenance,
 * preserved legacy Colab import and real Worker mode. Not a cloud dependency.
 *
 * Serve LSCNS repository root: python3 -m http.server 8765
 * Run: node tests/browser_m2_m3_v071.mjs
 */
import assert from "node:assert/strict";
import {createRequire} from "node:module";
const {chromium}=createRequire(import.meta.url)("playwright");
const browser=await chromium.launch({headless:true,args:["--no-sandbox","--disable-dev-shm-usage"]});
try {
 const page=await browser.newPage({viewport:{width:1560,height:960},acceptDownloads:true});
 const errors=[];page.on("pageerror",e=>errors.push(e.message));
 await page.goto(process.env.TEST_BASE_URL||
  "http://127.0.0.1:8765/COMSOL/universal_silica_v01/web/?v=v0.7.1",
  {waitUntil:"load",timeout:90000});
 if(await page.locator('[data-wizard="cancel"]').isVisible())
  await page.locator('[data-wizard="cancel"]').click();
 await page.waitForFunction(()=>window.LSWaveVectorWorkbench?.show&&window.LSWaveM3YeeQuickstart?.open);
 assert.equal(await page.locator("#m3-yee-quicklink").count(),1);
 assert.equal(await page.locator("#m2v6-nav").count(),1);
 assert.equal(await page.locator("#fv-json-file").count(),1);
 assert.ok((await page.locator("body").textContent()).includes("v0.7.1"));
 await page.locator("#m3-yee-quicklink").click();
 assert.ok(await page.locator('#vwb [data-vb-pane="m3"]').isVisible());
 await page.locator("#vb-m3-preset").click();
 assert.equal(await page.locator("#vb-kind").inputValue(),"radial");
 assert.equal(await page.locator("#vb-pml").isChecked(),false);
 assert.equal(await page.locator("#vb-bend").inputValue(),"0");
 assert.equal(await page.locator("#vb-kcore").inputValue(),"0");
 assert.equal(await page.locator("#vb-kclad").inputValue(),"0");
 assert.equal(await page.locator("#vb-material-source").inputValue(),"constant");
 assert.equal(await page.locator("#vb-grid").inputValue(),"48");
 assert.equal(await page.locator("#vb-modes").inputValue(),"2");
 // Small grid to keep UI smoke inexpensive. Default baseline remains 48.
 await page.locator("#vb-grid").fill("24");
 await page.locator("#vb-run").click();
 await page.waitForFunction(()=>window.LSWaveVectorWorkbench?.latest?.()?.result?.modes?.length===2,
  undefined,{timeout:180000});
 const data=await page.evaluate(()=>window.LSWaveVectorWorkbench.latest().result);
 assert.equal(data.schema,"LSCNS_vector_mode_v1");
 assert.equal(data.modes.length,2);
 assert.ok(data.modes.every(m=>Number.isFinite(m.neff)));
 assert.deepEqual(errors,[]);
 console.log("PASS v0.7.1 M2 source and strict Yee M3 entry, no-PML preset, 2-mode Worker, legacy Colab uploader preserved");
} finally {
 await browser.close();
}
