/**
 * Optional Chromium/Playwright M2/M3 UI smoke.
 * Serve LSCNS repo root: python3 -m http.server 8765
 * From web/: node tests/browser_m2_m3_v07.mjs
 * Not required by GitHub Pages at runtime (developer test only).
 */
import assert from "node:assert/strict";
import {createRequire} from "node:module";
const {chromium}=createRequire(import.meta.url)("playwright");
const browser=await chromium.launch({headless:true,args:["--no-sandbox","--disable-dev-shm-usage"]});
try{
 const page=await browser.newPage({viewport:{width:1520,height:950},acceptDownloads:true});
 const errors=[];page.on("pageerror",e=>errors.push(e.message));
 await page.goto(process.env.TEST_BASE_URL||"http://127.0.0.1:8765/COMSOL/universal_silica_v01/web/",{waitUntil:"load",timeout:90000});
 if(await page.locator('[data-wizard="cancel"]').isVisible())await page.locator('[data-wizard="cancel"]').click();
 const original=await page.evaluate(()=>({
  preset:document.querySelector("#preset")?.value,
  neff:document.querySelector("#neff")?.textContent,
  mfd:document.querySelector("#mfd")?.textContent,
  aeff:document.querySelector("#aeff")?.textContent
 }));
 assert.ok(await page.locator("#m2v6-nav").isVisible());
 await page.locator("#m2v6-nav").click();
 await page.locator("#m2v6-run").click();
 await page.waitForFunction(()=>window.LSWaveM2Preview?.latest()?.schemaVersion===2);
 assert.ok((await page.locator("#m2v6-rows tr").count())>=6);
 assert.ok((await page.locator("#m2v6-status").textContent()).includes("미리보기 성공"));
 const [m2download]=await Promise.all([page.waitForEvent("download"),page.locator("#m2v6-save").click()]);
 assert.equal(m2download.suggestedFilename(),"LS_Wave_M2_v0.6_material_preview.json");
 await page.locator("#m2v6-core-kind").selectOption("f_mol");
 await page.locator("#m2v6-core-mol").fill("3");
 await page.locator("#m2v6-run").click();
 assert.ok((await page.locator("#m2v6-status").textContent()).includes("requires user Sellmeier coefficients"));
 assert.ok(await page.locator("#m3v7-nav").isVisible());
 await page.locator("#m3v7-nav").click();
 for(const n of [41,51,69])await page.locator('#m3v7-grids input[value="'+n+'"]').uncheck();
 await page.locator("#m3v7-run").click();
 await page.waitForFunction(()=>window.LSWaveM3Study?.latest()?.schemaVersion===2,undefined,{timeout:90000});
 const result=await page.evaluate(()=>window.LSWaveM3Study.latest());
 assert.equal(result.convergence.length,1);
 assert.equal(result.portableVectorResult.schema,"LSCNS_vector_mode_v1");
 assert.ok(result.portableVectorResult.modes.every(m=>m.field_map?.intensity_normalized_to_max_total?.Ez));
 assert.equal(await page.locator("#m3v7-convergence tr").count(),1);
 const [vectorDownload]=await Promise.all([page.waitForEvent("download"),page.locator("#m3v7-export").click()]);
 assert.equal(vectorDownload.suggestedFilename(),"LS_Wave_Vector_v0.7_colab_compatible.json");
 await page.locator('#m3v7-grids input[value="29"]').uncheck();
 await page.locator('#m3v7-grids input[value="91"]').check();
 await page.locator("#m3v7-run").click();
 await page.locator("#m3v7-cancel").click();
 assert.ok((await page.locator("#m3v7-status").textContent()).includes("사용자 취소"));
 assert.equal(await page.locator("#m3v7-export").isDisabled(),true);
 const unchanged=await page.evaluate(()=>({
  preset:document.querySelector("#preset")?.value,
  neff:document.querySelector("#neff")?.textContent,
  mfd:document.querySelector("#mfd")?.textContent,
  aeff:document.querySelector("#aeff")?.textContent
 }));
 assert.deepEqual(unchanged,original,"Existing solver cards must not be modified by M2/M3.");
 assert.ok(await page.locator("#fv-json-file").count()===1);
 assert.ok(await page.locator("#ls-m1q-run").count()===1);
 assert.deepEqual(errors,[]);
 console.log("PASS M2/M3 v0.6/0.7 browser: chart, CSV, mol/wt warning, Worker, cancellation, old controls");
}finally{await browser.close();}
