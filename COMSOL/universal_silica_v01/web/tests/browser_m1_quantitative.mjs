/** Optional Playwright M1 v0.5.1 smoke/regression.
 * Serve LSCNS repository root: python3 -m http.server 8765
 * Run in web/: node tests/browser_m1_quantitative.mjs
 * Or TEST_BASE_URL=https://.../web/ node tests/browser_m1_quantitative.mjs
 */
import assert from "node:assert/strict";
import {createRequire} from "node:module";
const {chromium}=createRequire(import.meta.url)("playwright");
const browser=await chromium.launch({headless:true,args:["--no-sandbox","--disable-dev-shm-usage"]});
try{
 const page=await browser.newPage({viewport:{width:1500,height:920},acceptDownloads:true});
 const errors=[];
 page.on("pageerror",e=>errors.push(e.message));
 await page.goto(process.env.TEST_BASE_URL||"http://127.0.0.1:8765/COMSOL/universal_silica_v01/web/",{waitUntil:"load",timeout:90000});
 if(await page.locator('[data-wizard="cancel"]').isVisible())await page.locator('[data-wizard="cancel"]').click();
 await page.waitForFunction(()=>!["—",""].includes(document.getElementById("neff")?.textContent?.trim()||""),undefined,{timeout:90000});
 const legacy=await page.evaluate(()=>({
  preset:document.getElementById("preset").value,
  neff:document.getElementById("neff").textContent,
  mfd:document.getElementById("mfd").textContent,
  aeff:document.getElementById("aeff").textContent
 }));
 assert.ok(await page.locator("#ls-m1q-nav").isVisible());
 assert.equal(await page.locator("#ls-m1q-section").isHidden(),true);
 await page.locator("#ls-m1q-nav").click();
 assert.equal(await page.locator("#ls-m1q-section").isVisible(),true);
 await page.locator("#ls-m1q-run").click();
 await page.waitForFunction(()=>window.LSWaveM1Quantitative?.latest()?.caseCount===4,undefined,{timeout:90000});
 assert.equal(await page.locator("#ls-m1q-cases tr").count(),4);
 assert.equal(await page.locator("#ls-m1q-regression tr").count(),5);
 const result=await page.evaluate(()=>window.LSWaveM1Quantitative.latest());
 assert.equal(result.passCount,4);
 assert.equal(result.regressionPassCount,5);
 assert.ok(result.cases.every(c=>c.status==="PASS"));
 assert.equal(result.cases.find(c=>c.id==="requested_41").comsolStatus,"미검증");
 const still=await page.evaluate(()=>({
  preset:document.getElementById("preset").value,
  neff:document.getElementById("neff").textContent,
  mfd:document.getElementById("mfd").textContent,
  aeff:document.getElementById("aeff").textContent
 }));
 assert.deepEqual(still,legacy,"Original UI model / neff / MFD / Aeff must be unchanged.");
 const [json]=await Promise.all([page.waitForEvent("download"),page.locator("#ls-m1q-json").click()]);
 assert.equal(json.suggestedFilename(),"LS_Wave_M1_v0.5.1_report.json");
 const [csv]=await Promise.all([page.waitForEvent("download"),page.locator("#ls-m1q-csv").click()]);
 assert.equal(csv.suggestedFilename(),"LS_Wave_M1_v0.5.1_report.csv");
 assert.ok(await page.locator("#ls-m1-validation-tab").count()===1,"Existing M1 28-check tool still mounted.");
 assert.ok(await page.locator("#m2-run").count()===1,"M2 still exists.");
 assert.ok(await page.locator("#m3-run").count()===1,"M3 still exists.");
 assert.deepEqual(errors,[]);
 console.log("PASS M1 browser: 4-case PASS/FAIL, 5-preset before/after, JSON/CSV, and unchanged user results.");
} finally {await browser.close();}
