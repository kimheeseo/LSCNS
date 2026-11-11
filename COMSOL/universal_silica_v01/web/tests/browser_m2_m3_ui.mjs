/**
 * Optional browser E2E M2 & M3; run from repo root with HTTP server at :8765.
 * Requires playwright package and Chromium. Tests no mutation of original UI.
 */
import assert from "node:assert/strict";
import {createRequire} from "node:module";
const {chromium}=createRequire(import.meta.url)("playwright");
const browser=await chromium.launch({headless:true,args:["--no-sandbox","--disable-dev-shm-usage"]});
try{
 const page=await browser.newPage({viewport:{width:1520,height:970}});
 const errors=[];page.on("pageerror",e=>errors.push(e.message));
 await page.goto(process.env.TEST_BASE_URL||"http://127.0.0.1:8765/COMSOL/universal_silica_v01/web/",{waitUntil:"load"});
 if(await page.locator('[data-wizard="cancel"]').isVisible())await page.locator('[data-wizard="cancel"]').click();
 assert.ok(await page.locator("#ls-m2-nav").isVisible());
 const initialNeff=await page.locator("#neff").textContent();
 const initialProfile=await page.locator("#profile").inputValue();
 await page.locator("#ls-m2-nav").click();
 await page.locator("#m2-run").click();
 await page.waitForFunction(()=>window.LSWaveMaterialM2?.last()?.result?.n_eff!==undefined);
 const m2=await page.evaluate(()=>window.LSWaveMaterialM2.last());
 assert.ok(Number.isFinite(m2.result.n_eff));
 assert.ok(Math.abs(m2.result.delta_n-.005)<1e-9);
 assert.equal(await page.locator("#neff").textContent(),initialNeff);
 assert.equal(await page.locator("#profile").inputValue(),initialProfile);
 await page.locator("#preset").selectOption("vector");
 assert.ok(await page.locator("#ls-m3").isVisible());
 assert.ok(await page.locator("#fv-json-file").count()===1,"Existing Colab file import retained");
 await page.locator("#m3-run").click();
 await page.waitForFunction(()=>window.LSWaveVectorM3?.latest()?.modes?.length===2,undefined,{timeout:90000});
 const r=await page.evaluate(()=>window.LSWaveVectorM3.latest());
 assert.equal(r.schema,"LSCNS_vector_mode_v1");
 assert.equal(r.modes.length,2);
 assert.ok(r.modes.every(m=>m.residual<1e-7));
 assert.equal(await page.locator("#m3-numbers tr").count(),3);
 assert.deepEqual(errors,[]);
 console.log("PASS M2/M3 browser optional section: original scalar state unchanged, M2 and worker FDFD output");
}finally{await browser.close();}
