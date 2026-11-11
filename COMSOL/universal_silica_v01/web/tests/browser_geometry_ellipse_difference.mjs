/**
 * Ellipse / Difference / Delete Entities browser workflow (headless Playwright).
 * Requires the site served locally at http://127.0.0.1:8765 or TEST_BASE_URL.
 * Browser execution is intentionally separate from the unit test.
 */
import assert from "node:assert/strict";
import {createRequire} from "node:module";
const {chromium}=createRequire(import.meta.url)("playwright");
const browser=await chromium.launch({headless:true,args:["--no-sandbox","--disable-dev-shm-usage"]});
try{
 const page=await browser.newPage({viewport:{width:1500,height:900}});
 const errors=[];
 page.on("pageerror",e=>errors.push(e.message));
 await page.goto(process.env.TEST_BASE_URL||"http://127.0.0.1:8765/COMSOL/universal_silica_v01/web/",{waitUntil:"load",timeout:90000});
 await page.locator('[data-wizard="cancel"]').click();
 await page.evaluate(()=>document.dispatchEvent(new CustomEvent("ls-wizard-finished",{detail:{physics:"ewfd",study:"mode"}})));
 await page.locator('summary[data-action="geometry"]').click({button:"right"});
 await page.locator('#ls-cad-menu button[data-cmd="ellipse"]').click();
 assert.equal(await page.locator('#ls-cad-subtitle').textContent(),"Ellipse");
 await page.locator('#ls-cad-ellipse-a').fill('120[um]');
 await page.locator('#ls-cad-ellipse-b').fill('75[um]');
 await page.locator('#ls-cad-ellipse-sector').fill('360');
 await page.locator('#ls-cad-ellipse-rotation').fill('20');
 assert.equal(await page.locator('#ls-cad-view [data-cad-shape="ell1"] ellipse').count(),1);
 const ellipse=await page.evaluate(()=>window.LSWaveCADService.entities().find(s=>s.owner==='ell1'));
 assert.equal(ellipse?.kind,'ellipse');
 await page.locator('#ls-cad-add-circle').click();
 await page.locator('#ls-cad-radius').fill('20[um]');
 await page.locator('#ls-cad-add-difference').click();
 assert.equal(await page.locator('#ls-cad-subtitle').textContent(),'Difference');
 assert.equal(await page.locator('#ls-cad-diff-add-inputs input').count(),2);
 assert.equal(await page.locator('#ls-cad-diff-sub-inputs input').count(),2);
 await page.locator('#ls-cad-diff-pick-add').click();
 assert.equal(await page.locator('#ls-cad-diff-pick-add').getAttribute('aria-pressed'),'true');
 await page.locator('#ls-cad-diff-pick-add').click();
 await page.locator('#ls-cad-diff-add-inputs [data-cad-diff-ref="ell1"]').check();
 await page.locator('#ls-cad-diff-sub-inputs [data-cad-diff-ref="c1"]').check();
 await page.locator('#ls-cad-build-selected').click();
 const outputs=await page.evaluate(()=>window.LSWaveCADService.entities());
 assert.equal(outputs.length,1,'Selected source shapes are consumed by Difference by default');
 assert.equal(outputs[0].kind,'difference');
 assert.equal(outputs[0].contourCount,2,'Ellipse minus interior circle forms two contours');
 assert.equal(await page.locator('#ls-cad-view [data-cad-shape="dif1"] path').count(),1);
 await page.locator('#ls-cad-diff-keep').check();
 await page.locator('#ls-cad-build-selected').click();
 assert.equal(await page.evaluate(()=>window.LSWaveCADService.entities().length),3);
 // Delete the Difference result, leaving the two kept input objects visible.
 await page.locator('#ls-cad-add-delete-entities').click();
 assert.equal(await page.locator('#ls-cad-delete-level').inputValue(),'Object');
 await page.locator('#ls-cad-inputs [data-cad-input-ref="dif1"]').check();
 await page.locator('#ls-cad-build-selected').click();
 const retained=await page.evaluate(()=>window.LSWaveCADService.entities().map(s=>s.owner));
 assert.deepEqual(retained.sort(),['c1','ell1']);
 assert.deepEqual(errors,[]);
 console.log('PASS: Ellipse 2D rendering → Difference (2 contours) → Delete Entities Object');
}finally{await browser.close();}
