/**
 * Rectangle → Union → Delete Entities → Form Union browser regression.
 * Prerequisite: site served at :8765, Playwright installed, Chromium available.
 */
import assert from "node:assert/strict";
import {createRequire} from "node:module";
const {chromium}=createRequire(import.meta.url)("playwright");
const browser=await chromium.launch({headless:true,args:["--no-sandbox","--disable-dev-shm-usage"]});
try{
 const page=await browser.newPage({viewport:{width:1550,height:940}});
 const errors=[];page.on("pageerror",e=>errors.push(e.message));
 await page.goto(process.env.TEST_BASE_URL||"http://127.0.0.1:8765/COMSOL/universal_silica_v01/web/",{waitUntil:"load",timeout:90000});
 await page.locator('[data-wizard="cancel"]').click();
 await page.evaluate(()=>document.dispatchEvent(new CustomEvent("ls-wizard-finished",{detail:{physics:"ewfd",study:"mode"}})));
 await page.locator('summary[data-action="geometry"]').click({button:"right"});
 await page.locator('#ls-cad-menu button[data-cmd="rectangle"]').click();
 assert.equal(await page.locator('#ls-cad-subtitle').textContent(),"Rectangle");
 await page.locator('#ls-cad-rect-base').selectOption('Center');
 await page.locator('#ls-cad-rect-width').fill('40[um]');
 await page.locator('#ls-cad-rect-height').fill('20[um]');
 await page.locator('#ls-cad-rect-rotation').fill('15');
 assert.equal(await page.locator('#ls-cad-view [data-cad-shape="r1"] polygon').count(),1);
 assert.ok(await page.locator('#ls-cad-rect-layers').count(),'Rectangle Layers accordion is present');
 // Corner means the anchored vertex remains at x/y during rotation.
 await page.locator('#ls-cad-rect-base').selectOption('Corner');
 await page.locator('#ls-cad-rect-x').fill('0[um]');
 await page.locator('#ls-cad-rect-y').fill('0[um]');
 await page.locator('#ls-cad-rect-rotation').fill('90');
 const cornerRect=await page.evaluate(()=>window.LSWaveCADService.entities().find(s=>s.owner==='r1'));
 assert.ok(Math.abs(cornerRect.x+10e-6)<1e-10&&Math.abs(cornerRect.y-20e-6)<1e-10,'Corner rotates about the defined corner');
 await page.locator('#ls-cad-rect-base').selectOption('Center');
 await page.locator('#ls-cad-rect-rotation').fill('15');
 await page.locator('#ls-cad-add-circle').click();
 await page.locator('#ls-cad-radius').fill('25[um]');
 await page.locator('#ls-cad-x').fill('10[um]');
 await page.locator('#ls-cad-add-union').click();
 assert.equal(await page.locator('#ls-cad-subtitle').textContent(),"Union");
 const inputs=page.locator('#ls-cad-inputs input[type=checkbox]');
 assert.equal(await inputs.count(),2);
 await inputs.nth(0).check();
 await inputs.nth(1).check();
 await page.locator('#ls-cad-build-selected').click();
 const union=await page.evaluate(()=>window.LSWaveCADService.entities());
 assert.equal(union.length,1);
 assert.equal(union[0].kind,"union");
 assert.ok(union[0].contourCount>0);
 assert.equal(await page.locator('#ls-cad-view [data-cad-shape="uni1"] path').count()>0,true);
 await page.locator('#ls-cad-union-interior').check();
 await page.locator('#ls-cad-build-selected').click();
 assert.equal(await page.locator('#ls-cad-view [data-cad-shape="uni1"] path').count()>0,true);
 // Delete the generated preview Domain object.
 await page.locator('#ls-cad-add-delete-entities').click();
 assert.equal(await page.locator('#ls-cad-delete-level').inputValue(),'Object');
 const del=page.locator('#ls-cad-inputs input[data-cad-input-ref="uni1"]');
 await del.check();
 await page.locator('#ls-cad-build-selected').click();
 assert.equal(await page.evaluate(()=>window.LSWaveCADService.entities().length),0);
 // Create fresh finalization inputs and preview their combined silhouette.
 await page.locator('#ls-cad-add-rectangle').click();
 await page.locator('#ls-cad-rect-width').fill('60[um]');
 await page.locator('#ls-cad-rect-height').fill('60[um]');
 await page.locator('#ls-cad-add-circle').click();
 await page.locator('#ls-cad-radius').fill('24[um]');
 await page.locator('[data-cad-form-union]').click();
 assert.ok(await page.locator('#ls-cad-form-settings').isVisible());
 assert.equal(await page.locator('#ls-cad-subtitle').textContent(),'Form Union/Assembly');
 assert.equal(await page.locator('#ls-cad-form-action').inputValue(),'union');
 assert.equal(await page.locator('#ls-cad-object-select').inputValue(),'fin');
 assert.equal(await page.locator('#ls-cad-form-repair').inputValue(),'Automatic');
 await page.locator('#ls-cad-form-interior').uncheck();
 await page.locator('#ls-cad-build-selected').click();
 assert.equal(await page.locator('#ls-cad-view [data-cad-shape="root"] path').count()>0,true);
 await page.locator('#ls-cad-form-action').selectOption('assembly');
 await page.locator('#ls-cad-build-selected').click();
 assert.equal(await page.locator('#ls-cad-view [data-cad-shape="root"]').count(),0,'Assembly keeps inputs separate');
 assert.ok((await page.locator('#ls-cad-view [data-cad-shape="r2"]').count())>0);
 assert.deepEqual(errors,[]);
 console.log("PASS: Rectangle Corner rotation, Union 2D outline, Delete Entities Object, Form Union/Assembly selected settings");
}finally{await browser.close();}
