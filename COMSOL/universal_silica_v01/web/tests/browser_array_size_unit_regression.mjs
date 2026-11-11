/**
 * Regression for the reported "Array Build Selected draws no circles" case.
 * Run with the local site served at :8765:
 *   node tests/browser_array_size_unit_regression.mjs
 */
import assert from 'node:assert/strict';
import {createRequire} from 'node:module';
const {chromium}=createRequire(import.meta.url)('playwright');
const browser=await chromium.launch({headless:true,args:['--no-sandbox','--disable-dev-shm-usage']});
try{
 const page=await browser.newPage({viewport:{width:1500,height:900}});
 const errors=[];
 page.on('pageerror',error=>errors.push(error.message));
 await page.goto(process.env.TEST_BASE_URL||'http://127.0.0.1:8765/COMSOL/universal_silica_v01/web/',{waitUntil:'load',timeout:90000});
 await page.locator('[data-wizard="cancel"]').click();
 // Start an empty 2D Mode Analysis workspace.
 await page.evaluate(()=>document.dispatchEvent(new CustomEvent('ls-wizard-finished',{detail:{physics:'ewfd',study:'mode'}})));
 await page.locator('summary[data-action="geometry"]').click();
 assert.equal(await page.locator('#ls-cad-root-unit').inputValue(),'um','2D fiber geometry must default to micrometers');
 // Create exactly one preceding source Circle, then Array.
 await page.locator('#ls-cad-add-circle').click();
 await page.locator('#ls-cad-add-array').click();
 assert.equal(await page.locator('#ls-cad-inputs input:checked').count(),1,'Single preceding source should be selected automatically');
 await page.locator('#ls-cad-array-type').selectOption('Linear');
 await page.locator('#ls-cad-array-dx').fill('125');
 await page.locator('#ls-cad-build-selected').click();
 assert.equal(await page.locator('#ls-cad-view [data-cad-shape="arr1"]').count(),1);
 assert.match(await page.locator('#ls-cad-array-feedback').textContent(),/Size = 1/);
 // The reported workflow: Size=6 must create six circles after Build Selected.
 await page.locator('#ls-cad-array-size').fill('6');
 await page.locator('#ls-cad-build-selected').click();
 assert.equal(await page.locator('#ls-cad-view [data-cad-shape="arr1"]').count(),6);
 const objects=await page.evaluate(()=>window.LSWaveCADService.entities().filter(e=>e.owner==='arr1'));
 assert.equal(objects.length,6);
 for(let i=1;i<6;i++)assert.ok(Math.abs(objects[i].x-objects[i-1].x-125e-6)<1e-10);
 assert.match(await page.locator('#ls-cad-array-feedback').textContent(),/6 circles/);
 // Reproduce the other screenshot: x=2 in meters makes outputs microscopic.
 await page.locator('[data-cad-node="arr1"]').click();
 await page.locator('summary[data-action="geometry"]').click();
 await page.locator('#ls-cad-root-unit').selectOption('m');
 await page.locator('[data-cad-node="arr1"]').click();
 await page.locator('#ls-cad-array-dx').fill('2');
 assert.match(await page.locator('#ls-cad-array-feedback').textContent(),/2 meters/);
 assert.match(await page.locator('#ls-cad-array-feedback').textContent(),/1000/);
 await page.locator('#ls-cad-array-use-um').click();
 assert.equal(await page.locator('#ls-cad-array-dx').inputValue(),'2');
 await page.locator('#ls-cad-build-selected').click();
 assert.equal(await page.locator('#ls-cad-view [data-cad-shape="arr1"]').count(),6);
 assert.equal(await page.locator('#ls-cad-array-use-um').isHidden(),true);
 assert.deepEqual(errors,[]);
 console.log('PASS: Array Size=1 vs 6, unit mismatch warning, micrometer switch, Build Selected outputs');
}finally{await browser.close();}
