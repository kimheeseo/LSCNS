/** Lecture screenshots: Conversions submenu; Zoom Extents frames all six disks. */
import assert from 'node:assert/strict';
import {createRequire} from 'node:module';
const {chromium}=createRequire(import.meta.url)('playwright');
const browser=await chromium.launch({headless:true,args:['--no-sandbox','--disable-dev-shm-usage']});
try{
 const page=await browser.newPage({viewport:{width:1500,height:940}});
 const errors=[];page.on('pageerror',e=>errors.push(e.message));
 await page.goto(process.env.TEST_BASE_URL||'http://127.0.0.1:8765/COMSOL/universal_silica_v01/web/',{waitUntil:'load',timeout:90000});
 await page.locator('[data-wizard="cancel"]').click();
 await page.locator('summary[data-action="geometry"]').click({button:'right'});
 const group=page.locator('#ls-cad-menu > ul > li.has-children').filter({has:page.locator(':scope > button').filter({hasText:'Conversions'})});
 await group.locator(':scope > button').hover();assert.ok(await group.locator(':scope > ul').isVisible());
 assert.deepEqual(await group.locator(':scope > ul > li > button').allTextContents(),['Convert to Solid','Convert to Curve','Convert to Point','Split']);
 assert.equal(await group.locator(':scope > ul button:disabled').count(),4,'Keep operations without implemented settings disabled.');
 await page.screenshot({path:'/tmp/comsol_conversions_submenu.png'});await page.keyboard.press('Escape');
 await page.locator('[data-cad-node="c1"]').click();await page.keyboard.press('F3');
 await page.locator('[data-cad-node="c2"]').click();await page.locator('#ls-cad-radius').fill('200[um]');await page.locator('#ls-cad-build-selected').click();
 await page.locator('#ls-cad-add-array').click();await page.locator('#ls-cad-inputs [data-cad-input-ref="c2"]').check();
 await page.locator('#ls-cad-array-type').selectOption('Linear');await page.locator('#ls-cad-array-size').fill('6');
 await page.locator('#ls-cad-array-dx').fill('400[um]');await page.locator('#ls-cad-build-selected').click();
 const fit=page.getByRole('button',{name:'Zoom Extents',exact:true});assert.ok(await fit.isVisible());assert.equal(await fit.getAttribute('title'),'Zoom Extents');
 const circles=async(selector='#ls-cad-view')=>page.locator(selector+' .ls-cad-rendered-surface circle').evaluateAll(els=>els.map(c=>({x:Number(c.getAttribute('cx')),y:Number(c.getAttribute('cy')),r:Number(c.getAttribute('r'))})));
 const inside=async()=>{const cs=await circles();assert.equal(cs.length,6);for(const c of cs)assert.ok(c.x-c.r>=55-1e-7&&c.x+c.r<=618+1e-7&&c.y-c.r>=31-1e-7&&c.y+c.r<=452+1e-7,JSON.stringify(c));
  const min=Math.min(...cs.map(c=>c.x-c.r)),max=Math.max(...cs.map(c=>c.x+c.r));assert.ok(Math.abs((min+max)/2-(55+618)/2)<1e-7);return cs;
 };
 const model=await page.evaluate(()=>JSON.stringify(window.LSWaveCADService.entities()));
 const mini=await circles('#ls-cad-mini-svg');
 await page.locator('#ls-cad-plus').click();await page.locator('#ls-cad-plus').click();
 assert.deepEqual(await circles('#ls-cad-mini-svg'),mini,'The small overview keeps all geometry visible during main-view zoom.');
 const drag=await page.locator('#ls-cad-view').evaluate(svg=>{
  const screen=(x,y)=>{const p=svg.createSVGPoint();p.x=x;p.y=y;const q=p.matrixTransform(svg.getScreenCTM());return {x:q.x,y:q.y};};return {a:screen(260,410),b:screen(350,410)};
 });
 await page.mouse.move(drag.a.x,drag.a.y);await page.mouse.down();await page.mouse.move(drag.b.x,drag.b.y,{steps:5});await page.mouse.up();
 assert.ok((await circles()).some(c=>c.x-c.r<55||c.x+c.r>618),'Start from an actually cropped/translated array.');
 await page.screenshot({path:'/tmp/comsol_zoom_extents_before.png'});
 await fit.click();const framed=await inside();
 assert.equal(await page.evaluate(()=>JSON.stringify(window.LSWaveCADService.entities())),model,'Zoom Extents changes the camera, not dimensions/positions.');
 assert.equal(await page.locator('#ls-cad-array-size').inputValue(),'6');
 const labels=await page.locator('#ls-cad-view [data-axis="x"]').allTextContents();assert.ok(labels.includes('2000'),'World coordinates remain in micrometers across the long array: '+labels);
 await page.screenshot({path:'/tmp/comsol_zoom_extents_six_circles.png'});
 await fit.click();assert.deepEqual(await circles(),framed,'Repeated Zoom Extents is stable.');
 // Zoom after fitting stays centered on the fitted geometry, not the origin.
 await page.locator('#ls-cad-plus').click();const zoomed=await circles();
 assert.ok(Math.abs((Math.min(...zoomed.map(c=>c.x-c.r))+Math.max(...zoomed.map(c=>c.x+c.r)))/2-(55+618)/2)<1e-7);
 await fit.click();
 // Vertical and translated negative-diagonal arrays must also fit.
 await page.locator('#ls-cad-array-dx').fill('0');await page.locator('#ls-cad-array-dy').fill('400[um]');await page.locator('#ls-cad-build-selected').click();await fit.click();await inside();
 await page.locator('#ls-cad-array-dx').fill('-400[um]');await page.locator('#ls-cad-array-dy').fill('-250[um]');await page.locator('#ls-cad-build-selected').click();
 await page.locator('[data-cad-node="c2"]').click();await page.locator('#ls-cad-x').fill('5000[um]');await page.locator('#ls-cad-y').fill('-3000[um]');
 await page.locator('[data-cad-node="arr1"]').click();await page.locator('#ls-cad-build-selected').click();await fit.click();await inside();
 const pixels=async(entity,north=false)=>page.locator('#ls-cad-view [data-cad-entity="'+entity+'"] .ls-cad-rendered-surface circle').evaluate((c,north)=>{
  const p=c.ownerSVGElement.createSVGPoint();p.x=Number(c.getAttribute('cx'));p.y=Number(c.getAttribute('cy'))-(north?Number(c.getAttribute('r')):0);const q=p.matrixTransform(c.ownerSVGElement.getScreenCTM());return {x:q.x,y:q.y};
 },north);
 // Verify inverse screen/world mapping after nonzero camera centering.
 await page.locator('summary[data-action="waveMaterials"]').click({button:'right'});await page.locator('[data-material-command="blank"]').click();await page.locator('#ls-material-select-mode').selectOption('manual');
 let p=await pixels('arr1/c2/0/0');await page.mouse.click(p.x,p.y);
 assert.deepEqual(await page.evaluate(()=>window.LSWaveMaterialsEditor.list()[0].domainIds),[1]);
 await fit.click();assert.equal(await page.locator('#ls-material-domain-list input:checked').count(),1);
 await page.locator('[data-tab="definitions"]').click();await page.locator('#ls-ribbon-tools [data-action="definitionsAdjacent"]').click();
 p=await pixels('arr1/c2/0/0');await page.mouse.click(p.x,p.y);
 assert.deepEqual(await page.evaluate(()=>window.LSWaveAdjacentSelection.list()[0].selectedByLevel.Domain),['arr1/c2/0/0/disk']);
 await page.locator('#ls-adjacent-level').selectOption('Boundary');await fit.click();
 p=await pixels('arr1/c2/5/0',true);await page.mouse.click(p.x,p.y);
 assert.deepEqual(await page.evaluate(()=>window.LSWaveAdjacentSelection.list()[0].selectedByLevel.Boundary),['arr1/c2/5/0/outer-boundary']);
 await page.locator('#ls-ribbon-tools [data-action="definitionsPML"]').click();await fit.click();
 p=await pixels('arr1/c2/0/0');await page.mouse.click(p.x,p.y);
 assert.deepEqual(await page.evaluate(()=>window.LSWavePMLService.selected().map(d=>d.key)),['arr1/c2/0/0/disk']);
 await page.screenshot({path:'/tmp/comsol_zoom_extents_pml_pick.png'});
 await page.locator('[data-cad-node="c2"]').click();await page.locator('#ls-cad-reset').click();
 const reset=await circles();for(const c of reset){assert.ok(Math.abs(c.x-(55+618)/2)<1e-7);assert.ok(Math.abs(c.y-(31+452)/2)<1e-7);}
 assert.deepEqual(errors,[]);
 console.log('PASS: Conversions submenu, cropped-to-six-disk Zoom Extents, aspect/world centering, stable zoom/reset, overview, real Material/Adjacent/Boundary/PML picking after off-origin fit');
}finally{await browser.close();}
