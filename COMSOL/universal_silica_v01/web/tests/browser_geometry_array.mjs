/** User lecture screenshots: Transforms / Array Settings / Boolean menu. */
import assert from 'node:assert/strict';
import {createRequire} from 'node:module';
const {chromium}=createRequire(import.meta.url)('playwright');
const browser=await chromium.launch({headless:true,args:['--no-sandbox','--disable-dev-shm-usage']});
try{
 const page=await browser.newPage({viewport:{width:1500,height:940}});
 const errors=[];page.on('pageerror',e=>errors.push(e.message));
 await page.goto(process.env.TEST_BASE_URL||'http://127.0.0.1:8765/COMSOL/universal_silica_v01/web/',{waitUntil:'load',timeout:90000});
 await page.locator('[data-wizard="cancel"]').click();
 await page.locator('[data-cad-node="c2"]').click();
 await page.locator('#ls-cad-radius').fill('1[um]');await page.locator('#ls-cad-x').fill('6[um]');
 await page.locator('#ls-cad-layers-section summary').click();await page.locator('#ls-cad-layer').fill('0.2[um]');
 const openMenu=()=>page.locator('summary[data-action="geometry"]').click({button:'right'});
 const group=name=>page.locator('#ls-cad-menu > ul > li.has-children').filter({has:page.locator(':scope > button').filter({hasText:name})});
 await openMenu();
 const booleans=group('Booleans and Partitions');await booleans.locator(':scope > button').hover();
 assert.deepEqual(await booleans.locator(':scope > ul > li > button').allTextContents(),['Union','Intersection','Difference','Compose','Partition Objects','Partition Domains','Partition Edges']);
 assert.equal(await booleans.locator(':scope > ul button:disabled').count(),6,'Only Union is implemented for the 2D geometry preview.');
  assert.equal(await booleans.locator('[data-cmd="union"]').isDisabled(),false);
 assert.ok(await booleans.evaluate(el=>[...el.parentElement.children].indexOf(el)<[...el.parentElement.children].findIndex(x=>x.querySelector(':scope > button')?.textContent==='Transforms▸')));
 await page.screenshot({path:'/tmp/comsol_geometry_boolean_partition_menu.png'});
 const transforms=group('Transforms');await transforms.locator(':scope > button').hover();
 assert.deepEqual(await transforms.locator(':scope > ul > li > button').allTextContents(),['Array','Copy','Mirror','Move','Rotate','Scale']);
 assert.equal(await transforms.locator('[data-cmd="array"]').isDisabled(),false);
 await page.screenshot({path:'/tmp/comsol_geometry_transforms_array_menu.png'});
 await transforms.locator('[data-cmd="array"]').click();
 assert.equal(await page.locator('#ls-cad-subtitle').textContent(),'Array');
 assert.equal(await page.locator('#ls-cad-label').inputValue(),'Array 1');
 assert.ok(await page.locator('[data-cad-node="arr1"]').isVisible());
 assert.equal(await page.locator('#ls-cad-array-type').inputValue(),'Rectangular');
 assert.deepEqual(await page.locator('#ls-cad-array-type option').allTextContents(),['Rectangular','Linear']);
 for(const field of ['nx','ny'])assert.equal(await page.locator('#ls-cad-array-'+field).inputValue(),'1');
 for(const field of ['dx','dy'])assert.equal(await page.locator('#ls-cad-array-'+field).inputValue(),'0');
 assert.ok(await page.locator('#ls-cad-array-nx-row').isVisible());assert.ok(await page.locator('#ls-cad-array-ny-row').isVisible());
 assert.ok(await page.locator('#ls-cad-array-size-row').isHidden());
 assert.ok(await page.locator('#ls-cad-rotate-angle-section').isHidden());assert.ok(await page.locator('#ls-cad-keep-row').isHidden());
 assert.equal(await page.locator('#ls-cad-activate-selection').getAttribute('aria-pressed'),'true');
 assert.equal(await page.locator('#ls-cad-inputs input:checked').count(),0,'New Array input starts empty as in lecture.');
 await page.screenshot({path:'/tmp/comsol_array_rectangular_settings.png'});
 const input=page.locator('#ls-cad-inputs [data-cad-input-ref="c2"]');
 await input.check();assert.equal(await page.locator('#ls-cad-view [data-cad-entity="c2"].is-input-selected').count(),1);
 await page.locator('#ls-cad-array-nx').fill('2');await page.locator('#ls-cad-array-ny').fill('3');
 await page.locator('#ls-cad-array-dx').fill('8');await page.locator('#ls-cad-array-dy').fill('6[um]');
 await page.locator('#ls-cad-build-selected').click();
 assert.equal(await page.locator('#ls-cad-activate-selection').getAttribute('aria-pressed'),'false');
 assert.equal(await page.locator('#ls-cad-view [data-cad-shape="arr1"]').count(),6);
 assert.equal(await page.locator('#ls-cad-view [data-cad-shape="c2"]').count(),0,'Array includes origin and consumes its selected inputs.');
 const output=async owner=>page.evaluate(owner=>window.LSWaveCADService.entities().filter(e=>e.owner===owner).map(e=>({key:e.entityKey,x:e.x,y:e.y,r:e.r,layers:e.layers})),owner);
 const close=(a,b)=>assert.ok(Math.abs(a-b)<1e-12,`${a} != ${b}`);
 let copies=await output('arr1');assert.equal(copies.length,6);
 for(const [i,c] of copies.entries()){close(c.x,(6+Math.floor(i/3)*8)*1e-6);close(c.y,(i%3)*6e-6);close(c.r,1e-6);close(c.layers[0],.2e-6);}
 await page.screenshot({path:'/tmp/comsol_array_rectangular_six_outputs.png'});
 await page.locator('#ls-cad-array-type').selectOption('Linear');
 assert.ok(await page.locator('#ls-cad-array-nx-row').isHidden());assert.ok(await page.locator('#ls-cad-array-ny-row').isHidden());
 assert.ok(await page.locator('#ls-cad-array-size-row').isVisible());assert.equal(await page.locator('#ls-cad-array-size').inputValue(),'1');
 await page.locator('#ls-cad-array-size').fill('4');await page.locator('#ls-cad-array-dx').fill('7[um]');await page.locator('#ls-cad-array-dy').fill('-3');
 await page.locator('#ls-cad-build-all').click();copies=await output('arr1');assert.equal(copies.length,4);
 for(const [i,c] of copies.entries()){close(c.x,(6+i*7)*1e-6);close(c.y,-i*3e-6);}
 await page.screenshot({path:'/tmp/comsol_array_linear_settings_outputs.png'});
 // Invalid counts are rejected without allocating uncontrolled geometry.
 for(const size of ['0','1.5','451']){
  await page.locator('#ls-cad-array-size').fill(size);await page.locator('#ls-cad-build-selected').click();
  assert.ok((await page.locator('#ls-cad-status').textContent()).startsWith('Build failed:'));
 }
 await page.locator('#ls-cad-array-size').fill('4');await page.locator('#ls-cad-build-selected').click();
 const before=await page.locator('#outer_radius_um').inputValue();await page.locator('#ls-cad-apply').click();
 assert.ok((await page.locator('#ls-cad-status').textContent()).includes('Disable Array'));
 assert.equal(await page.locator('#outer_radius_um').inputValue(),before,'Preview Array must not silently alter the radial FEM.');
 // Rectangular dimensions survive changing type, while displacement remains shared.
 await page.locator('#ls-cad-array-type').selectOption('Rectangular');
 assert.equal(await page.locator('#ls-cad-array-nx').inputValue(),'2');assert.equal(await page.locator('#ls-cad-array-ny').inputValue(),'3');
 await page.locator('#ls-cad-build-selected').click();assert.equal((await output('arr1')).length,6);
 // Graphics input selection always shows preceding source geometry, even after consumption.
 await page.locator('#ls-cad-activate-selection').click();
 const source=page.locator('#ls-cad-view [data-cad-entity="c2"] .ls-cad-rendered-surface circle');
 await source.click({force:true});assert.equal(await input.isChecked(),false);
 await source.click({force:true});assert.ok(await input.isChecked());
 // Multiple independent inputs each produce nx*ny outputs; not one whole feature group.
 await page.locator('#ls-cad-inputs [data-cad-input-ref="c1"]').check();await page.locator('#ls-cad-build-selected').click();assert.equal((await output('arr1')).length,12);
 await page.locator('#ls-cad-inputs [data-cad-input-ref="c1"]').uncheck();await page.locator('#ls-cad-build-selected').click();assert.equal((await output('arr1')).length,6);
 // Existing named cumulative selection and Materials can reference the Array feature.
 await page.locator('#ls-cad-cumulative-new').click();await page.locator('#ls-cad-cumulative-name').fill('Array outputs');await page.locator('#ls-cad-cumulative-ok').click();
 await page.locator('#ls-cad-cumulative-list [data-cad-cumulative-id="csel1"]').waitFor({state:'visible'});
 assert.deepEqual(await page.evaluate(()=>window.LSWaveCADService.cumulativeSelections()[0].featureIds),['arr1']);
 await page.locator('summary[data-action="waveMaterials"]').click({button:'right'});await page.locator('[data-material-command="blank"]').click();
 await page.locator('#ls-material-select-mode').selectOption('cumulative:csel1');
 assert.equal(await page.locator('#ls-material-domain-list input:checked').count(),30,'Six layered array outputs each have five independent preview domains.');
 await page.locator('#ls-material-select-mode').selectOption('manual');
 await page.locator('#ls-material-domain-list input[data-material-domain="2"]').check();assert.equal(await page.locator('#ls-material-domain-list input:checked').count(),1);
 await page.locator('[data-cad-node="arr1"]').click();assert.equal(await page.locator('#ls-cad-array-nx').inputValue(),'2');
 // Array of one specific previous derived object, with mixed displacement units.
 await openMenu();await group('Transforms').locator(':scope > button').hover();await group('Transforms').locator('[data-cmd="array"]').click();
 const selectedKey='arr1/c2/1/2';await page.locator('#ls-cad-inputs [data-cad-input-ref="'+selectedKey+'"]').check();
 await page.locator('#ls-cad-array-type').selectOption('Linear');await page.locator('#ls-cad-array-size').fill('3');
 await page.locator('#ls-cad-array-dx').fill('1[mm]');await page.locator('#ls-cad-array-dy').fill('0');await page.locator('#ls-cad-build-selected').click();
 copies=await output('arr2');assert.equal(copies.length,3);assert.equal((await output('arr1')).length,5);
 for(const [i,c] of copies.entries())close(c.x,(6+7)*1e-6+i*.001);
 await page.locator('#ls-cad-array-size').fill('4');await page.locator('#ls-cad-undo').click();assert.equal(await page.locator('#ls-cad-array-size').inputValue(),'3');
 await page.locator('#ls-cad-redo').click();assert.equal(await page.locator('#ls-cad-array-size').inputValue(),'4');
 // Array configurations survive export as real feature data.
 const downloaded=page.waitForEvent('download');await page.locator('#ls-cad-export').click();const download=await downloaded;
 const stream=await download.createReadStream();let body='';for await(const chunk of stream)body+=chunk.toString();
 const doc=JSON.parse(body);assert.equal(doc.geometry.find(n=>n.id==='arr2').arrayType,'Linear');assert.deepEqual(doc.geometry.find(n=>n.id==='arr2').inputRefs,[selectedKey]);
 assert.deepEqual(errors,[]);
 console.log('PASS: screenshot menus, Array rectangular/linear settings, source Graphics picking, exact coordinates/layers/units, limits, cumulative/material integration, nesting, undo/redo/export and FEM guard');
}finally{await browser.close();}
