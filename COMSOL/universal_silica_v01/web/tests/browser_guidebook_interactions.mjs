/** Real UI regression for common COMSOL 5.5 guidebook operations, pp. 176–178. */
import assert from 'node:assert/strict';
import {createRequire} from 'node:module';
const {chromium}=createRequire(import.meta.url)('playwright');
const browser=await chromium.launch({headless:true,args:['--no-sandbox','--disable-dev-shm-usage']});
try{
 const page=await browser.newPage({viewport:{width:1500,height:940}});
 const errors=[];page.on('pageerror',e=>errors.push(e.message));
 await page.goto(process.env.TEST_BASE_URL||'http://127.0.0.1:8765/COMSOL/universal_silica_v01/web/',{waitUntil:'load',timeout:90000});
 await page.locator('[data-wizard="cancel"]').click();
 const cad=id=>page.locator('[data-cad-node="'+id+'"]');
 await cad('c1').click();await cad('c2').click();
 await page.keyboard.press('Alt+ArrowLeft');assert.ok(await cad('c1').evaluate(el=>el.classList.contains('active')));
 await page.keyboard.press('Alt+ArrowRight');assert.ok(await cad('c2').evaluate(el=>el.classList.contains('active')));
 // Arrow selection survives the editor rebuilding its buttons.
 await page.keyboard.press('ArrowUp');assert.ok(await cad('c1').evaluate(el=>el.classList.contains('active')));
 await page.keyboard.press('ArrowDown');assert.ok(await cad('c2').evaluate(el=>el.classList.contains('active')));
 await page.keyboard.press('F3');await page.keyboard.press('F3');
 assert.ok(await cad('c2').evaluate(el=>el.classList.contains('ls-cad-node-disabled')),'F3 must remain Disable when repeated');
 await page.keyboard.press('F4');await page.keyboard.press('F4');
 assert.equal(await cad('c2').evaluate(el=>el.classList.contains('ls-cad-node-disabled')),false,'F4 must remain Enable when repeated');
 const geometry=page.locator('summary[data-action="geometry"]');
 await geometry.focus();await page.keyboard.press('ArrowLeft');assert.equal(await geometry.evaluate(el=>el.parentElement.open),false);
 await page.keyboard.press('ArrowRight');assert.equal(await geometry.evaluate(el=>el.parentElement.open),true);
 await cad('c2').click();await cad('c2').focus();
 await page.keyboard.press('Control+ArrowUp');assert.equal(await page.locator('#ls-geometry-node-list button').first().getAttribute('data-cad-node'),'c2');
 await page.keyboard.press('Control+ArrowDown');assert.equal(await page.locator('#ls-geometry-node-list button').first().getAttribute('data-cad-node'),'c1');
 // F1 is local contextual help; Escape must not close the current editor.
 await page.keyboard.press('F1');assert.ok(await page.locator('#ls-guide-help').isVisible());
 assert.equal(await page.locator('#ls-guide-help h3').textContent(),'Geometry');
 await page.keyboard.press('Escape');assert.ok(await page.locator('#ls-cad-pane').isVisible());assert.ok(await page.locator('#ls-guide-help').isHidden());
 await cad('c1').click();
 await page.locator('#ls-cad-radius').fill('rb+wlr');
 await page.locator('#ls-cad-layers-section summary').click();await page.locator('#ls-cad-layer').fill('wlr');await page.locator('#ls-cad-build-selected').click();
 // Material: no implicit conversion from All/named modes to Manual.
 await page.locator('summary[data-action="waveMaterials"]').click({button:'right'});
 await page.locator('[data-material-command="blank"]').click();
 assert.equal(await page.locator('#ls-material-select-all').isDisabled(),true);
 await page.locator('#ls-material-select-mode').selectOption('manual');
 const count=await page.locator('#ls-material-domain-list input').count();assert.ok(count>=5);
 await page.locator('#ls-material-select-all').click();assert.equal(await page.locator('#ls-material-domain-list input:checked').count(),count);
 await page.locator('#ls-cad-view').focus();await page.keyboard.press('Control+d');assert.equal(await page.locator('#ls-material-domain-list input:checked').count(),0);
 await page.keyboard.press('Control+a');assert.equal(await page.locator('#ls-material-domain-list input:checked').count(),count);
 const m1=page.locator('#ls-material-domain-list input').first();await m1.focus();await page.keyboard.press('Delete');
 assert.equal(await page.locator('#ls-material-domain-list input:checked').count(),count-1);
 await page.locator('#ls-material-label').focus();await page.keyboard.press('Control+a');
 assert.equal(await page.locator('#ls-material-domain-list input:checked').count(),count-1,'Ctrl+A in text must select text, not domains');
 await page.keyboard.press('F1');assert.equal(await page.locator('#ls-guide-help h3').textContent(),'Materials · Selection');
 await page.screenshot({path:'/tmp/comsol_guidebook_material_help.png'});await page.keyboard.press('Escape');
 await page.locator('#ls-material-select-clear').click();assert.equal(await page.locator('#ls-material-domain-list input:checked').count(),0);
 // Adjacent: each input level keeps an independent selection.
 await page.locator('[data-tab="definitions"]').click();await page.locator('#ls-ribbon-tools [data-action="definitionsAdjacent"]').click();
 await page.locator('#ls-cad-view').focus();await page.keyboard.press('Control+a');
 const domainCount=(await page.evaluate(()=>window.LSWaveAdjacentSelection.list()[0].selectedByLevel.Domain)).length;assert.equal(domainCount,count);
 await page.locator('#ls-adjacent-level').selectOption('Boundary');
 await page.locator('#ls-adjacent-all').click();assert.ok((await page.evaluate(()=>window.LSWaveAdjacentSelection.list()[0].selectedByLevel.Boundary)).length>0);
 await page.locator('#ls-cad-view').focus();await page.keyboard.press('Control+d');
 assert.equal((await page.evaluate(()=>window.LSWaveAdjacentSelection.list()[0].selectedByLevel.Boundary)).length,0);
 await page.locator('#ls-adjacent-level').selectOption('Domain');assert.equal(await page.locator('#ls-adjacent-list input:checked').count(),domainCount);
 await page.locator('#ls-adjacent-list input').first().focus();await page.keyboard.press('Delete');assert.equal(await page.locator('#ls-adjacent-list input:checked').count(),domainCount-1);
 // PML: preserves its settings and honors All/Manual selection modes.
 await page.locator('#ls-ribbon-tools [data-action="definitionsPML"]').click();
 await page.locator('#ls-pml-type').selectOption('Cylindrical');await page.locator('#ls-pml-all').click();
 assert.equal(await page.locator('#ls-pml-domain-list input:checked').count(),count);
 await page.locator('#ls-cad-view').focus();await page.keyboard.press('Control+d');assert.equal(await page.locator('#ls-pml-domain-list input:checked').count(),0);
 await page.keyboard.press('Control+a');assert.equal(await page.locator('#ls-pml-domain-list input:checked').count(),count);
 await page.locator('#ls-pml-domain-list input').first().focus();await page.keyboard.press('Delete');assert.equal(await page.locator('#ls-pml-domain-list input:checked').count(),count-1);
 await page.locator('#ls-pml-selection').selectOption('all');
 await page.locator('#ls-cad-view').focus();await page.keyboard.press('Control+d');assert.equal(await page.locator('#ls-pml-domain-list input:checked').count(),count,'All mode must not silently become Manual');
 await page.locator('#ls-pml-selection').selectOption('manual');assert.equal(await page.locator('#ls-pml-domain-list input:checked').count(),count-1);
 await page.keyboard.press('F1');assert.equal(await page.locator('#ls-guide-help h3').textContent(),'Definitions · Perfectly Matched Layer');
 // Help follows node switches without closing the help pane or losing settings.
 await cad('c1').click();assert.equal(await page.locator('#ls-guide-help h3').textContent(),'Geometry');
 await page.locator('[data-pml-node="pml1"]').click();assert.equal(await page.locator('#ls-pml-type').inputValue(),'Cylindrical');
 await page.screenshot({path:'/tmp/comsol_guidebook_pml_help.png'});await page.keyboard.press('Escape');
 assert.ok(await page.locator('#ls-pml-editor').isVisible());
 assert.deepEqual(errors,[]);
 console.log('PASS: guidebook tree navigation/history, F3/F4, node reorder, scoped All/Clear/Delete, F1 help, existing Materials/Adjacent/PML state');
}finally{await browser.close();}
