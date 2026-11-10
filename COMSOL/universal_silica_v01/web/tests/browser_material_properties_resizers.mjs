/** COMSOL screenshot regression: Adjacent output, Material add/delete,
 * parameter-backed n_iso/ki_iso, property-context Add to Material and pane resize.
 */
import assert from "node:assert/strict";
import {createRequire} from "node:module";
const {chromium}=createRequire(import.meta.url)("playwright");
const browser=await chromium.launch({headless:true,args:["--no-sandbox","--disable-dev-shm-usage"]});
try{
 const page=await browser.newPage({viewport:{width:1650,height:980}});
 const errors=[];page.on("pageerror",e=>errors.push(e.message));
 const base=process.env.TEST_BASE_URL||"http://127.0.0.1:8765/COMSOL/universal_silica_v01/web/";
 await page.goto(base,{waitUntil:"load",timeout:90000});
 await page.locator('[data-dimension="2D"]').click();
 await page.locator('[data-interface="ewfd"]').click();
 await page.locator('[data-wizard="add"]').click();
 await page.locator('[data-wizard="select-added"]').click();
 await page.locator('[data-wizard="study"]').click();
 await page.locator('[data-study="mode"]').first().click();
 await page.locator('[data-wizard="done"]').click();
 await page.locator('summary[data-action="geometry"]').click({button:"right"});
 await page.locator('#ls-cad-menu > ul > li > [data-cmd="circle"]').click();
 await page.locator("#ls-cad-radius").fill("rb+wlr");
 await page.locator("#ls-cad-layers-section summary").click();
 await page.locator("#ls-cad-layer").fill("wlr");
 await page.locator("#ls-cad-build-selected").click();
 // 1. Adjacent Output Entities dropdown has three choices.
 await page.locator('[data-tab="definitions"]').click();
 await page.locator('#ls-ribbon-tools [data-action="definitionsAdjacent"]').click();
 assert.deepEqual(await page.locator("#ls-adjacent-output option").allTextContents(),["Adjacent domains","Adjacent boundaries","Adjacent points"]);
 assert.equal(await page.locator("#ls-adjacent-output").inputValue(),"Adjacent boundaries");
 await page.locator("#ls-adjacent-output").selectOption("Adjacent points");
 assert.equal((await page.evaluate(()=>window.LSWaveAdjacentSelection.list()[0].output)),"Adjacent points");
 await page.locator('[data-adjacent-node="adj1"]').click();
 assert.equal(await page.locator("#ls-adjacent-output").inputValue(),"Adjacent points");
 await page.screenshot({path:"/tmp/comsol_adjacent_output_levels.png"});
 // 2. Materials +/duplicate/delete via both tree node and Settings buttons.
 await page.locator('summary[data-action="waveMaterials"]').click({button:"right"});
 await page.locator('#ls-materials-context [data-material-command="blank"]').click();
 assert.equal((await page.evaluate(()=>window.LSWaveMaterialsEditor.list())).length,1);
 await page.locator('#ls-materials-editor [data-material-command="blank"]').click();
 assert.equal((await page.evaluate(()=>window.LSWaveMaterialsEditor.list())).length,2);
 assert.ok(await page.locator('[data-ls-blank-material="mat2"]').isVisible());
 await page.locator('[data-material-command="duplicate"]').click();
 assert.equal((await page.evaluate(()=>window.LSWaveMaterialsEditor.list())).length,3);
 const mat3=page.locator('#ls-cad-material-list [data-ls-blank-material="mat3"]');
 await mat3.click();
 assert.equal(await page.locator("#ls-material-label").inputValue(),"Material 3");
 await page.locator('[data-material-command="delete"]').click();
 assert.ok(await page.locator("#ls-material-delete-dialog").isVisible());
 await page.locator('#ls-material-delete-dialog [value="no"]').click();
 assert.equal((await page.evaluate(()=>window.LSWaveMaterialsEditor.list())).length,3);
 await page.locator('[data-material-command="delete"]').click();
 await page.locator('#ls-material-delete-dialog [value="yes"]').click();
 assert.equal((await page.evaluate(()=>window.LSWaveMaterialsEditor.list())).length,2);
 assert.equal(await page.locator('#ls-cad-material-list [data-ls-blank-material="mat3"]').count(),0);
 // 3. Refractive index expression applies numerically to material's domain data.
 await page.locator('[data-ls-blank-material="mat2"]').click();
 const real=page.locator("#ls-material-real"),imag=page.locator("#ls-material-imag");
 await real.fill("nbg");
 await imag.fill("0");
 const resolved=await page.evaluate(()=>window.LSWaveMaterialsEditor.resolve(window.LSWaveMaterialsEditor.list()[1]));
 assert.equal(resolved.valid,true);
 assert.ok(Math.abs(resolved.real-1.45)<1e-12);
 assert.equal(resolved.imag,0);
 assert.ok((await page.locator("#ls-material-evaluated").textContent()).includes("1.45"));
 assert.ok(await page.locator("#ls-material-evaluated").evaluate(el=>el.classList.contains("is-valid")));
 const assigned=await page.evaluate(()=>window.LSWaveMaterialsEditor.assignments());
 assert.ok(assigned.length>0);
 assert.equal(assigned[0].real,1.45);
 await real.fill("not_a_parameter");
 assert.ok(await page.locator("#ls-material-evaluated").evaluate(el=>el.classList.contains("is-invalid")));
 assert.equal((await page.evaluate(()=>window.LSWaveMaterialsEditor.resolve(window.LSWaveMaterialsEditor.list()[1]))).valid,false);
 await real.fill("nbg");
 // Domain-specific values: one material for the core, one for cladding, no cross-over.
 await page.locator("#ls-material-select-mode").selectOption("manual");
 await page.locator('#ls-material-domain-list [data-material-domain="1"]').check();
 assert.ok((await page.evaluate(()=>window.LSWaveMaterialsEditor.assignments())).some(a=>a.domainId===1&&a.real===1.45));
 // 4. Properties categories and Refractive Index context menu.
 await page.locator("#ls-material-properties-section > summary").click();
 const group=page.locator('.ls-material-property-group').filter({has:page.locator('summary:has-text("Electromagnetic Models")')});
 assert.ok(await group.isVisible());
 assert.ok(await group.locator('[data-ls-property="Refractive index"]').isVisible());
 await group.locator('[data-ls-property="Refractive index"]').click({button:"right"});
 assert.ok(await page.locator("#ls-material-property-context").isVisible());
 await page.locator("#ls-material-property-add").click();
 assert.equal(await page.locator("#ls-material-property-context").isHidden(),true);
 assert.ok((await page.evaluate(()=>window.LSWaveMaterialsEditor.list()[1].properties)).includes("refractive-index"));
 assert.equal(await page.locator("#ls-material-real").inputValue(),"nbg");
 // Selecting Refractive index now expands a pair of individually editable
 // child properties, while preserving the original right-click Add to Material.
 const propertyNode=page.locator(".ls-material-refractive-node");
 assert.equal(await propertyNode.isVisible(),true);
 assert.equal(await propertyNode.evaluate(el=>el.open),true);
 const realProperty=propertyNode.locator('[data-ls-refractive-focus="real"]');
 const imaginaryProperty=propertyNode.locator('[data-ls-refractive-focus="imag"]');
 assert.equal(await realProperty.count(),1);
 assert.equal(await imaginaryProperty.count(),1);
 await realProperty.click();
 assert.equal(await page.evaluate(()=>document.activeElement?.id),"ls-material-real");
 await page.locator("#ls-material-real").fill("nbg+0.005");
 assert.ok(Math.abs((await page.evaluate(()=>window.LSWaveMaterialsEditor.resolve(window.LSWaveMaterialsEditor.list()[1]))).real-1.455)<1e-12);
 await imaginaryProperty.click();
 assert.equal(await page.evaluate(()=>document.activeElement?.id),"ls-material-imag");
 await page.locator("#ls-material-imag").fill("0.0001");
 assert.ok(Math.abs((await page.evaluate(()=>window.LSWaveMaterialsEditor.resolve(window.LSWaveMaterialsEditor.list()[1]))).imag-0.0001)<1e-12);
 // Clicking the parent summary closes and reopens the same two child items.
 await propertyNode.locator('summary[data-ls-property="Refractive index"]').click();
 assert.equal(await propertyNode.evaluate(el=>el.open),false);
 await propertyNode.locator('summary[data-ls-property="Refractive index"]').click();
 assert.equal(await propertyNode.evaluate(el=>el.open),true);
 assert.equal(await realProperty.isVisible(),true);
 assert.equal(await imaginaryProperty.isVisible(),true);
 await realProperty.click();
 await page.locator("#ls-material-real").fill("nbg");
 await imaginaryProperty.click();
 await page.locator("#ls-material-imag").fill("0");
 await page.screenshot({path:"/tmp/comsol_material_properties_resolved.png"});
 // 5. Model Builder and Settings separators can be dragged; graphics width
 // adapts without losing the Material editing context.
 const model=page.locator(".ls-model-tree"),sidebar=page.locator(".layout > .sidebar");
 const initialModel=(await model.boundingBox()).width;
 const h1=page.locator("#ls-split-model");
 await h1.waitFor({state:"visible"});
 let b=await h1.boundingBox();
 await page.mouse.move(b.x+b.width/2,b.y+80);
 await page.mouse.down();
 await page.mouse.move(b.x+b.width/2+75,b.y+80,{steps:6});
 await page.mouse.up();
 assert.ok((await model.boundingBox()).width>initialModel+45,"Model Builder drag increases width");
 const initialSettings=(await sidebar.boundingBox()).width;
 const h2=page.locator("#ls-split-settings");
 b=await h2.boundingBox();
 await page.mouse.move(b.x+b.width/2,b.y+90);
 await page.mouse.down();
 await page.mouse.move(b.x+b.width/2+80,b.y+90,{steps:6});
 await page.mouse.up();
 assert.ok((await sidebar.boundingBox()).width>initialSettings+40,"Settings drag increases width");
 assert.ok(await page.locator("#ls-materials-editor").isVisible());
 assert.equal(await page.locator("#ls-material-real").inputValue(),"nbg");
 await page.screenshot({path:"/tmp/comsol_resized_material_settings.png"});
 assert.deepEqual(errors,[],"No browser JavaScript errors");
 console.log("PASS: Adjacent 3 output levels, Material add/duplicate/delete, Parameters nbg evaluation, invalid value feedback, right-click Add to Material and draggable pane widths");
}finally{await browser.close();}
