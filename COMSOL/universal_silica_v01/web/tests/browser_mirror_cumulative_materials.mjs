/** End-to-end GUI: individual rotated objects → Mirror, cumulative selection and Blank Material. */
import assert from "node:assert/strict";
import {createRequire} from "node:module";
const {chromium}=createRequire(import.meta.url)("playwright");
const browser=await chromium.launch({headless:true,args:["--no-sandbox","--disable-dev-shm-usage"]});
try{
 const page=await browser.newPage({viewport:{width:1550,height:980}});
 const errors=[];page.on("pageerror",e=>errors.push(e.message));
 await page.goto(process.env.TEST_BASE_URL||"http://127.0.0.1:8765/COMSOL/universal_silica_v01/web/",{waitUntil:"load",timeout:90000});
 await page.locator('[data-wizard="cancel"]').click();
 // The legacy sample starts with two circles. Put c2 off-axis for six visible rotated copies.
 await page.locator('[data-cad-node="c2"]').click();
 await page.locator("#ls-cad-radius").fill("2.5[um]");
 await page.locator("#ls-cad-x").fill("6[um]");
 await page.locator('summary[data-action="geometry"]').click({button:"right"});
 await page.locator('#ls-cad-menu > ul > li.has-children > button').filter({hasText:"Transforms"}).hover();
 await page.locator('#ls-cad-menu li.has-children > ul [data-cmd="rotate"]').click();
 assert.equal(await page.locator("#ls-cad-subtitle").textContent(),"Rotate");
 assert.equal(await page.locator("#ls-cad-angle").inputValue(),"range(0,60,300)");
 assert.equal(await page.locator('#ls-cad-view [data-cad-shape="rot1"]').count(),6);
 assert.equal(await page.locator('#ls-cad-view [data-cad-entity^="rot1/c2/"]').count(),6);
 await page.screenshot({path:"/tmp/comsol_six_rotated_circles.png"});
 // Cumulative Selection: New -> OK -> Model Builder node.
 await page.locator("#ls-cad-cumulative-new").click();
 assert.ok(await page.locator("#ls-cad-cumulative-dialog").isVisible());
 await page.locator("#ls-cad-cumulative-name").fill("Hole rotation outputs");
 await page.locator("#ls-cad-cumulative-ok").click();
 const cumulative=page.locator('#ls-cad-cumulative-list [data-cad-cumulative-id="csel1"]');
 await cumulative.waitFor({state:"visible",timeout:5000}).catch(async error=>{
  console.error("Cumulative create diagnostics",await page.evaluate(()=>({
   dialogReturn:document.querySelector("#ls-cad-cumulative-dialog")?.returnValue,
   dialogOpen:document.querySelector("#ls-cad-cumulative-dialog")?.open,
   groupHidden:document.querySelector("#ls-cad-cumulative-group")?.hidden,
   groupHTML:document.querySelector("#ls-cad-cumulative-group")?.outerHTML?.slice(0,900),
   status:document.querySelector("#ls-cad-status")?.textContent,
   contribution:document.querySelector("#ls-cad-contribute")?.value
  })));
  throw error;
 });
 assert.equal(await page.locator("#ls-cad-contribute").inputValue(),"csel1");
 await cumulative.click();
 assert.equal(await page.locator("#ls-cad-subtitle").textContent(),"Cumulative Selection");
 assert.ok((await page.locator("#ls-cad-cumulative-summary").textContent()).includes("Rotate 1"));
 // Mirror adds a new feature, but starts without selected input objects.
 await page.locator('summary[data-action="geometry"]').click({button:"right"});
 await page.locator('#ls-cad-menu > ul > li.has-children > button').filter({hasText:"Transforms"}).hover();
 await page.locator('#ls-cad-menu li.has-children > ul [data-cmd="mirror"]').click();
 assert.equal(await page.locator("#ls-cad-subtitle").textContent(),"Mirror");
 assert.equal(await page.locator('#ls-cad-view [data-cad-shape="mir1"]').count(),0);
 await page.locator("#ls-cad-activate-selection").click();
 assert.equal(await page.locator("#ls-cad-activate-selection").getAttribute("aria-pressed"),"true");
 const one='rot1/c2/1',two='rot1/c2/2';
 await page.locator('#ls-cad-view [data-cad-entity="'+one+'"] .ls-cad-rendered-surface circle').dispatchEvent("click");
 await page.locator('#ls-cad-view [data-cad-entity="'+two+'"] .ls-cad-rendered-surface circle').dispatchEvent("click");
 assert.ok(await page.locator('#ls-cad-inputs [data-cad-input-ref="'+one+'"]').isChecked());
 assert.ok(await page.locator('#ls-cad-inputs [data-cad-input-ref="'+two+'"]').isChecked());
 assert.equal(await page.locator('#ls-cad-view [data-cad-shape="mir1"]').count(),2);
 const all=await page.locator('#ls-cad-view [data-cad-shape="mir1"]').evaluateAll(els=>els.map(el=>({entity:el.getAttribute("data-cad-entity"),key:el.dataset.cadSource})));
 assert.equal(new Set(all.map(e=>e.entity)).size,2,"Mirror results must have distinct entity IDs");
 await page.screenshot({path:"/tmp/comsol_selected_two_mirrors.png"});
 // Check feature can keep/remove input objects.
 await page.locator("#ls-cad-keep").check();
 assert.equal(await page.locator('#ls-cad-view [data-cad-shape="mir1"]').count(),2);
 await page.locator("#ls-cad-keep").uncheck();
 // Navigate Component 1 Materials / Blank Material.
 await page.locator('summary[data-action="waveMaterials"]').click({button:"right"});
 assert.ok(await page.locator("#ls-materials-context").isVisible());
 await page.locator('[data-material-command="blank"]').click();
 assert.ok(await page.locator("#ls-materials-editor").isVisible());
 assert.ok(await page.locator('#ls-cad-material-list [data-ls-blank-material="mat1"]').isVisible());
 assert.equal(await page.locator("#ls-material-label").inputValue(),"Material 1");
 assert.equal(await page.locator("#ls-material-select-mode").inputValue(),"all");
 const available=await page.locator("#ls-material-domain-list [data-material-domain]").count();
 assert.ok(available>=5,"Material overview should expose generated Geometry regions; found "+available);
 await page.locator("#ls-material-real").fill("1.45");
 await page.locator("#ls-material-imag").fill("0");
 await page.locator("#ls-material-select-mode").selectOption("manual");
 const first=page.locator('#ls-material-domain-list [data-material-domain="1"]');
 await first.check();
 assert.equal(await page.evaluate(()=>window.LSWaveMaterialsEditor.list()[0].mode),"manual");
 assert.ok((await page.evaluate(()=>window.LSWaveMaterialsEditor.list()[0].domainIds)).includes(1));
 await page.screenshot({path:"/tmp/comsol_blank_material_domains.png"});
 // Material 1: COMSOL Geometric Entity Level must expose Domain/Boundary/Point.
 const level=page.locator("#ls-material-entity-level"),selection=page.locator("#ls-material-select-mode");
 assert.deepEqual(await level.locator("option").allTextContents(),["Domain","Boundary","Point"]);
 assert.equal(await level.inputValue(),"Domain");
 // Named Cumulative Selection contributed by Rotate 1 must appear under Selection.
 const namedOption=selection.locator('optgroup[label="Cumulative Selections"] option[value="cumulative:csel1"]');
 assert.equal(await namedOption.textContent(),"Hole rotation outputs");
 await selection.selectOption("cumulative:csel1");
 assert.equal(await selection.inputValue(),"cumulative:csel1");
 assert.equal(await page.evaluate(()=>window.LSWaveMaterialsEditor.list()[0].mode),"cumulative");
 assert.equal(await page.evaluate(()=>window.LSWaveMaterialsEditor.list()[0].cumulativeId),"csel1");
 const domainSelected=await page.locator("#ls-material-domain-list input:checked").count();
 assert.ok(domainSelected>0&&domainSelected<available,"Cumulative Selection should include only Rotation 1 output domains");
 assert.equal(await page.locator("#ls-material-pick").isDisabled(),true);
 await page.screenshot({path:"/tmp/comsol_material_cumulative_selection.png"});
 // Named cumulative selection remains available and selects matching entities
 // for Boundary and Point levels; this is a geometry-only preview.
 await level.selectOption("Boundary");
 assert.equal(await level.inputValue(),"Boundary");
 assert.equal(await selection.inputValue(),"cumulative:csel1");
 const boundaryCount=await page.locator("#ls-material-domain-list input").count();
 const boundaryChecked=await page.locator("#ls-material-domain-list input:checked").count();
 assert.ok(boundaryCount>=available&&boundaryChecked>0&&boundaryChecked<boundaryCount,"Boundary ring selection should filter to the named Geometry feature");
 assert.ok((await page.locator("#ls-material-domain-list").textContent()).includes("boundary"));
 await page.screenshot({path:"/tmp/comsol_material_boundary_level.png"});
 await level.selectOption("Point");
 assert.equal(await level.inputValue(),"Point");
 assert.equal(await selection.inputValue(),"cumulative:csel1");
 const pointsChecked=await page.locator("#ls-material-domain-list input:checked").count();
 assert.ok(pointsChecked>0,"Point-level named selection should include rotation results");
 assert.ok((await page.locator("#ls-material-domain-list").textContent()).includes("center point"));
 await level.selectOption("Domain");
 // Return to Manual and confirm its previous domain checkbox edit was preserved.
 await selection.selectOption("manual");
 assert.equal(await page.locator('#ls-material-domain-list [data-material-domain="1"]').isChecked(),true);
 assert.equal(await page.evaluate(()=>window.LSWaveMaterialsEditor.list()[0].mode),"manual");
 assert.equal(await page.evaluate(()=>window.LSWaveMaterialsEditor.list()[0].level),"Domain");

 await page.locator('summary[data-action="waveMaterials"]').click();
 assert.ok((await page.locator("#ls-materials-editor").textContent()).includes("Entities needing a material"));
 assert.equal((await page.evaluate(()=>window.LSWaveMaterialsEditor.list()[0].real)),"1.45");
 assert.deepEqual(errors,[],"No uncaught browser errors");
 console.log("PASS: Rotate/Mirror individual entities, named Cumulative Selection in Material 1, Domain/Boundary/Point levels and preserved manual assignments");
}finally{await browser.close();}
