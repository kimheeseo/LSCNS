/** Stage 3: linked Geometry → Materials → Mesh → Study → Results, actual browser FEM. */
import assert from "node:assert/strict";
import {createRequire} from "node:module";
const {chromium}=createRequire(import.meta.url)("playwright");
const browser=await chromium.launch({headless:true,args:["--no-sandbox","--disable-dev-shm-usage"]});
try{
 const page=await browser.newPage({viewport:{width:1440,height:900}});
 const errors=[];page.on("pageerror",e=>errors.push(e.message));
 const url=process.env.TEST_BASE_URL||"http://127.0.0.1:8765/COMSOL/universal_silica_v01/web/";
 await page.goto(url,{waitUntil:"load",timeout:90000});
 await page.waitForSelector("#ls-linked-workflow",{state:"visible"});
 assert.equal(await page.locator("#ls-wf-geometry-state").textContent(),"Ready");
 assert.equal(await page.locator("#ls-wf-geometry-state").getAttribute("data-state"),"ready");
 assert.equal(await page.evaluate(()=>window.LSWaveCADService.status().domains.length),2);
 // Geometry changes must be applied before the linked study.
 await page.locator("#ls-linked-workflow [data-wf-step='geometry']").click();
 assert.equal(await page.locator("#ls-cad-root-settings").isVisible(),true);
 await page.locator('[data-cad-node="c2"]').click();
 await page.locator("#ls-cad-radius").fill("7[um]");
 await page.locator("#ls-cad-close").click();
 assert.equal(await page.locator("#ls-wf-geometry-state").textContent(),"Apply");
 await page.locator("#ls-wf-compute").click();
 assert.ok((await page.locator("#ls-wf-status").textContent()).includes("not applied"));
 await page.locator("#ls-wf-apply-geometry").click();
 assert.equal(await page.locator("#ls-wf-geometry-state").textContent(),"Ready");
 assert.equal(Number(await page.locator("#core_radius_um").inputValue()),7);
 // Domain click opens actual Core / Cladding material authoring.
 await page.locator("#ls-linked-workflow [data-wf-step='materials']").click();
 assert.equal(await page.locator("#ls-wave-settings").isVisible(),true);
 assert.ok((await page.locator(".ls-wave-domain-map").textContent()).includes("D1 · Core"));
 await page.locator('[data-wave-domain="core"]').last().click();
 assert.equal(await page.locator("#ls-wave-settings h3").textContent(),"Core Material");
 await page.locator("#ls-wave-delta").fill("0.0042");
 await page.locator('[data-wave-command="applyDelta"]').click();
 assert.equal(Number(await page.locator("#delta_n_core").inputValue()),0.0042);
 await page.locator('[data-wave-domain="cladding"]').last().click();
 assert.equal(await page.locator("#ls-wave-settings h3").textContent(),"Cladding Material");
 await page.locator("#ls-wave-clad-mode").selectOption("fixed");
 await page.locator("#ls-wave-clad-index").fill("1.4378");
 await page.locator('[data-wave-command="applyClad"]').click();
 assert.equal(await page.locator("#clad_material").inputValue(),"fixed");
 assert.equal(Number(await page.locator("#clad_index_override").inputValue()),1.4378);
 // Mesh element count is the true 1D radial FEM mesh input.
 await page.locator("#ls-linked-workflow [data-wf-step='mesh']").click();
 assert.equal(await page.locator("#ls-wave-elements").count(),1);
 await page.locator("#ls-wave-elements").fill("900");
 await page.locator('[data-wave-command="applyMesh"]').click();
 assert.equal(Number(await page.locator("#elements").inputValue()),900);
 // Study runs the live numerical solver, not a static preview.
 await page.locator("#ls-linked-workflow [data-wf-step='study']").click();
 assert.equal(await page.locator("#ls-wave-wavelength").count(),1);
 await page.locator("#ls-wave-wavelength").fill("1550");
 await page.locator('[data-wave-command="applyCompute"]').click();
 await page.waitForFunction(()=>document.querySelector("#ls-wf-result-state")?.textContent==="Current",{timeout:30000});
 const n=Number(await page.locator("#neff").textContent());
 assert.ok(n>1.4378&&n<1.442,"Mode Analysis must calculate physically plausible n_eff");
 await page.locator("#ls-linked-workflow [data-wf-step='results']").click();
 assert.ok((await page.locator("#ls-wave-settings").textContent()).includes("MFD"));
 assert.ok((await page.locator("#ls-wave-settings").textContent()).includes(n.toFixed(9)));
 // Material or mesh input change invalidates linked result.
 await page.locator("#ls-linked-workflow [data-wf-step='mesh']").click();
 await page.locator("#ls-wave-elements").fill("950");
 await page.locator('[data-wave-command="applyMesh"]').click();
 assert.equal(await page.locator("#ls-wf-result-state").textContent(),"Stale");
 await page.locator("#ls-wf-compute").click();
 await page.waitForFunction(()=>document.querySelector("#ls-wf-result-state")?.textContent==="Current",{timeout:30000});
 // Arbitrary 2D CAD shapes are not silently fed into scalar FEM.
 await page.locator("#ls-linked-workflow [data-wf-step='geometry']").click();
 await page.locator("#ls-cad-add-rotate").click();
 await page.locator("#ls-cad-close").click();
 assert.equal(await page.locator("#ls-wf-geometry-state").textContent(),"Blocked");
 await page.locator("#ls-wf-compute").click();
 assert.ok((await page.locator("#ls-wf-status").textContent()).includes("Unsupported geometry"));
 // The unsupported operation is still available for visual preview but not linked Solve.
 await page.locator("#ls-linked-workflow [data-wf-step='geometry']").click();
 await page.locator('[data-cad-node="rot1"]').click({button:"right"});
 await page.locator('#ls-cad-menu [data-cmd="toggle"]').click();
 await page.locator("#ls-cad-close").click();
 assert.equal(await page.locator("#ls-wf-geometry-state").textContent(),"Apply");
 await page.locator("#ls-wf-apply-geometry").click();
 assert.equal(await page.locator("#ls-wf-geometry-state").textContent(),"Ready");
 // Parameterized geometry: editing d updates the resolved Circle 2 radius, not just its text expression.
 await page.locator("#ls-linked-workflow [data-wf-step='geometry']").click();
 await page.locator('[data-cad-node="c2"]').click();
 await page.locator("#ls-cad-radius").fill("d/2");
 await page.locator("#ls-cad-close").click();
 await page.locator("#ls-wf-apply-geometry").click();
 await page.locator('.ls-model-tree [data-action="settings"]').click();
 await page.locator('#ls-parameter-rows [data-field="expression"]').first().fill("16[um]");
 assert.equal(await page.locator("#ls-wf-geometry-state").textContent(),"Apply");
 await page.locator("#ls-wf-compute").click();
 assert.ok((await page.locator("#ls-wf-status").textContent()).includes("not applied"));
 await page.locator("#ls-wf-apply-geometry").click();
 assert.equal(await page.locator("#ls-wf-geometry-state").textContent(),"Ready");
 assert.equal(Number(await page.locator("#core_radius_um").inputValue()),8);
 await page.screenshot({path:"/tmp/ls_comsol_linked_workflow.png"});
 assert.deepEqual(errors,[],"No browser exceptions");
 console.log("PASS: Stage 3 Geometry-to-domain Materials-to-Mesh-to-Study-to-Results real FEM, stale result detection, unsupported CAD guard");
}finally{await browser.close();}
