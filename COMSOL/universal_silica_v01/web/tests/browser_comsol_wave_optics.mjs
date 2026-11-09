/** Regression: Wave Optics inspired ewfd Settings wired to existing FEM. */
import assert from "node:assert/strict";
import {createRequire} from "node:module";
const {chromium}=createRequire(import.meta.url)("playwright");
const base=process.env.TEST_BASE_URL||"http://127.0.0.1:8765/COMSOL/universal_silica_v01/web/";
const browser=await chromium.launch({headless:true,args:["--no-sandbox","--disable-dev-shm-usage"]});
try{
 const page=await browser.newPage({viewport:{width:1440,height:860}});
 const errors=[];page.on("pageerror",e=>errors.push(e.message));
 await page.goto(base,{waitUntil:"load",timeout:90000});
 await page.locator('[data-wizard="cancel"]').click(); // Open legacy workbench for existing solver regression.
 await page.waitForSelector("#ls-wave-settings",{state:"visible"});
 assert.equal(await page.locator(".ls-model-tree [data-action='wavePhysics']").count(),1);
 assert.equal(await page.locator(".ls-model-tree [data-action='waveEquation']").count(),1);
 assert.equal(await page.locator(".ls-model-tree [data-action='wavePEC']").count(),1);
 assert.equal(await page.locator(".ls-model-tree [data-action='waveInitial']").count(),1);
 assert.ok((await page.locator("#ls-wave-settings").textContent()).includes("Electromagnetic Waves, Frequency Domain"));
 await page.screenshot({path:"/tmp/ls_waveoptics_ewfd_desktop.png"});
 await page.locator(".ls-model-tree [data-action='waveEquation']").click();
 assert.equal(await page.locator("#ls-wave-wavelength").count(),1);
 assert.equal(await page.locator("#ls-wave-delta").count(),1);
 await page.locator("#ls-wave-delta").fill("0.0055");
 await page.locator("[data-wave-command='applyWaveEquation']").click();
 assert.equal(Number(await page.locator("#delta_n_core").inputValue()),0.0055);
 await page.locator(".ls-model-tree [data-action='waveCladding']").click();
 await page.locator("#ls-wave-clad-mode").selectOption("fixed");
 await page.locator("#ls-wave-clad-index").fill("1.4378");
 await page.locator("[data-wave-command='applyClad']").click();
 assert.equal(await page.locator("#clad_material").inputValue(),"fixed");
 assert.equal(Number(await page.locator("#clad_index_override").inputValue()),1.4378);
 await page.locator(".ls-model-tree [data-action='waveSize']").click();
 await page.locator("#ls-wave-elements").fill("900");
 await page.locator("[data-wave-command='applyMesh']").click();
 assert.equal(Number(await page.locator("#elements").inputValue()),900);
 await page.locator(".ls-model-tree [data-action='waveMode']").click();
 assert.equal(await page.locator("#ls-wave-wavelength").count(),1);
 await page.locator("#ls-wave-wavelength").fill("1550");
 await page.locator("[data-wave-command='applyCompute']").click();
 await page.waitForFunction(()=>Number(document.querySelector("#neff")?.textContent)>1.4,{timeout:30000});
 await page.locator(".ls-model-tree [data-action='waveDerived']").click();
 assert.ok((await page.locator("#ls-wave-settings").textContent()).includes("MFD"));
 assert.ok((await page.locator("#ls-wave-settings").textContent()).includes("n_eff"));
 await page.screenshot({path:"/tmp/ls_waveoptics_results_desktop.png"});
 // Switching from ewfd to Geometry and back should not trap or overlap panes.
 await page.locator(".ls-model-tree [data-cad-node='c2']").click();
 assert.equal(await page.locator("#ls-cad-pane").isVisible(),true);
 assert.equal(await page.locator("#ls-wave-settings").isVisible(),false);
 await page.locator(".ls-model-tree [data-action='wavePhysics']").click();
 assert.equal(await page.locator("#ls-cad-pane").isVisible(),false);
 assert.equal(await page.locator("#ls-cad-graphics").isVisible(),false);
 assert.equal(await page.locator("#ls-wave-settings").isVisible(),true);
 await page.locator(".ls-model-tree [data-action='settings']").click();
 assert.equal(await page.locator("#ls-parameters-view").isVisible(),true);
 assert.equal(await page.locator("#ls-wave-settings").isVisible(),false);
 await page.locator(".ls-model-tree [data-action='waveMesh']").click();
 assert.equal(await page.locator("#ls-wave-settings").isVisible(),true);
 assert.equal(await page.locator("#ls-config-pane").isVisible(),false);
 // Old HCF research preset remains accessible from same tree.
 await page.locator(".ls-model-tree [data-select-preset='hcf']").click();
 assert.equal(await page.locator("#ls-wave-settings").isVisible(),false);
 assert.equal(await page.locator("#hcf-designer").isVisible(),true);
 assert.deepEqual(errors,[],"No browser JS errors");
 console.log("PASS: Wave Optics ewfd Settings, material and mesh inputs, real Mode Analysis, Results, Geometry/Parameters handoff, HCF compatibility");
}finally{await browser.close();}
