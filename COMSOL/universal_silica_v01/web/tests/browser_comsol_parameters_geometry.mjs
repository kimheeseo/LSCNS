/** Browser test: COMSOL-like Parameters / dynamic Circle / Rotate / right-click / Radial FEM. */
import assert from "node:assert/strict";
import {createRequire} from "node:module";
const {chromium}=createRequire(import.meta.url)("playwright");
const base=process.env.TEST_BASE_URL||"http://127.0.0.1:8765/COMSOL/universal_silica_v01/web/";
const browser=await chromium.launch({headless:true,args:["--no-sandbox","--disable-dev-shm-usage"]});
try{
 const page=await browser.newPage({viewport:{width:1440,height:850}});
 const errors=[];page.on("pageerror",err=>errors.push(err.message));
 await page.goto(base,{waitUntil:"load",timeout:90000});
 await page.locator('[data-wizard="cancel"]').click(); // Open legacy workbench for existing solver regression.
 await page.waitForSelector("#ls-cad-pane",{state:"attached"});
 await page.locator('.ls-model-tree [data-action="settings"]').click();
 assert.equal(await page.locator("#ls-parameters-view").isVisible(),true);
 assert.equal(await page.locator("#ls-parameter-rows tr").count(),10);
 await page.locator('#ls-parameter-rows [data-field="expression"]').first().fill("12[um]");
 assert.equal((await page.locator('[data-param-value="0"]').textContent()).trim(),"1.2E-5 m");
 await page.locator('.ls-model-tree [data-cad-node="c2"]').click();
 assert.equal(await page.locator("#ls-cad-pane").isVisible(),true);
 assert.equal(await page.locator("#ls-cad-graphics").isVisible(),true);
 assert.equal(await page.locator("#ls-cad-radius").inputValue(),"d/2");
 assert.equal(await page.locator("#ls-cad-sector").inputValue(),"360");
 assert.equal(await page.locator("#ls-cad-color").inputValue(),"None");
 assert.equal(await page.locator("#ls-cad-physics").isDisabled(),true);
 assert.ok((await page.locator("#ls-cad-legend").textContent()).includes("2 object"));
 await page.locator("#ls-cad-apply").click();
 assert.ok((await page.locator("#ls-cad-status").textContent()).startsWith("Applied"));
 assert.equal(Number(await page.locator("#core_radius_um").inputValue()),6);
 await page.screenshot({path:"/tmp/ls_comsol_geometry_circle.png"});
 // Geometry 1 context menu and feature addition.
 await page.locator('summary[data-action="geometry"]').click({button:"right"});
 assert.equal(await page.locator("#ls-cad-menu").isVisible(),true);
 assert.equal(await page.locator('#ls-cad-menu [data-cmd="circle"]').isVisible(),true);
 assert.ok(await page.locator('#ls-cad-menu button').filter({hasText:"Transforms"}).count()>0);
 await page.locator('#ls-cad-menu > ul > li > [data-cmd="circle"]').click();
 assert.equal(await page.locator('#ls-geometry-node-list [data-cad-node="c3"]').count(),1);
 await page.locator("#ls-cad-radius").fill("2.5[um]");
 await page.locator("#ls-cad-y").fill("Delta");
 assert.ok((await page.locator("#ls-cad-legend").textContent()).includes("3 object"));
 // Add Rotate through the Geometry 1 right-click Transforms submenu.
 await page.locator('summary[data-action="geometry"]').click({button:"right"});
 await page.locator('#ls-cad-menu > ul > li.has-children > button').filter({hasText:"Transforms"}).hover();
 assert.equal(await page.locator('#ls-cad-menu li.has-children > ul [data-cmd="rotate"]').isVisible(),true);
 await page.locator('#ls-cad-menu li.has-children > ul [data-cmd="rotate"]').click();
 assert.equal(await page.locator("#ls-cad-subtitle").textContent(),"Rotate");
 assert.equal(await page.locator("#ls-cad-angle").inputValue(),"range(0,60,300)");
 assert.equal(await page.locator('#ls-cad-inputs [data-cad-input="c3"]').isChecked(),true);
 assert.ok((await page.locator("#ls-cad-legend").textContent()).includes("9 object"));
 await page.locator("#ls-cad-keep").uncheck();
 assert.ok((await page.locator("#ls-cad-legend").textContent()).includes("8 object"));
 await page.screenshot({path:"/tmp/ls_comsol_geometry_rotate.png"});
 await page.locator("#ls-cad-apply").click();
 assert.ok((await page.locator("#ls-cad-status").textContent()).includes("Cannot apply"));
 assert.equal(Number(await page.locator("#core_radius_um").inputValue()),6);
 // Right-click Rotate: build selected, duplicate, disable, rename, delete menu.
 await page.locator('[data-cad-node="rot1"]').click({button:"right"});
 for(const item of ["build-selected","build-preceding","duplicate","delete","toggle","rename","settings"]){
  assert.ok(await page.locator('#ls-cad-menu [data-cmd="'+item+'"]').count()>0,item);
 }
 await page.locator('#ls-cad-menu [data-cmd="toggle"]').click();
 assert.ok((await page.locator('[data-cad-node="rot1"]').getAttribute("class")).includes("ls-cad-node-disabled"));
 // Extra circle also blocks Radial FEM until disabled.
 await page.locator('[data-cad-node="c3"]').click({button:"right"});
 await page.locator('#ls-cad-menu [data-cmd="toggle"]').click();
 await page.locator("#ls-cad-apply").click();
 assert.ok((await page.locator("#ls-cad-status").textContent()).startsWith("Applied"));
 await page.locator('[data-cad-node="c2"]').click();
 await page.locator("#ls-cad-x").fill("1[um]");
 await page.locator("#ls-cad-apply").click();
 assert.ok((await page.locator("#ls-cad-status").textContent()).includes("concentric"));
 assert.equal(Number(await page.locator("#core_radius_um").inputValue()),6);
 await page.locator("#ls-cad-x").fill("0[um]");
 await page.locator("#ls-cad-close").click();
 assert.equal(await page.locator("#ls-cad-pane").isVisible(),false);
 assert.equal(await page.locator("#calculateBtn").isVisible(),true);
 await page.locator("#calculateBtn").click();
 await page.waitForFunction(()=>Number(document.querySelector("#neff")?.textContent)>1.4,{timeout:30000});
 assert.deepEqual(errors,[],"No browser JavaScript errors");
 console.log("PASS: COMSOL Parameters, multi-Circle Geometry, Rotate range/keep-inputs, right-click menus, FEM guards, original solver");
}finally{await browser.close();}
