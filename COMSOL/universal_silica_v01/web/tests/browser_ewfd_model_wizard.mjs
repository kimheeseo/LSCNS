/** New COMSOL-like ewfd wizard / Parameters import / first Circle browser smoke. */
import assert from "node:assert/strict";
import {createRequire} from "node:module";
import {fileURLToPath} from "node:url";
import path from "node:path";
const {chromium}=createRequire(import.meta.url)("playwright");
const browser=await chromium.launch({headless:true,args:["--no-sandbox","--disable-dev-shm-usage"]});
try{
 const page=await browser.newPage({viewport:{width:1365,height:768}});
 const errors=[];page.on("pageerror",e=>errors.push(e.message));
 await page.goto(process.env.TEST_BASE_URL||"http://127.0.0.1:8765/COMSOL/universal_silica_v01/web/",{waitUntil:"load",timeout:90000});
 await page.waitForSelector("#ls-startup-wizard",{state:"visible"});
 assert.ok(await page.getByRole("heading",{name:"Select Space Dimension"}).isVisible());
 await page.locator('[data-dimension="2D"]').click();
 assert.ok(await page.getByRole("heading",{name:"Select Physics"}).isVisible());
 assert.equal(await page.locator('#lsw-tree [data-toggle="optics"]').count(),1);
 assert.equal(await page.locator('#lsw-tree [data-toggle="wave"]').count(),1);
 assert.equal(await page.locator('#lsw-tree [data-toggle="ray"]').count(),1);
 const studyBtn=page.locator('[data-wizard="study"]');
 assert.ok(await studyBtn.isDisabled());
 await page.locator('[data-interface="ewfd"]').click();
 await page.locator('[data-wizard="add"]').click();
 assert.ok(await page.locator('[data-wizard="select-added"]').isVisible());
 assert.ok(await studyBtn.isDisabled(),"Study requires selecting added ewfd");
 await page.locator('[data-wizard="select-added"]').click();
 assert.ok(await studyBtn.isEnabled());
 await studyBtn.click();
 assert.ok(await page.getByRole("heading",{name:"Select Study"}).isVisible());
 assert.ok(await page.locator('[data-wizard="done"]').isDisabled());
 await page.locator('[data-study="mode"]').first().click();
 assert.ok((await page.locator(".lsw-added-small").textContent()).includes("Mode Analysis"));
 await page.locator('[data-wizard="done"]').click();
 assert.ok(await page.locator("#ls-startup-wizard").isHidden());
 assert.ok(await page.locator(".ls-model-tree").isVisible());
 assert.ok(await page.locator("#ls-cad-root-settings").isVisible());
 assert.equal(await page.locator("#ls-geometry-node-list [data-cad-node]").count(),0,"New model starts with empty geometry");
 await page.locator('.ls-model-tree [data-action="settings"]').first().click();
 assert.ok(await page.locator("#ls-parameters-view").isVisible());
 const defaults=[
  ["d","5[um]","5E-6 m"],["Delta","6.75[um]","6.75E-6 m"],
  ["rb","10[um]","1E-5 m"],["nbg","1.45","1.45"],
  ["nair","1","1"],["wl","1.45[um]","1.45E-6 m"],
  ["neffMax","1.4454","1.4454"],["neffMin","1.429","1.429"],
  ["wlr","wl/sqrt(nbg^2-neffMax^2)",null]
 ];
 assert.equal(await page.locator("#ls-param-label").inputValue(),"Parameters 1");
 assert.equal(await page.locator("#ls-parameter-rows tr").count(),9,"New ewfd model uses nine screenshot reference defaults");
 for(let i=0;i<defaults.length;i++){
  const row=page.locator("#ls-parameter-rows tr").nth(i);
  assert.equal(await row.locator('[data-field="name"]').inputValue(),defaults[i][0]);
  assert.equal(await row.locator('[data-field="expression"]').inputValue(),defaults[i][1]);
  if(defaults[i][2]!==null)assert.equal((await row.locator('[data-param-value]').textContent()).trim(),defaults[i][2]);
 }
 assert.ok(Math.abs(Number((await page.locator('#ls-parameter-rows tr').nth(8).locator('[data-param-value]').textContent()).replace(/E/,'e').replace(/ m$/,''))-1.2564e-5)<1e-8,"wlr derived value matches screenshot");
 // Values and descriptions remain user-editable; Restore defaults returns to nine rows.
 await page.locator("#ls-parameter-rows tr").first().locator('[data-field="expression"]').fill("6[um]");
 assert.equal((await page.locator('#ls-parameter-rows tr').first().locator('[data-param-value]').textContent()).trim(),"6E-6 m");
 await page.locator('#ls-parameter-rows tr').first().locator('[data-field="description"]').fill("New hole diameter");
 await page.locator("#ls-params-add").click();
 assert.equal(await page.locator("#ls-parameter-rows tr").count(),10);
 await page.locator("#ls-params-reset").click();
 assert.equal(await page.locator("#ls-parameter-rows tr").count(),9);
 assert.equal(await page.locator("#ls-parameter-rows tr").first().locator('[data-field="expression"]').inputValue(),"5[um]");
 const sample=path.resolve(path.dirname(fileURLToPath(import.meta.url)),"../examples/COMSOL_microstructured_parameters.txt");
 await page.locator("#ls-params-file").setInputFiles(sample);
 await page.waitForFunction(()=>/불러왔습니다|실패/.test(document.querySelector("#ls-config-status")?.textContent||""),{timeout:10000});
 assert.ok(!(await page.locator("#ls-config-status").textContent()).includes("실패"),"Import failed: "+await page.locator("#ls-config-status").textContent());
 assert.equal(await page.locator("#ls-parameter-rows tr").count(),10);
 assert.ok((await page.locator('#ls-parameter-rows [data-field="description"]').nth(1).inputValue()).includes("Hole pitch"));
 await page.locator('summary[data-action="geometry"]').click();
 assert.ok(await page.locator("#ls-cad-root-settings").isVisible());
 assert.ok(await page.locator("#ls-cad-angle-unit").isVisible());
 assert.ok(await page.locator("#ls-cad-build-leave").isVisible());
 await page.locator('summary[data-action="geometry"]').click({button:"right"});
 assert.ok(await page.locator("#ls-cad-menu").isVisible());
 await page.locator('#ls-cad-menu > ul > li > [data-cmd="circle"]').click();
 assert.ok(await page.locator('#ls-geometry-node-list [data-cad-node="c1"]').isVisible());
 assert.equal(await page.locator("#ls-cad-radius").inputValue(),"1[m]");
 await page.locator("#ls-cad-radius").fill("40[um]");
 assert.ok((await page.locator("#ls-cad-legend").textContent()).includes("1 object"));
 assert.deepEqual(errors,[],"No browser JavaScript errors");
 await page.screenshot({path:"/tmp/ls_comsol_ewfd_new_model_wizard.png"});
 console.log("PASS: new 2D ewfd wizard → Study → Mode Analysis → Parameters import → Geometry 1 → Circle 1");
}finally{await browser.close();}
