/** Browser regression: Model Builder Definitions context menu screenshot workflow. */
import assert from "node:assert/strict";
import {createRequire} from "node:module";
const {chromium}=createRequire(import.meta.url)("playwright");
const browser=await chromium.launch({headless:true,args:["--no-sandbox","--disable-dev-shm-usage"]});
try{
 const page=await browser.newPage({viewport:{width:1460,height:940}});
 const errors=[];page.on("pageerror",error=>errors.push(error.message));
 await page.goto(process.env.TEST_BASE_URL||"http://127.0.0.1:8765/COMSOL/universal_silica_v01/web/",{waitUntil:"load",timeout:90000});
 await page.locator('[data-wizard="cancel"]').click();
 const definitions=page.locator("#ls-component-definitions > summary");
 await definitions.click({button:"right"});
 const menu=page.locator("#ls-definitions-context");
 assert.ok(await menu.isVisible(),"Definitions right-click must open context menu");
 const sections=await menu.locator(":scope > .ls-def-context-group").count();
 assert.equal(sections,5);
 for(const label of ["Variables","Functions","Selections","Probes","Nonlocal Couplings",
 "Pairs","Coordinate Systems","Moving Mesh","Deformed Geometry","Topology Optimization",
 "Shape Optimization","Perfectly Matched Layer","Shared Properties","View",
 "Show More Options...","Node Group","Group by Type","Update Probes","Help"]){
  assert.ok((await menu.textContent()).includes(label),"Missing "+label);
 }
 const selections=menu.locator('[data-def-context-parent="Selections"]');
 await selections.locator(":scope > button").hover();
 assert.ok(await selections.locator(".ls-def-context-nested").isVisible());
 const adjacent=selections.locator('[data-definitions-command="adjacent"]');
 assert.equal(await adjacent.isDisabled(),false);
 const explicit=selections.locator('button').filter({hasText:"Explicit"});
 assert.equal(await explicit.isDisabled(),true,"Unimplemented Explicit selection remains disabled");
 await page.screenshot({path:"/tmp/comsol_definitions_rightclick_menu.png"});
 await adjacent.click();
 assert.ok(await menu.isHidden(),"Choosing Adjacent closes context menu");
 assert.ok(await page.locator("#ls-adjacent-editor").isVisible());
 assert.ok(await page.locator('[data-adjacent-node="adj1"]').isVisible());
 assert.equal(await page.locator("#ls-adjacent-label").inputValue(),"Adjacent 1");
 await definitions.click({button:"right"});
 assert.ok(await menu.isVisible());
 await page.keyboard.press("Escape");
 assert.ok(await menu.isHidden(),"Escape dismisses the right-click menu");
 assert.deepEqual(errors,[],"No uncaught browser errors");
 console.log("PASS: Component Definitions right-click opens COMSOL-like menu; Selections → Adjacent creates editable node; unsupported commands disabled and Escape closes");
}finally{await browser.close();}
