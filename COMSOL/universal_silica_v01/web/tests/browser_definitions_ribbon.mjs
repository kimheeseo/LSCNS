/**
 * Screenshot-based COMSOL Definitions ribbon regression:
 * clicking Definitions changes the upper ribbon, not the current model
 * or solver settings; unsupported features are visibly disabled.
 */
import assert from "node:assert/strict";
import {createRequire} from "node:module";
const {chromium}=createRequire(import.meta.url)("playwright");
const browser=await chromium.launch({headless:true,args:["--no-sandbox","--disable-dev-shm-usage"]});
try{
 const page=await browser.newPage({viewport:{width:1450,height:850}});
 const errors=[];
 page.on("pageerror",e=>errors.push(e.message));
 await page.goto(process.env.TEST_BASE_URL||"http://127.0.0.1:8765/COMSOL/universal_silica_v01/web/",{waitUntil:"load",timeout:90000});
 await page.locator('[data-wizard="cancel"]').click();
 const home=page.locator('[data-tab="home"]'),definitions=page.locator('[data-tab="definitions"]');
 assert.equal(await home.getAttribute("aria-pressed"),"true");
 // Geometry settings should survive a ribbon-only tab switch.
 await page.locator('[data-cad-node="c2"]').click();
 const titleBefore=await page.locator("#ls-cad-subtitle").textContent();
 await definitions.click();
 assert.equal(await definitions.getAttribute("aria-pressed"),"true");
 assert.equal(await home.getAttribute("aria-pressed"),"false");
 assert.equal(await page.locator("#ls-cad-subtitle").textContent(),titleBefore);
 assert.ok(await page.locator(".ls-ribbon").evaluate(el=>el.classList.contains("ls-ribbon-is-definitions")));
 assert.ok(await page.locator("body").evaluate(el=>el.classList.contains("ls-definitions-ribbon")));
 const names=["Variables","Functions","Selections","Probes","Coupling","Coordinate Systems","Coordinate Systems","Deformed Mesh","Optimization","View"];
 assert.deepEqual(await page.locator('#ls-ribbon-tools .ls-def-group-caption').allTextContents(),names);
 const selections=page.locator('#ls-ribbon-tools [data-def-group="2"]');
 for(const label of ["Explicit","Complement","Adjacent","Ball/Disk","Box","Cylinder","Union","Intersection","Difference","Colors"]){
  const cmd=selections.locator("button").filter({hasText:label});
  assert.equal(await cmd.count(),1,"Missing Definitions Selections command "+label);
  assert.equal(await cmd.isDisabled(),true,"Unimplemented "+label+" must be explicitly disabled");
 }
 for(const label of ["Analytic","Interpolation","Piecewise","More Functions"]){
  assert.ok(await page.locator('#ls-ribbon-tools .ls-def-command').filter({hasText:label}).count());
 }
 assert.equal(await page.locator('#ls-ribbon-tools [data-action="geometry"]').isDisabled(),false,"The existing View action should stay usable");
 const layoutHeight=await page.locator(".layout").evaluate(el=>el.getBoundingClientRect().height);
 assert.ok(layoutHeight>350,"Definitions ribbon should not collapse the workbench layout");
 await page.screenshot({path:"/tmp/comsol_definitions_tab_desktop.png",fullPage:true});
 await home.click();
 assert.equal(await home.getAttribute("aria-pressed"),"true");
 assert.equal(await page.locator(".ls-definitions-groups").count(),0);
 assert.ok(await page.locator('#ls-ribbon-tools [data-action="compute"]').isVisible());
 await definitions.click();
 await page.locator('#ls-ribbon-tools [data-action="geometry"]').click();
 assert.ok(await page.locator("#ls-cad-pane").isVisible());
 assert.deepEqual(errors,[]);
 const mobile=await browser.newPage({viewport:{width:390,height:850}});
 mobile.on("pageerror",e=>errors.push(e.message));
 await mobile.goto(process.env.TEST_BASE_URL||"http://127.0.0.1:8765/COMSOL/universal_silica_v01/web/",{waitUntil:"load",timeout:90000});
 await mobile.locator('[data-wizard="cancel"]').click();
 await mobile.locator('[data-tab="definitions"]').click();
 assert.ok(await mobile.locator('[data-def-group="2"]').count());
 assert.equal(await mobile.locator('#ls-ribbon-tools').evaluate(el=>getComputedStyle(el).overflowX),"auto");
 await mobile.screenshot({path:"/tmp/comsol_definitions_tab_mobile.png",fullPage:true});
 assert.deepEqual(errors,[]);
 console.log("PASS: screenshot-style Definitions tab with grouped commands, inactive placeholders, live View, state preservation, desktop/mobile responsive layout");
}finally{await browser.close();}
