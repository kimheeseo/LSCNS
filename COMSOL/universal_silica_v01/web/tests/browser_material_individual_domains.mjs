/** COMSOL screenshot: Material 1 > Domain > Manual => individual domain ID picking on Graphics. */
import assert from "node:assert/strict";
import {createRequire} from "node:module";
const {chromium}=createRequire(import.meta.url)("playwright");
const browser=await chromium.launch({headless:true,args:["--no-sandbox","--disable-dev-shm-usage"]});
const selected=async page=>page.evaluate(()=>window.LSWaveMaterialsEditor.list()[0].domainIds.sort((a,b)=>a-b));
try{
 const page=await browser.newPage({viewport:{width:1500,height:940}});
 const errors=[];page.on("pageerror",error=>errors.push(error.message));
 await page.goto(process.env.TEST_BASE_URL||"http://127.0.0.1:8765/COMSOL/universal_silica_v01/web/",{waitUntil:"load",timeout:90000});
 // New 2D ewfd model and photographed Parameters table.
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
 assert.equal(await page.locator('#ls-cad-view [data-cad-shape="c1"] circle.ls-cad-layer-boundary').count(),1);
 // 6 hole disks inside the inner rb circle.
 await page.locator('summary[data-action="geometry"]').click({button:"right"});
 await page.locator('#ls-cad-menu > ul > li > [data-cmd="circle"]').click();
 await page.locator("#ls-cad-radius").fill("2.5[um]");
 await page.locator("#ls-cad-x").fill("6[um]");
 await page.locator('summary[data-action="geometry"]').click({button:"right"});
 await page.locator('#ls-cad-menu > ul > li.has-children > button').filter({hasText:"Transforms"}).hover();
 await page.locator('#ls-cad-menu li.has-children > ul [data-cmd="rotate"]').click();
 assert.equal(await page.locator("#ls-cad-angle").inputValue(),"range(0,60,300)");
 await page.locator("#ls-cad-keep").uncheck();
 await page.locator("#ls-cad-build-selected").click();
 assert.equal(await page.locator('#ls-cad-view [data-cad-shape="rot1"]').count(),6);
 // New blank Material 1 with 11 independent preview regions:
 // 4 annular sectors + 1 inner disk + 6 air holes.
 await page.locator('summary[data-action="waveMaterials"]').click({button:"right"});
 await page.locator('[data-material-command="blank"]').click();
 assert.equal(await page.locator("#ls-material-entity-level").inputValue(),"Domain");
 assert.equal(await page.locator("#ls-material-select-mode").inputValue(),"all");
 const total=await page.locator('#ls-material-domain-list [data-material-domain]').count();
 assert.equal(total,11,"Expected four independent annular sectors, one core region, six air holes");
 await page.locator("#ls-material-select-mode").selectOption("manual");
 assert.deepEqual(await selected(page),[],"Manual selection should start empty");
 assert.equal(await page.locator('#ls-material-domain-list input:checked').count(),0);
 assert.equal(await page.locator('#ls-cad-view .ls-cad-material-region.is-selected').count(),0);
 assert.equal(await page.locator("#ls-material-pick").getAttribute("aria-pressed"),"true");
 const annulus=page.locator('#ls-cad-view [data-cad-shape="c1"] .ls-cad-rendered-surface circle');
 const coords=await annulus.evaluate(c=>{
  const cx=Number(c.getAttribute("cx")),cy=Number(c.getAttribute("cy")),r=Number(c.getAttribute("r"));
  const point=c.ownerSVGElement.createSVGPoint();
  point.x=cx+r*.75/Math.SQRT2;point.y=cy-r*.75/Math.SQRT2;
  const screen=point.matrixTransform(c.ownerSVGElement.getScreenCTM());
  return {x:screen.x,y:screen.y};
 });
 await page.mouse.click(coords.x,coords.y);
 assert.deepEqual(await selected(page),[1],"Physical mouse click in upper-right annulus must select just domain 1");
 assert.equal(await page.locator('#ls-material-domain-list [data-material-domain="1"]').isChecked(),true);
 assert.equal(await page.locator('#ls-cad-view .ls-cad-material-region.is-selected').count(),1);
 assert.equal(await page.locator('#ls-cad-view [data-cad-domain-id="1"].is-selected').count(),1);
 assert.equal(await page.locator('#ls-cad-view [data-cad-domain-id="2"].is-selected').count(),0);
 // A second click toggles ONLY the same domain.
 await page.mouse.click(coords.x,coords.y);
 assert.deepEqual(await selected(page),[]);
 // Checkbox and Graphics both update the same one-domain state.
 await page.locator('#ls-material-domain-list [data-material-domain="1"]').check();
 assert.deepEqual(await selected(page),[1]);
 assert.equal(await page.locator('#ls-cad-view .ls-cad-material-region.is-selected').count(),1);
 // Screenshot example: individually selected 1,2,4,5,6,7 (not all regions).
 for(const id of [2,4,5,6,7]){
  await page.locator('#ls-cad-view [data-cad-domain-id="'+id+'"]').dispatchEvent("click");
 }
 assert.deepEqual(await selected(page),[1,2,4,5,6,7]);
 assert.equal(await page.locator('#ls-cad-view .ls-cad-material-region.is-selected').count(),6);
 assert.equal(await page.locator('#ls-material-domain-list input:checked').count(),6);
 assert.equal(await page.locator('#ls-cad-view [data-cad-domain-id="3"].is-selected').count(),0);
 await page.screenshot({path:"/tmp/comsol_material_individual_selected_domains.png"});
 assert.deepEqual(errors,[],"No JavaScript errors");
 console.log("PASS: Material 1 Manual selection of 11 partitioned regions; true mouse click picks only one quadrant, blue fill and domain-number checkboxes stay synchronized");
}finally{await browser.close();}
