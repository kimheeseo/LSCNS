/**
 * User screenshot regression:
 * Blank Material -> Definitions ribbon -> Adjacent 1 ->
 * Geometric entity level Boundary -> select independent air-hole rings ->
 * Label Air Hole Boundaries -> persist tree and blue selections.
 */
import assert from "node:assert/strict";
import {createRequire} from "node:module";
const {chromium}=createRequire(import.meta.url)("playwright");
const browser=await chromium.launch({headless:true,args:["--no-sandbox","--disable-dev-shm-usage"]});
try{
 const page=await browser.newPage({viewport:{width:1500,height:940}});
 const errors=[];page.on("pageerror",error=>errors.push(error.message));
 await page.goto(process.env.TEST_BASE_URL||"http://127.0.0.1:8765/COMSOL/universal_silica_v01/web/",{waitUntil:"load",timeout:90000});
 await page.locator('[data-dimension="2D"]').click();
 await page.locator('[data-interface="ewfd"]').click();
 await page.locator('[data-wizard="add"]').click();
 await page.locator('[data-wizard="select-added"]').click();
 await page.locator('[data-wizard="study"]').click();
 await page.locator('[data-study="mode"]').first().click();
 await page.locator('[data-wizard="done"]').click();
 // Reference cross-section: concentric outer/inner ring and six rotated holes.
 await page.locator('summary[data-action="geometry"]').click({button:"right"});
 await page.locator('#ls-cad-menu > ul > li > [data-cmd="circle"]').click();
 await page.locator("#ls-cad-radius").fill("rb+wlr");
 await page.locator("#ls-cad-layers-section summary").click();
 await page.locator("#ls-cad-layer").fill("wlr");
 await page.locator("#ls-cad-build-selected").click();
 await page.locator('summary[data-action="geometry"]').click({button:"right"});
 await page.locator('#ls-cad-menu > ul > li > [data-cmd="circle"]').click();
 await page.locator("#ls-cad-radius").fill("2.5[um]");
 await page.locator("#ls-cad-x").fill("6[um]");
 await page.locator('summary[data-action="geometry"]').click({button:"right"});
 await page.locator('#ls-cad-menu > ul > li.has-children > button').filter({hasText:"Transforms"}).hover();
 await page.locator('#ls-cad-menu li.has-children > ul [data-cmd="rotate"]').click();
 await page.locator("#ls-cad-keep").uncheck();
 await page.locator("#ls-cad-build-selected").click();
 assert.equal(await page.locator('#ls-cad-view [data-cad-shape="rot1"]').count(),6);
 // Material 1 must continue to show Blank Material settings as in image 1.
 await page.locator('summary[data-action="waveMaterials"]').click({button:"right"});
 await page.locator('[data-material-command="blank"]').click();
 assert.ok(await page.locator("#ls-materials-editor").isVisible());
 assert.equal(await page.locator("#ls-material-label").inputValue(),"Material 1");
 await page.screenshot({path:"/tmp/comsol_adjacent_material1_before.png"});
 // Definitions click changes ribbon. Adjacent must now be usable.
 await page.locator('[data-tab="definitions"]').click();
 const btn=page.locator('#ls-ribbon-tools [data-action="definitionsAdjacent"]');
 assert.equal(await btn.count(),1);
 assert.equal(await btn.isDisabled(),false);
 await btn.click();
 assert.ok(await page.locator("#ls-adjacent-editor").isVisible());
 assert.ok(await page.locator("#ls-materials-editor").isHidden(),"Materials panel must close");
 assert.ok(await page.locator("#ls-cad-graphics").isVisible());
 assert.equal(await page.locator("#ls-adjacent-label").inputValue(),"Adjacent 1");
 assert.equal(await page.locator("#ls-adjacent-level").inputValue(),"Domain");
 assert.equal(await page.locator("#ls-adjacent-output").inputValue(),"Adjacent boundaries");
 assert.ok(await page.locator('#ls-definitions-selections-list [data-adjacent-node="adj1"]').isVisible());
 await page.screenshot({path:"/tmp/comsol_adjacent_created.png"});
 // Screenshot regression: Adjacent must show built geometry with no CAD
 // labels, dashed Rotate outlines or inherited feature highlights.
 assert.equal(await page.locator('#ls-cad-view [stroke-dasharray]').count(),0);
 assert.equal(await page.locator('#ls-cad-view text').filter({hasText:/^(c[0-9]+|rot[0-9]+)\*?$/}).count(),0);
 assert.equal(await page.locator('#ls-cad-view path[stroke="#deebf3"]').count(),0);
 assert.equal(await page.locator('#ls-cad-view .is-input-selected, #ls-cad-view .is-material-selected, #ls-cad-view .is-cumulative-selected').count(),0);
 // A real Domain mouse click must select the one small disk, regardless
 // of overlapping transparent hit areas from the surrounding core.
 const disk=page.locator('#ls-cad-view [data-cad-entity="rot1/c2/0"] .ls-cad-rendered-surface circle');
 const center=await disk.evaluate(el=>{
  const point=el.ownerSVGElement.createSVGPoint();
  point.x=Number(el.getAttribute("cx"));point.y=Number(el.getAttribute("cy"));
  const q=point.matrixTransform(el.ownerSVGElement.getScreenCTM());
  return {x:q.x,y:q.y};
 });
 await page.mouse.move(center.x,center.y);
 assert.equal(await page.locator('#ls-cad-view .ls-cad-adjacent-entity.is-hovered').count(),1);
 const hoverColor=await page.locator('#ls-cad-view .ls-cad-adjacent-entity.is-hovered .ls-cad-adjacent-mark').evaluate(el=>getComputedStyle(el).stroke);
 assert.equal(hoverColor,"rgb(209, 60, 60)");
 assert.equal(await page.locator('#ls-cad-view .ls-cad-rendered-shape.is-hovered').count(),0);
 await page.screenshot({path:"/tmp/comsol_adjacent_domain_hover.png"});
 await page.mouse.click(center.x,center.y);
 assert.deepEqual(await page.evaluate(()=>window.LSWaveAdjacentSelection.list()[0].selectedByLevel.Domain),["rot1/c2/0/disk"]);
 assert.equal(await page.locator('#ls-adjacent-list input:checked').count(),1);
 assert.equal(await page.locator('#ls-cad-view .is-adjacent-selected').count(),1);
 assert.equal(await page.locator('#ls-cad-view .is-adjacent-selected .ls-cad-adjacent-mark').evaluate(el=>getComputedStyle(el).stroke),"rgb(22, 75, 250)");
 await page.screenshot({path:"/tmp/comsol_adjacent_domain_selected.png"});
 // Pausing keeps the blue selection but prevents geometry editing/picking.
 await page.locator('#ls-adjacent-activate').click();
 await page.mouse.click(center.x,center.y);
 assert.ok(await page.locator('#ls-adjacent-editor').isVisible());
 assert.equal(await page.locator('#ls-cad-view .is-adjacent-selected').count(),1);
 assert.equal((await page.evaluate(()=>window.LSWaveAdjacentSelection.list()[0].selectedByLevel.Domain)).length,1);
 await page.locator('#ls-adjacent-activate').click();
 await page.mouse.click(center.x,center.y);
 assert.deepEqual(await page.evaluate(()=>window.LSWaveAdjacentSelection.list()[0].selectedByLevel.Domain),[]);
 // Empty plot margin never toggles any domain.
 const margin=await page.locator('#ls-cad-view').evaluate(el=>{
  const p=el.createSVGPoint();p.x=60;p.y=36;
  const q=p.matrixTransform(el.getScreenCTM());return {x:q.x,y:q.y};
 });
 await page.mouse.click(margin.x,margin.y);
 assert.deepEqual(await page.evaluate(()=>window.LSWaveAdjacentSelection.list()[0].selectedByLevel.Domain),[]);

 // Boundary mode: drawn edges are independent hit targets; selecting one
 // is NOT the same as selecting a filled Material domain.
 await page.locator("#ls-adjacent-level").selectOption("Boundary");
 const targets=page.locator('#ls-cad-view .ls-cad-adjacent-overlay [data-adjacent-key]');
 const num=await targets.count();
 assert.ok(num>=8,"Expect outer circle, inner ring, spokes and six air-hole boundaries. Got "+num);
 const a='rot1/c2/0/outer-boundary',b='rot1/c2/1/outer-boundary';
 assert.equal(await page.locator('#ls-cad-view [data-adjacent-key="'+a+'"]').count(),1);
 assert.equal(await page.locator('#ls-cad-view [data-adjacent-key="'+b+'"]').count(),1);
 const circle=page.locator('#ls-cad-view [data-adjacent-key="'+a+'"] circle').last();
 // Force physical mouse click on the ring's northern point.
 const pixel=await circle.evaluate(el=>{
  const svg=el.ownerSVGElement,point=svg.createSVGPoint();
  point.x=Number(el.getAttribute("cx"));
  point.y=Number(el.getAttribute("cy"))-Number(el.getAttribute("r"));
  const q=point.matrixTransform(svg.getScreenCTM());
  return {x:q.x,y:q.y};
 });
 await page.evaluate(()=>{
  window.__lsBoundaryClicks=[];
  document.addEventListener("click",event=>{
   window.__lsBoundaryClicks.push({
    target:event.target?.tagName,
    html:event.target?.outerHTML?.slice(0,250),
    adjacent:event.target?.closest?.("[data-adjacent-key]")?.getAttribute("data-adjacent-key")||null
   });
  },true);
 });
 await page.mouse.click(pixel.x,pixel.y);
 const diagnostic=await page.evaluate(({x,y})=>({
  pixel:{x,y},hitTag:document.elementFromPoint(x,y)?.tagName,
  hitHTML:document.elementFromPoint(x,y)?.outerHTML?.slice(0,300),
  clickEvents:window.__lsBoundaryClicks,
  paneVisible:!document.querySelector("#ls-adjacent-editor")?.hidden,
  active:document.querySelector("#ls-cad-view .ls-cad-adjacent-overlay")?.outerHTML?.slice(0,250),
  body:document.body.className,
  selection:window.LSWaveAdjacentSelection.list()[0].selectedByLevel.Boundary
 }),pixel);
 console.log("Adjacent real click diagnostics:",JSON.stringify(diagnostic));
 let saved=await page.evaluate(()=>window.LSWaveAdjacentSelection.list()[0].selectedByLevel.Boundary);
 assert.equal(saved.length,1,"One ring mouse click adds one boundary only");
 assert.equal(saved[0],a);
 assert.equal(await page.locator('#ls-cad-view .ls-cad-adjacent-entity.is-adjacent-selected').count(),1);
 await page.locator('#ls-cad-view [data-adjacent-key="'+b+'"]').dispatchEvent("click");
 saved=await page.evaluate(()=>window.LSWaveAdjacentSelection.list()[0].selectedByLevel.Boundary);
 assert.deepEqual(new Set(saved),new Set([a,b]));
 assert.equal(await page.locator('#ls-adjacent-list input:checked').count(),2);
 await page.screenshot({path:"/tmp/comsol_adjacent_two_boundaries_selected.png"});
 await page.locator("#ls-adjacent-label").fill("Air Hole Boundaries");
 const feature=page.locator('#ls-definitions-selections-list [data-adjacent-node="adj1"]');
 assert.ok((await feature.textContent()).includes("Air Hole Boundaries"));
 assert.equal(await page.evaluate(()=>window.LSWaveAdjacentSelection.list()[0].label),"Air Hole Boundaries");
 // Return to Home and click the named node; selected rings must be persistent.
 await page.locator('[data-tab="home"]').click();
 await feature.click();
 assert.equal(await page.locator("#ls-adjacent-label").inputValue(),"Air Hole Boundaries");
 assert.equal(await page.locator("#ls-adjacent-level").inputValue(),"Boundary");
 assert.equal(await page.locator('#ls-cad-view .ls-cad-adjacent-entity.is-adjacent-selected').count(),2);
 await page.locator('#ls-cad-view [data-adjacent-key="'+a+'"]').dispatchEvent("click");
 assert.equal((await page.evaluate(()=>window.LSWaveAdjacentSelection.list()[0].selectedByLevel.Boundary)).length,1,"Clicking selected edge toggles just that edge");
 assert.equal(await page.locator('#ls-cad-view .ls-cad-adjacent-entity.is-adjacent-selected').count(),1);
 // Existing CAD editor and named selection should survive leaving Adjacent.
 await page.locator('[data-cad-node="c1"]').click();
 assert.ok(await page.locator("#ls-cad-pane").isVisible());
 assert.ok(await page.locator("#ls-adjacent-editor").isHidden());
 await feature.click();
 assert.equal(await page.locator("#ls-adjacent-label").inputValue(),"Air Hole Boundaries");
 assert.equal((await page.evaluate(()=>window.LSWaveAdjacentSelection.list()[0].selectedByLevel.Boundary)).length,1);
 assert.deepEqual(errors,[],"No uncaught page errors");
 console.log("PASS: Blank Material -> Definitions Adjacent -> boundary-level blue ring click, two independent air holes, editable label, Model Builder selection persistence, CAD coexistence");
}finally{await browser.close();}
