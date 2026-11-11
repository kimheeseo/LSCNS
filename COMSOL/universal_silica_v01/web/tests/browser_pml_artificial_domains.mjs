/**
 * Screenshot-comparable PML UI:
 * Definitions right-click > Perfectly Matched Layer ->
 * Component Definitions > Artificial Domains > PML 1.
 * Validates Geometry-only domain selection and settings, NOT FEM absorption.
 */
import assert from "node:assert/strict";
import {createRequire} from "node:module";
const {chromium}=createRequire(import.meta.url)("playwright");
const browser=await chromium.launch({headless:true,args:["--no-sandbox","--disable-dev-shm-usage"]});
try{
 const page=await browser.newPage({viewport:{width:1500,height:940}});
 const errors=[];page.on("pageerror",e=>errors.push(e.message));
 await page.goto(process.env.TEST_BASE_URL||"http://127.0.0.1:8765/COMSOL/universal_silica_v01/web/",{waitUntil:"load",timeout:90000});
 await page.locator('[data-wizard="cancel"]').click();
 await page.locator('[data-cad-node="c1"]').click();
 await page.locator("#ls-cad-radius").fill("rb+wlr");
 await page.locator("#ls-cad-layers-section summary").click();
 await page.locator("#ls-cad-layer").fill("wlr");
 await page.locator("#ls-cad-build-selected").click();
 const definitions=page.locator("#ls-component-definitions > summary");
 await definitions.click({button:"right"});
 const menu=page.locator("#ls-definitions-context");
 const addPML=menu.locator('[data-definitions-command="pml"]');
 assert.ok(await addPML.isVisible());
 assert.equal(await addPML.isDisabled(),false);
 await page.screenshot({path:"/tmp/comsol_pml_definitions_context.png"});
 await addPML.click();
 assert.ok(await menu.isHidden());
 const editor=page.locator("#ls-pml-editor");
 assert.ok(await editor.isVisible());
 assert.ok(await page.locator("#ls-cad-graphics").isVisible());
 assert.ok(await page.locator("#ls-cad-pane").isHidden());
 assert.equal(await page.locator("#ls-pml-label").inputValue(),"Perfectly Matched Layer 1");
 assert.equal(await page.locator("#ls-pml-name").inputValue(),"pml1");
 assert.equal(await page.locator("#ls-pml-selection").inputValue(),"manual");
 assert.equal(await page.locator("#ls-pml-type").inputValue(),"Cartesian");
 assert.deepEqual(await page.locator("#ls-pml-type option").allTextContents(),["Cartesian","Cylindrical","User defined"]);
 assert.equal(await page.locator("#ls-pml-center-x").count(),0,"Cylindrical center coordinates should be hidden for Cartesian");
 assert.equal(await page.locator("#ls-pml-stretch").inputValue(),"Polynomial");
 assert.equal(await page.locator("#ls-pml-wavelength-source").inputValue(),"Physics interface");
 assert.equal(await page.locator("#ls-pml-physics").inputValue(),"ewfd");
 assert.ok((await page.locator("#ls-pml-physics").textContent()).includes("Electromagnetic Waves, Frequency Domain"));
 assert.equal(await page.locator("#ls-pml-factor").inputValue(),"1");
 assert.equal(await page.locator("#ls-pml-curvature").inputValue(),"1");
 assert.ok(await page.locator("#ls-definitions-artificial-group").isVisible());
 assert.ok(await page.locator('#ls-definitions-artificial-list [data-pml-node="pml1"]').isVisible());
 const total=await page.locator('#ls-pml-domain-list input[data-pml-domain-key]').count();
 assert.ok(total>=5,"Five or more Geometry domains are available from one layered circle");
 assert.equal(await page.locator('#ls-pml-domain-list input:checked').count(),0,"Manual PML selection starts empty");
 // Select three independent domains, as in the supplied COMSOL screenshot.
 const keys=await page.evaluate(()=>window.LSWavePMLService.domains().filter(d=>[1,3,4].includes(d.id)).map(d=>d.key));
 assert.equal(keys.length,3);
 for(const key of keys)await page.locator('#ls-pml-domain-list input[data-pml-domain-key="'+key+'"]').check();
 assert.equal(await page.locator('#ls-pml-domain-list input:checked').count(),3);
 assert.equal(await page.locator('#ls-cad-view .ls-cad-pml-region.is-pml-selected').count(),3);
 assert.deepEqual((await page.evaluate(()=>window.LSWavePMLService.selected().map(d=>d.id))).sort((a,b)=>a-b),[1,3,4]);
 await page.screenshot({path:"/tmp/comsol_pml_three_domains_settings.png"});
 // Actual mouse click in the middle of domain 3's annular quadrant
 // toggles ONLY that chosen PML region.
 const loc=await page.evaluate(()=>{
  const d=window.LSWavePMLService.domains().find(v=>v.id===3);
  const circle=document.querySelector('#ls-cad-view [data-cad-shape="c1"] .ls-cad-rendered-surface circle');
  if(!circle)throw Error("No outer Geometry circle");
  const pixelPerMeter=Number(circle.getAttribute("r"))/window.LSWavePMLService.domains()[0].outerR;
  const r=(d.innerR+d.outerR)/2*pixelPerMeter;
  const theta=(d.startAngle+d.endAngle)/2*Math.PI/180;
  const svg=circle.ownerSVGElement,p=svg.createSVGPoint();
  p.x=Number(circle.getAttribute("cx"))+r*Math.cos(theta);
  p.y=Number(circle.getAttribute("cy"))-r*Math.sin(theta);
  const s=p.matrixTransform(svg.getScreenCTM());
  return {x:s.x,y:s.y};
 });
 await page.mouse.click(loc.x,loc.y);
 assert.deepEqual((await page.evaluate(()=>window.LSWavePMLService.selected().map(d=>d.id))).sort((a,b)=>a-b),[1,4]);
 assert.equal(await page.locator('#ls-cad-view .ls-cad-pml-region.is-pml-selected').count(),2);
 await page.mouse.click(loc.x,loc.y);
 assert.deepEqual((await page.evaluate(()=>window.LSWavePMLService.selected().map(d=>d.id))).sort((a,b)=>a-b),[1,3,4]);
 // All domains option and returning to Manual preserves independent picks.
 await page.locator("#ls-pml-selection").selectOption("all");
 assert.equal(await page.locator('#ls-pml-domain-list input:checked').count(),total);
 await page.locator("#ls-pml-selection").selectOption("manual");
 assert.equal(await page.locator('#ls-pml-domain-list input:checked').count(),3);
 await page.locator("#ls-pml-label").fill("Outer Cladding PML");
 const feature=page.locator('#ls-definitions-artificial-list [data-pml-node="pml1"]');
 assert.ok((await feature.textContent()).includes("Outer Cladding PML"));
 await page.locator("#ls-pml-type").selectOption("Cylindrical");
 assert.equal(await page.locator("#ls-pml-center-x").inputValue(),"0");
 assert.equal(await page.locator("#ls-pml-center-y").inputValue(),"0");
 await page.locator("#ls-pml-center-x").fill("3e-6");
 await page.locator("#ls-pml-center-y").fill("-2e-6");
 await page.locator("#ls-pml-type").selectOption("User defined");
 assert.equal(await page.locator("#ls-pml-center-x").count(),0);
 await page.locator("#ls-pml-type").selectOption("Cylindrical");
 assert.equal(await page.locator("#ls-pml-center-x").inputValue(),"3e-6","Cylindrical center should survive mode switching");
 assert.equal(await page.locator("#ls-pml-center-y").inputValue(),"-2e-6");
 await page.locator("#ls-pml-stretch").selectOption("Rational");
 await page.locator("#ls-pml-wavelength-source").selectOption("User defined");
 await page.locator("#ls-pml-wavelength").fill("wl");
 await page.locator("#ls-pml-factor").fill("1.5");
 await page.locator("#ls-pml-curvature").fill("2");
 await page.screenshot({path:"/tmp/comsol_pml_custom_settings.png"});
 // Switch editors then return to PML; Geometry selection and settings persist.
 await page.locator('[data-cad-node="c1"]').click();
 assert.ok(await editor.isHidden());
 assert.ok(await page.locator("#ls-cad-pane").isVisible());
 await feature.click();
 assert.ok(await editor.isVisible());
 assert.equal(await page.locator("#ls-pml-label").inputValue(),"Outer Cladding PML");
 assert.equal(await page.locator("#ls-pml-type").inputValue(),"Cylindrical");
 assert.equal(await page.locator("#ls-pml-stretch").inputValue(),"Rational");
 assert.equal(await page.locator("#ls-pml-wavelength-source").inputValue(),"User defined");
 assert.equal(await page.locator("#ls-pml-wavelength").inputValue(),"wl");
 assert.equal(await page.locator("#ls-pml-center-x").inputValue(),"3e-6");
 assert.equal(await page.locator("#ls-pml-center-y").inputValue(),"-2e-6");
 assert.equal(await page.locator("#ls-pml-factor").inputValue(),"1.5");
 assert.equal(await page.locator("#ls-pml-curvature").inputValue(),"2");
 assert.equal(await page.locator("#ls-pml-physics").inputValue(),"ewfd");
 const saved=await page.evaluate(()=>window.LSWavePMLService.list()[0]);
 assert.equal(saved.scalingFactor,"1.5");
 assert.equal(saved.curvatureParameter,"2");
 assert.equal(saved.centerX,"3e-6");
 assert.equal(saved.centerY,"-2e-6");
 assert.deepEqual((await page.evaluate(()=>window.LSWavePMLService.selected().map(d=>d.id))).sort((a,b)=>a-b),[1,3,4]);
 // The alternate Definitions ribbon entry is connected to the same editor.
 await page.locator('[data-tab="definitions"]').click();
 const ribbonPML=page.locator('#ls-ribbon-tools [data-action="definitionsPML"]');
 assert.equal(await ribbonPML.isDisabled(),false);
 await ribbonPML.click();
 assert.ok(await page.locator('#ls-definitions-artificial-list [data-pml-node="pml2"]').isVisible());
 assert.equal(await page.locator("#ls-pml-name").inputValue(),"pml2");
 assert.equal(await page.locator("#ls-pml-type").inputValue(),"Cartesian");
 assert.equal(await page.locator("#ls-pml-factor").inputValue(),"1");
 assert.equal(await page.locator("#ls-pml-curvature").inputValue(),"1");
 assert.equal((await page.evaluate(()=>window.LSWavePMLService.list())).length,2);
 // Existing Material editor still opens normally after PML activation.
 await page.locator('summary[data-action="waveMaterials"]').click({button:"right"});
 await page.locator('#ls-materials-context [data-material-command="blank"]').click();
 assert.ok(await page.locator("#ls-materials-editor").isVisible());
 assert.ok(await editor.isHidden());
 assert.deepEqual(errors,[],"No uncaught browser errors");
 console.log("PASS: Definitions right-click PML -> Artificial Domains PML1, Domain 1/3/4 selection, mouse toggle, label/cartesian/scaling settings, persisted editor and ribbon PML2");
}finally{await browser.close();}
