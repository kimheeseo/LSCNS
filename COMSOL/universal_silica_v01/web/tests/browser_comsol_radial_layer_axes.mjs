/**
 * Regression driven by the user's COMSOL screenshot:
 * Parameters 1: rb=10[um], wl=1.45[um], neffMax=1.4454, nbg=1.45
 * Geometry Circle 1: radius=rb+wlr, Layer 1 thickness=wlr
 * Expected: two concentric circles (outer~22.56um, inner=10um),
 * labeled real-distance X/Y axes and changes after editing parameters.
 */
import assert from "node:assert/strict";
import {createRequire} from "node:module";
const {chromium}=createRequire(import.meta.url)("playwright");
const browser=await chromium.launch({headless:true,args:["--no-sandbox","--disable-dev-shm-usage"]});
const near=(a,b,tol=1e-8)=>Math.abs(a-b)<tol;
try{
 const page=await browser.newPage({viewport:{width:1500,height:920}});
 const errors=[];page.on("pageerror",e=>errors.push(e.message));
 await page.goto(process.env.TEST_BASE_URL||"http://127.0.0.1:8765/COMSOL/universal_silica_v01/web/",{waitUntil:"load",timeout:90000});
 await page.locator('[data-dimension="2D"]').click();
 await page.locator('[data-interface="ewfd"]').click();
 await page.locator('[data-wizard="add"]').click();
 await page.locator('[data-wizard="select-added"]').click();
 await page.locator('[data-wizard="study"]').click();
 await page.locator('[data-study="mode"]').first().click();
 await page.locator('[data-wizard="done"]').click();
 await page.locator('.ls-model-tree [data-action="settings"]').click();
 const defaults=await page.locator("#ls-parameter-rows tr").count();
 assert.equal(defaults,9,"Screenshot's nine parameters should load automatically");
 assert.equal(await page.locator("#ls-parameter-rows tr").nth(2).locator('[data-field="expression"]').inputValue(),"10[um]");
 const actualWlr=await page.evaluate(()=>window.LSWaveParameterService.calculate("wlr"));
 const actualOuter=await page.evaluate(()=>window.LSWaveParameterService.calculate("rb+wlr"));
 assert.ok(near(actualWlr,12.564e-6,8e-9),"wlr from screenshot must be ~12.564 um: "+actualWlr);
 assert.ok(near(actualOuter,22.564e-6,8e-9),"outer R ~22.564 um: "+actualOuter);
 await page.locator('summary[data-action="geometry"]').click({button:"right"});
 await page.locator('#ls-cad-menu > ul > li > [data-cmd="circle"]').click();
 assert.equal(await page.locator("#ls-cad-radius").inputValue(),"rb+wlr");
 await page.locator("#ls-cad-layers-section summary").click();
 await page.locator("#ls-cad-layer").fill("wlr");
 await page.locator("#ls-cad-build-selected").click();
 const shape=page.locator('#ls-cad-view [data-cad-shape="c1"]');
 assert.equal(await shape.count(),1,"One circle feature, with layers");
 assert.equal(await shape.locator(".ls-cad-rendered-surface circle").count(),1,"Outer circular boundary");
 assert.equal(await shape.locator("circle.ls-cad-layer-boundary").count(),1,"Inner circular boundary");
 assert.equal(await shape.locator("line.ls-cad-layer-boundary").count(),4,"Four COMSOL-style radial guides");
 const outer=Number(await shape.getAttribute("data-outer-radius-m"));
 const inner=Number(await shape.getAttribute("data-inner-radius-m"));
 assert.ok(near(outer,actualOuter,1e-11),"Outer radius expression not evaluated");
 assert.ok(near(inner,10e-6,1e-11),"Layer thickness should leave inner rb=10 um");
 const legend=await page.locator("#ls-cad-legend").textContent();
 assert.ok(legend.includes("22.")&&legend.includes("10 µm"),"Both radii must be displayed in legend: "+legend);
 const yLabels=await page.locator('#ls-cad-view [data-axis="y"]').allTextContents();
 const xLabels=await page.locator('#ls-cad-view [data-axis="x"]').allTextContents();
 assert.ok(yLabels.includes("25")&&yLabels.includes("-25"),"Y axis should have physical -25,+25 micrometer ticks: "+yLabels);
 assert.ok(xLabels.length>=5&&xLabels.includes("0"),"X axis has numeric physical ticks: "+xLabels);
 assert.ok((await page.locator('[data-axis-units="true"]').textContent()).includes("10"),"Scientific unit exponent label");
 await page.screenshot({path:"/tmp/comsol_reference_two_circles_axes.png"});
 // Increasing Layer 1 by 2 um must reduce the inner radius but not outer radius.
 await page.locator("#ls-cad-layer").fill("wlr+2[um]");
 await page.locator("#ls-cad-build-selected").click();
 assert.ok(near(Number(await shape.getAttribute("data-inner-radius-m")),8e-6,1e-11));
 assert.ok(near(Number(await shape.getAttribute("data-outer-radius-m")),actualOuter,1e-11));
 // Returning to wlr restores inner rb without needing to add another Circle.
 await page.locator("#ls-cad-layer").fill("wlr");
 await page.locator("#ls-cad-build-selected").click();
 assert.ok(near(Number(await shape.getAttribute("data-inner-radius-m")),10e-6,1e-11));
 // Edit rb in Global Definitions and check both radii update accordingly.
 await page.locator('.ls-model-tree [data-action="settings"]').click();
 await page.locator("#ls-parameter-rows tr").nth(2).locator('[data-field="expression"]').fill("12[um]");
 await page.locator('[data-cad-node="c1"]').click();
 assert.ok(near(Number(await shape.getAttribute("data-inner-radius-m")),12e-6,1e-11));
 assert.ok(near(Number(await shape.getAttribute("data-outer-radius-m")),actualOuter+2e-6,1e-11));
 // Reset to screenshot parameters can be performed at any time.
 await page.locator('.ls-model-tree [data-action="settings"]').click();
 await page.locator("#ls-params-reset").click();
 assert.equal(await page.locator("#ls-parameter-rows tr").nth(2).locator('[data-field="expression"]').inputValue(),"10[um]");
 await page.locator('[data-cad-node="c1"]').click();
 assert.ok(near(Number(await shape.getAttribute("data-inner-radius-m")),10e-6,1e-11));
 assert.deepEqual(errors,[],"No uncaught page errors");
 console.log("PASS: screenshot-exact rb+wlr outer / wlr layer gives two concentric circles (22.564/10 um) and physical X/Y ticks with dynamic parameter updates");
}finally{await browser.close();}
