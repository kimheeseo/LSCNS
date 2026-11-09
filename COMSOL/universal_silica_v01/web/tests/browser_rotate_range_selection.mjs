/** Browser regression for COMSOL-like Rotate Range dialog and Graphics input-object activation. */
import assert from "node:assert/strict";
import {createRequire} from "node:module";
const {chromium}=createRequire(import.meta.url)("playwright");
const browser=await chromium.launch({headless:true,args:["--no-sandbox","--disable-dev-shm-usage"]});
try{
 const page=await browser.newPage({viewport:{width:1450,height:900}});
 const errors=[];page.on("pageerror",e=>errors.push(e.message));
 await page.goto(process.env.TEST_BASE_URL||"http://127.0.0.1:8765/COMSOL/universal_silica_v01/web/",{waitUntil:"load",timeout:90000});
 await page.locator('[data-wizard="cancel"]').click();
 // Configure a small, off-axis Circle 2 so the 60-degree rotation is visible.
 await page.locator('[data-cad-node="c2"]').click();
 await page.locator("#ls-cad-radius").fill("2.5[um]");
 await page.locator("#ls-cad-x").fill("6[um]");
 await page.locator('summary[data-action="geometry"]').click({button:"right"});
 await page.locator('#ls-cad-menu > ul > li.has-children > button').filter({hasText:"Transforms"}).hover();
 await page.locator('#ls-cad-menu li.has-children > ul [data-cmd="rotate"]').click();
 assert.equal(await page.locator("#ls-cad-subtitle").textContent(),"Rotate");
 assert.equal(await page.locator("#ls-cad-angle").inputValue(),"range(0,60,300)");
 assert.ok(await page.locator("#ls-cad-range-open").isVisible(),"Angle's Range button missing");
 const input=page.locator('#ls-cad-inputs [data-cad-input="c2"]');
 assert.ok(await input.isChecked(),"The previous Circle 2 is initially selected as Rotate input");
 await page.screenshot({path:"/tmp/comsol_rotate_initial_settings.png"});
 // Range popup matches user reference: method=Step; Start/Step/Stop; Function=None; Replace/Add/Cancel.
 await page.locator("#ls-cad-range-open").click();
 assert.ok(await page.locator("#ls-cad-range-dialog").isVisible());
 assert.equal(await page.locator("#ls-cad-range-method").inputValue(),"Step");
 assert.equal(await page.locator("#ls-cad-range-start").inputValue(),"0");
 assert.equal(await page.locator("#ls-cad-range-step").inputValue(),"60");
 assert.equal(await page.locator("#ls-cad-range-stop").inputValue(),"300");
 assert.equal(await page.locator("#ls-cad-range-function").inputValue(),"none");
 await page.screenshot({path:"/tmp/comsol_rotate_range_dialog.png"});
 await page.locator("#ls-cad-range-cancel").click();
 assert.ok(await page.locator("#ls-cad-range-dialog").isHidden());
 assert.equal(await page.locator("#ls-cad-angle").inputValue(),"range(0,60,300)","Cancel keeps original angle");
 // Replace the default range with 0,60,120,180 = four transformed positions.
 await page.locator("#ls-cad-range-open").click();
 await page.locator("#ls-cad-range-stop").fill("180");
 await page.locator("#ls-cad-range-replace").click();
 assert.equal(await page.locator("#ls-cad-angle").inputValue(),"range(0,60,180)");
 assert.equal(await page.locator('#ls-cad-view [data-cad-shape="rot1"]').count(),4);
 // Add 240 and 300-degree positions without removing the preceding range.
 await page.locator("#ls-cad-range-open").click();
 await page.locator("#ls-cad-range-start").fill("240");
 await page.locator("#ls-cad-range-step").fill("60");
 await page.locator("#ls-cad-range-stop").fill("300");
 await page.locator("#ls-cad-range-add").click();
 assert.equal(await page.locator("#ls-cad-angle").inputValue(),"range(0,60,180),range(240,60,300)");
 assert.equal(await page.locator('#ls-cad-view [data-cad-shape="rot1"]').count(),6);
 // Invalid zero step stays in the dialog and does not change the rotation.
 await page.locator("#ls-cad-range-open").click();
 await page.locator("#ls-cad-range-step").fill("0");
 await page.locator("#ls-cad-range-replace").click();
 assert.ok(await page.locator("#ls-cad-range-dialog").isVisible());
 assert.ok((await page.locator("#ls-cad-range-error").textContent()).includes("step"));
 await page.locator("#ls-cad-range-cancel").click();
 assert.equal(await page.locator("#ls-cad-angle").inputValue(),"range(0,60,180),range(240,60,300)");
 // Activating the input selector preserves Rotate 1 settings and paints c2 blue.
 const button=page.locator("#ls-cad-activate-selection");
 await button.click();
 assert.equal(await button.getAttribute("aria-pressed"),"true");
 assert.ok(await input.isChecked());
 const original=page.locator('#ls-cad-view [data-cad-shape="c2"]');
 assert.equal(await original.count(),1);
 assert.ok(await original.evaluate(el=>el.classList.contains("is-input-selected")));
 assert.equal(await original.locator(".ls-cad-rendered-surface").evaluate(el=>getComputedStyle(el).fill),"rgb(194, 197, 243)");
 await page.screenshot({path:"/tmp/comsol_rotate_c2_active_selected.png"});
 // When selection mode is active, clicking the source shape changes Rotate's inputs,
 // not the selected Model Builder feature. No input remains a valid edit state.
 await original.locator(".ls-cad-rendered-surface circle").click({force:true});
 assert.equal(await page.locator("#ls-cad-subtitle").textContent(),"Rotate");
 assert.equal(await input.isChecked(),false);
 assert.equal(await page.locator('#ls-cad-view [data-cad-shape="rot1"]').count(),0);
 // Click c2 again to add it and confirm the blue highlight returns.
 await page.locator('#ls-cad-view [data-cad-shape="c2"] .ls-cad-rendered-surface circle').click({force:true});
 assert.ok(await input.isChecked());
 assert.ok(await page.locator('#ls-cad-view [data-cad-shape="c2"]').evaluate(el=>el.classList.contains("is-input-selected")));
 await button.click();
 assert.equal(await button.getAttribute("aria-pressed"),"false");
 assert.ok(!await page.locator('#ls-cad-view [data-cad-shape="c2"]').evaluate(el=>el.classList.contains("is-input-selected")));
 assert.deepEqual(errors,[],"No uncaught browser errors");
 console.log("PASS: Rotate Range Step dialog Replace/Add/Cancel/validation, c2 Activate Selection blue Graphics highlight and click-to-select");
}finally{await browser.close();}
