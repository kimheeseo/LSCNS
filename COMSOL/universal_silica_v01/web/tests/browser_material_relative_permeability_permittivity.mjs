/**
 * COMSOL Materials Basic Properties: Relative permeability/permittivity
 * Right-click -> Add to Material -> Status validation & duplication regression.
 * Requires Chromium/Playwright and local web app server (:8765).
 */
import assert from "node:assert/strict";
import {createRequire} from "node:module";
const {chromium}=createRequire(import.meta.url)("playwright");
const browser=await chromium.launch({headless:true,args:["--no-sandbox","--disable-dev-shm-usage"]});
try {
 const page=await browser.newPage({viewport:{width:1580,height:960}});
 const errors=[];
 page.on("pageerror",error=>errors.push(error.message));
 await page.goto(process.env.TEST_BASE_URL||"http://127.0.0.1:8765/COMSOL/universal_silica_v01/web/",{waitUntil:"load",timeout:90000});
 await page.locator('[data-dimension="2D"]').click();
 await page.locator('[data-interface="ewfd"]').click();
 await page.locator('[data-wizard="add"]').click();
 await page.locator('[data-wizard="select-added"]').click();
 await page.locator('[data-wizard="study"]').click();
 await page.locator('[data-study="mode"]').first().click();
 await page.locator('[data-wizard="done"]').click();
 await page.locator('summary[data-action="waveMaterials"]').click({button:"right"});
 await page.locator('#ls-materials-context [data-material-command="blank"]').click();
 await page.locator('#ls-material-properties-section > summary').click();
 const basic=page.locator('.ls-material-property-group[data-ls-property-group="Basic Properties"]');
 assert.equal(await basic.isVisible(),true);
 const mu=basic.locator('[data-ls-property="Relative permeability"]');
 const eps=basic.locator('[data-ls-property="Relative permittivity"]');
 assert.equal(await mu.isEnabled(),true);
 assert.equal(await eps.isEnabled(),true);
 await mu.click({button:"right"});
 assert.equal(await page.locator('#ls-material-property-context').isVisible(),true);
 await page.locator('#ls-material-property-add').click();
 assert.equal(await page.locator('#ls-material-property-context').isHidden(),true);
 assert.equal(await page.locator('[data-ls-material-basic-row="relative-permeability"]').count(),1);
 assert.equal(await page.locator('[data-ls-basic-state="relative-permeability"] .ls-material-invalid').count(),1);
 await eps.click({button:"right"});
 await page.locator('#ls-material-property-add').click();
 assert.equal(await page.locator('[data-ls-material-basic-row="relative-permittivity"]').count(),1);
 assert.equal(await page.locator('[data-ls-basic-state="relative-permittivity"] .ls-material-invalid').count(),1);
 const muInput=page.locator('[data-ls-basic-value="relative-permeability"]');
 const epsInput=page.locator('[data-ls-basic-value="relative-permittivity"]');
 await muInput.fill("1");
 assert.equal(await page.locator('[data-ls-basic-state="relative-permeability"] .ls-material-valid').count(),1);
 assert.equal(await page.locator('[data-ls-basic-state="relative-permittivity"] .ls-material-invalid').count(),1);
 await epsInput.fill("2.25");
 assert.equal(await page.locator('[data-ls-basic-state="relative-permittivity"] .ls-material-valid').count(),1);
 await epsInput.fill("unknown_parameter_xyz");
 assert.equal(await page.locator('[data-ls-basic-state="relative-permittivity"] .ls-material-invalid').count(),1);
 assert.equal(await page.locator('[data-ls-basic-state="relative-permeability"] .ls-material-valid').count(),1);
 await epsInput.fill("2.25");
 const value=await page.evaluate(()=>window.LSWaveMaterialsEditor.basicProperties("mat1"));
 assert.equal(value.length,2);
 assert.equal(value.find(x=>x.id==="relative-permittivity").value,2.25);
 assert.equal(value.find(x=>x.id==="relative-permeability").value,1);
 const material=await page.evaluate(()=>window.LSWaveMaterialsEditor.list()[0]);
 assert.ok(material.properties.includes("relative-permeability"));
 assert.ok(material.properties.includes("relative-permittivity"));
 assert.equal(material.basicPropertyValues["relative-permittivity"],"2.25");
 await mu.click({button:"right"});
 await page.locator('#ls-material-property-add').click();
 assert.equal(await page.locator('[data-ls-material-basic-row="relative-permeability"]').count(),1,"Duplicate Add to Material must not add a duplicate row");
 await page.locator('[data-material-command="duplicate"]').click();
 const duplicate=await page.evaluate(()=>window.LSWaveMaterialsEditor.list()[1]);
 assert.equal(duplicate.basicPropertyValues["relative-permittivity"],"2.25");
 await page.locator('[data-ls-basic-value="relative-permittivity"]').fill("3");
 const originals=await page.evaluate(()=>window.LSWaveMaterialsEditor.list());
 assert.equal(originals[0].basicPropertyValues["relative-permittivity"],"2.25","Duplicated material values must not alias the original");
 assert.equal(originals[1].basicPropertyValues["relative-permittivity"],"3");
 assert.deepEqual(errors,[]);
 console.log("PASS: Basic Properties right-click Add to Material, independent status and duplicate persistence");
} finally {
 await browser.close();
}
