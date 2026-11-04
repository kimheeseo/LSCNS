/** Full browser regression for COMSOL-inspired presentation layer. */
import assert from 'node:assert/strict';
import {createRequire} from 'node:module';
const require=createRequire(import.meta.url);
const {chromium}=require('playwright');
const base=process.env.TEST_BASE_URL||'http://127.0.0.1:8765/COMSOL/universal_silica_v01/web/';
const browser=await chromium.launch({headless:true,args:['--no-sandbox','--disable-dev-shm-usage']});
try{
 const page=await browser.newPage({viewport:{width:1277,height:695},deviceScaleFactor:1});
 const errors=[];
 page.on('pageerror',err=>errors.push(err.message));
 await page.goto(base,{waitUntil:'load',timeout:90000});
 await page.waitForSelector('.ls-model-tree',{timeout:30000});
 await page.waitForSelector('.ls-ribbon');
 await page.waitForSelector('.ls-study-pane');
 await page.waitForSelector('.ls-graphics-head');
 await page.waitForSelector('.ls-output-console');
 await page.waitForSelector('#hcf-designer',{timeout:30000});
 await page.waitForSelector('#full-vector-tool',{timeout:30000});
 await page.waitForSelector('#fiber-3d',{timeout:30000});
 const layouts=await page.evaluate(()=>['.ls-model-tree','.sidebar','main.workspace','.ls-study-pane','.ls-output-console'].map(s=>({s,rect:document.querySelector(s).getBoundingClientRect().toJSON()})));
 const [tree,settings,graphics,study,dock]=layouts.map(x=>x.rect);
 assert.ok(tree.x<settings.x && settings.x<graphics.x && graphics.x<study.x,'Model Builder → Settings → Graphics → Add Study pane order');
 assert.ok(dock.y>graphics.y,'Messages dock under graphics');
 assert.ok(graphics.width>360,'Graphics pane must remain useful at desktop reference width');
 const panelBg=await page.locator('.sidebar').evaluate(el=>getComputedStyle(el).backgroundColor);
 assert.equal(panelBg,'rgb(255, 255, 255)','Settings pane should be white/light');
 const allIds=await page.evaluate(()=>[...document.querySelectorAll('[id]')].map(x=>x.id));
 assert.equal(new Set(allIds).size,allIds.length,'Original document has no duplicate IDs after skin insertion');
 const preserved=['preset','profile','wavelength_nm','core_radius_um','delta_n_core','calculateBtn','sweepBtn','meshBtn','crossSection','modeChart','mfd','aeff','neff','fv-example','fem-import','hcf-coreDiameterUm','pmd-run'];
 for(const id of preserved)assert.ok(await page.locator('#'+id).count()===1,`Original functional id still present: ${id}`);
 await page.screenshot({path:'/tmp/ls_comsol_theme_desktop.png'});
 // Ribbon tabs show working commands rather than static decorative labels.
 await page.locator('[data-tab="geometry"]').click();
 assert.ok(await page.locator('#ls-ribbon-tools [data-action="hcf"]').isVisible());
 await page.locator('[data-tab="home"]').click();
 assert.ok(await page.locator('#ls-ribbon-tools [data-action="compute"]').isVisible());
 // The model tree uses the original selection/compute handlers.
 await page.locator('.ls-model-tree [data-select-preset="comsol_step"]').click();
 await page.waitForFunction(()=>document.querySelector('#comsol-reference-status')?.textContent.startsWith('PASS'),{timeout:30000});
 assert.equal(await page.locator('#clad_material').inputValue(),'fixed');
 const benchmarkNeff=Number(await page.locator('#neff').textContent());
 const benchmarkPct=Number((await page.locator('#comsol-reference-pct').textContent()).replace('%',''));
 assert.ok(Math.abs(benchmarkNeff-1.4444)<0.00005,'COMSOL same-input neff stays valid');
 assert.ok(benchmarkPct<5,'5% relative benchmark threshold');
 assert.ok((await page.locator('#ls-console-text').textContent()).length>15,'Live solver messages visible');
 // Existing HCF 2D geometry controls and separate complex FEM panel still work.
 await page.locator('.ls-model-tree [data-select-preset="hcf"]').click();
 assert.equal(await page.locator('#hcf-sidebar-controls').isVisible(),true);
 assert.equal(await page.locator('#hcf-designer').isVisible(),true);
 assert.equal(await page.locator('#hcf-fem').isVisible(),true);
 await page.locator('#hcf-quick-core').fill('40');
 await page.waitForFunction(()=>document.querySelector('#hcf-coreDiameterUm')?.value==='40');
 // Existing Full-vector Colab JSON displays the computed vector result.
 await page.locator('.ls-model-tree [data-select-preset="vector"]').click();
 assert.equal(await page.locator('#full-vector-tool').isVisible(),true);
 await page.locator('#fv-example').click();
 await page.waitForFunction(()=>document.querySelector('#fv-neff')?.textContent.trim()!=='—',{timeout:30000});
 assert.ok(Number(await page.locator('#fv-neff').textContent())>1);
 // Return to normal Scalar FEM and preserve Settings input and compute buttons.
 await page.locator('.ls-model-tree [data-select-preset="g654"]').click();
 assert.equal(await page.locator('#standard-fiber-options').isVisible(),true);
 assert.equal(await page.locator('#calculateBtn').isVisible(),true);
 await page.locator('#calculateBtn').click();
 await page.waitForFunction(()=>Number(document.querySelector('#neff')?.textContent)>1.4,{timeout:30000});
 // Study pane is genuinely interactive.
 await page.locator('#ls-add-study').click();
 assert.equal(await page.locator('#ls-study-more').isVisible(),true);
 await page.locator('#ls-console-collapse').click();
 assert.equal(await page.locator('.layout').evaluate(el=>el.classList.contains('ls-console-minimized')),true);
 await page.locator('#ls-console-collapse').click();
 assert.equal(await page.locator('.layout').evaluate(el=>el.classList.contains('ls-console-minimized')),false);
 // Main functionality retained in compact view, no desktop-only scroll traps.
 await page.setViewportSize({width:390,height:844});
 assert.equal(await page.locator('.ls-model-tree').isVisible(),false);
 assert.equal(await page.locator('.sidebar').isVisible(),true);
 assert.equal(await page.locator('main.workspace').isVisible(),true);
 await page.screenshot({path:'/tmp/ls_comsol_theme_mobile.png',fullPage:false});
 assert.deepEqual(errors,[],'No JavaScript runtime exceptions');
 console.log('PASS: COMSOL-inspired 4-pane UI, ribbon, live model tree, neff benchmark '+benchmarkPct.toFixed(6)+'%, HCF, vector JSON, Scalar FEM, Study panel, log dock and mobile');
} finally {await browser.close();}
