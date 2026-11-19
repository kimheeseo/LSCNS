const {chromium}=require('playwright');const fs=require('node:fs/promises');
(async()=>{const browser=await chromium.launch({headless:true,...(process.env.CHROMIUM_PATH?{executablePath:process.env.CHROMIUM_PATH}:{})});try{
 const page=await browser.newPage({viewport:{width:1280,height:1000}}),errors=[];page.on('pageerror',e=>errors.push(e.message));
 await page.goto(new URL('COMSOL/universal_silica_v01/web/#research-step4-v080',process.env.BASE_URL||'http://127.0.0.1:8769/').href);
 await page.waitForFunction(()=>document.querySelectorAll('#rf-preset option').length===25);
 const wizard=page.locator('#ls-startup-wizard');if(await wizard.isVisible())await wizard.locator('[data-wizard="cancel"]').click();
 await page.locator('#r4-example').waitFor();
 for(const mode of ['hcf-selection','vector-selection']){
  await page.evaluate(m=>document.body.classList.add(m),mode);await page.locator('#ls-step4-entry').click();await page.locator('#research-step4-v080').waitFor({state:'visible'});await page.evaluate(m=>document.body.classList.remove(m),mode);
 }
 await page.locator('#r4-example').click();await page.waitForFunction(()=>document.querySelectorAll('#r4-table tbody tr').length===2);
 if(!(await page.locator('#r4-status').textContent()).includes('실제 PEC'))throw Error('STEP4 example trace missing');
 await page.locator('#rf-json').fill((await page.locator('#rf-json').inputValue()).replace('1510','1511'));
 if(await page.locator('#r4-table tbody tr').count())throw Error('stale STEP4 result remained');
 await page.locator('#ls-step5-entry').click();await page.waitForFunction(()=>document.querySelector('#pc5-cutoff')?.value==='5');
 await page.locator('#pc5-tab-bands').click();await page.waitForFunction(()=>document.querySelector('#pc5-band-note').textContent.includes('벡터 Maxwell'));
 if((await page.locator('#pc5-candidate option').count())<1)throw Error('computed gap selector absent');
 await page.locator('#pc5-tab-defect').click();
 const pending=page.waitForEvent('download');await page.locator('#pc5-export-defect').click();const download=await pending;
 const design=JSON.parse(await fs.readFile(await download.path(),'utf8'));
 if(!Number.isFinite(design.target_neff)||design.family!=='pbgf')throw Error('finite FEM target missing');
 const current=JSON.parse(await page.locator('#rf-json').inputValue());if(current.target_neff!==design.target_neff)throw Error('finite handoff not applied to shared editor');
 await page.locator('#pc5-load-fem-example').click();await page.waitForFunction(()=>document.querySelectorAll('#rf-mode-select option').length===4&&document.querySelectorAll('#r4-table tbody tr').length===4);
 await page.locator('#ls-step5-entry').click();await page.locator('#pc5-tab-design').click();await page.locator('#pc5-hole').fill('0.85');await page.locator('#pc5-hole').dispatchEvent('change');
 if(!(await page.locator('#pc5-status').textContent()).includes('무효'))throw Error('stale Bloch results remained');
 await page.locator('#pc5-backend').selectOption('scalar_fdfd');const inputDownload=page.waitForEvent('download');await page.locator('#pc5-export-input').click();const input=JSON.parse(await fs.readFile(await (await inputDownload).path(),'utf8'));
 if(input.schema!=='LSCNS_STEP5_BLOCH_INPUT_V083')throw Error('legacy scalar input no longer supported');
 await page.setViewportSize({width:390,height:844});
 const overflow=await page.locator('#research-step4-v080').evaluate(e=>e.scrollWidth>e.clientWidth+2)||await page.locator('#research-step5-v083').evaluate(e=>e.scrollWidth>e.clientWidth+2);
 if(overflow)throw Error('new panel mobile overflow');
 const report={scope:'actual workbench STEP4/5',step4_actual_FEM_metrics_and_dispersion:true,step4_stale_guard:true,hcf_vector_navigation:true,
  vector_Bloch_example:true,computed_gap_handoff:true,actual_finite_FEM_and_metrics_import:true,target_neff:design.target_neff,shared_design_applied:true,legacy_scalar_export:true,step5_stale_guard:true,mobile_overflow:overflow,page_errors:errors};
 if(process.env.RESULT_PATH)await fs.writeFile(process.env.RESULT_PATH,JSON.stringify(report,null,2));console.log(JSON.stringify(report));if(errors.length)process.exitCode=1;
 }finally{await browser.close();}})().catch(e=>{console.error(e);process.exitCode=1;});
