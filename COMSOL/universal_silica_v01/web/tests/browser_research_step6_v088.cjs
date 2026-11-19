const {chromium}=require('playwright');const fs=require('node:fs/promises');
(async()=>{const browser=await chromium.launch({headless:true,...(process.env.CHROMIUM_PATH?{executablePath:process.env.CHROMIUM_PATH}:{})});try{
 const page=await browser.newPage({viewport:{width:1280,height:1000}}),errors=[];page.on('pageerror',e=>errors.push(e.message));
 await page.goto(new URL('COMSOL/universal_silica_v01/web/#research-step6-v088',process.env.BASE_URL||'http://127.0.0.1:8769/').href);
 await page.waitForFunction(()=>document.querySelectorAll('#rf-preset option').length===25);
 const wizard=page.locator('#ls-startup-wizard');if(await wizard.isVisible())await wizard.locator('[data-wizard="cancel"]').click();
 await page.locator('#r6-tab-bend').waitFor();
 for(const mode of ['hcf-selection','vector-selection']){await page.evaluate(m=>document.body.classList.add(m),mode);await page.locator('#ls-step6-entry').click();await page.locator('#research-step6-v088').waitFor({state:'visible'});await page.evaluate(m=>document.body.classList.remove(m),mode);}
 await page.locator('#r6-example').click();await page.waitForFunction(()=>document.querySelectorAll('#r6-table tbody tr').length===4);
 if(!(await page.locator('#r6-table').textContent()).includes('TRACKED_SUBSPACE_CANDIDATE'))throw Error('bend numeric evidence absent');
 await page.locator('#r6-export-config').click();
 await page.locator('#r6-directions').fill('0,45');if(await page.locator('#r6-table tbody tr').count())throw Error('stale bend result');
 const counts={};for(const [kind,count]of [['pm',1],['mcf',6],['splice',1],['sweep',4]]){await page.locator('#r6-tab-'+kind).click();await page.locator('#r6-example').click();await page.waitForFunction(n=>document.querySelectorAll('#r6-table tbody tr').length===n,count);counts[kind]=await page.locator('#r6-table tbody tr').count();}
 if(!(await page.locator('#r6-table').textContent()).includes('COMPUTED'))throw Error('actual FEM grid absent');
 await page.locator('#rf-json').fill((await page.locator('#rf-json').inputValue()).replace('"wavelength_nm": 1550','"wavelength_nm": 1551'));if(await page.locator('#r6-table tbody tr').count())throw Error('stale design result');
 await page.locator('#r6-tab-dgd').click();if(!(await page.locator('#r6-help').textContent()).includes('시연용'))throw Error('PMD demonstration label absent');
 await page.locator('#r6-export-template').click();const pending=page.waitForEvent('download');await page.locator('#r6-export-config').click();const downloaded=JSON.parse(await fs.readFile(await (await pending).path(),'utf8'));if(downloaded.operation!=='dgd'||downloaded.schema_version!=='1.0.0')throw Error('versioned DGD input absent');
 await page.setViewportSize({width:390,height:844});await page.locator('#r6-tab-mcf').click();await page.locator('#r6-example').click();await page.waitForFunction(()=>document.querySelectorAll('#r6-table tbody tr').length===6);
 const overflow=await page.locator('#research-step6-v088').evaluate(e=>e.scrollWidth>e.clientWidth+2);if(overflow)throw Error('STEP6 mobile overflow');
 const report={scope:'actual workbench STEP6 v088',bend_actual_FEM:true,actual_PM_MCF_splice_grid_rows:counts,stale_inputs_cleared:true,hcf_vector_navigation:true,versioned_execution_config:true,PMD_demonstration_visible:true,mobile_overflow:overflow,page_errors:errors};
 if(process.env.RESULT_PATH)await fs.writeFile(process.env.RESULT_PATH,JSON.stringify(report,null,2));console.log(JSON.stringify(report));if(errors.length)process.exitCode=1;
 }finally{await browser.close();}})().catch(e=>{console.error(e);process.exit(1)});
