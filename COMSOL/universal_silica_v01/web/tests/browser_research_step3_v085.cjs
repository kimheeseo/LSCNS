/* Full deployed workbench navigation + actual Python result input/field integration. */
const {chromium}=require('playwright');
const fs=require('node:fs/promises');
(async()=>{
 const browser=await chromium.launch({headless:true,...(process.env.CHROMIUM_PATH?{executablePath:process.env.CHROMIUM_PATH}:{})});
 try{
  const page=await browser.newPage({viewport:{width:1280,height:1000}}),errors=[];
  page.on('pageerror',e=>errors.push(e.message));
  const base=process.env.BASE_URL||'http://127.0.0.1:8769/';
  // Actual index and CSS: HCF/vector selector modes previously hid the research panel.
  await page.goto(new URL('COMSOL/universal_silica_v01/web/#research-step3-v085',base).href);
  await page.waitForFunction(()=>document.querySelectorAll('#rf-preset option').length===25);
  await page.waitForFunction(()=>document.querySelectorAll('#rf-step3-table tr').length===3);
  for(const mode of ['hcf-selection','vector-selection']){
    await page.evaluate(m=>document.body.classList.add(m),mode);
    await page.locator('#ls-step3-entry').click();
    if(!await page.locator('#research-step3-v085').isVisible())throw Error('STEP3 navigation hidden in '+mode);
    await page.evaluate(m=>document.body.classList.remove(m),mode);
  }
  await page.locator('#rf-step3-guided').click();
  await page.waitForFunction(()=>document.querySelectorAll('#rf-mode-select option').length===2);
  if(!await page.locator('#rf-field-panel').isVisible())throw Error('actual field panel hidden');
  if(!(await page.locator('#rf-verdict').textContent()).includes('현재 입력과 일치'))throw Error('input trace missing');
  await page.locator('#rf-mode-select').selectOption('1');
  const first=await page.locator('#rf-results').textContent();
  if(!first.includes('Re(n_eff)'))throw Error('selected mode metrics missing');
  const config=JSON.parse(await page.locator('#rf-json').inputValue());config.wavelength_nm=1549;
  await page.locator('#rf-json').fill(JSON.stringify(config));
  if(await page.locator('#rf-field-panel').isVisible())throw Error('stale fields remained visible');
  if(await page.locator('#rf-mode-select option').count())throw Error('stale mode selector remained');
  const fixture=await fs.readFile('COMSOL/universal_silica_v01/fiber_research_v1/benchmarks/step3_v085/guided_mode_result.json');
  await page.locator('#rf-result-open').setInputFiles({name:'mode_result.json',mimeType:'application/json',buffer:fixture});
  await page.waitForFunction(()=>document.querySelector('#rf-status').textContent.includes('불일치'));
  if(await page.locator('#rf-field-panel').isVisible())throw Error('mismatched result displayed');
  await page.locator('#rf-step3-radiative').click();
  await page.waitForFunction(()=>document.querySelectorAll('#rf-mode-select option').length===2);
  if(!(await page.locator('#rf-verdict').textContent()).includes('V5=PASS_SCOPED'))throw Error('computed radiative benchmark missing');
  await page.setViewportSize({width:390,height:844});
  const overflow=await page.locator('#research-step3-v085').evaluate(e=>e.scrollWidth>e.clientWidth+2);
  if(overflow)throw Error('STEP3 panel mobile overflow');
  if(process.env.SCREENSHOT_PATH)await page.locator('#research-step3-v085').screenshot({path:process.env.SCREENSHOT_PATH});
  const report={scope:'actual workbench index, navigation, CSS and STEP3 panel',preset_count:25,
    hidden_selector_navigation:true,actual_python_guided_and_radiative_examples:true,independent_mode_selector:true,
    stale_result_hidden:true,mismatched_input_rejected:true,mobile_panel_overflow:overflow,page_errors:errors};
  if(process.env.RESULT_PATH)await fs.writeFile(process.env.RESULT_PATH,JSON.stringify(report,null,2));
  console.log(JSON.stringify(report));if(errors.length)process.exitCode=1;
 }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exitCode=1;});
