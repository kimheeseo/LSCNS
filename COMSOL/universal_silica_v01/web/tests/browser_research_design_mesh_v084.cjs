/* Isolated research panel integration; provide a real Python mesh_preview.json. */
const {chromium}=require('playwright');
const fs=require('node:fs/promises');
(async()=>{
 const fixture=process.argv[2];if(!fixture)throw Error('usage: node browser_research_design_mesh_v084.cjs /absolute/mesh_preview.json');
 const browser=await chromium.launch({headless:true,...(process.env.CHROMIUM_PATH?{executablePath:process.env.CHROMIUM_PATH}: {})});
 try{
  const page=await browser.newPage({viewport:{width:1180,height:1000}});
  const errors=[];page.on('pageerror',e=>errors.push(e.message));
  const base=process.env.BASE_URL||'http://127.0.0.1:8769/';
  await page.route('**/__step12_isolated_harness__',route=>route.fulfill({status:200,contentType:'text/html',
    body:'<html><head><base href="/COMSOL/universal_silica_v01/web/"><style>*{box-sizing:border-box}body{margin:0}</style></head><body><main class="workspace"></main></body></html>'}));
  await page.goto(new URL('/__step12_isolated_harness__',base).href);
  await page.evaluate(()=>import('/COMSOL/universal_silica_v01/web/research_fiber_v1.mjs'));
  await page.waitForFunction(()=>document.querySelectorAll('#rf-preset option').length===25);
  const preview=JSON.parse(await fs.readFile(fixture,'utf8'));
  await page.locator('#rf-open').setInputFiles({name:'design.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(preview.input))});
  await page.locator('#rf-mesh-open').setInputFiles(fixture);
  await page.waitForFunction(()=>document.querySelector('#rf-mesh-status').textContent.includes('현재 입력과 일치'));
  if(!await page.locator('#rf-mesh-canvas').isVisible())throw Error('mesh canvas hidden');
  const status=await page.locator('#rf-mesh-status').textContent();
  if(process.env.SCREENSHOT_PATH)await page.locator('#rf-design-mesh-v084').screenshot({path:process.env.SCREENSHOT_PATH});
  await page.locator('#rf-json').fill(JSON.stringify({...preview.input,wavelength_nm:3400}));
  if(await page.locator('#rf-mesh-canvas').isVisible())throw Error('stale mesh remained visible');
  await page.setViewportSize({width:390,height:844});
  const overflow=await page.locator('#rf-design-mesh-v084').evaluate(el=>el.scrollWidth>el.clientWidth+2);
  const result={scope:'isolated research panel, not whole-app browser regression',preset_count:25,
    actual_python_mesh_import:status,stale_result_hidden:true,new_panel_mobile_overflow:overflow,page_errors:errors};
  if(process.env.RESULT_PATH)await fs.writeFile(process.env.RESULT_PATH,JSON.stringify(result,null,2));
  console.log(JSON.stringify(result));
  if(errors.length||overflow)process.exitCode=1;
 }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exitCode=1;});
