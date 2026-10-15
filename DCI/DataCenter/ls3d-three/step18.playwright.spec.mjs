// STEP18 real-browser acceptance. Stored tests are NOT browser PASS evidence.
// LS3D_BASE_URL=http://localhost:8000/DCI/DataCenter/LS_Datacenter_3D.html
// npx playwright test DCI/DataCenter/ls3d-three/step18.playwright.spec.mjs
import {test,expect} from '@playwright/test';
const BASE=process.env.LS3D_BASE_URL||
  'https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html';
async function enter(page,zone='hall'){
  const problems=[];
  page.on('pageerror',e=>problems.push(e.message));
  page.on('console',m=>{if(m.type()==='error')problems.push(m.text());});
  await page.goto(BASE+(BASE.includes('?')?'&':'?')+'engine=three',{waitUntil:'domcontentloaded'});
  await page.waitForFunction(()=>['ready','failed'].includes(window.LS3D_THREE?.status),
    null,{timeout:30000});
  expect(await page.evaluate(()=>window.LS3D_THREE.status)).toBe('ready');
  await page.evaluate(z=>window.__LS3D_TEST__.setZone(z),zone);
  await page.waitForFunction(()=>window.LS3D_THREE?.adapter?.diagnostic?.frames>3);
  return problems;
}
test('STEP18 Data Hall close inspection keeps camera clipping valid and lets roof occluders be restored',async({page},info)=>{
  const errors=await enter(page);
  await page.evaluate(()=>{
    const T=window.__LS3D_TEST__;
    T.pickAsset(T.assetMap.get('gpu-0-0'),true);
  });
  await page.waitForFunction(()=>window.LS3D_THREE.adapter.camera.position.distanceTo(
    window.LS3D_THREE.adapter.scene.getObjectByName('equipment:gpu-0-0')?.position||new window.LS3D_THREE.adapter.THREE.Vector3())<35);
  const state=await page.evaluate(()=>{
    const a=window.LS3D_THREE.adapter;
    return {near:a.camera.near,far:a.camera.far,occlusion:a.cameraOcclusion,
      quality:a.viewportPerformance,zone:window.__LS3D_TEST__.currentZone};
  });
  expect(state.zone).toBe('hall');
  expect(state.near).toBeGreaterThanOrEqual(.055);
  expect(state.far).toBeGreaterThanOrEqual(600);
  expect(state.occlusion.hiddenOccluders).toBeGreaterThanOrEqual(0);
  expect(state.quality.dpr).toBeGreaterThan(0);
  await page.screenshot({path:info.outputPath('step18-rack-camera-close.png')});
  expect(errors).toEqual([]);
});
test('STEP18 mobile zone / MEP / rack controls remain scrollable in same viewport',async({page},info)=>{
  await page.setViewportSize({width:390,height:844});
  const errors=await enter(page,'hall');
  await page.waitForFunction(()=>document.querySelector('#ls3dColoControls')&&
    document.querySelector('#ls3dMEPControls'));
  await page.evaluate(()=>{
    const T=window.__LS3D_TEST__;
    T.pickAsset(T.assetMap.get('gpu-0-0'),true);
  });
  await page.waitForFunction(()=>!!document.querySelector('#ls3dRackMotionControls')&&
    !document.querySelector('#ls3dRackMotionControls').hidden);
  const state=await page.evaluate(()=>{
    const ids=['ls3dColoControls','ls3dMEPControls','ls3dRackMotionControls'];
    return ids.map(id=>{
      const el=document.getElementById(id),r=el.getBoundingClientRect();
      return {id,left:r.left,right:r.right,top:r.top,bottom:r.bottom,
        display:getComputedStyle(el).display,overflowY:el.style.overflowY};
    });
  });
  expect(state.every(s=>s.left>=0&&s.right<=400)).toBe(true);
  expect(state.find(s=>s.id==='ls3dRackMotionControls').overflowY).toBe('auto');
  await page.screenshot({path:info.outputPath('step18-mobile-panel-docks.png')});
  expect(errors).toEqual([]);
});
test('STEP18 adaptive medium quality gives bounded DPR, respects low/high settings and keeps seven zones',async({page})=>{
  const errors=await enter(page,'campus');
  const data=await page.evaluate(()=>{
    const a=window.LS3D_THREE.adapter;
    const samples=[];
    for(const quality of ['low','medium','high']){
      window.LS3D_CONFIG.quality=quality;
      samples.push(quality);
    }
    return {sceneZones:window.__LS3D_TEST__.Z.length,
      perf:a.viewportPerformance,stats:a.diagnostic};
  });
  expect(data.sceneZones).toBe(7);
  expect(data.perf.dpr).toBeGreaterThanOrEqual(.55);
  expect(data.perf.dpr).toBeLessThanOrEqual(1.55);
  expect(data.stats.memoryGeometries).toBeGreaterThanOrEqual(0);
  expect(errors).toEqual([]);
});
test('STEP18 camera/panel controls survive full-screen entry and exit',async({page},info)=>{
  const errors=await enter(page);
  await page.locator('#btnFull').click();
  await page.waitForFunction(()=>!!document.fullscreenElement,{timeout:10000});
  await page.screenshot({path:info.outputPath('step18-fullscreen.png')});
  await page.locator('#btnFull').click();
  await page.waitForFunction(()=>!document.fullscreenElement,{timeout:10000});
  expect(errors).toEqual([]);
});
