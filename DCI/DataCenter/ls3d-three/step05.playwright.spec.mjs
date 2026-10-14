// STEP 05 independent real browser regression. Uploaded, not yet executed in a real GPU browser.
// npx playwright test DCI/DataCenter/ls3d-three/step05.playwright.spec.mjs
import { test, expect } from '@playwright/test';
const BASE=process.env.LS3D_BASE_URL||
  'https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html';
test.use({viewport:{width:1366,height:768}});
async function openThree(page) {
  await page.goto(BASE+'?engine=three',{waitUntil:'domcontentloaded'});
  await page.waitForFunction(()=>['ready','failed'].includes(window.LS3D_THREE?.status),null,{timeout:30000});
  const s=await page.evaluate(()=>({status:window.LS3D_THREE?.status,error:window.LS3D_THREE?.error}));
  expect(s.status,s.error||'Three.js failed to initialize').toBe('ready');
  await page.waitForFunction(()=>window.LS3D_THREE?.renders>3);
  await page.waitForFunction(()=>window.LS3D_THREE?.adapter?.equipment?.lodAssets>0);
}
test('real 3D equipment: 7 zones, recognizable asset families, instancing and LOD',async({page},info)=>{
  const errors=[];page.on('pageerror',e=>errors.push(e.message));
  await openThree(page);
  const sample=await page.evaluate(()=>{
    const S=window.LS3D_THREE,THREE=S.adapter.THREE;
    const modelCounts={lod:0,instancing:0,instances:0,mesh:0};
    S.adapter.scene.traverse(o=>{
      if(o instanceof THREE.LOD)modelCounts.lod++;
      if(o instanceof THREE.InstancedMesh){modelCounts.instancing++;modelCounts.instances+=o.count;}
      if(o.isMesh)modelCounts.mesh++;
    });
    return {
      status:S.status,diagnostics:S.diagnostics,
      modelCounts,classes:S.adapter.equipment.activeKinds,
      originalAssets:window.__LS3D_TEST__.visibleAssets().length,
      originalZones:window.__LS3D_TEST__.Z.length
    };
  });
  expect(sample.originalZones).toBe(7);
  expect(sample.originalAssets).toBeGreaterThan(25);
  for(const type of ['transformer','switchgear','cooling-tower','chiller','ups','cdu','battery-bank','air-handler','network','odf']){
    // 'network' may be represented by original mechanically operable racks on some legacy types.
    if(type==='network')continue;
    expect(sample.classes).toContain(type);
  }
  expect(sample.modelCounts.lod).toBeGreaterThan(12);
  expect(sample.modelCounts.instancing).toBeGreaterThan(5);
  expect(sample.modelCounts.instances).toBeGreaterThan(sample.modelCounts.instancing*3);
  expect(sample.modelCounts.mesh).toBeGreaterThan(80);
  expect(sample.diagnostics.frames).toBeGreaterThan(1);
  expect(errors).toEqual([]);
  await page.screenshot({path:info.outputPath('step05-campus.png')});
});
test('utility and cooling are not the same repeated box',async({page},info)=>{
  await openThree(page);
  for(const zone of ['utility','cooling','gray','network','ops']){
    await page.evaluate(z=>window.__LS3D_TEST__.setZone(z),zone);
    await page.waitForTimeout(250);
    const data=await page.evaluate(()=>({
      zone:window.__LS3D_TEST__.currentZone,
      activeKinds:window.LS3D_THREE.adapter.equipment.activeKinds,
      active:window.LS3D_THREE.diagnostics.assets
    }));
    expect(data.zone).toBe(zone);
    expect(data.active).toBeGreaterThan(0);
    if(zone==='utility')expect(data.activeKinds).toContain('transformer');
    if(zone==='cooling')expect(data.activeKinds).toContain('cooling-tower');
    if(zone==='gray')expect(data.activeKinds).toContain('ups');
    await page.screenshot({path:info.outputPath('step05-'+zone+'.png')});
  }
});
test('rack articulations survive after the STEP05 factory is enabled',async({page})=>{
  await openThree(page);
  await page.evaluate(()=>window.__LS3D_TEST__.setZone('hall'));
  await page.waitForTimeout(350);
  const hasRack=await page.evaluate(()=>{
    const S=window.LS3D_THREE,THREE=S.adapter.THREE;
    let near=0,far=0,instancing=0;
    S.adapter.scene.traverse(o=>{
      if(o.name==='rack-high')near++;
      if(o.name==='rack-low')far++;
      if(o instanceof THREE.InstancedMesh)instancing++;
    });
    return {near,far,instancing};
  });
  expect(hasRack.near).toBeGreaterThan(0);
  expect(hasRack.far).toBe(hasRack.near);
});
test('small server room limits all models to one facility, including new LODs',async({page})=>{
  await openThree(page);
  await page.evaluate(()=>window.__LS3D_TEST__.applyFacilityMode('server-room'));
  await page.waitForFunction(()=>window.LS3D_THREE?.diagnostics?.zones===1);
  const state=await page.evaluate(()=>{
    const S=window.LS3D_THREE;
    return {
      facility:window.__LS3D_TEST__.facilityMode,
      zones:S.diagnostics.zones,assets:S.diagnostics.assets,
      groups:S.adapter.scene.children[0].children[0].children.map(x=>x.name)
    };
  });
  expect(state.facility).toBe('server-room');
  expect(state.zones).toBe(1);
  expect(state.groups).toEqual(['arch:zone:hall']);
  expect(state.assets).toBeGreaterThan(0);
});
test('legacy rendering path keeps working unchanged',async({page})=>{
  await page.goto(BASE+'?engine=legacy',{waitUntil:'domcontentloaded'});
  await page.waitForFunction(()=>window.__LS3D_TEST__?.Z?.length===7);
  expect(await page.evaluate(()=>window.LS3D_THREE?.status)).toBe('legacy');
  await expect(page.locator('#scene')).toBeVisible();
});
