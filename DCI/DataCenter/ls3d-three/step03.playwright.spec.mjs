// STEP 03 browser regression — requires @playwright/test and deployed GitHub Pages.
// Run: npx playwright test DCI/DataCenter/ls3d-three/step03.playwright.spec.mjs
// Optional env: LS3D_BASE_URL=http://127.0.0.1:8765/DCI/DataCenter/LS_Datacenter_3D.html
import { test, expect } from '@playwright/test';
const URL3D = process.env.LS3D_BASE_URL ||
  'https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html';
const basicErrors = [];
test.use({ viewport: { width: 1366, height: 768 } });
test('legacy renderer remains default and 2D/BOM links exist', async ({page}) => {
  const errors=[];page.on('pageerror', e=>errors.push(e.message));
  await page.goto(URL3D + '?engine=legacy',{waitUntil:'domcontentloaded'});
  await page.waitForFunction(()=>window.__LS3D_TEST__?.Z?.length===7);
  expect(await page.evaluate(()=>window.LS3D_THREE?.status)).toBe('legacy');
  await expect(page.locator('#scene')).toBeVisible();
  await expect(page.locator('a[href="./LS_Datacenter_Campus.html"]').first()).toBeVisible();
  await expect(page.locator('a[href="./index.html"]').first()).toBeVisible();
  expect(errors).toEqual([]);
});
test('Three.js creates GPU objects and syncs all seven zones', async ({page}, info) => {
  const errors=[];page.on('pageerror', e=>errors.push(e.message));
  await page.goto(URL3D + '?engine=three',{waitUntil:'domcontentloaded'});
  await page.waitForFunction(()=>['ready','failed'].includes(window.LS3D_THREE?.status),null,{timeout:30000});
  const initial=await page.evaluate(()=>{
    const S=window.LS3D_THREE,T=window.__LS3D_TEST__;
    return {status:S.status,error:S.error,version:S.adapter?.THREE.REVISION,
      renders:S.renders,zones:S.diagnostics?.zones,assets:S.diagnostics?.assets,
      camera:!!S.adapter?.camera,scene:!!S.adapter?.scene,
      originalT:T?.Z?.length, hasRenderer:!!S.adapter?.renderer};
  });
  expect(initial.status,initial.error||'Three.js module did not initialize').toBe('ready');
  expect(initial.version).toBe('186');
  expect(initial.originalT).toBe(7);
  expect(initial.hasRenderer&&initial.camera&&initial.scene).toBe(true);
  await page.waitForFunction(()=>window.LS3D_THREE?.renders>2);
  expect(initial.assets).toBeGreaterThan(0);
  for(const id of ['campus','utility','cooling','gray','hall','network','ops']){
    await page.evaluate(id=>window.__LS3D_TEST__.setZone(id),id);
    await page.waitForTimeout(200);
    const check=await page.evaluate(()=>({zone:window.__LS3D_TEST__.currentZone,assets:window.LS3D_THREE.diagnostics?.assets}));
    expect(check.zone).toBe(id);
    expect(check.assets).toBeGreaterThan(0);
  }
  await page.screenshot({path:info.outputPath('step03-desktop.png'),fullPage:false});
  expect(errors).toEqual([]);
});
test('server-room restricts external scene content', async ({page})=>{
  await page.goto(URL3D+'?engine=three',{waitUntil:'domcontentloaded'});
  await page.waitForFunction(()=>window.LS3D_THREE?.status==='ready',{},{timeout:30000});
  const r=await page.evaluate(()=>{
    const T=window.__LS3D_TEST__;
    T.applyFacilityMode('server-room');T.setZone('hall');
    return {profile:T.facilityMode,zones:T.visibleZones().map(z=>z.id)};
  });
  expect(r.profile).toBe('server-room');
  expect(r.zones).toEqual(['hall']);
  await page.waitForTimeout(350);
  const counts=await page.evaluate(()=>{
    const S=window.LS3D_THREE;
    return {zones:S.diagnostics?.zones,routes:S.adapter?.scene.children[0].children[2].children.length};
  });
  expect(counts.zones).toBe(1);
  expect(counts.routes).toBe(0);
});
test('mobile canvas and Three scene remain usable', async ({page},info})=>{
  await page.setViewportSize({width:390,height:844});
  await page.goto(URL3D+'?engine=three',{waitUntil:'domcontentloaded'});
  await page.waitForFunction(()=>window.LS3D_THREE?.status==='ready',{},{timeout:30000});
  await expect(page.locator('#sceneThree')).toBeVisible();
  await page.waitForFunction(()=>window.LS3D_THREE?.renders>=3);
  await page.screenshot({path:info.outputPath('step03-mobile.png'),fullPage:false});
});
