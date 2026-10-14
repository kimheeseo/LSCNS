// STEP 04 real-browser regression for the architecture layer.
// Run locally after: npm install --no-save @playwright/test && npx playwright install chromium
// Then: npx playwright test DCI/DataCenter/ls3d-three/step04.playwright.spec.mjs
// WARNING: These checks have not been independently run on GPU hardware yet.
import { test, expect } from '@playwright/test';

const BASE = process.env.LS3D_BASE_URL ||
  'https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html';

test.use({ viewport: {width: 1366, height: 768} });
async function start(page) {
  await page.goto(BASE + '?engine=three', { waitUntil:'domcontentloaded' });
  await page.waitForFunction(() => ['ready','failed'].includes(window.LS3D_THREE?.status),null,{timeout:30000});
  const status = await page.evaluate(() => ({status:window.LS3D_THREE.status,error:window.LS3D_THREE.error}));
  expect(status.status,status.error||'Could not initialize Three.js').toBe('ready');
  await page.waitForFunction(() => window.LS3D_THREE?.renders > 3);
}

test('campus contains physically modeled envelopes, glazing, doors and trays',async ({page}, info)=>{
  const errors=[];page.on('pageerror',e=>errors.push(e.message));
  await start(page);
  const result=await page.evaluate(()=>{
    const S=window.LS3D_THREE;
    const kinds={};let boxMeshes=0,animatedDoors=0;
    S.adapter.scene.traverse(obj=>{
      if(obj.userData?.archKind) kinds[obj.userData.archKind]=(kinds[obj.userData.archKind]||0)+1;
      if(obj.isMesh && obj.name.startsWith('arch:')) boxMeshes++;
      if(obj.name.startsWith('arch:door-hinge'))animatedDoors++;
    });
    return {counts:S.adapter.architecture.counts,zoneCount:S.diagnostics.zones,
      kinds,boxMeshes,animatedDoors,frames:S.renders};
  });
  expect(result.zoneCount).toBe(7);
  expect(result.counts.wall).toBeGreaterThan(0);
  expect(result.counts.glass).toBeGreaterThan(0);
  expect(result.counts.floor).toBeGreaterThan(0);
  expect(result.counts.ceiling).toBeGreaterThan(0);
  expect(result.counts.tray).toBeGreaterThan(0);
  expect(result.counts.door).toBeGreaterThan(0);
  expect(result.animatedDoors).toBeGreaterThan(0);
  expect(result.boxMeshes).toBeGreaterThan(150);
  expect(errors).toEqual([]);
  await page.screenshot({path:info.outputPath('step04-campus.png')});
});
test('data hall roof toggles, section cuts, double-entry doors animate',async ({page},info)=>{
  await start(page);
  await page.evaluate(()=>window.__LS3D_TEST__.setZone('hall'));
  await page.waitForTimeout(700);
  const before=await page.evaluate(()=>{
    const S=window.LS3D_THREE, hinges=[];
    S.adapter.scene.traverse(o=>{if(o.name.startsWith('arch:door-hinge'))hinges.push(o.rotation.y);});
    return {hinges, roofs:S.adapter.architecture.counts.roof, walls:S.adapter.architecture.counts.wall};
  });
  expect(before.hinges.length).toBe(2);
  expect(Math.abs(before.hinges[0])).toBeGreaterThan(.3);
  expect(Math.sign(before.hinges[0])).toBe(-Math.sign(before.hinges[1]));
  expect(before.roofs).toBe(0);
  await page.locator('#btnRoof').click();
  await page.waitForFunction(()=>window.LS3D_THREE?.adapter?.architecture?.counts?.roof>0);
  await page.screenshot({path:info.outputPath('step04-hall-roof.png')});
  await page.locator('#btnSection').click();
  await page.waitForFunction(()=>window.LS3D_THREE?.adapter?.architecture?.counts?.wall===0);
  const section=await page.evaluate(()=>window.LS3D_THREE.adapter.architecture);
  expect(section.counts.roof).toBe(0);
  expect(section.counts.tray).toBeGreaterThan(0);
  await page.screenshot({path:info.outputPath('step04-hall-section.png')});
});
test('server-room renders only the local IT room and never entire campus',async ({page})=>{
  await start(page);
  await page.evaluate(()=>{
    const T=window.__LS3D_TEST__;T.applyFacilityMode('server-room');T.setZone('hall');
  });
  await page.waitForFunction(()=>window.LS3D_THREE?.diagnostics?.zones===1);
  const result=await page.evaluate(()=>{
    const S=window.LS3D_THREE,world=S.adapter.scene.children[0],
      zoneRoot=world.children[0];
    return {zones:S.diagnostics.zones,buildingGroups:zoneRoot.children.length,
      names:zoneRoot.children.map(g=>g.name),trays:S.adapter.architecture.counts.tray};
  });
  expect(result.buildingGroups).toBe(1);
  expect(result.names).toEqual(['arch:zone:hall']);
  expect(result.trays).toBeGreaterThan(0);
});
test('mobile viewport gets a real architectural Three scene',async ({page}, info)=>{
  await page.setViewportSize({width:390,height:844});
  await start(page);
  await expect(page.locator('#sceneThree')).toBeVisible();
  const count=await page.evaluate(()=>window.LS3D_THREE?.adapter?.architecture?.counts);
  expect(count.glass).toBeGreaterThan(0);
  expect(count.tray).toBeGreaterThan(0);
  await page.screenshot({path:info.outputPath('step04-mobile.png')});
});
