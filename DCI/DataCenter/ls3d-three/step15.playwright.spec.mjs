// STEP 15 browser acceptance suite. This file is NOT evidence that Chromium ran.
// LS3D_BASE_URL=http://localhost:8000/DCI/DataCenter/LS_Datacenter_3D.html
// npx playwright test DCI/DataCenter/ls3d-three/step15.playwright.spec.mjs
import {test,expect} from '@playwright/test';
const BASE=process.env.LS3D_BASE_URL||
  'https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html';
async function start(page){
  const errors=[];
  page.on('pageerror',e=>errors.push(e.message));
  page.on('console',m=>{if(m.type()==='error')errors.push(m.text());});
  await page.goto(BASE+(BASE.includes('?')?'&':'?')+'engine=three',{waitUntil:'domcontentloaded'});
  await page.waitForFunction(()=>['ready','failed'].includes(window.LS3D_THREE?.status),null,{timeout:30000});
  const mode=await page.evaluate(()=>window.LS3D_THREE?.status);
  expect(mode).toBe('ready');
  await page.evaluate(()=>window.__LS3D_TEST__.setZone('network'));
  await page.waitForFunction(()=>window.LS3D_THREE?.adapter?.mmr?.ready===true);
  await expect(page.locator('#ls3dMMRControls')).toBeVisible();
  return errors;
}
const sample=page=>page.evaluate(()=>{
  const a=window.LS3D_THREE.adapter,found=[];
  a.scene.traverse(o=>{if(o.name?.startsWith('mmr:'))found.push(o.name);});
  return {ready:a.mmr.ready,stats:a.mmr.stats,found};
});
test('STEP15 Carrier/ODF/DWDM/Core are distinct three-dimensional equipment compartments',async({page},info)=>{
  const errors=await start(page);
  const d=await sample(page);
  expect(d.ready).toBe(true);expect(d.stats.doors).toBeGreaterThanOrEqual(4);
  for(const name of ['mmr:carrier-demarc-plate','mmr:dwdm-muxponder-0',
    'mmr:odf-patch-shelf-0','mmr:core-interface-card-0'])
    expect(d.found).toContain(name);
  await page.screenshot({path:info.outputPath('step15-mmr-four-subrooms-desktop.png')});
  expect(errors).toEqual([]);
});
test('STEP15 hinged MMR door moves its Three.Group rather than swapping images',async({page},info)=>{
  const errors=await start(page);
  const button=page.locator('[data-mmr-door="carrier"]');
  await expect(button).toBeVisible();
  await button.click();
  await page.waitForFunction(()=>{
    let angle=0;
    window.LS3D_THREE.adapter.scene.traverse(o=>{
      if(o.name==='mmr:door-hinge:carrier')angle=o.rotation.y;
    });
    return angle<-.25;
  });
  await page.screenshot({path:info.outputPath('step15-carrier-entry-door-open.png')});
  await button.click();
  await page.waitForFunction(()=>window.LS3D_THREE.adapter.mmrController.summary().doors[0].opening<.08);
  expect(errors).toEqual([]);
});
test('STEP15 selects and dims optical paths while keeping connection metadata',async({page},info)=>{
  const errors=await start(page);
  const button=page.locator('[data-mmr-link="odf-dwdm"]');
  await button.click();
  const result=await page.evaluate(()=>{
    const root=window.LS3D_THREE.adapter.scene,mesh=[];
    root.traverse(o=>{if(o.name?.startsWith('mmr:fiber:'))mesh.push({
      name:o.name,opacity:o.material?.opacity,segment:o.userData.mmrSegment
    });});
    return {selected:window.LS3D_THREE.adapter.mmr.selectedLink,
      connections:window.LS3D_THREE.adapter.mmr.links,mesh};
  });
  expect(result.selected).toBe('odf-dwdm');
  expect(result.connections.every(l=>l.visible)).toBe(true);
  expect(result.mesh.find(m=>m.name==='mmr:fiber:odf-dwdm').opacity).toBeGreaterThan(.9);
  expect(result.mesh.find(m=>m.name==='mmr:fiber:carrier-odf').opacity).toBeLessThan(.2);
  await expect(page.locator('#ls3dMMRControls [data-summary]')).toContainText('ODF Cross-connect');
  await page.screenshot({path:info.outputPath('step15-odf-dwdm-connection.png')});
  expect(errors).toEqual([]);
});
test('STEP15 MMR hidden in server-room; mobile controls remain usable',async({page},info)=>{
  await page.setViewportSize({width:390,height:844});
  const errors=await start(page);
  await expect(page.locator('[data-mmr-link="carrier-odf"]')).toBeVisible();
  await page.locator('[data-mmr-link="carrier-odf"]').click();
  await page.screenshot({path:info.outputPath('step15-mmr-mobile.png')});
  await page.evaluate(()=>window.__LS3D_TEST__.applyFacilityMode('server-room'));
  await page.waitForFunction(()=>window.LS3D_THREE?.adapter?.mmr?.ready===false);
  await expect(page.locator('#ls3dMMRControls')).toBeHidden();
  expect(errors).toEqual([]);
});
