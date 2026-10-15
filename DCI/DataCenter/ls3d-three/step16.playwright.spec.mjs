// STEP 16 browser acceptance; requires Playwright Chromium.
// Source-based tests run separately; this file is NOT evidence of a browser PASS.
// npx playwright test DCI/DataCenter/ls3d-three/step16.playwright.spec.mjs
import {test,expect} from '@playwright/test';
const BASE=process.env.LS3D_BASE_URL||
  'https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html';
async function boot(page){
  const errors=[];page.on('pageerror',e=>errors.push(e.message));
  page.on('console',m=>{if(m.type()==='error')errors.push(m.text());});
  await page.goto(BASE+(BASE.includes('?')?'&':'?')+'engine=three',{waitUntil:'domcontentloaded'});
  await page.waitForFunction(()=>['ready','failed'].includes(window.LS3D_THREE?.status),null,{timeout:30000});
  expect(await page.evaluate(()=>window.LS3D_THREE.status)).toBe('ready');
  await page.evaluate(()=>window.__LS3D_TEST__.setZone('hall'));
  await page.waitForFunction(()=>window.LS3D_THREE?.adapter?.colocation?.ready===true,null,{timeout:15000});
  await expect(page.locator('#ls3dColoControls')).toBeVisible();
  return errors;
}
test('STEP16 customer A/B/C fenced bays avoid actual rack footprints',async({page},info)=>{
  const errors=await boot(page);
  const data=await page.evaluate(()=>{
    const sim=window.LS3D_THREE.adapter.colocationController;
    const nodes=[];sim.root.traverse(o=>nodes.push(o.name));
    return {s:sim.summary(),nodes};
  });
  expect(data.s.bays.map(x=>x.id)).toEqual(['A','B','C']);
  expect(data.s.layoutIssue).toBe(null);
  expect(data.s.stats.meshWires).toBeGreaterThan(150);
  expect(data.s.stats.doors).toBe(3);
  expect(data.s.stats.powerPaths).toBe(6);
  expect(data.s.stats.fiberPaths).toBe(3);
  expect(data.nodes).toContain('colo:fiber-adapter-A-0');
  expect(data.nodes).toContain('colo:rear-pdu-feed-B-B');
  await page.screenshot({path:info.outputPath('step16-tenant-mesh-cages-desktop.png')});
  expect(errors).toEqual([]);
});
test('STEP16 customer B power B filter fades other 3D fibers and power',async({page},info)=>{
  const errors=await boot(page);
  await page.locator('[data-colo-circuit="B:power-B"]').click();
  const data=await page.evaluate(()=>{
    const adapter=window.LS3D_THREE.adapter;
    const mesh=[];adapter.colocationController.root.traverse(o=>{
      if(o.name.startsWith('colo:wire:'))mesh.push({name:o.name,opacity:o.material.opacity});
    });
    return {s:adapter.colocationController.summary(),mesh};
  });
  expect(data.s.selected).toBe('B');
  expect(data.s.selectedCircuit).toBe('power-B');
  expect(data.s.bays.flatMap(b=>b.circuits).filter(c=>c.highlighted)).toHaveLength(1);
  expect(data.s.bays.flatMap(b=>b.circuits).every(c=>c.visible)).toBe(true);
  expect(data.mesh.find(m=>m.name==='colo:wire:power-B-B').opacity).toBeGreaterThan(.9);
  expect(data.mesh.find(m=>m.name==='colo:wire:fiber-A').opacity).toBeLessThan(.2);
  await expect(page.locator('#ls3dColoControls [data-summary]')).toContainText('Rack PDU B');
  await page.screenshot({path:info.outputPath('step16-power-b-focus.png')});
  await page.locator('[data-colo-tenant="all"]').click();
  expect(await page.evaluate(()=>window.LS3D_THREE.adapter.colocation.selectedCircuit)).toBe('all');
  expect(errors).toEqual([]);
});
test('STEP16 three real hinges animate individually and retain opening on sectional view',async({page},info)=>{
  const errors=await boot(page);
  await page.locator('[data-colo-gate="C"]').click();
  await page.waitForFunction(()=>{
    let y=0;window.LS3D_THREE.adapter.colocationController.root.traverse(o=>{
      if(o.name==='colo:gate-hinge:C')y=o.rotation.y;
    });return y<-.3;
  });
  const doors=await page.evaluate(()=>window.LS3D_THREE.adapter.colocation.bays);
  expect(doors.find(d=>d.id==='C').gateOpening).toBeGreaterThan(.2);
  expect(doors.find(d=>d.id==='A').gateOpening).toBeLessThan(.05);
  await page.screenshot({path:info.outputPath('step16-customer-c-gate.png')});
  expect(errors).toEqual([]);
});
test('STEP16 mobile controls, server-room isolation, legacy fallback',async({page},info)=>{
  await page.setViewportSize({width:390,height:844});
  const errors=await boot(page);
  await page.locator('[data-colo-circuit="A:fiber"]').click();
  await expect(page.locator('#ls3dColoControls [data-summary]')).toContainText('광 패치/분배');
  await page.screenshot({path:info.outputPath('step16-colocation-mobile.png')});
  await page.evaluate(()=>window.__LS3D_TEST__.applyFacilityMode('server-room'));
  await page.waitForFunction(()=>window.LS3D_THREE?.adapter?.colocation?.ready===false);
  await expect(page.locator('#ls3dColoControls')).toBeHidden();
  expect(errors).toEqual([]);
  await page.goto(BASE+(BASE.includes('?')?'&':'?')+'engine=legacy',{waitUntil:'domcontentloaded'});
  await page.waitForFunction(()=>window.__LS3D_TEST__?.Z?.length===7);
  expect(await page.evaluate(()=>window.LS3D_THREE?.status)).toBe('legacy');
});
