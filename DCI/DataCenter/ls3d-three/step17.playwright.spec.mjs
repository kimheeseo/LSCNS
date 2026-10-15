// STEP 17 browser acceptance (real Chromium run required; merely storing does NOT prove PASS).
// LS3D_BASE_URL=http://localhost:8000/DCI/DataCenter/LS_Datacenter_3D.html
// npx playwright test DCI/DataCenter/ls3d-three/step17.playwright.spec.mjs
import {test,expect} from '@playwright/test';
const BASE=process.env.LS3D_BASE_URL||
  'https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html';
async function setup(page){
  const errors=[];
  page.on('pageerror',e=>errors.push(e.message));
  page.on('console',m=>{if(m.type()==='error')errors.push(m.text());});
  await page.goto(BASE+(BASE.includes('?')?'&':'?')+'engine=three',{waitUntil:'domcontentloaded'});
  await page.waitForFunction(()=>['ready','failed'].includes(window.LS3D_THREE?.status),null,{timeout:30000});
  expect(await page.evaluate(()=>window.LS3D_THREE?.status)).toBe('ready');
  await page.evaluate(()=>window.__LS3D_TEST__.setZone('campus'));
  await page.waitForFunction(()=>window.LS3D_THREE?.adapter?.mep?.total>0);
  return errors;
}
test('STEP17 all three systems create pickable TubeGeometry and physical cable trays',async({page},info)=>{
  const errors=await setup(page);
  const d=await page.evaluate(()=>{
    const a=window.LS3D_THREE.adapter,o=[];
    a.mepController.root.traverse(x=>{
      if(x.name.startsWith('mep:tube:'))o.push({name:x.name,route:x.userData.mepRouteId,
        tube:x.geometry?.type,opacity:x.material?.opacity});
    });
    return {summary:a.mep,objects:o};
  });
  expect(new Set(d.summary.paths.map(p=>p.system))).toEqual(new Set(['power','cooling','fiber']));
  expect(d.objects).toHaveLength(d.summary.total);
  expect(d.objects.every(x=>!!x.route)).toBe(true);
  expect(d.summary.stats.traySegments).toBeGreaterThan(0);
  expect(d.summary.stats.pipeClamps).toBeGreaterThan(0);
  await page.screenshot({path:info.outputPath('step17-campus-all-systems.png')});
  expect(errors).toEqual([]);
});
test('STEP17 selected MMR to Colo optical path highlights and displays source and destination',async({page},info)=>{
  const errors=await setup(page);
  const target='fiber:mmr-hall:colo-rack-a';
  await page.locator('#ls3dMEPControls summary').click();
  await page.locator('#ls3dMEPRouteSelect').selectOption(target);
  const result=await page.evaluate(()=>{
    const a=window.LS3D_THREE.adapter;
    const tubes=[];a.mepController.root.traverse(o=>{
      if(o.name.startsWith('mep:tube:'))tubes.push({
        id:o.userData.mepRouteId,visible:o.visible,opacity:o.material?.opacity
      });
    });
    return {summary:a.mep,tubes};
  });
  expect(result.summary.selected).toBe(target);
  expect(result.summary.paths.filter(x=>x.highlighted)).toHaveLength(1);
  expect(result.tubes.find(x=>x.id===target).opacity).toBeGreaterThan(.9);
  expect(result.tubes.find(x=>x.id!==target).opacity).toBeLessThan(.2);
  await expect(page.locator('#ls3dMEPRouteInfo')).toContainText('Co-location Rack A');
  await page.screenshot({path:info.outputPath('step17-mmr-to-colo-fiber.png')});
  expect(errors).toEqual([]);
});
test('STEP17 Hall-local routes, system filter, server-room and inspect isolation',async({page},info)=>{
  const errors=await setup(page);
  await page.evaluate(()=>window.__LS3D_TEST__.setZone('hall'));
  const hall=await page.evaluate(()=>window.LS3D_THREE.adapter.mep);
  expect(hall.paths.every(x=>x.fromZone==='hall'||x.toZone==='hall')).toBe(true);
  await page.locator('#ls3dMEPControls summary').click();
  await page.locator('#ls3dMEPControls [data-mep-system="cooling"]').click();
  const focused=await page.evaluate(()=>window.LS3D_THREE.adapter.mep);
  expect(focused.paths.every(x=>x.highlighted===(x.system==='cooling'))).toBe(true);
  await page.evaluate(()=>window.__LS3D_TEST__.applyFacilityMode('server-room'));
  await page.waitForFunction(()=>window.LS3D_THREE.adapter.mep.visible===false);
  await expect(page.locator('#ls3dMEPControls')).toBeHidden();
  expect(errors).toEqual([]);
});
test('STEP17 390px mobile UI and no JS console errors',async({page},info)=>{
  await page.setViewportSize({width:390,height:844});
  const errors=await setup(page);
  await page.locator('#ls3dMEPControls summary').click();
  await expect(page.locator('#ls3dMEPRouteSelect')).toBeVisible();
  await page.locator('#ls3dMEPControls [data-mep-system="fiber"]').click();
  await page.screenshot({path:info.outputPath('step17-mobile-fiber-routes.png')});
  expect(errors).toEqual([]);
});
