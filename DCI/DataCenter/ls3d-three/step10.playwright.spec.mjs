// STEP 10 real Chromium rack hierarchy / interaction regression, not yet run.
import {test,expect} from '@playwright/test';
const BASE=process.env.LS3D_BASE_URL||'https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html';
async function open(page) {
 await page.goto(BASE+'?engine=three',{waitUntil:'domcontentloaded'});
 await page.waitForFunction(()=>['ready','failed'].includes(window.LS3D_THREE?.status),null,{timeout:30000});
 expect(await page.evaluate(()=>window.LS3D_THREE.status)).toBe('ready');
}
test('42/48U racks contain independently selectable server, switch, ODF and PDU objects',async({page},info)=>{
 await open(page);
 const r=await page.evaluate(()=>{
  const adapter=window.LS3D_THREE.adapter;
  const sample=[];adapter.scene.traverse(o=>{
   if(o.userData?.rackPart)sample.push(o.userData.rackPart);
  });
  const unique=[...new Set(sample.map(p=>p.kind))];
  const selected=sample.find(p=>p.kind==='tor');
  const chosen=selected&&adapter.selectRackPart(selected.assetId,selected.id);
  const overlay=document.getElementById('ls3dRackPartInfo');
  return {count:sample.length,types:unique,stats:adapter.rackComponents,
    selected:chosen,overlayVisible:overlay&&!overlay.hidden,
    label:overlay?.textContent||'',units:new Set(sample.map(p=>p.uHeight))};
 });
 for(const name of ['frame','server-1','tor','odf','patch-panel','pdu-a','pdu-b',
  'front-door','rear-door','fiber-cassette','cable-manager']){
  expect(r.types).toContain(name);
 }
 expect(r.stats.counts.units42).toBeGreaterThan(0);
 expect(r.stats.counts.units48).toBeGreaterThan(0);
 expect(r.stats.visibleParts).toBeGreaterThan(100);
 expect(r.overlayVisible).toBe(true);
 expect(r.selected.kind).toBe('tor');
 expect(r.label).toContain('ToR');
 await page.screenshot({path:info.outputPath('step10-rack-component-selection.png')});
});
test('rack door, slide and cassette remains operable in hall and company room',async({page},info)=>{
 await open(page);
 await page.evaluate(()=>window.__LS3D_TEST__.setZone('hall'));
 await page.waitForFunction(()=>window.LS3D_THREE?.diagnostics?.zones===2);
 const result=await page.evaluate(()=>{
  const adapter=window.LS3D_THREE.adapter;
  let front=0,slide=0,cassette=0,port=0;
  adapter.scene.traverse(o=>{
    if(o.name==='front-door-hinge')front++;
    if(o.name==='server-tray')slide++;
    if(o.name==='fiber-cassette')cassette++;
    if(o.name?.startsWith('rack-shape:server-port'))port++;
  });
  return {front,slide,cassette,port};
 });
 expect(result.front).toBeGreaterThan(0);
 expect(result.slide).toBeGreaterThan(0);
 expect(result.cassette).toBeGreaterThan(0);
 expect(result.port).toBeGreaterThan(0);
 await page.screenshot({path:info.outputPath('step10-hall-rack-parts.png')});
 await page.evaluate(()=>window.__LS3D_TEST__.applyFacilityMode('server-room'));
 await page.waitForFunction(()=>window.LS3D_THREE?.diagnostics?.zones===1);
 expect(await page.evaluate(()=>window.LS3D_THREE.adapter.rackComponents.rackCount)).toBeGreaterThan(0);
});
test('mobile and legacy default path remain separate',async({page},info)=>{
 await page.setViewportSize({width:390,height:844});
 await open(page);
 await expect(page.locator('#sceneThree')).toBeVisible();
 await page.screenshot({path:info.outputPath('step10-mobile.png')});
 await page.goto(BASE+'?engine=legacy',{waitUntil:'domcontentloaded'});
 await page.waitForFunction(()=>window.__LS3D_TEST__?.Z?.length===7);
 expect(await page.evaluate(()=>window.LS3D_THREE?.status)).toBe('legacy');
});
