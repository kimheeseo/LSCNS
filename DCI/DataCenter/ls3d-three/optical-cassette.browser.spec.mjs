import {test,expect} from '@playwright/test';
import {readFile,writeFile} from 'node:fs/promises';

const ENTRY='http://127.0.0.1:8765/DCI/DataCenter/LS_Datacenter_3D.html?engine=three';
const THREE_BUILD=new URL('../../../node_modules/three/build/',import.meta.url);

async function start(page){
 const errors=[];page.on('pageerror',e=>errors.push(e.message));
 await page.addInitScript(()=>localStorage.setItem('ls3d-first-guide-v481','done'));
 await page.route('**/npm/three@0.186.1/build/three.*.js',async route=>{
  const file=route.request().url().includes('three.core.js')?'three.core.js':'three.module.js';
  await route.fulfill({status:200,contentType:'text/javascript',
   headers:{'access-control-allow-origin':'*'},body:await readFile(new URL(file,THREE_BUILD))});
 });
 await page.goto(ENTRY,{waitUntil:'domcontentloaded'});
 await page.waitForFunction(()=>['ready','failed'].includes(window.LS3D_THREE?.status),null,{timeout:45000});
 const engine=await page.evaluate(()=>({status:window.LS3D_THREE.status,error:window.LS3D_THREE.error}));
 expect(engine.status,engine.error||'Three.js did not initialize').toBe('ready');
 await page.waitForFunction(()=>window.LS3D_THREE.renders>3);
 const choice=page.locator('#displayModeChoice');
 if(await choice.isVisible())await page.locator('[data-display-choice="'+(page.viewportSize().width<640?'mobile':'pc')+'"]').click();
 await page.evaluate(()=>document.getElementById('guide')?.classList.remove('open'));
 await page.evaluate(()=>{
  const T=window.__LS3D_TEST__;T.setZone('hall');T.pickAsset(T.assetMap.get('tor-a'));
 });
 if(page.viewportSize().width<640){
  await page.getByRole('button',{name:/설비 검사·상세 펼치기/}).click();
  await page.locator('#inspectorRackOpen').click();
 }else await page.locator('#detailBtn').click();
 await page.waitForFunction(()=>window.__LS3D_TEST__.rackIsolateActive&&!!window.LS3D_THREE.adapter.rackInteraction);
 await page.locator('[data-rack-action="optical-mtp"]').click();
 await page.waitForFunction(()=>{
  const s=window.LS3D_THREE.adapter.rackInteraction.selectedOptics;
  return s?.open>.92&&s?.detail&&s?.isolated;
 },null,{timeout:20000});
 return errors;
}

async function inspect(page){return page.evaluate(()=>{
 const a=window.LS3D_THREE.adapter,objects={modules:0,lc:0,mpo:0,fibers:0,covers:0};
 const rack=a.scene.getObjectByName('equipment:tor-a');
 rack.traverse(o=>{
  if(/^rack-optical:module-\d+$/.test(o.name))objects.modules++;
  if(o.name.startsWith('rack-optical:adapter-port-'))objects.lc++;
  if(o.name.startsWith('rack-optical:mpo-keyed-receptacle-'))objects.mpo++;
  if(o.name.startsWith('rack-optical:fanout-fiber-'))objects.fibers++;
  if(o.name.startsWith('rack-optical:module-cover-'))objects.covers++;
 });
 const high=rack.getObjectByName('rack-high');
 return {status:window.LS3D_THREE.status,renders:window.LS3D_THREE.renders,
  calls:a.diagnostic.calls,selected:a.selectedPart?.kind,parts:objects,
  visibleHigh:high.children.filter(o=>o.visible).map(o=>o.name),
  coverOffset:rack.getObjectByName('rack-optical:module-cover-0').position.z,
  camera:{pitch:window.__LS3D_TEST__.cameraDesired.pitch,yaw:window.__LS3D_TEST__.cameraDesired.yaw}};
});}

test('desktop WebGL2: cassette front, rear MPO, LC path and rack restoration',async({page},info)=>{
 const errors=await start(page),front=await inspect(page);
 expect(front.status).toBe('ready');expect(front.renders).toBeGreaterThan(3);
 expect(front.calls).toBeGreaterThan(0);
 expect(front.parts).toEqual({modules:4,lc:24,mpo:4,fibers:48,covers:4});
 expect(front.visibleHigh).toEqual(['fiber-cassette']);
 expect(front.camera.pitch).toBeGreaterThan(.6);
 await page.screenshot({path:info.outputPath('cassette-front-open.png')});
 await page.evaluate(()=>window.LS3D_THREE.adapter.selectRackPart('tor-a','tor-a:mpo-input-0'));
 await page.waitForFunction(()=>window.__LS3D_TEST__.cameraDesired.yaw===0);
 await page.screenshot({path:info.outputPath('cassette-rear-mpo.png')});
 await page.evaluate(()=>window.LS3D_THREE.adapter.selectRackPart('tor-a','tor-a:fiber-port:0'));
 await page.screenshot({path:info.outputPath('cassette-lc-path.png')});
 await page.evaluate(()=>window.LS3D_THREE.adapter.clearInspectionSelection());
 await page.waitForFunction(()=>!window.LS3D_THREE.adapter.rackInteraction.selectedOptics?.isolated);
 const restored=await page.evaluate(()=>{
  const rack=window.LS3D_THREE.adapter.scene.getObjectByName('equipment:tor-a');
  return rack.getObjectByName('rack-high').children.filter(o=>o.visible).length;
 });
 expect(restored).toBeGreaterThan(1);
 expect(errors).toEqual([]);
 await writeFile(info.outputPath('desktop-metrics.json'),JSON.stringify({front,restored,errors},null,2));
});

test.describe('mobile',()=>{
test.use({viewport:{width:390,height:844}});
test('mobile WebGL2: cassette opens and renders without clipping errors',async({page},info)=>{
 const errors=await start(page),mobile=await inspect(page);
 expect(mobile.parts.fibers).toBe(48);expect(mobile.calls).toBeGreaterThan(0);
 expect(mobile.visibleHigh).toEqual(['fiber-cassette']);expect(errors).toEqual([]);
 await page.locator('#viewport').screenshot({path:info.outputPath('cassette-mobile-open.png')});
 await writeFile(info.outputPath('mobile-metrics.json'),JSON.stringify({mobile,errors},null,2));
});
});
