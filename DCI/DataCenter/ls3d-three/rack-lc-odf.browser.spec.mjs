import {test,expect} from '@playwright/test';
import {readFile,writeFile} from 'node:fs/promises';
const ENTRY='http://127.0.0.1:8765/DCI/DataCenter/LS_Datacenter_3D.html?engine=three';
const THREE_BUILD=new URL('../../../node_modules/three/build/',import.meta.url);
async function start(page,id,zone){
 const errors=[];page.on('pageerror',e=>errors.push(e.message));
 await page.addInitScript(()=>localStorage.setItem('ls3d-first-guide-v481','done'));
 await page.route('**/npm/three@0.186.1/build/three.*.js',async route=>{
  const name=route.request().url().includes('three.core.js')?'three.core.js':'three.module.js';
  await route.fulfill({status:200,contentType:'text/javascript',
   headers:{'access-control-allow-origin':'*'},body:await readFile(new URL(name,THREE_BUILD))});
 });
 await page.goto(ENTRY,{waitUntil:'domcontentloaded'});
 await page.waitForFunction(()=>['ready','failed'].includes(window.LS3D_THREE?.status),null,{timeout:45000});
 const engine=await page.evaluate(()=>({status:window.LS3D_THREE.status,error:window.LS3D_THREE.error}));
 expect(engine.status,engine.error||'Three.js initialization failed').toBe('ready');
 const choice=page.locator('#displayModeChoice');
 if(await choice.isVisible())await page.locator('[data-display-choice="'+(page.viewportSize().width<640?'mobile':'pc')+'"]').click();
 await page.evaluate(()=>document.getElementById('guide')?.classList.remove('open'));
 await page.evaluate(({id,zone})=>{
  const T=window.__LS3D_TEST__;T.setZone(zone);T.pickAsset(T.assetMap.get(id));
 },{id,zone});
 if(page.viewportSize().width<640){
  await page.getByRole('button',{name:/설비 검사·상세 펼치기/}).click();
  await page.locator('#inspectorRackOpen').click();
 }else await page.locator('#detailBtn').click();
 await page.waitForFunction(()=>window.__LS3D_TEST__.rackIsolateActive);
 return errors;
}
async function metrics(page,id,prefix){return page.evaluate(({id,prefix})=>{
 const a=window.LS3D_THREE.adapter,rack=a.scene.getObjectByName('equipment:'+id);
 const found=[];rack.traverse(o=>{if(o.name?.startsWith(prefix))found.push(o.name)});
 return {status:window.LS3D_THREE.status,renders:window.LS3D_THREE.renders,
  calls:a.diagnostic.calls,objects:found,visibleHigh:rack.getObjectByName('rack-high').children.filter(o=>o.visible).map(o=>o.name),
  selected:a.selectedPart?.kind,lc:a.rackInteraction.selectedLcPanel,odf:a.rackInteraction.selectedOdfBox};
},{id,prefix});}
test('PC LC adapter front, rear and selected pigtail',async({page},info)=>{
 const errors=await start(page,'tor-a','hall');
 await page.locator('#viewport').screenshot({path:info.outputPath('network-rack-bottom-space.png')});
 await page.locator('[data-rack-action="optical-lc"]').click();
 await page.waitForFunction(()=>window.LS3D_THREE.adapter.rackInteraction.selectedLcPanel?.isolated);
 const front=await metrics(page,'tor-a','rack-lc:');
 expect(front.lc.portCount).toBe(12);expect(front.lc.fiberCount).toBe(24);
 expect(front.objects.filter(x=>x.startsWith('rack-lc:duplex-adapter-'))).toHaveLength(12);
 expect(front.objects.filter(x=>x.startsWith('rack-lc:rear-pigtail-'))).toHaveLength(24);
 expect(front.calls).toBeGreaterThan(0);
 await page.locator('#viewport').screenshot({path:info.outputPath('lc-front.png')});
 await page.locator('[data-rack-quick="lc-rear"]').click();
 await page.waitForFunction(()=>window.__LS3D_TEST__.cameraDesired.yaw===0);
 await page.locator('#viewport').screenshot({path:info.outputPath('lc-rear.png')});
 await page.locator('[data-rack-quick="lc-next-port"]').click();
 await page.waitForFunction(()=>window.LS3D_THREE.adapter.rackInteraction.selectedLcPanel?.selectedPort===0);
 const port=await metrics(page,'tor-a','rack-lc:');
 expect(port.selected).toBe('lc-adapter-port');expect(errors).toEqual([]);
 await page.locator('#viewport').screenshot({path:info.outputPath('lc-port-path.png')});
 await page.evaluate(()=>window.LS3D_THREE.adapter.clearInspectionSelection());
 await page.waitForFunction(()=>!window.LS3D_THREE.adapter.rackInteraction.selectedLcPanel?.isolated);
 await writeFile(info.outputPath('lc-metrics.json'),JSON.stringify({front,port,errors},null,2));
});
test('PC ODF enclosure closes, hinges open and highlights LC pigtail',async({page},info)=>{
 const errors=await start(page,'odf','network');
 await page.locator('[data-rack-action="odf-select"]').click();
 await page.waitForFunction(()=>window.LS3D_THREE.adapter.rackInteraction.selectedOdfBox?.isolated);
 expect(await page.evaluate(()=>window.__LS3D_TEST__.cameraDesired.yaw)).toBe(0);
 const closed=await metrics(page,'odf','odf-small:');
 expect(closed.odf.portCount).toBe(6);expect(closed.odf.open).toBe(0);
 expect(closed.objects.filter(x=>x.startsWith('odf-small:pigtail-'))).toHaveLength(6);
 await page.locator('#viewport').screenshot({path:info.outputPath('odf-closed.png')});
 await page.locator('[data-rack-quick="odf-cover"]').click();
 await page.waitForFunction(()=>window.LS3D_THREE.adapter.rackInteraction.selectedOdfBox?.open>.95);
 const opened=await metrics(page,'odf','odf-small:');
 expect(opened.odf.target).toBe(1);expect(opened.visibleHigh).toEqual(['odf-small:enclosure']);
 await page.locator('#viewport').screenshot({path:info.outputPath('odf-open.png')});
 await page.locator('[data-rack-quick="odf-next-port"]').click();
 await page.waitForFunction(()=>window.LS3D_THREE.adapter.rackInteraction.selectedOdfBox?.selectedPort===0);
 const port=await metrics(page,'odf','odf-small:');
 expect(port.selected).toBe('odf-small-port');
 await page.locator('#viewport').screenshot({path:info.outputPath('odf-port-path.png')});
 await page.locator('[data-rack-quick="odf-cover"]').click();
 await page.waitForFunction(()=>window.LS3D_THREE.adapter.rackInteraction.selectedOdfBox?.open<.05);
 expect(errors).toEqual([]);
 await writeFile(info.outputPath('odf-metrics.json'),JSON.stringify({closed,opened,port,errors},null,2));
});
test.describe('mobile',()=>{
 test.use({viewport:{width:390,height:844}});
 test('ODF opens and all six LC ports render on a narrow screen',async({page},info)=>{
  const errors=await start(page,'odf','network');
  await page.locator('[data-rack-quick="odf-select"]').click();
  await page.locator('[data-rack-quick="odf-cover"]').click();
  await page.waitForFunction(()=>window.LS3D_THREE.adapter.rackInteraction.selectedOdfBox?.open>.95);
  const mobile=await metrics(page,'odf','odf-small:');
  expect(mobile.objects.filter(x=>x.startsWith('odf-small:lc-adapter-'))).toHaveLength(6);
  expect(mobile.calls).toBeGreaterThan(0);expect(errors).toEqual([]);
  await page.locator('#viewport').screenshot({path:info.outputPath('odf-mobile-open.png')});
  await writeFile(info.outputPath('mobile-metrics.json'),JSON.stringify({mobile,errors},null,2));
 });
});
