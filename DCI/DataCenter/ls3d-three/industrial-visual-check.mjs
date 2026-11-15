import assert from 'node:assert/strict';
import {mkdir,readFile,writeFile} from 'node:fs/promises';
import {chromium} from 'playwright';
const out='/tmp/datacenter-industrial-qa';await mkdir(out,{recursive:true});
const browser=await chromium.launch({headless:true,args:['--no-sandbox','--enable-unsafe-swiftshader','--use-angle=swiftshader','--disable-dev-shm-usage']});
try{
 const page=await browser.newPage({viewport:{width:1600,height:1000}}),errors=[];
 page.on('pageerror',e=>errors.push(e.message));
 await page.addInitScript(()=>{localStorage.setItem('ls3d-first-guide-v481','done');localStorage.setItem('ls3d-display-mode-v1','pc');});
 const build=new URL('../../../node_modules/three/build/',import.meta.url);
 await page.route('**/npm/three@0.186.1/build/three.*.js',async route=>{
  const name=route.request().url().includes('three.core.js')?'three.core.js':'three.module.js';
  await route.fulfill({status:200,contentType:'text/javascript',headers:{'access-control-allow-origin':'*'},body:await readFile(new URL(name,build))});
 });
 await page.goto('http://127.0.0.1:8765/DCI/DataCenter/LS_Datacenter_3D.html?engine=three',{waitUntil:'load',timeout:90000});
 await page.waitForFunction(()=>window.LS3D_THREE?.status==='ready',null,{timeout:60000});
 if(await page.locator('#displayModeChoice').isVisible())await page.locator('[data-display-choice="pc"]').click();
 await page.evaluate(()=>document.getElementById('guide').classList.remove('open'));
 await page.waitForFunction(()=>window.LS3D_THREE.renders>5);
 await page.locator('#viewport').screenshot({path:out+'/01-campus.png'});
 await page.evaluate(()=>window.__LS3D_TEST__.setZone('hall'));
 await page.waitForFunction(()=>window.LS3D_THREE.adapter.diagnostic.assets>0);
 await page.locator('#viewport').screenshot({path:out+'/02-data-hall.png'});
 const results=[];
 async function assertMechanicalMotion(id){
  await page.locator('#ls3dRackPanelToggle').click();
  const door=page.locator('[data-rack-mechanical="front-door"]');
  const cassette=page.locator('[data-rack-mechanical="cassette-slide"]');
  assert.ok(await door.isVisible());assert.ok(await cassette.isVisible());
  assert.equal(await page.locator('#ls3dRackMotionBody').isHidden(),true);
  await page.waitForFunction(()=>{const p=document.getElementById('ls3dRackMotionControls').getBoundingClientRect();return [...document.querySelectorAll('[data-rack-mechanical]')].every(b=>b.getBoundingClientRect().bottom<=p.bottom&&b.getBoundingClientRect().top>=p.top);});
  await door.click();
  await page.waitForFunction(id=>{
   const root=window.LS3D_THREE.adapter.scene.getObjectByName('equipment:'+id);
   return Math.abs(root.getObjectByName('front-door-hinge').rotation.y)>1.5;
  },id);
  assert.equal(await door.textContent(),'랙 문 닫기');
  await cassette.click();
  await page.waitForFunction(id=>{
   const root=window.LS3D_THREE.adapter.scene.getObjectByName('equipment:'+id);
   return root.getObjectByName('fiber-cassette').position.z < -1;
  },id);
  assert.equal(await cassette.textContent(),'카세트 넣기');
  assert.equal(await page.evaluate(()=>window.LS3D_THREE.adapter.rackPanel.collapsed),true);
  if(id==='tor-a')await page.locator('#viewport').screenshot({path:out+'/mechanical-door-open-cassette-out.png'});
  await cassette.click();
  await page.waitForFunction(id=>Math.abs(window.LS3D_THREE.adapter.scene.getObjectByName('equipment:'+id).getObjectByName('fiber-cassette').position.z)<.02,id);
  assert.equal(await cassette.textContent(),'카세트 빼기');
  await door.click();
  await page.waitForFunction(id=>Math.abs(window.LS3D_THREE.adapter.scene.getObjectByName('equipment:'+id).getObjectByName('front-door-hinge').rotation.y)<.02,id);
  assert.equal(await door.textContent(),'랙 문 열기');
  // Closing the door retracts an extended cassette using the original motion controller.
  await cassette.click();
  await page.waitForFunction(()=>window.__LS3D_TEST__.mechanical(window.__LS3D_TEST__.selected).cassette>.95);
  await door.click();
  await page.waitForFunction(()=>{
   const m=window.__LS3D_TEST__.mechanical(window.__LS3D_TEST__.selected);return m.door<.01&&m.cassette<.01;
  });
  if(id==='tor-a')await page.locator('#viewport').screenshot({path:out+'/mechanical-door-closed-cassette-in.png'});
  await page.locator('#ls3dRackPanelToggle').click();
 }

 async function assertRackInventory(id){
  const inventory=await page.evaluate(id=>{
   const root=window.LS3D_THREE.adapter.scene.getObjectByName('equipment:'+id),ids=new Set();
   root.traverseVisible(o=>{
    if(!o.isMesh)return;
    if(o.userData.partId)ids.add(o.userData.partId);
    // Detailed optical meshes use subpart IDs; include their visible parent assembly.
    for(let parent=o.parent;parent&&parent!==root;parent=parent.parent)
     if(parent.userData.rackPart?.id)ids.add(parent.userData.rackPart.id);
   });
   return [...ids];
  },id);
  for(const key of [':tor:',':patch-panel:',':optical-housing',':cassette-module-0',':cassette-module-1',':cassette-module-2',':cassette-module-3',':cable-manager:',':pdu-a:',':pdu-b:'])
   assert.ok(inventory.some(part=>part.startsWith(id+key)),id+' missing visible original part '+key);
  if(id==='tor-a')assert.ok(inventory.some(part=>part.startsWith(id+':horizontal-manager-')),'Network rack retains horizontal cable management');
  return inventory;
 }

 for(const [id,kind] of [['gpu-0-0','gpu'],['compute-c','compute'],['storage','storage'],['tor-a','network']]){
  await page.evaluate(id=>{const T=window.__LS3D_TEST__;T.closeRackScene();T.setZone('hall');T.pickAsset(T.assetMap.get(id));},id);
  await page.locator('#detailBtn').click();
  await page.waitForFunction(()=>window.__LS3D_TEST__.rackIsolateActive);
  await assertMechanicalMotion(id);
  await page.locator('[data-rack-action="front-door"]').click();
  await page.waitForFunction(()=>window.__LS3D_TEST__.mechanical(window.__LS3D_TEST__.selected).door>.95);
  assert.equal(await page.evaluate(()=>window.LS3D_THREE.adapter.diagnostic.hardwareDetail),null,'Door opening must retain the original rack view');
  await assertRackInventory(id);
  await page.locator('#ls3dRackPanelToggle').click();
  await page.locator('#viewport').screenshot({path:out+'/original-'+kind+'-rack-open.png'});
  await page.locator('#ls3dRackPanelToggle').click();
  await page.locator('[data-rack-action="hardware-detail"]').click();
  await page.waitForFunction(kind=>window.LS3D_THREE.adapter.diagnostic.hardwareDetail?.kind===kind,kind);
  await page.waitForFunction(()=>window.__LS3D_TEST__.mechanical(window.__LS3D_TEST__.selected).slide>.95);
  await page.locator('#ls3dRackPanelToggle').click();
  await page.locator('#viewport').screenshot({path:out+'/03-'+kind+'-internal.png'});
  await page.locator('#ls3dRackPanelToggle').click();
  await page.locator('[data-rack-action="hardware-explode"]').click();
  await page.waitForFunction(()=>window.LS3D_THREE.adapter.diagnostic.hardwareDetail?.progress>.99);
  await page.locator('#ls3dRackPanelToggle').click();
  await page.locator('#viewport').screenshot({path:out+'/04-'+kind+'-exploded.png'});
  const state=await page.evaluate(()=>window.LS3D_THREE.adapter.diagnostic.hardwareDetail);
  assert.equal(state.kind,kind);assert.equal(state.exploded,true);results.push(state);
  await page.locator('#ls3dRackPanelToggle').click();
  await page.locator('[data-rack-action="hardware-explode"]').click();
  await page.waitForFunction(()=>window.LS3D_THREE.adapter.diagnostic.hardwareDetail?.progress<.01);
  await page.locator('#rackIsolateFit').click();
  await page.waitForFunction(()=>window.LS3D_THREE.adapter.diagnostic.rackViewMode==='rack-overview'&&window.__LS3D_TEST__.mechanical(window.__LS3D_TEST__.selected).slide<.01);
  assert.equal(await page.evaluate(()=>window.LS3D_THREE.adapter.diagnostic.hardwareDetail),null);
  await assertRackInventory(id);
  if(id==='tor-a'){
   for(const action of ['housing-select','cassette-select','manager-select','switch-select']){
    if(await page.locator('#ls3dRackMotionBody').isHidden())await page.locator('#ls3dRackPanelToggle').click();
    await page.locator('[data-rack-action="'+action+'"]').click();
    await page.waitForFunction(()=>!!window.LS3D_THREE.adapter.selectedPart);
    await page.locator('#rackIsolateFit').click();
    await page.waitForFunction(()=>!window.LS3D_THREE.adapter.selectedPart);
    await assertRackInventory(id);
   }
   await page.locator('#ls3dRackPanelToggle').click();
   await page.locator('#viewport').screenshot({path:out+'/restored-network-rack-all-components.png'});
  }
 }
 await page.setViewportSize({width:390,height:844});
 await page.waitForFunction(()=>window.LS3D_THREE.adapter.diagnostic.frames>10);
 assert.equal(await page.locator('#sceneThree').isVisible(),true);
 assert.equal(await page.evaluate(()=>window.__LS3D_TEST__.rackIsolateActive),true);
 if(!await page.locator('#ls3dRackMotionBody').isHidden())await page.locator('#ls3dRackPanelToggle').click();
 assert.ok(await page.locator('[data-rack-mechanical="front-door"]').isVisible());
 assert.ok(await page.locator('[data-rack-mechanical="cassette-slide"]').isVisible());
 await page.screenshot({path:out+'/05-mobile.png'});
 await page.evaluate(()=>window.__LS3D_TEST__.closeRackScene());
 await page.waitForFunction(()=>window.LS3D_THREE.adapter.diagnostic.hardwareDetail===null);
 assert.deepEqual(errors,[]);
 await writeFile(out+'/results.json',JSON.stringify({results,errors},null,2));
 console.log('PASS: neutral campus, four rack types, open tray, explode/reassemble, door open/close and cassette out/in transforms verified with collapsed controls; original assemblies preserved and restored, cleanup, mobile, no runtime errors');
}finally{await browser.close();}
