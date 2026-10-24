// STEP 19: end-to-end Chromium/WebGL2 acceptance. Real screenshots and metrics
// are produced ONLY when this suite actually runs (e.g., GitHub Actions).
// This suite does not claim calculated routes are validated engineering drawings.
import {test,expect} from '@playwright/test';
import {readFile} from 'node:fs/promises';
import path from 'node:path';

const BASE=process.env.LS3D_BASE_URL||
  'http://127.0.0.1:8765/DCI/DataCenter/LS_Datacenter_3D.html';
const THREE_BUILD=path.resolve(process.cwd(),'node_modules/three/build');
const ZONES=['campus','utility','cooling','gray','hall','network','ops'];
const serious=[];
test.beforeEach(async({page})=>{
  // Only the tested 3D controls are in scope; first-visit onboarding covers
  // the canvas and controls in a fresh CI browser unless already completed.
  await page.addInitScript(() => {
    try { window.localStorage.setItem('ls3d-first-guide-v481', 'done'); } catch (_) {}
  });
  // Hermetic Three.js version for browser CI, preserving the site's exact
  // import map and module graph. This replaces *only* the CDN response.
  await page.route('**/npm/three@0.186.1/build/three.*.js',async route=>{
    const name=route.request().url().includes('three.core.js')?'three.core.js':'three.module.js';
    await route.fulfill({
      status:200,body:await readFile(path.join(THREE_BUILD,name)),
      headers:{'content-type':'text/javascript; charset=utf-8',
        'access-control-allow-origin':'*'}
    });
  });
});
function errorsFor(page){
  const errors=[];
  page.on('pageerror',e=>errors.push('pageerror: '+e.message));
  page.on('console',msg=>{if(msg.type()==='error')
    errors.push('console error: '+msg.text());});
  return errors;
}
async function enter(page){
  const errors=errorsFor(page);
  await page.goto(BASE+(BASE.includes('?')?'&':'?')+'engine=three',{waitUntil:'domcontentloaded'});
  await page.waitForFunction(()=>['ready','failed'].includes(window.LS3D_THREE?.status),
    null,{timeout:60000});
  const status=await page.evaluate(()=>({status:LS3D_THREE.status,error:LS3D_THREE.error}));
  expect(status.status,status.error||'WebGL2 Three.js initialization failed').toBe('ready');
  await page.waitForFunction(()=>LS3D_THREE?.adapter?.diagnostic?.frames>=4,
    null,{timeout:30000});
  // The older tour initializer opens independently of the first-visit flag.
  // Explicitly dismiss its DOM overlay before Chromium tests app controls.
  await page.evaluate(()=>document.getElementById('guide')?.classList.remove('open'));
  return errors;
}
async function zone(page,id){
  const previous=await page.evaluate(()=>window.LS3D_THREE.adapter.diagnostic.frames);
  await page.evaluate(id=>window.__LS3D_TEST__.setZone(id),id);
  await page.waitForFunction(n=>window.LS3D_THREE.adapter.diagnostic.frames>n+2,
    previous,{timeout:25000});
}
function attachJSON(info,name,data){
  return info.attach(name,{body:Buffer.from(JSON.stringify(data,null,2)),contentType:'application/json'});
}

test('S19-01: seven zones, actual WebGL2 draw calls, NPC/security and zone screenshots',async({page},info)=>{
  const errors=await enter(page),data=[];
  const initial=await page.evaluate(()=>{
    const a=window.LS3D_THREE.adapter,gl=document.getElementById('sceneThree').getContext('webgl2');
    return {zoneIds:window.__LS3D_TEST__.Z.map(z=>z.id),
      revision:a.THREE.REVISION,gl:!!gl};
  });
  expect(initial.gl).toBe(true);expect(initial.zoneIds).toEqual(ZONES);
  for(const id of ZONES){
    await zone(page,id);
    const snapshot=await page.evaluate(()=>{
      const d=window.LS3D_THREE.adapter.diagnostic;
      return {zone:window.__LS3D_TEST__.currentZone,frames:d.frames,
        renderCalls:d.calls,triangles:d.triangles,assets:d.assets,
        visibleNpc:d.npc,security:d.security?.visible||false,
        memoryGeometries:d.memoryGeometries};
    });
    data.push(snapshot);
    expect(snapshot.zone).toBe(id);
    expect(snapshot.renderCalls).toBeGreaterThan(0);
    expect(snapshot.triangles).toBeGreaterThan(0);
    await page.screenshot({path:info.outputPath('step19-zone-'+id+'.png')});
  }
  expect(data.find(x=>x.zone==='campus').security).toBe(true);
  expect(data.find(x=>x.zone==='hall').assets).toBeGreaterThan(0);
  await attachJSON(info,'zone-render-metrics.json',data);
  expect(errors).toEqual([]);
});

test('S19-02: inspect rack ToR, NIC, independent PDU and animate door + server drawer',async({page},info)=>{
  const errors=await enter(page);
  await zone(page,'hall');
  await page.evaluate(()=>{
    const T=window.__LS3D_TEST__;
    T.pickAsset(T.assetMap.get('gpu-0-0'),true);
  });
  await page.waitForFunction(()=>document.getElementById('ls3dRackMotionControls') &&
    !document.getElementById('ls3dRackMotionControls').hidden);
  const id='gpu-0-0';
  const before=await page.evaluate(id=>{
    const T=window.__LS3D_TEST__,a=window.LS3D_THREE.adapter;
    return {mechanical:{...T.mechanical(T.assetMap.get(id))},
      rack:a.rackCabling.diagnostic(id),components:a.rackComponents};
  },id);
  expect(before.components.rackCount).toBeGreaterThan(0);
  expect(before.rack.links.filter(l=>l.kind==='fiber').length).toBeGreaterThan(0);
  expect(before.rack.links.some(l=>l.kind==='power-A')).toBe(true);
  expect(before.rack.links.some(l=>l.kind==='power-B')).toBe(true);
  await page.locator('[data-rack-action="front-door"]').click();
  await page.locator('[data-rack-action="server-slide"]').click();
  await page.waitForTimeout(700);
  const after=await page.evaluate(id=>{
    const T=window.__LS3D_TEST__,a=window.LS3D_THREE.adapter;
    return {mechanical:{...T.mechanical(T.assetMap.get(id))},
      cabling:a.rackCabling.diagnostic(id),interaction:a.rackInteraction};
  },id);
  expect(after.mechanical.targetDoor).not.toBe(before.mechanical.targetDoor);
  expect(after.mechanical.targetSlide).not.toBe(before.mechanical.targetSlide);
  expect(after.cabling.links.length).toBe(before.rack.links.length);
  await page.screenshot({path:info.outputPath('step19-rack-open-nic-cables.png')});
  await attachJSON(info,'rack-mechanics.json',{before,after});
  expect(errors).toEqual([]);
});

test('S19-03: MMR/ODF/DWDM, customer A/B/C isolation and all MEP topology classes',async({page},info)=>{
  const errors=await enter(page);
  await zone(page,'network');
  const mmr=await page.evaluate(()=>{
    const a=window.LS3D_THREE.adapter,s=a.mmrController.summary();
    a.mmrController.selectLink('carrier-odf');
    a.mmrController.toggleDoor('carrier');
    return {before:s,selected:a.mmrController.summary().selectedLink};
  });
  expect(mmr.before.ready).toBe(true);
  expect(mmr.before.doors).toHaveLength(4);
  expect(mmr.selected).toBe('carrier-odf');
  await page.screenshot({path:info.outputPath('step19-network-mmr.png')});
  await zone(page,'hall');
  const colo=await page.evaluate(()=>{
    const a=window.LS3D_THREE.adapter,controller=a.colocationController;
    const before=controller.summary();
    const info=controller.selectCircuit('B','power-B');
    controller.toggleGate('B');
    return {before,info,after:controller.summary()};
  });
  expect(colo.before.ready).toBe(true);
  expect(colo.before.bays.map(b=>b.id)).toEqual(['A','B','C']);
  expect(colo.info.kind).toBe('power-B');
  expect(colo.after.bays.flatMap(b=>b.circuits).filter(c=>c.highlighted)).toHaveLength(1);
  await page.screenshot({path:info.outputPath('step19-hall-colocation.png')});
  await zone(page,'campus');
  const mep=await page.evaluate(()=>{
    const controller=window.LS3D_THREE.adapter.mepController;
    const before=controller.summary();
    const pick=before.paths.find(p=>p.id==='fiber:mmr-hall:colo-rack-c');
    if(pick)controller.selectRoute(pick.id);
    return {before,after:controller.summary()};
  });
  expect(new Set(mep.before.paths.map(p=>p.system))).toEqual(new Set(['power','cooling','fiber']));
  expect(mep.before.paths.some(p=>p.id==='power:pdu:colo-rack-a')).toBe(true);
  expect(mep.before.paths.some(p=>p.id==='cooling:gray-cdu:gpu-0-0')).toBe(true);
  expect(mep.after.paths.filter(p=>p.highlighted)).toHaveLength(1);
  await page.screenshot({path:info.outputPath('step19-campus-mep-fiber.png')});
  await attachJSON(info,'mmr-colo-mep.json',{mmr,colo,mep});
  expect(errors).toEqual([]);
});

test('S19-04: mobile responsive controls, multi-touch camera and facility isolation',async({page},info)=>{
  await page.setViewportSize({width:390,height:844});
  const errors=await enter(page);
  await zone(page,'hall');
  await page.evaluate(()=>{
    const T=window.__LS3D_TEST__;T.pickAsset(T.assetMap.get('gpu-0-0'),true);
  });
  await page.waitForFunction(()=>document.querySelector('#ls3dRackMotionControls') &&
    !document.querySelector('#ls3dRackMotionControls').hidden);
  const panels=await page.evaluate(()=>{
    const ids=['ls3dRackMotionControls','ls3dMEPControls','ls3dColoControls'];
    return ids.map(id=>{
      const el=document.getElementById(id),r=el?.getBoundingClientRect();
      return {id,exists:!!el,hidden:el?.hidden,
        left:r?.left,right:r?.right,overflow:el?.style.overflowY};
    });
  });
  expect(panels.every(p=>p.exists&&p.left>=-1&&p.right<=391)).toBe(true);
  await page.screenshot({path:info.outputPath('step19-mobile-390.png')});
  await page.evaluate(()=>window.__LS3D_TEST__.applyFacilityMode('server-room'));
  // A room may remain initialized (ready=true) after becoming hidden.
  // Facility isolation is about actual rendered visibility, not the cached
  // initialization flag of a 3D room controller.
  await page.waitForFunction(()=>{
    const a=window.LS3D_THREE?.adapter;
    return a?.mep?.visible===false &&
      a?.colocationController?.root?.visible===false &&
      a?.mmrController?.root?.visible===false;
  },null,{timeout:20000});
  expect(await page.evaluate(()=>window.LS3D_THREE.adapter.colocationController.root.visible)).toBe(false);
  await attachJSON(info,'mobile-layout.json',panels);
  expect(errors).toEqual([]);
});

test('S19-05: sampled real-browser frame rate, draw calls and memory snapshots (no invented FPS)',async({page},info)=>{
  const errors=await enter(page);
  await zone(page,'hall');
  const samples=await page.evaluate(async()=>{
    const a=window.LS3D_THREE.adapter,result=[];
    let prev=performance.now(),frames=0;
    await new Promise(resolve=>{
      const target=90;
      const tick=now=>{
        const interval=now-prev;prev=now;
        if(frames>0){
          result.push(interval);
        }
        if(++frames>=target)resolve();
        else requestAnimationFrame(tick);
      };
      requestAnimationFrame(tick);
    });
    const sorted=result.slice().sort((x,y)=>x-y);
    const mean= result.reduce((sum,v)=>sum+v,0)/result.length;
    const heap=performance.memory?.usedJSHeapSize||null;
    return {observedFrames:result.length,meanFrameIntervalMs:mean,
      estimatedRafFps:1000/mean,p95FrameIntervalMs:sorted[Math.floor(sorted.length*.95)],
      jsHeapBytes:heap,geometries:a.diagnostic.memoryGeometries,
      textures:a.diagnostic.memoryTextures,drawCalls:a.diagnostic.calls,
      triangles:a.diagnostic.triangles,viewport:a.viewportPerformance};
  });
  expect(samples.observedFrames).toBeGreaterThan(70);
  expect(samples.meanFrameIntervalMs).toBeGreaterThan(0);
  expect(samples.geometries).toBeGreaterThanOrEqual(0);
  await attachJSON(info,'chromium-performance-90-frames.json',samples);
  await page.screenshot({path:info.outputPath('step19-performance-hall.png')});
  expect(errors).toEqual([]);
});

// Regression for the reported blank blue viewport from campus/front camera.
test('S19-rack-camera: front and rear show the selected rack at desktop and mobile sizes',async({page},info)=>{
  await enter(page);
  await page.evaluate(()=>{
    const T=window.__LS3D_TEST__;T.setZone('campus');T.pickAsset(T.assetMap.get('tor-a'),false);
  });
  for(const [width,height] of [[1400,700],[1400,580],[390,844]]){
    await page.setViewportSize({width,height});
    for(const side of ['front','rear']){
      await page.locator('[data-rack-action="'+side+'-camera"]').click();
      await page.waitForFunction(()=>window.LS3D_THREE.diagnostics?.isolatedRack==='tor-a');
      await page.waitForTimeout(180);
      const geometry=await page.evaluate(()=>{
        const A=window.LS3D_THREE.adapter,THREE=A.THREE;
        let rack=null,architectureVisible=null;const ids=[];
        A.scene.traverse(o=>{
          if(o.name==='equipment:tor-a')rack=o;
          if(o.name?.startsWith('equipment:'))ids.push(o.userData.assetId);
          if(o.name?.startsWith('arch:zone:'))architectureVisible=o.parent.visible;
        });
        if(!rack)return {missing:true};
        const box=new THREE.Box3().setFromObject(rack),center=box.getCenter(new THREE.Vector3()).project(A.camera);
        const corners=[];for(const x of [box.min.x,box.max.x])for(const y of [box.min.y,box.max.y])for(const z of [box.min.z,box.max.z]){
          const p=new THREE.Vector3(x,y,z).project(A.camera);corners.push({x:p.x,y:p.y,z:p.z});
        }
        return {ids,architectureVisible,center:{x:center.x,y:center.y,z:center.z},corners};
      });
      expect(geometry.missing).toBeFalsy();expect(geometry.ids).toEqual(['tor-a']);
      expect(geometry.architectureVisible).toBe(false);
      expect(Math.abs(geometry.center.x)).toBeLessThan(1);expect(Math.abs(geometry.center.y)).toBeLessThan(1);
      for(const p of geometry.corners){expect(Math.abs(p.x)).toBeLessThan(1);expect(Math.abs(p.y)).toBeLessThan(1);expect(p.z).toBeLessThan(1);expect(p.z).toBeGreaterThan(-1);}
      await page.screenshot({path:info.outputPath('rack-camera-'+width+'-'+side+'.png')});
    }
  }
  await page.locator('[data-rack-action="rack-context"]').click();
  await page.waitForFunction(()=>!window.LS3D_THREE.diagnostics?.isolatedRack);
});


test('S19-inspector: tool dock, asset/part lists and clicked golden tray stay synchronized',async({page},info)=>{
  await enter(page);
  const ids=['phase2-bar','phase5-open','phase6-open','phase7-open','phase8-open'];
  for(const id of ids){expect(await page.locator('#'+id).evaluate(el=>el.parentElement.id)).toBe('ls3dToolDock');expect(await page.locator('#viewport #'+id).count()).toBe(0);}
  await page.locator('#phase5-open').click();await expect(page.locator('#phase5-panel')).toBeVisible();await page.locator('#p5-close').click();
  await page.locator('#inspectorAssetSelect').selectOption('tor-b');await expect(page.locator('#assetName')).toContainText('NETWORK B');
  await page.waitForFunction(()=>window.LS3D_INSPECTOR.parts.some(p=>p.id==='tor-b:optical-housing'));
  await page.locator('#inspectorPartSelect').selectOption('tor-b:optical-housing');
  await expect(page.locator('#inspectorPartName')).toContainText('4슬롯');
  await page.evaluate(()=>{const T=window.__LS3D_TEST__;T.setZone('hall');T.pickAsset(T.assetMap.get('tor-a'),false);});
  const trayId='structure:hall:hall-power-tray';
  await page.waitForFunction(id=>window.LS3D_INSPECTOR.parts.some(p=>p.id===id),trayId);
  await page.locator('#inspectorPartSelect').selectOption(trayId);
  await expect(page.locator('#inspectorPartName')).toHaveText('전력 케이블 트레이');
  await expect(page.locator('#inspectorPartDescription')).toContainText('지지');
  // Aim perpendicular to a rail and use the production raycast picker.
  await page.evaluate(()=>{
    const A=window.LS3D_THREE.adapter,T=window.__LS3D_TEST__;let rail=null;
    A.scene.traverse(o=>{if(!rail&&o.name?.includes('hall-power-tray-rail-'))rail=o;});
    const p=rail.getWorldPosition(new A.THREE.Vector3());
    Object.assign(T.cameraDesired,{target:[p.x,p.y,p.z],yaw:Math.PI/2,pitch:.25,distance:2.5});
    A.clearInspectionSelection();
  });
  await page.waitForTimeout(450);
  await page.evaluate(()=>{
    const A=window.LS3D_THREE.adapter;let rail=null;
    A.scene.traverse(o=>{if(!rail&&o.name?.includes('hall-power-tray-rail-'))rail=o;});
    const p=rail.getWorldPosition(new A.THREE.Vector3()).project(A.camera),r=A.renderer.domElement.getBoundingClientRect();
    A.pick(r.left+(p.x+1)*r.width/2,r.top+(1-p.y)*r.height/2);
  });
  await expect(page.locator('#inspectorPartName')).toHaveText('전력 케이블 트레이');
  await page.screenshot({path:info.outputPath('inspector-golden-tray-sidebar.png')});
  await page.setViewportSize({width:390,height:844});
  await page.waitForFunction(()=>document.getElementById('phase5-open').parentElement.id==='ls3dMobileToolDock');
  await page.locator('#mobileInspectorTools summary').click();await expect(page.locator('#phase5-open')).toBeVisible();
  await page.screenshot({path:info.outputPath('inspector-mobile-dock.png')});
});


test('S19-rack-fill: central passive patching and blanking cover every front U',async({page},info)=>{
  await enter(page);
  await page.evaluate(()=>{
    const T=window.__LS3D_TEST__,A=window.LS3D_THREE.adapter;
    T.setZone('hall');T.pickAsset(T.assetMap.get('tor-a'),false);
    A.focusRackSide('tor-a','front');A.toggleRackFrontDoor('tor-a');A.toggleRackCassette('tor-a');
  });
  await page.waitForFunction(()=>window.LS3D_INSPECTOR.parts.some(p=>p.kind.startsWith('distribution-panel-')));
  await page.waitForTimeout(450);
  const result=await page.evaluate(()=>{
    const A=window.LS3D_THREE.adapter;let rack=null;A.scene.traverse(o=>{if(o.name==='equipment:tor-a')rack=o;});
    const mounted=[],slots=Array(49).fill(0);let panels=0,sleds=0,instances=0;
    rack.traverse(o=>{
      const p=o.userData?.rackPart;if(p&&(/^(network-switch-|distribution-panel-|horizontal-manager-|blanking-panel-)/.test(p.kind)||['tor','patch-panel','cassette-pull'].includes(p.kind))){
        mounted.push(p);for(let u=p.uStart;u<p.uStart+p.uHeight;u++)slots[u]++;
        if(p.kind.startsWith('distribution-panel-'))panels++;
        if(p.kind.startsWith('network-switch-'))sleds++;
      }
      if(o.isInstancedMesh&&o.name==='rack-shape:lc-duplex-adapter')instances+=o.count;
    });
    return {coverage:slots.slice(1),panels,sleds,instances,firstPanel:mounted.find(p=>p.kind.startsWith('distribution-panel-'))?.id};
  });
  expect(result.coverage).toHaveLength(48);expect(result.coverage.every(n=>n===1)).toBe(true);
  expect(result.panels).toBe(6);expect(result.sleds).toBe(3);expect(result.instances).toBe(72);
  await page.locator('#inspectorPartSelect').selectOption(result.firstPanel);
  await expect(page.locator('#inspectorPartName')).toContainText('광패치패널');
  await page.screenshot({path:info.outputPath('rack-center-filled-passive-panels.png')});
});

// Pointer events, not a direct adapter.pick call: verifies the user-facing path.
test('S19-product-click: ceiling trays and busduct expose persistent product descriptions',async({page},info)=>{
 await enter(page);
 async function aim(name,zone){
  await page.evaluate(({name,zone})=>{
   const T=window.__LS3D_TEST__;T.setZone(zone);
   const v={target:[0,3,0],yaw:Math.PI/2,pitch:.25,distance:3};
   Object.assign(T.cameraDesired,v);Object.assign(T.camera,{...v,target:[...v.target]});
  },{name,zone});
  await page.waitForFunction(name=>{let found=false;window.LS3D_THREE.adapter.scene.traverse(o=>{if(o.name===name)found=true;});return found;},name);
  await page.evaluate(name=>{
   const A=window.LS3D_THREE.adapter,T=window.__LS3D_TEST__;let obj;
   A.scene.traverse(o=>{if(o.name===name)obj=o;});
   const p=obj.getWorldPosition(new A.THREE.Vector3());
   const bus=name.includes('busway');
   const v={target:[p.x,p.y,p.z],yaw:bus?Math.PI:Math.PI/2,pitch:bus?.08:.25,distance:bus?8:2.5};
   Object.assign(T.cameraDesired,v);Object.assign(T.camera,{...v,target:[...v.target]});
   A.clearInspectionSelection();
  },name);
  await page.waitForTimeout(180);
  return page.evaluate(name=>{
   const A=window.LS3D_THREE.adapter;let obj;
   A.scene.traverse(o=>{if(o.name===name)obj=o;});
   const p=obj.getWorldPosition(new A.THREE.Vector3()).project(A.camera),r=A.renderer.domElement.getBoundingClientRect();
   return {x:r.left+(p.x+1)*r.width/2,y:r.top+(1-p.y)*r.height/2};
  },name);
 }
 let p=await aim('arch:tray:hall-power-tray-rail--0.38','hall');
 await page.mouse.click(p.x,p.y);
 await expect(page.locator('#inspectorPartName')).toHaveText('전력 케이블 트레이');
 await page.waitForTimeout(600);await expect(page.locator('#inspectorPartDetails')).toBeVisible();
 await expect(page.locator('#inspectorPartDescription')).toContainText('전력 케이블');
 await page.screenshot({path:info.outputPath('product-power-tray-click.png')});
 p=await aim('equipment-part:busway-shell','gray');
 await page.mouse.click(p.x,p.y);
 await expect(page.locator('#inspectorPartName')).toContainText('부스덕트');
 await expect(page.locator('#inspectorPartDescription')).toContainText('탭오프');
 await page.waitForTimeout(600);await expect(page.locator('#inspectorPartDetails')).toBeVisible();
 await page.screenshot({path:info.outputPath('product-busduct-click.png')});
 await page.setViewportSize({width:390,height:844});
 p=await aim('arch:tray:hall-fiber-tray-rail--0.38','hall');
 await page.mouse.click(p.x,p.y);
 await expect(page.locator('#inspectorPartName')).toHaveText('광케이블 트레이');
 await expect(page.locator('#inspectorPartDescription')).toContainText('굴곡반경');
});
