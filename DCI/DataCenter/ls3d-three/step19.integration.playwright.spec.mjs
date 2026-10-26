// LS Datacenter 3D STEP 19: integrated Chromium/WebGL acceptance.
// Tests run against a local, checked-out repository at LS3D_BASE_URL.
// No deployed GitHub Pages result is claimed by these tests.
// Run: npx playwright test DCI/DataCenter/ls3d-three/step19.integration.playwright.spec.mjs
import {test,expect} from '@playwright/test';
import {readFile} from 'node:fs/promises';
import {mkdir,writeFile} from 'node:fs/promises';

const BASE=process.env.LS3D_BASE_URL||
 'http://127.0.0.1:8765/DCI/DataCenter/LS_Datacenter_3D.html';
const THREE_BUILD=new URL('../../../node_modules/three/build/',import.meta.url);
async function ready(page,zone='campus'){
  const issues=[];
  page.on('pageerror',e=>issues.push('pageerror: '+e.message));
  page.on('console',msg=>{
    if(msg.type()==='error' && /three|webgl|module|uncaught|failed/i.test(msg.text()))
      issues.push('console: '+msg.text());
  });
  // STEP 19 validates 3D controls after onboarding. A fresh CI browser has no
  // visitor history, so the first-visit guide otherwise overlays buttons and
  // intercepts both clicks and mobile touch gestures. Keep onboarding enabled
  // for real visitors; skip it only in these app-control acceptance tests.
  await page.addInitScript(() => {
    try { window.localStorage.setItem('ls3d-first-guide-v481', 'done'); } catch (_) {}
  });
  // Local known version, no dependence on third-party CDN availability.
  await page.route('**/npm/three@0.186.1/build/three.*.js',async route=>{
    const name=route.request().url().includes('three.core.js')?'three.core.js':'three.module.js';
    await route.fulfill({status:200,contentType:'text/javascript',
      headers:{'access-control-allow-origin':'*'},
      body:await readFile(new URL(name,THREE_BUILD))});
  });
  await page.goto(BASE+(BASE.includes('?')?'&':'?')+'engine=three',{waitUntil:'domcontentloaded'});
  await page.waitForFunction(()=>['ready','failed'].includes(window.LS3D_THREE?.status),
    null,{timeout:45000});
  const status=await page.evaluate(()=>({
    state:window.LS3D_THREE.status,error:window.LS3D_THREE.error,
    gl:!!document.querySelector('#sceneThree')
  }));
  expect(status.state,status.error||'Three bootstrap failure').toBe('ready');
  expect(status.gl).toBe(true);
  await page.waitForFunction(()=>window.LS3D_THREE.renders>=3,null,{timeout:30000});
  const displayChoice=page.locator('#displayModeChoice');
  if(await displayChoice.isVisible())
    await page.locator('[data-display-choice="'+(page.viewportSize()?.width<=790?'mobile':'pc')+'"]').click();
  // The legacy tour dialog can be opened by a separate initializer even when
  // the newer first-visit localStorage flag is pre-seeded. Dismiss the actual
  // overlay before testing independent rack/MMR/mobile controls.
  await page.evaluate(()=>document.getElementById('guide')?.classList.remove('open'));
  await page.evaluate(z=>window.__LS3D_TEST__.setZone(z),zone);
  await page.waitForFunction(z=>window.__LS3D_TEST__.currentZone===z,zone);
  return issues;
}
function verify(issues){expect(issues,issues.join('; ')).toEqual([]);}
async function snapshot(page,info,file){
  await page.screenshot({path:info.outputPath(file),fullPage:false,animations:'disabled'});
}
const settle=page=>page.waitForTimeout(380);

test('19-A: seven zones, real Three scene frames, GPU geometry counters and no JS errors',async({page},info)=>{
  const errors=await ready(page);
  const zones=await page.evaluate(()=>window.__LS3D_TEST__.Z.map(z=>z.id));
  expect(zones).toEqual(['campus','utility','cooling','gray','hall','network','ops']);
  const snapshots=[];
  for(const z of zones){
    await page.evaluate(x=>window.__LS3D_TEST__.setZone(x),z);
    await settle(page);
    const d=await page.evaluate(()=>({
      zone:window.__LS3D_TEST__.currentZone,
      calls:window.LS3D_THREE.adapter.diagnostic.calls,
      geometry:window.LS3D_THREE.adapter.diagnostic.memoryGeometries,
      frames:window.LS3D_THREE.renders,
      assets:window.LS3D_THREE.adapter.diagnostic.assets,
      cameraNear:window.LS3D_THREE.adapter.camera.near
    }));
    expect(d.zone).toBe(z);expect(d.calls).toBeGreaterThan(0);
    expect(d.geometry).toBeGreaterThan(0);
    expect(d.cameraNear).toBeGreaterThan(0);
    snapshots.push(d);
    if(['campus','hall','network'].includes(z))
      await snapshot(page,info,'step19-zone-'+z+'.png');
  }
  await writeFile(info.outputPath('step19-seven-zones-metrics.json'),
    JSON.stringify(snapshots,null,2),'utf8');
  verify(errors);
});

test('19-B: campus NPC rigs and eight-stage visitor checkpoint security',async({page},info)=>{
  const errors=await ready(page,'campus');
  const result=await page.evaluate(()=>{
    const a=window.LS3D_THREE.adapter,s=a.securityController;
    const rigs=[];
    a.scene.traverse(o=>{if(o.name?.startsWith('npc:'))rigs.push(o.name);});
    const phases=s.phases.map(x=>x.name);
    const first=s.simulate(1);
    const later=s.simulate(48);
    return {phases,rigs,first,later,diagnostic:s.diagnostic(),
      npc:a.diagnostic.npc,npcActions:a.diagnostic.npcAnimation?.actions};
  });
  expect(result.phases).toHaveLength(8);
  expect(result.first.active).toBe(true);
  expect(result.later.active).toBe(true);
  expect(result.later.gateOpen).toBe(true);
  expect(result.diagnostic.stageCount).toBe(8);
  expect(result.npc).toBeGreaterThanOrEqual(1);
  await snapshot(page,info,'step19-campus-security.png');
  verify(errors);
});

test('19-C: rack front/rear, selected sled, cassette, ToR and independent PDU A/B',async({page},info)=>{
  const errors=await ready(page,'hall');
  const assetId=await page.evaluate(()=>{
    const T=window.__LS3D_TEST__,rack=T.visibleAssets().find(a=>
      a.zone==='hall'&&T.isMechanicalRack(a));
    if(!rack)throw Error('interactive rack missing');
    T.pickAsset(rack,true);return rack.id;
  });
  await page.waitForFunction(()=>!!document.querySelector('#ls3dRackMotionControls')&&
    !document.querySelector('#ls3dRackMotionControls').hidden);
  // Browser actions below are real DOM interactions.
  await page.locator('[data-rack-action="rear-camera"]').click();
  await page.locator('[data-rack-action="rear-door"]').click();
  await settle(page);
  await page.locator('[data-rack-action="next-server"]').click();
  await page.locator('[data-rack-action="server-slide"]').click();
  await page.locator('[data-rack-action="cassette-slide"]').click();
  await settle(page);
  const state=await page.evaluate(id=>{
    const adapter=window.LS3D_THREE.adapter;
    let door=null,slide=null,cassette=null,tor=0,nic=0;
    adapter.scene.traverse(o=>{
      if(o.name==='equipment:'+id)o.traverse(p=>{
        if(p.name==='rear-door-hinge')door=p.rotation.y;
        if(p.name==='fiber-cassette')cassette=p.position.z;
        if(p.name.startsWith('rack-link:tor-port-'))tor++;
        if(p.name.startsWith('rack-server:nic-board'))nic++;
        if(p.userData?.rackPart?.id===adapter.rackInteraction.selectedServer?.selected)
          slide=p.position.z;
      });
    });
    return {door,slide,cassette,tor,nic,
      rack:adapter.rackComponents,links:adapter.rackCabling.diagnostic(id),
      focus:adapter.rackInteraction.focus?.side};
  },assetId);
  expect(state.rack.rackCount).toBeGreaterThan(0);
  expect(state.focus).toBe('rear');
  expect(state.door).not.toBeNull();
  expect(state.tor).toBeGreaterThanOrEqual(8);
  expect(state.nic).toBeGreaterThan(0);
  expect(state.links.links.filter(x=>x.kind==='power-A').length).toBeGreaterThan(0);
  expect(state.links.links.filter(x=>x.kind==='power-B').length).toBeGreaterThan(0);
  await snapshot(page,info,'step19-rack-rear-sled-cassette.png');
  await page.locator('[data-rack-action="front-camera"]').click();
  verify(errors);
});

test('19-D: Network MMR Carrier/ODF/DWDM/Core with 3D door and optical route selection',async({page},info)=>{
  const errors=await ready(page,'network');
  await page.waitForFunction(()=>window.LS3D_THREE.adapter.mmr?.ready===true);
  await expect(page.locator('#ls3dMMRControls')).toBeVisible();
  await page.locator('[data-mmr-door="carrier"]').click();
  await page.locator('[data-mmr-link="odf-dwdm"]').click();
  await settle(page);
  const result=await page.evaluate(()=>{
    const adapter=window.LS3D_THREE.adapter,s=adapter.mmrController.summary(),names=[];
    adapter.mmrController.root.traverse(o=>names.push(o.name));
    return {s,names};
  });
  expect(result.s.bays).toHaveLength(4);
  expect(result.s.selectedLink).toBe('odf-dwdm');
  expect(result.s.selectedRoute.from).toContain('ODF');
  expect(result.s.doors.find(x=>x.id==='carrier').opening).toBeGreaterThan(0);
  for(const token of ['carrier-demarc-plate','dwdm-muxponder-0',
    'odf-patch-shelf-0','core-interface-card-0'])
    expect(result.names).toContain('mmr:'+token);
  await snapshot(page,info,'step19-mmr-dwdm.png');verify(errors);
});

test('19-E: Co-location A/B/C fences, doors, independent optical and dual power routes',async({page},info)=>{
  const errors=await ready(page,'hall');
  await page.waitForFunction(()=>window.LS3D_THREE.adapter.colocation?.ready===true);
  await expect(page.locator('#ls3dColoControls')).toBeVisible();
  await page.locator('[data-colo-gate="C"]').click();
  await page.locator('[data-colo-circuit="B:power-B"]').click();
  await settle(page);
  const d=await page.evaluate(()=>{
    const s=window.LS3D_THREE.adapter.colocationController.summary();
    return {s,visible:window.LS3D_THREE.adapter.colocationController.root.visible};
  });
  expect(d.visible).toBe(true);
  expect(d.s.bays).toHaveLength(3);
  expect(d.s.selected).toBe('B');
  expect(d.s.selectedCircuit).toBe('power-B');
  expect(d.s.bays.flatMap(x=>x.circuits).filter(x=>x.highlighted)).toHaveLength(1);
  expect(d.s.bays.find(x=>x.id==='C').gateOpening).toBeGreaterThan(0);
  expect(d.s.layoutIssue).toBeNull();
  await snapshot(page,info,'step19-colo-fence-and-power.png');verify(errors);
});

test('19-F: MEP fiber/power/cooling 3D paths, selective highlight and no fabricated ratings',async({page},info)=>{
  const errors=await ready(page,'campus');
  await page.waitForFunction(()=>window.LS3D_THREE.adapter.mep?.total>0);
  const before=await page.evaluate(()=>window.LS3D_THREE.adapter.mep);
  expect(new Set(before.paths.map(x=>x.system))).toEqual(new Set(['power','cooling','fiber']));
  expect(before.paths.every(p=>p.verified===false&&p.lengthMeters===null)).toBe(true);
  const chosen='fiber:mmr-hall:colo-rack-a';
  const focus=await page.evaluate(id=>window.LS3D_THREE.adapter.mepController.selectRoute(id),chosen);
  expect(focus?.id).toBe(chosen);
  const selected=await page.evaluate(()=>window.LS3D_THREE.adapter.mepController.summary());
  expect(selected.paths.filter(p=>p.highlighted)).toHaveLength(1);
  expect(selected.paths.every(p=>p.visible)).toBe(true);
  await snapshot(page,info,'step19-campus-mep-fiber.png');
  await page.evaluate(()=>window.LS3D_THREE.adapter.mepController.selectSystem('cooling'));
  const cool=await page.evaluate(()=>window.LS3D_THREE.adapter.mepController.summary());
  expect(cool.paths.every(p=>p.highlighted===(p.system==='cooling'))).toBe(true);
  verify(errors);
});

test('19-G: narrow mobile screen and two-finger pan/pinch with no false rack pick',async({browser},info)=>{
  const context=await browser.newContext({viewport:{width:390,height:844},isMobile:true,
    hasTouch:true,deviceScaleFactor:2});
  const page=await context.newPage(),errors=await ready(page,'hall');
  await page.waitForFunction(()=>window.LS3D_THREE.adapter.colocation?.ready===true);
  const before=await page.evaluate(()=>({
    desired:JSON.stringify(window.__LS3D_TEST__.cameraDesired),
    selected:window.__LS3D_TEST__.selected?.id
  }));
  // On mobile the metrics and introductory sections put the 3D scene below
  // the initial fold. A user scrolls to the viewport before touching it.
  // CDP touch coordinates are viewport-relative, so scroll first as well.
  await page.locator('#sceneThree').scrollIntoViewIfNeeded();
  await page.waitForTimeout(200);
  // Locate two genuinely uncovered points on the mobile Three.js canvas.
  // Fixed viewport fractions may fall under Co-location/MEP HUD overlays,
  // so those touches never reach the real canvas pointer handlers.
  const gesture=await page.evaluate(()=>{
    const canvas=document.getElementById('sceneThree'),r=canvas.getBoundingClientRect();
    const ys=[.56,.68,.76,.44,.34,.84];
    const xs=[.22,.30,.38,.46,.54,.62,.70];
    const exposed=(x,y)=>document.elementFromPoint(x,y)===canvas;
    for(const fy of ys)for(const fx of xs){
      const p1={x:r.x+r.width*fx,y:r.y+r.height*fy};
      for(const fx2 of xs){
        if(fx2-fx<.16)continue;
        const p2={x:r.x+r.width*fx2,y:p1.y};
        if(p2.x+44>r.right-6||p1.y+14>r.bottom-6)continue;
        if(exposed(p1.x,p1.y)&&exposed(p2.x,p2.y)&&
          exposed(p1.x+22,p1.y+12)&&exposed(p2.x+42,p2.y+12)){
          return {p1,p2,canvas:{width:r.width,height:r.height}};
        }
      }
    }
    return null;
  });
  expect(gesture,'No unoccluded region available for a two-finger canvas gesture').toBeTruthy();
  const {p1,p2}=gesture;
  const cdp=await context.newCDPSession(page);
  await cdp.send('Input.dispatchTouchEvent',{type:'touchStart',touchPoints:[
    {...p1,id:1},{...p2,id:2}
  ]});
  // Send intermediate movement so real pointer state has a pinch baseline.
  await cdp.send('Input.dispatchTouchEvent',{type:'touchMove',touchPoints:[
    {x:p1.x+10,y:p1.y+5,id:1},{x:p2.x+18,y:p2.y+5,id:2}
  ]});
  await cdp.send('Input.dispatchTouchEvent',{type:'touchMove',touchPoints:[
    {x:p1.x+22,y:p1.y+12,id:1},{x:p2.x+42,y:p2.y+12,id:2}
  ]});
  await cdp.send('Input.dispatchTouchEvent',{type:'touchEnd',touchPoints:[]});
  const after=await page.evaluate(()=>({
    desired:JSON.stringify(window.__LS3D_TEST__.cameraDesired),
    selected:window.__LS3D_TEST__.selected?.id,
    viewport:window.LS3D_THREE.adapter.viewportPerformance,
    overlay:window.LS3D_THREE.adapter.diagnostic.overlay
  }));
  expect(after.desired).not.toBe(before.desired);
  expect(after.selected).toBe(before.selected);
  expect(after.viewport.mobile).toBe(true);
  await snapshot(page,info,'step19-mobile-hall.png');
  verify(errors);await context.close();
});

test('19-H: company server-room hides outside 3D routes and legacy mode still runs',async({page},info)=>{
  const errors=await ready(page,'hall');
  await page.evaluate(()=>window.__LS3D_TEST__.applyFacilityMode('server-room'));
  await page.waitForFunction(()=>{
    const a=window.LS3D_THREE?.adapter;
    return a?.mep?.visible===false &&
      a?.mmrController?.root?.visible===false &&
      a?.colocationController?.root?.visible===false;
  });
  const state=await page.evaluate(()=>({
    facility:window.__LS3D_TEST__.facilityMode,
    mep:window.LS3D_THREE.adapter.mep,
    mmrVisible:window.LS3D_THREE.adapter.mmrController.root.visible,
    coloVisible:window.LS3D_THREE.adapter.colocationController.root.visible
  }));
  expect(state.facility).toBe('server-room');
  expect(state.mmrVisible).toBe(false);
  expect(state.coloVisible).toBe(false);
  await snapshot(page,info,'step19-company-server-room.png');
  verify(errors);
  await page.goto(BASE+(BASE.includes('?')?'&':'?')+'engine=legacy',{waitUntil:'domcontentloaded'});
  await page.waitForFunction(()=>window.LS3D_THREE?.status==='legacy');
  expect(await page.evaluate(()=>window.__LS3D_TEST__.Z.length)).toBe(7);
  expect(await page.locator('#back2D').getAttribute('href')).toContain('LS_Datacenter_Campus.html');
});


test('19-I: mobile rack controls fold/unfold and keep rack mechanisms intact',async({browser},info)=>{
  const context=await browser.newContext({
    viewport:{width:390,height:844},isMobile:true,hasTouch:true,deviceScaleFactor:2
  });
  const page=await context.newPage();
  const errors=await ready(page,'hall');
  await page.evaluate(()=>{
    const T=window.__LS3D_TEST__;T.pickAsset(T.assetMap.get('gpu-0-0'),true);
  });
  await page.waitForFunction(()=>!!document.getElementById('ls3dRackPanelToggle')&&
    !document.getElementById('ls3dRackMotionControls').hidden);
  const btn=page.locator('#ls3dRackPanelToggle');
  const body=page.locator('#ls3dRackMotionBody');
  await expect(btn).toHaveAttribute('aria-expanded','false');
  await expect(body).toBeHidden();
  await expect(page.locator('#ls3dRackMotionControls')).toHaveAttribute('data-rack-collapsed','true');
  await snapshot(page,info,'step20-mobile-rack-collapsed.png');
  await btn.click();
  await expect(btn).toHaveAttribute('aria-expanded','true');
  await expect(body).toBeVisible();
  const before=await page.evaluate(()=>{
    const T=window.__LS3D_TEST__;
    return T.mechanical(T.assetMap.get('gpu-0-0')).targetDoor;
  });
  await page.locator('[data-rack-action="front-door"]').click();
  const after=await page.evaluate(()=>{
    const T=window.__LS3D_TEST__;
    return T.mechanical(T.assetMap.get('gpu-0-0')).targetDoor;
  });
  expect(after).not.toBe(before);
  await snapshot(page,info,'step20-mobile-rack-expanded.png');
  await btn.click();
  await expect(body).toBeHidden();
  await page.waitForTimeout(200);
  const preserved=await page.evaluate(()=>{
    const T=window.__LS3D_TEST__;
    return {selected:T.selected?.id,
      targetDoor:T.mechanical(T.assetMap.get('gpu-0-0')).targetDoor,
      panel:window.LS3D_THREE.adapter.rackPanel};
  });
  expect(preserved.selected).toBe('gpu-0-0');
  expect(preserved.targetDoor).toBe(after);
  expect(preserved.panel.collapsed).toBe(true);
  verify(errors);
  await context.close();
});

test('19-J: selected GPU module extracts and reinstalls inside the isolated rack',async({page},info)=>{
  const errors=await ready(page,'hall');
  await page.evaluate(()=>window.__LS3D_TEST__.pickAsset(window.__LS3D_TEST__.assetMap.get('gpu-0-0')));
  await page.locator('#detailBtn').click();
  const action=page.locator('[data-rack-action="gpu-toggle"]');
  await expect(action).toBeVisible();
  const initial=await page.evaluate(()=>window.LS3D_THREE.adapter.rackInteraction.selectedServer);
  expect(initial.gpuCount).toBe(16);
  expect(initial.gpuTarget).toBe(0);
  await action.click();
  await page.waitForFunction(()=>window.LS3D_THREE.adapter.rackInteraction.selectedServer.gpuProgress>.95);
  const extracted=await page.evaluate(()=>({
    state:window.LS3D_THREE.adapter.rackInteraction.selectedServer,
    z:window.LS3D_THREE.adapter.scene.getObjectByName('rack-part:gpu-module-0-0:gpu-0-0')?.position.z,
    door:window.__LS3D_TEST__.mechanical(window.__LS3D_TEST__.assetMap.get('gpu-0-0')).targetDoor
  }));
  expect(extracted.state.gpuTarget).toBe(1);
  expect(extracted.z).toBeLessThan(-1);
  expect(extracted.door).toBe(1);
  await snapshot(page,info,'step19-gpu-module-extracted.png');
  await action.click();
  await page.waitForFunction(()=>window.LS3D_THREE.adapter.rackInteraction.selectedServer.gpuProgress<.05);
  expect(await page.evaluate(()=>window.LS3D_THREE.adapter.rackInteraction.selectedServer.gpuTarget)).toBe(0);
  verify(errors);
});

test('19-K: only Loading Bay faces the cargo-side camera in Gray Zone',async({page},info)=>{
  const errors=await ready(page,'gray');
  await page.evaluate(()=>window.__LS3D_TEST__.pickAsset(window.__LS3D_TEST__.assetMap.get('visitor-loading')));
  await page.waitForFunction(()=>!!window.LS3D_THREE.adapter.scene.getObjectByName('equipment:visitor-loading'));
  const world=await page.evaluate(()=>{
    const root=window.LS3D_THREE.adapter.scene;
    const bay=root.getObjectByName('equipment:visitor-loading');
    const staging=root.getObjectByName('equipment:visitor-staging');
    const carton=bay.getObjectByName('equipment-part:loading-box-0');
    const forklift=bay.getObjectByName('equipment-part:loading-forklift-chassis');
    const boxWorld=new window.LS3D_THREE.adapter.THREE.Vector3();
    carton.getWorldPosition(boxWorld);
    return {bayYaw:bay.rotation.y,stagingYaw:staging.rotation.y,
      cartonZ:boxWorld.z,bayZ:bay.position.z,forklift:!!forklift};
  });
  expect(world.bayYaw).toBe(0);
  expect(world.stagingYaw).toBeCloseTo(Math.PI);
  expect(world.cartonZ).toBeGreaterThan(world.bayZ);
  expect(world.forklift).toBe(true);
  await page.evaluate(()=>window.__LS3D_VISITOR__.start());
  await page.locator('[data-visit-stage="3"]').click();
  const tour=await page.evaluate(()=>({
    state:window.__LS3D_VISITOR__.state,
    bay:window.__LS3D_TEST__.assetMap.get('visitor-loading'),
    yaw:window.__LS3D_TEST__.cameraDesired.yaw
  }));
  expect(tour.state.index).toBe(3);
  expect(tour.state.position[2]).toBeGreaterThan(tour.bay.z+tour.bay.d/2);
  expect(Math.cos(tour.yaw)).toBeLessThan(-.9);
  await snapshot(page,info,'step19-loading-bay-front.png');
  verify(errors);
});
