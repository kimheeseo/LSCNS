// STEP 11-13 — real browser acceptance tests (not executed in this authoring session).
// Set LS3D_BASE_URL to a trusted local server URL to test unpublished changes.
// npx playwright test DCI/DataCenter/ls3d-three/step11to13.playwright.spec.mjs
import {test,expect} from '@playwright/test';
const BASE=process.env.LS3D_BASE_URL ||
  'https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html';
test.use({viewport:{width:1366,height:768}});
async function start(page){
  await page.goto(BASE+'?engine=three',{waitUntil:'domcontentloaded'});
  await page.waitForFunction(()=>['ready','failed'].includes(window.LS3D_THREE?.status),null,{timeout:30000});
  const r=await page.evaluate(()=>({state:window.LS3D_THREE?.status,error:window.LS3D_THREE?.error}));
  expect(r.state,r.error||'Three bootstrap failed').toBe('ready');
  await page.waitForFunction(()=>window.LS3D_THREE?.renders>4);
  await page.evaluate(()=>window.__LS3D_TEST__.setZone('hall'));
  await page.waitForFunction(()=>window.LS3D_THREE?.adapter?.rackComponents?.rackCount>0);
}
async function chooseRack(page){
  return page.evaluate(()=>{
    const T=window.__LS3D_TEST__;
    const a=T.visibleAssets().find(x=>x.zone==='hall'&&T.isMechanicalRack(x));
    if(!a)throw Error('No interactive rack visible');
    T.pickAsset(a);
    return {id:a.id,type:a.type};
  });
}
test('STEP11 front/rear rack camera targets independent sides, doors hinge separately',async({page},info)=>{
  const errors=[];page.on('pageerror',e=>errors.push(e.message));
  await start(page);
  const rack=await chooseRack(page);
  await page.waitForFunction(()=>!!document.getElementById('ls3dRackMotionControls') && !document.getElementById('ls3dRackMotionControls').hidden);
  await page.locator('[data-rack-action="rear-camera"]').click();
  const rear=await page.evaluate(()=>({
    yaw:window.__LS3D_TEST__.cameraDesired.yaw,
    side:window.LS3D_THREE.adapter.rackInteraction.focus?.side
  }));
  expect(rear.side).toBe('rear');expect(rear.yaw).toBeCloseTo(0,3);
  await page.locator('[data-rack-action="rear-door"]').click();
  await page.waitForFunction(id=>{
    let angle=null;
    window.LS3D_THREE.adapter.scene.traverse(o=>{
      if(o.name==='equipment:'+id)o.traverse(c=>{if(c.name==='rear-door-hinge')angle=c.rotation.y;});
    });
    return angle>.3;
  },rack.id);
  const independent=await page.evaluate(id=>{
    let front=null,rear=null;
    window.LS3D_THREE.adapter.scene.traverse(o=>{if(o.name==='equipment:'+id)o.traverse(c=>{
      if(c.name==='front-door-hinge')front=c.rotation.y;
      if(c.name==='rear-door-hinge')rear=c.rotation.y;
    });});
    return {front,rear};
  },rack.id);
  expect(independent.rear).toBeGreaterThan(.3);
  expect(independent.front).toBeCloseTo(0,2);
  await page.locator('[data-rack-action="front-camera"]').click();
  const yaw=await page.evaluate(()=>window.__LS3D_TEST__.cameraDesired.yaw);
  expect(yaw).toBeCloseTo(Math.PI,3);
  await page.locator('[data-rack-action="front-door"]').click();
  await page.waitForFunction(id=>{
    let angle=null;
    window.LS3D_THREE.adapter.scene.traverse(o=>{
      if(o.name==='equipment:'+id)o.traverse(c=>{if(c.name==='front-door-hinge')angle=c.rotation.y;});
    });
    return angle<-.3;
  },rack.id);
  await page.screenshot({path:info.outputPath('step11-rear-front-camera-doors.png')});
  expect(errors).toEqual([]);
});
test('STEP12 only selected server sled slides; distinct NIC/port geometry exists',async({page},info)=>{
  await start(page);const rack=await chooseRack(page);
  const servers=await page.evaluate(id=>{
    const a=window.LS3D_THREE.adapter,partIds=[];
    a.scene.traverse(o=>{
      if(o.name==='equipment:'+id)o.traverse(p=>{
        if(p.userData?.rackPart?.kind?.startsWith('server-'))partIds.push(p.userData.rackPart.id);
      });
    });
    return partIds;
  },rack.id);
  expect(servers.length).toBeGreaterThanOrEqual(2);
  const result=await page.evaluate(({id,first,second})=>{
    const a=window.LS3D_THREE.adapter;
    const selected=a.setRackServerSelection(id,second);
    a.toggleRackServerSlide(id);
    return {selected,first,second};
  },{id:rack.id,first:servers[0],second:servers[1]});
  expect(result.selected).toBe(servers[1]);
  await page.waitForFunction(({id,second})=>{
    let z=0;
    window.LS3D_THREE.adapter.scene.traverse(o=>{
      if(o.name==='equipment:'+id)o.traverse(p=>{
        if(p.userData?.rackPart?.id===second)z=p.position.z;
      });
    });
    return z<-.3;
  },{id:rack.id,second:servers[1]});
  const data=await page.evaluate(({id,first,second})=>{
    const pos={};let nic=0,port=0;
    window.LS3D_THREE.adapter.scene.traverse(o=>{
      if(o.name==='equipment:'+id)o.traverse(p=>{
        if(p.userData?.rackPart?.id===first)pos.first=p.position.z;
        if(p.userData?.rackPart?.id===second)pos.second=p.position.z;
        if(p.name?.startsWith('rack-server:nic-board'))nic++;
        if(p.name?.startsWith('rack-server:nic-cage'))port++;
      });
    });
    return {pos,nic,port,session:window.LS3D_THREE.adapter.rackInteraction.selectedServer};
  },{id:rack.id,first:servers[0],second:servers[1]});
  expect(data.pos.first).toBeCloseTo(0,3);
  expect(data.pos.second).toBeLessThan(-.3);
  expect(data.nic).toBeGreaterThan(0);expect(data.port).toBeGreaterThan(0);
  expect(data.session.selected).toBe(servers[1]);
  await page.screenshot({path:info.outputPath('step12-selected-server-nic.png')});
  await page.locator('[data-rack-action="server-slide"]').click();
  await page.waitForFunction(({id,second})=>{
    let z=null;window.LS3D_THREE.adapter.scene.traverse(o=>{
      if(o.name==='equipment:'+id)o.traverse(p=>{if(p.userData?.rackPart?.id===second)z=p.position.z;});
    });return z!==null&&Math.abs(z)<.05;
  },{id:rack.id,second:servers[1]});
});
test('STEP13 optical cassette slides, exposes 24 pickable LC duplex ports, lid hinges and 3D fibers',async({page},info)=>{
  await start(page);const rack=await chooseRack(page);
  await page.locator('[data-rack-action="cassette-slide"]').click();
  await page.waitForFunction(id=>{
    let slide=null,lid=null;
    window.LS3D_THREE.adapter.scene.traverse(o=>{
      if(o.name==='equipment:'+id)o.traverse(p=>{
        if(p.name==='fiber-cassette')slide=p.position.z;
        if(p.name?.startsWith('rack-optical:module-cover-'))lid=p.rotation.x;
      });
    });
    return slide<-.2&&lid<-.4;
  },rack.id);
  const data=await page.evaluate(id=>{
    const counts={adapters:0,connectors:0,fibers:0},portIds=[];
    window.LS3D_THREE.adapter.scene.traverse(o=>{
      if(o.name==='equipment:'+id)o.traverse(p=>{
        if(p.name?.startsWith('rack-optical:adapter-port-'))counts.adapters++;
        if(p.name?.startsWith('rack-optical:connector-boot-'))counts.connectors++;
        if(p.name?.startsWith('rack-optical:fanout-fiber-'))counts.fibers++;
        if(p.name?.startsWith('rack-optical:adapter-port-'))portIds.push(p.userData.partId);
      });
    });
    const selected=window.LS3D_THREE.adapter.selectRackPart(id,portIds[0]);
    return {counts,portIds,selected};
  },rack.id);
  expect(data.counts).toEqual({adapters:24,connectors:8,fibers:48});
  expect(new Set(data.portIds).size).toBe(24);
  expect(data.selected.kind).toBe('fiber-port');
  await expect(page.locator('#ls3dRackPartInfo')).toContainText('LC duplex');
  await page.locator('[data-rack-action="optical-explode"]').click();
  expect(await page.evaluate(id=>window.LS3D_THREE.adapter.rackInteraction.selectedOptics?.exploded,rack.id)).toBe(true);
  await page.screenshot({path:info.outputPath('step13-optical-cassette-fibers.png')});
  await page.locator('[data-rack-action="cassette-slide"]').click();
  await page.waitForFunction(id=>{
    let slide=null;window.LS3D_THREE.adapter.scene.traverse(o=>{
      if(o.name==='equipment:'+id)o.traverse(p=>{if(p.name==='fiber-cassette')slide=p.position.z;});
    });return slide!==null&&Math.abs(slide)<.05;
  },rack.id);
});
test('STEP11-13 company server-room, mobile/legacy and original asset/BOM path remain',async({page},info)=>{
  await page.setViewportSize({width:390,height:844});await start(page);
  await page.evaluate(()=>window.__LS3D_TEST__.applyFacilityMode('server-room'));
  await page.waitForFunction(()=>window.LS3D_THREE?.diagnostics?.zones===1);
  const r=await page.evaluate(()=>({
    racks:window.LS3D_THREE.adapter.rackComponents.rackCount,
    groupNames:window.LS3D_THREE.adapter.scene.children[0].children[0].children.map(g=>g.name)
  }));
  expect(r.racks).toBeGreaterThan(0);expect(r.groupNames).toEqual(['arch:zone:hall']);
  await page.screenshot({path:info.outputPath('step13-company-server-room-mobile.png')});
  await page.goto(BASE+'?engine=legacy',{waitUntil:'domcontentloaded'});
  await page.waitForFunction(()=>window.__LS3D_TEST__?.Z?.length===7);
  expect(await page.evaluate(()=>window.LS3D_THREE?.status)).toBe('legacy');
});
