// STEP 14 browser acceptance — requires Playwright/Chromium; do not mistake file creation for execution.
// LS3D_BASE_URL=http://localhost:8000/DCI/DataCenter/LS_Datacenter_3D.html
// npx playwright test DCI/DataCenter/ls3d-three/step14.playwright.spec.mjs
import {test,expect} from '@playwright/test';
const BASE=process.env.LS3D_BASE_URL||
  'https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html';
async function boot(page){
  const errors=[];
  page.on('pageerror',e=>errors.push(e.message));
  page.on('console',msg=>{if(msg.type()==='error')errors.push(msg.text());});
  await page.goto(BASE+(BASE.includes('?')?'&':'?')+'engine=three',{waitUntil:'domcontentloaded'});
  await page.waitForFunction(()=>['ready','failed'].includes(window.LS3D_THREE?.status),null,{timeout:30000});
  const state=await page.evaluate(()=>({status:window.LS3D_THREE?.status,error:window.LS3D_THREE?.error}));
  expect(state.status,state.error||'Three.js failed').toBe('ready');
  await page.evaluate(()=>window.__LS3D_TEST__.setZone('hall'));
  await page.waitForFunction(()=>window.LS3D_THREE?.adapter?.rackComponents?.rackCount>0);
  const id=await page.evaluate(()=>{
    const T=window.__LS3D_TEST__;
    const rack=T.visibleAssets().find(x=>x.zone==='hall'&&T.isMechanicalRack(x));
    if(!rack)throw Error('No mechanical rack');
    T.pickAsset(rack);return rack.id;
  });
  await page.waitForFunction(()=>!document.getElementById('ls3dRackMotionControls')?.hidden);
  return {id,errors};
}
test('STEP14 42U/48U ToR placement, selectable eight ports, independent PDU A/B',async({page},info)=>{
  const {id,errors}=await boot(page);
  const d=await page.evaluate(id=>{
    const a=window.LS3D_THREE.adapter,parts=[];
    a.scene.traverse(o=>{
      if(o.name==='equipment:'+id)o.traverse(c=>{
        if(c.userData?.rackPart)parts.push(c.userData.rackPart);
      });
    });
    const tor=parts.find(p=>p.kind==='tor');
    const panel=parts.find(p=>p.kind==='patch-panel');
    const portIds=[];
    a.scene.traverse(o=>{
      if(o.name==='equipment:'+id)o.traverse(p=>{
        if(p.name.startsWith('rack-link:tor-port-'))portIds.push(p.userData.partId);
      });
    });
    return {tor,panel,portIds,diag:a.rackCabling.diagnostic(id)};
  },id);
  expect(d.tor).toBeTruthy();expect(d.panel).toBeTruthy();
  expect(d.tor.uStart).toBeGreaterThan(d.panel.uStart);
  expect(new Set(d.portIds).size).toBe(8);
  expect(d.diag.links.some(l=>l.kind==='fiber')).toBe(true);
  expect(d.diag.links.some(l=>l.kind==='power-A')).toBe(true);
  expect(d.diag.links.some(l=>l.kind==='power-B')).toBe(true);
  const chosen=await page.evaluate(({id,pid})=>window.LS3D_THREE.adapter.selectRackPart(id,pid),{id,pid:d.portIds[0]});
  expect(chosen.kind).toBe('tor-port');
  await page.screenshot({path:info.outputPath('step14-tor-ports-3d.png')});
  expect(errors).toEqual([]);
});
test('STEP14 front/rear power and optical filter dims unselected physical routes',async({page},info)=>{
  const {id,errors}=await boot(page);
  await page.locator('[data-rack-action="rear-camera"]').click();
  await page.locator('[data-rack-action="link-power-a"]').click();
  const read=()=>page.evaluate(id=>{
    const a=window.LS3D_THREE.adapter;
    const info=a.rackCabling.diagnostic(id),powerA=[],powerB=[],fiber=[];
    a.scene.traverse(o=>{
      if(o.name==='equipment:'+id)o.traverse(m=>{
        if(m.name.startsWith('rack-link:cable:')){
          const group=m.parent;
          const dst=group.userData.cableSystem==='power-A'?powerA:
            group.userData.cableSystem==='power-B'?powerB:fiber;
          dst.push(m.material.opacity);
        }
      });
    });
    return {info,powerA,powerB,fiber,side:a.rackInteraction.focus?.side};
  },id);
  const d=await read();
  expect(d.side).toBe('rear');
  expect(d.info.filter).toBe('power-A');
  expect(d.info.links.every(x=>x.visible)).toBe(true);
  expect(d.powerA.every(x=>x>.8)).toBe(true);
  expect(d.powerB.every(x=>x<.2)).toBe(true);
  expect(d.fiber.every(x=>x<.2)).toBe(true);
  await page.screenshot({path:info.outputPath('step14-rear-pdu-a.png')});
  await page.locator('[data-rack-action="link-power-b"]').click();
  const e=await read();
  expect(e.powerB.every(x=>x>.8)).toBe(true);
  expect(e.powerA.every(x=>x<.2)).toBe(true);
  await page.locator('[data-rack-action="link-all"]').click();
  expect((await read()).info.links.every(x=>x.emphasized)).toBe(true);
  expect(errors).toEqual([]);
});
test('STEP14 attached NIC cable follows selected server rail travel',async({page},info)=>{
  const {id,errors}=await boot(page);
  const servers=await page.evaluate(id=>{
    const x=[];
    window.LS3D_THREE.adapter.scene.traverse(o=>{
      if(o.name==='equipment:'+id)o.traverse(p=>{
        if(p.userData?.rackPart?.kind?.startsWith('server-'))x.push(p.userData.rackPart.id);
      });
    });
    return x;
  },id);
  expect(servers.length).toBeGreaterThan(1);
  const chosen=servers[1];
  const before=await page.evaluate(({id,chosen})=>{
    const a=window.LS3D_THREE.adapter;
    function cable(){
      let z=null;
      a.scene.traverse(o=>{if(o.name==='equipment:'+id)o.traverse(p=>{
        if(p.userData?.cableSystem==='fiber'&&p.userData.from===chosen)
          z=p.children[0]?.geometry?.parameters?.path?.points?.[0]?.z??null;
      });});
      return z;
    }
    const origin=cable();
    a.setRackServerSelection(id,chosen);
    a.toggleRackServerSlide(id);
    return {origin};
  },{id,chosen});
  expect(before.origin).not.toBeNull();
  await page.waitForFunction(({id,chosen,origin})=>{
    const a=window.LS3D_THREE.adapter;
    let cableZ=null,serverZ=null;
    a.scene.traverse(o=>{if(o.name==='equipment:'+id)o.traverse(p=>{
      if(p.userData?.rackPart?.id===chosen)serverZ=p.position.z;
      if(p.userData?.cableSystem==='fiber'&&p.userData.from===chosen)
        cableZ=p.children[0]?.geometry?.parameters?.path?.points?.[0]?.z??null;
    });});
    return serverZ<-.3&&cableZ!==null&&Math.abs((cableZ-origin)-serverZ)<.04;
  },{id,chosen,origin:before.origin});
  await page.screenshot({path:info.outputPath('step14-nic-service-loop-pulled.png')});
  await page.locator('[data-rack-action="server-slide"]').click();
  expect(errors).toEqual([]);
});
test('STEP14 mobile rack controls remain reachable',async({page},info)=>{
  await page.setViewportSize({width:390,height:844});
  const {errors}=await boot(page);
  await page.locator('details:has([data-rack-action="link-fiber"]) summary').click();
  await expect(page.locator('[data-rack-action="link-fiber"]')).toBeVisible();
  await page.locator('[data-rack-action="link-fiber"]').click();
  await page.screenshot({path:info.outputPath('step14-mobile-cabling-controls.png')});
  expect(errors).toEqual([]);
});
