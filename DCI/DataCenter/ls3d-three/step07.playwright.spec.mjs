// STEP 07 browser regression (real Chromium/WebGL2 required). Saved script != executed PASS.
// Run: npx playwright test DCI/DataCenter/ls3d-three/step07.playwright.spec.mjs
import {test,expect} from '@playwright/test';
const BASE=process.env.LS3D_BASE_URL ||
  'https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html';
test.use({viewport:{width:1366,height:768}});
async function load(page) {
  await page.goto(BASE+'?engine=three',{waitUntil:'domcontentloaded'});
  await page.waitForFunction(()=>['ready','failed'].includes(window.LS3D_THREE?.status),null,{timeout:30000});
  const s=await page.evaluate(()=>({state:window.LS3D_THREE.status,error:window.LS3D_THREE.error}));
  expect(s.state,s.error||'Three.js not available').toBe('ready');
  await page.waitForFunction(()=>window.LS3D_THREE?.renders>3);
}
test('A* path planner uses real existing 74-asset model and no blocked segments',async({page},info)=>{
  const errors=[];page.on('pageerror',e=>errors.push(e.message));
  await load(page);
  const result=await page.evaluate(()=>{
    const T=window.__LS3D_TEST__,S=window.LS3D_THREE;
    const nav=S.adapter.navigation,staff=T.npcs.filter(n=>n.route);
    return staff.map(n=>{
      const plan=nav.getPlan(n);
      return {id:n.id,legs:plan.legs.length,failed:plan.fails,
        total:plan.total,correctSegments:plan.legs.every(l=>l.blocked||
          l.points.every(p=>!plan.grid.pointBlocked(p))&&
          l.points.slice(1).every((p,i)=>plan.grid.lineClear(l.points[i],p)))};
    });
  });
  expect(result.length).toBeGreaterThanOrEqual(4);
  expect(result.every(x=>x.failed===0&&x.correctSegments&&x.legs>=4)).toBe(true);
  expect(errors).toEqual([]);
  await page.screenshot({path:info.outputPath('step07-campus-navigation.png')});
});
test('Hall worker navigates through the real door portal and avoids racks',async({page},info)=>{
  await load(page);
  await page.evaluate(()=>window.__LS3D_TEST__.setZone('hall'));
  const state=await page.evaluate(()=>{
    const T=window.__LS3D_TEST__,S=window.LS3D_THREE,n=T.npcs.find(n=>n.id==='eng-hall');
    const plan=S.adapter.navigation.getPlan(n);
    const z=T.Z.find(z=>z.id==='hall'),front=z.pos[1]+z.size[1]/2;
    return {failed:plan.fails,door:front,outerWaypoint:plan.route.some(p=>p[1]>front+1),
      innerWaypoint:plan.route.some(p=>p[1]<front-1),
      detours:plan.legs.filter(l=>l.points.length>2).length,
      safe:plan.legs.every(l=>l.blocked||
        l.points.slice(1).every((p,i)=>plan.grid.lineClear(l.points[i],p)))};
  });
  expect(state.failed).toBe(0);
  expect(state.outerWaypoint&&state.innerWaypoint).toBe(true);
  expect(state.detours).toBeGreaterThan(0);
  expect(state.safe).toBe(true);
  await page.screenshot({path:info.outputPath('step07-hall-door.png')});
});
test('worker-follow camera sees same navigated pose as 3D avatar',async({page})=>{
  await load(page);
  await page.evaluate(()=>window.__LS3D_TEST__.setZone('hall'));
  await page.waitForFunction(()=>window.LS3D_THREE?.diagnostics?.npc===1);
  const r=await page.evaluate(()=>{
    const T=window.__LS3D_TEST__,S=window.LS3D_THREE,n=T.npcs.find(n=>n.id==='eng-hall');
    const p=T.npcPose(n);
    let avatar=null;S.adapter.scene.traverse(o=>{
      if(o.userData?.npcId===n.id)avatar=o;
    });
    return {x:p.x,z:p.z,avatarX:avatar?.position.x,avatarZ:avatar?.position.z};
  });
  expect(r.avatarX).toBeCloseTo(r.x,1);
  expect(r.avatarZ).toBeCloseTo(r.z,1);
});
test('server-room routes stay within its single building with no outsider model',async({page})=>{
  await load(page);
  await page.evaluate(()=>window.__LS3D_TEST__.applyFacilityMode('server-room'));
  await page.waitForFunction(()=>window.LS3D_THREE?.diagnostics?.zones===1);
  const r=await page.evaluate(()=>{
    const T=window.__LS3D_TEST__,nav=window.LS3D_THREE.adapter.navigation;
    const n=T.npcs.find(p=>p.id==='eng-hall'),plan=nav.getPlan(n);
    const hall=T.Z.find(z=>z.id==='hall');
    return {mode:T.facilityMode,failed:plan.fails,
      points:plan.legs.flatMap(l=>l.points),
      cx:hall.pos[0],cz:hall.pos[1],external:
        window.LS3D_THREE.adapter.scene.children[0].children[0].children.map(g=>g.name)};
  });
  expect(r.mode).toBe('server-room');
  expect(r.failed).toBe(0);
  expect(r.external).toEqual(['arch:zone:hall']);
  expect(r.points.every(p=>Math.abs(p[0]-r.cx)<13&&Math.abs(p[1]-r.cz)<12)).toBe(true);
});
test('On-Premises route replan, mobile scene and original legacy path',async({page},info)=>{
  await load(page);
  await page.evaluate(()=>window.__LS3D_TEST__.applyFacilityMode('onprem'));
  const r=await page.evaluate(()=>{
    const T=window.__LS3D_TEST__,nav=window.LS3D_THREE.adapter.navigation;
    return T.npcs.filter(n=>n.id==='eng-hall'||n.id==='eng-ops')
      .map(n=>({id:n.id,failed:nav.getPlan(n).fails}));
  });
  expect(r.every(p=>p.failed===0)).toBe(true);
  await page.setViewportSize({width:390,height:844});
  await expect(page.locator('#sceneThree')).toBeVisible();
  await page.screenshot({path:info.outputPath('step07-onprem-mobile.png')});
  await page.goto(BASE+'?engine=legacy',{waitUntil:'domcontentloaded'});
  await page.waitForFunction(()=>window.__LS3D_TEST__?.Z?.length===7);
  expect(await page.evaluate(()=>window.LS3D_THREE?.status)).toBe('legacy');
});
