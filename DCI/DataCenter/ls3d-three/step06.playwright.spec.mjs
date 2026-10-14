// STEP 06 — browser QA. Prepared for Chromium; not yet executed on an independent GPU browser.
// Run: npx playwright test DCI/DataCenter/ls3d-three/step06.playwright.spec.mjs
// Optional LS3D_BASE_URL=http://127.0.0.1:8765/DCI/DataCenter/LS_Datacenter_3D.html
import {test,expect} from '@playwright/test';
const BASE=process.env.LS3D_BASE_URL||
 'https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html';
test.use({viewport:{width:1366,height:768}});
async function start(page){
  await page.goto(BASE+'?engine=three',{waitUntil:'domcontentloaded'});
  await page.waitForFunction(()=>['ready','failed'].includes(window.LS3D_THREE?.status),null,{timeout:30000});
  const initial=await page.evaluate(()=>({state:window.LS3D_THREE?.status,error:window.LS3D_THREE?.error}));
  expect(initial.state,initial.error||'Three engine not ready').toBe('ready');
  await page.evaluate(()=>{
    const T=window.__LS3D_TEST__;
    T.applyFacilityMode('aidc');
    T.setZone('campus');
  });
  await page.waitForFunction(()=>window.LS3D_THREE?.diagnostics?.npc===6);
}
test('campus renders four field workers and two security guards with 3D safety gear',async({page},info)=>{
  const errors=[];page.on('pageerror',e=>errors.push(e.message));
  await start(page);
  const r=await page.evaluate(()=>{
    const S=window.LS3D_THREE,npcs=[],parts={};
    S.adapter.scene.traverse(obj=>{
      if(obj.userData?.npcId)npcs.push({id:obj.userData.npcId,role:obj.userData.npcRole,rig:!!obj.userData.rig});
      if(obj.name?.startsWith('npc-part:'))parts[obj.name]=(parts[obj.name]||0)+1;
    });
    return {npcs,parts,summary:S.adapter.npcAnimation};
  });
  expect(r.npcs.length).toBe(6);
  expect(r.summary.roles.guard).toBe(2);
  expect(r.summary.roles.engineer).toBe(2);
  expect(r.summary.roles.operator).toBe(1);
  expect(r.summary.roles.technician).toBe(1);
  expect(r.npcs.every(n=>n.rig)).toBe(true);
  for(const name of ['badge-card','badge-photo','radio-case','safety-helmet-shell','security-cap-crown','upper-trouser','safety-boot']){
    expect(r.parts['npc-part:'+name],name).toBeGreaterThan(0);
  }
  expect(errors).toEqual([]);
  await page.screenshot({path:info.outputPath('step06-campus-staff.png')});
});
test('inspect/walk/idle poses manipulate real articulation hierarchy',async({page},info)=>{
  await start(page);
  await page.evaluate(()=>{
    const T=window.__LS3D_TEST__;
    window.__step06OldPose=T.npcPose;
    T.npcPose=n=>({x:n.id==='eng-hall'?11:n.id==='eng-campus'?-14:n.id==='guard-a'?30:0,
      z:n.id==='eng-hall'?8:n.id==='eng-campus'?34:45,
      heading:.33,action:n.id==='eng-hall'?'inspect':n.id==='eng-campus'?'walk':'pause'});
  });
  try{
    await page.waitForFunction(()=>{
      const S=window.LS3D_THREE;
      return S?.adapter?.npcAnimation?.actions?.inspect>=1 && S?.adapter?.npcAnimation?.actions?.walk>=1;
    });
    const check=await page.evaluate(()=>{
      let walking,inspecting,idle;
      window.LS3D_THREE.adapter.scene.traverse(o=>{
        if(o.userData?.npcId==='eng-campus')walking=o.userData.rig;
        if(o.userData?.npcId==='eng-hall')inspecting=o.userData.rig;
        if(o.userData?.npcId==='guard-a')idle=o.userData.rig;
      });
      return {
        walkingJoint:walking?.legs?.left?.upper?.name,
        walkingAngle:Math.abs(walking?.legs?.left?.upper?.rotation.x||0),
        inspectTablet:inspecting?.tablet?.visible,
        inspectLeftArm:inspecting?.arms?.left?.upper?.rotation.x,
        inspectLeftElbow:inspecting?.arms?.left?.lower?.rotation.x,
        guardTablet:idle?.tablet?.visible,
        actions:window.LS3D_THREE.adapter.npcAnimation.actions
      };
    });
    expect(check.walkingJoint).toBe('npc-joint:left-hip');
    expect(check.inspectTablet).toBe(true);
    expect(check.inspectLeftArm).toBeLessThan(-.75);
    expect(check.inspectLeftElbow).toBeLessThan(-.75);
    expect(check.guardTablet).toBe(false);
    expect(check.actions.inspect).toBeGreaterThanOrEqual(1);
    await page.screenshot({path:info.outputPath('step06-pose-samples.png')});
  }finally{
    await page.evaluate(()=>{
      if(window.__step06OldPose){
        window.__LS3D_TEST__.npcPose=window.__step06OldPose;
        delete window.__step06OldPose;
      }
    });
  }
});
test('zone and facility filter personnel without spawning outside security guards',async({page},info)=>{
  await start(page);
  await page.evaluate(()=>window.__LS3D_TEST__.setZone('hall'));
  await page.waitForFunction(()=>window.LS3D_THREE?.diagnostics?.npc===1);
  expect(await page.evaluate(()=>window.LS3D_THREE.adapter.npcAnimation.ids)).toEqual(['eng-hall']);
  await page.evaluate(()=>window.__LS3D_TEST__.applyFacilityMode('server-room'));
  await page.waitForFunction(()=>window.LS3D_THREE?.diagnostics?.zones===1);
  const filtered=await page.evaluate(()=>window.LS3D_THREE.adapter.npcAnimation);
  expect(filtered.ids).toEqual(['eng-hall']);
  expect(filtered.roles.guard||0).toBe(0);
  await page.screenshot({path:info.outputPath('step06-server-room-worker.png')});
  await page.evaluate(()=>window.__LS3D_TEST__.applyFacilityMode('aidc'));
  await page.evaluate(()=>window.__LS3D_TEST__.setZone('campus'));
  await page.waitForFunction(()=>window.LS3D_THREE?.diagnostics?.npc===6);
});
test('Three.js mobile viewport retains NPCs and works with fallback legacy route',async({page},info)=>{
  await page.setViewportSize({width:390,height:844});
  await start(page);
  await expect(page.locator('#sceneThree')).toBeVisible();
  expect(await page.evaluate(()=>window.LS3D_THREE?.diagnostics?.npc)).toBe(6);
  await page.screenshot({path:info.outputPath('step06-mobile.png')});
  await page.goto(BASE+'?engine=legacy',{waitUntil:'domcontentloaded'});
  await page.waitForFunction(()=>window.__LS3D_TEST__?.Z?.length===7);
  expect(await page.evaluate(()=>window.LS3D_THREE?.status)).toBe('legacy');
  await expect(page.locator('#scene')).toBeVisible();
});
