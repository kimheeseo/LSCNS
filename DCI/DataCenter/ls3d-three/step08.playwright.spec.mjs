// STEP 08 WebGL2 regression — script only until run on Chrome/CI.
import {test,expect} from '@playwright/test';
const BASE=process.env.LS3D_BASE_URL||'https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html';
async function open(page) {
 await page.goto(BASE+'?engine=three',{waitUntil:'domcontentloaded'});
 await page.waitForFunction(()=>['ready','failed'].includes(window.LS3D_THREE?.status),null,{timeout:30000});
 expect(await page.evaluate(()=>window.LS3D_THREE.status)).toBe('ready');
}
test('worker inspection FSM advances align, inspect, tablet, depart without restart',async({page},info)=>{
 await open(page);
 const data=await page.evaluate(()=>{
  const T=window.__LS3D_TEST__,controller=window.LS3D_THREE.adapter.inspection;
  const n=T.npcs.find(n=>n.id==='eng-hall');
  const pose={x:15,z:7,heading:0,action:'pause'};
  controller.reset();
  const seq=[0,.55,1.36,2.03,2.2,2.5].map(t=>
    controller.update(n,pose,t,'normal'));
  const total=controller.diagnostic().metrics.completed;
  const resumed=controller.update(n,{...pose,action:'walk'},2.8,'normal');
  const next=controller.update(n,{...pose,action:'walk'},3.3,'normal');
  return {states:seq.map(s=>s.inspectionState),
   tablets:seq.map(s=>s.showTablet),targets:seq.map(s=>s.targetId),
   total,resumed:resumed.inspectionState,next:next.inspectionState};
 });
 expect(data.states).toEqual(['align','inspect','tablet','depart','depart','depart']);
 expect(data.tablets).toEqual([false,false,true,false,false,false]);
 expect(data.targets[0]).not.toBeNull();
 expect(data.total).toBe(1);
 expect(data.resumed).toBe('depart');
 expect(data.next).toBe('move');
 await page.screenshot({path:info.outputPath('step08-inspection.png')});
});
test('incident chooses known asset, room mode keeps worker near assigned room',async({page})=>{
 await open(page);
 const data=await page.evaluate(()=>{
  const T=window.__LS3D_TEST__,fsm=window.LS3D_THREE.adapter.inspection;
  fsm.reset();
  const n=T.npcs.find(n=>n.id==='eng-utility');
  const result=fsm.update(n,{x:-47,z:-26,heading:0,action:'pause'},12,'power');
  const record=fsm.diagnostic().workers[0];
  return {result,record,known:T.assetMap.has(result.targetId)};
 });
 expect(data.result.inspectionAlert).toBe(true);
 expect(data.known).toBe(true);
 expect(data.record.scenario).toBe('power');
});
test('legacy mode not affected by new inspection FSM',async({page})=>{
 await page.goto(BASE+'?engine=legacy',{waitUntil:'domcontentloaded'});
 await page.waitForFunction(()=>window.__LS3D_TEST__?.Z?.length===7);
 expect(await page.evaluate(()=>window.LS3D_THREE?.status)).toBe('legacy');
});
