// STEP 09 browser QA. Do not claim pass without actual Chromium run logs.
import {test,expect} from '@playwright/test';
const BASE=process.env.LS3D_BASE_URL||'https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html';
async function open(page) {
 await page.goto(BASE+'?engine=three',{waitUntil:'domcontentloaded'});
 await page.waitForFunction(()=>['ready','failed'].includes(window.LS3D_THREE?.status),null,{timeout:30000});
 expect(await page.evaluate(()=>window.LS3D_THREE.status)).toBe('ready');
 await page.evaluate(()=>{window.__LS3D_TEST__.applyFacilityMode('aidc');window.__LS3D_TEST__.setZone('campus')});
 await page.waitForFunction(()=>window.LS3D_THREE?.adapter?.securityController?.diagnostic()?.ready===true);
}
test('visitor makes complete 8-stage check-in; real scene has suitcase, X-ray and card reader',async({page},info)=>{
 await open(page);
 const r=await page.evaluate(()=>{
  const S=window.LS3D_THREE.adapter.securityController;
  const stages=[];for(let time=0;time<65;time+=.5)stages.push(S.simulate(time).stage);
  const nodes=[];
  S.root.traverse(o=>{if(o.name?.startsWith('security:'))nodes.push(o.name)});
  return {stageList:S.phases.map(x=>x.name),stages:[...new Set(stages)],nodes,
    diag:S.diagnostic()};
 });
 expect(r.stageList).toEqual(r.stages);
 for(const name of ['security:visitor','security:baggage','security:baggage:suitcase',
 'security:scanner-ready','security:scanner-running','security:card-reader-approved',
 'security:gate-barrier-pivot','security:access-barrier']){
  expect(r.nodes).toContain(name);
 }
 await page.screenshot({path:info.outputPath('step09-campus-entry.png')});
});
test('security barrier only opens after validation, baggage moves along conveyor',async({page})=>{
 await open(page);
 const x=await page.evaluate(()=>{
  const s=window.LS3D_THREE.adapter.securityController;
  const early=s.simulate(16),earlyAngle=s.diagnostic().barrierAngle;
  const beltA=s.simulate(22),p1=s.diagnostic().baggagePosition;
  const beltB=s.simulate(27),p2=s.diagnostic().baggagePosition;
  const late=s.simulate(54),lateAngle=s.diagnostic().barrierAngle;
  return {early,earlyAngle,beltA,beltB,p1,p2,late,lateAngle};
 });
 expect(x.earlyAngle).toBe(0);
 expect(x.p1[0]).not.toBe(x.p2[0]);
 expect(x.late.gateOpen).toBe(true);
 expect(Math.abs(x.lateAngle)).toBeGreaterThan(.1);
});
test('company server room hides visitors and outside security zone',async({page})=>{
 await open(page);
 await page.evaluate(()=>window.__LS3D_TEST__.applyFacilityMode('server-room'));
 await page.waitForFunction(()=>window.LS3D_THREE?.diagnostics?.zones===1);
 const r=await page.evaluate(()=>window.LS3D_THREE.adapter.securityController.diagnostic());
 expect(r.visible).toBe(false);
 expect(r.ready).toBe(false);
});
