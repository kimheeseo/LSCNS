// STEP 09 — visibly animated, conceptual campus security check-in sequence.
// Does NOT perform real identity verification, physical scanner analysis or access authorization.
// Uses existing gate, reception, X-ray conveyor, metal detector, security personnel positions.
export function createVisitorSecurity(THREE,T,{cube,characterFactory}) {
 const root=new THREE.Group();root.name='security:checkpoint';
 const phases=[
  {name:'reception-approach',duration:8,action:'walk'},
  {name:'visitor-registration',duration:6,action:'pause'},
  {name:'xray-queue',duration:8,action:'walk'},
  {name:'baggage-conveyor',duration:7,action:'pause'},
  {name:'metal-screening',duration:8,action:'walk'},
  {name:'access-card-validation',duration:8,action:'pause'},
  {name:'gate-passage',duration:8,action:'walk'},
  {name:'exit',duration:7,action:'walk'}
 ];
 const cycle=phases.reduce((s,p)=>s+p.duration,0);
 const visitor=characterFactory.create({id:'visitor-demo',role:'visitor',speed:2.1});
 visitor.name='security:visitor';root.add(visitor);
 const luggage=new THREE.Group();luggage.name='security:baggage';root.add(luggage);
 const caseBox=cube(luggage,0,.14,0,.72,.47,.55,'#aa7f56',.22,.64);
 caseBox.name='security:baggage:suitcase';
 cube(luggage,0,.62,-.14,.38,.17,.09,'#d8be8c',.53,.4).name='security:baggage:handle';
 const lamp=new THREE.Group();lamp.name='security:scanner-lamp';root.add(lamp);
 const lampMesh=cube(lamp,0,2.8,0,.40,.15,.20,'#7cecc0',.14,.31);
 lampMesh.name='security:scanner-ready';
 const lampWarning=cube(lamp,0,2.8,0,.40,.15,.20,'#f6c66d',.14,.31);
 lampWarning.name='security:scanner-running';lampWarning.visible=false;
 const gate=new THREE.Group();gate.name='security:gate-barrier-pivot';root.add(gate);
 const bar=cube(gate,2.2,0,0,4.4,.16,.19,'#e8c982',.57,.36);
 bar.name='security:access-barrier';
 const scanner=cube(root,0,1.15,0,.2,.34,.19,'#53d1bd',.33,.42);
 scanner.name='security:card-reader-approved';
 const readerWaiting=cube(root,0,1.15,0,.2,.34,.19,'#e8b26d',.33,.42);
 readerWaiting.name='security:card-reader-waiting';
 const zones={},stats={cycles:0,updates:0,stage:'inactive',ready:false,stageCount:phases.length};
 let lastCycle=-1,mode='',lastScene='',visible=false;
 function available() {
   const aids=new Map((T.visibleAssets?.()||[]).map(a=>[a.id,a]));
   return {desk:aids.get('entry-desk'),xray:aids.get('xray'),
     metal:aids.get('metal-detector'),gate:aids.get('gate')};
 }
 function lerp(a,b,t){return [a[0]+(b[0]-a[0])*t,a[1]+(b[1]-a[1])*t];}
 function point(a,dx=0,dz=0){return [a.x+dx,a.z+dz];}
 function simulate(clock=T.simTime||0) {
   const components=available();
   visible=T.facilityMode==='aidc' && T.currentZone==='campus' &&
     Object.values(components).every(Boolean);
   root.visible=visible;stats.ready=visible;
   if(!visible){stats.stage='inactive';return {active:false,stage:'inactive'};}
   const {desk,xray,metal,gate:entry}=components;
   const way=[
    point(desk,-8,desk.d*.5+1.3),point(desk,0,desk.d*.5+1.3),
    point(xray,-5,xray.d*.5+1.0),point(xray,0,xray.d*.5+1.0),
    point(metal,-3,0),
    point(metal,0,0),point(entry,-3.5,0),
    point(entry,0,0),point(entry,8,1.7)
   ];
   const time=Math.max(0,Number(clock)||0),iteration=Math.floor(time/cycle),t=time%cycle;
   if(lastCycle!==iteration){lastCycle=iteration;stats.cycles++;}
   let step=phases.length-1,start=0;
   for(let i=0;i<phases.length;i++){
     if(t<start+phases[i].duration){step=i;break;}start+=phases[i].duration;
   }
   const phase=phases[step],ratio=Math.max(0,Math.min(1,(t-start)/phase.duration));
   const from=way[step],to=way[step+1];
   const smooth=ratio*ratio*(3-2*ratio);
   const pos=lerp(from,to,phase.action==='walk'?smooth:0);
   visitor.position.set(pos[0],.04,pos[1]);
   const travel=Math.hypot(to[0]-from[0],to[1]-from[1]);
   visitor.rotation.y=travel>.0001?Math.atan2(to[0]-from[0],to[1]-from[1]):0;
   characterFactory.update(visitor,{speed:2.1},{
     action:phase.action==='walk'?'walk':step===1||step===5?'inspect':'pause'},time);
   // Conveyor motion is represented as a separate bag object, never parented to the visitor rig.
   const bagStart=point(xray,-xray.w*.60,0),bagEnd=point(xray,xray.w*.60,0);
   const conveyor=step===3?ratio:step>3?1:0;
   const bp=lerp(bagStart,bagEnd,conveyor);
   luggage.position.set(bp[0],.64,bp[1]);luggage.visible=step>=2&&step<=6;
   const approved=step>=6;
   lamp.position.set(xray.x,0,xray.z);
   lampMesh.visible=!(step>=3&&step<=4);
   lampWarning.visible=!lampMesh.visible;
   scanner.position.set(entry.x-3.5,0,entry.z);
   scanner.visible=approved;
   readerWaiting.visible=!approved;
   readerWaiting.position.set(entry.x-3.5,0,entry.z);
   gate.position.set(entry.x-1.8,1.1,entry.z);
   gate.rotation.y=approved?Math.PI*.45*Math.min(1,(t-(start-.001))/1.7):0;
   stats.stage=phase.name;stats.updates++;
   return {active:true,stage:phase.name,phaseIndex:step,progress:ratio,
     visitor:{x:pos[0],z:pos[1]},bag:{x:bp[0],z:bp[1]},
     gateOpen:approved,guardAAction:step===3?'inspect':'idle',
     guardBAction:step===5||step===6?'inspect':'idle'};
 }
 function diagnostic(){
   return {...stats,cycleSeconds:cycle,visible:root.visible,
     visitorPosition:[visitor.position.x,visitor.position.z],
     barrierAngle:gate.rotation.y,
     baggagePosition:[luggage.position.x,luggage.position.z]};
 }
 return {root,phases,simulate,diagnostic};
}
