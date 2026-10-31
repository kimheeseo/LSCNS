import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
const html=fs.readFileSync(new URL('../LS_Datacenter_3D.html',import.meta.url),'utf8');
const shared=html.slice(html.indexOf('// Shared camera controls'),html.indexOf('\nfunction render3D'));
const controls=html.slice(html.indexOf('function initControls()'),html.indexOf('\nfunction bindUI()'));
function harness(){
 const make=()=>({listeners:{},style:{},addEventListener(type,fn){(this.listeners[type]??=[]).push(fn)},emit(type,event={}){for(const fn of this.listeners[type]??[])fn({preventDefault(){event.prevented=true},...event})},setPointerCapture(){},focus(){ctx.document.activeElement=this},tagName:'CANVAS'});
 const canvas=make(),mode=make(),hint={};
 const document=Object.assign(make(),{activeElement:{tagName:'BODY'},hidden:false,querySelectorAll:()=>[]});
 const ctx=vm.createContext({document,window:make(),canvas,$:id=>id==='cameraDragMode'?mode:id==='cameraHint'?hint:null,
  cameraDesired:{yaw:0,pitch:.7,distance:40,target:[0,0,0]},pressedKeys:new Set(),walk:false,workerViewId:null,
  walkPos:[0,1.7,0],walkPitch:0,workerViewYaw:0,workerViewPitch:0,clamp:(x,a,b)=>Math.min(b,Math.max(a,x)),
  phase1Inspecting:()=>false,LS3D_ZOOM_MIN:1e-10,collide:()=>false,toggleWalk(){},pickAt(){},
 });
 vm.runInContext(shared+'\n'+controls+'\ninitControls();',ctx);
 return ctx;
}
const near=(a,b)=>assert.ok(Math.abs(a-b)<1e-9,`${a} != ${b}`);
test('all classic inline scripts parse and removed modules cannot create MEP/FPS UI',()=>{
 for(const m of html.matchAll(/<script([^>]*)>([\s\S]*?)<\/script>/g)){if(!/importmap|type="module"/.test(m[1]))new vm.Script(m[2]);}
 assert.doesNotMatch(html,/<script[^>]+src="[^\"]*Phase45/);
 const optical=fs.readFileSync(new URL('../LS_Datacenter_3D_Phase2A.js',import.meta.url),'utf8');
 new vm.Script(optical);assert.doesNotMatch(optical,/phase2-mep|LS3D_PHASE2_A/);assert.match(optical,/id="phase2-opt-open"/);
 assert.match(html,/__LS3D_TEST__=\{panCamera,zoomCamera,dragCamera,/);
});
test('arrows pan without walking and prevent browser scrolling',()=>{
 const c=harness(),e={key:'ArrowRight'};c.document.emit('keydown',e);c.updateWalk(.05);
 assert.ok(e.prevented);assert.ok(c.cameraDesired.target[0]>0);near(c.cameraDesired.target[2],0);
 c.document.emit('keyup',{key:'ArrowRight'});const x=c.cameraDesired.target[0];c.updateWalk(.05);near(c.cameraDesired.target[0],x);
 c.cameraDesired.yaw=Math.PI/2;c.pressedKeys.add('arrowup');c.updateWalk(.05);assert.ok(c.cameraDesired.target[0]<x);
});
test('diagonal speed is normalized and Shift doubles travel',()=>{
 const a=harness(),b=harness();a.pressedKeys.add('arrowright');b.pressedKeys.add('arrowright');b.pressedKeys.add('arrowup');a.updateWalk(.05);b.updateWalk(.05);
 near(Math.hypot(...a.cameraDesired.target),Math.hypot(...b.cameraDesired.target));
 const distance=a.cameraDesired.target[0];a.pressedKeys.add('shift');a.updateWalk(.05);near(a.cameraDesired.target[0],3*distance);
});
test('walking arrows move relative to view and respect collision',()=>{
 const c=harness();c.walk=true;c.pressedKeys.add('arrowright');c.updateWalk(.05);assert.ok(c.walkPos[0]<0);
 c.pressedKeys.clear();c.pressedKeys.add('arrowup');c.updateWalk(.05);assert.ok(c.walkPos[2]>0);
 c.collide=()=>true;const before=c.walkPos.slice();c.updateWalk(.05);assert.deepEqual(c.walkPos,before);
});
test('input fields, contenteditable and visible dialogs suppress navigation',()=>{
 const c=harness();for(const el of [{tagName:'INPUT'},{tagName:'SELECT'},{tagName:'DIV',isContentEditable:true}]){
 c.document.activeElement=el;c.document.emit('keydown',{key:'ArrowUp'});assert.equal(c.pressedKeys.size,0);
 }
 c.document.activeElement={tagName:'BODY'};c.document.querySelectorAll=()=>[{getClientRects:()=>[{}]}];c.document.emit('keydown',{key:'ArrowUp'});assert.equal(c.pressedKeys.size,0);
 c.document.querySelectorAll=()=>[{getClientRects:()=>[]}];c.document.emit('keydown',{key:'ArrowUp'});assert.equal(c.pressedKeys.size,1);
});
test('blur, hidden tab and input focus clear held movement keys',()=>{
 const c=harness();c.pressedKeys.add('arrowup');c.window.emit('blur');assert.equal(c.pressedKeys.size,0);
 c.pressedKeys.add('arrowup');c.document.hidden=true;c.document.emit('visibilitychange');assert.equal(c.pressedKeys.size,0);
 c.pressedKeys.add('arrowup');c.document.activeElement={tagName:'INPUT'};c.document.emit('focusin');assert.equal(c.pressedKeys.size,0);
});
test('right drag pans on the camera screen plane at rotated viewpoints',()=>{
 const c=harness();c.cameraDesired.yaw=Math.PI/2;c.dragCamera(10,0,{button:2});near(c.cameraDesired.target[0],0);assert.ok(c.cameraDesired.target[2]>0);
 const before=c.cameraDesired.target.slice();c.dragCamera(0,10,{button:2});assert.ok(c.cameraDesired.target[0]<before[0]);assert.ok(c.cameraDesired.target[1]>before[1]);
});
test('middle/Ctrl drag and zoom mode zoom both ways with distance limits',()=>{
 const c=harness();c.dragCamera(0,-20,{button:1});assert.ok(c.cameraDesired.distance<40);
 const before=c.cameraDesired.distance;c.dragCamera(0,20,{button:0,ctrlKey:true});assert.ok(c.cameraDesired.distance>before);
 vm.runInContext("cameraDragMode='zoom'",c);c.dragCamera(0,-20,{button:0});assert.ok(c.cameraDesired.distance<40);
 for(let i=0;i<100;i++)c.zoomCamera(1);near(c.cameraDesired.distance,260);
 for(let i=0;i<100;i++)c.zoomCamera(-1);near(c.cameraDesired.distance,.12);
});
test('left drag pan mode, rotation and walking look remain usable',()=>{
 const c=harness();vm.runInContext("cameraDragMode='pan'",c);c.dragCamera(10,0,{button:0});assert.ok(c.cameraDesired.target[0]<0);near(c.cameraDesired.yaw,0);
 vm.runInContext("cameraDragMode='rotate'",c);c.dragCamera(10,0,{button:0});assert.ok(c.cameraDesired.yaw<0);
 c.walk=true;const before=c.cameraDesired.distance;c.dragCamera(0,-20,{button:1});near(c.cameraDesired.distance,before);assert.ok(c.walkPitch<0);
});
test('Three.js pointer events delegate to shared controls; pinch does not select',async()=>{
 const source=fs.readFileSync(new URL('./bootstrap.mjs',import.meta.url),'utf8');
 const calls=[],elements={};
 const make=()=>({style:{},listeners:{},setAttribute(){},appendChild(){},insertAdjacentElement(){},getContext:()=>null,setPointerCapture(){},focus(){calls.push(['focus'])},addEventListener(type,fn){this.listeners[type]=fn}});
 const T={cameraDesired:{distance:40},dragCamera:(...args)=>calls.push(['drag',...args]),panCamera:(...args)=>calls.push(['pan',...args]),zoomCamera:(...args)=>calls.push(['zoom',...args])};
 let canvas;const ctx=vm.createContext({URLSearchParams,console:{warn(){}},window:{location:{search:'?engine=three'},__LS3D_TEST__:T},document:{getElementById:id=>elements[id]??=(make()),querySelector:()=>make(),createElement(tag){const el=make();if(tag==='canvas'&&!canvas)canvas=el;return el;}}});
 vm.runInContext(source,ctx);await Promise.resolve();const status=ctx.window.LS3D_THREE;status.status='ready';status.adapter={pick:()=>calls.push(['pick'])};
 const emit=(type,extra)=>canvas.listeners[type]({preventDefault(){},pointerId:1,clientX:0,clientY:0,button:0,...extra});
 emit('pointerdown',{button:2});emit('pointermove',{clientX:30});emit('pointerup',{clientX:30});assert.equal(calls.filter(c=>c[0]==='drag').length,1);assert.equal(calls.filter(c=>c[0]==='pick').length,0);
 emit('pointerdown',{});emit('pointerup',{});assert.equal(calls.filter(c=>c[0]==='pick').length,1);
 emit('pointerdown',{});emit('pointerdown',{pointerId:2,clientX:20});emit('pointermove',{pointerId:2,clientX:30});emit('pointerup',{pointerId:2,clientX:30});emit('pointerup',{});
 assert.ok(calls.some(c=>c[0]==='zoom'));assert.ok(calls.some(c=>c[0]==='pan'));assert.equal(calls.filter(c=>c[0]==='pick').length,1);
});
