import test from 'node:test';import assert from 'node:assert/strict';
import {rackCameraView,applyRackCamera,rackFocusActive} from './rack-camera-policy.mjs';
const a={id:'tor-a',x:23,z:-17,w:3,d:3,h:6};
for(const [width,height] of [[1400,333],[1000,650],[390,320],[390,700]])for(const side of ['front','rear']){
 test(`rack fits ${width}x${height} ${side}`,()=>{
  const v=rackCameraView(a,{width,height,side});
  assert.ok(v.target.every(Number.isFinite)&&Number.isFinite(v.distance));
  assert.equal(v.yaw,side==='front'?Math.PI:0);
  const tan=Math.tan(Math.PI/8),aspect=width/height,narrow=width<=640;
  const usableW=width-(narrow?0:Math.min(322,width*.34))-32;
  const usableH=height-(narrow?Math.min(height*.31,260):0)-32;
  const nearDistance=v.distance-a.d*.65;
  assert.ok(nearDistance*tan*usableH/height>a.h*.54,'height fits above/beside controls');
  assert.ok(nearDistance*tan*aspect*usableW/width>a.w*.76,'width fits beside controls');
 });
}
test('front camera exits microscope, snaps live camera, isolates campus and restores on navigation',()=>{
 const T={selected:a,currentZone:'campus',facilityMode:'aidc',cameraDesired:{target:[0,0,0],distance:.1},camera:{target:[0,0,0],distance:.1},closeDetail(){this.closed=true}};
 const f=applyRackCamera(T,a,{width:1400,height:333,side:'front'});
 assert.ok(T.closed);assert.deepEqual(T.camera,T.cameraDesired);assert.notEqual(T.camera.target,T.cameraDesired.target);
 assert.ok(rackFocusActive(f,T));T.currentZone='hall';assert.equal(rackFocusActive(f,T),false);
 T.currentZone='campus';T.cameraDesired.distance=161;assert.equal(rackFocusActive(f,T),false);
});
test('no camera takeover during walking or worker view',()=>{
 for(const state of [{walkActive:true},{workerViewId:'worker-1'}])assert.equal(applyRackCamera({...state,cameraDesired:{}},a),null);
});
