// STEP 03 opt-in bootstrap; legacy remains the default renderer.
// Query: ?engine=three or ?engine=legacy. No extra requestAnimationFrame is created.
const params = new URLSearchParams(window.location.search);
const requested = params.get('engine') === 'three';
const T = window.__LS3D_TEST__;
const status = window.LS3D_THREE = {
  status: requested ? 'loading' : 'legacy',
  mode: requested ? 'three' : 'legacy',
  error: null, adapter: null, diagnostics: null, renders: 0,
  render() {}, fallback() {}
};
if (requested && T) {
  const viewport = document.getElementById('viewport');
  const oldCanvas = document.getElementById('scene');
  const canvas = document.createElement('canvas');
  canvas.id = 'sceneThree';
  canvas.setAttribute('aria-label','Three.js 3D 데이터센터 화면: 드래그 회전, 휠 확대, 설비 선택');
  canvas.style.cssText='position:absolute;inset:0;width:100%;height:100%;display:none;z-index:1;touch-action:none;cursor:grab;';
  oldCanvas.insertAdjacentElement('afterend',canvas);
  const label = document.createElement('span');
  label.id='ls3dEngineStatus';
  label.style.cssText='display:inline-block;margin-left:8px;padding:3px 6px;border-radius:5px;border:1px solid #5f8395;color:#d9f4fa;font:750 9px system-ui;';
  const heading = document.querySelector('.vp-title b');
  if (heading) heading.appendChild(label);
  function mark(text) { label.textContent=text; }
  mark('Three.js 불러오는 중');
  let pointer = null, second = null, lastPinch = null, lastMid = null, multiTouch = false;
  function clamp(x,a,b){return Math.min(b,Math.max(a,x))}
  function fallback(error) {
    status.error=String(error?.message || error);
    status.status='failed';
    status.mode='legacy';
    try {status.adapter?.dispose();} catch (_) {}
    status.adapter=null;
    canvas.style.display='none';
    canvas.style.pointerEvents='none';
    oldCanvas.style.visibility='';
    oldCanvas.style.pointerEvents='';
    mark('기존 WebGL 복귀');
    label.title=status.error;
    console.warn('LS3D Three.js fallback:',status.error);
  }
  status.fallback=fallback;
  function setPointers() {
    if(pointer&&second){
      lastPinch=Math.hypot(pointer.x-second.x,pointer.y-second.y);
      lastMid={x:(pointer.x+second.x)/2,y:(pointer.y+second.y)/2};
    }else {lastPinch=null;lastMid=null;}
  }
  canvas.addEventListener('pointerdown',e=>{
    if (status.status!=='ready')return;
    e.preventDefault();
    const p={id:e.pointerId,x:e.clientX,y:e.clientY,originX:e.clientX,originY:e.clientY,button:e.button,shift:e.shiftKey,moved:false};
    if(!pointer)pointer=p;
    else if(!second){second=p;multiTouch=true;pointer.moved=true;second.moved=true;}
    canvas.setPointerCapture(e.pointerId);
    setPointers();
  });
  canvas.addEventListener('pointermove',e=>{
    if(status.status!=='ready')return;
    const p=pointer?.id===e.pointerId?pointer:second?.id===e.pointerId?second:null;
    if(!p)return;
    const dx=e.clientX-p.x,dy=e.clientY-p.y;
    p.x=e.clientX;p.y=e.clientY;
    if(Math.abs(p.x-p.originX)+Math.abs(p.y-p.originY)>5)p.moved=true;
    if(pointer && second) {
      const pinch=Math.hypot(pointer.x-second.x,pointer.y-second.y);
      const center={x:(pointer.x+second.x)/2,y:(pointer.y+second.y)/2};
      if(!T.walkActive&&!T.workerViewId){
        if(lastPinch && pinch>3)
          T.cameraDesired.distance=clamp(T.cameraDesired.distance*lastPinch/pinch,.8,260);
        if(lastMid){
          const shiftX=center.x-lastMid.x,shiftY=center.y-lastMid.y;
          const yaw=T.cameraDesired.yaw;
          const factor=clamp(T.cameraDesired.distance*.0017,.012,.35);
          // Two-finger translation pans the camera target independently of pinch.
          T.cameraDesired.target[0]+=(-shiftX*Math.cos(yaw)+shiftY*Math.sin(yaw)*.30)*factor;
          T.cameraDesired.target[2]+=(shiftX*Math.sin(yaw)+shiftY*Math.cos(yaw)*.30)*factor;
        }
      }
      lastPinch=pinch;lastMid=center;
      return;
    }
    if(T.workerViewId) {
      T.workerYaw=T.workerYaw-dx*.005;
      T.workerPitch=clamp(T.workerPitch+dy*.004,-.72,.72);
    } else if(T.walkActive) {
      T.cameraDesired.yaw-=dx*.005;
      T.walkPitch=clamp(T.walkPitch+dy*.004,-1.15,1.15);
    } else if(p.button===2 || p.shift) {
      const factor=clamp(T.cameraDesired.distance*.0024,.02,.55);
      const yaw=T.cameraDesired.yaw;
      T.cameraDesired.target[0]+=(-dx*Math.cos(yaw)+dy*Math.sin(yaw)*.25)*factor;
      T.cameraDesired.target[2]+=(dx*Math.sin(yaw)+dy*Math.cos(yaw)*.25)*factor;
      T.cameraDesired.target[1]+=dy*factor*.55;
    } else {
      T.cameraDesired.yaw-=dx*.006;
      T.cameraDesired.pitch=clamp(T.cameraDesired.pitch+dy*.006,.08,1.56);
    }
  });
  canvas.addEventListener('pointerup',e=>{
    const p=pointer?.id===e.pointerId?pointer:second?.id===e.pointerId?second:null;
    if(p&&!p.moved&&!multiTouch&&!second&&p.button===0)
      status.adapter?.pick(e.clientX,e.clientY);
    if(pointer?.id===e.pointerId)pointer=second,second=null;
    else if(second?.id===e.pointerId)second=null;
    if(!pointer&&!second)multiTouch=false;
    setPointers();
  });
  canvas.addEventListener('pointercancel',()=>{
    pointer=null;second=null;lastPinch=null;lastMid=null;multiTouch=false;
  });
  canvas.addEventListener('lostpointercapture',e=>{
    if(pointer?.id===e.pointerId)pointer=second,second=null;
    else if(second?.id===e.pointerId)second=null;
    if(!pointer&&!second)multiTouch=false;
    setPointers();
  });
  canvas.addEventListener('contextmenu',e=>e.preventDefault());
  canvas.addEventListener('wheel',e=>{
    e.preventDefault();
    if(status.status!=='ready')return;
    if(T.workerViewId)T.workerPitch=clamp(T.workerPitch+Math.sign(e.deltaY)*.035,-.72,.72);
    else if(!T.walkActive)T.cameraDesired.distance=clamp(T.cameraDesired.distance*Math.exp(e.deltaY*(T.cameraDesired.distance<12?.0013:.001)),.12,260);
  },{passive:false});
  status.render=function(dt,view) {
    if(status.status!=='ready')return;
    try {
      status.adapter.render(dt,view);
      status.renders++;
      status.diagnostics={...status.adapter.diagnostic};
    } catch(e){ fallback(e);throw e; }
  };
  (async function(){
    try {
      const support=document.createElement('canvas').getContext('webgl2');
      if(!support)throw new Error('Three.js 0.186.1 requires WebGL2; browser only supports legacy WebGL.');
      const {createSceneAdapter}=await import('./scene-adapter.mjs?v=20261025-carton-samsung-sdi');
      const adapter=createSceneAdapter(T,canvas);
      status.adapter=adapter;
      status.status='ready';
      status.mode='three';
      canvas.style.display='block';
      canvas.style.pointerEvents='auto';
      oldCanvas.style.visibility='hidden';
      oldCanvas.style.pointerEvents='none';
      mark('Three.js r'+adapter.THREE.REVISION+' · PBR');
    } catch(e){ fallback(e); }
  })();
} else if(requested && !T) {
  status.status='failed';status.error='LS3D host API unavailable';
  console.warn(status.error);
}

