// STEP 18 — Three.js camera, responsive HUD, and sustained-frame render policy.
// No timers, new requestAnimationFrame, browser globals, or host data changes.
export function createViewportPolicy({mobileBreakpoint=640}={}){
  let averageMs=16.7,slowFrames=0,fastFrames=0,degraded=false,samples=0;
  let ratio=1,requested='medium',mobile=false;
  const clamp=(x,lo,hi)=>Math.max(lo,Math.min(hi,x));
  function sample(deltaSeconds){
    const ms=deltaSeconds*1000;
    if(!Number.isFinite(ms)||ms<5||ms>220)return summary();
    averageMs=samples?averageMs*.94+ms*.06:ms;
    samples++;
    if(averageMs>38){slowFrames++;fastFrames=0;}
    else if(averageMs<21){fastFrames++;slowFrames=0;}
    else{slowFrames=0;fastFrames=0;}
    // Hysteresis avoids per-frame pixelRatio resize thrashing.
    if(!degraded&&slowFrames>=48){degraded=true;slowFrames=fastFrames=0;}
    else if(degraded&&fastFrames>=210){degraded=false;slowFrames=fastFrames=0;}
    return summary();
  }
  function settings({width=800,quality='medium',devicePixelRatio=1}={}){
    requested=['low','medium','high'].includes(quality)?quality:'medium';
    mobile=width<=mobileBreakpoint;
    const caps={low:.8,medium:1.1,high:1.55};
    const cap=mobile?{low:.8,medium:1.0,high:1.3}[requested]:caps[requested];
    // Manual low/high settings are never silently reclassified as medium.
    ratio=clamp(Math.min(devicePixelRatio||1,cap)*(degraded&&requested==='medium'?.77:1),.55,1.55);
    return {dpr:ratio,quality:requested,mobile,adaptive:degraded,
      frameAverageMs:averageMs};
  }
  function summary(){return {sampleCount:samples,averageFrameMs:averageMs,
    degraded,slowFrames,fastFrames,dpr:ratio,quality:requested,mobile};}
  return {sample,settings,summary};
}

// Only occluding architectural walls/roof/ceiling along the camera->selected
// asset line of sight are temporarily hidden; geometry and cached materials
// stay untouched. Restore on zoom-out, section change, zone change and dispose.
export function createCameraOcclusion(THREE,{zoneRoot}){
  const hidden=new Set(),ray=new THREE.Raycaster();
  const direction=new THREE.Vector3(),origin=new THREE.Vector3(),destination=new THREE.Vector3();
  let frame=0,lastZone='',lastSelected='';
  const relevant=new Set(['wall','roof','ceiling']);
  function restore(){for(const mesh of hidden)mesh.visible=true;hidden.clear();}
  function update({eye,selected,currentZone,facilityMode,inspect=false,showSection=false}={}){
    frame++;
    if(inspect||showSection||!selected||currentZone==='campus'||
      selected.zone!==currentZone||facilityMode==='server-room'||
      !Array.isArray(eye)||eye.some(v=>!Number.isFinite(v))){
      restore();return summary();
    }
    const changed=lastZone!==currentZone||lastSelected!==selected.id;
    lastZone=currentZone;lastSelected=selected.id;
    const targetY=Math.min(Math.max((selected.h||4)*.50,.85),4.5);
    origin.set(eye[0],eye[1],eye[2]);
    destination.set(selected.x,targetY,selected.z);
    direction.subVectors(destination,origin);
    const length=direction.length();
    // Restore immediately at zoom-out, even between 6-frame visibility scans.
    if(length>24||length<1){restore();return summary();}
    // Redraw at 10 Hz-ish rather than raycasting every GPU frame.
    if(!changed&&frame%6!==0)return summary();
    restore();
    let room=null;
    for(const child of zoneRoot.children){
      if(child.name==='arch:zone:'+currentZone){room=child;break;}
    }
    if(!room)return summary();
    const candidates=[];
    room.traverse(mesh=>{
      if(mesh.isMesh&&relevant.has(mesh.userData?.archKind))candidates.push(mesh);
    });
    if(!candidates.length)return summary();
    // Geometry is assumed already in the scene graph; update matrices only
    // when performing the infrequent, targeted visibility query.
    room.updateWorldMatrix?.(true,true);
    ray.set(origin,direction.normalize());
    ray.near=0;ray.far=Math.max(.1,length-.35);
    for(const hit of ray.intersectObjects(candidates,false)){
      const object=hit.object;
      if(object.visible){object.visible=false;hidden.add(object);}
    }
    return summary();
  }
  function summary(){return {hiddenOccluders:hidden.size,checkedFrame:frame};}
  function dispose(){restore();}
  return {update,restore,summary,dispose};
}

// The host already owns overlay DOM. This function only moves existing
// control panels and prevents them from covering one another on phones.
export function layoutSceneOverlays({width,height,mmr,colo,mep,rack}){
  const narrow=width<=640,short=height<520;
  const zone=mmr&&!mmr.hidden?mmr:colo&&!colo.hidden?colo:null;
  const rackVisible=!!rack&&!rack.hidden;
  if(zone){
    zone.style.maxWidth=narrow?'calc(100% - 16px)':'min(330px,calc(100% - 24px))';
    zone.style.left=narrow?'8px':'12px';
    zone.style.top=narrow?'58px':'73px';
    zone.style.right='auto';
    zone.style.maxHeight=narrow?(short?'min(30%,140px)':'min(38%,250px)'):'55%';
    zone.style.overflowY='auto';
  }
  if(mep){
    mep.style.left='auto';
    mep.style.right=narrow?'8px':'12px';
    mep.style.maxWidth=narrow?'calc(100% - 16px)':'min(322px,calc(100% - 24px))';
    mep.style.top=narrow&&zone?'auto':narrow?'54px':'83px';
    mep.style.bottom=narrow&&zone?(rackVisible?'calc(32% + 16px)':'12px'):'auto';
    // Keep the collapsed MEP control reachable while the zone panel is open.
    const details=mep.querySelector?.('details');
    if(details){
      details.style.maxHeight=narrow?'min(32vh,220px)':'min(52vh,420px)';
      const congested=!!narrow&&!!zone;
      if(congested&&details.open&&mep.dataset.step18Collapsed!=='1'){
        details.open=false;mep.dataset.step18Collapsed='1';
      }
      if(!congested)mep.dataset.step18Collapsed='0';
    }
  }
  if(rack){
    rack.style.maxWidth=narrow?'calc(100% - 16px)':'min(302px,calc(100% - 20px))';
    rack.style.right=narrow?'8px':'10px';
    rack.style.bottom=narrow?'8px':'10px';
    if(narrow){
      rack.style.maxHeight='min(31%,260px)';
      rack.style.overflowY='auto';
    }else{
      rack.style.maxHeight='min(76vh,550px)';rack.style.overflowY='auto';
    }
  }
  return {mobile:narrow,short,zonePanel:!!zone,rackPanel:rackVisible};
}
