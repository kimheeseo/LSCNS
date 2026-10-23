// Rack cameras use world coordinates and reserve space for the floating controls.
export function rackCameraView(a,{width=1000,height=600,side='front',fov=Math.PI/4}={}){
 const safeWidth=Math.max(160,Number(width)||1000),safeHeight=Math.max(120,Number(height)||600);
 const narrow=safeWidth<=640,sidePanel=narrow?0:Math.min(322,safeWidth*.34);
 const bottomPanel=narrow?Math.min(safeHeight*.31,260):0;
 const tan=Math.tan(fov*.5),aspect=safeWidth/safeHeight;
 const usableW=Math.max(80,safeWidth-sidePanel-32),usableH=Math.max(70,safeHeight-bottomPanel-32);
 const halfH=(a.h||6)*.54,halfW=(a.w||3)*.76;
 const distance=Math.max(halfH/(tan*usableH/safeHeight),halfW/(tan*aspect*usableW/safeWidth))*1.12+(a.d||3)*.65;
 const yaw=side==='rear'?0:Math.PI;
 const offsetX=sidePanel/safeWidth*distance*tan*aspect;
 const offsetY=bottomPanel/safeHeight*distance*tan;
 return {target:[a.x+Math.cos(yaw)*offsetX,(a.h||6)*.52-offsetY,a.z],distance,pitch:.12,yaw};
}
export function applyRackCamera(T,a,options={}){
 if(!a||!T.cameraDesired||T.workerViewId||T.walkActive)return null;
 // Closing the legacy microscope inspector prevents its local-origin target from
 // replacing the rack's world-space target on the following animation frame.
 T.closeDetail?.();
 if(T.selected?.id!==a.id)T.pickAsset?.(a,false);
 const view=rackCameraView(a,options);
 Object.assign(T.cameraDesired,view,{target:view.target.slice()});
 // Do not sweep through rooms, walls, or an old microscope position en route.
 if(T.camera)Object.assign(T.camera,view,{target:view.target.slice()});
 return {id:a.id,side:options.side==='rear'?'rear':'front',zone:T.currentZone,mode:T.facilityMode,distance:view.distance};
}
export function rackFocusActive(focus,T){
 return !!focus&&T.selected?.id===focus.id&&T.currentZone===focus.zone&&T.facilityMode===focus.mode&&
  !T.workerViewId&&!T.walkActive&&T.cameraDesired.distance<=Math.max(45,focus.distance*3);
}
