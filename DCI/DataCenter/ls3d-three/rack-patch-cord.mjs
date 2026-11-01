// Educational LC patch-cord demonstration for the existing 3D rack.
// One pair of cords travels along -Z (in front of the patch panel), then
// inserts into the panel's physical port positions. No extra RAF/timers.
export function createRackPatchCord(THREE){
 const registry=new Map();
 const metrics={created:0,connections:0,updates:0};
 const smooth=x=>{const t=Math.max(0,Math.min(1,x));return t*t*(3-2*t)};
 function cableMesh(name,color){
   const material=new THREE.MeshStandardMaterial({color,metalness:.04,roughness:.67,side:THREE.DoubleSide});
   const mesh=new THREE.Mesh(new THREE.TubeGeometry(
     new THREE.CatmullRomCurve3([new THREE.Vector3(0,0,0),new THREE.Vector3(0,.04,.1)]),
     16,.018,6,false),material);
   mesh.name=name;mesh.userData.noPick=true;
   return mesh;
 }
 function replaceCurve(mesh,points){
   const curve=new THREE.CatmullRomCurve3(points.map(p=>new THREE.Vector3(...p)));
   const old=mesh.geometry;
   mesh.geometry=new THREE.TubeGeometry(curve,20,.018,6,false);
   old?.dispose?.();
 }
 function addShape(parent,name,x,y,z,w,h,d,color){
   const obj=new THREE.Mesh(new THREE.BoxGeometry(w,h,d),
     new THREE.MeshStandardMaterial({color,metalness:.25,roughness:.43}));
   obj.name=name;obj.position.set(x,y,z);obj.userData.noPick=true;
   parent.add(obj);return obj;
 }
 function decorate(asset,built,root){
   if(registry.has(asset.id))drop(asset.id);
   const panel=built.parts.find(p=>p.kind==='patch-panel');
   if(!panel)return null;
   const w=asset.w||3,d=asset.d||3,h=asset.h||6;
   const units=built.rackUnits||42,unit=(h-.46)/units,lower=.23;
   const y=lower+(panel.uStart-1)*unit+Math.max(.045,unit*.37);
   const faceZ=-d*.512;
   const colors=[0x69e7c3,0xf1cb6a];
   const lines=[],group=new THREE.Group();
   group.name='rack-patch-cord-demo:'+asset.id;
   group.userData.assetId=asset.id;root.add(group);
   for(let index=0;index<2;index++){
     // These are positions of the model's real port-4 and port-5 slots.
     const slot=4+index;
     const x=-w*.27+slot*w*.049;
     const guideX=index===0?-w*.35:w*.35;
     const plug=new THREE.Group();plug.name='lc-connector-moving-'+index;
     plug.userData.noPick=true;group.add(plug);
     addShape(plug,'LC housing',0,0,-.045,.086,.054,.14,index===0?0x5fe5cf:0xe5d279);
     addShape(plug,'LC latch',0,.035,-.027,.048,.013,.085,0xf1f7fb);
     addShape(plug,'LC ferrule',0,0,.034,.033,.025,.036,0xd8e4ed);
     const jacket=cableMesh('optical-patch-cord-'+index,colors[index]);group.add(jacket);
     lines.push({x,y,guideX,index,plug,jacket,lastStep:-1});
   }
   const record={assetId:asset.id,group,lines,faceZ,d,target:0,progress:0,plugged:false,panelId:panel.id};
   registry.set(asset.id,record);metrics.created++;paint(record,true);
   return {panelId:panel.id};
 }
 function paint(record,force=false){
   const p=smooth(record.progress),distance=.76;
   // tip moves toward the patch-panel mouth; the cable follows in real time.
   const plugZ=record.faceZ-.025-distance*(1-p);
   record.lines.forEach(line=>{
     line.plug.position.set(line.x,line.y,plugZ);
     const progressStep=Math.round(p*50);
     if(force||progressStep!==line.lastStep){
       // Leave an arched slack loop from the internal cable manager to
       // the LC connector; its final segment extends/shortens with insertion.
       const bendZ=record.faceZ-.58;
       const route=[
         [line.guideX,line.y-.37,-record.d*.21],
         [line.guideX,line.y-.15,-record.d*.40],
         [line.guideX,line.y+.08,bendZ],
         [line.x,line.y+.16,bendZ-.16],
         [line.x,line.y+.015,plugZ-.20],
         [line.x,line.y,plugZ-.10]
       ];
       replaceCurve(line.jacket,route);
       line.lastStep=progressStep;
     }
   });
 }
 function toggle(id){
   const r=registry.get(id);if(!r)return null;
   r.target=r.target?0:1;metrics.connections++;
   return {connected:!!r.target,portCount:r.lines.length,panelId:r.panelId};
 }
 function update(id,dt=.016){
   const r=registry.get(id);if(!r)return null;
   const t=1-Math.exp(-Math.min(.15,Math.max(0,dt))*3);
   const next=r.progress+(r.target-r.progress)*t;
   if(Math.abs(next-r.progress)>.0002){
     r.progress=next;paint(r);metrics.updates++;
   }else if(Math.abs(r.progress-r.target)>.0002){
     r.progress=r.target;paint(r);
   }
   r.plugged=r.progress>.97;
   return state(id);
 }
 function state(id){
   const r=registry.get(id);
   return r?{connected:!!r.target,progress:r.progress,plugged:r.plugged,panelId:r.panelId,portCount:r.lines.length}:null;
 }
 function drop(id){
   const r=registry.get(id);if(!r)return;
   r.group.parent?.remove(r.group);
   r.group.traverse(o=>{
     if(o.geometry)o.geometry.dispose?.();
     if(o.material){
       if(Array.isArray(o.material))o.material.forEach(m=>m.dispose?.());
       else o.material.dispose?.();
     }
   });
   registry.delete(id);
 }
 function reset(){for(const id of [...registry.keys()])drop(id);}
 return {decorate,toggle,update,state,drop,reset,metrics};
}
