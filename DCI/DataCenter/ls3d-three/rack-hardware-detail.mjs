// Selected-tray hardware only: teaching geometry, never a BOM or vendor CAD.
// One session is built on demand; no per-rack high-detail scene or extra RAF.
export function createRackHardwareDetail(THREE,{cube,material}){
 let active=null,requested=false;
 function box(g,n,x,y,z,w,h,d,c,metal=.48,rough=.45){const m=cube(g,x,y,z,w,h,d,c,metal,rough);m.name='hardware:'+n;return m;}
 function layer(root,name,offset){const g=new THREE.Group();g.name='hardware-layer:'+name;g.userData.lift=offset;root.add(g);return g;}
 function batch(g,name,points,size,color){
  if(!points.length)return;
  const sample=box(g,name,0,0,0,...size,color);g.remove(sample);
  const mesh=new THREE.InstancedMesh(sample.geometry,sample.material,points.length),matrix=new THREE.Matrix4();
  points.forEach(([x,y,z],i)=>{matrix.makeScale(...size);matrix.setPosition(x,y+size[1]/2,z);mesh.setMatrixAt(i,matrix);});
  mesh.name='hardware:'+name;mesh.instanceMatrix.needsUpdate=true;g.add(mesh);
 }
 function fan(g,x,y,z,r){
  box(g,'fan-cartridge',x,y,z,r*2.25,.13,r*2.25,'#242b30');
  const ring=new THREE.Mesh(new THREE.TorusGeometry(r,.012,5,20),material('#aeb6ba',.6,.4));
  ring.name='hardware:fan-grille';ring.rotation.x=Math.PI/2;ring.position.set(x,y+.14,z);g.add(ring);
  for(let b=0;b<5;b++){
   const blade=box(g,'fan-blade',x+Math.cos(b*1.257)*r*.42,y+.12,z+Math.sin(b*1.257)*r*.42,r*.8,.025,r*.28,'#555e63');blade.rotation.y=-b*1.257;
  }
  const hub=new THREE.Mesh(new THREE.CylinderGeometry(r*.24,r*.24,.04,12),material('#939da2',.7,.4));hub.position.set(x,y+.155,z);g.add(hub);
 }
 function clear(){
  if(!active)return;
  for(const [mesh,visible] of active.hidden)mesh.visible=visible;
  for(const [g,y] of active.external)g.position.y=y;
  active.root.removeFromParent();
  active.root.traverse(o=>{if(o.geometry&&o.geometry!==active.sharedGeometry)o.geometry.dispose();if(o.material?.map){o.material.map.dispose();o.material.dispose();}});
  active=null;requested=false;
 }
 function build(asset,module,rackUnits){
  clear();
  const info=module.userData.rackPart,w=asset.w,d=asset.d,unit=(asset.h-.46)/rackUnits,y=.23+(info.uStart-1)*unit;
  const kind=/gpu/i.test(asset.type)?'gpu':/storage/i.test(asset.type)?'storage':/switch|core|spine|leaf|mmr/i.test(asset.type)?'network':'compute';
  const root=new THREE.Group();root.name='rack-hardware-detail:'+asset.id;module.add(root);
  const hidden=[];
  // Replace only the selected module's solid body. Front ports and live subparts stay selectable.
  for(const m of module.children)if(m.isMesh&&/^rack-shape:sled\d+$/.test(m.name)){hidden.push([m,m.visible]);m.visible=false;}
  const base=layer(root,'chassis',0),pcb=layer(root,'pcb',d*.10),chips=layer(root,'packages',d*.22),thermal=layer(root,'cooling',d*.40),cover=layer(root,'cover',d*.58);
  const floor=box(base,'silver-chassis',0,y,-d*.14,w*.71,.035,d*.58,'#bac2c6');
  for(const side of [-1,1])box(base,'chassis-flange',side*w*.345,y+.035,-d*.14,.035,.12,d*.58,'#d4dadc');
  box(pcb,'green-mainboard',0,y+.055,-d*.13,w*.65,.030,d*.48,'#267d59',.12,.7);
  // Gold traces and mounting pads make the board readable as a physical assembly.
  batch(pcb,'copper-traces',Array.from({length:16},(_,i)=>[-w*.30+i*w*.039,y+.087,-d*.12]),[.009,.005,d*.39],'#c3a76b');
  batch(pcb,'mounting-pads',[-1,1].flatMap(x=>[-1,1].map(z=>[x*w*.29,y+.088,-d*.13+z*d*.20])),[.055,.008,.055],'#d4b563');
  const positions=kind==='gpu'?[]:kind==='storage'?[[-w*.17,-d*.08],[w*.17,-d*.08]]:kind==='network'?[[0,-d*.12]]:[[-w*.16,-d*.10],[w*.16,-d*.10]];
  positions.forEach(([x,z],i)=>{
   box(chips,'chip-package-'+i,x,y+.095,z,w*.14,.045,d*.14,'#242a30');
   box(chips,'metal-package-cap-'+i,x,y+.14,z,w*.12,.012,d*.12,'#bcc3c6');
   box(thermal,'heatsink-base-'+i,x,y+.16,z,w*.17,.02,d*.18,'#b1b8bc');
   batch(thermal,'heatsink-fins-'+i,Array.from({length:10},(_,j)=>[x-w*.078+j*w*.017,y+.18,z]),[w*.007,.10,d*.18],'#d0d5d8');
  });
  if(kind==='gpu'){
   // Existing four selectable GPU packages are preserved; their cooling plates lift in the exploded view.
   box(pcb,'gpu-carrier-board',0,y+.09,-d*.15,w*.62,.02,d*.32,'#278868',.12,.65);
  }
  if(kind!=='storage')for(const side of [-1,1]){
   box(pcb,'dimm-socket',side*w*.278,y+.093,-d*.05,w*.031,.03,d*.28,'#252c30');
   for(let row=0;row<3;row++){
    const x=side*(w*.225+row*w*.028);
    box(chips,'memory-dimm',x,y+.12,-d*.05,.018,.10,d*.24,'#2d7556',.1,.7);
    batch(chips,'dimm-chips',Array.from({length:5},(_,i)=>[x-.012,y+.15,-d*.14+i*d*.043]),[.026,.045,.05],'#1e2528');
   }
  }
  for(let i=0;i<4;i++)fan(thermal,-w*.24+i*w*.16,y+.10,-d*.345,w*.060);
  for(const side of [-1,1]){
   box(chips,'redundant-psu',side*w*.25,y+.10,d*.045,w*.13,.14,d*.12,'#aeb7bd');
   box(chips,'psu-receptacle',side*w*.25,y+.13,d*.113,w*.065,.07,.025,'#222a30');
  }
  if(kind==='storage')for(let drive=0;drive<6;drive++){
   const x=-w*.25+drive*w*.10;
   box(chips,'storage-drive-can',x,y+.12,-d*.17,w*.082,.10,d*.23,'#adb7be');
   box(chips,'drive-controller-pcb',x,y+.10,-d*.17,w*.072,.012,d*.19,'#246b52');
  }
  box(chips,kind==='network'?'switch-io-bank':'network-io-bank',0,y+.12,d*.075,w*.35,.11,d*.07,'#bac2c5');
  batch(chips,'io-port-cages',Array.from({length:8},(_,i)=>[-w*.15+i*w*.043,y+.14,d*.115]),[w*.033,.06,.035],'#202a31');
  box(cover,'lifted-metal-cover',0,y+.28,-d*.14,w*.70,.025,d*.55,'#d4d9dc');
  // Closed detail mode is an open tray. The lid appears only while exploded.
  cover.visible=false;
  const external=[];
  if(kind==='gpu')for(const child of module.children)if(child.userData?.rackPart?.kind?.startsWith('gpu-module-'))external.push([child,child.position.y]);
  root.traverse(o=>{if(o.isMesh){o.userData.assetId=asset.id;o.userData.partId=info.id;}});
  active={assetId:asset.id,partId:info.id,module,root,hidden,external,layers:[pcb,chips,thermal,cover],cover,progress:0,kind,sharedGeometry:floor.geometry};
 }
 function update(asset,module,enabled,dt,rackUnits=42){
  if(!enabled||!module){clear();return;}
  if(active?.assetId!==asset.id||active?.module!==module)build(asset,module,rackUnits);
  active.progress+=(Number(requested)-active.progress)*Math.min(1,dt*7);
  for(const g of active.layers)g.position.y=g.userData.lift*active.progress;
  active.cover.visible=active.progress>.015;
  for(const [g,y] of active.external)g.position.y=y+asset.d*.22*active.progress;
 }
 function explode(){requested=!requested;return requested;}
 function state(){return active?{assetId:active.assetId,partId:active.partId,kind:active.kind,exploded:requested,progress:active.progress,layers:5}:null;}
 return {update,explode,clear,state};
}
