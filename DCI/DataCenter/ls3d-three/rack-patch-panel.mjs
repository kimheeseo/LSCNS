// STEP 1: inspectable 1U patch panel. Geometry is owned by the rack and disposed with it.
export function createRackPatchPanel(THREE){
 const sessions=new Map();
 const clamp=x=>Math.max(0,Math.min(1,x));
 function decorate(asset,built){
  const part=built.parts.find(p=>p.kind==='patch-panel');
  const panel=built.root?.children?.find(g=>g.userData?.rackPart?.id===part?.id);
  if(!part||!panel)return null;
  const w=asset.w||3,d=asset.d||3,u=(asset.h-.46)/built.rackUnits;
  const y=.23+(part.uStart-1)*u, z0=-d*.29, front=-d*.512, rear=-d*.06;
  const body=panel.children.find(o=>o.name==='rack-shape:'+part.label+'-body');
  if(body)body.visible=false;
  function box(parent,name,x,yy,z,sx,sy,sz,color){
   const mesh=new THREE.Mesh(new THREE.BoxGeometry(sx,sy,sz),new THREE.MeshStandardMaterial({color,metalness:.25,roughness:.55}));
   mesh.name='patch-panel:'+name;mesh.position.set(x,yy,z);mesh.userData.assetId=asset.id;mesh.userData.partId=part.id;parent.add(mesh);return mesh;
  }
  // Existing front port face/labels stay in place; replace the opaque shell by a tray.
  const width=w*.76,depth=d*.46, height=Math.max(.09,u-.02);
  box(panel,'floor',0,y-height*.42,z0,width,.014,depth,'#354959');
  for(const side of [-1,1])box(panel,'side-'+side,side*width*.495,y,z0,.016,height,depth,'#536575');
  box(panel,'rear',0,y,rear,width,height,.018,'#43586a');
  const lid=new THREE.Group();lid.name='patch-panel:hinged-lid';lid.position.set(0,y+height*.5,rear);panel.add(lid);
  box(lid,'lid',0,0,-depth*.5,width,.014,depth,'#405666');
  for(const sx of [-1,1])box(panel,'rack-ear-'+sx,sx*w*.39,y,front+.025,w*.065,height*.78,.025,'#9caeb9');
  const tray=box(panel,'splice-tray',0,y-height*.16,z0,width*.46,.015,depth*.33,'#adb8b8');
  for(let i=0;i<6;i++)box(panel,'splice-sleeve-'+(i+1),(i-2.5)*width*.055,y-height*.105,z0,width*.041,.012,depth*.10,'#f0e8d5');
  for(const sx of [-1,1])box(panel,'cable-clip-'+sx,sx*width*.37,y-height*.16,rear-depth*.2,.040,.025,.055,'#a7bac6');
  const fibers=[];const palette=[0xf1ca3b,0x63afee,0xd481ad,0x85d2a2,0xe6a65b,0xb8a1e2];
  for(let i=0;i<6;i++){
   const x=-w*.27+(i+3)*w*.049,cy=y-height*.08,zc=z0;
   const curve=new THREE.CatmullRomCurve3([
    new THREE.Vector3(x,cy,front+.05),new THREE.Vector3(x*.82,cy,zc-depth*.19),
    new THREE.Vector3(x*.55,cy,zc+depth*.05),new THREE.Vector3(x*.67,cy,zc+depth*.17),
    new THREE.Vector3((i-2.5)*width*.055,cy,zc+depth*.12)
   ],false,'centripetal');
   const mesh=new THREE.Mesh(new THREE.TubeGeometry(curve,24,.006,4,false),new THREE.MeshStandardMaterial({color:palette[i],emissive:0x000000}));
   mesh.name='patch-panel:fiber-path-'+(i+1);mesh.userData.assetId=asset.id;mesh.userData.partId=part.id;panel.add(mesh);fibers.push(mesh);
   const port={id:asset.id+':patch-panel-port-'+i,assetId:asset.id,kind:'patch-panel-port',label:'광 패치패널 LC duplex 포트 '+(i+1),description:'전면 LC duplex '+(i+1)+'번 포트에서 내부 피그테일과 융착 보호슬리브로 이어지는 교육용 광경로.',uStart:part.uStart,uHeight:1,portIndex:i};
   built.parts.push(port);built.partMap.set(port.id,port);
   const hit=box(panel,'lc-duplex-pick-'+i,x,y+height*.10,front-.034,w*.033,.047,.015,'#377eac');
   hit.userData.partId=port.id;
  }
  part.description='19인치 1U 수동 광 패치패널. 전면 LC duplex 포트, 상단 힌지 커버, 내부 피그테일 루프·융착 트레이·보호슬리브·고정 클립을 확인할 수 있습니다. 포트 선택 시 대응 광경로가 강조됩니다. 교육용 배치이며 실제 제조사 치수가 아닙니다.';
  const s={panel,lid,fibers,targetSlide:0,slide:0,targetCover:0,cover:0,port:0,part};sessions.set(asset.id,s);return s;
 }
 function toggleSlide(id){const s=sessions.get(id);if(!s)return false;if(s.targetCover)s.targetCover=0;s.targetSlide=1-s.targetSlide;return true;}
 function toggleCover(id){const s=sessions.get(id);if(!s||s.targetSlide<1)return false;s.targetCover=1-s.targetCover;return true;}
 function selectPort(id,n){const s=sessions.get(id);if(!s)return false;s.port=(n+6)%6;s.fibers.forEach((f,i)=>{f.material.emissive.setHex(i===s.port?0xffa600:0);f.material.emissiveIntensity=i===s.port?.8:0;});return true;}
 function update(id,dt){const s=sessions.get(id);if(!s)return null;const a=1-Math.exp(-Math.min(.1,Math.max(0,dt))*5);
  s.slide+= (s.targetSlide-s.slide)*a;s.cover+=(s.targetCover-s.cover)*a;
  s.panel.position.z=-s.slide*.48;s.lid.rotation.x=-s.cover*1.55;
  return state(id);
 }
 function state(id){const s=sessions.get(id);return s?{slide:s.slide,targetSlide:s.targetSlide,cover:s.cover,targetCover:s.targetCover,port:s.port,partId:s.part.id}:null;}
 function drop(id){sessions.delete(id);}function reset(){sessions.clear();}
 return {decorate,toggleSlide,toggleCover,selectPort,update,state,drop,reset};
}
