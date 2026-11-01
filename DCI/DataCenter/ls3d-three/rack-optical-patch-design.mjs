// Visual-only STEP 1 refresh for the existing Network Rack patch-panel slot.
// Retains the original part ID, U allocation, LC cable insertion coordinates and UI.
export function createRackOpticalPatchDesign(THREE,{cube}){
 const sessions=new Map();
 function decorate(asset,built,rackRoot){
  if(asset.type!=='Top-of-Rack switch')return null;
  const part=built.parts.find(p=>p.kind==='patch-panel');
  const panel=rackRoot?.children.find(g=>g.userData?.rackPart?.id===part?.id);
  if(!part||!panel)return null;
  // Replace only the old patch-panel primitive meshes. Other rack parts stay intact.
  for(const child of panel.children)child.visible=false;
  const {w,d,h,id}=asset,unit=(h-.46)/built.rackUnits;
  const lower=.23+(part.uStart-1)*unit,height=Math.max(.10,unit*.88),width=w*.76;
  const front=-d*.512,rear=front+d*.33,depth=rear-front,mid=(front+rear)*.5;
  const path=new THREE.Group();path.name='rack-optical-patch:physical-assembly';panel.add(path);
  function box(parent,name,x,base,z,sx,sy,sz,color){
   const mesh=cube(parent,x,base,z,sx,sy,sz,color,.38,.46);
   mesh.name='rack-optical-patch:'+name;mesh.userData.assetId=id;mesh.userData.partId=part.id;
   return mesh;
  }
  // 19-inch rack tray, mounting ears, side walls and cable gland.
  box(path,'bottom-tray',0,lower+.008,mid,width,.016,depth,'#354756');
  box(path,'rear-wall',0,lower+.015,rear,width,height-.02,.020,'#344c5d');
  for(const sign of [-1,1]){
   box(path,'side-wall-'+sign,sign*width*.49,lower+.012,mid,.018,height-.02,depth,'#657684');
   box(path,'mounting-ear-'+sign,sign*w*.395,lower+.004,front,w*.07,height-.008,.030,'#a7b9c1');
   box(path,'fastener-'+sign,sign*w*.395,lower+height*.43,front-.020,.029,.028,.012,'#24313b');
   box(path,'cable-clip-'+sign,sign*width*.37,lower+height*.18,rear-depth*.23,.040,.029,.055,'#b4c4c7');
  }
  box(path,'rear-entry-gland',-width*.29,lower+height*.21,rear+.024,.082,.067,.070,'#586c76');
  box(path,'front-bezel',0,lower+.009,front+.020,width*.94,height-.015,.040,'#203947');
  box(path,'port-label-strip',0,lower+height*.77,front-.008,width*.88,.018,.014,'#dde5e6');
  // Twelve duplex adapters. The original LC cord still enters ports 4 and 5.
  const portY=lower+Math.max(.045,unit*.37);
  for(let i=0;i<12;i++){
   const x=-w*.27+i*w*.049;
   box(path,'blue-lc-duplex-'+i,x,portY-.024,front-.010,w*.039,.058,.042,'#1788b4');
   for(const side of [-1,1]){
    box(path,'lc-socket-'+i+'-'+side,x+side*w*.009,portY-.003,front-.038,w*.012,.024,.011,'#102c3b');
    box(path,'lc-retainer-'+i+'-'+side,x+side*w*.009,portY+.023,front-.037,w*.013,.008,.014,'#ccdee3');
   }
  }
  // Splice area and slack loops are enclosed by the cover while idle.
  box(path,'splice-tray',0,lower+height*.15,mid,width*.42,.020,depth*.37,'#b0bfc1');
  for(let i=0;i<6;i++){
   box(path,'protection-sleeve-'+i,(i-2.5)*width*.055,lower+height*.22,mid+depth*.07,
    width*.037,.014,depth*.15,'#e4ddd0');
  }
  const palette=[0xe9bf36,0x5faad8,0xd891ad,0x7acfa5,0xe5a165,0xb3a3d8];
  for(let i=0;i<6;i++){
   const x=-w*.27+(i+3)*w*.049;
   const y=lower+height*.34;
   const curve=new THREE.CatmullRomCurve3([
    new THREE.Vector3(x,y,front+.08),
    new THREE.Vector3(x*.78,y,mid-depth*.16),
    new THREE.Vector3(x*.56,y,mid+depth*.20),
    new THREE.Vector3(-x*.52,y,mid+depth*.26),
    new THREE.Vector3((i-2.5)*width*.055,y,mid+depth*.11)
   ],false,'centripetal');
   const fiber=new THREE.Mesh(new THREE.TubeGeometry(curve,24,.005,4,false),
    new THREE.MeshStandardMaterial({color:palette[i],roughness:.63,metalness:.02}));
   fiber.name='rack-optical-patch:pigtail-'+i;fiber.userData.assetId=id;
   fiber.userData.partId=part.id;path.add(fiber);
  }
  // Side-pivot removable top cover swings in the horizontal plane. It stays below
  // the ToR occupying the next U instead of clipping vertically through it.
  const cover=new THREE.Group();cover.name='rack-optical-patch:top-cover-hinge';
  cover.position.set(-width*.5,lower+height,rear);path.add(cover);
  box(cover,'opaque-lid',width*.5,0,-depth*.5,width,.018,depth,'#506474');
  box(cover,'front-cover-lip',width*.5,-height*.12,-depth,width,.032,.022,'#718594');
  box(path,'hinge-pin',-width*.5,lower+height-.012,rear,.036,.032,.040,'#bdcbd0');
  part.description='Network Rack의 19인치 1U 수동 광 패치패널입니다. 전면 LC duplex 12포트, 랙 고정 브래킷, 상단 힌지 커버와 내부 다색 피그테일·융착 트레이·보호슬리브를 구분해 표시합니다. 기존 케이블 연결/분리 버튼은 4·5번 LC 포트에 그대로 작동합니다. 교육용 형상입니다.';
  const s={cover,target:0,angle:0,partId:part.id};sessions.set(id,s);return s;
 }
 function select(id,active){const s=sessions.get(id);if(!s)return false;s.target=active?1:0;return true;}
 function update(id,dt){const s=sessions.get(id);if(!s)return null;
  const alpha=1-Math.exp(-Math.min(.1,Math.max(0,dt))*5);
  s.angle+=(s.target-s.angle)*alpha;s.cover.rotation.y=s.angle*1.52;
  return {open:s.angle,target:s.target};
 }
 function state(id){const s=sessions.get(id);return s?{open:s.angle,target:s.target,partId:s.partId}:null;}
 return {decorate,select,update,state,drop:id=>sessions.delete(id),reset:()=>sessions.clear()};
}
