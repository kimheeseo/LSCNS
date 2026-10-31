// STEP 4: inspectable small fiber termination box; separate from the 1U patch panel
// and the MPO cassette carrier. All objects belong to the rack scene lifecycle.
export function createRackSmallFiberBox(THREE){
 const sessions=new Map();
 const palette=[0xf2d348,0x5baeee,0xd781ac,0x79c7a2,0xeaa873,0xb8a4e5];
 function decorate(asset,built){
  if(asset.type!=='Top-of-Rack switch')return null;
  const {id,w,d,h}=asset,unit=(h-.46)/built.rackUnits,uStart=built.rackUnits-9;
  const bottom=.23+(uStart-1)*unit,boxHeight=unit*1.78,width=w*.59;
  const front=-d*.515,rear=front+d*.35,depth=rear-front,mid=(front+rear)/2;
  const root=new THREE.Group();root.name='small-fiber-box:assembly';built.root.add(root);
  const create=(key,kind,label,description,extra={})=>{
   const p={id:id+':fiber-box-'+key,assetId:id,kind,label,description,uStart,uHeight:2,...extra};
   built.parts.push(p);built.partMap.set(p.id,p);return p;
  };
  const housing=create('housing','small-fiber-box','ODF 소형 광단자함 · 6 LC',
   '검은색 소형 광단자함입니다. 측면 LC 어댑터 6개, 전면 힌지 커버, 피그테일 여장 루프, 중앙 splice holder·융착 보호슬리브, 인입 케이블 고정 클립을 구분해 볼 수 있습니다. MMR/ODF 배선 교육용 예시이며 실제 제품 치수·최소 굴곡반경을 뜻하지 않습니다.');
  root.userData.rackPart=housing;
  const coverPart=create('cover','fiber-box-cover','소형 광단자함 · 힌지 커버',
   '검은 커버의 왼쪽 세로 힌지가 회전하며 전면·상단을 함께 엽니다. 닫으면 내부 배선이 가려지고 열면 광섬유 루프와 융착접속부가 드러납니다.');
  const holder=create('splice-holder','fiber-box-splice-holder','Splice holder · 융착 보호슬리브',
   '중앙의 융착접속 보호슬리브 6개를 고정하는 홀더입니다. 외부 인입 광섬유와 LC 피그테일의 접속부를 보호합니다. 본 모델은 교육용 위치 표시이며 현장 융착 품질 측정을 대체하지 않습니다.');
  const entry=create('cable-entry','fiber-box-cable-entry','케이블 인입부 · 고정 클립/타이',
   '후면 인입 케이블의 외피를 고정하고 여장 루프의 장력을 완화하는 구조입니다. 타이와 가이드는 광섬유를 과도하게 압착하지 않도록 사용합니다.');
  const boxGeometry=new THREE.BoxGeometry(1,1,1),materials=new Map();
  function block(parent,name,x,y,z,sx,sy,sz,color,p=housing){
   let mat=materials.get(color);
   if(!mat){mat=new THREE.MeshStandardMaterial({color,roughness:.57,metalness:.28});materials.set(color,mat);}
   const m=new THREE.Mesh(boxGeometry,mat);m.name='small-fiber-box:'+name;
   m.position.set(x,y,z);m.scale.set(sx,sy,sz);
   m.userData.assetId=id;m.userData.partId=p.id;parent.add(m);return m;
  }
  // Thin-walled case: front and top openings are covered by the hinged L-shaped lid.
  block(root,'floor',0,bottom+.013,mid,width,.026,depth,'#161e28');
  block(root,'back-wall',0,bottom+boxHeight*.5,rear,width,boxHeight,.027,'#151c26');
  for(const side of [-1,1]){
   block(root,'side-wall-'+side,side*width*.49,bottom+boxHeight*.5,mid,.024,boxHeight,depth,'#1a222d');
   block(root,'mounting-bracket-'+side,side*w*.392,bottom+boxHeight*.48,rear-.025,w*.076,boxHeight*.66,.034,'#72838b');
   block(root,'bracket-fastener-'+side,side*w*.392,bottom+boxHeight*.48,rear-.049,.033,.033,.012,'#cad2d4');
  }
  block(root,'lower-lip',0,bottom+.022,front,width,.044,.025,'#2d3740');
  const lid=new THREE.Group();lid.name='small-fiber-box:hinged-cover';
  lid.position.set(-width*.5,bottom+boxHeight*.5,front-.032);root.add(lid);
  block(lid,'opaque-front',width*.5,0,0,width,boxHeight,.032,'#111820',coverPart);
  block(lid,'opaque-top',width*.5,boxHeight*.5,depth*.5,width,.026,depth,'#161e27',coverPart);
  block(lid,'cover-latch',width*.93,-boxHeight*.17,-.022,width*.074,.036,.018,'#9aadb4',coverPart);
  for(const y of [-boxHeight*.29,boxHeight*.29])
   block(root,'hinge-pin',-width*.505,bottom+boxHeight*.5+y,front-.035,.036,boxHeight*.13,.040,'#8799a5',coverPart);
  // Six blue LC adapters on the right side face outwards (+X), unlike front ports.
  const ports=[],paths=[];
  for(let n=0;n<6;n++){
   const z=front+depth*(n+.75)/7.5,yy=bottom+boxHeight*.43;
   const port=create('lc-'+n,'fiber-box-lc-port','소형 광단자함 · 측면 LC 어댑터 '+(n+1),
    '측면 수동 LC 어댑터 '+(n+1)+'번입니다. 해당 내부 피그테일이 중앙 splice holder의 '+(n+1)+'번 보호슬리브와 이어집니다. 광파워를 증폭하지 않습니다.',{portIndex:n});
   ports.push(port);
   block(root,'blue-lc-adapter-'+n,width*.5+.016,yy,z,.054,boxHeight*.36,depth*.079,'#178dc7',port);
   block(root,'lc-socket-'+n,width*.5+.052,yy,z,.015,boxHeight*.17,depth*.043,'#0d2635',port);
   const points=[
    new THREE.Vector3(width*.43,yy,z),
    new THREE.Vector3(width*.28,bottom+boxHeight*.39,z-.035),
    new THREE.Vector3(width*.18,bottom+boxHeight*.36,mid-depth*.20),
    new THREE.Vector3(-width*.17,bottom+boxHeight*.36,mid-depth*.22),
    new THREE.Vector3(-width*.29,bottom+boxHeight*.37,mid),
    new THREE.Vector3(-width*.17,bottom+boxHeight*.39,mid+depth*.20),
    new THREE.Vector3((n-2.5)*width*.050,bottom+boxHeight*.37,mid+depth*.10)
   ];
   const curve=new THREE.CatmullRomCurve3(points,false,'centripetal');
   const fiber=new THREE.Mesh(new THREE.TubeGeometry(curve,28,.005,5,false),
    new THREE.MeshStandardMaterial({color:palette[n],roughness:.67,emissive:0x000000}));
   fiber.name='small-fiber-box:pigtail-'+n;fiber.userData.assetId=id;fiber.userData.partId=port.id;
   fiber.userData.skipPartEmphasis=true;root.add(fiber);paths.push(fiber);
  }
  // Central splice holder and individual white protection sleeves.
  block(root,'splice-base',0,bottom+boxHeight*.24,mid+depth*.11,width*.35,.022,depth*.26,'#8b9da1',holder);
  for(let n=0;n<6;n++){
   block(root,'splice-sleeve-'+n,(n-2.5)*width*.050,bottom+boxHeight*.34,mid+depth*.11,width*.032,.019,depth*.16,'#e5e2d5',holder);
   block(root,'sleeve-clip-'+n,(n-2.5)*width*.050,bottom+boxHeight*.31,mid+depth*.02,width*.036,.018,.024,'#637982',holder);
  }
  // Cable entry, strain relief, fiber guides and two non-crushing retention ties.
  block(root,'entry-gland',-width*.28,bottom+boxHeight*.48,rear+.025,.095,.077,.077,'#4b5964',entry);
  const incoming=new THREE.CatmullRomCurve3([
   new THREE.Vector3(-width*.28,bottom+boxHeight*.48,rear+.18),
   new THREE.Vector3(-width*.28,bottom+boxHeight*.47,rear+.06),
   new THREE.Vector3(-width*.36,bottom+boxHeight*.39,mid+depth*.22),
   new THREE.Vector3(-width*.23,bottom+boxHeight*.37,mid+depth*.12)
  ],false,'centripetal');
  const cable=new THREE.Mesh(new THREE.TubeGeometry(incoming,20,.017,6,false),new THREE.MeshStandardMaterial({color:'#e3cc54'}));
  cable.name='small-fiber-box:incoming-cable';cable.userData.assetId=id;cable.userData.partId=entry.id;root.add(cable);
  for(const side of [-1,1]){
   block(root,'fiber-guide-'+side,side*width*.34,bottom+boxHeight*.25,mid,.044,.040,depth*.13,'#889ca1',entry);
   block(root,'retention-tie-'+side,side*width*.34,bottom+boxHeight*.44,mid,.056,.012,depth*.14,'#c0cacf',entry);
  }
  const s={root,lid,ports,paths,holder,housing,targetCover:0,cover:0,selected:-1,highlight:false,uStart};
  sessions.set(id,s);return s;
 }
 function setHighlight(s){s.paths.forEach((p,i)=>{
  const active=s.highlight&&i===s.selected;p.material.emissive.setHex(active?0xe8a936:0);p.material.emissiveIntensity=active?1.15:0;
 });}
 function selectPort(id,n){const s=sessions.get(id);if(!s)return null;
  s.selected=((n%6)+6)%6;s.highlight=true;s.targetCover=1;setHighlight(s);return s.ports[s.selected];}
 function selectHolder(id){const s=sessions.get(id);if(!s)return null;s.targetCover=1;return s.holder;}
 function togglePath(id){const s=sessions.get(id);if(!s)return false;
  if(s.selected<0)s.selected=0;s.targetCover=1;s.highlight=!s.highlight;setHighlight(s);return s.highlight;}
 function toggleCover(id){const s=sessions.get(id);if(!s)return false;
  s.targetCover=1-s.targetCover;return s.targetCover;}
 function update(id,dt){const s=sessions.get(id);if(!s)return null;
  const t=1-Math.exp(-Math.min(.1,Math.max(0,dt))*5);
  s.cover+=(s.targetCover-s.cover)*t;s.lid.rotation.y=s.cover*1.65;
  return state(id);
 }
 function state(id){const s=sessions.get(id);return s?{targetCover:s.targetCover,cover:s.cover,selected:s.selected,highlight:s.highlight,portCount:6,partId:s.housing.id,holderId:s.holder.id,uStart:s.uStart}:null;}
 return {decorate,toggleCover,selectPort,selectHolder,togglePath,update,state,drop:id=>sessions.delete(id),reset:()=>sessions.clear()};
}
