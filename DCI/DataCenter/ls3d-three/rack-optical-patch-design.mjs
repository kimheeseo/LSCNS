// 3D depth refresh for the existing Network Rack patch panel; preserves U and LC coordinates.
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
  const subpart=(key,label,description)=>{
   const entry={id:id+':patch-'+key,assetId:id,kind:'patch-internal',parentPartId:part.id,
    label,description,uStart:part.uStart,uHeight:1};
   built.parts.push(entry);built.partMap.set(entry.id,entry);return entry;
  };
  const splice=subpart('splice-holder','패치패널 · 융착접속 트레이와 보호슬리브',
   '트렁크 광섬유와 LC 피그테일의 융착접속부를 보호슬리브 및 홀더에 정리합니다.');
  const guides=subpart('fiber-guides','패치패널 · 광섬유 여장과 고정 클립',
   '다색 광섬유 여장을 완만한 루프로 정리하고 클립·케이블 타이로 인입 케이블을 고정합니다.');
  const ports=Array.from({length:12},(_,i)=>subpart('lc-port-'+(i+1),
   '패치패널 · LC duplex '+(i+1)+'번',
   '전면 LC duplex 수동 어댑터입니다. 내부 피그테일을 융착 트레이의 접속부와 연결해 표시합니다.'));
  function box(parent,name,x,base,z,sx,sy,sz,color,selected=part){
   const mesh=cube(parent,x,base,z,sx,sy,sz,color,.38,.46);
   mesh.name='rack-optical-patch:'+name;mesh.userData.assetId=id;mesh.userData.partId=selected.id;mesh.castShadow=true;mesh.receiveShadow=true;
   return mesh;
  }
  // 19-inch rack tray, mounting ears, side walls and cable gland.
  box(path,'bottom-tray',0,lower+.008,mid,width,.016,depth,'#1d2830');
  box(path,'rear-wall',0,lower+.015,rear,width,height-.02,.020,'#29343b');
  for(const sign of [-1,1]){
   box(path,'side-wall-'+sign,sign*width*.49,lower+.012,mid,.018,height-.02,depth,'#2d3b42');
   box(path,'mounting-ear-'+sign,sign*w*.395,lower+.004,front,w*.07,height-.008,.030,'#a7b9c1');
   box(path,'fastener-'+sign,sign*w*.395,lower+height*.43,front-.020,.029,.028,.012,'#24313b');
   box(path,'cable-clip-'+sign,sign*width*.37,lower+height*.18,rear-depth*.23,.040,.029,.055,'#b4c4c7',guides);
  }
  box(path,'rear-entry-gland',-width*.29,lower+height*.21,rear+.024,.082,.067,.070,'#586c76',guides);
  for(const sign of [-1,1])box(path,'fiber-retainer-'+sign,sign*width*.28,lower+height*.16,mid,
   .025,.026,depth*.11,'#c9c8b8',guides);
  box(path,'cable-tie',-width*.29,lower+height*.37,rear-depth*.12,.013,.012,.075,'#e3e2d9',guides);
  box(path,'front-bezel',0,lower+.009,front+.020,width*.94,height-.015,.040,'#1c2c35');
  box(path,'port-label-strip',0,lower+height*.77,front-.008,width*.88,.018,.014,'#dde5e6');
  // Twelve duplex adapters. The original LC cord still enters ports 4 and 5.
  const portY=lower+Math.max(.045,unit*.37);
  for(let i=0;i<12;i++){
   const x=-w*.27+i*w*.049;
   box(path,'blue-lc-duplex-'+i,x,portY-.024,front-.010,w*.039,.058,.042,'#1788b4',ports[i]);
   for(const side of [-1,1]){
    box(path,'lc-socket-'+i+'-'+side,x+side*w*.009,portY-.003,front-.038,w*.012,.024,.011,'#102c3b',ports[i]);
    box(path,'lc-retainer-'+i+'-'+side,x+side*w*.009,portY+.023,front-.037,w*.013,.008,.014,'#ccdee3',ports[i]);
   }
  }
  // Stamped-metal lips, recessed base and multiple levels establish physical depth.
  box(path,'inset-inner-floor',0,lower+.025,mid,width*.93,.007,depth*.90,'#40535c');
  for(const sign of [-1,1]){
   box(path,'raised-side-rim-'+sign,sign*width*.484,lower+height*.75,mid,.025,.024,depth,'#a6b4bc');
   box(path,'internal-support-rail-'+sign,sign*width*.375,lower+height*.14,mid,.032,.029,depth*.77,'#667e88');
   box(path,'front-latch-'+sign,sign*width*.44,lower+height*.42,front-.021,.045,.032,.026,'#c5ced0');
   for(let k=0;k<3;k++)box(path,'rim-rivet-'+sign+'-'+k,sign*width*.483,
    lower+height*.81,front+depth*(.18+k*.30),.014,.012,.014,'#d7e0e1');
  }
  for(let i=0;i<12;i++){
   const x=-w*.27+i*w*.049;
   box(path,'adapter-upper-lip-'+i,x,portY+.030,front-.030,w*.037,.010,.016,'#bad8e6',ports[i]);
   box(path,'adapter-socket-recess-'+i,x,portY-.017,front-.042,w*.029,.039,.010,'#09283b',ports[i]);
  }
  // Splice area and slack loops are enclosed by the cover while idle.
  box(path,'splice-tray',0,lower+height*.15,mid,width*.44,.029,depth*.39,'#607985',splice);
  box(path,'splice-holder-platform',0,lower+height*.27,mid+depth*.075,width*.40,.018,depth*.30,'#b4c4c6',splice);
  for(const sign of [-1,1])box(path,'raised-splice-wall-'+sign,sign*width*.213,
   lower+height*.29,mid,width*.015,.022,depth*.39,'#d6dfdd',splice);
  for(let i=0;i<5;i++)box(path,'splice-comb-'+i,0,lower+height*.39,
   mid-depth*.10+i*depth*.053,width*.34,.012,.011,'#d0d9d9',splice);
  for(let i=0;i<12;i++){
   box(path,'protection-sleeve-'+i,(i-5.5)*width*.031,lower+height*.22,mid+depth*.07,
    width*.021,.014,depth*.15,'#e4ddd0',splice);
  }
  // Two elevated oval routing guides make the 12 separate pigtails legible.
  for(const ratio of [1,1.12]){
   const guide=new THREE.Mesh(new THREE.TorusGeometry(1,.011,8,80),
    new THREE.MeshStandardMaterial({color:ratio===1?0xc9d4d1:0x7e959f,roughness:.38,metalness:.38}));
   guide.name='rack-optical-patch:raised-guide-'+ratio;
   guide.rotation.x=Math.PI/2;guide.scale.set(width*.23*ratio,depth*.305*ratio,1);
   guide.position.set(0,lower+height*(ratio===1?.34:.29),mid);
   guide.castShadow=true;guide.receiveShadow=true;
   guide.userData.assetId=id;guide.userData.partId=guides.id;path.add(guide);
  }
  for(const side of [-1,1])for(let i=0;i<3;i++)
   box(path,'fiber-comb-'+side+'-'+i,side*width*.305,lower+height*.30,
    front+depth*(.31+i*.17),.023,.023,.055,'#b6c7c9',guides);
  const palette=[0xe9bf36,0x5faad8,0xd891ad,0x7acfa5,0xe5a165,0xb3a3d8,
   0xd7cc65,0x869ce0,0xe2a0a0,0x72c6c8,0xb4b2dd,0xe2d9ad];
  for(let i=0;i<12;i++){
   const x=-w*.27+i*w*.049,y=lower+height*(.48+(i%3)*.015);
   // The oval slack path is visible from above; each route has a gentle bend
   // and ends at its own splice sleeve rather than abruptly crossing the tray.
   const radiusX=width*(.195+i*.0045),radiusZ=depth*(.247+i*.004);
   const phase=(i-5.5)*.067;
   const points=[new THREE.Vector3(x,y,front+.07),new THREE.Vector3(x*.95,y,front+depth*.14)];
   for(let j=0;j<=28;j++){
    const angle=-Math.PI*.5+phase+j*Math.PI*2/28;
    points.push(new THREE.Vector3(radiusX*Math.cos(angle),y,mid+radiusZ*Math.sin(angle)));
   }
   points.push(new THREE.Vector3((i-5.5)*width*.031,y,mid+depth*.11));
   const curve=new THREE.CatmullRomCurve3(points,false,'centripetal');
   const fiber=new THREE.Mesh(new THREE.TubeGeometry(curve,76,.0053,6,false),
    new THREE.MeshStandardMaterial({color:palette[i],roughness:.63,metalness:.02}));
   fiber.name='rack-optical-patch:pigtail-'+i;fiber.userData.assetId=id;
   fiber.userData.partId=ports[i].id;fiber.castShadow=true;path.add(fiber);
  }
  // Full-width rear piano hinge: the opened steel lid stands behind the work tray,
  // rather than appearing as a flat wing to its right.
  const cover=new THREE.Group();cover.name='rack-optical-patch:top-cover-hinge';
  cover.position.set(-width*.5,lower+height,rear);path.add(cover);
  box(cover,'brushed-metal-lid',width*.5,-.018,-depth*.5,width,.026,depth,'#839398');
  box(cover,'pressed-inner-stamping',width*.5,-.021,-depth*.5,width*.87,.008,depth*.80,'#5e747d');
  for(const sign of [-1,1]){
   box(cover,'lid-rib-'+sign,width*(.5+sign*.46),-.040,-depth*.5,.025,.034,depth*.87,'#b4bfc0');
   for(let j=0;j<3;j++)box(cover,'lid-rivet-'+sign+'-'+j,width*(.5+sign*.46),-.026,
    -depth*(.22+j*.28),.018,.011,.018,'#d9dfdf');
  }
  box(cover,'lid-front-return',width*.5,-height*.15,-depth,width,.036,.026,'#c3cfd0');
  box(cover,'front-grip',width*.5,-height*.25,-depth-.022,width*.17,.038,.050,'#283d48');
  box(path,'rear-piano-hinge',0,lower+height-.014,rear,.028,.032,.045,'#b2c0c3');
  for(let i=-5;i<=5;i++)box(path,'hinge-knuckle-'+i,i*width*.083,
   lower+height-.019,rear+.010,width*.047,.024,.040,'#d0d9d9');
  part.description='Network Rack의 19인치 1U 수동 광 패치패널입니다. 선택하면 패널만 분리해 사선 시점에서 보며 후면 힌지 커버가 열립니다. 파란 LC duplex 12포트, 다색 피그테일 여장 루프, 융착 트레이·보호슬리브, 케이블 클립과 인입부를 클릭해 확인할 수 있습니다. 기존 케이블 연결/분리 버튼은 4·5번 LC 포트에 그대로 작동합니다. 첨부 소형 광단자함은 내부 배치 참고용이며 교육용 형상입니다.';
  const s={cover,target:0,angle:0,partId:part.id,root:rackRoot,panel,isolated:false,visibility:null,cableVisible:false};
  sessions.set(id,s);return s;
 }
 function isolate(id,active){
  const s=sessions.get(id);if(!s)return false;
  if(active&&!s.isolated){s.visibility=s.root.children.map(child=>[child,child.visible]);s.isolated=true;}
  if(!active&&s.isolated){for(const [child,visible] of s.visibility)child.visible=visible;s.visibility=null;s.isolated=false;}
  if(s.isolated)for(const child of s.root.children)
   child.visible=child===s.panel||(s.cableVisible&&child.name.startsWith('rack-patch-cord-demo:'));
  return true;
 }
 function includeCable(id,active){const s=sessions.get(id);if(!s)return false;s.cableVisible=!!active;return isolate(id,s.isolated);}
 function owns(info){const s=info&&sessions.get(info.assetId);return !!s&&(info.id===s.partId||info.parentPartId===s.partId);}
 function select(id,active){const s=sessions.get(id);if(!s)return false;s.target=active?1:0;return true;}
 function update(id,dt){const s=sessions.get(id);if(!s)return null;
  const alpha=1-Math.exp(-Math.min(.1,Math.max(0,dt))*5);
  s.angle+=(s.target-s.angle)*alpha;s.cover.rotation.x=s.angle*1.40;
  if(s.isolated)isolate(id,true);
  return {open:s.angle,target:s.target,isolated:s.isolated};
 }
 function state(id){const s=sessions.get(id);return s?{open:s.angle,target:s.target,partId:s.partId,isolated:s.isolated}:null;}
 function drop(id){isolate(id,false);sessions.delete(id);}
 function reset(){for(const id of [...sessions.keys()])drop(id);}
 return {decorate,select,isolate,includeCable,owns,update,state,drop,reset};
}
