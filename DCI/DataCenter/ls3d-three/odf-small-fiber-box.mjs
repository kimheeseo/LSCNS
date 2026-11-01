// Educational 6-port termination box in the ODF rack. This is a distinct
// small enclosure, not a second 19-inch patch panel or an MPO cassette.
export function createOdfSmallFiberBox(THREE,{cube}){
 const sessions=new Map(),stats={built:0,fibers:0,updates:0};
 function decorate(asset,built,high){
  if(asset.id!=='odf')return null;
  const {id,w,d,h}=asset,unit=(h-.46)/built.rackUnits,uStart=29;
  const width=Math.min(1.18,w*.56),height=Math.max(.48,unit*4.65);
  const root=new THREE.Group();root.name='odf-small:enclosure';
  root.position.set(0,.23+(uStart-1)*unit,-d*.465);
  high.add(root);
  const parts=new Set(),paths=[],portMeshes=[];
  function info(key,kind,label,description,extra={}){
   const p={id:id+':odf-small:'+key,assetId:id,kind,label,description,uStart,uHeight:5,...extra};
   built.parts.push(p);built.partMap.set(p.id,p);parts.add(p.id);return p;
  }
  function box(parent,p,name,x,y,z,bw,bh,bd,color){
   const mesh=cube(parent,x,y,z,bw,bh,bd,color,.57,.31);
   mesh.name='odf-small:'+name;mesh.userData.assetId=id;mesh.userData.partId=p.id;
   return mesh;
  }
  const shell=info('housing','odf-small-box','ODF 소형 광단자함 · 6 LC 어댑터',
   '검은 소형 하우징 안에 6색 피그테일 여장 루프, 융착접속 보호슬리브·홀더, 케이블 고정 클립과 측면 LC 어댑터 6개가 있는 교육용 광단자함입니다.');
  const coverPart=info('cover','odf-small-cover','힌지 커버','왼쪽 수직 힌지로 여닫는 불투명 커버입니다. 닫으면 내부 광섬유와 접속부를 가립니다.');
  const splice=info('splice-holder','odf-small-splice','융착접속 홀더 · 보호슬리브',
   '중앙 홀더에 6개의 융착접속 보호슬리브를 고정합니다. 접속 손실 및 실제 굽힘 반경은 제품 사양과 측정값을 확인해야 합니다.');
  const guide=info('fiber-guide','odf-small-guide','피그테일 여장 루프 · 고정 클립',
   '케이블 인입부에서 나온 광섬유를 클립과 타이로 고정하고 완만한 루프로 여장을 확보합니다.');
  const back=box(root,shell,'back-plate',0,0,.058,width,height,.024,'#11191e');
  back.userData.rackPart=shell;
  for(const sign of [-1,1]){
   box(root,shell,'top-bottom-wall-'+sign,0,sign<0?0:height-.027,.009,width,.027,.12,'#222a30');
   box(root,shell,'mount-bracket-'+sign,sign*width*.57,height*.43,.09,width*.13,.10,.036,'#606d72');
   box(root,shell,'bracket-screw-'+sign,sign*width*.57,height*.47,.068,.029,.029,.014,'#b6bdba');
  }
  box(root,shell,'left-wall',-width*.49,.02,.012,.022,height-.04,.115,'#242e34');
  box(root,shell,'right-adapter-rail',width*.48,.02,.014,.030,height-.04,.107,'#27323a');
  box(root,guide,'entry-gland',-width*.52,height*.69,.017,.081,.094,.12,'#606a69');
  box(root,guide,'strain-relief',-width*.42,height*.66,-.01,.068,.10,.033,'#181d21');
  for(const [i,x,y] of [[0,-.31,.72],[1,-.27,.25],[2,.15,.16]]){
   box(root,guide,'fiber-clip-'+i,x*width,y*height,-.014,.053,.025,.028,'#c4c8ba');
  }
  box(root,guide,'cable-tie',-width*.28,height*.64,-.032,.018,.095,.019,'#d3d0b8');
  box(root,splice,'splice-holder',0,height*.34,-.016,width*.38,height*.30,.037,'#454d50');
  for(let i=0;i<6;i++){
   box(root,splice,'splice-sleeve-'+i,0,height*(.365+i*.037),-.047,width*.28,.012,.014,
    i%2?'#ccd5d1':'#e9d9ae');
  }
  const colors=['#f4d148','#60bbdf','#ef8e72','#67bf8c','#ba9ed2','#e4e2cd'];
  const colorNames=['노란색','파란색','주황색','초록색','보라색','흰색'];
  for(let i=0;i<6;i++){
   const p=info('lc-port-'+i,'odf-small-port','LC 어댑터 '+(i+1)+' / 내부 피그테일',
    '수동 LC 어댑터 '+(i+1)+'번입니다. 대응 피그테일은 '+colorNames[i]+' 경로로 표시되며 광 증폭 기능은 없습니다.',{portIndex:i});
   const py=height*(.15+i*.135);
   portMeshes.push(box(root,p,'lc-adapter-'+i,width*.52,py,-.018,width*.16,height*.095,.10,'#1886ba'));
   box(root,p,'lc-socket-'+i,width*.62,py,-.018,.014,height*.052,.055,'#103445');
   // Shape the slack into a smooth, nearly full loop before the splice holder.
   const cx=-width*.095,cy=height*.49,rx=width*(.275-i*.004),ry=height*(.395-i*.007);
   const pts=[new THREE.Vector3(-width*.44,height*.68,-.047)];
   for(let k=0;k<=16;k++){
    const t=2.45+k*(Math.PI*1.82)/16;
    pts.push(new THREE.Vector3(cx+rx*Math.cos(t),cy+ry*Math.sin(t),-.050-i*.001));
   }
   pts.push(new THREE.Vector3(-width*.12,height*(.38+i*.035),-.054));
   pts.push(new THREE.Vector3(width*.10,height*(.38+i*.035),-.052));
   pts.push(new THREE.Vector3(width*.36,py,-.047));
   const curve=new THREE.CatmullRomCurve3(pts);
   const fiber=new THREE.Mesh(new THREE.TubeGeometry(curve,48,.0038,5,false),
    new THREE.MeshStandardMaterial({color:colors[i],roughness:.56,metalness:0}));
   fiber.name='odf-small:pigtail-'+i;fiber.userData.assetId=id;fiber.userData.partId=p.id;
   root.add(fiber);paths.push(fiber);stats.fibers++;
  }
  const hinge=box(root,coverPart,'hinge-barrel',-width*.50,.01,-.109,.025,height-.02,.027,'#657078');
  hinge.userData.rackPart=coverPart;
  const cover=new THREE.Group();cover.name='odf-small:hinged-cover';
  cover.position.set(-width*.50,0,-.127);root.add(cover);
  box(cover,coverPart,'opaque-cover',width*.5,0,0,width*.985,height,.038,'#161d22');
  box(cover,coverPart,'cover-edge',width*.95,.025,-.019,.029,height-.05,.030,'#3d484c');
  box(cover,coverPart,'cover-latch',width*.93,height*.52,-.044,.058,.060,.025,'#9ea9a8');
  for(const [i,x] of [[0,width*.23],[1,width*.77]])
   box(cover,coverPart,'cover-fastener-'+i,x,height*.84,-.044,.026,.026,.009,'#7b8587');
  const session={id,root,high,cover,paths,portMeshes,parts,height,open:0,target:0,selectedPort:-1,
   isolated:false,visibility:null};
  sessions.set(id,session);stats.built++;
  return session;
 }
 function owns(info){const s=info&&sessions.get(info.assetId);return !!s?.parts.has(info.id);}
 function isolate(id,active){
  const s=sessions.get(id);if(!s)return false;
  if(active&&!s.isolated){s.visibility=s.high.children.map(child=>[child,child.visible]);s.isolated=true;}
  if(!active&&s.isolated){for(const [child,visible] of s.visibility)child.visible=visible;s.visibility=null;s.isolated=false;}
  if(s.isolated)for(const child of s.high.children)child.visible=child===s.root;
  return true;
 }
 function highlight(id,index){
  const s=sessions.get(id);if(!s)return false;
  s.selectedPort=Number.isInteger(index)&&index>=0&&index<6?index:-1;
  s.paths.forEach((path,i)=>{
   path.material.emissive.set(i===s.selectedPort?0xffcc46:0x000000);
   path.material.emissiveIntensity=i===s.selectedPort?.9:0;
   path.material.transparent=s.selectedPort>=0&&i!==s.selectedPort;
   path.material.opacity=s.selectedPort>=0&&i!==s.selectedPort?.26:1;
  });
  return true;
 }
 function toggle(id){const s=sessions.get(id);if(!s)return false;s.target=s.target?0:1;if(!s.target)highlight(id,-1);return !!s.target;}
 function open(id){const s=sessions.get(id);if(s)s.target=1;return !!s;}
 function update(id,dt=.016){
  const s=sessions.get(id);if(!s)return null;
  s.open+=(s.target-s.open)*Math.min(1,Math.max(0,dt)*7.5);
  if(Math.abs(s.target-s.open)<.001)s.open=s.target;
  s.cover.rotation.y=s.open*2.52;
  stats.updates++;return state(id);
 }
 function state(id){const s=sessions.get(id);return s?{open:s.open,target:s.target,selectedPort:s.selectedPort,isolated:s.isolated,portCount:6}:null;}
 function drop(id){isolate(id,false);sessions.delete(id);}
 function reset(){for(const id of [...sessions.keys()])drop(id);}
 return {decorate,owns,isolate,highlight,toggle,open,update,state,drop,reset,stats};
}
