// A passive 1U LC adapter panel, separate from the patch panel and MPO cassette.
export function createRackLcAdapterDesign(THREE,{cube}){
 const sessions=new Map(),stats={built:0,ports:0,fibers:0};
 function decorate(asset,built,high){
  if(asset.type!=='Top-of-Rack switch')return null;
  const part=built.parts.find(p=>p.kind==='lc-adapter-panel');
  const panel=high.children.find(g=>g.userData?.rackPart?.id===part?.id);
  if(!panel||!part)return null;
  for(const child of panel.children)child.visible=false;
  const {id,w,d,h}=asset,unit=(h-.46)/built.rackUnits;
  const lower=.23+(part.uStart-1)*unit,height=unit*.88,width=w*.76;
  const front=-d*.529,rear=front+d*.31,mid=(front+rear)*.5;
  const root=new THREE.Group();root.name='rack-lc:passive-adapter-panel';panel.add(root);
  const ids=new Set([part.id]),paths=[],portMeshes=[];
  part.description='독립된 1U 수동 LC duplex 어댑터 패널입니다. 전면 LC 커넥터와 후면 피그테일을 물리적으로 이어 주며, 융착접속·광 증폭 기능은 없습니다. 포트·광경로는 교육용 예시입니다.';
  function info(i){
   const p={id:id+':lc-adapter-port-'+i,assetId:id,kind:'lc-adapter-port',portIndex:i,
    label:'LC 어댑터 패널 · duplex 포트 '+(i+1),
    description:'전면 '+(i+1)+'번 수동 LC duplex 접속부입니다. 후면 커넥터와 피그테일 두 가닥으로 이어집니다. 광파워를 증폭하지 않습니다.',
    uStart:part.uStart,uHeight:1};
   ids.add(p.id);built.parts.push(p);built.partMap.set(p.id,p);return p;
  }
  function box(parent,p,name,x,y,z,bw,bh,bd,color){
   const mesh=cube(parent,x,y,z,bw,bh,bd,color,.52,.26);
   mesh.name='rack-lc:'+name;mesh.userData.assetId=id;mesh.userData.partId=p.id;return mesh;
  }
  box(root,part,'chassis',0,lower,mid,width,height*.88,rear-front,'#2e3b43');
  box(root,part,'faceplate',0,lower,front,width*.98,height*.86,.032,'#8b9ba2');
  box(root,part,'port-number-strip',0,lower+height*.70,front-.028,width*.85,.010,.012,'#e2e7e2');
  for(const side of [-1,1]){
   box(root,part,'mounting-ear-'+side,side*w*.398,lower,front+.010,w*.08,height*.86,.035,'#b8c5c8');
   box(root,part,'mounting-screw-'+side,side*w*.398,lower+height*.39,front-.018,.025,.025,.014,'#25343c');
   box(root,part,'rear-cable-guide-'+side,side*width*.40,lower+height*.14,rear+.045,.032,.043,.13,'#d4c9a9');
  }
  const colors=['#f4d15a','#66aee2','#ef8e79','#75c690','#bba3d9','#d7dad7'];
  for(let i=0;i<12;i++){
   const p=info(i),x=(i-5.5)*width*.073,py=lower+height*.22;
   portMeshes.push(box(root,p,'duplex-adapter-'+i,x,py,front-.040,width*.066,height*.53,.059,'#168bbd'));
   box(root,p,'port-number-'+(i+1),x,lower+height*.72,front-.042,width*.035,.008,.010,'#1d4e5c');
   for(let side=0;side<2;side++){
    const sx=x+(side?.014:-.014);
    box(root,p,'front-lc-socket-'+i+'-'+side,sx,py+height*.05,front-.078,.021,height*.28,.014,'#102c3b');
    box(root,p,'rear-lc-connector-'+i+'-'+side,sx,py+height*.05,rear+.044,.022,height*.25,.065,'#d7e2df');
    const pts=[
     new THREE.Vector3(sx,py+height*.16,rear+.084),
     new THREE.Vector3(sx+(i%2?.014:-.014),py+height*.32,rear+.15),
     new THREE.Vector3(sx+(i-5.5)*.006,py+height*.47,rear+.27),
     new THREE.Vector3(sx+(i-5.5)*.009,py+height*.20,rear+.39)
    ];
    const fiber=new THREE.Mesh(
     new THREE.TubeGeometry(new THREE.CatmullRomCurve3(pts),16,.0035,4,false),
     new THREE.MeshStandardMaterial({color:colors[i%6],roughness:.64}));
    fiber.name='rack-lc:rear-pigtail-'+i+'-'+side;
    fiber.userData.assetId=id;fiber.userData.partId=p.id;
    root.add(fiber);paths.push({port:i,mesh:fiber});stats.fibers++;
   }
   stats.ports++;
  }
  const s={id,part,high,panel,root,ids,paths,portMeshes,selectedPort:-1,isolated:false,visibility:null,front,rear,lower,height};
  sessions.set(id,s);stats.built++;return s;
 }
 function owns(info){const s=info&&sessions.get(info.assetId);return !!s?.ids.has(info.id);}
 function isolate(id,active){
  const s=sessions.get(id);if(!s)return false;
  if(active&&!s.isolated){s.visibility=s.high.children.map(child=>[child,child.visible]);s.isolated=true;}
  if(!active&&s.isolated){for(const [child,visible] of s.visibility)child.visible=visible;s.visibility=null;s.isolated=false;}
  if(s.isolated)for(const child of s.high.children)child.visible=child===s.panel;
  return true;
 }
 function highlight(id,index){
  const s=sessions.get(id);if(!s)return false;
  s.selectedPort=Number.isInteger(index)&&index>=0&&index<12?index:-1;
  for(const {port,mesh} of s.paths){
   mesh.material.emissive.set(port===s.selectedPort?0xffd259:0x000000);
   mesh.material.emissiveIntensity=port===s.selectedPort?.9:0;
   mesh.material.transparent=s.selectedPort>=0&&port!==s.selectedPort;
   mesh.material.opacity=s.selectedPort>=0&&port!==s.selectedPort?.26:1;
  }
  return true;
 }
 function state(id){const s=sessions.get(id);return s?{portCount:12,fiberCount:24,selectedPort:s.selectedPort,isolated:s.isolated}:null;}
 function drop(id){isolate(id,false);sessions.delete(id);}
 function reset(){for(const id of [...sessions.keys()])drop(id);}
 return {decorate,owns,isolate,highlight,state,drop,reset,stats};
}
