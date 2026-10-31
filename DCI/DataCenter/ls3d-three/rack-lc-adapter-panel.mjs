// STEP 3: an independent passive 1U LC adapter panel, separate from the patch-panel tray.
// Geometry and textures are owned by the rack scene and released by disposeAsset.
export function createRackLcAdapterPanel(THREE){
 const sessions=new Map();
 function decorate(asset,built){
  if(asset.type!=='Top-of-Rack switch')return null;
  const {w,d,h,id}=asset,unit=(h-.46)/built.rackUnits,uStart=built.rackUnits-6;
  const y=.23+(uStart-1)*unit,front=-d*.50,rear=-d*.25;
  const width=w*.75,height=unit*.78;
  const panel=new THREE.Group();panel.name='lc-adapter:1u-passive-panel';built.root.add(panel);
  const part={id:id+':lc-adapter-panel:'+uStart,assetId:id,kind:'lc-adapter-panel',label:'LC duplex 어댑터 패널 · 수동 접속부',uStart,uHeight:1,
   description:'독립형 1U 수동 LC duplex 어댑터 패널입니다. 전면 12개 포트(24심)와 후면 커넥터·피그테일, 번호 및 여장 가이드를 확인합니다. 광 신호를 증폭하거나 융착하지 않으며 패치패널 전체 또는 MPO 카세트가 아닙니다. 포트 구성은 교육용 예시입니다.'};
  panel.userData.rackPart=part;
  const ports=[],paths=[],mats=new Map(),boxGeo=new THREE.BoxGeometry(1,1,1);
  function add(parent,name,x,yy,z,sx,sy,sz,color,p=part){
   const key=color;let mat=mats.get(key);
   if(!mat){mat=new THREE.MeshStandardMaterial({color,metalness:.32,roughness:.53});mats.set(key,mat);}
   const mesh=new THREE.Mesh(boxGeo,mat);mesh.name='lc-adapter:'+name;
   mesh.position.set(x,yy+sy*.5,z);mesh.scale.set(sx,sy,sz);
   mesh.userData.assetId=id;mesh.userData.partId=p.id;parent.add(mesh);return mesh;
  }
  function register(p){built.parts.push(p);built.partMap.set(p.id,p);return p;}
  register(part);
  // An adapter panel is a thin pass-through plate, without a lid or splice tray.
  add(panel,'front-face',0,y,front,width,height,.038,'#344c60');
  add(panel,'rear-rail',0,y,rear,width,.035,.045,'#819dab');
  for(const side of [-1,1]){
   add(panel,'mount-ear-'+side,side*w*.397,y-.01,front,w*.064,height+.02,.044,'#a5bbc5');
   add(panel,'rear-cable-guide-'+side,side*width*.44,y+.02,rear-.10,.034,height*.66,.045,'#91aaba');
  }
  // Canvas text is oriented towards the negative-Z front (no mirrored Korean/number labels).
  if(typeof document!=='undefined'){
   const c=document.createElement('canvas');c.width=1536;c.height=128;
   const ctx=c.getContext('2d');
   if(ctx){
    ctx.fillStyle='#e5eef2';ctx.fillRect(0,0,c.width,c.height);
    ctx.fillStyle='#143044';ctx.font='700 52px Arial,sans-serif';ctx.textAlign='center';
    for(let n=0;n<12;n++)ctx.fillText(String(n+1).padStart(2,'0'),(n+.5)*128,85);
    const texture=new THREE.CanvasTexture(c);texture.colorSpace=THREE.SRGBColorSpace;
    const label=new THREE.Mesh(new THREE.PlaneGeometry(width*.83,height*.19),new THREE.MeshBasicMaterial({map:texture,side:THREE.DoubleSide}));
    label.name='lc-adapter:port-numbers';label.position.set(0,y+height*.83,front-.029);label.rotation.y=Math.PI;
    label.userData.assetId=id;label.userData.partId=part.id;label.userData.ownedTexture=texture;panel.add(label);
   }
  }
  const colors=[0xf1ca3b,0x56bce9,0xe19ab3,0x83d1a4,0xf0a273,0xb5a4e1];
  for(let n=0;n<12;n++){
   const x=(n-5.5)*width*.064,py=y+height*.23;
   const port=register({id:id+':lc-adapter-port-'+n,assetId:id,kind:'lc-adapter-port',label:'LC 어댑터 패널 · 전면 duplex '+(n+1),
    description:'수동 LC duplex '+(n+1)+'번 포트입니다. 전면 LC 커넥터와 후면 접속 커넥터 사이를 광학적으로 통과시키며 피그테일 2심이 후면 여장 가이드로 이어집니다.',uStart,uHeight:1,portIndex:n});
   const back=register({id:id+':lc-adapter-rear-'+n,assetId:id,kind:'lc-adapter-rear-port',label:'LC 어댑터 패널 · 후면 연결 '+(n+1),
    description:'전면 '+(n+1)+'번 LC duplex 포트에 대응하는 후면 2심 커넥터·피그테일 연결입니다. 어댑터는 수동 접속부이며 후면 경로는 다른 패널의 융착 트레이와 구분됩니다.',uStart,uHeight:1,portIndex:n});
   add(panel,'front-port-'+n,x,py,front-.036,width*.054,height*.48,.048,'#168bb5',port);
   for(const side of [-1,1]){
    add(panel,'front-lc-socket-'+n+'-'+side,x+side*width*.012,py+height*.12,front-.065,width*.013,height*.19,.012,'#0b2030',port);
    add(panel,'rear-connector-'+n+'-'+side,x+side*width*.012,py+height*.12,rear+.039,width*.014,height*.20,.088,'#3c91ad',back);
    const start=new THREE.Vector3(x+side*width*.012,py+height*.20,rear+.085);
    const end=new THREE.Vector3(x+side*width*.012,py+height*.22,rear+.27);
    const bend=side*width*.018;
    const curve=new THREE.CatmullRomCurve3([start,
     new THREE.Vector3(start.x+bend,py+height*.38,rear+.14),
     new THREE.Vector3(start.x+bend,py+height*.54,rear+.19),end],false,'centripetal');
    const material=new THREE.MeshStandardMaterial({color:colors[n%colors.length],roughness:.65,emissive:0x000000});
    const fiber=new THREE.Mesh(new THREE.TubeGeometry(curve,18,.004,4,false),material);
    fiber.name='lc-adapter:pigtail-'+n+'-'+side;fiber.userData.assetId=id;fiber.userData.partId=back.id;
    fiber.userData.skipPartEmphasis=true;panel.add(fiber);
    (paths[n]??=[]).push(fiber);
   }
   ports.push({front:port,back,x});
  }
  const session={panel,part,ports,paths,selected:-1,highlight:false,uStart};sessions.set(id,session);
  return session;
 }
 function selectPort(id,n,highlight=true){const s=sessions.get(id);if(!s)return null;
  s.selected=((n%12)+12)%12;s.highlight=highlight;updatePaths(s);return s.ports[s.selected];}
 function updatePaths(s){s.paths.forEach((pair,i)=>pair.forEach(f=>{
  const selected=s.highlight&&i===s.selected;f.material.emissive.setHex(selected?0xf2ab30:0);
  f.material.emissiveIntensity=selected?1.25:0;
 }));}
 function togglePath(id){const s=sessions.get(id);if(!s)return false;
  if(s.selected<0)s.selected=0;s.highlight=!s.highlight;updatePaths(s);return s.highlight;}
 function state(id){const s=sessions.get(id);return s?{partId:s.part.id,portCount:s.ports.length,selected:s.selected,highlight:s.highlight,uStart:s.uStart}:null;}
 return {decorate,selectPort,togglePath,state,drop:id=>sessions.delete(id),reset:()=>sessions.clear()};
}
