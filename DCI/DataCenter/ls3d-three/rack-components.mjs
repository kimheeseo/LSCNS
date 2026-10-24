// STEP 10 — independent, selectable rack frame and assemblies, 42/48U conceptual layout.
// This is a physical hierarchy for teaching. Not a manufacturer-certified U allocation/BOM.
// Allocate only unused front U positions. Passive panels are illustrative and
// never change active server/NIC counts or the BOM engine.
export function planRackFrontFill(rackUnits,occupied,network=false){
 const used=Array(rackUnits+1).fill(false),added=[];
 function reserve(p){
  for(let u=p.uStart;u<p.uStart+p.uHeight;u++){
   if(u<1||u>rackUnits||used[u])throw new Error('Overlapping or out-of-range rack unit '+u);
   used[u]=true;
  }
 }
 occupied.forEach(reserve);
 if(network){
  const lastActive=Math.max(2,...occupied.filter(p=>/^(server|storage|network-switch)-/.test(p.kind)).map(p=>p.uStart+p.uHeight-1));
  let serial=0;
  for(let start=lastActive+1;start+2<=rackUnits-5;start+=5){
   if([start,start+1,start+2].some(u=>used[u]))continue;
   const panel={kind:'distribution-panel',uStart:start,uHeight:2,serial:++serial};
   const manager={kind:'horizontal-manager',uStart:start+2,uHeight:1,serial};
   reserve(panel);reserve(manager);added.push(panel,manager);
  }
 }
 for(let u=1;u<=rackUnits;){
  if(used[u]){u++;continue;}
  let units=1;while(units<3&&u+units<=rackUnits&&!used[u+units])units++;
  const panel={kind:'blanking-panel',uStart:u,uHeight:units};reserve(panel);added.push(panel);u+=units;
 }
 return {added,coverage:used.slice(1),occupiedUnits:used.slice(1).filter(Boolean).length};
}

export function createRackComponentKit(THREE,{cube}){
 const stats={built:0,parts:0,units42:0,units48:0,byType:{}};
 function box(parent,label,x,y,z,w,h,d,color,metal=.55,rough=.39){
  const m=cube(parent,x,y,z,w,h,d,color,metal,rough);
  m.name='rack-shape:'+label;return m;
 }
 function group(parent,assetId,kind,label,uStart,uHeight){
  const g=new THREE.Group();
  g.name='rack-part:'+kind+':'+assetId;
  g.userData.partId=assetId+':'+kind+':'+uStart;
  g.userData.rackPart={id:g.userData.partId,assetId,kind,label,uStart,uHeight};
  parent.add(g);stats.parts++;stats.byType[kind]=(stats.byType[kind]||0)+1;
  return g;
 }
 function attach(g,assetId){
  const part=g.userData.rackPart;
  g.traverse(o=>{if(o.isMesh){o.userData.assetId=assetId;o.userData.partId=part.id;}});
 }
 function build(root,a){
  const id=a.id,w=a.w||3,d=a.d||3,h=a.h||6;
  const count=/gpu|tor-a|tor-b|server rack/i.test(a.type+' '+a.id)?48:42;
  const unit=(h-.46)/count,lower=.23;
  const pieces=[],map=new Map();
  function retain(g){const p=g.userData.rackPart;if(p){pieces.push(p);map.set(p.id,p);attach(g,id)}return g}
  function part(kind,label,start,units,width,height,depth,z,col){
    const pp=group(root,id,kind,label,start,units);
    const y=lower+(start-1)*unit;
    box(pp,label+'-body',0,y,z,width,height||Math.max(.09,units*unit-.02),depth||d*.67,col);
    return retain(pp);
  }
  // Structural frame and independent uprights; no manufacturer dimension certification.
  const shell=group(root,id,'frame','랙 프레임',1,count);
  for(const sx of [-1,1])for(const sz of [-1,1]){
    box(shell,'upright-'+sx+'-'+sz,sx*w*.46,.06,sz*d*.44,w*.065,h-.12,d*.065,'#829db1',.74,.28);
  }
  box(shell,'top-plate',0,h-.13,0,w,.14,d,'#9db2c3',.66,.32);
  box(shell,'bottom-plate',0,.04,0,w,.14,d,'#65839b');
  box(shell,'right-side',w*.48,.14,0,w*.038,h-.30,d*.78,'#3c5b75');
  box(shell,'left-side',-w*.48,.14,0,w*.038,h-.30,d*.78,'#3c5b75');
  retain(shell);
  // Each rack unit is proportional to the existing educational scene h.
  const unitGroup=group(root,id,'u-rails','42U/48U 마운팅 레일',1,count);
  for(const xx of [-w*.395,w*.395])
    box(unitGroup,'u-mounting-rail-'+xx,xx,.24,-d*.412,.085,h-.51,.105,'#c1d0d5',.74,.27);
  // Repeated mounting apertures expose the U pitch and actual attachment points.
  for(const xx of [-w*.395,w*.395])for(let u=0;u<count;u++){
    box(unitGroup,'mounting-aperture-'+u,xx,lower+u*unit+unit*.34,-d*.470,.035,.026,.009,'#182a39',.1,.7);
  }
  retain(unitGroup);
  const pickSlots=[
    // STEP 14: reserve the physically highest U positions for the ToR.
    {kind:'tor',label:'Top-of-Rack 네트워크 스위치',top:count,units:2,color:'#36718a'},
    {kind:'patch-panel',label:'광 패치패널',top:count-2,units:1,color:'#be9e68'},
    {kind:'odf',label:'ODF 및 광 카세트',top:count-3,units:2,color:'#537e94'}
  ];
  const network=/core|spine|leaf|odf|mmr|switch/i.test(a.type);
  const storage=/storage/i.test(a.type);
  const isGpu=/gpu/i.test(a.type);
  const serverType=network?'network-switch':storage?'storage':'server';
  for(const slot of pickSlots){
    const start=Math.max(1,slot.top-slot.units+1);
    if(slot.kind==='odf')continue; // Open, four-slot chassis is supplied by rack-optical-cassette.
    const g=part(slot.kind,slot.label,start,slot.units,w*.76,null,d*.46,-d*.29,slot.color);
    const yh=(slot.units*unit)*.5;
    const face=box(g,'port-face',0,lower+(start-1)*unit+.015,-d*.555,w*.68,
      Math.max(.055,slot.units*unit-.04),.09,'#1a364a',.42,.38);
    for(let i=0;i<12;i++){
      box(g,'port-'+i,-w*.27+i*w*.049,lower+(start-1)*unit+Math.max(.045,slot.units*unit*.37),
        -d*.613,w*.033,.05,.022,i%2?'#72d3e6':'#69e5b9',.12,.24);
    }
    if(slot.kind==='odf'){
      const cassette=group(root,id,'fiber-cassette','광 카세트',start,slot.units);
      box(cassette,'cassette-body',0,lower+(start-1)*unit+.08,-d*.39,w*.62,.17,d*.18,'#d9b270');
      retain(cassette);
    }
  }
  const slide=new THREE.Group();slide.name='server-tray';root.add(slide);
  // Server sled groups are independent pick targets and move together with original slide state.
  const maxServers=network?3:storage?7:isGpu?7:6;
  const modules=[],reserved=9;
  for(let i=0;i<maxServers;i++){
    const u=3+i*4;
    if(u+3>=count-reserved)break;
    const units=storage?3:4,y=lower+(u-1)*unit;
    const kind=serverType,label=(network?'네트워크 모듈':storage?'스토리지 서버':'GPU/Compute 서버')+' '+(i+1);
    const partGroup=group(slide,id,kind+'-'+(i+1),label,u,units);
    box(partGroup,'sled'+i,0,y,-d*.15,w*.72,units*unit-.035,d*.57,
      network?'#304e6d':storage?'#587a82':'#304659',.63,.33);
    box(partGroup,'server-bezel'+i,0,y+.015,-d*.47,w*.67,Math.max(.08,units*unit-.09),.08,
      '#1d354e',.49,.37);
    for(let j=0;j<8;j++){
      box(partGroup,'server-port-'+i+'-'+j,-w*.26+j*w*.075,y+units*unit*.38,
        -d*.52,.057,.075,.027,j%2?'#68d8e9':'#63eaba',.2,.26);
    }
    for(const sign of [-1,1]){
      box(partGroup,'server-mount-ear',sign*w*.373,y,-d*.48,w*.05,units*unit-.035,.038,'#92a9b5');
      box(partGroup,'server-pull-handle',sign*w*.304,y+unit*.6,-d*.54,w*.028,unit*1.7,.05,'#acbdc7');
    }
    for(let bay=0;bay<4;bay++){
      box(partGroup,'hot-swap-drive-bay-'+bay,-w*.17+bay*w*.105,y+unit*.45,-d*.52,w*.09,unit*1.9,.025,'#536c7d');
      box(partGroup,'drive-release-tab-'+bay,-w*.17+bay*w*.105,y+unit*.65,-d*.54,w*.046,.022,.015,'#b4c1c7');
    }
    retain(partGroup);modules.push(partGroup);
  }
  // Account for active assemblies and the top optical drawer before filling gaps.
  const occupied=pieces.filter(p=>p.kind==='tor'||p.kind==='patch-panel'||/^(server|storage|network-switch)-/.test(p.kind));
  occupied.push({kind:'optical-drawer',uStart:count-4,uHeight:2});
  const frontFill=planRackFrontFill(count,occupied,network);
  function batch(g,name,points,ww,hh,dd,color){
    // Instanced sockets/fingers keep repeated passive patch panels inexpensive.
    if(THREE.InstancedMesh&&THREE.Matrix4&&THREE.Vector3&&THREE.Quaternion){
      const sample=box(g,name+'-sample',0,0,0,ww,hh,dd,color);
      g.remove(sample);
      const mesh=new THREE.InstancedMesh(sample.geometry,sample.material,points.length);
      const q=new THREE.Quaternion(),scale=new THREE.Vector3(ww,hh,dd);
      points.forEach(([x,y,z],i)=>mesh.setMatrixAt(i,new THREE.Matrix4().compose(new THREE.Vector3(x,y+hh*.5,z),q,scale)));
      mesh.name='rack-shape:'+name;mesh.instanceMatrix.needsUpdate=true;g.add(mesh);return mesh;
    }
    points.forEach(([x,y,z],i)=>box(g,name+'-'+i,x,y,z,ww,hh,dd,color));
  }
  for(const fill of frontFill.added){
    const y=lower+(fill.uStart-1)*unit,hh=fill.uHeight*unit-.018;
    const label=fill.kind==='distribution-panel'?'광패치패널 '+fill.serial+' · LC duplex × 12 (예시)':fill.kind==='horizontal-manager'?'수평 케이블 정리대 '+fill.serial:fill.uHeight+'U 블랭킹 패널 · 예비 공간';
    const g=group(root,id,fill.kind+'-'+fill.uStart,label,fill.uStart,fill.uHeight);
    g.userData.rackPart.description=fill.kind==='distribution-panel'?'중앙 광배선 구역의 수동형 패치패널입니다. 12개 LC duplex 아답터와 모듈 하우징을 표시한 교육용 배치이며 실제 접속·용량·BOM 수량을 추가한 결과가 아닙니다.':fill.kind==='horizontal-manager'?'광패치코드와 케이블을 수평으로 정리하는 핑거형 관리 부품입니다.': '미사용 U 공간을 막아 빈 슬롯을 구별하는 마감 패널입니다. 서버나 광모듈이 들어 있는 부품이 아닙니다.';
    if(fill.kind==='distribution-panel'){
      box(g,'distribution-housing',0,y,-d*.24,w*.76,hh,d*.52,'#405c70');
      box(g,'distribution-front',0,y+.012,-d*.525,w*.75,hh-.025,.055,'#8ba2b0');
      const row=Array.from({length:12},(_,i)=>[-w*.305+i*w*.0555,y+hh*.24,-d*.57]);
      batch(g,'lc-duplex-adapter',row,w*.045,hh*.47,.045,'#2689b2');
      const sockets=row.flatMap(([x,yy,z])=>[-1,1].map(side=>[x+side*w*.010,yy+hh*.09,z-.027]));
      batch(g,'lc-duplex-socket',sockets,w*.012,hh*.23,.012,'#102432');
      batch(g,'module-divider',[-1,0,1].map(i=>[i*w*.19,y+.02,-d*.59]),.015,hh*.78,.015,'#cad8dd');
      box(g,'patch-label-strip',0,y+hh*.79,-d*.574,w*.62,.019,.009,'#dce5e7');
    }else if(fill.kind==='horizontal-manager'){
      box(g,'manager-back',0,y,-d*.48,w*.76,hh,.08,'#2c4558');
      batch(g,'manager-fingers',Array.from({length:10},(_,i)=>[-w*.32+i*w*.071,y+.018,-d*.54]),w*.021,hh*.70,.11,'#8aa0ad');
      box(g,'manager-lower-lip',0,y,-d*.59,w*.75,.018,.025,'#b4c2c9');
    }else{
      box(g,'blanking-plate',0,y,-d*.50,w*.76,hh,.035,'#344c5f');
      box(g,'blanking-top-edge',0,y+hh-.014,-d*.523,w*.74,.011,.01,'#6c8394');
      batch(g,'blanking-fasteners',[-1,1].map(side=>[side*w*.35,y+hh*.45,-d*.526]),.020,.020,.012,'#b1bec7');
    }
    retain(g);
  }
  const cassette=new THREE.Group();cassette.name='fiber-cassette';root.add(cassette);
  const cassetteParts=group(cassette,id,'cassette-pull','인출형 광카세트',count-4,2);
  // Drawer is an empty carrier; optical kit builds the housing and all modules.
  retain(cassetteParts);
  for(const side of [-1,1]){
    const pdu=group(root,id,side<0?'pdu-a':'pdu-b',side<0?'후면 Rack PDU A':'후면 Rack PDU B',1,count);
    box(pdu,'pdu-case',side*w*.40,.26,d*.42,w*.075,h*.82,.12,'#c7a96c',.51,.42);
    for(let i=0;i<11;i++){
      box(pdu,'outlet-'+i,side*w*.40,.38+i*(h*.70/11),d*.495,w*.043,.11,.045,
        i%2?'#2b475c':'#2b7568',.2,.46);
    }
    retain(pdu);
  }
  const management=group(root,id,'cable-manager','수직 케이블 관리자',1,count);
  box(management,'duct',w*.34,.29,-d*.40,w*.085,h*.72,.14,'#708a9d',.57,.44);
  retain(management);
  const frontDoor=new THREE.Group();frontDoor.name='front-door-hinge';
  frontDoor.position.set(-w*.5,.06,-d*.56);root.add(frontDoor);
  const frontMesh=group(frontDoor,id,'front-door','전면 랙 도어',1,count);
  box(frontMesh,'door-frame',w*.5,0,0,w*.95,h-.13,.075,'#516e86',.59,.44);
  // Perforations/visible strips hint mesh door. Do not equate to real airflow resistance.
  for(let i=0;i<15;i++)box(frontMesh,'perforation-'+i,w*.5,
    .30+i*(h-.75)/15,-.055,w*.75,.046,.034,'#182e43',.19,.68);
  retain(frontMesh);
  const rearDoor=new THREE.Group();rearDoor.name='rear-door-hinge';
  rearDoor.position.set(w*.5,.06,d*.55);root.add(rearDoor);
  const rearPart=group(rearDoor,id,'rear-door','후면 랙 도어',1,count);
  box(rearPart,'rear-door-leaf',-w*.5,0,0,w*.90,h-.13,.085,'#466078',.57,.46);
  retain(rearPart);
  // Keep STEP 03 controls: returning exactly the original moving mechanical API.
  const moving={hinge:frontDoor,slide,cassette,reach:d*.67,door:null,rearDoor};
  stats.built++;if(count===48)stats.units48++;else stats.units42++;
  return {moving,parts:pieces,partMap:map,rackUnits:count,serverModules:modules.length,
    serverGroups:modules,frontFill,cassetteGroup:cassette,cassetteDrawer:cassetteParts,
    frontDoor,rearDoor,slideGroup:slide};
 }
 return {build,stats};
}

