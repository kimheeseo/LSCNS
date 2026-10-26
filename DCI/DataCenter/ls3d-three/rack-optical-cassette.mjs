// LS structured-cabling teaching model. Reference architecture, not product CAD.
// Four 12F MPO-to-6-LC-duplex cassettes: 48 fibers / 24 duplex adapters in this example.
export function createRackOpticalCassette(THREE,{cube}) {
 const registry=new Map(),stats={built:0,adapters:0,connectors:0,fibers:0,updates:0};
 const source='https://www.lssimple.co.kr/kr/product/prodView.asp?cate1=&cate2=&cate3=238&prd_seq=1213';
 function decorate(asset,built){
  if(!built.cassetteDrawer)return null;
  const {w,d,h,id}=asset,unit=(h-.46)/built.rackUnits;
  const cassetteU=built.cassetteU||built.rackUnits-4;
  const base=.23+(cassetteU-1)*unit,ht=unit*1.80,front=-d*.52,back=-d*.07;
  const tray=new THREE.Group();tray.name='rack-optical:LS-four-slot-tray';built.cassetteDrawer.add(tray);
  const session={id,drawer:built.cassetteDrawer,tray,portDefs:[],modules:[],covers:[],open:0,detail:false,exploded:false,base,ht,front,back,w,d,unit,built};
  function info(key,kind,label,description,extra={}){
   const p={id:id+':'+key,assetId:id,kind,label,description,uStart:cassetteU,uHeight:2,source,...extra};
   built.parts.push(p);built.partMap.set(p.id,p);return p;
  }
  function box(parent,p,name,x,y,z,ww,hh,dd,color,transparent=false){
   const mesh=cube(parent,x,y,z,ww,hh,dd,color,.48,.36,transparent);
   mesh.name='rack-optical:'+name;mesh.userData.assetId=id;mesh.userData.partId=p.id;return mesh;
  }
  session.box=box;session.info=info;
  const housing=info('optical-housing','optical-housing','LS 배선 솔루션 · 4슬롯 광패널 하우징','금속 외함·모듈 가이드·케이블 관리 구조. LS 공개 FDF 구성을 참고한 교육용 2U 배치이며 특정 제품의 CAD 치수와 다릅니다.');
  box(tray,housing,'floor',0,base, (front+back)/2,w*.75,.024,back-front,'#8c9ca6');
  for(const sign of [-1,1]){
   box(tray,housing,'side-wall',sign*w*.378,base, (front+back)/2,.026,ht,back-front,'#2b4153');
   box(tray,housing,'pull-handle-post',sign*w*.355,base+ht*.25,front-.055,.025,ht*.52,.034,'#c4d0d6');
   box(tray,housing,'rack-ear',sign*w*.402,base,front,w*.043,ht,.035,'#b1c2ce');
   for(const yy of [.20,.78])box(tray,housing,'rack-ear-fastener',sign*w*.402,base+ht*yy,front-.025,.026,.026,.017,'#1b2530');
  }
  box(tray,housing,'back-wall',0,base,back,w*.75,ht,.025,'#23394b');
  for(let m=0;m<4;m++){
   const x=(m-1.5)*w*.181,mw=w*.17;
   const p=info('cassette-module-'+m,'optical-module','카세트 '+(m+1)+' · MPO 12F → LC duplex × 6','후면 MPO 1개에서 12개 광섬유를 전면 LC duplex 6개로 분기합니다. 4개 모듈 합계 48심/24 duplex인 교육용 예시입니다.',{moduleIndex:m});
   const group=new THREE.Group();group.name='rack-optical:module-'+m;tray.add(group);
   box(group,p,'module-base',x,base+.028,(front+back)/2,mw,.018,back-front-.07,'#b5c4ca');
   for(const side of [-1,1])box(group,p,'module-side',x+side*mw*.49,base+.038,(front+back)/2,.013,ht*.76,back-front-.07,'#253e52');
   box(group,p,'adapter-plate',x,base+.035,front+.025,mw,ht*.73,.020,'#192d3e');
   box(group,p,'release-tab',x,base+.025,front-.040,mw*.38,.023,.055,'#ba3658');
   box(group,p,'id-label-'+(m+1),x,base+ht*.80,front+.013,mw*.76,.021,.009,'#e7edf0');
   // Six duplex adapters, each with two visible LC sockets; no false light sources.
   for(let j=0;j<6;j++){
    const px=x+(j-2.5)*mw*.151,py=base+ht*.24;
    const port=info('fiber-port:'+(m*6+j),'fiber-port','카세트 '+(m+1)+' / LC duplex '+(j+1),'LC 아답터는 두 광섬유를 접속합니다. 예시: MPO 12F → LC duplex 6개. 극성·손실·굴곡반경은 선정 제품 사양으로 확인합니다.',{portIndex:m*6+j,moduleIndex:m,duplexIndex:j,fiberCount:2});
    session.portDefs.push(port);
    box(group,port,'adapter-port-'+port.portIndex,px,py,front-.008,mw*.133,ht*.43,.045,'#1688b3');
    for(const side of [-1,1]){
     box(group,port,'lc-socket',px+side*mw*.033,py+ht*.06,front-.034,mw*.043,ht*.19,.012,'#071722');
     box(group,port,'lc-latch',px+side*mw*.033,py+ht*.29,front-.035,mw*.045,.012,.022,'#d9e4e7');
    }
    stats.adapters++;stats.connectors+=2;
   }
   const mpo=info('mpo-input-'+m,'mpo-input','카세트 '+(m+1)+' · 후면 MPO 12F 입력','사전 종단된 다심 트렁크 케이블이 후면 MPO 아답터에 연결됩니다. MTP는 MPO 계열의 상표이며 적용 제품에 따라 선택합니다.',{moduleIndex:m});
   box(group,mpo,'mpo-adapter',x,base+ht*.18,back-.033,mw*.48,ht*.45,.070,'#323842');
   box(group,mpo,'mpo-key',x,base+ht*.58,back-.025,mw*.16,.014,.036,'#c9d5df');
   box(group,mpo,'mpo-connector-boot',x,base+ht*.25,back+.04,mw*.33,ht*.27,.08,'#d7b648');
   const lid=new THREE.Group();lid.name='rack-optical:module-cover-'+m;group.add(lid);
   lid.position.set(x,base+ht*.90,back-.045);
   box(lid,p,'module-lid',0,0,-(back-front)*.45,mw,.015,(back-front)*.88,'#2f5065');
   session.modules.push({group,x,mw,p,mpo});session.covers.push(lid);
  }
  // Fixed guide rails are attached to the rack, not the moving drawer.
  for(const sign of [-1,1])box(built.cassetteGroup.parent,housing,'fixed-slide-rail',sign*w*.39,base+.035,(front+back)/2,.028,.045,back-front,'#bac8ce');
  registry.set(id,session);stats.built++;return session;
 }
 function details(s){
  if(s.detail)return;s.detail=true;
  const {w,d,base,ht,front,back,box,info}=s;
  function tube(parent,p,name,points,r,color){
   const curve=new THREE.CatmullRomCurve3(points.map(v=>new THREE.Vector3(...v)));
   const mesh=new THREE.Mesh(new THREE.TubeGeometry(curve,24,r,4,false),new THREE.MeshStandardMaterial({color,roughness:.58,metalness:.02}));
   mesh.name='rack-optical:'+name;mesh.userData.assetId=s.id;mesh.userData.partId=p.id;parent.add(mesh);return mesh;
  }
  const colors=['#54b8ee','#f6a34b','#65b98b','#b9926b','#d2dbe2','#f0f1dd','#e16b65','#383f52','#eedb57','#b68bd9','#ee98bb','#5bd3d2'];
  for(let m=0;m<4;m++){
   const {group,x,mw,p,mpo}=s.modules[m],zmid=(front+back)/2;
   const detail=new THREE.Group();detail.name='rack-optical:fanout-detail-'+m;group.add(detail);
   const spool=info('fiber-manager-'+m,'fiber-management','카세트 '+(m+1)+' · 여장 및 굴곡 관리','광섬유 팬아웃과 여장을 가이드 안에 정리합니다. 이 화면의 곡률과 치수는 실제 최소 굴곡반경 검증값이 아닙니다.',{moduleIndex:m});
   for(const side of [-1,1]){
    box(detail,spool,'fiber-guide',x+side*mw*.28,base+ht*.22,zmid,mw*.10,ht*.29,d*.14,'#d9cdb2');
    box(detail,spool,'retention-finger',x+side*mw*.28,base+ht*.52,zmid,mw*.19,.014,d*.14,'#e1d8c4');
   }
   // A preterminated fan-out is shown, with no unnecessary splice holder.
   for(let f=0;f<12;f++){
    const px=x+(Math.floor(f/2)-2.5)*mw*.151+(f%2?1:-1)*mw*.030;
    const yy=base+ht*.34+(f%3)*.004;
    tube(detail,s.portDefs[m*6+Math.floor(f/2)],'fanout-fiber-'+m+'-'+f,[[x+(f-5.5)*mw*.008,yy,back-.09],[x+mw*.31,yy,zmid+d*.065],[x+mw*.28,yy,zmid-d*.075],[px,yy,front+.065]],.0035,colors[f]);stats.fibers++;
   }
   tube(detail,mpo,'mpo-trunk-'+m,[[x,base+ht*.40,back+.08],[x,base+ht*.48,back+d*.12],[x+w*.10,base+ht*.62,back+d*.21]],.019,'#ebd45e');
   box(detail,mpo,'trunk-strain-relief',x,base+ht*.20,back+.10,mw*.42,ht*.45,.030,'#161f2d');
   // Two illustrative patched duplex outputs per module; remaining ports remain accessible.
   for(let j=0;j<2;j++){
    const port=s.portDefs[m*6+j],px=x+(j-2.5)*mw*.151;
    box(detail,port,'connector-boot-'+port.portIndex,px,base+ht*.27,front-.070,mw*.11,ht*.32,.065,'#228fb7');
    tube(detail,port,'lc-patch-'+port.portIndex,[[px,base+ht*.40,front-.11],[px,base-ht*.45,front-.18],[x-mw*.48,base-ht*.56,front-.14],[x-mw*.48,base+ht*1.5,front-.10]],.008,'#f5d65e');
   }
  }
 }
 function update(id,moving,legacy){
  const s=registry.get(id);if(!s)return null;
  const open=Math.max(0,Math.min(1,Number(legacy?.cassette)||0));
  if(open>.02)details(s);
  if(moving?.cassette)moving.cassette.position.z=-open*(moving.reach||1)*.70;
  s.covers.forEach(lid=>{lid.rotation.x=-open*1.38;});
  s.modules.forEach((m,i)=>{m.group.position.y=s.exploded?open*(i===1?.42:.12):0;});
  s.open=open;stats.updates++;return {open,portCount:24,moduleCount:4,drawerZ:moving?.cassette?.position.z||0};
 }
 function state(id){const s=registry.get(id);return s?{open:s.open,ports:s.portDefs.map(p=>p.id),moduleCount:4,fibers:48,duplexPorts:24,detail:s.detail,exploded:s.exploded}:null;}
 function explode(id){const s=registry.get(id);if(!s)return false;s.exploded=!s.exploded;return s.exploded;}
 return {decorate,update,state,explode,drop:id=>registry.delete(id),reset:()=>registry.clear(),stats};
}
