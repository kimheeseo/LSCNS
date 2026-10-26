// STEP 16 — independently fenced, gated and addressable Customer A/B/C Co-location cages.
// The original asset ids colo, colo-rack-a/b/c are authoritative; never create fake BOM racks.
// Color-coded rack-local illustrative A/B power + fiber distribution is not a certified circuit.
export function createColocationCages(THREE,{cube}) {
  const root=new THREE.Group();root.name='colo:tenant-cages';root.visible=false;
  const tenants=[
    {id:'A',rack:'colo-rack-a',color:'#5fc9d3',power:'#ecc469',powerB:'#e89662',fiber:'#79ddc8'},
    {id:'B',rack:'colo-rack-b',color:'#e0b56c',power:'#e6c15d',powerB:'#db9676',fiber:'#7ec7ee'},
    {id:'C',rack:'colo-rack-c',color:'#b89bdc',power:'#efca6a',powerB:'#e79779',fiber:'#b2a2ec'}
  ];
  const stats={builds:0,bays:0,doors:0,meshPanels:0,chainPanels:0,powerPaths:0,fiberPaths:0,selectionChanges:0,meshWires:0};
  const routeKinds=['fiber','power-A','power-B'];
  const labels={fiber:'광 패치/분배', 'power-A':'Rack PDU A 전원', 'power-B':'Rack PDU B 전원'};
  // This is an illustrative visual connection model, NOT electrical/optical engineering proof.
  const routeMeta=(tenant,kind,rackId)=>({
    tenant,kind,label:labels[kind],from:'공용 분배 경계 (개념)',to:rackId,
    verified:false,connector:'미검증 / 장비별 확인 필요',capacity:'가정값 없음'
  });
  const gates=new Map(),links=new Map();
  let selected='all',selectedCircuit='all',ready=false,lastSignature='',layoutIssue=null,boundaries=[];
  const part=(parent,name,x,y,z,w,h,d,color,transparent=false)=>{
    const obj=cube(parent,x,y,z,w,h,d,color,.41,.47,transparent);
    obj.name='colo:'+name;return obj;
  };
  // One batched line mesh per fence face makes the black diamond wire visible
  // without hundreds of separate draw calls or an opaque texture over racks.
  function chainMesh(parent,name,axis,fixed,min,max,height,bottom=.18){
    if(!THREE.BufferGeometry||!THREE.Float32BufferAttribute||!THREE.LineSegments||!THREE.LineBasicMaterial)return null;
    const vertices=[],pitch=.34;
    for(const slope of [-1,1])for(let intercept=min-height;intercept<=max+height;intercept+=pitch){
      const y0=Math.max(0,slope>0?min-intercept:intercept-max);
      const y1=Math.min(height,slope>0?max-intercept:intercept-min);
      if(y1-y0<.04)continue;
      const u0=intercept+slope*y0,u1=intercept+slope*y1;
      if(axis==='x')vertices.push(fixed,bottom+y0,u0,fixed,bottom+y1,u1);
      else vertices.push(u0,bottom+y0,fixed,u1,bottom+y1,fixed);
    }
    const geometry=new THREE.BufferGeometry();
    geometry.setAttribute('position',new THREE.Float32BufferAttribute(vertices,3));
    const wire=new THREE.LineSegments(geometry,new THREE.LineBasicMaterial({color:'#19242a',transparent:true,opacity:.78,depthWrite:false}));
    wire.name='colo:chain-mesh:'+name;wire.userData.noPick=true;parent.add(wire);
    stats.meshWires+=vertices.length/6;stats.chainPanels++;
    return wire;
  }
  function nameSign(parent,name,title,subtitle,x,y,z,width,height){
    if(typeof document==='undefined')return null;
    const canvas=document.createElement('canvas');canvas.width=1024;canvas.height=256;
    const ctx=canvas.getContext('2d');if(!ctx)return null;
    ctx.fillStyle='#17242e';ctx.fillRect(0,0,1024,256);
    ctx.strokeStyle='#8ca5aa';ctx.lineWidth=9;ctx.strokeRect(6,6,1012,244);
    ctx.textAlign='center';ctx.textBaseline='middle';ctx.fillStyle='#f5f7f6';
    ctx.font='bold 94px Arial,sans-serif';ctx.fillText(title,512,subtitle?104:131,940);
    if(subtitle){ctx.fillStyle='#8cd4dd';ctx.font='bold 43px Arial,sans-serif';ctx.fillText(subtitle,512,196,930);}
    const texture=new THREE.CanvasTexture(canvas);texture.colorSpace=THREE.SRGBColorSpace;
    const sign=new THREE.Mesh(new THREE.PlaneGeometry(width,height),
      new THREE.MeshBasicMaterial({map:texture,side:THREE.DoubleSide,transparent:true,depthWrite:false}));
    sign.name='colo:sign:'+name;sign.position.set(x,y,z);sign.userData.noPick=true;
    sign.userData.coloSign=title;parent.add(sign);return sign;
  }
  function cable(parent,name,points,color,radius=.035,meta=null){
    const layer=new THREE.Group();layer.name='colo:circuit:'+name;parent.add(layer);
    if(meta)Object.assign(layer.userData,{coloTenant:meta.tenant,coloSystem:meta.kind,coloRoute:meta});
    if(THREE.CatmullRomCurve3&&THREE.TubeGeometry&&THREE.Vector3&&THREE.MeshStandardMaterial){
      const curve=new THREE.CatmullRomCurve3(points.map(p=>new THREE.Vector3(...p)));
      const fiber=new THREE.Mesh(new THREE.TubeGeometry(curve,15,radius,6,false),
        new THREE.MeshStandardMaterial({color,metalness:.10,roughness:.58,transparent:true,opacity:.95,depthWrite:true}));
      fiber.name='colo:wire:'+name;
      if(meta)Object.assign(fiber.userData,{coloTenant:meta.tenant,coloSystem:meta.kind,coloRoute:meta});
      layer.add(fiber);
    }else for(let i=1;i<points.length;i++){
      const a=points[i-1],b=points[i];
      part(layer,'test-wire-'+name+'-'+i,(a[0]+b[0])/2,(a[1]+b[1])/2,(a[2]+b[2])/2,.045,.045,.045,color);
    }
    return layer;
  }
  function clear(){
    root.traverse(obj=>{
      if(/^(colo:wire:|colo:chain-mesh:|colo:sign:)/.test(obj.name||'')){
        obj.geometry?.dispose();obj.material?.map?.dispose();obj.material?.dispose();
      }
    });
    while(root.children.length)root.remove(root.children[0]);
    gates.clear();links.clear();lastSignature='';ready=false;boundaries=[];
  }
  function build(assets,mode,flags={}){
    const lookup=new Map((assets||[]).map(a=>[a.id,a]));
    const cage=lookup.get('colo');
    const valid=mode!=='server-room'&&!!cage&&tenants.every(t=>lookup.has(t.rack));
    if(!valid){clear();layoutIssue=null;root.visible=false;return summary();}
    const racks=tenants.map(t=>lookup.get(t.rack));
    const signature=JSON.stringify([mode,cage.x,cage.z,cage.w,cage.d,cage.h,
      racks.map(r=>[r.id,r.x,r.z,r.w,r.d]),!!flags.showSection]);
    if(lastSignature===signature){root.visible=true;return summary();}
    const oldGates=new Map([...gates].map(([id,g])=>[id,{opening:g.opening,target:g.target}]));
    clear();lastSignature=signature;root.visible=true;ready=true;layoutIssue=null;stats.builds++;
    root.position.set(cage.x,0,cage.z);
    // The enclosure rises above the 6 m customer racks, as in a full-height cage.
    const w=cage.w,d=cage.d,h=Math.max(3.5,cage.h,...racks.map(r=>r.h+.45));
    const front=d/2,back=-d/2,barH=h-.2;
    // Derive fence positions from ORIGINAL rack footprints so partitions cannot
    // slice through customer racks; equal thirds would cut across Customer A.
    const thresholds=[-w/2];
    for(let i=0;i<2;i++){
      const rightEdge=racks[i].x+racks[i].w/2;
      const leftEdge=racks[i+1].x-racks[i+1].w/2;
      thresholds.push((rightEdge+leftEdge)/2-cage.x);
    }
    thresholds.push(w/2);
    // Never fall back to equal-thirds and drive security walls through racks.
    const margin=.055;
    const invalid=racks.some((r,i)=>r.x-r.w/2<cage.x-w/2+margin||
      r.x+r.w/2>cage.x+w/2-margin||
      r.z-r.d/2<cage.z-d/2+margin||r.z+r.d/2>cage.z+d/2-margin||
      (i>0 && racks[i-1].x+racks[i-1].w/2+margin>=r.x-r.w/2))||
      thresholds.slice(1).some((x,i)=>x-thresholds[i]<1.6);
    if(invalid){
      layoutIssue='고객 랙 좌표와 펜스 간격이 충돌합니다. 배치 확인 전 임의 구획을 생성하지 않습니다.';
      clear();root.visible=false;return summary();
    }
    boundaries=thresholds.slice();
    const floor=part(root,'raised-tenant-floor',0,-.07,0,w+.23,.12,d+.25,'#587884');
    floor.userData.noPick=true;
    // Per-bay enclosure: linked steel mesh panels and posts with real 3D depth.
    for(let partition=0;partition<=3;partition++){
      const xx=thresholds[partition];
      for(let s=0;s<8;s++){
        const zz=back+(s+.5)*d/8;
        const pole=part(root,'mesh-stile-'+partition+'-'+s,xx,.08,zz,.060,barH,.065,'#273239');
        pole.userData.noPick=true;
      }
      // Open steel lattice, not an opaque wall or a textured 2D plane.
      for(let row=0;row<9;row++){
        const y=.19+row*barH/9;
        part(root,'partition-horizontal-'+partition+'-'+row,xx,y,0,.045,.038,d-.10,'#273239').userData.noPick=true;
        stats.meshWires++;
      }
      for(let col=0;col<14;col++){
        const z=back+(col+.5)*(d-.08)/14;
        part(root,'partition-vertical-'+partition+'-'+col,xx,.21,z,.039,barH*.93,.039,'#273239').userData.noPick=true;
        stats.meshWires++;
      }
      stats.meshPanels++;
      chainMesh(root,'side-'+partition,'x',xx,back+.04,front-.04,barH-.32);
      part(root,'side-top-rail-'+partition,xx,h-.13,0,.075,.075,d,'#1e292f').userData.noPick=true;
    }
    nameSign(root,'main','COLOCATION CAGE','CUSTOMER A · B · C',0,h-.33,front+.20,4.9,.60);
    for(let section=0;section<3;section++){
      const bayWidth=thresholds[section+1]-thresholds[section];
      const center=(thresholds[section]+thresholds[section+1])/2;
      const cfg=tenants[section],rack=racks[section],localRackX=rack.x-cage.x;
      const tenantGroup=new THREE.Group();tenantGroup.name='colo:tenant:'+cfg.id;
      tenantGroup.userData.coloTenant=cfg.id;root.add(tenantGroup);
      // Fence behind and two front gate jambs, leaving a real 1.48m concept gap.
      for(let j=0;j<=12;j++){
        const xx=thresholds[section]+.07+(bayWidth-.14)*j/12;
        part(tenantGroup,'rear-mesh-vertical-'+cfg.id+'-'+j,xx,.2,back,.045,barH*.93,.045,'#263238').userData.noPick=true;
        stats.meshWires++;
      }
      for(let j=0;j<9;j++){
        part(tenantGroup,'rear-mesh-horizontal-'+cfg.id+'-'+j,center,.23+j*barH*.76/9,back,
          bayWidth-.14,.04,.048,'#263238').userData.noPick=true;
        stats.meshWires++;
      }
      chainMesh(tenantGroup,'rear-'+cfg.id,'z',back+.012,thresholds[section]+.06,thresholds[section+1]-.06,barH-.32);
      const sideWidth=Math.max(.15,(bayWidth-1.48)/2);
      for(const sign of [-1,1]){
        const xx=center+sign*(.74+sideWidth/2);
        part(tenantGroup,'front-rail-'+cfg.id+'-'+sign,xx,h-.13,front,sideWidth,.08,.068,'#1e292f').userData.noPick=true;
        for(let j=0;j<5;j++){
          part(tenantGroup,'front-lattice-h-'+cfg.id+'-'+sign+'-'+j,xx,.2+j*barH*.70/5,front,
            sideWidth,.035,.043,'#263238').userData.noPick=true;
          stats.meshWires++;
        }
        for(let j=0;j<Math.max(2,Math.ceil(sideWidth/.26));j++){
          const count=Math.max(2,Math.ceil(sideWidth/.26)),px=xx-sideWidth/2+(j+.5)*sideWidth/count;
          part(tenantGroup,'front-lattice-v-'+cfg.id+'-'+sign+'-'+j,px,.22,front,
            .042,barH*.93,.042,'#263238').userData.noPick=true;stats.meshWires++;
        }
        chainMesh(tenantGroup,'front-'+cfg.id+'-'+sign,'z',front+.012,
          xx-sideWidth/2+.045,xx+sideWidth/2-.045,barH-.32);
      }
      const entry=new THREE.Group();entry.name='colo:gate-hinge:'+cfg.id;
      entry.userData.coloTenant=cfg.id;
      entry.position.set(center-.72,.08,front);
      tenantGroup.add(entry);
      const door=part(entry,'gate-leaf-'+cfg.id,.70,barH-.20,0,1.40,.075,.067,'#1e292f');
      door.userData.coloTenant=cfg.id;
      part(entry,'gate-bottom-'+cfg.id,.70,.03,0,1.40,.075,.067,'#1e292f').userData.coloTenant=cfg.id;
      for(const edge of [.03,1.37])part(entry,'gate-post-'+cfg.id+'-'+edge,edge,.06,0,.075,barH-.19,.075,'#1e292f').userData.coloTenant=cfg.id;
      chainMesh(entry,'gate-'+cfg.id,'z',.004,.07,1.33,barH-.38);
      part(entry,'gate-handle-'+cfg.id,1.28,1.25,-.09,.095,.30,.13,'#e3d3b4').userData.coloTenant=cfg.id;
      const label=part(tenantGroup,'tenant-nameplate-'+cfg.id,center,h-.92,front+.06,1.76,.38,.07,'#17242e');
      label.userData.coloTenant=cfg.id;
      nameSign(tenantGroup,'customer-'+cfg.id,'CUSTOMER '+cfg.id,'PRIVATE CAGE',center,h-.72,front+.12,1.73,.37);
      for(const zz of [back,front])part(tenantGroup,'roof-rail-'+cfg.id+'-'+zz,center,h-.04,zz,bayWidth,.08,.07,'#1e292f').userData.noPick=true;
      const prior=oldGates.get(cfg.id);
      gates.set(cfg.id,{entry,target:prior?.target||0,opening:prior?.opening||0});
      entry.rotation.y=-(prior?.opening||0)*1.28;
      stats.doors++;
      // Dedicated rack pathways; each cable remains within its own x-bay.
      // Distinct optical raceway, fiber adapters and A/B rear PDU feeder interfaces.
      part(tenantGroup,'fiber-overhead-raceway-'+cfg.id,center,3.78,0,
        Math.max(1.35,bayWidth*.72),.12,.23,'#6aa6ac').userData.noPick=true;
      for(let k=0;k<8;k++){
        const connector=part(tenantGroup,'fiber-adapter-'+cfg.id+'-'+k,
          localRackX-.45+k*.128,3.13,rack.z-cage.z+1.20,.085,.13,.13,k%2?'#75e4d1':'#b8dddf');
        connector.userData.coloTenant=cfg.id;connector.userData.coloSystem='fiber';
      }
      for(const side of ['A','B']){
        const x=localRackX+(side==='A'?-1:1)*rack.w*.37;
        const pdu=part(tenantGroup,'rear-pdu-feed-'+cfg.id+'-'+side,
          x,.34,rack.z-cage.z+rack.d*.50,.16,2.05,.14,side==='A'?cfg.power:cfg.powerB);
        pdu.userData.coloTenant=cfg.id;pdu.userData.coloSystem='power-'+side;
      }
      const fiber=cable(tenantGroup,'fiber-'+cfg.id,[
        [center+.44,3.42,front-.12],[center+.44,3.65,.50],
        [localRackX+.26,3.45,-.04],[localRackX+.26,3.18,rack.z-cage.z+1.20]
      ],cfg.fiber,.035,routeMeta(cfg.id,'fiber',rack.id));
      const powerA=cable(tenantGroup,'power-A-'+cfg.id,[
        [center-.47,.64,front-.16],[center-.47,.70,.68],
        [localRackX-.23,.73,.07],[localRackX-.23,.75,rack.z-cage.z+.22]
      ],cfg.power,.045,routeMeta(cfg.id,'power-A',rack.id));
      const powerB=cable(tenantGroup,'power-B-'+cfg.id,[
        [center+.14,.98,front-.14],[center+.14,1.02,.88],
        [localRackX+.12,1.10,.04],[localRackX+.12,1.12,rack.z-cage.z+.20]
      ],cfg.powerB,.044,routeMeta(cfg.id,'power-B',rack.id));
      for(const link of [fiber,powerA,powerB])link.userData.coloTenant=cfg.id;
      links.set(cfg.id,{fiber,'power-A':powerA,'power-B':powerB,powerA,powerB,rackId:rack.id,
        metadata:Object.fromEntries(routeKinds.map(k=>[k,routeMeta(cfg.id,k,rack.id)]))});
      stats.fiberPaths++;stats.powerPaths+=2;stats.bays++;
      // Colored, independent power/fiber entry nodes with tenant labels for picking.
      for(const [key,x,y,z,col] of [
        ['power-A',center-.47,.65,front-.1,cfg.power],
        ['power-B',center+.14,1.0,front-.1,cfg.powerB],
        ['fiber',center+.44,3.4,front-.1,cfg.fiber]]){
        const badge=part(tenantGroup,'customer-'+cfg.id+'-'+key,x,y,z,.34,.25,.21,col);
        badge.userData.coloTenant=cfg.id;badge.userData.coloSystem=key;
      }
    }
    applySelection();
    return summary();
  }
  function applySelection(){
    for(const [id,paths] of links){
      for(const kind of routeKinds){
        const match=(selected==='all'||selected===id)&&
          (selectedCircuit==='all'||selectedCircuit===kind);
        const group=paths[kind];group.visible=true;
        group.userData.highlighted=match;
        group.traverse(o=>{
          if(o.name?.startsWith('colo:wire:')&&o.material){
            o.material.opacity=match?.95:.11;
            o.material.depthWrite=match;
          }
        });
      }
    }
  }
  function selectTenant(id='all'){
    selected=['A','B','C'].includes(id)?id:'all';
    selectedCircuit='all';stats.selectionChanges++;applySelection();
    return selected;
  }
  function selectCircuit(id,kind){
    if(!tenants.some(t=>t.id===id)||!routeKinds.includes(kind))return null;
    selected=id;selectedCircuit=kind;stats.selectionChanges++;
    applySelection();return links.get(id)?.metadata[kind]||null;
  }
  function toggleGate(id){
    const gate=gates.get(id);if(!gate)return null;
    gate.target=gate.target?0:1;return gate.target;
  }
  function tick(dt){
    const blend=1-Math.exp(-Math.max(0,Math.min(.12,Number(dt)||0))*7.0);
    for(const gate of gates.values()){
      gate.opening+=(gate.target-gate.opening)*blend;
      gate.entry.rotation.y=-gate.opening*1.28;
    }
  }
  function focusRackId(id){
    const tenant=tenants.find(t=>t.id===id);
    return tenant?.rack||null;
  }
  function summary(){
    return {ready,selected,selectedCircuit,layoutIssue,boundaries:boundaries.slice(),
      bays:tenants.map(t=>({id:t.id,rackId:t.rack,
        gateTarget:gates.get(t.id)?.target||0,gateOpening:gates.get(t.id)?.opening||0,
        linksVisible:!!links.get(t.id)?.fiber.visible,
        circuits:routeKinds.map(kind=>({
          ...links.get(t.id)?.metadata?.[kind],
          highlighted:!!links.get(t.id)?.[kind]?.userData.highlighted,
          visible:!!links.get(t.id)?.[kind]?.visible
        }))
      })),stats:{...stats}};
  }
  function dispose(){clear();selected='all';selectedCircuit='all';layoutIssue=null;root.visible=false;}
  return {root,build,selectTenant,selectCircuit,toggleGate,tick,focusRackId,summary,dispose,
    tenantIds:tenants.map(t=>t.id),circuitKinds:routeKinds.slice()};
}
