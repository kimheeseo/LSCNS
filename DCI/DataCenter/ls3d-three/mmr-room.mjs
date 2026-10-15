// STEP 15 — independently navigable/inspectable Meet-Me Room visual layer.
// Network zone owns original asset coordinates; MMR geometry never mutates host assets.
// Carrier external entry -> termination -> ODF cross-connect -> conceptual DWDM -> Core
// links are LOCAL explanatory fibers, not verified physical service routes (STEP 17).
export function createMMRRoom(THREE,{cube}) {
  const root=new THREE.Group();root.name='mmr:network-subroom';root.visible=false;
  const stats={builds:0,bays:0,modules:0,patchPorts:0,localFibers:0,selections:0,doors:0};
  const doors=new Map();
  const segmentInfo={
    'external-carrier':{from:'외부 Carrier/OSP',to:'Carrier Demarc / ER',system:'외부 광 인입',example:true},
    'carrier-odf':{from:'Carrier Demarc / ER',to:'ODF Cross-connect',system:'광 패치 / 점퍼',example:true},
    'odf-dwdm':{from:'ODF Cross-connect',to:'DWDM/OTN Handoff',system:'광 인터커넥트',example:true},
    'dwdm-core':{from:'DWDM/OTN Handoff',to:'Core Router',system:'광 네트워크 인터페이스',example:true}
  };
  const baySpecs=[
    {id:'carrier',label:'Carrier Entrance / ER',x:-9,color:'#75b7d1'},
    {id:'handoff',label:'Meet-Me / Carrier Handoff',x:-3,color:'#b3c1cc'},
    {id:'cross-connect',label:'ODF / Cross-connect',x:3,color:'#61d3bb'},
    {id:'core',label:'Core Router / WAN',x:9,color:'#e9c27a'}
  ];
  const segmentIds=['all','external-carrier','carrier-odf','odf-dwdm','dwdm-core'];
  let selection='all',selectedBay=null,links=[],lastSig='';
  function part(parent,name,x,y,z,w,h,d,color,transparent=false){
    const obj=cube(parent,x,y,z,w,h,d,color,.35,.45,transparent);
    obj.name='mmr:'+name;
    obj.userData.mmrItem=name;obj.userData.mmrArea=selectedBay;
    return obj;
  }
  function fiber(parent,name,points,color){
    const g=new THREE.Group();g.name='mmr:link:'+name;parent.add(g);
    g.userData.mmrSegment=name;
    g.userData.mmrRoute=segmentInfo[name];
    if(THREE.CatmullRomCurve3&&THREE.TubeGeometry&&THREE.Vector3&&THREE.MeshStandardMaterial){
      const geometry=new THREE.TubeGeometry(
        new THREE.CatmullRomCurve3(points.map(p=>new THREE.Vector3(...p))),20,.043,6,false);
      const mesh=new THREE.Mesh(geometry,new THREE.MeshStandardMaterial({
        color,metalness:.08,roughness:.53,transparent:true,opacity:.95,depthWrite:true
      }));
      mesh.name='mmr:fiber:'+name;
      // Route tubes are raycast targets. Keep the segment's exact identity.
      mesh.userData.mmrSegment=name;mesh.userData.mmrRoute=segmentInfo[name];
      g.add(mesh);
    }else for(let i=1;i<points.length;i++){
      const a=points[i-1],b=points[i];
      part(g,'segment-'+name+'-'+i,(a[0]+b[0])/2,(a[1]+b[1])/2,(a[2]+b[2])/2,.09,.10,.09,color);
    }
    stats.localFibers++;links.push({id:name,group:g});return g;
  }
  function clear(){
    // shared cube/material are owned by Scene Adapter; dispose only generated
    // curve geometries/materials to avoid GPU memory growth on zone changes.
    root.traverse(o=>{
      if(o.name?.startsWith('mmr:fiber:')){
        o.geometry?.dispose();o.material?.dispose();
      }
    });
    while(root.children.length)root.remove(root.children[0]);
    links=[];doors.clear();lastSig='';
  }
  function build(zone,assets,mode,flags={}){
    const found=new Map((assets||[]).filter(a=>a.zone==='network').map(a=>[a.id,a]));
    const ready=!!zone&&mode!=='server-room'&&
      ['carrier','mmr','odf','core'].every(k=>found.has(k));
    if(!ready){clear();root.visible=false;return {ready:false};}
    const signature=JSON.stringify([zone.pos,zone.size,mode,
      [...found.keys()].sort(),!!flags.showSection]);
    if(signature===lastSig){root.visible=true;return summary();}
    clear();lastSig=signature;root.visible=true;stats.builds++;
    root.position.set(zone.pos[0],0,zone.pos[1]);
    const h=4.15,front=1.3,back=-9.15,left=-13.25,right=13.25;
    const room=new THREE.Group();room.name='mmr:physical-room';root.add(room);
    // Four separate front-edge rooms, partitioned with see-through security
    // walls; leave front-open doors and aisle access instead of solid slabs.
    for(const x of [-13.25,-6,0,6,13.25]){
      const wall=part(room,'room-divider-'+x,x,.06,(front+back)/2,.12,h,front-back,'#9dc0cb',true);
      wall.userData.noPick=true;
      for(let j=0;j<7;j++){
        const z=back+(j+.5)*(front-back)/7;
        const stile=part(room,'divider-post-'+x+'-'+j,x,.09,z,.09,h,.09,'#6d98a9');
        stile.userData.noPick=true;
      }
    }
    part(room,'rear-boundary',0,.08,back,right-left,h,.13,'#7fa9ba',true).userData.noPick=true;
    for(const bay of baySpecs){
      const g=new THREE.Group();g.name='mmr:bay:'+bay.id;room.add(g);
      g.userData.mmrBay=bay.id;
      const x=bay.x,tag=bay.id;
      // Front opening is 2.3 units wide, with two jambs and color-coded lintel.
      for(const dir of [-1,1]){
        const edge=x+dir*2.13;
        part(g,'bay-jamb-'+tag+'-'+dir,edge,.07,front,.11,h,.13,bay.color).userData.noPick=true;
        part(g,'bay-front-rail-'+tag+'-'+dir,x+dir*2.68,.12,front,1.05,.09,.13,bay.color).userData.noPick=true;
      }
      const header=part(g,'bay-nameplate-'+tag,x,h-.24,front,3.92,.27,.21,bay.color);
      header.userData.mmrBay=tag;
      // Controlled doorway is a real articulated transparent 3D panel, not a flat sprite.
      const hinge=new THREE.Group();hinge.name='mmr:door-hinge:'+tag;
      hinge.position.set(x-1.04,.08,front+.045);
      hinge.userData.mmrBay=tag;g.add(hinge);
      const leaf=part(hinge,'door-leaf-'+tag,1.00,.08,0,1.98,2.55,.065,'#75b5c8',true);
      leaf.userData.mmrBay=tag;
      part(hinge,'door-handle-'+tag,1.85,1.20,.085,.12,.26,.12,'#d8bb81').userData.mmrBay=tag;
      part(g,'card-reader-'+tag,x+1.35,1.10,front+.14,.25,.48,.16,'#2a6070').userData.mmrBay=tag;
      doors.set(tag,{hinge,target:0,opening:0});stats.doors++;
      // Front-mounted panels lie at the actual 5-high host rack front (z=-4+3/2).
      // Each bay has distinct working equipment, not a generic 3D box relabeled.
      const board=part(g,'termination-'+tag,x,3.45,-2.36,3.55,.67,.16,
        tag==='carrier'?'#6b8190':tag==='cross-connect'?'#608f92':tag==='core'?'#556987':'#47718e');
      board.userData.mmrBay=tag;stats.modules++;
      for(let i=0;i<24;i++){
        const xx=x-1.55+(i%12)*.28,yy=3.56+Math.floor(i/12)*.25;
        const p=part(g,'patch-'+tag+'-'+i,xx,yy,-2.19,.15,.115,.065,
          tag==='core'?(i%4?'#68a9d4':'#d1b16f'):(i%2?'#6dd9c6':'#b4cddd'));
        p.userData.mmrBay=tag;p.userData.mmrPort=i;stats.patchPorts++;
      }
      // Visual equipment identity: demarc block, DWDM optics, ODF patch coils, core port cage.
      if(tag==='carrier'){
        part(g,'carrier-demarc-plate',x,2.78,-2.35,2.9,.33,.20,'#d3bc7b').userData.mmrBay=tag;
        for(let k=0;k<4;k++)
          part(g,'carrier-entry-gland-'+k,x-1.12+k*.75,2.26,-2.34,.23,.20,.22,'#7896a8').userData.mmrBay=tag;
      }else if(tag==='cross-connect'){
        for(let k=0;k<3;k++){
          const tray=part(g,'odf-patch-shelf-'+k,x,2.15+k*.31,-2.31,2.94,.21,.35,'#6d9c9b');
          tray.userData.mmrBay=tag;stats.modules++;
          for(let j=0;j<8;j++)part(g,'odf-adapter-'+k+'-'+j,x-1.19+j*.34,2.22+k*.31,-2.10,.15,.09,.08,'#c2ddda').userData.mmrBay=tag;
        }
      }else if(tag==='core'){
        for(let k=0;k<3;k++){
          part(g,'core-interface-card-'+k,x,2.02+k*.34,-2.32,3.0,.23,.27,'#315d79').userData.mmrBay=tag;
          for(let j=0;j<8;j++)part(g,'core-cage-'+k+'-'+j,x-1.21+j*.35,2.09+k*.34,-2.13,.17,.095,.09,j%2?'#e2c27c':'#6cd3bf').userData.mmrBay=tag;
          stats.modules++;
        }
      }
      for(const sx of [-1,1]){
        part(g,'grounded-cable-tray-'+tag+'-'+sx,x+sx*1.8,5.45,(back+front)/2,.09,.14,front-back,'#8098a9').userData.noPick=true;
      }
      // Short curved drop from overhead raceway toward the original rack's front.
      fiber(g,'vertical-drop-'+tag,[
        [x,5.7,-4.3],[x,5.40,-3.5],[x,4.80,-2.7],[x,4.12,-2.22]
      ],'#81decf');
      stats.bays++;
    }
    // DWDM/OTN overlay mounts on the existing handoff rack at z=-4 (front=-2.5).
    const dwdm=new THREE.Group();dwdm.name='mmr:dwdm-shelf';room.add(dwdm);
    for(let k=0;k<4;k++){
      const unit=part(dwdm,'dwdm-muxponder-'+k,-3,1.20+k*.55,-2.38,3.2,.38,.25,
        k%2?'#46778c':'#315d79');
      unit.userData.mmrBay='handoff';stats.modules++;
      for(let j=0;j<8;j++){
        const port=part(dwdm,'dwdm-port-'+k+'-'+j,-4.25+j*.36,1.34+k*.55,-2.20,.19,.085,.09,
          j%2?'#58cfe0':'#b8d3dd');
        port.userData.mmrBay='handoff';stats.patchPorts++;
      }
    }
    // A local external-plant handoff enters through the WEST boundary. It
    // does not represent a factual operator POP or full-campus fiber network.
    part(room,'external-entry-sleeve',left+.20,5.35,-4.3,.50,.49,.62,'#edcb81').userData.mmrBay='carrier';
    fiber(room,'external-carrier',[
      [left-1,5.35,-4.3],[left+.7,5.36,-4.3],[-10.7,5.58,-4.3],[-9,5.6,-4.3]
    ],'#e2cb83');
    fiber(room,'carrier-odf',[
      [-9,5.57,-4.3],[-9,5.85,-6.35],[-2.5,5.85,-6.35],[3,5.54,-4.3]
    ],'#6fd8c0');
    fiber(room,'odf-dwdm',[
      [3,5.54,-4.3],[3,5.85,-6.9],[-2.3,5.85,-7.55],[-3,5.25,-8.1]
    ],'#61d9ef');
    fiber(room,'dwdm-core',[
      [-3,5.25,-8.1],[0,5.80,-7.55],[9,5.80,-7.55],[9,5.54,-4.3]
    ],'#b0a4ed');
    // Keep selection consistent when rebuilding after a mode, section, or zone change.
    applyLinkSelection();
    return summary();
  }
  function applyLinkSelection(){
    for(const link of links){
      const active=selection==='all'||link.id===selection||link.id.startsWith('vertical-drop-');
      link.group.visible=true; // maintain topology context
      link.group.traverse(object=>{
        if(object.name?.startsWith('mmr:fiber:')&&object.material){
          object.material.opacity=active?.95:.10;
          object.material.depthWrite=active;
        }
      });
      link.emphasized=active;
    }
  }
  function selectLink(id){
    selection=segmentIds.includes(id)?id:'all';
    stats.selections++;applyLinkSelection();
    return selection;
  }
  function toggleDoor(id){
    const door=doors.get(id);
    if(!door)return null;
    door.target=door.target?0:1;
    return door.target;
  }
  function tick(dt=.016){
    const blend=1-Math.exp(-Math.max(0,Math.min(.1,Number(dt)||0))*7);
    for(const door of doors.values()){
      door.opening+=(door.target-door.opening)*blend;
      door.hinge.rotation.y=-door.opening*1.3;
    }
  }
  function selectBay(id){
    if(!baySpecs.some(b=>b.id===id))return null;
    selectedBay=id;stats.selections++;
    return baySpecs.find(b=>b.id===id);
  }
  function summary(){
    return {ready:root.visible,selectedBay,selectedLink:selection,
      bays:baySpecs.map(({id,label})=>({id,label})),
      links:links.map(l=>({id:l.id,visible:l.group.visible,
        emphasized:!!l.emphasized,metadata:segmentInfo[l.id]||null})),
      doors:baySpecs.map(b=>({id:b.id,target:doors.get(b.id)?.target||0,
        opening:doors.get(b.id)?.opening||0})),
      selectedRoute:segmentInfo[selection]||null,stats:{...stats}};
  }
  function dispose(){clear();root.visible=false;selection='all';selectedBay=null;}
  return {root,build,selectLink,selectBay,toggleDoor,tick,dispose,summary,
    linkIds:segmentIds.slice(),bayIds:baySpecs.map(b=>b.id)};
}
