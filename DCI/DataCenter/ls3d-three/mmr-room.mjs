// STEP 15 — independently navigable/inspectable Meet-Me Room visual layer.
// Network zone owns original asset coordinates; MMR geometry never mutates host assets.
// Carrier external entry -> termination -> ODF cross-connect -> conceptual DWDM -> Core
// links are LOCAL explanatory fibers, not verified physical service routes (STEP 17).
export function createMMRRoom(THREE,{cube}) {
  const root=new THREE.Group();root.name='mmr:network-subroom';root.visible=false;
  const stats={builds:0,bays:0,modules:0,patchPorts:0,localFibers:0,selections:0};
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
    if(THREE.CatmullRomCurve3&&THREE.TubeGeometry&&THREE.Vector3&&THREE.MeshStandardMaterial){
      const geometry=new THREE.TubeGeometry(
        new THREE.CatmullRomCurve3(points.map(p=>new THREE.Vector3(...p))),20,.043,6,false);
      const mesh=new THREE.Mesh(geometry,new THREE.MeshStandardMaterial({color,metalness:.08,roughness:.53}));
      mesh.name='mmr:fiber:'+name;
      mesh.userData.noPick=true;g.add(mesh);
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
    links=[];lastSig='';
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
      // Dedicated structured distribution tray and 24 separate illustrated patch slots.
      const board=part(g,'termination-'+tag,x,4.45,back+.45,4.1,.77,.40,'#3b657e');
      board.userData.mmrBay=tag;stats.modules++;
      for(let i=0;i<24;i++){
        const xx=x-1.72+(i%12)*.31,yy=4.62+Math.floor(i/12)*.28;
        const p=part(g,'patch-'+tag+'-'+i,xx,yy,back+.21,.17,.12,.075,i%2?'#6dd9c6':'#b4cddd');
        p.userData.mmrBay=tag;p.userData.mmrPort=i;stats.patchPorts++;
      }
      for(const sx of [-1,1]){
        part(g,'grounded-cable-tray-'+tag+'-'+sx,x+sx*1.8,5.45,(back+front)/2,.09,.14,front-back,'#8098a9').userData.noPick=true;
      }
      stats.bays++;
    }
    // Explicit supplementary DWDM modules next to the existing "mmr"
    // asset: these are non-duplicated shelf overlays at the rear service wall.
    const dwdm=new THREE.Group();dwdm.name='mmr:dwdm-shelf';room.add(dwdm);
    for(let k=0;k<4;k++){
      const unit=part(dwdm,'dwdm-muxponder-'+k,-3,1.20+k*.55,back+.90,3.4,.38,.61,
        k%2?'#46778c':'#315d79');
      unit.userData.mmrBay='handoff';stats.modules++;
      for(let j=0;j<8;j++){
        const port=part(dwdm,'dwdm-port-'+k+'-'+j,-4.25+j*.36,1.34+k*.55,back+.51,.19,.085,.09,
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
    return summary();
  }
  function selectLink(id){
    selection=segmentIds.includes(id)?id:'all';
    stats.selections++;
    for(const l of links)l.group.visible=selection==='all'||l.id===selection;
    return selection;
  }
  function selectBay(id){
    if(!baySpecs.some(b=>b.id===id))return null;
    selectedBay=id;stats.selections++;
    return baySpecs.find(b=>b.id===id);
  }
  function summary(){
    return {ready:root.visible,selectedBay,selectedLink:selection,
      bays:baySpecs.map(({id,label})=>({id,label})),
      links:links.map(l=>({id:l.id,visible:l.group.visible})),stats:{...stats}};
  }
  function dispose(){clear();root.visible=false;selection='all';selectedBay=null;}
  return {root,build,selectLink,selectBay,dispose,summary,
    linkIds:segmentIds.slice(),bayIds:baySpecs.map(b=>b.id)};
}
