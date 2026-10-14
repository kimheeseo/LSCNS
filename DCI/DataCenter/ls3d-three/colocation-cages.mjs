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
  const stats={builds:0,bays:0,doors:0,meshPanels:0,powerPaths:0,fiberPaths:0,selectionChanges:0};
  const gates=new Map(),links=new Map();
  let selected='all',ready=false,lastSignature='';
  const part=(parent,name,x,y,z,w,h,d,color,transparent=false)=>{
    const obj=cube(parent,x,y,z,w,h,d,color,.41,.47,transparent);
    obj.name='colo:'+name;return obj;
  };
  function cable(parent,name,points,color,radius=.035){
    const layer=new THREE.Group();layer.name='colo:circuit:'+name;parent.add(layer);
    if(THREE.CatmullRomCurve3&&THREE.TubeGeometry&&THREE.Vector3&&THREE.MeshStandardMaterial){
      const curve=new THREE.CatmullRomCurve3(points.map(p=>new THREE.Vector3(...p)));
      const fiber=new THREE.Mesh(new THREE.TubeGeometry(curve,15,radius,6,false),
        new THREE.MeshStandardMaterial({color,metalness:.10,roughness:.58}));
      fiber.name='colo:wire:'+name;fiber.userData.noPick=true;layer.add(fiber);
    }else for(let i=1;i<points.length;i++){
      const a=points[i-1],b=points[i];
      part(layer,'test-wire-'+name+'-'+i,(a[0]+b[0])/2,(a[1]+b[1])/2,(a[2]+b[2])/2,.045,.045,.045,color);
    }
    return layer;
  }
  function clear(){
    root.traverse(obj=>{
      if(obj.name?.startsWith('colo:wire:')){
        obj.geometry?.dispose();obj.material?.dispose();
      }
    });
    while(root.children.length)root.remove(root.children[0]);
    gates.clear();links.clear();lastSignature='';ready=false;
  }
  function build(assets,mode,flags={}){
    const lookup=new Map((assets||[]).map(a=>[a.id,a]));
    const cage=lookup.get('colo');
    const valid=mode!=='server-room'&&!!cage&&tenants.every(t=>lookup.has(t.rack));
    if(!valid){clear();root.visible=false;return summary();}
    const racks=tenants.map(t=>lookup.get(t.rack));
    const signature=JSON.stringify([mode,cage.x,cage.z,cage.w,cage.d,cage.h,
      racks.map(r=>[r.id,r.x,r.z,r.w,r.d]),!!flags.showSection]);
    if(lastSignature===signature){root.visible=true;return summary();}
    clear();lastSignature=signature;root.visible=true;ready=true;stats.builds++;
    root.position.set(cage.x,0,cage.z);
    const w=cage.w,d=cage.d,h=Math.min(4.6,Math.max(3.5,cage.h));
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
    if(thresholds.slice(1).some((x,i)=>x-thresholds[i]<1.95)){
      thresholds.splice(0,4,-w/2,-w/6,w/6,w/2);
    }
    const floor=part(root,'raised-tenant-floor',0,-.07,0,w+.23,.12,d+.25,'#587884');
    floor.userData.noPick=true;
    // Per-bay enclosure: linked steel mesh panels and posts with real 3D depth.
    for(let partition=0;partition<=3;partition++){
      const xx=thresholds[partition];
      for(let s=0;s<8;s++){
        const zz=back+(s+.5)*d/8;
        const pole=part(root,'mesh-stile-'+partition+'-'+s,xx,.08,zz,.060,barH,.065,'#8ca5b0');
        pole.userData.noPick=true;
      }
      const side=part(root,'mesh-wall-'+partition,xx,.23,0,.035,barH*.79,d-.10,
        partition===0||partition===3?'#6da2b5':'#79a9b4',true);
      side.userData.noPick=true;stats.meshPanels++;
      for(let y=0;y<4;y++)
        part(root,'mesh-rail-'+partition+'-'+y,xx,.16+y*barH/4,0,.067,.07,d-.1,'#93b3bd').userData.noPick=true;
    }
    for(let section=0;section<3;section++){
      const bayWidth=thresholds[section+1]-thresholds[section];
      const center=(thresholds[section]+thresholds[section+1])/2;
      const cfg=tenants[section],rack=racks[section],localRackX=rack.x-cage.x;
      const tenantGroup=new THREE.Group();tenantGroup.name='colo:tenant:'+cfg.id;
      tenantGroup.userData.coloTenant=cfg.id;root.add(tenantGroup);
      // Fence behind and two front gate jambs, leaving a real 1.48m concept gap.
      part(tenantGroup,'rear-mesh-'+cfg.id,center,.25,back,bayWidth-.12,barH*.76,.043,cfg.color,true).userData.noPick=true;
      const sideWidth=Math.max(.15,(bayWidth-1.48)/2);
      for(const sign of [-1,1]){
        const xx=center+sign*(.74+sideWidth/2);
        part(tenantGroup,'front-mesh-'+cfg.id+'-'+sign,xx,.23,front,sideWidth,barH*.80,.041,cfg.color,true).userData.noPick=true;
        part(tenantGroup,'front-rail-'+cfg.id+'-'+sign,xx,2.87,front,sideWidth,.08,.068,'#b0c2ca').userData.noPick=true;
      }
      const entry=new THREE.Group();entry.name='colo:gate-hinge:'+cfg.id;
      entry.userData.coloTenant=cfg.id;
      entry.position.set(center-.72,.08,front);
      tenantGroup.add(entry);
      const door=part(entry,'gate-leaf-'+cfg.id,.70,.02,0,1.40,barH-.17,.067,cfg.color,true);
      door.userData.coloTenant=cfg.id;
      part(entry,'gate-handle-'+cfg.id,1.28,1.25,-.09,.095,.30,.13,'#e3d3b4').userData.coloTenant=cfg.id;
      const label=part(tenantGroup,'tenant-nameplate-'+cfg.id,center,h-.10,front+.05,1.67,.32,.16,cfg.color);
      label.userData.coloTenant=cfg.id;
      gates.set(cfg.id,{entry,target:0,opening:0});
      stats.doors++;
      // Dedicated rack pathways; each cable remains within its own x-bay.
      const fiber=cable(tenantGroup,'fiber-'+cfg.id,[
        [center+.44,3.42,front-.12],[center+.44,3.46,.50],
        [localRackX+.26,3.38,-.04],[localRackX+.26,3.18,rack.z-cage.z+.12]
      ],cfg.fiber,.035);
      const powerA=cable(tenantGroup,'power-A-'+cfg.id,[
        [center-.47,.64,front-.16],[center-.47,.70,.68],
        [localRackX-.23,.73,.07],[localRackX-.23,.75,rack.z-cage.z+.22]
      ],cfg.power,.045);
      const powerB=cable(tenantGroup,'power-B-'+cfg.id,[
        [center+.14,.98,front-.14],[center+.14,1.02,.88],
        [localRackX+.12,1.10,.04],[localRackX+.12,1.12,rack.z-cage.z+.20]
      ],cfg.powerB,.044);
      for(const link of [fiber,powerA,powerB])link.userData.coloTenant=cfg.id;
      links.set(cfg.id,{fiber,powerA,powerB,rackId:rack.id});
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
    selectTenant(selected);
    return summary();
  }
  function selectTenant(id='all'){
    selected=['A','B','C'].includes(id)?id:'all';
    stats.selectionChanges++;
    for(const [tenant,paths] of links){
      const visible=selected==='all'||selected===tenant;
      paths.fiber.visible=visible;paths.powerA.visible=visible;paths.powerB.visible=visible;
    }
    return selected;
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
    return {ready,selected,bays:tenants.map(t=>({id:t.id,rackId:t.rack,
      gateTarget:gates.get(t.id)?.target||0,gateOpening:gates.get(t.id)?.opening||0,
      linksVisible:!!links.get(t.id)?.fiber.visible,
      circuits:links.has(t.id)?3:0})),stats:{...stats}};
  }
  function dispose(){clear();selected='all';root.visible=false;}
  return {root,build,selectTenant,toggleGate,tick,focusRackId,summary,dispose,
    tenantIds:tenants.map(t=>t.id)};
}
