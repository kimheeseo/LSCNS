// STEP 13 — independently extensible optical cassette with lid/adapter/connector/fiber paths.
// Demonstration-scale geometry only; no standardized fiber bend-radius or optical-loss claim.
export function createRackOpticalCassette(THREE,{cube}) {
  const registry=new Map(),stats={built:0,adapters:0,connectors:0,fibers:0,updates:0};
  function part(group,name,x,y,z,w,h,d,color,transparency=false){
    const m=cube(group,x,y,z,w,h,d,color,.26,.42,transparency);
    m.name='rack-optical:'+name;return m;
  }
  function decorate(asset,built){
    const drawer=built.cassetteDrawer;
    if(!drawer)return null;
    const id=asset.id,w=asset.w,d=asset.d,h=asset.h;
    const top=h*.78+.10,frontZ=-d*.42;
    const tray=new THREE.Group();tray.name='rack-optical:tray';drawer.add(tray);
    part(tray,'sliding-tray-floor',0,top-.07,frontZ,w*.58,.08,d*.33,'#b8cbd0');
    for(const side of [-1,1])part(tray,'tray-side-'+side,side*w*.29,top,frontZ,.06,.17,d*.34,'#6d8d9e');
    const faceZ=frontZ-d*.125;
    const portDefs=[];
    for(let i=0;i<8;i++){
      const x=-w*.225+i*w*.065,y=top+.04;
      const frame=part(tray,'adapter-port-'+i,x,y,faceZ,w*.053,.135,.065,'#327f8e');
      const ferrule=part(tray,'mt-ferrule-'+i,x,y+.015,faceZ-.058,w*.035,.070,.044,'#e2cfad');
      const connector=part(tray,'connector-boot-'+i,x,y+.015,faceZ-.102,w*.035,.076,.075,
        i%2?'#6ad6d0':'#edc27c');
      const p={id:id+':fiber-port:'+i,assetId:id,kind:'fiber-port',label:'광 카세트 '+(i+1)+'번 포트 / 커넥터',uStart:Math.max(1,built.rackUnits-5),uHeight:2,portIndex:i};
      portDefs.push(p);built.parts.push(p);built.partMap.set(p.id,p);
      for(const mesh of [frame,ferrule,connector]){mesh.userData.assetId=id;mesh.userData.partId=p.id;}
      // 3D tube fiber within the tray: colored service-loop from adapter to splice holder.
      if(THREE.CatmullRomCurve3&&THREE.TubeGeometry&&THREE.Vector3&&THREE.MeshStandardMaterial){
        const curve=new THREE.CatmullRomCurve3([
          new THREE.Vector3(x,y,faceZ+.02),
          new THREE.Vector3(x*.77,y+.07,frontZ+.03),
          new THREE.Vector3(x*.55,y+.09,frontZ+d*.05),
          new THREE.Vector3(x*.18,y+.04,frontZ+d*.095)
        ]);
        const path=new THREE.Mesh(new THREE.TubeGeometry(curve,12,.018,5,false),
          new THREE.MeshStandardMaterial({color:i%2?'#f4b55b':'#62d7d7',metalness:.03,roughness:.63}));
        path.name='rack-optical:fiber-'+i;
        path.userData.assetId=id;path.userData.partId=p.id;tray.add(path);stats.fibers++;
      }
      stats.adapters++;stats.connectors++;
    }
    const lid=new THREE.Group();lid.name='rack-optical:cover-hinge';tray.add(lid);
    lid.position.set(0,top+.22,frontZ+d*.15);
    part(lid,'transparent-cover',0,-.16,-d*.14,w*.55,.045,d*.29,'#79b7c8',true);
    const splice=part(tray,'splice-holder',0,top+.03,frontZ+d*.115,w*.35,.10,.10,'#d8bd89');
    const session={id,drawer,cover:lid,portDefs,open:0};
    registry.set(id,session);stats.built++;
    return session;
  }
  function update(id,moving,legacy){
    const s=registry.get(id);if(!s)return null;
    const open=Math.max(0,Math.min(1,Number(legacy?.cassette)||0));
    // Existing cassette drawer group is the real sliding carriage; its children
    // include the adapter, 8 connector ferrules, patch fibers and hinged lid.
    if(moving?.cassette)moving.cassette.position.z=-open*(moving.reach||1)*.70;
    s.cover.rotation.x=-open*1.18;s.open=open;stats.updates++;
    return {open,portCount:s.portDefs.length,drawerZ:moving?.cassette?.position.z||0};
  }
  function state(id){const s=registry.get(id);return s?{open:s.open,ports:s.portDefs.map(p=>p.id)}:null;}
  function drop(id){registry.delete(id);}
  function reset(){registry.clear();}
  return {decorate,update,state,drop,reset,stats};
}
