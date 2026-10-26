// STEP 14 — visually connected ToR, server NICs, dual rear PDU A/B and cable management.
// All endpoints use the STEP 10–12 rack-local groups; not a wiring BOM or certified physical design.
// Separate illustrative rack-level links from STEP 17 inter-zone routes.
export function createRackNetworkCabling(THREE,{cube}) {
  const records=new Map();
  const stats={built:0,torSwitches:0,torPorts:0,nicLinks:0,powerALinks:0,powerBLinks:0,
    ductSegments:0,highlights:0,cableGeometryUpdates:0};
  const box=(p,name,x,y,z,w,h,d,color)=>{
    const m=cube(p,x,y,z,w,h,d,color,.42,.42);
    m.name='rack-link:'+name;return m;
  };
  const wire=(parent,name,points,color,radius=.022)=>{
    const g=new THREE.Group();g.name='rack-link:'+name;parent.add(g);
    if(THREE.CatmullRomCurve3&&THREE.TubeGeometry&&THREE.Vector3&&THREE.MeshStandardMaterial){
      const base=points.map(p=>p.slice());
      const curveFor=(offset=0)=>new THREE.CatmullRomCurve3(base.map((p,i)=>
        new THREE.Vector3(p[0],p[1],p[2]+(i===0?offset:i===1?offset*.6:0))));
      const mesh=new THREE.Mesh(new THREE.TubeGeometry(curveFor(),14,radius,5,false),
        new THREE.MeshStandardMaterial({color,metalness:.08,roughness:.61,transparent:true,opacity:.94}));
      mesh.name='rack-link:cable:'+name;mesh.userData.noPick=true;g.add(mesh);
      // A pulled sled carries the NIC endpoint; rebuild its *own* service-loop
      // curve only when the travel changes enough to matter visually.
      let previousOffset=0;
      g.userData.followServer=(offset=0)=>{
        if(!Number.isFinite(offset)||Math.abs(offset-previousOffset)<.025)return false;
        const previous=mesh.geometry;
        mesh.geometry=new THREE.TubeGeometry(curveFor(offset),14,radius,5,false);
        previous.dispose?.();previousOffset=offset;stats.cableGeometryUpdates++;
        return true;
      };
      g.userData.emphasize=(active)=>{
        mesh.material.opacity=active?.94:.11;
        mesh.material.depthWrite=!!active;
      };
    } else {
      // Non-WebGL test doubles can still inspect independent links.
      for(let j=1;j<points.length;j++){
        const a=points[j-1],b=points[j],mid=a.map((v,i)=>(v+b[i])/2);
        box(g,name+'-segment-'+j,mid[0],mid[1],mid[2],.038,.04,.038,color);
      }
    }
    return g;
  };
  function decorate(asset,built,high){
    const {id,w,d,h}=asset;
    const u=(h-.46)/built.rackUnits,low=.23;
    const torPart=built.parts.find(p=>p.kind==='tor');
    const torU=torPart?.uStart||built.rackUnits-4;
    const torY=low+(torU-1)*u+u*.76;
    const leafBPart=built.parts.find(p=>p.kind==='leaf-b');
    const leafBY=leafBPart?low+(leafBPart.uStart-1)*u+u*.76:null;
    const network=new THREE.Group();network.name='rack-link:assembly:'+id;
    high.add(network);
    const tor= new THREE.Group();tor.name='rack-link:tor-chassis';
    network.add(tor);
    box(tor,'tor-top',0,torY+.03,-d*.30,w*.73,u*1.55,d*.43,'#2a6583');
    box(tor,'tor-face',0,torY+.07,-d*.548,w*.66,u*.93,.06,'#163d53');
    for(let i=0;i<8;i++){
      const x=-w*.27+i*w*.078;
      const port=box(tor,'tor-port-'+i,x,torY+.14,-d*.585,w*.061,.076,.040,
        i%2?'#5dcad4':'#6ce6ae');
      const record={id:id+':tor-port:'+i,assetId:id,kind:'tor-port',
        label:'Top-of-Rack 포트 '+(i+1)+' (교육용)',uStart:torU,uHeight:2,portIndex:i};
      port.userData.assetId=id;port.userData.partId=record.id;
      built.parts.push(record);built.partMap.set(record.id,record);
      stats.torPorts++;
    }
    for(const sign of [-1,1])
      box(tor,'tor-uplink-'+sign,sign*w*.31,torY+.14,-d*.587,w*.045,.097,.047,'#e7bf63');
    // Distinct rear PDU A/B bus/labels and port-side harnesses. Each rail is separately pickable.
    const pdus={}, colors={A:'#f6c86a',B:'#67c8e6'};
    for(const side of ['A','B']){
      const x=(side==='A'?-1:1)*w*.405,z=d*.432;
      const bus=new THREE.Group();bus.name='rack-link:pdu-'+side;network.add(bus);
      const face=box(bus,'pdu-'+side+'-face',x,.30,z,w*.092,h*.79,.087,colors[side]);
      const strip=box(bus,'pdu-'+side+'-identity',x,h*.85,z,w*.102,.22,.10,colors[side]);
      for(const mesh of [face,strip]){
        mesh.userData.assetId=id;
        mesh.userData.partId=built.parts.find(p=>p.kind===('pdu-'+side.toLowerCase()))?.id;
      }
      for(let k=0;k<8;k++){
        const yy=.5+k*(h*.67/8);
        box(bus,'pdu-'+side+'-outlet-'+k,x,yy,z+.055,w*.054,.12,.040,'#203d4e');
      }
      pdus[side]=bus;
    }
    // Rack-local separate front fiber riser, rear power trunks, top entry cable management.
    const duct=box(network,'fiber-side-duct',w*.34,.33,-d*.42,w*.075,h*.75,.15,'#557c90');
    box(network,'rear-power-duct',0,.25,d*.48,w*.74,h*.71,.085,'#597287');
    for(let k=0;k<8;k++){
      box(network,'duct-separator-'+k,w*.34,.45+k*h*.085,-d*.51,w*.10,.055,.075,'#94b4c7');
      stats.ductSegments++;
    }
    const links=[];
    const modules=built.serverGroups||[];
    modules.forEach((module,index)=>{
      const info=module.userData.rackPart;
      if(!info)return;
      const yy=low+(info.uStart-1)*u+info.uHeight*u*.26;
      const sx=index%2?1:-1;
      const isDac=index%2===0;
      const cable=wire(network,'nic-to-tor-'+index,[
        [sx*w*.18,yy,d*.382],[sx*w*.31,yy+.13,d*.28],
        [sx*w*.34,yy+.35,-d*.04],[sx*w*.33,torY+.12,-d*.14],
        [sx*w*.22,torY+.15,-d*.57]
      ],isDac?'#77b8dc':'#e7ce62',.025);
      const cableRecord={id:id+':nic-link:'+index,assetId:id,kind:isDac?'dac-cable':'fiber-patch-cable',
        label:(isDac?'DAC twinax':'광 패치 코드')+' · 서버 NIC '+(index+1)+' → 상단 ToR',
        description:isDac?'서버 NIC와 랙 최상단 ToR 포트를 짧은 구리 직결 케이블(DAC)로 연결하는 교육용 경로입니다.':'서버 NIC와 상단 ToR을 광 점퍼로 연결하는 교육용 경로입니다.',
        uStart:info.uStart,uHeight:info.uHeight};
      built.parts.push(cableRecord);built.partMap.set(cableRecord.id,cableRecord);
      const markCable=object=>{if(object.isMesh){object.userData.assetId=id;object.userData.partId=cableRecord.id;object.userData.noPick=false;}};
      if(typeof cable.traverse==='function')cable.traverse(markCable);
      else for(const child of cable.children||[])markCable(child);
      cable.userData.cableSystem=isDac?'dac':'fiber';
      cable.userData.from=info.id;
      cable.userData.to=torPart?.id;
      stats.nicLinks++;
      links.push({kind:isDac?'dac':'fiber',from:info.id,to:torPart?.id,group:cable});
      if(leafBPart){
        const redundant=wire(network,'nic-to-leaf-b-'+index,[
          [-sx*w*.18,yy,d*.382],[-sx*w*.30,yy+.12,d*.26],
          [-sx*w*.33,yy+.35,-d*.04],[-sx*w*.30,leafBY+.12,-d*.14],
          [-sx*w*.20,leafBY+.15,-d*.57]
        ],'#54c3d4',.023);
        const route={id:id+':leaf-b-link:'+index,assetId:id,kind:'fiber-patch-cable',
          label:'서버 NIC '+(index+1)+' → Leaf B (ToR)',
          description:'Leaf A와 별개로 Leaf B에 연결하는 이중화 데이터 경로 예시입니다. Leaf A/B는 각각 Spine 1/2로 상향 연결합니다.',
          uStart:info.uStart,uHeight:info.uHeight};
        built.parts.push(route);built.partMap.set(route.id,route);
        redundant.traverse(o=>{if(o.isMesh){o.userData.assetId=id;o.userData.partId=route.id;o.userData.noPick=false;}});
        redundant.userData.cableSystem='fiber';redundant.userData.from=info.id;
        redundant.userData.to=leafBPart.id;
        links.push({kind:'fiber',from:info.id,to:leafBPart.id,group:redundant});
        stats.nicLinks++;
      }
      for(const side of ['A','B']){
        const px=(side==='A'?-1:1)*w*.405;
        const cord=wire(network,'server-'+index+'-pdu-'+side,[
          [sx*w*.22,yy,d*.33],[sx*w*.29,yy+.09,d*.40],
          [px,yy+.12,d*.445],[px,yy+.02,d*.50]
        ],colors[side],.026);
        cord.userData.cableSystem='power-'+side;
        cord.userData.from=info.id;
        cord.userData.to=(built.parts.find(p=>p.kind==='pdu-'+side.toLowerCase()))?.id;
        links.push({kind:'power-'+side,from:info.id,to:cord.userData.to,group:cord});
        if(side==='A')stats.powerALinks++;else stats.powerBLinks++;
      }
    });
    const record={id,torY,links,tor,pdus,selectedServer:null,filter:'all',built,
      serverOffsets:new Map(),lastEmphasis:null};
    records.set(id,record);stats.built++;stats.torSwitches++;
    return record;
  }
  function focus(id,filter='all'){
    const rec=records.get(id);if(!rec)return null;
    const value=['all','fiber','dac','power-A','power-B'].includes(filter)?filter:'all';
    rec.filter=value;stats.highlights++;
    update(id,rec.selectedServer);
    return value;
  }
  function update(id,selectedServer=null) {
    const rec=records.get(id);if(!rec)return null;
    rec.selectedServer=selectedServer;
    const emphasisKey=rec.filter+'|'+(selectedServer||'');
    const needsEmphasis=emphasisKey!==rec.lastEmphasis;
    const serverGroups=rec.built.serverGroups||[];
    for(const group of serverGroups){
      const serverId=group.userData?.rackPart?.id;
      if(!serverId)continue;
      const offset=Number(group.position?.z)||0;
      if(Math.abs(offset-(rec.serverOffsets.get(serverId)||0))>=.025){
        for(const link of rec.links)if(link.from===serverId)
          link.group.userData.followServer?.(offset);
        rec.serverOffsets.set(serverId,offset);
      }
    }
    if(needsEmphasis){
      for(const l of rec.links){
        const matches=(rec.filter==='all'||rec.filter===l.kind)&&
          (rec.filter==='all'||!selectedServer||l.from===selectedServer);
        // Keep all routes in view: dim unrelated ones instead of removing them.
        l.group.visible=true;
        l.group.userData.emphasize?.(matches);
        l.emphasized=matches;
      }
      rec.lastEmphasis=emphasisKey;
    }
    return {id,mode:rec.filter,selectedServer,highlighted:rec.links.filter(l=>l.emphasized).length,
      visible:rec.links.filter(l=>l.group.visible).length,links:rec.links.length,
      torU:rec.built.parts.find(p=>p.kind==='tor')?.uStart};
  }
  function diagnostic(id){
    const r=records.get(id);
    return r?{id,torU:r.built.parts.find(p=>p.kind==='tor')?.uStart,
      torY:r.torY,filter:r.filter,selectedServer:r.selectedServer,
      links:r.links.map(l=>({kind:l.kind,from:l.from,to:l.to,visible:l.group.visible,
        emphasized:!!l.emphasized}))}:null;
  }
  function drop(id){records.delete(id);}
  function reset(){records.clear();}
  return {decorate,focus,update,diagnostic,drop,reset,stats};
}
