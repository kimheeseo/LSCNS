// LS Datacenter 3D — STEP 17: anchored physical 3D campus MEP overlay.
// Educational 3D geometry only: asset coordinates are NOT surveyed route lengths,
// rated cable trays, hydraulic pipe calculations or certified equipment interfaces.
export function createCampusMEPRoutes(THREE,T,{cube,root,canvas}){
  const palette={power:'#f2bf66',cooling:'#52bdda',fiber:'#63dbc0'};
  const names={power:'전력',cooling:'냉각',fiber:'광통신'};
  const yLevels={power:5.25,cooling:6.12,fiber:7.36};
  const labels={power:'전력 피더/Busway (개념)',cooling:'냉각수 순환 배관 (개념)',
    fiber:'광 트렁크/패치케이블 (개념)'};
  // Source-of-truth equipment IDs: no fabricated catalog items/BOM quantities.
  // Each tuple is a physical teaching example, not guaranteed cabling architecture.
  const chains=[
    ['power','utility-to-rack',['xfmr-a','switchgear','ups-a','pdu','tor-a']],
    ['power','power-to-colo-A',['pdu','colo-rack-a']],
    ['power','power-to-colo-B',['pdu','colo-rack-b']],
    ['power','power-to-colo-C',['pdu','colo-rack-c']],
    ['cooling','coolant-to-gpu',['tower-a','chiller','pump','cdu-plant','gray-cdu','gpu-0-0']],
    ['cooling','coolant-to-crah',['pump','crah-a']],
    ['fiber','carrier-to-tor',['carrier','mmr','odf','core','spine-a','leaf-a','tor-a']],
    ['fiber','odf-to-hall',['odf','mmr-hall','odf-hall']],
    ['fiber','hall-to-customer-A',['mmr-hall','colo-rack-a']],
    ['fiber','hall-to-customer-B',['mmr-hall','colo-rack-b']],
    ['fiber','hall-to-customer-C',['mmr-hall','colo-rack-c']]
  ];
  let records=[],signature='',selectionSystem='all',selectionId='all';
  let panel=null,details=null,summaryNode=null,selectNode=null;
  const stats={builds:0,links:0,tubes:0,traySegments:0,pipeClamps:0,
    discardedGeometries:0,clicks:0};
  const available=()=>T.visibleAssets?.()||[];
  const enabled=(flags,key)=>flags?.flow?.[key]!==false;
  const getZone=()=>T.currentZone||'campus';
  const owned=[];
  function material(color){
    return new THREE.MeshStandardMaterial({color,metalness:.23,roughness:.47,
      transparent:true,opacity:.94,depthWrite:true});
  }
  function addTube(parent,name,points,radius,color,route){
    if(points.length<2)return null;
    const geometry=new THREE.TubeGeometry(
      new THREE.CatmullRomCurve3(points.map(p=>new THREE.Vector3(...p))),
      Math.max(18,points.length*9),radius,8,false);
    const mat=material(color),tube=new THREE.Mesh(geometry,mat);
    tube.name='mep:tube:'+name;
    tube.userData.mepRouteId=route.id;
    tube.userData.mepSystem=route.system;
    parent.add(tube);owned.push({geometry,material:mat});
    stats.tubes++;
    return tube;
  }
  function markTerminal(parent,x,y,z,system,id){
    const node=cube(parent,x,y-.08,z,.27,.23,.27,palette[system],.25,.54);
    node.name='mep:terminal:'+id;
    node.userData.mepRouteId=id;
    return node;
  }
  function tray(parent,a,b,system){
    if(system==='cooling')return;
    if(Math.abs(a[1]-b[1])>.001)return;
    const dx=b[0]-a[0],dz=b[2]-a[2];
    const length=Math.hypot(dx,dz);
    if(length<.5 || (Math.abs(dx)>.01&&Math.abs(dz)>.01))return;
    const alongX=Math.abs(dx)>Math.abs(dz),cx=(a[0]+b[0])/2,cz=(a[2]+b[2])/2;
    const w=alongX?length:.44,d=alongX?.44:length;
    const tint=system==='power'?'#9c835e':'#5c9b94';
    const base=cube(parent,cx,a[1]-.23,cz,w,.075,d,tint,.64,.37);
    base.name='mep:ladder-tray:base';base.userData.noPick=true;
    for(const sign of [-1,1]){
      const side=cube(parent,cx+(alongX?0:sign*.20),a[1]-.21,
        cz+(alongX?sign*.20:0),alongX?length:.045,.22,alongX?.045:length,tint,.58,.39);
      side.name='mep:ladder-tray:side';side.userData.noPick=true;
    }
    const n=Math.min(16,Math.max(2,Math.floor(length/2.8)));
    for(let i=1;i<n;i++){
      const t=i/n;
      const cross=cube(parent,a[0]+dx*t,a[1]-.17,a[2]+dz*t,
        alongX?.055:.40,.055,alongX?.40:.055,'#acbdc2',.54,.42);
      cross.name='mep:ladder-tray:rung';cross.userData.noPick=true;
    }
    stats.traySegments++;
  }
  function makePoints(a,b,key,serial){
    // Overhead multi-level pathway with independent source/target drops.
    // Grid-aligned service elbows are curved by CatmullRom; no paper-thin polyline.
    const y=yLevels[key]+(serial%3)*.14;
    const ay=Math.min(3.2,Math.max(1.4,(a.h||4)*.53));
    const by=Math.min(3.2,Math.max(1.4,(b.h||4)*.53));
    const mid=(a.x+b.x)/2+(serial%2?.48:-.48);
    return [
      [a.x,ay,a.z],[a.x,y,a.z],
      [mid,y,a.z],[mid,y,b.z],
      [b.x,y,b.z],[b.x,by,b.z]
    ];
  }
  function dropGeometry(){
    for(const obj of owned){
      obj.geometry?.dispose?.();obj.material?.dispose?.();
      stats.discardedGeometries++;
    }
    owned.length=0;
    while(root.children.length)root.remove(root.children[0]);
    records=[];
  }
  function selectable(a,b,system,chain){
    return {
      id:system+':'+a.id+':'+b.id,system,source:a.id,target:b.id,
      from:a.name||a.id,to:b.name||b.id,fromZone:a.zone,toZone:b.zone,
      name:chain,label:labels[system],connector:'미지정·미검증',
      ratedSpeed:null,lengthMeters:null,verified:false
    };
  }
  function physicalBuild(flags){
    dropGeometry();stats.builds++;
    const map=new Map(available().map(a=>[a.id,a]));
    const all=[];
    for(const [system,chain,ids] of chains){
      if(!enabled(flags,system))continue;
      for(let i=0;i<ids.length-1;i++){
        const a=map.get(ids[i]),b=map.get(ids[i+1]);
        if(!a||!b)continue;
        const z=getZone();
        if(z!=='campus'&&a.zone!==z&&b.zone!==z)continue;
        if(a.id===b.id)continue;
        const model=selectable(a,b,system,chain);
        if(all.some(m=>m.id===model.id))continue;
        const points=makePoints(a,b,system,i);
        const assembly=new THREE.Group();
        assembly.name='mep:link:'+model.id;
        assembly.userData.mepRouteId=model.id;
        root.add(assembly);
        addTube(assembly,model.id,points,
          system==='cooling'?.16:system==='power'?.095:.059,palette[system],model);
        markTerminal(assembly,points[0][0],points[0][1],points[0][2],system,model.id);
        markTerminal(assembly,points.at(-1)[0],points.at(-1)[1],points.at(-1)[2],system,model.id);
        if(system==='cooling'){
          // Insulating sleeve / pipe couplers visually distinguish hydraulic paths.
          for(let k=1;k<points.length-1;k++){
            const p=points[k];
            const sleeve=cube(assembly,p[0],p[1]-.11,p[2],.32,.2,.32,'#7baec3',.34,.52);
            sleeve.name='mep:pipe-coupler';sleeve.userData.noPick=true;
            stats.pipeClamps++;
          }
        }else for(let k=2;k<=4;k++)tray(assembly,points[k-1],points[k],system);
        all.push({...model,group:assembly,points});
      }
    }
    records=all;stats.links=records.length;
    applyHighlights();
    return summary();
  }
  function applyHighlights(){
    if(selectionId!=='all'&&!records.some(r=>r.id===selectionId))
      selectionId='all';
    for(const rec of records){
      const strong=(selectionId!=='all'?rec.id===selectionId:
        selectionSystem==='all'||rec.system===selectionSystem);
      rec.highlighted=strong;
      rec.group.traverse(o=>{
        if(o.name?.startsWith('mep:tube:')&&o.material){
          o.material.opacity=strong?.94:.09;
          o.material.depthWrite=strong;
        }
      });
      rec.group.visible=true;
    }
  }
  function selectSystem(system='all'){
    if(!['all','power','cooling','fiber'].includes(system))return false;
    selectionSystem=system;selectionId='all';stats.clicks++;
    applyHighlights();renderInfo();return true;
  }
  function selectRoute(id='all'){
    if(id==='all')return selectSystem('all');
    const rec=records.find(r=>r.id===id);
    if(!rec)return null;
    selectionId=id;selectionSystem=rec.system;stats.clicks++;
    applyHighlights();if(details)details.open=true;renderInfo();
    return {...rec,group:undefined,points:undefined};
  }
  function pick(object){
    for(let o=object;o;o=o.parent){
      const id=o.userData?.mepRouteId;
      if(id)return selectRoute(id);
      if(o===root)break;
    }
    return null;
  }
  function ensurePanel(){
    if(typeof document==='undefined'||!canvas?.parentElement)return;
    if(panel)return;
    panel=document.createElement('div');panel.id='ls3dMEPControls';
    panel.style.cssText='position:absolute;right:12px;top:83px;z-index:24;max-width:min(322px,calc(100% - 24px));color:#e9f6fa;font:11px system-ui;pointer-events:auto';
    details=document.createElement('details');
    details.style.cssText='background:rgba(11,29,44,.95);border:1px solid #5492a4;border-radius:9px;box-shadow:0 7px 24px #0007;max-height:min(52vh,420px);overflow:auto';
    const heading=document.createElement('summary');
    heading.style.cssText='cursor:pointer;padding:9px 11px;font-weight:700;user-select:none;';
    heading.textContent='◈ 전력 · 냉각 · 광 3D 연결 보기';details.appendChild(heading);
    const container=document.createElement('div');
    container.style.cssText='padding:3px 10px 10px;display:grid;gap:8px;';
    const buttons=document.createElement('div');
    buttons.style.cssText='display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:4px;';
    for(const [kind,name] of [['all','전체'],['power','전력'],['cooling','냉각'],['fiber','광']]){
      const b=document.createElement('button');b.type='button';b.dataset.mepSystem=kind;b.textContent=name;
      b.style.cssText='font:inherit;cursor:pointer;color:#e8f7fa;background:#23536a;border:1px solid #608fa0;border-radius:5px;padding:7px 3px;';
      buttons.appendChild(b);
    }
    container.appendChild(buttons);
    const label=document.createElement('label');label.textContent='연결 구간 선택';
    label.style.cssText='font-size:10px;display:grid;gap:4px;';
    selectNode=document.createElement('select');selectNode.id='ls3dMEPRouteSelect';
    selectNode.style.cssText='width:100%;min-width:0;max-width:100%;background:#193c54;border:1px solid #507b91;border-radius:5px;padding:7px;color:white;font:inherit;';
    label.appendChild(selectNode);container.appendChild(label);
    summaryNode=document.createElement('div');summaryNode.id='ls3dMEPRouteInfo';
    summaryNode.style.cssText='font-size:10px;line-height:1.5;color:#b8e3ee;white-space:normal';
    summaryNode.setAttribute('aria-live','polite');container.appendChild(summaryNode);
    const legend=document.createElement('div');legend.style.cssText='display:flex;gap:11px;flex-wrap:wrap;color:#aac8d2;font-size:10px;';
    for(const [system,label,color] of [['power','전력','#f2bf66'],['cooling','냉각','#52bdda'],['fiber','광','#63dbc0']]){
      const l=document.createElement('span');l.textContent='● '+label;l.style.color=color;legend.appendChild(l);
    }
    container.appendChild(legend);details.appendChild(container);panel.appendChild(details);
    panel.addEventListener('click',e=>{
      const val=e.target?.closest?.('button[data-mep-system]')?.dataset.mepSystem;
      if(val)selectSystem(val);
    });
    selectNode.addEventListener('change',()=>selectRoute(selectNode.value));
    canvas.parentElement.appendChild(panel);
  }
  function renderInfo(){
    if(!panel||panel.hidden||!summaryNode)return;
    const active=records.find(r=>r.id===selectionId),source=active?.from||'전체 연결',
      destination=active?.to||'계통별 경로';
    const info=active
      ?names[active.system]+' · '+source+' → '+destination+' · '+active.label+
        ' · 길이·규격·정격: 미검증 (시뮬레이션 예시)'
      :'총 '+records.length+'개 설비 간 연결 · 단면/경로: 개념 형상 · 실측 길이·정격 미제공';
    summaryNode.textContent=info;
    for(const b of panel.querySelectorAll('[data-mep-system]')){
      const match=b.dataset.mepSystem===selectionSystem;
      b.style.borderColor=match?'#7be5d8':'#608fa0';
      b.setAttribute('aria-pressed',String(match));
    }
    if(selectNode){
      const previous=selectNode.dataset.signature;
      const sig=records.map(r=>r.id).join('|');
      if(sig!==previous){
        selectNode.replaceChildren();
        const opt=document.createElement('option');opt.value='all';opt.textContent='모든 경로';selectNode.appendChild(opt);
        for(const r of records){
          const option=document.createElement('option');option.value=r.id;
          option.textContent=names[r.system]+' · '+r.source+' → '+r.target;
          selectNode.appendChild(option);
        }
        selectNode.dataset.signature=sig;
      }
      selectNode.value=active?.id||'all';
    }
  }
  function update(flags={}){
    const visible=T.facilityMode!=='server-room'&&!flags.inspect;
    root.visible=visible;
    if(!visible){if(panel)panel.hidden=true;return summary();}
    const assets=available();
    const sig=JSON.stringify([T.facilityMode,getZone(),['power','cooling','fiber'].map(k=>enabled(flags,k)),
      assets.map(a=>[a.id,a.zone,a.x,a.z,a.h])]);
    if(sig!==signature){
      signature=sig;physicalBuild(flags);
    }
    if(records.length){ensurePanel();if(panel){panel.hidden=false;renderInfo();}}
    else if(panel)panel.hidden=true;
    return summary();
  }
  function summary(){
    return {visible:root.visible,system:selectionSystem,selected:selectionId,
      total:records.length,builds:stats.builds,
      paths:records.map(r=>({
        id:r.id,system:r.system,source:r.source,target:r.target,
        fromZone:r.fromZone,toZone:r.toZone,
        visible:r.group.visible,highlighted:r.highlighted,
        nodeCount:r.points.length,verified:r.verified,
        lengthMeters:r.lengthMeters
      })),stats:{...stats}};
  }
  function dispose(){
    dropGeometry();signature='';if(panel)panel.remove();
    panel=null;details=null;summaryNode=null;selectNode=null;root.visible=false;
  }
  return {root,update,selectSystem,selectRoute,pick,summary,dispose,
    systems:Object.keys(palette),get records(){return records;}};
}
