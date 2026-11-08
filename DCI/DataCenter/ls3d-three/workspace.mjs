// UI integrates with the host and its render loop; it never owns equipment data.
export function createWorkspace(T, status) {
  const $=id=>document.getElementById(id), shell=$('shell');
  const button=(label,fn)=>{const b=document.createElement('button');b.type='button';b.textContent=label;b.addEventListener('click',fn);return b;};
  const side=document.querySelector('.side'), right=document.querySelector('.right');
  side.id='workspaceExplorer';right.id='workspaceInspector';
  document.body.classList.add('twin-workspace');
  if(matchMedia('(max-width:790px)').matches){shell.classList.add('explorer-closed');side.inert=true;}
  const controls=document.createElement('div');controls.className='workspace-panels';
  for(const [panel,label,cls] of [[side,'탐색','explorer-closed'],[right,'정보','inspector-closed']]){
    const toggle=button(label,()=>{const closed=shell.classList.toggle(cls);panel.inert=closed;toggle.setAttribute('aria-expanded',String(!closed));});
    toggle.setAttribute('aria-controls',panel.id);toggle.setAttribute('aria-expanded',String(!shell.classList.contains(cls)));controls.append(toggle);
  }
  document.querySelector('.top').prepend(controls);
  const operations=$('mobileOperations');
  if(operations){
    shell.classList.add('workspace-operations-collapsed');
    const toggle=button('운영 설정',()=>{const closed=shell.classList.toggle('workspace-operations-collapsed');toggle.setAttribute('aria-expanded',String(!closed));});
    toggle.setAttribute('aria-controls','mobileOperations');toggle.setAttribute('aria-expanded','false');controls.append(toggle);
  }
  const tree=document.createElement('section');tree.className='workspace-tree';tree.setAttribute('aria-label','현재 구역의 랙 및 장비');
  const heading=document.createElement('h2');heading.textContent='구역 내 장비';tree.append(heading);
  const list=document.createElement('div');tree.append(list);side.insertBefore(tree,side.querySelector('.side-bottom'));
  const path=document.createElement('div');path.className='workspace-path';path.setAttribute('aria-label','현재 탐색 경로');$('viewport').append(path);
  const tools=document.createElement('nav');tools.className='workspace-tools';tools.setAttribute('aria-label','3D 화면 조작');
  const reset=()=>{if(T.rackIsolateActive&&T.selected)status.adapter?.focusRackSide(T.selected.id,'front');else T.setZone(T.currentZone);};
  tools.append(button('회전 ↶',()=>T.dragCamera(45,0,{button:0})),button('이동 ←',()=>T.panCamera(35,0)),button('이동 →',()=>T.panCamera(-35,0)),button('확대 +',()=>T.zoomCamera(-.18)),button('축소 −',()=>T.zoomCamera(.18)),button('시점 초기화',reset),button('전체화면',()=>$('btnFull').click()));
  const xray=button('투시 꺼짐',()=>{const enabled=status.adapter?.setRackXray(!status.adapter.rackXray);xray.setAttribute('aria-pressed',String(!!enabled));xray.textContent=enabled?'투시 켜짐':'투시 꺼짐';});xray.setAttribute('aria-pressed','false');tools.append(xray);
  $('viewport').append(tools);
  // Rack roles follow the host model; labels stay anchored to rack coordinates.
  const style=document.createElement('style');
  style.textContent=`
    .netarch-toggle{position:absolute;z-index:18;top:58px;right:12px;border:1px solid #66ded0;border-radius:8px;background:#103c4aeF;color:#effffc;padding:9px 12px;font:700 12px system-ui;cursor:pointer}
    .netarch-toggle[hidden],.netarch-layer[hidden],.netarch-legend[hidden]{display:none!important}
    .netarch-toggle[aria-pressed="true"]{background:#13736b}
    .netarch-legend{position:absolute;z-index:18;top:103px;right:12px;max-width:270px;padding:10px 12px;border:1px solid #456579;border-radius:9px;background:#081b2deb;color:#e8f7ff;font:11px/1.5 system-ui;pointer-events:none}
    .netarch-legend b{color:#76ecdb}.netarch-legend em{color:#ff9ba6;font-style:normal;font-weight:800}
    .netarch-layer{position:absolute;inset:0;z-index:16;pointer-events:none;overflow:hidden}
    .netarch-marker{position:absolute;transform:translate(-50%,-100%);padding:3px 7px;border:2px solid #63e3d2;border-radius:5px;background:#062c34eF;color:white;font:800 10px/1.2 system-ui;white-space:nowrap;box-shadow:0 2px 8px #0009}
    .netarch-marker.spine{border-color:#ff6d7b;background:#661d2beF}
    @media(max-width:790px){.netarch-toggle{top:52px;right:8px;font-size:11px}.netarch-legend{top:94px;right:8px;max-width:210px;font-size:10px}}
  `;
  document.head.append(style);
  const viewport=$('viewport'),arch=button('네트워크 아키텍처 보기',()=>{
    architectureOn=!architectureOn;
    arch.setAttribute('aria-pressed',String(architectureOn));
    arch.textContent=architectureOn?'네트워크 아키텍처 닫기':'네트워크 아키텍처 보기';
    syncArchitecture();
  });
  arch.className='netarch-toggle';arch.id='networkArchitectureToggle';arch.hidden=true;arch.setAttribute('aria-pressed','false');viewport.append(arch);
  const legend=document.createElement('div');legend.className='netarch-legend';legend.hidden=true;
  legend.innerHTML='<b>LEAF</b> · GPU 랙 상단 ToR 스위치<br><em>SPINE</em> · 네트워크 랙 내부 스위치<br><small>색상은 랙 내부 스위치의 역할을 나타냅니다.</small>';
  viewport.append(legend);
  const layer=document.createElement('div');layer.className='netarch-layer';layer.hidden=true;layer.setAttribute('aria-hidden','true');viewport.append(layer);
  let architectureOn=false,architectureSignature='',architectureMarkers=[];
  function syncArchitecture(){
    const adapter=status.adapter;
    const inHall=T.currentZone==='hall'&&!T.rackIsolateActive;
    arch.hidden=!inHall;
    if(!inHall&&architectureOn){architectureOn=false;arch.setAttribute('aria-pressed','false');arch.textContent='네트워크 아키텍처 보기';}
    arch.disabled=status.status!=='ready'||!adapter;
    arch.title=arch.disabled?'Three.js 3D 화면에서 사용할 수 있습니다.':'GPU 랙과 네트워크 랙의 스위치 역할을 표시합니다.';
    const show=inHall&&architectureOn&&!!adapter&&status.status==='ready';
    layer.hidden=legend.hidden=!show;
    if(!show)return;
    const racks=T.visibleAssets().filter(a=>a.zone==='hall'&&(a.type==='GPU server rack'||a.id==='tor-a'||a.id==='tor-b'));
    const sig=racks.map(a=>a.id).join('|');
    if(sig!==architectureSignature){
      architectureSignature=sig;layer.replaceChildren();architectureMarkers=racks.map(a=>{
        const el=document.createElement('span'),spine=a.id==='tor-a'||a.id==='tor-b';
        el.className='netarch-marker'+(spine?' spine':'');el.textContent=spine?'SPINE':'LEAF';el.title=a.name;
        layer.append(el);return {id:a.id,el};
      });
    }
    const byId=new Map(racks.map(a=>[a.id,a])),width=viewport.clientWidth,height=viewport.clientHeight;
    for(const marker of architectureMarkers){
      const a=byId.get(marker.id);if(!a)continue;
      const point=new adapter.THREE.Vector3(a.x,a.h+0.35,a.z).project(adapter.camera);
      const visible=point.z>=-1&&point.z<=1&&Math.abs(point.x)<=1&&Math.abs(point.y)<=1;
      marker.el.hidden=!visible;
      if(visible){marker.el.style.left=((point.x+1)*width/2)+'px';marker.el.style.top=((1-point.y)*height/2)+'px';}
    }
  }
  const stats=document.createElement('output');stats.className='workspace-stats';stats.setAttribute('aria-label','현재 렌더링 성능');$('viewport').append(stats);
  let signature='',selection='',lastStats=0,rows=new Map();
  function update(now=performance.now()) {
    syncArchitecture();
    const adapter=status.adapter,assets=T.visibleAssets().filter(a=>T.currentZone==='campus'||a.zone===T.currentZone);
    const sig=T.facilityMode+'|'+T.currentZone+'|'+assets.map(a=>a.id).join(',');
    if(signature!==sig){signature=sig;list.replaceChildren();rows=new Map();
      for(const a of assets){const b=button(a.name,()=>{if(a.zone!==T.currentZone)T.setZone(a.zone);T.pickAsset(a,true);});b.dataset.assetId=a.id;b.title=a.id+' · '+a.type;list.append(b);rows.set(a.id,b);}
    }
    const selected=T.selected,part=adapter?.selectedPart;
    const next=[T.currentZone,selected?.id,part?.id].join('|');
    if(selection!==next){selection=next;path.replaceChildren();path.append(button('전체 공간',()=>{T.closeRackScene();T.setZone('campus');}));
      const zone=T.Z.find(z=>z.id===T.currentZone);path.append(button(zone?.ko||T.currentZone,()=>{T.closeRackScene();T.setZone(T.currentZone);}));
      if(selected){path.append(document.createTextNode(' / '+selected.name));if(part)path.append(document.createTextNode(' / '+part.label));}
      for(const [id,b] of rows)b.setAttribute('aria-current',id===selected?.id?'true':'false');
    }
    xray.disabled=!adapter||!T.rackIsolateActive||!T.isMechanicalRack(selected);
    xray.setAttribute('aria-pressed',String(!!adapter?.rackXray));xray.textContent=adapter?.rackXray?'투시 켜짐':'투시 꺼짐';
    if(now-lastStats>750){lastStats=now;const d=status.diagnostics||{},fps=d.fps;
      stats.textContent=status.status==='ready'?`${Number.isFinite(fps)?fps.toFixed(1):'—'} FPS · ${d.calls??0} calls · ${d.triangles??0} tris · G ${d.memoryGeometries??0} / T ${d.memoryTextures??0}`:'레거시 렌더러';
    }
  }
  update();return {update};
}
