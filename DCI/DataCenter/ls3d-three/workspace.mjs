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
  const stats=document.createElement('output');stats.className='workspace-stats';stats.setAttribute('aria-label','현재 렌더링 성능');$('viewport').append(stats);
  let signature='',selection='',lastStats=0,rows=new Map();
  function update(now=performance.now()) {
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
