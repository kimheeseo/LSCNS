// STEP 15/16 scene + DOM bridge. No additional animation frame, no BOM mutation.
import {createMMRRoom} from './mmr-room.mjs?v=20261015-step15-detail-b';
import {createColocationCages} from './colocation-cages.mjs?v=20261015-step16';
export function createZoneSpecializations(THREE,T,{cube,world,canvas}){
  const mmr=createMMRRoom(THREE,{cube});
  const colocation=createColocationCages(THREE,{cube});
  world.add(mmr.root,colocation.root);
  let mmrPanel=null,coloPanel=null,statusMMR=null,statusColo=null,lastMMR='',lastColo='';
  const getDocument=()=>typeof document==='undefined'?null:document;
  const asset=id=>T.assetMap?.get(id);
  const canUseCamera=()=>!!T.cameraDesired&&!T.workerViewId&&!T.walkActive;
  const bayAsset={carrier:'carrier',handoff:'mmr','cross-connect':'odf',core:'core'};
  const statusLabel=(panel,text,prev)=>{
    if(panel&&text!==prev)panel.querySelector('[data-summary]').textContent=text;
    return text;
  };
  function pickMMRBay(id){
    const bay=mmr.selectBay(id);if(!bay)return null;
    const associated=asset(bayAsset[id]);
    if(associated)T.pickAsset(associated,true);
    return bay;
  }
  function pickMMRLink(id){
    const accepted=mmr.linkIds.includes(id)?mmr.selectLink(id):null;
    if(accepted)renderStatus();
    return accepted;
  }
  function pickTenant(id){
    const selected=colocation.selectTenant(id);
    if(selected!=='all'){
      const a=asset(colocation.focusRackId(selected));
      if(a)T.pickAsset(a,false);
    }
    return selected;
  }
  function pickTenantCircuit(id,kind){
    const detail=colocation.selectCircuit(id,kind);
    if(!detail)return null;
    renderStatus();
    return detail;
  }
  function viewMMR(){
    const doc=getDocument();
    if(!doc||!canvas.parentElement)return null;
    if(mmrPanel)return mmrPanel;
    mmrPanel=doc.createElement('section');
    mmrPanel.id='ls3dMMRControls';
    mmrPanel.setAttribute('aria-label','MMR 광회선 인입 및 상호연결 보기');
    mmrPanel.style.cssText=
      'position:absolute;left:12px;top:73px;z-index:25;max-width:min(330px,calc(100% - 24px));max-height:55%;overflow-y:auto;padding:10px 12px;border:1px solid #68b4c9;border-radius:10px;background:rgba(14,34,52,.94);color:#ecf9ff;font:12px system-ui;pointer-events:auto;box-shadow:0 4px 22px #0006';
    const title=doc.createElement('strong');
    title.textContent='◎ Network / MMR — 광회선 연결';
    mmrPanel.appendChild(title);
    const summary=doc.createElement('div');summary.dataset.summary='mmr';
    summary.style.cssText='font-size:11px;margin:5px 0;color:#b9d9e7';
    mmrPanel.appendChild(summary);
    const grid=doc.createElement('div');
    grid.style.cssText='display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:5px';
    const actions=[
      ['mmrBay','carrier','① Carrier 인입'],
      ['mmrBay','handoff','② MMR / DWDM'],
      ['mmrBay','cross-connect','③ ODF Cross-connect'],
      ['mmrBay','core','④ Core Router'],
      ['mmrDoor','carrier','Carrier 출입문'],
      ['mmrDoor','handoff','DWDM 출입문'],
      ['mmrDoor','cross-connect','ODF 출입문'],
      ['mmrDoor','core','Core 출입문'],
      ['mmrLink','all','전체 광회선'],
      ['mmrLink','external-carrier','외부→Carrier'],
      ['mmrLink','carrier-odf','Carrier→ODF'],
      ['mmrLink','odf-dwdm','ODF→DWDM'],
      ['mmrLink','dwdm-core','DWDM→Core']
    ];
    for(const [type,id,label] of actions){
      const btn=doc.createElement('button');btn.type='button';btn.dataset[type]=id;
      btn.textContent=label;
      btn.style.cssText='background:#22465c;color:#eaf8ff;border:1px solid #54899b;border-radius:5px;padding:6px 4px;font:inherit;cursor:pointer;text-align:left';
      grid.appendChild(btn);
    }
    mmrPanel.appendChild(grid);
    mmrPanel.addEventListener('click',event=>{
      const b=event.target.closest?.('button');
      if(!b)return;
      if(b.dataset.mmrBay)pickMMRBay(b.dataset.mmrBay);
      if(b.dataset.mmrDoor)mmr.toggleDoor(b.dataset.mmrDoor);
      if(b.dataset.mmrLink)pickMMRLink(b.dataset.mmrLink);
      renderStatus();
    });
    canvas.parentElement.appendChild(mmrPanel);
    return mmrPanel;
  }
  function viewColo(){
    const doc=getDocument();
    if(!doc||!canvas.parentElement)return null;
    if(coloPanel)return coloPanel;
    coloPanel=doc.createElement('section');
    coloPanel.id='ls3dColoControls';
    coloPanel.setAttribute('aria-label','Co-location 고객 A/B/C 격리 구역');
    coloPanel.style.cssText=
      'position:absolute;left:12px;top:73px;z-index:25;max-width:min(320px,calc(100% - 24px));max-height:55%;overflow-y:auto;padding:10px 12px;border:1px solid #83c4b6;border-radius:10px;background:rgba(14,34,52,.94);color:#edfaff;font:12px system-ui;pointer-events:auto;box-shadow:0 4px 22px #0006';
    const title=doc.createElement('strong');title.textContent='▣ Customer Co-location A / B / C';coloPanel.appendChild(title);
    const summary=doc.createElement('div');summary.dataset.summary='colo';
    summary.style.cssText='font-size:11px;margin:5px 0;color:#c9e5e6';
    coloPanel.appendChild(summary);
    const row=doc.createElement('div');
    row.style.cssText='display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:5px';
    const options=[['coloTenant','all','전체 고객 배선']];
    for(const id of ['A','B','C']){
      options.push(['coloTenant',id,'고객 '+id+' 배선 격리'],
        ['coloGate',id,id+' 출입문 열기/닫기'],
        ['coloRack',id,id+' 고객 랙 선택']);
      for(const [system,label] of [['fiber','광 회선'],['power-A','전력 A'],['power-B','전력 B']])
        options.push(['coloCircuit',id+':'+system,'고객 '+id+' · '+label]);
    }
    for(const [kind,id,label] of options){
      const btn=doc.createElement('button');btn.type='button';btn.dataset[kind]=id;btn.textContent=label;
      btn.style.cssText='background:#22475a;color:#f0fcff;border:1px solid #588b97;border-radius:5px;padding:6px 4px;font:inherit;cursor:pointer;text-align:left';
      row.appendChild(btn);
    }
    coloPanel.appendChild(row);
    coloPanel.addEventListener('click',e=>{
      const b=e.target.closest?.('button');if(!b)return;
      if(b.dataset.coloTenant)pickTenant(b.dataset.coloTenant);
      if(b.dataset.coloCircuit){
        const [id,system]=b.dataset.coloCircuit.split(':');
        pickTenantCircuit(id,system);
      }
      if(b.dataset.coloGate)colocation.toggleGate(b.dataset.coloGate);
      if(b.dataset.coloRack){
        const a=asset(colocation.focusRackId(b.dataset.coloRack));
        if(a)T.pickAsset(a,true);
      }
      renderStatus();
    });
    canvas.parentElement.appendChild(coloPanel);
    return coloPanel;
  }
  function renderStatus(){
    if(mmrPanel&&!mmrPanel.hidden){
      const s=mmr.summary(),route=s.selectedRoute;
      const text='선택 회선: '+s.selectedLink+
        (route?' · '+route.from+' → '+route.to+' ('+route.system+')':' · 전체 연결')+
        ' · 4개 출입 구획 / DWDM · 교육용 예시';
      lastMMR=statusLabel(mmrPanel,text,lastMMR);
    }
    if(coloPanel&&!coloPanel.hidden){
      const s=colocation.summary(),opened=s.bays.filter(b=>b.gateOpening>.1).map(b=>b.id);
      const current=s.bays.find(b=>b.id===s.selected)?.circuits.find(c=>c.kind===s.selectedCircuit);
      const text=s.layoutIssue||('선택 고객: '+(s.selected==='all'?'전체':s.selected)+
        (current?' · '+current.label+' · '+current.from+' → '+current.to:' · 전체 광·전력 A/B')+
        ' · 열린 출입문 '+(opened.join(', ')||'없음')+' · 예시 배선/정격 미검증');
      lastColo=statusLabel(coloPanel,text,lastColo);
    }
  }
  function update(flags={},dt=.016){
    const mode=T.facilityMode,visible=T.visibleAssets();
    const inNetwork=mode!=='server-room'&&T.currentZone==='network'&&!flags.inspect;
    const inHall=mode!=='server-room'&&T.currentZone==='hall'&&!flags.inspect;
    if(inNetwork){
      mmr.build(T.Z.find(z=>z.id==='network'),visible,mode,flags);
      mmr.root.visible=mmr.summary().ready;
      if(mmr.root.visible)mmr.tick(dt);
    }else mmr.root.visible=false;
    if(inHall){
      colocation.build(visible,mode,flags);
      colocation.root.visible=colocation.summary().ready;
      if(colocation.root.visible)colocation.tick(dt);
    }else colocation.root.visible=false;
    if(mmr.root.visible)viewMMR();
    if(colocation.root.visible)viewColo();
    if(mmrPanel)mmrPanel.hidden=!mmr.root.visible;
    if(coloPanel)coloPanel.hidden=!colocation.root.visible;
    renderStatus();
    return diagnostic();
  }
  function pick(object){
    for(let o=object;o;o=o.parent){
      if(o.userData?.coloTenant){
        const id=o.userData.coloTenant;
        const system=o.userData.coloSystem;
        if(system&&colocation.circuitKinds.includes(system)){
          const details=pickTenantCircuit(id,system);
          if(details)return {kind:'customer-link',id,system,
            assetId:colocation.focusRackId(id)};
        }
        pickTenant(id);renderStatus();
        return {kind:'customer',id,assetId:colocation.focusRackId(id)};
      }
      if(o.userData?.mmrSegment&&mmr.linkIds.includes(o.userData.mmrSegment)){
        const id=o.userData.mmrSegment;
        pickMMRLink(id);
        return {kind:'mmr-link',id,assetId:
          id==='external-carrier'?'carrier':id==='carrier-odf'?'odf':
          id==='odf-dwdm'?'mmr':'core'};
      }
      if(o.userData?.mmrBay){
        const id=o.userData.mmrBay;
        pickMMRBay(id);renderStatus();
        return {kind:'mmr',id,assetId:bayAsset[id]};
      }
    }
    return null;
  }
  function diagnostic(){return {mmr:mmr.summary(),colocation:colocation.summary()};}
  function dispose(){
    mmr.dispose();colocation.dispose();
    mmrPanel?.remove();coloPanel?.remove();
    mmrPanel=null;coloPanel=null;lastMMR='';lastColo='';
    mmr.root.removeFromParent?.();colocation.root.removeFromParent?.();
  }
  return {mmr,colocation,update,pick,dispose,diagnostic,pickMMRBay,pickTenant,pickTenantCircuit};
}
