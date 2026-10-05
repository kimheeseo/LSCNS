(() => {
  'use strict';
  const $=id=>document.getElementById(id);
  const esc=x=>String(x??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const fmt=n=>Number(n||0).toLocaleString('ko-KR');
  const sum=(xs,k)=>xs.reduce((v,x)=>v+Number(x?.[k]||0),0);
  function normalizeLegacy(r){
    if(!r?.usable)return null;
    return {racks:r.summary?.computeRacks||0,leaf:r.fabric?.leafCount||0,spine:r.fabric?.spineCount||0,core:r.fabric?.coreCount||0,
      endpointLinks:r.fabric?.serverLeafLinks||r.fabric?.endpointLinks||0,fabricLinks:r.fabric?.leafSpineLinks||0,coreLinks:r.fabric?.spineCoreLinks||0,
      speed:r.fabric?.linkSpeedGbps||'',protocol:r.fabric?.protocol||r.fabricProtocol||'',mode:(r.fabric?.coreCount||0)>0?'3-tier':'2-tier'};
  }
  function normalizeCapacity(r){
    if(!r?.summary)return null;
    const xs=(r.networks||[]).filter(x=>x.network==='compute');
    const core=sum(r.coreRows||[],'count');
    return {racks:r.summary.racks||0,leaf:sum(xs,'leaf'),spine:sum(xs,'spine'),core,
      endpointLinks:sum(xs,'endpoints'),fabricLinks:sum(xs,'fabricLinks'),coreLinks:sum(xs,'coreLinks'),
      speed:xs[0]?.speed||'',protocol:xs[0]?.protocol||'',mode:core>0?'3-tier':'2-tier'};
  }
  function node(x,y,w,h,title,value,sub,active=true){
    return '<g opacity="'+(active?1:.38)+'"><rect x="'+x+'" y="'+y+'" width="'+w+'" height="'+h+'" rx="16" fill="'+(active?'#142944':'#253244')+'" stroke="'+(active?'#5d91d8':'#7b8796')+'" stroke-width="2"/>'+
      '<text x="'+(x+w/2)+'" y="'+(y+31)+'" text-anchor="middle" fill="#eef6ff" font-size="16" font-weight="800">'+esc(title)+'</text>'+
      '<text x="'+(x+w/2)+'" y="'+(y+60)+'" text-anchor="middle" fill="#74b9ff" font-size="22" font-weight="900">'+esc(value)+'</text>'+
      '<text x="'+(x+w/2)+'" y="'+(y+82)+'" text-anchor="middle" fill="#aabbd0" font-size="11">'+esc(sub)+'</text></g>';
  }
  function line(x1,y1,x2,y2,label,active=true){
    return '<g opacity="'+(active?1:.35)+'"><line x1="'+x1+'" y1="'+y1+'" x2="'+x2+'" y2="'+y2+'" stroke="'+(active?'#6ea8ff':'#738296')+'" stroke-width="5" marker-end="url(#arrow)"/>'+
      '<rect x="'+((x1+x2)/2-52)+'" y="'+((y1+y2)/2-15)+'" width="104" height="22" rx="11" fill="#081423"/>'+
      '<text x="'+((x1+x2)/2)+'" y="'+((y1+y2)/2)+'" dominant-baseline="middle" text-anchor="middle" fill="#c9ddf4" font-size="10">'+esc(label)+'</text></g>';
  }
  function paint(d){
    const svg=$('fabricArchitectureSvg'),meta=$('fabricArchitectureMeta'); if(!svg||!meta||!d)return;
    const coreActive=d.core>0;
    let out='<defs><marker id="arrow" markerWidth="9" markerHeight="9" refX="8" refY="3" orient="auto"><path d="M0,0 L0,6 L9,3 z" fill="#6ea8ff"/></marker></defs>'+
      '<rect x="0" y="0" width="1100" height="430" rx="22" fill="#0a1626"/>'+
      '<text x="42" y="42" fill="#f3f8ff" font-size="22" font-weight="900">AI / Cloud Data Center Fabric</text>'+
      '<text x="42" y="66" fill="#8fa9c5" font-size="12">현재 설계 수량 기반 logical topology · 실제 포트 breakout/FEC/cabling은 BOM에서 검증</text>';
    out+=node(45,145,170,110,'Compute / Rack',fmt(d.racks)+' rack','GPU / Server endpoints');
    out+=node(300,145,170,110,'Leaf',fmt(d.leaf)+' sw','ToR / EoR access');
    out+=node(555,145,170,110,'Spine',fmt(d.spine)+' sw','fabric aggregation');
    out+=node(810,145,190,110,'Core / Super-spine',coreActive?fmt(d.core)+' sw':'Optional','inter-block / 3-tier',coreActive);
    out+=line(215,200,300,200,(d.endpointLinks?fmt(d.endpointLinks)+' links':'server links'));
    out+=line(470,200,555,200,(d.fabricLinks?fmt(d.fabricLinks)+' links':'fabric links'));
    out+=line(725,200,810,200,(d.coreLinks?fmt(d.coreLinks)+' links':'3-tier uplinks'),coreActive);
    out+='<g><rect x="300" y="305" width="425" height="70" rx="14" fill="#0f2137" stroke="#2f577f"/>'+
      '<text x="512" y="330" text-anchor="middle" fill="#dcecff" font-size="13" font-weight="800">Traffic flow</text>'+
      '<text x="512" y="354" text-anchor="middle" fill="#9fb5cd" font-size="12">Server → Leaf → Spine'+(coreActive?' → Core / Super-spine':'')+' → DCI / External network</text></g>';
    svg.innerHTML=out;
    meta.innerHTML='<div class="legendCard"><b>Topology</b><span>'+esc(d.mode)+' · '+esc(d.protocol||'protocol 미확정')+(d.speed?' · '+esc(d.speed)+'G':'')+'</span></div>'+
      '<div class="legendCard"><b>Leaf 역할</b><span>서버/NIC의 첫 번째 스위칭 계층. rack/row 단위 endpoint를 집선합니다.</span></div>'+
      '<div class="legendCard"><b>Spine 역할</b><span>Leaf 간 east-west traffic을 집선합니다. 모든 Leaf↔Spine 연결/oversubscription을 검토합니다.</span></div>'+
      '<div class="legendCard"><b>Core / Super-spine</b><span>'+(coreActive?'현재 설계에 포함. 블록/Pod 또는 fabric 간 상위 계층입니다.':'현재 2-tier 설계에서는 미사용. 규모 확대 또는 multi-block 3-tier 시 활성화됩니다.')+'</span></div>';
  }
  function renderNow(){if(window.DCCapacityDesign)paint(normalizeCapacity(window.DCCapacityDesign));else if(window.DCDesign)paint(normalizeLegacy(window.DCDesign));}
  document.addEventListener('dc:design',e=>paint(normalizeLegacy(e.detail)));
  document.addEventListener('dc:capacity-bom',e=>paint(normalizeCapacity(e.detail.r)));
  document.addEventListener('dc:capacity-invalid',()=>{const s=$('fabricArchitectureSvg');if(s)s.innerHTML='';});
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',renderNow);else renderNow();
})();