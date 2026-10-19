/* Power path explainer for the DC BOM Design Tool. Planning diagram, not a construction SLD. */
(function(){
  'use strict';
  var scenario='normal';
  var svg=document.getElementById('powerArchitectureSvg');
  var notes=document.getElementById('powerArchitectureNotes');
  var badge=document.getElementById('powerModeBadge');
  if(!svg||!notes) return;
  function esc(v){return String(v==null?'—':v).replace(/[&<>"']/g,function(c){return ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'})[c];});}
  function fmt(v,unit){return Number.isFinite(Number(v))?Number(v).toLocaleString('en-US',{maximumFractionDigits:1})+' '+unit:'—';}
  function block(x,y,w,h,title,sub,side,state){
    var cls='pwrBlock '+(side?'side'+side:'')+' '+(state||'');
    return '<g class="'+cls+'"><rect x="'+x+'" y="'+y+'" width="'+w+'" height="'+h+'" rx="12"/><text x="'+(x+w/2)+'" y="'+(y+27)+'" text-anchor="middle" class="pwrTitle">'+esc(title)+'</text><text x="'+(x+w/2)+'" y="'+(y+47)+'" text-anchor="middle" class="pwrSub">'+esc(sub)+'</text></g>';
  }
  function wire(x1,y1,x2,y2,side,active,dash){
    return '<path class="pwrWire side'+side+(active?' active':' muted')+(dash?' dashed':'')+'" d="M'+x1+' '+y1+' H'+((x1+x2)/2)+' V'+y2+' H'+x2+'"/>';
  }
  function draw(r){
    r=r||window.DCDesign||{};
    var f=r.facility||{}, input=r.input||{};
    var activeA=scenario==='normal'||scenario==='ups-a-maintenance';
    var utilityA=scenario==='normal'||scenario==='ups-a-maintenance';
    var generatorA=scenario==='utility-a-outage';
    var upsA=scenario!=='ups-a-maintenance';
    var label=scenario==='normal'?'정상 운전':scenario==='utility-a-outage'?'A 계통 정전 · 발전기 전환':'UPS A 정비 · 유지보수 바이패스';
    if(badge){badge.textContent=scenario==='normal'?'NORMAL':scenario==='utility-a-outage'?'UTILITY A OUTAGE':'UPS A MAINTENANCE';badge.className='powerBadge '+scenario;}
    var aState=utilityA?'':'failed', upsState=upsA?'':'bypass';
    var out='<defs><marker id="pwrArrowA" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="5" markerHeight="5" orient="auto-start-reverse"><path d="M 0 0 L 10 5 L 0 10 z" class="arrowA"/></marker><marker id="pwrArrowB" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="5" markerHeight="5" orient="auto-start-reverse"><path d="M 0 0 L 10 5 L 0 10 z" class="arrowB"/></marker></defs>';
    out+='<text x="560" y="26" text-anchor="middle" class="pwrHeader">'+esc(label)+'</text>';
    out+='<text x="225" y="63" text-anchor="middle" class="pwrLane">A PATH</text><text x="895" y="63" text-anchor="middle" class="pwrLane">B PATH</text>';
    out+=block(55,80,170,58,'Utility Grid A',utilityA?'Primary utility':'LOSS OF UTILITY','A',aState);
    out+=block(895,80,170,58,'Utility Grid B','Independent utility','B','');
    out+=block(55,184,170,58,'Transformer A',fmt(f.transformerDesignKva,'kVA'),'A',aState);
    out+=block(895,184,170,58,'Transformer B',fmt(f.transformerDesignKva,'kVA'),'B','');
    out+=block(55,288,170,58,'ATS / Switchgear A',generatorA?'GENERATOR SUPPLY':'Utility preferred','A',generatorA?'generator':'');
    out+=block(895,288,170,58,'ATS / Switchgear B','Utility preferred','B','');
    out+=block(275,288,170,58,'Generator A',generatorA?fmt(f.generatorDesignKw,'kW')+' · RUNNING':fmt(f.generatorDesignKw,'kW')+' · standby','A',generatorA?'generator':'standby');
    out+=block(675,288,170,58,'Generator B',fmt(f.generatorDesignKw,'kW')+' · standby','B','standby');
    out+=block(55,392,170,58,upsA?'UPS A':'Maintenance Bypass',upsA?fmt(f.upsDesignKva,'kVA'):'UPS isolated','A',upsState);
    out+=block(895,392,170,58,'UPS B',fmt(f.upsDesignKva,'kVA'),'B','');
    out+=block(285,392,170,58,'RPP A / Busway','A-side distribution','A','');
    out+=block(665,392,170,58,'RPP B / Busway','B-side distribution','B','');
    out+=block(452,460,216,48,'Dual-corded IT racks','A PDU + B PDU','', '');
    out+=wire(140,138,140,184,'A',utilityA); out+=wire(980,138,980,184,'B',true);
    out+=wire(140,242,140,288,'A',utilityA||generatorA); out+=wire(980,242,980,288,'B',true);
    out+=wire(360,346,140,288,'A',generatorA,true); out+=wire(760,346,980,288,'B',false,true);
    out+=wire(140,346,140,392,'A',activeA||generatorA); out+=wire(980,346,980,392,'B',true);
    out+=wire(225,421,285,421,'A',activeA||generatorA); out+=wire(895,421,835,421,'B',true);
    out+=wire(455,421,520,460,'A',activeA||generatorA); out+=wire(665,421,600,460,'B',true);
    if(scenario==='ups-a-maintenance') out+='<path class="pwrBypass" d="M140 346 V370 H370 V392"/><text x="265" y="367" text-anchor="middle" class="pwrCallout">maintenance bypass</text>';
    if(scenario==='utility-a-outage') out+='<text x="360" y="274" text-anchor="middle" class="pwrCallout">ATS starts Generator A · UPS bridges transfer</text>';
    svg.innerHTML=out;
    var rows=[
      ['Current design load',fmt(f.itPowerKw,'kW'),'IT load used for UPS sizing; facility load '+fmt(f.facilityPowerKw,'kW')],
      ['Calculated equipment', 'UPS '+(f.upsQty||'—')+' · Gen '+(f.generatorQty||'—')+' · Transformer '+(f.transformerQty||'—'),'Existing BOM engine; redundancy '+esc(input.facilityRedundancy||'N+1')],
      ['Rack delivery','A PDU + B PDU','Dual-corded equipment keeps one live feed during a single-path event'],
      ['Operating interpretation',label,scenario==='utility-a-outage'?'Loss of Grid A triggers ATS A and Generator A; B remains on utility.':scenario==='ups-a-maintenance'?'UPS A is isolated only through an approved maintenance-bypass procedure.':'Both utility paths are available; generators remain in standby.']
    ];
    notes.innerHTML=rows.map(function(x){return '<div class="legendCard powerNote"><b>'+esc(x[0])+'</b><strong>'+esc(x[1])+'</strong><span>'+esc(x[2])+'</span></div>';}).join('');
  }
  document.querySelectorAll('[data-power-scenario]').forEach(function(btn){btn.addEventListener('click',function(){scenario=btn.dataset.powerScenario;document.querySelectorAll('[data-power-scenario]').forEach(function(b){b.classList.toggle('active',b===btn);});draw();});});
  window.addEventListener('dc-design-updated',function(e){draw(e.detail);});
  document.querySelector('[data-power-scenario="normal"]')?.classList.add('active');
  draw();
})();