(function browserPatch(){
'use strict';
function $v(id){return document.getElementById(id)}
function txtV(e){return(e&&(e.innerText||e.textContent)||'').replace(/\s+/g,' ').trim()}
function numV(v,d){var m=String(v==null?'':v).replace(/,/g,'').match(/-?\d+(?:\.\d+)?/);return m?Number(m[0]):(d||0)}
function fmtV(v,u,d){return Number.isFinite(v)?v.toFixed(d==null?1:d)+(u?' '+u:''):'—'}
function selectedV(){if(window.DCDesign?.systemProfile)return window.DCDesign.systemProfile;try{if(typeof getSelectedSystem==='function')return getSelectedSystem()}catch(e){}return{}}
function unitsV(){if(window.DCDesign?.usable)return window.DCDesign.summary.systemUnits;return Math.max(0,numV(txtV($v('mUnits')),0))}
function racksV(){if(window.DCDesign?.usable)return window.DCDesign.summary.computeRacks||window.DCDesign.summary.totalRacks;return Math.max(1,numV(txtV($v('mComputeRacks')),1))}
function rackUV(){return Math.max(1,numV($v('rackRU')?$v('rackRU').value:48,48))}
function powerV(){if(window.DCDesign?.usable){const n=window.DCDesign.summary.totalItPowerKw;return {typical:null,design:n,peak:n};}
 var typical=null,design=0,peak=0,p=$v('v48-power-model');
 if(p){
  Array.from(p.querySelectorAll('.v48-power-cell')).forEach(function(c){
   var z=txtV(c),m=z.match(/(\d+(?:\.\d+)?)\s*kW/i);if(!m)return;
   if(/Typical IT/i.test(z))typical=Number(m[1]);
   if(/Design-Max/i.test(z))design=Number(m[1]);
   if(/Peak-Provisioning/i.test(z))peak=Number(m[1]);
  });
 }
 var s=selectedV(),u=unitsV();
 if(!design&&s){var sd=Number(s.powerDesignMax!=null?s.powerDesignMax:s.power||0);design=sd*u}
 if(!peak&&s){var sp=Number(s.powerPeakProvisioning!=null?s.powerPeakProvisioning:(s.powerDesignMax!=null?s.powerDesignMax:s.power||0));peak=sp*u}
 if(typical==null&&s&&s.powerTypical!=null)typical=Number(s.powerTypical)*u;
 return{typical:typical,design:design,peak:peak||design};
}
function roleSpeedV(id,fallback){
 var e=$v(id);if(!e)return fallback;
 if(e.value&&e.value!=='auto')return Number(e.value)||fallback;
 var note=txtV($v('v48-role-note')),label=id.indexOf('compute')>=0?'Compute':id.indexOf('storage')>=0?'Storage':'In-Band';
 var re=new RegExp(label+'\\s+(\\d+)G','i'),m=note.match(re);return m?Number(m[1]):fallback;
}
function valV(id){var e=$v(id);return e?e.value:''}
function pueV(){if($v('pue'))return Number($v('pue').value)||1.2;var v=numV(valV('v60-pue'),1.2);return Math.max(1,v||1.2)}
function wueV(){var raw=valV('v60-wue');return raw===''?null:Math.max(0,numV(raw,0))}
function mountV(){
 var p=$v('v60-system-engineering');if(p)return p;
 var anchor=$v('v53-cooling-architecture')||$v('v51-cooling-architecture')||$v('v48-power-model')||$v('v48-rack-twin');
 if(!anchor)return null;
 p=document.createElement('section');p.id='v60-system-engineering';
 p.innerHTML=
  '<h3>v6 System Engineering · Facility / PUE / Reliability / 5E</h3>'+
  '<div class="v60-sub">기존 Compute–Network–Optical BOM 위에 Facility Power, PUE/WUE, Rack constraint, Reliability, 5E engineering status를 추가합니다. PUE는 설계 시나리오이며 실제 운영 PUE 측정값과 구분합니다.</div>'+
  '<div class="v60-controls">'+
   '<div class="v60-control"><label>PUE Scenario</label><input id="v60-pue" type="number" min="1" step="0.01" value="1.20"><div class="v60-note">Editable design assumption · not measured PUE</div></div>'+
   '<div class="v60-control"><label>PUE Measurement Boundary</label><select id="v60-pue-cat"><option value="0">Category 0 · instantaneous · UPS output</option><option value="1">Category 1 · annual · UPS output</option><option value="2">Category 2 · annual · RPP output</option><option value="3">Category 3 · annual · Rack PDU output</option></select></div>'+
   '<div class="v60-control"><label>WUE (L/kWh, optional)</label><input id="v60-wue" type="number" min="0" step="0.01" placeholder="optional"></div>'+
   '<div class="v60-control"><label>Cooling Mode</label><select id="v60-cooling-mode"><option value="review">Review / Not fixed</option><option>Air Cooling</option><option>Direct-to-Chip Cold Plate</option><option>Air + Liquid Hybrid</option><option>Immersion Cooling</option></select></div>'+
   '<div class="v60-control"><label>Rack Power Limit (kW, optional)</label><input id="v60-rack-power-limit" type="number" min="0" step="1" placeholder="site/design limit"></div>'+
   '<div class="v60-control"><label>Rack Cooling Limit (kW, optional)</label><input id="v60-rack-cooling-limit" type="number" min="0" step="1" placeholder="thermal limit"></div>'+
   '<div class="v60-control"><label>Power Redundancy</label><select id="v60-power-red"><option value="">Not set</option><option>N</option><option>N+1</option><option>2N</option><option>2N+1</option></select></div>'+
   '<div class="v60-control"><label>Cooling Redundancy</label><select id="v60-cooling-red"><option value="">Not set</option><option>N</option><option>N+1</option><option>2N</option></select></div>'+
   '<div class="v60-control"><label>Network Redundancy</label><select id="v60-network-red"><option value="">Not set</option><option>Single Fabric</option><option>Dual Fabric</option></select></div>'+
   '<div class="v60-control"><label>Availability Target</label><select id="v60-avail-target"><option value="">Not set</option><option>99.9%</option><option>99.99%</option><option>99.999%</option></select></div>'+
   '<div class="v60-control"><label>MTBF (h, optional)</label><input id="v60-mtbf" type="number" min="0" step="1" placeholder="component MTBF"></div>'+
   '<div class="v60-control"><label>MTTR (h, optional)</label><input id="v60-mttr" type="number" min="0" step="0.1" placeholder="component MTTR"></div>'+
  '</div>'+
  '<div class="v60-section"><h4>Facility Power & PUE</h4><div class="v60-grid" id="v60-facility-grid"></div><div class="v60-formula" id="v60-pue-note"></div></div>'+
  '<div class="v60-section"><h4>Network / Optical Architecture Lens</h4><div class="v60-fabrics" id="v60-fabrics"></div></div>'+
  '<div class="v60-section"><div class="v60-two"><div><h4>Rack Constraint Review</h4><table class="v60-table" id="v60-rack-table"></table><div class="v60-note">Recommended systems/rack = minimum of available RU, optional rack power limit, and optional rack cooling limit. Empty limits are not assumed.</div></div><div><h4>Reliability / Resilience Review</h4><table class="v60-table" id="v60-rel-table"></table><div class="v60-note">Availability proxy is calculated only when MTBF and MTTR are supplied. Redundancy selections are configuration checks, not a certified site-availability calculation.</div></div></div></div>'+
  '<div class="v60-section"><h4>5E Engineering Dashboard</h4><div class="v60-5e" id="v60-5e"></div><div class="v60-note">5E is used here as a transparent engineering-status framework, not as a Huawei certification or an industry-standard score.</div></div>'+
  '<div class="v60-section"><h4>Supply Chain / Operations Readiness</h4><div class="v60-grid" id="v60-supply-grid"></div></div>'+
  '<div class="v60-formula">Reference basis incorporated in v6: Huawei Data Center 2030 concepts (5E, flexible resources, peer-to-peer interconnection), DEFOG PUE measurement boundaries, Siemens cooling/digital-twin engineering, uploaded GPU-infrastructure and resilience/supply-chain white papers. Where the sources do not provide a design value, v6 leaves the input editable or unset rather than inventing a default.</div>';
 anchor.insertAdjacentElement('afterend',p);
 p.querySelectorAll('input,select').forEach(function(e){e.addEventListener('input',renderV);e.addEventListener('change',renderV)});
 return p;
}
function metricV(k,v,s){return'<div class="v60-metric"><div class="k">'+k+'</div><div class="v">'+v+'</div>'+(s?'<div class="s">'+s+'</div>':'')+'</div>'}
function bomUpsertV(cat,item,qty,key){
 var body=$v('bom');if(!body)return;
 var row=body.querySelector('tr[data-v60-row="'+key+'"]');
 if(!row){row=document.createElement('tr');row.setAttribute('data-v60-row',key);row.innerHTML='<td></td><td></td><td></td>';body.appendChild(row)}
 var td=row.querySelectorAll('td');if(td.length>=3){td[0].textContent=cat;td[1].textContent=item;td[2].textContent=qty}
}
function renderV(){
 var p=mountV();if(!p)return;if($v('pue'))$v('v60-pue').value=$v('pue').value;
 var pw=powerV(),pue=pueV(),wue=wueV(),facility=pw.design*pue,overhead=Math.max(0,facility-pw.design),peakFacility=pw.peak*pue;
 var water=(wue==null?null:wue*pw.design),cat=valV('v60-pue-cat')||'0';
 $v('v60-facility-grid').innerHTML=
  metricV('Design-Max IT',fmtV(pw.design,'kW',1),'3-level power model')+
  metricV('PUE Scenario',pue.toFixed(2),'editable assumption')+
  metricV('Facility Design Power',fmtV(facility,'kW',1),'IT × PUE')+
  metricV('Non-IT Overhead',fmtV(overhead,'kW',1),'facility minus IT')+
  metricV('Peak Facility Envelope',fmtV(peakFacility,'kW',1),'peak IT × PUE')+
  metricV('Water-use Proxy',water==null?'—':fmtV(water,'L/h',1),wue==null?'enter WUE to calculate':'WUE × design IT load');
 var catText={0:'Category 0 uses instantaneous kW and IT power at UPS output.',1:'Category 1 uses annual energy and IT energy at UPS output.',2:'Category 2 uses annual energy and IT energy at RPP output.',3:'Category 3 uses annual energy and IT energy at Rack PDU output.'}[cat];
 $v('v60-pue-note').innerHTML='<b>PUE = Total Facility Power / IT Equipment Power.</b> Current design estimate: '+fmtV(facility,'kW',1)+' / '+fmtV(pw.design,'kW',1)+' = <b>'+pue.toFixed(2)+'</b>. '+catText+' Category 0 is the closest fit to an instantaneous design calculation; Categories 1–3 are annual metering concepts.';
 var c=roleSpeedV('v48-compute-speed',0),st=roleSpeedV('v48-storage-speed',0),ib=roleSpeedV('v48-inband-speed',0),phys=!!$v('v48-physical-audit');
 $v('v60-fabrics').innerHTML=
  '<div class="v60-fabric"><b>Scale-Up Fabric</b><span>Vendor/system-internal accelerator fabric. v6 does not invent a speed where the BOM profile does not expose one.</span></div>'+
  '<div class="v60-fabric"><b>Scale-Out Compute</b><span>'+(c?c+'G role-selected fabric':'system reference / review')+'</span></div>'+
  '<div class="v60-fabric"><b>Storage Fabric</b><span>'+(st?st+'G role-selected fabric':'review')+'</span></div>'+
  '<div class="v60-fabric"><b>In-Band / OOB</b><span>'+(ib?ib+'G':'review')+' in-band · 1G OOB reference</span></div>'+
  '<div class="v60-fabric"><b>Optical Physicalization</b><span>'+(phys?'Logical links separated from cages / optics / cables':'Physical cage/optic audit not available')+'</span></div>';
 var s=selectedV(),u=unitsV(),r=racksV(),per=r?Math.ceil(u/r):0,ru=Number(s.ru||0),rackU=rackUV(),ruCap=ru>0?Math.floor((rackU-Number($v('reservedRU')?.value||0))/ru):null,sp=Number(s.powerDesignMax!=null?s.powerDesignMax:s.power||0);
 var pLimRaw=valV('rackPowerKw'),cLimRaw=valV('rackCoolingKw'),pLim=pLimRaw===''?null:numV(pLimRaw,0),cLim=cLimRaw===''?null:numV(cLimRaw,0);
 var pCap=pLim!=null&&sp>0?Math.floor(pLim*(1-Number($v('rackHeadroomPct')?.value||0)/100)/sp):null,cCap=cLim!=null&&sp>0?Math.floor(cLim*(1-Number($v('rackHeadroomPct')?.value||0)/100)/sp):null,caps=[ruCap,pCap,cCap].filter(function(x){return x!=null&&Number.isFinite(x)&&x>=0}),recommended=caps.length?Math.min.apply(null,caps):null;
 var bottleneck='Not enough constraints';if(recommended!=null){if(ruCap===recommended)bottleneck='RU';if(pCap===recommended)bottleneck='Power';if(cCap===recommended)bottleneck='Cooling'}
 $v('v60-rack-table').innerHTML='<thead><tr><th>Constraint</th><th>Capacity</th><th>Basis</th></tr></thead><tbody>'+
  '<tr><td>Current placement</td><td>'+per+' systems/rack</td><td>'+u+' systems / '+r+' racks</td></tr>'+
  '<tr><td>RU gross cap</td><td>'+(ruCap==null?'—':ruCap+' systems')+'</td><td>'+rackU+'U / '+(ru||'—')+'U per system</td></tr>'+
  '<tr><td>Power cap</td><td>'+(pCap==null?'Not set':pCap+' systems')+'</td><td>'+(pLim==null?'Enter rack limit':pLim+' kW / '+sp.toFixed(1)+' kW')+'</td></tr>'+
  '<tr><td>Cooling cap</td><td>'+(cCap==null?'Not set':cCap+' systems')+'</td><td>'+(cLim==null?'Enter thermal limit':cLim+' kW / '+sp.toFixed(1)+' kW')+'</td></tr>'+
  '<tr><td><b>Recommended cap</b></td><td><b>'+(recommended==null?'—':recommended+' systems')+'</b></td><td>'+bottleneck+' bottleneck</td></tr></tbody>';
 var pr=valV('v60-power-red'),cr=valV('v60-cooling-red'),nr=valV('v60-network-red'),at=valV('v60-avail-target'),mtbf=numV(valV('v60-mtbf'),0),mttr=numV(valV('v60-mttr'),0);
 var avail=mtbf>0&&mttr>=0?mtbf/(mtbf+mttr)*100:null,configured=!!(pr&&cr&&nr),robust=(pr&&pr!=='N'&&cr&&cr!=='N'&&nr==='Dual Fabric');
 $v('v60-rel-table').innerHTML='<thead><tr><th>Layer</th><th>Selection</th><th>Review</th></tr></thead><tbody>'+
  '<tr><td>Power</td><td>'+(pr||'Not set')+'</td><td>'+(pr==='N'?'SPOF review':'—')+'</td></tr>'+
  '<tr><td>Cooling</td><td>'+(cr||'Not set')+'</td><td>'+(cr==='N'?'SPOF review':'—')+'</td></tr>'+
  '<tr><td>Network</td><td>'+(nr||'Not set')+'</td><td>'+(nr==='Single Fabric'?'SPOF review':'—')+'</td></tr>'+
  '<tr><td>Availability target</td><td>'+(at||'Not set')+'</td><td>target only</td></tr>'+
  '<tr><td>MTBF/MTTR proxy</td><td>'+(avail==null?'Not calculated':avail.toFixed(5)+'%')+'</td><td>A=MTBF/(MTBF+MTTR)</td></tr>'+
  '<tr><td><b>Configuration status</b></td><td><b>'+(robust?'REDUNDANT':configured?'CONFIGURED':'REVIEW')+'</b></td><td>not certification</td></tr></tbody>';
 function card(name,desc,ok){return'<div class="v60-5e-card"><b>'+name+'</b><span>'+desc+'</span><i class="v60-status '+(ok?'v60-ok':'v60-review')+'">'+(ok?'MODELED':'REVIEW')+'</i></div>'}
 var storage=!!$v('storageFabricType');
 $v('v60-5e').innerHTML=
  card('Energy Efficiency','PUE '+pue.toFixed(2)+(wue==null?' · WUE optional':' · WUE '+wue),true)+
  card('Computing Efficiency','Selected system + 3-level power envelope',pw.design>0)+
  card('Data Efficiency',storage?'Storage fabric input present':'Storage fabric input not detected',storage)+
  card('Transmission Efficiency',(c?c+'G compute role':'compute role review')+' · '+(phys?'physical audit active':'audit review'),!!(c&&phys))+
  card('Operation Efficiency',(configured?'Redundancy inputs configured':'Redundancy / availability inputs incomplete'),configured);
 var spare=numV($v('sparePct')?$v('sparePct').value:'',0),sc=!!$v('supply-chain-btn'),mode=valV('v60-cooling-mode');
 $v('v60-supply-grid').innerHTML=
  metricV('Spare Policy',spare?spare.toFixed(1)+'%':'0 / not set','existing BOM spare input')+
  metricV('Vendor / Product Map',sc?'ACTIVE':'REVIEW','Supply Chain UI')+
  metricV('Related Vendors',sc?'REFERENCE LAYER':'REVIEW','not counted as BOM selections')+
  metricV('Cooling Mode',mode==='review'?'NOT FIXED':mode,'user-selected architecture')+
  metricV('Power Redundancy',pr||'NOT SET','facility design input')+
  metricV('Cooling Redundancy',cr||'NOT SET','facility design input');
 bomUpsertV('Facility / Efficiency','PUE scenario · Category '+cat,pue.toFixed(2),'pue');
 bomUpsertV('Facility / Efficiency','Design facility power',fmtV(facility,'kW',1),'facility-power');
 if(wue!=null)bomUpsertV('Facility / Efficiency','WUE scenario',wue.toFixed(2)+' L/kWh','wue');
 bomUpsertV('Cooling Design','Cooling mode',mode==='review'?'Review / Not fixed':mode,'cooling-mode');
 bomUpsertV('Reliability','Power / Cooling / Network', (pr||'—')+' / '+(cr||'—')+' / '+(nr||'—'),'reliability');
 if(recommended!=null)bomUpsertV('Rack Constraint','Recommended systems/rack cap',String(recommended),'rack-cap');
 try{
  if(window.currentDesignSnapshot){
   window.currentDesignSnapshot.results=window.currentDesignSnapshot.results||{};
   window.currentDesignSnapshot.results.facilityEngineering={pue:pue,pueCategory:Number(cat),wue:wue,designItKW:pw.design,facilityDesignKW:facility,facilityPeakKW:peakFacility,nonItOverheadKW:overhead};
   window.currentDesignSnapshot.results.reliability={powerRedundancy:pr||null,coolingRedundancy:cr||null,networkRedundancy:nr||null,availabilityTarget:at||null,availabilityProxyPct:avail};
   window.currentDesignSnapshot.results.rackConstraints={currentSystemsPerRack:per,ruCap:ruCap,powerCap:pCap,coolingCap:cCap,recommendedCap:recommended,bottleneck:bottleneck};
  }
 }catch(e){}
}
function startV(){
 if(!mountV()){setTimeout(startV,400);return}
 renderV();
 document.addEventListener('dc:design',renderV);
 for(const id of ['v60-rack-power-limit','v60-rack-cooling-limit','v60-cooling-mode','v60-power-red']){const el=$v(id);if(el){el.disabled=true;el.title='Use the main design inputs; this field is a reference only';}}
 document.addEventListener('click',function(e){var x=e.target;if(x&&x.closest&&x.closest('button'))setTimeout(renderV,120)},true);
 document.addEventListener('change',function(){setTimeout(renderV,80)},true);
 setTimeout(renderV,900);
}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',startV,{once:true});else startV();
})();