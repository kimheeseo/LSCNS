(() => {
  'use strict';
  const by=id=>document.getElementById(id);
  const input=(id,label,value,step='any')=>`<label for="${id}"><span>${label}</span><input id="${id}" type="number" value="${value}" step="${step}"></label>`;
  const select=(id,label,values)=>`<label for="${id}"><span>${label}</span><select id="${id}">${values.map(([v,t])=>`<option value="${v}">${t}</option>`).join('')}</select></label>`;
  const box=document.createElement('details');box.id='engineering-inputs';
  box.innerHTML='<summary>설계 가정 · 범용 서버 · 광손실 · 전력 (v7.4)</summary><p>아래 가정은 실제 계산 및 Excel에 반영됩니다. 400/800G AI fabric과 콜로케이션 용량 계획을 지원합니다. 다른 속도·DC 전원·특수 토폴로지는 별도 설계가 필요합니다.</p><div class="fields sub">'+
    input('reservedRU','Rack reserved RU',4,1)+input('rackHeadroomPct','Rack power / cooling headroom %',20)+input('oversub','Leaf oversubscription target',1)+
    input('coreDistanceM','Spine→Core distance m',100)+select('coreCabling','Spine→Core cabling',[['structured','Structured'],['p2p','Point-to-point']])+
    input('pue','PUE scenario (unmeasured)',1.2)+input('facilityHeadroomPct','Facility capacity headroom %',20)+input('powerFactor','Transformer power factor',.9)+
    input('upsBlockKw','UPS block kW',1000)+input('generatorBlockKw','Generator block kW',2500)+input('transformerBlockKva','Transformer block kVA',3000)+
    input('upsUtilPct','UPS utilization %',80)+input('generatorUtilPct','Generator utilization %',80)+input('transformerUtilPct','Transformer utilization %',80)+
    input('pduRatedKw','PDU rated kW',46)+input('pduUtilPct','PDU utilization %',80)+input('pduOutlets','PDU outlets (per side)',24,1)+
    select('opticalBudgetMode','Optical budget',[['application','Planning preset (RFQ)'],['project','Project budget']])+input('projectBudgetDb','Project loss budget dB',3)+
    input('fiberAttenDbKm','OS2 attenuation dB/km',.35)+input('matedPairs','Structured mated connector pairs',6,1)+input('mpoLossDb','MPO pair loss dB',.35)+input('lcLossDb','LC pair loss dB',.25)+
    input('spliceCount','Splices / route',0,1)+input('spliceLossDb','Splice loss dB',.1)+input('marginDb','Loss margin dB',.5)+input('panelChannels','Panel channel capacity',24,1)+input('panelRU','Panel RU',1,1)+
    input('storageArrayTb','Storage assumed usable TB/array',1000)+input('storageArrayKw','Storage array kW assumption',10)+input('storageArrayRU','Storage array RU assumption',4,1)+
    input('storageSwitchPorts','Storage switch ports assumption',64,1)+input('storageSwitchKw','Storage switch kW assumption',1)+input('storageSwitchRU','Storage switch RU assumption',2,1)+
    input('opsUnitKw','Operations appliance kW assumption',1)+input('opsUnitRU','Operations appliance RU assumption',1,1)+input('cduBlockKw','CDU kW assumption',600)+
    '</div><div id="custom-profile" hidden><h4>Custom server — user assumptions</h4><div class="fields sub"><label for="customName">Server name<input id="customName" value="Custom server"></label>'+
    input('customGpus','GPU/node',8,1)+input('customRU','RU/node',10,1)+input('customKw','Design kW/node',14.3)+input('customLinks','NIC logical links/node (1–8)',8,1)+
    select('customSpeed','Compute link rate',[['400','400G'],['800','800G']])+input('customLogicalPerCage','Logical ports / physical cage',1,1)+input('customCords','AC cords/node',6,1)+'</div></div>';
  by('run').parentElement.insertBefore(box,by('run'));
  by('systemId').insertAdjacentHTML('beforeend','<option value="custom">Custom server (user specifications)</option>');
  const custom=()=>{by('custom-profile').hidden=by('systemId').value!=='custom';if(by('systemId').value==='custom')box.open=true;if(by('systemId').value==='b300'&&Number(by('trunkFiberCount').value)<24){by('trunkFiberCount').value='24';box.open=true;}};by('systemId').addEventListener('change',custom);
  window.getEngineeringInputs=()=>{
    const e={};for(const el of box.querySelectorAll('input,select'))e[el.id]=el.type==='number'?el.value.trim()===''?null:Number(el.value):el.value;
    const mode=by('v65-mode');e.scenario=mode?.value==='colo'?'colo':'ai';
    for(const [to,from] of [['tenantCount','v65-tenants'],['racksPerTenant','v65-racks'],['avgRackKw','v65-density'],['coloReservePct','v65-reserve'],['carrierPaths','v65-carriers']])if(by(from))e[to]=by(from).value.trim()===''?null:Number(by(from).value);
    return e;
  };
  const audit=document.createElement('section');audit.id='engineering-audit';by('kpis').insertAdjacentElement('afterend',audit);
  const escape=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const update=()=>{
    const r=window.DCDesign;if(!r?.usable){audit.innerHTML='';return;}
    const f=r.facility;
    audit.innerHTML='<h3>설계 검증 · 계산 기준 v7.4</h3><p>REVIEW는 계산 가능한 계획안이며, 제품 호환성·시공 조건의 검토가 남아 있다는 뜻입니다.</p>'+
      `<div class="detailGrid"><div class="detail"><b>Total racks / IT</b><span>${r.summary.totalRacks} racks / ${r.summary.totalItPowerKw} kW</span></div><div class="detail"><b>Facility (scenario)</b><span>${f?.facilityPowerKw} kW = IT × PUE ${f?.pue}</span></div></div>`+
      (r.portAudit?`<p>Port/cage conservation: PASS · Server↔Leaf ${r.fabric.serverLinks} · Leaf↔Spine ${r.fabric.totalLeafUplinks} · Spine↔Core ${r.fabric.spineToCoreLinks} links · ${r.fabric.stageCount} stages. Installed GPU ${r.summary.installedGPU} / requested ${r.summary.targetGPU}.</p>`:'')+
      '<ul>'+(r.warnings||[]).map(x=>'<li>'+escape(x)+'</li>').join('')+'</ul>'+
      (r.racks?'<details><summary>실제 랙 배분 / 전력 / PDU</summary><table><thead><tr><th>Rack</th><th>Role</th><th>Equipment</th><th>RU</th><th>kW</th><th>A/B PDU</th></tr></thead><tbody>'+r.racks.map(x=>`<tr><td>${escape(x.id)}</td><td>${escape(x.role)}</td><td>${x.units}</td><td>${x.ru}</td><td>${x.powerKw}</td><td>${x.pduQty}</td></tr>`).join('')+'</tbody></table></details>':'');
  };
  document.addEventListener('dc:design',update);
  document.addEventListener('change',ev=>{
    const id=ev.target?.id;if(!id||ev.target.dataset.referenceOnly)return;
    if(id==='pue'&&by('v60-pue'))by('v60-pue').value=by('pue').value;
    if(id==='v60-pue'){by('pue').value=by('v60-pue').value;}
    if(ev.target.closest('#engineering-inputs')||['v65-mode','v65-tenants','v65-racks','v65-density','v65-reserve','v65-carriers'].includes(id)||ev.target.closest('.requirements')||by(id)?.closest('.fields')) {
      window.DCDesignStale=true;if(by('status'))by('status').textContent='RECALCULATE';
    }
  });
})();
