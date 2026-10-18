/* Additive UI: legacy KO/JA/EN/ZH/DE controls retained. New engineering labels localized. */
(()=>{
'use strict';
const $=id=>document.getElementById(id),E=window.DCEngineering;
const WORDS={ko:['설계 가정 · 즉시 재계산','가정값 · 제조사 데이터시트 확인 필요','입력 방식','GPU 수','IT 전력','NVL72 랙 수','계층 수','광링크 검증','서버–Leaf 연결','계산 근거','액냉 / 공랭','결정 제약','트레이 경로 + 수직 상승 + 여유장','액냉 필수: NVL72는 DLC/Immersion 사용'],en:['Design assumptions · live calculation','Assumptions · verify manufacturer datasheet','Input mode','GPU count','IT power','NVL72 rack count','Fabric tiers','Link budget validation','Server–Leaf attachment','Calculation basis','Liquid / air','Binding constraint','Tray route + vertical rise + slack','Liquid required: NVL72 uses DLC/Immersion'],ja:['設計仮定 · 即時再計算','仮定値 · メーカー仕様書確認が必要','入力方式','GPU数','IT電力','NVL72ラック数','階層数','光リンク検証','サーバー–Leaf接続','計算根拠','液冷 / 空冷','決定制約','トレイ経路 + 垂直立上り + 余長','液冷必須: NVL72はDLC/Immersion'],zh:['设计假设 · 即时重算','假设值 · 需核对厂商数据表','输入模式','GPU数量','IT功率','NVL72机柜数','网络层数','光链路验证','服务器–Leaf连接','计算依据','液冷 / 风冷','限制因素','桥架路径 + 垂直高度 + 余量','NVL72必须使用液冷'],de:['Planungsannahmen · Live-Berechnung','Annahmen · Herstellerdatenblatt prüfen','Eingabemodus','GPU-Anzahl','IT-Leistung','NVL72-Racks','Fabric-Ebenen','Optisches Linkbudget','Server–Leaf-Anbindung','Berechnungsbasis','Flüssigkeit / Luft','Begrenzung','Trassenweg + Steigung + Reserve','NVL72 benötigt Flüssigkeitskühlung']};
const text=k=>(WORDS[document.documentElement.lang]||WORDS.ko)[k],esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const extraKeys=[],fields=document.createElement('div');fields.className='fields sub';
function number(id,label,v,min=0){if($(id))return;extraKeys.push(id);fields.insertAdjacentHTML('beforeend',`<label><span>${label}</span><input id="${id}" type="number" min="${min}" step="any" value="${v}"></label>`);}
function choose(id,label,opts){extraKeys.push(id);fields.insertAdjacentHTML('beforeend',`<label><span>${label}</span><select id="${id}">${opts.map(([v,t])=>`<option value="${v}">${t}</option>`).join('')}</select></label>`);}
function check(id,label,v=false){extraKeys.push(id);fields.insertAdjacentHTML('beforeend',`<label class="check"><input id="${id}" type="checkbox" ${v?'checked':''}><span>${label}</span></label>`);}
$('topology').innerHTML='<option value="rail">Rail-optimized</option><option value="single">Single ToR</option><option value="dual">Dual ToR · split NIC A/B</option>';
const tier=document.createElement('label');tier.innerHTML='<span data-eng-label="6"></span><select id="fabricTier"><option value="2">2-tier</option><option value="3">3-tier · Super-spine</option></select>';$('topology').closest('label').after(tier);extraKeys.push('fabricTier');
const inputMode=document.createElement('label');inputMode.innerHTML='<span data-eng-label="2"></span><select id="designInputMode"><option value="gpu" data-eng-label="3"></option><option value="power" data-eng-label="4"></option><option value="racks" data-eng-label="5"></option></select>';$('targetGPU').closest('label').before(inputMode);
number('itInputMw','IT input MW · reserved headroom applied',100);number('nvlRackCount','NVL72 rack count',4,1);
number('verticalRiseM','Vertical rise m',3);number('p2pPairs','P2P connector pairs',2);check('lowLoss','Low-loss MPO · 0.25 dB');choose('connectorPolish','Connector polish',[['auto','Profile auto · confirm SKU'],['APC','APC'],['UPC','UPC']]);check('twinPort400','Qualified 2×400G switch module sharing');
choose('serverLogicalPerModule','Server logical 400G ports / physical module',[['1','1 · generic'],['2','2 · only with qualified twin-port module']]);
const opticChoices=[['auto','Generic planning · SKU unverified'],...Object.entries(E.OPTIC_SKUS).map(([id,p])=>[id,p.name+' · '+p.reach+'m · '+p.connector+'/'+p.polish])];
choose('serverOpticSku','Server ↔ Leaf optic at one endpoint',opticChoices);
choose('uplinkOpticSku','Leaf ↔ Spine optic at one endpoint',opticChoices);
choose('coreOpticSku','Spine ↔ Core optic at one endpoint',opticChoices);
choose('mpoEndGender','MPO cable at transceiver',[['review','Pinning / gender not confirmed'],['female','Female · unpinned'],['male','Male · pinned']]);
number('upsEfficiency','UPS efficiency η',.96);number('upsBlockKva','UPS module kVA',1000);number('liquidPct','DLC / rear-door liquid heat %',75);number('waterDensity','Fluid density kg/m³',1000);number('waterCp','Fluid cp J/(kg·K)',4186);number('deltaTK','Liquid ΔT K',10);choose('cduRedundancy','CDU redundancy',[['N+1','N+1'],['N','N'],['2N','2N']]);number('airLimitKw','Air planning limit kW/rack',30);number('rearDoorLimitKw','Rear-door planning limit kW/rack',50);
choose('storageEncoding','Storage protection',[['ec','EC 8+2 · usable 80%'],['replica3','3 replicas · usable 33.3%']]);number('storageReadGbpsPerGpu','Read GB/s per GPU',0);number('storageRawTbPerNode','Raw TB / storage node',1000);number('storageReadGbpsPerNode','Read GB/s / storage node',50);number('cassetteChannels','MPO cassette channels',12);number('consolePorts','Console ports/device',32);number('trayWidthMm','Tray width mm',300);number('trayDepthMm','Tray depth mm',50);number('cableOdMm','Cable OD mm',5);number('trayRouteM','Measured tray route m · 0=RFQ',0);
const reason=document.createElement('p');reason.id='nvl72-cooling-reason';reason.className='muted';reason.hidden=true;$('coolingMode').closest('label').append(reason);
const opts=$('storageProtocol');opts.insertAdjacentHTML('beforeend','<option>NVMe-oF/RoCE</option>');for(const o of opts.options)if(o.value==='Fibre Channel')o.textContent='Fibre Channel · legacy SAN';
for(const key of ['gb200','gb300'])$('systemId').insertAdjacentHTML('beforeend',`<option value="${key}">${key.toUpperCase()} NVL72 · rack-scale</option>`);
const box=document.createElement('section');box.id='engineering-v82-inputs';box.className='inputGroup';box.innerHTML='<h3 data-eng-label="0"></h3><p data-eng-label="1"></p>';box.append(fields);$('engineering-inputs').prepend(box);
const profile=document.createElement('details');profile.open=false;profile.id='engineering-system-table';profile.innerHTML='<summary>System assumptions · GPU / RU / kW / NIC / speed</summary><div class="tableWrap"><table><thead><tr><th>System</th><th>GPU/unit</th><th>RU/unit</th><th>kW/unit</th><th>NIC ports/unit</th><th>Gbit/s</th></tr></thead><tbody>'+Object.entries(E.SYSTEMS).map(([key,p])=>'<tr><td>'+key.toUpperCase()+'</td>'+['gpus','ru','power','links','linkSpeed'].map(k=>`<td><input type="number" step="any" min="1" data-system="${key}" data-param="${k}" value="${p[k]}" aria-label="${key} ${k}"></td>`).join('')+'</tr>').join('')+'</tbody></table></div>';$('engineering-inputs').prepend(profile);
const opticTable=document.createElement('details');opticTable.innerHTML='<summary>Optics planning table · reach / loss · vendor confirmation required</summary><div class="tableWrap"><table><thead><tr><th>Optics</th><th>Fiber / connector</th><th>Reach m</th><th>Budget dB</th></tr></thead><tbody>'+E.OPTICS.map(p=>`<tr><td>${p.standard||p.speed+'G '+p.id.toUpperCase()}</td><td>${p.fiber} · ${p.connector} ${p.polish||''}</td><td><input type="number" min="0" data-optic="${p.speed}-${p.id}" data-param="reach" value="${p.reach}"></td><td>${p.loss===null?'Integrated cable':`<input type="number" min="0" step="any" data-optic="${p.speed}-${p.id}" data-param="loss" value="${p.loss}">`}</td></tr>`).join('')+'</tbody></table></div>';$('engineering-inputs').prepend(opticTable);
const audit=document.createElement('section');audit.id='engineering-v82-results';audit.className='panel';document.querySelector('.resultPanel').append(audit);
const oldGet=window.getEngineeringInputs;
window.getEngineeringInputs=()=>{
 const i=oldGet();for(const id of extraKeys){const e=$(id);if(e)i[id]=e.type==='checkbox'?e.checked:e.type==='number'?Number(e.value):e.value;}
 i.inputMode=$('designInputMode').value;i.systemAssumptions={};i.opticsAssumptions={};
 profile.querySelectorAll('input').forEach(e=>{(i.systemAssumptions[e.dataset.system]??={})[e.dataset.param]=Number(e.value);});
 opticTable.querySelectorAll('input').forEach(e=>{(i.opticsAssumptions[e.dataset.optic]??={})[e.dataset.param]=Number(e.value);});
 // Capture mounted legacy fields too. The one assumption panel is the source of truth.
 for(const id of ['oversub','slackPct','matedPairs','fiberPolarity','upsUtilPct','generatorUtilPct','transformerUtilPct'])if($(id))i[id]=$(id).type==='number'?Number($(id).value):$(id).value;
 return i;
};
function sync(){
 const nvl=['gb200','gb300'].includes($('systemId').value);
 for(const o of $('coolingMode').options)o.disabled=nvl&&!['dlc','immersion'].includes(o.value);
 if(nvl&&!['dlc','immersion'].includes($('coolingMode').value))$('coolingMode').value='dlc';
 const reason=$('nvl72-cooling-reason');if(reason){reason.hidden=!nvl;reason.textContent=text(13);}
 const mode=$('designInputMode').value;
 $('targetGPU').closest('label').hidden=mode!=='gpu';$('itInputMw').closest('label').hidden=mode!=='power';$('nvlRackCount').closest('label').hidden=mode!=='racks';
 $('designInputMode').querySelector('[value="racks"]').disabled=!nvl;
 if(!nvl&&mode==='racks'){$('designInputMode').value='gpu';sync();}
 profile.querySelectorAll('tr').forEach(row=>{const e=row.querySelector('input');if(e)row.hidden=e.dataset.system!==$('systemId').value;});
 document.querySelectorAll('[data-eng-label]').forEach(e=>e.textContent=text(Number(e.dataset.engLabel)));
 $('topology').closest('label').querySelector('span').textContent=text(8);
}
let pending=0;
function schedule(){clearTimeout(pending);pending=setTimeout(()=>{if(typeof run==='function')run();},220);}
document.addEventListener('input',e=>{if(e.target.closest('#engineering-inputs,.formPanel')&&!e.target.dataset.referenceOnly){sync();schedule();}});
document.addEventListener('change',e=>{if(e.target.closest('#engineering-inputs,.formPanel')&&!e.target.dataset.referenceOnly){sync();schedule();}});
new MutationObserver(()=>{sync();if(window.DCDesign)paint(window.DCDesign);}).observe(document.documentElement,{attributes:true,attributeFilter:['lang']});
function paint(r){if(!r?.usable||r.version!=='8.2.0-engineering'){audit.innerHTML='';return;}
 audit.innerHTML='<h3>'+text(7)+'</h3><p>SKU 선택은 해당 모듈의 거리·매체·연마 방식을 제한합니다. 상대편 모듈, FEC, 광손실 예산, 케이블 극성은 별도 확인이 필요합니다.</p><div class="tableWrap"><table><thead><tr><th>Link</th><th>Optics / endpoint SKU</th><th>Length m</th><th>Active / channel / trunk F</th><th>'+text(9)+'</th><th>Result</th></tr></thead><tbody>'+Object.entries(r.optical).filter(([,o])=>o.links).map(([seg,o])=>'<tr><td>'+seg+'</td><td>'+esc(o.profile.media)+' · '+esc(o.profile.opticSku||'Generic / RFQ')+' · '+esc(o.profile.connector)+'/'+esc(o.profile.polish||'-')+'</td><td>'+o.lengthM+'</td><td>'+o.profile.fibers+' / '+o.profile.installedChannelFibers+' / '+o.profile.trunkFiberCount+(o.profile.unusedTrunkFibers?' (unused '+o.profile.unusedTrunkFibers+')':'')+'</td><td>'+o.loss.attenDbKm+'×'+o.lengthM+'/1000 + '+o.loss.matedPairs+'×'+o.loss.connectorDb+' = '+o.loss.estimatedDb+' dB; assumed budget '+o.loss.budgetDb+'; margin '+o.loss.marginDb+'</td><td>'+(!o.loss.applicable?'Cable reach: ':'')+(o.loss.pass?'PRELIMINARY / RFQ':'FAIL')+'</td></tr>').join('')+'</tbody></table></div><p>'+text(11)+': '+esc(r.summary.bindingConstraint.join(' + '))+' · '+r.summary.systemsPerRack+' / rack · '+JSON.stringify(r.summary.rackLimits)+'</p><p>'+text(10)+': '+r.cooling.liquidKw+' / '+r.cooling.airKw+' kW · '+r.cooling.flowLMin+' L/min · CDU '+r.cooling.cduDuty+' duty / '+r.cooling.cduQty+' installed</p><p>UPS '+r.facility.upsDesignKva+' kVA · Generator '+r.facility.generatorDesignKw+' kW · Transformer '+r.facility.transformerDesignKva+' kVA</p><p>'+text(12)+': (route + '+r.input.verticalRiseM+' m + 2×'+r.input.terminationM+' m) × '+(1+r.input.slackPct/100)+' → standard length</p><p>Independent domains: '+r.fabric.domains.length+' · R='+r.input.switchRadix+' · 2-tier R²/2='+r.fabric.domains[0].twoLimit+' / 3-tier R³/4='+r.fabric.domains[0].threeLimit+' NIC endpoints/domain</p>';
}
document.addEventListener('dc:design',e=>paint(e.detail));
// Reuse the same engine for click, live update, sensitivity and exports; colocation remains on existing API.
const nativeFetch=window.fetch.bind(window);
window.fetch=async(input,init={})=>{
 const url=new URL(typeof input==='string'?input:input.url,location.href);
 if(url.pathname.endsWith('/api/design')){
  const raw=JSON.parse(init.body||'{}');if(raw.scenario!=='colo'){
   try{const r=E.calculate(raw);r.region={name:raw.regionId||'Project',voltage:raw.siteVoltage,hz:raw.siteFrequency,phases:raw.sitePhases,framework:raw.regionalFramework,pduRatedKw:r.input.pduRatedKw,status:'Planning assumption; verify site specifications'};return new Response(JSON.stringify(r),{headers:{'Content-Type':'application/json'}});}catch(e){return new Response(JSON.stringify({usable:false,status:'INVALID',error:e.message,input:raw}),{status:400,headers:{'Content-Type':'application/json'}});}
  }
 }
 return nativeFetch(input,init);
};
function finishMount(){
 // Legacy UI adds labels later. Keep all assumptions editable and visible in every result tab.
 const panel=$('engineering-inputs');panel.open=false;const advanced=panel.closest('details.advanced');if(advanced)advanced.open=false;
 const ownSummary=panel.querySelector(':scope > summary');if(ownSummary){ownSummary.textContent=text(0);panel.prepend(ownSummary);}
 for(const [id,val]of [['slackPct',15],['matedPairs',4],['marginDb',0],['fiberPolarity','B']])if($(id))$(id).value=val;
 if($('upsBlockKw'))$('upsBlockKw').closest('label').hidden=true;
 if($('storageArrayTb'))$('storageArrayTb').closest('label').hidden=true;
 // Oversubscription is one of the two topology axes' companion inputs, not buried in legacy advanced fields.
 $('oversub').closest('label').querySelector('span').textContent='Oversubscription · x:1';
 if($('targetPodEgress')){$('targetPodEgress').closest('label').hidden=true;$('targetPodEgress').disabled=true;}
 if($('powerFactor'))$('powerFactor').closest('label').querySelector('span').textContent='AC load power factor · UPS / Transformer';
 sync();document.dispatchEvent(new CustomEvent('dc:inputs-ready'));schedule();
}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',()=>requestAnimationFrame(()=>requestAnimationFrame(()=>requestAnimationFrame(()=>requestAnimationFrame(finishMount)))));else finishMount();sync();
})();

