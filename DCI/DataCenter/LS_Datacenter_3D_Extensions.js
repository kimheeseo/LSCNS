/* v4.1 · 2026-10-10: CONFIG 기반 IT/시설 전력·PUE·가용성 계산 */
(function(){
'use strict';const T=window.__LS3D_TEST__;if(!T)return;const $=x=>document.getElementById(x);
const defs={
 aidc:{rackCount:20,rackPowerKw:60,gpuRatio:80,cooling:'mixed',air:20,rear:20,dlc:60,redundancy:'N+1',targetPue:1.30,pueMin:1.10,pueMax:1.60,outdoorC:24,coolingKw:1600,generatorKw:1800,upsKw:1600,batteryMin:12,upsEfficiencyPct:96,distributionLossPct:2.5,coolingCop:4,otherPct:4,portCount:128,plugKw:18,cpoKw:14},
 'server-room':{rackCount:3,rackPowerKw:6.5,gpuRatio:0,cooling:'air',air:100,rear:0,dlc:0,redundancy:'N',targetPue:1.45,pueMin:1.10,pueMax:2,outdoorC:24,coolingKw:35,generatorKw:0,upsKw:30,batteryMin:15,upsEfficiencyPct:94,distributionLossPct:3,coolingCop:3.5,otherPct:6,portCount:32,plugKw:.8,cpoKw:.6},
 onprem:{rackCount:12,rackPowerKw:8,gpuRatio:15,cooling:'rear',air:60,rear:40,dlc:0,redundancy:'N+1',targetPue:1.40,pueMin:1.10,pueMax:2,outdoorC:24,coolingKw:220,generatorKw:600,upsKw:500,batteryMin:10,upsEfficiencyPct:95,distributionLossPct:2.8,coolingCop:3.8,otherPct:5,portCount:64,plugKw:2.5,cpoKw:1.8}
};
const CONFIG=window.LS3D_CONFIG={rackCount:0,rackPowerKw:0,gpuRatio:0,cooling:'air',air:100,rear:0,dlc:0,redundancy:'N+1',targetPue:1.3,pueMin:1.1,pueMax:2,outdoorC:24,coolingKw:0,generatorKw:0,upsKw:0,batteryMin:12,upsEfficiencyPct:96,distributionLossPct:2.5,coolingCop:4,otherPct:4,mode:'pluggable',plugW:8,cpoW:4,plugKw:18,cpoKw:14,portCount:128,cpoPue:0,density:1.35,selected:'tor-spine-a',broken:null,reroute:false,edited:false,quality:'medium',baselineRackInletC:22,availabilityModel:{normal:99.99,power:{N:96,'N+1':99.5,'2N':99.95},cooling:{N:93,'N+1':98.5,'2N':99.9},network:{N:85,'N+1':98,'2N':99.9},other:{N:94,'N+1':98.5,'2N':99.9}}};
const c=CONFIG;
function preset(id){Object.assign(c,defs[id]||defs.aidc);c.broken=null;c.reroute=false;c.edited=false}
preset(T.facilityMode);
const LOSS={smfDbKm:.22,mmfDbKm:3,connectorDb:.25,spliceDb:.10,cpoCouplingDb:.15};
const nodes=[['server','GPU NIC',[3,0],'gpu-0-0'],['torA','ToR / Leaf A',[15,0],'tor-a'],['torB','ToR / Leaf B',[22,5],'tor-b'],['spA','Spine A',[8,18],'spine-a'],['spB','Spine B',[22,18],'spine-b'],['core','Core',[15,28],'core'],['odf','ODF / MMR',[-5,28],'odf'],['isp','Campus / ISP',[-22,28],'carrier']].map(x=>({id:x[0],name:x[1],pos:x[2],asset:x[3]}));
const links=[
['server-tor','server','torA','DAC/AEC','MMF',0,'QSFP electrical','400G',5,2,0,0,'A'],
['server-tor-b','server','torB','DAC/AEC','MMF',0,'QSFP electrical','400G',7,2,0,0,'B'],
['tor-spine-a','torA','spA','SMF trunk','SMF',24,'MPO-16','800G assumed',34,4,2,4,'A'],
['tor-spine-a-b','torA','spB','SMF redundant','SMF',24,'MPO-16','800G assumed',42,4,2,4,'B'],
['tor-spine-b','torB','spB','SMF trunk','SMF',24,'MPO-16','800G assumed',36,4,2,4,'B'],
['tor-spine-b-a','torB','spA','SMF redundant','SMF',24,'MPO-16','800G assumed',44,4,2,4,'A'],
['spine-core-a','spA','core','SMF trunk','SMF',24,'MPO-16','800G assumed',48,4,2,4,'A'],
['spine-core-b','spB','core','SMF redundant','SMF',24,'MPO-16','800G assumed',54,4,2,4,'B'],
['core-odf','core','odf','SMF patch','SMF',48,'LC duplex','400/800G assumed',62,4,2,4.5,'A'],
['odf-isp','odf','isp','Carrier cross-connect','SMF',96,'MPO/LC','400G assumed',85,4,4,5,'A'],
['campus-dci','isp','core','Campus DCI SMF','SMF',24,'LC duplex','400G assumed',950,4,12,6,'B']
].map(x=>({id:x[0],from:x[1],to:x[2],type:x[3],media:x[4],cores:x[5],connector:x[6],speed:x[7],length:x[8],conn:x[9],splice:x[10],budget:x[11],route:x[12],down:false}));
const N=id=>nodes.find(x=>x.id===id),L=id=>links.find(x=>x.id===id),A=id=>T.assets.find(x=>x.id===id),esc=x=>String(x).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
function loss(l){let alpha=l.media==='MMF'?LOSS.mmfDbKm:LOSS.smfDbKm,val=l.length/1000*alpha+l.conn*LOSS.connectorDb+l.splice*LOSS.spliceDb+(c.mode==='cpo'&&l.media==='SMF'?LOSS.cpoCouplingDb:0);return [val,l.budget?l.budget-val:null,alpha]}
function event(sev,msg,l){if(T.addEventLog)T.addEventLog(sev,msg,A(N(l?.from)?.asset)||A('odf')||T.selected,'fiber')}
function flow(){T.flowPaths.fiber.splice(0,T.flowPaths.fiber.length,...links.filter(x=>!x.down).map(x=>[N(x.from).pos,N(x.to).pos]));let e=$('optState');if(e)e.textContent=c.broken?(c.reroute?'우회 경로 · 지연 +18 µs · 가용 대역폭 65%':'연결 끊김 · 대체 경로 없음'):'정상 경로 · 교육용 가정값'}
function calc(){
 const rackLoad=c.rackCount*c.rackPowerKw;
 const opticalPower=(c.mode==='cpo'?c.cpoW:c.plugW)*c.portCount/1000+(c.mode==='cpo'?c.cpoKw:c.plugKw);
 let load=rackLoad+opticalPower;
 const elapsed=T.scenarioElapsed?T.scenarioElapsed():0,sc=T.scenario;
 if(sc==='gpu-surge')load*=1+.55*Math.min(1,elapsed/180);
 if(sc==='dlc-leak'&&elapsed>45)load*=.75;
 if(sc==='rack-hotspot'&&elapsed>100)load*=.85;
 if(sc==='ups-battery'&&elapsed>=c.batteryMin*60)load*=.2;
 const mixTotal=Math.max(1,c.air+c.rear+c.dlc);
 const coolingMultiplier=(c.air*1+c.rear*1.08+c.dlc*1.28)/mixTotal;
 const outdoorDerate=Math.max(.55,1-Math.max(0,c.outdoorC-20)*.01);
 const coolingLoadRatio=load/Math.max(1,c.coolingKw);
 const loadDerate=Math.max(.75,1-Math.max(0,coolingLoadRatio-.6)*.12);
 const effectiveCop=Math.max(1,c.coolingCop*coolingMultiplier*outdoorDerate*loadDerate);
 const coolingPower=load/effectiveCop;
 const upsLoss=load*(1/Math.max(.8,Math.min(.999,c.upsEfficiencyPct/100))-1);
 const distributionLoss=(load+coolingPower+upsLoss)*c.distributionLossPct/100;
 const otherPower=load*c.otherPct/100;
 let facility=load+coolingPower+upsLoss+distributionLoss+otherPower;
 const pue=load>0?facility/load:0;
 let temperature=c.baselineRackInletC+Math.max(0,load/Math.max(1,c.coolingKw)-.7)*4+Math.max(0,c.outdoorC-24)*.08;
 if(sc==='fire')temperature+=13;
 if(sc==='dlc-leak')temperature+=9;
 if(sc==='rack-hotspot')temperature+=Math.min(17,elapsed/10);
 if(sc==='cooling')temperature+=Math.min(16,elapsed/70);
 const availabilityClass=c.broken?'network':sc==='power'||sc==='ups-battery'?'power':sc==='cooling'||sc==='dlc-leak'||sc==='rack-hotspot'?'cooling':sc==='network'||sc==='fiber-cut'?'network':sc==='fire'?'other':'normal';
 let availability=c.availabilityModel[availabilityClass]?.[c.redundancy]??c.availabilityModel.normal;
 if(c.broken&&!c.reroute)availability=0;
 if(sc==='ups-battery'&&elapsed>=c.batteryMin*60)availability=0;
 if(sc==='rack-hotspot'&&temperature>=30)availability=Math.max(0,availability-3);
 if(sc==='cooling'&&temperature>=32)availability=Math.max(0,availability-5);
 const m={load,facility,pue,temperature,availability,latency:c.broken?(c.reroute?18:999):sc==='fiber-cut'?18:0,bandwidth:c.broken?(c.reroute?65:0):sc==='fiber-cut'?65:100,rackLoad,opticalPower,coolingPower,effectiveCop,upsLoss,distributionLoss,otherPower};
 const set=(id,value)=>{const e=$(id);if(e)e.innerHTML=value};
 set('kpiLoad',load.toFixed(0)+'<small>kW</small>');
 set('kpiFacility',facility.toFixed(0)+'<small>kW</small>');
 const fm=$('kpiFacilityMeta');if(fm)fm.textContent='계산 PUE '+pue.toFixed(2)+' · 가정값';
 set('kpiAvailability',availability.toFixed(2)+'<small>%</small>');
 set('kpiTemp',temperature.toFixed(1)+'<small>°C</small>');
 if($('thermalOut'))$('thermalOut').textContent='평균 랙 입구 '+temperature.toFixed(1)+'°C · 냉각 용량 부하율 '+(load/Math.max(1,c.coolingKw)*100).toFixed(1)+'% · 외기 '+c.outdoorC.toFixed(1)+'°C (가정)';
 const badMix=Math.abs(c.air+c.rear+c.dlc-100)>.01;
 if($('warn'))$('warn').textContent=(c.pueMin>c.pueMax?'경고: 목표 PUE 하한은 상한 이하여야 합니다. ':badMix?'경고: 냉각 방식 비율 합계를 100%로 맞춰주세요. ':load>c.coolingKw?'경고: IT 부하가 냉각 용량을 초과합니다. ':c.generatorKw<load?'주의: 발전기 가정 용량이 IT 부하 미만입니다. ':c.upsKw<load?'주의: UPS 가정 용량이 IT 부하 미만입니다. ':'')+'계산값은 교육·설계 검토용 가정입니다.';
 if($('cpoOut'))$('cpoOut').textContent='플러거블 '+(c.portCount*c.plugW/1000+c.plugKw).toFixed(1)+' kW · CPO '+(c.portCount*c.cpoW/1000+c.cpoKw).toFixed(1)+' kW · 광 IT 부하 비중 '+(opticalPower/Math.max(1,load)*100).toFixed(2)+'% · 포트 밀도 '+(c.mode==='cpo'?c.density:1).toFixed(2)+'× (가정)';
 const rows=[['랙 IT 부하','랙 수 × 랙당 IT 전력',rackLoad.toFixed(1)+' kW'],['광 네트워크 IT 부하','광 포트 전력 + 스위치 전력',opticalPower.toFixed(1)+' kW'],['IT 부하 합계','랙 IT + 광 네트워크',load.toFixed(1)+' kW'],['유효 냉각 COP','기본 COP × 방식 계수 × 외기·부하 보정',effectiveCop.toFixed(2)],['냉각 전력','IT 부하 ÷ 유효 COP',coolingPower.toFixed(1)+' kW'],['UPS 손실','IT 부하 × (UPS 효율⁻¹ − 1)',upsLoss.toFixed(1)+' kW'],['배전 손실','(IT + 냉각 + UPS 손실) × 손실률',distributionLoss.toFixed(1)+' kW'],['기타 전력','IT 부하 × 기타 전력률',otherPower.toFixed(1)+' kW'],['시설 전력','IT + 냉각 + UPS/배전 손실 + 기타',facility.toFixed(1)+' kW'],['PUE','시설 전력 ÷ IT 부하',pue.toFixed(3)],['가용성','장애 종류 × 이중화 등급 가정',availability.toFixed(2)+'%']];
 const tbody=$('basisRows');if(tbody)tbody.innerHTML=rows.map(r=>'<tr><th>'+r[0]+'</th><td>'+r[1]+'</td><td>'+r[2]+'</td></tr>').join('');
 c.calculate=calc;
 return m
}
function form(){let vals=[['rackCount','랙 수',1,500],['rackPowerKw','랙당 IT 전력 kW',1,200],['gpuRatio','GPU 서버 비율 %',0,100],['air','공랭 비율 %',0,100],['rear','후면열교환기 비율 %',0,100],['dlc','DLC 비율 %',0,100],['targetPue','목표 PUE (참고)',1,3],['pueMin','목표 PUE 하한',1,3],['pueMax','목표 PUE 상한',1,3],['outdoorC','외기 온도 °C',-20,50],['coolingKw','냉각 용량 kW',1,1000000],['coolingCop','기본 냉각 COP',1,12],['upsEfficiencyPct','UPS 효율 %',80,99.9],['distributionLossPct','배전 손실률 %',0,15],['otherPct','기타 전력률 %',0,20],['generatorKw','발전기 용량 kW',0,1000000],['upsKw','UPS 용량 kW',0,1000000],['batteryMin','UPS 백업 분',1,180],['portCount','광 포트 수',1,100000],['plugW','플러거블 포트 W',.1,100],['cpoW','CPO 포트 W',.1,100],['plugKw','플러거블 스위치 kW',0,100000],['cpoKw','CPO 스위치 kW',0,100000],['density','CPO 포트 밀도 배율',.1,10]];$('cfgFields').innerHTML=vals.map(x=>'<label>'+x[1]+'<input data-k="'+x[0]+'" type="number" min="'+x[2]+'" max="'+x[3]+'" value="'+c[x[0]]+'"></label>').join()+'<label>냉각 방식 (가정)<select data-k="cooling"><option value="air">공랭 (Air)</option><option value="rear">후면열교환기 (Rear Door)</option><option value="dlc">직접액체냉각 (DLC)</option><option value="mixed">하이브리드 (Hybrid)</option></select></label><label>이중화<select data-k="redundancy"><option>N</option><option>N+1</option><option>2N</option></select></label><label>CPO 비교<select data-k="mode"><option value="pluggable">플러거블 광모듈</option><option value="cpo">CPO</option></select></label><label>렌더 품질<select data-k="quality"><option value="low">품질 낮음</option><option value="medium">품질 보통</option><option value="high">품질 높음</option></select></label>';$('cfgFields').querySelectorAll('[data-k]').forEach(e=>{e.value=c[e.dataset.k];e.oninput=()=>{let v=e.type==='number'?+e.value:e.value;if(e.type==='number'&&(v<+e.min||v>+e.max))return;c[e.dataset.k]=v;c.edited=true;if(e.dataset.k==='redundancy'){$('redundancySelect').value=v;$('redundancySelect').dispatchEvent(new Event('change',{bubbles:true}));T.redundancy=v}if(e.dataset.k==='batteryMin')battery();if(e.dataset.k==='mode')optics();if(e.dataset.k==='rackCount'||e.dataset.k==='rackPowerKw')layout();calc()}});$('redundancySelect').value=c.redundancy;$('redundancySelect').dispatchEvent(new Event('change',{bubbles:true}))}
function build(){let loading=document.createElement('div');loading.textContent='3D 데이터센터 씬 초기화 중 · 가정값 모델 로딩';loading.style='position:fixed;z-index:100;inset:0;display:grid;place-items:center;background:#071523;color:#d8f4ff;font:700 12px system-ui';document.body.append(loading);setTimeout(()=>loading.remove(),650);let st=document.createElement('style');st.textContent='.twin-modal{display:none;position:fixed;z-index:80;inset:0;background:#020b14d9;justify-content:flex-end}.twin-modal.open{display:flex}.twin-panel{width:min(520px,100vw);height:100dvh;overflow:auto;background:#0c1d2d;padding:16px;color:#e5f7ff;border-left:1px solid #345}.twin-panel h2{font-size:17px}.twin-panel p,.twin-panel small{font-size:10px;line-height:1.6;color:#9fb8c8}.twin-panel label{display:block;font-size:10px;color:#adc3d0}.twin-panel input,.twin-panel select{width:100%;padding:7px;margin:4px 0 8px;background:#081928;color:white;border:1px solid #345;border-radius:6px}.calc-basis{margin:10px 0;padding:9px;background:#10283a;border:1px solid #35536a;border-radius:8px;font-size:10px}.calc-basis summary{cursor:pointer;font-weight:700;color:#9fe8df}.table-scroll{max-width:100%;overflow-x:auto}.calc-basis table{width:100%;border-collapse:collapse;margin:8px 0}.calc-basis th,.calc-basis td{padding:5px;border-bottom:1px solid #294559;text-align:left;vertical-align:top}.calc-basis td:last-child{white-space:nowrap;text-align:right}.calc-basis small{display:block;line-height:1.5}.top-actions>span{display:flex;gap:4px}.top-actions>span button{white-space:nowrap}.twin-panel button{background:#153d50;color:#e6ffff;border:1px solid #398397;border-radius:7px;padding:8px;margin:3px;font-size:10px}.twin-grid{display:grid;grid-template-columns:1fr 1fr;gap:7px}.twin-readout{padding:8px;background:#132c40;border-radius:7px;font-size:10px}.optical-svg{width:100%;background:#081928;border:1px solid #345}.optical-svg path{fill:none;stroke:#55ddc4;stroke-width:4;cursor:pointer}.optical-svg path.alt{stroke:#7aa8ff;stroke-dasharray:5 4}.optical-svg path.bad{stroke:#ff7086}.optical-svg rect{fill:#173950;stroke:#7aa}.optical-svg text{fill:white;font:9px system-ui}#optList button{display:block;width:100%;text-align:left}@media(max-width:420px){.top-actions>span button{font-size:8px;padding:5px 6px}.top-actions{gap:3px}}@media(max-width:790px){.twin-modal{align-items:flex-end}.twin-panel{height:90dvh;width:100%;border-radius:15px 15px 0 0}}';document.head.append(st);
let bar=document.querySelector('.top-actions'),b=document.createElement('span');b.innerHTML='<button id="cfgOpen">⚙ 설정</button><button id="optOpen">◎ 광</button><button id="guideOpen">? 안내</button>';bar.prepend(b);
document.body.insertAdjacentHTML('beforeend','<div class="twin-modal" id="twinSettings"><section class="twin-panel"><h2>AI 데이터센터 가정값</h2><button data-close>닫기</button><p>교육·설계 검토용 값 · 변경 즉시 계산 반영</p><div><button data-p="server-room">소규모 서버실</button><button data-p="onprem">엔터프라이즈</button><button data-p="aidc">AI 학습 캠퍼스</button></div><div id="cfgFields" class="twin-grid"></div><details class="calc-basis" open><summary>계산 근거 보기 · 수식과 중간값</summary><div class="table-scroll"><table><thead><tr><th>항목</th><th>계산식</th><th>결과</th></tr></thead><tbody id="basisRows"></tbody></table></div><small>외기 20°C 초과 시 COP를 °C당 1% 낮추고, 냉각 방식·부하 보정 계수는 개념 모델 가정값으로 적용합니다. 목표 PUE는 비교용 입력값이며 KPI는 계산 결과를 표시합니다.</small></details><p id="warn"></p><p id="thermalOut"></p><b>CPO / 플러거블 비교</b><p id="cpoOut"></p><p>시설전력=(IT+광 네트워크)×PUE. 냉각 유량=열부하/(4.186×ΔT)×60. 외기 20°C 초과분 1°C당 PUE +0.012 가정.</p><button id="jsonOut">JSON 내보내기</button><button id="jsonIn">JSON 불러오기</button><input id="jsonFile" type="file" accept=".json" hidden><button id="bomOut">BOM 툴로 전달 ↗</button><button id="reset">초기값 복원</button></section></div><div class="twin-modal" id="optModal"><section class="twin-panel"><h2>광 네트워크 토폴로지</h2><button data-close>닫기</button><p>Rack NIC → ToR/Leaf → Spine → Core → ODF/MMR → ISP · 가정값</p><div id="optState"></div><svg id="optSvg" class="optical-svg" viewBox="0 0 650 145"></svg><div id="optList"></div><div id="optDetail"></div></section></div><div class="twin-modal" id="guide"><section class="twin-panel"><h2>빠른 안내 · 5단계</h2><p>1 구역 선택 → 2 설비 클릭 → 3 전력·냉각·광 흐름 토글 → 4 장애 시나리오와 링크 직접 단선 → 5 작업자 시점 보기.</p><p>PC: 드래그 회전·휠 줌·우클릭 이동. 모바일: 한 손가락 회전·두 손가락 이동/줌. 키보드 1–7 구역 이동, ESC 닫기.</p><button data-close>닫기</button></section></div>');
$('cfgOpen').onclick=()=>{form();$('twinSettings').classList.add('open')};$('optOpen').onclick=()=>{$('optModal').classList.add('open');optics()};$('guideOpen').onclick=()=>$('guide').classList.add('open');document.querySelectorAll('[data-close]').forEach(x=>x.onclick=()=>x.closest('.twin-modal').classList.remove('open'));
document.querySelectorAll('[data-p]').forEach(x=>x.onclick=()=>{T.applyFacilityMode(x.dataset.p);preset(x.dataset.p);layout();form();calc()});
$('jsonOut').onclick=saveJson;$('jsonIn').onclick=()=>$('jsonFile').click();$('jsonFile').onchange=e=>{let f=e.target.files[0];if(f){let r=new FileReader();r.onload=()=>{try{let o=JSON.parse(r.result);if(o.schemaVersion!=='lsdc-twin-bom/1.0')throw Error('schema mismatch');Object.assign(c,o.facility,{cooling:o.facility.cooling.mode,air:o.facility.cooling.airPct,rear:o.facility.cooling.rearDoorPct,dlc:o.facility.cooling.dlcPct,gpuRatio:o.facility.gpuServerRatio*100,pueMin:o.facility.pueRange?.min||c.pueMin,pueMax:o.facility.pueRange?.max||c.pueMax,edited:true});if(o.optical){c.mode=o.optical.cpoMode||c.mode;Object.assign(c,{plugW:o.optical.assumptions.pluggablePortW||c.plugW,cpoW:o.optical.assumptions.cpoPortW||c.cpoW,plugKw:o.optical.assumptions.pluggableSwitchKw||c.plugKw,cpoKw:o.optical.assumptions.cpoSwitchKw||c.cpoKw,portCount:o.optical.assumptions.portCount||c.portCount});(o.optical.topology?.links||[]).forEach(q=>{let l=L(q.id);if(l){l.length=q.lengthM??l.length;l.conn=q.connectors??l.conn;l.splice=q.splices??l.splice;l.down=!!q.down}})}form();calc();optics()}catch(z){alert('JSON 오류: '+z.message)}};r.readAsText(f)}};
$('bomOut').onclick=()=>{saveJson();window.open('./index.html','_blank','noopener')};$('reset').onclick=()=>{preset(T.facilityMode);form();calc()};$('redundancySelect').addEventListener('change',e=>{c.redundancy=e.target.value;flow()});
$('facilityModeSelect').onchange=e=>{T.applyFacilityMode(e.target.value);preset(e.target.value);form();calc()};
document.addEventListener('keydown',e=>{if(e.key==='Escape')document.querySelectorAll('.twin-modal').forEach(x=>x.classList.remove('open'));if(/^[1-7]$/.test(e.key)&&!['INPUT','SELECT'].includes(document.activeElement.tagName)){let z=T.visibleZones()[+e.key-1];if(z)T.setZone(z.id)}});
document.addEventListener('visibilitychange',()=>{if(document.hidden)T.simRunning=false});
const oldScenario=T.scenarioApply;T.scenarioApply=function(id){oldScenario(id);if(id==='normal'&&c.broken){let l=L(c.broken);if(l)l.down=false;c.broken=null;c.reroute=false;flow();optics()}};
T.metrics=function(){return calc()};
let up=T.scenarioDefs.find(x=>x.id==='ups-battery');const battery=()=>{if(up){let t=Math.max(60,c.batteryMin*60);up.timeline[2][0]=t/2;up.timeline[3][0]=t;up.timeline[4][0]=t+120}};
battery();
let tempAlarm=false,capacityAlarm=false;setInterval(()=>{let live=calc();if(live.temperature>27&&!tempAlarm){event('WARN','랙 입구 온도 가정값 27°C 초과',L(c.selected));tempAlarm=true}if(live.temperature<=27)tempAlarm=false;if(c.rackCount*c.rackPowerKw>c.coolingKw&&!capacityAlarm){event('CRITICAL','IT 부하가 냉각 용량 가정을 초과',L(c.selected));capacityAlarm=true}if(c.rackCount*c.rackPowerKw<=c.coolingKw)capacityAlarm=false;if(c.broken&&T.simTime-c.brokenAt>180){let l=L(c.broken);l.down=false;event('INFO','광 링크 복구',l);c.broken=null;c.reroute=false;flow();optics()}},400);

/* Incident response fleet: four field engineers ride service vans to the affected zone.
   Campus roads and travel speed are illustrative training assumptions. */
let responseDispatch=null;
const originalPose=T.npcPose;
const roadRoutes={
 utility:[[-3,29],[-3,-35],[-47,-35],[-47,-39]],
 cooling:[[35,29],[35,-35],[47,-35],[47,-39]],
 gray:[[-3,29],[-25,29],[-25,26]],
 hall:[[-3,29],[15,29],[15,27]],
 network:[[35,29],[44,29],[44,28]],
 ops:[[-3,29],[-43,29],[-43,27]],
 campus:[[26,29],[26,47]]
};
function incidentStage(){const d=T.scenarioDefs.find(x=>x.id===T.scenario);if(!d)return null;const t=T.scenarioElapsed();let st=d.timeline[0];for(const row of d.timeline)if(t>=row[0])st=row;return {def:d,stage:st,elapsed:t,index:d.timeline.indexOf(st)}}
function incidentAsset(){
 const q=incidentStage();if(!q)return null;
 const id=q.stage[4],visible=T.visibleAssets(),a=(id&&T.assetMap.get(id)&&visible.some(x=>x.id===id)&&T.assetMap.get(id))||q.def.assets.map(x=>T.assetMap.get(x)).find(x=>x&&visible.some(y=>y.id===x.id));
 return a||T.selected;
}
function distance2(a,b){return Math.hypot(b[0]-a[0],b[1]-a[1])}
function routeLen(route){return route.slice(1).reduce((n,p,i)=>n+distance2(route[i],p),0)}
function pointOnRoute(route,d){for(let i=1;i<route.length;i++){const a=route[i-1],b=route[i],len=distance2(a,b);if(d<=len){const t=len?d/len:1;return {x:a[0]+(b[0]-a[0])*t,z:a[1]+(b[1]-a[1])*t,heading:Math.atan2(b[1]-a[1],b[0]-a[0])}}d-=len}const b=route[route.length-1],a=route[route.length-2]||b;return {x:b[0],z:b[1],heading:Math.atan2(b[1]-a[1],b[0]-a[0])}}
function responseTarget(a,index){
 const z=T.Z.find(x=>x.id===a.zone),dx=a.x-z.pos[0],dz=a.z-z.pos[1],n=Math.hypot(dx,dz)||1;
 const side=Math.max(a.w,a.d)*.72+2;
 return [a.x-dx/n*side+(index-1.5)*.72,a.z-dz/n*side];
}
function makeResponseRoute(n,index,a){
 const p=originalPose(n),start=[p.x,p.z],roads=roadRoutes[a.zone]||roadRoutes.campus,park=responseTarget(a,index);
 return [start,...roads,park];
}
function startResponse(){
 responseDispatch=null;if(T.scenario==='normal')return;if(T.scenario==='fire'){const safeTarget=incidentAsset();responseDispatch={scenario:'fire',start:T.simTime,asset:safeTarget,people:[]};if(safeTarget){T.setZone(safeTarget.zone);T.pickAsset(safeTarget,true);T.cameraDesired.distance=55;T.cameraDesired.target=[safeTarget.x,2,safeTarget.z]}return;}
 const a=incidentAsset();if(!a)return;
 const staff=T.npcs.filter(n=>n.role!=='guard'&&T.facilityProfiles[T.facilityMode].staff.includes(n.id)).slice(0,4);
 const people=staff.map((n,i)=>{const route=makeResponseRoute(n,i,a),drive=routeLen(route)/7.5;return {n,index:i,route,drive,arrive:drive+2.5,asset:a,park:route[route.length-1],inspect:[a.x+(i-1.5)*.8,a.z+a.d*.65]}});
 responseDispatch={scenario:T.scenario,start:T.simTime,asset:a,people};
 T.setZone('campus');T.pickAsset(a,true);T.cameraDesired.distance=88;T.cameraDesired.target=[a.x,2,a.z];
 T.addEventLog('INFO','현장 출동 차량 4대 배차 · 엔지니어 이동 시작',a,'response');
}
function responsePose(n){
 if(!responseDispatch||responseDispatch.scenario!==T.scenario)return originalPose(n);
 const p=responseDispatch.people.find(x=>x.n.id===n.id);if(!p)return originalPose(n);
 const t=Math.max(0,T.simTime-responseDispatch.start);
 if(t<p.drive){const v=pointOnRoute(p.route,t*7.5);return {...v,action:'ride'}}
 if(t<p.arrive){const v=pointOnRoute(p.route,routeLen(p.route));return {...v,action:'ride'}}
 const walk=Math.min(1,(t-p.arrive)/2.8),start=p.park,end=p.inspect;
 return {x:start[0]+(end[0]-start[0])*walk,z:start[1]+(end[1]-start[1])*walk,heading:Math.atan2(end[1]-start[1],end[0]-start[0]),action:walk<1?'walk':'inspect'}
}
T.npcPose=responsePose;
const responseTask=workerTask;
workerTask=function(n,p){if(responseDispatch&&responseDispatch.scenario===T.scenario){const r=responseDispatch.people.find(x=>x.n.id===n.id);if(r){const t=Math.max(0,T.simTime-responseDispatch.start);return t<r.drive?'출동 차량 탑승 · 장애 구역으로 이동 중':t<r.arrive?'현장 도착 · 차량 하차 중':'장애 설비 점검 중 · '+r.asset.name}}return responseTask(n,p||T.npcPose(n))};
window.LS3D_RESPONSE_ACTIVE=function(id){return T.scenario==='fire'&&!!responseDispatch&&responseDispatch.scenario==='fire'&&T.facilityProfiles[T.facilityMode].staff.includes(id)||!!responseDispatch&&responseDispatch.people.some(x=>x.n.id===id)};
function drawResponse(ctx){
 const q=incidentStage(),active=T.scenario!=='normal'&&q&&!q.stage[5]?.includes('recovery');
 if(responseDispatch&&responseDispatch.scenario===T.scenario){
  responseDispatch.people.forEach(function(p){const t=Math.max(0,T.simTime-responseDispatch.start),dist=Math.min(routeLen(p.route),t*7.5),v=pointOnRoute(p.route,dist),rot=v.heading;
   for(let ri=1;ri<p.route.length;ri++)ctx.line([p.route[ri-1][0],.1,p.route[ri-1][1]],[p.route[ri][0],.1,p.route[ri][1]],'#e6bf67');
   ctx.box(v.x,.04,v.z,2.55,.82,1.12,'#d8e1e7',rot);ctx.box(v.x,.82,v.z-.03,1.12,.55,.9,'#7ba2b5',rot);ctx.box(v.x+.62,.84,v.z-.03,.38,.12,.66,'#f4c86c',rot);
   ctx.box(v.x-.84,.10,v.z-.60,.38,.38,.17,'#172532',rot);ctx.box(v.x-.84,.10,v.z+.60,.38,.38,.17,'#172532',rot);ctx.box(v.x+.84,.10,v.z-.60,.38,.38,.17,'#172532',rot);ctx.box(v.x+.84,.10,v.z+.60,.38,.38,.17,'#172532',rot);
   ctx.box(v.x,.04,v.z,2.18,.075,1.2,'#f8cc67',rot);ctx.box(v.x+.72,.9,v.z,.25,.16,.22,T.simTime%2<1?'#ff596b':'#ffd369',rot);
   if(t<p.arrive+16&&T.scenarioElapsed()>Math.max(45,p.drive)){ctx.box(v.x,.025,v.z,2.6,.025,1.4,[.2,.85,.72,.62],0,'glass')}
  });
 }
 if(!active||!q)return;
 const a=incidentAsset();if(!a)return;const t=q.elapsed,phase=q.stage[5],smokeColor=[.69,.75,.81,.74];
 if(T.scenario==='dlc-leak')ctx.box(a.x,.055,a.z,a.w*1.5,.055,a.d*1.5,[.18,.72,.93,.55],0,'glass');
 if(T.scenario==='rack-hotspot')ctx.ring(a.x,.18,a.z,Math.max(a.w,a.d)*.88,t%1<.5?'#fb806f':'#ffc17a',34);
 if(T.scenario==='fiber-cut'){for(let i=0;i<5;i++){const x=a.x+Math.sin(t*.22+i*2)*.9,z=a.z+Math.cos(t*.19+i*2)*.8;ctx.box(x,1+i%2*.65,z,.16,.65,.16,'#ffd476',t*.15+i)}}
 if(T.scenario==='fire'||T.scenario==='power'||T.scenario==='cooling'||T.scenario==='dlc-leak'||T.scenario==='rack-hotspot'||T.scenario==='ups-battery'){
  const sx=a.x+a.w*.18,sz=a.z-a.d*.18;
  for(let i=0;i<7;i++){const rise=(t*.62+i*1.55)%10,x=sx+Math.sin(t*.08+i*2)*(.45+rise*.15),y=a.h+.4+rise*.44,z=sz+Math.cos(t*.07+i)*.42;const size=1.15+rise*.20;ctx.box(x,y,z,size,size*.78,size,smokeColor,0,'glass')} if(T.scenario==='fire'){const blink=t%1<.5;ctx.box(a.x,.12,a.z,a.w+3,.1,a.d+3,blink?'#ed5965':'#923847',0,'glass');ctx.box(a.x,.3,a.z,.9,2.2,.9,blink?'#ff704f':'#f5bb50',0,'glass')}
 }
 if(T.scenario==='fire'&&q.index>=4&&q.index<6){const pulse=t%1<.5;ctx.box(a.x,.1,a.z,a.w+2,.08,a.d+2,pulse?'#d7e4ec':'#83cbd1',0,'glass');for(let i=0;i<5;i++)ctx.cylinder(a.x-1.8+i*.9,.2+(t%3)*.55,a.z,.5,1.1,[.78,.86,.9,.2],12)}
 if(phase==='recovery')responseDispatch=null;
}
window.LS3D_RESPONSE_VISUALS=window.LS3D_RESPONSE_VISUALS||[];window.LS3D_RESPONSE_VISUALS.push(drawResponse);
const baseResponseScenario=scenarioApply;
scenarioApply=function(id){baseResponseScenario(id);if(id==='normal'){responseDispatch=null}else startResponse()};
T.scenarioApply=scenarioApply;
if(!T.gl){let a=document.createElement('a');a.href='./LS_Datacenter_Campus.html';a.textContent='WebGL 미지원 · 2D Gold Pixel Tour';a.style='position:fixed;bottom:8px;z-index:99;background:#432;color:white;padding:10px';document.body.append(a)}
layout();try{if(!sessionStorage.getItem('twin-guide')){$('guide').classList.add('open');sessionStorage.setItem('twin-guide','1')}}catch(e){}
layout();form();flow();optics();calc()}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',build);else build();
})();