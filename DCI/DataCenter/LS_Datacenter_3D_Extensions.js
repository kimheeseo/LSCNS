/* v4.5.2 · 2026-10-10: 설정 품질 변경 시 상단 표기 직접 갱신 */
(function(){
'use strict';const T=window.__LS3D_TEST__;if(!T)return;const $=x=>document.getElementById(x);
const defs={
 aidc:{rackCount:20,rackPowerKw:60,gpuRatio:80,cooling:'mixed',air:20,rear:20,dlc:60,redundancy:'N+1',targetPue:1.30,pueMin:1.10,pueMax:1.60,outdoorC:24,coolingKw:1600,generatorKw:1800,upsKw:1600,batteryMin:12,upsEfficiencyPct:96,distributionLossPct:2.5,coolingCop:4,otherPct:4,portCount:128,plugKw:18,cpoKw:14},
 'server-room':{rackCount:3,rackPowerKw:6.5,gpuRatio:0,cooling:'air',air:100,rear:0,dlc:0,redundancy:'N',targetPue:1.45,pueMin:1.10,pueMax:2,outdoorC:24,coolingKw:35,generatorKw:0,upsKw:30,batteryMin:15,upsEfficiencyPct:94,distributionLossPct:3,coolingCop:3.5,otherPct:6,portCount:32,plugKw:.8,cpoKw:.6},
 onprem:{rackCount:12,rackPowerKw:8,gpuRatio:15,cooling:'rear',air:60,rear:40,dlc:0,redundancy:'N+1',targetPue:1.40,pueMin:1.10,pueMax:2,outdoorC:24,coolingKw:220,generatorKw:600,upsKw:500,batteryMin:10,upsEfficiencyPct:95,distributionLossPct:2.8,coolingCop:3.8,otherPct:5,portCount:64,plugKw:2.5,cpoKw:1.8}
};
const CONFIG=window.LS3D_CONFIG={rackCount:0,rackPowerKw:0,gpuRatio:0,cooling:'air',air:100,rear:0,dlc:0,redundancy:'N+1',targetPue:1.3,pueMin:1.1,pueMax:2,outdoorC:24,coolingKw:0,generatorKw:0,upsKw:0,batteryMin:12,upsEfficiencyPct:96,distributionLossPct:2.5,coolingCop:4,otherPct:4,mode:'pluggable',plugW:8,cpoW:4,plugKw:18,cpoKw:14,portCount:128,topology:'spine-leaf',radix:64,linkSpeed:'800G',linkDistanceM:34,brokenLatencyUs:18,brokenBandwidthPct:72,cpoPue:0,density:1.35,selected:'tor-spine-a',broken:null,reroute:false,edited:false,quality:'medium',baselineRackInletC:22,availabilityModel:{normal:99.99,power:{N:96,'N+1':99.5,'2N':99.95},cooling:{N:93,'N+1':98.5,'2N':99.9},network:{N:85,'N+1':98,'2N':99.9},other:{N:94,'N+1':98.5,'2N':99.9}}};
const c=CONFIG;const LAYOUT_DATA={schemaVersion:'lsdc-layout/1.0',get zones(){return T.visibleZones().map(z=>({id:z.id,name:z.name}))},get assets(){return T.assets.filter(a=>!a.hidden).map(a=>({id:a.id,zone:a.zone,name:a.name,type:a.type||a.name,x:a.x,z:a.z,width:a.w,depth:a.d,height:a.h,powerKw:(String(a.power||'').match(/[0-9]+(?:\.[0-9]+)?/)||[])[0]?Number(String(a.power).match(/[0-9]+(?:\.[0-9]+)?/)[0]):null}))}};
function preset(id){Object.assign(c,defs[id]||defs.aidc);c.topology=c.topology||'spine-leaf';c.radix=c.radix||64;c.linkSpeed=c.linkSpeed||'800G';c.linkDistanceM=c.linkDistanceM||34;c.broken=null;c.reroute=false;c.edited=false}
preset(T.facilityMode);
const LOSS={smfDbKm:.22,mmfDbKm:3,connectorDb:.25,spliceDb:.10,cpoCouplingDb:.15};
const OPTICAL_PROFILE={'400G':{txDbm:2,rxSensitivityDbm:-7.0},'800G':{txDbm:1,rxSensitivityDbm:-5.5},'1.6T':{txDbm:0,rxSensitivityDbm:-4.0}};
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
const TOPOLOGY_BASE_NODES=nodes.slice(),TOPOLOGY_BASE_LINKS=links.map(x=>({...x}));
const N=id=>nodes.find(x=>x.id===id),L=id=>links.find(x=>x.id===id),A=id=>T.assets.find(x=>x.id===id),esc=x=>String(x).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
function loss(l){if(['DAC/AEC','AOC'].includes(l.type))return [null,null,0];let alpha=l.media==='MMF'?LOSS.mmfDbKm:LOSS.smfDbKm,val=l.length/1000*alpha+l.conn*LOSS.connectorDb+l.splice*LOSS.spliceDb+(c.mode==='cpo'&&l.media==='SMF'?LOSS.cpoCouplingDb:0),profile=OPTICAL_PROFILE[c.linkSpeed]||OPTICAL_PROFILE['800G'];return [val,profile.txDbm-profile.rxSensitivityDbm-val,alpha]}
function event(sev,msg,l){if(T.addEventLog)T.addEventLog(sev,msg,A(N(l?.from)?.asset)||A('odf')||T.selected,'fiber')}
function flow(){T.flowPaths.fiber.splice(0,T.flowPaths.fiber.length,...links.filter(x=>!x.down).map(x=>[N(x.from).pos,N(x.to).pos]));let e=$('optState');if(e)e.textContent=(c.topology==='fat-tree'?'Fat-tree':'Spine-Leaf')+' · radix '+c.radix+' · '+c.linkSpeed+' · '+(c.broken?(c.reroute?'우회 적용 · 지연 +'+Number(c.brokenLatencyUs||18).toFixed(1)+' µs · 가용 대역폭 '+Number(c.brokenBandwidthPct||72).toFixed(0)+'%':'연결 끊김 · 대체 경로 없음'):'정상 경로 · 교육용 가정값')}
function optics(){
 const svg=$('optSvg'),list=$('optList'),detail=$('optDetail'),tbody=$('optPortTable')?.querySelector('tbody');
 if(!svg||!list||!detail)return;
 const pt=n=>({x:28+(n.pos[0]+22)/50*590,y:148-n.pos[1]*4.1});
 const port=(l,id)=>{if(id==='server')return 'NIC-0';return 'Eth1/'+Math.max(1,links.filter(x=>x.from===id||x.to===id).indexOf(l)+1)};
 const setSelected=e=>{const q=L(e.dataset.link);if(q){c.selected=q.id;optics()}};
 svg.innerHTML=links.map(l=>{const a=pt(N(l.from)),b=pt(N(l.to)),cl=(l.down?'bad ':l.route==='B'?'alt ':'')+(c.selected===l.id?'selected':'');return '<path data-link="'+esc(l.id)+'" class="'+cl+'" d="M '+a.x.toFixed(1)+' '+a.y.toFixed(1)+' L '+b.x.toFixed(1)+' '+b.y.toFixed(1)+'"></path>'}).join('')+nodes.map(n=>{const q=pt(n);return '<rect x="'+(q.x-36)+'" y="'+(q.y-10)+'" width="72" height="20" rx="4"></rect><text x="'+q.x+'" y="'+(q.y+3)+'" text-anchor="middle">'+esc(n.name)+'</text>'}).join('');
 const prof=OPTICAL_PROFILE[c.linkSpeed]||OPTICAL_PROFILE['800G'];
 const show=l=>{if(!l)return;const v=loss(l),from=N(l.from),to=N(l.to),fiber=l.length/1000*(l.media==='MMF'?LOSS.mmfDbKm:LOSS.smfDbKm),connector=l.conn*LOSS.connectorDb,splice=l.splice*LOSS.spliceDb,coupling=c.mode==='cpo'&&l.media==='SMF'?LOSS.cpoCouplingDb:0;
 detail.innerHTML='<div class="twin-readout"><b>'+esc(from.name)+' → '+esc(to.name)+'</b><br>From 포트: '+port(l,l.from)+' · To 포트: '+port(l,l.to)+'<br>케이블: '+esc(l.type)+' · 매체: '+esc(l.media)+' · '+esc(c.linkSpeed)+' 가정<br>거리: '+Number(l.length).toLocaleString()+' m · 코어: '+(l.cores||'전기/능동 케이블')+'<br>커넥터: '+esc(l.connector)+' · 접속 '+l.conn+'개 · 스플라이스 '+l.splice+'개<br>손실: '+(v[0]==null?'광 예산 해당 없음 (전기/능동 케이블)':('섬유 '+fiber.toFixed(3)+' + 커넥터 '+connector.toFixed(2)+' + 스플라이스 '+splice.toFixed(2)+' + 결합 '+coupling.toFixed(2)+' = '+v[0].toFixed(2)+' dB'))+'<br>링크 마진: '+(v[1]==null?'계산 대상 없음':v[1].toFixed(2)+' dB · TX '+prof.txDbm+' dBm − RX 감도 '+prof.rxSensitivityDbm+' dBm − 손실')+'<br>상태: '+(l.down?'단선':v[1]!=null&&v[1]<0?'⚠ 링크 마진 음수 · 예산 부족 가정':'정상')+' · 교육·설계 검토용 가정값</div>';
 const d=$('optDistance'),cab=$('optCable');if(d)d.value=l.length;if(cab)cab.value=['DAC/AEC','AOC','MMF','SMF'].includes(l.type)?l.type:(l.media==='SMF'?'SMF':'MMF');
 const cut=$('cutLinkBtn');if(cut)cut.textContent=l.down?'선택 링크 복구':'선택 링크 단선';
 };
 svg.querySelectorAll('[data-link]').forEach(e=>e.onclick=()=>setSelected(e));
 list.innerHTML=links.map(l=>'<button type="button" data-link="'+esc(l.id)+'">'+esc(N(l.from).name)+' ('+port(l,l.from)+') → '+esc(N(l.to).name)+' ('+port(l,l.to)+') · '+esc(l.type)+(l.down?' · 단선':'')+'</button>').join('');
 list.querySelectorAll('[data-link]').forEach(e=>e.onclick=()=>setSelected(e));
 if(tbody)tbody.innerHTML=links.map(l=>'<tr><td>'+esc(N(l.from).name)+' / '+port(l,l.from)+'</td><td>'+esc(N(l.to).name)+' / '+port(l,l.to)+'</td><td>'+esc(l.type)+'</td><td>'+Number(l.length).toLocaleString()+' m</td><td>'+esc(c.linkSpeed)+'</td><td>'+(l.down?'단선':loss(l)[1]!=null&&loss(l)[1]<0?'마진 경고':'정상')+'</td></tr>').join('');
 show(L(c.selected)||links[0]);flow();
}function calc(){
 const rackLoad=c.rackCount*c.rackPowerKw;
 const opticalPower=(c.mode==='cpo'?c.cpoW:c.plugW)*c.portCount/1000+(c.mode==='cpo'?c.cpoKw:c.plugKw);
 let load=rackLoad+opticalPower;
 const elapsed=T.scenarioElapsed?T.scenarioElapsed():0,sc=T.scenario;
 if(sc==='gpu-surge')load*=1+.55*Math.min(1,elapsed/180);
 if(sc==='dlc-leak'&&elapsed>45)load*=.75;
 if(sc==='rack-hotspot'&&elapsed>100)load*=.85;
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

 const availabilityClass=c.broken?'network':sc==='power'||sc==='ups-battery'?'power':sc==='cooling'||sc==='dlc-leak'||sc==='rack-hotspot'?'cooling':sc==='network'||sc==='fiber-cut'?'network':sc==='fire'?'other':'normal';
 let availability=c.availabilityModel[availabilityClass]?.[c.redundancy]??c.availabilityModel.normal;
 if(c.broken&&!c.reroute)availability=0;
 if(sc==='ups-battery'&&elapsed>=c.batteryMin*60)availability=0;
 if(sc==='rack-hotspot'&&temperature>=30)availability=Math.max(0,availability-3);
 if(sc==='cooling'&&temperature>=32)availability=Math.max(0,availability-5);
 const m={load,facility,pue,temperature,availability,latency:c.broken?(c.reroute?(c.brokenLatencyUs||18):999):sc==='fiber-cut'?(c.brokenLatencyUs||18):0,bandwidth:c.broken?(c.reroute?(c.brokenBandwidthPct||72):0):sc==='fiber-cut'?(c.brokenBandwidthPct||72):100,rackLoad,opticalPower,coolingPower,effectiveCop,upsLoss,distributionLoss,otherPower};
 const set=(id,value)=>{const e=$(id);if(e)e.innerHTML=value};
 if(T.scenario==='normal'){set('kpiLoad',load.toFixed(0)+'<small>kW</small>');
 set('kpiFacility',facility.toFixed(0)+'<small>kW</small>');
 const fm=$('kpiFacilityMeta');if(fm)fm.textContent='계산 PUE '+pue.toFixed(2)+' · 가정값';
 set('kpiAvailability',availability.toFixed(2)+'<small>%</small>');
 set('kpiTemp',temperature.toFixed(1)+'<small>°C</small>')}
 if($('thermalOut'))$('thermalOut').textContent='평균 랙 입구 '+temperature.toFixed(1)+'°C · 냉각 용량 부하율 '+(load/Math.max(1,c.coolingKw)*100).toFixed(1)+'% · 외기 '+c.outdoorC.toFixed(1)+'°C (가정)';
 const badMix=Math.abs(c.air+c.rear+c.dlc-100)>.01;
 if($('warn'))$('warn').textContent=(c.pueMin>c.pueMax?'경고: 목표 PUE 하한은 상한 이하여야 합니다. ':badMix?'경고: 냉각 방식 비율 합계를 100%로 맞춰주세요. ':load>c.coolingKw?'경고: IT 부하가 냉각 용량을 초과합니다. ':c.generatorKw<load?'주의: 발전기 가정 용량이 IT 부하 미만입니다. ':c.upsKw<load?'주의: UPS 가정 용량이 IT 부하 미만입니다. ':'')+'계산값은 교육·설계 검토용 가정입니다.';
 if($('cpoOut'))$('cpoOut').textContent='플러거블 '+(c.portCount*c.plugW/1000+c.plugKw).toFixed(1)+' kW · CPO '+(c.portCount*c.cpoW/1000+c.cpoKw).toFixed(1)+' kW · 광 IT 부하 비중 '+(opticalPower/Math.max(1,load)*100).toFixed(2)+'% · 포트 밀도 '+(c.mode==='cpo'?c.density:1).toFixed(2)+'× (가정)';
 const rows=[['랙 IT 부하','랙 수 × 랙당 IT 전력',rackLoad.toFixed(1)+' kW'],['광 네트워크 IT 부하','광 포트 전력 + 스위치 전력',opticalPower.toFixed(1)+' kW'],['IT 부하 합계','랙 IT + 광 네트워크',load.toFixed(1)+' kW'],['유효 냉각 COP','기본 COP × 방식 계수 × 외기·부하 보정',effectiveCop.toFixed(2)],['냉각 전력','IT 부하 ÷ 유효 COP',coolingPower.toFixed(1)+' kW'],['UPS 손실','IT 부하 × (UPS 효율⁻¹ − 1)',upsLoss.toFixed(1)+' kW'],['배전 손실','(IT + 냉각 + UPS 손실) × 손실률',distributionLoss.toFixed(1)+' kW'],['기타 전력','IT 부하 × 기타 전력률',otherPower.toFixed(1)+' kW'],['시설 전력','IT + 냉각 + UPS/배전 손실 + 기타',facility.toFixed(1)+' kW'],['PUE','시설 전력 ÷ IT 부하',pue.toFixed(3)],['가용성','장애 종류 × 이중화 등급 가정',availability.toFixed(2)+'%']];
 const tbody=$('basisRows');if(tbody)tbody.innerHTML=rows.map(r=>'<tr><th>'+r[0]+'</th><td>'+r[1]+'</td><td>'+r[2]+'</td></tr>').join('');
 return m
}
c.calculate=calc;
function layout(){const desired=Math.max(1,Math.min(40,Math.round(Number(c.rackCount)||20))),rows=Math.ceil(desired/5),rackType='GPU server rack';for(let i=0;i<desired;i++){let rack=T.assets.find(a=>a.id==='gpu-'+Math.floor(i/5)+'-'+(i%5));if(!rack){const id='gpu-dyn-'+i;rack=T.assetMap.get(id)||T.addAsset('hall',id,'GPU ROW · R'+(i+1),rackType,0,0,3.2,3.2,6,Number(c.rackPowerKw).toFixed(1)+' kW 가정','설정 패널의 랙 수·랙당 전력을 반영한 대표 3D 배치입니다.')}rack.hidden=false;rack.x=-15+(i%5)*7.2;rack.z=-12+Math.floor(i/5)*7.2;rack.w=3.2;rack.d=3.2;rack.h=6;rack.power=Number(c.rackPowerKw).toFixed(1)+' kW 가정';rack.desc='설정값 반영 · '+Number(c.rackPowerKw).toFixed(1)+' kW/랙 · '+c.gpuRatio+'% GPU 서버 비율 (교육·설계 검토용 가정값).'}T.assets.filter(a=>a.type===rackType).forEach(a=>{const i=T.assets.filter(x=>x.type===rackType).indexOf(a);if(i>=desired)a.hidden=true});const note=$('warn');if(note&&c.rackCount>40)note.textContent='3D 랙은 성능을 위해 40개까지만 대표 표시하며, 전력 계산에는 입력한 '+c.rackCount+'개 랙 전체를 반영합니다. 교육·설계 검토용 가정값.';if(T.selected&&T.selected.hidden)T.pickAsset(T.visibleAssets().find(a=>!a.hidden)||T.selected,false)}
function buildPayload(){return {schemaVersion:'lsdc-twin-bom/1.0',layoutData:LAYOUT_DATA,createdAt:new Date().toISOString(),source:'LS Datacenter 3D Twin v4.4',units:{power:'kW',length:'m',temperature:'degC',flow:'L/min',loss:'dB',time:'min'},assumptionNotice:'교육·설계 검토용 가정값 · 공식 제품 사양 아님',facility:{profile:T.facilityMode,name:{aidc:'AI 데이터센터', 'server-room':'회사 서버실',onprem:'On-Premise'}[T.facilityMode]||T.facilityMode,rackCount:c.rackCount,rackPowerKw:c.rackPowerKw,gpuServerRatio:c.gpuRatio/100,redundancy:c.redundancy,cooling:{mode:c.cooling,airPct:c.air,rearDoorPct:c.rear,dlcPct:c.dlc,capacityKw:c.coolingKw},pueRange:{min:c.pueMin,max:c.pueMax},targetPue:c.targetPue,outdoorC:c.outdoorC,generatorKw:c.generatorKw,upsKw:c.upsKw,batteryMinutes:c.batteryMin,upsEfficiencyPct:c.upsEfficiencyPct,distributionLossPct:c.distributionLossPct,coolingCop:c.coolingCop},optical:{cpoMode:c.mode,assumptions:{pluggablePortW:c.plugW,cpoPortW:c.cpoW,pluggableSwitchKw:c.plugKw,cpoSwitchKw:c.cpoKw,portCount:c.portCount,portSpeed:c.linkSpeed,radix:c.radix},topology:{type:c.topology,radix:c.radix,portSpeed:c.linkSpeed,nodes:nodes.map(n=>({id:n.id,name:n.name,layer:n.id==='server'?'Rack':n.id.startsWith('tor')?'Leaf':n.id.startsWith('sp')?'Spine':n.id==='core'?'Core':n.id==='odf'?'MMR':'External',assetId:n.asset})),links:links.map(l=>({id:l.id,from:l.from,to:l.to,type:l.type,fiber:l.media,media:l.media,cores:l.cores,connectorType:l.connector,speed:l.speed,lengthM:l.length,connectors:l.conn,splices:l.splice,budgetDb:l.budget,route:l.route,down:!!l.down}))}}}}
function saveJson(){const blob=new Blob([JSON.stringify(buildPayload(),null,2)],{type:'application/json'}),url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download='LS_Datacenter_3D_Config.json';a.click();setTimeout(()=>URL.revokeObjectURL(url),500);if(window.toast)toast('공유 스키마 v1.0 JSON 저장 완료')}
function shareToBom(){const raw=unescape(encodeURIComponent(JSON.stringify(buildPayload()))),bytes=Uint8Array.from(raw,x=>x.charCodeAt(0));let binary='';bytes.forEach(x=>binary+=String.fromCharCode(x));const token=btoa(binary).replace(/\+/g,'-').replace(/\//g,'_').replace(/=+$/,'');window.open('./index.html?source=3d-v44#twin='+token,'_blank','noopener');if(window.toast)toast('현재 구성의 BOM 공유 링크를 열었습니다')}
function exportBomCsv(){const assets=T.visibleAssets(),rows=[['설비ID','구역','설비유형','수량','정격전력(kW)','냉각용량(kW)','연결 대상','비고'],...assets.map(a=>{const connections=links.filter(l=>nodes.some(n=>n.asset===a.id&&(n.id===l.from||n.id===l.to))).map(l=>{const n=N(l.from),q=N(l.to);return (n?.asset===a.id?q?.name:n?.name)}).filter(Boolean);const p=String(a.power||'').match(/[0-9]+(?:\.[0-9]+)?/),cool=String(a.cooling||'').match(/[0-9]+(?:\.[0-9]+)?/);return [a.id,a.zone,a.type||a.name,1,p?Number(p[0]):'',cool?Number(cool[0]):'',connections.join(' / '),String(a.desc||'교육·설계 검토용 가정값').replace(/\s+/g,' ').slice(0,240)]})];const raw='\uFEFF'+rows.map(r=>r.map(x=>'"'+String(x??'').replaceAll('"','""')+'"').join(',')).join('\r\n'),url=URL.createObjectURL(new Blob([raw],{type:'text/csv;charset=utf-8'})),a=document.createElement('a');a.href=url;a.download='LS_Datacenter_3D_BOM_Equipment.csv';a.click();setTimeout(()=>URL.revokeObjectURL(url),500);if(window.toast)toast('BOM 공통 컬럼 CSV 저장 완료')}
function form(){let vals=[['rackCount','랙 수',1,500],['rackPowerKw','랙당 IT 전력 kW',1,200],['gpuRatio','GPU 서버 비율 %',0,100],['air','공랭 비율 %',0,100],['rear','후면열교환기 비율 %',0,100],['dlc','DLC 비율 %',0,100],['targetPue','목표 PUE (참고)',1,3],['pueMin','목표 PUE 하한',1,3],['pueMax','목표 PUE 상한',1,3],['outdoorC','외기 온도 °C',-20,50],['coolingKw','냉각 용량 kW',1,1000000],['coolingCop','기본 냉각 COP',1,12],['upsEfficiencyPct','UPS 효율 %',80,99.9],['distributionLossPct','배전 손실률 %',0,15],['otherPct','기타 전력률 %',0,20],['generatorKw','발전기 용량 kW',0,1000000],['upsKw','UPS 용량 kW',0,1000000],['batteryMin','UPS 백업 분',1,180],['portCount','광 포트 수',1,100000],['plugW','플러거블 포트 W',.1,100],['cpoW','CPO 포트 W',.1,100],['plugKw','플러거블 스위치 kW',0,100000],['cpoKw','CPO 스위치 kW',0,100000],['density','CPO 포트 밀도 배율',.1,10]];$('cfgFields').innerHTML=vals.map(x=>'<label>'+x[1]+'<input data-k="'+x[0]+'" type="number" min="'+x[2]+'" max="'+x[3]+'" value="'+c[x[0]]+'"></label>').join()+'<label>냉각 방식 (가정)<select data-k="cooling"><option value="air">공랭 (Air)</option><option value="rear">후면열교환기 (Rear Door)</option><option value="dlc">직접액체냉각 (DLC)</option><option value="mixed">하이브리드 (Hybrid)</option></select></label><label>이중화<select data-k="redundancy"><option>N</option><option>N+1</option><option>2N</option></select></label><label>CPO 비교<select data-k="mode"><option value="pluggable">플러거블 광모듈</option><option value="cpo">CPO</option></select></label><label>렌더 품질<select data-k="quality"><option value="low">품질 낮음</option><option value="medium">품질 보통</option><option value="high">품질 높음</option></select></label>';$('cfgFields').querySelectorAll('[data-k]').forEach(e=>{e.value=c[e.dataset.k];e.oninput=()=>{let v=e.type==='number'?+e.value:e.value;if(e.type==='number'&&(v<+e.min||v>+e.max))return;c[e.dataset.k]=v;c.edited=true;if(e.dataset.k==='redundancy'){$('redundancySelect').value=v;$('redundancySelect').dispatchEvent(new Event('change',{bubbles:true}));T.redundancy=v}if(e.dataset.k==='batteryMin')battery();if(e.dataset.k==='mode')optics();if(e.dataset.k==='quality'){const qualityButton=document.getElementById('qualityToggle'),qualityLabels={low:'낮음',medium:'보통',high:'높음'};if(qualityButton){qualityButton.textContent='품질: '+qualityLabels[v];qualityButton.setAttribute('aria-label','화면 품질: '+qualityLabels[v]+' · 눌러 변경')}}if(e.dataset.k==='rackCount'||e.dataset.k==='rackPowerKw')layout();calc()};e.onchange=e.oninput});$('redundancySelect').value=c.redundancy;$('redundancySelect').dispatchEvent(new Event('change',{bubbles:true}))}
function build(){let loading=document.createElement('div');loading.className='ux-init-loading';loading.setAttribute('role','status');loading.setAttribute('aria-live','polite');loading.innerHTML='<i></i><b>3D 캠퍼스 준비 중</b><span>설비와 시뮬레이션 데이터를 불러오는 중입니다.</span>';document.body.append(loading);setTimeout(()=>loading.remove(),900);let st=document.createElement('style');st.textContent='.twin-modal{display:none;position:fixed;z-index:80;inset:0;background:#020b14d9;justify-content:flex-end}.twin-modal.open{display:flex}.twin-panel{width:min(520px,100vw);height:100dvh;overflow:auto;background:#0c1d2d;padding:16px;color:#e5f7ff;border-left:1px solid #345}.twin-panel h2{font-size:17px}.twin-panel p,.twin-panel small{font-size:10px;line-height:1.6;color:#9fb8c8}.twin-panel label{display:block;font-size:10px;color:#adc3d0}.twin-panel input,.twin-panel select{width:100%;padding:7px;margin:4px 0 8px;background:#081928;color:white;border:1px solid #345;border-radius:6px}.calc-basis{margin:10px 0;padding:9px;background:#10283a;border:1px solid #35536a;border-radius:8px;font-size:10px}.calc-basis summary{cursor:pointer;font-weight:700;color:#9fe8df}.table-scroll{max-width:100%;overflow-x:auto}.calc-basis table{width:100%;border-collapse:collapse;margin:8px 0}.calc-basis th,.calc-basis td{padding:5px;border-bottom:1px solid #294559;text-align:left;vertical-align:top}.calc-basis td:last-child{white-space:nowrap;text-align:right}.calc-basis small{display:block;line-height:1.5}.top-actions>span{display:flex;gap:4px}.top-actions>span button{white-space:nowrap}.twin-panel button{background:#153d50;color:#e6ffff;border:1px solid #398397;border-radius:7px;padding:8px;margin:3px;font-size:10px}.twin-grid{display:grid;grid-template-columns:1fr 1fr;gap:7px}.twin-readout{padding:8px;background:#132c40;border-radius:7px;font-size:10px}.optical-svg{width:100%;background:#081928;border:1px solid #345}.optical-svg path{fill:none;stroke:#55ddc4;stroke-width:4;cursor:pointer}.optical-svg path.alt{stroke:#7aa8ff;stroke-dasharray:5 4}.optical-svg path.bad{stroke:#ff7086}.optical-svg rect{fill:#173950;stroke:#7aa}.optical-svg text{fill:white;font:9px system-ui}#optList button{display:block;width:100%;text-align:left}@media(max-width:420px){.top-actions>span button{font-size:8px;padding:5px 6px}.top-actions{gap:3px}}@media(max-width:790px){.twin-modal{align-items:flex-end}.twin-panel{height:90dvh;width:100%;border-radius:15px 15px 0 0}}';document.head.append(st);
let bar=document.querySelector('.top-actions'),b=document.createElement('span');b.className='ux-top-tools';b.innerHTML='<button id="cfgOpen" aria-label="가정값 설정">⚙ 설정</button><button id="optOpen" aria-label="광 네트워크 상세">◎ 광</button><button id="guideOpen" aria-label="사용 안내 다시 보기" title="사용 안내 다시 보기">? 안내</button><button id="qualityToggle" aria-label="화면 품질 변경">품질: 보통</button>';bar.prepend(b);
const qualityNames={low:'낮음',medium:'보통',high:'높음'},qualityOrder=['low','medium','high'];
function syncQuality(){const q=qualityNames[c.quality]||'보통';$('qualityToggle').textContent='품질: '+q;$('qualityToggle').setAttribute('aria-label','화면 품질: '+q+' · 눌러 변경');const input=document.querySelector('[data-k="quality"]');if(input)input.value=c.quality}
$('qualityToggle').onclick=()=>{c.quality=qualityOrder[(qualityOrder.indexOf(c.quality)+1)%qualityOrder.length];syncQuality();toast('렌더 품질: '+qualityNames[c.quality])};syncQuality();
const mobileToggle=document.createElement('button');mobileToggle.type='button';mobileToggle.id='mobileControlsToggle';mobileToggle.className='mobile-controls-toggle';mobileToggle.setAttribute('aria-expanded','false');mobileToggle.setAttribute('aria-label','3D 화면 조작 메뉴 펼치기');mobileToggle.textContent='화면 조작 ▾';
const vpTop=document.querySelector('.vp-top');if(vpTop){vpTop.insertBefore(mobileToggle,vpTop.querySelector('.vp-tools'));const viewport=$('viewport');const setMobileControls=()=>{const mobile=matchMedia('(max-width:790px)').matches;viewport.classList.toggle('ux-controls-collapsed',mobile);mobileToggle.hidden=!mobile};mobileToggle.onclick=()=>{const open=viewport.classList.toggle('ux-controls-open');mobileToggle.setAttribute('aria-expanded',String(open));mobileToggle.setAttribute('aria-label',open?'3D 화면 조작 메뉴 접기':'3D 화면 조작 메뉴 펼치기');mobileToggle.textContent=open?'화면 조작 ▴':'화면 조작 ▾'};setMobileControls();window.addEventListener('resize',setMobileControls)}

['btnHome','btnIso','btnTop','btnRoof','btnSection','btnWorkers','btnWalk','btnFull','playBtn','resetBtn','searchBtn','exportBtn'].forEach(id=>{const el=$(id);if(el&&!el.hasAttribute('aria-label'))el.setAttribute('aria-label',el.textContent.trim()||id)});
const sceneCanvas=document.getElementById('scene');if(sceneCanvas)sceneCanvas.style.touchAction='none';
document.body.insertAdjacentHTML('beforeend','<div class="twin-modal" id="twinSettings"><section class="twin-panel"><h2>AI 데이터센터 가정값</h2><button data-close>닫기</button><p>교육·설계 검토용 값 · 변경 즉시 계산 반영</p><div><button data-p="server-room">소규모 서버실</button><button data-p="onprem">엔터프라이즈</button><button data-p="aidc">AI 학습 캠퍼스</button></div><div id="cfgFields" class="twin-grid"></div><details class="calc-basis" open><summary>계산 근거 보기 · 수식과 중간값</summary><div class="table-scroll"><table><thead><tr><th>항목</th><th>계산식</th><th>결과</th></tr></thead><tbody id="basisRows"></tbody></table></div><small>외기 20°C 초과 시 COP를 °C당 1% 낮추고, 냉각 방식·부하 보정 계수는 개념 모델 가정값으로 적용합니다. 목표 PUE는 비교용 입력값이며 KPI는 계산 결과를 표시합니다.</small></details><p id="warn"></p><p id="thermalOut"></p><b>CPO / 플러거블 비교</b><p id="cpoOut"></p><p>시설전력=(IT+광 네트워크)×PUE. 냉각 유량=열부하/(4.186×ΔT)×60. 외기 20°C 초과분 1°C당 PUE +0.012 가정.</p><button id="jsonOut">JSON 내보내기</button><button id="jsonIn">JSON 불러오기</button><input id="jsonFile" type="file" accept=".json" hidden><button id="bomOut">BOM 툴로 전달 ↗</button><button id="reset">초기값 복원</button></section></div><div class="twin-modal" id="optModal"><section class="twin-panel"><h2>광 네트워크 토폴로지</h2><button data-close>닫기</button><p>Rack NIC → ToR/Leaf → Spine → Core → ODF/MMR → ISP · 가정값</p><div class="opt-controls"><label>토폴로지<select id="optTopology"><option value="spine-leaf">Spine-Leaf</option><option value="fat-tree">Fat-tree</option></select></label><label>Radix<select id="optRadix"><option>32</option><option>64</option><option>128</option></select></label><label>포트 속도<select id="optSpeed"><option>400G</option><option>800G</option><option>1.6T</option></select></label><label>광 모듈<select id="optMode"><option value="pluggable">Pluggable</option><option value="cpo">CPO</option></select></label><label>선택 링크 거리 (m)<input id="optDistance" type="number" min="0" max="100000" step="1"></label><label>케이블 종류<select id="optCable"><option>DAC/AEC</option><option>AOC</option><option>MMF</option><option>SMF</option></select></label></div><div id="optState" class="opt-state"></div><svg id="optSvg" class="optical-svg" viewBox="0 0 650 170" role="img" aria-label="광 네트워크 계층도"></svg><div id="optDetail"></div><div class="opt-actions"><button id="cutLinkBtn" type="button">선택 링크 단선 / 복구</button><button id="portCsvBtn" type="button">포트 연결 CSV</button></div><div id="optList"></div><div class="opt-table-wrap"><table id="optPortTable"><thead><tr><th>From 장비 / 포트</th><th>To 장비 / 포트</th><th>케이블</th><th>거리</th><th>속도</th><th>상태</th></tr></thead><tbody></tbody></table></div></section></div><div class="twin-modal" id="guide" role="dialog" aria-modal="true" aria-labelledby="guideTitle"><section class="twin-panel twin-tour-panel"><div class="twin-head"><div><h2 id="guideTitle">처음 사용하는 분을 위한 안내</h2><p id="tourCount" class="twin-note">1 / 5 단계</p></div><button id="tourClose" class="twin-close" type="button" aria-label="안내 닫기">닫기 ✕</button></div><div class="twin-tour-progress" aria-hidden="true"><i id="tourProgress"></i></div><h3 id="tourStepTitle"></h3><p id="tourStepBody" class="twin-tour-copy"></p><div id="tourTip" class="twin-tip"></div><label class="twin-dont-show"><input type="checkbox" id="tourDontShow"> 다시 보지 않기</label><footer class="twin-tour-actions"><button id="tourSkip" class="twin-action" type="button">건너뛰기</button><span><button id="tourPrev" class="twin-close" type="button">이전</button> <button id="tourNext" class="twin-action" type="button">다음</button></span></footer></section></div>');
$('cfgOpen').onclick=()=>{form();$('twinSettings').classList.add('open')};$('optOpen').onclick=()=>{$('optModal').classList.add('open');optics()};
const tourSteps=[
{title:'1 · 구역 선택',body:'왼쪽 구역 목록 또는 상단 구역 선택에서 7개 구역을 둘러보세요.',tip:'키보드 숫자 1–7로도 구역을 바꿀 수 있습니다.'},
{title:'2 · 설비 살펴보기',body:'3D 장면의 설비를 클릭하거나 오른쪽 설비 목록에서 선택하면 위치가 강조되고 상세 패널이 열립니다.',tip:'드래그로 회전, 휠로 확대·축소, Shift+드래그 또는 우클릭 드래그로 이동합니다.'},
{title:'3 · 운영 시나리오',body:'오른쪽 장애 시나리오에서 전원·냉각·네트워크 장애를 선택하고 설비 반응과 이벤트 로그를 확인합니다.',tip:'1× / 5× / 15× 배속은 시뮬레이션 시간에 적용됩니다.'},
{title:'4 · 흐름 표시',body:'전력, 냉각, 광통신 버튼을 켜거나 끄며 캠퍼스 내 연결 경로를 비교하세요.',tip:'데이터와 상태는 교육·설계 검토용 가정값입니다.'},
{title:'5 · 작업자 시점',body:'작업자 4명 버튼에서 담당 업무를 고른 뒤 직원 시점으로 이동해 순찰과 점검을 관찰하세요.',tip:'모바일: 한 손가락 회전, 두 손가락 드래그 이동 및 핀치 확대·축소.'}
];
let tourIndex=0;
function renderTour(){const s=tourSteps[tourIndex];$('tourCount').textContent=(tourIndex+1)+' / '+tourSteps.length+' 단계';$('tourStepTitle').textContent=s.title;$('tourStepBody').textContent=s.body;$('tourTip').textContent=s.tip;$('tourProgress').style.width=((tourIndex+1)/tourSteps.length*100)+'%';$('tourPrev').disabled=tourIndex===0;$('tourNext').textContent=tourIndex===tourSteps.length-1?'완료':'다음';}
function openTour(){tourIndex=0;renderTour();$('guide').classList.add('open');$('tourNext').focus()}
function closeTour(completed){try{if(completed||$('tourDontShow').checked)localStorage.setItem('ls3d-tour-v45','done')}catch(e){}$('guide').classList.remove('open');$('guideOpen').focus()}
$('guideOpen').onclick=openTour;$('tourClose').onclick=()=>closeTour(false);$('tourSkip').onclick=()=>closeTour(true);$('tourPrev').onclick=()=>{tourIndex=Math.max(0,tourIndex-1);renderTour()};$('tourNext').onclick=()=>{if(tourIndex===tourSteps.length-1)closeTour(true);else{tourIndex++;renderTour()}};
document.querySelectorAll('[data-close]').forEach(x=>x.onclick=()=>x.closest('.twin-modal').classList.remove('open'));
document.querySelectorAll('[data-p]').forEach(x=>x.onclick=()=>{T.applyFacilityMode(x.dataset.p);preset(x.dataset.p);layout();form();calc()});
$('jsonOut').onclick=saveJson;$('jsonIn').onclick=()=>$('jsonFile').click();$('jsonFile').onchange=e=>{let f=e.target.files[0];if(f){let r=new FileReader();r.onload=()=>{try{let o=JSON.parse(r.result);if(o.schemaVersion!=='lsdc-twin-bom/1.0')throw Error('schema mismatch');const f=o.facility||{},cool=f.cooling||{},profile=f.profile||f.facilityType;if(['aidc','server-room','onprem'].includes(profile)&&profile!==T.facilityMode){T.applyFacilityMode(profile);preset(profile)}Object.assign(c,{rackCount:f.rackCount??c.rackCount,rackPowerKw:f.rackPowerKw??c.rackPowerKw,gpuRatio:(f.gpuServerRatio??c.gpuRatio/100)*100,redundancy:f.redundancy||c.redundancy,cooling:cool.mode||'air',air:cool.airPct??c.air,rear:cool.rearDoorPct??c.rear,dlc:cool.dlcPct??c.dlc,coolingKw:cool.capacityKw??f.coolingKw??c.coolingKw,targetPue:f.targetPue??c.targetPue,pueMin:f.pueRange?.min??c.pueMin,pueMax:f.pueRange?.max??c.pueMax,outdoorC:f.outdoorC??c.outdoorC,generatorKw:f.generatorKw??c.generatorKw,upsKw:f.upsKw??c.upsKw,batteryMin:f.batteryMinutes??f.batteryMin??c.batteryMin,upsEfficiencyPct:f.upsEfficiencyPct??c.upsEfficiencyPct,distributionLossPct:f.distributionLossPct??c.distributionLossPct,coolingCop:f.coolingCop??c.coolingCop,edited:true});if(o.optical){c.mode=o.optical.cpoMode||c.mode;Object.assign(c,{plugW:o.optical.assumptions.pluggablePortW||c.plugW,cpoW:o.optical.assumptions.cpoPortW||c.cpoW,plugKw:o.optical.assumptions.pluggableSwitchKw||c.plugKw,cpoKw:o.optical.assumptions.cpoSwitchKw||c.cpoKw,portCount:o.optical.assumptions.portCount||c.portCount});c.topology=o.optical.topology?.type||c.topology;c.radix=o.optical.topology?.radix||c.radix;c.linkSpeed=o.optical.topology?.portSpeed||c.linkSpeed;syncTopology();(o.optical.topology?.links||[]).forEach(q=>{let l=L(q.id);if(l){l.length=q.lengthM??l.length;l.conn=q.connectors??l.conn;l.splice=q.splices??l.splice;l.down=!!q.down}})}form();wireOpticalControls();calc();optics()}catch(z){alert('JSON 오류: '+z.message)}};r.readAsText(f)}};
$('bomOut').onclick=shareToBom;$('reset').onclick=()=>{preset(T.facilityMode);form();calc()};$('redundancySelect').addEventListener('change',e=>{c.redundancy=e.target.value;flow()});
$('facilityModeSelect').onchange=e=>{T.applyFacilityMode(e.target.value);preset(e.target.value);form();calc()};
document.addEventListener('keydown',e=>{if(e.key==='Escape')document.querySelectorAll('.twin-modal').forEach(x=>x.classList.remove('open'));if(/^[1-7]$/.test(e.key)&&!['INPUT','SELECT'].includes(document.activeElement.tagName)){let z=T.visibleZones()[+e.key-1];if(z)T.setZone(z.id)}});
let resumeAfterHidden=false;document.addEventListener('visibilitychange',()=>{if(document.hidden){resumeAfterHidden=!!T.simRunning;if(resumeAfterHidden)T.simRunning=false}else if(resumeAfterHidden){T.simRunning=true;resumeAfterHidden=false}});
const oldScenario=T.scenarioApply;T.scenarioApply=function(id){oldScenario(id);if(id==='normal'&&c.broken){let l=L(c.broken);if(l)l.down=false;c.broken=null;c.reroute=false;flow();optics()}};
T.metrics=function(){return calc()};
let up=T.scenarioDefs.find(x=>x.id==='ups-battery'),power=T.scenarioDefs.find(x=>x.id==='power');const battery=()=>{let t=Math.max(60,c.batteryMin*60);if(up){up.timeline[2][0]=Math.max(30,t*.2);up.timeline[3][0]=t;up.timeline[4][0]=t+60;up.timeline[5][0]=t+180}if(power)power.timeline[3][0]=Math.max(90,t+180)};
battery();
let tempAlarm=false,capacityAlarm=false,thermalThrottleLogged=false,thermalShutdownLogged=false,networkCutoverLogged=false,powerTransferLogged=false,powerBatteryDepletedLogged=false,upsBatteryDepletedLogged=false;setInterval(()=>{let live=calc(),elapsed=T.scenarioElapsed?T.scenarioElapsed():0;if(T.scenario==='cooling'){const thermalScale=c.redundancy==='N'?1:c.redundancy==='N+1'?.36:.12;live=Object.assign({},live,{temperature:live.temperature+Math.min(24,elapsed*.055*thermalScale)})}if(T.scenario==='power'||T.scenario==='ups-battery'){const generatorReady=T.scenario==='power'&&c.generatorKw>=live.load&&c.generatorKw>0&&c.upsKw>=live.load&&elapsed>=12;live=Object.assign({},live,{upsRemaining:generatorReady?c.batteryMin:Math.max(0,c.batteryMin-elapsed/60)})}if(live.temperature>27&&!tempAlarm){event('WARN','랙 입구 온도 가정값 27°C 초과',L(c.selected));tempAlarm=true}if(live.temperature<=27)tempAlarm=false;if(c.rackCount*c.rackPowerKw>c.coolingKw&&!capacityAlarm){event('CRITICAL','IT 부하가 냉각 용량 가정을 초과',L(c.selected));capacityAlarm=true}if(c.rackCount*c.rackPowerKw<=c.coolingKw)capacityAlarm=false;if(T.scenario==='cooling'&&live.temperature>=32&&!thermalThrottleLogged){event('WARN','랙 입구 32°C 가정 임계값 초과 · GPU/서버 서멀 스로틀링',A('tor-a')||T.selected,'stage');thermalThrottleLogged=true}if(T.scenario!=='cooling'||live.temperature<32)thermalThrottleLogged=false;if(T.scenario==='cooling'&&live.temperature>=41&&!thermalShutdownLogged){event('CRITICAL','랙 입구 41°C 가정 임계값 초과 · 부하 차단',A('company-rack-a')||T.selected,'stage');thermalShutdownLogged=true}if(T.scenario!=='cooling'||live.temperature<41)thermalShutdownLogged=false;if(T.scenario==='network'&&elapsed>=15&&!networkCutoverLogged){event('WARN',c.redundancy==='N'?'Core 장애 · 우회 경로 없음 · 연결 가용 대역폭 0%':`Core 장애 · Spine 우회 적용 · 지연 ${c.redundancy==='N+1'?18:10} µs · 대역폭 ${c.redundancy==='N+1'?72:90}%`,A('spine-a')||T.selected,'stage');networkCutoverLogged=true}if(T.scenario!=='network')networkCutoverLogged=false;if(T.scenario==='power'&&elapsed>=12&&!powerTransferLogged){event('INFO','발전기 기동 신호 · UPS 배터리 가교 진행 (기동 지연 12초)',A('generator')||T.selected,'stage');powerTransferLogged=true}if(T.scenario==='power'&&live.upsRemaining<=0&&elapsed>12&&!powerBatteryDepletedLogged){event('CRITICAL','UPS 백업 잔여시간 소진 · 부하 차단',A('ups-a')||T.selected,'stage');powerBatteryDepletedLogged=true}if(T.scenario==='ups-battery'&&live.upsRemaining<=0&&!upsBatteryDepletedLogged){event('CRITICAL','UPS 배터리 방전 · 부하 차단',A('ups-a')||T.selected,'stage');upsBatteryDepletedLogged=true}if(T.scenario!=='power'){powerTransferLogged=false;powerBatteryDepletedLogged=false}if(T.scenario!=='ups-battery')upsBatteryDepletedLogged=false;if(c.broken&&T.simTime-c.brokenAt>180){let l=L(c.broken);l.down=false;event('INFO','광 링크 복구',l);c.broken=null;c.reroute=false;flow();optics()}},400);

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
if(!T.gl){const f=document.createElement('div');f.className='ux-fallback';f.setAttribute('role','status');f.innerHTML='<strong>3D 그래픽을 사용할 수 없습니다</strong><p>브라우저의 WebGL 지원을 확인하거나 2D 탐방 페이지에서 캠퍼스를 살펴보세요.</p><a href="./LS_Datacenter_Campus.html">2D Gold Pixel Tour 열기 ↗</a>';document.getElementById('viewport').append(f)}
layout();try{if(!localStorage.getItem('ls3d-tour-v45'))openTour()}catch(e){openTour()}

function syncTopology(){
 const keep=TOPOLOGY_BASE_NODES.filter(n=>c.topology!=='fat-tree'||!['spA','spB'].includes(n.id));
 const extra=c.topology==='fat-tree'?[{id:'aggA',name:'Aggregation A',pos:[8,15],asset:null},{id:'aggB',name:'Aggregation B',pos:[22,15],asset:null}]:[];
 nodes.splice(0,nodes.length,...keep,...extra);
 let next;
 if(c.topology==='fat-tree'){
  const fixed=TOPOLOGY_BASE_LINKS.filter(l=>['server-tor','server-tor-b','core-odf','odf-isp','campus-dci'].includes(l.id)).map(l=>({...l}));
  const tiers=[['torA-aggA','torA','aggA','A',22],['torA-aggB','torA','aggB','B',27],['torB-aggA','torB','aggA','B',27],['torB-aggB','torB','aggB','A',22],['aggA-core','aggA','core','A',24],['aggB-core','aggB','core','B',24]].map(x=>({id:x[0],from:x[1],to:x[2],type:'SMF trunk',media:'SMF',cores:24,connector:'MPO-16',speed:c.linkSpeed,length:x[4],conn:4,splice:2,budget:0,route:x[3],down:false}));
  next=[...fixed,...tiers];
 }else next=TOPOLOGY_BASE_LINKS.map(l=>({...l}));
 next.forEach(l=>l.speed=c.linkSpeed);links.splice(0,links.length,...next);c.broken=null;c.reroute=false;c.selected=links.some(l=>l.id===c.selected)?c.selected:(links[0]?.id||null);
}
function opticalPathExists(){const seen=new Set(['server']),queue=['server'];while(queue.length){const u=queue.shift();if(u==='isp')return true;links.filter(l=>!l.down&&(l.from===u||l.to===u)).forEach(l=>{const v=l.from===u?l.to:l.from;if(!seen.has(v)){seen.add(v);queue.push(v)}})}return false}
function cutSelectedLink(){const l=L(c.selected);if(!l)return;if(l.down){l.down=false;c.broken=null;c.reroute=false;event('INFO','광 링크 복구 · '+N(l.from).name+' → '+N(l.to).name,l)}else{l.down=true;c.broken=l.id;c.reroute=opticalPathExists();c.brokenLatencyUs=Math.max(10,12+l.length/100);c.brokenBandwidthPct=c.reroute?Math.max(35,100-18-Math.round(100/Math.max(1,c.radix))):0;event(c.reroute?'WARN':'CRITICAL','광 링크 단선 · '+N(l.from).name+' → '+N(l.to).name+(c.reroute?' · 대체 경로 우회':' · 대체 경로 없음'),l)}flow();optics();calc()}
function exportPortCsv(){const rows=[['From 장비','From 포트','To 장비','To 포트','토폴로지','케이블 종류','매체','거리 (m)','속도','커넥터','상태','링크 손실 (dB)','링크 마진 (dB)'],...links.map(l=>{const q=loss(l);return [N(l.from).name,portName(l,l.from),N(l.to).name,portName(l,l.to),c.topology==='fat-tree'?'Fat-tree':'Spine-Leaf',l.type,l.media,l.length,c.linkSpeed,l.connector,l.down?'단선':'정상',q[0]==null?'N/A':q[0].toFixed(3),q[1]==null?'N/A':q[1].toFixed(3)]})];const raw=String.fromCharCode(65279)+rows.map(r=>r.map(x=>'\"'+String(x).replaceAll('\"','\"\"')+'\"').join(',')).join(String.fromCharCode(13,10)),url=URL.createObjectURL(new Blob([raw],{type:'text/csv;charset=utf-8'})),a=document.createElement('a');a.href=url;a.download='LS_Datacenter_Optical_Port_Map.csv';a.click();setTimeout(()=>URL.revokeObjectURL(url),500)}
function portName(l,id){if(id==='server')return 'NIC-0';return 'Eth1/'+Math.max(1,links.filter(x=>x.from===id||x.to===id).indexOf(l)+1)}
function wireOpticalControls(){
 const top=$('optTopology'),rad=$('optRadix'),speed=$('optSpeed'),mode=$('optMode'),dist=$('optDistance'),cable=$('optCable');
 if(top){top.value=c.topology;top.onchange=()=>{c.topology=top.value;syncTopology();optics();flow();calc()};}
 if(rad){rad.value=String(c.radix);rad.onchange=()=>{c.radix=+rad.value;calc();flow()};}
 if(speed){speed.value=c.linkSpeed;speed.onchange=()=>{c.linkSpeed=speed.value;links.forEach(l=>l.speed=c.linkSpeed);optics();calc()};}
 if(mode){mode.value=c.mode;mode.onchange=()=>{c.mode=mode.value;const q=document.querySelector('[data-k="mode"]');if(q)q.value=c.mode;optics();calc()};}
 if(dist){const applyDistance=()=>{const l=L(c.selected);if(l){l.length=Math.max(0,Math.min(100000,+dist.value||0));c.linkDistanceM=l.length;optics();calc()}};dist.oninput=applyDistance;dist.onchange=applyDistance;}
 if(cable)cable.onchange=()=>{const l=L(c.selected);if(l){l.type=cable.value;l.media=['SMF','MMF'].includes(cable.value)?cable.value:'MMF';optics();calc()}};
 const cut=$('cutLinkBtn');if(cut)cut.onclick=cutSelectedLink;
 const csv=$('portCsvBtn');if(csv)csv.onclick=exportPortCsv;
}

syncTopology();layout();form();wireOpticalControls();flow();optics();calc();const exportBtn=$('exportBtn');if(exportBtn)exportBtn.onclick=exportBomCsv;
const phase3ScenarioApply=scenarioApply;scenarioApply=function(id){phase3ScenarioApply(id);if(id==='fiber-cut'){if(!c.broken)cutSelectedLink()}else if(id==='normal'&&c.broken){const q=L(c.broken);if(q)q.down=false;c.broken=null;c.reroute=false;flow();optics()}calc()};T.scenarioApply=scenarioApply;
}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',build);else build();
})();