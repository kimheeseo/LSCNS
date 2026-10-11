(function(root){
'use strict';
const profiles={dr4:{name:'400G DR4',gbps:400,active:8,perFiber:100},fr4:{name:'400G FR4',gbps:400,active:2,perFiber:400}};
function calculate(i){
 const p=profiles[i.profile];if(!p)throw Error('광 인터페이스를 선택하세요.');
 const n=(v,label,min,integer=true)=>{if(v==null||typeof v==='boolean'||String(v).trim()==='')throw Error(label+'을 입력하세요.');const x=Number(v);if(!Number.isFinite(x)||x<min||(integer&&!Number.isSafeInteger(x)))throw Error(label+' 입력을 확인하세요.');return x;};
 const links=n(i.links,'경로당 링크 수',1),routes=n(i.routes,'독립 경로 수',1),allocated=n(i.allocated,'링크당 배정 심수',p.active),fibers=n(i.fibers,'케이블 심수',1),spare=n(i.spare,'추가 예비 링크율',0,false);
 const capacity=Math.floor(fibers/allocated);if(!capacity)throw Error('케이블 심수가 링크당 배정 심수보다 작습니다.');
 const plannedLinks=Math.ceil(links*(1+spare/100)),cablesPerRoute=Math.ceil(plannedLinks/capacity);
 const r={profile:p.name,perFiberGbps:p.perFiber,activeFibers:links*p.active,requiredFibers:plannedLinks*allocated,capacity,plannedLinks,cablesPerRoute,totalCables:cablesPerRoute*routes,installedFibers:cablesPerRoute*routes*fibers,trafficTbps:links*p.gbps/1000,totalTrafficTbps:links*routes*p.gbps/1000};
 if(Object.values(r).some(v=>typeof v==='number'&&(!Number.isFinite(v)||Math.abs(v)>Number.MAX_SAFE_INTEGER)))throw Error('계산 범위를 초과했습니다.');return r;
}
const model={profiles,calculate};if(typeof module!=='undefined'&&module.exports)module.exports=model;root.DCFiberCapacity=model;
if(typeof document==='undefined')return;
const $=id=>document.getElementById(id),fmt=n=>n.toLocaleString('ko-KR',{maximumFractionDigits:3});
function update(){try{const r=calculate({profile:$('fc-profile').value,links:$('fc-links').value,routes:$('fc-routes').value,allocated:$('fc-allocated').value,fibers:$('fc-fibers').value,spare:$('fc-spare').value});$('fc-result').textContent=`한 심당 ${fmt(r.perFiberGbps)} Gb/s (한 방향) · 케이블 1개당 최대 ${fmt(r.capacity)}링크 · 경로당 활성 ${fmt(r.activeFibers)}심 / 예비 포함 최소 배정 ${fmt(r.requiredFibers)}심 · 경로당 ${fmt(r.cablesPerRoute)}개, 전체 ${fmt(r.totalCables)}개 케이블 (${fmt(r.installedFibers)}심 설치) · 현재 통신용 링크 전체 ${fmt(r.totalTrafficTbps)} Tb/s/방향`; }catch(e){$('fc-result').textContent='산정 대기: '+e.message;}}
function mount(){const target=$('view-optical');if(!target)return false;if($('fiber-capacity-panel'))return true;const panel=document.createElement('section');panel.id='fiber-capacity-panel';panel.className='panel';panel.innerHTML=`
<h2>LS전선 관점 · 한 심당 전송량과 1,024F 케이블 선정</h2>
<p><b>심수 선정 기준은 W/심이 아니라 Gb/s/심, 양방향 링크당 심수, 연결 경로입니다.</b> 광섬유의 속도가 고정된 것은 아니며 양 끝 광모듈·광엔진의 병렬 레인 또는 WDM 방식에 따라 달라집니다. F는 광섬유 심수입니다.</p>
<div class="tableWrap"><table><thead><tr><th>광 인터페이스</th><th>한 심당 전송량 · 한 방향</th><th>400G 양방향 링크당 활성 심수</th><th>1,024F 산술 최대 · 예비 없음</th></tr></thead><tbody>
<tr><td>400G DR4 · SMF · 최대 500 m</td><td>100 Gb/s</td><td>송신 4 + 수신 4 = 8심</td><td>128링크 · 51.2 Tb/s/방향</td></tr>
<tr><td>400G FR4 · SMF · 최대 2 km</td><td>400 Gb/s (4파장 × 100G)</td><td>송신 1 + 수신 1 = 2심</td><td>512링크 · 204.8 Tb/s/방향</td></tr></tbody></table></div>
<p>위 값은 명목 통신 속도이며 실제 사용자 데이터 처리량과 다릅니다. <b>DR4의 MPO-12에는 활성 8심과 미사용 4위치가 있습니다.</b> 12심 전체를 트렁크에서 링크마다 배정하는 배선은 1,024 ÷ 12의 정수부인 <b>85링크</b>만 수용합니다. Base-8 변환·카세트 구성으로 8심을 효율적으로 배정할 수 있는지 실제 접속도와 극성으로 확인하세요.</p>
<h3>경로별 광케이블 심수 참고 계산</h3>
<p>독립 경로는 동일 양 끝 접속점 사이의 케이블 경로입니다. 서로 다른 랙·스위치 쌍이나 이중화 경로의 심수를 한 케이블로 합쳐 계산하지 않습니다. 아래 경로들은 같은 링크 수를 갖는 것으로 가정합니다. 링크 수가 다르면 경로마다 계산 후 케이블 수를 합산하세요.</p>
<div class="fields sub">
<label><span>광 인터페이스</span><select id="fc-profile" data-reference-only="true"><option value="dr4">400G DR4 · 활성 8심</option><option value="fr4">400G FR4 · 활성 2심</option></select></label>
<label><span>경로당 현재 양방향 링크 수</span><input id="fc-links" data-reference-only="true" type="number" min="1" step="1" value="128"></label>
<label><span>같은 규모의 독립 경로 수</span><input id="fc-routes" data-reference-only="true" type="number" min="1" step="1" value="1"></label>
<label><span>링크당 트렁크 배정 심수 (활성 + 미사용 위치)</span><input id="fc-allocated" data-reference-only="true" type="number" min="8" step="1" value="8"></label>
<label><span>후보 케이블 1개 심수 (F)</span><input id="fc-fibers" data-reference-only="true" type="number" min="1" step="1" value="1024"></label>
<label><span>현재 링크 수 대비 추가 예비 링크율 (%)</span><input id="fc-spare" data-reference-only="true" type="number" min="0" step="any" value="0"></label>
</div><p id="fc-result" role="status"></p>
<p>계산: 예비 포함 링크 = ceil[현재 링크 × (1 + 예비율/100)]; 케이블당 수용 링크 = floor[케이블 심수 ÷ 링크당 배정 심수]; 경로당 케이블 = ceil[예비 포함 링크 ÷ 케이블당 수용 링크]. 하나의 링크 배정 블록은 한 케이블에 수용하는 가정입니다. 예비율은 <b>설치 후 남겨둘 확장 링크</b>이며 구매 예비 재고와 별개입니다.</p>
<p>예: DR4 128링크에 8심씩 배정하면 1,024심. 확장 링크 20%를 추가하면 154링크 × 8 = <b>최소 1,232 배정심</b>이 필요해 1,024F 한 개로는 부족하며 이 가정에서는 두 개가 필요합니다. 12심 배정이면 예비 없이도 1,536심이 필요합니다.</p>
<p><b>현재 설계와의 관계:</b> 이 계산은 비교 참고용이며 기존 BOM·트렁크 설정을 변경하지 않습니다. 실제 설계 결과의 활성 심수, 채널 점유 심수, 설치 케이블 심수는 서로 다를 수 있습니다. CPO 데이터 경로는 제품별 Gb/s/레인·레인 수·WDM·양방향 구성으로 판단하며, ELS→광엔진의 PM fiber는 레이저 공급 경로이므로 위 데이터 링크 환산을 적용하지 않습니다.</p>
<p>1,024F 입력은 용량 비교 후보이며 LS전선의 해당 심수 SKU 공급·적합 판정을 뜻하지 않습니다. 실제 선정에는 광섬유 종류, 길이·손실 예산, 커넥터·극성·카세트, 관로·굽힘, 접속점, 유지보수·예비·이중화를 확인하세요.</p>
<p>공식 근거: <a href="https://networking-docs.nvidia.com/cablemanagfaq/cable-and-connector-definitions" target="_blank" rel="noopener">NVIDIA DR4/FR4 심수·도달거리·MPO 정의</a> · <a href="https://networking-docs.nvidia.com/mms4x00ns400/specifications" target="_blank" rel="noopener">400G DR4 제품 사양</a>. 확인일 2026-10-11.</p>`;
target.prepend(panel);panel.querySelectorAll('input').forEach(e=>e.addEventListener('input',update));$('fc-profile').onchange=()=>{const a=profiles[$('fc-profile').value].active;$('fc-allocated').value=a;$('fc-allocated').min=a;update();};update();return true;}
function start(){if(mount())return;let count=0;const timer=setInterval(()=>{if(mount()||++count>30)clearInterval(timer);},100);}if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',start,{once:true});else start();
})(typeof window!=='undefined'?window:globalThis);
