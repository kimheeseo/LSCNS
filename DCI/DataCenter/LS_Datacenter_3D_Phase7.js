/* Phase 7 v4.7.0 · NVIDIA catalog-grounded draft BOM, no procurement approval.
 * Source of hardware specs: existing product_catalog/NVIDIA/.../catalog.json. */
(function(){
'use strict';
const T=window.__LS3D_TEST__,C=window.LS3D_CONFIG,view=document.getElementById('viewport');if(!T||!C||!view)return;
const $=id=>document.getElementById(id),esc=s=>String(s??'').replace(/[&<>"']/g,x=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[x]));
const base='./product_catalog/NVIDIA/';const paths=['InfiniBand Switch Models/catalog.json','Networking/catalog.json'];
let catalog=null,last=null,loading=null,openState=false;
function select(j,id){return j?.products?.[id]||null}
async function load(){
 if(loading)return loading;
 loading=Promise.all(paths.map(async p=>{const r=await fetch(base+p,{cache:'no-store'});if(!r.ok)throw Error('제품 카탈로그 불러오기 실패: '+p+' HTTP '+r.status);return r.json()}))
 .then(j=>{catalog={switches:j[0],networking:j[1],loadedAt:new Date().toISOString()};return catalog}).catch(e=>{loading=null;throw e});
 return loading;
}
function estimate(o,cat=catalog){
 if(!cat)throw Error('NVIDIA 제품 카탈로그를 먼저 불러오세요.');
 const racks=Number(o.racks),uplink=Number(o.uplinkPerRack),length=Number(o.lengthM);
 const speed=Number(o.speedG||800),proto=o.protocol==='Ethernet'?'Ethernet':'InfiniBand',mode=o.mode==='cpo'?'cpo':'pluggable';
 if(!Number.isInteger(racks)||racks<1||racks>200000||!Number.isInteger(uplink)||uplink<1||uplink>8||!Number.isFinite(length)||length<.1||length>10000||![400,800].includes(speed))throw Error('랙 1~200,000, 링크/랙 1~8, 거리 0.1~10,000m, 속도 400/800G 범위를 확인하세요.');
 const srcSwitch=proto==='InfiniBand'?(mode==='cpo'?select(cat.switches,'Q3450-LD'):select(cat.switches,'Q3400-RA')):select(cat.networking,'nvidia-spectrum-4-sn5000');
 const srcNic=select(cat.networking,'nvidia-connectx-8-supernic');
 if(!srcSwitch||!srcNic)throw Error('NVIDIA 제품 레코드가 없습니다. 카탈로그를 확인하세요.');
 const linkCount=racks*uplink,cpo=mode==='cpo',logicalPortCapacity=proto==='InfiniBand'?144:null;
 const notes=[],blocking=[];
 if(proto==='Ethernet'&&cpo)blocking.push('Ethernet SN5000에 Q3450-LD InfiniBand CPO 인터페이스를 적용할 수 없습니다. CPO/이더넷 조합 재선정 필요.');
 if(proto==='Ethernet'&&speed===800)blocking.push('ConnectX-8 Ethernet은 포트당 최대 400G(합계 800G) 사례입니다. 800G 단일 포트 NIC로 확정할 수 없습니다.');
 if(proto==='InfiniBand'&&speed===400)notes.push('스위치 800G 포트의 400G breakout·재구성 및 광학 인터페이스 호환 검토가 필요합니다.');
 if(proto==='InfiniBand'&&cpo)notes.push('Q3450-LD 전면 MPO12 CPO: 스위치 측 플러거블 광모듈 수량은 0. 호스트 측 광모듈은 별도 규격 검증.');
 if(proto==='InfiniBand'&&!cpo)notes.push('Q3400-RA: 144 논리 800G 포트 / 72 OSFP cage를 구분해야 합니다. 포트마다 OSFP 하나라는 산식은 사용하지 않습니다.');
 if(proto==='Ethernet')notes.push('Spectrum-4 SN5000은 제품군 레퍼런스이며 고정 모델의 포트 수/속도 조합이 확정되지 않아 스위치 수량 자동 확정 불가.');
 const switches=logicalPortCapacity?Math.ceil(linkCount/logicalPortCapacity):null;
 const switchSidePluggable=cpo?0:null;
 const fibersCables=linkCount;const installedLength=length*linkCount;
 const rows=[
 {id:'NIC',name:srcNic.name,qty:linkCount,unit:'개 (NIC 1개/링크 가정)',status:'가정값 · 슬롯/지원 모델 검증',url:srcNic.officialUrl},
 {id:'SWITCH',name:srcSwitch.name,qty:switches,unit:'대',status:logicalPortCapacity?'논리 포트 기준 하한 · 실제 cage/breakout 검증':'모델/SKU 및 포트 불명 · 미산정',url:srcSwitch.officialUrl},
 {id:'CABLE',name:cpo?'MPO12 SMF 케이블 후보 (CPO switch-side)':'광 트렁크/점퍼 (커넥터 SKU 미선정)',qty:fibersCables,unit:'본 (1본/가상 링크)',status:'거리/극성/광규격/정격 확인 필요',url:srcSwitch.officialUrl},
 {id:'LENGTH',name:'광케이블 포설 총 길이 (여유 미포함)',qty:Number(installedLength.toFixed(2)),unit:'m',status:'도면 실측 아님 · 입력 거리 × 링크 수',url:null},
 {id:'OPT_SWITCH',name:'Switch-side 플러거블 광모듈',qty:switchSidePluggable,unit:'개',status:cpo?'Q3450 CPO측 모듈 불필요':'모듈·breakout 구조 미확정 · 미산정',url:srcSwitch.officialUrl},
 {id:'OPT_HOST',name:'NIC-side 호스트 광모듈',qty:null,unit:'개',status:'실물 NIC 포트·FEC·reach 확인 전 미산정',url:srcNic.officialUrl}
 ];
 return{schema:'lsdc/phase7-bom-audit/1.0',input:{racks,uplinkPerRack:uplink,lengthM:length,protocol:proto,mode,speedG:speed},
 products:{switch:{name:srcSwitch.name,url:srcSwitch.officialUrl,source:proto==='InfiniBand'?paths[0]:paths[1]},nic:{name:srcNic.name,url:srcNic.officialUrl,source:paths[1]}},
 logicalLinks:linkCount,switchesMinimum:switches,rows,notes,blocking,orderable:false,warning:'교육용 설계 초안. 기존 BOM 산출 엔진을 대체하지 않으며 실제 구매 수량이나 포트 호환을 보증하지 않습니다.'};
}
const panel=document.createElement('aside');panel.id='phase7-panel';panel.className='phase58-panel';panel.hidden=true;
panel.innerHTML=[
'<header><b>PHASE 7 · NVIDIA 제품/BOM 정합성 검사</b><button id="p7-close" type="button">✕</button></header>',
'<p class="p58-note">제품명·속도·인터페이스는 기존 NVIDIA 카탈로그(JSON)에서 로드합니다. 수량은 화면 입력으로 생성하는 검토용 초안이며 구매 승인 아님.</p>',
'<label>네트워크 프로토콜<select id="p7-proto"><option>InfiniBand</option><option>Ethernet</option></select></label>',
'<label>광 인터커넥트 모드<select id="p7-mode"><option value="pluggable">Pluggable</option><option value="cpo">CPO</option></select></label>',
'<label>링크 속도 (Gb/s)<select id="p7-speed"><option value="800">800</option><option value="400">400</option></select></label>',
'<label>GPU 랙 수<input type="number" id="p7-racks" min="1" max="200000" step="1"></label>',
'<label>랙당 uplink 수<input type="number" id="p7-links" value="2" min="1" max="8" step="1"></label>',
'<label>링크당 가상 거리 (m)<input type="number" id="p7-length" value="34" min=".1" max="10000" step=".1"></label>',
'<div class="p58-actions"><button id="p7-sync" type="button">제품 자료 새로고침</button><button id="p7-calc" type="button">BOM 감사 실행</button><button id="p7-json" type="button">검증 JSON</button><button id="p7-to-bom" type="button">기존 3D 구성 → BOM 도구</button></div>',
'<p id="p7-status" aria-live="polite">제품 카탈로그 로드 전.</p><div id="p7-summary"></div><div id="p7-table"></div>',
'<small class="p58-foot">정확한 주문형 SKU, MTP/MPO polarity, host-side 광모듈·reach, FEC, breakout 및 switch cage 소모량은 RFQ 이전 제조사 호환성 검토 필요.</small>'
].join('');document.body.appendChild(panel);
const btn=document.createElement('button');btn.id='phase7-open';btn.className='phase58-launch p7';btn.textContent='▣ NVIDIA / BOM';btn.type='button';view.appendChild(btn);
function inputs(){return{racks:+$('p7-racks').value,uplinkPerRack:+$('p7-links').value,lengthM:+$('p7-length').value,protocol:$('p7-proto').value,mode:$('p7-mode').value,speedG:+$('p7-speed').value}}
function render(){
 try{
  last=estimate(inputs());
  $('p7-status').textContent='검토 완료 · '+(last.blocking.length?'호환성 경고 '+last.blocking.length+'건':'명시적 차단 없음 (호환성 확정 아님)');
  $('p7-summary').innerHTML='<strong>링크 '+last.logicalLinks+'개 · 스위치 논리포트 기반 '+(last.switchesMinimum??'미산정')+'대</strong>'
   +last.blocking.map(x=>'<p class="p58-err">'+esc(x)+'</p>').join('')
   +last.notes.map(x=>'<p class="p58-note">'+esc(x)+'</p>').join('');
  $('p7-table').innerHTML='<table><thead><tr><th>BOM 품목</th><th>수량</th><th>검증 상태</th></tr></thead><tbody>'
   +last.rows.map(x=>'<tr><td>'+(x.url?'<a target="_blank" rel="noopener" href="'+esc(x.url)+'">'+esc(x.name)+'</a>':esc(x.name))+'</td><td>'+(x.qty===null?'미산정':x.qty.toLocaleString('ko-KR'))+' '+esc(x.unit)+'</td><td>'+esc(x.status)+'</td></tr>').join('')+'</tbody></table>';
 }catch(e){last=null;$('p7-status').textContent='검증 불가 · '+e.message;$('p7-table').innerHTML='';$('p7-summary').innerHTML=''}
}
function toggle(v){openState=v;panel.hidden=!v;btn.setAttribute('aria-expanded',String(v));if(!v)return;
 $('p7-racks').value=C.rackCount||20;$('p7-mode').value=C.mode==='cpo'?'cpo':'pluggable';
 load().then(render).catch(e=>$('p7-status').textContent=e.message);
}
btn.onclick=()=>toggle(!openState);$('p7-close').onclick=()=>toggle(false);$('p7-calc').onclick=render;
$('p7-sync').onclick=()=>{loading=null;load().then(render).catch(e=>$('p7-status').textContent=e.message)};
$('p7-mode').onchange=e=>{C.mode=e.target.value;const m=$('optMode');if(m){m.value=C.mode;m.dispatchEvent(new Event('change',{bubbles:true}))}render()};
$('p7-json').onclick=()=>{if(!last){$('p7-status').textContent='먼저 BOM 검증을 실행하세요.';return}
 const url=URL.createObjectURL(new Blob([JSON.stringify(last,null,2)],{type:'application/json'})),a=document.createElement('a');
 a.href=url;a.download='LS_Datacenter_NVIDIA_BOM_audit.json';a.click();setTimeout(()=>URL.revokeObjectURL(url),1200);
};
$('p7-to-bom').onclick=()=>{const old=$('bomOut');if(!old){$('p7-status').textContent='기존 BOM 공유 버튼을 찾을 수 없습니다.';return} $('p7-status').textContent='기존 3D 전체 구성(JSON)을 BOM 도구로 전달합니다. Phase 7 가상 링크 수량은 별도 감사 JSON으로 확인하십시오.';old.click()};
window.LS3D_PHASE7={open:()=>toggle(true),load,estimate,get catalog(){return catalog},get audit(){return last}};
})();