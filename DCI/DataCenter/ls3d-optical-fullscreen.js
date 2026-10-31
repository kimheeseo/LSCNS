/* Full-screen optical network topology, rack hierarchy and engineering link budget.
   Optical module numbers are illustrative user-editable examples, not vendor specs. */
(function(){
'use strict';
const $=id=>document.getElementById(id);
const esc=x=>String(x).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const presets={
 '400-fr4':{label:'400G FR4 · SMF (가정)',speed:'400G',tx:-2,rx:-8,alpha:.35},
 '800-dr8':{label:'800G DR8 · SMF (가정)',speed:'800G',tx:-1,rx:-7,alpha:.35},
 '800-fr4':{label:'800G FR4 · SMF (가정)',speed:'800G',tx:0,rx:-7.5,alpha:.35},
 '1600-dr8':{label:'1.6T 2×DR8 · SMF (개념)',speed:'1.6T',tx:0,rx:-6,alpha:.35},
 custom:{label:'사용자 지정 모듈',speed:null,tx:0,rx:-7,alpha:.35}
};
const cfg={link:'tor-spine-a',speed:'800G',optic:'800-dr8',media:'SMF',length:34,tx:-1,rx:-7,alpha:.35,connectors:4,connLoss:.25,splices:2,spliceLoss:.1,extra:.15,reserve:1.5};
const numeric=x=>Number.isFinite(Number(x))?Number(x):0;
const fixed=x=>numeric(x).toFixed(3);
function topo(){try{return window.__LS3D_TEST__?.exportConfig?.()?.optical?.topology||null;}catch(_){return null;}}
function activeLink(){return (topo()?.links||[]).find(l=>l.id===cfg.link);}
function profile(id){return presets[id]||presets.custom;}
function numField(id,label,step,min){
 return '<label class="lo-field"><span>'+label+'</span><input type="number" id="lo-'+id+'" step="'+step+'" min="'+min+'"></label>';
}
function node(x,y,w,h,label,detail,type){
 const colors={cloud:['#234c69','#b9e2ff'],security:['#522d34','#ffd0c3'],router:['#254359','#bad9ff'],core:['#423954','#e4cfff'],spine:['#175262','#a1fff0'],leaf:['#254b63','#a9e5ff'],server:['#354253','#ecf1f8'],odf:['#514630','#ffe6a6']};
 const c=colors[type]||colors.server;
 let ports='';
 if(['router','core','spine','leaf'].includes(type))for(let i=0;i<9;i++)ports+='<rect x="'+(x-w/2+16+i*(w-32)/9).toFixed(1)+'" y="'+(y+h/2-10)+'" width="6" height="3" fill="'+c[1]+'" opacity=".8"/>';
 return '<g><rect x="'+(x-w/2)+'" y="'+(y-h/2)+'" width="'+w+'" height="'+h+'" rx="9" stroke="'+c[1]+'" stroke-width="1.4" fill="'+c[0]+'"/>'+ports+
 '<text x="'+x+'" y="'+(y-4)+'" text-anchor="middle" fill="#f3fdff" font-weight="750" font-size="15">'+esc(label)+'</text>'+
 '<text x="'+x+'" y="'+(y+12)+'" text-anchor="middle" fill="'+c[1]+'" font-size="11">'+esc(detail)+'</text></g>';
}
function schematic(){
 let wires=[],items=[];
 function edge(x1,y1,x2,y2,id,backup){
  const selected=id&&id===cfg.link;
  const color=selected?'#ffe28f':backup?'#8ba8ff':id?'#64e0d3':'#688da2';
  const attrs=id?' data-lo-link="'+esc(id)+'" tabindex="0" role="button" class="lo-pick-link" aria-label="'+esc(id)+' 링크 선택"':'';
  wires.push('<path d="M'+x1+' '+y1+' L'+x2+' '+y2+'" fill="none" stroke="'+color+'" stroke-width="'+(selected?5:id?2.4:1.8)+'"'+(backup?' stroke-dasharray="6 5"':'')+attrs+'/>');
 }
 edge(550,73,330,123);edge(550,73,770,123);
 edge(330,165,330,209);edge(770,165,770,209);
 edge(330,251,550,308);edge(770,251,550,308);
 edge(550,350,340,412);edge(550,350,760,412);
 edge(170,318,448,330);
 const xLeaf=[170,420,680,930];
 const paths=[['tor-spine-a','tor-spine-a-b'],['tor-spine-a','tor-spine-a-b'],['tor-spine-b-a','tor-spine-b'],['tor-spine-b-a','tor-spine-b']];
 xLeaf.forEach((x,i)=>{edge(340,454,x,533,paths[i][0],false);edge(760,454,x,533,paths[i][1],true);edge(x,581,x,642,'server-tor')});
 items.push(node(550,49,180,48,'Internet / ISP','외부 망','cloud'));
 items.push(node(330,144,165,42,'Firewall A','보안 경계 예시','security'),node(770,144,165,42,'Firewall B','보안 경계 예시','security'));
 items.push(node(330,230,165,42,'Edge Router A','BGP 예시','router'),node(770,230,165,42,'Edge Router B','BGP 예시','router'));
 items.push(node(550,329,200,42,'Core / Gateway','툴 내 Core 1대','core'));
 items.push(node(170,318,140,42,'ODF / MMR','광 패치 영역','odf'));
 items.push(node(340,433,172,42,'Spine A','IP Fabric · ECMP','spine'),node(760,433,172,42,'Spine B','IP Fabric · ECMP','spine'));
 ['Leaf A / ToR','Leaf B / ToR','Leaf C · 예시','Leaf D · 예시'].forEach((label,i)=>items.push(node(xLeaf[i],559,175,45,label,i<2?'모델 A/B':'확장 개념','leaf')));
 ['GPU Rack','Storage Rack','Compute Rack','Server Rack'].forEach((label,i)=>items.push(node(xLeaf[i],667,175,46,label,'NIC · Fiber / DAC','server')));
 return '<svg id="loHierarchy" viewBox="0 0 1100 725" role="img" aria-label="외부 인터넷과 방화벽, 라우터, 코어, Spine A/B, Leaf, 서버랙의 네트워크 계층도">'+wires.join('')+items.join('')+
 '<text x="550" y="495" text-anchor="middle" fill="#8bebd6" font-size="13">Spine ↔ Leaf · L3 ECMP · 선택적 EVPN/VXLAN</text>'+
 '<text x="550" y="708" text-anchor="middle" fill="#bad4e0" font-size="13">실선: 주 경로 예시 · 점선: 이중 경로 · 선을 클릭해 광 예산 확인</text></svg>';
}
function calc(){
 const fiber=cfg.length/1000*cfg.alpha,conn=cfg.connectors*cfg.connLoss,splice=cfg.splices*cfg.spliceLoss;
 const loss=fiber+conn+splice+cfg.extra,rx=cfg.tx-loss,budget=cfg.tx-cfg.rx,gross=rx-cfg.rx,margin=gross-cfg.reserve;
 return {fiber,conn,splice,loss,rx,budget,gross,margin,optical:cfg.media==='SMF'||cfg.media==='MMF'};
}
function sync(){
 for(const key of Object.keys(cfg)){const el=$('lo-'+key);if(el)el.value=cfg[key];}
}
function setLink(id){
 const l=(topo()?.links||[]).find(x=>x.id===id);if(!l)return;
 cfg.link=id;cfg.length=Math.max(0,numeric(l.lengthM));
 cfg.connectors=Math.max(0,numeric(l.connectors));
 cfg.splices=Math.max(0,numeric(l.splices));
 cfg.media=l.type==='DAC/AEC'?'DAC/AEC':l.type==='AOC'?'AOC':l.media==='MMF'?'MMF':'SMF';
 cfg.alpha=cfg.media==='MMF'?3.0:.35;
 const old=$('optDistance');if(old&&window.LS3D_CONFIG?.selected===id)old.value=cfg.length;
 sync();render();
}
function render(){
 const links=topo()?.links||[],sel=$('lo-link');
 if(sel){sel.innerHTML=links.map(l=>'<option value="'+esc(l.id)+'">'+esc(l.from+' → '+l.to+' · '+l.type)+'</option>').join('');sel.value=cfg.link;}
 $('loHierarchy')?.querySelectorAll('[data-lo-link]').forEach(path=>{
  const b=path.dataset.loLink===cfg.link;path.setAttribute('stroke',b?'#ffe28f':path.hasAttribute('stroke-dasharray')?'#8ba8ff':'#64e0d3');path.setAttribute('stroke-width',b?5:2.4);
 });
 const v=calc(),p=profile(cfg.optic),l=activeLink();
 const warning=[];
 if(l?.down)warning.push('원본 토폴로지에서 이 링크는 단선 상태입니다.');
 if(p.speed&&p.speed!==cfg.speed)warning.push('모듈 예시 속도와 포트 속도가 일치하지 않습니다.');
 if(cfg.media==='AOC'||cfg.media==='DAC/AEC')warning.push('DAC/AEC 또는 AOC는 이 수동 광손실 계산식의 대상이 아닙니다. 능동 케이블의 도달거리와 BER 데이터시트를 확인하세요.');
 if(cfg.length>2000&&cfg.optic.includes('dr8'))warning.push('DR 계열로 장거리 전송을 가정했습니다. 실제 모듈 최대 도달거리 및 분산·BER 조건을 확인하세요.');
 const amount=(label,value)=>'<div class="lo-amount"><small>'+label+'</small><b>'+value+'</b></div>';
 const step=(label,expression,result)=>'<div class="lo-math"><span class="lo-key">'+label+'</span><br>'+expression+'<br><b>'+result+'</b></div>';
 $('lo-results').innerHTML='<p><b>'+esc(l?l.from+' → '+l.to+' · '+l.type:cfg.link)+'</b><br>'+esc(cfg.speed)+' · '+esc(p.label)+' · '+esc(cfg.media)+' · '+cfg.length+'m</p>'+
 '<div class="lo-amounts">'+amount('총 광 손실',v.optical?fixed(v.loss)+' dB':'N/A')+
 amount('수신 최소 파워',v.optical?fixed(v.rx)+' dBm':'N/A')+
 amount('설계 Link Margin',v.optical?fixed(v.margin)+' dB':'N/A')+'</div>'+
 step('① 광섬유 감쇠', 'L × α = '+fixed(cfg.length/1000)+' km × '+fixed(cfg.alpha)+' dB/km',fixed(v.fiber)+' dB')+
 step('② 총 광손실', 'L×α + Nconn×Lconn + Nsplice×Lsplice + Lextra<br> = '+fixed(v.fiber)+' + '+fixed(v.conn)+' + '+fixed(v.splice)+' + '+fixed(cfg.extra),fixed(v.loss)+' dB')+
 step('③ 수신 최소 광파워', 'Ptx(min) − Ltotal = '+fixed(cfg.tx)+' − '+fixed(v.loss),fixed(v.rx)+' dBm')+
 step('④ 여유 링크 마진', 'Ptx(min) − Prx(sensitivity) − Ltotal − Reserve<br> = '+fixed(cfg.tx)+' − ('+fixed(cfg.rx)+') − '+fixed(v.loss)+' − '+fixed(cfg.reserve),fixed(v.margin)+' dB')+
 '<div class="lo-math">광모듈 예산 = Ptx(min) − Prx(sensitivity) = <b>'+fixed(v.budget)+' dB</b><br>설계 예비 차감 전 여유 = '+fixed(v.gross)+' dB<br>광 마진 판정: <b>'+(v.optical?(v.margin>=0?'예시 계산 기준 만족':'예시 예산 부족'):'적용 대상 아님')+'</b></div>'+
 (warning.length?'<div class="lo-warning">'+warning.map(esc).join(' ')+'</div>':'')+
 '<p style="font-size:11px;color:#b8cbd8">광모듈 Tx/Rx 수치는 설명용 가정값이며 제조사 보증 사양이 아닙니다. 데이터센터 고속 광링크의 OMA, TDECQ, FEC, lane 구조, 분산·모듈 호환성은 별도 검증이 필요합니다.</p>';
}
function init(){
 const modal=$('optModal'),button=$('optOpen'),panel=modal?.querySelector('.twin-panel');
 if(!modal||!button||!panel||$('loDashboard'))return;
 modal.setAttribute('role','dialog');modal.setAttribute('aria-modal','true');modal.setAttribute('aria-label','전체화면 광 네트워크 설계');
 const section=document.createElement('section');section.id='loDashboard';
 const moduleOptions=Object.keys(presets).map(k=>'<option value="'+k+'">'+esc(presets[k].label)+'</option>').join('');
 const entries=[
  numField('length','케이블 거리 (m)',1,0),numField('tx','광 모듈 최소 Tx (dBm)',.1,-50),
  numField('rx','광 모듈 Rx 감도 (dBm)',.1,-50),numField('alpha','섬유 감쇠 (dB/km)',.01,0),
  numField('connectors','커넥터 결합 수 (개)',1,0),numField('connLoss','결합당 손실 (dB)',.01,0),
  numField('splices','융착접속 수 (개)',1,0),numField('spliceLoss','융착당 손실 (dB)',.01,0),
  numField('extra','기타 결합·패널 손실 (dB)',.01,0),numField('reserve','설계 예비 마진 (dB)',.1,0)
 ].join('');
 section.innerHTML='<div class="lo-topline"><div><h3>데이터센터 네트워크 · 전체 계층 구조</h3><div class="lo-sub">참고 그림 기반의 설계 개념도입니다. Firewall·Edge Router·Leaf C/D는 확장 예시로, 현재 툴의 실제 설치 수량에 포함되지 않습니다.</div></div><button id="lo-close" type="button">전체화면 닫기</button></div>'+
 '<div class="lo-grid"><section class="lo-section"><h4>Internet → Firewall → Router → Core → Spine → Leaf → Server</h4>'+schematic()+
 '<div class="lo-bottom"><b>전체 네트워크 랙 · 계층별 역할</b><table><thead><tr><th>계층</th><th>랙/장비</th><th>역할</th></tr></thead><tbody>'+
 '<tr><td>보안 · WAN</td><td>Firewall / Edge Router</td><td>외부 트래픽 보안·경계 라우팅 (개념)</td></tr>'+
 '<tr><td>Core / Spine</td><td>Core Router · Spine A/B</td><td>IP Fabric·ECMP 트래픽 집선</td></tr>'+
 '<tr><td>Leaf / ToR</td><td>Leaf A/B · Network Rack</td><td>서버 NIC 직접 연결 및 Uplink</td></tr>'+
 '<tr><td>서버</td><td>GPU / Compute / Storage Rack</td><td>연산·처리·저장, NIC 탑재</td></tr>'+
 '<tr><td>광 배선</td><td>ODF / MMR / Patch Panel</td><td>수동 광 접속·분배, L3 스위칭 아님</td></tr></tbody></table>'+
 '<p>MLAG는 일부 Leaf 쌍에 적용할 수 있는 L2 이중화 기술이며, Spine 간 모든 연결에 MLAG가 필요한 것은 아닙니다. VXLAN/EVPN 및 BGP는 구성에 따라 선택됩니다.</p></div></section>'+
 '<section class="lo-section"><h4>광모듈·포트속도별 링크 마진 계산</h4><div class="lo-fields">'+
 '<label class="lo-field"><span>링크 선택</span><select id="lo-link"></select></label>'+
 '<label class="lo-field"><span>포트 속도</span><select id="lo-speed"><option>400G</option><option>800G</option><option>1.6T</option></select></label>'+
 '<label class="lo-field"><span>광 모듈 가정</span><select id="lo-optic">'+moduleOptions+'</select></label>'+
 '<label class="lo-field"><span>매체</span><select id="lo-media"><option>SMF</option><option>MMF</option><option>DAC/AEC</option><option>AOC</option></select></label>'+entries+'</div>'+
 '<div id="lo-results" aria-live="polite"></div>'+
 '<div class="lo-bottom">포트 속도만으로 광모듈의 Tx 최소출력·Rx 감도를 확정할 수 없습니다. 각 값은 직접 수정할 수 있는 교육용 예시이며, 실제 설계는 광모듈 데이터시트로 검증해야 합니다.</div></section></div>';
 const after=panel.querySelector('p');
 if(after)after.insertAdjacentElement('afterend',section);else panel.prepend(section);
 $('lo-close').onclick=()=>modal.classList.remove('open');
 for(const k of ['length','tx','rx','alpha','connectors','connLoss','splices','spliceLoss','extra','reserve']){
  $('lo-'+k).addEventListener('input',()=>{cfg[k]=Math.max(k==='tx'||k==='rx'?-50:0,numeric($('lo-'+k).value));render()});
 }
 $('lo-optic').addEventListener('change',()=>{cfg.optic=$('lo-optic').value;const p=profile(cfg.optic);cfg.tx=p.tx;cfg.rx=p.rx;cfg.alpha=p.alpha;sync();render();});
 $('lo-media').addEventListener('change',()=>{cfg.media=$('lo-media').value;cfg.alpha=cfg.media==='MMF'?3:.35;sync();render();});
 $('lo-speed').addEventListener('change',()=>{cfg.speed=$('lo-speed').value;cfg.optic=cfg.speed==='400G'?'400-fr4':cfg.speed==='1.6T'?'1600-dr8':'800-dr8';const p=profile(cfg.optic);cfg.tx=p.tx;cfg.rx=p.rx;cfg.alpha=p.alpha;
   const original=$('optSpeed');if(original){original.value=cfg.speed;original.dispatchEvent(new Event('change',{bubbles:true}));}sync();render();});
 $('lo-link').addEventListener('change',()=>setLink($('lo-link').value));
 section.querySelectorAll('[data-lo-link]').forEach(path=>{
  path.addEventListener('click',()=>setLink(path.dataset.loLink));
  path.addEventListener('keydown',e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();setLink(path.dataset.loLink)}});
 });
 button.addEventListener('click',()=>{
  if($('optSpeed')&&$('optSpeed').value!==cfg.speed){cfg.speed=$('optSpeed').value;cfg.optic=cfg.speed==='400G'?'400-fr4':cfg.speed==='1.6T'?'1600-dr8':'800-dr8';const p=profile(cfg.optic);cfg.tx=p.tx;cfg.rx=p.rx;cfg.alpha=p.alpha;}
  const id=window.LS3D_CONFIG?.selected;
  setLink((topo()?.links||[]).some(l=>l.id===id)?id:cfg.link);
 });
 document.addEventListener('keydown',e=>{if(e.key==='Escape'&&modal.classList.contains('open'))modal.classList.remove('open')});
 const current=topo()?.links||[];if(current.length)setLink(current.some(l=>l.id===cfg.link)?cfg.link:current[0].id);
}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init,{once:true});else init();
})();