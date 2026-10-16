/* Mount after legacy editors have finished, then calculate the visible default example. */
(()=>{
 'use strict';
 const $=id=>document.getElementById(id),esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
 const num=n=>Number.isFinite(Number(n))?Number(n).toLocaleString('ko-KR',{maximumFractionDigits:2}):'미산정';
 const presets=[
  {name:'1k GPU · B200 Rail',values:{targetGPU:1024,systemId:'b200',topology:'rail',designInputMode:'gpu'}},
  {name:'256 GPU · H200 Dual',values:{targetGPU:256,systemId:'h200',topology:'dual',designInputMode:'gpu'}},
  {name:'4랙 · GB200 NVL72',values:{targetGPU:288,systemId:'gb200',topology:'rail',designInputMode:'racks',nvlRackCount:4,rackPowerKw:150,rackCoolingKw:150,coolingMode:'dlc'}}
 ];
 let mounted=false,applying=false,example=presets[0].name;
 function group(parent,title){const d=document.createElement('details');d.className='entry-options';const s=document.createElement('summary');s.textContent=title;d.append(s);parent.append(d);return d;}
 function mount(){
  if(mounted)return;
  const form=document.querySelector('.formPanel'),basic=$('basicInputs'),advanced=form?.querySelector('details.advanced');
  if(!basic||!advanced)return;
  mounted=true;document.body.classList.add('entry-ready');
  const fields=basic.querySelector('.inputGroup>.fields'),more=document.createElement('div');more.className='fields sub';
  const policy=group(advanced,'입력 방식 · 패브릭 · 배선');policy.append(more);
  const core=new Set(['targetGPU','systemId','topology']);
  [...fields.children].forEach(label=>{if(![...label.querySelectorAll('input,select')].some(x=>core.has(x.id)))more.append(label);});
  for(const id of ['itInputMw','nvlRackCount'])fields.append($(id).closest('label'));
  for(const id of ['network-policy','media-policy','route-procurement'])if($(id))policy.append($(id));
  const scope=$('wb-scenario')?.closest('label');if(scope)more.prepend(scope);
  // Colocation inputs must remain accessible when the user switches scope.
   // Moving them under a closed Advanced group makes the capacity workflow unreachable.
   const colo=form.querySelector(':scope>details.coloInputs');
   if(colo)colo.hidden=$('wb-scenario')?.value!=='colo';
   for(const d of [...form.querySelectorAll(':scope>details')])if(d!==advanced&&d!==colo){d.open=false;advanced.append(d);}
   if(colo)form.addEventListener('change',e=>{
     if(e.target.id!=='wb-scenario')return;
     colo.hidden=e.target.value!=='colo';
     if(!colo.hidden)colo.open=true;
   });
  // Preserve the input collector's containment contract: assumptions remain inside engineering-inputs.
  $('engineering-inputs').open=false;advanced.open=false;
  advanced.querySelector(':scope>summary').textContent='고급 옵션 · 랙·냉각 / Facility / 보조장비 / 가정값';
  $('inputEvidenceGuide').querySelector('summary').textContent='배선·예비 수량 계산 근거';
  for(const code of document.querySelectorAll('#inputEvidenceGuide code,#resultFormulaEvidence code'))if(code.textContent.startsWith('Leff ='))code.textContent='Leff = (Lroute + vertical rise + 2 × end allowance) × (1 + slack%)';
  const presetBox=document.createElement('section');presetBox.className='entry-presets';presetBox.innerHTML='<h3>예제로 시작하기</h3><p>프리셋을 누르면 고급 가정도 초기값으로 돌아가고 자동 계산됩니다.</p><div class="entry-preset-buttons">'+presets.map((p,n)=>`<button type="button" data-preset="${n}" aria-pressed="${n===0}">${esc(p.name)}</button>`).join('')+'</div><span id="entry-state" role="status">기본 예제 자동 계산 중 · 1,024 GPU</span>';
  basic.before(presetBox);
  const defaults=[...form.querySelectorAll('input,select')].map(el=>({el,value:el.value,checked:el.checked}));
  presetBox.addEventListener('click',e=>{
   const button=e.target.closest('[data-preset]');if(!button)return;applying=true;
   defaults.forEach(({el,value,checked})=>{el.value=value;if(el.type==='checkbox')el.checked=checked;});
   const p=presets[Number(button.dataset.preset)];for(const [id,value]of Object.entries(p.values))if($(id))$(id).value=value;
   example=p.name;
   // Reuse existing scope/custom-profile/cooling synchronizers and the debounced calculation.
   $('wb-scenario').dispatchEvent(new Event('change',{bubbles:true}));
   $('systemId').dispatchEvent(new Event('change',{bubbles:true}));
   presetBox.querySelectorAll('button').forEach(b=>b.setAttribute('aria-pressed',String(b===button)));
   advanced.open=false;applying=false;$('entry-state').textContent='예제 적용 중 · '+example;
  });
  const banner=document.createElement('section');banner.id='designConstraintSummary';banner.setAttribute('aria-label','설계 제약 검토');banner.innerHTML='<h2>설계 제약 검토</h2><p id="entry-check-state" role="status">기본 예제 계산 중…</p><div id="entry-check-cards" class="entry-check-grid"></div><div id="entry-alerts"></div><details id="entry-review-details"><summary>검토 사항과 모델 가정</summary><div id="entry-review-body"></div></details>';
  $('workbench-shell').prepend(banner);
  const evidence=document.createElement('details');evidence.id='entry-calculation-evidence';evidence.className='entry-evidence';evidence.innerHTML='<summary>산출 근거 · 전력·냉각 계산식 / 가정 / 출처</summary><div id="entry-formulas" class="entry-formula-grid"></div><div id="entry-sources"><p>프로파일과 PUE·예비율·냉각 기준은 계획 가정입니다. 고급 옵션 → 설계 가정에서 수정하면 결과에 반영됩니다. 부품 정격·호환성·시공 조건의 인증 판정이 아닙니다.</p><ul><li><a href="https://www.thegreengrid.org/resources/glossary?combine=pue" target="_blank" rel="noopener">The Green Grid · PUE 정의</a> — 에너지 비율 정의를 정전력 계획 모델에 적용합니다.</li><li><a href="https://docs.nvidia.com/dgx/dgxb200-user-guide/introduction-to-dgxb200.html" target="_blank" rel="noopener">NVIDIA DGX B200 사용자 가이드</a> — 8 GPU / 10U / 최대 14.3 kW 비교 출처. 다른 시스템·수정 프로파일은 해당 제조사 사양 확인이 필요합니다.</li><li><a href="https://github.com/kimheeseo/LSCNS/blob/main/DCI/DataCenter/engineering-v82.js" target="_blank" rel="noopener">실행 중인 계산 모델의 소스 코드</a> · <a href="https://github.com/kimheeseo/LSCNS/blob/main/DCI/DataCenter/design-checks.js" target="_blank" rel="noopener">제약 검사 코드</a></li></ul></div>';
  const result=document.querySelector('.resultPanel');$('kpis').after(evidence);
  if($('twinBomImportPanel')&&!location.hash.includes('twin=')){const bridge=group($('view-evidence'),'3D Digital Twin 구성 가져오기');bridge.append($('twinBomImportPanel'));bridge.open=location.hash.includes('twin=');}
  const supporting=group(result,'상세 수량·배선 검증 펼치기');supporting.classList.add('entry-supporting');
  for(const id of ['completeness-summary','planningSummary','engineering-audit','planningValidation','engineering-v82-results','resultFormulaEvidence'])if($(id)){
   const node=$(id);if(id==='planningSummary'&&node.nextElementSibling?.classList.contains('muted'))supporting.append(node.nextElementSibling);
   supporting.append(node);
  }
  form.addEventListener('input',changed);form.addEventListener('change',changed);
  function changed(e){if(e.target.dataset.referenceOnly)return;if(!applying){example='';presetBox.querySelectorAll('button').forEach(b=>b.setAttribute('aria-pressed','false'));$('entry-state').textContent='사용자 설정 · 자동 재계산';}setPending('입력이 변경되었습니다. 이전 결과를 재계산 중입니다.');}
  if(window.DCDesign)paint(window.DCDesign);
 }
 function setPending(message){if(!$('entry-check-state'))return;$('entry-check-state').textContent=message;$('designConstraintSummary').setAttribute('aria-busy','true');}
 function formula(title,code,detail){return `<article class="entry-formula"><h3>${esc(title)}</h3><code>${esc(code)}</code><p>${esc(detail)}</p></article>`;}
 function paint(r){
  if(!mounted)return;const c=window.DCDesignChecks.evaluate(r),banner=$('designConstraintSummary');banner.setAttribute('aria-busy','false');
   // Never present stored sample figures as if they were outputs for the user's input.
   const preview=$('entry-sample-result');
   if(preview){
     preview.hidden=Boolean(c.valid);
     const note=$('entry-sample-state');
     if(note&&!c.valid)note.textContent='계산 결과를 확인할 수 없어 저장된 예시를 표시합니다. 현재 입력값의 결과가 아니며, 계산 API 연결과 입력값을 확인해 주세요.';
   }
  $('topology').closest('label').querySelector('span').textContent='토폴로지 · 서버–Leaf 연결';
  const errors=c.issues.filter(x=>x.level==='error'),notes=c.issues.filter(x=>x.level!=='error');
  banner.dataset.severity=!c.valid||errors.length||c.errors.length?'error':notes.length||c.warnings.length?'warning':'ok';
  $('entry-check-state').textContent=!c.valid?'계산하지 못했습니다. 입력값을 확인하세요.':errors.length?`한도 초과 ${errors.length}종 · 아래 랙 전력·냉각·공간을 확인하세요.`:c.errors.length?`모델 오류 ${c.errors.length}건 · 상세 검토 필요`:c.metrics.some(m=>!m.known)?'랙별 데이터 부족 · 전력·냉각·RU 한도 검증 미완료':notes.length?`물리 한도 이내 · 계획 기준 주의 ${notes.length}종`:'전력·냉각·RU 입력 한도 이내 · 제품·시공 조건은 별도 검토';
  $('entry-state').textContent=(example?'예제 · '+example:'사용자 설정')+(c.valid?' · 계산 완료':' · 계산 실패');
  $('entry-check-cards').innerHTML=c.metrics.map(m=>`<div class="entry-check-card" data-failed="${m.count>0}"><span>${m.label} · ${m.known?(m.count?'초과':'한도 이내'):'미검증'}</span><b>${m.known?`${num(m.max)} / ${num(m.limit)}`:'데이터 없음'}</b><small>${m.unit} · ${m.key==='ru'?'사용 + 예약 RU':'랙별 최대 부하'}${m.known?' · '+esc(m.rack):''}</small></div>`).join('');
  $('entry-alerts').innerHTML=c.issues.map(x=>`<p class="entry-alert ${x.level}">${esc(x.message)}</p>`).join('');
  if(c.errors.length)$('entry-alerts').insertAdjacentHTML('beforeend',`<p class="entry-alert error">계산 모델 오류 ${c.errors.length}건 · ${esc(c.errors[0])}${c.errors.length>1?' · 전체 목록은 아래 펼치기':''}</p>`);
  const all=[...c.errors,...c.warnings];$('entry-review-details').querySelector('summary').textContent=`모델 오류 ${c.errors.length}건 · 검토 사항 ${c.warnings.length}건 펼치기`;
  $('entry-review-body').innerHTML='<p>랙 공간은 장비 RU + 예약 RU를 비교합니다. NVL72 Compute 랙은 OEM 전체 랙 RU를 비교합니다. 냉각 부하는 IT 전력이 모두 열로 전환된다는 계획 가정입니다.</p><ul>'+all.map(x=>'<li>'+esc(x)+'</li>').join('')+'</ul>';
  if(!c.valid){$('entry-formulas').textContent='유효한 결과가 없어 계산식에 값을 대입할 수 없습니다.';return;}
  if(r.version!=='8.2.0-engineering'){$('entry-formulas').textContent='이 결과는 별도 용량 계획 API에서 계산되었습니다. 근거 탭에서 해당 시나리오의 산정 기준을 확인하세요.';return;}
  const i=r.input,s=r.summary,p=r.systemProfile,f=r.facility,k=r.cooling,down=i.switchRadix/2,up=Math.max(1,Math.ceil(down/i.oversub));
  $('entry-formulas').innerHTML=
   formula('IT 전력 합계',`PIT = Σ(랙별 장비 수 × 장비 kW) = ${num(s.totalItPowerKw)} kW`,`Compute ${num(s.systemUnits)} × ${num(p.power)} = ${num(s.computePowerKw)} kW. 네트워크·스토리지·보조장비 부하를 합산합니다. 설치 수량 기준이며 예비 구매품은 제외합니다.`)+
   formula('시설전력 · PUE',`${num(s.totalItPowerKw)} kW × PUE ${num(i.pue)} = ${num(f.facilityPowerKw)} kW`,`추가 시설 부하 = ${num(f.facilityPowerKw-s.totalItPowerKw)} kW. PUE는 실측값이 아닌 입력 시나리오이며 냉각 전력과 시설 손실을 포괄합니다.`)+
   formula('랙 수 · 수용 한도',p.rackScale?`OEM 랙 수 = ${num(s.computeRacks)}`:`서버/랙 = max(1, min(RU ${s.rackLimits.RU}, 전력 ${s.rackLimits.power}, 냉각 ${s.rackLimits.cooling}${s.rackLimits.NIC!==undefined?', NIC '+s.rackLimits.NIC:''})) = ${num(s.systemsPerRack)}`,`RU 한도 = floor((${num(i.rackRU)} − ${num(i.reservedRU)}) / ${num(p.ru)}). 전력·냉각 한도 = floor(정격 × (1 − ${num(i.rackHeadroomPct)}%) / ${num(p.power)}). NVL72는 OEM 랙 단위입니다. 하나도 들어가지 않으면 검토용 1개를 표시하며 초과 경고를 냅니다.`)+
   formula('냉각 열부하 · 유량',`액냉 ${num(k.liquidKw)} + 공랭 ${num(k.airKw)} = IT ${num(s.totalItPowerKw)} kW`,`액냉 = Compute × 액냉 분담률(공랭 0%, 침지 100%, DLC/Rear-door ${num(i.liquidPct)}%). 유량 = ${num(k.liquidKw)} × 1000 × 60000 / (${num(i.waterDensity)} × ${num(i.waterCp)} × ${num(i.deltaTK)}) = ${num(k.flowLMin)} L/min. 냉각 전력 소비량과 열제거 용량은 다릅니다.`)+
   formula('Fabric · Oversubscription',`목표 ${num(i.oversub)}:1 → Leaf down ${num(down)} / up ${num(up)} = ${num(down/up)}:1`,`동일 링크 속도 기준 포트 비율. Up = ceil(Down / 목표 비율). 정수 포트로 올림하며 부분 점유 랙의 실효 비율·워크로드 처리량은 다를 수 있습니다. ${num(i.fabricTier)}-tier 요청, 실제 도메인 구성은 네트워크 탭에서 확인합니다.`)+
   formula('UPS · 시설 용량 가정',`UPS = PIT / PF / η × (1 + 시설 여유율) = ${num(f.upsDesignKva)} kVA`,`PF ${num(i.powerFactor)}, UPS 효율 ${num(i.upsEfficiency)}, 시설 여유 ${num(i.facilityHeadroomPct)}%. 현재 모델의 보수적 용량 근사입니다. 출력 정격·배터리 잔여시간 계산과 구분하며 실제 선정은 공급사 부하 조건을 확인해야 합니다. Facility BOM ${i.facilityEnabled?'포함':'미포함'}; 값은 용량 참고용입니다.`);
 }
 document.addEventListener('dc:inputs-ready',mount,{once:true});
 document.addEventListener('dc:calculating',()=>setPending('계산 중… 이전 결과는 새 계산이 끝나면 갱신됩니다.'));
 document.addEventListener('dc:design',e=>paint(e.detail));
})();
