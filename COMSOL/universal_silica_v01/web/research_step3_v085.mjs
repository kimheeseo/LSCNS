/* Display reproducible Python computations; never run Maxwell FEM in the browser. */
const ROOT='../fiber_research_v1/benchmarks/step3_v085/';
export function validateModeResult(d){
  if(d?.schema!=='LSCNS_fiber_research_vector_v1'||!d.input||!Array.isArray(d.modes)||!d.modes.length||d.modes.length>64)throw Error('지원하는 벡터 결과 JSON과 input이 필요합니다');
  for(const m of d.modes){
    for(const k of ['neff_real','neff_imag','residual_relative','weak_gauss_relative'])if(!Number.isFinite(m[k]))throw Error(k+' 유한성 오류');
    const f=m.field_samples,points=f?.coordinates_um;
    if(!Array.isArray(points)||!points.length||points.length>8000||!points.every(p=>Array.isArray(p)&&p.length===2&&p.every(Number.isFinite)))throw Error('유한한 필드 좌표가 필요합니다 (최대 8,000점)');
    for(const k of ['Ex','Ey','Ez','Hx','Hy','Hz']){
      const v=f.components?.[k];
      if(!Array.isArray(v)||v.length!==points.length||!v.every(z=>Array.isArray(z)&&z.length===2&&z.every(Number.isFinite)))throw Error(k+' 복소 필드 배열 오류');
    }
  }
  return d;
}
async function getJSON(name){
  const r=await fetch(ROOT+name,{cache:'no-store'});if(!r.ok)throw Error('HTTP '+r.status);return r.json();
}
export function mountStep3(section,{updateConfig,displayResult}){
  const panel=document.createElement('div');panel.id='research-step3-v085';
  panel.style.cssText='margin:20px 0;padding:16px;border:1px solid #4e879d;border-radius:10px;scroll-margin-top:90px;overflow-wrap:anywhere';
  panel.innerHTML='<h3>STEP 3 · 풀벡터 모드 솔버 v0.8.5</h3>'+
    '<p>Nédélec edge + H1 혼합 FEM · 복소 β² sparse shift-invert · 선택적 companion 교차 검사 · 원형 PML / 1차 SBC. Python(로컬/Colab)에서 실행하며 브라우저는 결과를 표시합니다.</p>'+
    '<div class="rf-line"><button id="rf-step3-guided">실계산 HE11 예시 적용</button><button id="rf-step3-radiative">실계산 방사 모드 예시 적용</button></div>'+
    '<p class="rf-muted">버튼은 합성 기준 시험의 입력과 이미 계산된 JSON을 함께 적용합니다. HE11: n=1.50/1.44, 코어 반경 1.3 µm. 방사 시험: n=1.00/1.45, 반경 3 µm. 모두 λ=1.55 µm의 수치 검증용 조건입니다. 실제 저손실 HCF의 정확도는 별도 검증이 필요합니다.</p>'+
    '<p id="rf-step3-state" role="status">독립 기준 검증표 불러오는 중…</p>'+
    '<div style="overflow-x:auto"><table style="width:100%;border-collapse:collapse"><thead><tr><th>시험</th><th>결과</th><th>비교값</th></tr></thead><tbody id="rf-step3-table"></tbody></table></div>'+
    '<p>V1: 정확한 원형 HE11 해 · V4: 복소 TE11, 양의 전력, 잔차와 독립 편광 쌍 · V5: outgoing Hankel 방사 해와 메시/창/PML 두께·강도·차수 비교. PASS_SCOPED는 해당 합성 시험에 한정됩니다.</p>'+
    '<p><a href="'+ROOT+'step3_validation.json" target="_blank" rel="noopener">실제 검증 수치 JSON</a> · <a href="https://github.com/kimheeseo/LSCNS/blob/main/COMSOL/validation_report/STEP3_V085_VERIFICATION_20261119.md" target="_blank" rel="noopener">검증표·실행 방법</a></p>';
  section.insertBefore(panel,section.querySelector('#rf-result-open').parentElement.previousElementSibling);
  const state=panel.querySelector('#rf-step3-state');
  const fmt=(x,percent=false)=>Number.isFinite(x)?(percent?(100*x).toPrecision(4)+'%':x.toExponential(3)):'미확인';
  getJSON('step3_validation.json').then(d=>{
    if(d.schema!=='LSCNS_step3_validation_v085')throw Error('검증표 형식 불일치');
    const rows=[['V1',d.V1?.status,'최대 |ΔRe| '+fmt(d.V1?.maximum_absolute_Re_error)],
      ['V4',d.V4?.status,'복소 TE11 최대 오차 '+fmt(d.V4?.uniform_TE11_reference?.maximum_complex_error)],
      ['V5',d.V5?.status,'최대 Im 상대오차 '+fmt(d.V5?.maximum_relative_Im_error,true)+' / 분산 '+fmt(d.V5?.relative_Im_spread,true)]];
    const body=panel.querySelector('#rf-step3-table');
    for(const row of rows){const tr=document.createElement('tr');for(const value of row){const td=document.createElement('td');td.textContent=value??'미확인';td.style.padding='8px';tr.append(td);}body.append(tr);}
    state.textContent='검증 실행: '+d.timestamp_utc+' · 제품 손실 인증: '+(d.product_loss_certified===true?'기록 검토 필요':'아님');
  }).catch(e=>{state.textContent='검증표 읽기 실패: '+e.message;});
  for(const kind of ['guided','radiative'])panel.querySelector('#rf-step3-'+kind).addEventListener('click',async e=>{
    const b=e.currentTarget;b.disabled=true;
    try{
      const result=validateModeResult(await getJSON(kind+'_mode_result.json'));
      updateConfig(result.input);displayResult(result);
      state.textContent='실제 Python '+kind+' 예시 적용 완료 · 현재 입력과 일치 · 제품 손실 인증 아님';
    }catch(err){state.textContent='예시 읽기 실패: '+err.message;}finally{b.disabled=false;}
  });
}
