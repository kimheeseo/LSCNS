/**
 * Full-vector (2D coupled H-field FDFD) results viewer for existing scalar FEM Web.
 * This module never claims that GitHub Pages runs the vector solver.
 * It accepts only schema LSCNS_vector_mode_v1 from the independent Colab/Python solver.
 */
const NOTEBOOK='https://colab.research.google.com/github/kimheeseo/LSCNS/blob/main/COMSOL/universal_silica_v01/full_vector/LS_Full_Vector_Fiber_v0_1.ipynb';
const EXAMPLE='../full_vector/results/vector_mode_result.json';
const $=(id)=>document.getElementById(id);
const fmt=(v,d=5)=>Number.isFinite(v)?Number(v).toFixed(d):'—';
const css=`
#full-vector-tool{margin-top:24px;scroll-margin-top:85px}
#full-vector-tool .fv-summary{color:#b5c8db;font-size:12px;line-height:1.65;margin:12px 0}
#full-vector-tool .fv-actions{display:flex;gap:10px;flex-wrap:wrap;margin:16px 0}
#full-vector-tool .fv-actions a,#full-vector-tool .fv-actions button,#full-vector-tool .fv-upload{padding:10px 16px;border-radius:9px;background:#193a58;color:#eaf8ff;border:1px solid #39758f;font-size:12px;text-decoration:none;font-weight:650;cursor:pointer}
#full-vector-tool .fv-actions a:hover,#full-vector-tool .fv-actions button:hover,#full-vector-tool .fv-upload:hover{background:#225575}
#full-vector-tool .fv-badge{border:1px solid #5b7189;color:#b5c7d9;border-radius:20px;padding:5px 9px;font-size:10px}
#full-vector-tool .fv-status{font-size:12px;color:#8ee7d6;min-height:18px;line-height:1.65;margin:4px 0 14px}
#full-vector-tool .fv-metrics{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:10px;margin:8px 0 18px}
#full-vector-tool .fv-metrics article{padding:15px 12px;background:#10233a;border:1px solid #284967;border-radius:10px;min-width:0}
#full-vector-tool .fv-metrics small{font-size:10px;color:#a9bdd2;display:block;margin-bottom:9px}
#full-vector-tool .fv-metrics strong{display:block;font-size:clamp(14px,1.8vw,24px);letter-spacing:-.5px;color:#f2f9ff;white-space:nowrap}
#full-vector-tool .fv-metrics span{display:block;font-size:10px;color:#829eb8;margin-top:5px}
#full-vector-tool .fv-plots{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px}
#full-vector-tool .fv-map{padding:10px;background:#0a1c31;border:1px solid #29415e;border-radius:10px;min-width:0}
#full-vector-tool .fv-map h4{margin:4px 0 8px;font-size:12px;color:#cfebf1}
#full-vector-tool .fv-map canvas{display:block;width:100%;height:auto;background:#081527;border-radius:5px}
#full-vector-tool .fv-help{margin-top:10px;font-size:11px;line-height:1.7;color:#9db2cc}
#full-vector-tool .fv-alert{color:#ffcf91}
@media(max-width:950px){#full-vector-tool .fv-metrics{grid-template-columns:repeat(2,minmax(0,1fr))}}
@media(max-width:610px){#full-vector-tool .fv-plots{grid-template-columns:1fr}}
`;
function mount(){
 const main=document.querySelector('main.workspace')||document.querySelector('main');if(!main)return;
 const style=document.createElement('style');style.textContent=css;document.head.appendChild(style);
 const sec=document.createElement('section');sec.id='full-vector-tool';sec.className='engineering-panel';
 sec.innerHTML=`
 <div class="eng-head"><div><span class="kicker">07 / VECTOR MODE LAB</span><h2>Full-vector 모드 / 전기장 적분</h2></div><span class="fv-badge">2D VECTOR FDFD · Colab</span></div>
 <p class="fv-summary">원형 실리카 광섬유의 <b>결합된 Hx·Hy 벡터 고유모드</b>를 Python Colab에서 실제 계산하고, Maxwell 방정식으로 복원한 Ex·Ey·Ez의 강도 적분에서 MFD와 Aeff를 구합니다. 기존 웹의 <b>Scalar LP01 FEM</b>과는 다른 독립 엔진입니다.</p>
 <div class="fv-actions"><a href="${NOTEBOOK}" target="_blank" rel="noopener noreferrer">↗ Colab에서 Full-vector 계산</a><button type="button" id="fv-example">저장된 기본 결과 보기</button><label class="fv-upload" for="fv-json-file">결과 JSON 가져오기</label><input id="fv-json-file" type="file" accept=".json,application/json" hidden/></div>
 <div id="fv-status" class="fv-status" role="status" aria-live="polite">Colab 계산 결과 JSON을 선택하거나 기본 결과를 불러오세요.</div>
 <div class="fv-metrics"><article><small>Vector n_eff</small><strong id="fv-neff">—</strong><span>HE-like guided mode</span></article><article><small>MFD / second moment</small><strong id="fv-mfd">—</strong><span>µm · ∫ |E|²</span></article><article><small>Aeff / vector intensity</small><strong id="fv-aeff">—</strong><span>µm² · ∫ |E|² / ∫ |E|⁴</span></article><article><small>MFD / PII-like</small><strong id="fv-pii">—</strong><span>µm · near-field estimate</span></article></div>
 <div class="fv-plots">
 <div class="fv-map"><h4>전체 |E|²</h4><canvas id="fv-map-total" width="360" height="360"></canvas></div>
 <div class="fv-map"><h4>|Ex|²</h4><canvas id="fv-map-ex" width="360" height="360"></canvas></div>
 <div class="fv-map"><h4>|Ey|²</h4><canvas id="fv-map-ey" width="360" height="360"></canvas></div>
 <div class="fv-map"><h4>|Ez|²</h4><canvas id="fv-map-ez" width="360" height="360"></canvas></div>
 </div>
 <p id="fv-details" class="fv-help">각 성분 그림은 패턴 확인을 위해 개별 최대값으로 정규화됩니다. 성분의 상대 에너지 비율은 결과 메타데이터를 확인하십시오.</p>
 <p class="fv-help fv-alert">주의: 위의 기존 Colab 결과 뷰어는 계산하지 않고 <b>JSON을 불러와 표시</b>합니다. 아래 M3 섹션에서 별도 브라우저 벡터 FDFD를 실행할 수 있습니다. Scalar 측면 입력값을 바꾸더라도 가져온 결과는 자동 재계산되지 않습니다. 이 FDFD는 FEM이 아니며 Full-vector 손실/PML/HCF 복소 모드 해석도 아닙니다.</p>`;
 const anchor=$('pmd-tool');if(anchor)anchor.before(sec);else main.appendChild(sec);
 $('fv-example').addEventListener('click',async()=>{
   $('fv-status').textContent='기본 벡터 해석 결과를 불러오는 중…';
   try{const res=await fetch(EXAMPLE,{cache:'no-store'});if(!res.ok)throw Error('HTTP '+res.status);show(await res.json(),'GitHub에 저장된 기준 계산');}
   catch(err){$('fv-status').textContent='불러오기 실패: '+err.message;}
 });
 $('fv-json-file').addEventListener('change',async e=>{
   const f=e.target.files?.[0];if(!f)return;
   if(f.size>12_000_000){$('fv-status').textContent='JSON 크기 제한 12 MB';return;}
   try{show(JSON.parse(await f.text()),f.name);}catch(err){$('fv-status').textContent='결과 읽기 실패: '+err.message;}
 });
 const nav=document.querySelector('.top-link');
 // Do not modify the existing model navigator; append only if a sidebar nav exists.
 const toolbar=document.querySelector('.top-right');if(toolbar && !$('fv-nav')){
  const a=document.createElement('a');a.id='fv-nav';a.href='#full-vector-tool';a.className='top-link';a.textContent='Vector 모드';toolbar.prepend(a);
 }
}
function show(data,source){
 if(data?.schema!=='LSCNS_vector_mode_v1'||!Array.isArray(data.modes)||!data.modes.length)throw Error('지원하지 않는 Full-vector JSON schema');
 const m=data.modes[0]; const f=m.field_map;
 if(!f?.intensity_normalized_to_max_total?.Itotal)throw Error('JSON에 전기장 강도 맵이 없습니다');
 for(const k of ['neff','MFD_2nd_um','Aeff_um2','MFD_PII_um']) if(!Number.isFinite(m[k]))throw Error(k+' 값이 유효하지 않습니다');
 $('fv-neff').textContent=fmt(m.neff,9);$('fv-mfd').textContent=fmt(m.MFD_2nd_um,3);
 $('fv-aeff').textContent=fmt(m.Aeff_um2,3);$('fv-pii').textContent=fmt(m.MFD_PII_um,3);
 const keys=[['fv-map-total','Itotal'],['fv-map-ex','Ex'],['fv-map-ey','Ey'],['fv-map-ez','Ez']];
 for(const [canvas,key] of keys)drawField($(canvas),f.intensity_normalized_to_max_total[key]);
 const me=data.input||{};
 $('fv-status').textContent=`✓ ${source} · ${fmt(me.wavelength_nm,0)} nm / ${me.profile||'unknown'} / ${me.grid_n||'?'}×${me.grid_n||'?'} · 벡터 계산 스냅샷`;
 const nErr=data.validation?.abs_neff_error_vs_exact;
 $('fv-details').textContent=`모드 1 · 외곽 에너지 분율 ${m.edge_power_fraction?.toExponential?.(2)||'—'} · Ez 상대 강도 적분비 ${(100*(m.Ez_intensity_fraction||0)).toFixed(3)}% · ${Number.isFinite(nErr)?'독립 HE11 neff 오차 '+nErr.toExponential(2):'HE11 기준 해당 없음'} · 성분별 이미지 자체 정규화. 새 입력은 Colab에서 재실행해야 합니다.`;
}
function drawField(canvas,map){
 if(!Array.isArray(map)||map.length<2||!Array.isArray(map[0]))throw Error('필드 맵 형식 오류');
 const n=map.length,ctx=canvas.getContext('2d');if(!ctx)return;
 const max=Math.max(1e-30,...map.map(r=>Math.max(...r)));
 const size=canvas.width,px=size/n;
 ctx.fillStyle='#071527';ctx.fillRect(0,0,size,size);
 for(let y=0;y<n;y++)for(let x=0;x<n;x++){
  const z=Math.pow(Math.max(0,Math.min(1,Number(map[y][x]||0)/max)),.58);
  const R=Math.floor(7+239*z**1.65),G=Math.floor(27+183*z),B=Math.floor(52+90*(1-z));
  ctx.fillStyle=`rgb(${R},${G},${B})`;ctx.fillRect(x*px,y*px,px+0.8,px+0.8);
 }
 ctx.strokeStyle='#d4efff88';ctx.strokeRect(.5,.5,size-1,size-1);
}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',mount,{once:true});else mount();
