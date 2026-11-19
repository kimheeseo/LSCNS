/* STEP6 Python execution and result viewer. Browser never substitutes a scalar FEM. */
const ROOT='../fiber_research_v1/benchmarks/step6_v088/';
const $=id=>document.getElementById(id);
const stable=v=>v&&typeof v==='object'?Array.isArray(v)?v.map(stable):Object.fromEntries(Object.keys(v).sort().map(k=>[k,stable(v[k])])):v;
const equal=(a,b)=>JSON.stringify(stable(a))===JSON.stringify(stable(b));
const esc=v=>String(v??'—').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const fmt=v=>typeof v==='number'?Number.isFinite(v)?v.toPrecision(7):'—':v==null?'—':String(v);
const kinds=['bend','pm','mcf','splice','sweep','dgd'];
const names=['6a 굽힘','6b PM 열응력','6c MCF','6d 접속','6e 스윕·Pareto','6f PMD 경로'];
const examples=['bend.json','pm.json','mcf.json','splice_fem_smf_hcf.json','grid.json',null];
export function validateStep6Result(r){
 if(!/^LSCNS_step6_(bend|pm|mcf|splice|grid|dgd)_v088$/.test(r?.schema))throw Error('STEP6 v0.8.8 결과 JSON이 필요합니다');
 if(r.certified!==false&&r.validation?.certified!==false)throw Error('검증 상태가 명시되지 않았습니다');
 const walk=(v,depth=0)=>{if(depth>60)throw Error('JSON 깊이 초과');if(typeof v==='number'&&!Number.isFinite(v))throw Error('유한하지 않은 결과');if(v&&typeof v==='object')for(const x of Object.values(v))walk(x,depth+1);};walk(r);
 const type=r.schema.split('_')[2];
 if(['bend','grid'].includes(type)&&!Array.isArray(r.cases))throw Error('cases 누락');
 if(type==='mcf'&&!Array.isArray(r.modes))throw Error('modes 누락');
 if(type==='pm'&&!r.modal_birefringence)throw Error('modal_birefringence 누락');
 if(type==='splice'&&(!Number.isFinite(r.eta_Et)||r.eta_Et<0||r.eta_Et>1))throw Error('0≤eta≤1 필요');
 return r;
}
function download(obj,name){const u=URL.createObjectURL(new Blob([JSON.stringify(obj,null,2)+'\n'],{type:'application/json'}));const a=document.createElement('a');a.href=u;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(u),8000);}
function numbers(id){const raw=$(id).value.split(',').map(s=>s.trim());const v=raw.map(Number);if(raw.some(x=>!x)||v.some(x=>!Number.isFinite(x)))throw Error('쉼표로 구분한 유한 숫자가 필요합니다');return v;}
function table(headers,rows){$('r688-table').innerHTML='<thead><tr>'+headers.map(s=>'<th>'+esc(s)+'</th>').join('')+'</tr></thead><tbody>'+rows.map(row=>'<tr>'+row.map(v=>'<td>'+esc(fmt(v))+'</td>').join('')+'</tr>').join('')+'</tbody>';}
function plot(points,xLabel,yLabel){const c=$('r688-chart');const ctx=c.getContext('2d');ctx.clearRect(0,0,c.width,c.height);ctx.fillStyle='#0a1d2a';ctx.fillRect(0,0,c.width,c.height);ctx.fillStyle='#d4e4ea';ctx.font='13px sans-serif';ctx.fillText(xLabel+' / '+yLabel,15,20);const good=points.filter(p=>p.slice(0,2).every(Number.isFinite));if(!good.length)return;const xs=good.map(p=>p[0]),ys=good.map(p=>p[1]),a=Math.min(...xs),b=Math.max(...xs),lo=Math.min(...ys),hi=Math.max(...ys);for(const [x,y,front]of good){ctx.beginPath();ctx.fillStyle=front?'#ffc968':'#6bc9e0';ctx.arc(65+(x-a)/(b-a||1)*(c.width-100),c.height-45-(y-lo)/(hi-lo||1)*(c.height-95),front?6:4,0,7);ctx.fill();}const groups=[...new Set(good.map(p=>p[3]).filter(v=>v!==undefined))];groups.forEach((g,i)=>{const series=good.filter(p=>p[3]===g).sort((a,b)=>a[0]-b[0]);ctx.strokeStyle=['#6bc9e0','#ffc968','#c5a0ff','#88dc9c'][i%4];ctx.beginPath();series.forEach(([x,y],j)=>{const px=65+(x-a)/(b-a||1)*(c.width-100),py=c.height-45-(y-lo)/(hi-lo||1)*(c.height-95);if(j)ctx.lineTo(px,py);else ctx.moveTo(px,py);});ctx.stroke();ctx.fillStyle=ctx.strokeStyle;ctx.fillText('θ='+g+'°',c.width-200,40+i*18);});ctx.fillStyle='#d4e4ea';ctx.fillText(fmt(a),40,c.height-15);ctx.fillText(fmt(b),c.width-100,c.height-15);ctx.fillText(fmt(hi),5,45);ctx.fillText(fmt(lo),5,c.height-50);}
function mount(){
 const parent=$('research-fiber-v1');if(!parent)return;
 const section=document.createElement('section');section.id='research-step6-v088';section.style.cssText='margin-top:24px;border-top:2px solid #487287;padding-top:18px;min-width:0;overflow-wrap:anywhere';
 section.innerHTML='<h3>STEP 6 · 확장 해석 v0.8.8</h3><p>Python 혼합 벡터 FEM 실행 → JSON 가져오기. 손실·PM·MCF·접속 정확도는 독립 검증 전까지 후보 값입니다.</p>'+
 '<div class="rf-line">'+kinds.map((k,i)=>'<button id="r688-tab-'+k+'" data-step6="'+k+'">'+names[i]+'</button>').join('')+'</div>'+
 '<p id="r688-help"></p><div id="r688-settings" class="rf-line"><label>굽힘 반경 mm<input id="r688-radii" value="100,20" style="max-width:150px"></label><label>방향 deg<input id="r688-directions" value="0,90" style="max-width:140px"></label><label>모드 수<input id="r688-modes" type="number" value="6" min="1" max="64" style="max-width:75px"></label></div>'+
 '<label for="r688-extra">추가 입력 JSON (PM 재료 / 접속 필드 / 스윕 축·목적 / DGD 데이터)</label><textarea id="r688-extra" rows="6" spellcheck="false" style="width:100%;box-sizing:border-box;font:12px monospace"></textarea>'+
 '<div class="rf-line"><button id="r688-export-config">실행 입력 JSON 저장</button><button id="r688-export-template">입력 템플릿</button><button id="r688-example">실제 계산 예시 적용</button><label for="r688-open" class="rf-file">결과 JSON 가져오기</label><input type="file" id="r688-open" accept=".json" hidden><button id="r688-save-result" disabled>현재 결과 저장</button></div>'+
 '<p><code id="r688-command"></code></p><p id="r688-status" role="status">계산 전 · FEM은 로컬/Colab에서 실행합니다.</p><p id="r688-summary"></p><div style="overflow:auto;max-width:100%"><table id="r688-table" style="width:100%;font-size:12px"></table></div><canvas id="r688-chart" width="800" height="270" style="width:100%;height:auto"></canvas>'+
 '<details><summary>공식·단위·검증 한계</summary><p>n_eq=n(1+x_out/R_eff), R[mm]. α=(20/ln10)k0 Im(neff)[dB/m]. B=|nx−ny|, L_B=λ/B[m]. κ=(β_even−β_odd)/2[1/m]. η=|∫Et1·Et2*|²/(∫|Et1|²∫|Et2|²). Δng=Δn−λ dΔn/dλ, DGD=L|Δng|/c.</p><p>R_eff/R 경험 계수, PM 상수, Fresnel 계면 굴절률의 기본 추정값은 없습니다. PM 합성 예시는 실제 PANDA 재료 상수가 아닙니다. Et 접속 근사는 SMF↔HCF 융착 변형·표면 산란을 포함하지 않습니다. PMD 통계는 시연용이며 B만으로 PMD 계수[ps/√km]를 결정하지 않습니다.</p><a href="../fiber_research_v1/STEP6_V088.md">실행·출처·한계 문서</a></details>';
 parent.appendChild(section);
 let kind='bend',output=null;
 const status=s=>$('r688-status').textContent=s;
 const clear=()=>{output=null;$('r688-table').innerHTML='';$('r688-summary').textContent='';$('r688-save-result').disabled=true;plot([],'','');status('입력이 변경되어 결과 무효 · 다시 계산/가져오기');};
 const help={bend:'반경·방향별 PML 고유해를 실제 계산하고 퇴화 부분공간을 overlap으로 추적합니다. 평균 총손실 차이를 순수 굽힘 손실로 인증하지 않습니다.',pm:'E, ν, α, ΔT, p11/p12 및 각각의 출처가 필요합니다. 열응력→굴절률 텐서→같은 공간 분포의 두 편광 후보를 비교합니다.',mcf:'N코어 실 FEM supermode. 1~64개 요청 가능하며 DOF/메모리 한도가 적용됩니다. κ는 동일 대칭 2코어의 같은 편광 even/odd 후보만 산출합니다.',splice:'실제 FEM field-grid 2개 또는 명시적 Gaussian 검증 입력을 사용합니다. 오프셋[µm], 각도[deg], 갭[µm]와 선택적 scalar Fresnel을 적용합니다.',sweep:'다차원 그리드의 각 점에서 실제 MCF/굽힘 FEM을 실행합니다. 실패 점은 최적점에서 제외하고 Pareto 및 목적별 최적 설계를 저장합니다.',dgd:'현재 PMD 통계는 시연용입니다. 부호가 유지된 nx−ny를 파장별 추적한 데이터로 균일 무결합 섬유 DGD를 계산합니다. 랜덤 PMD 모델과 연결하려면 길이방향 상관·결합 자료가 필요합니다.'};
 function choose(k){kind=k;clear();$('r688-help').textContent=help[k];$('r688-command').textContent='python step6_cli.py '+k+' step6_'+k+'_input.json --out step6_'+k+'_result.json';$('r688-example').disabled=k==='dgd';$('r688-settings').hidden=!['bend','pm','mcf'].includes(k);for(const btn of section.querySelectorAll('[data-step6]'))btn.setAttribute('aria-pressed',String(btn.dataset.step6===k));}
 async function template(){
  if(kind==='pm'){const r=await fetch('../fiber_research_v1/benchmarks/step6b_panda_materials_TEMPLATE.json');if(!r.ok)throw Error('템플릿 HTTP '+r.status);return {materials:await r.json()};}
  if(kind==='bend')return {settings:{reff_multiplier:null}};
  if(kind==='mcf')return {settings:{target_neff:null}};
  if(kind==='splice')return {field1:{TODO:'step6_cli.py field-grid 로 생성한 JSON 전체'},field2:{TODO:'같은 파장 field-grid JSON 전체'},settings:{offset_um:[0,0],angle_deg:[0,0],gap_um:0,gap_n:1,interface_indices:null}};
  if(kind==='sweep')return {axes:{'parameters.pitch_um':[6,7],'parameters.trench_width_um':[.3,.5]},objectives:{mean_core_power_fraction:{direction:'max',unit:'1'},max_total_loss_dB_m:{direction:'min',unit:'dB/m'}},evaluator:{kind:'mcf',polarization:'auto',settings:{n_modes:6}},constraints:{}};
  return {wavelength_nm:[],signed_delta_n:[],length_m:null,tracking_evidence:{polarization_branches_tracked:false,source:'TODO 실제 편광 branch overlap 추적 결과'}};
 }
 function render(r){validateStep6Result(r);const d=r.source_design||r.configuration?.design;if(d&&!equal(d,JSON.parse($('rf-json').value)))throw Error('현재 설계와 결과 입력이 다릅니다. 해당 설계를 먼저 불러오세요.');clear();output=r;$('r688-save-result').disabled=false;
 const type=r.schema.split('_')[2];if(r.settings?.n_modes)$('r688-modes').value=r.settings.n_modes;if(r.settings?.radii_mm)$('r688-radii').value=r.settings.radii_mm.join(',');if(r.settings?.directions_deg)$('r688-directions').value=r.settings.directions_deg.join(',');$('r688-summary').textContent='미검증 후보 · '+(r.example_basis||r.validation?.material_basis||r.validation?.limits?.join(' · ')||r.limitation||r.status||'');
 if(type==='bend'){table(['R mm','방향 deg','총손실 dB/m','직선 대비 차 dB/m','상태'],r.cases.map(c=>[c.radius_mm,c.direction_deg,c.mean_alpha_total_dB_m,c.delta_mean_total_dB_m,c.status]));plot(r.cases.map(c=>[c.radius_mm,c.mean_alpha_total_dB_m,false,c.direction_deg]),'반경 mm','평균 총손실 dB/m');}
 if(type==='pm'){const b=r.modal_birefringence;table(['B','L_B m','공간 overlap','상태'],[[b.B,b.L_B_m,b.spatial_intensity_overlap,b.status]]);}
 if(type==='mcf'){table(['모드','Re neff','Im neff','편광','core power 합','envelope parity'],r.modes.map(m=>[m.index,m.neff_real,m.neff_imag,m.polarization,m.core_power_fractions.reduce((a,b)=>a+b,0),m.envelope_parity]));$('r688-summary').textContent+=' · κ 후보: '+JSON.stringify(r.coupling_pairs);plot(r.modes.map(m=>[m.index,m.neff_real]),'모드','Re neff');}
 if(type==='splice')table(['η Et','overlap 손실 dB','Fresnel T','결합 손실 dB'],[[r.eta_Et,r.overlap_loss_dB,r.fresnel_T,r.combined_loss_dB]]);
 if(type==='grid'){const keys=Object.keys(r.configuration.objectives);table(['case','파라미터',...keys,'Pareto','상태'],r.cases.map(c=>[c.case_index,JSON.stringify(c.parameters),...keys.map(k=>c.metrics?.[k]),r.pareto_case_indices.includes(c.case_index)?'●':'',c.status]));plot(r.cases.map(c=>[c.metrics?.[keys[0]],c.metrics?.[keys[1]],r.pareto_case_indices.includes(c.case_index)]),keys[0],keys[1]||'');}
 if(type==='dgd'){table(['λ nm','signed Δn','Δng','DGD ps'],r.wavelength_nm.map((l,i)=>[l,r.signed_delta_n[i],r.delta_ng[i],r.DGD_ps[i]]));plot(r.wavelength_nm.map((l,i)=>[l,r.DGD_ps[i]]),'파장 nm','DGD ps');}
 status('Python 결과 적용 · '+r.schema+' · 정량 인증 아님');
 }
 section.querySelectorAll('[data-step6]').forEach(b=>b.addEventListener('click',()=>choose(b.dataset.step6)));
 $('r688-export-template').onclick=async()=>{try{$('r688-extra').value=JSON.stringify(await template(),null,2);clear();status('템플릿 적용 · TODO를 실제 출처 값으로 입력하세요.');}catch(e){status(e.message);}};
 $('r688-export-config').onclick=()=>{try{const d=JSON.parse($('rf-json').value),extra=JSON.parse($('r688-extra').value||'{}'),s=extra.settings||{},m=Number($('r688-modes').value);if(!Number.isInteger(m)||m<1||m>64)throw Error('모드 수 1~64');let c=extra;
  if(['bend','pm','mcf'].includes(kind)){c={...extra,design:d,settings:{...s,n_modes:m}};if(kind==='bend'){if(m>8)throw Error('굽힘 추적은 최대 8모드');c.settings={...c.settings,radii_mm:numbers('r688-radii'),directions_deg:numbers('r688-directions')};}if(kind==='pm'&&!extra.materials)throw Error('PM 재료 템플릿과 출처를 입력하세요.');if(kind==='mcf'&&d.family!=='mcf')throw Error('MCF 설계를 선택하세요.');}
  if(kind==='sweep')c={...extra,design:d};download({schema:'LSCNS_step6_input_v088',schema_version:'1.0.0',operation:kind,...c},'step6_'+kind+'_input.json');status('입력 JSON 저장 · 표시된 Python 명령으로 계산하세요.');}catch(e){status('입력 오류: '+e.message);}};
 $('r688-example').onclick=async()=>{const selected=kind;try{status('실제 계산 결과 불러오는 중…');const response=await fetch(ROOT+examples[kinds.indexOf(selected)]);if(!response.ok)throw Error('HTTP '+response.status);const r=await response.json();if(kind!==selected)return;const d=r.source_design||r.configuration?.design;if(d)document.dispatchEvent(new CustomEvent('ls-step6-design-import',{detail:d}));render(r);}catch(e){clear();status('예시 오류: '+e.message);}};
 $('r688-open').onchange=async e=>{try{const f=e.target.files[0];if(!f)return;if(f.size>20000000)throw Error('20 MB 제한');render(JSON.parse(await f.text()));}catch(err){clear();status('가져오기 실패: '+err.message);}};
 $('r688-save-result').onclick=()=>{if(output)download(output,'step6_result.json');};
 $('rf-json').addEventListener('input',()=>{if(output)clear();});$('r688-extra').addEventListener('input',()=>{if(output)clear();});$('r688-settings').addEventListener('input',()=>{if(output)clear();});choose('bend');
}
if(typeof document!=='undefined'){if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',mount);else mount();}
