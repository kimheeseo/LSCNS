/** STEP6b v0.8.2: source-guarded PM material JSON editor and result-only stress-optic maps.
 * Does not run thermoelastic/anisotropic vector FEM in the browser.
 * Distinguishes spatial material tensor birefringence from actual modal birefringence.
 */
const $=id=>document.getElementById(id);
const esc=x=>String(x??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;':'&quot;',"'":'&#39;'}[c]));
const fmt=(v,d=5)=>typeof v==='number'&&Number.isFinite(v)?v.toPrecision(d):'근거 필요';
const sorted=o=>Array.isArray(o)?o.map(sorted):o&&typeof o==='object'?
  Object.fromEntries(Object.keys(o).sort().map(k=>[k,sorted(o[k])])):o;
const same=(a,b)=>JSON.stringify(sorted(a))===JSON.stringify(sorted(b));
let summary=null,stress=null,pmConfig=null;
function status(t){const e=$('r6b-status');if(e)e.textContent=t;}
function getDesign(){
  const editor=$('rf-json');if(!editor)throw Error('STEP 1 설계 JSON 패널이 없습니다');
  return JSON.parse(editor.value);
}
function getMaterials(){return JSON.parse($('r6b-materials').value);}
function designMatches(){
  if(!summary)return false;
  try{return same(getDesign(),summary.source_design)&&same(getMaterials(),summary.material_config);}
  catch{return false;}
}
function validateMaterial(m,partial=false){
  if(m?.schema!=='LSCNS_step6b_materials_v1'||typeof m.materials!=='object'||
     m.materials===null)throw Error('STEP6b materials schema가 아닙니다');
  if(!partial&&(!Number.isFinite(m.delta_T_K)||Math.abs(m.delta_T_K)>2000))
    throw Error('냉각 ΔT[K] 수치/범위를 확인하세요');
  const keys=(getDesign()?.variant==='elliptic_core')?['background','core']:
    ['background','core','stress_rod'];
  for(const name of keys){
    const p=m.materials[name];
    if(!p)throw Error(name+': 재료값 누락');
    if(partial)continue;
    for(const k of ['E_Pa','poisson_ratio','alpha_1_per_K']){
      if(!Number.isFinite(p[k]))throw Error(name+' '+k+': 수치 및 문헌 근거 필요');
    }
    if(p.E_Pa<=0||p.poisson_ratio<=-.99||p.poisson_ratio>=.4999)
      throw Error(name+': E/ν 유효 범위 위반');
    if(!/^https?:\/\//.test(p.mechanical_source_url||''))
      throw Error(name+': mechanical_source_url 필요');
    const missing=p.p11==null||p.p12==null;
    if(!missing&&(!Number.isFinite(p.p11)||!Number.isFinite(p.p12)||
       !/^https?:\/\//.test(p.photoelastic_source_url||'')))
      throw Error(name+': 유효한 strain-optic p11,p12/source_url 필요');
  }
  return m;
}
function downloadJSON(text,name){
  const u=URL.createObjectURL(new Blob([text],{type:'application/json'}));
  const a=document.createElement('a');a.href=u;a.download=name;a.click();
  setTimeout(()=>URL.revokeObjectURL(u),3000);
}
function showMath(){
  const e=$('r6b-math');if(!e)return;
  let config;
  try{config=getMaterials();}catch{e.textContent='재료 JSON 오류: 값을 확인하세요';return;}
  const para=[];
  para.push('일반화 평면변형률: σ=2μ(ε−αΔT·I)+λ[tr(ε)−3αΔT]·I, λ=Eν/[(1+ν)(1−2ν)], μ=E/[2(1+ν)], ∫σzz dA=0.');
  para.push('광탄성: ΔQii=p11 εii+p12(εjj+εkk), ΔQxy=(p11−p12)εxy, Q=εr⁻¹=n₀⁻²I+ΔQ, εr=Q⁻¹.');
  para.push('코어 재료 복굴절 후보: B_material=|n_principal_1−n_principal_2| (core Q 면적평균). 실제 모드 복굴절 B_modal=|Re(neff,x)−Re(neff,y)|은 FEM 두 모드 검증을 별도로 요구.');
  para.push('냉각 ΔT='+fmt(config.delta_T_K)+' K; 부호는 T_operating−T_reference. 냉각이면 음수.');
  for(const [key,v] of Object.entries(config.materials||{})){
    para.push('재료 '+key+': E='+fmt(v.E_Pa)+' Pa; ν='+fmt(v.poisson_ratio)+
              '; α='+fmt(v.alpha_1_per_K)+' 1/K; p11='+fmt(v.p11)+
              '; p12='+fmt(v.p12)+
              '; 출처(기계) '+(v.mechanical_source_url||'TODO')+
              '; 출처(광탄성) '+(v.photoelastic_source_url||'TODO'));
  }
  e.textContent=para.join('\n\n');
}
function draw(){
  const canvas=$('r6b-map');if(!canvas)return;
  const ctx=canvas.getContext('2d');const W=canvas.width,H=canvas.height;
  ctx.fillStyle='#071a2b';ctx.fillRect(0,0,W,H);
  const samples=stress?.field_samples;
  if(!summary||!designMatches()||!Array.isArray(samples)||!samples.length){
    ctx.fillStyle='#d3e8f1';ctx.font='14px sans-serif';
    ctx.fillText('PM stress result JSON을 가져오면 응력/광탄성 지도 표시',24,52);
    return;
  }
  const metric=$('r6b-select').value;
  const pts=samples.filter(p=>Array.isArray(p.point_um)&&p.point_um.length===2);
  const getValue=p=>metric==='diff'?p.stress_Pa?.[0]-p.stress_Pa?.[1]:
    metric==='shear'?p.stress_Pa?.[3]:
    Array.isArray(p.local_principal_n)?p.local_principal_n[0]-p.local_principal_n[1]:NaN;
  const valid=pts.filter(p=>Number.isFinite(getValue(p)));
  if(!valid.length){ctx.fillStyle='#f2cf9c';ctx.fillText('해당 물리량 없음: 광탄성 p11,p12 미입력 또는 미검증',24,52);return;}
  const a=valid.map(p=>p.point_um[0]),b=valid.map(p=>p.point_um[1]);
  let xmin=Math.min(...a),xmax=Math.max(...a),ymin=Math.min(...b),ymax=Math.max(...b);
  if(xmax===xmin){xmin-=1;xmax+=1;}if(ymax===ymin){ymin-=1;ymax+=1;}
  const pad=Math.max(xmax-xmin,ymax-ymin)*.04||1;
  xmin-=pad;xmax+=pad;ymin-=pad;ymax+=pad;
  const range=valid.map(getValue);let lo=Math.min(...range),hi=Math.max(...range);
  if(lo===hi){lo-=1e-15;hi+=1e-15;}
  const left=56,top=22,w=W-90,h=H-66;
  const side=Math.min(w,h);
  const xpx=x=>left+(x-xmin)/(xmax-xmin)*w;
  const ypx=y=>top+(ymax-y)/(ymax-ymin)*h;
  const radius=Math.max(2,Math.min(6,side/Math.sqrt(valid.length)/2));
  for(const p of valid){
    const val=(getValue(p)-lo)/(hi-lo);
    const hue=220-220*val;
    ctx.fillStyle='hsl('+hue.toFixed(1)+' 92% 58%)';
    ctx.beginPath();ctx.arc(xpx(p.point_um[0]),ypx(p.point_um[1]),radius,0,2*Math.PI);ctx.fill();
  }
  ctx.strokeStyle='#7092a4';ctx.strokeRect(left,top,w,h);
  ctx.font='12px sans-serif';ctx.fillStyle='#d7e6f3';
  ctx.fillText('x [µm]',left+w/2,top+h+37);
  ctx.fillText('y [µm]',5,top+24);
  ctx.fillText('min '+fmt(lo)+'   max '+fmt(hi)+' '+(metric==='b'?'[dimensionless]':'[Pa]'),left+10,top+18);
  ctx.fillText('※ element centroid scatter, not FEM continuous field',left+10,top+h+17);
}
function update(){
  showMath();
  if(!summary){draw();return;}
  if(!designMatches()){
    status('원본 설계 또는 재료 입력 변경됨 → 기존 계산 결과 무효화. Python에서 다시 계산해야 합니다.');
    $('r6b-metrics').textContent='계산 결과 무효 / 입력 불일치';
    draw();return;
  }
  const proxy=summary.photoelastic?.core_material_birefringence||{};
  const modal=summary.modal_birefringence||{};
  const mech=summary.mechanics||{};
  const box=[
    ['최대 응력 |σ| [Pa]',fmt(mech.max_abs_stress_Pa)],
    ['축방향 일반화 변형률 εzz',fmt(mech.axial_strain_uniform)],
    ['코어 재료 B (NOT modal)',fmt(proxy.B_core_area_weighted)],
    ['코어 재료 beat-length proxy [m]',fmt(proxy.beat_length_proxy_m)],
    ['모드 B candidate (NOT certified)',fmt(modal.B_modal_candidate)],
    ['모드 beat-length candidate [m]',fmt(modal.beat_length_candidate_m)],
    ['검증',summary.validation?.COMSOL_comparison||'EVIDENCE_REQUIRED'],
    ['버전/검증',summary.engine_version+' / '+(modal.status||'TODO')]
  ];
  $('r6b-metrics').innerHTML=box.map(([k,v])=>
    '<div style="border:1px solid #3b6173;padding:9px;border-radius:7px;background:#0b2236">'+
    '<small style="display:block;color:#bdcbd7">'+esc(k)+'</small><strong>'+esc(v)+'</strong></div>').join('');
  status('입력 일치 · 기계 FEM / 재료 tensor 후보. 최종 PM 모드 B는 COMSOL·PML·메시 검증 전까지 미인증.');
  $('r6b-hash').textContent='입력 SHA-256: '+summary.input_sha256+
    ' | UTC: '+(summary.created_utc||'unknown');
  draw();
}
function mount(){
  const anchor=$('research-step6a-v081')||$('research-fiber-v1')||document.querySelector('main');
  if(!anchor||$('research-step6b-v082'))return;
  const s=document.createElement('section');
  s.id='research-step6b-v082';s.className='engineering-panel';
  s.style.cssText='padding:18px;margin:22px 0;background:#10283c;color:#e5f4ff;border:1px solid #3d6985;border-radius:11px';
  s.innerHTML='<h2 style="font-size:18px;margin:0 0 7px">STEP 6b · PM 섬유 열응력·광탄성 <small>v0.8.2</small></h2>'+
    '<p style="font-size:12px;color:#e9cca4">Panda / Bow-tie 별도 응력봉 영역 · 일반화 평면변형률 열응력 · 등방성 strain-optic tensor → 복굴절 후보. FEM/COMSOL 정량 검증 미완료.</p>'+
    '<p style="font-size:12px">모든 E, ν, α, p11, p12 값과 출처는 별도 입력 JSON으로 관리합니다. 확인되지 않은 재료 상수는 TODO를 유지하고 광학 계산을 실행하지 않습니다.</p>'+
    '<div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin:12px 0">'+
    '<label style="border:1px solid #668aa1;border-radius:7px;padding:8px;cursor:pointer">재료 JSON 가져오기<input id="r6b-mat-open" type="file" accept=".json" hidden></label>'+
    '<button id="r6b-mat-save" style="background:#20425a;border:1px solid #668aa1;color:#f0f8ff;border-radius:7px;padding:8px;cursor:pointer">재료 JSON 저장</button>'+
    '<label style="border:1px solid #668aa1;border-radius:7px;padding:8px;cursor:pointer">PM 결과 요약 JSON 가져오기<input id="r6b-summary-open" type="file" accept=".json" hidden></label>'+
    '<label style="border:1px solid #668aa1;border-radius:7px;padding:8px;cursor:pointer">응력 지도 JSON 가져오기<input id="r6b-stress-open" type="file" accept=".json" hidden></label></div>'+
    '<textarea id="r6b-materials" aria-label="PM 소재 물성 입력 JSON" spellcheck="false" style="width:100%;min-height:160px;max-height:350px;background:#061b2b;color:#f2f9ff;padding:10px;border:1px solid #52778d;border-radius:7px;font:12px/1.4 Consolas,monospace"></textarea>'+
    '<p id="r6b-status" role="status" style="font-size:12px">문헌 근거와 재료값을 입력하세요.</p>'+
    '<div id="r6b-metrics" style="display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:8px;margin:10px 0"></div>'+
    '<label style="font-size:12px">물리 지도 <select id="r6b-select" style="background:#183c56;color:#edf6fd;border:1px solid #537e98;padding:7px">'+
    '<option value="diff">σxx − σyy [Pa]</option><option value="shear">σxy [Pa]</option>'+
    '<option value="b">지역 재료 복굴절 (모드가 아님)</option></select></label>'+
    '<canvas id="r6b-map" width="680" height="310" style="width:100%;max-width:820px;height:auto;display:block"></canvas>'+
    '<p id="r6b-hash" style="font-size:11px;overflow-wrap:anywhere;color:#b4c8d9"></p>'+
    '<details><summary style="cursor:pointer">계산 근거 보기 (공식, 적용값 및 문헌 출처)</summary>'+
    '<pre id="r6b-math" style="font-size:11px;white-space:pre-wrap;overflow-wrap:anywhere;background:#081b2c;padding:10px"></pre>'+
    '<p style="font-size:11px"><a target="_blank" rel="noopener" href="https://doc.comsol.com/6.4/doc/com.comsol.help.models.sme.stress_optical_generalized/stress_optical_generalized.html">COMSOL Generalized Plane Strain</a> · '+
    '<a target="_blank" rel="noopener" href="https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.stress_optical/stress_optical.html">COMSOL Stress-Optical Effects</a></p></details>'+
    '<p style="font-size:11px;color:#ead5aa">해당 웹은 읽기 전용 시각화/재료 JSON 편집 도구입니다. Gmsh·FEM은 Python/Colab에서 실행합니다. 재료 평균 B는 실제 모드 유효 굴절률 간 차이가 아닙니다.</p>';
  anchor.after(s);
  $('r6b-mat-open').addEventListener('change',async e=>{
    try{
      const f=e.target.files?.[0];if(!f||f.size>1000000)throw Error('파일 크기/미선택');
      const cfg=validateMaterial(JSON.parse(await f.text()),true);
      $('r6b-materials').value=JSON.stringify(cfg,null,2);update();
    }catch(err){status('재료 JSON 오류: '+err.message);}
  });
  $('r6b-mat-save').addEventListener('click',()=>{
    try{
      const c=validateMaterial(getMaterials(),false);
      downloadJSON(JSON.stringify(c,null,2)+'\n','step6b_pm_materials.json');
      status('재료 JSON 저장. 광학 모드 검증과 실제 문헌값 적합성은 별도 필요.');
    }catch(err){status('재료 JSON 저장 실패: '+err.message);}
  });
  $('r6b-summary-open').addEventListener('change',async e=>{
    try{
      const f=e.target.files?.[0];if(!f||f.size>4000000)throw Error('4MB 제한');
      const v=JSON.parse(await f.text());
      if(v?.schema!=='LSCNS_step6b_summary_v082'||v.engine_version!=='0.8.2'||
         !v.source_design||!v.material_config||!/^[a-f0-9]{64}$/i.test(v.input_sha256||'')){
        throw Error('STEP6b v0.8.2 결과가 아닙니다');
      }
      summary=v;update();
    }catch(err){status('결과 요약 오류: '+err.message);}
  });
  $('r6b-stress-open').addEventListener('change',async e=>{
    try{
      const f=e.target.files?.[0];if(!f||f.size>14000000)throw Error('14MB 제한');
      const v=JSON.parse(await f.text());
      if(v?.schema!=='LSCNS_step6b_stress_result_v082'||!Array.isArray(v.field_samples))
        throw Error('STEP6b stress JSON 아님');
      if(summary&&v.input_sha256!==summary.input_sha256)throw Error('요약 결과와 응력 지도 입력 해시 불일치');
      stress=v;update();
    }catch(err){status('응력 지도 오류: '+err.message);}
  });
  $('r6b-materials').addEventListener('input',update);
  $('rf-json')?.addEventListener('input',update);
  $('r6b-select').addEventListener('change',draw);
  fetch('../fiber_research_v1/benchmarks/step6b_panda_materials_TEMPLATE.json',{cache:'no-store'})
    .then(r=>{if(!r.ok)throw Error('HTTP '+r.status);return r.json();})
    .then(c=>{$('r6b-materials').value=JSON.stringify(c,null,2);update();})
    .catch(e=>status('재료 TODO 템플릿 가져오기 실패: '+e.message));
  draw();
}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',mount);
else mount();
