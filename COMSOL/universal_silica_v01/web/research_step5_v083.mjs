/* LSCNS STEP5 v0.8.3. Browser visualizes Python scalar eigenvalue JSON;
 * it cannot solve the Bloch eigenproblem. No fictitious sample bands.
 */
const R5='../fiber_research_v1/';
const id5=k=>document.getElementById(k);
const keys5=['lattice','pitch_um','hole_diameter_over_pitch','n_host','n_air',
             'grid_n','bands','kmesh_n','path_points_per_segment','beta_norms','backend','pwe_cutoff','schema_version'];
const sites5={triangular:[[0,0]],hexagonal:[[1/3,1/3],[2/3,2/3]],
              square:[[0,0]],kagome:[[.5,0],[0,.5],[.5,.5]]};
const desc5={
  triangular:'삼각 격자(1-site triangular air-hole lattice)',
  hexagonal:'육각 벌집(honeycomb): 2-site basis; 삼각 격자와 구분',
  square:'사각 격자(1-site square air-hole lattice)',
  kagome:'이상화 Kagome 3-site 원형 구멍; 실제 strut Kagome HCF가 아님'
};
let result5=null;
const fixed5=(x,d=3)=>Number.isFinite(Number(x))?Number(x).toFixed(d):'—';
function read5(){
  const vector=id5('pc5-backend').value==='vector_pwe';
  return {...(vector?{schema_version:'1.0.0',backend:'vector_pwe',pwe_cutoff:Number(id5('pc5-cutoff').value)}:{}),schema:vector?'LSCNS_STEP5_BLOCH_INPUT_V087':'LSCNS_STEP5_BLOCH_INPUT_V083',
    lattice:id5('pc5-lattice').value,
    pitch_um:Number(id5('pc5-pitch').value),
    hole_diameter_over_pitch:Number(id5('pc5-hole').value),
    n_host:Number(id5('pc5-host').value),n_air:Number(id5('pc5-air').value),
    grid_n:Number(id5('pc5-grid').value),bands:Number(id5('pc5-bands').value),
    kmesh_n:Number(id5('pc5-kmesh').value),
    path_points_per_segment:Number(id5('pc5-path').value),
    beta_norms:id5('pc5-betas').value.split(',').map(x=>Number(x.trim())),
    input_basis:'Explicit constant refractive indices; non-dispersive STEP5; see backend'};
}
function guard5(c){
  if(c.backend==='vector_pwe'&&(!Number.isInteger(c.pwe_cutoff)||c.pwe_cutoff<1||c.pwe_cutoff>8))return 'PWE cutoff 1..8';
  if(!sites5[c.lattice])return '격자 이름 오류';
  if(!(c.pitch_um>=.1&&c.pitch_um<=1000))return 'Λ는 0.1–1000 µm';
  if(!(c.hole_diameter_over_pitch>.01&&c.hole_diameter_over_pitch<.95))return 'd/Λ는 (0.01, 0.95)';
  if(!([c.n_host,c.n_air].every(x=>x>=.9&&x<=5)))return 'n 범위: 0.9–5';
  const mind=c.lattice==='hexagonal'?1/Math.sqrt(3):c.lattice==='kagome'?.5:1;
  if(c.hole_diameter_over_pitch>=.98*mind)return '구멍 지름이 격자 최근접 거리에 비해 큽니다(중첩).';
  for(const [key,lo,hi] of [['grid_n',10,96],['bands',2,16],['kmesh_n',2,12],
                             ['path_points_per_segment',2,30]]){
    if(!Number.isInteger(c[key])||c[key]<lo||c[key]>hi)return key+' 허용: '+lo+'–'+hi;
  }
  if(c.bands>=c.grid_n**2-1)return '모드 수가 행렬 차원보다 큽니다.';
  if(!Array.isArray(c.beta_norms)||c.beta_norms.length<1||c.beta_norms.length>10||
     c.beta_norms.some(x=>!Number.isFinite(x)||x<(c.backend==='vector_pwe'?0:.001)||x>5))return 'βΛ/2π: 0.001–5, 쉼표로 구분한 1–10개';
  if(c.beta_norms.length*c.kmesh_n**2+3*c.path_points_per_segment>260)return '고유해석 260회 초과: k-mesh 또는 β값을 줄여주세요.';
  return '';
}
function validResult5(d){
  if(!(['LSCNS_STEP5_BLOCH_RESULT_V083','LSCNS_STEP5_BLOCH_RESULT_V087'].includes(d?.schema))||!['SCALAR_SAMPLED_UNVERIFIED','VECTOR_PWE_SAMPLED_UNVERIFIED'].includes(d.status)||
     !d.input||!Array.isArray(d.path?.distance_norm)||!Array.isArray(d.path?.bands_f)||
     !Array.isArray(d.gap_map?.candidates)||!Array.isArray(d.gap_map?.light_lines_beta_over_k0))
    throw Error('v0.8.3 Bloch 계산 결과 JSON 스키마가 아닙니다.');
  const error=guard5(d.input);if(error)throw Error(error);
  if(d.path.distance_norm.length!==d.path.bands_f.length||
     d.path.bands_f.some(b=>!Array.isArray(b)||b.length!==d.input.bands||b.some(x=>!Number.isFinite(x)||x<(d.input.backend==='vector_pwe'?0:Number.MIN_VALUE))))
    throw Error('밴드 행렬 구조 오류');
  for(const g of d.gap_map.candidates){
    if(!Number.isFinite(g.f_low)||!Number.isFinite(g.f_high)||g.f_low<=0||g.f_high<=g.f_low||
       !Number.isFinite(g.beta_norm)||g.status!=='SAMPLED_GAP_CANDIDATE')
      throw Error('후보 간극 데이터가 잘못되었습니다');
  }
}
function same5(a,b){return keys5.every(k=>JSON.stringify(a?.[k])===JSON.stringify(b?.[k]));}
function status5(text,error=false){
  const el=id5('pc5-status');if(el){el.textContent=text;el.style.color=error?'#ffb4a6':'#a6e2e1';}
}
function download5(obj,name){
  const blob=new Blob([JSON.stringify(obj,null,2)+'\n'],{type:'application/json'});
  const uri=URL.createObjectURL(blob),a=document.createElement('a');
  a.href=uri;a.download=name;document.body.appendChild(a);a.click();a.remove();
  setTimeout(()=>URL.revokeObjectURL(uri),800);
}
function switch5(k){
  for(const t of ['design','bands','defect']){
    id5('pc5-pane-'+t).hidden=t!==k;
    id5('pc5-tab-'+t).setAttribute('aria-pressed',String(t===k));
  }
}
function a5(cfg){
  const p=cfg.pitch_um;
  return cfg.lattice==='square'?[[p,0],[0,p]]:[[p,p/2],[0,Math.sqrt(3)*p/2]];
}
function drawPreview5(){
  const canvas=id5('pc5-preview'),ctx=canvas.getContext('2d'),W=canvas.width,H=canvas.height;
  ctx.clearRect(0,0,W,H);ctx.fillStyle='#0b1e33';ctx.fillRect(0,0,W,H);
  const c=read5(),error=guard5(c);
  id5('pc5-unit-description').textContent=desc5[c.lattice]+(error?' | '+error:'');
  if(error)return;
  const a=a5(c),v=(u,z)=>[a[0][0]*u+a[0][1]*z,a[1][0]*u+a[1][1]*z];
  const vertices=[[0,0],[1,0],[1,1],[0,1]],points=vertices.map(x=>v(...x));
  const xs=points.map(p=>p[0]),ys=points.map(p=>p[1]);
  const mid=[(Math.min(...xs)+Math.max(...xs))/2,(Math.min(...ys)+Math.max(...ys))/2];
  const size=Math.max(Math.max(...xs)-Math.min(...xs),Math.max(...ys)-Math.min(...ys))*1.4;
  const s=Math.min(W,H)/size;
  const tr=(u,z)=>{const p=v(u,z);return [W/2+(p[0]-mid[0])*s,H/2-(p[1]-mid[1])*s];};
  ctx.save();ctx.beginPath();
  vertices.forEach(([u,z],i)=>{if(i===0)ctx.moveTo(...tr(u,z));else ctx.lineTo(...tr(u,z));});
  ctx.closePath();ctx.fillStyle='#c2cad7';ctx.fill();ctx.clip();
  for(const [u,z] of sites5[c.lattice]){
    for(let dx=-1;dx<=1;dx++)for(let dy=-1;dy<=1;dy++){
      const p=tr(u+dx,z+dy);
      ctx.beginPath();ctx.arc(p[0],p[1],c.hole_diameter_over_pitch*c.pitch_um*s/2,0,Math.PI*2);
      ctx.fillStyle='#154366';ctx.fill();
    }
  }
  ctx.restore();ctx.beginPath();
  vertices.forEach(([u,z],i)=>{if(!i)ctx.moveTo(...tr(u,z));else ctx.lineTo(...tr(u,z));});
  ctx.closePath();ctx.lineWidth=2;ctx.strokeStyle='#52dcd9';ctx.stroke();
  ctx.font='14px sans-serif';ctx.fillStyle='#dbf3ff';
  ctx.fillText('단위셀 · Periodic air-hole motif',18,25);
  ctx.font='12px sans-serif';ctx.fillStyle='#9dc8d9';
  ctx.fillText('E(r+aᵢ)=exp(i kᴮ·aᵢ) E(r)',18,H-16);
}
function chart5(id,title,xlabel,ylabel){
  const c=id5(id),ctx=c.getContext('2d'),w=c.width,h=c.height;
  ctx.clearRect(0,0,w,h);ctx.fillStyle='#0b1e33';ctx.fillRect(0,0,w,h);
  const r={x0:65,x1:w-20,y0:38,y1:h-45};r.ctx=ctx;
  ctx.strokeStyle='#5d7b96';ctx.beginPath();ctx.moveTo(r.x0,r.y0);ctx.lineTo(r.x0,r.y1);ctx.lineTo(r.x1,r.y1);ctx.stroke();
  ctx.fillStyle='#d7ebf8';ctx.font='bold 14px sans-serif';ctx.fillText(title,r.x0,23);
  ctx.font='12px sans-serif';ctx.fillText(ylabel,10,26);ctx.fillText(xlabel,w-175,h-13);
  return r;
}
function placeholder5(id,message){
  const c=chart5(id,'Python 고유해석 결과 미로드','','');
  c.ctx.fillStyle='#95b5cb';c.ctx.font='15px sans-serif';c.ctx.fillText(message,90,id5(id).height/2);
}
function drawBands5(d){
  const r=chart5('pc5-bands-canvas','Γ–M–K–Γ / Γ–X–M–Γ Band','Bloch 경로','Λ/λ');
  const path=d.path,flat=path.bands_f.flat();
  const mn=Math.min(...flat)*.95,mx=Math.max(...flat)*1.05,dist=Math.max(...path.distance_norm)||1;
  const x=z=>r.x0+(r.x1-r.x0)*z/dist;
  const y=z=>r.y1-(r.y1-r.y0)*(z-mn)/(mx-mn||1);
  for(let i=0;i<=4;i++){let f=mn+(mx-mn)*i/4;r.ctx.strokeStyle='#2a4159';
    r.ctx.beginPath();r.ctx.moveTo(r.x0,y(f));r.ctx.lineTo(r.x1,y(f));r.ctx.stroke();
    r.ctx.fillStyle='#b8cede';r.ctx.fillText(fixed5(f),3,y(f)+4);}
  const colors=['#65e3d7','#ffbd86','#c2aff5','#87b5fd','#eddf83','#e8969d'];
  for(let b=0;b<d.input.bands;b++){
    r.ctx.beginPath();r.ctx.strokeStyle=colors[b%colors.length];r.ctx.lineWidth=1.8;
    path.bands_f.forEach((row,i)=>{if(!i)r.ctx.moveTo(x(path.distance_norm[i]),y(row[b]));
      else r.ctx.lineTo(x(path.distance_norm[i]),y(row[b]));});r.ctx.stroke();
  }
  path.labels.forEach((label,i)=>{if(!label)return;r.ctx.fillStyle='#edf8ff';
    r.ctx.fillText(label,x(path.distance_norm[i])-5,r.y1+19);});
  id5('pc5-band-note').textContent='βΛ/2π='+fixed5(path.beta_norm,3)+
    ', 밴드 '+d.input.bands+'개 · '+(d.input.backend==='vector_pwe'?'벡터 Maxwell PWE; cutoff/k 수렴 별도':'기존 스칼라 FDFD 근사');
}
function drawMap5(d){
  const r=chart5('pc5-map-canvas','샘플 k점 간극 후보','normalized Λ/λ','β/k₀');
  const gaps=d.gap_map.candidates;
  const nums=d.path.bands_f.flat().concat(gaps.flatMap(g=>[g.f_low,g.f_high]));
  const mn=Math.min(...nums)*.90,mx=Math.max(...nums)*1.08;
  const maxY=Math.max(1.8,d.input.n_host*1.2,...gaps.map(g=>g.beta_norm/g.f_low*1.1));
  const x=z=>r.x0+(r.x1-r.x0)*(z-mn)/(mx-mn||1);
  const y=z=>r.y1-(r.y1-r.y0)*z/maxY;
  for(let j=0;j<=4;j++){const val=maxY*j/4;r.ctx.fillStyle='#aac7d6';
    r.ctx.fillText(fixed5(val,2),12,y(val)+3);}
  for(let j=0;j<=4;j++){const val=mn+(mx-mn)*j/4;r.ctx.fillStyle='#aac7d6';
    r.ctx.fillText(fixed5(val,2),x(val)-12,r.y1+17);}
  for(const [label,n,color] of [['air',d.input.n_air,'#a0d5fb'],['host',d.input.n_host,'#ffd3a0']]){
    r.ctx.strokeStyle=color;r.ctx.setLineDash([5,4]);r.ctx.beginPath();r.ctx.moveTo(r.x0,y(n));r.ctx.lineTo(r.x1,y(n));r.ctx.stroke();r.ctx.setLineDash([]);
    r.ctx.fillStyle=color;r.ctx.fillText(label+' light line n='+fixed5(n),r.x1-198,y(n)-7);
  }
  gaps.forEach((g,i)=>{r.ctx.beginPath();r.ctx.strokeStyle=['#6cebd6','#ffba85','#c8a8fc','#f5dc82'][i%4];
    r.ctx.lineWidth=6;r.ctx.globalAlpha=.8;
    for(let k=0;k<=30;k++){const f=g.f_low+(g.f_high-g.f_low)*k/30;
      if(!k)r.ctx.moveTo(x(f),y(g.beta_norm/f));else r.ctx.lineTo(x(f),y(g.beta_norm/f));}
    r.ctx.stroke();r.ctx.globalAlpha=1;});
  id5('pc5-gap-count').textContent=gaps.length+'개 후보 · cutoff/grid·k sampling 수렴 및 complete-gap 인증 별도';
  const s=id5('pc5-candidate');s.replaceChildren();
  if(!gaps.length){s.add(new Option('후보 없음',''));}
  gaps.forEach((g,i)=>s.add(new Option('후보 '+i+' · band '+g.band_below+'/'+g.band_above+
    ' · βΛ/2π='+fixed5(g.beta_norm,2)+' · f '+fixed5(g.f_low)+'–'+fixed5(g.f_high),String(i))));
  showGap5();
}
function showGap5(){
  const output=id5('pc5-gap-detail');
  if(!result5){output.textContent='고유밴드 결과가 없습니다.';return;}
  const g=result5.gap_map.candidates[Number(id5('pc5-candidate').value)];
  if(!g){output.textContent='계산된 유효 후보 갭이 없습니다. 결함 설계는 만들 수 없습니다.';return;}
  const avg=(g.f_low+g.f_high)/2;
  output.textContent='βΛ/(2π)='+fixed5(g.beta_norm,4)+
     ' · f=Λ/λ: '+fixed5(g.f_low,6)+'–'+fixed5(g.f_high,6)+
     ' · 상대 간극 '+fixed5(100*g.relative_width,2)+'% · λ중앙='+
     fixed5(1000*result5.input.pitch_um/avg,2)+' nm. 완전 BZ 갭은 미확정.';
}
function applyFile5(d){
  validResult5(d);
  const src=d.input;id5('pc5-backend').value=src.backend==='vector_pwe'?'vector_pwe':'scalar_fdfd';id5('pc5-cutoff').value=src.pwe_cutoff??3;
  for(const [k,el] of Object.entries({
    lattice:'pc5-lattice',pitch_um:'pc5-pitch',hole_diameter_over_pitch:'pc5-hole',
    n_host:'pc5-host',n_air:'pc5-air',grid_n:'pc5-grid',bands:'pc5-bands',
    kmesh_n:'pc5-kmesh',path_points_per_segment:'pc5-path'})){id5(el).value=src[k];}
  id5('pc5-betas').value=src.beta_norms.join(', ');
  drawPreview5();result5=d;drawBands5(d);drawMap5(d);switch5('bands');
  status5('해석 결과 불러오기 성공: 설계값을 JSON 기준으로 동기화했습니다. 표본 간극은 complete gap 인증이 아닙니다.');
}
async function loadExample5(auto=false){
  const resp=await fetch('./research_step5_example_v087.json?v=20261119-v087');
  if(!resp.ok)throw Error('GitHub Pages 예시 데이터 읽기 실패 '+resp.status);
  const demo=await resp.json();
  applyFile5(demo);
  status5('실제 Python 수치해석 예시(참조값 아님) · 입력 N='+demo.input.grid_n+
    ', 밴드 '+demo.input.bands+'개, 표본 갭 후보 '+demo.gap_map.candidates.length+
    '개. 벡터/스칼라 여부는 backend 표기를 확인하세요. complete gap/PCF 제품 검증은 별도입니다.');
  if(auto)switch5('design');
}
async function exportDefect5(){
  if(!result5||!same5(result5.input,read5()))throw Error('현재 설계와 일치하는 해석 결과가 필요합니다.');
  if(result5.input.lattice!=='triangular')throw Error('기존 Full-section FEM과 연동 가능한 격자는 삼각 격자뿐입니다.');
  const idx=Number(id5('pc5-candidate').value),gap=result5.gap_map.candidates[idx];
  if(!gap)throw Error('선택한 계산된 밴드갭 후보가 없습니다.');
  const rings=Number(id5('pc5-rings').value),cells=Number(id5('pc5-defect-size').value),sur=Number(id5('pc5-surround').value);
  if(!Number.isInteger(rings)||rings<2||rings>12||![7,19].includes(cells)||
     !(sur>=0&&sur<=result5.input.pitch_um))throw Error('rings 2..12, cell 7/19, surround 0..Λ 입력 확인');
  const inp=result5.input, mid=(gap.f_low+gap.f_high)/2,wave=1000*inp.pitch_um/mid;
  if(wave<250||wave>3500)throw Error('갭 중심 λ가 기존 Full-section solver의 250–3500 nm 입력 범위를 벗어납니다.');
  const res=await fetch(R5+'presets.json');if(!res.ok)throw Error('기존 프리셋 로드 실패');
  const presets=await res.json(),base=presets['pbgf__'+cells+'_cell'];
  if(!base)throw Error('기존 pbgf 프리셋이 없습니다.');
  base.target_neff=gap.beta_norm/mid;
  if(!(base.target_neff>0&&base.target_neff<3.5))throw Error('β/k₀ target 범위 초과');
  base.name='STEP5 proposed '+cells+'-cell triangular PBGF';base.wavelength_nm=wave;
  Object.assign(base.parameters,{pitch_um:inp.pitch_um,d_over_pitch:inp.hole_diameter_over_pitch,
                               rings,defect_cells:cells,core_surround_um:sur});
  if(cells===19&&rings<3)throw Error('19-cell에는 최소 3개 링이 필요합니다.');
  base.domain_radius_um=inp.backend==='vector_pwe'?(rings+3)*inp.pitch_um:Math.max(base.domain_radius_um,(rings+2)*inp.pitch_um);
  base.materials.background={model:'constant',n:inp.n_host};
  base.materials.silica={model:'constant',n:inp.n_host};
  base.materials.air={model:'constant',n:inp.n_air};
  base.boundary={type:'pml',thickness_um:Math.max(3,inp.pitch_um*1.5),strength:2,order:3};
  base.metadata={status:'proposal_not_validated',
    provenance:'STEP5 computed sampled gap ('+(inp.backend||'scalar_fdfd')+') -> triangular finite PBGF proposal',target_beta_over_k0:base.target_neff,
    source_input_sha256:result5.input_sha256,selected_gap:gap,
    fixed_beta_um_inverse:2*Math.PI*gap.beta_norm/inp.pitch_um,domain_radius_basis:'numerical window; window/PML convergence required',
    limitations:'no proof of vector confinement, full-band gap, mesh/PML independence'};
  download5(base,'step5_finite_pbgf_design.json');
  const editor=document.getElementById('rf-json');if(editor){editor.value=JSON.stringify(base,null,2);editor.dispatchEvent(new Event('input',{bubbles:true}));}
  id5('pc5-handoff-note').textContent='결함 설계를 연구 JSON 편집기에 적용하고 다운로드했습니다. 기존 Python에서 python cli.py solve step5_finite_pbgf_design.json --out out_step5/core_modes --modes 4 ; 생성되는 mode_result.json은 기존 Full-vector 결과 뷰어로 가져올 수 있습니다. 수치 검증 상태는 NOT VERIFIED.';
}
async function init5(){
  const main=document.querySelector('main.workspace');if(!main||id5('research-step5-v083'))return;
  const link=document.createElement('link');link.rel='stylesheet';link.href='./research_step5_v083.css?v=20261118-v083';document.head.append(link);
  const sec=document.createElement('section');sec.id='research-step5-v083';
  const resp=await fetch('./research_step5_v083_panel.html?v=20261119-v087');
  if(!resp.ok)throw Error('STEP5 패널 로드 실패 '+resp.status);
  sec.innerHTML=await resp.text();main.append(sec);
  for(const key of ['design','bands','defect'])id5('pc5-tab-'+key).addEventListener('click',()=>switch5(key));
  sec.querySelectorAll('.pc5-inputs input:not([type=file]), #pc5-pane-design select').forEach(el=>{
    el.addEventListener('change',()=>{
      if(result5&&!same5(result5.input,read5())){
        result5=null;placeholder5('pc5-bands-canvas','설계가 변경되어 기존 결과 무효');
        placeholder5('pc5-map-canvas','Python 재계산이 필요합니다.');
        id5('pc5-candidate').replaceChildren();
        id5('pc5-gap-detail').textContent='설계가 변경되었습니다. 과거 후보는 더 이상 유효하지 않습니다.';
        status5('설계 변경: 이전 결과 무효화됨',true);
      }
      drawPreview5();
    });
  });
  id5('pc5-export-input').addEventListener('click',()=>{
    const cfg=read5(),error=guard5(cfg);
    if(error){status5(error,true);return;}
    download5(cfg,'step5_design.json');
    status5('설계 입력 JSON 저장. Python CLI에서 실제 고유해석을 수행하세요.');
  });
  id5('pc5-import-result').addEventListener('change',async e=>{
    try{const file=e.target.files[0];if(file)applyFile5(JSON.parse(await file.text()));}
    catch(err){status5('해석 JSON 읽기 실패: '+err.message,true);}
  });
  id5('pc5-load-example').addEventListener('click',()=>{
    loadExample5().catch(err=>status5('예시 로드 실패: '+err.message,true));
  });
  id5('pc5-candidate').addEventListener('change',showGap5);
  id5('pc5-export-defect').addEventListener('click',async()=>{
    try{await exportDefect5();status5('결함 설계 JSON 내보내기 성공 · 수치/물리 미검증');}
    catch(err){status5(err.message,true);id5('pc5-handoff-note').textContent=err.message;}
  });
  id5('pc5-load-fem-example').addEventListener('click',async()=>{
    try{
      const get=async n=>{const r=await fetch(R5+'benchmarks/step5_v087/'+n,{cache:'no-store'});if(!r.ok)throw Error('HTTP '+r.status);return r.json();};
      const [r,m]=await Promise.all([get('finite_mode_result.json'),get('finite_metrics.json')]);
      document.dispatchEvent(new CustomEvent('ls-research-result-import',{detail:r}));document.dispatchEvent(new CustomEvent('ls-step4-import',{detail:m}));
      id5('pc5-rings').value=r.input.parameters.rings;id5('pc5-surround').value=r.input.parameters.core_surround_um;id5('pc5-defect-size').value=r.input.parameters.defect_cells;
      location.hash='#research-fiber-v1';status5('사전 계산 2-ring 예시 적용. 코어 파워 1.76–7.70%이며 core-confined mode 확인 사례가 아닙니다. 새 계산·제품 손실 인증 아님.');
    }catch(e){status5('FEM 예시 적용 실패: '+e.message,true);}
  });
  drawPreview5();placeholder5('pc5-bands-canvas','해석 결과 JSON 없음');placeholder5('pc5-map-canvas','밴드갭 결과 JSON 없음');
  loadExample5(true).catch(err=>status5('예시 로드 실패(직접 JSON 업로드 가능): '+err.message,true));
  if(location.hash==='#research-step5-v083')sec.scrollIntoView();
}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',()=>{init5().catch(err=>console.error(err));});
else init5().catch(err=>console.error(err));
