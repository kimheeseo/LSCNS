import {mountDesignMesh} from './research_design_mesh_v084.mjs?v=20261119-v084';
/** LSCNS research fiber v1 — JSON preset editor + Python vector FEM result viewer.
 * Static GitHub Pages: no scikit-fem/gmsh eigen solve in browser.
 */
const ROOT='../fiber_research_v1/';
const $=id=>document.getElementById(id);
const escapeHTML=x=>String(x??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const jsonName=s=>String(s||'fiber').replace(/[^a-zA-Z0-9_-]+/g,'_').slice(0,64);
let presets={},current=null,output=null;
function save(text,name){
  const a=document.createElement('a'); const url=URL.createObjectURL(new Blob([text],{type:'application/json'}));
  a.href=url;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(url),8000);
}
function validate(d){
  if(d?.schema_version!=='1.0.0')throw Error('schema_version 1.0.0 파일이 필요합니다');
  if(!d.family||!d.variant||!d.name)throw Error('family/variant/name 누락');
  if(!Number.isFinite(d.wavelength_nm)||d.wavelength_nm<250||d.wavelength_nm>3500)throw Error('파장은 250~3500 nm');
  if(!Number.isFinite(d.domain_radius_um)||d.domain_radius_um<=1)throw Error('계산창 반지름이 잘못됨');
  if(!d.mesh||!d.boundary||!d.parameters||!d.materials)throw Error('materials/mesh/boundary/parameters 누락');
  if(d.mesh.symmetry!=='full')throw Error('C6v 축약은 벡터 경계조건이 검증되기 전까지 비활성');
  if(!['pec','sbc','pml'].includes(d.boundary.type))throw Error('PEC/SBC/PML 중 선택');
  for(const [k,v]of Object.entries(d.parameters))if(typeof v==='number'&&!Number.isFinite(v))throw Error(k+'가 유한수가 아님');
  for(const [k,m]of Object.entries(d.materials)){
    if(m.model==='ge_doped_sellmeier'||m.model==='f_doped_sellmeier')throw Error(k+' Ge/F 도핑 Sellmeier 계수 미검증(TODO)');
  }
  return d;
}
function status(t){$('rf-status').textContent=t;}

/* Literature-only ESM indicator for conventional triangular hole PCF.
   The d/Lambda=0.406 cutoff is an approximate ideal-lattice guide, NOT a mode calculation.
   Reference: Birks et al., Opt. Lett. 22, 961 (1997), doi:10.1364/OL.22.000961;
   quantitative guide: Folkenberg et al. (2003), https://arxiv.org/abs/physics/0307009 */
function drawESM(design){
  const root=$('rf-esm-panel'),canvas=$('rf-esm-canvas'),explanation=$('rf-esm-interpretation');
  if(!root||!canvas)return;
  const eligible=design?.family==='solid_pcf'&&['hex','lma','high_nonlinearity'].includes(design.variant)&&design.parameters?.defect_cells===1&&(design.parameters?.ellipse_ratio??1)===1;
  root.hidden=!eligible;
  if(!eligible)return;
  const pitch=Number(design.parameters?.pitch_um),ratio=Number(design.parameters?.d_over_pitch);
  const lw=Number(design.wavelength_nm)/1000/pitch;
  const ctx=canvas.getContext('2d'),w=canvas.width,h=canvas.height;
  ctx.clearRect(0,0,w,h);
  ctx.fillStyle='#091728';ctx.fillRect(0,0,w,h);
  const box={x:53,y:20,w:w-76,h:h-58};
  const xs=q=>box.x+q/.9*box.w;
  const ys=q=>box.y+(1.8-q)/1.7*box.h;
  ctx.fillStyle='#143d43';ctx.fillRect(box.x,box.y,xs(.406)-box.x,box.h);
  ctx.fillStyle='#423029';ctx.fillRect(xs(.406),box.y,xs(.9)-xs(.406),box.h);
  ctx.strokeStyle='#9edce4';ctx.lineWidth=1;
  ctx.beginPath();ctx.moveTo(box.x,box.y);ctx.lineTo(box.x,box.y+box.h);ctx.lineTo(box.x+box.w,box.y+box.h);ctx.stroke();
  ctx.strokeStyle='#e0c487';ctx.setLineDash([4,5]);ctx.beginPath();ctx.moveTo(xs(.406),box.y);ctx.lineTo(xs(.406),box.y+box.h);ctx.stroke();ctx.setLineDash([]);
  ctx.fillStyle='#e9f6ff';ctx.font='12px Arial';
  for(const x of [.2,.4,.6,.8])ctx.fillText(x.toFixed(1),xs(x)-8,box.y+box.h+17);
  for(const y of [.2,.6,1,1.4,1.8])ctx.fillText(y.toFixed(1),8,ys(y)+4);
  ctx.fillText('d / Λ',box.x+box.w/2,box.y+box.h+38);
  ctx.save();ctx.translate(12,box.y+box.h/2);ctx.rotate(-Math.PI/2);ctx.fillText('λ / Λ',-20,0);ctx.restore();
  ctx.font='11px Arial';ctx.fillStyle='#d2ede2';ctx.fillText('ESM 근사',box.x+7,box.y+20);
  ctx.fillStyle='#f8dec8';ctx.fillText('추가 모드 가능',xs(.406)+9,box.y+20);
  ctx.fillStyle='#fff';ctx.fillText('0.406',xs(.406)+3,box.y+box.h-8);
  const px=Math.min(.9,Math.max(0,ratio)),py=Math.min(1.8,Math.max(.1,lw));
  ctx.beginPath();ctx.arc(xs(px),ys(py),6,0,2*Math.PI);ctx.fillStyle='#6ce3fb';ctx.fill();ctx.strokeStyle='#061723';ctx.stroke();
  explanation.textContent='입력 d/Λ='+ratio.toFixed(3)+', λ/Λ='+lw.toFixed(3)+'. '+(ratio<=.406?'일반적인 육각 단일 결함 PCF의 ESM 근사 구역입니다.':'기준 초과: 고차모드 가능성을 분석해야 합니다.')+' 이 도표는 FEM/고유모드/컷오프 검증 결과가 아닙니다.';
}

function updateConfig(d){
  current=structuredClone(validate(d));$('rf-json').value=JSON.stringify(current,null,2);drawESM(current);
  // Inform optional result panels: replacing a preset must invalidate stale calculations.
  $('rf-json').dispatchEvent(new Event('input',{bubbles:true}));
  status('JSON 프리셋 로드 완료. 이는 형상 설정이며 모드 계산값이 아닙니다. Python에서 실행해야 합니다.');
}
function summaryMode(mode){
  const fmt=(x,n=8)=>typeof x==='number'&&Number.isFinite(x)?x.toPrecision(n):'—';
  $('rf-results').innerHTML=
    '<div class="rf-metrics">'+[
      ['Re(n_eff)',fmt(mode.neff_real,11)],['Im(n_eff)',fmt(mode.neff_imag,5)],
      ['총손실 후보 dB/km',fmt(mode.alpha_total_dB_km,6)],['고유값 상대 잔차',fmt(mode.residual_relative,5)],
      ['약형 Gauss 잔차',fmt(mode.weak_gauss_relative,5)],
      ['PML 강도분율',fmt(mode.classification?.PML_intensity_fraction,5)]
    ].map(([a,b])=>'<div><small>'+escapeHTML(a)+'</small><strong>'+escapeHTML(b)+'</strong></div>').join('')+'</div>'+
    '<p>모드 상태: <strong>'+escapeHTML(mode.status)+'</strong> · 편광: '+escapeHTML(mode.classification?.polarization_class)+
    ' · 대칭: '+escapeHTML(mode.classification?.symmetry_class)+'</p>';
}
function heat(canvas,coords,fields,part){
  const ctx=canvas.getContext('2d');ctx.clearRect(0,0,canvas.width,canvas.height);
  ctx.fillStyle='#071526';ctx.fillRect(0,0,canvas.width,canvas.height);
  const values=fields?.[part];
  if(!Array.isArray(coords)||!Array.isArray(values)||coords.length!==values.length)return;
  const radius=Math.max(1,...coords.map(p=>Math.max(Math.abs(p[0]),Math.abs(p[1]))));
  const intensity=values.map(v=>Array.isArray(v)&&v.length===2?Math.hypot(v[0],v[1])**2:0);
  const peak=Math.max(...intensity,1e-20);
  for(let i=0;i<coords.length;i++){
    const x=(coords[i][0]/radius*.46+.5)*canvas.width;
    const y=(-coords[i][1]/radius*.46+.5)*canvas.height;
    const q=Math.min(1,Math.sqrt(intensity[i]/peak));
    const hue=208-160*q,light=18+60*q;
    ctx.fillStyle='hsla('+hue+',94%,'+light+'%,0.93)';
    ctx.fillRect(x-2,y-2,4,4);
  }
}
function displayResult(data){
  if(data?.schema!=='LSCNS_fiber_research_vector_v1'||!Array.isArray(data.modes)||!data.modes.length)throw Error('지원하는 Python FEM 결과 schema가 아님');
  const mode=data.modes[0];
  for(const k of ['neff_real','neff_imag','residual_relative','weak_gauss_relative'])if(!Number.isFinite(mode[k]))throw Error(k+' 유한성 검사 실패');
  const comps=mode.field_samples?.components;
  if(!comps||!['Ex','Ey','Ez','Hx','Hy','Hz'].every(k=>Array.isArray(comps[k])))throw Error('복소 6성분 필드 샘플 누락');
  output=data;summaryMode(mode);
  for(const key of ['Ex','Ey','Ez','Hx','Hy','Hz'])heat($('rf-map-'+key),mode.field_samples.coordinates_um,comps,key);
  $('rf-verdict').textContent='수치 결과: '+(data.validation?.level||'검증 메타데이터 없음')+
    ' / 메시 수렴 '+String(data.validation?.mesh_converged===true?'PASS':'미검증')+
    ' / PML 수렴 '+String(data.validation?.pml_converged===true?'PASS':'미검증')+
    '. 샘플 점 산점도이며 정밀 필드 지도가 아닙니다.';
  status('외부 Python JSON 표시 완료. 브라우저에서 벡터 고유값을 계산하지 않았습니다.');
}
function mount(){
  const main=document.querySelector('main.workspace')||document.querySelector('main');if(!main)return;
  const style=document.createElement('style');style.textContent=
  '#research-fiber-v1{margin-top:22px;scroll-margin-top:85px;padding:20px;background:#10213a;color:#e4f2ff;border:1px solid #385573;border-radius:12px;font-size:12px;line-height:1.6}' +
  '#research-fiber-v1 h2{margin:0 0 6px;font-size:19px;color:#f3fcff}#research-fiber-v1 a{color:#8eeae0}' +
  '#research-fiber-v1 .rf-line{display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin:14px 0}' +
  '#research-fiber-v1 select,#research-fiber-v1 button,#research-fiber-v1 textarea,#research-fiber-v1 label.rf-file{border:1px solid #46758d;border-radius:8px;color:#eafbff;background:#1a3853;padding:9px 11px}' +
  '#research-fiber-v1 button,#research-fiber-v1 label.rf-file{cursor:pointer}' +
  '#research-fiber-v1 textarea{display:block;width:100%;font:11px/1.5 Consolas,monospace;min-height:180px;max-height:430px;resize:vertical}' +
  '#research-fiber-v1 .rf-metrics{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:10px;margin:12px 0}' +
  '#research-fiber-v1 .rf-metrics div{border:1px solid #30536a;padding:10px;background:#0a2235;border-radius:8px;min-width:0}' +
  '#research-fiber-v1 .rf-metrics small{display:block;color:#abc9d9}#research-fiber-v1 .rf-metrics strong{font-size:15px;word-break:break-all}' +
  '#research-fiber-v1 .rf-fields{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:10px}' +
  '#research-fiber-v1 .rf-fields figure{margin:0;min-width:0;padding:7px;border:1px solid #31516c;border-radius:7px}' +
  '#research-fiber-v1 .rf-fields canvas{width:100%;height:auto;display:block;aspect-ratio:1}#research-fiber-v1 .rf-muted{color:#d0bb92}' +
  '@media(max-width:700px){#research-fiber-v1 .rf-fields{grid-template-columns:repeat(2,minmax(0,1fr))}#research-fiber-v1 .rf-metrics{grid-template-columns:repeat(2,minmax(0,1fr))}}';
  document.head.appendChild(style);
  const section=document.createElement('section');section.id='research-fiber-v1';section.className='engineering-panel';
  section.innerHTML='<h2>광섬유 연구 · 설계/메시 v0.8.4</h2><p>앱 v0.8.4 · 기본 scalar 엔진 v0.7.1 유지 · Python 벡터 FEM 별도 실행</p>'+
    '<p>8개 계열 · 25개 버전형 JSON 프리셋을 저장/불러옵니다. 지오메트리·복소 풀벡터 FEM은 로컬 또는 Colab Python에서 실행합니다. 웹은 결과 JSON만 시각화합니다.</p>'+
    '<p class="rf-muted">주의: Kagome/negative-curvature/PM 응력과 PBGF 밴드갭은 일부 형상 근사 또는 미구현입니다. 계산된 손실은 메시·PML 독립 검증 전까지 미검증 수치입니다.</p>'+
    '<div class="rf-line"><label>설계 프리셋 <select id="rf-preset"><option>프리셋 불러오는 중…</option></select></label><button id="rf-use">프리셋 적용</button><button id="rf-check">설계 점검</button><button id="rf-save">JSON 저장</button><label for="rf-open" class="rf-file">JSON 열기</label><input type="file" accept=".json" id="rf-open" hidden></div>'+
    '<div id="rf-esm-panel" hidden style="margin:12px 0;border:1px solid #3b6878;padding:14px;border-radius:10px">'+
    '<h3 style="margin:0">PCF Endlessly Single-Mode · 문헌 근사 구역</h3><canvas id="rf-esm-canvas" width="620" height="285" style="width:100%;max-width:620px;height:auto;display:block"></canvas>'+
    '<p id="rf-esm-interpretation"></p><p class="rf-muted">일반적인 triangular-lattice 단일 결함 PCF에서의 참고 기준 d/Λ ≲ 0.406. 유한 링, 다중 결함, 타원 구멍, 고비선형 구조에는 직접 적용되지 않습니다. <a href="https://arxiv.org/abs/physics/0307009" target="_blank" rel="noopener">Folkenberg et al. (2003)</a></p></div>'+
    '<textarea id="rf-json" spellcheck="false" aria-label="설계 JSON 수정"></textarea>'+
    '<p id="rf-status" role="status" aria-live="polite">프리셋 로드 대기</p>'+
    '<p><a href="https://github.com/kimheeseo/LSCNS/tree/main/COMSOL/universal_silica_v01/fiber_research_v1" target="_blank" rel="noopener">Python 설치/실행 가이드 및 코드</a> · <a href="'+ROOT+'fiber_design.schema.json" target="_blank" rel="noopener">JSON Schema</a></p>'+
    '<hr><h3>풀벡터 Python 결과 JSON</h3>'+
    '<div class="rf-line"><label for="rf-result-open" class="rf-file">mode_result.json 가져오기</label><input id="rf-result-open" type="file" accept=".json" hidden><button id="rf-example">결과 내보내기</button></div>'+
    '<p id="rf-verdict" class="rf-muted">결과가 아직 없습니다. Python에서 계산한 JSON을 선택하세요.</p>'+
    '<div id="rf-results"></div>'+
    '<div class="rf-fields">'+['Ex','Ey','Ez','Hx','Hy','Hz'].map(k=>'<figure><figcaption>|'+k+'|² · 개별 정규화</figcaption><canvas id="rf-map-'+k+'" width="260" height="260"></canvas></figure>').join('')+'</div>'+
    '<p class="rf-muted">필드 샘플은 희소 삼각형 중심점의 복소값입니다. 그림은 산점 시각화이며 실제 FEM 재계산, 정량 손실 인증 또는 밴드갭 지도는 아닙니다.</p>';
  main.appendChild(section);
  const revealResearch=()=>{
    const active=location.hash==='#research-fiber-v1'||location.hash==='#rf-design-mesh-v084';
    document.body.classList.toggle('ls-research-view',active);
    if(active)requestAnimationFrame(()=>section.scrollIntoView({block:'start'}));
  };
  window.addEventListener('hashchange',revealResearch);revealResearch();
  mountDesignMesh(section);
  $('rf-use').addEventListener('click',()=>{try{updateConfig(presets[$('rf-preset').value]);}catch(e){status('로드 실패: '+e.message);}});
  $('rf-check').addEventListener('click',()=>{try{const d=validate(JSON.parse($('rf-json').value));drawESM(d);status('브라우저 기본 검사 통과. Python schema.py 상세 검증이 별도로 필요합니다.');}catch(e){status('검사 실패: '+e.message);}});
  $('rf-save').addEventListener('click',()=>{try{const d=validate(JSON.parse($('rf-json').value));save(JSON.stringify(d,null,2)+'\n',jsonName(d.name)+'.json');status('설계 JSON 다운로드 요청');}catch(e){status('저장 실패: '+e.message);}});
  $('rf-open').addEventListener('change',async e=>{try{const f=e.target.files[0];if(!f)return;if(f.size>300000)throw Error('300 KB 제한');updateConfig(JSON.parse(await f.text()));}catch(err){status('파일 읽기 실패: '+err.message);}});
  $('rf-result-open').addEventListener('change',async e=>{try{const f=e.target.files[0];if(!f)return;if(f.size>15000000)throw Error('15 MB 제한');displayResult(JSON.parse(await f.text()));}catch(err){status('결과 읽기 실패: '+err.message);}});
  $('rf-example').addEventListener('click',()=>{if(!output){status('먼저 결과 JSON을 불러와야 합니다.');return;}save(JSON.stringify(output,null,2),'fiber_mode_result_copy.json');});
  fetch(ROOT+'presets.json',{cache:'no-store'}).then(r=>{if(!r.ok)throw Error('HTTP '+r.status);return r.json();})
    .then(obj=>{presets=obj;const options=Object.keys(obj).sort();$('rf-preset').innerHTML=options.map(k=>'<option value="'+escapeHTML(k)+'">'+escapeHTML(k.replace('__',' / '))+'</option>').join('');updateConfig(obj[options[0]]);})
    .catch(e=>status('프리셋 다운로드 실패: '+e.message));
}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',mount);else mount();

