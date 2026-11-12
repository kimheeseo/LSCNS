import {BASE,config,buildYeeGrid,sampler} from './vector_geometry.mjs?v=20261112-m3m6-v2';
import {normalizeReference,compareReference,twoCoreMetrics,nCoreMetrics,jobsFor,rowsCSV,migrateProject,SCHEMA} from './vector_study.mjs?v=20261112-m3m6-v2';
import {DEFAULT as HCF_DEFAULT} from './hcf_designer.mjs';
import {silicaIndex} from './solver.mjs';
import {evaluateM2Material} from './m2_preview_engine_v06.mjs?v=20261112-m3m6-v1';
const $=id=>document.getElementById(id),esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const fmt=(n,d=9)=>Number.isFinite(n)?Number(n).toPrecision(d):'미검증',exp=n=>Number.isFinite(n)?n.toExponential(3):'미검증';
const read=id=>Number($(id).value),field=(id,label,value,extra='')=>`<label>${label}<input id="${id}" type="number" value="${value}" ${extra}></label>`;
let latest=null,history=[],reference=null,worker=null,request=0,busy=false,currentSignature=null,geometry={kind:'radial'},polygonPoints=[],activeTab='m3',imageSource=null;
const feature=globalThis.LSWaveFeatures?.vectorWorkbench!==false;
function save(name,data,type='application/json'){
 const url=URL.createObjectURL(new Blob([typeof data==='string'?data:JSON.stringify(data,null,2)],{type})),a=document.createElement('a');a.href=url;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(url),10000);
}
const geomOptions='<option value="radial">기존 동심원 입력</option><option value="regions">폴리곤 / 원 영역</option><option value="mask">SVG / SEM 이미지 마스크</option><option value="mcf">N코어 직접 해석</option><option value="hcf">HCF / NANF 설계기</option>';
function init(){
 if(!feature||$('vwb'))return;
 const root=document.createElement('section');root.id='vwb';root.className='engineering-panel';root.hidden=true;
 root.innerHTML=`<header><small>M3 → M6 · BROWSER MAXWELL WORKBENCH · 2026.11.12</small><h2>Full-vector · PML · 임의 단면</h2><p>브라우저에서 2D Yee 격자 고유모드를 계산합니다. 기존 스칼라 FEM·Colab 뷰어는 각자의 계산 경로를 유지합니다.</p><p><a style="color:#78e9d7" href="https://github.com/kimheeseo/LSCNS/blob/main/COMSOL/universal_silica_v01/web/M3_M6_YEE_VALIDATION_20261112.md" target="_blank" rel="noopener">수용 기준·비교표·미검증 항목 보고서 ↗</a></p></header>
 <nav role="tablist">${[['m3','M3 · Mode analysis'],['m4','M4 · PML / Bend'],['m5','M5 · Cross section'],['m6','M6 · Study / Project'],['results','Results / Compare']].map(([v,t])=>`<button data-vb-tab="${v}" id="vb-tab-${v}" role="tab" aria-selected="${v==='m3'}">${t}</button>`).join('')}</nav>
 <div class="vb-pane" data-vb-pane="m3"><h3>2D Full-vector Mode analysis</h3><div class="vb-grid">
 ${field('vb-wavelength','파장 λ (nm)',1550,'min="400" max="1700"')}${field('vb-radius','코어 반경 a (µm)',4.1,'min="0.1" step=".1"')}${field('vb-core','코어 n',BASE.n_core,'step=".00001"')}${field('vb-clad','클래딩 n',BASE.n_clad,'step=".00001"')}${field('vb-half','계산창 반폭 (µm)',18,'min="1"')}${field('vb-grid','Yee 격자 N × N (24–192)',48,'min="24" max="192" step="1"')}${field('vb-modes','요청 모드 수 (1–16)',2,'min="1" max="16"')}
 <label>목표 neff (빈칸: 자동)<input id="vb-target" type="number" placeholder="자동" step="any"></label>
 <label>형상 입력<select id="vb-kind">${geomOptions}</select></label><label>재료 입력<select id="vb-material-source"><option value="constant">직접 n / κ</option><option value="m2">기존 M2 선택 재료 연결</option><option value="json">아래 영역별 M2 JSON</option></select></label>
 ${field('vb-kcore','코어 κ ≥ 0',0,'min="0" step="any"')}${field('vb-kclad','클래딩 κ ≥ 0',0,'min="0" step="any"')}
 </div><div class="vb-actions"><button id="vb-from-legacy">기존 스칼라 입력 가져오기</button><button id="vb-python-case">저장된 Colab 비교 케이스</button><button id="vb-open-m2">기존 M2 재료 패널 열기</button></div>
 <p id="vb-estimate" class="vb-note"></p><p class="vb-help">Yee를 선택한 이유: 임의 단면·이미지와 동일한 Cartesian 격자에서 Maxwell curl을 이산화하고 PML stretching을 각 미분에 적용할 수 있습니다. 전역 행렬은 희소 CSR이며 dense 행렬을 만들지 않습니다.</p>
 <details><summary>알고리즘 · 수렴 기준 · M3 범위</summary><p>전파 규약 exp(iβz−iωt). 횡방향 전기장 Ex/Ey의 β²/k₀²=PQ 고유값 문제. Block shift-invert 역반복, 복소 ILU(0)-GMRES(재시작 36), Rayleigh–Ritz. 외측 잔차 기본 10⁻⁸, PML/흡수는 10⁻¹⁰. 내측 잔차 min(2×10⁻¹⁰, 외측 허용값×0.02). 제한 횟수에서 실패하면 결과를 생성하지 않습니다.</p><p>M3 단독은 유한 사각 경계·실수 재료이며 PML/HCF는 M4를 켜야 지원합니다. 필드 맵은 Yee 위치에서 동일 좌표로 보간한 복소 Ex/Ey/Ez입니다. MFD=2√(2⟨r²⟩), Aeff=(∫I)²/∫I², I=|Ex|²+|Ey|²+|Ez|². PML을 제외한 영역에서 적분합니다.</p></details></div>
 <div class="vb-pane" data-vb-pane="m4" hidden><h3>PML · 복소 neff · 굽힘</h3><label class="vb-check"><input id="vb-pml" type="checkbox"> M4 복소 stretching PML 활성화</label><div class="vb-grid">
 ${field('vb-pml-thickness','PML 두께 (µm)',4,'min="0" step=".5"')}${field('vb-pml-strength','PML 강도 σmax',3,'min="0.1" max="30" step=".5"')}${field('vb-pml-order','PML 프로파일 차수 m',3,'min="1" max="6"')}${field('vb-bend','굽힘 반경 R (µm), 0=직선',0,'min="0" step="100"')}
 </div><p class="vb-help">sₓ=1+iσmax·tₓᵐ, sᵧ=1+iσmax·tᵧᵐ, ∂→s⁻¹∂. 외곽에서 PEC로 종단합니다. PML 두께는 3셀 이상입니다. κ는 M2의 α(dB/km)→κ 변환도 그대로 연결합니다.</p>
 <details open><summary>손실 환산식 · 굽힘 모델</summary><p>k₀=2π/λ [m⁻¹], λ=λnm×10⁻⁹ m.<br>αtotal [dB/km]=(20/ln 10)·k₀·Im(neff)·1000.<br>ñ=n+iκ, 양의 Im(neff)가 감쇠입니다. 음수값은 절댓값으로 바꾸지 않습니다.<br>Conformal 등가 굴절률 ñbend(x,y)=ñ(x,y)·exp(x/R). R≥계산창 반폭×10. 이는 등방성 등가 굴절률 근사이며 완전한 곡선 좌표 벡터 해석은 아닙니다.<br>Loss[dB/turn]=α[dB/km]·2πR[µm]×10⁻⁹.</p><p>κ≠0 결과는 흡수+복사 총손실입니다. 같은 구조의 κ=0 결과를 별도로 계산해 총손실 차이를 비교해야 합니다. 굽힘 추가손실은 같은 구조·격자·PML의 직선 대비 차이로 표시합니다.</p></details>
 <div class="vb-actions"><button id="vb-pml-study">PML 두께·강도 3×3 수렴 검사</button><button id="vb-bend-study">R 의존성 + 직선 4점 검사</button><button id="vb-abs-study">현재 κ vs κ=0 비교</button></div>
 <p class="vb-note">가짜 모드 후보 제외: PML 영역의 ∫|E|² 비율 &gt;35%, 물리 영역 &lt;65%, 음의 Im(neff) &lt;−10⁻⁸, 고유값 잔차 초과. 이 기준만으로 물리 모드임을 보증하지 않습니다. 격자·PML·계산창 변화와 필드 위치도 확인하세요. 잔차로 계산한 손실 민감도 규모는 엄밀한 오차 상한이 아닙니다.</p>
 <div class="vb-actions"><button id="vb-from-hcf">기존 HCF 설계기 입력 재사용</button><label class="vb-upload">HCF 설계 JSON<input type="file" id="vb-hcf-file" accept=".json" hidden></label></div><p id="vb-hcf-state" class="vb-help">HCF 고유값이 계산되어도 튜브 벽 해상도와 Colab 복소 neff 비교가 없으면 손실은 미검증입니다.</p></div>
 <div class="vb-pane" data-vb-pane="m5" hidden><h3>공통 2D 단면 → Yee 유전율 격자</h3>
 <div class="vb-grid"><label>형상 선택<select id="vb-kind-copy">${geomOptions}</select></label><label>MCF 코어 수 N<input id="vb-mcf-count" type="number" min="2" max="19" value="2"></label>${field('vb-mcf-pitch','MCF 코어 간격 (µm)',16,'min="1" step="1"')}${field('vb-mcf-length','이상적 결합 관측 길이 (mm)',1,'min="0" max="100" step=".1"')}</div>
 <label class="vb-check"><input id="vb-subpixel" type="checkbox" checked> 경계 subpixel averaging (Yee 각 위치에서 4×4 ε 면적 샘플 평균)</label><p class="vb-help">ε=(n+iκ)²를 평균합니다. 계단형 경계와 비교할 수 있으며, 경계 텐서의 정확한 해석적 homogenization은 아닙니다. N≥3은 중심+육각 격자로 생성합니다. N코어 초모드 전체를 얻으려면 요청 모드 수와 목표 neff를 조정하세요.</p>
 <div class="vb-two"><div><h3>폴리곤 편집</h3><p class="vb-help">캔버스 클릭으로 꼭짓점을 추가합니다. 아래 x,y 좌표를 수정하여 꼭짓점을 이동·삭제할 수 있습니다. 마지막 영역이 앞선 영역 위에 놓입니다.</p><canvas id="vb-geometry-canvas" width="410" height="410" aria-label="단면 편집: 클릭하여 polygon vertex 추가"></canvas><label>꼭짓점 x,y (µm) · 한 줄에 한 점<textarea id="vb-points" rows="5" placeholder="-4,-4&#10;4,-4&#10;4,4&#10;-4,4"></textarea></label><div class="vb-grid"><label>영역 재료 ID<select id="vb-region-material"><option>core</option><option>clad</option><option>air</option></select></label><label>편집할 영역<select id="vb-region-list"><option value="-1">새 영역</option></select></label></div><div class="vb-actions"><button id="vb-add-region">영역 저장</button><button id="vb-delete-region">영역 삭제</button><button id="vb-clear-points">꼭짓점 초기화</button><button id="vb-circle-region">원형 참조 영역 만들기</button></div></div>
 <div><h3>SVG / SEM 이미지 마스크</h3><label class="vb-upload">파일 선택 · SVG/PNG/JPEG<input id="vb-mask-file" type="file" accept="image/svg+xml,image/png,image/jpeg" hidden></label><div class="vb-grid"><label>색상 분류<select id="vb-mask-mode"><option value="bw">흑백 임계값</option><option value="color">색상 군집 (최대 8색)</option></select></label>${field('vb-threshold','흑백 임계값 0–255',128,'min="0" max="255"')}${field('vb-mask-half','이미지 전체 반폭 (µm)',18,'min=".1" step="1"')}</div><button id="vb-rebuild-mask">마스크 재생성</button><p class="vb-help">최대 512×512. 이미지 폭 전체를 물리 단위로 지정하세요. SEM 스케일바는 자동 해석하지 않습니다. SVG 도형·path는 512px 래스터 마스크로 가져옵니다. 외부 참조·스크립트·애니메이션 SVG는 거부합니다.</p><div id="vb-palette"></div><p id="vb-mask-state" class="vb-help">이미지 파일 대기</p>
 <label>영역 JSON (고급 · circle/polygon)<textarea id="vb-regions-json" rows="4" spellcheck="false"></textarea></label><button id="vb-apply-regions">영역 JSON 적용</button></div></div>
 <details><summary>영역별 M2 재료 정의 JSON</summary><p>예: {"core":{"kind":"geO2","molPercent":5,"absorption":{"kind":"alpha","value":0.2}},"clad":{"kind":"silica"},"air":{"kind":"constant","n":1.00027,"kappa":0}}. 직접 n/κ 입력 또는 M2 선택 시 기본 core/clad/air ID를 제공합니다.</p><textarea id="vb-materials-json" rows="6" spellcheck="false">{"core":{"kind":"constant","n":1.449023621703261,"kappa":0},"clad":{"kind":"constant","n":1.444023621703261,"kappa":0},"air":{"kind":"constant","n":1.00027,"kappa":0}}</textarea><button id="vb-materials-apply">재료 ID 목록 갱신</button></details>
 <div class="vb-actions"><button id="vb-subpixel-study">Subpixel OFF / ON 비교</button><button id="vb-grid-study">현재 형상 격자 수렴 3점</button></div></div>
 <div class="vb-pane" data-vb-pane="m6" hidden><h3>Parametric sweep · Project</h3><p class="vb-help">1–2개 파라미터, 축별 2–18점·전체 36점 이하. 각 점을 동일 Worker에서 순서대로 계산합니다. 결과 맵은 마지막 성공 지점, 지점별 neff/메타데이터는 표와 JSON에 저장합니다. 스윕 도중 취소하면 미완료 연구 결과를 확정하지 않습니다.</p>
 ${[1,2].map(i=>`<fieldset><legend>스윕 축 ${i}</legend>${i===2?'<label class="vb-check"><input id="vb-axis2" type="checkbox"> 두 번째 파라미터 활성화</label>':''}<div class="vb-grid"><label>파라미터<select id="vb-param${i}">${[['a','코어 반경 a (µm)'],['delta_n','Δn'],['lambda','파장 λ (nm)'],['R','굽힘 R (µm)'],['grid','격자 N (정수)'],['pml_thickness','PML 두께'],['pml_strength','PML 강도'],['pitch','MCF pitch']].map(([v,t])=>`<option value="${v}">${t}</option>`).join('')}</select></label>${field('vb-start'+i,'시작',i===1?3.5:1500,'step="any"')}${field('vb-end'+i,'종료',i===1?5:1600,'step="any"')}${field('vb-count'+i,'점 수',3,'min="2" max="18"')}</div></fieldset>`).join('')}
 <div class="vb-actions"><button id="vb-sweep-run">파라미터 스윕 실행</button><button id="vb-project-save">전체 프로젝트 JSON 저장</button><label class="vb-upload">프로젝트 / 구버전 설계 불러오기<input id="vb-project-open" type="file" accept=".json" hidden></label></div><details><summary>프로젝트 JSON 붙여넣기</summary><textarea id="vb-project-paste" rows="5" placeholder="프로젝트 JSON 또는 구버전 {version,model} JSON"></textarea><button id="vb-project-paste-load">붙여넣은 JSON 불러오기</button></details><p id="vb-project-state" class="vb-help">Parameters, CAD, Materials, PML, 기존 입력, 새 격자/형상/재료, 저장 결과를 함께 저장합니다.</p>
 <details><summary>이전 계산 결과 저장 · UI 기능 플래그</summary><div class="vb-actions"><button id="vb-history-save">현재 결과 저장</button><button id="vb-history-clear">저장 결과 비우기</button></div><p id="vb-history-state">저장 결과 없음</p><p>window.LSWaveFeatures.vectorWorkbench=false 로 새 UI를 비활성화할 수 있습니다. 상단 ‘기존 단일 페이지’ 버튼은 계산 엔진을 바꾸지 않고 레이아웃을 즉시 전환합니다. URL ?workflow=legacy 도 지원합니다.</p></details></div>
 <div class="vb-progress"><div class="vb-actions"><button class="vb-primary" id="vb-run">▶ 현재 단면 계산</button><button id="vb-cancel" disabled>■ 취소</button><button id="vb-result-json" disabled>Colab 호환 JSON</button><button id="vb-report-json" disabled>연구 보고서 JSON</button><button id="vb-csv" disabled>스윕 CSV</button></div><progress id="vb-progress" value="0" max="100"></progress> <span id="vb-percent">0%</span><p id="vb-status" role="status" aria-live="polite">대기 · Worker에서 실행합니다.</p></div>
 <div class="vb-pane" data-vb-pane="results" hidden><h3>모드 · 필드 · 비교</h3><p id="vb-mode-summary"></p><div class="vb-scroll"><table><thead><tr><th>모드 / 편광</th><th>Re(neff)</th><th>Im(neff)</th><th>고유 잔차</th><th>MFD µm</th><th>Aeff µm²</th><th>총손실 dB/km</th><th>잔차 민감도 규모</th><th>dB/turn</th><th>PML |E|² 비율</th></tr></thead><tbody id="vb-mode-rows"></tbody></table></div>
 <div class="vb-grid"><label>표시 모드<select id="vb-mode-select"></select></label><label>맵 표시<select id="vb-field-type"><option value="intensity">|E|² 강도</option><option value="real">Re(E) 부호 포함</option><option value="imag">Im(E) 부호 포함</option></select></label></div><div class="vb-maps">${['Itotal','Ex','Ey','Ez'].map(k=>`<div><h4>${k}</h4><canvas id="vb-map-${k}" width="240" height="240"></canvas></div>`).join('')}</div>
 <h3>격자 / PML / R / 스윕 수렴 표</h3><div id="vb-study-table" class="vb-scroll"></div><canvas id="vb-study-chart" width="850" height="300" aria-label="Sweep result with saved curves" style="width:100%;margin-top:12px"></canvas><p id="vb-curve-state" class="vb-help"></p><p id="vb-loss-difference" class="vb-note"></p>
 <h3>Colab · 이전 결과 오버레이</h3><div class="vb-actions"><button id="vb-load-python">저장된 Python 기준 불러오기</button><button id="vb-load-hcf-reference">기존 Colab HCF 기준 불러오기</button><label class="vb-upload">비교 JSON 선택<input id="vb-reference-file" type="file" accept=".json" hidden></label><label>저장 결과 선택<select id="vb-history-select"><option value="">선택 없음</option></select></label></div><p id="vb-reference-state" class="vb-help">기준 JSON 없음 · 미검증</p><div id="vb-reference-table" class="vb-scroll"></div><div class="vb-grid"><label>오버레이 필드<select id="vb-overlay-field"><option>Itotal</option><option>Ex</option><option>Ey</option><option>Ez</option></select></label><label>비교 필드 혼합비<input type="range" id="vb-overlay-opacity" min="0" max="100" value="50"></label></div><canvas id="vb-overlay" width="340" height="340" aria-label="Coordinate-resampled field overlay"></canvas><p id="vb-overlay-state" class="vb-help"></p>
 <h3>MCF 초모드 분리 / 이상적 결합</h3><div id="vb-mcf-table" class="vb-scroll"></div><p class="vb-help">κ=π|Δneff|/λ, Lc=π/(2κ), P₂=sin²(κL). 이는 동일한 2코어·손실 없는 결합 모델입니다. N≥3은 코어 내부 Ex 평균으로 U를 구성하여 H=U·diag(β)·U⁻¹의 비대각 결합과 초모드 전파의 코어별 투영 누화 지표를 표시합니다. 각 편광에서 N개 모드가 필요하고 기저 수렴은 미검증입니다. 확률적 케이블 XT와는 다릅니다.</p>
 </div>`;
 document.querySelector('main.workspace').append(root);
 const nav=document.querySelector('nav.eng-shortcuts'),a=document.createElement('a');a.id='vwb-nav';a.href='#vwb';a.textContent='M3–M6 · YEE / PML / 2D';nav?.append(a);a.addEventListener('click',e=>{e.preventDefault();show(activeTab)});
 try{mountTree();events();estimate();geometryPreview();}catch(error){$("vb-status").textContent="UI initialization failed: "+error.message;console.error(error);return;}
 globalThis.LSWaveVectorWorkbench=Object.freeze({show,latest:()=>latest,config:readConfig,run:(axes=[])=>run(axes),cancel,snapshot:projectSnapshot,loadProject,history:()=>structuredClone(history)});
 if(location.hash==='#vwb')show('m3');
}
function show(tab='m3'){
 $('vwb').hidden=false;activeTab=tab;
 for(const el of document.querySelectorAll('[data-vb-pane]'))el.hidden=el.dataset.vbPane!==tab;
 for(const el of document.querySelectorAll('[data-vb-tab]'))el.setAttribute('aria-selected',String(el.dataset.vbTab===tab));
 $('vwb').scrollIntoView({block:'start',behavior:'instant'});
 if(tab==='m5')geometryPreview();
}
function readConfig(){
 let materials=null;
 const source=$('vb-material-source').value;
 if(source==='m2'){
  const m=globalThis.LSWaveM2Preview?.snapshot();if(!m)throw Error('M2 editor unavailable');materials={...m,air:{kind:'constant',n:1.00027}};
 }else if(source==='json')materials=JSON.parse($('vb-materials-json').value);
 let g={...geometry,kind:$('vb-kind').value};
 if(g.kind==='mcf')g={kind:'mcf',count:read('vb-mcf-count'),pitch_um:read('vb-mcf-pitch')};
 if(g.kind==='regions')g={kind:'regions',regions:geometry.regions??[],background:geometry.background??'clad'};
 if(g.kind==='mask'&&!g.labels)throw Error('Import an image/SVG mask first');
 const c=config({wavelength_nm:read('vb-wavelength'),core_radius_um:read('vb-radius'),n_core:read('vb-core'),n_clad:read('vb-clad'),
  grid_n:read('vb-grid'),window_halfwidth_um:read('vb-half'),modes:read('vb-modes'),target_neff:$('vb-target').value===''?null:read('vb-target'),
  k_core:read('vb-kcore'),k_clad:read('vb-kclad'),subpixel:$('vb-subpixel').checked,geometry:g,materials,
  profile:geometry.profile??'step',shells:geometry.shells,trench_start_um:geometry.trench_start_um,trench_end_um:geometry.trench_end_um,trench_delta_n:geometry.trench_delta_n,graded_alpha:geometry.graded_alpha,
  pml:{enabled:$('vb-pml').checked,thickness_um:read('vb-pml-thickness'),strength:read('vb-pml-strength'),order:read('vb-pml-order')},bend_radius_um:read('vb-bend')});
 if(materials){c.n_core=materials.core?evaluateMaterial(materials.core,c.wavelength_nm).n:c.n_core;c.n_clad=materials.clad?evaluateMaterial(materials.clad,c.wavelength_nm).n:c.n_clad;}
 return c;
}
function evaluateMaterial(m,w){return m.kind==='constant'?{n:m.n,kappa:m.kappa??0}:evaluateM2Material(m,w);}
function writeConfig(c){
 const map={wavelength_nm:'vb-wavelength',core_radius_um:'vb-radius',n_core:'vb-core',n_clad:'vb-clad',grid_n:'vb-grid',window_halfwidth_um:'vb-half',modes:'vb-modes',k_core:'vb-kcore',k_clad:'vb-kclad',bend_radius_um:'vb-bend'};
 for(const [k,id]of Object.entries(map))$(id).value=c[k]??0;
 $('vb-target').value=c.target_neff??'';$('vb-subpixel').checked=c.subpixel!==false;
 for(const [k,id]of [['thickness_um','vb-pml-thickness'],['strength','vb-pml-strength'],['order','vb-pml-order']])$(id).value=c.pml?.[k]??BASE.pml[k];
 $('vb-pml').checked=!!c.pml?.enabled;geometry=structuredClone(c.geometry??{kind:'radial'});geometry.profile=c.profile??'step';
 for(const k of ['shells','trench_start_um','trench_end_um','trench_delta_n','graded_alpha'])if(c[k]!=null)geometry[k]=c[k];
 $('vb-kind').value=$('vb-kind-copy').value=geometry.kind;
 if(geometry.kind==='mcf'){$('vb-mcf-count').value=geometry.count;$('vb-mcf-pitch').value=geometry.pitch_um;}
 if(c.materials){$('vb-material-source').value='json';$('vb-materials-json').value=JSON.stringify(c.materials,null,2);}else $('vb-material-source').value='constant';
 updateRegions();updatePalette();estimate();geometryPreview();
}
function setBusy(value){
 busy=value;for(const id of ['vb-run','vb-sweep-run','vb-grid-study','vb-pml-study','vb-bend-study','vb-subpixel-study','vb-abs-study','vb-project-open'])$(id).disabled=value;
 $('vb-cancel').disabled=!value;
 for(const id of ['vb-result-json','vb-report-json','vb-csv'])$(id).disabled=value||!latest;
}
function cancel(){if(!busy)return;++request;worker?.terminate();worker=null;setBusy(false);$('vb-status').textContent='사용자 취소 · Worker 종료. 미완료 결과는 저장하지 않았습니다.';$('vb-percent').textContent='취소';}
function estimate(){
 try{
  const n=read('vb-grid'),m=read('vb-modes'),bytes=n*n*(m+50)*32/1048576,seconds=Math.max(1,(n/48)**3*(m/2)*5);
  $('vb-estimate').textContent=`상한 192×192 · ${m} modes · 작업 배열 모델 약 ${bytes.toFixed(1)} MiB. 사전 시간 모델 ${seconds.toFixed(0)}–${(seconds*5).toFixed(0)}초/점 (CPU·구조에 따라 초과 가능, 실측 아님). PML·HCF·큰 스윕은 수 분 이상 걸릴 수 있습니다.`;
 }catch{}
}
function run(axes=[],explicitJobs=null){
 if(busy)return;
 let input,jobs;try{input=readConfig();sampler(input);jobs=jobsFor(input,axes);if(explicitJobs)jobs=explicitJobs;}
 catch(error){$('vb-status').textContent='입력 오류: '+error.message;return;}
 const id=++request;setBusy(true);$('vb-progress').value=0;$('vb-percent').textContent='0%';$('vb-status').textContent=`Worker 시작 · ${jobs.length}점 · 시간 표시는 계산 후 실측합니다.`;
 currentSignature=JSON.stringify(readConfig());
 try{
  worker=new Worker(new URL('./vector_study_worker.mjs?v=20261112-m3m6-v2',import.meta.url),{type:'module'});
  worker.onmessage=e=>{
   const p=e.data;if(p.id!==id||id!==request)return;
   if(p.type==='progress'){
    $('vb-progress').value=p.percent??0;$('vb-percent').textContent=Math.round(p.percent??0)+'% (반복 횟수 기반 추정)';
    $('vb-status').textContent=`${p.job}/${p.total} · ${p.phase} · iter ${p.iteration??'—'} · 잔차 ${exp(p.residual)} ${p.row?.status??''}`;
   }else if(p.type==='error')finishError(p.message);
   else if(p.type==='result'){
    worker.terminate();worker=null;latest=p.report;setBusy(false);render();show('results');$('vb-progress').value=100;$('vb-percent').textContent='100%';
    const ok=latest.rows.filter(r=>r.neff!=null).length;
    $('vb-status').textContent=`${ok}/${latest.rows.length}점 성공 · 실측 총 시간 ${(latest.wallMs/1000).toFixed(2)} s. `+(JSON.stringify(readConfig())!==currentSignature?'입력이 변경됨: 표는 실행 시 입력 결과입니다.':'');
   }
  };
  worker.onerror=e=>finishError(e.message||'Worker loading failed');
  worker.postMessage({id,input,axes,reference,jobs:explicitJobs});
 }catch(error){finishError(error.message);}
}
function finishError(message){worker?.terminate();worker=null;setBusy(false);$('vb-status').textContent='계산 실패 / 미검증: '+message;}
function table(target,rows,columns){
 const el=$(target);el.replaceChildren();const t=document.createElement('table'),head=t.createTHead().insertRow();
 for(const [key,title]of columns){const th=document.createElement('th');th.textContent=title;head.append(th);}
 const body=t.createTBody();
 for(const row of rows){const tr=body.insertRow();for(const [key]of columns){const td=tr.insertCell();td.textContent=typeof row[key]==='number'?fmt(row[key],11):(row[key]??'미검증');}}
 el.append(t);
}
function render(){
 if(!latest)return;const result=latest.result;
 $('vb-mode-rows').innerHTML=(result?.modes??[]).map((m,i)=>`<tr><td>${i+1} · ${esc(m.polarization)}</td><td>${fmt(m.neff,13)}</td><td>${exp(m.neff_imag)}</td><td>${exp(m.residual)}</td><td>${fmt(m.MFD_2nd_um,7)}</td><td>${fmt(m.Aeff_um2,7)}</td><td>${fmt(m.loss_dB_km,7)}</td><td>${fmt(m.loss_residual_scale_dB_km,5)}${m.loss_below_residual_scale?' · 미검증':''}</td><td>${fmt(m.loss_dB_turn,6)}</td><td>${fmt(m.pml_power_fraction*100,5)}%</td></tr>`).join('');
 $('vb-mode-select').innerHTML=(result?.modes??[]).map((m,i)=>`<option value="${i}">Mode ${i+1} · ${esc(m.polarization)}</option>`).join('');
 $('vb-mode-summary').textContent=result?`${result.method} · ${result.input.grid_n}×${result.input.grid_n} · ${result.diagnostics.warnings.join(' / ')} · ${result.validation.status}`:'모든 점 실패: 수치 결과 없음';
 const axes=latest.axes??[],axisKeys=axes.map(a=>a.parameter),extra=['variant','R','pml_thickness','pml_strength','grid'].filter(k=>!axisKeys.includes(k)&&latest.rows.some(r=>r[k]!=null));
 table('vb-study-table',latest.rows,[['job','#'],...axisKeys.map(k=>[k,k]),...extra.map(k=>[k,k]),['grid_n','N'],['neff','neff'],['neff_imag','Im(neff)'],['analytic','M1 LP01'],['scalarFEM','Scalar FEM'],['deltaAnalytic','Δ M1'],['deltaScalar','Δ scalar'],['deltaPrevious','Δ 이전'],['loss_dB_km','dB/km'],['residual','잔차'],['wallMs','실측 ms'],['workspaceEstimateBytes','작업 배열 추정 B'],['heapMeasuredBytes','heap 실측 B'],['status','상태']]);
 drawStudyChart();renderFields();renderReference();renderMCF();
 const successes=latest.rows.filter(r=>Number.isFinite(r.neff));
 if(axes.some(a=>a.parameter==='grid')&&successes.length>=2){const d=Math.abs(successes.at(-1).neff-successes.at(-2).neff);$('vb-curve-state').textContent=`마지막 격자 |Δneff|=${exp(d)} · ${d<=1e-6?'neff 기준 ≤1e-6 통과':'neff 기준 미수렴'} · continuum/loss 정확도 보증 아님.`;}
 const straight=successes.find(r=>r.R===0),bent=successes.filter(r=>r.R>0),zero=successes.find(r=>r.variant==='kappa=0'),actual=successes.find(r=>r.variant==='current kappa');
 $('vb-loss-difference').textContent=zero&&actual?`총손실 − κ=0 복사손실 = ${fmt(actual.loss_dB_km-zero.loss_dB_km,7)} dB/km (재료 변화에 따른 필드 차이 포함; 미검증).`:straight&&bent.length?bent.map(r=>`R ${r.R}µm: 직선 대비 ${fmt(r.loss_dB_km-straight.loss_dB_km,7)} dB/km`).join(' · '):'손실값은 mesh/PML/창 크기 수렴을 확인하기 전까지 미검증입니다. κ=0과 비교하면 복사 성분을 따로 확인할 수 있습니다.';
 if(straight&&bent.length){const sorted=bent.slice().sort((a,b)=>a.R-b.R),monotonic=sorted.every((r,i)=>i===0||r.loss_dB_km<=sorted[i-1].loss_dB_km);$('vb-loss-difference').textContent+=' · R 감소 시 손실 증가: '+(monotonic?'해당 계산점에서 관찰됨':'미검증 / 비단조: 모드 교차·수렴 확인 필요');}
}
function mapCanvas(canvas,grid,signed=false){
 const ctx=canvas.getContext('2d'),H=grid?.length??0,W=grid?.[0]?.length??0;ctx.clearRect(0,0,canvas.width,canvas.height);if(!H||!W)return;
 let max=1e-30;for(const row of grid)for(const v of row)max=Math.max(max,Math.abs(v));
 for(let j=0;j<H;j++)for(let i=0;i<W;i++){const v=grid[j][i]/max,t=Math.max(0,v);ctx.fillStyle=signed?(v>=0?`rgb(${60+190*v},${70+100*v},${100-60*v})`:`rgb(${70+50*v},${80-70*v},${100-145*v})`):`rgb(${Math.round(255*Math.max(0,(t-.25)/.75))},${Math.round(30+220*Math.sqrt(t))},${Math.round(65+150*(1-t))})`;ctx.fillRect(i*canvas.width/W,(H-1-j)*canvas.height/H,Math.ceil(canvas.width/W),Math.ceil(canvas.height/H));}
}
function renderFields(){
 const m=latest?.result?.modes[read('vb-mode-select')||0],mode=$('vb-field-type').value;
 for(const key of ['Itotal','Ex','Ey','Ez']){
  const map=key==='Itotal'||mode==='intensity'?m?.field_map?.intensity_normalized_to_max_total?.[key]:m?.field_map?.field_components?.[key+'_'+mode];mapCanvas($('vb-map-'+key),map,key!=='Itotal'&&mode!=='intensity');
 }
 overlay();
}
function drawStudyChart(){
 const can=$('vb-study-chart'),ctx=can.getContext('2d'),W=can.width,H=can.height;ctx.fillStyle='#102c40';ctx.fillRect(0,0,W,H);
 const axes=latest?.axes??[],key=axes[0]?.parameter??(latest?.rows.some(r=>r.R!=null)?'R':'job'),loss=key==='R'||key.startsWith('pml'),ykey=loss?'loss_dB_km':'neff',groups=new Map();
 for(const r of latest?.rows??[]){if(!Number.isFinite(r[key])||!Number.isFinite(r[ykey]))continue;const group=axes[1]?String(r[axes[1].parameter]):'current';if(!groups.has(group))groups.set(group,[]);groups.get(group).push(r);}
 const curves=[...groups].map(([label,rows])=>({label:'current '+label,rows}));
 for(const h of history)if(h.axes?.[0]?.parameter===axes[0]?.parameter&&h.rows)curves.push({label:h.label,rows:h.rows.filter(r=>Number.isFinite(r[key])&&Number.isFinite(r[ykey]))});
 const all=curves.flatMap(c=>c.rows);if(!all.length)return;
 let xmin=Math.min(...all.map(r=>r[key])),xmax=Math.max(...all.map(r=>r[key])),ymin=Math.min(...all.map(r=>r[ykey])),ymax=Math.max(...all.map(r=>r[ykey]));if(xmax===xmin)xmax=xmin+1;if(ymax===ymin)ymax=ymin+Math.max(1e-8,Math.abs(ymin)*1e-5);
 const margin=.05*(ymax-ymin);ymin-=margin;ymax+=margin;
 const xx=v=>86+(v-xmin)/(xmax-xmin)*(W-110),yy=v=>35+(ymax-v)/(ymax-ymin)*(H-85);
 ctx.font='12px sans-serif';ctx.fillStyle='#c2dae7';ctx.fillText(key+' → '+ykey+(loss?' (raw total loss; convergence unverified)':''),16,20);
 for(let i=0;i<5;i++){const y=35+i*(H-85)/4,v=ymax-i*(ymax-ymin)/4;ctx.strokeStyle='#294b61';ctx.beginPath();ctx.moveTo(86,y);ctx.lineTo(W-24,y);ctx.stroke();ctx.fillStyle='#c2dae7';ctx.fillText(v.toPrecision(7),3,y+4);const x=86+i*(W-110)/4;ctx.fillText((xmin+i*(xmax-xmin)/4).toPrecision(4),x-12,H-22);}
 for(let i=0;i<curves.length;i++){const curve=curves[i];ctx.strokeStyle=['#60e6c2','#ffb66d','#94c5ff','#e99fff'][i%4];ctx.fillStyle=ctx.strokeStyle;ctx.lineWidth=2;ctx.beginPath();curve.rows.slice().sort((a,b)=>a[key]-b[key]).forEach((r,j)=>{if(j)ctx.lineTo(xx(r[key]),yy(r[ykey]));else ctx.moveTo(xx(r[key]),yy(r[ykey]));});ctx.stroke();ctx.fillText(curve.label,90+i*170,H-5);}
 $('vb-curve-state').textContent=`곡선 ${curves.length}개 · 저장 결과와 같은 축에서 오버레이. 모드 1 index 추적; 모드 교차는 필드 확인 필요.`;
}
function renderReference(){
 if(!latest?.result)return;
 const check=compareReference(latest.result,reference);$('vb-reference-state').textContent=check.status;
 table('vb-reference-table',check.rows,[['mode','Mode index'],['neff','Browser neff'],['referenceNeff','Colab/저장 neff'],['deltaNeff','차이'],['deltaImag','Im 차이'],['status','대응 기준']]);overlay();
}
function sampleBilinear(map,xs,ys,x,y){
 if(!map?.length||!xs?.length||!ys?.length)return null;
 const dx=xs[1]-xs[0],dy=ys[1]-ys[0];if(!(dx>0&&dy>0))return null;
 const u=(x-xs[0])/dx,v=(y-ys[0])/dy;if(u<0||v<0||u>xs.length-1||v>ys.length-1)return null;
 const i=Math.min(xs.length-2,Math.floor(u)),j=Math.min(ys.length-2,Math.floor(v)),a=u-i,b=v-j;
 if(![map[j]?.[i],map[j]?.[i+1],map[j+1]?.[i],map[j+1]?.[i+1]].every(Number.isFinite))return null;
 return map[j][i]*(1-a)*(1-b)+map[j][i+1]*a*(1-b)+map[j+1][i]*(1-a)*b+map[j+1][i+1]*a*b;
}
function overlay(){
 const can=$('vb-overlay'),ctx=can.getContext('2d');ctx.clearRect(0,0,can.width,can.height);
 const result=latest?.result;if(!result||!reference)return;
 const check=compareReference(result,reference);if(!check.rows.length){$('vb-overlay-state').textContent=check.status;return;}
 const i=read('vb-mode-select')||0,a=result.modes[i]?.field_map,b=reference.modes[i]?.field_map,key=$('vb-overlay-field').value,alpha=read('vb-overlay-opacity')/100;
 const am=a?.intensity_normalized_to_max_total?.[key],bm=b?.intensity_normalized_to_max_total?.[key];
 if(!am||!bm){$('vb-overlay-state').textContent='미검증: 해당 모드의 비교 필드 맵 없음';return;}
 const xs=a.x_um,ys=a.y_um??a.x_um,bx=b.x_um,by=b.y_um??b.x_um,blend=[];let missing=0;
 for(let j=0;j<am.length;j++){const row=[];for(let k=0;k<am[j].length;k++){const v=sampleBilinear(bm,bx,by,xs[k],ys[j]);if(v==null)missing++;row.push(v==null?am[j][k]:am[j][k]*(1-alpha)+v*alpha);}blend.push(row);}
 mapCanvas(can,blend);$('vb-overlay-state').textContent=`동일 물리 조건 · ${key} 좌표 기반 bilinear 재표본화 혼합 · 격자 ${a.grid_size} vs ${b.grid_size} · 비중첩 ${missing}셀. 전역 위상 비교는 강도맵에 영향 없음. 편광/모드 대응은 직접 확인하세요.`;
}
function renderMCF(){
 const r=latest?.result;if(!r)return;const m=twoCoreMetrics(r,{length_mm:read('vb-mcf-length')});
 if(r.input.geometry.kind==='mcf'&&r.input.geometry.count===2){table('vb-mcf-table',[m],[['delta_neff','직접 Δneff'],['legacyDelta','기존 보간 Δneff'],['deltaToLegacy','차이'],['kappa_per_m','κ /m'],['coupling_length_mm','Lc mm'],['xt_db','XT dB'],['legacyStatus','참조 비교'],['status','모델']]);}
 else if(r.input.geometry.kind==='mcf'){
  table('vb-mcf-table',r.modes.map((m,i)=>({mode:i+1,neff:m.neff,polarization:m.polarization,delta:i?m.neff-r.modes[0].neff:null,beat:i?r.input.wavelength_nm*1e-6/Math.abs(m.neff-r.modes[0].neff):null})),[['mode','Supermode'],['neff','neff'],['polarization','편광'],['delta','Δneff vs mode1'],['beat','Beat length mm']]);
 }else $('vb-mcf-table').textContent='MCF 직접 해석을 선택하면 결과를 표시합니다.';
 if(r.input.geometry.kind==='mcf'){
  const n=nCoreMetrics(r,{length_mm:read('vb-mcf-length')}),p=document.createElement('p');p.textContent=n.status;$('vb-mcf-table').append(p);
  if(n.coupling){const target=document.createElement('div');target.id='vb-ncore-coupling';$('vb-mcf-table').append(target);table('vb-ncore-coupling',n.coupling,[['fromCore','입력 코어'],['toCore','출력 코어'],['kappaMagnitude_per_m','|Hᵢⱼ| /m']]);
   const t=document.createElement('div');t.id='vb-ncore-transfer';$('vb-mcf-table').append(t);table('vb-ncore-transfer',n.transfer,[['core','Core'],['normalizedProjectedFraction','Projected intensity fraction'],['xt_dB','XT relative to input dB']]);}
 }

}
function geometryPreview(){
 const can=$('vb-geometry-canvas');if(!can)return;
 try{
  const c=readConfig(),s=sampler(c),N=120,L=c.window_halfwidth_um,grid=[];
  for(let j=0;j<N;j++){const row=[];for(let i=0;i<N;i++)row.push(s.at(-L+(i+.5)*2*L/N,-L+(j+.5)*2*L/N).n);grid.push(row);}
  // Show index contrast, not absolute n, so weakly guiding cores are visible.
  const lo=Math.min(...grid.flat());mapCanvas(can,grid.map(r=>r.map(v=>v-lo+1e-8)));
 }catch{const ctx=can.getContext('2d');ctx.fillStyle='#0a2338';ctx.fillRect(0,0,can.width,can.height);}
 if(polygonPoints.length){
  const L=read('vb-half'),ctx=can.getContext('2d');ctx.strokeStyle='#ffb15b';ctx.fillStyle='#ffb15b';ctx.lineWidth=2;ctx.beginPath();polygonPoints.forEach(([x,y],i)=>{const px=(x/L+1)*can.width/2,py=(1-y/L)*can.height/2;if(i)ctx.lineTo(px,py);else ctx.moveTo(px,py);});ctx.stroke();
  for(const [x,y]of polygonPoints){ctx.beginPath();ctx.arc((x/L+1)*can.width/2,(1-y/L)*can.height/2,3,0,2*Math.PI);ctx.fill();}
 }
}
function parsePoints(){
 const rows=$('vb-points').value.trim().split(/\n/).filter(Boolean).map(r=>r.trim().split(/[,\s]+/).map(Number));
 if(rows.some(r=>r.length!==2||!r.every(Number.isFinite)))throw Error('Each vertex requires finite x,y');return rows;
}
function updateRegions(){
 const rows=geometry.regions??[];$('vb-region-list').innerHTML='<option value="-1">새 영역</option>'+rows.map((r,i)=>`<option value="${i}">영역 ${i+1}: ${esc(r.material)} · ${r.kind??'polygon'}</option>`).join('');
 $('vb-regions-json').value=JSON.stringify(rows,null,2);
}
function updatePalette(){
 const el=$('vb-palette');el.replaceChildren();if(geometry.kind!=='mask')return;
 const ids=materialIDs();
 for(const [key,value]of Object.entries(geometry.palette??{})){
  const row=document.createElement('div');row.className='vb-palette';const chip=document.createElement('span');chip.className='vb-chip';chip.style.background=geometry.colors?.[key]??(Number(key)?'#ffffff':'#000000');const label=document.createElement('span');label.textContent='Color '+key;const select=document.createElement('select');
  for(const id of ids){const option=document.createElement('option');option.value=id;option.textContent=id;select.append(option);}select.value=value;
  select.addEventListener('change',()=>{geometry.palette[key]=select.value;invalidate();geometryPreview();});row.append(chip,label,select);el.append(row);
 }
}
function materialIDs(){try{return ['core','clad','air',...Object.keys(JSON.parse($('vb-materials-json').value))].filter((v,i,a)=>a.indexOf(v)===i);}catch{return ['core','clad','air'];}}
function sanitizedSVG(text){
 const doc=new DOMParser().parseFromString(text,'image/svg+xml');if(doc.querySelector('parsererror')||doc.documentElement.localName!=='svg')throw Error('Invalid SVG');
 const allowed=new Set(['svg','g','path','polygon','polyline','circle','ellipse','rect','line','defs','clipPath','title','desc']);
 for(const e of [...doc.querySelectorAll('*')]){
  if(!allowed.has(e.localName))throw Error('Unsupported SVG element: '+e.localName+'; export plain paths first');
  for(const attr of [...e.attributes])if(/^on/i.test(attr.name)||/href/i.test(attr.name)||/url\s*\(|@import|javascript:/i.test(attr.value))throw Error('External/active SVG references are unsupported');
 }
 doc.documentElement.setAttribute('width','512');doc.documentElement.setAttribute('height','512');
 return new XMLSerializer().serializeToString(doc);
}
async function importMask(file){
 if(!file||file.size>10e6)throw Error('Image maximum 10MB');
 const isSVG=file.type==='image/svg+xml'||file.name.toLowerCase().endsWith('.svg'),blob=isSVG?new Blob([sanitizedSVG(await file.text())],{type:'image/svg+xml'}):file;
 const url=URL.createObjectURL(blob),img=new Image();
 try{await new Promise((resolve,reject)=>{img.onload=resolve;img.onerror=()=>reject(Error('Image decoding failed'));img.src=url;});
  const scale=Math.min(1,512/img.naturalWidth,512/img.naturalHeight),w=Math.max(1,Math.round(img.naturalWidth*scale)),h=Math.max(1,Math.round(img.naturalHeight*scale)),can=document.createElement('canvas');can.width=w;can.height=h;const ctx=can.getContext('2d');ctx.fillStyle='white';ctx.fillRect(0,0,w,h);ctx.drawImage(img,0,0,w,h);imageSource={width:w,height:h,data:ctx.getImageData(0,0,w,h).data,name:file.name};rebuildMask();
 }finally{URL.revokeObjectURL(url);}
}
function rebuildMask(){
 if(!imageSource)throw Error('Select image first');const {width,height,data}=imageSource,labels=new Array(width*height),palette={},colors={};
 if($('vb-mask-mode').value==='bw'){
  const t=read('vb-threshold');if(t<0||t>255)throw Error('Threshold 0–255');
  for(let i=0;i<labels.length;i++)labels[i]=.2126*data[4*i]+.7152*data[4*i+1]+.0722*data[4*i+2]<t?0:1;
  Object.assign(palette,{0:'core',1:'clad'});Object.assign(colors,{0:'#000000',1:'#ffffff'});
 }else{
  // Deterministic RGB histogram and farthest-first palette; no pixel values interpreted as refractive index.
  const histogram=new Map();for(let i=0;i<labels.length;i++){const rgb=[0,1,2].map(k=>Math.round(data[4*i+k]/32)*32),key=rgb.join(',');histogram.set(key,{rgb,count:(histogram.get(key)?.count??0)+1});}
  const bins=[...histogram.values()].sort((a,b)=>b.count-a.count),selected=[bins[0].rgb],distance=(a,b)=>a.reduce((s,v,i)=>s+(v-b[i])**2,0);
  while(selected.length<8&&selected.length<bins.length){let best=null,score=-1;for(const bin of bins){const d=Math.min(...selected.map(c=>distance(bin.rgb,c)))*Math.sqrt(bin.count);if(d>score){score=d;best=bin.rgb;}}if(score<=0)break;selected.push(best);}
  selected.forEach((rgb,i)=>{palette[i]=i===0?'clad':'core';colors[i]='rgb('+rgb.map(v=>Math.min(255,v)).join(',')+')';});
  for(let i=0;i<labels.length;i++){const rgb=[data[4*i],data[4*i+1],data[4*i+2]];let best=0,d=Infinity;selected.forEach((v,j)=>{const s=distance(v,rgb);if(s<d){d=s;best=j;}});labels[i]=best;}
 }
 geometry={kind:'mask',width,height,labels,palette,colors,halfwidth_um:read('vb-mask-half'),halfheight_um:read('vb-mask-half')*height/width,sourceName:imageSource.name};
 if(!(geometry.halfwidth_um>0))throw Error('Set positive physical image halfwidth');
 $('vb-kind').value=$('vb-kind-copy').value='mask';$('vb-mask-state').textContent=`${imageSource.name} · ${width}×${height} · ${Object.keys(palette).length}색 · 전체 폭 ${geometry.halfwidth_um*2}µm; assign each color to material.`;
 updatePalette();invalidate();geometryPreview();
}
function invalidate(){
 if(latest)$('vb-status').textContent='입력이 변경됨 · 표시 결과는 이전 실행 입력 기준. 다시 계산하세요.';
 estimate();
}
const stateServices={parameters:'LSWaveParameterService',cad:'LSWaveCADService',materials:'LSWaveMaterialsEditor',pml:'LSWavePMLService',m2:'LSWaveM2Preview',scalar:'LSWaveScalarProject'};
function uiSnapshot(){
 const out={};for(const el of document.querySelectorAll('input[id],select[id],textarea[id]')){
  if(el.type==='file'||el.closest('#vwb')||el.closest('#ls-cad-pane')||el.closest('#ls-config-pane')||el.closest('#ls-materials-editor')||el.closest('#ls-pml-editor'))continue;
  out[el.id]=el.type==='checkbox'?{checked:el.checked}:{value:el.value};
 }return out;
}
function projectSnapshot(){
 const services={};for(const [key,name]of Object.entries(stateServices))if(globalThis[name]?.snapshot)services[key]=globalThis[name].snapshot();
 return {schema:'LSCNS_wave_project',schemaVersion:SCHEMA,generatedAt:new Date().toISOString(),ui:uiSnapshot(),services,
  vector:{config:readConfig(),history:structuredClone(history),latest:structuredClone(latest),reference:structuredClone(reference),activeTab,
   sweep:[1,2].map(i=>({parameter:$('vb-param'+i).value,start:read('vb-start'+i),end:read('vb-end'+i),count:read('vb-count'+i)})),axis2:$('vb-axis2').checked,length_mm:read('vb-mcf-length')},
  layout:{legacy:document.body.classList.contains('vb-single-page')},warnings:['Solver results retain input snapshots; loading a project never recomputes automatically']};
}
function restoreUI(ui){
 // Restore preset/profile before other fields because their existing listeners write dependent inputs.
 const ordered=['preset','profile',...Object.keys(ui).filter(k=>k!=='preset'&&k!=='profile')];
 for(const id of ordered){const el=$(id),v=ui[id];if(!el||!v||!['INPUT','SELECT','TEXTAREA'].includes(el.tagName)||el.type==='file')continue;if(Object.hasOwn(v,'checked'))el.checked=!!v.checked;else el.value=String(v.value??'');if(id!=='preset'&&id!=='profile'){el.dispatchEvent(new Event('change',{bubbles:true}));el.dispatchEvent(new Event('input',{bubbles:true}));}}
}
function loadProject(doc){
 if(busy)throw Error('Cancel active study before loading project');
 const {project:p,warnings}=migrateProject(doc),backup=projectSnapshot();
 function apply(project){
  if(project.legacyModel){
   const m=project.legacyModel,ui={};for(const [k,v]of Object.entries(m))if($(k))ui[k]={value:k==='shells'?v.map(s=>s.join(',')).join('\n'):k==='n2_m2_per_W'?v/1e-20:v};
   ui.clad_material={value:m.clad_index_override==null?'silica':'fixed'};restoreUI(ui);globalThis.LSWaveScalarProject?.loadProject({model:m,last:null});
  }else if(project.ui)restoreUI(project.ui);
  for(const [key,name]of Object.entries(stateServices))if(project.services?.[key])globalThis[name]?.loadProject?.(project.services[key]);
  writeConfig(project.vector.config);history=structuredClone(project.vector.history??[]);latest=structuredClone(project.vector.latest??null);reference=structuredClone(project.vector.reference??null);
  for(const [j,a]of (project.vector.sweep??[]).entries()){const i=j+1;if(i>2)break;$('vb-param'+i).value=a.parameter;for(const k of ['start','end','count'])$('vb-'+k+i).value=a[k];}
  $('vb-axis2').checked=!!project.vector.axis2;$('vb-mcf-length').value=project.vector.length_mm??1;
  renderHistory();setBusy(false);if(latest)render();setLayout(!!project.layout?.legacy,false);show(project.vector.activeTab??'m6');
 }
 try{apply(p);}catch(error){try{apply(backup);}catch{}throw error;}
 $('vb-project-state').textContent='프로젝트 불러오기 완료 · '+(warnings.join(' · ')||'schema v3; inputs and retained results restored');
 return warnings;
}
function renderHistory(){
 $('vb-history-select').innerHTML='<option value="">선택 없음</option>'+history.map((h,i)=>`<option value="${i}">${esc(h.label)}</option>`).join('');
 $('vb-history-state').textContent='저장 결과 '+history.length+'개 · 결과별 입력 스냅샷 보존';
}
function saveHistory(){
 if(!latest?.result)throw Error('No completed result');
 const h={label:new Date().toLocaleString('ko-KR',{timeZone:'Asia/Seoul'}),result:structuredClone(latest.result),rows:structuredClone(latest.rows),axes:structuredClone(latest.axes)};
 const candidate=[...history,h].slice(-5);if(JSON.stringify(candidate).length>30e6)throw Error('Saved field maps exceed 30 MB; clear earlier results or lower mesh/modes');history=candidate;renderHistory();drawStudyChart();
}
function setLayout(legacy,persist=true){
 document.body.classList.toggle('vb-legacy',legacy);document.body.classList.toggle('vb-single-page',legacy);document.body.classList.toggle('ls-comsol-shell',!legacy);
 $('vb-toggle').textContent=legacy?'신규 Model tree UI':'기존 단일 페이지';
 if(persist)try{localStorage.setItem('LSWave.workflow.legacy',String(legacy));}catch{}
}
function mountTree(){
 const tree=document.querySelector('.ls-model-tree'),root=document.createElement('div');root.id='vb-tree';root.innerHTML='<strong>M3–M6 · Model workflow</strong>'+[['geometry','1 · Geometry'],['materials','2 · Materials'],['physics','3 · Physics · Mode analysis'],['mesh','4 · Mesh'],['study','5 · Study'],['results','6 · Results']].map(([id,label])=>`<button data-vb-node="${id}">${label}</button>`).join('');tree?.prepend(root);
 root.addEventListener('click',e=>{
  const id=e.target.closest('[data-vb-node]')?.dataset.vbNode;if(!id)return;
  if(id==='geometry'){document.dispatchEvent(new CustomEvent('ls-open-cad',{detail:{id:'root'}}));show('m5');}
  if(id==='materials'){document.dispatchEvent(new CustomEvent('ls-open-wave-optics',{detail:{panel:'waveMaterials'}}));globalThis.LSWaveM2Preview?.show();}
  if(id==='physics'){document.dispatchEvent(new CustomEvent('ls-open-wave-optics',{detail:{panel:'waveMode'}}));show('m3');}
  if(id==='mesh'){document.dispatchEvent(new CustomEvent('ls-open-wave-optics',{detail:{panel:'waveSize'}}));show('m3');$('vb-grid').focus();}
  if(id==='study')show('m6');if(id==='results')show('results');
 });
 const open=document.createElement('button');open.id='vb-open';open.textContent='M3–M6';document.querySelector('.topbar')?.append(open);open.addEventListener('click',()=>{if(document.body.classList.contains('vb-single-page'))setLayout(false);show('m3');});
 const toggle=document.createElement('button');toggle.id='vb-toggle';toggle.type='button';toggle.textContent='기존 단일 페이지';document.querySelector('.topbar')?.append(toggle);toggle.addEventListener('click',()=>setLayout(!document.body.classList.contains('vb-single-page')));
 let legacy=new URL(location.href).searchParams.get('workflow')==='legacy';try{legacy=legacy||localStorage.getItem('LSWave.workflow.legacy')==='true';}catch{}setLayout(legacy,false);
}
function events(){
 const safe=fn=>async e=>{try{await fn(e);}catch(error){$('vb-status').textContent='오류: '+error.message;}};
 for(const el of document.querySelectorAll('[data-vb-tab]'))el.addEventListener('click',()=>show(el.dataset.vbTab));
 $('vwb').addEventListener('change',e=>{if(e.target.matches('input,select,textarea'))invalidate();});
 for(const id of ['vb-kind','vb-kind-copy'])$(id).addEventListener('change',e=>{$('vb-kind').value=$('vb-kind-copy').value=e.target.value;geometryPreview();});
 $('vb-run').addEventListener('click',()=>run());$('vb-cancel').addEventListener('click',cancel);
 $('vb-sweep-run').addEventListener('click',()=>{const axes=[1,...($('vb-axis2').checked?[2]:[])].map(i=>({parameter:$('vb-param'+i).value,start:read('vb-start'+i),end:read('vb-end'+i),count:read('vb-count'+i)}));run(axes);});
 $('vb-grid-study').addEventListener('click',()=>{const n=read('vb-grid'),start=Math.max(24,Math.min(144,n));run([{parameter:'grid',start,end:Math.min(192,start+48),count:3}]);});
 $('vb-pml-study').addEventListener('click',safe(()=>{$('vb-pml').checked=true;const c=readConfig(),h=2*c.window_halfwidth_um/c.grid_n,start=Math.max(3*h,c.pml.thickness_um*.8),end=Math.min(c.window_halfwidth_um*.55,c.pml.thickness_um*1.2);if(end<=start)throw Error('Increase grid/window to vary PML thickness');run([{parameter:'pml_thickness',start,end,count:3},{parameter:'pml_strength',start:1,end:5,count:3}]);}));
 $('vb-bend-study').addEventListener('click',safe(()=>{
  $('vb-pml').checked=true;const c=readConfig(),R=c.bend_radius_um||10000,start=Math.max(10*c.window_halfwidth_um,R/2),jobs=[0,start,R,R*2].map(v=>({config:{...c,bend_radius_um:v},parameters:{R:v}}));run([],jobs);
 }));
 $('vb-subpixel-study').addEventListener('click',safe(()=>{const c=readConfig();run([],[false,true].map(v=>({config:{...c,subpixel:v},parameters:{variant:'subpixel '+v}})));}));
 $('vb-abs-study').addEventListener('click',safe(()=>{const c=readConfig(),zero=structuredClone(c);zero.k_core=zero.k_clad=0;if(zero.materials)for(const m of Object.values(zero.materials)){m.kappa=0;m.absorption={kind:'k',value:0};if(m.table)m.table=m.table.map(row=>({...row,kappa:0}));}run([],[{config:zero,parameters:{variant:'kappa=0'}},{config:c,parameters:{variant:'current kappa'}}]);}));
 $('vb-result-json').addEventListener('click',()=>{if(latest?.result)save('LS_Wave_Yee_Colab_compatible.json',latest.result);});$('vb-report-json').addEventListener('click',()=>save('LS_Wave_M3_M6_study.json',latest));$('vb-csv').addEventListener('click',()=>save('LS_Wave_sweep.csv',rowsCSV(latest.rows),'text/csv;charset=utf-8'));
 $('vb-project-save').addEventListener('click',safe(()=>save('LS_Wave_project_v3.json',projectSnapshot())));
 $('vb-project-paste-load').addEventListener('click',safe(()=>loadProject(JSON.parse($('vb-project-paste').value))));
 $('vb-project-open').addEventListener('change',safe(async e=>{const f=e.target.files?.[0];if(!f)return;if(f.size>60e6)throw Error('Project maximum 60MB');loadProject(JSON.parse(await f.text()));e.target.value='';}));
 $('vb-history-save').addEventListener('click',safe(saveHistory));$('vb-history-clear').addEventListener('click',()=>{history=[];renderHistory();drawStudyChart();});$('vb-history-select').addEventListener('change',()=>{const raw=$('vb-history-select').value;if(raw===''){reference=null;}else reference=history[Number(raw)]?.result??null;renderReference();});
 $('vb-open-m2').addEventListener('click',()=>globalThis.LSWaveM2Preview?.show());
 $('vb-from-legacy').addEventListener('click',safe(()=>{
  const num=id=>Number($(id)?.value),n=$('clad_material')?.value==='fixed'?num('clad_index_override'):silicaIndex(num('wavelength_nm'));
  const c=config({wavelength_nm:num('wavelength_nm'),core_radius_um:num('core_radius_um'),n_core:n+num('delta_n_core'),n_clad:n,window_halfwidth_um:Math.max(18,num('core_radius_um')*3),profile:$('profile').value,shells:$('shells').value.split(/\n/).filter(Boolean).map(r=>r.split(',').map(Number)),trench_start_um:num('trench_start_um'),trench_end_um:num('trench_end_um'),trench_delta_n:num('trench_delta_n'),graded_alpha:num('graded_alpha')});writeConfig(c);invalidate();
 }));
 $('vb-python-case').addEventListener('click',()=>{writeConfig(config({wavelength_nm:1550,core_radius_um:5.5,n_core:1.4475022010878765,n_clad:1.444023621703261,window_halfwidth_um:22,grid_n:68}));invalidate();});
 const loadPython=async()=>{const response=await fetch(new URL('../full_vector/results/vector_mode_result.json',import.meta.url));if(!response.ok)throw Error('Reference HTTP '+response.status);const ref=normalizeReference(await response.json());ref.input.n_clad=1.444023621703261;ref.provenance='repository Python collocated H-field FDFD fixture; not Yee same-grid oracle';reference=ref;renderReference();$('vb-reference-state').textContent+=' · 저장된 Python collocated 방식; 동일 이산화 아님';};
 $('vb-load-python').addEventListener('click',safe(loadPython));
 $('vb-load-hcf-reference').addEventListener('click',safe(async()=>{const r=await fetch(new URL('../hcf_complex_fem/results/default_convergence.json',import.meta.url));if(!r.ok)throw Error('HCF reference HTTP '+r.status);reference=normalizeReference(await r.json());renderReference();}));$('vb-reference-file').addEventListener('change',safe(async e=>{const f=e.target.files?.[0];if(!f)return;if(f.size>40e6)throw Error('Reference maximum 40MB');reference=normalizeReference(JSON.parse(await f.text()));renderReference();e.target.value='';}));
 for(const id of ['vb-mode-select','vb-field-type'])$(id).addEventListener('change',renderFields);
 for(const id of ['vb-overlay-field','vb-overlay-opacity'])$(id).addEventListener('input',overlay);
 $('vb-from-hcf').addEventListener('click',safe(()=>{
  const hcf={...HCF_DEFAULT};for(const key of Object.keys(hcf)){const el=$('hcf-'+key);if(el)hcf[key]=el.type==='checkbox'?el.checked:Number(el.value);}useHCF(hcf);
 }));
 $('vb-hcf-file').addEventListener('change',safe(async e=>{const f=e.target.files?.[0];if(!f)return;if(f.size>1e6)throw Error('HCF design maximum 1MB');const d=JSON.parse(await f.text());useHCF(d.input??d);e.target.value='';}));
 $('vb-mask-file').addEventListener('change',safe(async e=>{await importMask(e.target.files?.[0]);e.target.value='';}));$('vb-rebuild-mask').addEventListener('click',safe(rebuildMask));
 $('vb-geometry-canvas').addEventListener('pointerdown',e=>{const rect=e.currentTarget.getBoundingClientRect(),L=read('vb-half'),x=((e.clientX-rect.left)/rect.width*2-1)*L,y=(1-(e.clientY-rect.top)/rect.height*2)*L;polygonPoints.push([Number(x.toFixed(4)),Number(y.toFixed(4))]);$('vb-points').value=polygonPoints.map(p=>p.join(',')).join('\n');geometryPreview();});
 $('vb-points').addEventListener('input',()=>{try{polygonPoints=parsePoints();geometryPreview();}catch{}});
 $('vb-clear-points').addEventListener('click',()=>{polygonPoints=[];$('vb-points').value='';geometryPreview();});
 $('vb-add-region').addEventListener('click',safe(()=>{const points=parsePoints();if(points.length<3)throw Error('Polygon ≥3 vertices');const i=read('vb-region-list'),regions=geometry.regions??[],r={kind:'polygon',points,material:$('vb-region-material').value};if(i>=0)regions[i]=r;else regions.push(r);geometry={kind:'regions',regions};$('vb-kind').value=$('vb-kind-copy').value='regions';updateRegions();invalidate();geometryPreview();}));
 $('vb-region-list').addEventListener('change',()=>{const r=geometry.regions?.[read('vb-region-list')];if(!r)return;polygonPoints=structuredClone(r.points??[]);$('vb-points').value=polygonPoints.map(p=>p.join(',')).join('\n');$('vb-region-material').value=r.material;geometryPreview();});
 $('vb-delete-region').addEventListener('click',()=>{const i=read('vb-region-list');if(i>=0)geometry.regions?.splice(i,1);updateRegions();invalidate();geometryPreview();});
 $('vb-circle-region').addEventListener('click',()=>{geometry={kind:'regions',regions:[{kind:'circle',x:0,y:0,radius:read('vb-radius'),material:'core'}]};$('vb-kind').value=$('vb-kind-copy').value='regions';updateRegions();invalidate();geometryPreview();});
 $('vb-apply-regions').addEventListener('click',safe(()=>{const regions=JSON.parse($('vb-regions-json').value);if(!Array.isArray(regions))throw Error('Region array required');const c=readConfig();sampler({...c,geometry:{kind:'regions',regions}});geometry={kind:'regions',regions};$('vb-kind').value=$('vb-kind-copy').value='regions';updateRegions();invalidate();geometryPreview();}));
 $('vb-materials-apply').addEventListener('click',safe(()=>{const materials=JSON.parse($('vb-materials-json').value);for(const m of Object.values(materials))evaluateMaterial(m,read('vb-wavelength'));$('vb-material-source').value='json';$('vb-region-material').innerHTML=materialIDs().map(id=>`<option>${esc(id)}</option>`).join('');updatePalette();invalidate();geometryPreview();}));
 $('vb-mcf-length').addEventListener('change',renderMCF);
}
function useHCF(hcf){
 const c=readConfig(),radius=hcf.coreDiameterUm/2,extent=radius+(hcf.coreDiameterUm*hcf.outerTubeRatio)+4;
 c.geometry={kind:'hcf',hcf};c.core_radius_um=radius;c.window_halfwidth_um=Math.max(54,extent*1.2);c.pml={enabled:true,thickness_um:8,strength:3,order:3};c.target_neff=hcf.airIndex-.0007;c.bend_radius_um=0;writeConfig(c);$('vb-hcf-state').textContent='기존 HCF/NANF 입력을 재사용했습니다. 실제 튜브 extent와 PML 간격을 계산 시 검증합니다.';invalidate();
}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init,{once:true});else init();
