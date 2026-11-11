/**
 * M3 v0.7 optional full-vector FDFD study panel.
 * Feature flag: window.LSWaveFeatures.fullVectorStudyV07=false.
 * Existing Colab JSON viewer and old M3 panel are untouched.
 */
import {M3_STUDY_VERSION,M3_PYTHON_REF,validateM3Study,
 estimatedTypedArrayMiB,normalizeColabJSON,referenceRunMatch} from "./m3_study_metrics_v07.mjs";
const ENABLED=globalThis.LSWaveFeatures?.fullVectorStudyV07!==false;
const $=id=>document.getElementById(id);
const esc=v=>String(v??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const num=(n,d=10)=>Number.isFinite(n)?Number(n).toPrecision(d):"—";
const exp=n=>Number.isFinite(n)?Number(n).toExponential(3):"미검증";
const measuredV8Ms=Object.freeze({29:1276,41:4088,51:6611,69:14816,81:22835});
let worker=null,request=0,busy=false,last=null,colab=null,colabSource="없음",modeSelected=0;
function estimate(){
 const ns=checkedGrids(),target=ns.map(n=>measuredV8Ms[n]??Math.round(22835*(n/81)**3.2));
 const sum=target.reduce((a,b)=>a+b,0),memory=ns.length?Math.max(...ns.map(estimatedTypedArrayMiB)):0;
 const extra=ns.includes(91)?" (91 extrapolated)":"";
 $("m3v7-estimate").textContent="참고 V8 환경 시간 약 "+(sum/1000).toFixed(1)+"초"+extra+
  " · 최고 격자 Float64 배열 사용량 모델 약 "+memory.toFixed(2)+" MiB"+
  " (실측 RAM 아님; 브라우저/기기별 차이 큼)";
}
function checkedGrids(){
 return [...document.querySelectorAll("#m3v7-grids input:checked")].map(x=>Number(x.value)).sort((a,b)=>a-b);
}
function config(){
 return {profile:"step",wavelength_nm:Number($("m3v7-wave").value),
 core_radius_um:Number($("m3v7-radius").value),window_halfwidth_um:Number($("m3v7-half").value),
 n_core:Number($("m3v7-core").value),n_clad:Number($("m3v7-clad").value)};
}
function setBusy(on){
 busy=on;$("m3v7-run").disabled=on;$("m3v7-cancel").disabled=!on;
 $("m3v7-export").disabled=on||!last;$("m3v7-report").disabled=on||!last;
}
function download(name,data,type="application/json"){
 const blob=new Blob([JSON.stringify(data,null,2)],{type:type+";charset=utf-8"});
 const url=URL.createObjectURL(blob),a=document.createElement("a");
 a.href=url;a.download=name;document.body.appendChild(a);a.click();a.remove();
 setTimeout(()=>URL.revokeObjectURL(url),1000);
}
function drawField(canvas,map){
 const ctx=canvas.getContext("2d");if(!ctx||!Array.isArray(map)||!map.length)return;
 const n=map.length,size=canvas.width,max=Math.max(1e-20,...map.flat().map(Number));
 ctx.fillStyle="#08192b";ctx.fillRect(0,0,size,size);
 for(let j=0;j<n;j++)for(let i=0;i<n;i++){
  const v=Math.pow(Math.max(0,Math.min(1,Number(map[j][i])/max)),.56),px=size/n;
  ctx.fillStyle="rgb("+Math.round(13+230*v)+","+Math.round(32+160*v)+","+Math.round(73-36*v)+")";
  ctx.fillRect(i*px,j*px,px+1,px+1);
 }
}
function drawOverlay(canvas,computed,reference,opacity){
 const ctx=canvas.getContext("2d");if(!ctx)return;
 const n=computed.length,N=canvas.width,px=N/n;
 const ma=Math.max(1e-20,...computed.flat().map(Number)),mb=Math.max(1e-20,...reference.flat().map(Number));
 ctx.fillStyle="#061729";ctx.fillRect(0,0,N,N);
 for(let j=0;j<n;j++)for(let i=0;i<n;i++){
  const a=Math.pow(Math.max(0,Math.min(1,Number(computed[j][i])/ma)),.56),
   b=Math.pow(Math.max(0,Math.min(1,Number(reference[j][i])/mb)),.56);
  const v=(1-opacity)*a+opacity*b;
  // Computed = cyan, reference = red. Mixing slides between both.
  ctx.fillStyle="rgb("+Math.round(10+230*v*opacity)+","+
   Math.round(25+195*v*(1-opacity))+","+Math.round(48+190*v*(1-opacity))+")";
  ctx.fillRect(i*px,j*px,px+1,px+1);
 }
}
function displayMode(){
 if(!last)return;
 const result=last.portableVectorResult,mode=result?.modes[modeSelected];
 const field=mode?.field_map?.intensity_normalized_to_max_total;
 if(!field){$("m3v7-overlay-status").textContent="모드 필드 맵 없음";return}
 const keys=["Itotal","Ex","Ey","Ez"];
 for(const [i,key] of keys.entries())drawField($("m3v7-map-"+key),field[key]);
 const ref=colab?.modes[modeSelected]?.field_map?.intensity_normalized_to_max_total;
 const config0=last.input,grid=result?.input?.grid_n;
 if(ref&&referenceRunMatch(config0,grid,colab.input)&&
    Array.isArray(ref.Itotal)&&ref.Itotal.length===field.Itotal.length){
  const key=$("m3v7-overlay-key").value;
  drawOverlay($("m3v7-overlay"),field[key],ref[key],Number($("m3v7-opacity").value)/100);
  $("m3v7-overlay-status").textContent="동일 격자·입력: "+colabSource+
   " · "+key+" 색상 혼합 오버레이 ("+$("m3v7-opacity").value+"% Colab)";
 }else{
  const c=$("m3v7-overlay").getContext("2d");c.clearRect(0,0,300,300);
  c.fillStyle="#9eb8ca";c.font="13px sans-serif";c.fillText("비교 조건 불일치",70,155);
  $("m3v7-overlay-status").textContent="미검증: 동일 grid/재료/형상 Colab 맵이 없거나 선택 모드 맵이 없음.";
 }
}
function render(report){
 last=report;
 $("m3v7-summary").textContent="LP₀₁ 해석해 neff = "+num(report.scalarLP01.analyticNeff,13)+
  " · 기존 스칼라 FEM neff = "+num(report.scalarLP01.scalarFEMNeff,13)+
  " · Colab = "+(report.storedColabReference?num(report.storedColabReference.neff,13):"미검증");
 const rows=report.convergence;
 $("m3v7-convergence").innerHTML=rows.map(r=>'<tr>'+
  '<td>'+r.grid_n+'×'+r.grid_n+'</td><td>'+num(r.neff,13)+'</td>'+
  '<td>'+exp(r.residual)+'</td><td>'+exp(r.diffToLP01)+'</td>'+
  '<td>'+exp(r.diffToScalar)+'</td><td>'+exp(r.diffToPreviousGrid)+'</td>'+
  '<td>'+exp(r.deltaToColab)+'<small>'+esc(r.colabComparisonStatus)+'</small></td>'+
  '<td>'+r.wallMs.toFixed(1)+' ms</td>'+
  '<td>'+(r.workerHeapUsedBytes==null?"실측 미지원":(r.workerHeapUsedBytes/1048576).toFixed(2)+" MiB")+
  '<small>Float64 배열 모델 ≈ '+r.estimatedTypedArrayMiB.toFixed(2)+' MiB</small></td>'+
  '<td>'+esc(r.convergenceStatus)+'</td></tr>').join("");
 const modes=report.portableVectorResult.modes;
 $("m3v7-modes").innerHTML=modes.map((m,i)=>
  '<tr><td>'+[i+1,m.polarization].join(" · ")+'</td><td>'+num(m.neff,12)+'</td>'+
  '<td>'+exp(m.residual)+'</td><td>'+num(m.MFD_2nd_um,8)+'</td>'+
  '<td>'+num(m.Aeff_um2,8)+'</td><td>'+
  num(100*m.Ex_intensity_fraction,6)+' / '+num(100*m.Ey_intensity_fraction,6)+
  ' / '+num(100*m.Ez_intensity_fraction,6)+'%</td></tr>').join("");
 $("m3v7-mode").innerHTML=modes.map((m,i)=>'<option value="'+i+'">Mode '+(i+1)+
  ' · '+esc(m.polarization)+'</option>').join("");
 modeSelected=0;displayMode();
 $("m3v7-result-status").textContent=report.finalGridConvergenceConfirmed?
  "기준 통과: 마지막 두 격자 neff 차이 ≤ 1×10⁻⁶ (연속체 해 검증은 아님)":
  "미수렴: 마지막 두 격자 neff 차이 > 1×10⁻⁶. Staircase 경계 때문에 비단조 오차 발생 가능.";
}
async function loadDefault(){
 const url=new URL("../full_vector/results/vector_mode_result.json",import.meta.url);
 const response=await fetch(url,{cache:"no-store"});
 if(!response.ok)throw Error("HTTP "+response.status);
 colab=normalizeColabJSON(await response.json());
 colabSource="저장된 Python 69×69";$("m3v7-refstate").textContent="Colab 기준 로드됨: "+num(colab.modes[0].neff,13);
}
function init(){
 if(!ENABLED||$("m3v7-study"))return;
 const parent=$("full-vector-tool"),nav=document.querySelector("nav.eng-shortcuts");
 if(!parent||!nav)return;
 const tab=document.createElement("a");tab.href="#m3v7-study";tab.id="m3v7-nav";
 tab.textContent="M3 · VECTOR STUDY v0.7";nav.append(tab);
 const sec=document.createElement("section");sec.id="m3v7-study";sec.className="m3v7-study engineering-panel";sec.hidden=true;
 sec.innerHTML=[
 '<header><small>M3 · BROWSER FULL-VECTOR MODE STUDY · v'+M3_STUDY_VERSION+'</small>',
 '<h2>2D Hx/Hy FDFD · 격자 수렴 · Colab 비교</h2>',
 '<p>기존 Colab 결과 뷰어와 이전 M3는 유지됩니다. 이 새 해석은 <b>결합 Hx/Hy collocated FDFD</b>',
 '로 Yee staggered/edge FEM이 아니므로 물리적 정확도를 동일하게 주장하지 않습니다.</p></header>',
 '<div class="m3v7-config">',
 '<label>파장 (nm)<input id="m3v7-wave" type="number" value="1550" min="401" max="1699"></label>',
 '<label>Core radius (µm)<input id="m3v7-radius" type="number" value="5.5" min="0.5" step="0.1"></label>',
 '<label>n_core<input id="m3v7-core" type="number" value="1.4475022010878765" step="0.00001"></label>',
 '<label>n_clad<input id="m3v7-clad" type="number" value="1.444023621703261" step="0.00001"></label>',
 '<label>Calculation halfwidth (µm)<input id="m3v7-half" type="number" value="22" min="5" step="1"></label>',
 '</div>',
 '<fieldset id="m3v7-grids"><legend>FDFD 격자 선택 (odd, max 91×91; 최대 5개)</legend>',
 ...[29,41,51,69,81,91].map(n=>'<label><input type="checkbox" value="'+n+'" '+
   ([29,41,51,69].includes(n)?"checked":"")+'>'+n+'×'+n+'</label>'),
 '</fieldset>',
 '<p id="m3v7-estimate" class="m3v7-estimate">예상 자원 확인 중</p>',
 '<p id="m3v7-refstate">Colab 저장 결과는 계산 시 불러오며, 실패해도 계산은 계속합니다.</p>',
 '<div class="m3v7-actions">',
 '<button id="m3v7-run">▶ Full-vector 격자 스윕 실행</button>',
 '<button id="m3v7-cancel" disabled>■ 계산 취소</button>',
 '<button id="m3v7-export" disabled>기존 Colab 호환 JSON 저장</button>',
 '<button id="m3v7-report" disabled>수렴 보고서 JSON 저장</button>',
 '<label>비교 Colab JSON 선택<input id="m3v7-upload" type="file" accept=".json" hidden></label></div>',
 '<div class="m3v7-progress"><progress id="m3v7-progress" value="0" max="100"></progress><span id="m3v7-pct">0%</span></div>',
 '<p id="m3v7-status" role="status" aria-live="polite">대기 · Worker가 계산하며 취소는 즉시 종료됩니다.</p>',
 '<p id="m3v7-summary"></p>',
 '<h3>격자 수렴 · M1 해석해/스칼라 FEM/Colab neff 비교</h3>',
 '<div class="m3v7-scroll"><table><thead><tr><th>grid</th><th>Vector neff</th><th>잔차</th>',
 '<th>|Δ LP₀₁|</th><th>|Δ scalar|</th><th>|Δ 이전 격자|</th><th>|Δ Colab|</th>',
 '<th>실측 시간</th><th>메모리</th><th>수렴 판정</th></tr></thead>',
 '<tbody id="m3v7-convergence"><tr><td colspan="10">계산 대기</td></tr></tbody></table></div>',
 '<h3>편광별 Full-vector 출력</h3>',
 '<div class="m3v7-scroll"><table><thead><tr><th>모드</th><th>neff</th><th>residual</th>',
 '<th>MFD (µm)</th><th>Aeff (µm²)</th><th>Ex/Ey/Ez 분율</th></tr></thead>',
 '<tbody id="m3v7-modes"><tr><td colspan="6">계산 대기</td></tr></tbody></table></div>',
 '<div class="m3v7-controls"><label>모드<select id="m3v7-mode"><option value="0">Mode 1</option></select></label>',
 '<label>오버레이 성분<select id="m3v7-overlay-key"><option>Itotal</option><option>Ex</option><option>Ey</option><option>Ez</option></select></label>',
 '<label>Colab 오버레이 투명도<input id="m3v7-opacity" type="range" min="0" max="100" value="50"></label></div>',
 '<div class="m3v7-maps">',
 ...["Itotal","Ex","Ey","Ez"].map(k=>'<article><h4>|'+k+'|²</h4><canvas id="m3v7-map-'+k+'" width="240" height="240"></canvas></article>'),
 '<article><h4>계산 결과 ↔ Colab 오버레이</h4><canvas id="m3v7-overlay" width="300" height="300"></canvas></article></div>',
 '<p id="m3v7-overlay-status">동일 격자·형상의 Colab JSON에서만 오버레이를 활성화합니다.</p>',
 '<p id="m3v7-result-status"></p>',
 '<details open><summary>수치해석 가정과 미지원 항목</summary>',
 '<p>이 버전의 고유값 연산은 <b>Yee staggered 격자가 아닌 collocated Hx/Hy FDFD</b>입니다.',
 ' 최대 격자 91×91, 실수 ε, Dirichlet 사각경계, 계단형 코어; PML/HCF/흡수/복소 neff 미지원.</p>',
 '<p>고유값: matrix-free shift-invert, 대각 사전조건 BiCGSTAB 내측 상대잔차 2×10⁻⁹,',
 ' 역반복 고유값 잔차 1×10⁻⁷. 초기 모드는 x/y 방향 가우시안, 편광 직교화.</p>',
 '<p>시간은 선택한 브라우저 Worker의 <b>실측 performance.now()</b>이며,',
 ' 메모리는 Worker 성능 API가 노출될 때만 실측 수치를 표시합니다.',
 ' Float64 배열 모델은 추정치로 분리합니다. n_eff 격자 수렴은 자동 보장되지 않습니다.</p></details>'
 ].join("");
 parent.appendChild(sec);
 const show=()=>{sec.hidden=false;sec.scrollIntoView({behavior:"smooth",block:"start"});};
 tab.addEventListener("click",e=>{e.preventDefault();show()});
 $("m3v7-grids").addEventListener("change",estimate);
 for(const n of ["m3v7-wave","m3v7-radius","m3v7-core","m3v7-clad","m3v7-half"])
  $(n).addEventListener("input",estimate);
 $("m3v7-mode").addEventListener("change",()=>{modeSelected=Number($("m3v7-mode").value);displayMode()});
 $("m3v7-overlay-key").addEventListener("change",displayMode);
 $("m3v7-opacity").addEventListener("input",displayMode);
 $("m3v7-run").addEventListener("click",async()=>{
  if(busy)return;
  let cfg,grids;
  try{cfg=config();grids=checkedGrids();validateM3Study(cfg,grids)}
  catch(error){$("m3v7-status").textContent="입력 오류: "+error.message;return}
  const id=++request;last=null;setBusy(true);
  $("m3v7-convergence").innerHTML='<tr><td colspan="10">계산 진행 중…</td></tr>';
  $("m3v7-progress").value=0;$("m3v7-pct").textContent="0%";
  $("m3v7-status").textContent="Worker 초기화 중…";
  if(!colab){try{await loadDefault()}catch(e){
   $("m3v7-refstate").textContent="Colab 기준 불러오기 실패: "+e.message+" · Colab 차이는 미검증";}}
  if(!busy||id!==request)return; // cancelled during async reference load
  try{
   worker?.terminate();
   worker=new Worker(new URL("./m3_study_worker_v07.mjs?v=v0.7.0",import.meta.url),{type:"module"});
   worker.onmessage=e=>{
    if(e.data?.id!==id||id!==request||!busy)return;
    const p=e.data;
    if(p.op==="progress"){
     $("m3v7-progress").value=p.percent??0;
     $("m3v7-pct").textContent=(p.percent??0)+"% (추정)";
     $("m3v7-status").textContent=(p.grid_n||"")+"×"+(p.grid_n||"")+
      " · "+p.phase+" "+(p.mode||"")+" "+
      (p.residual!=null?"잔차 "+exp(p.residual):"");
    }else if(p.op==="error"){
     worker?.terminate();worker=null;setBusy(false);
     $("m3v7-status").textContent="FDFD 계산 실패: "+p.message;
    }else if(p.op==="result"){
     worker?.terminate();worker=null;last=p.report;setBusy(false);render(last);
     $("m3v7-progress").value=100;$("m3v7-pct").textContent="100%";
     $("m3v7-status").textContent="완료 · 실제 계산 시간 합계 "+
      (last.wallMsTotal/1000).toFixed(2)+" s · 보고서/필드 JSON 저장 가능";
    }
   };
   worker.onerror=e=>{
    if(id!==request)return;
    worker?.terminate();worker=null;setBusy(false);
    $("m3v7-status").textContent="Worker 예외: "+(e.message||"unknown");
   };
   worker.postMessage({id,op:"run",config:cfg,grids,colabReference:colab});
  }catch(error){worker?.terminate();worker=null;setBusy(false);
   $("m3v7-status").textContent="Worker 생성 실패: "+error.message}
 });
 $("m3v7-cancel").addEventListener("click",()=>{
  if(!busy)return;
  ++request;worker?.terminate();worker=null;last=null;
  setBusy(false);$("m3v7-status").textContent="사용자 취소: Web Worker 종료됨 (결과는 생성하지 않음).";
  $("m3v7-pct").textContent="취소";$("m3v7-export").disabled=true;$("m3v7-report").disabled=true;
 });
 $("m3v7-export").addEventListener("click",()=>{
  if(last)download("LS_Wave_Vector_v0.7_colab_compatible.json",last.portableVectorResult);
 });
 $("m3v7-report").addEventListener("click",()=>{
  if(last)download("LS_Wave_M3_v0.7_study.json",last);
 });
 $("m3v7-upload").addEventListener("change",async e=>{
  const f=e.target.files?.[0];if(!f)return;
  if(f.size>12000000){$("m3v7-refstate").textContent="Colab JSON 12 MB 초과";return}
  try{
   colab=normalizeColabJSON(JSON.parse(await f.text()));colabSource=f.name;
   $("m3v7-refstate").textContent="Colab 기준: "+f.name+" neff "+num(colab.modes[0].neff);
   if(last)displayMode();
  }catch(error){$("m3v7-refstate").textContent="Colab 파일 오류: "+error.message}
  e.target.value="";
 });
 estimate();setBusy(false);
 globalThis.LSWaveM3Study=Object.freeze({enabled:true,latest:()=>last,show,cancel:()=>{
  if(busy)$("m3v7-cancel").click();
 }});
}
if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",init,{once:true});else init();
