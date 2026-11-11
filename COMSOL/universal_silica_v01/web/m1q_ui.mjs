/**
 * Optional M1 v0.5.1 step-index verification section.
 * Does not alter existing model inputs, cards, FEM code, old M1 or IO schema.
 * Feature flag: window.LSWaveFeatures.m1Quantitative=false before module load.
 */
import {runM1Quantitative} from "./m1q_engine.mjs";
import {m1qReportJSON,m1qReportCSV} from "./m1q_export.mjs";
import {M1Q_VERSION,M1Q_TOLERANCES,M1Q_COMSOL_URL} from "./m1q_cases.mjs";
const ENABLED=globalThis.LSWaveFeatures?.m1Quantitative!==false;
const $=id=>document.getElementById(id);
const esc=s=>String(s??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const fmt=n=>typeof n==="number"&&Number.isFinite(n)?n.toPrecision(12):"—";
const sci=n=>typeof n==="number"&&Number.isFinite(n)?n.toExponential(3):"—";
let latest=null;
function download(name,content,type){
 const blob=new Blob([content],{type}),url=URL.createObjectURL(blob);
 const link=document.createElement("a");link.href=url;link.download=name;
 document.body.appendChild(link);link.click();link.remove();
 setTimeout(()=>URL.revokeObjectURL(url),1000);
}
function caseRows(report){
 return report.cases.map(c=>{
  const x=c.inputs;
  const params="λ="+x.wavelength_nm+" nm; a="+x.core_radius_um+
   " µm; Δn="+x.delta_n_core+"; n_clad="+x.n_clad+
   "; R="+x.outer_radius_um+" µm; N="+x.elements;
  return '<tr data-m1q-case="'+esc(c.id)+'" class="'+(c.passed?"is-pass":"is-fail")+'">'+
   '<td><b>'+esc(c.name)+'</b><small>'+esc(c.sourceType)+'</small></td>'+
   '<td>'+esc(params)+'<small>V='+fmt(c.V)+'</small></td>'+
   '<td>'+fmt(c.analyticNeff)+'</td><td>'+fmt(c.femNeff)+'</td>'+
   '<td>'+sci(c.absError)+'</td><td>'+sci(c.relativePercent)+'%</td>'+
   '<td>≤ '+sci(c.toleranceAbsNeff)+'</td><td><b>'+esc(c.status)+'</b>'+
   (c.error?'<small>'+esc(c.error)+'</small>':'')+'</td></tr>';
 }).join("");
}
function regressionRows(report){
 return report.regression.map(r=>{
  const keys=["neff","mfd","aeff","D"];
  const cells=[
   r.id,
   keys.map(k=>fmt(r.before[k])).join(" / "),
   keys.map(k=>fmt(r.after[k])).join(" / "),
   keys.map(k=>sci(r.absError[k])).join(" / "),
   r.status
  ];
  return '<tr data-m1q-regression="'+esc(r.id)+'" class="'+(r.passed?"is-pass":"is-fail")+'">'+
   cells.map(x=>'<td>'+esc(x)+'</td>').join("")+'</tr>';
 }).join("");
}
function init(){
 if(!ENABLED||$("ls-m1q-section"))return;
 const nav=document.querySelector("nav.eng-shortcuts");
 if(!nav)return;
 const tab=document.createElement("a");tab.id="ls-m1q-nav";
 tab.href="#ls-m1q-section";tab.textContent="06 · NUMERICAL VERIFICATION";
 tab.setAttribute("aria-controls","ls-m1q-section");nav.append(tab);
 const section=document.createElement("section");
 section.id="ls-m1q-section";section.className="ls-m1q engineering-panel";section.hidden=true;
 section.setAttribute("aria-label","06 NUMERICAL VERIFICATION 해석해 대비");
 section.innerHTML=[
  '<div class="ls-m1q-head"><div><small>06 · NUMERICAL VERIFICATION(해석해 대비) · v'+esc(M1Q_VERSION)+'</small>',
  '<h2>Scalar Radial FEM ↔ LP₀₁ Bessel 해석해</h2></div><span>독립 검증 · 기존 설계 보존</span></div>',
  '<p class="ls-m1q-intro">4개 step-index 기준에서 고정 굴절률의 LP₀₁ 특성방정식과 기존 FEM 고유값을 비교합니다. ',
  'a=4.1 µm / Δn=0.005는 사용자 요청 조건이며 COMSOL 공식 App 145와 구분합니다.</p>',
  '<div class="ls-m1q-buttons">',
  '<button id="ls-m1q-run">▶ 검증 실행</button>',
  '<button id="ls-m1q-json" disabled>결과 JSON 저장</button>',
  '<button id="ls-m1q-csv" disabled>결과 CSV 저장</button>',
  '<button id="ls-m1q-hide">접기</button></div>',
  '<p id="ls-m1q-status" role="status" aria-live="polite">검증 실행을 누르면 수치 비교를 시작합니다.</p>',
  '<div id="ls-m1q-summary" aria-live="polite"></div>',
  '<div class="ls-m1q-scroll"><table class="ls-m1q-table"><thead><tr>',
  '<th>케이스 / 기준 출처</th><th>입력</th><th>해석해 neff</th><th>FEM neff</th>',
  '<th>절대오차</th><th>상대오차 (%)</th><th>허용오차</th><th>판정</th>',
  '</tr></thead><tbody id="ls-m1q-cases"><tr><td colspan="8">계산 대기</td></tr></tbody></table></div>',
  '<details class="ls-m1q-regression" open><summary>기존 5개 프리셋 회귀 비교 · neff / MFD(PII,µm) / Aeff(µm²) / D</summary>',
  '<div class="ls-m1q-scroll"><table class="ls-m1q-table"><thead><tr><th>프리셋</th>',
  '<th>업데이트 전 v0.5.0 (neff / MFD / Aeff / D)</th>',
  '<th>업데이트 후 (동일 순서)</th><th>|변화량| (동일 순서)</th><th>판정</th>',
  '</tr></thead><tbody id="ls-m1q-regression">',
  '<tr><td colspan="5">계산 대기</td></tr></tbody></table></div></details>',
  '<div class="ls-m1q-notes"><h3>허용오차 · 해석 근거 · 미검증 사항</h3>',
  '<p><b>FEM ↔ LP₀₁</b>: |Δneff| ≤ '+sci(M1Q_TOLERANCES.absNeff)+
  '. 독립 SciPy ↔ JavaScript 해석해: ≤ '+sci(M1Q_TOLERANCES.analyticSciPy)+
  '. FEM 고유값 상대잔차 ≤ '+sci(M1Q_TOLERANCES.eigenResidual)+
  '. 상대오차는 |FEM−해석해|/|해석해|×100(%)이며 판정은 절대오차와 잔차를 적용합니다.</p>',
  '<p><b>해석해</b>: 원형 step-index, 약유도 LP₀₁, 무한 클래딩.',
  ' 0&lt;u&lt;min(V,2.4048255577)에서 이분법 최대 75회, 수렴 기준 구간폭 4×10⁻¹⁴.',
  ' 기존 FEM은 유한 R 경계·P1 요소이므로 해석해와 차이가 발생합니다.</p>',
  '<p><b>COMSOL 공식 출처</b>: <a href="'+esc(M1Q_COMSOL_URL)+'" target="_blank" rel="noopener noreferrer">',
  'Step-Index Fiber App 145</a>, λ=1550 nm, a=8 µm, ncore=1.4457, nclad=1.4378,',
  ' 공개 neff≈1.4444 (벡터 HE₁₁, 소수 4자리 반올림).',
  ' 공식값과의 반올림 범위 비교는 스칼라 해석해 검증과 구분합니다.</p>',
  '<p><b>미검증</b>: a=4.1 µm / Δn=0.005의 COMSOL 직접 실행값, 1310 nm의 실제 재료 분산,',
  ' 벡터 모드·PML·HCF·MCF·굽힘/누설 손실·실측값.</p>',
  '<p><b>회귀 허용오차</b>: neff '+sci(M1Q_TOLERANCES.regressionNeff)+
  ', MFD '+sci(M1Q_TOLERANCES.regressionMfdUm)+' µm, Aeff '+
  sci(M1Q_TOLERANCES.regressionAeffUm2)+' µm², D '+
  sci(M1Q_TOLERANCES.regressionD)+' ps/(nm·km).</p></div>'
 ].join("");
 nav.after(section);
 const status=$("ls-m1q-status"),run=$("ls-m1q-run"),json=$("ls-m1q-json"),csv=$("ls-m1q-csv");
 const show=()=>{section.hidden=false;tab.setAttribute("aria-current","true");};
 const hide=()=>{section.hidden=true;tab.removeAttribute("aria-current");};
 tab.addEventListener("click",event=>{event.preventDefault();show();section.scrollIntoView({behavior:"smooth",block:"start"});});
 $("ls-m1q-hide").addEventListener("click",hide);
 run.addEventListener("click",async()=>{
  if(run.disabled)return;
  show();run.disabled=true;json.disabled=true;csv.disabled=true;latest=null;
  status.textContent="계산 중 · 기존 사용자 설계와 결과값은 변경하지 않습니다.";
  $("ls-m1q-cases").innerHTML='<tr><td colspan="8">계산 중…</td></tr>';
  try{
   await new Promise(resolve=>setTimeout(resolve,0));
   const result=runM1Quantitative({onProgress:step=>{status.textContent="검증 중 · "+step;}});
   latest=result;
   $("ls-m1q-cases").innerHTML=caseRows(result);
   $("ls-m1q-regression").innerHTML=regressionRows(result);
   $("ls-m1q-summary").textContent="해석해 비교 "+result.passCount+"/"+result.caseCount+
    " PASS · 기존 결과 회귀 "+result.regressionPassCount+"/"+result.regressionCount+
    " PASS · "+(result.passed?"정의된 기준 충족":"FAIL 항목 존재");
   status.textContent="검증 종료 · COMSOL 직접 비교는 공식 App 145 반올림 참조값 이외 미검증";
   json.disabled=false;csv.disabled=false;
  }catch(error){
   status.textContent="검증 실패: "+String(error?.message||error);
   $("ls-m1q-cases").innerHTML='<tr><td colspan="8">오류 · 위 메시지 확인</td></tr>';
  }finally{run.disabled=false;}
 });
 json.addEventListener("click",()=>{if(latest)download("LS_Wave_M1_v0.5.1_report.json",m1qReportJSON(latest),"application/json;charset=utf-8");});
 csv.addEventListener("click",()=>{if(latest)download("LS_Wave_M1_v0.5.1_report.csv",m1qReportCSV(latest),"text/csv;charset=utf-8");});
 globalThis.LSWaveM1Quantitative=Object.freeze({enabled:true,run:runM1Quantitative,latest:()=>latest,show,hide});
}
if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",init,{once:true});else init();
