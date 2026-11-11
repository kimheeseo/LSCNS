# M1 v0.5.1 — Full new source files (copy-ready)

These eight new files are reproduced in full from the committed main branch. Copy each code block into the named path relative to web/. Original-file edits are documented as minimal diffs in [M1Q_VALIDATION.md](M1Q_VALIDATION.md); the original scalar FEM code has not been rewritten.

## 1. m1q_cases.mjs

~~~javascript
/**
 * M1 v0.5.1 quantitative verification fixtures. User-configured a=4.1 um,
 * delta n=0.005 is NOT COMSOL's published App 145 geometry (a=8 um).
 * Fixed n_clad=1.4440 in user cases is an explicit benchmark assumption,
 * not a claimed COMSOL value or manufacturer measurement.
 * Analytic baseline values are independently reproduced with SciPy J0/J1/K0/K1
 * + Brent root solve; they are not injected into the browser eigensolver.
 */
export const M1Q_VERSION="0.5.1";
export const M1Q_SCHEMA=1;
export const M1Q_COMSOL_URL="https://doc.comsol.com/6.4/doc/com.comsol.help.models.rf.step_index_fiber/step_index_fiber.html";
export const M1Q_TOLERANCES=Object.freeze({
  absNeff:1e-6,
  relativePercent:1e-4, // Display-only reference: |FEM-analytic| / |analytic| * 100
  analyticSciPy:1e-10,
  eigenResidual:1e-7,
  comsolRounded:5e-5,
  regressionNeff:1e-10,
  regressionMfdUm:1e-6,
  regressionAeffUm2:1e-6,
  regressionD:1e-6
});
export const M1Q_CASES=Object.freeze([
 Object.freeze({id:"official_comsol_145",name:"COMSOL 공식 App 145",
  sourceType:"COMSOL 공식 입력 + SciPy LP01 해석해",sourceUrl:M1Q_COMSOL_URL,
  wavelength_nm:1550,core_radius_um:8,n_clad:1.4378,delta_n_core:0.0079,
  outer_radius_um:40,elements:750,
  scipy_neff:1.4444018444092133,comsol_neff_rounded:1.4444}),
 Object.freeze({id:"requested_41",name:"사용자 지정 4.1 µm / Δn 0.005",
  sourceType:"사용자 지정 조건 + SciPy LP01 해석해; COMSOL 공식 결과 아님",
  wavelength_nm:1550,core_radius_um:4.1,n_clad:1.4440,delta_n_core:0.005,
  outer_radius_um:40,elements:750,scipy_neff:1.446081132401304}),
 Object.freeze({id:"low_v",name:"저V 단일모드 · V < 2.405",
  sourceType:"사용자 검증용 고정 재료 + SciPy LP01 해석해",
  wavelength_nm:1550,core_radius_um:4.1,n_clad:1.4440,delta_n_core:0.0032,
  outer_radius_um:40,elements:750,scipy_neff:1.444863128995567}),
 Object.freeze({id:"different_wavelength",name:"다른 파장 · 1310 nm",
  sourceType:"고정 굴절률 1310 nm 비교 + SciPy LP01 해석해; 실제 재료 분산 미포함",
  wavelength_nm:1310,core_radius_um:4.1,n_clad:1.4440,delta_n_core:0.005,
  outer_radius_um:40,elements:750,scipy_neff:1.446608620045317})
]);
/**
 * Captured BEFORE M1 changes from v0.5.0 solver.mjs
 * blob SHA: 6d774b556b83829c77a3b6f3d208eb8ea8da6a3a.
 * Same copy/preset, same execution/runtime; these are snapshots, not
 * independent physical-material experimental references.
 */
export const M1Q_PREUPDATE=Object.freeze([
 Object.freeze({id:"g654",neff:1.4457189679091202,mfd:12.500052522796532,aeff:121.9238953090936,D:19.825664231522584}),
 Object.freeze({id:"g652",neff:1.446104728174543,mfd:10.182785332592127,aeff:79.29336129096376,D:16.9928875566912}),
 Object.freeze({id:"graded",neff:1.445828815012516,mfd:10.192939528309598,aeff:76.8841326978277,D:14.919171784564464}),
 Object.freeze({id:"trench",neff:1.4460435091028088,mfd:9.839323631718809,aeff:74.82460400259428,D:17.80387925775665}),
 Object.freeze({id:"multi",neff:1.447001785935374,mfd:9.582447747061527,aeff:71.68879956968254,D:19.162570615215405})
]);
~~~

## 2. m1q_engine.mjs

~~~javascript
/**
 * M1 v0.5.1 read-only quantitative verification. No shared UI/model mutation.
 * FEM uses the *unchanged* original scalar P1 radial solver.
 * Independent LP01 Bessel oracle is imported from validation_lp01.mjs.
 * Oracle numerical method:
 *   F(u)=u J1(u)/J0(u) - w K1(w)/K0(w), w=sqrt(V²-u²).
 *   Bracket: 0<u<min(V, first zero J0=2.4048255577).
 *   Bisection: at most 75 iterations, stop when |upper-lower|<4e-14
 *   (dimensionless u). J functions use convergent series, K ratio uses
 *   1024-interval Simpson quadrature of the positive Bessel-K integral.
 *   Check independent SciPy Brent reference at |delta neff|<=1e-10 and
 *   characteristic F(u) residual <=1e-8 (not a FEM residual).
 */
import {PRESETS,solve,dispersion} from "./solver.mjs";
import {lp01Characteristic} from "./validation_lp01.mjs";
import {M1Q_VERSION,M1Q_SCHEMA,M1Q_CASES,M1Q_PREUPDATE,M1Q_TOLERANCES,M1Q_COMSOL_URL} from "./m1q_cases.mjs";
const T=M1Q_TOLERANCES;
function finite(n){return typeof n==="number"&&Number.isFinite(n);}
export function m1qFixtureModel(c){
 return {...PRESETS.comsol_step,profile:"step",wavelength_nm:c.wavelength_nm,
  core_radius_um:c.core_radius_um,delta_n_core:c.delta_n_core,
  clad_index_override:c.n_clad,outer_radius_um:c.outer_radius_um,elements:c.elements};
}
export function verifyM1QCases(){
 return M1Q_CASES.map(c=>{
  const base={
   id:c.id,name:c.name,sourceType:c.sourceType,sourceUrl:c.sourceUrl??null,
   inputs:{wavelength_nm:c.wavelength_nm,core_radius_um:c.core_radius_um,
    delta_n_core:c.delta_n_core,n_clad:c.n_clad,n_core:c.n_clad+c.delta_n_core,
    outer_radius_um:c.outer_radius_um,elements:c.elements},
   referenceNeffSciPy:c.scipy_neff,toleranceAbsNeff:T.absNeff,
   comsolPublishedNeff:c.comsol_neff_rounded??null,
   comsolStatus:c.comsol_neff_rounded==null?"미검증":"미검증",
   analyticNeff:null,femNeff:null,absError:null,relativePercent:null,
   V:null,analyticResidual:null,femEigenResidual:null,
   passed:false,status:"FAIL",error:null
  };
  try{
   const anal=lp01Characteristic({
    wavelength_nm:c.wavelength_nm,core_radius_um:c.core_radius_um,
    n_core:c.n_clad+c.delta_n_core,n_clad:c.n_clad
   });
   const fem=solve(m1qFixtureModel(c));
   const absError=Math.abs(fem.n_eff-anal.n_eff);
   const relativePercent=absError/Math.abs(anal.n_eff)*100;
   const sciPyAbs=Math.abs(anal.n_eff-c.scipy_neff);
   const valid=finite(anal.n_eff)&&finite(fem.n_eff)&&finite(fem.residual)&&
    Math.abs(anal.residual)<=1e-8&&sciPyAbs<=T.analyticSciPy&&
    fem.residual<=T.eigenResidual&&fem.guided_proxy===true;
   const passed=valid&&absError<=T.absNeff;
   let comsolStatus="미검증";
   if(c.comsol_neff_rounded!=null){
    comsolStatus=Math.abs(fem.n_eff-c.comsol_neff_rounded)<=T.comsolRounded?"반올림 범위 내":"반올림 범위 외";
   }
   return {...base,analyticNeff:anal.n_eff,femNeff:fem.n_eff,absError,relativePercent,
    V:anal.V,analyticResidual:anal.residual,femEigenResidual:fem.residual,
    sciPyAbsError:sciPyAbs,comsolStatus,status:passed?"PASS":"FAIL",passed};
  }catch(error){return {...base,error:String(error?.message||error)};}
 });
}
export function verifyM1QLegacy(){
 return M1Q_PREUPDATE.map(c=>{
  const sol=solve({...PRESETS[c.id]}),D=dispersion({...PRESETS[c.id]});
  const now={neff:sol.n_eff,mfd:sol.MFD_PII_um,aeff:sol.Aeff_um2,D};
  const before={neff:c.neff,mfd:c.mfd,aeff:c.aeff,D:c.D};
  const deviations={
   neff:Math.abs(now.neff-before.neff),
   mfd:Math.abs(now.mfd-before.mfd),
   aeff:Math.abs(now.aeff-before.aeff),
   D:Math.abs(now.D-before.D)
  };
  const tolerances={
   neff:T.regressionNeff,mfd:T.regressionMfdUm,
   aeff:T.regressionAeffUm2,D:T.regressionD
  };
  const passed=Object.keys(tolerances).every(key=>finite(now[key])&&deviations[key]<=tolerances[key]);
  return {id:c.id,before,after:now,absError:deviations,tolerances,passed,status:passed?"PASS":"FAIL"};
 });
}
export function runM1Quantitative({onProgress=()=>{}}={}){
 onProgress("Step-index LP01 / 독립 SciPy 기준 4개");
 const cases=verifyM1QCases();
 onProgress("기존 5개 프리셋 계산 결과 비교");
 const regression=verifyM1QLegacy();
 const allPass=cases.every(x=>x.passed)&&regression.every(x=>x.passed);
 return {
  schemaVersion:M1Q_SCHEMA,version:M1Q_VERSION,kind:"LSWaveM1QuantitativeVerification",
  generatedAt:new Date().toISOString(),
  source:{comsol:M1Q_COMSOL_URL,
   analytic:"step-index scalar LP01 Bessel characteristic equation; SciPy scipy.special + scipy.optimize.brentq independent fixture",
   legacy:"v0.5.0 solver.mjs blob SHA 6d774b556b83829c77a3b6f3d208eb8ea8da6a3a"},
  passed:allPass,passCount:cases.filter(x=>x.passed).length,
  caseCount:cases.length,regressionPassCount:regression.filter(x=>x.passed).length,
  regressionCount:regression.length,cases,regression,tolerances:{...T},
  unverified:[
   "사용자 지정 a=4.1 µm, Δn=0.005는 COMSOL 공식 공개 예제값이 아니므로 COMSOL 직접 결과 미검증",
   "LP01 스칼라 해석해는 무한 클래딩·약유도 가정. 유한 R 경계와 FEM 메시 오차 존재",
   "1310 nm 케이스의 n_clad=1.4440 및 Δn=0.005는 고정 상수; 재료 분산 미검증",
   "벡터 HE11 편광, PML, HCF, MCF, PMD, 실측 감쇠, 제조사 규격은 이 M1 사례로 검증되지 않음"
  ]
 };
}
~~~

## 3. m1q_export.mjs

~~~javascript
/**
 * M1 v0.5.1 independent report writers; design JSON and existing export
 * formats are untouched.
 */
function csvCell(value){
 const raw=value==null?"":String(value);
 // Escape user-supplied text and prevent spreadsheet formula injection.
 const safe=/^[=+@\-]/.test(raw.trimStart())?"'"+raw:raw;
 return '"'+safe.replace(/"/g,'""')+'"';
}
export function m1qReportCSV(report){
 if(report?.kind!=="LSWaveM1QuantitativeVerification"||!Array.isArray(report.cases)||!Array.isArray(report.regression))
  throw Error("Invalid M1 report");
 const columns=["row_type","case_id","source","wavelength_nm","core_radius_um","delta_n_core",
  "n_clad","n_core","V","lp01_analytic_neff","fem_neff","abs_error_neff","relative_error_percent",
  "tolerance_abs_neff","comsol_direct_status","metric","preupdate","postupdate",
  "regression_abs_error","regression_tolerance","status"];
 const rows=[];
 for(const c of report.cases){
  rows.push(["case",c.id,c.sourceType,c.inputs.wavelength_nm,c.inputs.core_radius_um,
   c.inputs.delta_n_core,c.inputs.n_clad,c.inputs.n_core,c.V,
   c.analyticNeff,c.femNeff,c.absError,c.relativePercent,c.toleranceAbsNeff,
   c.comsolStatus,"","","","","",c.status]);
 }
 for(const r of report.regression){
  for(const key of ["neff","mfd","aeff","D"])
   rows.push(["regression",r.id,"v0.5.0 before/after","","","","","","","","","","","","",
    key,r.before[key],r.after[key],r.absError[key],r.tolerances[key],
    r.absError[key]<=r.tolerances[key]?"PASS":"FAIL"]);
 }
 return "\uFEFF"+[columns,...rows].map(row=>row.map(csvCell).join(",")).join("\r\n")+"\r\n";
}
export function m1qReportJSON(report){
 if(report?.kind!=="LSWaveM1QuantitativeVerification")throw Error("Invalid M1 report");
 return JSON.stringify(report,null,2)+"\n";
}
~~~

## 4. m1q_ui.mjs

~~~javascript
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
   '<td><b>'+esc(c.name)+'</b><small>'+esc(c.sourceType)+'</small>'+
   '<small>COMSOL 직접 비교: '+esc(c.comsolStatus)+'</small></td>'+
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
~~~

## 5. m1q_ui.css

~~~css
/* M1 v0.5.1 independent verification section; do not override legacy UI. */
#ls-m1q-section[hidden]{display:none!important}
#ls-m1q-section{scroll-margin-top:85px;margin:16px 0 23px;padding:19px;
 color:#d8edfa;background:#0e2339;border:1px solid #3e667c;border-radius:12px}
#ls-m1q-section .ls-m1q-head{display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap}
#ls-m1q-section .ls-m1q-head small{color:#7ad7d8;font-size:10px;letter-spacing:.07em}
#ls-m1q-section h2{margin:6px 0 12px;font-size:19px;color:#f0f9ff}
#ls-m1q-section .ls-m1q-head span{color:#9fcbd8;font-size:11px;border:1px solid #476d80;border-radius:18px;padding:7px}
#ls-m1q-section .ls-m1q-intro,#ls-m1q-section .ls-m1q-notes p{font-size:12px;line-height:1.75;color:#c1d2e0}
#ls-m1q-section .ls-m1q-buttons{display:flex;flex-wrap:wrap;gap:9px;margin:14px 0}
#ls-m1q-section .ls-m1q-buttons button{padding:9px 13px;font-size:12px;font-weight:600;cursor:pointer;
 color:#eaf8fc;background:#1a4763;border:1px solid #4a869f;border-radius:8px}
#ls-m1q-section .ls-m1q-buttons button:hover:not(:disabled){background:#296381}
#ls-m1q-section .ls-m1q-buttons button:disabled{opacity:.44;cursor:default}
#ls-m1q-status{color:#91dcd0;font-size:12px;min-height:22px;margin:10px 0}
#ls-m1q-summary{font-size:12px;color:#b7f0d9;font-weight:650;margin:10px 0}
#ls-m1q-section .ls-m1q-scroll{max-width:100%;max-height:440px;overflow:auto;border:1px solid #30526a;border-radius:7px}
#ls-m1q-section .ls-m1q-table{min-width:930px;width:100%;border-collapse:collapse;line-height:1.5;font-size:11px}
#ls-m1q-section .ls-m1q-table th{color:#b6dbe8;background:#18344d;position:sticky;top:0;z-index:1}
#ls-m1q-section .ls-m1q-table td,#ls-m1q-section .ls-m1q-table th{
 padding:10px 9px;text-align:left;border-bottom:1px solid #345168;vertical-align:top}
#ls-m1q-section .ls-m1q-table td{color:#d8eaf4;font-variant-numeric:tabular-nums;white-space:nowrap}
#ls-m1q-section .ls-m1q-table td small{display:block;font-size:10px;color:#97b9cc;white-space:normal;max-width:230px;margin-top:5px}
#ls-m1q-section .ls-m1q-table .is-pass td:last-child{font-weight:750;color:#7ce5b7}
#ls-m1q-section .ls-m1q-table .is-fail td:last-child{font-weight:750;color:#f8a6a6}
#ls-m1q-section .ls-m1q-regression{margin:16px 0}
#ls-m1q-section .ls-m1q-regression summary{cursor:pointer;color:#b6e9f0;margin:10px 0;font-size:12px}
#ls-m1q-section .ls-m1q-notes{border-top:1px solid #3c6073;margin-top:14px;padding-top:10px}
#ls-m1q-section .ls-m1q-notes h3{font-size:13px;color:#bce4ef}
#ls-m1q-section .ls-m1q-notes a{color:#7ed9e7}
@media(max-width:800px){#ls-m1q-section{padding:12px}#ls-m1q-section h2{font-size:16px}}
~~~

## 6. tests/test_m1_quantitative.mjs

~~~javascript
/**
 * M1 v0.5.1 independent LP01 quantitative verification and legacy regression.
 * Run: node tests/test_m1_quantitative.mjs
 */
import assert from "node:assert/strict";
import {PRESETS,solve} from "../solver.mjs";
import {M1Q_CASES,M1Q_PREUPDATE,M1Q_TOLERANCES} from "../m1q_cases.mjs";
import {m1qFixtureModel,runM1Quantitative} from "../m1q_engine.mjs";
import {m1qReportJSON,m1qReportCSV} from "../m1q_export.mjs";
const before={...PRESETS.g652};
const report=runM1Quantitative();
assert.equal(report.version,"0.5.1");
assert.equal(report.schemaVersion,1);
assert.equal(report.kind,"LSWaveM1QuantitativeVerification");
assert.equal(report.caseCount,4);
assert.equal(report.regressionCount,5);
assert.equal(report.passCount,4);
assert.equal(report.regressionPassCount,5);
assert.equal(report.passed,true);
for(const [i,c] of report.cases.entries()){
 const source=M1Q_CASES[i];
 assert.equal(c.id,source.id);
 assert.equal(c.status,"PASS");
 assert.ok(c.absError<=M1Q_TOLERANCES.absNeff,c.id+" LP01 neff tolerance");
 assert.ok(c.sciPyAbsError<=M1Q_TOLERANCES.analyticSciPy,c.id+" SciPy independent benchmark");
 assert.ok(c.femEigenResidual<=M1Q_TOLERANCES.eigenResidual,c.id+" FEM algebraic residual");
 assert.ok(c.relativePercent>=0&&c.relativePercent<.0001,c.id+" relative error");
 assert.equal(c.inputs.n_clad,source.n_clad);
 assert.equal(c.inputs.n_core,source.n_clad+source.delta_n_core);
 assert.equal(m1qFixtureModel(source).profile,"step");
 if(c.id!=="official_comsol_145")assert.equal(c.comsolStatus,"미검증");
 else assert.equal(c.comsolStatus,"반올림 범위 내");
}
assert.ok(report.cases.find(c=>c.id==="low_v").V<2.4048255577);
assert.equal(report.cases.find(c=>c.id==="different_wavelength").inputs.wavelength_nm,1310);
assert.equal(report.cases.find(c=>c.id==="requested_41").inputs.core_radius_um,4.1);
for(const item of report.regression){
 const ref=M1Q_PREUPDATE.find(x=>x.id===item.id);
 assert.equal(item.status,"PASS");
 assert.equal(item.before.neff,ref.neff);
 for(const key of ["neff","mfd","aeff","D"]){
  assert.ok(item.absError[key]<=item.tolerances[key],item.id+" "+key+" changed");
 }
}
assert.deepEqual(PRESETS.g652,before,"Running verification must not modify original presets");
const old=solve({...PRESETS.g652});
assert.ok(Math.abs(old.n_eff-M1Q_PREUPDATE[1].neff)<=M1Q_TOLERANCES.regressionNeff);
const json=JSON.parse(m1qReportJSON(report));
assert.equal(json.caseCount,4);
const csv=m1qReportCSV(report);
assert.ok(csv.startsWith("\uFEFF"));
assert.equal(csv.trim().split("\r\n").length,25,"1 header + 4 cases + 20 regression metric rows");
assert.ok(csv.includes("relative_error_percent"));
assert.ok(csv.includes("미검증"));
console.log("PASS M1 v0.5.1: 4 independent LP01 cases, 5 legacy presets / 20 metric regressions, JSON+CSV exports");
~~~

## 7. tests/m1q_reference_scipy.py

~~~python
"""M1 v0.5.1 independent step-index LP01 roots (development-only).

Run from web/: python tests/m1q_reference_scipy.py
Requires SciPy; browser app has no Python/Colab/server dependency.
This checks stored *analytic* reference values, not COMSOL vector-mode neff.
"""
from __future__ import annotations
import math
from scipy.special import j0,j1,k0,k1
from scipy.optimize import brentq

CASES=(
    ("official_comsol_145",1550.,8.,1.4378,.0079,1.4444018444092133),
    ("requested_41",1550.,4.1,1.4440,.005,1.446081132401304),
    ("low_v",1550.,4.1,1.4440,.0032,1.444863128995567),
    ("different_wavelength",1310.,4.1,1.4440,.005,1.446608620045317),
)
def lp01_neff(wavelength_nm,a_um,n_clad,delta_n):
    n_core=n_clad+delta_n
    k0=2*math.pi/(wavelength_nm/1000.)
    v=k0*a_um*math.sqrt(n_core*n_core-n_clad*n_clad)
    def characteristic(u):
        w=math.sqrt(v*v-u*u)
        return u*j1(u)/j0(u)-w*k1(w)/k0(w)
    # Independent SciPy Brent bracketing; root tolerance 1e-13 in u.
    upper=min(v,2.404825557695773)*(1-1e-9)
    u=brentq(characteristic,1e-8,upper,xtol=1e-13,rtol=1e-14,maxiter=100)
    neff=math.sqrt(n_core*n_core-(u/(k0*a_um))**2)
    return neff,v,characteristic(u)

if __name__=="__main__":
    for name,w,a,ncl,dn,reference in CASES:
        actual,v,res=lp01_neff(w,a,ncl,dn)
        difference=abs(actual-reference)
        assert difference<=1e-10,(name,difference)
        assert abs(res)<=1e-9,(name,res)
        print(f"{name:22s}  V={v:.9f}  LP01_neff={actual:.13f}  diff={difference:.3e} PASS")
    print("PASS 4/4 independent SciPy root references.")
~~~

## 8. tests/browser_m1_quantitative.mjs

~~~javascript
/** Optional Playwright M1 v0.5.1 smoke/regression.
 * Serve LSCNS repository root: python3 -m http.server 8765
 * Run in web/: node tests/browser_m1_quantitative.mjs
 * Or TEST_BASE_URL=https://.../web/ node tests/browser_m1_quantitative.mjs
 */
import assert from "node:assert/strict";
import {createRequire} from "node:module";
const {chromium}=createRequire(import.meta.url)("playwright");
const browser=await chromium.launch({headless:true,args:["--no-sandbox","--disable-dev-shm-usage"]});
try{
 const page=await browser.newPage({viewport:{width:1500,height:920},acceptDownloads:true});
 const errors=[];
 page.on("pageerror",e=>errors.push(e.message));
 await page.goto(process.env.TEST_BASE_URL||"http://127.0.0.1:8765/COMSOL/universal_silica_v01/web/",{waitUntil:"load",timeout:90000});
 if(await page.locator('[data-wizard="cancel"]').isVisible())await page.locator('[data-wizard="cancel"]').click();
 await page.waitForFunction(()=>!["—",""].includes(document.getElementById("neff")?.textContent?.trim()||""),undefined,{timeout:90000});
 const legacy=await page.evaluate(()=>({
  preset:document.getElementById("preset").value,
  neff:document.getElementById("neff").textContent,
  mfd:document.getElementById("mfd").textContent,
  aeff:document.getElementById("aeff").textContent
 }));
 assert.ok(await page.locator("#ls-m1q-nav").isVisible());
 assert.equal(await page.locator("#ls-m1q-section").isHidden(),true);
 await page.locator("#ls-m1q-nav").click();
 assert.equal(await page.locator("#ls-m1q-section").isVisible(),true);
 await page.locator("#ls-m1q-run").click();
 await page.waitForFunction(()=>window.LSWaveM1Quantitative?.latest()?.caseCount===4,undefined,{timeout:90000});
 assert.equal(await page.locator("#ls-m1q-cases tr").count(),4);
 assert.equal(await page.locator("#ls-m1q-regression tr").count(),5);
 const result=await page.evaluate(()=>window.LSWaveM1Quantitative.latest());
 assert.equal(result.passCount,4);
 assert.equal(result.regressionPassCount,5);
 assert.ok(result.cases.every(c=>c.status==="PASS"));
 assert.equal(result.cases.find(c=>c.id==="requested_41").comsolStatus,"미검증");
 const still=await page.evaluate(()=>({
  preset:document.getElementById("preset").value,
  neff:document.getElementById("neff").textContent,
  mfd:document.getElementById("mfd").textContent,
  aeff:document.getElementById("aeff").textContent
 }));
 assert.deepEqual(still,legacy,"Original UI model / neff / MFD / Aeff must be unchanged.");
 const [json]=await Promise.all([page.waitForEvent("download"),page.locator("#ls-m1q-json").click()]);
 assert.equal(json.suggestedFilename(),"LS_Wave_M1_v0.5.1_report.json");
 const [csv]=await Promise.all([page.waitForEvent("download"),page.locator("#ls-m1q-csv").click()]);
 assert.equal(csv.suggestedFilename(),"LS_Wave_M1_v0.5.1_report.csv");
 assert.ok(await page.locator("#ls-m1-validation-tab").count()===1,"Existing M1 28-check tool still mounted.");
 assert.ok(await page.locator("#m2-run").count()===1,"M2 still exists.");
 assert.ok(await page.locator("#m3-run").count()===1,"M3 still exists.");
 assert.deepEqual(errors,[]);
 console.log("PASS M1 browser: 4-case PASS/FAIL, 5-preset before/after, JSON/CSV, and unchanged user results.");
} finally {await browser.close();}
~~~

## Existing-file diffs

See [M1Q_VALIDATION.md](M1Q_VALIDATION.md#④-minimal-diff-against-v050) and [CHANGELOG.md](CHANGELOG.md).
