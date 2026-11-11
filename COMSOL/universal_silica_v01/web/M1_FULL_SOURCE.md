# M1 v0.3.0 — Complete new source files (copy-ready)

All content below was retrieved from committed GitHub `main` files. Original solver/App/UI code is not reproduced or rewritten. Existing-file diffs are in [M1_VALIDATION.md](M1_VALIDATION.md).

## 1. `validation_lp01.mjs`

```javascript
/**
 * M1 independent scalar LP01 oracle for a lossless circular step-index fiber.
 * No dependency on solver.mjs and no reference numbers used in root finding.
 * Solves u*J1(u)/J0(u) = w*K1(w)/K0(w), u^2+w^2=V^2.
 * J0/J1: convergent Taylor series (0 < u < first J0 root).
 * K1/K0: positive integral representation, composite Simpson quadrature.
 */
const J0_FIRST_ZERO=2.404825557695773;
function assertFinitePositive(v,name){
 if(!Number.isFinite(v)||v<=0)throw new Error(name+" must be finite and positive.");
}
function j01(x){
 const z=-(x*x)/4;
 let t0=1,s0=1,t1=x/2,s1=t1;
 for(let k=1;k<=90;k++){
  t0*=z/(k*k);
  t1*=z/(k*(k+1));
  s0+=t0;s1+=t1;
  if(Math.max(Math.abs(t0),Math.abs(t1))<1e-17)break;
 }
 return [s0,s1];
}
/** Returns K1(x)/K0(x) without subtracting exponentially small numbers. */
function k1OverK0(x){
 assertFinitePositive(x,"w");
 // exp(-x*(cosh(t)-1)) cancels exp(-x) in K0 and K1.
 // x*cosh(T) > ~50 makes the neglected integral exponentially small.
 const T=Math.max(9,Math.acosh(1+50/x)),N=1024,h=T/N;
 let a=0,b=0;
 for(let i=0;i<=N;i++){
  const ch=Math.cosh(i*h),f=Math.exp(-x*(ch-1));
  const weight=(i===0||i===N)?1:(i%2?4:2);
  a+=weight*f;b+=weight*f*ch;
 }
 if(!(a>0&&Number.isFinite(a)&&Number.isFinite(b)))throw new Error("K-Bessel quadrature failed.");
 return b/a;
}
export function lp01Characteristic({wavelength_nm,core_radius_um,n_core,n_clad}){
 for(const [name,v] of Object.entries({wavelength_nm,core_radius_um,n_core,n_clad}))assertFinitePositive(v,name);
 if(!(n_core>n_clad))throw new Error("Core index must exceed cladding index.");
 const k0=2*Math.PI/(wavelength_nm/1000);
 const V=k0*core_radius_um*Math.sqrt(n_core*n_core-n_clad*n_clad);
 if(!(V>=0.5&&V<=15))throw new Error("LP01 oracle is conditioned for 0.5 <= V <= 15.");
 const hi=Math.min(V,J0_FIRST_ZERO)*(1-1e-12);
 const residual=u=>{
  const [J0,J1]=j01(u),w=Math.sqrt(Math.max(0,V*V-u*u));
  if(!(J0>0&&w>0))throw new Error("LP01 root bracket became singular.");
  return u*J1/J0-w*k1OverK0(w);
 };
 let lo=Math.max(1e-10,hi*1e-9),upper=hi;
 if(!(residual(lo)<0&&residual(upper)>0))throw new Error("Cannot bracket fundamental LP01 root.");
 for(let i=0;i<75;i++){
  const m=(lo+upper)/2;
  if(residual(m)>0)upper=m;else lo=m;
  if(upper-lo<4e-14)break;
 }
 const u=(lo+upper)/2,w=Math.sqrt(V*V-u*u);
 const neff=Math.sqrt(n_core*n_core-(u/(k0*core_radius_um))**2);
 const b=(neff*neff-n_clad*n_clad)/(n_core*n_core-n_clad*n_clad);
 return {n_eff:neff,V,u,w,b,residual:residual(u),method:"LP01 scalar characteristic equation"};
}
```

## 2. `validation_cases.mjs`

```javascript
/**
 * M1 independent public references and pre-update numeric regression baselines.
 * The COMSOL value is rounded to 4 decimals; never use it as a solver input.
 */
export const COMSOL_PUBLIC = Object.freeze({
 id:"comsol-step-index-145",name:"COMSOL Step-Index Fiber · Application ID 145",
 source:"https://doc.comsol.com/6.4/doc/com.comsol.help.models.rf.step_index_fiber/step_index_fiber.html",
 wavelength_nm:1550,core_radius_um:8,outer_radius_um:40,
 n_core:1.4457,n_clad:1.4378,reference_neff_rounded:1.4444,
 roundingHalfWidth:0.00005,
 comment:"Published full-vector HE11 result rounded to 4 decimals; not a high-precision FEM golden value."
});
export const LP01_CASES=Object.freeze([
 Object.freeze({
  id:"comsol_step",name:"COMSOL matching geometry & indices",
  wavelength_nm:1550,core_radius_um:8,outer_radius_um:40,n_core:1.4457,n_clad:1.4378,
  elements:750,expectedAnalytic:1.4444018444092133,expectedV:4.894537730479192,
  source:"COMSOL geometry + independent SciPy Bessel reference"
 }),
 Object.freeze({
  id:"single_mode",name:"Independent single-mode step-index (V < 2.405)",
  wavelength_nm:1550,core_radius_um:4.1,outer_radius_um:40,n_core:1.4472,n_clad:1.444,
  elements:750,expectedAnalytic:1.444863128995567,expectedV:1.598623098047904,
  source:"Independent SciPy scipy.special + brentq, not a COMSOL published number"
 })
]);
export const REGRESSION=Object.freeze([
 Object.freeze({id:"g654",neff:1.445718967909122,mfd:12.500052522436938,aeff:121.92389530338414,D:19.825664345368367}),
 Object.freeze({id:"g652",neff:1.4461047281745436,mfd:10.182785332301778,aeff:79.29336128699885,D:16.99288790683873}),
 Object.freeze({id:"graded",neff:1.4458288150125134,mfd:10.192939527752662,aeff:76.88413268997199,D:14.919171294740607}),
 Object.freeze({id:"trench",neff:1.4460435091028094,mfd:9.839323631520045,aeff:74.82460399998054,D:17.8038792615834}),
 Object.freeze({id:"multi",neff:1.447001785935375,mfd:9.582447746912562,aeff:71.68879956793322,D:19.16257095579606})
]);
export const TOL=Object.freeze({
 lp01Neff:1e-6,comsolRoundedNeff:5e-5,analyticVsSciPy:1e-10,
 neff:1e-8,mfdUm:1e-4,aeffUm2:1e-4,dispersion:0.005,residual:1e-7
});
```

## 3. `validation_engine.mjs`

```javascript
/**
 * M1 automated validation runner. Read-only: passes copies to the existing
 * solver and never mutates presets, user design, FEM numerical kernels, or JSON.
 */
import {PRESETS,solve,dispersion,meshStudy} from "./solver.mjs";
import {lp01Characteristic} from "./validation_lp01.mjs";
import {COMSOL_PUBLIC,LP01_CASES,REGRESSION,TOL} from "./validation_cases.mjs";

function check(group,test,value,reference,tolerance,note=""){
 const absError=Math.abs(value-reference),valid=Number.isFinite(value)&&Number.isFinite(absError);
 return {group,test,value,reference,tolerance,absError,passed:valid&&absError<=tolerance,note};
}
export function fixtureModel(c){
 return {...PRESETS.comsol_step,profile:"step",wavelength_nm:c.wavelength_nm,
  core_radius_um:c.core_radius_um,outer_radius_um:c.outer_radius_um,
  delta_n_core:c.n_core-c.n_clad,clad_index_override:c.n_clad,elements:c.elements};
}
export function validateLP01Fixtures(){
 const rows=[],details=[];
 for(const c of LP01_CASES){
  const analytic=lp01Characteristic(c),fem=solve(fixtureModel(c));
  details.push({id:c.id,name:c.name,analytic,femNeff:fem.n_eff,femResidual:fem.residual});
  rows.push(check("해석해 독립 기준",c.id+" · Bessel vs SciPy",
   analytic.n_eff,c.expectedAnalytic,TOL.analyticVsSciPy,"Analytic oracle independent of FEM"));
  rows.push(check("FEM vs LP01",c.id+" · n_eff",
   fem.n_eff,analytic.n_eff,TOL.lp01Neff,"Scalar vs scalar; same indices and geometry"));
  rows.push(check("FEM 잔차",c.id+" · residual",fem.residual,0,TOL.residual,
   "Algebraic relative residual; not a physical-accuracy proof"));
  if(c.id==="comsol_step"){
   rows.push(check("COMSOL 공개값",c.id+" · 4-decimal n_eff",
    fem.n_eff,COMSOL_PUBLIC.reference_neff_rounded,TOL.comsolRoundedNeff,
    "Rounded published vector HE11 reference; within rounding window only"));
  }
 }
 return {rows,details};
}
export function validateLegacyPresets({includeDispersion=true}={}){
 const rows=[];
 for(const gold of REGRESSION){
  const result=solve({...PRESETS[gold.id]});
  rows.push(check("기존 회귀",gold.id+" · n_eff",result.n_eff,gold.neff,TOL.neff));
  rows.push(check("기존 회귀",gold.id+" · MFD PII [um]",result.MFD_PII_um,gold.mfd,TOL.mfdUm));
  rows.push(check("기존 회귀",gold.id+" · Aeff [um²]",result.Aeff_um2,gold.aeff,TOL.aeffUm2));
  if(includeDispersion){
   const D=dispersion({...PRESETS[gold.id]});
   rows.push(check("기존 회귀",gold.id+" · D [ps/(nm km)]",D,gold.D,TOL.dispersion));
  }
 }
 return rows;
}
export function validateMesh(){
 const c=LP01_CASES[0],reference=lp01Characteristic(c).n_eff;
 const meshes=meshStudy(fixtureModel(c),[200,350,550,850]);
 const errors=meshes.map(m=>Math.abs(m.n_eff-reference));
 const monotonic=errors.every((v,i)=>i===0||v<=errors[i-1]+1e-12);
 const fine=errors.at(-1)<=TOL.lp01Neff;
 return {meshes:meshes.map((m,i)=>({...m,analyticAbsError:errors[i]})),
  rows:[
   {group:"메시 수렴",test:"200→350→550→850 elements · |Δn_eff| 감소",value:errors.at(-1),
    reference:0,tolerance:TOL.lp01Neff,absError:errors.at(-1),passed:monotonic&&fine,
    note:"Monotonic convergence and final error <= 1e-6; four meshes only"}
  ]};
}
/** Yields to the browser before expensive baseline dispersion checks. */
export async function runM1Validation({onProgress=()=>{},includeDispersion=true}={}){
 const rows=[],details=[],emit=(step)=>onProgress({step,completed:rows.length});
 onProgress({step:"LP01 analytical references",completed:0});
 const core=validateLP01Fixtures();rows.push(...core.rows);details.push(...core.details);
 emit("2 analytical cases and COMSOL rounded reference");
 await Promise.resolve(); // browser can move to next task on UI's own scheduling
 const mesh=validateMesh();rows.push(...mesh.rows);emit("mesh study");
 const baseline=validateLegacyPresets({includeDispersion});
 rows.push(...baseline);emit("legacy preset regression");
 return {schemaVersion:1,version:"0.3.0",kind:"M1-validation-report",
  source:COMSOL_PUBLIC.source,generatedAt:new Date().toISOString(),
  passed:rows.every(r=>r.passed),passedCount:rows.filter(r=>r.passed).length,
  totalCount:rows.length,rows,details,mesh:mesh.meshes,assumptions:[
   "COMSOL four-decimal HE11 is a rounded vector-mode public benchmark",
   "Independent analytic LP01 assumes weak guidance, infinite cladding and step index",
   "Finite-radius scalar radial FEM uses zero field at R=40 um",
   "No PML, vector solver, leakage/bend-loss physics or new materials are validated"
  ]};
}
```

## 4. `validation_ui.mjs`

```javascript
/**
 * M1 validation tab/section. Feature flag: window.LSWaveFeatures.m1Validation.
 * Defaults ON; does not read/modify active design, existing result cards or files.
 */
import {runM1Validation} from "./validation_engine.mjs";
import {COMSOL_PUBLIC,TOL} from "./validation_cases.mjs";
const FEATURES=Object.freeze({m1Validation:globalThis.LSWaveFeatures?.m1Validation!==false});
const esc=s=>String(s??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const fmt=n=>Number.isFinite(n)?Number(n).toPrecision(9):"—";
let last=null;
function init(){
 if(!FEATURES.m1Validation)return;
 const nav=document.querySelector("nav.eng-shortcuts"),main=document.querySelector("main.workspace");
 if(!nav||!main||document.getElementById("ls-m1-validation"))return;
 const tab=document.createElement("a");
 tab.href="#ls-m1-validation";tab.id="ls-m1-validation-tab";
 tab.textContent="M1 · 검증";
 tab.setAttribute("aria-controls","ls-m1-validation");
 nav.append(tab);
 const section=document.createElement("section");
 section.id="ls-m1-validation";section.className="ls-m1-validation";section.hidden=true;
 section.setAttribute("aria-label","M1 Step-index LP01 validation");
 section.innerHTML=`
 <div class="ls-m1-head"><div><small>VALIDATION / M1 · v0.3.0</small>
  <h2>Step-index FEM · LP₀₁ · COMSOL 공개값 검증</h2></div>
  <span class="ls-m1-badge">독립 검사 · 원본 설계 변경 없음</span></div>
 <p class="ls-m1-description">COMSOL 공식 step-index 기준(λ = 1550 nm, a = 8 µm, R = 40 µm, n₁ = 1.4457, n₂ = 1.4378)을 고정 입력으로 사용합니다.
  FEM–LP₀₁은 스칼라 간 비교이며, COMSOL의 1.4444는 벡터 HE₁₁ 결과를 소수 4자리로 공개한 값입니다.
  계산값이 반올림 범위에 들어가도 정밀 벡터 FEM 인증은 아닙니다.</p>
 <div class="ls-m1-controls">
  <button type="button" id="ls-m1-run">▶ 검증 28항목 실행</button>
  <button type="button" id="ls-m1-export" disabled>검증 보고서 JSON</button>
  <button type="button" id="ls-m1-hide">접기</button>
 </div>
 <p id="ls-m1-status" role="status" aria-live="polite">검증을 실행하면 이 표에 현재 솔버 계산 결과가 표시됩니다.</p>
 <div id="ls-m1-summary" class="ls-m1-summary" hidden></div>
 <div class="ls-m1-table-scroll">
  <table><thead><tr><th>그룹</th><th>검사</th><th>계산값</th><th>기준값</th><th>|오차|</th><th>허용오차</th><th>판정</th></tr></thead>
  <tbody id="ls-m1-results"><tr><td colspan="7">계산 대기</td></tr></tbody></table>
 </div>
 <details class="ls-m1-limits"><summary>검증 기준 · 가정 · 출처</summary>
  <p>FEM–해석해: |Δn_eff| ≤ ${TOL.lp01Neff}. COMSOL 공개 반올림값: ±${TOL.comsolRoundedNeff}.
   기존 회귀: neff ${TOL.neff}, MFD/Aeff ${TOL.mfdUm}, D ${TOL.dispersion}; FEM residual ≤ ${TOL.residual}.</p>
  <p>해석해는 원형 step-index, 약유도 스칼라 LP₀₁, 무한 클래딩을 가정합니다. 실제 FEM은 R = 40 µm에서 Dirichlet 경계이며,
   COMSOL은 HE₁₁ 벡터 모델입니다. 따라서 오차가 0이 될 필요는 없습니다.</p>
  <p>기존 G.654/G.652/Graded/Trench/Multilayer 5개 프로파일은 저장된 v0.2 수치 스냅샷과 비교합니다.
   별도의 전자기 벡터 모드, HCF, MCF 다중코어 물리모델, PML, 굽힘 손실은 M1 검증 대상이 아닙니다.</p>
  <a href="${esc(COMSOL_PUBLIC.source)}" target="_blank" rel="noopener noreferrer">COMSOL Step-Index Fiber · Application ID 145 (공식 문서)</a>
 </details>`;
 nav.after(section); // New section, without rearranging existing UI cards.
 const tbody=section.querySelector("#ls-m1-results");
 const status=section.querySelector("#ls-m1-status");
 const runBtn=section.querySelector("#ls-m1-run");
 const exportBtn=section.querySelector("#ls-m1-export");
 const show=()=>{section.hidden=false;tab.setAttribute("aria-current","true");};
 const hide=()=>{section.hidden=true;tab.removeAttribute("aria-current");};
 tab.addEventListener("click",event=>{event.preventDefault();show();section.scrollIntoView({behavior:"smooth",block:"start"});});
 section.querySelector("#ls-m1-hide").addEventListener("click",hide);
 runBtn.addEventListener("click",async()=>{
  if(runBtn.disabled)return;
  show();runBtn.disabled=true;exportBtn.disabled=true;last=null;
  status.textContent="검증 중 · 입력 설계와 기존 결과는 변경하지 않습니다.";
  tbody.innerHTML='<tr><td colspan="7">계산 중…</td></tr>';
  try{
   // Yield once so the busy status is painted before the synchronous FEM batch.
   await new Promise(resolve=>setTimeout(resolve,0));
   const report=await runM1Validation({onProgress:state=>{status.textContent="검증 중 · "+state.step;}});
   last=report;
   tbody.innerHTML=report.rows.map(r=>'<tr class="'+(r.passed?"is-pass":"is-fail")+'">'+
    '<td>'+esc(r.group)+'</td><td title="'+esc(r.note)+'">'+esc(r.test)+'</td>'+
    '<td>'+fmt(r.value)+'</td><td>'+fmt(r.reference)+'</td><td>'+fmt(r.absError)+'</td>'+
    '<td>'+fmt(r.tolerance)+'</td><td>'+ (r.passed?"PASS":"FAIL")+'</td></tr>').join("");
   const summary=section.querySelector("#ls-m1-summary");
   summary.hidden=false;
   summary.textContent=report.passedCount+" / "+report.totalCount+" 통과 · "+
    (report.passed?"정의된 M1 수치 기준 만족":"일부 검사 실패 — 항목별 오차 확인")+
    " · COMSOL 공개값은 반올림 정밀도 기준만 적용";
   status.textContent="검증 완료 · "+new Date(report.generatedAt).toLocaleString("ko-KR")+" · 입력 설계 및 기존 계산 결과는 보존되었습니다.";
   exportBtn.disabled=false;
  }catch(error){
   status.textContent="검증 실행 실패: "+(error?.message||error);
   tbody.innerHTML='<tr><td colspan="7">검증 실패 — 계산 결과를 확인하세요.</td></tr>';
  }finally{runBtn.disabled=false;}
 });
 exportBtn.addEventListener("click",()=>{
  if(!last)return;
  const blob=new Blob([JSON.stringify(last,null,2)],{type:"application/json"});
  const url=URL.createObjectURL(blob),a=document.createElement("a");
  a.href=url;a.download="ls_wave_optics_m1_validation.json";a.click();
  setTimeout(()=>URL.revokeObjectURL(url),1000);
 });
 globalThis.LSWaveM1Validation=Object.freeze({enabled:true,run:runM1Validation,latest:()=>last,show,hide});
}
if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",init,{once:true});
else init();
```

## 5. `validation_ui.css`

```css
/* M1 optional validation section. Existing workbench styles are untouched. */
#ls-m1-validation[hidden]{display:none!important}
.ls-m1-validation{scroll-margin-top:74px;margin:16px 0 21px;padding:20px;
 border:1px solid #386b83;border-radius:13px;background:#10243a;color:#e8f6fd;
 font:12px "Noto Sans KR","Malgun Gothic",sans-serif}
.ls-m1-head{display:flex;justify-content:space-between;align-items:center;gap:14px;flex-wrap:wrap}
.ls-m1-head small{color:#77cdd6;letter-spacing:.07em;font-size:10px}
.ls-m1-head h2{font-size:18px;margin:6px 0 10px;color:#f0f8ff}
.ls-m1-badge{font-size:11px;color:#a2d5d8;border:1px solid #3c6e78;border-radius:8px;padding:7px 10px}
.ls-m1-description,.ls-m1-limits p{color:#b7ccdd;line-height:1.7;margin:10px 0}
.ls-m1-controls{display:flex;flex-wrap:wrap;gap:9px;margin:14px 0}
.ls-m1-controls button{cursor:pointer;border:1px solid #3f7c9c;border-radius:8px;padding:9px 12px;
 background:#194467;color:#eef9ff;font-size:12px;font-weight:650}
.ls-m1-controls button:hover:not(:disabled){background:#276489}
.ls-m1-controls button:disabled{cursor:default;opacity:.48}
#ls-m1-status{min-height:20px;font-size:12px;color:#c1d4e2;margin:10px 0}
.ls-m1-summary{padding:10px 12px;border:1px solid #39726f;border-radius:7px;color:#b8f0df;font-size:12px;margin:9px 0}
.ls-m1-table-scroll{overflow:auto;max-height:500px;border:1px solid #31556e;border-radius:7px}
.ls-m1-validation table{width:100%;border-collapse:collapse;min-width:850px;text-align:left;font-size:11px}
.ls-m1-validation th,.ls-m1-validation td{padding:9px 8px;border-bottom:1px solid #2e5065;white-space:nowrap}
.ls-m1-validation th{position:sticky;top:0;z-index:1;background:#193952;color:#c7e5f2}
.ls-m1-validation td{color:#d5e9f2;font-variant-numeric:tabular-nums}
.ls-m1-validation .is-pass td:last-child{color:#8ee1bd;font-weight:750}
.ls-m1-validation .is-fail td:last-child{color:#ffafa9;font-weight:750}
.ls-m1-limits{margin:14px 0 0;padding:8px 11px;border-top:1px solid #34576e}
.ls-m1-limits summary{cursor:pointer;color:#9cdbed;font-size:12px}
.ls-m1-limits a{color:#7adadd}
@media(max-width:750px){.ls-m1-validation{padding:13px}.ls-m1-head h2{font-size:16px}}
```

## 6. `tests/test_m1_validation.mjs`

```javascript
/**
 * M1 independent oracle + legacy numerical regression.
 * Run: node tests/test_m1_validation.mjs
 * Also run existing: node tests/test_solver.mjs
 */
import assert from "node:assert/strict";
import {PRESETS,solve} from "../solver.mjs";
import {lp01Characteristic} from "../validation_lp01.mjs";
import {COMSOL_PUBLIC,LP01_CASES,TOL} from "../validation_cases.mjs";
import {fixtureModel,validateMesh,runM1Validation} from "../validation_engine.mjs";

assert.equal(COMSOL_PUBLIC.reference_neff_rounded,1.4444);
assert.equal(COMSOL_PUBLIC.core_radius_um,8);
assert.equal(COMSOL_PUBLIC.outer_radius_um,40);
assert.equal(COMSOL_PUBLIC.n_core,1.4457);
assert.equal(COMSOL_PUBLIC.n_clad,1.4378);
for(const c of LP01_CASES){
 const analytic=lp01Characteristic(c);
 assert.ok(Math.abs(analytic.n_eff-c.expectedAnalytic)<=TOL.analyticVsSciPy,c.id+" analytic vs independently computed SciPy");
 assert.ok(Math.abs(analytic.V-c.expectedV)<1e-11,c.id+" V");
 assert.ok(Math.abs(analytic.residual)<1e-9,c.id+" characteristic-equation residual");
 assert.ok(analytic.n_eff>c.n_clad&&analytic.n_eff<c.n_core,c.id+" guided interval");
 const m=fixtureModel(c);
 const fem=solve(m);
 assert.ok(Math.abs(fem.n_eff-analytic.n_eff)<=TOL.lp01Neff,c.id+" FEM vs LP01");
 assert.ok(fem.residual<=TOL.residual,c.id+" FEM relative residual");
 assert.equal(PRESETS.comsol_step.elements,750,"Presets must not be mutated");
}
assert.throws(()=>lp01Characteristic({...LP01_CASES[0],n_core:1.43}),/Core index/);
assert.throws(()=>lp01Characteristic({...LP01_CASES[0],core_radius_um:0}),/positive/);
const mesh=validateMesh();
assert.ok(mesh.rows[0].passed);
assert.equal(mesh.meshes.length,4);
for(let i=1;i<mesh.meshes.length;i++)
 assert.ok(mesh.meshes[i].analyticAbsError<mesh.meshes[i-1].analyticAbsError,"Mesh accuracy should improve");
const result=await runM1Validation();
assert.equal(result.schemaVersion,1);
assert.equal(result.kind,"M1-validation-report");
assert.equal(result.totalCount,28);
assert.equal(result.passedCount,28);
assert.equal(result.passed,true);
assert.ok(result.rows.some(row=>row.group==="COMSOL 공개값"&&row.passed));
assert.ok(result.rows.some(row=>row.group==="FEM vs LP01"&&row.passed));
assert.equal(result.rows.filter(row=>row.group==="기존 회귀").length,20);
const comsolFem=solve({...PRESETS.comsol_step});
assert.ok(Math.abs(comsolFem.n_eff-COMSOL_PUBLIC.reference_neff_rounded)<=TOL.comsolRoundedNeff);
assert.equal(result.details[0].femNeff,comsolFem.n_eff);
console.log("PASS M1: 28/28 numeric comparisons, independent LP01 oracle, COMSOL rounding, legacy baselines and mesh trend.");
```

## 7. `tests/browser_m1_validation.mjs`

```javascript
/**
 * Browser M1 smoke + old solver UI non-regression (requires Playwright).
 * TEST_BASE_URL can override http://127.0.0.1:8765/COMSOL/universal_silica_v01/web/
 */
import assert from "node:assert/strict";
import {createRequire} from "node:module";
const {chromium}=createRequire(import.meta.url)("playwright");
const browser=await chromium.launch({headless:true,args:["--no-sandbox","--disable-dev-shm-usage"]});
try{
 const page=await browser.newPage({viewport:{width:1550,height:1000},acceptDownloads:true});
 const errors=[];page.on("pageerror",e=>errors.push(e.message));
 await page.goto(process.env.TEST_BASE_URL||"http://127.0.0.1:8765/COMSOL/universal_silica_v01/web/",{waitUntil:"load",timeout:90000});
 // Startup wizard can cover the workbench; it must not be modified by M1.
 if(await page.locator('[data-wizard="cancel"]').isVisible())await page.locator('[data-wizard="cancel"]').click();
 const tab=page.locator("#ls-m1-validation-tab");
 assert.ok(await tab.isVisible(),"M1 nav shortcut available by default");
 assert.ok(await page.locator("#ls-m1-validation").isHidden(),"Validation section starts hidden");
 const oldNeff=await page.locator("#neff").textContent();
 const oldPreset=await page.locator("#preset").inputValue();
 const oldRadius=await page.locator("#core_radius_um").inputValue();
 await tab.click();
 assert.ok(await page.locator("#ls-m1-validation").isVisible());
 await page.locator("#ls-m1-run").click();
 await page.waitForFunction(()=>window.LSWaveM1Validation?.latest()?.totalCount===28,undefined,{timeout:90000});
 assert.equal(await page.locator("#ls-m1-results tr").count(),28);
 const report=await page.evaluate(()=>window.LSWaveM1Validation.latest());
 assert.equal(report.passed,true);
 assert.equal(report.passedCount,28);
 assert.equal(await page.locator("#preset").inputValue(),oldPreset,"M1 must not switch presets");
 assert.equal(await page.locator("#core_radius_um").inputValue(),oldRadius,"M1 must not change active design");
 assert.equal(await page.locator("#neff").textContent(),oldNeff,"M1 must not change original FEM cards");
 const [download]=await Promise.all([
  page.waitForEvent("download"),
  page.locator("#ls-m1-export").click()
 ]);
 assert.equal(download.suggestedFilename(),"ls_wave_optics_m1_validation.json");
 await page.locator("#ls-m1-hide").click();
 assert.equal(await page.locator("#ls-m1-validation").isHidden(),true);
 // Legacy controls remain available.
 assert.ok(await page.locator("#calculateBtn").isEnabled());
 assert.ok(await page.locator("#meshBtn").isEnabled());
 assert.ok(await page.locator("#sweepBtn").isEnabled());
 assert.deepEqual(errors,[]);
 console.log("PASS M1 browser: 28 rows, report download, all original controls and model values unchanged");
}finally{await browser.close();}
```

## 8. `tests/m1_lp01_scipy_reference.py`

```python
"""Independent cross-check of the browser M1 LP01 Bessel root.

Local use only: python tests/m1_lp01_scipy_reference.py
Requires numpy and scipy in the developer's validation environment.
No Python/Colab dependency is added to the deployed GitHub Pages site.
"""
from __future__ import annotations
import math
from scipy.optimize import brentq
from scipy.special import j0,j1,k0,k1

CASES = (
    ("comsol_step", 1550.0, 8.0, 1.4457, 1.4378, 1.4444018444092133),
    ("single_mode", 1550.0, 4.1, 1.4472, 1.4440, 1.4448631289955670),
)
COMSOL_PUBLIC_ROUNDED=1.4444

def lp01(wavelength_nm:float, a_um:float, n_core:float, n_clad:float):
    k0=2*math.pi/(wavelength_nm/1000)
    v=k0*a_um*math.sqrt(n_core*n_core-n_clad*n_clad)
    limit=min(v,2.404825557695773)*(1-1e-12)
    def equation(u):
        w=math.sqrt(v*v-u*u)
        return u*j1(u)/j0(u)-w*k1(w)/k0(w)
    u=brentq(equation,1e-10,limit,xtol=1e-13,rtol=1e-14)
    n_eff=math.sqrt(n_core*n_core-(u/(k0*a_um))**2)
    return {"V":v,"u":u,"w":math.sqrt(v*v-u*u),"n_eff":n_eff,
            "characteristic_residual":equation(u)}

if __name__=="__main__":
    for name,lam,a,n1,n2,expected in CASES:
        actual=lp01(lam,a,n1,n2)
        delta=abs(actual["n_eff"]-expected)
        assert delta<=1e-10,(name,delta)
        print(f"{name}: neff={actual['n_eff']:.13f}, V={actual['V']:.10f}, |Δ|={delta:.3e} [PASS]")
        if name=="comsol_step":
            error=abs(actual["n_eff"]-COMSOL_PUBLIC_ROUNDED)
            assert error<=5e-5
            print(f"COMSOL rounded comparison: |Δ|={error:.3e} [within 4-decimal rounding interval]")
```

---

## Existing-file minimal diffs
See [M1_VALIDATION.md](M1_VALIDATION.md#3-existing-file-diffs-only) and [CHANGELOG.md](CHANGELOG.md).
