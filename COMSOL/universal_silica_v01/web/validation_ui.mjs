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
