/**
 * v0.4 M2 optional, strictly isolated material-model lab.
 * Preserves legacy scalar FEM, original app state and JSON.
 */
import {materialPairAt,parseIndexTable,upgradeMaterialSpec,perturbativeAbsorptionDbKm} from "./material_models.mjs";
import {PRESETS,solve} from "./solver.mjs";
const FEATURES=Object.freeze({materialM2:globalThis.LSWaveFeatures?.materialM2!==false});
const $=(id)=>document.getElementById(id);
const esc=v=>String(v).replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const opts=[["offset","Pure silica + Δn (legacy default)"],["geO2","GeO₂–SiO₂ Sellmeier interpolation"],["f","F–SiO₂ measured 1%/2% interpolation"],["p2o5","P₂O₅–SiO₂ 9.1% reference interpolation"],["table","User n(λ) table"]];
const fmt=n=>Number.isFinite(n)?n.toPrecision(10):"—";
let last=null;
function buildMaterial(prefix,defaultD){
 return '<fieldset class="m2-material"><legend>'+prefix.toUpperCase()+' material</legend>'+
 '<label>Index model<select id="m2-'+prefix+'-mode">'+opts.map(([key,name])=>'<option value="'+key+'">'+esc(name)+'</option>').join("")+'</select></label>'+
 '<div class="m2-material-fields"><label>Δn offset<input id="m2-'+prefix+'-delta" type="number" value="'+defaultD+'" step=".001"/></label>'+
 '<label>Dopant (Ge/P: mol%; F: wt%)<input id="m2-'+prefix+'-percent" type="number" value="1" min="0" step=".1"/></label>'+
 '<label>Imaginary index κ (absorption only)<input id="m2-'+prefix+'-kappa" type="number" value="0" min="0" max=".02" step=".000001"/></label></div>'+
 '<details><summary>Custom wavelength, n, κ table (CSV/TSV)</summary><p>Sorted: wavelength_nm,n_real[,kappa]. At least 2 rows. No extrapolation.</p>'+
 '<textarea id="m2-'+prefix+'-table" spellcheck="false" rows="4" placeholder="1500,1.45,0&#10;1550,1.449,0&#10;1600,1.448,0"></textarea></details></fieldset>';
}
function settings(prefix){
 const mode=$("m2-"+prefix+"-mode").value;
 const material={mode};
 if(mode==="offset")material.deltaN=Number($("m2-"+prefix+"-delta").value);
 if(["geO2","f","p2o5"].includes(mode))material.percent=Number($("m2-"+prefix+"-percent").value);
 if(mode==="table")material.table=parseIndexTable($("m2-"+prefix+"-table").value);
 else material.kappa=Number($("m2-"+prefix+"-kappa").value);
 return material;
}
function init(){
 if(!FEATURES.materialM2)return;
 const nav=document.querySelector("nav.eng-shortcuts");if(!nav||$("ls-m2"))return;
 const button=document.createElement("a");button.id="ls-m2-nav";button.href="#ls-m2";button.textContent="M2 · 재료 모델";nav.append(button);
 const section=document.createElement("section");section.id="ls-m2";section.className="ls-m2-lab engineering-panel";section.hidden=true;
 section.innerHTML='<div class="eng-head"><div><span class="kicker">M2 · MATERIAL DISPERSION · v0.4</span><h2>재료 n(λ) / 도핑 / 흡수 실험</h2></div><span class="eng-tag">독립 계산 · 기본 Δn 유지</span></div>'+
 '<p class="eng-help">기존 Radial FEM, 결과 카드, 설계 JSON은 변경되지 않습니다. 아래의 선택적 step-index 재료 모델에서 계산한 실수 n(λ)만 기존 FEM에 전달합니다. κ는 1차 근사 감쇠 추정에만 사용합니다.</p>'+
 '<div class="m2-controls"><label>λ (nm)<input id="m2-lambda" type="number" value="1550" min="400" max="1700"/></label>'+
 '<label>Core radius (µm)<input id="m2-radius" type="number" value="4.1" min=".3" max="35" step=".1"/></label>'+
 '<label>Outer radius (µm)<input id="m2-window" type="number" value="40" min="15" max="100"/></label>'+
 '<label>Radial elements<input id="m2-elements" type="number" value="750" min="80" max="1500" step="50"/></label></div>'+
 '<div class="m2-material-grid">'+buildMaterial("core",.005)+buildMaterial("clad",0)+'</div>'+
 '<div class="m2-actions"><button id="m2-run">▶ n(λ) 및 Scalar FEM 비교 계산</button><button id="m2-json" disabled>재료 구성 JSON 저장</button><label class="m2-file">재료 JSON 불러오기<input type="file" id="m2-import" accept=".json" hidden/></label></div>'+
 '<div id="m2-message" role="status">실행 전 · 기본값은 Pure silica + Δn offset 입니다.</div><div id="m2-results"></div>'+
 '<p class="eng-note">GeO₂ 계수는 fused GeO₂ 기준 농도 보간(0–30 mol%), F는 1/2 wt% 실측계수 보간(0–2 wt%), P₂O₅는 9.1% 참고 조성 보간(0–9.1%)입니다. 혼합 조성의 실제 도핑·제조공정·흡수 데이터 검증값이 아닙니다. 현재 400–1700 nm 범위로 제한합니다.</p>'+
 '<button type="button" id="m2-hide">접기</button>';
 nav.after(section);
 const show=()=>{section.hidden=false;section.scrollIntoView({block:"start",behavior:"smooth"});};
 button.addEventListener("click",e=>{e.preventDefault();show();});
 $("m2-hide").addEventListener("click",()=>{section.hidden=true;});
 $("m2-run").addEventListener("click",()=>{
  $("m2-message").textContent="M2 재료 모델 및 Scalar FEM 계산 중…";
  try{
   const wavelength_nm=Number($("m2-lambda").value),core_radius_um=Number($("m2-radius").value),
    outer_radius_um=Number($("m2-window").value),elements=Number($("m2-elements").value);
   const spec={schemaVersion:1,core:settings("core"),clad:settings("clad")};
   const pair=materialPairAt(spec,wavelength_nm);
   const model={...PRESETS.g652,profile:"step",wavelength_nm,core_radius_um,outer_radius_um,elements,
    clad_index_override:pair.clad.n,delta_n_core:pair.deltaN};
   const r=solve(model);
   const absorption=perturbativeAbsorptionDbKm({...r,core_radius_um},pair);
   last={schemaVersion:1,kind:"ls-wave-optics-m2-material-spec",version:"0.4.0",spec,
    model:{wavelength_nm,core_radius_um,outer_radius_um,elements},result:{
    n_core:pair.core.n,n_clad:pair.clad.n,delta_n:pair.deltaN,n_eff:r.n_eff,
    MFD_PII_um:r.MFD_PII_um,Aeff_um2:r.Aeff_um2,
    absorption_estimate_db_per_km:absorption.absorption_dB_km,
    kappa_core:pair.core.kappa,kappa_clad:pair.clad.kappa}};
   const display=last.result;
   $("m2-results").innerHTML='<table class="m2-table"><thead><tr><th>n_core</th><th>n_clad</th><th>Δn</th><th>Scalar FEM neff</th><th>MFD (µm)</th><th>Aeff (µm²)</th><th>Estimated material loss (dB/km)</th></tr></thead><tbody><tr>'+
    ["n_core","n_clad","delta_n","n_eff","MFD_PII_um","Aeff_um2","absorption_estimate_db_per_km"].map(k=>'<td>'+fmt(display[k])+'</td>').join("")+'</tr></tbody></table>';
   $("m2-message").textContent="완료 · 별도 step-index 실수 굴절률 계산. κ 추정값은 복소 고유모드 손실이 아닙니다.";
   $("m2-json").disabled=false;
  }catch(error){
   $("m2-message").textContent="M2 계산 실패: "+error.message;
   $("m2-json").disabled=true;last=null;
  }
 });
 $("m2-json").addEventListener("click",()=>{
  if(!last)return;
  const blob=new Blob([JSON.stringify(last,null,2)],{type:"application/json"});
  const url=URL.createObjectURL(blob),a=document.createElement("a");a.href=url;a.download="ls_wave_m2_materials.json";a.click();
  setTimeout(()=>URL.revokeObjectURL(url),1000);
 });
 $("m2-import").addEventListener("change",async e=>{
  const f=e.target.files?.[0];if(!f)return;
  if(f.size>200000){$("m2-message").textContent="200 KB 초과";return;}
  try{
   const doc=JSON.parse(await f.text()),spec=upgradeMaterialSpec(doc.spec||doc);
   for(const prefix of ["core","clad"]){
    const m=spec[prefix];if(!opts.some(([key])=>key===m?.mode))throw Error(prefix+" unknown mode");
    $("m2-"+prefix+"-mode").value=m.mode;
    if(m.deltaN!==undefined)$("m2-"+prefix+"-delta").value=m.deltaN;
    if(m.percent!==undefined)$("m2-"+prefix+"-percent").value=m.percent;
    if(m.kappa!==undefined)$("m2-"+prefix+"-kappa").value=m.kappa;
    if(m.table)$("m2-"+prefix+"-table").value=m.table.map(r=>[r.wavelength_nm,r.n,r.kappa].join(",")).join("\n");
   }
   if(doc.model)for(const [k,id] of [["wavelength_nm","m2-lambda"],["core_radius_um","m2-radius"],["outer_radius_um","m2-window"],["elements","m2-elements"]]){
    if(doc.model[k]!==undefined)$(id).value=doc.model[k];
   }
   $("m2-message").textContent="재료 구성 불러옴 · 계산은 다시 실행해야 합니다.";
  }catch(error){$("m2-message").textContent="불러오기 실패: "+error.message;}
 });
 globalThis.LSWaveMaterialM2=Object.freeze({enabled:true,last:()=>last,show});
}
if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",init,{once:true});else init();
