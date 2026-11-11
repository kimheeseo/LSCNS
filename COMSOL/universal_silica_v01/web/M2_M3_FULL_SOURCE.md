# LS Wave Optics M2 + M3 v0.5.0 — Full copy-ready source files

This file contains the **entire committed content** of all newly added M2/M3 JS/CSS/tests. Copy each fenced block into its named path under `web/`. Existing-file minimal diffs are provided in [M2_M3_VALIDATION.md](M2_M3_VALIDATION.md); original numerical solver routines remain unchanged.

## 1. `material_models.mjs`

```javascript
/**
 * M2 optional material dispersion library. Wavelength nm; Sellmeier resonances um.
 * Does NOT modify scalar solver default offset mode.
 * Binary composition interpolation is a material model, not a measured product fit.
 */
export const MATERIAL_SCHEMA_VERSION=1;
export const SELLMEIER=Object.freeze({
 silica:{B:[.69616630,.40794260,.89747940],L:[.0684043,.1162414,9.896161],maxPercent:0,source:"ITU-T G Suppl.47 (2025), Table I.1"},
 geO2:{B:[.80686642,.71815848,.85416831],L:[.068972606,.15396605,11.841931],maxPercent:100,source:"Fleming (1984), Applied Optics 23, 4486"},
 f1:{B:[.69325,.39720,.86008],L:[.06724,.11714,9.7761],percent:1,source:"Fluorine doped silica, Applied Optics 22, 3102 (1983), 1%"},
 f2:{B:[.67744,.40101,.87193],L:[.06135,.12030,9.8563],percent:2,source:"Fluorine doped silica, Applied Optics 22, 3102 (1983), 2%"},
 p9:{B:[.695790,.452497,.712513],L:[.061568,.119921,8.656641],percent:9.1,source:"Optical Fibers: Signal Attenuation and Dispersion, 9.1 P2O5:90.0 SiO2 fit"}
});
export const MATERIAL_MODES=Object.freeze(["offset","geO2","f","p2o5","table"]);
function finite(v,name){if(!Number.isFinite(v))throw Error(name+" must be finite.");return v;}
export function sellmeierIndex(coeff,wavelengthNm){
 const w=finite(Number(wavelengthNm),"wavelength")/1000;
 if(!(w>=.4&&w<=1.7))throw Error("Doped Sellmeier calibration is restricted to 400–1700 nm in M2.");
 const sum=coeff.B.reduce((a,b,i)=>a+b*w*w/(w*w-coeff.L[i]**2),1);
 if(!(sum>0))throw Error("Invalid Sellmeier expression.");
 return Math.sqrt(sum);
}
function mixCoeff(left,right,f){
 return {B:left.B.map((v,i)=>v*(1-f)+right.B[i]*f),
  L:left.L.map((v,i)=>v*(1-f)+right.L[i]*f)};
}
/**
 * GeO2: interpolate between pure SiO2 and pure GeO2 coefficients 0–30 mol%.
 * F: interpolate SiO2 -> measured 1% -> measured 2% compositions.
 * P2O5: SiO2 -> measured 9.1% composition; no extrapolation.
 */
export function compositionIndex(material,wavelengthNm){
 const mode=material?.mode||"offset";
 if(mode==="offset")return sellmeierIndex(SELLMEIER.silica,wavelengthNm)+finite(Number(material?.deltaN??0),"deltaN");
 if(mode==="table")return interpolateTable(material.table,wavelengthNm).n;
 const percent=finite(Number(material.percent),"dopant concentration");
 if(percent<0)throw Error("Dopant concentration cannot be negative.");
 if(mode==="geO2"){
  if(percent>30)throw Error("GeO2 coefficient interpolation limited to <=30 mol% in M2.");
  return sellmeierIndex(mixCoeff(SELLMEIER.silica,SELLMEIER.geO2,percent/100),wavelengthNm);
 }
 if(mode==="f"){
  if(percent>2)throw Error("F interpolation uses only measured 0–2% composition.");
  const coeff=percent<=1?mixCoeff(SELLMEIER.silica,SELLMEIER.f1,percent):
   mixCoeff(SELLMEIER.f1,SELLMEIER.f2,percent-1);
  return sellmeierIndex(coeff,wavelengthNm);
 }
 if(mode==="p2o5"){
  if(percent>9.1)throw Error("P2O5 interpolation limited to 0–9.1 mol%.");
  return sellmeierIndex(mixCoeff(SELLMEIER.silica,SELLMEIER.p9,percent/9.1),wavelengthNm);
 }
 throw Error("Unsupported material model "+mode);
}
export function parseIndexTable(text){
 if(typeof text!=="string"||text.length>200000)throw Error("n(lambda) table requires text <=200KB.");
 const rows=[];
 for(const raw of text.split(/\r?\n/)){
  const line=raw.trim();if(!line||line.startsWith("#"))continue;
  const cols=line.split(/[,;\t]+/).map(s=>s.trim());
  if(cols[0].toLowerCase().includes("wavelength"))continue;
  if(cols.length<2||cols.length>3)throw Error("Table columns: wavelength_nm,n_real[,kappa].");
  const r=cols.map(Number);
  if(!r.every(Number.isFinite)||r[0]<=0||r[1]<1||r[1]>3||(r[2]??0)<0)throw Error("Invalid wavelength, n, or kappa in table.");
  rows.push({wavelength_nm:r[0],n:r[1],kappa:r[2]??0});
 }
 if(rows.length<2||rows.length>2000)throw Error("Table needs 2–2000 sorted rows.");
 for(let i=1;i<rows.length;i++)if(rows[i].wavelength_nm<=rows[i-1].wavelength_nm)throw Error("Table wavelengths must strictly increase.");
 return rows;
}
export function interpolateTable(rows,wavelengthNm){
 if(!Array.isArray(rows)||rows.length<2)throw Error("n(lambda) table not loaded.");
 const w=finite(Number(wavelengthNm),"wavelength");
 if(w<rows[0].wavelength_nm||w>rows.at(-1).wavelength_nm)throw Error("Table wavelength outside measured range; no extrapolation.");
 let lo=0,hi=rows.length-1;
 while(hi-lo>1){const m=(lo+hi)>>1;if(rows[m].wavelength_nm>w)hi=m;else lo=m;}
 const a=rows[lo],b=rows[hi],f=(w-a.wavelength_nm)/(b.wavelength_nm-a.wavelength_nm);
 return {n:a.n+(b.n-a.n)*f,kappa:a.kappa+(b.kappa-a.kappa)*f};
}
export function materialIndex(material,wavelengthNm){
 if(material?.mode==="table")return interpolateTable(material.table,wavelengthNm);
 const n=compositionIndex(material,wavelengthNm);
 const kappa=finite(Number(material?.kappa??0),"kappa");
 if(kappa<0||kappa>0.02)throw Error("kappa must lie in [0,0.02].");
 return {n,kappa};
}
export function materialPairAt(spec,wavelengthNm){
 const core=materialIndex(spec.core,wavelengthNm),clad=materialIndex(spec.clad,wavelengthNm);
 if(!(core.n>clad.n))throw Error("Guided step-index requires n_core > n_clad.");
 return {core,clad,deltaN:core.n-clad.n,wavelength_nm:wavelengthNm};
}
/** Separate JSON schema: does not replace existing {version,model} design file. */
export function upgradeMaterialSpec(o){
 if(!o||typeof o!=="object")throw Error("Material spec must be an object.");
 if(o.schemaVersion===1)return structuredClone(o);
 if(o.schemaVersion!==undefined)throw Error("Unknown materials schema version.");
 return {schemaVersion:MATERIAL_SCHEMA_VERSION,core:o.core??{mode:"offset",deltaN:.005},
  clad:o.clad??{mode:"offset",deltaN:0}};
}
/** M2 only: scalar step-index wrapper. Historical solve() is never modified.
 * Complex n is returned as first-order absorption estimate, not complex eigenmode.
 */
export function perturbativeAbsorptionDbKm(solution,pair){
 const field=solution.field,r=solution.r,coreRadius=solution.core_radius_um;
 if(!Number.isFinite(coreRadius))throw Error("Provide core_radius_um in material mode.");
 let pc=0,pt=0;
 for(let i=0;i<r.length-1;i++){
  const x=(r[i]+r[i+1])/2,d=r[i+1]-r[i],v=(field[i]+field[i+1])/2,weight=v*v*x*d;
  pt+=weight;if(x<coreRadius)pc+=weight;
 }
 const fraction=pt>0?Math.min(1,Math.max(0,pc/pt)):0;
 const kappaEff=fraction*pair.core.kappa+(1-fraction)*pair.clad.kappa;
 // alpha [dB/km] = (4*pi*kappa/lambda[m])*1000*(10/log(10)).
 return {absorption_dB_km:4*Math.PI*kappaEff/(pair.wavelength_nm*1e-9)*1000*10/Math.log(10),
  core_power_fraction:fraction,approximation:"weak-loss first-order intensity weighting; not complex-neff eigenmode"};
}
```

## 2. `material_m2_ui.mjs`

```javascript
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
 '<label>Dopant concentration (mol%)<input id="m2-'+prefix+'-percent" type="number" value="1" min="0" step=".1"/></label>'+
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
 '<p class="eng-note">GeO₂ 계수는 fused GeO₂ 기준 농도 보간(0–30 mol%), F는 1%/2% 실측계수 보간(0–2%), P₂O₅는 9.1% 참고 조성 보간(0–9.1%)입니다. 혼합 조성의 실제 도핑·제조공정·흡수 데이터 검증값이 아닙니다. 현재 400–1700 nm 범위로 제한합니다.</p>'+
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
```

## 3. `material_m2_ui.css`

```css
#ls-m2[hidden]{display:none!important}
#ls-m2{scroll-margin-top:80px}
#ls-m2 .m2-controls,#ls-m2 .m2-material-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px;margin:14px 0}
#ls-m2 .m2-controls{grid-template-columns:repeat(4,minmax(0,1fr))}
#ls-m2 .m2-material{border:1px solid #426880;border-radius:9px;padding:12px;min-width:0}
#ls-m2 .m2-material legend{font-size:12px;color:#9ce5de}
#ls-m2 label{display:flex;flex-direction:column;gap:5px;font-size:11px;color:#b9d5e8;min-width:0}
#ls-m2 .m2-material-fields{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:9px;margin:10px 0}
#ls-m2 input,#ls-m2 select,#ls-m2 textarea{width:100%;min-width:0;color:#f0f8ff;background:#12263e;border:1px solid #496780;border-radius:5px;padding:8px;font-size:12px}
#ls-m2 select{font-size:11px}
#ls-m2 details{margin:11px 0}
#ls-m2 summary{cursor:pointer;color:#b0d9e4}
#ls-m2 details p{font-size:11px;color:#a5c3d0}
#ls-m2 .m2-actions{display:flex;align-items:center;gap:9px;flex-wrap:wrap;margin:12px 0}
#ls-m2 button,#ls-m2 .m2-file{display:inline-block;background:#21516a;border:1px solid #3e90a6;border-radius:8px;color:#eefcff;padding:10px 13px;cursor:pointer;font-size:12px}
#ls-m2 button:disabled{opacity:.45;cursor:not-allowed}
#ls-m2 #m2-message{margin:13px 0;color:#b8e8d2;font-size:12px}
#ls-m2 #m2-results{overflow:auto}
#ls-m2 .m2-table{min-width:760px;width:100%;border-collapse:collapse;color:#d9edf8;font-size:11px}
#ls-m2 .m2-table td,#ls-m2 .m2-table th{padding:10px;border:1px solid #385871}
#ls-m2 .m2-table th{color:#b7dbe8;background:#143149}
@media(max-width:950px){#ls-m2 .m2-controls,#ls-m2 .m2-material-fields{grid-template-columns:repeat(2,minmax(0,1fr))}}
@media(max-width:750px){#ls-m2 .m2-material-grid,#ls-m2 .m2-controls{grid-template-columns:1fr}}
```

## 4. `vector_fdfd_browser.mjs`

```javascript
/**
 * M3 2D coupled transverse H-field full-vector FDFD (browser / JS).
 * Port of full_vector/vector_modes.py's Maxwell Hx-Hy operator.
 * Isotropic lossless z-invariant dielectrics, real epsilon only, Dirichlet
 * square boundary. No PML, radiation attenuation, anisotropy, or complex n.
 * Uses matrix-free BiCGSTAB shifted-inverse iteration with residual checks.
 * Fields: Ex/Ey real and Ez imaginary, consistent for e^(i beta z-i omega t).
 */
export const VECTOR_ENGINE_VERSION="0.5.0";
const dot=(a,b)=>{let t=0;for(let i=0;i<a.length;i++)t+=a[i]*b[i];return t;};
function norm(a){return Math.sqrt(dot(a,a));}
function finite(v,k){if(!Number.isFinite(v))throw Error(k+" must be finite");return v;}
function defaults(cfg){
 const a={wavelength_nm:1550,profile:"step",n_core:1.449023621703261,n_clad:1.444023621703261,
 core_radius_um:4.1,window_halfwidth_um:18,grid_n:45,graded_alpha:2,
 trench_start_um:6,trench_end_um:9,trench_delta_n:-.001,shells:[],...cfg};
 for(const k of ["wavelength_nm","core_radius_um","window_halfwidth_um","n_core","n_clad"])finite(a[k],k);
 if(!(a.wavelength_nm>400&&a.wavelength_nm<1700&&a.n_core>a.n_clad&&a.n_clad>1))
  throw Error("Requires real guided core/cladding indexes and 400–1700 nm.");
 if(!(a.window_halfwidth_um>2.3*a.core_radius_um))throw Error("FDFD window halfwidth > 2.3 × core radius.");
 if(!Number.isInteger(a.grid_n)||a.grid_n<29||a.grid_n>91||a.grid_n%2!==1)throw Error("grid_n must be odd integer [29,91].");
 if(!["step","graded","trench","multilayer"].includes(a.profile))throw Error("Unsupported radial profile.");
 return a;
}
function indexAt(c,r){
 if(c.profile==="graded")return c.n_clad+(r<c.core_radius_um?(c.n_core-c.n_clad)*Math.max(0,1-(r/c.core_radius_um)**c.graded_alpha):0);
 if(c.profile==="trench")return r<c.core_radius_um?c.n_core:r>=c.trench_start_um&&r<c.trench_end_um?c.n_clad+c.trench_delta_n:c.n_clad;
 if(c.profile==="multilayer"){
  for(const [limit,delta] of c.shells)if(r<limit)return c.n_clad+delta;
  return c.n_clad;
 }
 return r<c.core_radius_um?c.n_core:c.n_clad;
}
export function makeVectorGrid(cfg){
 const c=defaults(cfg),n=c.grid_n,h=2*c.window_halfwidth_um/(n-1),K=n*n;
 const index=new Float64Array(K),lnEps=new Float64Array(K);
 for(let j=0;j<n;j++)for(let i=0;i<n;i++){
  const x=-c.window_halfwidth_um+i*h,y=-c.window_halfwidth_um+j*h;
  const nn=indexAt(c,Math.hypot(x,y)),id=j*n+i;
  if(!(nn>=1&&nn<=3))throw Error("Unphysical refractive index in grid.");
  index[id]=nn;lnEps[id]=Math.log(nn*nn);
 }
 return {config:c,n,h,index,lnEps};
}
function operator(g){
 const {n,h,index,lnEps,config}=g,k=n-2,L=k*k,N=2*L,inv2h=1/(2*h),invh2=1/(h*h);
 const eps=new Float64Array(L),gx=new Float64Array(L),gy=new Float64Array(L);
 const k0=2*Math.PI/(config.wavelength_nm/1000),sigma=(k0*index.reduce((a,b)=>Math.max(a,b),0))**2+.02;
 for(let j=0;j<k;j++)for(let i=0;i<k;i++){
  const id=j*k+i,full=(j+1)*n+i+1;
  eps[id]=index[full]**2;
  gx[id]=(lnEps[full+1]-lnEps[full-1])*inv2h;
  gy[id]=(lnEps[full+n]-lnEps[full-n])*inv2h;
 }
 const diag=new Float64Array(N);
 for(let q=0;q<L;q++)diag[q]=diag[L+q]=sigma+4*invh2-k0*k0*eps[q];
 const lap=(x,off,id,i,j)=>{
  const center=x[off+id];
  return ((i>0?x[off+id-1]:0)+(i<k-1?x[off+id+1]:0)+
   (j>0?x[off+id-k]:0)+(j<k-1?x[off+id+k]:0)-4*center)*invh2;
 };
 const derivative=(x,off,id,i,j,axis)=>{
  if(axis==="x")return ((i<k-1?x[off+id+1]:0)-(i>0?x[off+id-1]:0))*inv2h;
  return ((j<k-1?x[off+id+k]:0)-(j>0?x[off+id-k]:0))*inv2h;
 };
 function applyM(x,out){
  for(let j=0;j<k;j++)for(let i=0;i<k;i++){
   const q=j*k+i,C=derivative(x,0,q,i,j,"y")-derivative(x,L,q,i,j,"x");
   const V=k0*k0*eps[q];
   out[q]=lap(x,0,q,i,j)+V*x[q]-gy[q]*C;
   out[L+q]=lap(x,L,q,i,j)+V*x[L+q]+gx[q]*C;
  }
  return out;
 }
 function applyA(x,out){
  applyM(x,out);
  for(let i=0;i<N;i++)out[i]=sigma*x[i]-out[i];
  return out;
 }
 return {applyM,applyA,N,L,k,k0,sigma,diag};
}
/** Diagonally preconditioned BiCGSTAB; fails explicitly instead of fabricating a result. */
function bicgstab(op,b,{tolerance=1e-8,maxIterations=850}={}){
 const N=b.length,x=new Float64Array(N),r=Float64Array.from(b),shadow=Float64Array.from(b),
  p=new Float64Array(N),v=new Float64Array(N),tmp=new Float64Array(N),
  ph=new Float64Array(N),sh=new Float64Array(N),s=new Float64Array(N),t=new Float64Array(N);
 const scale=Math.max(1e-30,norm(b));let rho=1,alpha=1,omega=1;
 for(let iter=0;iter<maxIterations;iter++){
  const next=dot(shadow,r);
  if(Math.abs(next)<1e-35)throw Error("FDFD linear BiCGSTAB breakdown (rho).");
  const beta=(next/rho)*(alpha/omega);rho=next;
  for(let i=0;i<N;i++)p[i]=r[i]+beta*(p[i]-omega*v[i]);
  for(let i=0;i<N;i++)ph[i]=p[i]/op.diag[i];
  op.applyA(ph,v);const denom=dot(shadow,v);
  if(Math.abs(denom)<1e-35)throw Error("FDFD linear BiCGSTAB breakdown (alpha).");
  alpha=rho/denom;
  for(let i=0;i<N;i++)s[i]=r[i]-alpha*v[i];
  if(norm(s)/scale<tolerance){
   for(let i=0;i<N;i++)x[i]+=alpha*ph[i];
   return {x,iterations:iter+1,error:norm(s)/scale};
  }
  for(let i=0;i<N;i++)sh[i]=s[i]/op.diag[i];
  op.applyA(sh,t);
  const tt=dot(t,t);if(tt<1e-35)throw Error("FDFD linear BiCGSTAB breakdown (omega).");
  omega=dot(t,s)/tt;
  if(Math.abs(omega)<1e-25)throw Error("FDFD linear BiCGSTAB stagnated.");
  for(let i=0;i<N;i++){x[i]+=alpha*ph[i]+omega*sh[i];r[i]=s[i]-omega*t[i];}
  if(norm(r)/scale<tolerance)return {x,iterations:iter+1,error:norm(r)/scale};
 }
 throw Error("FDFD shifted inverse linear solver did not converge.");
}
function eigenMode(op,g,polarization,orthogonal=[],{maxOuter=105,onProgress=()=>{}}={}){
 const {k,L,N,k0}=op,h=g.h,c=g.config;
 const x=new Float64Array(N),prev=new Float64Array(N),Mx=new Float64Array(N);
 for(let j=0;j<k;j++)for(let i=0;i<k;i++){
  const u=(i+1-k/2)*h,v=(j+1-k/2)*h,rad2=u*u+v*v;
  const q=j*k+i;const amp=Math.exp(-rad2/(c.core_radius_um*c.core_radius_um*1.2));
  x[polarization==="x"?q:L+q]=amp*(1+.02*u/c.core_radius_um);
 }
 const orth=y=>{
  for(const mode of orthogonal){const projection=dot(y,mode);for(let i=0;i<N;i++)y[i]-=projection*mode[i];}
  return y;
 };
 orth(x);let xnorm=norm(x);if(xnorm<1e-10)throw Error("Initial polarization vector lost during orthogonalization.");
 for(let i=0;i<N;i++)x[i]/=xnorm;
 let eig=NaN,residual=Infinity,converged=false,innerIterations=0,outer=0;
 for(outer=0;outer<maxOuter;outer++){
  const rhs=Float64Array.from(x);
  const sol=bicgstab(op,rhs,{tolerance:2e-9,maxIterations:900});
  innerIterations+=sol.iterations;
  const y=orth(sol.x);const ynorm=norm(y);
  if(!(ynorm>1e-16))throw Error("Vector inverse iteration collapsed.");
  for(let i=0;i<N;i++)x[i]=y[i]/ynorm;
  op.applyM(x,Mx);eig=dot(x,Mx);
  let res=0;for(let i=0;i<N;i++){const d=Mx[i]-eig*x[i];res+=d*d;}
  residual=Math.sqrt(res)/Math.max(Math.abs(eig),1e-30);
  if(outer%10===0)onProgress({mode:polarization,iteration:outer,residual});
  if(residual<2.5e-5){converged=true;break;}
 }
 if(!(eig>0))throw Error("FDFD did not obtain positive beta^2.");
 return {field:x,eigenvalue:eig,neff:Math.sqrt(eig)/k0,
  residual,converged,outerIterations:outer+1,innerIterations};
}
function reconstruct(eigen,g,op){
 const {n,h,index,config}=g,k=n-2,L=k*k,N=n*n,k0=op.k0;
 const Hx=new Float64Array(N),Hy=new Float64Array(N),div=new Float64Array(N),
  ex=new Float64Array(N),ey=new Float64Array(N),ez=new Float64Array(N);
 for(let j=0;j<k;j++)for(let i=0;i<k;i++){
  const q=j*k+i,full=(j+1)*n+(i+1);
  Hx[full]=eigen.field[q];Hy[full]=eigen.field[L+q];
 }
 const derivative=(field,id,axis)=>{
  if(axis==="x")return ((id%n<n-1?field[id+1]:0)-(id%n>0?field[id-1]:0))/(2*h);
  return ((id+n<N?field[id+n]:0)-(id>=n?field[id-n]:0))/(2*h);
 };
 const beta=Math.sqrt(eigen.eigenvalue);
 for(let id=0;id<N;id++)div[id]=derivative(Hx,id,"x")+derivative(Hy,id,"y");
 let total=0,power2=0,xsum=0,ysum=0,core=0,edge=0,sumx=0,sumy=0,sumz=0;
 const I=new Float64Array(N),Ix=new Float64Array(N),Iy=new Float64Array(N),Iz=new Float64Array(N);
 for(let id=0;id<N;id++){
  const epsilon=index[id]*index[id];
  ex[id]=(beta*Hy[id]-derivative(div,id,"y")/beta)/(k0*epsilon);
  ey[id]=(derivative(div,id,"x")/beta-beta*Hx[id])/(k0*epsilon);
  ez[id]=(derivative(Hy,id,"x")-derivative(Hx,id,"y"))/(k0*epsilon);
  Ix[id]=ex[id]*ex[id];Iy[id]=ey[id]*ey[id];Iz[id]=ez[id]*ez[id];
  const intensity=Ix[id]+Iy[id]+Iz[id];I[id]=intensity;
  const x=-config.window_halfwidth_um+(id%n)*h,
   y=-config.window_halfwidth_um+Math.floor(id/n)*h;
  total+=intensity;power2+=intensity*intensity;
  xsum+=x*intensity;ysum+=y*intensity;
  if(Math.hypot(x,y)<=config.core_radius_um)core+=intensity;
  if(id%n<3||id%n>=n-3||Math.floor(id/n)<3||Math.floor(id/n)>=n-3)edge+=intensity;
  sumx+=Ix[id];sumy+=Iy[id];sumz+=Iz[id];
 }
 if(!(total>0&&power2>0))throw Error("No nonzero E field reconstructed.");
 const cx=xsum/total,cy=ysum/total;
 let moment=0,grad=0;
 for(let id=0;id<N;id++){
  const x=-config.window_halfwidth_um+(id%n)*h-cx,
   y=-config.window_halfwidth_um+Math.floor(id/n)*h-cy;
  moment+=(x*x+y*y)*I[id];
  const v=Math.sqrt(I[id]);
  const left=id%n>0?Math.sqrt(I[id-1]):0,right=id%n<n-1?Math.sqrt(I[id+1]):0,
   low=id>=n?Math.sqrt(I[id-n]):0,high=id+n<N?Math.sqrt(I[id+n]):0;
  grad+=((right-left)/(2*h))**2+((high-low)/(2*h))**2;
 }
 const map={Itotal:I,Ex:Ix,Ey:Iy,Ez:Iz};
 return {mode_index:0,neff:eigen.neff,residual:eigen.residual,iteration_converged:eigen.converged,
  MFD_2nd_um:2*Math.sqrt(2*moment/total),
  MFD_PII_um:grad>0?2*Math.sqrt(2*total/grad):NaN,
  Aeff_um2:total*total/power2*h*h,
  Ex_intensity_fraction:sumx/total,Ey_intensity_fraction:sumy/total,Ez_intensity_fraction:sumz/total,
  core_power_fraction:core/total,edge_power_fraction:edge/total,
  polarization:String(sumx>sumy?"Ex dominant":"Ey dominant"),
  field_maps:map,field_components:{Ex_real:ex,Ey_real:ey,Ez_imag:ez}};
}
export function solveVectorFDFD(input,{modes=2,onProgress=()=>{}}={}){
 const g=makeVectorGrid(input),op=operator(g),vectors=[],results=[];
 for(const pol of ["x","y"].slice(0,modes)){
  const m=eigenMode(op,g,pol,vectors,{onProgress});
  if(!m.converged)throw Error("FDFD eigenmode residual too large ("+m.residual.toExponential(2)+"); increase mesh/window or reduce contrast.");
  if(!(m.neff>g.config.n_clad+1e-6&&m.neff<g.config.n_core+.002))
   throw Error("FDFD mode is not guided; improve resolution/window.");
  vectors.push(m.field);results.push({...reconstruct(m,g,op),mode_index:results.length+1,
   eigenIterations:m.outerIterations,linearIterations:m.innerIterations});
 }
 return {config:g.config,grid_n:g.n,h_um:g.h,index_map:Array.from(g.index),modes:results,
  method:"Coupled transverse Hx Hy 2D full-vector FDFD, shifted inverse BiCGSTAB"};
}
export function vectorResultJSON(result,{stride=1}={}){
 const {grid_n:n,config:c,modes}=result;
 if(!Number.isInteger(stride)||stride<1)throw Error("field stride >=1 required");
 const points=[];for(let i=0;i<n;i+=stride)points.push(i);
 const mapToGrid=arr=>points.map(j=>points.map(i=>Number(arr[j*n+i].toPrecision(8))));
 const normalized=arr=>{const max=Math.max(1e-30,...arr);return mapToGrid(Float64Array.from(arr,v=>v/max));};
 return {schema:"LSCNS_vector_mode_v1",schemaVersion:1,
  method:result.method,input:{...c,grid_n:n},modes:result.modes.map((m,idx)=>{
   const {field_maps,field_components,...meta}=m;
   if(idx!==0)return meta;
   const {Itotal,Ex,Ey,Ez}=field_maps,max=Math.max(1e-30,...Itotal);
   return {...meta,field_map:{grid_size:points.length,
    x_um:points.map(i=>Number((-c.window_halfwidth_um+i*result.h_um).toFixed(6))),
    intensity_normalized_to_max_total:{
     Itotal:mapToGrid(Float64Array.from(Itotal,v=>v/max)),
     Ex:mapToGrid(Float64Array.from(Ex,v=>v/max)),
     Ey:mapToGrid(Float64Array.from(Ey,v=>v/max)),
     Ez:mapToGrid(Float64Array.from(Ez,v=>v/max))
    },index_map:mapToGrid(result.index_map),
    field_component_phase_convention:"Ex and Ey are real; Ez = i*Ez_imag; arbitrary global eigenmode phase",
    field_components:{Ex_real:mapToGrid(field_components.Ex_real),
      Ey_real:mapToGrid(field_components.Ey_real),Ez_imag:mapToGrid(field_components.Ez_imag)}}};
  }),validation:{domain:"finite square, Dirichlet at outer edge; NO PML",
   scope:"2D full-vector coupled Hx Hy FDFD, Ex Ey Ez derived from Maxwell; M3 numerical prototype",
   limitation:"Finite difference staircase index and finite window; loss and complex neff not computed"}};
}
```

## 5. `vector_fdfd_worker.mjs`

```javascript
/** M3 worker: keep expensive FDFD eigen-iteration off the UI thread. */
import {solveVectorFDFD,vectorResultJSON} from "./vector_fdfd_browser.mjs";
self.onmessage=event=>{
 const {id,input,modes=2}=event.data||{};
 try{
  const result=solveVectorFDFD(input,{modes,onProgress:detail=>self.postMessage({id,type:"progress",detail})});
  const json=vectorResultJSON(result,{stride:1});
  self.postMessage({id,type:"result",result:json});
 }catch(error){self.postMessage({id,type:"error",message:String(error?.message||error)});}
};
```

## 6. `vector_m3_ui.mjs`

```javascript
/**
 * M3 browser full-vector study, separate from existing Colab JSON viewer.
 * Feature flag: window.LSWaveFeatures.vectorM3=false before module import.
 * Worker is static-GitHub-Pages-compatible; no server/API/WASM dependency.
 */
const ENABLED=globalThis.LSWaveFeatures?.vectorM3!==false && globalThis.LSWaveFeatures?.fullVectorWeb!==false;
let worker=null,latest=null,busy=false,sequence=0;
const $=id=>document.getElementById(id);
const fmt=(n,d=8)=>Number.isFinite(n)?n.toFixed(d):"—";
function draw(canvas,matrix){
 if(!canvas||!Array.isArray(matrix))return;
 const n=matrix.length,ctx=canvas.getContext("2d"),size=canvas.width;
 const max=Math.max(1e-30,...matrix.flat().map(Number));
 ctx.fillStyle="#071527";ctx.fillRect(0,0,size,size);
 const pixel=size/n;
 for(let j=0;j<n;j++)for(let i=0;i<n;i++){
  const ratio=Math.pow(Math.max(0,Math.min(1,Number(matrix[j][i])/max)),.55);
  ctx.fillStyle="rgb("+Math.round(10+235*ratio)+","+Math.round(32+160*ratio)+","+Math.round(50+70*(1-ratio))+")";
  ctx.fillRect(i*pixel,j*pixel,pixel+1,pixel+1);
 }
}
function init(){
 if(!ENABLED||$("ls-m3"))return;
 const parent=$("full-vector-tool");if(!parent)return;
 const section=document.createElement("section");section.id="ls-m3";section.className="ls-m3-panel";
 section.innerHTML='<h3>브라우저 내 Full-vector FDFD · M3 v0.5</h3>'+
 '<p>실제 결합 Hx·Hy 2D Maxwell 고유모드 계산 → Ex, Ey, Ez 복원. 기존 Colab JSON 불러오기 기능은 위에 그대로 남아 있습니다. 아래 해석은 PML/복소 굴절률/굽힘손실을 포함하지 않습니다.</p>'+
 '<div class="ls-m3-controls">'+
 '<label>Wavelength nm<input id="m3-lambda" type="number" value="1550" min="401" max="1699"/></label>'+
 '<label>Core radius µm<input id="m3-radius" type="number" value="4.1" step=".1"/></label>'+
 '<label>n_core<input id="m3-core-n" type="number" value="1.4490236217" step=".0001"/></label>'+
 '<label>n_clad<input id="m3-clad-n" type="number" value="1.4440236217" step=".0001"/></label>'+
 '<label>Half-width µm<input id="m3-half" type="number" value="18" step="1"/></label>'+
 '<label>FDFD grid<select id="m3-grid"><option>29</option><option>41</option><option selected>51</option><option>61</option><option>71</option></select></label>'+
 '</div>'+
 '<div class="ls-m3-actions"><button id="m3-run">▶ 브라우저 Full-vector 계산</button>'+
 '<button id="m3-use-m2">M2 재료 계산값 가져오기</button>'+
 '<button id="m3-json" disabled>Colab 호환 결과 JSON 저장</button></div>'+
 '<p id="m3-state" role="status">브라우저 계산 준비 · 별도 Web Worker (로컬 JS) · 2개 편광 모드</p>'+
 '<div id="m3-numbers"></div><div class="ls-m3-plots">'+
 '<article><h4>|E|²</h4><canvas width="240" height="240" id="m3-total"></canvas></article>'+
 '<article><h4>|Ex|²</h4><canvas width="240" height="240" id="m3-ex"></canvas></article>'+
 '<article><h4>|Ey|²</h4><canvas width="240" height="240" id="m3-ey"></canvas></article>'+
 '<article><h4>|Ez|²</h4><canvas width="240" height="240" id="m3-ez"></canvas></article></div>'+
 '<p class="ls-m3-warning">제한: 격자 경계는 Dirichlet, 각 성분 맵은 개별 최대값 기준 색상입니다. neff는 grid 및 window 수렴 검증이 필요합니다. PML과 손실은 M4 범위입니다. M2의 κ는 아직 이 실수 벡터 FDFD에 적용되지 않습니다.</p>';
 parent.querySelector(".fv-actions")?.after(section);
 const run=$("m3-run"),state=$("m3-state"),save=$("m3-json");
 run.addEventListener("click",()=>{
  if(busy)return;
  try{
   const input={
    profile:"step",wavelength_nm:Number($("m3-lambda").value),core_radius_um:Number($("m3-radius").value),
    window_halfwidth_um:Number($("m3-half").value),grid_n:Number($("m3-grid").value),
    n_core:Number($("m3-core-n").value),n_clad:Number($("m3-clad-n").value)
   };
   if(!globalThis.Worker)throw Error("Module Web Worker unavailable in this browser.");
   worker?.terminate();worker=new Worker(new URL("./vector_fdfd_worker.mjs",import.meta.url),{type:"module"});
   const id=++sequence;busy=true;latest=null;run.disabled=true;save.disabled=true;
   state.textContent="2D Maxwell FDFD 고유값 계산 중…";
   worker.onmessage=event=>{
    const packet=event.data;if(packet?.id!==id)return;
    if(packet.type==="progress"){
     state.textContent="FDFD "+packet.detail.mode+" 편광 · "+packet.detail.iteration+"회 · 고유값 상대잔차 "+packet.detail.residual.toExponential(2);
     return;
    }
    busy=false;run.disabled=false;
    if(packet.type==="error"){state.textContent="FDFD 실패: "+packet.message;worker.terminate();return;}
    latest=packet.result;save.disabled=false;worker.terminate();
    const rows=latest.modes;
    $("m3-numbers").innerHTML='<table><thead><tr><th>Mode</th><th>neff</th><th>Eigen residual</th><th>Ex energy</th><th>Ey energy</th><th>Ez energy</th><th>MFD µm</th><th>Aeff µm²</th></tr></thead><tbody>'+
     rows.map(m=>'<tr><td>'+m.mode_index+'</td><td>'+fmt(m.neff,9)+'</td><td>'+m.residual.toExponential(2)+'</td>'+
     '<td>'+fmt(m.Ex_intensity_fraction,5)+'</td><td>'+fmt(m.Ey_intensity_fraction,5)+'</td><td>'+fmt(m.Ez_intensity_fraction,5)+'</td>'+
     '<td>'+fmt(m.MFD_2nd_um,3)+'</td><td>'+fmt(m.Aeff_um2,3)+'</td></tr>').join("")+'</tbody></table>';
    const field=rows[0].field_map.intensity_normalized_to_max_total;
    for(const [canvas,key] of [["m3-total","Itotal"],["m3-ex","Ex"],["m3-ey","Ey"],["m3-ez","Ez"]])draw($(canvas),field[key]);
    state.textContent="2개 편광 모드 계산 완료 · 스칼라 FEM과 다른 Maxwell 벡터 연산 · 실제 HE₁₁ 고정밀 검증은 별도 필요";
   };
   worker.onerror=e=>{busy=false;run.disabled=false;save.disabled=true;state.textContent="FDFD Worker 오류: "+(e.message||"unknown");worker?.terminate();};
   worker.postMessage({id,input,modes:2});
  }catch(error){busy=false;run.disabled=false;state.textContent="FDFD 입력/실행 실패: "+error.message;}
 });
 $("m3-use-m2").addEventListener("click",()=>{
  const result=globalThis.LSWaveMaterialM2?.last?.();
  if(!result?.result){state.textContent="M2에서 먼저 재료 n(λ) 계산을 실행하세요.";return;}
  const v=result.result,m=result.model;
  for(const [id,val] of [["m3-lambda",m.wavelength_nm],["m3-radius",m.core_radius_um],["m3-core-n",v.n_core],["m3-clad-n",v.n_clad]])
   $(id).value=val;
  state.textContent="M2 실수 굴절률 가져옴 · κ(흡수)는 M3 고유값 계산에 적용되지 않음.";
 });
 save.addEventListener("click",()=>{
  if(!latest)return;
  const blob=new Blob([JSON.stringify(latest,null,2)],{type:"application/json"});
  const url=URL.createObjectURL(blob),a=document.createElement("a");a.href=url;
  a.download="ls_wave_full_vector_browser_m3.json";a.click();
  setTimeout(()=>URL.revokeObjectURL(url),1000);
 });
 globalThis.LSWaveVectorM3=Object.freeze({enabled:true,latest:()=>latest});
}
if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",init,{once:true});else init();
```

## 7. `vector_m3_ui.css`

```css
/* M3 extra section nested within existing Colab Full-vector viewer */
#ls-m3{margin:18px 0;padding:17px;background:#0d2336;border:1px solid #335f7e;border-radius:11px}
#ls-m3 h3{color:#dcf7ff;font-size:17px;margin:3px 0 10px}
#ls-m3 p{color:#b3d0e1;font-size:12px;line-height:1.65}
#ls-m3 .ls-m3-controls{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:9px;margin:15px 0}
#ls-m3 label{display:flex;flex-direction:column;gap:6px;font-size:11px;color:#bad7e6}
#ls-m3 input,#ls-m3 select{width:100%;min-width:0;padding:9px;background:#162e48;color:#ebfaff;border:1px solid #416980;border-radius:6px}
#ls-m3 .ls-m3-actions{display:flex;gap:8px;flex-wrap:wrap;margin:12px 0}
#ls-m3 button{background:#215a76;border:1px solid #4599aa;color:#f0fdff;border-radius:8px;padding:9px 13px;font-size:12px;cursor:pointer}
#ls-m3 button:disabled{opacity:.5;cursor:not-allowed}
#ls-m3 #m3-state{min-height:24px;font-size:12px;color:#97decf}
#ls-m3 #m3-numbers{overflow:auto;margin:12px 0}
#ls-m3 #m3-numbers table{min-width:780px;width:100%;border-collapse:collapse;font-size:11px;color:#d0ebf7}
#ls-m3 #m3-numbers td,#ls-m3 #m3-numbers th{padding:9px;border:1px solid #2b526c}
#ls-m3 .ls-m3-plots{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px}
#ls-m3 .ls-m3-plots article{min-width:0;background:#071b2b;padding:8px;border:1px solid #294a62;border-radius:8px}
#ls-m3 .ls-m3-plots h4{font-size:11px;color:#cee9f5;margin:4px 0}
#ls-m3 .ls-m3-plots canvas{width:100%;max-width:240px;height:auto;display:block;image-rendering:auto}
#ls-m3 .ls-m3-warning{color:#f1c391}
@media(max-width:900px){#ls-m3 .ls-m3-plots{grid-template-columns:repeat(2,minmax(0,1fr))}}
@media(max-width:680px){#ls-m3 .ls-m3-controls{grid-template-columns:repeat(2,minmax(0,1fr))}}
```

## 8. `tests/test_material_m2.mjs`

```javascript
import assert from "node:assert/strict";
import {SELLMEIER,sellmeierIndex,compositionIndex,parseIndexTable,interpolateTable,
 materialIndex,materialPairAt,upgradeMaterialSpec} from "../material_models.mjs";
import {PRESETS,solve} from "../solver.mjs";
const silica=sellmeierIndex(SELLMEIER.silica,1550);
assert.ok(Math.abs(silica-1.444023621703261)<1e-12);
assert.ok(Math.abs(compositionIndex({mode:"offset",deltaN:.005},1550)-(silica+.005))<1e-12);
assert.ok(compositionIndex({mode:"geO2",percent:6.3},1550)>silica);
assert.ok(compositionIndex({mode:"f",percent:1},1550)<silica);
assert.ok(compositionIndex({mode:"f",percent:2},1550)<compositionIndex({mode:"f",percent:1},1550));
assert.ok(compositionIndex({mode:"p2o5",percent:9.1},1550)>silica);
assert.throws(()=>compositionIndex({mode:"geO2",percent:40},1550),/30/);
assert.throws(()=>compositionIndex({mode:"f",percent:2.1},1550),/2%/);
assert.throws(()=>compositionIndex({mode:"p2o5",percent:10},1550),/9.1/);
assert.throws(()=>sellmeierIndex(SELLMEIER.silica,2000),/400/);
const rows=parseIndexTable("wavelength_nm,n,kappa\n1500,1.5,0.000001\n1600,1.4,0.000003");
assert.equal(rows.length,2);
assert.ok(Math.abs(interpolateTable(rows,1550).n-1.45)<1e-12);
assert.ok(Math.abs(interpolateTable(rows,1550).kappa-2e-6)<1e-14);
assert.throws(()=>interpolateTable(rows,1700),/range/);
assert.throws(()=>parseIndexTable("1550,1.4\n1500,1.5"),/increase/);
const spec=upgradeMaterialSpec({core:{mode:"offset",deltaN:.005},clad:{mode:"offset",deltaN:0}});
assert.equal(spec.schemaVersion,1);
const pair=materialPairAt(spec,1550);
assert.ok(Math.abs(pair.deltaN-.005)<1e-12);
const legacy=solve(PRESETS.g652);
const newInput={...PRESETS.g652,clad_index_override:pair.clad.n,delta_n_core:pair.deltaN};
const newFEM=solve(newInput);
assert.ok(Math.abs(legacy.n_eff-newFEM.n_eff)<1e-12,"default M2 offset should reproduce original FEM");
assert.throws(()=>materialIndex({mode:"offset",kappa:-1},1550),/kappa/);
console.log("PASS M2: references, concentration bounds, n(lambda) table, schema and legacy default preservation");
```

## 9. `tests/test_vector_m3.mjs`

```javascript
/**
 * M3 browser vector Maxwell operator regression against the *existing* Python
 * full_vector/results/vector_mode_result.json baseline (same 69x69 mesh).
 * Run with Node JS, independent of Colab and network.
 */
import assert from "node:assert/strict";
import {silicaIndex} from "../solver.mjs";
import {solveVectorFDFD,makeVectorGrid,vectorResultJSON} from "../vector_fdfd_browser.mjs";
const ncl=silicaIndex(1550),dn=.0034785793846155;
const input={profile:"step",wavelength_nm:1550,core_radius_um:5.5,n_clad:ncl,n_core:ncl+dn,
 window_halfwidth_um:22,grid_n:69};
assert.equal(makeVectorGrid(input).index.length,69*69);
const result=solveVectorFDFD(input,{modes:2});
assert.equal(result.modes.length,2);
const PYTHON_NEFF=1.445711888509771; // Existing Python 2D H-field FDFD result, same grid/window/profile.
for(const m of result.modes){
 assert.ok(m.neff>ncl&&m.neff<ncl+dn);
 assert.ok(Math.abs(m.neff-PYTHON_NEFF)<2e-6,"FDFD agreement with existing Python same-grid eigs");
 assert.ok(m.residual<2.5e-5,"Eigenpair relative residual");
 assert.ok(m.Ez_intensity_fraction>0&&m.Ez_intensity_fraction<.05);
 assert.ok(Math.abs(m.Ex_intensity_fraction+m.Ey_intensity_fraction+m.Ez_intensity_fraction-1)<1e-10);
 assert.ok(m.Aeff_um2>0&&m.MFD_2nd_um>0);
 assert.ok(m.edge_power_fraction<.01);
}
assert.ok(result.modes[0].polarization!==result.modes[1].polarization,"Polarization partners distinct");
assert.ok(Math.abs(result.modes[0].neff-result.modes[1].neff)<5e-5,"Circular fiber pair near-degenerate");
const portable=vectorResultJSON(result);
assert.equal(portable.schema,"LSCNS_vector_mode_v1");
assert.equal(portable.schemaVersion,1);
assert.ok(Array.isArray(portable.modes[0].field_map.intensity_normalized_to_max_total.Ex));
assert.ok(portable.modes[0].field_map.intensity_normalized_to_max_total.Ez.length===69);
assert.equal(portable.modes[0].field_map.field_components.Ex_real.length,69);
assert.equal(portable.modes[0].field_map.field_components.Ey_real.length,69);
assert.equal(portable.modes[0].field_map.field_components.Ez_imag.length,69);
assert.ok(portable.modes[0].field_map.field_component_phase_convention.includes("Ez = i"));
assert.throws(()=>makeVectorGrid({...input,grid_n:30}),/odd/);
assert.throws(()=>makeVectorGrid({...input,n_core:ncl-.001}),/guided/);
console.log("PASS M3: real 2D coupled Maxwell Hx-Hy FDFD, Ex/Ey/Ez, two polarizations and Python same-grid neff");
```

## 10. `tests/browser_m2_m3_ui.mjs`

```javascript
/**
 * Optional browser E2E M2 & M3; run from repo root with HTTP server at :8765.
 * Requires playwright package and Chromium. Tests no mutation of original UI.
 */
import assert from "node:assert/strict";
import {createRequire} from "node:module";
const {chromium}=createRequire(import.meta.url)("playwright");
const browser=await chromium.launch({headless:true,args:["--no-sandbox","--disable-dev-shm-usage"]});
try{
 const page=await browser.newPage({viewport:{width:1520,height:970}});
 const errors=[];page.on("pageerror",e=>errors.push(e.message));
 await page.goto(process.env.TEST_BASE_URL||"http://127.0.0.1:8765/COMSOL/universal_silica_v01/web/",{waitUntil:"load"});
 if(await page.locator('[data-wizard="cancel"]').isVisible())await page.locator('[data-wizard="cancel"]').click();
 assert.ok(await page.locator("#ls-m2-nav").isVisible());
 const initialNeff=await page.locator("#neff").textContent();
 const initialProfile=await page.locator("#profile").inputValue();
 await page.locator("#ls-m2-nav").click();
 await page.locator("#m2-run").click();
 await page.waitForFunction(()=>window.LSWaveMaterialM2?.last()?.result?.n_eff!==undefined);
 const m2=await page.evaluate(()=>window.LSWaveMaterialM2.last());
 assert.ok(Number.isFinite(m2.result.n_eff));
 assert.ok(Math.abs(m2.result.delta_n-.005)<1e-9);
 assert.equal(await page.locator("#neff").textContent(),initialNeff);
 assert.equal(await page.locator("#profile").inputValue(),initialProfile);
 await page.locator("#preset").selectOption("vector");
 assert.ok(await page.locator("#ls-m3").isVisible());
 assert.ok(await page.locator("#fv-json-file").count()===1,"Existing Colab file import retained");
 await page.locator("#m3-run").click();
 await page.waitForFunction(()=>window.LSWaveVectorM3?.latest()?.modes?.length===2,undefined,{timeout:90000});
 const r=await page.evaluate(()=>window.LSWaveVectorM3.latest());
 assert.equal(r.schema,"LSCNS_vector_mode_v1");
 assert.equal(r.modes.length,2);
 assert.ok(r.modes.every(m=>m.residual<2.5e-5));
 assert.equal(await page.locator("#m3-numbers tr").count(),3);
 assert.deepEqual(errors,[]);
 console.log("PASS M2/M3 browser optional section: original scalar state unchanged, M2 and worker FDFD output");
}finally{await browser.close();}
```

## Minimal existing code patches

See [M2_M3_VALIDATION.md](M2_M3_VALIDATION.md#minimal-changes-to-original-files) and [CHANGELOG.md](CHANGELOG.md).
