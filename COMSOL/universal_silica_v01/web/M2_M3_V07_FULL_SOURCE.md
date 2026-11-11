# LS Wave Optics v0.7.0 — M2 & M3 full copy-ready source files

Complete committed source for the 11 new files below. For each file, create it under COMSOL/universal_silica_v01/web/ using the exact relative path. Existing code modifications were limited to metadata/import URLs and fluorine unit descriptions; see [M2_M3_V07_VALIDATION.md](M2_M3_V07_VALIDATION.md) for exact minimal diffs.

## 1. m2_preview_engine_v06.mjs

~~~javascript
/**
 * M2 v0.6.0 — independent optical materials PREVIEW, not coupled to an eigensolver.
 * All coefficients below have a public reference; uncertain F mol%-specific
 * coefficients are deliberately NOT invented.
 *
 * SiO2 + GeO2: Fleming (1984) Applied Optics 23, 4486, Table II
 *   https://doi.org/10.1364/AO.23.004486
 * Fluorine 1 and 2 WEIGHT percent (wt%, NOT mol%): Fleming & Wood (1983),
 *   Applied Optics 22, 3102, Table II; https://doi.org/10.1364/AO.22.003102
 * P2O5 9.1 mol% anchor: Optical Fiber Communication Systems with MATLAB
 *   and Simulink Models, Table 3.4 (reproduced); original citation therein
 *   Fleming, Electronics Letters 14, 326 (1978).
 *   https://www.routledge.com/rsc/downloads/Pages_from_K22108_Book_rev.pdf
 *
 * Intermediate concentrations obtained by linear Sellmeier-coefficient
 * interpolation are approximate MIXING RULES, NOT measured predictions.
 * Core/cladding original pure-silica + fixed Δn path is preserved exactly.
 */
import {silicaIndex} from "./solver.mjs";
import {SELLMEIER,sellmeierIndex,parseIndexTable,interpolateTable} from "./material_models.mjs";
export const M2_VERSION="0.6.0";
export const M2_SCHEMA=2;
export const M2_SOURCES=Object.freeze({
 silica:{label:"Malitson pure fused silica; Fleming 1984 Table II",url:"https://doi.org/10.1364/AO.23.004486"},
 geO2:{label:"Fleming 1984 Table II: pure GeO₂ & pure SiO₂; interpolated mixture",url:"https://doi.org/10.1364/AO.23.004486"},
 fwt:{label:"Fleming & Wood 1983 Table II: 1 and 2 wt% F (NOT mol%)",url:"https://doi.org/10.1364/AO.22.003102"},
 p2o5:{label:"Optical Fiber Communication Systems, Table 3.4: 9.1 mol% P₂O₅ mixture",url:"https://www.routledge.com/rsc/downloads/Pages_from_K22108_Book_rev.pdf"},
 fmol:{label:"No verified F mol%-specific Sellmeier mapping: user must enter six sourced coefficients",url:null},
 table:{label:"User supplied wavelength-index measurement table; not manufacturer verified",url:null}
});
export const M2_CATALOG=Object.freeze([
 {id:"silica",label:"Pure SiO₂",model:{kind:"silica"}},
 {id:"geO2",label:"GeO₂ 5 mol% · modeled",model:{kind:"geO2",molPercent:5}},
 {id:"f_wt",label:"F 1 wt% · measured fit",model:{kind:"f_wt",weightPercent:1}},
 {id:"p2o5",label:"P₂O₅ 5 mol% · modeled",model:{kind:"p2o5",molPercent:5}}
]);
const finite=(x,name)=>{const n=Number(x);if(!Number.isFinite(n)||x==null||x==="")throw Error(name+" must be finite.");return n};
const range=(v,lo,hi,name)=>{if(v<lo||v>hi)throw Error(name+" must be "+lo+"–"+hi+".");return v};
function mixCoefficients(a,b,w){
 return {B:a.B.map((v,i)=>(1-w)*v+w*b.B[i]),
  L:a.L.map((v,i)=>(1-w)*v+w*b.L[i])};
}
function validateUserCoefficients(c){
 if(!c||!Array.isArray(c.B)||!Array.isArray(c.L)||c.B.length!==3||c.L.length!==3)
  throw Error("Sourced custom Sellmeier requires B1,B2,B3,L1,L2,L3.");
 const B=c.B.map((v,i)=>finite(v,"B"+(i+1)));
 const L=c.L.map((v,i)=>finite(v,"L"+(i+1)));
 for(const v of L)if(v<=0)throw Error("Sellmeier resonance wavelengths must be positive µm.");
 return {B,L};
}
function pchipSlopes(x,y){
 const n=x.length,slopes=new Array(n),d=new Array(n-1),h=new Array(n-1);
 for(let i=0;i<n-1;i++){h[i]=x[i+1]-x[i];d[i]=(y[i+1]-y[i])/h[i]}
 if(n===2)return [d[0],d[0]];
 const endpoint=(h0,h1,d0,d1)=>{
  let m=((2*h0+h1)*d0-h0*d1)/(h0+h1);
  if(Math.sign(m)!==Math.sign(d0))m=0;
  else if(Math.sign(d0)!==Math.sign(d1)&&Math.abs(m)>3*Math.abs(d0))m=3*d0;
  return m;
 };
 slopes[0]=endpoint(h[0],h[1],d[0],d[1]);
 slopes[n-1]=endpoint(h[n-2],h[n-3],d[n-2],d[n-3]);
 for(let i=1;i<n-1;i++){
  if(d[i-1]===0||d[i]===0||Math.sign(d[i-1])!==Math.sign(d[i]))slopes[i]=0;
  else{
   const w1=2*h[i]+h[i-1],w2=h[i]+2*h[i-1];
   slopes[i]=(w1+w2)/(w1/d[i-1]+w2/d[i]);
  }
 }
 return slopes;
}
export function interpolateMeasured(rows,lambdaNm,method="linear"){
 if(!["linear","pchip"].includes(method))throw Error("Interpolation method must be linear or PCHIP.");
 if(!Array.isArray(rows)||rows.length<2)throw Error("Measurement table needs at least 2 points.");
 const x=rows.map(r=>r.wavelength_nm);
 for(let i=0;i<x.length;i++)if(!Number.isFinite(x[i])||(i&&x[i]<=x[i-1]))throw Error("Wavelengths must strictly increase.");
 const lambda=finite(lambdaNm,"λ");
 if(lambda<x[0]||lambda>x.at(-1))throw Error("Out of table range "+x[0]+"–"+x.at(-1)+" nm (NO extrapolation).");
 if(method==="linear")return interpolateTable(rows,lambda);
 let lo=0,hi=x.length-1;
 while(hi-lo>1){const m=(hi+lo)>>1;if(x[m]>lambda)hi=m;else lo=m;}
 const h=x[hi]-x[lo],t=(lambda-x[lo])/h,t2=t*t,t3=t2*t;
 const interp=key=>{
  const yy=rows.map(r=>finite(r[key]??(key==="kappa"?0:NaN),key));
  const deriv=pchipSlopes(x,yy);
  return (2*t3-3*t2+1)*yy[lo]+(t3-2*t2+t)*h*deriv[lo]+
   (-2*t3+3*t2)*yy[hi]+(t3-t2)*h*deriv[hi];
 };
 return {n:interp("n"),kappa:Math.max(0,interp("kappa"))};
}
export function kFromAlpha(alphaDbKm,lambdaNm){
 const alpha=range(finite(alphaDbKm,"α dB/km"),0,1e12,"α dB/km");
 const lambda=range(finite(lambdaNm,"λ nm"),1,1e7,"λ nm");
 // ñ=n+iκ; field ~ exp(i2πñz/λ); intensity decays exp(-4πκ z/λ).
 // α[dB/km] = (10/ln10)(4π κ/λ[m]) (1000 m/km).
 return alpha*Math.log(10)*lambda*1e-9/(4*Math.PI*1000*10);
}
export function alphaFromK(kappa,lambdaNm){
 const k=range(finite(kappa,"κ"),0,1,"κ");
 const lambda=range(finite(lambdaNm,"λ nm"),1,1e7,"λ nm");
 return 4*Math.PI*k/(lambda*1e-9)*1000*10/Math.log(10);
}
export function evaluateM2Material(material,lambdaNm){
 const cfg=material||{kind:"offset",deltaN:0};
 const kind=cfg.kind||"offset",lambda=finite(lambdaNm,"λ");
 let n;
 if(kind==="offset")n=silicaIndex(lambda)+finite(cfg.deltaN??0,"Δn");
 else if(kind==="silica")n=silicaIndex(lambda); // exact legacy solver implementation
 else if(kind==="geO2"){
  const p=range(finite(cfg.molPercent??0,"GeO₂ mol%"),0,30,"GeO₂ mol%");
  n=sellmeierIndex(mixCoefficients(SELLMEIER.silica,SELLMEIER.geO2,p/100),lambda);
 }else if(kind==="p2o5"){
  const p=range(finite(cfg.molPercent??0,"P₂O₅ mol%"),0,9.1,"P₂O₅ mol%");
  n=sellmeierIndex(mixCoefficients(SELLMEIER.silica,SELLMEIER.p9,p/9.1),lambda);
 }else if(kind==="f_wt"){
  const p=range(finite(cfg.weightPercent??0,"F wt%"),0,2,"F wt%");
  const ab=p<=1?[SELLMEIER.silica,SELLMEIER.f1,p]:[SELLMEIER.f1,SELLMEIER.f2,p-1];
  n=sellmeierIndex(mixCoefficients(...ab),lambda);
 }else if(kind==="f_mol"){
  const p=range(finite(cfg.molPercent??0,"F mol%"),0,100,"F mol%");
  if(p===0)n=silicaIndex(lambda);
  else{
   if(!cfg.coefficients?.source?.trim())throw Error("F mol% > 0 requires user Sellmeier coefficients AND a source; no wt%→mol% assumption.");
   n=sellmeierIndex(validateUserCoefficients(cfg.coefficients),lambda);
  }
 }else if(kind==="custom_sellmeier"){
  if(!cfg.coefficients?.source?.trim())throw Error("Custom Sellmeier coefficients require source/reference text.");
  n=sellmeierIndex(validateUserCoefficients(cfg.coefficients),lambda);
 }else if(kind==="table"){
  n=interpolateMeasured(cfg.table,lambda,cfg.interpolation||"linear").n;
 }else throw Error("Unsupported M2 model "+kind);
 if(!(n>1&&n<3))throw Error("Material index n outside supported 1–3 range.");
 const absorption=cfg.absorption||{kind:"none"};
 let kappa;
 if(absorption.kind==="k")kappa=range(finite(absorption.value,"κ"),0,1,"κ");
 else if(absorption.kind==="alpha")kappa=kFromAlpha(absorption.value,lambda);
 else if(absorption.kind==="none"||absorption.kind==null)
  kappa=kind==="table"?interpolateMeasured(cfg.table,lambda,cfg.interpolation||"linear").kappa:0;
 else throw Error("Unknown absorption mode.");
 if(!(Number.isFinite(kappa)&&kappa>=0))throw Error("Invalid κ.");
 return {n,kappa,alphaDbKm:alphaFromK(kappa,lambda),lambdaNm:lambda,
  complex:{real:n,imag:kappa},convention:"n_tilde=n+i*kappa; exp(i beta z - i omega t)"};
}
export function sampleM2Material(material,{startNm=1300,endNm=1650,count=71}={}){
 if(!Number.isInteger(count)||count<2||count>301)throw Error("Preview sample count 2–301.");
 const start=finite(startNm,"λ start"),end=finite(endNm,"λ end");
 if(!(end>start))throw Error("λ end must exceed λ start.");
 const samples=[];
 for(let i=0;i<count;i++){
  const lambda=start+(end-start)*i/(count-1);
  samples.push(evaluateM2Material(material,lambda));
 }
 return samples;
}
/**
 * M2 preview JSON lives separately from legacy design JSON and M2 v0.4 JSON.
 * Legacy F 'percent' in earlier UI was erroneously described as mol%.
 * Never map it silently to molPercent. Mark unresolved until confirmed.
 */
export function migrateM2Preview(src){
 const o=src?.spec&&src.kind==="ls-wave-optics-m2-material-spec"?src.spec:src;
 if(!o||typeof o!=="object")throw Error("M2 JSON object expected.");
 if(o.schemaVersion===2)return structuredClone(o);
 if(o.schemaVersion!=null&&o.schemaVersion!==1)throw Error("Unsupported M2 schemaVersion.");
 const convert=material=>{
  if(!material)return {kind:"offset",deltaN:0};
  const mode=material.kind??material.mode??"offset";
  if(mode==="f")return {kind:"f_legacy_unresolved",legacyPercent:material.percent,
   warning:"Legacy F percent ambiguity: source fit uses wt%, prior UI labeled mol%. Choose wt% or custom mol%-specific coefficients."};
  const kinds={offset:"offset",geO2:"geO2",p2o5:"p2o5",table:"table",silica:"silica"};
  if(!(mode in kinds))throw Error("Unknown legacy mode "+mode);
  return {kind:kinds[mode],deltaN:material.deltaN??0,
   molPercent:material.percent,table:material.table,
   interpolation:"linear",absorption:material.kappa!=null?{kind:"k",value:material.kappa}:{kind:"none"}};
 };
 return {schemaVersion:2,kind:"LSWaveM2MaterialPreview",version:M2_VERSION,
  core:convert(o.core),clad:convert(o.clad)};
}
export function makeM2PreviewDocument(core,clad){
 return {schemaVersion:M2_SCHEMA,kind:"LSWaveM2MaterialPreview",version:M2_VERSION,
  createdAt:new Date().toISOString(),core:structuredClone(core),clad:structuredClone(clad)};
}
~~~

## 2. m2_preview_ui_v06.mjs

~~~javascript
/**
 * M2 v0.6 optional, self-contained n(lambda)/k(lambda) preview.
 * No write to original scalar solver, M2 v0.4 editor, M3, or project JSON.
 * Feature flag: window.LSWaveFeatures.materialPreviewV06=false.
 */
import {M2_VERSION,M2_SOURCES,M2_CATALOG,evaluateM2Material,sampleM2Material,
 parseIndexTable,migrateM2Preview,makeM2PreviewDocument} from "./m2_preview_engine_v06.mjs";
const ENABLED=globalThis.LSWaveFeatures?.materialPreviewV06!==false;
const $=id=>document.getElementById(id);
const safe=s=>String(s??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const fmt=(v,d=10)=>Number.isFinite(v)?Number(v).toPrecision(d):"—";
const modes=[
 ["offset","기본값: SiO₂ + Δn (기존 계산식)"],
 ["silica","순수 SiO₂"],
 ["geO2","GeO₂ 도핑 (mol% 보간)"],
 ["f_wt","F 도핑: 논문 wt% 기준 (0–2)"],
 ["f_mol","F 도핑: mol% + 사용자 Sellmeier 계수"],
 ["p2o5","P₂O₅ 도핑 (mol% 보간)"],
 ["table","사용자 n(λ) CSV / TSV"],
 ["custom_sellmeier","사용자 Sellmeier (B1–B3, L1–L3)"]
];
const modelDefault={core:{kind:"offset",deltaN:0.005,absorption:{kind:"none"}},
 clad:{kind:"silica",absorption:{kind:"none"}}};
const state=structuredClone(modelDefault);
let latest=null;
function form(role){
 const label=role==="core"?"CORE / 코어":"CLADDING / 클래딩";
 return '<fieldset class="m2v6-material" data-role="'+role+'"><legend>'+label+'</legend>'+
 '<label>재료 모델<select id="m2v6-'+role+'-kind">'+modes.map(([v,t])=>
  '<option value="'+v+'" '+(v===state[role].kind?"selected":"")+'>'+safe(t)+'</option>').join("")+'</select></label>'+
 '<div class="m2v6-fields">'+
 '<label>Δn (기본 방식)<input id="m2v6-'+role+'-offset" type="number" step=".0001" value="'+(role==="core"?.005:0)+'"/></label>'+
 '<label>도핑 농도 (mol%)<input id="m2v6-'+role+'-mol" type="number" min="0" max="100" step=".1" value="5"/></label>'+
 '<label>F 실험농도 (wt%)<input id="m2v6-'+role+'-wt" type="number" min="0" max="2" step=".1" value="1"/></label>'+
 '<label>보간 방식<select id="m2v6-'+role+'-interp"><option value="linear">Linear</option><option value="pchip">PCHIP (shape-preserving)</option></select></label>'+
 '<label>흡수 입력 방식<select id="m2v6-'+role+'-abs"><option value="none">기본 / 테이블 k</option><option value="k">κ 직접 입력</option><option value="alpha">α (dB/km) 입력</option></select></label>'+
 '<label>κ 또는 α 입력값<input id="m2v6-'+role+'-absvalue" type="number" value="0" min="0" step="any"/></label>'+
 '</div><details id="m2v6-'+role+'-csvbox"><summary>사용자 n(λ) 테이블 (CSV 붙여넣기)</summary>'+
 '<p>열 순서: wavelength_nm,n_real[,kappa]. Linear/PCHIP 선택 가능. 범위 밖 외삽 금지.</p>'+
 '<textarea id="m2v6-'+role+'-csv" rows="4" spellcheck="false" placeholder="1300,1.450,0&#10;1550,1.445,0&#10;1700,1.443,0"></textarea></details>'+
 '<details id="m2v6-'+role+'-coeffbox"><summary>사용자 Sellmeier 계수 및 출처 (F mol%는 필수)</summary>'+
 '<p>3-term B1,B2,B3 / L1,L2,L3 [µm] (실측 또는 확인 가능한 출처를 직접 입력)</p>'+
 '<label>B1, B2, B3<input id="m2v6-'+role+'-B" placeholder="B1,B2,B3"/></label>'+
 '<label>L1, L2, L3 [µm]<input id="m2v6-'+role+'-L" placeholder="L1,L2,L3"/></label>'+
 '<label>계수 출처 / 실험 조건<input id="m2v6-'+role+'-source" placeholder="논문 DOI/표준/측정 보고서"/></label></details>'+
 '<p class="m2v6-hint" id="m2v6-'+role+'-hint">기본 Δn 모델 · 파장 독립 추가값</p></fieldset>';
}
function readCoefficients(role){
 const B=$("m2v6-"+role+"-B").value.split(/[,\s]+/).filter(Boolean).map(Number);
 const L=$("m2v6-"+role+"-L").value.split(/[,\s]+/).filter(Boolean).map(Number);
 return {B,L,source:$("m2v6-"+role+"-source").value.trim()};
}
function readMaterial(role){
 const root="m2v6-"+role+"-",kind=$(root+"kind").value;
 const m={kind,absorption:{kind:$(root+"abs").value}};
 if(m.absorption.kind!=="none"){
  const raw=$(root+"absvalue").value.trim();
  if(raw==="")throw Error(role+" κ/α value required.");
  m.absorption.value=Number(raw);
 }
 if(kind==="offset")m.deltaN=Number($(root+"offset").value);
 else if(kind==="geO2"||kind==="p2o5"||kind==="f_mol")m.molPercent=Number($(root+"mol").value);
 else if(kind==="f_wt")m.weightPercent=Number($(root+"wt").value);
 if(kind==="f_mol"||kind==="custom_sellmeier")m.coefficients=readCoefficients(role);
 if(kind==="table"){
  m.table=parseIndexTable($(root+"csv").value);
  m.interpolation=$(root+"interp").value;
 }
 return m;
}
function writeMaterial(role,m){
 if(m.kind==="f_legacy_unresolved"){
  $("m2v6-"+role+"-hint").textContent=m.warning;
  $("m2v6-"+role+"-kind").value="f_mol";
  return;
 }
 if(!modes.some(([k])=>k===m.kind))throw Error("Unknown material kind "+m.kind);
 const root="m2v6-"+role+"-";
 $(root+"kind").value=m.kind;
 if(m.deltaN!=null)$(root+"offset").value=m.deltaN;
 if(m.molPercent!=null)$(root+"mol").value=m.molPercent;
 if(m.weightPercent!=null)$(root+"wt").value=m.weightPercent;
 if(m.absorption){
  $(root+"abs").value=m.absorption.kind||"none";
  $(root+"absvalue").value=m.absorption.value??0;
 }
 if(m.table)$(root+"csv").value=m.table.map(x=>[x.wavelength_nm,x.n,x.kappa??0].join(",")).join("\n");
 if(m.interpolation)$(root+"interp").value=m.interpolation;
 if(m.coefficients){
  $(root+"B").value=(m.coefficients.B||[]).join(",");
  $(root+"L").value=(m.coefficients.L||[]).join(",");
  $(root+"source").value=m.coefficients.source||"";
 }
}
function provenance(role){
 const kind=$("m2v6-"+role+"-kind").value;
 const sourceKey=kind==="f_wt"?"fwt":kind==="f_mol"?"fmol":kind==="table"?"table":kind==="offset"?"silica":kind;
 const src=M2_SOURCES[sourceKey]||M2_SOURCES.table;
 const hints={
  offset:"SiO₂ Sellmeier + 파장 무관 Δn · 기존 기본 모델(변경 없음)",
  silica:"순수 SiO₂ · 기존 Sellmeier 함수와 완전 동일",
  geO2:"0–30 mol% · SiO₂/GeO₂ endpoint 계수의 선형 혼합(중간 농도 실측 검증 아님)",
  f_wt:"원문 1/2 wt% 실측 피팅값 · mol%와 혼용 금지",
  f_mol:"F mol% 사용자 지정 · 공식 mol%-계수 매핑은 확인되지 않아 직접 계수+출처 필요",
  p2o5:"0–9.1 mol% · 9.1 mol% 실측 기준 계수 보간(중간 농도 실측 검증 아님)",
  table:"CSV/PCHIP/선형 · 범위 밖 외삽 거부 · 사용자 데이터 책임",
  custom_sellmeier:"사용자 실측 Sellmeier 6계수와 출처 필요"
 };
 $("m2v6-"+role+"-hint").textContent=hints[kind]||"";
 const info=$("m2v6-"+role+"-sourceinfo");
 if(info)info.textContent=src.label;
}
function draw(canvas,series,field,title){
 const ctx=canvas.getContext("2d");if(!ctx)return;
 const {width:W,height:H}=canvas;ctx.clearRect(0,0,W,H);
 ctx.fillStyle="#0b1f31";ctx.fillRect(0,0,W,H);
 const margins={l:68,r:14,t:26,b:40},pw=W-margins.l-margins.r,ph=H-margins.t-margins.b;
 const all=series.flatMap(s=>s.rows.map(v=>v[field])).filter(Number.isFinite);
 if(!all.length){ctx.fillStyle="#c2d5e0";ctx.fillText("No valid samples",35,H/2);return;}
 let lo=Math.min(...all),hi=Math.max(...all),span=hi-lo;
 if(span<1e-15){const d=Math.max(Math.abs(hi)*.03,field==="kappa"?1e-12:.001);lo-=d;hi+=d}
 else{lo-=.08*span;hi+=.08*span}
 const xvalues=series.flatMap(s=>s.rows.map(v=>v.lambdaNm)),xmin=Math.min(...xvalues),xmax=Math.max(...xvalues);
 const xx=x=>margins.l+(x-xmin)/(xmax-xmin)*pw,yy=y=>margins.t+(hi-y)/(hi-lo)*ph;
 ctx.strokeStyle="#294659";ctx.lineWidth=1;ctx.font="11px Arial";ctx.fillStyle="#a9c5d7";
 for(let i=0;i<=4;i++){
  const y=margins.t+i*ph/4;
  ctx.beginPath();ctx.moveTo(margins.l,y);ctx.lineTo(W-margins.r,y);ctx.stroke();
  const v=hi-(hi-lo)*i/4;
  ctx.fillText(field==="kappa"?v.toExponential(1):v.toFixed(5),6,y+3);
 }
 for(let i=0;i<=4;i++){
  const x=margins.l+i*pw/4;
  ctx.beginPath();ctx.moveTo(x,margins.t);ctx.lineTo(x,H-margins.b);ctx.stroke();
  ctx.fillText(Math.round(xmin+i*(xmax-xmin)/4),x-12,H-20);
 }
 ctx.fillText("λ (nm)",W-58,H-5);ctx.fillText(title,margins.l,16);
 for(const s of series){
  ctx.strokeStyle=s.color;ctx.lineWidth=2;ctx.beginPath();
  s.rows.forEach((v,i)=>{if(i===0)ctx.moveTo(xx(v.lambdaNm),yy(v[field]));else ctx.lineTo(xx(v.lambdaNm),yy(v[field]));});
  ctx.stroke();
 }
}
function output(){
 const status=$("m2v6-status"),errors=[],start=Number($("m2v6-start").value),end=Number($("m2v6-end").value);
 const active=[];
 for(const role of ["core","clad"]){
  const model=readMaterial(role);
  state[role]=structuredClone(model);
  provenance(role);
  try{
   const rows=sampleM2Material(model,{startNm:start,endNm:end,count:71});
   const at1550=evaluateM2Material(model,1550);
   active.push({name:role==="core"?"CORE":"CLADDING",color:role==="core"?"#73d7f8":"#ffa77a",rows,at1550,role,model});
  }catch(e){errors.push((role==="core"?"Core":"Cladding")+": "+e.message)}
 }
 const refs=[];
 for(const [i,preset] of M2_CATALOG.entries()){
  try{
   const rows=sampleM2Material(preset.model,{startNm:start,endNm:end,count:71});
   refs.push({name:preset.label,color:["#9fc5d2","#93d5ad","#ba9dd2","#f1c68a"][i],rows,
    at1550:evaluateM2Material(preset.model,1550)});
  }catch(e){errors.push(preset.label+": "+e.message)}
 }
 const series=[...refs,...active],body=$("m2v6-rows");
 body.innerHTML=series.map(s=>'<tr><td>'+safe(s.name)+'</td>'+
  '<td>'+fmt(s.at1550.n,11)+'</td><td>'+s.at1550.kappa.toExponential(4)+'</td>'+
  '<td>'+fmt(s.at1550.alphaDbKm,7)+'</td><td>'+safe(s.role?"선택 재료":"참고곡선")+'</td></tr>').join("");
 draw($("m2v6-nchart"),series,"n","Real n(λ) · Sellmeier / table");
 draw($("m2v6-kchart"),series,"kappa","Imaginary κ(λ) · alpha/k preview");
 $("m2v6-legend").innerHTML=series.map(s=>'<span><i style="background:'+s.color+'"></i>'+safe(s.name)+'</span>').join("");
 latest=makeM2PreviewDocument(state.core,state.clad);
 latest.preview={rangeNm:[start,end],rows:series.map(s=>({name:s.name,at1550:s.at1550}))};
 status.textContent=errors.length?"미리보기 일부 실패 · "+errors.join(" / "):
  "미리보기 성공 · 1550 nm 표 및 n(λ)/κ(λ) 그래프. M3/Scalar FEM에는 적용하지 않음.";
 $("m2v6-save").disabled=false;
}
function download(){
 if(!latest)return;
 const blob=new Blob([JSON.stringify(latest,null,2)],{type:"application/json"}),url=URL.createObjectURL(blob);
 const a=document.createElement("a");a.href=url;a.download="LS_Wave_M2_v0.6_material_preview.json";
 document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),1000);
}
function mount(){
 if(!ENABLED||$("m2v6-preview"))return;
 const nav=document.querySelector("nav.eng-shortcuts");if(!nav)return;
 const tab=document.createElement("a");tab.id="m2v6-nav";tab.href="#m2v6-preview";
 tab.textContent="M2 · MATERIAL PREVIEW v0.6";nav.append(tab);
 const section=document.createElement("section");
 section.id="m2v6-preview";section.hidden=true;section.className="m2v6-preview engineering-panel";
 section.innerHTML='<header><small>M2 · MATERIAL MODEL PREVIEW · v'+M2_VERSION+
 '</small><h2>파장 의존 n(λ) · 도핑 Sellmeier · 복소 굴절률</h2>'+
 '<p>기존 SiO₂+Δn 계산은 기본값으로 유지. 아래 모델은 미리보기/JSON 전용이며 벡터/스칼라 고유값 솔버에 자동 연결되지 않습니다.</p></header>'+
 '<div class="m2v6-range"><label>파장 시작 (nm)<input id="m2v6-start" type="number" value="1300" min="400" max="1700"/></label>'+
 '<label>파장 종료 (nm)<input id="m2v6-end" type="number" value="1650" min="400" max="1700"/></label></div>'+
 '<div class="m2v6-materials">'+form("core")+form("clad")+'</div>'+
 '<div class="m2v6-actions"><button type="button" id="m2v6-run">▶ 재료 곡선 미리보기</button>'+
 '<button type="button" id="m2v6-save" disabled>재료 JSON 저장</button>'+
 '<label class="m2v6-import">재료 JSON 불러오기<input type="file" id="m2v6-open" accept=".json" hidden/></label></div>'+
 '<p id="m2v6-status" role="status">미리보기 실행 대기 · F wt%/mol% 단위 구분 필수</p>'+
 '<div id="m2v6-legend" class="m2v6-legend"></div>'+
 '<div class="m2v6-plots"><canvas id="m2v6-nchart" width="590" height="300" aria-label="n wavelength chart"></canvas>'+
 '<canvas id="m2v6-kchart" width="590" height="300" aria-label="k wavelength chart"></canvas></div>'+
 '<h3>1550 nm 재료 비교</h3><div class="m2v6-scroll"><table><thead><tr>'+
 '<th>재료 / 구성</th><th>n_real @1550</th><th>κ @1550</th><th>α (dB/km)</th><th>구분</th>'+
 '</tr></thead><tbody id="m2v6-rows"><tr><td colspan="5">계산 전</td></tr></tbody></table></div>'+
 '<details class="m2v6-provenance" open><summary>근거·단위·기능 한계</summary>'+
 '<p>SiO₂ 및 GeO₂: <a href="'+M2_SOURCES.geO2.url+'" target="_blank">Fleming, Applied Optics 23 (1984), Table II</a>; Ge 도핑은 단순 계수 보간 모델입니다.</p>'+
 '<p>F: <a href="'+M2_SOURCES.fwt.url+'" target="_blank">Fleming & Wood, Applied Optics 22 (1983), Table II</a>의 1/2 <b>wt%</b> 실험. 몰농도 대응 계수는 확인되지 않아 사용자 mol% 선택 시 B/L 6개와 출처를 필수 입력합니다.</p>'+
 '<p>P₂O₅: <a href="'+M2_SOURCES.p2o5.url+'" target="_blank">광섬유 교재 Table 3.4</a>의 9.1 mol% 참고 피팅. 중간 농도는 보간 추정이며 실측 검증 아님.</p>'+
 '<p>복소 굴절률은 ñ=n+iκ, exp(iβz−iωt) 사용. α(dB/km)=4πκ×(1000/λ[m])×10/ln(10). 흡수/전파 손실은 <b>미리보기만</b> 수행하며 고유모드 감쇠 계산 아님.</p>'+
 '<p>CSV 입력은 n(λ) 두 점 이상, Linear/PCHIP 선택. 파장 범위 밖 외삽 금지(경고). 구 M2 schemaVersion 1의 F 입력 단위 불명은 수동 확인을 요청합니다.</p></details>';
 nav.after(section);
 const show=()=>{section.hidden=false;section.scrollIntoView({block:"start",behavior:"smooth"});};
 tab.addEventListener("click",e=>{e.preventDefault();show()});
 $("m2v6-run").addEventListener("click",()=>{try{output()}catch(e){$("m2v6-status").textContent="미리보기 실패: "+e.message;}});
 $("m2v6-save").addEventListener("click",download);
 for(const role of ["core","clad"]) $("m2v6-"+role+"-kind").addEventListener("change",()=>provenance(role));
 $("m2v6-open").addEventListener("change",async e=>{
  const f=e.target.files?.[0];if(!f)return;
  if(f.size>250000){$("m2v6-status").textContent="재료 파일 제한 250KB";return}
  try{
   const doc=migrateM2Preview(JSON.parse(await f.text()));
   for(const role of ["core","clad"])writeMaterial(role,doc[role]);
   $("m2v6-status").textContent="재료 JSON 불러옴 · 계산은 다시 실행. F 단위 모호한 옛 파일은 수동 확인.";
  }catch(error){$("m2v6-status").textContent="재료 불러오기 실패: "+error.message;}
  e.target.value="";
 });
 globalThis.LSWaveM2Preview=Object.freeze({enabled:true,run:output,latest:()=>latest,show});
 for(const role of ["core","clad"])provenance(role);
}
if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",mount,{once:true});
else mount();
~~~

## 3. m2_preview_ui_v06.css

~~~css
/* v0.6 M2 material preview, scoped: leaves all legacy layouts untouched. */
#m2v6-preview[hidden]{display:none!important}
#m2v6-preview{scroll-margin-top:82px;margin:16px 0;padding:18px;
 background:#0e2438;color:#d8ebf4;border:1px solid #3f697c;border-radius:12px}
#m2v6-preview small{font-size:10px;color:#7bced8;letter-spacing:.04em}
#m2v6-preview h2{font-size:18px;margin:8px 0 6px;color:#eaf8ff}
#m2v6-preview h3{font-size:14px;margin:15px 0 7px;color:#bfe7f1}
#m2v6-preview p{font-size:11px;line-height:1.6;color:#c1cfdd}
#m2v6-preview .m2v6-range{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px;max-width:430px;margin:12px 0}
#m2v6-preview .m2v6-materials{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px;margin:12px 0}
#m2v6-preview .m2v6-material{min-width:0;border:1px solid #3e5a72;border-radius:8px;padding:12px}
#m2v6-preview legend{font-size:12px;color:#95e4db;padding:0 6px}
#m2v6-preview label{font-size:11px;color:#bbd8e8;display:flex;flex-direction:column;gap:6px}
#m2v6-preview .m2v6-fields{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px;margin:10px 0}
#m2v6-preview input,#m2v6-preview select,#m2v6-preview textarea{min-width:0;width:100%;padding:8px;
 border:1px solid #4a728b;background:#102f46;color:#ebf9ff;border-radius:6px;font-size:12px;box-sizing:border-box}
#m2v6-preview select{font-size:11px}
#m2v6-preview details{margin:9px 0}
#m2v6-preview details summary{cursor:pointer;color:#a7dce8;font-size:12px}
#m2v6-preview .m2v6-hint{padding:7px 8px;border-radius:5px;background:#19364b;color:#d6bda7;font-size:11px;min-height:28px}
#m2v6-preview .m2v6-actions{display:flex;gap:9px;flex-wrap:wrap;margin:13px 0}
#m2v6-preview .m2v6-actions button,#m2v6-preview .m2v6-actions .m2v6-import{
 padding:9px 12px;border-radius:7px;background:#26536c;border:1px solid #4f96a6;
 color:#ebfaff;cursor:pointer;font-size:12px;display:inline-block}
#m2v6-preview button:disabled{opacity:.45;cursor:default}
#m2v6-preview #m2v6-status{font-size:12px;color:#9fe2cd;min-height:24px}
#m2v6-preview .m2v6-plots{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin:12px 0}
#m2v6-preview canvas{width:100%;height:auto;display:block;border:1px solid #2c536e;border-radius:7px}
#m2v6-preview .m2v6-legend{display:flex;gap:12px;flex-wrap:wrap;font-size:11px;color:#b8cbd6}
#m2v6-preview .m2v6-legend span{display:flex;align-items:center;gap:5px}
#m2v6-preview .m2v6-legend i{width:9px;height:9px;border-radius:50%;display:inline-block}
#m2v6-preview .m2v6-scroll{overflow:auto;border:1px solid #33546e;border-radius:7px}
#m2v6-preview table{min-width:550px;border-collapse:collapse;width:100%;font-variant-numeric:tabular-nums;font-size:11px}
#m2v6-preview th,#m2v6-preview td{padding:9px;border-bottom:1px solid #2c4c63;text-align:left}
#m2v6-preview th{background:#1c3a50;color:#b8e4f3}
#m2v6-preview td{color:#d3e8f3}
#m2v6-preview .m2v6-provenance{border-top:1px solid #39617b;padding:10px 0}
#m2v6-preview a{color:#82e4ea}
@media(max-width:950px){#m2v6-preview .m2v6-plots,#m2v6-preview .m2v6-materials{grid-template-columns:1fr}}
@media(max-width:500px){#m2v6-preview .m2v6-fields{grid-template-columns:1fr}}
~~~

## 4. tests/test_m2_preview_v06.mjs

~~~javascript
/**
 * M2 v0.6.0: reference coefficient, interpolation and schema regression.
 * Run: node tests/test_m2_preview_v06.mjs
 * Does not use the FEM solver to compute doped n(lambda).
 */
import assert from "node:assert/strict";
import {silicaIndex,solve,PRESETS} from "../solver.mjs";
import {parseIndexTable} from "../material_models.mjs";
import {M2_CATALOG,evaluateM2Material,interpolateMeasured,
 kFromAlpha,alphaFromK,sampleM2Material,migrateM2Preview,
 makeM2PreviewDocument} from "../m2_preview_engine_v06.mjs";
const ref=[
 ["silica",1.4440236217032607],
 ["geO2",1.4515267771427718],
 ["f_wt",1.4393958509698328],
 ["p2o5",1.452224423108019]
];
for(const [id,gold] of ref){
 const m=M2_CATALOG.find(m=>m.id===id),actual=evaluateM2Material(m.model,1550).n;
 assert.ok(Math.abs(actual-gold)<1e-11,id+" Sellmeier 1550 reference mismatch");
 assert.equal(sampleM2Material(m.model,{startNm:1500,endNm:1600,count:11}).length,11);
}
assert.equal(evaluateM2Material({kind:"silica"},1550).n,silicaIndex(1550),"exact original silica");
assert.equal(evaluateM2Material({kind:"offset",deltaN:.005},1550).n,silicaIndex(1550)+.005);
assert.equal(evaluateM2Material({kind:"geO2",molPercent:0},1550).n,silicaIndex(1550));
assert.equal(evaluateM2Material({kind:"p2o5",molPercent:0},1550).n,silicaIndex(1550));
assert.equal(evaluateM2Material({kind:"f_wt",weightPercent:0},1550).n,silicaIndex(1550));
assert.throws(()=>evaluateM2Material({kind:"f_mol",molPercent:1},1550),/requires user Sellmeier coefficients/);
const custom={kind:"f_mol",molPercent:1,coefficients:{
 B:[.69616630,.40794260,.89747940],L:[.0684043,.1162414,9.896161],
 source:"User-provided test coefficients only; not validated fluorine data."}};
assert.ok(Math.abs(evaluateM2Material(custom,1550).n-silicaIndex(1550))<1e-12);
assert.throws(()=>evaluateM2Material({kind:"f_wt",weightPercent:3},1550),/0–2/);
assert.throws(()=>evaluateM2Material({kind:"p2o5",molPercent:10},1550),/0–9.1/);
assert.throws(()=>evaluateM2Material({kind:"geO2",molPercent:40},1550),/0–30/);
const table=parseIndexTable("1300,1.45,0\n1500,1.445,0.0000001\n1700,1.44,0");
for(const interpolation of ["linear","pchip"]){
 const center=evaluateM2Material({kind:"table",table,interpolation},1550);
 assert.ok(center.n<1.445&&center.n>1.44);
 assert.ok(center.kappa>=0&&center.kappa<=1e-7);
}
assert.throws(()=>evaluateM2Material({kind:"table",table,interpolation:"pchip"},1800),/Out of table range/);
assert.throws(()=>interpolateMeasured(table,1550,"spline"),/Linear|PCHIP|linear/);
for(const alpha of [0,.01,.2,20]){
 const k=kFromAlpha(alpha,1550),again=alphaFromK(k,1550);
 assert.ok(Math.abs(alpha-again)<Math.max(1e-12,alpha*1e-13));
}
assert.ok(Math.abs(kFromAlpha(.2,1550)-5.6802508913155675e-12)<1e-20);
const absorption=evaluateM2Material({kind:"silica",absorption:{kind:"alpha",value:.2}},1550);
assert.ok(Math.abs(absorption.alphaDbKm-.2)<1e-12);
assert.ok(absorption.kappa>0);
const doc=makeM2PreviewDocument({kind:"offset",deltaN:.005},{kind:"silica"});
assert.equal(doc.schemaVersion,2);
const loaded=migrateM2Preview(doc);
assert.equal(loaded.core.deltaN,.005);
const old=migrateM2Preview({schemaVersion:1,core:{mode:"f",percent:1},clad:{mode:"offset",deltaN:0}});
assert.equal(old.core.kind,"f_legacy_unresolved","legacy wt vs mol ambiguity must not be guessed");
const before=solve({...PRESETS.g652});
const after=solve({...PRESETS.g652});
for(const key of ["n_eff","MFD_PII_um","Aeff_um2"]){
 assert.equal(before[key],after[key],"legacy FEM cannot change due to M2 preview");
}
console.log("PASS M2 v0.6: coefficient sources, wt/mol separation, linear/PCHIP, complex alpha/k, schema migration, legacy FEM");
~~~

## 5. m3_study_metrics_v07.mjs

~~~javascript
/**
 * M3 v0.7 study comparisons and instrumentation helpers. Does not modify
 * existing FDFD discretization or Colab JSON schema.
 * The solver is a COLOCATED transverse Hx/Hy coupled FDFD (central differences),
 * NOT a Yee-staggered grid or edge-element FEM. Do not label it a Yee solver.
 */
export const M3_STUDY_VERSION="0.7.0";
export const M3_STUDY_SCHEMA_VERSION=2;
export const M3_STUDY_MAX_GRID=91;
export const M3_STUDY_MAX_MESHES=5;
export const M3_PYTHON_REF=Object.freeze({
 title:"Stored Python/Colab Hx-Hy FDFD (69×69 matched-grid)",
 path:"../full_vector/results/vector_mode_result.json",
 neff:1.445711888509771,
 wavelength_nm:1550,core_radius_um:5.5,delta_n:0.0034785793846155,
 window_halfwidth_um:22,grid_n:69,n_clad:1.444023621703261
});
export function validateM3Study(config,grids){
 if(!config||!Array.isArray(grids)||!grids.length||grids.length>M3_STUDY_MAX_MESHES)
  throw Error("Select 1–5 grid resolutions.");
 for(const [i,n] of grids.entries()){
  if(!Number.isInteger(n)||n<29||n>M3_STUDY_MAX_GRID||n%2!==1)
   throw Error("Grid must be odd 29–"+M3_STUDY_MAX_GRID+".");
  if(i>0&&n<=grids[i-1])throw Error("Study grid list must increase strictly.");
 }
 for(const k of ["wavelength_nm","core_radius_um","n_core","n_clad","window_halfwidth_um"]){
  if(!(Number.isFinite(config[k])&&config[k]>0))throw Error("Invalid "+k);
 }
 if(!(config.n_core>config.n_clad&&config.window_halfwidth_um>2.3*config.core_radius_um))
  throw Error("Guided core/cladding and window >2.3 × radius required.");
 return true;
}
/** Conservative accounting estimate ONLY for major Float64 arrays.
 * This is not measured process RSS or complete JS-heap usage. Browser reports
 * measured worker JS heap only when an eligible performance.memory API exists.
 */
export function estimatedTypedArrayMiB(gridN){
 const inner=gridN-2;
 const estimatedBytes=8*(12*gridN*gridN+31*inner*inner);
 return estimatedBytes/(1024*1024);
}
export function referenceRunMatch(config,gridN,refInput){
 const b=refInput||M3_PYTHON_REF;
 if(!b)return false;
 const tol=(a,c)=>Number.isFinite(a)&&Number.isFinite(c)&&Math.abs(a-c)<1e-8;
 return gridN===b.grid_n&&tol(config.wavelength_nm,b.wavelength_nm)&&
  tol(config.core_radius_um,b.core_radius_um)&&
  tol(config.window_halfwidth_um,b.window_halfwidth_um)&&
  tol(config.n_core-config.n_clad,b.delta_n)&&
  (!Number.isFinite(b.n_clad)||tol(config.n_clad,b.n_clad));
}
export function normalizeColabJSON(raw){
 if(!raw||raw.schema!=="LSCNS_vector_mode_v1"||!Array.isArray(raw.modes)||!raw.modes.length)
  throw Error("Colab JSON must have schema LSCNS_vector_mode_v1 and at least one mode.");
 if(!raw.input||!Number.isFinite(raw.modes[0].neff))
  throw Error("Colab JSON input or neff invalid.");
 return raw;
}
export function meshComparisonRows(runs,scalarNeff,analyticNeff,colab=null){
 if(!Array.isArray(runs))throw Error("M3 run array required.");
 return runs.map((r,i)=>{
  const n=r.modes?.[0]?.neff;
  if(!Number.isFinite(n))throw Error("Invalid vector effective index.");
  const matched=colab&&referenceRunMatch(r.config,r.grid_n,colab.input);
  const colabNeff=matched?colab.modes[0].neff:null;
  const deltaPrevious=i?Math.abs(n-runs[i-1].modes[0].neff):null;
  return {
   grid_n:r.grid_n,neff:n,residual:r.modes[0].residual,wallMs:r.wallMs,
   workerHeapUsedBytes:r.workerHeapUsedBytes??null,
   estimatedTypedArrayMiB:estimatedTypedArrayMiB(r.grid_n),
   diffToLP01:Number.isFinite(analyticNeff)?Math.abs(n-analyticNeff):null,
   diffToScalar:Number.isFinite(scalarNeff)?Math.abs(n-scalarNeff):null,
   diffToPreviousGrid:deltaPrevious,
   colabNeff,deltaToColab:matched?Math.abs(n-colabNeff):null,
   colabComparisonStatus:matched?"동일 격자·입력":"미검증 (입력/격자 상이)",
   convergenceStatus:deltaPrevious==null?"기준 격자":deltaPrevious<=1e-6?"인접 격자차 ≤1e-6":"미수렴"
  };
 });
}
/** Progress reports physical inner eigen iterations, not arbitrary frame timers.
 * Two polarizations per grid, max 105 outer inverse iterations per mode.
 * Actual solve lengths vary, so percentages are indicative only.
 */
export function estimateProgressPercent(gridIndex,gridCount,mode,iteration){
 const modeIndex=mode==="y"?1:0,local=Math.min(1,Math.max(0,(iteration+1)/105));
 return Math.min(99,Math.max(0,Math.round((gridIndex+(modeIndex+local)/2)/gridCount*100)));
}
export function studyDocument({config,grids,runs,scalarNeff,analyticNeff,colab,vectorJSON}){
 const rows=meshComparisonRows(runs,scalarNeff,analyticNeff,colab);
 const last=rows.at(-1);
 return {kind:"LSWaveFullVectorStudy",schemaVersion:M3_STUDY_SCHEMA_VERSION,version:M3_STUDY_VERSION,
  generatedAt:new Date().toISOString(),input:config,requestedGrids:grids,
  method:"2D full-vector, coupled transverse Hx/Hy COLOCATED FDFD (not Yee) / matrix-free shift-invert BiCGSTAB",
  tolerance:{vectorRelativeEigenResidual:1e-7,linearBiCGSTAB:2e-9,adjacentGridAbsNeff:1e-6},
  scalarLP01:{analyticNeff,scalarFEMNeff:scalarNeff,model:"fixed real index, circular step-index, no PML"},
  storedColabReference:colab?{neff:colab.modes[0].neff,grid_n:colab.input?.grid_n,source:M3_PYTHON_REF.path}:null,
  convergence:rows,wallMsTotal:runs.reduce((sum,r)=>sum+r.wallMs,0),
  finalGridConvergenceConfirmed:last?.convergenceStatus==="인접 격자차 ≤1e-6",
  portableVectorResult:vectorJSON,
  unverified:[
   "Different grid sizes are not a same-discretization cross-check; staircased core boundary causes non-monotone neff.",
   "Grid convergence and full-vector continuum accuracy not assumed unless grid-difference criterion passes.",
   "This is a colocation finite difference Maxwell operator, not Yee staggered FDFD or edge FEM.",
   "No PML, HCF, bend/leakage loss, complex eigenvalue, dispersion or mesh-validated manufacturer fiber data."
  ]};
}
~~~

## 6. m3_study_worker_v07.mjs

~~~javascript
/**
 * M3 v0.7.0: Worker-based, cancellable (by Worker.terminate in caller)
 * multi-grid 2D vector FDFD study. No paid services / server.
 *
 * IMPORTANT: original full_vector/vector_modes.py and browser
 * vector_fdfd_browser.mjs solve a *collocated* Hx/Hy finite-difference
 * Maxwell eigenproblem, not an explicitly Yee-staggered FDFD grid.
 * Both modes are full-vector coupled, but a strict Yee certification
 * is NOT claimed.
 */
import {solveVectorFDFD,vectorResultJSON} from "./vector_fdfd_browser.mjs?v=v0.7.0";
import {PRESETS,solve} from "./solver.mjs?v=v0.7.0";
import {lp01Characteristic} from "./validation_lp01.mjs";
import {validateM3Study,estimateProgressPercent,studyDocument,
 normalizeColabJSON} from "./m3_study_metrics_v07.mjs";
const heapUsed=()=>Number.isFinite(self.performance?.memory?.usedJSHeapSize)?
 self.performance.memory.usedJSHeapSize:null;
function vectorAllModes(result){
 // The historical portable JSON exporter maps only first mode to field arrays.
 // Generate each mode's field map individually and attach it to that mode,
 // retaining schema="LSCNS_vector_mode_v1" and old first mode contents.
 const first=vectorResultJSON(result);
 for(let i=1;i<result.modes.length;i++){
  const one=vectorResultJSON({...result,modes:[result.modes[i]]});
  first.modes[i].field_map=one.modes[0].field_map;
 }
 return first;
}
self.onmessage=e=>{
 const {op,id,config,grids,colabReference}=e.data||{};
 if(op!=="run")return;
 try{
  validateM3Study(config,grids);
  const colab=colabReference?normalizeColabJSON(colabReference):null;
  const prefix={...PRESETS.g654,profile:"step",wavelength_nm:config.wavelength_nm,
   core_radius_um:config.core_radius_um,outer_radius_um:40,elements:750,
   delta_n_core:config.n_core-config.n_clad,clad_index_override:config.n_clad};
  let scalarNeff=null,analyticNeff=null,comparisonWarning=null;
  try{
   scalarNeff=solve(prefix).n_eff;
   analyticNeff=lp01Characteristic({wavelength_nm:config.wavelength_nm,
    core_radius_um:config.core_radius_um,n_core:config.n_core,n_clad:config.n_clad}).n_eff;
  }catch(error){comparisonWarning="Scalar/LP01 reference unavailable: "+error.message;}
  const runs=[];
  let finalPortable=null;
  for(let i=0;i<grids.length;i++){
   const grid_n=grids[i],gridConfig={...config,profile:"step",grid_n};
   self.postMessage({id,op:"progress",percent:Math.round(100*i/grids.length),
    grid_n,phase:"grid-start",detail:"Preparing "+grid_n+"×"+grid_n});
   const before=heapUsed(),start=performance.now();
   const vector=solveVectorFDFD(gridConfig,{
    modes:2,
    onProgress:({mode,iteration,residual})=>
     self.postMessage({id,op:"progress",grid_n,phase:"inverse-iteration",
      mode,iteration,residual,percent:estimateProgressPercent(i,grids.length,mode,iteration)})
   });
   const elapsed=performance.now()-start,endHeap=heapUsed();
   const metaModes=vector.modes.map(m=>{
    const {field_maps,field_components,...meta}=m;
    return meta;
   });
   runs.push({grid_n,config:gridConfig,modes:metaModes,wallMs:elapsed,
    workerHeapUsedBytes:endHeap,workerHeapBeforeBytes:before,
    heapStatus:endHeap==null?"not-available":"performance.memory worker heap (experimental)"});
   finalPortable=vectorAllModes(vector);
   self.postMessage({id,op:"progress",grid_n,phase:"grid-finished",
    percent:Math.round(100*(i+1)/grids.length),wallMs:elapsed,
    detail:"Completed "+grid_n+"×"+grid_n});
  }
  const report=studyDocument({config,grids,runs,scalarNeff,analyticNeff,
   colab,vectorJSON:finalPortable});
  if(comparisonWarning)report.unverified.push(comparisonWarning);
  self.postMessage({id,op:"result",report});
 }catch(error){
  self.postMessage({id,op:"error",message:String(error?.message||error)});
 }
};
~~~

## 7. m3_study_ui_v07.mjs

~~~javascript
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
 // The checked-in Python fixture uses pure-silica Sellmeier cladding.
 // Stored JSON predates explicit n_clad metadata; annotate only this known fixture.
 colab.input={...colab.input,n_clad:M3_PYTHON_REF.n_clad};
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
~~~

## 8. m3_study_ui_v07.css

~~~css
/* New M3 v0.7 vector FDFD convergence UI: scoped to #m3v7-study */
#m3v7-study[hidden]{display:none!important}
#m3v7-study{scroll-margin-top:75px;background:#10243a;border:1px solid #385a70;
 color:#d4e8f4;border-radius:12px;margin:22px 0;padding:16px}
#m3v7-study small{font-size:10px;color:#81ccdc;letter-spacing:.04em}
#m3v7-study h2{font-size:18px;color:#eff9ff;margin:8px 0}
#m3v7-study h3{font-size:14px;color:#c1edf6;margin:16px 0 8px}
#m3v7-study p{font-size:12px;line-height:1.6;color:#b9cfdf}
#m3v7-study .m3v7-config{display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:9px;margin:14px 0}
#m3v7-study label{font-size:11px;color:#bbd9e9;display:flex;flex-direction:column;gap:5px;min-width:0}
#m3v7-study input[type="number"],#m3v7-study select{width:100%;min-width:0;box-sizing:border-box;
 background:#122e46;color:#f1f9ff;border:1px solid #456d85;border-radius:6px;padding:8px}
#m3v7-study fieldset{border:1px solid #426277;border-radius:7px;padding:9px;display:flex;flex-wrap:wrap;gap:12px}
#m3v7-study legend{font-size:11px;color:#b2dfed}
#m3v7-study fieldset label{display:flex;flex-direction:row;gap:5px;align-items:center}
#m3v7-study .m3v7-estimate{color:#e7bd91;background:#243344;padding:8px;border-radius:5px}
#m3v7-study .m3v7-actions{display:flex;flex-wrap:wrap;gap:8px;margin:13px 0}
#m3v7-study .m3v7-actions button,#m3v7-study .m3v7-actions label{
 display:inline-block;cursor:pointer;border:1px solid #5594aa;border-radius:8px;background:#235670;
 color:#eefaff;padding:9px 12px;font-size:12px}
#m3v7-study button:disabled{opacity:.42;cursor:default}
#m3v7-study .m3v7-progress{display:flex;align-items:center;gap:12px}
#m3v7-study progress{height:13px;flex:1;accent-color:#46b5aa}
#m3v7-study .m3v7-scroll{overflow:auto;max-height:410px;border:1px solid #355572;border-radius:7px;margin:12px 0}
#m3v7-study table{border-collapse:collapse;min-width:1050px;width:100%;font-size:11px;font-variant-numeric:tabular-nums}
#m3v7-study th,#m3v7-study td{padding:9px;border-bottom:1px solid #31516b;text-align:left;white-space:nowrap}
#m3v7-study th{background:#1a3850;color:#b9e0ec;position:sticky;top:0}
#m3v7-study td small{display:block;font-size:10px;color:#89acbf;margin-top:4px}
#m3v7-study .m3v7-controls{display:flex;gap:12px;flex-wrap:wrap;margin:14px 0}
#m3v7-study .m3v7-maps{display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:10px}
#m3v7-study .m3v7-maps article{background:#0a1d30;border:1px solid #2b506b;border-radius:7px;padding:7px;min-width:0}
#m3v7-study .m3v7-maps h4{font-size:11px;margin:3px 0 8px;color:#cee5f0}
#m3v7-study canvas{display:block;width:100%;height:auto}
#m3v7-study details{border-top:1px solid #315772;padding:12px 0;margin-top:14px}
#m3v7-study details summary{cursor:pointer;font-size:12px;color:#a2dde5}
@media(max-width:1000px){#m3v7-study .m3v7-config{grid-template-columns:repeat(3,minmax(0,1fr))}#m3v7-study .m3v7-maps{grid-template-columns:repeat(3,minmax(0,1fr))}}
@media(max-width:620px){#m3v7-study .m3v7-config{grid-template-columns:repeat(2,minmax(0,1fr))}#m3v7-study .m3v7-maps{grid-template-columns:repeat(2,minmax(0,1fr))}}
~~~

## 9. tests/test_m3_study_v07.mjs

~~~javascript
/**
 * M3 v0.7 regression, same-grid independent Python reference.
 * Run: node tests/test_m3_study_v07.mjs
 * n=69 benchmark is computation-heavy. Console time is measured on local host.
 */
import assert from "node:assert/strict";
import {readFileSync} from "node:fs";
import {performance} from "node:perf_hooks";
import {PRESETS,solve,silicaIndex} from "../solver.mjs";
import {lp01Characteristic} from "../validation_lp01.mjs";
import {solveVectorFDFD,vectorResultJSON} from "../vector_fdfd_browser.mjs";
import {M3_PYTHON_REF,validateM3Study,estimatedTypedArrayMiB,
 meshComparisonRows,normalizeColabJSON,referenceRunMatch,studyDocument} from "../m3_study_metrics_v07.mjs";
const raw=JSON.parse(readFileSync(new URL("../../full_vector/results/vector_mode_result.json",import.meta.url),"utf8"));
const colab=normalizeColabJSON({...raw,input:{...raw.input,n_clad:M3_PYTHON_REF.n_clad}});
const ncl=silicaIndex(1550),config={profile:"step",wavelength_nm:1550,core_radius_um:5.5,
 window_halfwidth_um:22,n_core:ncl+.0034785793846155,n_clad:ncl};
validateM3Study(config,[29,41,69]);
assert.throws(()=>validateM3Study(config,[30]),/odd/);
assert.throws(()=>validateM3Study(config,[91,29]),/increase/);
assert.ok(estimatedTypedArrayMiB(69)>0);
assert.ok(referenceRunMatch(config,69,colab.input));
assert.equal(referenceRunMatch(config,51,colab.input),false);
assert.equal(referenceRunMatch({...config,n_clad:ncl+.001,n_core:config.n_core+.001},69,colab.input),false);
const sc=solve({...PRESETS.g654,clad_index_override:ncl});
const lp=lp01Characteristic({wavelength_nm:1550,core_radius_um:config.core_radius_um,
 n_core:config.n_core,n_clad:ncl});
const runs=[],start=performance.now();
for(const grid_n of [29,41,69]){
 const t=performance.now();
 const v=solveVectorFDFD({...config,grid_n},{modes:2});
 const elapsed=performance.now()-t;
 runs.push({grid_n,config:{...config,grid_n},modes:v.modes,wallMs:elapsed});
 assert.ok(v.modes.every(m=>m.residual<1e-7));
 const portable=vectorResultJSON(v);
 if(grid_n===69){
  assert.equal(portable.schema,"LSCNS_vector_mode_v1");
  assert.ok(Math.abs(v.modes[0].neff-colab.modes[0].neff)<1e-8,
   "same-grid Python/JS mode match");
  assert.ok(portable.modes[0].field_map.intensity_normalized_to_max_total.Ex.length===69);
 }
 console.log("grid",grid_n,"ms",elapsed.toFixed(1),"neff",v.modes[0].neff);
}
const rows=meshComparisonRows(runs,sc.n_eff,lp.n_eff,colab);
assert.equal(rows[0].colabComparisonStatus,"미검증 (입력/격자 상이)");
assert.equal(rows[2].colabComparisonStatus,"동일 격자·입력");
assert.ok(rows[2].deltaToColab<1e-8);
assert.equal(rows[1].convergenceStatus,"미수렴","Do not claim monotone convergence for staircase index");
const report=studyDocument({config,grids:[29,41,69],runs,scalarNeff:sc.n_eff,
 analyticNeff:lp.n_eff,colab,vectorJSON:vectorResultJSON(solveVectorFDFD({...config,grid_n:29},{modes:2}))});
assert.equal(report.schemaVersion,2);
assert.equal(report.version,"0.7.0");
assert.equal(report.storedColabReference.grid_n,69);
assert.equal(report.finalGridConvergenceConfirmed,false);
console.log("PASS M3 v0.7: LP01/scalar/Colab neff, mesh convergence correctly not claimed, cost",
 (performance.now()-start).toFixed(1),"ms");
~~~

## 10. tests/benchmark_m3_v07.mjs

~~~javascript
/**
 * Resource benchmark for the user's actual Node runtime (NOT browser memory).
 * Node >=20; optionally run with --expose-gc:
 * node --expose-gc tests/benchmark_m3_v07.mjs
 * M3_GRIDS=29,41,51,69,81 node --expose-gc tests/benchmark_m3_v07.mjs
 * Reports process heap/RSS/external/ArrayBuffer before-after and peak RSS
 * as MEASURED host numbers. This cannot infer browser Worker RSS.
 */
import {performance} from "node:perf_hooks";
import {resourceUsage} from "node:process";
import {silicaIndex} from "../solver.mjs";
import {solveVectorFDFD} from "../vector_fdfd_browser.mjs";
import {estimatedTypedArrayMiB} from "../m3_study_metrics_v07.mjs";
const grids=(process.env.M3_GRIDS||"29,41,51,69").split(",").map(Number);
const input={profile:"step",wavelength_nm:1550,core_radius_um:5.5,window_halfwidth_um:22,
 n_clad:silicaIndex(1550),n_core:silicaIndex(1550)+.0034785793846155};
const mb=bytes=>+(bytes/1048576).toFixed(3);
const result=[];
for(const grid_n of grids){
 global.gc?.();
 const before=process.memoryUsage(),rssBefore=resourceUsage().maxRSS;
 const start=performance.now(),v=solveVectorFDFD({...input,grid_n},{modes:2}),
  ms=performance.now()-start,after=process.memoryUsage(),rssAfter=resourceUsage().maxRSS;
 const item={grid_n,measuredWallMs:+ms.toFixed(1),
  measuredProcessBeforeMiB:{rss:mb(before.rss),heap:mb(before.heapUsed),external:mb(before.external),arrayBuffers:mb(before.arrayBuffers)},
  measuredProcessAfterMiB:{rss:mb(after.rss),heap:mb(after.heapUsed),external:mb(after.external),arrayBuffers:mb(after.arrayBuffers)},
  measuredPeakRSSMiB:process.platform==="linux"?+(rssAfter/1024).toFixed(3):null,
  typedArrayModelEstimateMiB:+estimatedTypedArrayMiB(grid_n).toFixed(3),
  effectiveIndex:v.modes[0].neff,eigenResidual:v.modes[0].residual};
 result.push(item);
 console.log(JSON.stringify(item));
}
console.log("M3 v0.7 resource benchmark finished; host OS measurements are NOT browser heap measurements.");
~~~

## 11. tests/browser_m2_m3_v07.mjs

~~~javascript
/**
 * Optional Chromium/Playwright M2/M3 UI smoke.
 * Serve LSCNS repo root: python3 -m http.server 8765
 * From web/: node tests/browser_m2_m3_v07.mjs
 * Not required by GitHub Pages at runtime (developer test only).
 */
import assert from "node:assert/strict";
import {createRequire} from "node:module";
const {chromium}=createRequire(import.meta.url)("playwright");
const browser=await chromium.launch({headless:true,args:["--no-sandbox","--disable-dev-shm-usage"]});
try{
 const page=await browser.newPage({viewport:{width:1520,height:950},acceptDownloads:true});
 const errors=[];page.on("pageerror",e=>errors.push(e.message));
 await page.goto(process.env.TEST_BASE_URL||"http://127.0.0.1:8765/COMSOL/universal_silica_v01/web/",{waitUntil:"load",timeout:90000});
 if(await page.locator('[data-wizard="cancel"]').isVisible())await page.locator('[data-wizard="cancel"]').click();
 const original=await page.evaluate(()=>({
  preset:document.querySelector("#preset")?.value,
  neff:document.querySelector("#neff")?.textContent,
  mfd:document.querySelector("#mfd")?.textContent,
  aeff:document.querySelector("#aeff")?.textContent
 }));
 assert.ok(await page.locator("#m2v6-nav").isVisible());
 await page.locator("#m2v6-nav").click();
 await page.locator("#m2v6-run").click();
 await page.waitForFunction(()=>window.LSWaveM2Preview?.latest()?.schemaVersion===2);
 assert.ok((await page.locator("#m2v6-rows tr").count())>=6);
 assert.ok((await page.locator("#m2v6-status").textContent()).includes("미리보기 성공"));
 const [m2download]=await Promise.all([page.waitForEvent("download"),page.locator("#m2v6-save").click()]);
 assert.equal(m2download.suggestedFilename(),"LS_Wave_M2_v0.6_material_preview.json");
 await page.locator("#m2v6-core-kind").selectOption("f_mol");
 await page.locator("#m2v6-core-mol").fill("3");
 await page.locator("#m2v6-run").click();
 assert.ok((await page.locator("#m2v6-status").textContent()).includes("requires user Sellmeier coefficients"));
 assert.ok(await page.locator("#m3v7-nav").isVisible());
 await page.locator("#m3v7-nav").click();
 for(const n of [41,51,69])await page.locator('#m3v7-grids input[value="'+n+'"]').uncheck();
 await page.locator("#m3v7-run").click();
 await page.waitForFunction(()=>window.LSWaveM3Study?.latest()?.schemaVersion===2,undefined,{timeout:90000});
 const result=await page.evaluate(()=>window.LSWaveM3Study.latest());
 assert.equal(result.convergence.length,1);
 assert.equal(result.portableVectorResult.schema,"LSCNS_vector_mode_v1");
 assert.ok(result.portableVectorResult.modes.every(m=>m.field_map?.intensity_normalized_to_max_total?.Ez));
 assert.equal(await page.locator("#m3v7-convergence tr").count(),1);
 const [vectorDownload]=await Promise.all([page.waitForEvent("download"),page.locator("#m3v7-export").click()]);
 assert.equal(vectorDownload.suggestedFilename(),"LS_Wave_Vector_v0.7_colab_compatible.json");
 await page.locator('#m3v7-grids input[value="29"]').uncheck();
 await page.locator('#m3v7-grids input[value="91"]').check();
 await page.locator("#m3v7-run").click();
 await page.locator("#m3v7-cancel").click();
 assert.ok((await page.locator("#m3v7-status").textContent()).includes("사용자 취소"));
 assert.equal(await page.locator("#m3v7-export").isDisabled(),true);
 const unchanged=await page.evaluate(()=>({
  preset:document.querySelector("#preset")?.value,
  neff:document.querySelector("#neff")?.textContent,
  mfd:document.querySelector("#mfd")?.textContent,
  aeff:document.querySelector("#aeff")?.textContent
 }));
 assert.deepEqual(unchanged,original,"Existing solver cards must not be modified by M2/M3.");
 assert.ok(await page.locator("#fv-json-file").count()===1);
 assert.ok(await page.locator("#ls-m1q-run").count()===1);
 assert.deepEqual(errors,[]);
 console.log("PASS M2/M3 v0.6/0.7 browser: chart, CSV, mol/wt warning, Worker, cancellation, old controls");
}finally{await browser.close();}
~~~

## Original files and rollout

Read [M2_M3_V07_VALIDATION.md](M2_M3_V07_VALIDATION.md) and [CHANGELOG.md](CHANGELOG.md). Do not replace any existing solver, notebook, M1 or JSON schema with this material.
