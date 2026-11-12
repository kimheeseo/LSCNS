/**
 * M2 v0.6.0 — independent optical materials PREVIEW, not coupled to an eigensolver.
 * All coefficients below have a public reference; uncertain F mol%-specific
 * coefficients are deliberately NOT invented.
 *
 * Pure SiO2: I. H. Malitson, JOSA 55, 1205 (1965), Eq. 1; DOI:10.1364/JOSA.55.001205
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
 silica:{label:"I. H. Malitson, JOSA 55 (1965), pure fused SiO₂ Sellmeier Eq. 1; 20 °C",url:"https://doi.org/10.1364/JOSA.55.001205"},
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
  n=p===0?silicaIndex(lambda):sellmeierIndex(mixCoefficients(SELLMEIER.silica,SELLMEIER.geO2,p/100),lambda);
 }else if(kind==="p2o5"){
  const p=range(finite(cfg.molPercent??0,"P₂O₅ mol%"),0,9.1,"P₂O₅ mol%");
  n=p===0?silicaIndex(lambda):sellmeierIndex(mixCoefficients(SELLMEIER.silica,SELLMEIER.p9,p/9.1),lambda);
 }else if(kind==="f_wt"){
  const p=range(finite(cfg.weightPercent??0,"F wt%"),0,2,"F wt%");
  const ab=p<=1?[SELLMEIER.silica,SELLMEIER.f1,p]:[SELLMEIER.f1,SELLMEIER.f2,p-1];
  n=p===0?silicaIndex(lambda):sellmeierIndex(mixCoefficients(...ab),lambda);
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
