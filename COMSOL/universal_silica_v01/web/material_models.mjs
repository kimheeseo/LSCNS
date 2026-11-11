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
