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
