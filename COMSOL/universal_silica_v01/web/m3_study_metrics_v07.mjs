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
