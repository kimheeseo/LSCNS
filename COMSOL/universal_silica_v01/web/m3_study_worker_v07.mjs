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
