import {config,LIMITS} from './vector_geometry.mjs?v=20261112-m3m6-v2';
import {mcfCenters} from './vector_geometry.mjs?v=20261112-m3m6-v2';
import {solveYee,portableYee,denseSolve} from './yee_solver.mjs?v=20261112-m3m6-v2';
import {solve,PRESETS,claddingIndex} from './solver.mjs';
import {lp01Characteristic} from './validation_lp01.mjs';
import {calculateMCF} from './mcf_xt.mjs';
export const SCHEMA=3;
const close=(a,b,t=1e-8)=>Number.isFinite(a)&&Number.isFinite(b)&&Math.abs(a-b)<=t;
export function normalizeReference(doc){
 const d=doc.portableVectorResult??doc.result??doc;
 const hcf=d.base??(Array.isArray(d.results)?d.results.find(x=>Number.isFinite(x.neff_real)):null)??d;
 if(hcf.full_vector===false&&Number.isFinite(hcf.neff_real))return {schema:'LSCNS_vector_mode_v1',referenceKind:'legacy_hcf_scalar',method:hcf.model,
  input:{geometry:{kind:'hcf',hcf:hcf.input},wavelength_nm:hcf.input.wavelengthNm},modes:[{neff:hcf.neff_real,neff_imag:hcf.neff_imag}],legacyResults:d.results??[hcf]};
 if(d.schema!=='LSCNS_vector_mode_v1'||!Array.isArray(d.modes)||!d.input||!d.modes.every(m=>Number.isFinite(m.neff)))throw Error('LSCNS_vector_mode_v1 result with input/modes required');
 return d;
}
/** Metadata must establish the physical model. Missing metadata never means agreement. */
export function compareReference(result,reference){
 const input=result.input,c=config(input),ref=reference?.input;
 if(reference?.referenceKind==='legacy_hcf_scalar'){
  const r=reference.legacyResults.find(r=>close(r.input.wavelengthNm,c.wavelength_nm)),keys=['coreDiameterUm','tubeCount','nesting','outerTubeRatio','innerRatio','thirdRatio','gapCoreUm','minNeighborGapUm','nestedClearanceUm','wall1Um','wall2Um','wall3Um','airIndex'];
  const same=c.geometry.kind==='hcf'&&r&&keys.every(k=>close(c.geometry.hcf?.[k],r.input[k],1e-6));
  return {status:'미검증: 기존 HCF Colab은 scalar P1 FEM/PML. '+(same?'기하·파장 일치; grid/PML/domain/벡터 물리 차이 있음':'기하/파장 불일치 또는 누락'),rows:result.modes.map((m,i)=>({mode:i+1,neff:m.neff,referenceNeff:r?.neff_real??null,deltaNeff:same?m.neff-r.neff_real:null,deltaImag:same?m.neff_imag-r.neff_imag:null,status:'미검증: vector vs scalar / mode identity unconfirmed'}))};
 }
 if(!ref)return {status:'미검증: 비교 JSON 없음',rows:[]};
 const issues=[],g=c.geometry.kind,rg=ref.geometry?.kind??'radial';
 if(g!==rg)issues.push('geometry kind');
 const rc=ref.n_clad??ref.clad_index_override;
 for(const [k,a,b]of [['λ',c.wavelength_nm,ref.wavelength_nm],['a',c.core_radius_um,ref.core_radius_um],['window',c.window_halfwidth_um,ref.window_halfwidth_um],['n_clad',c.n_clad,rc],['n_core',c.n_core,ref.n_core??(rc!=null?rc+(ref.delta_n??ref.delta_n_core):undefined)]])if(!close(a,b))issues.push(k);
 if(JSON.stringify(c.materials??null)!==JSON.stringify(ref.materials??null))issues.push('M2 materials');
 if(g!=='radial'&&JSON.stringify(c.geometry)!==JSON.stringify(ref.geometry))issues.push('geometry details');
 if(c.pml.enabled&&JSON.stringify(c.pml)!==JSON.stringify(ref.pml))issues.push('PML');
 if((c.k_core??0)!==(ref.k_core??0)||(c.k_clad??0)!==(ref.k_clad??0))issues.push('absorption');
 if((c.bend_radius_um??0)!==(ref.bend_radius_um??0))issues.push('bend R');
 if(issues.length)return {status:'미검증: 입력 불일치/누락 · '+issues.join(', '),rows:[]};
 return {status:'동일 물리 입력; grid/algorithm may differ',rows:result.modes.map((m,i)=>{
  // By mode index only. Field overlap is a separate inspection; no automatic mode identity certification.
  const r=reference.modes[i];return {mode:i+1,neff:m.neff,referenceNeff:r?.neff??null,
   deltaNeff:r?m.neff-r.neff:null,deltaImag:r&&Number.isFinite(r.neff_imag)?m.neff_imag-r.neff_imag:null,
   status:r?'index pairing; check polarization/field overlap':'미검증: reference mode missing'};
 })};
}
export function scalarReferences(c){
 if(c.geometry.kind!=='radial'||c.profile!=='step'||c.pml.enabled||c.bend_radius_um||(c.k_core??0)||(c.k_clad??0)||c.materials)return {status:'미검증: lossless straight radial step-index only'};
 try{
  const analytic=lp01Characteristic({wavelength_nm:c.wavelength_nm,core_radius_um:c.core_radius_um,n_core:c.n_core,n_clad:c.n_clad}).n_eff;
  const fem=solve({...PRESETS.g654,profile:'step',wavelength_nm:c.wavelength_nm,core_radius_um:c.core_radius_um,outer_radius_um:Math.max(40,c.window_halfwidth_um),elements:1000,delta_n_core:c.n_core-c.n_clad,clad_index_override:c.n_clad}).n_eff;
  return {analytic,scalarFEM:fem,status:'M1 weak-guidance LP01 characteristic; legacy radial scalar FEM (different discretization/finite domain)'};
 }catch(error){return {status:'미검증: '+error.message};}
}
export const PARAMS=Object.freeze({a:'core_radius_um',delta_n:'delta_n',lambda:'wavelength_nm',R:'bend_radius_um',grid:'grid_n',pml_thickness:'pml.thickness_um',pml_strength:'pml.strength',pitch:'geometry.pitch_um'});
export function setParameter(c,name,value){
 const out=structuredClone(c),path=PARAMS[name]??name;
 if(!Object.values(PARAMS).includes(path))throw Error('Unsupported sweep parameter: '+name);
 if(path==='delta_n'){if(out.materials)throw Error('Δn sweep requires constant material mode');out.n_core=out.n_clad+value;}
 else {const keys=path.split('.');let ref=out;for(const k of keys.slice(0,-1))ref=ref[k];ref[keys.at(-1)]=value;}
 return out;
}
export function sweepValues(axis){
 if(!axis||!Object.hasOwn(PARAMS,axis.parameter))throw Error('Choose valid sweep parameter');
 const {start,end,count}=axis;if(![start,end].every(Number.isFinite)||!Number.isInteger(count)||count<2||count>18)throw Error('Sweep: finite range, 2–18 points per axis');
 return Array.from({length:count},(_,i)=>start+(end-start)*i/(count-1));
}
export function jobsFor(c,axes=[]){
 if(axes.length>2)throw Error('At most 2 sweep axes');
 if(axes.length===2&&axes[0].parameter===axes[1].parameter)throw Error('Sweep axes must differ');
 let jobs=[{config:config(c),parameters:{}}];
 for(const axis of axes){
  const next=[];for(const job of jobs)for(const v of sweepValues(axis))next.push({config:config(setParameter(job.config,axis.parameter,v)),parameters:{...job.parameters,[axis.parameter]:v}});jobs=next;
 }
 if(jobs.length>LIMITS.maxSweep)throw Error('At most 36 total sweep points');return jobs;
}
export function runStudy(input,{axes=[],reference=null,jobs:explicitJobs=null,onProgress=()=>{}}={}){
 const jobs=explicitJobs??jobsFor(input,axes),rows=[],results=[],start=performance.now();
 if(!Array.isArray(jobs)||!jobs.length||jobs.length>36)throw Error("1–36 study jobs required");
 for(const job of jobs)job.config=config(job.config);
 for(let i=0;i<jobs.length;i++){
  const job=jobs[i];onProgress({phase:'job',job:i+1,total:jobs.length,percent:100*i/jobs.length});
  try{
   const r=solveYee(job.config,{onProgress:p=>onProgress({...p,job:i+1,total:jobs.length,percent:(i+p.percent/100)*100/jobs.length})}),portable=portableYee(r),refs=scalarReferences(job.config);
   const first=portable.modes[0],prev=rows.at(-1);
   rows.push({job:i+1,...job.parameters,grid_n:r.grid_n,neff:first.neff,neff_imag:first.neff_imag,
    loss_dB_km:first.loss_dB_km,loss_dB_turn:first.loss_dB_turn,MFD_um:first.MFD_2nd_um,Aeff_um2:first.Aeff_um2,
    residual:first.residual,wallMs:r.wallMs,workspaceEstimateBytes:r.memory.typedArrayWorkspaceModelBytes,
    heapMeasuredBytes:r.memory.heapUsedBytes,analytic:refs.analytic??null,scalarFEM:refs.scalarFEM??null,
    deltaAnalytic:refs.analytic!=null?first.neff-refs.analytic:null,deltaScalar:refs.scalarFEM!=null?first.neff-refs.scalarFEM:null,
    deltaPrevious:prev?.neff!=null?first.neff-prev.neff:null,status:r.warnings.join(' / ')||'Eigen residual passed; continuum/loss convergence 미검증'});
   results.push({parameters:job.parameters,result:portable,reference:refs,comparison:reference?compareReference(portable,reference):null});
   onProgress({phase:'job-complete',job:i+1,total:jobs.length,row:rows.at(-1),percent:100*(i+1)/jobs.length});
  }catch(error){rows.push({job:i+1,...job.parameters,grid_n:job.config.grid_n,status:'FAILED / 미검증: '+error.message});onProgress({phase:'job-failed',job:i+1,total:jobs.length,row:rows.at(-1),percent:100*(i+1)/jobs.length});}
 }
 // Keep result metadata for every job; full field maps only at last successful point to bound message size.
 const latest=results.at(-1)?.result??null;
 const metadata=results.map(({result,...other})=>({...other,result:{...result,modes:result.modes.map(({field_map,...m})=>m)}}));
 return {schema:'LSCNS_vector_study',schemaVersion:SCHEMA,generatedAt:new Date().toISOString(),input:config(input),axes,rows,
  result:latest,runs:metadata,wallMs:performance.now()-start,
  unverified:['Colab HCF baseline absent unless an identical-input file is supplied','PML power <35% is a candidate filter, not a proof of physical mode','Mesh/PML sweeps require field/mode tracking; mode index can swap at crossings']};
}
export function rowsCSV(rows){
 const keys=[...new Set(rows.flatMap(row=>Object.keys(row)))],quote=x=>'"'+String(x??'').replaceAll('"','""')+'"';
 return '\uFEFF'+[keys.map(quote).join(','),...rows.map(row=>keys.map(k=>quote(row[k])).join(','))].join('\n')+'\n';
}
export function twoCoreMetrics(result,{length_mm=1}={}){
 const c=result.input;if(c.geometry?.kind!=='mcf'||c.geometry.count!==2)return {status:'For N-core use supermode spectrum/beat lengths; two-core formula only for two identical cores'};
 const modes=result.modes.filter(m=>m.Ex_intensity_fraction>.6).sort((a,b)=>b.neff-a.neff);
 if(modes.length<2)return {status:'미검증: two Ex supermodes required (request ≥4 modes)'};
 const delta=Math.abs(modes[0].neff-modes[1].neff),kappa=Math.PI*delta/(c.wavelength_nm*1e-9),phi=kappa*length_mm/1000,p2=Math.sin(phi)**2,p1=1-p2;
 let legacy=null;
 const valid=close(c.wavelength_nm,1550)&&close(c.core_radius_um,4.1)&&close(c.n_core-c.n_clad,.005)&&!c.pml.enabled&&!c.bend_radius_um&&!c.materials&&(c.geometry.pitch_um>=12&&c.geometry.pitch_um<=24);
 if(valid)legacy=calculateMCF(c.geometry.pitch_um,length_mm);
 return {status:'Direct vector supermode splitting; coherent identical-core model, no stochastic cable XT',delta_neff:delta,kappa_per_m:kappa,coupling_length_mm:Math.PI/(2*kappa)*1000,length_mm,p2,xt_db:10*Math.log10(p2/Math.max(p1,1e-300)),legacyDelta:legacy?.delta_neff??null,deltaToLegacy:legacy?delta-legacy.delta_neff:null,legacyStatus:legacy?'same λ/a/Δn; old scalar reference interpolation':'미검증: legacy λ/a/Δn/pitch assumptions mismatch'};
}
export function migrateProject(doc){
 if(!doc||typeof doc!=='object')throw Error('Project JSON object required');
 if(doc.schema==='LSCNS_wave_project'){
  if(doc.schemaVersion>SCHEMA)throw Error('Newer project schema; upgrade application first');
  if(!doc.vector?.config)throw Error('Project missing vector config');config(doc.vector.config);
  return {project:structuredClone({...doc,schemaVersion:SCHEMA}),warnings:doc.schemaVersion===SCHEMA?[]:['Earlier project schema migrated; missing fields use defaults']};
 }
 const m=doc.model??doc;
 if(m.profile==='mcf'){
  if(!Number.isFinite(m.pitch_um)||!Number.isFinite(m.length_mm))throw Error('Invalid legacy MCF design');
  return {project:{schema:'LSCNS_wave_project',schemaVersion:SCHEMA,legacyModel:structuredClone(m),vector:{config:config({wavelength_nm:1550,core_radius_um:4.1,window_halfwidth_um:30,grid_n:48,modes:4,geometry:{kind:'mcf',count:2,pitch_um:m.pitch_um}}),history:[]}},warnings:['구버전 MCF 보간 설계 이관; 기존 보간 모드는 보존하고 직접 벡터 해석 입력은 기본 고정조건으로 생성했습니다']};
 }
 if(!m.profile||!Number.isFinite(m.core_radius_um)||!Number.isFinite(m.wavelength_nm))throw Error('Unrecognized legacy design JSON');
 return {project:{schema:'LSCNS_wave_project',schemaVersion:SCHEMA,legacyModel:structuredClone(m),vector:{config:config({wavelength_nm:m.wavelength_nm,core_radius_um:m.core_radius_um,profile:m.profile,n_clad:claddingIndex({...PRESETS.g654,...m}),n_core:claddingIndex({...PRESETS.g654,...m})+(m.delta_n_core??.005),window_halfwidth_um:Math.max(18,m.core_radius_um*3),shells:m.shells}),history:[]}},warnings:['구버전 {version,model} 설계 이관 완료; CAD/M2/PML/과거 결과가 없는 필드는 현재 세션 기본값으로 유지','기존 scalar 경로의 모델 입력은 원래 값 그대로 복원됩니다']};
}
/** Project supermodes onto core-local transverse E integrals.
 * H_eff = U diag(beta) U^-1. Off-diagonal terms estimate coherent coupling.
 * This is a reduced modal model, not stochastic cable XT. Incomplete/ill-conditioned
 * mode sets are explicitly rejected; arbitrary scalar interpolation is never used.
 */
export function nCoreMetrics(result,{length_mm=1,inputCore=0}={}){
 const c=result.input,g=c.geometry,N=g?.count;
 if(g?.kind!=='mcf'||!Number.isInteger(N))return {status:'MCF only'};
 const centers=mcfCenters(N,g.pitch_um),modes=result.modes.filter(m=>m.Ex_intensity_fraction>.6).sort((a,b)=>b.neff-a.neff).slice(0,N);
 if(modes.length<N||modes.some(m=>!m.field_map?.field_components))return {status:'미검증: ≥N Ex-dominant supermodes with field maps required',cores:N,found:modes.length};
 const U=Array.from({length:N},()=>new Float64Array(2*N)),beta=modes.map(m=>[2*Math.PI*m.neff/(c.wavelength_nm*1e-9),2*Math.PI*m.neff_imag/(c.wavelength_nm*1e-9)]);
 for(let col=0;col<N;col++){
  const f=modes[col].field_map,xs=f.x_um,ys=f.y_um??xs,re=f.field_components.Ex_real,im=f.field_components.Ex_imag;
  for(let core=0;core<N;core++){
   const [cx,cy]=centers[core];let total=0,r=0,z=0;
   for(let j=0;j<ys.length;j++)for(let i=0;i<xs.length;i++)if(Math.hypot(xs[i]-cx,ys[j]-cy)<c.core_radius_um){r+=re[j][i];z+=im?.[j]?.[i]??0;total++;}
   U[core][2*col]=total?r/total:0;U[core][2*col+1]=total?z/total:0;
  }
 }
 const inverse=Array.from({length:N},()=>new Float64Array(2*N));
 for(let col=0;col<N;col++){const e=new Float64Array(2*N);e[2*col]=1;const x=denseSolve(U,e);for(let row=0;row<N;row++){inverse[row][2*col]=x[2*row];inverse[row][2*col+1]=x[2*row+1];}}
 const fnorm=A=>Math.sqrt(A.reduce((s,r)=>s+r.reduce((a,b)=>a+b*b,0),0)),condition=fnorm(U)*fnorm(inverse);
 if(!Number.isFinite(condition)||condition>1e6)return {status:'미검증: ill-conditioned core-mode projection',condition};
 const mul=(a,b)=>[a[0]*b[0]-a[1]*b[1],a[0]*b[1]+a[1]*b[0]],coupling=[],transfer=[];
 let powerTotal=0;
 for(let j=0;j<N;j++){
  for(let k=0;k<N;k++){
   let r=0,z=0;for(let m=0;m<N;m++){const b=mul([U[j][2*m],U[j][2*m+1]],mul(beta[m],[inverse[m][2*k],inverse[m][2*k+1]]));r+=b[0];z+=b[1];}
   if(j!==k)coupling.push({fromCore:k+1,toCore:j+1,kappaMagnitude_per_m:Math.hypot(r,z),real_per_m:r,imag_per_m:z});
  }
  let r=0,z=0;
  for(let m=0;m<N;m++){const L=length_mm/1000,attenuation=Math.exp(-beta[m][1]*L),phase=[attenuation*Math.cos(beta[m][0]*L),attenuation*Math.sin(beta[m][0]*L)],p=mul([U[j][2*m],U[j][2*m+1]],mul(phase,[inverse[m][2*inputCore],inverse[m][2*inputCore+1]]));r+=p[0];z+=p[1];}
  const power=r*r+z*z;transfer.push({core:j+1,projectedIntensity:power});powerTotal+=power;
 }
 const input=transfer[inputCore]?.projectedIntensity??0;
 for(const t of transfer){t.normalizedProjectedFraction=powerTotal?t.projectedIntensity/powerTotal:0;t.xt_dB=10*Math.log10(Math.max(1e-300,t.projectedIntensity)/Math.max(input,1e-300));}
 return {status:'미검증 reduced coherent model: core-average E projection; needs modal-basis convergence',cores:N,length_mm,inputCore:inputCore+1,condition,coupling,transfer,supermodeNeff:modes.map(m=>m.neff)};
}
