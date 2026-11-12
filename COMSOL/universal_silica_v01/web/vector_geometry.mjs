/** M3–M5 common 2D geometry. Length µm; wavelength nm; e^(iβz-iωt). */
import {lp01Characteristic} from './validation_lp01.mjs';
import {evaluateM2Material} from './m2_preview_engine_v06.mjs?v=20261112-m3m6-v1';
import {DEFAULT as HCF_DEFAULT,buildGeometry} from './hcf_designer.mjs';
export const LIMITS=Object.freeze({minGrid:24,maxGrid:192,maxModes:16,maxSweep:36,maxMask:512});
export const BASE={wavelength_nm:1550,grid_n:48,window_halfwidth_um:18,core_radius_um:4.1,
 n_core:1.449023621703261,n_clad:1.444023621703261,profile:'step',subpixel:true,
 modes:2,target_neff:null,pml:{enabled:false,thickness_um:4,strength:3,order:3},bend_radius_um:0,
 geometry:{kind:'radial'},materials:null};
export function config(input={}){
 const c={...BASE,...structuredClone(input),pml:{...BASE.pml,...input.pml},geometry:{...BASE.geometry,...input.geometry}};
 for(const k of ['wavelength_nm','window_halfwidth_um','core_radius_um','n_core','n_clad','bend_radius_um'])
  if(!Number.isFinite(c[k]))throw Error(k+' must be finite');
 if(!(c.wavelength_nm>=400&&c.wavelength_nm<=1700))throw Error('M2 wavelength range: 400–1700 nm');
 if(!(c.window_halfwidth_um>0&&c.core_radius_um>0))throw Error('Positive domain/radius required');
 if(!Number.isInteger(c.grid_n)||c.grid_n<LIMITS.minGrid||c.grid_n>LIMITS.maxGrid)throw Error('Yee grid: 24–192 cells per axis');
 if(!Number.isInteger(c.modes)||c.modes<1||c.modes>LIMITS.maxModes)throw Error('Modes: 1–16');
 if(c.bend_radius_um<0||(c.bend_radius_um>0&&c.bend_radius_um<10*c.window_halfwidth_um))throw Error('Conformal bend requires R ≥ 10 × halfwidth; R=0 is straight');
 if(c.target_neff!=null&&!(Number.isFinite(c.target_neff)&&c.target_neff>0&&c.target_neff<=3))throw Error('target neff must be 0–3');
 if(c.pml.enabled){
  const p=c.pml,h=2*c.window_halfwidth_um/c.grid_n;
  if(!Number.isFinite(p.thickness_um)||p.thickness_um<3*h||p.thickness_um>=c.window_halfwidth_um*.6)throw Error('PML thickness: ≥ 3 cells and < 60% halfwidth');
  if(!(p.strength>0&&p.strength<=30&&Number.isInteger(p.order)&&p.order>=1&&p.order<=6))throw Error('PML strength (0,30], profile order 1–6');
 }
 if(!['radial','regions','mask','mcf','hcf'].includes(c.geometry.kind))throw Error('Unknown geometry kind');
 if(c.geometry.kind==='hcf'&&!c.pml.enabled)throw Error('HCF/NANF requires M4 PML; M3 alone does not support HCF');
 return c;
}
export function materialValue(m,w){
 if(m?.kind&&m.kind!=='constant')return evaluateM2Material(m,w);
 const n=Number(m?.n??1),kappa=Number(m?.kappa??0);
 if(!(Number.isFinite(n)&&n>=1&&n<=3&&Number.isFinite(kappa)&&kappa>=0&&kappa<.1))throw Error('Material n [1,3], κ [0,0.1) required');
 return {n,kappa};
}
export function polygonContains(p,x,y){
 let inside=false;
 for(let i=0,j=p.length-1;i<p.length;j=i++){
  const a=p[i],b=p[j];
  if((a[1]>y)!==(b[1]>y)&&x<(b[0]-a[0])*(y-a[1])/(b[1]-a[1])+a[0])inside=!inside;
 }
 return inside;
}
export function mcfCenters(count,pitch){
 if(!Number.isInteger(count)||count<2||count>19||!Number.isFinite(pitch)||pitch<=0)throw Error('MCF: 2–19 cores, positive pitch');
 if(count===2)return [[-pitch/2,0],[pitch/2,0]];
 const hex=[[0,0]];
 for(let radius=1;hex.length<count;radius++){
  const ring=[];
  for(let q=-radius;q<=radius;q++)for(let r=-radius;r<=radius;r++)if(Math.max(Math.abs(q),Math.abs(r),Math.abs(q+r))===radius)ring.push([pitch*(q+r/2),pitch*Math.sqrt(3)*r/2]);
  ring.sort((a,b)=>Math.atan2(a[1],a[0])-Math.atan2(b[1],b[0]));hex.push(...ring);
 }
 return hex.slice(0,count);
}
/** Analytic shapes and pixel masks use the same sampling path at all Yee offsets. */
export function sampler(c){
 const g=c.geometry,w=c.wavelength_nm;
 const materials={clad:{kind:'constant',n:c.n_clad,kappa:c.k_clad??0},core:{kind:'constant',n:c.n_core,kappa:c.k_core??0},air:{kind:'constant',n:g.hcf?.airIndex??1.00027,kappa:0},...c.materials};
 const values=Object.fromEntries(Object.entries(materials).map(([key,m])=>[key,materialValue(m,w)]));
 const get=id=>{if(!values[id])throw Error('Unassigned material: '+id);return values[id];};
 let centers=null,hcf=null;
 if(g.kind==='mcf'){
  centers=mcfCenters(g.count??2,g.pitch_um??16);
  if((g.pitch_um??16)<=2*c.core_radius_um)throw Error('MCF cores overlap');
 }
 if(g.kind==='regions'){
  if(!Array.isArray(g.regions)||g.regions.length>128)throw Error('0–128 polygon/circle regions required');
  for(const region of g.regions){
   get(region.material);
   if(region.kind==='circle'){
    if(![region.x??0,region.y??0,region.radius].every(Number.isFinite)||region.radius<=0)throw Error('Invalid circle region');
   }else if(!Array.isArray(region.points)||region.points.length<3||region.points.length>4096||!region.points.every(p=>p.length===2&&p.every(Number.isFinite)))throw Error('Invalid polygon vertices');
  }
 }
 if(g.kind==='mask'){
  if(!Number.isInteger(g.width)||!Number.isInteger(g.height)||g.width<1||g.height<1||g.width>512||g.height>512||g.labels?.length!==g.width*g.height||!(g.halfwidth_um>0))throw Error('Mask: ≤512×512 with physical halfwidth and palette');
  for(const m of Object.values(g.palette||{}))get(m);
 }
 if(g.kind==='hcf')hcf=buildGeometry({...HCF_DEFAULT,...g.hcf,wavelengthNm:w});
 function at(x,y){
  if(g.kind==='mask'){
   const i=Math.floor((x/g.halfwidth_um+1)*g.width/2),j=Math.floor((1-y/(g.halfheight_um??g.halfwidth_um*g.height/g.width))*g.height/2);
   if(i<0||j<0||i>=g.width||j>=g.height)return get(g.background??'clad');
   return get(g.palette[g.labels[j*g.width+i]]??g.background??'clad');
  }
  if(g.kind==='hcf'){
   for(const ring of hcf.rings)for(const layer of ring.layers){const r=Math.hypot(x-layer.x,y-layer.y);if(r<layer.radius&&r>=layer.radius-layer.wall)return get(g.glass??'clad');}
   return get(g.air??'air');
  }
  if(g.kind==='regions'){
   for(let i=g.regions.length-1;i>=0;i--){const p=g.regions[i];if(p.kind==='circle'?Math.hypot(x-(p.x??0),y-(p.y??0))<p.radius:polygonContains(p.points,x,y))return get(p.material);}
   return get(g.background??'clad');
  }
  if(g.kind==='mcf')return centers.some(([cx,cy])=>Math.hypot(x-cx,y-cy)<c.core_radius_um)?get(g.core??'core'):get(g.background??'clad');
  const r=Math.hypot(x,y),core=get('core'),clad=get('clad');
  if(c.profile==='graded')return {n:clad.n+(core.n-clad.n)*Math.max(0,1-(r/c.core_radius_um)**(c.graded_alpha??2)),kappa:r<c.core_radius_um?core.kappa:clad.kappa};
  if(c.profile==='trench'&&r>=(c.trench_start_um??6)&&r<(c.trench_end_um??9))return {n:clad.n+(c.trench_delta_n??-.001),kappa:clad.kappa};
  if(c.profile==='multilayer'){for(const [end,delta]of c.shells||[])if(r<end)return {n:clad.n+delta,kappa:clad.kappa};return clad;}
  return r<c.core_radius_um?core:clad;
 }
 const warnings=[];
 const extent=centers?Math.max(...centers.map(p=>Math.hypot(...p)))+c.core_radius_um:hcf?.extentUm;
 if(extent&&extent>=c.window_halfwidth_um-(c.pml.enabled?c.pml.thickness_um:0))throw Error('Geometry intersects PML/domain edge; enlarge window');
 if(hcf&&Math.min(...hcf.wallsUm)<3*2*c.window_halfwidth_um/c.grid_n)warnings.push('미검증: HCF wall has fewer than 3 cells. Refine mesh; subpixel cannot establish loss accuracy.');
 return {at,values,warnings,centers,hcf};
}
export function buildYeeGrid(input){
 const c=config(input),s=sampler(c),n=c.grid_n,K=n*n,h=2*c.window_halfwidth_um/n;
 const arrays=Array.from({length:3},()=>new Float64Array(2*K)),index=new Float64Array(K);
 const count=c.subpixel?4:1;
 for(let axis=0;axis<3;axis++)for(let j=0;j<n;j++)for(let i=0;i<n;i++){
  const x=-c.window_halfwidth_um+(i+.5+(axis===0?.5:0))*h,y=-c.window_halfwidth_um+(j+.5+(axis===1?.5:0))*h;
  let er=0,ei=0;
  for(let b=0;b<count;b++)for(let a=0;a<count;a++){
   const xx=x+((a+.5)/count-.5)*h,yy=y+((b+.5)/count-.5)*h,m=s.at(xx,yy);
   const bend=c.bend_radius_um?Math.exp(2*xx/c.bend_radius_um):1;
   er+=(m.n*m.n-m.kappa*m.kappa)*bend;ei+=2*m.n*m.kappa*bend;
  }
  const q=j*n+i;arrays[axis][2*q]=er/(count*count);arrays[axis][2*q+1]=ei/(count*count);
  if(axis===2)index[q]=Math.sqrt((Math.hypot(er,ei)+er)/2/(count*count));
 }
 let bendTarget=null;
 if(c.bend_radius_um>0&&c.geometry.kind==='radial'&&c.profile==='step'){try{bendTarget=lp01Characteristic({wavelength_nm:c.wavelength_nm,core_radius_um:c.core_radius_um,n_core:s.values.core.n,n_clad:s.values.clad.n}).n_eff;}catch{}}
 const target=c.target_neff??bendTarget??(c.geometry.kind==='hcf'?s.values.air.n-.0007:Math.max(...Object.values(s.values).map(m=>m.n))+.0002);
 return {config:c,n,K,h,epsX:arrays[0],epsY:arrays[1],epsZ:arrays[2],index,target,warnings:s.warnings,centers:s.centers};
}
export function lossDbKm(imagNeff,wavelengthNm){
 if(!Number.isFinite(imagNeff)||!(wavelengthNm>0))throw Error('Finite Im(neff), positive wavelength required');
 // k0 [m^-1] = 2π / (λ_nm × 1e-9); power loss = 20/ln(10) k0 Im(neff) ×1000.
 return 20/Math.log(10)*(2*Math.PI/(wavelengthNm*1e-9))*imagNeff*1000;
}
