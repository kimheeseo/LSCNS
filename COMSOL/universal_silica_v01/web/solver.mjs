/**
 * Universal Silica Fiber v0.2 — client-side radial P1 FEM.
 * Mirrors the physical model in universal_silica_v01/universal_radial.py.
 * Only rotationally symmetric, weak-guidance scalar LP0 (dominant LP01).
 * Units: transverse micrometres, wavelength nanometres, n2 m²/W.
 * No COMSOL code, server, static lookup, or manufacturer-profile claim.
 */
export const VERSION = '0.5.0';
export const C0 = 299792458;
const B=[0.6961663,0.4079426,0.8974794];
const C=[0.0684043**2,0.1162414**2,9.896161**2];
const qx=[0.046910077030668,0.230765344947158,0.5,0.769234655052841,0.953089922969332];
const qw=[0.118463442528095,0.239314335249683,0.284444444444444,0.239314335249683,0.118463442528095];
export const PRESETS = {
  'comsol_step':{profile:'step',wavelength_nm:1550,core_radius_um:8,delta_n_core:0.0079,outer_radius_um:40,clad_index_override:1.4378,graded_alpha:2,trench_start_um:10,trench_end_um:12,trench_delta_n:-0.001,shells:[[4,0.006],[5,0.002],[7,0],[9,-0.0008]],n2_m2_per_W:2.6e-20,elements:750},
  'g654':{profile:'step', wavelength_nm:1550,core_radius_um:5.5,delta_n_core:0.0034785793846155,outer_radius_um:40,graded_alpha:2,trench_start_um:7,trench_end_um:10,trench_delta_n:-0.001,shells:[[4,0.006],[5,0.002],[7,0],[9,-0.0008]],n2_m2_per_W:2.2e-20,elements:750},
  'g652':{profile:'step',wavelength_nm:1550,core_radius_um:4.1,delta_n_core:0.005,outer_radius_um:40,graded_alpha:2,trench_start_um:7,trench_end_um:10,trench_delta_n:-0.001,shells:[[4,0.006],[5,0.002],[7,0],[9,-0.0008]],n2_m2_per_W:2.6e-20,elements:750},
  'graded':{profile:'graded',wavelength_nm:1550,core_radius_um:5.0,delta_n_core:0.006,outer_radius_um:40,graded_alpha:2,trench_start_um:7,trench_end_um:10,trench_delta_n:-0.001,shells:[[4,0.006],[5,0.002],[7,0],[9,-0.0008]],n2_m2_per_W:2.6e-20,elements:750},
  'trench':{profile:'trench',wavelength_nm:1550,core_radius_um:4.1,delta_n_core:0.005,outer_radius_um:40,graded_alpha:2,trench_start_um:6,trench_end_um:9,trench_delta_n:-0.001,shells:[[4,0.006],[5,0.002],[7,0],[9,-0.0008]],n2_m2_per_W:2.6e-20,elements:750},
  'multi':{profile:'multilayer',wavelength_nm:1550,core_radius_um:4.0,delta_n_core:0.006,outer_radius_um:40,graded_alpha:2,trench_start_um:7,trench_end_um:10,trench_delta_n:-0.001,shells:[[4,0.006],[5,0.002],[7,0],[9,-0.0008]],n2_m2_per_W:2.6e-20,elements:750}
};
export function silicaIndex(wavelengthNm){const w=Number(wavelengthNm)/1000;if(!Number.isFinite(w)||w<=0.21||w>=3.7)throw Error('Silica Sellmeier 유효 범위는 210~3700 nm입니다.'); const w2=w*w;return Math.sqrt(1+B.reduce((v,b,i)=>v+b*w2/(w2-C[i]),0));}
/** Optional independently specified cladding index for matched material benchmarks.
 * Standard production presets still use wavelength-dependent silica Sellmeier.
 * Fixed-index mode must not be interpreted as a measured dispersion model.
 */
export function claddingIndex(m){
 if(m.clad_index_override!==undefined&&m.clad_index_override!==null){
  const n=m.clad_index_override;
  if(!Number.isFinite(n)||n<1||n>2)throw Error('고정 클래딩 굴절률은 1~2 범위의 유한한 수여야 합니다.');
  return n;
 }
 return silicaIndex(m.wavelength_nm);
}
export function validate(m){
 if(!['step','graded','trench','multilayer'].includes(m.profile))throw Error('지원하지 않는 단면입니다.');
 for(const [key,min] of [['wavelength_nm',210],['core_radius_um',0],['outer_radius_um',0],['graded_alpha',0],['n2_m2_per_W',0]]){if(!(Number.isFinite(m[key])&&m[key]>min))throw Error(`${key} 입력값을 확인하세요.`);}
 if(m.wavelength_nm>=3700)throw Error('파장이 3700 nm를 초과합니다.');
 if(m.outer_radius_um<=2.2*m.core_radius_um)throw Error('계산 영역 반경은 코어 반경의 2.2배보다 커야 합니다.');
 if(!Number.isFinite(m.delta_n_core)||!Number.isFinite(m.trench_delta_n))throw Error('굴절률 차이가 유효하지 않습니다.');
 if(!Number.isInteger(m.elements)||m.elements<80||m.elements>5000)throw Error('Mesh 수는 80~5000 사이의 정수여야 합니다.');
 if(m.profile==='trench'&&!(m.core_radius_um<m.trench_start_um&&m.trench_start_um<m.trench_end_um&&m.trench_end_um<m.outer_radius_um))throw Error('Trench 경계는 코어 < 시작 < 종료 < 계산 영역이어야 합니다.');
 if(m.profile==='multilayer'){
   if(!Array.isArray(m.shells)||!m.shells.length)throw Error('다층 쉘 데이터를 입력하세요.');
   let p=0;for(const [end,dn] of m.shells){if(!(Number.isFinite(end)&&end>p&&end<m.outer_radius_um&&Number.isFinite(dn)))throw Error('쉘 외경은 증가해야 하며 계산 영역보다 작아야 합니다.');p=end;}
 }
 const ncl=claddingIndex(m);let maxDn=m.delta_n_core;
 if(m.profile==='multilayer')maxDn=Math.max(...m.shells.map(x=>x[1]));
 if(maxDn<=0)throw Error('기본 유도 모드를 위한 양의 코어 Δn이 필요합니다.');
 if(m.profile==='trench'&&m.trench_delta_n>=0)throw Error('Trench Δn은 음수로 입력하세요.');
 if(ncl+Math.min(-0.01,m.trench_delta_n)<0)throw Error('물리적으로 유효하지 않은 굴절률입니다.');
 return m;
}
export function edges(m){const arr=m.profile==='trench'?[m.core_radius_um,m.trench_start_um,m.trench_end_um]:m.profile==='multilayer'?m.shells.map(x=>x[0]):[m.core_radius_um];return [0,...arr,m.outer_radius_um];}
export function deltaIndex(m,r){switch(m.profile){case 'step':return r<m.core_radius_um?m.delta_n_core:0;case 'graded':return r<m.core_radius_um?m.delta_n_core*Math.max(0,1-(r/m.core_radius_um)**m.graded_alpha):0;case 'trench':return r<m.core_radius_um?m.delta_n_core:(r>=m.trench_start_um&&r<m.trench_end_um?m.trench_delta_n:0);case 'multilayer':{let p=0;for(const [end,dn] of m.shells){if(r>=p&&r<end)return dn;p=end;}return 0;}default:return 0;}}
export function radialGrid(m){const br=edges(m), intervals=br.slice(1).map((x,i)=>x-br[i]);const total=intervals.reduce((a,b)=>a+b,0);const r=[];for(let j=0;j<intervals.length;j++){const n=Math.max(12,Math.round(m.elements*intervals[j]/total));for(let i=0;i<n;i++)r.push(br[j]+intervals[j]*i/n);}r.push(br.at(-1));return Float64Array.from(r);}
function tridiagSolve(diag,off,rhs){
 const N=diag.length;const p=new Float64Array(N),z=new Float64Array(N);
 let d=diag[0];if(!(d>0))throw Error('FEM 선형 시스템이 양의 정부호가 아닙니다.');p[0]=N>1?off[0]/d:0;z[0]=rhs[0]/d;
 for(let i=1;i<N;i++){d=diag[i]-off[i-1]*p[i-1];if(!(d>0))throw Error('Cholesky 분해 실패: σ 또는 구조 조건을 확인하세요.');p[i]=i<N-1?off[i]/d:0;z[i]=(rhs[i]-off[i-1]*z[i-1])/d;}
 for(let i=N-2;i>=0;i--)z[i]-=p[i]*z[i+1];return z;
}
function tridiagMv(d,o,x){const n=d.length,y=new Float64Array(n);for(let i=0;i<n;i++)y[i]=d[i]*x[i]+(i?o[i-1]*x[i-1]:0)+(i<n-1?o[i]*x[i+1]:0);return y;}
function dot(a,b){let s=0;for(let i=0;i<a.length;i++)s+=a[i]*b[i];return s;}
export function solve(m){
 validate(m);
 const r=radialGrid(m),N=r.length-1,k0=2*Math.PI/(m.wavelength_nm/1000),ncl=claddingIndex(m);
 const ad=new Float64Array(N),ao=new Float64Array(N-1),bd=new Float64Array(N),bo=new Float64Array(N-1);
 let maxDn=-Infinity;
 for(let e=0;e<N;e++){
   const r0=r[e],h=r[e+1]-r0;
   const stiff=(r[e+1]*r[e+1]-r0*r0)/(2*h*h);
   let mass00=0,mass01=0,mass11=0,n00=0,n01=0,n11=0;
   for(let j=0;j<5;j++){
      const t=qx[j],radius=r0+h*t,wt=qw[j]*h*radius,pa=1-t,pb=t;const dn=deltaIndex(m,radius);const n=ncl+dn;
      maxDn=Math.max(maxDn,dn);
      mass00+=wt*pa*pa;mass01+=wt*pa*pb;mass11+=wt*pb*pb;
      n00+=wt*n*n*pa*pa;n01+=wt*n*n*pa*pb;n11+=wt*n*n*pb*pb;
   }
   ad[e]+=k0*k0*n00-stiff;bd[e]+=mass00;
   if(e+1<N){ad[e+1]+=k0*k0*n11-stiff;bd[e+1]+=mass11;ao[e]+=k0*k0*n01+stiff;bo[e]+=mass01;}
 }
 const sigma=(k0*(ncl+maxDn))**2+0.05;
 const hd=Float64Array.from(ad,(x,i)=>sigma*bd[i]-x),ho=Float64Array.from(ao,(x,i)=>sigma*bo[i]-x);
 let x=new Float64Array(N).fill(1),lambda=0,iter=0,conv=false;
 const norm=x=>Math.sqrt(dot(x,tridiagMv(bd,bo,x)));
 {const s=norm(x);for(let i=0;i<N;i++)x[i]/=s;}
 for(iter=0;iter<180;iter++){
   const rhs=tridiagMv(bd,bo,x),y=tridiagSolve(hd,ho,rhs),sn=norm(y);
   for(let i=0;i<N;i++)y[i]/=sn;
   if(y[0]<0)for(let i=0;i<N;i++)y[i]=-y[i];
   let dx=0;for(let i=0;i<N;i++){const d=y[i]-x[i];dx+=d*d;}
   x=y;
   if(dx<1e-20){conv=true;break;}
 }
 const Ax=tridiagMv(ad,ao,x),Bx=tridiagMv(bd,bo,x);
 lambda=dot(x,Ax)/dot(x,Bx);if(!(lambda>0))throw Error('양의 β² 고유모드를 찾지 못했습니다.');
 let residual2=0,den2=0;for(let i=0;i<N;i++){residual2+=(Ax[i]-lambda*Bx[i])**2;den2+=(lambda*Bx[i])**2;}
 const residual=Math.sqrt(residual2/den2),neff=Math.sqrt(lambda)/k0;
 const psi=new Float64Array(r.length);psi.set(x);
 let P=0,Q=0,Rs=0,grad=0,P0=0;
 for(let e=0;e<N;e++){
   const h=r[e+1]-r[e],slope=(psi[e+1]-psi[e])/h;
   for(let j=0;j<5;j++){
      const t=qx[j],rad=r[e]+h*t,wt=h*qw[j]*rad,field=(1-t)*psi[e]+t*psi[e+1],intensity=field*field;
      P0+=wt*intensity;P+=2*Math.PI*wt*intensity;Q+=2*Math.PI*wt*intensity*intensity;
      Rs+=2*Math.PI*wt*intensity*rad*rad;grad+=wt*slope*slope;
   }
 }
 if(!(P>0&&Q>0&&grad>0))throw Error('모드 적분에 실패했습니다.');
 const Aeff=P*P/Q, MFD2=2*Math.sqrt(2*Rs/P),MFDPII=2*Math.sqrt(2*P0/grad);
 const normP=Math.sqrt(P);for(let i=0;i<psi.length;i++)psi[i]/=normP;
 const gamma=2*Math.PI/(m.wavelength_nm*1e-9)*m.n2_m2_per_W/(Aeff*1e-12)*1000;
 const guided=neff>ncl&&neff<ncl+maxDn+1e-7;
 return {profile:m.profile,wavelength_nm:m.wavelength_nm,n_clad:ncl,n_eff:neff,n_max:ncl+maxDn,Aeff_um2:Aeff,MFD_2nd_um:MFD2,MFD_PII_um:MFDPII,gamma_Wkm:gamma,residual,elements:N,iterations:iter+1,iteration_converged:conv,guided_proxy:guided,r:Array.from(r),field:Array.from(psi)};
}
export function dispersion(m,hNm=10){
 if(!(hNm>0&&m.wavelength_nm-2*hNm>211&&m.wavelength_nm+2*hNm<3690))throw Error('파장 미분 간격을 줄이세요.');
 const ns=[-2,-1,0,1,2].map(j=>solve({...m,wavelength_nm:m.wavelength_nm+j*hNm}).n_eff);
 const d2=(-ns[4]+16*ns[3]-30*ns[2]+16*ns[1]-ns[0])/(12*(hNm/1000)**2);
 return -(m.wavelength_nm/1000)/C0*1e12*d2;
}
export function spectrum(m,start,end,step){
 if(!(Number.isFinite(start)&&Number.isFinite(end)&&Number.isFinite(step)&&start<end&&step>=2&&end-start<=600))throw Error('스윕 범위(최대 600 nm) 및 스텝(최소 2 nm)을 확인하세요.');
 const count=Math.floor((end-start)/step)+1;if(count>70)throw Error('파장 스윕은 최대 70포인트입니다.');
 const rows=[];for(let i=0;i<count;i++){
  const nm=start+i*step;if(nm-2*step<=211||nm+2*step>=3690)throw Error('스윕 범위가 재료 모델 파장 제한 밖입니다.');
  const mm={...m,wavelength_nm:nm}; const s=solve(mm);
  const D=dispersion(mm,step);
  rows.push({wavelength_nm:nm,n_eff:s.n_eff,MFD_PII_um:s.MFD_PII_um,Aeff_um2:s.Aeff_um2,D_ps_nm_km:D,gamma_Wkm:s.gamma_Wkm});
 }
 return rows;
}
export function meshStudy(m,levels=[200,350,550,850]){
 return levels.map(el=>{const x=solve({...m,elements:el});return {elements:x.elements,n_eff:x.n_eff,MFD_PII_um:x.MFD_PII_um,Aeff_um2:x.Aeff_um2,residual:x.residual};});
}
/**
 * Dispersion slopes for the current hypothetical circular silica-index model.
 * S_1550 = dD/dlambda at 1550 nm, S0 = dD/dlambda where D(lambda0)=0.
 * The two quantities must never be used interchangeably.
 * Not a measured manufacturer RIP or product qualification.
 * Differentiation uses the same five-point n_eff stencil as dispersion().
 */
export function dispersionDiagnostics(m, opts={}) {
 validate(m);
 const from=opts.fromNm??1100, to=opts.toNm??1550;
 const step=opts.scanStepNm??50, h=opts.diffStepNm??5, ds=opts.slopeStepNm??10;
 if(!(from>=250 && to<=3400 && to>from && step>=5 && step<=150 && h>=2 && ds>=5))
  throw Error('영분산 스캔 범위/미분 간격을 확인하세요.');
 const neffCache=new Map(), dCache=new Map();
 const nAt=nm=>{
  const k=nm.toFixed(8);
  if(!neffCache.has(k)){
   const s=solve({...m,wavelength_nm:nm});
   if(!s.guided_proxy || !Number.isFinite(s.n_eff))throw Error('파장 스윕에서 유도 모드를 찾지 못했습니다.');
   neffCache.set(k,s.n_eff);
  }
  return neffCache.get(k);
 };
 const dAt=nm=>{
  const k=nm.toFixed(8);
  if(!dCache.has(k)){
   if(nm-2*h<=211||nm+2*h>=3690)throw Error('분산 계산이 Sellmeier 파장 유효범위를 벗어납니다.');
   const vals=[-2,-1,0,1,2].map(j=>nAt(nm+j*h));
   const d2=(-vals[4]+16*vals[3]-30*vals[2]+16*vals[1]-vals[0])/(12*(h/1000)**2);
   dCache.set(k,-(nm/1000)/C0*1e12*d2);
  }
  return dCache.get(k);
 };
 const slopeAt=nm=>(dAt(nm+ds)-dAt(nm-ds))/(2*ds);
 const s1550=slopeAt(1550);
 let low=from, dl=dAt(low), hi=null, dh=null;
 for(let nm=from+step; nm<=to+1e-9; nm+=step){
   const dv=dAt(nm);
   if(dl===0 || dv===0 || dl*dv<0){hi=nm;dh=dv;break;}
   low=nm;dl=dv;
 }
 if(hi===null) {
  return {slope_1550_ps_nm2_km:s1550,lambda0_nm:null,slope_zero_ps_nm2_km:null,
   zero_found:false,search_nm:[from,to],difference_step_nm:ds,neff_step_nm:h};
 }
 let root=null;
 if(dl===0)root=low;
 else if(dh===0)root=hi;
 else for(let i=0;i<16;i++){
   const sec=(low*dh-hi*dl)/(dh-dl);
   const mid=(low+hi)/2;
   const x=Number.isFinite(sec)&&sec>low+0.05*(hi-low)&&sec<hi-0.05*(hi-low)?sec:mid;
   const d=dAt(x);
   if(Math.abs(d)<1e-4 || hi-low<0.025){root=x;break;}
   if(dl*d<=0){hi=x;dh=d;}else{low=x;dl=d;}
 }
 if(root===null)root=(low+hi)/2;
 return {slope_1550_ps_nm2_km:s1550,lambda0_nm:root,slope_zero_ps_nm2_km:slopeAt(root),
  zero_found:true,search_nm:[from,to],difference_step_nm:ds,neff_step_nm:h};
}

export function outputSummary(sol){const {r,field,...rest}=sol;return rest;}
