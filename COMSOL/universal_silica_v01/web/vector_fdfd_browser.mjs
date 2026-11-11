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
