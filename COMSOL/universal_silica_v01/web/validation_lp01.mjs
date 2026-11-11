/**
 * M1 independent scalar LP01 oracle for a lossless circular step-index fiber.
 * No dependency on solver.mjs and no reference numbers used in root finding.
 * Solves u*J1(u)/J0(u) = w*K1(w)/K0(w), u^2+w^2=V^2.
 * J0/J1: convergent Taylor series (0 < u < first J0 root).
 * K1/K0: positive integral representation, composite Simpson quadrature.
 */
const J0_FIRST_ZERO=2.404825557695773;
function assertFinitePositive(v,name){
 if(!Number.isFinite(v)||v<=0)throw new Error(name+" must be finite and positive.");
}
function j01(x){
 const z=-(x*x)/4;
 let t0=1,s0=1,t1=x/2,s1=t1;
 for(let k=1;k<=90;k++){
  t0*=z/(k*k);
  t1*=z/(k*(k+1));
  s0+=t0;s1+=t1;
  if(Math.max(Math.abs(t0),Math.abs(t1))<1e-17)break;
 }
 return [s0,s1];
}
/** Returns K1(x)/K0(x) without subtracting exponentially small numbers. */
function k1OverK0(x){
 assertFinitePositive(x,"w");
 // exp(-x*(cosh(t)-1)) cancels exp(-x) in K0 and K1.
 // x*cosh(T) > ~50 makes the neglected integral exponentially small.
 const T=Math.max(9,Math.acosh(1+50/x)),N=1024,h=T/N;
 let a=0,b=0;
 for(let i=0;i<=N;i++){
  const ch=Math.cosh(i*h),f=Math.exp(-x*(ch-1));
  const weight=(i===0||i===N)?1:(i%2?4:2);
  a+=weight*f;b+=weight*f*ch;
 }
 if(!(a>0&&Number.isFinite(a)&&Number.isFinite(b)))throw new Error("K-Bessel quadrature failed.");
 return b/a;
}
export function lp01Characteristic({wavelength_nm,core_radius_um,n_core,n_clad}){
 for(const [name,v] of Object.entries({wavelength_nm,core_radius_um,n_core,n_clad}))assertFinitePositive(v,name);
 if(!(n_core>n_clad))throw new Error("Core index must exceed cladding index.");
 const k0=2*Math.PI/(wavelength_nm/1000);
 const V=k0*core_radius_um*Math.sqrt(n_core*n_core-n_clad*n_clad);
 if(!(V>=0.5&&V<=15))throw new Error("LP01 oracle is conditioned for 0.5 <= V <= 15.");
 const hi=Math.min(V,J0_FIRST_ZERO)*(1-1e-12);
 const residual=u=>{
  const [J0,J1]=j01(u),w=Math.sqrt(Math.max(0,V*V-u*u));
  if(!(J0>0&&w>0))throw new Error("LP01 root bracket became singular.");
  return u*J1/J0-w*k1OverK0(w);
 };
 let lo=Math.max(1e-10,hi*1e-9),upper=hi;
 if(!(residual(lo)<0&&residual(upper)>0))throw new Error("Cannot bracket fundamental LP01 root.");
 for(let i=0;i<75;i++){
  const m=(lo+upper)/2;
  if(residual(m)>0)upper=m;else lo=m;
  if(upper-lo<4e-14)break;
 }
 const u=(lo+upper)/2,w=Math.sqrt(V*V-u*u);
 const neff=Math.sqrt(n_core*n_core-(u/(k0*core_radius_um))**2);
 const b=(neff*neff-n_clad*n_clad)/(n_core*n_core-n_clad*n_clad);
 return {n_eff:neff,V,u,w,b,residual:residual(u),method:"LP01 scalar characteristic equation"};
}
