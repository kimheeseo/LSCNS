/** Browser Maxwell Yee solver. No scalar field copied into vector components.
 * eps E / H curls are staggered, complex stretching replaces each derivative.
 * Matrix-free β²/k0² eigenproblem PQ E = neff² E. All lengths µm internally.
 */
import {sparseYee} from './yee_sparse.mjs';
import {buildYeeGrid,lossDbKm} from './vector_geometry.mjs';
export const ENGINE='Yee-FDFD 1.0 experimental';
const cmul=(a,b)=>[a[0]*b[0]-a[1]*b[1],a[0]*b[1]+a[1]*b[0]];
const cdiv=(a,b)=>{const d=b[0]*b[0]+b[1]*b[1];return [(a[0]*b[0]+a[1]*b[1])/d,(a[1]*b[0]-a[0]*b[1])/d];};
const csqrt=a=>{const r=Math.hypot(...a),re=Math.sqrt(Math.max(0,(r+a[0])/2));return [re,re>1e-30?a[1]/(2*re):Math.sign(a[1]||1)*Math.sqrt(r)];};
const norm=a=>{let s=0;for(const v of a)s+=v*v;return Math.sqrt(s);};
function dot(a,b){let r=0,i=0;for(let k=0;k<a.length;k+=2){r+=a[k]*b[k]+a[k+1]*b[k+1];i+=a[k]*b[k+1]-a[k+1]*b[k];}return [r,i];}
function axpy(y,a,x){const [r,i]=a;for(let k=0;k<y.length;k+=2){y[k]+=r*x[k]-i*x[k+1];y[k+1]+=r*x[k+1]+i*x[k];}return y;}
function scale(a,s){for(let k=0;k<a.length;k++)a[k]*=s;return a;}
function mulPoint(out,a,b,sign=1){for(let k=0;k<a.length;k+=2){out[k]=sign*(a[k]*b[k]-a[k+1]*b[k+1]);out[k+1]=sign*(a[k]*b[k+1]+a[k+1]*b[k]);}return out;}
function orth(v,basis){for(let pass=0;pass<2;pass++)for(const q of basis){const d=dot(q,v);axpy(v,[-d[0],-d[1]],q);}const r=norm(v);if(r<1e-15)throw Error('Invariant subspace collapsed');scale(v,1/r);return v;}
export function yeeOperator(grid){
 const {n,K,h,config:c,epsX,epsY,epsZ}=grid,k0=2*Math.PI/(c.wavelength_nm/1000),N=4*K;
 const sx=[],sy=[];
 for(const delta of [0,.5]){
  const s=new Float64Array(2*n);
  for(let i=0;i<n;i++){
   const pos=-c.window_halfwidth_um+(i+.5+delta)*h;
   const t=c.pml.enabled?Math.max(0,(Math.abs(pos)-c.window_halfwidth_um+c.pml.thickness_um)/c.pml.thickness_um):0;
   const a=c.pml.enabled?c.pml.strength*Math.min(1,t)**c.pml.order:0;
   s[2*i]=1/(1+a*a);s[2*i+1]=-a/(1+a*a);
  }
  sx.push(s);sy.push(s);
 }
 const invZ=new Float64Array(2*K);
 for(let q=0;q<K;q++){const a=2*q,d=epsZ[a]**2+epsZ[a+1]**2;invZ[a]=epsZ[a]/d;invZ[a+1]=-epsZ[a+1]/d;}
 function derivative(v,axis,forward,out){
  const s=(axis===0?sx:sy)[forward?1:0],step=axis===0?1:n;
  for(let j=0;j<n;j++)for(let i=0;i<n;i++){
   const q=j*n+i,a=2*q,p=axis===0?i:j,neighbor=forward?p<n-1:p>0;
   const b=2*(q+(forward?step:-step));
   const r=(forward?(neighbor?v[b]:0)-v[a]:v[a]-(neighbor?v[b]:0))/(h*k0),
    z=(forward?(neighbor?v[b+1]:0)-v[a+1]:v[a+1]-(neighbor?v[b+1]:0))/(h*k0),sr=s[2*p],si=s[2*p+1];
   out[a]=sr*r-si*z;out[a+1]=sr*z+si*r;
  }
  return out;
 }
 const tmp=Array.from({length:8},()=>new Float64Array(2*K));
 function magnetic(e,hx,hy){
  const ex=e.subarray(0,2*K),ey=e.subarray(2*K),[curl,d,dx,dy,mx,my]=tmp;
  derivative(ey,0,true,curl);derivative(ex,1,true,d);for(let a=0;a<curl.length;a++)curl[a]-=d[a];
  derivative(curl,0,false,dx);derivative(curl,1,false,dy);
  mulPoint(mx,epsY,ey);mulPoint(my,epsX,ex);
  for(let a=0;a<2*K;a++){hx[a]=-dx[a]-mx[a];hy[a]=-dy[a]+my[a];}
 }
 const hx=new Float64Array(2*K),hy=new Float64Array(2*K);
 function apply(e,out){
  magnetic(e,hx,hy);
  const [curl,d,dx,dy,, ,z]=tmp;
  derivative(hy,0,false,curl);derivative(hx,1,false,d);for(let a=0;a<2*K;a++)curl[a]-=d[a];
  mulPoint(z,invZ,curl);
  derivative(z,0,true,dx);derivative(z,1,true,dy);
  for(let a=0;a<2*K;a++){out[a]=dx[a]+hy[a];out[2*K+a]=dy[a]-hx[a];}
  return out;
 }
 const sigma=grid.target**2;
 const diagonal=new Float64Array(N);
 for(let q=0;q<K;q++)for(let field=0;field<2;field++){
  const a=2*q,eps=field===0?epsX:epsY;
  diagonal[field*2*K+a]=sigma-eps[a]+4/(h*k0)**2;
  diagonal[field*2*K+a+1]=-eps[a+1];
 }
 function shifted(e,out){apply(e,out);for(let a=0;a<N;a++)out[a]=sigma*e[a]-out[a];return out;}
 function precondition(v,out){for(let a=0;a<N;a+=2){const r=diagonal[a],i=diagonal[a+1],d=Math.max(1e-24,r*r+i*i);out[a]=(v[a]*r+v[a+1]*i)/d;out[a+1]=(v[a+1]*r-v[a]*i)/d;}return out;}
 return {apply,shifted,precondition,derivative,magnetic,N,K,k0,sigma,
  allocatedBytes:invZ.byteLength+diagonal.byteLength+tmp.reduce((s,a)=>s+a.byteLength,0)+hx.byteLength+hy.byteLength+sx.reduce((s,a)=>s+a.byteLength,0)};
}
/** Restarted left-preconditioned complex GMRES with verified unpreconditioned residual. */
export function gmres(op,b,{tolerance=2e-9,maxIterations=1000,restart=32}={}){
 const N=b.length,x=new Float64Array(N),r=new Float64Array(N),work=new Float64Array(N),tmp=new Float64Array(N);
 const bnorm=Math.max(1e-30,norm(b));let iterations=0;
 while(iterations<maxIterations){
  op.shifted(x,work);for(let k=0;k<N;k++)r[k]=b[k]-work[k];
  const actual=norm(r)/bnorm;if(actual<tolerance)return {x,iterations,residual:actual};
  op.precondition(r,work);const beta=norm(work);if(!Number.isFinite(beta)||beta===0)throw Error('GMRES invalid residual');
  const V=[scale(Float64Array.from(work),1/beta)],H=Array.from({length:restart+1},()=>new Float64Array(2*restart)),
   g=new Float64Array(2*(restart+1)),rot=[];g[0]=beta;let used=0;
  for(let j=0;j<restart&&iterations<maxIterations;j++,iterations++){
   op.shifted(V[j],tmp);op.precondition(tmp,work);
   for(let pass=0;pass<2;pass++)for(let i=0;i<=j;i++){
    const a=dot(V[i],work);H[i][2*j]+=a[0];H[i][2*j+1]+=a[1];axpy(work,[-a[0],-a[1]],V[i]);
   }
   const hn=norm(work);H[j+1][2*j]=hn;
   V.push(hn>1e-28?scale(Float64Array.from(work),1/hn):new Float64Array(N));
   for(let i=0;i<j;i++){
    const {c,s}=rot[i],a=[H[i][2*j],H[i][2*j+1]],b0=[H[i+1][2*j],H[i+1][2*j+1]],sb=cmul(s,b0),sa=cmul([s[0],-s[1]],a);
    H[i][2*j]=c*a[0]+sb[0];H[i][2*j+1]=c*a[1]+sb[1];H[i+1][2*j]=-sa[0]+c*b0[0];H[i+1][2*j+1]=-sa[1]+c*b0[1];
   }
   const a=[H[j][2*j],H[j][2*j+1]],b0=[H[j+1][2*j],H[j+1][2*j+1]],an=Math.hypot(...a),bn=Math.hypot(...b0),length=Math.hypot(an,bn),phase=an>1e-30?[a[0]/an,a[1]/an]:[1,0],c=an/length,s=cmul(phase,[b0[0]/length,-b0[1]/length]);
   rot.push({c,s});H[j][2*j]=phase[0]*length;H[j][2*j+1]=phase[1]*length;H[j+1][2*j]=H[j+1][2*j+1]=0;
   const ga=[g[2*j],g[2*j+1]],sg=cmul([s[0],-s[1]],ga);g[2*j]*=c;g[2*j+1]*=c;g[2*j+2]=-sg[0];g[2*j+3]=-sg[1];used=j+1;
   if(Math.hypot(g[2*j+2],g[2*j+3])<tolerance*beta||hn<1e-27){iterations++;break;}
  }
  const y=new Float64Array(2*used);
  for(let i=used-1;i>=0;i--){let a=[g[2*i],g[2*i+1]];for(let j=i+1;j<used;j++){const v=cmul([H[i][2*j],H[i][2*j+1]],[y[2*j],y[2*j+1]]);a[0]-=v[0];a[1]-=v[1];}const v=cdiv(a,[H[i][2*i],H[i][2*i+1]]);y[2*i]=v[0];y[2*i+1]=v[1];}
  for(let j=0;j<used;j++)axpy(x,[y[2*j],y[2*j+1]],V[j]);
  if(!Number.isFinite(norm(x)))throw Error('GMRES nonfinite iterate');
 }
 op.shifted(x,work);for(let k=0;k<N;k++)r[k]=b[k]-work[k];
 throw Error('GMRES did not converge: residual '+(norm(r)/bnorm).toExponential(2)+'; adjust shift/grid/PML');
}
/** Small dense complex eigensystem for Rayleigh–Ritz: shifted QR + inverse iteration. */
export function denseSolve(A,b){
 const n=b.length/2,a=A.map(r=>Float64Array.from(r)),x=Float64Array.from(b);
 for(let k=0;k<n;k++){
  let p=k;for(let i=k+1;i<n;i++)if(Math.hypot(a[i][2*k],a[i][2*k+1])>Math.hypot(a[p][2*k],a[p][2*k+1]))p=i;
  [a[k],a[p]]=[a[p],a[k]];for(let d=0;d<2;d++){const v=x[2*k+d];x[2*k+d]=x[2*p+d];x[2*p+d]=v;}
  if(Math.hypot(a[k][2*k],a[k][2*k+1])<1e-25)a[k][2*k]+=1e-20;
  for(let i=k+1;i<n;i++){
   const f=cdiv([a[i][2*k],a[i][2*k+1]],[a[k][2*k],a[k][2*k+1]]);
   for(let j=k;j<n;j++){const v=cmul(f,[a[k][2*j],a[k][2*j+1]]);a[i][2*j]-=v[0];a[i][2*j+1]-=v[1];}
   const v=cmul(f,[x[2*k],x[2*k+1]]);x[2*i]-=v[0];x[2*i+1]-=v[1];
  }
 }
 for(let i=n-1;i>=0;i--){let v=[x[2*i],x[2*i+1]];for(let j=i+1;j<n;j++){const s=cmul([a[i][2*j],a[i][2*j+1]],[x[2*j],x[2*j+1]]);v[0]-=s[0];v[1]-=s[1];}v=cdiv(v,[a[i][2*i],a[i][2*i+1]]);x[2*i]=v[0];x[2*i+1]=v[1];}return x;
}
function denseEigen(A){
 const n=A.length,B=A.map(r=>Float64Array.from(r)),vals=[];
 for(let active=n;active>0;active--){
  if(active===1){vals.unshift([B[0][0],B[0][1]]);break;}
  for(let iter=0;iter<500;iter++){
   let below=0;for(let j=0;j<active-1;j++)below+=Math.hypot(B[active-1][2*j],B[active-1][2*j+1]);
   if(below<1e-12)break;
   const a=[B[active-2][2*(active-2)],B[active-2][2*(active-2)+1]],d=[B[active-1][2*(active-1)],B[active-1][2*(active-1)+1]],b=[B[active-2][2*(active-1)],B[active-2][2*(active-1)+1]],cc=[B[active-1][2*(active-2)],B[active-1][2*(active-2)+1]],delta=[(a[0]-d[0])/2,(a[1]-d[1])/2],dsq=cmul(delta,delta),bc=cmul(b,cc),root=csqrt([dsq[0]+bc[0],dsq[1]+bc[1]]),mid=[(a[0]+d[0])/2,(a[1]+d[1])/2],v1=[mid[0]+root[0],mid[1]+root[1]],v2=[mid[0]-root[0],mid[1]-root[1]],shift=Math.hypot(v1[0]-d[0],v1[1]-d[1])<Math.hypot(v2[0]-d[0],v2[1]-d[1])?v1:v2;
   const Q=[],R=Array.from({length:active},()=>new Float64Array(2*active));
   for(let j=0;j<active;j++){
    const v=new Float64Array(2*active);for(let i=0;i<active;i++){v[2*i]=B[i][2*j]-(i===j?shift[0]:0);v[2*i+1]=B[i][2*j+1]-(i===j?shift[1]:0);}
    for(let i=0;i<j;i++){const a=dot(Q[i],v);R[i][2*j]=a[0];R[i][2*j+1]=a[1];axpy(v,[-a[0],-a[1]],Q[i]);}
    const vn=norm(v);R[j][2*j]=vn;
    if(vn<1e-18){v.fill(0);v[2*j]=1;orth(v,Q);}else scale(v,1/vn);Q.push(v);
   }
   for(let i=0;i<active;i++)for(let j=0;j<active;j++){
    let r=0,z=0;for(let k=0;k<active;k++){const p=cmul([R[i][2*k],R[i][2*k+1]],[Q[j][2*k],Q[j][2*k+1]]);r+=p[0];z+=p[1];}B[i][2*j]=r+(i===j?shift[0]:0);B[i][2*j+1]=z+(i===j?shift[1]:0);
   }
  }
  vals.unshift([B[active-1][2*(active-1)],B[active-1][2*(active-1)+1]]);
 }
 const eigen=[];
 for(let j=0;j<n;j++){
  const val=vals[j],near=eigen.filter(e=>Math.hypot(e.value[0]-val[0],e.value[1]-val[1])<1e-9).map(e=>e.vector);
  let v=new Float64Array(2*n);v[2*j]=1;for(let i=0;i<n;i++)v[2*i]+=.01*Math.sin((j+1)*(i+1));
  const shifted=A.map((row,i)=>{const r=Float64Array.from(row);r[2*i]-=val[0]-1e-11;r[2*i+1]-=val[1];return r;});
  for(let iter=0;iter<6;iter++)v=orth(denseSolve(shifted,v),near);
  eigen.push({value:val,vector:v});
 }
 return eigen;
}
function seed(grid,pol,order){
 const {n,K,h,config:c}=grid,v=new Float64Array(4*K),centers=grid.centers;
 for(let j=0;j<n;j++)for(let i=0;i<n;i++){
  const x=-c.window_halfwidth_um+(i+.5)*h,y=-c.window_halfwidth_um+(j+.5)*h;
  let amp=0;
  if(centers)centers.forEach(([cx,cy],k)=>{amp+=Math.cos(Math.PI*(order%centers.length)*(k+.5)/centers.length)*Math.exp(-((x-cx)**2+(y-cy)**2)/(c.core_radius_um**2))*(order<centers.length?1:(x-cx)/c.core_radius_um);});
  else{
   const width=c.core_radius_um*1.2;
   amp=Math.exp(-(x*x+y*y)/(width*width))*(order===0?1:order===1?x/width:order===2?y/width:Math.sin(order*x/width)+.3*Math.cos(order*y/width));
  }
  v[pol*2*K+2*(j*n+i)]=amp;
 }
 return v;
}
function ritz(op,Q){
 const MQ=Q.map(v=>op.apply(v,new Float64Array(op.N))),A=Q.map(q=>{const row=new Float64Array(2*Q.length);for(let j=0;j<Q.length;j++){const p=dot(q,MQ[j]);row[2*j]=p[0];row[2*j+1]=p[1];}return row;});
 return denseEigen(A).map(e=>{
  const v=new Float64Array(op.N),mv=new Float64Array(op.N);
  for(let j=0;j<Q.length;j++){const a=[e.vector[2*j],e.vector[2*j+1]];axpy(v,a,Q[j]);axpy(mv,a,MQ[j]);}
  const vnorm=norm(v);scale(v,1/vnorm);scale(mv,1/vnorm);const value=dot(v,mv);
  axpy(mv,[-value[0],-value[1]],v);
  return {vector:v,value,residual:norm(mv)/Math.max(1e-20,Math.hypot(...value))};
 }).sort((a,b)=>Math.hypot(a.value[0]-op.sigma,a.value[1])-Math.hypot(b.value[0]-op.sigma,b.value[1]));
}
function fields(e,grid,op){
 const {K,n,h,config:c}=grid,neff=csqrt(e.value),hx=new Float64Array(2*K),hy=new Float64Array(2*K),z=new Float64Array(2*K),tmp=new Float64Array(2*K),ez=new Float64Array(2*K);
 op.magnetic(e.vector,hx,hy); // Q E = neff H; divide by neff before curl H.
 for(const f of [hx,hy])for(let q=0;q<K;q++){const v=cdiv([f[2*q],f[2*q+1]],neff);f[2*q]=v[0];f[2*q+1]=v[1];}
 op.derivative(hy,0,false,z);op.derivative(hx,1,false,tmp);
 for(let q=0;q<K;q++){const a=2*q,v=cdiv([-(z[a+1]-tmp[a+1]),z[a]-tmp[a]],[grid.epsZ[a],grid.epsZ[a+1]]);ez[a]=v[0];ez[a+1]=v[1];}
 const components={},maps={Itotal:new Float64Array(K)},fractions=[];
 for(const [f,offset]of [['Ex',0],['Ey',2*K],['Ez',null]]){
  const re=new Float64Array(K),im=new Float64Array(K),intensity=new Float64Array(K);
  for(let j=0;j<n;j++)for(let i=0;i<n;i++){
   const q=j*n+i,a=2*q;
   if(offset===null){re[q]=ez[a];im[q]=ez[a+1];}
   else{const prev=f==='Ex'?(i>0?q-1:-1):(j>0?q-n:-1);re[q]=.5*(e.vector[offset+a]+(prev<0?0:e.vector[offset+2*prev]));im[q]=.5*(e.vector[offset+a+1]+(prev<0?0:e.vector[offset+2*prev+1]));}
   intensity[q]=re[q]**2+im[q]**2;maps.Itotal[q]+=intensity[q];
  }
  components[f+'_real']=re;components[f+'_imag']=im;maps[f]=intensity;fractions.push(intensity.reduce((a,b)=>a+b,0));
 }
 let total=0,i2=0,xsum=0,ysum=0,inside=0,pml=0,edge=0;
 const bound=c.window_halfwidth_um-(c.pml.enabled?c.pml.thickness_um:0),I=maps.Itotal;
 for(let j=0;j<n;j++)for(let i=0;i<n;i++){
  const q=j*n+i,x=-c.window_halfwidth_um+(i+.5)*h,y=-c.window_halfwidth_um+(j+.5)*h,a=I[q];
  total+=a;if(Math.max(Math.abs(x),Math.abs(y))>=bound){pml+=a;continue;}
  inside+=a;i2+=a*a;xsum+=x*a;ysum+=y*a;
  if(i<2||i>=n-2||j<2||j>=n-2)edge+=a;
 }
 let moment=0;const cx=xsum/inside,cy=ysum/inside;
 for(let j=0;j<n;j++)for(let i=0;i<n;i++){const x=-c.window_halfwidth_um+(i+.5)*h,y=-c.window_halfwidth_um+(j+.5)*h;if(Math.max(Math.abs(x),Math.abs(y))<bound)moment+=((x-cx)**2+(y-cy)**2)*I[j*n+i];}
 const pmlFraction=pml/total,accepted=pmlFraction<=.35&&inside/total>=.65&&neff[1]>=-1e-8&&e.residual<1e-7;
 const loss=lossDbKm(neff[1],c.wavelength_nm),noise=lossDbKm(e.residual*Math.hypot(...e.value)/(2*Math.max(neff[0],1e-20)),c.wavelength_nm);
 return {neff:neff[0],neff_imag:neff[1],neff_complex:{real:neff[0],imag:neff[1]},residual:e.residual,
  loss_dB_km:loss,loss_residual_scale_dB_km:noise,loss_below_residual_scale:Math.abs(loss)<noise,loss_dB_turn:c.bend_radius_um?loss*2*Math.PI*c.bend_radius_um*1e-9:null,
  loss_status:c.pml.enabled?'미검증: mesh/PML convergence required':neff[1]===0?'Lossless closed boundary':'Material absorption; no radiation boundary',
  pml_power_fraction:pmlFraction,physical_power_fraction:inside/total,edge_power_fraction:edge/total,
  accepted,rejection:accepted?null:'PML energy >35%, negative attenuation, or eigen residual failed',
  MFD_2nd_um:2*Math.sqrt(2*moment/inside),Aeff_um2:inside*inside/i2*h*h,
  Ex_intensity_fraction:fractions[0]/total,Ey_intensity_fraction:fractions[1]/total,Ez_intensity_fraction:fractions[2]/total,
  polarization:fractions[0]>fractions[1]?'Ex dominant':'Ey dominant',field_maps:maps,field_components:components};
}
export function solveYee(input,{onProgress=()=>{}}={}){
 const start=performance.now(),g=buildYeeGrid(input),base=yeeOperator(g),sparse=sparseYee(g),op={...base,...sparse,allocatedBytes:base.allocatedBytes+sparse.allocatedBytes},count=g.config.modes,
  block=Math.min(20,Math.max(2,count+2)),Q=[];
 for(let j=0;j<block;j++)Q.push(orth(seed(g,j%2,Math.floor(j/2)),Q));
 let eigen=[],iterations=0,linearIterations=0;
 const tolerance=input.eigen_tolerance??(g.config.pml.enabled||g.config.k_core||g.config.k_clad?1e-10:1e-8),maxOuter=input.max_outer??120;
 for(iterations=0;iterations<maxOuter;iterations++){
  const next=[];
  for(let j=0;j<block;j++){
   const solution=gmres(op,Q[j],{maxIterations:input.max_linear??1400,restart:input.gmres_restart??36,tolerance:Math.min(2e-10,tolerance*.02)});linearIterations+=solution.iterations;
   next.push(orth(solution.x,next));
  }
  Q.splice(0,Q.length,...next);
  if(iterations%3===0||iterations===maxOuter-1){
   eigen=ritz(op,Q);const residual=Math.max(...eigen.slice(0,count).map(e=>e.residual));
   onProgress({phase:'shift-invert',iteration:iterations+1,residual,percent:Math.min(96,Math.round(96*(iterations+1)/maxOuter)),linearIterations});
   if(eigen.slice(0,count).every(e=>e.residual<tolerance))break;
  }
 }
 const converged=eigen.filter(e=>e.residual<tolerance),raw=converged.slice(0,count).map(e=>fields(e,g,op));
 if(raw.length<count)throw Error('Eigenproblem did not converge all requested modes; best residuals: '+eigen.slice(0,count).map(e=>e.residual.toExponential(2)).join(', '));
 const modes=raw.filter(m=>m.accepted).map((m,i)=>({...m,mode_index:i+1})),rejected=raw.filter(m=>!m.accepted);
 if(!modes.length)throw Error('All eigenmodes rejected by PML/physical-domain filter');
 const arraysBytes=op.allocatedBytes+(g.epsX.byteLength*3+g.index.byteLength)+op.N*8*(block*3+44)+modes.length*g.K*8*10;
 return {config:g.config,grid_n:g.n,h_um:g.h,index_map:g.index,modes,rejected,
  method:ENGINE,iterations:iterations+1,linearIterations,wallMs:performance.now()-start,
  memory:{typedArrayWorkspaceModelBytes:arraysBytes,kind:'estimated live workspace; not measured heap',heapUsedBytes:globalThis.performance?.memory?.usedJSHeapSize??null},
  warnings:[...g.warnings,...(rejected.length?['PML/internal spurious modes rejected: '+rejected.length]:[])],
  convergence:{eigenTolerance:tolerance,linearTolerance:Math.min(2e-10,tolerance*.02),discretization:'Yee staggered transverse E; β²/k0²=P Q',algorithm:'block shift-invert subspace iteration; complex ILU(0)-GMRES; Rayleigh–Ritz'}};
}
export function portableYee(result){
 const {config:c,grid_n:n,h_um:h}=result;
 const matrix=a=>Array.from({length:n},(_,j)=>Array.from(a.subarray(j*n,(j+1)*n)));
 return {schema:'LSCNS_vector_mode_v1',schemaVersion:1,engineVersion:'1.0.0',method:result.method,
  input:{...c,delta_n:c.n_core-c.n_clad},modes:result.modes.map(m=>{
   const {field_maps,field_components,...meta}=m,max=Math.max(...field_maps.Itotal,1e-30);
   return {...meta,field_map:{grid_size:n,x_um:Array.from({length:n},(_,i)=>-c.window_halfwidth_um+(i+.5)*h),y_um:Array.from({length:n},(_,i)=>-c.window_halfwidth_um+(i+.5)*h),
    intensity_normalized_to_max_total:Object.fromEntries(Object.entries(field_maps).map(([k,a])=>[k,matrix(Float64Array.from(a,v=>v/max))])),
    index_map:matrix(Float64Array.from(result.index_map)),field_components:Object.fromEntries(Object.entries(field_components).map(([k,a])=>[k,matrix(a)])),
    field_component_phase_convention:'exp(i beta z - i omega t); complex E on Yee lattice, interpolated to cell nodes; arbitrary global phase'}};
  }),diagnostics:{wallMs:result.wallMs,memory:result.memory,convergence:result.convergence,rejected:result.rejected.map(({field_maps,field_components,...m})=>m),warnings:result.warnings},
  validation:{status:'미검증 until mesh/PML/reference comparisons pass',domain:c.pml.enabled?'complex-stretched PML':'closed finite square',lossConvention:'n+iκ; positive Im(neff) means power attenuation'}};
}
