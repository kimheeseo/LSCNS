/** Local sparse Yee PQ stencil + complex ILU(0). No dense global matrix. */
const mul=(a,b)=>[a[0]*b[0]-a[1]*b[1],a[0]*b[1]+a[1]*b[0]];
const div=(a,b)=>{const d=b[0]**2+b[1]**2;return [(a[0]*b[0]+a[1]*b[1])/d,(a[1]*b[0]-a[0]*b[1])/d];};
export function sparseYee(g){
 const {n,K,h,epsX,epsY,epsZ,config:c}=g,k0=2*Math.PI/(c.wavelength_nm/1000),N=2*K,sigma=g.target**2;
 const stretch=[[],[]];
 for(let dir=0;dir<2;dir++)for(let i=0;i<n;i++){
  const x=-c.window_halfwidth_um+(i+.5+(dir?.5:0))*h,t=c.pml.enabled?Math.max(0,(Math.abs(x)-c.window_halfwidth_um+c.pml.thickness_um)/c.pml.thickness_um):0,a=c.pml.enabled?c.pml.strength*Math.min(t,1)**c.pml.order:0;
  stretch[dir].push([1/((1+a*a)*h*k0),-a/((1+a*a)*h*k0)]);
 }
 const invZ=Array.from({length:K},(_,q)=>div([1,0],[epsZ[2*q],epsZ[2*q+1]]));
 function D(q,axis,f){const p=axis===0?q%n:Math.floor(q/n),step=axis===0?1:n,s=stretch[f?1:0][p],out=[[q,f?[-s[0],-s[1]]:s]];
  if(f?p<n-1:p>0)out.push([q+(f?step:-step),f?s:[-s[0],-s[1]]]);return out;}
 const rowptr=new Int32Array(N+1),cols=[],data=[];
 for(let row=0;row<N;row++){
  const q=row%K,field=Math.floor(row/K),map=new Map();
  function add(col,v){const old=map.get(col)||[0,0];old[0]+=v[0];old[1]+=v[1];map.set(col,old);}
  function product(target,a1,f1,a2,f2,middle,last,sign=1){
   for(const [mid,v1]of D(q,a1,f1))for(const [col,v2]of D(mid,a2,f2)){
    let v=mul(v1,v2);if(middle)v=mul(v,middle[mid]);if(last)v=mul(v,[last[2*col],last[2*col+1]]);add(target*K+col,[sign*v[0],sign*v[1]]);
   }
  }
  if(field===0){
   product(0,0,true,0,false,invZ,epsX);product(0,1,false,1,true);
   product(1,0,true,1,false,invZ,epsY);product(1,1,false,0,true,null,null,-1);
  }else{
   product(0,1,true,0,false,invZ,epsX);product(0,0,false,1,true,null,null,-1);
   product(1,1,true,1,false,invZ,epsY);product(1,0,false,0,true);
  }
  const eps=field===0?epsX:epsY;add(row,[eps[2*q],eps[2*q+1]]);
  rowptr[row]=cols.length;
  for(const [col,v]of [...map].sort((a,b)=>a[0]-b[0]))if(col===row||Math.hypot(...v)>1e-16){cols.push(col);data.push((col===row?sigma:0)-v[0],-v[1]);}
 }
 rowptr[N]=cols.length;const column=Int32Array.from(cols),values=Float64Array.from(data),lu=Float64Array.from(data),diagonal=new Int32Array(N);
 for(let i=0;i<N;i++){for(let p=rowptr[i];p<rowptr[i+1];p++)if(column[p]===i)diagonal[i]=p;}
 const find=(row,col)=>{let lo=rowptr[row],hi=rowptr[row+1]-1;while(lo<=hi){const mid=(lo+hi)>>1;if(column[mid]===col)return mid;if(column[mid]<col)lo=mid+1;else hi=mid-1;}return -1;};
 let regularized=0;
 for(let i=0;i<N;i++){
  for(let p=rowptr[i];p<diagonal[i];p++){
   const j=column[p],dp=2*diagonal[j],a=2*p,v=div([lu[a],lu[a+1]],[lu[dp],lu[dp+1]]);lu[a]=v[0];lu[a+1]=v[1];
   for(let r=diagonal[j]+1;r<rowptr[j+1];r++){
    const s=find(i,column[r]);if(s<0)continue;const t=mul(v,[lu[2*r],lu[2*r+1]]);lu[2*s]-=t[0];lu[2*s+1]-=t[1];
   }
  }
  const d=2*diagonal[i];if(Math.hypot(lu[d],lu[d+1])<1e-10){lu[d]+=1e-8;regularized++;}
 }
 function shifted(x,out){
  for(let i=0;i<N;i++){let r=0,z=0;for(let p=rowptr[i];p<rowptr[i+1];p++){const j=2*column[p],a=values[2*p],b=values[2*p+1];r+=a*x[j]-b*x[j+1];z+=a*x[j+1]+b*x[j];}out[2*i]=r;out[2*i+1]=z;}return out;
 }
 function apply(x,out){shifted(x,out);for(let i=0;i<out.length;i++)out[i]=sigma*x[i]-out[i];return out;}
 function precondition(x,out){
  out.set(x);
  for(let i=0;i<N;i++)for(let p=rowptr[i];p<diagonal[i];p++){
   const j=2*column[p],a=lu[2*p],b=lu[2*p+1];out[2*i]-=a*out[j]-b*out[j+1];out[2*i+1]-=a*out[j+1]+b*out[j];
  }
  for(let i=N-1;i>=0;i--){
   for(let p=diagonal[i]+1;p<rowptr[i+1];p++){
    const j=2*column[p],a=lu[2*p],b=lu[2*p+1];out[2*i]-=a*out[j]-b*out[j+1];out[2*i+1]-=a*out[j+1]+b*out[j];
   }
   const d=2*diagonal[i],a=lu[d],b=lu[d+1],den=a*a+b*b,r=out[2*i],z=out[2*i+1];out[2*i]=(r*a+z*b)/den;out[2*i+1]=(z*a-r*b)/den;
  }
  return out;
 }
 return {apply,shifted,precondition,regularized,allocatedBytes:rowptr.byteLength+column.byteLength+values.byteLength+lu.byteLength+diagonal.byteLength,nonzeros:cols.length};
}
