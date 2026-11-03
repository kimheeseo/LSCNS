"""M7 2D P1 FEM cross-check. Solver routines transcribed from repo
COMSOL/waveoptics_m7/src/fiber_modes.py (checked 2026-10-19).
"""
import importlib.util,json
from pathlib import Path
import numpy as np
from scipy.sparse import coo_matrix
from scipy.sparse.linalg import eigsh
loc=Path(__file__).resolve().parent
spec=importlib.util.spec_from_file_location('g654',loc/'g654e_validation.py')
g=importlib.util.module_from_spec(spec);import sys;sys.modules[spec.name]=g;spec.loader.exec_module(g)
def grid(n,h):
 if not isinstance(n,(int,np.integer)) or isinstance(n,bool) or n<15 or n%2!=1:raise ValueError('odd mesh count>=15')
 if not np.isfinite(h) or h<=0:raise ValueError('positive finite halfwidth')
 xy=np.stack(np.meshgrid(np.linspace(-h,h,n),np.linspace(-h,h,n),indexing='xy'),-1).reshape(-1,2)
 i,j=np.meshgrid(np.arange(n-1),np.arange(n-1),indexing='xy');p=(j*n+i).ravel()
 tris=np.concatenate([np.stack([p,p+1,p+n+1],1),np.stack([p,p+n+1,p+n],1)]).astype(int)
 free=np.flatnonzero((np.abs(xy[:,0])<h-1e-9)&(np.abs(xy[:,1])<h-1e-9))
 return xy,tris,free
def assemble_fem(xy,tri,index):
 p=xy[tri];dx=p[:,1,0]-p[:,0,0];dy=p[:,1,1]-p[:,0,1];ex=p[:,2,0]-p[:,0,0];ey=p[:,2,1]-p[:,0,1]
 det=dx*ey-dy*ex
 if np.any(np.abs(det)<1e-14):raise ValueError('degenerate triangle')
 area=np.abs(det)/2
 gx=np.stack([p[:,1,1]-p[:,2,1],p[:,2,1]-p[:,0,1],p[:,0,1]-p[:,1,1]],1)/det[:,None]
 gy=np.stack([p[:,2,0]-p[:,1,0],p[:,0,0]-p[:,2,0],p[:,1,0]-p[:,0,0]],1)/det[:,None]
 K=area[:,None,None]*(gx[:,:,None]*gx[:,None,:]+gy[:,:,None]*gy[:,None,:])
 b=np.ones((3,3));np.fill_diagonal(b,2)
 M=area[:,None,None]*b[None,:,:]/12
 ni=np.asarray(index(p.mean(1)),float)
 if len(ni)!=len(area) or np.any(ni<=0):raise ValueError('invalid index')
 N=M*ni[:,None,None]**2
 rows=np.broadcast_to(tri[:,:,None],(len(tri),3,3)).ravel()
 cols=np.broadcast_to(tri[:,None,:],(len(tri),3,3)).ravel()
 def csr(v):return coo_matrix((v.ravel(),(rows,cols)),shape=(len(xy),len(xy))).tocsr()
 return csr(K),csr(M),csr(N)
def solve_modes(f,n=81,cores=None,modes=2):
 centers=[(0.,0.)] if cores is None else list(cores)
 c=np.asarray(centers,float)
 xy,tri,free=grid(n,f.domain_halfwidth_um)
 def index(points):
  r2=np.full(len(points),np.inf)
  for xx,yy in centers:r2=np.minimum(r2,(points[:,0]-xx)**2+(points[:,1]-yy)**2)
  return np.where(r2<=f.radius_um**2,f.core,f.clad)
 K,M,N=assemble_fem(xy,tri,index);k0=2*np.pi/f.wavelength_um;A=(k0*k0*N-K)[free][:,free];MM=M[free][:,free]
 vals,vectors=eigsh(A,k=modes,M=MM,sigma=(k0*f.core)**2+.02,which='LM',tol=1e-10,v0=np.random.default_rng(0).normal(size=len(free)))
 order=np.argsort(vals)[::-1];vals=vals[order];vectors=vectors[:,order]
 fields=np.zeros((len(xy),modes));fields[free,:]=vectors
 neff=np.sqrt(np.maximum(0,vals))/k0
 resid=[float(np.linalg.norm(A@vectors[:,i]-vals[i]*(MM@vectors[:,i]))/np.linalg.norm(vals[i]*(MM@vectors[:,i]))) for i in range(modes)]
 return dict(fiber=f,n=n,xy=xy,triangles=tri,free=free,mass=M,neff=neff,fields=fields,guided=(neff>f.clad)&(neff<f.core),residuals=resid,cores=centers)
def scalar_metrics(sol,k=0):
 b=np.array([[1/3,1/3,1/3],[.059715871789770,.470142064105115,.470142064105115],[.470142064105115,.059715871789770,.470142064105115],[.470142064105115,.470142064105115,.059715871789770],[.797426985353087,.101286507323456,.101286507323456],[.101286507323456,.797426985353087,.101286507323456],[.101286507323456,.101286507323456,.797426985353087]])
 weights=np.array([.225,.132394152788506,.132394152788506,.132394152788506,.125939180544827,.125939180544827,.125939180544827])
 tri=sol['triangles'];xyz=sol['xy'][tri];p=sol['fields'][tri,k]
 area=np.abs((xyz[:,1,0]-xyz[:,0,0])*(xyz[:,2,1]-xyz[:,0,1])-(xyz[:,2,0]-xyz[:,0,0])*(xyz[:,1,1]-xyz[:,0,1]))/2
 intensity=np.abs(p@b.T)**2;points=np.einsum('qi,tij->tqj',b,xyz)
 wt=area[:,None]*weights
 power=float(np.sum(wt*intensity));four=float(np.sum(wt*intensity**2))
 center=np.sum(wt[:,:,None]*intensity[:,:,None]*points,axis=(0,1))/power
 r2=np.sum((points-center)**2,axis=2);moment=np.sum(wt*intensity*r2)/power
 return dict(n_eff=float(sol['neff'][k]),Aeff_um2=power**2/four,MFD_2nd_um=float(2*np.sqrt(2*moment)),residual=float(sol['residuals'][k]),centroid_um=center.tolist(),power_integral=power)
if __name__=='__main__':
 data=json.loads((loc/'g654e_results.json').read_text());params=data['parameters']
 f=g.Fiber(radius_um=params['radius_um'],delta_n=params['delta_n'],domain_halfwidth_um=30.)
 ref=g.lp01_analytic(f);ora=g.radial_lp01_aeff(f)
 tests=[]
 for n,h in [(61,24),(81,24),(121,24),(161,24),(121,32)]:
  f=g.replace(f,domain_halfwidth_um=h)
  sol=solve_modes(f,n=n,modes=1);m=scalar_metrics(sol)
  d=dict(n=n,outer_halfwidth_um=h,triangles=len(sol['triangles']),**m)
  d['neff_error']=abs(m['n_eff']-ref['n_eff'])
  d['Aeff_error_pct']=100*(m['Aeff_um2']/ora['Aeff_um2']-1)
  d['MFD_second_error_pct']=100*(m['MFD_2nd_um']/ora['MFD_2nd_um']-1)
  print(json.dumps({k:v for k,v in d.items() if k not in ('centroid_um','power_integral')}),flush=True)
  tests.append(d)
 data['m7_2d_fem']=tests
 data['checks']['2d_neff_error_fine_lt_1e_5']=tests[3]['neff_error']<1e-5
 data['checks']['2d_aeff_error_fine_lt_2pct']=abs(tests[3]['Aeff_error_pct'])<2
 data['checks']['2d_mfd2_error_fine_lt_2pct']=abs(tests[3]['MFD_second_error_pct'])<2
 data['checks']['2d_fine_residual_lt_1e_8']=tests[3]['residual']<1e-8
 (loc/'g654e_results.json').write_text(json.dumps(data,ensure_ascii=False,indent=2))