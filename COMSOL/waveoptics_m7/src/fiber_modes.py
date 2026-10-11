"""M7 optical fiber 2D scalar P1 FEM, independent circular LP01 Bessel validator.
Units in micrometers. Weak-guidance modes only, no vector PML / confinement loss.
"""
import numpy as np
from dataclasses import dataclass
from scipy.sparse import coo_matrix
from scipy.sparse.linalg import eigsh
from scipy.special import j0,j1,k0 as K0,k1 as K1
from scipy.optimize import brentq
from scipy.integrate import simpson
C0=299792458.
SB=(.6961663,.4079426,.8974794);SC=(.0684043**2,.1162414**2,9.896161**2)
def silica_index(w):
 w=float(w)
 if not .21<w<3.7:raise ValueError("Sellmeier domain")
 return float(np.sqrt(1+sum(b*w*w/(w*w-c) for b,c in zip(SB,SC))))
@dataclass(frozen=True)
class Fiber:
 wavelength_um:float=1.55
 radius_um:float=4.1
 delta_n:float=.005
 domain_halfwidth_um:float=18.
 n_cladding:float|None=None
 @property
 def clad(self):return silica_index(self.wavelength_um) if self.n_cladding is None else self.n_cladding
 @property
 def core(self):return self.clad+self.delta_n
 @property
 def v(self):return 2*np.pi*self.radius_um/self.wavelength_um*np.sqrt(self.core**2-self.clad**2)
def lp01_analytic(f):
 v=f.v
 if not 0<v<2.404825557695772:raise ValueError("LP01 oracle limited to V<2.405")
 def eq(u):
  w=np.sqrt(v*v-u*u)
  return u*j1(u)/j0(u)-w*K1(w)/K0(w)
 u=brentq(eq,max(v*1e-6,1e-8),v*(1-1e-10),xtol=2e-13)
 w=np.sqrt(v*v-u*u)
 return dict(V=float(v),u=float(u),w=float(w),n_eff=float(np.sqrt((2*np.pi*f.core/f.wavelength_um)**2-(u/f.radius_um)**2)*f.wavelength_um/(2*np.pi)))
def radial_lp01_aeff(f):
 a=lp01_analytic(f);r=np.linspace(0,max(f.domain_halfwidth_um*2,12*f.radius_um),12001)
 shape=np.where(r<=f.radius_um,j0(a['u']*r/f.radius_um)/j0(a['u']),K0(a['w']*np.maximum(r,1e-9)/f.radius_um)/K0(a['w']))
 I=shape*shape;total=2*np.pi*simpson(I*r,x=r);four=2*np.pi*simpson(I*I*r,x=r)
 mom=2*np.pi*simpson(I*r**3,x=r)/total
 return dict(Aeff_um2=float(total**2/four),MFD_2nd_um=float(2*np.sqrt(2*mom)))
def grid(n,h):
 if n<15 or n%2!=1:raise ValueError("odd mesh count>=15")
 xy=np.stack(np.meshgrid(np.linspace(-h,h,n),np.linspace(-h,h,n),indexing='xy'),-1).reshape(-1,2)
 i,j=np.meshgrid(np.arange(n-1),np.arange(n-1),indexing='xy');p=(j*n+i).ravel()
 tris=np.concatenate([np.stack([p,p+1,p+n+1],1),np.stack([p,p+n+1,p+n],1)]).astype(int)
 free=np.flatnonzero((np.abs(xy[:,0])<h-1e-9)&(np.abs(xy[:,1])<h-1e-9))
 return xy,tris,free
def assemble_fem(xy,tri,index):
 p=xy[tri];dx=p[:,1,0]-p[:,0,0];dy=p[:,1,1]-p[:,0,1];ex=p[:,2,0]-p[:,0,0];ey=p[:,2,1]-p[:,0,1]
 det=dx*ey-dy*ex
 if np.any(np.abs(det)<1e-14):raise ValueError("degenerate triangle")
 area=np.abs(det)/2
 gx=np.stack([p[:,1,1]-p[:,2,1],p[:,2,1]-p[:,0,1],p[:,0,1]-p[:,1,1]],1)/det[:,None]
 gy=np.stack([p[:,2,0]-p[:,1,0],p[:,0,0]-p[:,2,0],p[:,1,0]-p[:,0,0]],1)/det[:,None]
 K=area[:,None,None]*(gx[:,:,None]*gx[:,None,:]+gy[:,:,None]*gy[:,None,:])
 b=np.ones((3,3));np.fill_diagonal(b,2)
 M=area[:,None,None]*b[None,:,:]/12
 ni=np.asarray(index(p.mean(1)),float)
 if len(ni)!=len(area) or np.any(ni<=0):raise ValueError("invalid index")
 N=M*ni[:,None,None]**2
 rows=np.broadcast_to(tri[:,:,None],(len(tri),3,3)).ravel()
 cols=np.broadcast_to(tri[:,None,:],(len(tri),3,3)).ravel()
 def csr(v):return coo_matrix((v.ravel(),(rows,cols)),shape=(len(xy),len(xy))).tocsr()
 return csr(K),csr(M),csr(N)
def solve_modes(f,n=81,cores=None,modes=2):
 if f.radius_um<=0 or f.delta_n<=0 or f.domain_halfwidth_um<=f.radius_um*2:raise ValueError("invalid fiber")
 centers=[(0.,0.)] if cores is None else list(cores)
 xy,tri,free=grid(n,f.domain_halfwidth_um)
 def index(points):
  r2=np.full(len(points),np.inf)
  for xx,yy in centers:r2=np.minimum(r2,(points[:,0]-xx)**2+(points[:,1]-yy)**2)
  return np.where(r2<=f.radius_um**2,f.core,f.clad)
 K,M,N=assemble_fem(xy,tri,index);k0=2*np.pi/f.wavelength_um;A=(k0*k0*N-K)[free][:,free];MM=M[free][:,free]
 vals,vectors=eigsh(A,k=modes,M=MM,sigma=(k0*f.core)**2+.02,which='LM',tol=1e-9)
 order=np.argsort(vals)[::-1];vals=vals[order];vectors=vectors[:,order]
 fields=np.zeros((len(xy),modes));fields[free,:]=vectors
 neff=np.sqrt(np.maximum(0,vals))/k0
 resid=[float(np.linalg.norm(A@vectors[:,i]-vals[i]*(MM@vectors[:,i]))/np.linalg.norm(vals[i]*(MM@vectors[:,i]))) for i in range(modes)]
 return dict(fiber=f,n=n,xy=xy,triangles=tri,free=free,mass=M,neff=neff,fields=fields,guided=(neff>f.clad)&(neff<f.core),residuals=resid,cores=centers)
def scalar_metrics(sol,k=0):
 p=sol['fields'][:,k];xy=sol['xy'];w=np.asarray(sol['mass'].sum(axis=1)).ravel();I=p*p
 area=np.sum(w*I);mom=np.sum(w*I*(xy[:,0]**2+xy[:,1]**2))/area
 return dict(n_eff=float(sol['neff'][k]),Aeff_um2=float(area**2/np.sum(w*I**2)),MFD_2nd_um=float(2*np.sqrt(2*mom)),residual=float(sol['residuals'][k]))
def dispersion_lp01(w=1.55,step=.003):
 def ne(x):return lp01_analytic(Fiber(wavelength_um=float(x)))['n_eff']
 h=step;d2=(-ne(w+2*h)+16*ne(w+h)-30*ne(w)+16*ne(w-h)-ne(w-2*h))/(12*h*h)
 return float(-(w/C0)*1e12*d2)
