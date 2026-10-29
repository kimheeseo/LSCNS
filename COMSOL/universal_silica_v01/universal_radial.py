"""Research prototype v0.1: rotationally symmetric solid-silica fibers.

Scalar weak-guidance radial P1 finite element mode solver for LP01.
Materials are silica baseline Sellmeier plus *constant* delta_n offsets;
this is an illustrative index model, not a verified Ge/F doped Sellmeier law.
This solver does NOT model HCF/NANF, noncircular MCF, birefringence, loss, bends,
or fiber manufacturing variation. All transverse dimensions are in micrometres.
"""
from __future__ import annotations
from dataclasses import dataclass, replace, asdict
from typing import Sequence
import json
import numpy as np
from scipy.special import j0,j1,k0 as K0,k1 as K1
from scipy.optimize import brentq
from scipy.integrate import quad
from scipy.sparse import coo_matrix
from scipy.sparse.linalg import eigsh
from numpy.polynomial.legendre import leggauss

C0=299792458.0
_SELL_B=(0.6961663,0.4079426,0.8974794)
_SELL_C=(0.0684043**2,0.1162414**2,9.896161**2)

@dataclass(frozen=True)
class Model:
    profile: str = 'step'                  # step, graded, trench, multilayer
    wavelength_nm: float = 1550.0
    core_radius_um: float = 5.5
    delta_n_core: float = 0.0034785793846155
    outer_radius_um: float = 40.0       # computational, not physical cladding radius
    graded_alpha: float = 2.0
    trench_start_um: float = 7.0
    trench_end_um: float = 10.0
    trench_delta_n: float = -0.001
    # custom shells are (outer_radius_um, absolute delta_n relative to silica baseline).
    # First one starts at r=0; last listed shell may end before outer_radius_um.
    shells: tuple[tuple[float,float],...] = ((5.0,0.004), (7.0,0.0), (9.0,-0.0007))
    n2_m2_per_W: float = 2.2e-20
    def validate(self):
        if self.profile not in ('step','graded','trench','multilayer'):
            raise ValueError('profile must be step, graded, trench, or multilayer')
        for name in ('wavelength_nm','core_radius_um','outer_radius_um','graded_alpha','n2_m2_per_W'):
            x=float(getattr(self,name))
            if not np.isfinite(x) or x<=0: raise ValueError(f'{name} must be positive/finite')
        if not 250 < self.wavelength_nm < 3600: raise ValueError('wavelength outside silica Sellmeier domain')
        if self.outer_radius_um<=2.2*self.core_radius_um: raise ValueError('outer radius must exceed 2.2 x core radius')
        for name in ('delta_n_core','trench_delta_n'):
            if not np.isfinite(getattr(self,name)): raise ValueError(name)
        if self.profile=='trench' and not (self.core_radius_um < self.trench_start_um < self.trench_end_um < self.outer_radius_um):
            raise ValueError('trench radii must be core < start < end < outer')
        if self.profile=='multilayer':
            if not self.shells: raise ValueError('shells required')
            rr=np.array([x[0] for x in self.shells]); nn=np.array([x[1] for x in self.shells]);
            if np.any(~np.isfinite(rr)) or np.any(~np.isfinite(nn)) or np.any(np.diff(rr)<=0) or rr[0]<=0 or rr[-1]>=self.outer_radius_um:
                raise ValueError('shell boundaries must increase and lie below outer radius')
        return self

def silica_index(wavelength_um):
    w=np.asarray(wavelength_um,dtype=float)
    if np.any((w<=0.21)|(w>=3.7)): raise ValueError('silica Sellmeier only 0.21–3.7 µm')
    return np.sqrt(1+sum(b*w*w/(w*w-c) for b,c in zip(_SELL_B,_SELL_C)))

def boundaries(m:Model):
    m.validate()
    if m.profile=='trench': marks=[m.core_radius_um,m.trench_start_um,m.trench_end_um]
    elif m.profile=='multilayer': marks=[r for r,d in m.shells]
    else: marks=[m.core_radius_um]
    return np.r_[0.0,marks,m.outer_radius_um]

def delta_n_r(m:Model,r):
    r=np.asarray(r,dtype=float)
    if m.profile=='step': return np.where(r<m.core_radius_um,m.delta_n_core,0.0)
    if m.profile=='graded': return np.where(r<m.core_radius_um,m.delta_n_core*np.maximum(1-(r/m.core_radius_um)**m.graded_alpha,0),0.0)
    if m.profile=='trench':
        return np.where(r<m.core_radius_um,m.delta_n_core,
                        np.where((r>=m.trench_start_um)&(r<m.trench_end_um),m.trench_delta_n,0.0))
    out=np.zeros_like(r)
    left=0.0
    for right,dn in m.shells:
        out=np.where((r>=left)&(r<right),dn,out)
        left=right
    return out

def index_r(m:Model,r):
    return silica_index(m.wavelength_nm/1000)+delta_n_r(m,r)

def radial_grid(m:Model,elements=750):
    if not isinstance(elements,(int,np.integer)) or elements<80:raise ValueError('elements>=80')
    br=boundaries(m); intervals=np.diff(br)
    base=np.maximum(12,np.round(elements*intervals/np.sum(intervals)).astype(int))
    # allocate more near interfaces by maintaining at least 12 nodes per ring
    out=[np.linspace(br[i],br[i+1],int(base[i])+1)[:-1] for i in range(len(base))]
    return np.r_[np.concatenate(out),br[-1]]

def solve_lp01(m:Model,elements=750):
    m.validate(); r=radial_grid(m,elements)
    nodes=len(r); n_el=nodes-1
    t0,ww=leggauss(5); t=(t0+1)/2; wq=ww/2
    h=np.diff(r); rq=r[:-1,None]+h[:,None]*t
    weights=h[:,None]*wq
    basis=np.stack((1-t,t),axis=1)
    # quadratic-form angular integration has 2pi on both sides, cancels for eigenvalue.
    mass=np.einsum('eq,qi,qj->eij',weights*rq,basis,basis)
    deriv=(np.diff(r**2)/(2*h*h))[:,None,None]*np.array([[1.,-1.],[-1.,1.]])
    nvals=index_r(m,rq)
    n_mass=np.einsum('eq,eq,qi,qj->eij',weights*rq,nvals**2,basis,basis)
    k0=2*np.pi/(m.wavelength_nm/1000)
    mats=(k0*k0*n_mass-deriv,mass)
    edge=np.stack((np.arange(n_el),np.arange(1,nodes)),axis=1)
    ii=np.broadcast_to(edge[:,:,None],(n_el,2,2)).ravel()
    jj=np.broadcast_to(edge[:,None,:],(n_el,2,2)).ravel()
    A,B=[coo_matrix((v.ravel(),(ii,jj)),shape=(nodes,nodes)).tocsr()[:-1,:-1] for v in mats]
    n_max=float(np.max(index_r(m,rq)))
    eigen,ev=eigsh(A,k=1,M=B,sigma=(k0*n_max)**2+0.05,which='LM',tol=1e-10,v0=np.ones(nodes-1))
    beta2=float(eigen[0]); psi=np.r_[ev[:,0],0.0]
    if psi[0]<0:psi=-psi
    if beta2<=0: raise RuntimeError('nonpositive beta²')
    neff=float(np.sqrt(beta2)/k0)
    rel_res=float(np.linalg.norm(A@ev[:,0]-beta2*(B@ev[:,0]))/np.linalg.norm(beta2*(B@ev[:,0])))
    I=(psi[edge]@basis.T)**2
    wtr=weights*rq
    P=float(2*np.pi*np.sum(wtr*I)); Q=float(2*np.pi*np.sum(wtr*I**2));
    radial_moment=float(2*np.pi*np.sum(wtr*I*rq**2)/P)
    slope=np.diff(psi)/h
    nrm_grad=float(np.sum(weights*rq*slope[:,None]**2))
    nrm_psi=float(np.sum(weights*rq*I))
    mfd_pii=float(2*np.sqrt(2*nrm_psi/nrm_grad))
    aeff=P**2/Q
    tail_power=float(2*np.pi*np.sum(wtr[rq[:,0]>0.8*m.outer_radius_um]*I[rq[:,0]>0.8*m.outer_radius_um])/P)
    return {'model':m,'neff':neff,'Aeff_um2':aeff,'MFD_2nd_um':2*np.sqrt(2*radial_moment),'MFD_PII_um':mfd_pii,
            'gamma_Wkm':2*np.pi/(m.wavelength_nm*1e-9)*m.n2_m2_per_W/(aeff*1e-12)*1000,
            'residual':rel_res,'tail_power':tail_power,'guided_proxy':bool(neff>float(silica_index(m.wavelength_nm/1000))),
            'r_um':r,'field':psi/np.sqrt(P),'elements':n_el}

def bessel_step_oracle(m:Model):
    if m.profile!='step':raise ValueError('Bessel oracle only valid for step-index')
    k0=2*np.pi/(m.wavelength_nm/1000)
    ncl=float(silica_index(m.wavelength_nm/1000)); nco=ncl+m.delta_n_core
    if nco<=ncl:raise ValueError('index-guided core required')
    V=k0*m.core_radius_um*np.sqrt(nco*nco-ncl*ncl)
    if not (0<V<2.40482555769):raise ValueError('LP01 reference restricted to V<2.405')
    def fn(u):
        w=np.sqrt(V*V-u*u)
        return u*j1(u)/j0(u)-w*K1(w)/K0(w)
    u=brentq(fn,max(1e-8,V*1e-6),V*(1-1e-10),xtol=3e-13)
    neff=float(np.sqrt(nco*nco-(u/(k0*m.core_radius_um))**2))
    return {'neff':neff,'V':V}

def spectral_neff(m:Model,lambda_nm,elements=750):
    return solve_lp01(replace(m,wavelength_nm=float(lambda_nm)),elements)['neff']

def chromatic_dispersion(m:Model,elements=750,h_nm=10.0):
    # D ps/(nm.km) = -(lambda/c) d2 n_eff/d(lambda_um)2 * 1e12
    # All eigenvalue samples computed on identical radial mesh (wavelength-independent geometry).
    w=m.wavelength_nm/1000; h=h_nm/1000
    n=np.array([spectral_neff(m,(w+j*h)*1000,elements) for j in (-2,-1,0,1,2)])
    d2=(-n[4]+16*n[3]-30*n[2]+16*n[1]-n[0])/(12*h*h)
    return float(-w/C0*1e12*d2)

def serialize(sol):
    return {k:(asdict(v) if isinstance(v,Model) else float(v) if isinstance(v,(np.floating,np.integer)) else v)
            for k,v in sol.items() if k not in ('r_um','field')}

PRESETS={
 'G.654.E-like (illustrative)':dict(profile='step',core_radius_um=5.5,delta_n_core=0.0034785793846155),
 'G.652.D-like (illustrative)':dict(profile='step',core_radius_um=4.1,delta_n_core=0.005),
 'Graded core (illustrative)':dict(profile='graded',core_radius_um=5.0,delta_n_core=0.006,graded_alpha=2.0),
 'Trench-assisted (illustrative)':dict(profile='trench',core_radius_um=4.1,delta_n_core=0.005,trench_start_um=6.0,trench_end_um=9.0,trench_delta_n=-0.001),
 'Multilayer (illustrative)':dict(profile='multilayer',outer_radius_um=40.0,shells=((4.0,0.006),(5.0,0.002),(7.0,0.0),(9.0,-0.0008)))
}
if __name__=='__main__':
    for name,vals in PRESETS.items():
        model=Model(**vals); out=solve_lp01(model,elements=500)
        print(name,{k:round(out[k],8) for k in ('neff','Aeff_um2','MFD_2nd_um','MFD_PII_um','residual','tail_power')})
