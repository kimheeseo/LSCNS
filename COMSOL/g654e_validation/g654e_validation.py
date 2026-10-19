"""Independent repeatable verification of the M7 scalar Step-index assumptions.
Numerical radial P1 FEM is directly transcribed from COMSOL/waveoptics_m7/src/radial_fem.py
(with equivalent input class). LP01 Bessel oracle is transcribed from fiber_modes.py.
This DOES NOT simulate or validate the actual undisclosed LE125 index profile.
"""
from __future__ import annotations
from dataclasses import dataclass, replace
import numpy as np
from scipy.special import j0,j1,k0 as K0,k1 as K1
from scipy.optimize import brentq, root, minimize_scalar
from scipy.integrate import quad
from numpy.polynomial.legendre import leggauss
from scipy.sparse import coo_matrix
from scipy.sparse.linalg import eigsh
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import json
from pathlib import Path
C0=299792458.
SB=(.6961663,.4079426,.8974794);SC=(.0684043**2,.1162414**2,9.896161**2)
def silica_index(w):
 w=float(w)
 return float(np.sqrt(1+sum(b*w*w/(w*w-c) for b,c in zip(SB,SC))))
@dataclass(frozen=True)
class Fiber:
 wavelength_um:float=1.55
 radius_um:float=5.0
 delta_n:float=.003
 domain_halfwidth_um:float=35.
 n_cladding:float|None=None
 @property
 def clad(self):return silica_index(self.wavelength_um) if self.n_cladding is None else self.n_cladding
 @property
 def core(self):return self.clad+self.delta_n
 @property
 def v(self):return 2*np.pi*self.radius_um/self.wavelength_um*np.sqrt(self.core**2-self.clad**2)
def lp01_analytic(f):
 v=f.v
 if not 0<v<2.404825557695772:raise ValueError(f"LP01 oracle limited V<2.405: V={v}")
 def eq(u):
  w=np.sqrt(v*v-u*u)
  return u*j1(u)/j0(u)-w*K1(w)/K0(w)
 u=brentq(eq,max(v*1e-6,1e-8),v*(1-1e-10),xtol=2e-13)
 w=np.sqrt(v*v-u*u)
 ne=float(np.sqrt((2*np.pi*f.core/f.wavelength_um)**2-(u/f.radius_um)**2)*f.wavelength_um/(2*np.pi))
 return dict(V=float(v),u=float(u),w=float(w),n_eff=ne)
def radial_lp01_aeff(f):
 a=lp01_analytic(f);u,w=a['u'],a['w'];radius=f.radius_um
 def shape(t):return j0(u*t)/j0(u) if t<=1 else K0(w*t)/K0(w)
 def d_shape(t):return -u*j1(u*t)/j0(u) if t<=1 else -w*K1(w*t)/K0(w)
 def integrate(power,moment,deriv=False):
  s=d_shape if deriv else shape
  fun=lambda t:s(t)**power*t**moment
  return quad(fun,0,1,epsabs=1e-11,epsrel=1e-10)[0]+quad(fun,1,np.inf,epsabs=1e-11,epsrel=1e-10)[0]
 total=2*np.pi*radius**2*integrate(2,1)
 four=2*np.pi*radius**2*integrate(4,1)
 mom=2*np.pi*radius**4*integrate(2,3)/total
 # Petermann-II: MFD=2 sqrt(2 * \int E^2 r dr / \int (dE/dr)^2 r dr).
 # With t=r/a, denominator has no a^2 and numerator has a^2.
 petermann=2*radius*np.sqrt(2*integrate(2,1)/integrate(2,1,deriv=True))
 return dict(Aeff_um2=float(total**2/four),MFD_2nd_um=float(2*np.sqrt(2*mom)), MFD_petermann_ii_um=float(petermann))
def solve_radial(fiber, core_elements=160, clad_elements=640, outer_radius_um=None):
 outer = fiber.domain_halfwidth_um if outer_radius_um is None else float(outer_radius_um)
 a=fiber.radius_um
 r=np.r_[np.linspace(0,a,core_elements+1),np.linspace(a,outer,clad_elements+1)[1:]]
 x,w=leggauss(4);t=(x+1)/2;phi=np.stack((1-t,t),axis=1)
 h=np.diff(r);rq=r[:-1,None]+h[:,None]*t;weights=h[:,None]*w/2
 mass=np.einsum('eq,qi,qj->eij',weights*rq,phi,phi)
 stiff=((r[1:]**2-r[:-1]**2)/(2*h**2))[:,None,None]*np.array([[1.,-1.],[-1.,1.]])
 indices=np.r_[np.full(core_elements,fiber.core),np.full(clad_elements,fiber.clad)]
 k0=2*np.pi/fiber.wavelength_um
 local=k0**2*indices[:,None,None]**2*mass-stiff
 edges=np.stack((np.arange(len(h)),np.arange(len(h))+1),axis=1)
 rows=np.broadcast_to(edges[:,:,None],local.shape).ravel()
 cols=np.broadcast_to(edges[:,None,:],local.shape).ravel()
 def matrix(v):return coo_matrix((v.ravel(),(rows,cols)),shape=(len(r),len(r))).tocsr()[:-1,:-1]
 A,M=matrix(local),matrix(mass)
 val,vec=eigsh(A,k=1,M=M,sigma=(k0*fiber.core)**2,which='LM',v0=np.ones(len(r)-1),tol=1e-11)
 beta2=float(val[0]);q=vec[:,0]
 residual=np.linalg.norm(A@q-beta2*(M@q))/np.linalg.norm(beta2*(M@q))
 field=np.r_[q,0.]
 if field[0]<0:field=-field
 fieldq=field[edges]@phi.T;intensity=fieldq**2
 wt=2*np.pi*weights*rq;power=np.sum(wt*intensity)
 aeff=power**2/np.sum(wt*intensity**2)
 moment=np.sum(wt*intensity*rq**2)/power
 neff=np.sqrt(beta2)/k0 if beta2>0 else float('nan')
 return dict(n_eff=float(neff),Aeff_um2=float(aeff),MFD_2nd_um=float(2*np.sqrt(2*moment)),residual=float(residual),guided=bool(fiber.clad<neff<fiber.core),outer_radius_um=outer,core_elements=core_elements,clad_elements=clad_elements, radius_um=r,field=field/np.sqrt(power))
def dispersion_lp01(f,step=.003):
 def ne(x):return lp01_analytic(replace(f,wavelength_um=float(x)))['n_eff']
 w=f.wavelength_um;h=step
 d2=(-ne(w+2*h)+16*ne(w+h)-30*ne(w)+16*ne(w-h)-ne(w-2*h))/(12*h*h)
 return float(-(w/C0)*1e12*d2)
def run():
 out=Path(__file__).resolve().parent
 target_mfd=12.5 # LS PremierLink LE125 catalog; only fit target, NOT external validation
 # IMPORTANT: MFD target is Petermann-II estimate, not second-moment.
 # Fit only delta n given assumed core radius of 5.5 um; do not fit 2 unknowns to one measurement.
 radius=5.5
 def func(dn):return radial_lp01_aeff(Fiber(radius_um=radius,delta_n=dn))['MFD_petermann_ii_um']-target_mfd
 # choose bracket in physically plausible V<2.405 domain
 trial=[]
 for dn in np.linspace(.0010,.0045,50):
  f=Fiber(radius_um=radius,delta_n=float(dn))
  if f.v<2.404825557:trial.append((dn,func(float(dn))))
 br=None
 for (x,y),(xx,yy) in zip(trial,trial[1:]):
  if y*yy<0:br=(x,xx);break
 if br is None:raise RuntimeError('no calibration solution',trial)
 delta=brentq(func,*br,xtol=1e-13)
 f=Fiber(radius_um=radius,delta_n=float(delta));ore=lp01_analytic(f);metrics=radial_lp01_aeff(f)
 disp={str(k):dispersion_lp01(f,step=k) for k in (.006,.003,.0015)}
 meshes=[(40,160,30),(80,320,35),(160,640,40),(320,1280,45)]
 fem=[]
 for nc,ns,outer in meshes:
  d=solve_radial(f,core_elements=nc,clad_elements=ns,outer_radius_um=outer)
  row={k:v for k,v in d.items() if k not in ('radius_um','field')}
  row['abs_neff_error']=abs(d['n_eff']-ore['n_eff']);row['rel_aeff_error_pct']=100*(d['Aeff_um2']/metrics['Aeff_um2']-1)
  row['rel_mfd2_error_pct']=100*(d['MFD_2nd_um']/metrics['MFD_2nd_um']-1)
  fem.append(row)
  if len(fem)==3: radial_plot=d
 # check sensitivity to unknown profile: change core radius; recalibrate Delta n each time to MFD=12.5
 sensitivity=[]
 for rr in [5.0,5.25,5.5,5.75,6.0]:
  cand=[]
  for dn in np.linspace(.0006,.005,55):
   g=Fiber(radius_um=rr,delta_n=float(dn))
   if g.v<2.4047:cand.append((dn,radial_lp01_aeff(g)['MFD_petermann_ii_um']-target_mfd))
  bracket=next(((c[0],dd[0]) for c,dd in zip(cand,cand[1:]) if c[1]*dd[1]<0),None)
  if bracket:
   dnn=brentq(lambda n:radial_lp01_aeff(Fiber(radius_um=rr,delta_n=n))['MFD_petermann_ii_um']-target_mfd,*bracket)
   ff=Fiber(radius_um=rr,delta_n=dnn)
   sensitivity.append(dict(radius_um=rr,delta_n=dnn,V=ff.v,Aeff_um2=radial_lp01_aeff(ff)['Aeff_um2'],D_ps_nm_km=dispersion_lp01(ff)))
 # V-cutoff analytic ideal proxy. Cable cutoff NOT equivalent to theoretical cutoff.
 cutoff_proxy_nm=1550*f.v/2.4048255577
 wavelengths=np.linspace(1.53,1.625,50)
 nefs=[lp01_analytic(replace(f,wavelength_um=float(w)))['n_eff'] for w in wavelengths]
 Ds=[dispersion_lp01(replace(f,wavelength_um=float(w))) for w in wavelengths]
 fig,axs=plt.subplots(1,2,figsize=(11,4.2),constrained_layout=True)
 axs[0].plot(radial_plot['radius_um'],radial_plot['field']**2,label='radial FEM')
 axs[0].axvline(f.radius_um,ls='--',color='grey',label='core boundary')
 axs[0].set(xlabel='Radius (µm)',ylabel='Intensity (normalized)',title='Calibrated ideal step-index LP01')
 axs[0].legend()
 axs[1].plot([x['core_elements']+x['clad_elements'] for x in fem],[x['abs_neff_error'] for x in fem],'o-')
 axs[1].set_yscale('log');axs[1].set(xlabel='Radial elements',ylabel='|neff FEM − neff Bessel|',title='Radial FEM convergence')
 fig.savefig(out/'g654e_mode_convergence.png',dpi=160);plt.close(fig)
 fig,ax=plt.subplots(figsize=(7,4.5),constrained_layout=True)
 ax.plot(np.array(wavelengths)*1000,Ds,label='Ideal step-index predicted D')
 ax.axhspan(17,23,alpha=.16,label='ITU G.654.E D1550 specification range (1550nm only)')
 ax.axvline(1550,color='grey',ls='--')
 ax.set(xlabel='Wavelength (nm)',ylabel='D [ps/(nm km)]',title='Model dispersion; 1550 nm range benchmark')
 ax.legend(loc='best')
 fig.savefig(out/'g654e_dispersion.png',dpi=160);plt.close(fig)
 obj={'sources':{'fiber':'LS PremierLink LE125 brochure (2025): nominal MFD 12.5 um at 1550nm; loss <=0.17 dB/km','standard':'ITU-T G.654 (08/2024), Table 5: D1550 17–23 ps/(nm km), S1550 .050–.070; max loss 1550 0.23 dB/km','code':'COMSOL/waveoptics_m7/src/radial_fem.py and fiber_modes.py (equivalent radial FEM verification code)'},'caveat':'The manufactured LE125 refractive index profile is undisclosed; 12.5um MFD is a calibration TARGET, not an independent validation of product match. G.654.E MFD measurement convention may not equal model moment measure. No complex loss or cable cutoff predicted.','parameters':{'lambda_um':1.55,'radius_um':f.radius_um,'delta_n':delta,'n_core':f.core,'n_clad':f.clad,'V':f.v,'theoretical_step_index_cutoff_proxy_nm':cutoff_proxy_nm},'analytic':{**ore,**metrics,'D_ps_nm_km':disp[str(.003)]},'dispersion_step_convergence':disp,'radial_fem':fem,'sensitivity':sensitivity,'checks':{'radial_neff_abs_lt_1e_6':fem[-1]['abs_neff_error']<1e-6,'radial_aeff_rel_lt_1pct':abs(fem[-1]['rel_aeff_error_pct'])<1,'radial_mfd_2nd_rel_lt_1pct':abs(fem[-1]['rel_mfd2_error_pct'])<1,'radial_neff_error_refines':fem[-1]['abs_neff_error']<fem[0]['abs_neff_error'],'D1550_in_ITU_17_23':17<=disp[str(.003)]<=23,'MFD_petermann_calibration_lt_0p01':abs(metrics['MFD_petermann_ii_um']-12.5)<.01}}
 (out/'g654e_results.json').write_text(json.dumps(obj,ensure_ascii=False,indent=2))
 print(json.dumps(obj,ensure_ascii=False,indent=2))
if __name__=='__main__':run()