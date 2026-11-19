"""STEP6d formulas, SI/um units fixed.
eta=|int Et1 dot conj(Et2) dA|²/(int|Et1|² int|Et2|²), IL=-10log10(eta).
Offset [um]; angular phase exp(i*k_gap*(x*sin(tx)+y*sin(ty))) with x [m].
Gap angular spectrum: H(kx,ky)=exp(i*z*sqrt(k_gap²-kx²-ky²)), outgoing branch.
Normal scalar Fresnel T=4*n1*n2/(n1+n2)² for real positive indices.
Gaussian MFD=2w: eta_size=(2*w1*w2/(w1²+w2²))².
This requested Et metric is a weak-reflection coupling estimate, NOT full E/H
mode matching across high-contrast SMF/HCF ends; no splice-collapse/scattering model.
Sources: COMSOL single_mode_fiber_coupling; Ansys Understanding mode overlap;
Matsushima/Shimobaba, doi:10.1364/OE.17.019662 (sampling limits of angular spectrum).
"""
import math
import numpy as np
from scipy.interpolate import RegularGridInterpolator

SOURCES=['https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.single_mode_fiber_coupling/single_mode_fiber_coupling.html',
 'https://optics.ansys.com/hc/en-us/articles/360034396834-Understanding-the-mode-overlap-calculation','https://doi.org/10.1364/OE.17.019662']

def _grid(g):
 if g.get('schema')!='LSCNS_step6_field_grid_v088':raise ValueError('STEP6 field grid schema required')
 x=np.asarray(g['x_um'],float);y=np.asarray(g['y_um'],float);a=np.asarray(g['Et'],float)
 if len(x)<9 or len(y)<9 or a.shape!=(len(y),len(x),2,2) or not np.isfinite(a).all() or not np.isfinite(x).all() or not np.isfinite(y).all() or np.any(np.diff(x)<=0) or np.any(np.diff(y)<=0) or not np.allclose(np.diff(x),np.diff(x)[0]) or not np.allclose(np.diff(y),np.diff(y)[0]):raise ValueError('finite uniform grid required')
 lam=float(g['wavelength_nm'])
 if not math.isfinite(lam) or lam<=0:raise ValueError('positive wavelength')
 return x,y,a[...,0]+1j*a[...,1],lam

def gaussian_grid(mfd_um,wavelength_nm=1550,half_width_um=30,side=129,polarization=(1,0)):
 if mfd_um<=0 or half_width_um<=0 or side<9:raise ValueError('positive MFD/grid')
 x=np.linspace(-half_width_um,half_width_um,side);xx,yy=np.meshgrid(x,x);pol=np.asarray(polarization,complex)
 if pol.shape!=(2,) or not np.isfinite(pol).all() or np.linalg.norm(pol)==0:raise ValueError('finite nonzero Jones vector')
 e=np.exp(-(xx*xx+yy*yy)/(mfd_um/2)**2)[...,None]*pol
 return {'schema':'LSCNS_step6_field_grid_v088','wavelength_nm':wavelength_nm,'x_um':x.tolist(),'y_um':x.tolist(),
         'Et':np.stack((e.real,e.imag),axis=-1).tolist(),'basis':'analytic Gaussian verification only; not a simulated SMF or HCF'}

def fresnel(n1,n2):
 if not np.isfinite([n1,n2]).all() or min(n1,n2)<=0:raise ValueError('positive finite real interface indices required')
 return float(4*n1*n2/(n1+n2)**2)

def splice(g1,g2,offset_um=(0,0),angle_deg=(0,0),gap_um=0,gap_n=1.,interface_indices=None,padding=3):
 x,y,e,lam=_grid(g1);x2,y2,f,lam2=_grid(g2)
 values=np.array([*offset_um,*angle_deg,gap_um,gap_n],float)
 if values.shape!=(6,) or not np.isfinite(values).all() or gap_um<0 or gap_n<=0 or np.any(abs(np.asarray(angle_deg))>10):raise ValueError('finite offset, |angle|<=10deg, nonnegative gap and positive gap index')
 if abs(lam-lam2)>1e-8:raise ValueError('coupling modes must share wavelength')
 if padding not in (2,3,4):raise ValueError('FFT padding factor 2..4')
 xx,yy=np.meshgrid(x,y);k=2*np.pi*gap_n/(lam*1e-3);tilt=np.sin(np.deg2rad(angle_deg))
 if abs(k*tilt[0]*(x[1]-x[0]))>=np.pi/2 or abs(k*tilt[1]*(y[1]-y[0]))>=np.pi/2:raise ValueError('angular phase undersampled; refine field grid')
 e=e*np.exp(1j*k*(xx*tilt[0]+yy*tilt[1]))[...,None]
 incoming=float(np.sum(abs(e)**2)*np.diff(x)[0]*np.diff(y)[0])
 capture=1.
 if gap_um:
  ny,nx=e.shape[:2];py=padding*ny;px=padding*nx;oy=(py-ny)//2;ox=(px-nx)//2
  field=np.zeros((py,px,2),complex);field[oy:oy+ny,ox:ox+nx]=e
  kx=2*np.pi*np.fft.fftfreq(px,x[1]-x[0]);ky=2*np.pi*np.fft.fftfreq(py,y[1]-y[0]);kx,ky=np.meshgrid(kx,ky)
  kz=np.sqrt((k*k-kx*kx-ky*ky).astype(complex))
  propagated=np.fft.ifft2(np.fft.fft2(field,axes=(0,1))*np.exp(1j*kz*gap_um)[...,None],axes=(0,1))
  e=propagated[oy:oy+ny,ox:ox+nx]
  capture=float(np.sum(abs(e)**2)*np.diff(x)[0]*np.diff(y)[0]/incoming)
 query=np.column_stack(((yy-offset_um[1]).ravel(),(xx-offset_um[0]).ravel()))
 interpolator=RegularGridInterpolator((y2,x2),f,bounds_error=False,fill_value=0)
 receiver=interpolator(query).reshape(e.shape)
 # Normalize receiving mode on ITS full aperture; do not renormalize a clipped shifted mode.
 wx=np.ones(len(x));wy=np.ones(len(y));wx[[0,-1]]=.5;wy[[0,-1]]=.5
 w=np.outer(wy,wx)*(x[1]-x[0])*(y[1]-y[0])
 vx=np.ones(len(x2));vy=np.ones(len(y2));vx[[0,-1]]=.5;vy[[0,-1]]=.5
 wf=np.outer(vy,vx)*(x2[1]-x2[0])*(y2[1]-y2[0])
 source_norm=float(np.sum(np.sum(abs(_grid(g1)[2])**2,axis=-1)*w));dest_norm=float(np.sum(np.sum(abs(f)**2,axis=-1)*wf))
 if min(source_norm,dest_norm)<=0:raise ValueError('zero mode norm')
 amplitude=np.sum(np.sum(e*np.conj(receiver),axis=-1)*w)
 eta=float(abs(amplitude)**2/(source_norm*dest_norm))
 if eta>1+2e-3:raise ValueError('overlap >1: grid/interpolation inconsistent; refine apertures and sampling')
 eta=min(1.,eta);T=1.
 if interface_indices is not None:
  if not isinstance(interface_indices,list) or len(interface_indices)!= (3 if gap_um else 2):raise ValueError('explicit [n1,n2] for butt or [n1,ngap,n2] for gap')
  if gap_um and not np.isclose(interface_indices[1],gap_n):raise ValueError('interface gap index must equal propagation gap_n')
  T=float(np.prod([fresnel(a,b) for a,b in zip(interface_indices[:-1],interface_indices[1:])]))
 total=eta*T
 return {'schema':'LSCNS_step6_splice_v088','schema_version':'1.0.0','engine_version':'0.8.8','source_designs':[g.get('source_design_sha256') for g in (g1,g2)],
 'settings':{'offset_um':list(offset_um),'angle_deg':list(angle_deg),'gap_um':gap_um,'gap_n':gap_n,'interface_indices':interface_indices,'padding':padding},
 'eta_Et':eta,'fresnel_T':T if interface_indices is not None else None,'eta_with_requested_fresnel':total if interface_indices is not None else None,
 'overlap_loss_dB':-10*math.log10(eta) if eta>0 else None,'combined_loss_dB':-10*math.log10(total) if total>0 and interface_indices is not None else None,
 'gap_aperture_power_fraction':capture,'validation':{'certified':False,'sources':SOURCES,
 'limitations':['Et overlap approximation; full E/H matching required for high contrast SMF-HCF','finite aperture, interpolation, FFT padding and gap step require convergence','normal scalar Fresnel is optional user-input estimate; no multiple reflections or angled coating/interface solve','no fusion collapse, roughness, taper, surface scattering or radiation continuum'],
 'input_field_edge_peak_fraction':[g1.get('edge_peak_intensity_fraction'),g2.get('edge_peak_intensity_fraction')],
 'aperture_screen':'FAILED_OR_UNLOCALIZED_MODE' if any(g.get('edge_peak_intensity_fraction',0)>.01 for g in (g1,g2)) else 'NOT_CERTIFIED_WINDOW_CONVERGENCE_REQUIRED',
 'fresnel_status':'USER_REQUESTED_SCALAR_ESTIMATE' if interface_indices is not None else 'TODO_USER_INTERFACE_INDICES_NO_ASSUMED_NEFF'}}
