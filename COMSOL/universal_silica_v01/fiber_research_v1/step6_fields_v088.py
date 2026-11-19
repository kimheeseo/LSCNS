"""v0.8.8 common actual-FE probes and mode comparisons; lengths [um], weights [um²]."""
import numpy as np
try:
 from .modal_diagnostics import disk_probes,weighted_columns,principal_correlations
 from .step4_v086 import quadrature,digest
except ImportError:
 from modal_diagnostics import disk_probes,weighted_columns,principal_correlations
 from step4_v086 import quadrature,digest

def unpack(mode):
 q=mode['tracking_fields'];p=np.asarray(q['points_um'],float);w=np.asarray(q['weights_um2'],float);a=np.asarray(q['E'],float)
 if p.shape!=(len(w),2) or a.shape!=(len(w),3,2) or not np.isfinite(p).all() or not np.isfinite(a).all() or not np.isfinite(w).all() or np.any(w<=0):raise ValueError('finite matching FE probes required')
 return p,w,a[...,0]+1j*a[...,1]

def common(modes):
 p,w,e=unpack(modes[0]);fields=[e]
 for m in modes[1:]:
  pp,ww,ee=unpack(m)
  if not np.array_equal(p,pp) or not np.array_equal(w,ww):raise ValueError('identical physical probes and weights required')
  fields.append(ee)
 return p,w,np.stack(fields,axis=-1)

def subspace_overlap(a,b):
 _,w,f=common(a+b);n=len(a)
 return principal_correlations(weighted_columns(f[...,:n],w),weighted_columns(f[...,n:],w))

def spatial_overlap(a,b):
 _,w,f=common([a,b]);ia=np.sum(abs(f[:,:2,0])**2,axis=1);ib=np.sum(abs(f[:,:2,1])**2,axis=1)
 den=(ia@w)*(ib@w)
 if den<=0:raise ValueError('zero mode intensity')
 return float(np.clip((np.sqrt(ia*ib)@w)**2/den,0,1))

def polarization_pair(result,indices=None,minimum_spatial_overlap=.9):
 modes=result['modes'];lam=result['input']['wavelength_nm']
 if indices is None:
  candidates=[]
  for i,a in enumerate(modes):
   for j,b in enumerate(modes[:i]):
    labels={a['classification']['polarization_class'],b['classification']['polarization_class']}
    if labels=={'quasi_x','quasi_y'} and all(m.get('status')=='candidate_unverified' for m in (a,b)):
     candidates.append((spatial_overlap(a,b),j,i))
  if not candidates:return {'status':'TODO_IDENTIFY_ORTHOGONAL_SPATIAL_PAIR','B':None,'L_B_m':None}
  _,i,j=max(candidates);indices=[i,j]
 if len(indices)!=2 or indices[0]==indices[1]:raise ValueError('two distinct mode indices required')
 a,b=[modes[i] for i in indices];overlap=spatial_overlap(a,b)
 labels={a['classification']['polarization_class'],b['classification']['polarization_class']}
 if overlap<minimum_spatial_overlap or labels!={'quasi_x','quasi_y'} or any(m.get('status')!='candidate_unverified' for m in (a,b)):
  return {'status':'PAIR_REJECTED','indices':indices,'spatial_intensity_overlap':overlap,'B':None,'L_B_m':None}
 mx=a if a['classification']['polarization_class']=='quasi_x' else b;my=b if mx is a else a
 delta=mx['neff_real']-my['neff_real'];B=abs(delta)
 return {'status':'SPATIALLY_MATCHED_POLARIZATION_PAIR_CANDIDATE','indices':indices,'spatial_intensity_overlap':overlap,
         'signed_nx_minus_ny':delta,'B':B,'L_B_m':lam*1e-9/B if B>0 else None,
         'certified':False,'limitation':'spatial/polarization screening is not independent modal identification or mesh/PML convergence'}

def cartesian_probes(half_width_um,side=65):
 if not isinstance(side,int) or not 9<=side<=99 or not np.isfinite(half_width_um) or half_width_um<=0:raise ValueError('positive half width; grid side 9..99')
 axis=np.linspace(-half_width_um,half_width_um,side);x,y=np.meshgrid(axis,axis);v=np.ones(side);v[[0,-1]]=.5
 return np.column_stack((x.ravel(),y.ravel())),(np.outer(v,v)*(axis[1]-axis[0])**2).ravel()

def export_grid(result,mode_index=0):
 p,w,e=unpack(result['modes'][mode_index]);x=np.unique(p[:,0]);y=np.unique(p[:,1]);xx,yy=np.meshgrid(x,y)
 if len(x)<9 or len(y)<9 or p.shape!=(len(x)*len(y),2) or not np.array_equal(p,np.column_stack((xx.ravel(),yy.ravel()))) or not np.allclose(np.diff(x),np.diff(x)[0]) or not np.allclose(np.diff(y),np.diff(y)[0]):raise ValueError('ordered uniform rectangular actual-FE probe grid required')
 I=np.sum(abs(e[:,:2])**2,axis=1).reshape(len(y),len(x));edge=np.zeros_like(I,dtype=bool);edge[[0,-1],:]=True;edge[:,[0,-1]]=True
 return {'schema':'LSCNS_step6_field_grid_v088','wavelength_nm':result['input']['wavelength_nm'],'x_um':x.tolist(),'y_um':y.tolist(),
         'Et':np.stack((e[:,:2].real,e[:,:2].imag),axis=-1).reshape(len(y),len(x),2,2).tolist(),
         'source_design_sha256':digest(result['input']),'mode_index':mode_index,'family':result['input']['family'],
         'edge_peak_intensity_fraction':float(np.max(I[edge])/np.max(I)),
         'basis':'actual HCurl FE interpolation; finite rectangular aperture inside physical domain; PML excluded','certified':False}
