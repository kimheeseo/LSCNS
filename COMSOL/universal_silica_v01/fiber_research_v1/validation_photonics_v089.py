"""STEP7 independent parameterizations and actual FSM/HCF probes.
V_SIF=2*pi*a/lambda*sqrt(nc²-ncl²); b=(neff²-ncl²)/(nc²-ncl²).
V_PCF=2*pi*Lambda/lambda*sqrt(n_FM²-n_FSM²), cutoff pi (Mortensen 2003).
Only a finite sampled modal assessment is possible here; never declare ESM proof.
ARROW lambda_m=2*t*sqrt(ng²-na²)/m; a formula line is NOT a FEM resonance.
"""
import copy,itertools,math
import numpy as np
from scipy.optimize import brentq
from schema import load_presets
from vector_fem import solve
from step3_reference import he11_exact
from step5_vector_pwe import example_config,solve_point
from modal_diagnostics import disk_probes
from step6_fields_v088 import unpack,subspace_overlap
from step4_v086 import digest
from physical_metrics import confinement_loss


def bv_curve(V_samples=(1.8,2.,2.2,2.35)):
 nc,ncl,a=1.5,1.44,1.3;records=[]
 for V in V_samples:
  lam=2*np.pi*a*np.sqrt(nc*nc-ncl*ncl)/V;ref=he11_exact(nc,ncl,a,lam)
  d=copy.deepcopy(load_presets()['index_guiding__step']);d.update(name='V1 synthetic b-V exact HE11 benchmark',wavelength_nm=float(lam*1000),domain_radius_um=9.)
  d['parameters']={'core_radius_um':a,'delta_n':nc-ncl};d['materials']={'core':{'model':'constant','n':nc},'background':{'model':'constant','n':ncl}};d['boundary'].update(type='pec',thickness_um=0)
  d['mesh'].update(elements_per_wavelength=5,curvature_segments=192,bulk_max_um=.8);d['solver']={'element_order':2,'max_dofs':180000,'tolerance':1e-11}
  row={'V':V,'wavelength_nm':float(lam*1000),'reference_neff':ref['neff_real'],'reference_b':(ref['neff_real']**2-ncl*ncl)/(nc*nc-ncl*ncl),'input':d,'input_sha256':digest(d)}
  try:
   result=solve(d,n_modes=2,target_neff=ref['neff_real'],max_sample_points=10)
   modes=[m for m in result['modes'] if m['status']=='candidate_unverified']
   if len(modes)!=2:raise ValueError('missing two residual-qualified HE11 polarization candidates')
   row.update(status='COMPUTED',neff_real=[m['neff_real'] for m in modes],b=[(m['neff_real']**2-ncl*ncl)/(nc*nc-ncl*ncl) for m in modes],neff_error_ppm=max(abs(m['neff_real']-ref['neff_real'])/ref['neff_real']*1e6 for m in modes),solver=result['solver_diagnostics'])
  except (ValueError,RuntimeError,MemoryError) as e:row.update(status='FAILED',error=str(e))
  records.append(row);print('V1 b-V',V,row['status'],flush=True)
 return {'cases':records,'reference':'exact vector HE11 cylindrical J/K interface characteristic, infinite cladding; explicit synthetic n=1.5/1.44','neff_ppm_tolerance':10.,'tolerance_basis':'declared engineering acceptance, not a published COMSOL tolerance','curve_is_fundamental_HE11_only':True}


def fsm_index(pitch_um,wavelength_nm,ratio,cutoff):
 c=example_config();c.update(n_host=1.444,n_air=1.,pitch_um=pitch_um,hole_diameter_over_pitch=ratio,pwe_cutoff=cutoff,bands=2,kmesh_n=2,path_points_per_segment=2,beta_norms=[1.])
 f=pitch_um/(wavelength_nm*.001)
 def fun(beta):return solve_point(c,[0,0],beta)['frequencies_f'][0]-f
 beta=brentq(fun,max(.001*f,1e-6),c['n_host']*f,xtol=1e-10)
 s=solve_point(c,[0,0],beta)
 return {'n_FSM':beta/f,'pwe_cutoff':cutoff,'beta_norm':beta,'fixed_frequency_f':f,'frequency_root_residual':float(abs(fun(beta))),'eigen_residual':s['residual_normwise'],'divergence_relative':s['divergence_relative'],'basis':'fundamental transverse-H Maxwell Gamma-cell mode at fixed vacuum wavelength; beta root, two transverse polarizations','input':c}


def pcf_fsm_study():
 rows=[]
 for lam in (1200.,1800.):
  pitch=2.;ratio=.6;row={'wavelength_nm':lam,'pitch_um':pitch,'diameter_ratio':ratio,'lambda_over_pitch':lam/(1000*pitch)}
  try:
   cell=[fsm_index(pitch,lam,ratio,c) for c in (2,3,4)];row['FSM_convergence']=cell
   d=copy.deepcopy(load_presets()['solid_pcf__hex']);d.update(name='V6 Mortensen constant-index triangular one-cell defect numerical study',wavelength_nm=lam,domain_radius_um=8.)
   d['parameters'].update(pitch_um=pitch,d_over_pitch=ratio,rings=3,defect_cells=1);d['materials']['background']={'model':'constant','n':1.444};d['materials']['air']={'model':'air'}
   d['mesh'].update(elements_per_wavelength=4,through_wall_elements=3,bulk_max_um=.8,curvature_segments=64);d['boundary'].update(type='pml',thickness_um=1.5,strength=3.,order=3);d['solver']={'max_dofs':150000,'max_iterations':700,'tolerance':1e-9}
   p,w=disk_probes(.94*d['domain_radius_um'],pitch*.7,radial_order=16)
   result=solve(d,n_modes=6,target_neff=1.425,probe_points=p,probe_weights=w,max_sample_points=8)
   modes=[]
   for i,m in enumerate(result['modes']):
    points,weights,e=unpack(m);I=np.sum(abs(e[:,:2])**2,axis=1);fraction=float(np.sum(I[np.linalg.norm(points,axis=1)<pitch*.7]*weights[np.linalg.norm(points,axis=1)<pitch*.7])/(I@weights))
    modes.append({'index':i,'neff_real':m['neff_real'],'neff_imag':m['neff_imag'],'core_disk_Et_intensity_fraction':fraction,'status':m['status']})
   core=[m for m in modes if m['core_disk_Et_intensity_fraction']>.25 and m['status']=='candidate_unverified']
   row.update(input=d,input_sha256=digest(d),modes=modes)
   if not core:raise ValueError('no localized residual-qualified FM candidate')
   fm=max(core,key=lambda m:m['neff_real']);nf=cell[-1]['n_FSM'];rad=fm['neff_real']**2-nf**2
   if rad<=0:raise ValueError('FM candidate neff <= n_FSM; cannot apply guided-PCF V parameter')
   v=2*np.pi*pitch/(lam*.001)*np.sqrt(rad)
   row.update(status='CANDIDATE_NOT_CERTIFIED',n_FM=fm['neff_real'],n_FSM=nf,V_PCF=float(v),V_cutoff=float(np.pi),FSM_last_neff_change=abs(cell[-1]['n_FSM']-cell[-2]['n_FSM']),sampled_regime='BELOW_V_PI' if v<np.pi else 'ABOVE_V_PI')
  except (RuntimeError,ValueError,MemoryError) as e:row.update(status='FAILED',error=str(e))
  rows.append(row);print('V6 FSM',lam,row['status'],flush=True)
 return {'cases':rows,'status':'NEEDS_EVIDENCE','reference_url':'https://arxiv.org/pdf/physics/0307010','missing':'FM mesh/window/PML and wavelength/ring convergence, identified HOM cutoff, quantitative primary-reference digitization; two wavelengths cannot establish ESM','ESM_certified':False,'localization_basis':'Et intensity disk proxy, not core power identity proof'}


def arf_resonance_study():
 d=copy.deepcopy(load_presets()['anti_resonant__single_ring']);d.update(name='V7 synthetic single-ring ARF resonance study',domain_radius_um=13.)
 d['parameters'].update(core_radius_um=4.,tube_radius_um=3.5,wall_thickness_um=.55,tube_count=6,gap_um=.4)
 d['materials']['silica']={'model':'constant','n':1.45};d['mesh'].update(elements_per_wavelength=4,through_wall_elements=3,curvature_segments=64,bulk_max_um=.9);d['boundary'].update(thickness_um=1.5,strength=3.,order=3);d['solver']={'max_dofs':150000,'max_iterations':600,'tolerance':1e-9}
 arrow=2000*.55*np.sqrt(1.45**2-1);rows=[];previous=None;p,w=disk_probes(.93*13,4.,radial_order=16)
 for lam in arrow*np.array([.85,.9,.95,1.,1.05,1.1,1.15]):
  design=copy.deepcopy(d);design['wavelength_nm']=float(lam);target=float(np.sqrt(1-(2.4048255577*(lam*.001)/(2*np.pi*4.4))**2));row={'wavelength_nm':float(lam),'status':'FAILED','alpha_total_dB_m':None,'search_target_neff':target,'target_basis':'empty-capillary J0 scalar root seed only; not a validation value'}
  try:
   modes=solve(design,n_modes=6,target_neff=target,probe_points=p,probe_weights=w,max_sample_points=6)['modes'];good=[]
   for i,m in enumerate(modes):
    pp,ww,e=unpack(m);I=np.sum(abs(e[:,:2])**2,axis=1);inside=np.linalg.norm(pp,axis=1)<4.
    frac=float(np.sum(I[inside]*ww[inside])/(I@ww))
    if m['status']=='candidate_unverified':good.append((i,m,frac))
   groups=[]
   for a,b in itertools.combinations(good,2):
    if abs(a[1]['neff_real']-b[1]['neff_real'])>1e-4:continue
    corr=subspace_overlap(previous,[a[1],b[1]]) if previous else [1.,1.]
    groups.append((min(corr) if corr else 0.,(a[2]+b[2])/2,[a,b]))
   if not groups:raise ValueError('no valid two-polarization candidate subspace')
   quality,corefrac,pair=max(groups,key=lambda g:g[0]*g[1])
   if quality<.8 or corefrac<.25:raise ValueError('core localization/overlap failed: '+str((quality,corefrac)))
   previous=[m for _,m,_ in pair];row.update(status='TRACKED_CANDIDATE',core_disk_Et_fraction=corefrac,minimum_subspace_correlation=quality,alpha_total_dB_m=float(np.mean([m['alpha_total_dB_m'] for m in previous])),neff_real=[m['neff_real'] for m in previous],neff_imag=[m['neff_imag'] for m in previous])
  except (RuntimeError,ValueError,MemoryError) as e:row['error']=str(e)
  rows.append(row);print('V7 ARF',int(lam),row['status'],flush=True)
 good=[r for r in rows if r['status']=='TRACKED_CANDIDATE'];peak=None
 if len(good)==len(rows):
  j=int(np.argmax([r['alpha_total_dB_m'] for r in rows]))
  if 0<j<len(rows)-1:peak=rows[j]['wavelength_nm']
 return {'cases':rows,'input':d,'input_sha256':digest(d),'arrow_lambda_nm':float(arrow),'observed_interior_peak_nm':peak,'resonance_relative_error':abs(peak-arrow)/arrow if peak else None,
 'tolerance_relative':.05,'tolerance_basis':'declared numerical acceptance, not a published uncertainty','status':'NEEDS_EVIDENCE' if peak else 'FAIL',
 'missing':'same-branch independent PML/mesh sweep, spectral peak refinement and physical FM identity; analytic ARROW line alone is not a FEM validation','loss_certified':False}
