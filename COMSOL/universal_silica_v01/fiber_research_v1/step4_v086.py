"""STEP4 v0.8.6 physical quantities. SI formulas fixed here; no fitted constants.

exp(i beta z-i omega t), k0=2*pi/lambda_m, c=299792458 m/s.
alpha[dB/m]=(20/ln(10))*k0*Im(neff), dB/km=1000*alpha. NO absolute value.
Et²=|Ex|²+|Ey|²; Aeff[um²]=(sum Et²*w)^2/sum Et^4*w (w in um²).
MFD2nd[um]=2*sqrt(2*<|r-rbar|²>_Et²); directional MFD=4*sqrt(<x²>).
Petermann-II[um]=2*sqrt(2*int |E|²*r dr/int |dE/dr|²*r dr),
ONLY for independently supplied axisymmetric scalar radial amplitude.
D[ps/(nm km)]=-lambda_nm/c*1e15*n''_nm;
S[ps/(nm² km)]=-1e15/c*(n''_nm+lambda_nm*n'''_nm).
gamma[1/(W m)]=2*pi*n2_eff/(lambda_m*Aeff_um2*1e-12).
n2_eff=sum n2_material*|Et|^4*w/sum |Et|^4*w: scalar Kerr proxy.
P_material=sum Sz*w*1e-12, Sz=.5*Re(Ex*conj(Hy)-Ey*conj(Hx)).
B=|Re(nx)-Re(ny)|; Lb[m]=lambda_m/B. HOMER=min(alpha_HOM)/max(alpha_FM).
ARROW: lambda_nm=2000*t_um*sqrt(ng(lambda)^2-na(lambda)^2)/m.
All refractive indices/n2/identities supplied or sourced; no air/glass n2 defaults.
"""
from __future__ import annotations
import copy,hashlib,json,math
from pathlib import Path
import numpy as np
from scipy.optimize import linear_sum_assignment,brentq
try:
 from .physical_metrics import confinement_loss,nonlinear_gamma,petermann_ii_radial,birefringence,surface_anticrossing_candidates
 from .materials import index,provenance
 from .modal_diagnostics import weighted_columns,principal_correlations
except ImportError:
 from physical_metrics import confinement_loss,nonlinear_gamma,petermann_ii_radial,birefringence,surface_anticrossing_candidates
 from materials import index,provenance
 from modal_diagnostics import weighted_columns,principal_correlations
VERSION='0.8.6'
C=299792458.

def digest(d):return hashlib.sha256(json.dumps(d,sort_keys=True,separators=(',',':'),allow_nan=False).encode()).hexdigest()
def _complex(v):
 a=np.asarray(v,float)
 if a.ndim!=2 or a.shape[1]!=2 or not np.isfinite(a).all():raise ValueError('finite complex [N,2] required')
 return a[:,0]+1j*a[:,1]
def quadrature(mode):
 q=mode.get('field_quadrature')
 if not q:raise ValueError('actual full field_quadrature required; field_samples cannot be integrated')
 p=np.asarray(q['points_um'],float);w=np.asarray(q['weights_um2'],float)
 f={k:_complex(q[k]) for k in ('Ex','Ey','Ez','Hx','Hy','Hz')}
 labels=np.asarray(q['material_labels'],str)
 if p.shape!=(len(w),2) or len(w)<3 or not np.isfinite(p).all() or not np.isfinite(w).all() or np.any(w<=0) or labels.shape!=w.shape or any(v.shape!=w.shape for v in f.values()):raise ValueError('matching finite physical quadrature/labels required')
 if np.any(labels=='pml'):raise ValueError('PML quadrature must be excluded')
 return p,w,f,labels

def integrate_mode(mode,design,context=None):
 ctx=context or {};p,w,f,labels=quadrature(mode)
 I=abs(f['Ex'])**2+abs(f['Ey'])**2;P=float(I@w);Q=float((I*I)@w)
 if min(P,Q)<=0 or not np.isfinite([P,Q]).all():raise ValueError('nonzero finite Et integrals required')
 center=np.sum(p*(w*I)[:,None],axis=0)/P
 var=np.sum((p-center)**2*(w*I)[:,None],axis=0)/P
 try:
  from .lswave.postprocess.formulas import area as canonical_area
 except ImportError:
  from lswave.postprocess.formulas import area as canonical_area
 ae=canonical_area(P,Q)
 sz=.5*np.real(f['Ex']*np.conj(f['Hy'])-f['Ey']*np.conj(f['Hx']))
 power=float(sz@w*1e-12)
 if not np.isfinite(power) or power<=0:raise ValueError('positive net physical Poynting power required')
 fractions={str(m):float(np.sum(sz[labels==m]*w[labels==m])*1e-12/power) for m in np.unique(labels)}
 quartic={str(m):float(np.sum((I[labels==m]**2)*w[labels==m])/Q) for m in np.unique(labels)}
 roles=dict(ctx.get('material_roles',{}))
 for label in np.unique(labels):
  model=design.get('materials',{}).get(label,{}).get('model')
  if label not in roles:
   if model=='air' or label=='air':roles[str(label)]='air'
   elif model=='pure_silica' or label=='silica':roles[str(label)]='glass'
 role_known=all(roles.get(str(m)) in ('air','glass') for m in np.unique(labels))
 loss=confinement_loss(mode['neff_imag'],design['wavelength_nm'])
 absorption=any(float(m.get('kappa',0))!=0 for m in design['materials'].values())
 out={'neff_real':mode['neff_real'],'neff_imag':mode['neff_imag'],**loss,
      'alpha_confinement_candidate_dB_m':None if absorption else loss['alpha_total_dB_m'],
      'loss_status':'TOTAL_INCLUDES_MATERIAL_ABSORPTION' if absorption else 'RADIATION_CANDIDATE_REQUIRES_MESH_PML_STUDY',
      'Aeff_Et_um2':ae,'MFD_second_moment_Et_um':float(2*np.sqrt(2*sum(var))),
      'MFD_x_um':float(4*np.sqrt(var[0])),'MFD_y_um':float(4*np.sqrt(var[1])),
      'centroid_um':center.tolist(),'power_W':power,'material_power_fractions':fractions,
      'quartic_material_weights':quartic,'P_glass_fraction':sum(v for k,v in fractions.items() if roles.get(k)=='glass') if role_known else None,
      'material_role_status':'EXPLICIT_OR_MATERIAL_MODEL' if role_known else 'TODO_MATERIAL_ROLES',
      'Poynting_fraction_warning':any(x<0 or x>1 for x in fractions.values()),
      'MFD_PetermannII_um':None,'MFD_PetermannII_status':'TODO_AXISYMMETRIC_SCALAR_RADIAL_OR_VERIFIED_FAR_FIELD',
      'n2_eff_m2_W':None,'gamma_W_inv_m_inv':None,'nonlinearity_status':'TODO_SOURCED_N2_FOR_EVERY_PRESENT_MATERIAL',
      'basis':'actual physical FE Gauss quadrature, Et-only area/MFD; PML excluded',
      'confinement_loss_certified':False}
 radial=ctx.get('radial_profile')
 if radial:
  if radial.get('basis')!='axisymmetric_scalar' or not radial.get('source'):raise ValueError('PII requires explicit axisymmetric scalar radial basis and source')
  rr=petermann_ii_radial(radial['r_um'],_complex(radial['amplitude']))
  out.update(rr);out['MFD_PetermannII_status']='RADIAL_FORMULA_WINDOW_DERIVATIVE_CONVERGENCE_REQUIRED'
 nonlinear=ctx.get('nonlinear_materials',{})
 if all(k in nonlinear and nonlinear[k].get('source') for k in quartic):
  n2=0.
  for label,weight in quartic.items():
   value=nonlinear[label].get('n2_m2_W')
   if isinstance(value,bool) or not isinstance(value,(int,float)) or not np.isfinite(value) or value<0:raise ValueError('finite nonnegative sourced n2 required')
   n2+=weight*value
  out['n2_eff_m2_W']=n2;out.update(nonlinear_gamma(n2,design['wavelength_nm'],ae))
  out['nonlinearity_status']='QUARTIC_SCALAR_KERR_PROXY_NOT_FULL_VECTOR_TENSOR';out['n2_sources']=nonlinear
 return out

def arrow_dispersive(design,orders=(1,2,3,4),interval_nm=(250.,3500.)):
 p=design.get('parameters',{});m=design.get('materials',{})
 if 'wall_thickness_um' not in p or not all(k in m for k in ('silica','air')):return {'status':'TODO_WALL_GLASS_AIR','resonance_lines':[]}
 t=float(p['wall_thickness_um']);lo,hi=map(float,interval_nm)
 if not (np.isfinite(t) and t>0 and 250<=lo<hi<=3500):raise ValueError('positive wall and supported wavelength interval required')
 lines=[]
 for order in orders:
  if isinstance(order,bool) or not isinstance(order,int) or not 1<=order<=50:raise ValueError('ARROW integer orders 1..50')
  def eq(lam):
   ng=index(m['silica'],lam).real;na=index(m['air'],lam).real
   if ng<=na:raise ValueError('glass n must exceed air n')
   return lam-2000*t*np.sqrt(ng*ng-na*na)/order
  if eq(lo)*eq(hi)<=0:
   root=brentq(eq,lo,hi,xtol=1e-9);lines.append({'m':order,'lambda_nm':float(root),'equation_residual_nm':float(eq(root))})
 return {'status':'DISPERSIVE_ARROW_WALL_APPROXIMATION_NOT_FEM','resonance_lines':lines,
         'orders_without_root_in_interval':[i for i in orders if i not in [x['m'] for x in lines]],
         'sources':{k:provenance(m[k]) for k in ('silica','air')}}

def postprocess(data,context=None):
 if data.get('schema')!='LSCNS_fiber_research_vector_v1' or not data.get('modes'):raise ValueError('Python vector FEM result required')
 ctx=context or {};d=data['input'];modes=[dict(mode_index=i,**integrate_mode(m,d,{**ctx,**ctx.get('per_mode',{}).get(str(i),{})})) for i,m in enumerate(data['modes'])]
 pair=ctx.get('polarization_pair');B={'status':'TODO_VERIFIED_X_Y_PAIR'}
 if pair is not None:
  if not ctx.get('polarization_identity_source') or len(pair)!=2 or pair[0]==pair[1] or any(type(i)!=int or not 0<=i<len(modes) for i in pair):raise ValueError('verified distinct x/y indices and identity source required')
  B=dict(status='IDENTITY_SUPPLIED_PHASE_BIREFRINGENCE',**birefringence(modes[pair[0]]['neff_real'],modes[pair[1]]['neff_real'],d['wavelength_nm']),identity_source=ctx['polarization_identity_source'])
 hm={'status':'TODO_FM_HOM_ID_AND_CONVERGED_LOSSES','HOMER_ratio':None}
 ids=ctx.get('mode_identities',{})
 if ctx.get('identities_source') and ctx.get('losses_converged') is True:
  fm=[x['alpha_total_dB_m'] for i,x in enumerate(modes) if ids.get(str(i))=='FM'];hom=[x['alpha_total_dB_m'] for i,x in enumerate(modes) if ids.get(str(i))=='HOM']
  if fm and hom and min(fm+hom)>0:hm={'status':'INPUT_LABELLED_TOTAL_LOSS_RATIO_NOT_INDEPENDENT_CERTIFICATION','HOMER_ratio':min(hom)/max(fm),'identity_source':ctx['identities_source']}
 spacing=[{'i':i,'j':j,'delta_Re_neff':abs(a['neff_real']-b['neff_real']),'delta_Im_neff':abs(a['neff_imag']-b['neff_imag']),
           'abs_complex_delta_neff':abs(complex(a['neff_real']-b['neff_real'],a['neff_imag']-b['neff_imag']))} for i,a in enumerate(modes) for j,b in enumerate(modes) if j>i]
 return {'schema':'LSCNS_step4_metrics_v086','engine_version':VERSION,'input':d,'input_digest':digest(d),'context':ctx,
         'wavelength_nm':d['wavelength_nm'],'modes':modes,'birefringence':B,'HOMER':hm,'mode_neff_spacing':spacing,'arrow_overlay':arrow_dispersive(d),
         'validation':{'physical_loss_certified':False,'scope':'field integration and explicit supplied identity/material context; no product accuracy certification'}}

def dispersion(lam,neff,overlaps,threshold=.75):
 x=np.asarray(lam,float);n=np.asarray(neff,float);o=np.asarray(overlaps,float)
 if x.ndim!=1 or len(x)<5 or n.shape!=x.shape or o.shape!=(len(x)-1,) or not np.isfinite(np.r_[x,n,o]).all() or np.any(np.diff(x)<=0) or not 0<threshold<=1 or np.any(o<threshold) or np.any(o>1+1e-12):raise ValueError('>=5 ordered samples and verified adjacent overlap required')
 n2=[];n3=[]
 for i in range(len(x)):
  start=min(max(i-2,0),len(x)-5);sl=slice(start,start+5);scale=float(max(abs(x[sl]-x[i])));z=(x[sl]-x[i])/scale
  a=np.polynomial.polynomial.polyfit(z,n[sl]-n[i],4);n2.append(2*a[2]/scale**2);n3.append(6*a[3]/scale**3)
 n2=np.asarray(n2);n3=np.asarray(n3);D=-x/C*1e15*n2;S=-1e15/C*(n2+x*n3)
 floor=float(100*np.finfo(float).eps*max(abs(n))*max(x)/C*1e15/min(np.diff(x))**2)
 zero=[]
 if np.max(abs(D))>floor:
  for i in range(len(x)-1):
   if abs(D[i])<=floor:
    if not zero or x[i]-zero[-1]>min(np.diff(x))*.5:zero.append(float(x[i]))
   elif D[i]*D[i+1]<0 and abs(D[i+1])>floor:zero.append(float(x[i]-D[i]*(x[i+1]-x[i])/(D[i+1]-D[i])))
  if abs(D[-1])<=floor:zero.append(float(x[-1]))
 return {'wavelength_nm':x.tolist(),'neff_real':n.tolist(),'D_ps_nm_km':D.tolist(),'S_ps_nm2_km':S.tolist(),'zero_dispersion_lambda_nm':zero,
         'zero_status':'NO_ISOLATED_ZERO_NUMERICALLY_FLAT' if np.max(abs(D))<=floor else 'SAMPLED_ZERO_CANDIDATES_REQUIRE_STEP_CONVERGENCE',
         'D_roundoff_screen_ps_nm_km':floor,'derivative_method':'local scaled five-point quartic; analytic 2nd/3rd derivatives; no optical-accuracy certification'}

def _probe(mode):
 q=mode.get('tracking_fields')
 if q:
  e=np.asarray(q['E'],float);f=e[...,0]+1j*e[...,1];p=np.asarray(q['points_um']);w=np.asarray(q['weights_um2'])
 else:
  q=mode.get('field_quadrature')
  if not q:raise ValueError('common FE probes or full quadrature required, display samples are insufficient')
  p=np.asarray(q['points_um']);w=np.asarray(q['weights_um2']);f=np.column_stack([_complex(q[k]) for k in ('Ex','Ey')])
 if p.shape!=(len(w),2) or f.shape[0]!=len(w) or not np.isfinite(p).all():raise ValueError('invalid physical FE probe points')
 weighted_columns(f[:,:,None],w)
 return p,w,f

def _groups(modes,tol):
 remaining=set(range(len(modes)));groups=[]
 while remaining:
  group={min(remaining)};remaining-=group
  while True:
   add={j for j in remaining if any(abs(complex(modes[j]['neff_real'],modes[j]['neff_imag'])-complex(modes[i]['neff_real'],modes[i]['neff_imag']))<=tol for i in group)}
   if not add:break
   group|=add;remaining-=add
  groups.append(sorted(group))
 return groups

def track_sweep(raw,minimum_overlap=.75,degeneracy_tolerance=1e-6,parameter=None):
 runs=raw.get('runs',[])
 if not 0<minimum_overlap<=1 or not 0<degeneracy_tolerance<=.001 or len(runs)<2:raise ValueError('valid overlap/tolerance and >=2 runs required')
 key=parameter or raw.get('parameter','wavelength_nm');values=raw.get('parameter_values',[r['input']['wavelength_nm'] for r in runs])
 if len(values)!=len(runs) or not np.isfinite(values).all() or np.any(np.diff(values)<=0):raise ValueError('ordered finite sweep values required')
 reference=copy.deepcopy(runs[0]['input'])
 def remove(d):
  e=copy.deepcopy(d)
  if key=='wavelength_nm':e.pop(key,None)
  elif key.startswith('parameters.'):e['parameters'].pop(key.split('.',1)[1],None)
  else:raise ValueError('parameter must be wavelength_nm or parameters.<name>')
  return e
 if any(remove(r['input'])!=remove(reference) for r in runs):raise ValueError('non-swept physical/numerical inputs differ')
 for run,value in zip(runs,values):
  actual=run['input']['wavelength_nm'] if key=='wavelength_nm' else run['input']['parameters'][key.split('.',1)[1]]
  if actual!=value:raise ValueError('sweep values must match actual input')
 if key=='wavelength_nm' and len(runs)<5:raise ValueError('dispersion requires >=5 wavelengths')
 allprobes=[[_probe(m) for m in r['modes']] for r in runs]
 p,w=allprobes[0][0][:2]
 for probes in allprobes:
  for pp,ww,_ in probes:
   if not np.array_equal(pp,p) or not np.array_equal(ww,w):raise ValueError('identical physical probe points and weights required; verified interpolation needed')
 groups=[_groups(r['modes'],degeneracy_tolerance) for r in runs]
 order=list(range(len(groups[0])));tracks=[{'mode_id':'initial_group_'+str(i+1),'kind':'degenerate_subspace_centroid' if len(g)>1 else 'individual_mode','members':[],'neff_real':[],'neff_imag':[],'overlap_power':[],'overlap_verified':True} for i,g in enumerate(groups[0])]
 def cols(k,g):return weighted_columns(np.stack([allprobes[k][j][2] for j in g],axis=-1),w)
 for k,run in enumerate(runs):
  if k:
   old,new=groups[k-1],groups[k]
   if sorted(map(len,old))!=sorted(map(len,new)):raise ValueError('degenerate topology/mode count changed; dispersion BLOCKED')
   score=np.zeros((len(old),len(new)))
   for i,g in enumerate(old):
    for j,h in enumerate(new):
     if len(g)==len(h):
      c=principal_correlations(cols(k-1,g),cols(k,h));score[i,j]=min(c,default=0)**2
   rr,cc=linear_sum_assignment(-score);lookup=dict(zip(rr,cc));next_order=[]
   for i,prev in enumerate(order):
    j=int(lookup[prev]);s=float(score[prev,j])
    if s<minimum_overlap:raise ValueError(f'overlap {s:.4g} below threshold; dispersion BLOCKED at step {k}')
    tracks[i]['overlap_power'].append(s);next_order.append(j)
   order=next_order
  for i,j in enumerate(order):
   members=groups[k][j];tracks[i]['members'].append(members)
   for keyn in ('neff_real','neff_imag'):tracks[i][keyn].append(float(np.mean([run['modes'][m][keyn] for m in members])))
 for t in tracks:
  t['parameter_values']=list(map(float,values))
  if key=='wavelength_nm':t.update(dispersion(values,t['neff_real'],t['overlap_power'],minimum_overlap));t['alpha_total_dB_m']=[confinement_loss(im,lam)['alpha_total_dB_m'] for im,lam in zip(t['neff_imag'],values)]
 return {'schema':'LSCNS_step4_tracked_sweep_v086','engine_version':VERSION,'input':reference,'parameter':key,'parameter_values':list(map(float,values)),
         'minimum_overlap':minimum_overlap,'degeneracy_tolerance':degeneracy_tolerance,'curves':tracks,
         'arrow_overlay':arrow_dispersive(reference) if key=='wavelength_nm' else None,
         'validation':{'tracking':'PASS_FIELD_OR_SUBSPACE_OVERLAP','dispersion_accuracy_certified':False,'degenerate_individual_polarization_identity':'NOT_ASSIGNED'},
         'surface_anticrossing':{'status':'TODO_VERIFIED_GLASS_LOCALIZED_BRANCH_PAIR'}}

def detect_surface_pair(tracked,first,second,glass_fractions,identity_source):
 if not identity_source or tracked.get('parameter')!='wavelength_nm':raise ValueError('wavelength sweep and surface identity source required')
 a,b=tracked['curves'][first],tracked['curves'][second]
 if a['kind']!='individual_mode' or b['kind']!='individual_mode':raise ValueError('resolved individual branches needed for surface anticrossing')
 pair={'overlap_verified':True,'glass_mode_identified':True,'wavelength_nm':tracked['parameter_values'],'neff_core':a['neff_real'],'neff_glass':b['neff_real'],
       'glass_fraction_core':glass_fractions[first],'glass_fraction_glass':glass_fractions[second]}
 return dict(surface_anticrossing_candidates(pair),identity_source=identity_source)
