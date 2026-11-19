"""STEP6a-c additive orchestration of existing Gmsh / mixed HCurl FEM.
Bending n_eq=n(1+x_out/R_eff); losses signed dB/m, R [mm].
PM: delta(Q)=p:strain, epsilon_r=Q^-1; B=|nx-ny|, L_B=lambda_m/B.
MCF: beta=k0*neff [1/m], kappa=(beta_even-beta_odd)/2 [1/m].
Lc=pi/(2|Re kappa|) only for identical symmetric, negligible-loss 2-core model.
Twist is not a rotated 2D eigenproblem: helicoidal Maxwell/Jones z evolution TODO.
No output of this module implies independent loss/birefringence certification.
"""
import copy,itertools,math
import numpy as np
try:
 from .geometry import build_mesh,mcf_centers
 from .vector_fem import solve
 from .bending import configure_bend,bend_per_turn
 from .stress_optic import solve_thermal_stress
 from .step6_fields_v088 import disk_probes,unpack,subspace_overlap,polarization_pair,quadrature,digest
except ImportError:
 from geometry import build_mesh,mcf_centers
 from vector_fem import solve
 from bending import configure_bend,bend_per_turn
 from stress_optic import solve_thermal_stress
 from step6_fields_v088 import disk_probes,unpack,subspace_overlap,polarization_pair,quadrature,digest

VERSION='0.8.8'

def _root(kind,design):return {'schema':'LSCNS_step6_'+kind+'_v088','schema_version':'1.0.0','engine_version':VERSION,'source_design':design,'source_design_sha256':digest(design),'certified':False}

def bend_sweep(design,radii_mm,directions_deg,n_modes=4,target_neff=None,reference_indices=None,minimum_correlation=.8,mesh=None,solve_fn=None,**stress_correction):
 if design['boundary']['type']!='pml' or design.get('_bend'):raise ValueError('straight design with PML required')
 if not 1<=len(radii_mm)<=30 or not 1<=len(directions_deg)<=12 or len(radii_mm)*len(directions_deg)>100:raise ValueError('bend grid budget: <=100 cases')
 radii=sorted([float(r) for r in radii_mm],reverse=True);angles=[float(a)%360 for a in directions_deg]
 if not np.isfinite(radii+angles).all() or min(radii)<=0 or len(set(radii))!=len(radii) or len(set(angles))!=len(angles) or not 0<minimum_correlation<=1:raise ValueError('unique finite radii/directions and valid correlation required')
 if not 1<=n_modes<=8:raise ValueError('bend tracking spectra limited to 8 modes')
 for r,a in itertools.product(radii,angles):configure_bend(design,r,a,**stress_correction)
 mesh=mesh or build_mesh(design);solver=solve_fn or solve
 p,w=disk_probes(.94*design['domain_radius_um'],design['parameters'].get('core_radius_um'))
 def run(d):return solver(d,mesh,n_modes=n_modes,target_neff=target_neff,max_sample_points=25,probe_points=p,probe_weights=w)
 straight=run(design);modes=straight['modes']
 if reference_indices is None:
  tolerance=design.get('solver',{}).get('degeneracy_tolerance',1e-5)
  reference_indices=[j for j,m in enumerate(modes) if abs(complex(m['neff_real'],m['neff_imag'])-complex(modes[0]['neff_real'],modes[0]['neff_imag']))<=tolerance]
 if not 1<=len(reference_indices)<=2 or len(set(reference_indices))!=len(reference_indices):raise ValueError('select 1 mode or a 2D degenerate reference subspace explicitly')
 reference=[modes[j] for j in reference_indices]
 if any(m['status']!='candidate_unverified' for m in reference):raise ValueError('unstable straight reference')
 baseline=float(np.mean([m['alpha_total_dB_m'] for m in reference]));root=_root('bend',design)
 root.update(reference={'indices':reference_indices,'rank':len(reference),'straight_mean_alpha_total_dB_m':baseline,'basis':'tracked subspace; degenerate polarizations not individually labeled'},settings={'radii_mm':radii,'directions_deg':angles,'n_modes':n_modes,'target_neff':target_neff,'minimum_correlation':minimum_correlation,'stress_correction':stress_correction},cases=[])
 for angle in angles:
  previous=reference
  for radius in radii:
   case={'radius_mm':radius,'direction_deg':angle,'mean_alpha_total_dB_m':None,'delta_mean_total_dB_m':None}
   try:
    d=configure_bend(design,radius,angle,**stress_correction);result=run(d)
    good=[(j,m) for j,m in enumerate(result['modes']) if m['status']=='candidate_unverified']
    options=[]
    for group in itertools.combinations(good,len(previous)):
     corr=subspace_overlap(previous,[m for _,m in group]);options.append((min(corr) if corr else 0,group,corr))
    if not options:raise ValueError('insufficient residual-qualified modes')
    quality,group,corr=max(options,key=lambda x:x[0]);case.update(principal_correlations=corr,mode_indices=[j for j,_ in group],bend=d['_bend'])
    if quality<minimum_correlation:case['status']='MODE_TRACKING_FAILED'
    else:
     previous=[m for _,m in group];losses=[m['alpha_total_dB_m'] for m in previous];loss=float(np.mean(losses))
     case.update(status='TRACKED_SUBSPACE_CANDIDATE',mean_alpha_total_dB_m=loss,alpha_total_range_dB_m=[min(losses),max(losses)],
                 delta_mean_total_dB_m=loss-baseline,total_dB_per_turn=bend_per_turn(loss,radius),
                 mean_neff_real=float(np.mean([m['neff_real'] for m in previous])),neff_imag=[m['neff_imag'] for m in previous],
                 loss_resolution='NOT_ESTABLISHED_INDEPENDENT_PML_MESH_STUDY_REQUIRED')
   except (ValueError,RuntimeError) as err:case.update(status='FAILED',error=str(err))
   root['cases'].append(case)
 root['validation']={'loss_certified':False,'V3_COMSOL':'NOT_VALIDATED','limits':['2D weak-curvature equivalent index; not COMSOL axisymmetric bend','delta_total is a total-loss difference, not independently isolated bend radiation','subspace mean can track a degenerate pair; individual polarization loss needs branch identification','R_eff correction has no default empirical coefficient']}
 return root

def pm_analysis(design,config,n_modes=6,target_neff=None,indices=None,mesh=None):
 mesh=mesh or build_mesh(design);stress,delta=solve_thermal_stress(design,mesh,config,max_samples=120)
 root=_root('pm',design);root.update(material_config=config,stress=stress,settings={'n_modes':n_modes,'target_neff':target_neff,'indices':indices},modal_birefringence={'status':'TODO_PHOTOELASTIC_CONSTANTS','B':None,'L_B_m':None})
 optical=None
 if delta is not None:
  p,w=disk_probes(.94*design['domain_radius_um'],design['parameters'].get('core_radius_um'))
  optmesh=dict(mesh,pm_eps_delta=delta,pm_source_sha256=stress['input_sha256'])
  optical=solve(design,optmesh,n_modes=n_modes,target_neff=target_neff,probe_points=p,probe_weights=w,max_sample_points=80)
  root['modal_birefringence']=polarization_pair(optical,indices)
  root['modal_candidates']=[{k:m[k] for k in ('neff_real','neff_imag','status','residual_relative','classification')} for m in optical['modes']]
 root['validation']={'certified':False,'material_basis':'synthetic numerical test only' if config.get('test_only') else 'user-supplied sourced material constants; applicability requires verification',
  'limits':['generalized plane strain isotropic linear elasticity and uniform cooldown','free thermal strain is included; consistent dn/dT reference subtraction is not implemented','no viscoelastic stress relaxation, drawing history, full 3D warping or general xz/yz epsilon','pair spatial overlap screens candidates; independent PM fiber benchmark remains TODO']}
 return root,optical

def mcf_analysis(design,n_modes=None,target_neff=None,mesh=None,bend=None,export_full_quadrature=False):
 if design['family']!='mcf':raise ValueError('MCF design required')
 d=configure_bend(design,**bend) if bend else copy.deepcopy(design)
 par=d['parameters'];count=par['core_count'];centers=np.asarray(mcf_centers(count,par['pitch_um'],d['variant']),float)
 requested=n_modes if n_modes is not None else min(64,2*count)
 if not isinstance(requested,int) or not 1<=requested<=64:raise ValueError('1..64 requested modes')
 p,w=disk_probes(.94*d['domain_radius_um'],None,radial_order=16,angular_points=32)
 # Additional actual FE probes at core centers (positive negligible diagnostic weights).
 points=np.vstack((p,centers));weights=np.r_[w,np.full(count,1e-12)]
 result=solve(d,mesh,n_modes=requested,target_neff=target_neff,max_sample_points=60,export_full_quadrature=True,probe_points=points,probe_weights=weights)
 root=_root('mcf',d);root.update(settings={'n_modes':requested,'target_neff':target_neff,'bend':bend},core_centers_um=centers.tolist(),core_count=count,requested_modes=requested,returned_modes=len(result['modes']),modes=[],coupling_pairs=[])
 root['nearest_core_distance_um']=float(min(np.linalg.norm(a-b) for j,a in enumerate(centers) for b in centers[:j])) if count>1 else None
 for j,m in enumerate(result['modes']):
  pp,ww,f,labels=quadrature(m);sz=.5*np.real(f['Ex']*np.conj(f['Hy'])-f['Ey']*np.conj(f['Hx']));power=sz@ww
  fractions=[float(np.sum(sz[(np.linalg.norm(pp-c,axis=1)<par['core_radius_um']) & (labels=='core')]*ww[(np.linalg.norm(pp-c,axis=1)<par['core_radius_um']) & (labels=='core')])/power) for c in centers]
  _,probe_w,field=unpack(m);core_field=field[-count:,:2]
  coherency=np.einsum('ni,n,nj->ij',field[:,:2].conj(),probe_w,field[:,:2]);eig,jones=np.linalg.eigh(coherency);axis=jones[:,-1]
  purity=float(eig[-1]/sum(eig))
  record={k:m[k] for k in ('neff_real','neff_imag','alpha_total_dB_m','status','residual_relative')};record.update(index=j,polarization=m['classification']['polarization_class'],core_power_fractions=fractions,
   core_Et=[[[float(v.real),float(v.imag)] for v in row] for row in core_field],envelope_parity=None,
   Jones_principal_axis=[[float(v.real),float(v.imag)] for v in axis],polarization_purity=purity)
  if count==2 and not bend and np.linalg.norm(centers[0]+centers[1])<1e-8:
   den=np.linalg.norm(core_field[0])*np.linalg.norm(core_field[1]);corr=np.vdot(core_field[0],core_field[1])/den if den else 0
   balance=min(np.linalg.norm(core_field[0]),np.linalg.norm(core_field[1]))/max(np.linalg.norm(core_field[0]),np.linalg.norm(core_field[1]),1e-30)
   record.update(center_envelope_correlation_real=float(np.real(corr)),center_amplitude_balance=float(balance))
   if balance>.85 and abs(np.imag(corr))<.15:record['envelope_parity']='even' if np.real(corr)>.9 else 'odd' if np.real(corr)<-.9 else None
  root['modes'].append(record)
 if count==2 and not bend:
  even=[m for m in root['modes'] if m['envelope_parity']=='even' and m['status']=='candidate_unverified' and m['polarization_purity']>.95 and sum(m['core_power_fractions'])>.3]
  odd=[m for m in root['modes'] if m['envelope_parity']=='odd' and m['status']=='candidate_unverified' and m['polarization_purity']>.95 and sum(m['core_power_fractions'])>.3]
  matches=[]
  for a,b in itertools.product(even,odd):
   va=np.asarray(a['Jones_principal_axis']);vb=np.asarray(b['Jones_principal_axis']);za=va[:,0]+1j*va[:,1];zb=vb[:,0]+1j*vb[:,1]
   overlap=float(abs(np.vdot(za,zb))**2)
   if overlap>.95:matches.append((a,b,overlap))
  for a,b,overlap in matches:
   if sum(x['index']==a['index'] for x,_,_ in matches)!=1 or sum(y['index']==b['index'] for _,y,_ in matches)!=1:continue
   kappa=(2*np.pi/(d['wavelength_nm']*1e-9))*complex(a['neff_real']-b['neff_real'],a['neff_imag']-b['neff_imag'])/2
   root['coupling_pairs'].append({'polarization':a['polarization'] if a['polarization']==b['polarization'] and a['polarization'] in ('quasi_x','quasi_y') else 'matched_Jones_axis',
    'polarization_axis_overlap':overlap,'even_index':a['index'],'odd_index':b['index'],'kappa_real_per_m':float(kappa.real),'kappa_imag_per_m':float(kappa.imag),
    'Lc_candidate_m':math.pi/(2*abs(kappa.real)) if kappa.real else None,
    'basis':'center-sampled envelope parity and unique principal-Jones overlap candidate; identical symmetric 2-core, lossless coupling-length approximation; near-degenerate polarization axes need convergence'})
 root['validation']={'certified':False,'spectrum_complete':False,'requested_count_returned':len(result['modes'])>=requested,'limits':['requested count is not proof of all N-core guided supermodes; target shifts, per-core powers and mode identification required','kappa only for a unique same-polarization even/odd pair in an unbent origin-symmetric identical 2-core geometry','ring pitch parameter is ring radius; nearest_core_distance_um gives actual center spacing','bend detunes cores; beta splitting is not automatically kappa','optional twist requires helicoidal/longitudinal model; not implemented']}
 if not export_full_quadrature:
  for m in result['modes']:m.pop('field_quadrature',None)
 return root,result
