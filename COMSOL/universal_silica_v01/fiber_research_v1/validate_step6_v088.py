"""Actual STEP6 numerical evidence. Synthetic designs/constants are deliberately labeled.
No PCF V2/V6 or COMSOL PM/bend benchmark pass is inferred from a smoke run.
"""
import argparse,copy,json,time
from pathlib import Path
import numpy as np
from schema import load_presets
from geometry import build_mesh
from step6_analysis_v088 import bend_sweep,pm_analysis,mcf_analysis
from step6_fields_v088 import cartesian_probes,export_grid
from step6_splice_v088 import gaussian_grid,splice
from step6_sweep_v088 import grid_sweep
from step6_pmd_v088 import deterministic_dgd
from pcf_reference_v088 import reference_audit
from vector_fem import solve

def write(dst,name,data):
 (dst/name).write_text(json.dumps(data,indent=2,ensure_ascii=False,allow_nan=False)+'\n')
def small_mcf(count=2,trench=False):
 d=copy.deepcopy(load_presets()['mcf__ring']);d.update(name='Synthetic MCF numerical example; not commercial fiber',domain_radius_um=15)
 d['parameters'].update(core_count=count,core_radius_um=2.2,pitch_um=5.2,trench_width_um=.5)
 d['materials']['core']={'model':'constant','n':1.46};d['boundary']['thickness_um']=2
 d['mesh'].update(elements_per_wavelength=4,bulk_max_um=1.,through_wall_elements=3)
 d['target_neff']=1.45
 if trench:d['variant']='trench_assisted';d['parameters']['core_count']=3
 d['metadata']={'test_only':True,'constant_index_basis':'explicit synthetic optical contrast; no commercial specification or measured material claim'}
 return d

def write_inputs(dst):
 from tests.test_step6b_vector_smoke import small_pm,fake_props
 straight=json.loads((Path(__file__).parent/'benchmarks/step6a_step_index_smoke_input.json').read_text())
 def config(op,**kw):return {'schema':'LSCNS_step6_input_v088','schema_version':'1.0.0','operation':op,**kw}
 write(dst,'bend_input.json',config('bend',design=straight,settings={'radii_mm':[100,20],'directions_deg':[0,90],'n_modes':4,'target_neff':1.4448}))
 write(dst,'pm_input.json',config('pm',design=small_pm(),materials=fake_props(),settings={'n_modes':6,'target_neff':1.446}))
 write(dst,'mcf_input.json',config('mcf',design=small_mcf(),settings={'n_modes':6,'target_neff':1.45}))
 write(dst,'grid_input.json',config('sweep',design=small_mcf(trench=True),axes={'parameters.pitch_um':[6.,7.],'parameters.trench_width_um':[.3,.5]},objectives={'max_total_loss_dB_m':{'direction':'min','unit':'dB/m'},'mean_core_power_fraction':{'direction':'max','unit':'1'}},evaluator={'kind':'mcf','polarization':'auto','settings':{'n_modes':2,'target_neff':1.45}}))
 write(dst,'field_grid_input.json',config('field-grid',design=straight,settings={'half_width_um':5.5,'side':49,'n_modes':2,'target_neff':1.4448}))
 if (dst/'smf_field_grid.json').exists() and (dst/'hcf_field_grid.json').exists():
  write(dst,'splice_input.json',config('splice',field1=json.loads((dst/'smf_field_grid.json').read_text()),field2=json.loads((dst/'hcf_field_grid.json').read_text()),settings={'offset_um':[0,0],'angle_deg':[0,0],'gap_um':0}))

def main():
 p=argparse.ArgumentParser();p.add_argument('--out',default='benchmarks/step6_v088');p.add_argument('--skip-fem',action='store_true');args=p.parse_args();dst=Path(args.out);dst.mkdir(parents=True,exist_ok=True);start=time.monotonic()
 refs=reference_audit();write(dst,'pcf_reference.json',refs)
 a=gaussian_grid(10,side=129);b=gaussian_grid(14,side=129)
 analytic=splice(a,b,offset_um=(1,0),angle_deg=(.2,0),gap_um=20,interface_indices=[1.45,1,1.45]);analytic['example_basis']='analytic Gaussian demonstration; not an HCF mode';write(dst,'splice_gaussian.json',analytic)
 report={'schema':'LSCNS_step6_validation_v088','engine_version':'0.8.8','external_physical_validation':{'V2':refs['V2']['status'],'V6':refs['V6']['status'],'COMSOL_BEND_PM_MCF_SPLICE':'NOT_VALIDATED'},'checks':{},'certified':False}
 if not args.skip_fem:
  straight=json.loads((Path(__file__).parent/'benchmarks/step6a_step_index_smoke_input.json').read_text());mesh=build_mesh(straight)
  bend=bend_sweep(straight,[100,20],[0,90],mesh=mesh,n_modes=4,target_neff=1.4448)
  write(dst,'bend.json',bend);report['checks']['bend']={'cases':len(bend['cases']),'tracked':sum(c['status']=='TRACKED_SUBSPACE_CANDIDATE' for c in bend['cases']),'reference_rank':bend['reference']['rank']}
  print('bend',report['checks']['bend'],flush=True)
  from tests.test_step6b_vector_smoke import small_pm,fake_props
  pm,opt=pm_analysis(small_pm(),fake_props(),n_modes=6,target_neff=1.446)
  write(dst,'pm.json',pm);write(dst,'pm_modes.json',opt)
  report['checks']['pm']={'pair':pm['modal_birefringence'],'max_abs_stress_Pa':pm['stress']['mechanics']['max_abs_stress_Pa'],'basis':'synthetic mechanical/photoelastic constants, not measured silica/boron rods'}
  print('pm',report['checks']['pm']['pair'],flush=True)
  mcf,mo=mcf_analysis(small_mcf(),n_modes=6,target_neff=1.45)
  write(dst,'mcf.json',mcf);write(dst,'mcf_modes.json',mo)
  report['checks']['mcf']={'modes':mcf['returned_modes'],'coupling_pairs':mcf['coupling_pairs']}
  print('mcf',report['checks']['mcf'],flush=True)
  # N=3 actual supermodes, not a two-core interpolation table.
  n3,_=mcf_analysis(small_mcf(3),n_modes=6,target_neff=1.45);write(dst,'mcf_three_core.json',n3)
  report['checks']['three_core']={'modes':n3['returned_modes'],'core_count':n3['core_count']}
  # Actual pitch and trench-width grid (4 FEM evaluations); explicit objectives retain failures.
  td=small_mcf(trench=True)
  def evaluate(d):
   r,_=mcf_analysis(d,n_modes=2,target_neff=1.45)
   good=[m for m in r['modes'] if m['status']=='candidate_unverified']
   if not good:raise ValueError('no stable candidate')
   return {'max_total_loss_dB_m':max(m['alpha_total_dB_m'] for m in good),'mean_core_power_fraction':float(np.mean([sum(m['core_power_fractions']) for m in good]))}
  sweep=grid_sweep(td,{'parameters.pitch_um':[6.,7.],'parameters.trench_width_um':[.3,.5]},
   {'max_total_loss_dB_m':{'direction':'min','unit':'dB/m'},'mean_core_power_fraction':{'direction':'max','unit':'1'}},evaluate,evaluator_id='actual-trench-mcf-v088-two-candidates')
  write(dst,'grid.json',sweep);report['checks']['grid']={'cases':len(sweep['cases']),'computed':sum(c['status']=='COMPUTED' for c in sweep['cases']),'pareto':sweep['pareto_case_indices']}
  print('grid',report['checks']['grid'],flush=True)
  # Et source fields interpolated directly from HCurl FEM, distinct physical designs.
  pts,weights=cartesian_probes(10,33);splice_smf=copy.deepcopy(straight);splice_smf['domain_radius_um']=16;splice_smf['parameters']['core_radius_um']=3.2;splice_smf['materials']['core']['n']=1.46
  smf=solve(splice_smf,n_modes=2,target_neff=1.454,probe_points=pts,probe_weights=weights,max_sample_points=20)
  g1=export_grid(smf,0);write(dst,'smf_field_grid.json',g1)
  hcf=copy.deepcopy(load_presets()['anti_resonant__single_ring']);hcf.update(domain_radius_um=16,wavelength_nm=1550)
  hcf['parameters'].update(core_radius_um=5,tube_radius_um=3,wall_thickness_um=.8,tube_count=6)
  hcf['mesh'].update(elements_per_wavelength=4,bulk_max_um=1,through_wall_elements=3);hcf['boundary']['thickness_um']=2
  hcf['solver']={'max_iterations':500,'tolerance':1e-9}
  hcf['metadata']={'test_only':True,'basis':'small numerical HCF geometry, not certified guided core mode'}
  hopt=solve(hcf,n_modes=2,target_neff=.99,probe_points=pts,probe_weights=weights,max_sample_points=20)
  candidates=[export_grid(hopt,j) for j in range(len(hopt['modes']))];g2=max(candidates,key=lambda g:splice(g1,g)['eta_Et']);write(dst,'hcf_field_grid.json',g2)
  coupling=splice(g1,g2);coupling['example_basis']='FEM SMF/HCF candidate pipeline only: HCF grid edge is not negligible; guided core identity and physical splice loss NOT validated';write(dst,'splice_fem_smf_hcf.json',coupling)
  report['checks']['SMF_HCF_Et_overlap']={'eta':coupling['eta_Et'],'loss_dB':coupling['overlap_loss_dB'],'source_grid_edge_fractions':[g1['edge_peak_intensity_fraction'],g2['edge_peak_intensity_fraction']]}
  # Independent actual stress run with zero cooldown verifies mechanical -> optical path.
  zero=fake_props();zero['delta_T_K']=0.;pzero,_=pm_analysis(small_pm(),zero,n_modes=2,target_neff=1.446);write(dst,'pm_zero_cooldown.json',pzero)
  report['checks']['zero_cooldown_max_stress_Pa']=pzero['stress']['mechanics']['max_abs_stress_Pa']
 if not args.skip_fem:
  assert report['checks']['bend']['tracked']==4
  assert report['checks']['pm']['pair']['B'] is not None
  assert report['checks']['mcf']['coupling_pairs'], 'expected screened two-core coupling candidate'
  assert report['checks']['three_core']['modes']>=6
  assert report['checks']['grid']['computed']==4
  assert report['checks']['zero_cooldown_max_stress_Pa']==0
 write_inputs(dst)
 report['elapsed_seconds']=time.monotonic()-start;write(dst,'validation.json',report);print(json.dumps(report,ensure_ascii=False,indent=2))
if __name__=='__main__':main()
