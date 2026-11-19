"""Independent STEP5 checks and real example generation; sampled gaps only."""
import argparse,copy
from pathlib import Path
import numpy as np
from step5_vector_pwe import *
from step5_reference import reference_te_tm
from step4_v086_cli import write

def run(out):
 out=Path(out);out.mkdir(parents=True,exist_ok=True);uniform_error=0.;residual=0.;divergence=0.
 for lattice in MOTIFS:
  c=example_config();c.update(lattice=lattice,hole_diameter_over_pitch=.2,n_air=1.45,bands=8)
  for beta in (0.,.4,1.2):
   k=np.array([.13,.21]);s=solve_point(c,k,beta);g=reciprocal_indices(c['pwe_cutoff']);B=np.linalg.inv(primitive(lattice,1.)).T
   exact=np.sort(np.repeat(np.sqrt(np.sum(((g+k)@B.T)**2,axis=1)+beta**2)/1.45,2))[:8]
   uniform_error=max(uniform_error,max(abs(np.array(s['frequencies_f'])-exact)));residual=max(residual,s['residual_normwise']);divergence=max(divergence,s['divergence_relative'])
 c=example_config();c.update(lattice='square',hole_diameter_over_pitch=.42,bands=6,pwe_cutoff=6)
 reference_rows=[];p=np.array(solve_point(c,[.13,.21],0)['frequencies_f'])
 for N in (32,64,96):
  ref=reference_te_tm(c,[.13,.21],N);reference_rows.append({'real_space_N':N,'reference_f':ref.tolist(),'PWE_f':p.tolist(),'max_relative_difference':float(max(abs(p-ref)/ref))})
 c=example_config();c['pwe_cutoff']=5
 r=analyze(c);conv=convergence(c,(3,4,5),(3,5,7));write(r,out/'example_bloch.json');write(conv,out/'convergence.json')
 handoffs=[]
 for i,gap in enumerate(r['gap_map']['candidates']):
  d=defect_handoff(r,i,rings=2);write(d,out/f'finite_design_{i}.json');handoffs.append({'candidate':i,'wavelength_nm':d['wavelength_nm'],'target_neff':d['target_neff']})
 finite=None
 if r['gap_map']['candidates']:
  from vector_fem import solve
  from step4_v086 import postprocess,quadrature
  d=defect_handoff(r,0,rings=2,core_surround_um=0.)
  d['mesh'].update(elements_per_wavelength=4,through_wall_elements=3,curvature_segments=48)
  d['metadata']['example_scope']='explicit 2-ring, zero core-surround, low-order numerical illustration; NOT a converged PCF product'
  result=solve(d,n_modes=4,target_neff=d['target_neff'],export_full_quadrature=True,max_sample_points=80)
  metrics=postprocess(result,{'material_roles':{'air':'air','silica':'glass','background':'glass'}})
  core_fractions=[]
  for mode in result['modes']:
   pts,w,fields,labels=quadrature(mode);sz=.5*np.real(fields['Ex']*np.conj(fields['Hy'])-fields['Ey']*np.conj(fields['Hx']))
   core=np.linalg.norm(pts,axis=1)<.87*d['parameters']['pitch_um']
   core_fractions.append(float(np.sum(sz[core]*w[core])/np.sum(sz*w)))
   mode.pop('field_quadrature',None)
  write(d,out/'finite_design.json');write(result,out/'finite_mode_result.json');write(metrics,out/'finite_metrics.json')
  finite={'status':'COMPUTED_CANDIDATES_NOT_LOSS_CERTIFIED','returned_modes':len(result['modes']),'mixed_dofs':result['solver_diagnostics']['mixed_dofs'],
          'neff_real':[m['neff_real'] for m in result['modes']],'residual_max':max(m['residual_relative'] for m in result['modes']),
          'core_disk_Poynting_fractions':core_fractions,'rings':2,'core_surround_um':0.,'mesh_elements_per_wavelength':4}
 report={'schema':'LSCNS_step5_validation_v087','status':'PASS_SCOPED' if uniform_error<=1e-10 and max(residual,divergence)<=1e-10 and reference_rows[-1]['max_relative_difference']<=.02 else 'FAIL',
         'uniform_max_absolute_f_error':float(uniform_error),'max_residual':float(residual),'max_divergence':float(divergence),
         'independent_beta0_TE_TM_reference':reference_rows,'actual_sampled_gap_count':len(r['gap_map']['candidates']),
         'convergence_cases':len(conv['cases']),'handoff_proposals':handoffs,'finite_FEM_example':finite,'complete_gap_certified':False,'MPB_external_comparison':'NOT_RUN',
         'scope':'vector Maxwell mathematics, uniform exact and independent beta0 TE/TM finite-difference cross-check; no experimental PCF/HCF accuracy certification'}
 write(report,out/'validation.json');print(report,flush=True)
 if report['status']!='PASS_SCOPED':raise RuntimeError('STEP5 acceptance failed')
 return report
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('--out',required=True);run(p.parse_args().out)
