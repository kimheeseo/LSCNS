"""Independent STEP4 FEM dispersion benchmark; does not certify HCF product loss."""
import argparse,copy,json
from pathlib import Path
import numpy as np
from scipy.special import jnp_zeros
from validate_step3 import benchmark_design
from step4_v086_cli import solve_sweep,write
from step4_v086 import track_sweep,postprocess
from vector_fem import solve

def run(out):
 out=Path(out);out.mkdir(parents=True,exist_ok=True)
 d=benchmark_design('uniform_absorbing');d['solver']['max_dofs']=120000;d['mesh']['curvature_segments']=192
 tracked,metrics=solve_sweep(d,np.arange(1510.,1600.,10.),out/'uniform',2,1.442,{'material_roles':{'core':'glass','background':'glass'}})
 x=np.asarray(tracked['parameter_values']);a=jnp_zeros(1,1)[0]/(2*np.pi*3e-6);n=1.45+1e-4j;ne=np.sqrt(n*n-(a*x*1e-9)**2)
 n2=-a*a*n*n/ne**3;n3=-3*a**4*n*n*(x*1e-9)/ne**5
 exactD=-(x*1e-9)/299792458*n2.real*1e6;exactS=-(n2.real+(x*1e-9)*n3.real)/299792458*1e-3
 errors={'D':max(abs(np.array(t['D_ps_nm_km'])-exactD).max() for t in tracked['curves']),
         'S':max(abs(np.array(t['S_ps_nm2_km'])-exactS).max() for t in tracked['curves'])}
 raw=json.loads((out/'uniform/raw_sweep.json').read_text());coarse=copy.deepcopy(raw);coarse['runs']=raw['runs'][::2];coarse['parameter_values']=raw['parameter_values'][::2];c=track_sweep(coarse)
 step_error=max(abs(np.array(t['D_ps_nm_km'])[::2]-np.array(s['D_ps_nm_km'])).max() for t,s in zip(tracked['curves'],c['curves']))
 leak=benchmark_design('radiative');r=solve(leak,n_modes=2,target_neff=.981265,export_full_quadrature=True,max_sample_points=32)
 hcf=postprocess(r,{'material_roles':{'core':'air','background':'glass'}});write(hcf,out/'radiative_metrics.json')
 for m in hcf['modes']:
  if m['P_glass_fraction'] is None or abs(sum(m['material_power_fractions'].values())-1)>1e-10:raise RuntimeError('material partition failed')
 report={'schema':'LSCNS_step4_validation_v086','status':'PASS_SCOPED' if errors['D']<=.1 and errors['S']<=.01 and step_error<=.1 else 'FAIL',
         'maximum_D_error_ps_nm_km':float(errors['D']),'maximum_S_error_ps_nm2_km':float(errors['S']),
         'coarse_fine_D_change_ps_nm_km':float(step_error),'minimum_overlap':min(min(t['overlap_power']) for t in tracked['curves']),
         'reference':'exact complex circular PEC TE11 sqrt(n²-(jprime11 lambda/(2 pi R))²), R=3um, n=1.45+i1e-4 synthetic',
         'scope':'formula/integration/tracking and circular PEC dispersion; HCF loss/HOMER/anticrossing products NOT certified',
         'HCF_actual_radiative_partition_test':True,'HCF_n2_without_source':hcf['modes'][0]['nonlinearity_status']}
 write(report,out/'validation.json');write(metrics[0],out/'example_metrics.json');write(tracked,out/'example_sweep.json');print(report,flush=True)
 if report['status']!='PASS_SCOPED':raise RuntimeError('STEP4 published acceptance failed')
 return report
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('--out',required=True);run(p.parse_args().out)
