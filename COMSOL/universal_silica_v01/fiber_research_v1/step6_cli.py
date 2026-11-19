"""v0.8.8 STEP6 local/Colab runner. See STEP6_V088.md for schemas and examples."""
import argparse,json
from pathlib import Path
try:
 from .step6_analysis_v088 import bend_sweep,pm_analysis,mcf_analysis
 from .step6_fields_v088 import cartesian_probes,export_grid
 from .step6_splice_v088 import splice
 from .step6_sweep_v088 import grid_sweep
 from .step6_pmd_v088 import deterministic_dgd
 from .pcf_reference_v088 import reference_audit
 from .vector_fem import solve
except ImportError:
 from step6_analysis_v088 import bend_sweep,pm_analysis,mcf_analysis
 from step6_fields_v088 import cartesian_probes,export_grid
 from step6_splice_v088 import splice
 from step6_sweep_v088 import grid_sweep
 from step6_pmd_v088 import deterministic_dgd
 from pcf_reference_v088 import reference_audit
 from vector_fem import solve

def read(path):return json.loads(Path(path).read_text())
def write(path,value):
 p=Path(path);p.parent.mkdir(parents=True,exist_ok=True);p.write_text(json.dumps(value,ensure_ascii=False,indent=2,allow_nan=False)+'\n')
def main():
 p=argparse.ArgumentParser(description=__doc__);p.add_argument('operation',choices=['bend','pm','mcf','field-grid','splice','sweep','dgd','pcf-reference']);p.add_argument('config',nargs='?');p.add_argument('--out',required=True);a=p.parse_args()
 c=read(a.config) if a.config else {}
 if c.get('schema_version','1.0.0')!='1.0.0':raise ValueError('unsupported STEP6 schema_version')
 c={k:v for k,v in c.items() if k not in ('schema','schema_version','operation')}
 d=c.get('design');settings=c.get('settings',{})
 if a.operation=='bend':r=bend_sweep(d,**settings)
 elif a.operation=='pm':r,opt=pm_analysis(d,c['materials'],**settings);write(str(Path(a.out).with_suffix(''))+'_modes.json',opt)
 elif a.operation=='mcf':r,opt=mcf_analysis(d,**settings);write(str(Path(a.out).with_suffix(''))+'_modes.json',opt)
 elif a.operation=='field-grid':
  half=settings.pop('half_width_um',d['domain_radius_um']*.65);side=settings.pop('side',65);idx=settings.pop('mode_index',0)
  points,weights=cartesian_probes(half,side);opt=solve(d,probe_points=points,probe_weights=weights,**settings);r=export_grid(opt,idx)
 elif a.operation=='splice':r=splice(c['field1'],c['field2'],**settings)
 elif a.operation=='dgd':r=deterministic_dgd(**c)
 elif a.operation=='pcf-reference':r=reference_audit()
 else:
  task=c['evaluator'];kind=task['kind']
  if kind not in ('mcf','bend'):raise ValueError('sweep CLI evaluator must be mcf or bend')
  def evaluate(design):
   if kind=='mcf':
    result,_=mcf_analysis(design,**task.get('settings',{}));pairs=result['coupling_pairs']
    pol=task.get('polarization','auto')
    chosen=pairs[0] if pol=='auto' and len(pairs)==1 else next((v for v in pairs if v['polarization']==pol),None)
    good=[m for m in result['modes'] if m['status']=='candidate_unverified']
    if not good:raise ValueError('no residual-qualified stable MCF candidates')
    metrics={'max_total_loss_dB_m':max(m['alpha_total_dB_m'] for m in good),
             'mean_core_power_fraction':sum(sum(m['core_power_fractions']) for m in good)/len(good),
             'nearest_core_distance_um':result['nearest_core_distance_um']}
    if chosen:metrics['abs_kappa_per_m']=abs(chosen['kappa_real_per_m'])
    return metrics
   result=bend_sweep(design,**task['settings']);cases=result['cases']
   if any(v['status']!='TRACKED_SUBSPACE_CANDIDATE' for v in cases):raise ValueError('bend tracking failed')
   return {'max_total_loss_dB_m':max(v['mean_alpha_total_dB_m'] for v in cases)}
  r=grid_sweep(d,c['axes'],c['objectives'],evaluate,constraints=c.get('constraints'),checkpoint=str(Path(a.out).with_suffix('.checkpoint.json')),evaluator_id=json.dumps(task,sort_keys=True)+':step6-v088-vector086')
 write(a.out,r);print(json.dumps({'operation':a.operation,'out':a.out,'schema':r['schema']},ensure_ascii=False))
if __name__=='__main__':main()
