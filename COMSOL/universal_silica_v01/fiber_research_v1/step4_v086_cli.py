"""STEP4 v0.8.6 actual FE integration and common-probe sweep CLI."""
import argparse,copy,json
from pathlib import Path
import numpy as np
try:
 from .step4_v086 import postprocess,track_sweep,detect_surface_pair
 from .schema import validate_design
except ImportError:
 from step4_v086 import postprocess,track_sweep,detect_surface_pair
 from schema import validate_design

def write(d,path):
 p=Path(path);p.parent.mkdir(parents=True,exist_ok=True);p.write_text(json.dumps(d,indent=2,ensure_ascii=False,allow_nan=False)+'\n');return p

def solve_sweep(design,values,out,n_modes=2,target=None,context=None,parameter='wavelength_nm'):
 try:
  from .geometry import build_mesh
  from .vector_fem import solve
  from .modal_diagnostics import disk_probes
 except ImportError:
  from geometry import build_mesh
  from vector_fem import solve
  from modal_diagnostics import disk_probes
 values=np.asarray(values,float)
 if values.ndim!=1 or not 2<=len(values)<=101 or not np.isfinite(values).all() or np.any(np.diff(values)<=0):raise ValueError('2..101 increasing finite sweep values')
 if parameter=='wavelength_nm' and (len(values)<5 or min(values)<250 or max(values)>3500):raise ValueError('wavelength sweep >=5, 250..3500 nm')
 out=Path(out);out.mkdir(parents=True,exist_ok=True);cases=[]
 for value in values:
  d=copy.deepcopy(design)
  if parameter=='wavelength_nm':d['wavelength_nm']=float(value)
  elif parameter.startswith('parameters.'):d['parameters'][parameter.split('.',1)[1]]=float(value)
  else:raise ValueError('unsupported parameter')
  validate_design(d);cases.append(d)
 p,w=disk_probes(.94*min(d['domain_radius_um'] for d in cases),design['parameters'].get('core_radius_um'))
 common=build_mesh(cases[0],out/'common_mesh') if parameter=='wavelength_nm' else None
 runs=[];metrics=[]
 for i,d in enumerate(cases):
  mesh=common if common is not None else build_mesh(d,out/f'mesh_{i}')
  r=solve(d,mesh,n_modes=n_modes,target_neff=target,probe_points=p,probe_weights=w,max_sample_points=64,export_full_quadrature=True)
  m=postprocess(r,context);write(m,out/f'metrics_{i}.json');metrics.append(m)
  for mode in r['modes']:mode.pop('field_quadrature',None)
  runs.append(r);write({'schema':'LSCNS_step4_raw_sweep_v086','parameter':parameter,'parameter_values':values[:i+1].tolist(),'runs':runs},out/'partial_raw_sweep.json')
  print('STEP4',i+1,len(values),parameter,float(values[i]),flush=True)
 raw={'schema':'LSCNS_step4_raw_sweep_v086','parameter':parameter,'parameter_values':values.tolist(),'runs':runs}
 tracked=track_sweep(raw)
 for curve in tracked['curves']:
  curve['P_glass_fraction']=[float(np.mean([metrics[k]['modes'][j]['P_glass_fraction'] for j in group])) if all(metrics[k]['modes'][j]['P_glass_fraction'] is not None for j in group) else None for k,group in enumerate(curve['members'])]
 ctx=context or {}
 if ctx.get('surface_pair') is not None:
  a,b=ctx['surface_pair'];tracked['surface_anticrossing']=detect_surface_pair(tracked,a,b,[c['P_glass_fraction'] for c in tracked['curves']],ctx.get('surface_identity_source'))
 write(raw,out/'raw_sweep.json');write(tracked,out/'tracked_sweep.json');return tracked,metrics

def main(argv=None):
 p=argparse.ArgumentParser(description=__doc__);sp=p.add_subparsers(dest='cmd',required=True)
 for name in ('result','track','solve-sweep'):
  a=sp.add_parser(name);a.add_argument('input');a.add_argument('--out',required=True);a.add_argument('--context')
  if name=='solve-sweep':
   a.add_argument('--start',type=float,required=True);a.add_argument('--stop',type=float,required=True);a.add_argument('--points',type=int,default=9);a.add_argument('--parameter',default='wavelength_nm');a.add_argument('--modes',type=int,default=2);a.add_argument('--target-neff',type=float)
 a=p.parse_args(argv);d=json.loads(Path(a.input).read_text());ctx=json.loads(Path(a.context).read_text()) if a.context else {}
 if a.cmd=='solve-sweep':solve_sweep(d,np.linspace(a.start,a.stop,a.points),a.out,a.modes,a.target_neff,ctx,a.parameter)
 else:write(postprocess(d,ctx) if a.cmd=='result' else track_sweep(d),a.out)
if __name__=='__main__':main()
