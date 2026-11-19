"""Stable entry: python -m lswave.cli solve design.json --out result.json.
Run from fiber_research_v1 in local Python/Colab; old CLIs remain supported.
"""
import argparse,json
from pathlib import Path
from .reports import verify,stamp

def write(path,data):
 p=Path(path);p.parent.mkdir(parents=True,exist_ok=True);p.write_text(json.dumps(data,ensure_ascii=False,allow_nan=False,indent=2)+'\n')
def main():
 p=argparse.ArgumentParser(description=__doc__);p.add_argument('operation',choices=['solve','verify','bloch','sweep']);p.add_argument('input');p.add_argument('--out');p.add_argument('--modes',type=int,default=2);p.add_argument('--target-neff',type=float);p.add_argument('--retain-quadrature',action='store_true');p.add_argument('--start',type=float);p.add_argument('--stop',type=float);p.add_argument('--points',type=int,default=5);a=p.parse_args();d=json.loads(Path(a.input).read_text())
 if a.operation=='verify':verify(d);print('INPUT_HASH_AND_METADATA_VERIFIED_ONLY_NOT_PHYSICS');return
 if not a.out:p.error('--out required')
 if a.operation=='solve':
  from .solver import modes
  r=modes(d,n_modes=a.modes,target_neff=a.target_neff,retain_quadrature=a.retain_quadrature,max_sample_points=150)
 elif a.operation=='bloch':
  from step5_vector_pwe import analyze
  r=stamp(analyze(d),d,'0.8.7')
 else:
  import numpy as np
  from step4_v086_cli import solve_sweep
  if a.start is None or a.stop is None:p.error('sweep --start and --stop required')
  r,_=solve_sweep(d,np.linspace(a.start,a.stop,a.points),Path(a.out).with_suffix(''),a.modes,a.target_neff)
  r['input']=d;r['input_manifest']={'design':d,'run_options':{'wavelengths_nm':np.linspace(a.start,a.stop,a.points).tolist(),'n_modes':a.modes,'target_neff':a.target_neff}};r=stamp(r,r['input_manifest'],'0.8.6')
 write(a.out,r);print(json.dumps({'output':a.out,'input_sha256':r['input_sha256'],'solver_version':r['solver_version']}))
if __name__=='__main__':main()
