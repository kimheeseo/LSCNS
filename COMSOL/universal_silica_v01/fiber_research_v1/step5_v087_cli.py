"""STEP5 vector Maxwell PWE CLI, existing scalar step5_cli.py remains available."""
import argparse,json
from pathlib import Path
from step5_vector_pwe import analyze,convergence,defect_handoff,example_config
from step4_v086_cli import write

def main(argv=None):
 p=argparse.ArgumentParser(description=__doc__);sp=p.add_subparsers(dest='cmd',required=True)
 ex=sp.add_parser('example');ex.add_argument('--out',required=True)
 for name in ('analyze','converge','handoff'):
  q=sp.add_parser(name);q.add_argument('input');q.add_argument('--out',required=True)
  if name=='handoff':
   q.add_argument('--candidate-index',type=int,required=True);q.add_argument('--defect-cells',type=int,choices=(7,19),default=7);q.add_argument('--rings',type=int,default=4);q.add_argument('--core-surround-um',type=float,default=.15);q.add_argument('--domain-radius-um',type=float);q.add_argument('--solve-fem',action='store_true');q.add_argument('--export-full-quadrature',action='store_true')
 a=p.parse_args(argv)
 if a.cmd=='example':write(example_config(),a.out);return
 d=json.loads(Path(a.input).read_text())
 if a.cmd=='analyze':write(analyze(d),a.out)
 elif a.cmd=='converge':write(convergence(d),a.out)
 else:
  design=defect_handoff(d,a.candidate_index,a.rings,a.defect_cells,a.core_surround_um,a.domain_radius_um);write(design,a.out)
  if a.solve_fem:
   from vector_fem import solve
   r=solve(design,n_modes=4,target_neff=design['target_neff'],export_full_quadrature=a.export_full_quadrature)
   write(r,Path(a.out).with_name('mode_result.json'))
   print('actual finite modes computed; existence/localization/PML convergence are NOT gap certification')
if __name__=='__main__':main()
