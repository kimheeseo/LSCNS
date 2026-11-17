"""STEP 5 v0.8.3 CLI. Python/Colab numeric solver; web only visualizes saved JSON."""
import argparse
import json
from pathlib import Path
try:
    from .step5_bloch import analyze, defect_handoff, example_config, validate_config
except ImportError:
    from step5_bloch import analyze, defect_handoff, example_config, validate_config

def dump(path, value):
    output=Path(path)
    output.parent.mkdir(parents=True,exist_ok=True)
    output.write_text(json.dumps(value,ensure_ascii=False,indent=2,allow_nan=False)+"\n",encoding="utf-8")
    print(str(output.resolve()))
    return output

def main(argv=None):
    p=argparse.ArgumentParser(description="STEP5 fixed beta scalar 2D Bloch Floquet band structure")
    sp=p.add_subparsers(dest="command",required=True)
    ex=sp.add_parser("example",help="write illustrative input config")
    ex.add_argument("--out",default="step5_design.json")
    run=sp.add_parser("analyze",help="solve eigenfrequency over Bloch k path and reciprocal mesh")
    run.add_argument("design",help="STEP5_INPUT JSON")
    run.add_argument("--out",default="out_step5/step5_bloch_result.json")
    hand=sp.add_parser("handoff",help="generate pbgf finite-section full-vector design from selected candidate")
    hand.add_argument("result",help="STEP5_RESULT JSON")
    hand.add_argument("--candidate-index",type=int,required=True,help="0-based index in gap_map.candidates")
    hand.add_argument("--defect-cells",type=int,choices=(7,19),default=7)
    hand.add_argument("--rings",type=int,default=4)
    hand.add_argument("--core-surround-um",type=float,default=.15)
    hand.add_argument("--out",default="out_step5/finite_pbgf_design.json")
    args=p.parse_args(argv)
    if args.command=="example":
        dump(args.out,example_config())
    elif args.command=="analyze":
        cfg=validate_config(json.loads(Path(args.design).read_text(encoding="utf-8")))
        result=analyze(cfg)
        dump(args.out,result)
        print("Sampled gaps:",len(result["gap_map"]["candidates"]),
              "— CANDIDATES, not complete-gap certification")
    elif args.command=="handoff":
        result=json.loads(Path(args.result).read_text(encoding="utf-8"))
        design=defect_handoff(result,args.candidate_index,args.rings,
                             args.defect_cells,args.core_surround_um)
        dump(args.out,design)
        print("Only conceptual finite full-section FEM handoff. Not validated against COMSOL/PML.")
if __name__=="__main__":
    main()
