"""CLI for versioned design, gmsh geometry and experimental mixed vector FEM."""
import argparse
import copy
import json
from pathlib import Path
from datetime import datetime, timezone

try:
    from .schema import load_design, save_design, load_presets, validate_design
except ImportError:
    from schema import load_design, save_design, load_presets, validate_design


def dump(data,path):
    Path(path).write_text(json.dumps(data,indent=2,ensure_ascii=False,allow_nan=False)+"\n",encoding="utf-8")


def run_study(d,out,n_modes=1,target_neff=None):
    try:
        from .geometry import build_mesh
        from .vector_fem import solve
    except ImportError:
        from geometry import build_mesh
        from vector_fem import solve
    if d["boundary"]["type"]!="pml":raise ValueError("PML sensitivity requires boundary.type=pml")
    try:from .modal_study import make_sensitivity_cases,run_cases,sensitivity_verdict,input_digest
    except ImportError:from modal_study import make_sensitivity_cases,run_cases,sensitivity_verdict,input_digest
    cases=make_sensitivity_cases(d)
    rows,_=run_cases(d,cases,out,n_modes=max(2,n_modes),target_neff=target_neff)
    return {"schema":"LSCNS_fiber_research_study_v1","engine_version":"0.8.5","input":d,
            "input_digest":input_digest(d),"cases":rows,**sensitivity_verdict(rows),
            "tracking":"fixed physical-domain FE probes; degenerate-subspace principal correlations >=0.98"}


def main(argv=None):
    parser=argparse.ArgumentParser(description=__doc__)
    sub=parser.add_subparsers(dest="command",required=True)
    q=sub.add_parser("presets",help="write all A-H example JSON designs")
    q.add_argument("--out",default="presets")
    v=sub.add_parser("validate")
    v.add_argument("design")
    for name in ("mesh","solve","study","bend","inspect","mesh-study"):
        u=sub.add_parser(name)
        u.add_argument("design")
        u.add_argument("--out",default="output")
        if name in ("solve","study","bend"):
            u.add_argument("--modes",type=int,default=2 if name in ("solve","bend") else 1)
            u.add_argument("--target-neff",type=float,default=None)
    for name in ("mesh","mesh-study"):
        sub.choices[name].add_argument("--resolution-policy",choices=("boundary_refined","wavelength"),default=None)
    # STEP6a: no radius values are assumed; require explicit engineering inputs.
    bend=sub.choices["bend"]
    bend.add_argument("--radii-mm",type=float,nargs="+",required=True)
    bend.add_argument("--directions-deg",type=float,nargs="+",default=[0.0])
    bend.add_argument("--reff-multiplier",type=float,default=None)
    bend.add_argument("--reff-source-url",default=None)
    bend.add_argument("--reff-source-note",default=None)
    bend.add_argument("--min-overlap",type=float,default=.75)
    # STEP4: retain full physical FEM quadrature only when explicitly requested.
    # This may create very large results, but is required for quantitative overlap matching.
    sub.choices["solve"].add_argument("--export-full-quadrature", action="store_true")
    a=parser.parse_args(argv)
    if a.command=="presets":
        out=Path(a.out)
        for name,preset in load_presets().items():
            save_design(preset,out/(name+".json"))
        print("wrote",len(load_presets()),"designs to",out)
        return
    design=load_design(a.design)
    if a.command=="validate":
        print("VALID",design["schema_version"],design["family"],design["variant"]);return
    if getattr(a,"resolution_policy",None) is not None:
        design["mesh"]["resolution_policy"]=a.resolution_policy
    out=Path(a.out);out.mkdir(parents=True,exist_ok=True)
    save_design(design,out/"input.json")
    if a.command=="inspect":
        try: from .design_checks import design_diagnostics
        except ImportError: from design_checks import design_diagnostics
        report=design_diagnostics(design);dump(report,out/"design_diagnostics.json")
        print(json.dumps(report,ensure_ascii=False,indent=2));return
    if a.command=="mesh-study":
        try: from .mesh_study import run_mesh_study
        except ImportError: from mesh_study import run_mesh_study
        report=run_mesh_study(design,out)
        print(report["status"]);return
    if a.command=="mesh":
        try: from .geometry import build_mesh
        except ImportError: from geometry import build_mesh
        m=build_mesh(design,out)
        print(json.dumps(m["quality"],ensure_ascii=False,indent=2));return
    if a.command=="solve":
        try:
            from .geometry import build_mesh
            from .vector_fem import solve
        except ImportError:
            from geometry import build_mesh
            from vector_fem import solve
        m=build_mesh(design,out)
        s=solve(design,m,n_modes=a.modes,target_neff=a.target_neff,export_full_quadrature=a.export_full_quadrature)
        dump(s,out/"mode_result.json")
        print("EXPERIMENTAL (not loss-validated):",[(x["neff_real"],x["neff_imag"],x["residual_relative"]) for x in s["modes"]])
        return
    if a.command=="bend":
        try: from .bending import bend_radius_direction_sweep
        except ImportError: from bending import bend_radius_direction_sweep
        report=bend_radius_direction_sweep(
            design,a.radii_mm,a.directions_deg,
            reff_multiplier=a.reff_multiplier,
            stress_source_url=a.reff_source_url,
            stress_source_note=a.reff_source_note,
            n_modes=a.modes,target_neff=a.target_neff,
            minimum_overlap=a.min_overlap,output_dir=out)
        dump(report,out/"bend_sweep_result.json")
        print("STEP6a v0.8.1 experimental (NOT V3 validated):",
              sum(x["status"]=="COMPUTED_UNVERIFIED" for x in report["cases"]),
              "candidate curves",out/"bend_sweep_result.json")
        return
    report=run_study(design,out,a.modes,a.target_neff)
    dump(report,out/"study_report.json")
    print(report["status"])


if __name__=="__main__":
    main()

