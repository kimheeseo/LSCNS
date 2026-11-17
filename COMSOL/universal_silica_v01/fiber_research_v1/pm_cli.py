"""v0.8.2 STEP6b CLI: sourced thermal stress -> photoelastic index -> vector FEM.

python pm_cli.py solve <design.json> <pm_materials.json> --out out_pm --modes 4
The template fields are all TODO; no implicit E/nu/alpha/p11/p12 defaults.
A material-average refractive tensor is not a certified *modal* birefringence.
"""
import argparse
import json
import math
from pathlib import Path
import numpy as np
try:
    from .schema import load_design
    from .geometry import build_mesh
    from .stress_optic import solve_thermal_stress, ENGINE_VERSION
    from .vector_fem import solve as solve_optical
except ImportError:
    from schema import load_design
    from geometry import build_mesh
    from stress_optic import solve_thermal_stress, ENGINE_VERSION
    from vector_fem import solve as solve_optical


def derive_modal_birefringence(modes,wavelength_nm):
    """Only label *candidate* modal neff splitting when both orthogonal modes identified."""
    x=[m for m in modes if m["classification"]["polarization_class"]=="quasi_x"
       and m["status"]=="candidate_unverified"]
    y=[m for m in modes if m["classification"]["polarization_class"]=="quasi_y"
       and m["status"]=="candidate_unverified"]
    if not x or not y:
        return {"status":"TODO_POLARIZATION_MODE_ID",
                "needed":"converged, eigenvector-overlap-validated quasi-x and quasi-y fundamental modes in anisotropic Maxwell FEM",
                "B_modal_candidate":None,"beat_length_candidate_m":None}
    # Candidate pair only, not guaranteed HE11 fundamental pair.
    mx=min(x,key=lambda m:m["residual_relative"])
    my=min(y,key=lambda m:m["residual_relative"])
    b=abs(float(mx["neff_real"])-float(my["neff_real"]))
    return {"status":"UNVERIFIED_VECTOR_MODE_PAIR","B_modal_candidate":b,
            "beat_length_candidate_m":(float(wavelength_nm)*1e-9/b if b>0 else None),
            "neff_quasi_x":float(mx["neff_real"]),"neff_quasi_y":float(my["neff_real"]),
            "limitation":"provisional x/y labels; not independently benchmarked for degenerate modes, mesh/PML or PM fiber n_eff"}


def write_json(data,path):
    p=Path(path)
    p.write_text(json.dumps(data,ensure_ascii=False,indent=2,allow_nan=False)+"\n",encoding="utf-8")


def run_pm(design,config,mesh=None,n_modes=4,target_neff=None,out=None,run_optical=True):
    if mesh is None:
        mesh=build_mesh(design,Path(out)/"mesh" if out else None)
    stress,delta=solve_thermal_stress(design,mesh,config)
    stress["validation"]["V3_COMSOL_STRESS_OPTICAL"]="EVIDENCE_REQUIRED"
    stress["validation"]["fiber_production_qualification"]="NOT_CERTIFIED"
    output={"stress":stress,"optical":None,
            "modal_birefringence":{"status":"TODO_OPTICAL_SOLVE",
                                    "needed":"complete p11/p12 source values and full vector polarization-resolved mode solve"}}
    if run_optical and delta is not None:
        # Preserve the source geometry and mechanical materials; solver sees only the
        # relative anisotropic dielectric tensor per triangle.
        optmesh=dict(mesh)
        optmesh["pm_eps_delta"]=delta
        optmesh["pm_source_sha256"]=stress["input_sha256"]
        modes=solve_optical(design,optmesh,n_modes=n_modes,target_neff=target_neff,
                            max_sample_points=120)
        modes["input_sha256"]=stress["input_sha256"]
        modes["timestamp_utc"]=stress["timestamp_utc"]
        modes["validation"]["pm_modal_birefringence_certified"]=False
        output["optical"]=modes
        output["modal_birefringence"]=derive_modal_birefringence(modes["modes"],design["wavelength_nm"])
    if out is not None:
        dst=Path(out);dst.mkdir(parents=True,exist_ok=True)
        write_json(stress,dst/"pm_stress_result.json")
        if delta is not None:
            np.savez_compressed(dst/"pm_tensor.npz",
                points_um=np.asarray(mesh["points_um"]),
                triangles=np.asarray(mesh["triangles"]),
                material_ids=np.asarray(mesh["material_ids"]),
                epsilon_tensor_delta=np.asarray(delta),
                input_sha256=np.asarray(stress["input_sha256"]))
        if output["optical"] is not None:
            write_json(output["optical"],dst/"pm_mode_result.json")
        write_json({"schema":"LSCNS_step6b_summary_v082","engine_version":ENGINE_VERSION,
                    "input_sha256":stress["input_sha256"],
                    "created_utc":stress["timestamp_utc"],
                    "mechanics":stress["mechanics"],
                    "photoelastic":stress["photoelastic"],
                    "modal_birefringence":output["modal_birefringence"],
                    "validation":stress["validation"],
                    "source_design":design,"material_config":config},dst/"pm_summary.json")
    return output


def main(argv=None):
    p=argparse.ArgumentParser(description=__doc__)
    s=p.add_subparsers(dest="cmd",required=True)
    x=s.add_parser("solve",help="generalized plane strain + optional full vector Maxwell")
    x.add_argument("design")
    x.add_argument("materials")
    x.add_argument("--out",default="pm_result")
    x.add_argument("--modes",type=int,default=4)
    x.add_argument("--target-neff",type=float,default=None)
    x.add_argument("--mechanics-only",action="store_true",
                   help="skip Maxwell eigensolve; stress tensor and material-B proxy only")
    a=p.parse_args(argv)
    design=load_design(a.design)
    config=json.loads(Path(a.materials).read_text(encoding="utf-8"))
    result=run_pm(design,config,n_modes=a.modes,target_neff=a.target_neff,
                  out=a.out,run_optical=not a.mechanics_only)
    print(json.dumps({"engine_version":ENGINE_VERSION,
                      "stress":result["stress"]["photoelastic"]["core_material_birefringence"]["status"],
                      "modal":result["modal_birefringence"]["status"],
                      "validation":"NOT_CERTIFIED",
                      "out":a.out},ensure_ascii=False,indent=2))

if __name__=="__main__":main()
