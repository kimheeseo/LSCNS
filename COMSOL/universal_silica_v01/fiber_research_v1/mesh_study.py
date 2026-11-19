"""Three-level geometry/mesh refinement. Does NOT execute a mode solver."""
import copy
import json
import math
from pathlib import Path

try:
    from .schema import validate_design, save_design
    from .geometry import build_mesh
except ImportError:
    from schema import validate_design, save_design
    from geometry import build_mesh


def run_mesh_study(design, output_dir, factors=(1., 1.3, 1.6)):
    validate_design(design)
    if len(factors)<3 or any(isinstance(x,bool) or not isinstance(x,(int,float)) or not math.isfinite(x) or x<=0 for x in factors) or len(set(factors))!=len(factors):
        raise ValueError("at least three distinct positive refinement factors required")
    out=Path(output_dir);out.mkdir(parents=True,exist_ok=True)
    cases=[]
    for i,factor in enumerate(sorted(factors)):
        d=copy.deepcopy(design)
        mesh=d["mesh"]
        mesh["bulk_max_um"]=design["mesh"]["bulk_max_um"]/factor
        for key,limit in (("through_wall_elements",30),("elements_per_wavelength",50),("curvature_segments",256)):
            mesh[key]=int(math.ceil(design["mesh"][key]*factor))
            if mesh[key]>limit:
                raise ValueError("refinement exceeds schema limit for "+key)
        path=out/("level_"+str(i))
        save_design(d,path/"input.json")
        try:
            raw=build_mesh(d,path)
            q=raw["quality"]
            cases.append({"factor":factor,"status":"COMPUTED","quality":q,
                          "area_relative_errors":{k:abs(q["material_area_um2"][k]-v)/v
                                                  for k,v in q["cad_material_area_um2"].items() if v>0}})
        except Exception as exc:
            cases.append({"factor":factor,"status":"FAILED","error":str(exc)})
    successful=[x for x in cases if x["status"]=="COMPUTED"]
    counts=[x["quality"]["triangles"] for x in successful]
    distinct=len(set(counts))==len(cases) and len(successful)==len(cases)
    result={"schema":"LSCNS_mesh_study_v084","engine_version":"0.8.4","input":design,"cases":cases,
            "distinct_meshes":distinct,
            "status":"MESH_REFINEMENT_COMPUTED" if distinct else "FAIL_INCOMPLETE_OR_IDENTICAL_MESHES",
            "optical_convergence":False,"pml_loss_converged":False,
            "note":"CAD/linear-triangle area discrepancy and sampled wall resolution only; no neff/Im(neff) computed"}
    (out/"mesh_study.json").write_text(json.dumps(result,indent=2,allow_nan=False)+"\n")
    return result
