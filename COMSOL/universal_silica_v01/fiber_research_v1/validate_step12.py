"""Reproduce STEP 1–2 geometry checks; never calls the vector eigenmode solver."""
import argparse
import importlib.metadata
import json
import platform
from pathlib import Path
from datetime import datetime, timezone

try:
    from .schema import load_presets, validate_design
    from .geometry import build_mesh
except ImportError:
    from schema import load_presets, validate_design
    from geometry import build_mesh


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out",required=True)
    args=parser.parse_args()
    cases=[]
    for name,design in load_presets().items():
        try:
            validate_design(design)
            q=build_mesh(design)["quality"]
            area_errors={k:abs(q["material_area_um2"][k]-v)/v for k,v in q["cad_material_area_um2"].items() if v>0}
            walls=q["wall_resolution"]
            cases.append({"preset":name,"status":"GENERATED","triangles":q["triangles"],
                          "materials":q["materials"],"area_relative_errors":area_errors,
                          "max_area_relative_error":max(area_errors.values()),
                          "edge_min_um":q["edge_min_um"],"edge_max_um":q["edge_max_um"],
                          "aspect_ratio_max":q["aspect_ratio_max"],
                          "boundary_labels":sorted(q["boundary_groups"]),
                          "wall_status":walls["status"],"wall_count":len(walls["walls"]),
                          "wall_min_crossed_triangles":min((w["min_crossed_triangles"] for w in walls["walls"]),default=None),
                          "wall_min_coverage":min((w["min_covered_fraction"] for w in walls["walls"]),default=None),
                          "wavelength_resolution":q["wavelength_resolution"]["status"],
                          "geometry_fidelity":q["design_diagnostics"]["geometry_fidelity"]})
            print(name,q["triangles"],walls["status"],flush=True)
        except Exception as exc:
            cases.append({"preset":name,"status":"FAILED","error":str(exc)})
            print(name,"FAILED",str(exc),flush=True)
    generated=[c for c in cases if c["status"]=="GENERATED"]
    result={"schema":"LSCNS_step12_validation_v084","timestamp_utc":datetime.now(timezone.utc).isoformat(),
            "environment":{"python":platform.python_version(),**{p:importlib.metadata.version(p) for p in ("numpy","scipy","gmsh","scikit-fem","jsonschema")}},
            "cases":cases,"generated_count":len(generated),"expected_count":len(cases),
            "linear_triangle_area_error_policy":.02,
            "all_geometry_checks_passed":len(generated)==len(cases) and all(c["max_area_relative_error"]<.02 for c in generated),
            "sampled_walls_all_passed":all(c["wall_status"]!="FAIL_SAMPLED" for c in generated) and len(generated)==len(cases),
            "optical_accuracy_certified":False,
            "note":"Presets intentionally retain compatibility boundary_refined policy; wavelength FAIL_POLICY is exposed, not hidden. Named approximate CAD remains approximate."}
    out=Path(args.out);out.parent.mkdir(parents=True,exist_ok=True)
    out.write_text(json.dumps(result,indent=2,allow_nan=False)+"\n")
    if not result["all_geometry_checks_passed"] or not result["sampled_walls_all_passed"]:
        raise SystemExit(1)


if __name__=="__main__":main()
