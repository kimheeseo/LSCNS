"""Field/subspace-tracked mesh/window/PML studies for STEP 3 v0.8.5."""
import copy
import hashlib
import itertools
import json
from pathlib import Path
import numpy as np

try:
    from .modal_diagnostics import disk_probes,weighted_columns,principal_correlations
    from .vector_fem import solve
except ImportError:
    from modal_diagnostics import disk_probes,weighted_columns,principal_correlations
    from vector_fem import solve


def input_digest(design):
    return hashlib.sha256(json.dumps(design,sort_keys=True,separators=(",",":"),allow_nan=False).encode()).hexdigest()


def tracking_columns(result,indices):
    records=[result["modes"][i]["tracking_fields"] for i in indices]
    points=np.asarray(records[0]["points_um"]);weights=np.asarray(records[0]["weights_um2"])
    fields=[]
    for record in records:
        if not np.array_equal(points,record["points_um"]) or not np.array_equal(weights,record["weights_um2"]):raise ValueError("identical physical probes/weights required")
        raw=np.asarray(record["E"],dtype=float);fields.append(raw[...,0]+1j*raw[...,1])
    return weighted_columns(np.stack(fields,axis=-1),weights)


def match_subspace(reference,candidate,dimension=2):
    if len(reference["modes"])<dimension or len(candidate["modes"])<dimension:
        return {"status":"FAIL_MISSING_MODES","minimum_correlation":0.,"indices":[]}
    ref=tracking_columns(reference,range(dimension));best=None
    for subset in itertools.combinations(range(len(candidate["modes"])),dimension):
        correlation=principal_correlations(ref,tracking_columns(candidate,subset))
        score=min(correlation,default=0.)
        if best is None or score>best[0]:best=(score,subset,correlation)
    return {"status":"PASS_TRACKING" if best[0]>=.98 else "FAIL_TRACKING",
            "minimum_correlation":best[0],"indices":list(best[1]),"principal_correlations":best[2]}


def summarize(result,indices):
    modes=[result["modes"][i] for i in indices]
    return {"neff_real":[m["neff_real"] for m in modes],"neff_imag":[m["neff_imag"] for m in modes],
            "max_qep_residual":max(m["residual_relative"] for m in modes),
            "max_pencil_residual":max(m["linearized_residual_relative"] for m in modes),
            "max_gauss_residual":max(m["weak_gauss_relative"] for m in modes),
            "max_pml_intensity_fraction":max(m["classification"]["PML_intensity_fraction"] for m in modes),
            "all_finite":all(m["classification"]["finite_six_components"] for m in modes),
            "power_W":[m["classification"]["power_W"] for m in modes],
            "solver":result["solver_diagnostics"],"triangles":result["mesh_quality"]["triangles"]}


def run_cases(base,cases,output_dir,n_modes=2,target_neff=None):
    out=Path(output_dir);out.mkdir(parents=True,exist_ok=True)
    # All studies use the smallest physical window, not a different quadrature per mesh.
    radius=.94*min(d["domain_radius_um"] for _,d in cases)
    points,weights=disk_probes(radius,base["parameters"].get("core_radius_um"),radial_order=16)
    rows=[];reference=None;first_result=None
    for name,design in cases:
        print(name,flush=True)
        try:
            result=solve(design,n_modes=max(2,n_modes),target_neff=target_neff,max_sample_points=64,
                         probe_points=points,probe_weights=weights)
            if reference is None:
                reference=result;first_result=result
            match=match_subspace(reference,result,dimension=min(2,n_modes))
            if not match["indices"]:raise RuntimeError(match["status"])
            row={"name":name,"state":"COMPUTED","input_digest":input_digest(design),
                 "tracking":dict(match,tracked_dimension=min(2,n_modes)),**summarize(result,match["indices"])}
        except Exception as exc:row={"name":name,"state":"FAILED","error":str(exc)}
        rows.append(row)
        (out/"progress.json").write_text(json.dumps({"cases":rows},indent=2,allow_nan=False)+"\n")
    return rows,first_result


def make_sensitivity_cases(design):
    cases=[("baseline",copy.deepcopy(design))]
    for count,curvature in zip((4,6,8),(96,144,192)):
        d=copy.deepcopy(design);d["mesh"].update(elements_per_wavelength=count,curvature_segments=curvature)
        cases.append(("mesh_"+str(count),d))
    for factor in (1.,1.15,1.3):
        d=copy.deepcopy(design);d["domain_radius_um"]*=factor
        cases.append(("window_"+str(factor),d))
    if design["boundary"]["type"]=="pml":
        for field,values in (("thickness_um",(1.5,2.,2.5)),("strength",(3.,4.,5.)),("order",(2,3,4))):
            for value in values:
                d=copy.deepcopy(design);d["boundary"][field]=value;cases.append(("pml_"+field+"_"+str(value),d))
    return cases


def sensitivity_verdict(rows,imaginary_floor=1e-10):
    complete=all(r["state"]=="COMPUTED" and r["tracking"]["status"]=="PASS_TRACKING" for r in rows)
    good=[r for r in rows if r["state"]=="COMPUTED"]
    if not complete:return {"status":"FAIL_INCOMPLETE_OR_TRACKING","loss_certified":False}
    re=np.asarray([sorted(r["neff_real"]) for r in good]);im=np.asarray([value for r in good for value in r["neff_imag"]])
    real_spread=float(np.max(np.ptp(re,axis=0)))
    noise=max(abs(im));positive=bool(np.min(im)>imaginary_floor)
    spread=float(np.ptp(im)/max(abs(np.mean(im)),imaginary_floor))
    stable=positive and spread<=.05
    return {"status":"PASS_SCOPED_SENSITIVITY" if real_spread<=1e-5 and stable else
                     "UNRESOLVED_NEAR_ZERO_LOSS" if noise<=imaginary_floor else "FAIL_SENSITIVITY",
            "Re_spread":real_spread,"Im_fractional_spread":spread,"minimum_Im":float(im.min()),
            "maximum_abs_Im":float(noise),"imaginary_noise_floor":imaginary_floor,
            "loss_certified":False,"note":"stability alone is not an independent leakage reference validation"}
