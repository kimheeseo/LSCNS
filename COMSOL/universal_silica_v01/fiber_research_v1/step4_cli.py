"""v0.8.0 STEP4 optical post-processing CLI for *existing* Python FEM JSON.

Usage: python step4_cli.py result mode_result.json --out step4_result.json --n2-m2-W VALUE
       python step4_cli.py sweep wavelength_sweep.json --out step4_dispersion.json
No FEM computation is run here; PML loss is NOT validated by post-processing.
"""
import argparse
import json
import math
from pathlib import Path
import numpy as np
try:
    from .physical_metrics import (ENGINE_VERSION,confinement_loss,mode_neff_spacing,
        nonlinear_gamma,arrow_resonance_lines,overlay_arrow_with_fem,
        match_modes_by_overlap,dispersion_from_tracked_sweep)
    from .materials import index,provenance
except ImportError:
    from physical_metrics import (ENGINE_VERSION,confinement_loss,mode_neff_spacing,
        nonlinear_gamma,arrow_resonance_lines,overlay_arrow_with_fem,
        match_modes_by_overlap,dispersion_from_tracked_sweep)
    from materials import index,provenance


def _write(obj,out):
    path=Path(out)
    path.parent.mkdir(parents=True,exist_ok=True)
    path.write_text(json.dumps(obj,ensure_ascii=False,indent=2,allow_nan=False)+"\n",encoding="utf-8")
    return path


def postprocess_result(data,n2_m2_W=None):
    if data.get("schema")!="LSCNS_fiber_research_vector_v1" or not data.get("modes"):
        raise ValueError("STEP4 requires experimental Python vector FEM result JSON")
    d=data["input"]
    lam=float(d["wavelength_nm"])
    modes=[]
    for i,m in enumerate(data["modes"]):
        loss=confinement_loss(m["neff_imag"],lam)
        info=m.get("classification",{})
        ae=info.get("Aeff_um2")
        fm={"mode_index":i+1,"neff_real":m["neff_real"],"neff_imag":m["neff_imag"],
            "alpha_total_dB_m":loss["alpha_total_dB_m"],
            "alpha_total_dB_km":loss["alpha_total_dB_km"],
            "alpha_validation":"UNVERIFIED_FEM_PML; material absorption may be included",
            "Aeff_Et_um2":ae if info.get("Aeff_basis","").startswith("Et=") else None,
            "MFD_second_moment_Et_um":info.get("MFD_2nd_um") if info.get("Aeff_basis","").startswith("Et=") else None,
            "MFD_PetermannII_um":None,
            "MFD_PetermannII_needed":"axisymmetric radial near-field or rigorous far-field measurement and convergence",
            "gamma_W_inv_m_inv":None,
            "gamma_needed":"n2 at wavelength from a cited material source; Et-only Aeff",
            "P_glass_fraction":None,
            "P_glass_needed":"untruncated Poynting E/H quadrature and material-labelled glass mask",
            "HOMER_ratio":None,
            "HOMER_needed":"verified FM/HOM identities and independently PML-converged losses",
            "surface_anticrossing":"TODO: overlap-tracked multi-mode wavelength sweep + glass localization"}
        if n2_m2_W is not None and fm["Aeff_Et_um2"] is not None and
           d["family"] not in ("anti_resonant","pbgf","kagome"):
            fm["gamma_W_inv_m_inv"]=nonlinear_gamma(n2_m2_W,lam,fm["Aeff_Et_um2"])["gamma_W_inv_m_inv"]
            fm["gamma_needed"]=None
        modes.append(fm)
    is_hcf=d["family"] in ("anti_resonant","pbgf","kagome")
    arrow=None
    if is_hcf:
        p=d["parameters"]
        if "wall_thickness_um" in p and "silica" in d["materials"] and "air" in d["materials"]:
            ng=index(d["materials"]["silica"],lam).real
            na=index(d["materials"]["air"],lam).real
            arrow=overlay_arrow_with_fem(arrow_resonance_lines(
                p["wall_thickness_um"],ng,na))
            arrow["provenance"]={"n_glass":provenance(d["materials"]["silica"]),
                                 "n_air":provenance(d["materials"]["air"]),
                                 "fixed_reference_lambda_nm":lam}
            arrow["status"]="approximate fixed-index ARROW; no validated FEM wavelength curves"
        else:
            arrow={"status":"TODO_WALL_OR_MATERIAL","needed":"wall_thickness_um and sourced glass/air indices"}
    return {"schema":"LSCNS_step4_metrics_v080","engine_version":ENGINE_VERSION,
            "wavelength_nm":lam,"family":d["family"],"modes":modes,
            "mode_neff_spacing":mode_neff_spacing([m["neff_real"] for m in data["modes"]]) if len(data["modes"])>1 else [],
            "arrow_overlay":arrow,
            "dispersion":{"status":"TODO_TRACKED_SWEEP","needed":"at least five wavelength samples; overlap-tracked same physical mode"},
            "birefringence":{"status":"TODO_POLARIZATION_ID","needed":"field-validated polarization x/y mode pair"},
            "validation":{"formulas_unit_tests":"independent; see CI",
                          "actual_fem_confinement_loss":"UNVERIFIED",
                          "unavailable_quantity_policy":"null with TODO/needed, never guessed"}}


def _mode_for_overlap(mode):
    raw=mode.get("field_quadrature")
    if raw is None:
        raise ValueError("full field_quadrature absent. Re-run FEM with export_full_quadrature=True")
    arrays={}
    for k in ("Ex","Ey"):
        reim=np.asarray(raw[k],dtype=float)
        if reim.ndim!=2 or reim.shape[1]!=2:
            raise ValueError("invalid full complex quadrature sample")
        arrays[k]=reim[:,0]+1j*reim[:,1]
    return {"points_um":raw["points_um"],"weights_um2":raw["weights_um2"],**arrays}


def track_wavelength_sweep(data,min_overlap=.75):
    """Refuse to evaluate D if any adjacent vector-field overlap is ambiguous."""
    if data.get("schema")!="LSCNS_step4_raw_sweep_v1":
        raise ValueError("expected LSCNS_step4_raw_sweep_v1")
    runs=data.get("runs")
    if not isinstance(runs,list) or len(runs)<5:
        raise ValueError("at least five consecutive FEM results required")
    lam=np.array([r["input"]["wavelength_nm"] for r in runs],float)
    if not np.isfinite(lam).all() or np.any(np.diff(lam)<=0):
        raise ValueError("strictly increasing wavelength results required")
    def geometry_without_lambda(r):
        d=dict(r["input"])
        d.pop("wavelength_nm",None)
        return json.dumps(d,sort_keys=True)
    if len(set(geometry_without_lambda(r) for r in runs))!=1:
        raise ValueError("FEM geometry/boundary settings differ; separate sweep required")
    for r in runs:
        if r.get("schema")!="LSCNS_fiber_research_vector_v1" or not r.get("modes"):
            raise ValueError("each run must be LSCNS_fiber_research_vector_v1")
    n=len(runs[0]["modes"])
    if any(len(r["modes"])!=n for r in runs):
        raise ValueError("mode count changed; cannot certify identity")
    order=list(range(n))
    neff_by_initial=[[] for _ in range(n)]
    overlap_by_initial=[[] for _ in range(n)]
    for i,idx in enumerate(order):
        neff_by_initial[i].append(runs[0]["modes"][idx]["neff_real"])
    for k in range(len(runs)-1):
        prev=[_mode_for_overlap(m) for m in runs[k]["modes"]]
        curr=[_mode_for_overlap(m) for m in runs[k+1]["modes"]]
        result=match_modes_by_overlap(prev,curr,min_overlap)
        if not result["all_tracked"]:
            raise ValueError(f"mode identity uncertain between lambda {lam[k]} and {lam[k+1]}; D/S BLOCKED")
        lookup={a["previous_index"]:a for a in result["assignments"]}
        next_order=[]
        for i,prev_idx in enumerate(order):
            match=lookup[prev_idx]
            curr_idx=match["current_index"]
            next_order.append(curr_idx)
            overlap_by_initial[i].append(match["overlap_power"])
            neff_by_initial[i].append(runs[k+1]["modes"][curr_idx]["neff_real"])
        order=next_order
    curves=[]
    for i in range(n):
        d=dispersion_from_tracked_sweep(lam,neff_by_initial[i],overlap_by_initial[i],min_overlap)
        curves.append({"mode_id":"initial_mode_"+str(i+1),"overlap_verified":True,
                       "neff_real":neff_by_initial[i],"overlap_power":overlap_by_initial[i],**d})
    return {"schema":"LSCNS_step4_tracked_dispersion_v080","engine_version":ENGINE_VERSION,
            "minimum_overlap_heuristic":min_overlap,"curves":curves,
            "validation":{"field_tracking":"PASS under fixed shared quadrature and overlap threshold",
                          "physical_accuracy":"UNVERIFIED; FEM/PML/difference-step and degeneracy need independent tests"}}


def main(argv=None):
    p=argparse.ArgumentParser(description=__doc__)
    sub=p.add_subparsers(dest="command",required=True)
    r=sub.add_parser("result",help="postprocess a Python FEM mode_result.json")
    r.add_argument("input");r.add_argument("--out",default="step4_metrics.json")
    r.add_argument("--n2-m2-W",type=float,default=None)
    s=sub.add_parser("sweep",help="overlap-track actual FEM full-quadrature wavelength runs")
    s.add_argument("input");s.add_argument("--out",default="step4_dispersion.json")
    s.add_argument("--min-overlap",type=float,default=.75)
    a=p.parse_args(argv)
    d=json.loads(Path(a.input).read_text(encoding="utf-8"))
    result=postprocess_result(d,a.n2_m2_W) if a.command=="result" else track_wavelength_sweep(d,a.min_overlap)
    _write(result,a.out)
    print(f"STEP4 v{ENGINE_VERSION}: {a.out}; FEM physics remains unverified.")


if __name__=="__main__": main()
