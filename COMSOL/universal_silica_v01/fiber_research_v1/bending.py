"""STEP 6a / engine v0.8.1 — weak-curvature equivalent-index bend analysis.

Coordinates x,y are in um, physical bend radius R is in mm. Angle theta
(degrees counterclockwise from +x) points OUTWARD from centre of curvature.
x_out = x cos(theta) + y sin(theta) [um]
n_eq(x,y) = n(x,y) * (1 + x_out / R_eff_um)
eps_eq = eps_material * (1 + x_out / R_eff_um)^2 (complex eps supported).

If stress correction is requested, R_eff = f_source * R (f_source dimensionless).
f_source must be explicitly entered with a source URL; there is NO assumed
universal photoelastic factor. Source examples:
- https://www.comsol.com/paper/study-of-bending-losses-in-optical-fibers-using-comsol-multiphysics-61111
- https://optiwave.wiki/kb/optical_fiber_bending_loss_with_stress_effect/
- https://doi.org/10.1109/JLT.2019.2912916
DO NOT apply quoted silica/SMF factor to HCF, PCF, MCF, Ge/F doped glass or other
wavelength/temperature without model-specific evidence.
COMSOL step_index_fiber_bend uses axisymmetric geometry; this conformal
2D Cartesian approximation is NOT the identical COMSOL formulation.

Validation contract: R_eff > outer computation radius (mathematical positivity).
R_eff comparable to transverse extent may violate the weak-curvature assumption
even if accepted numerically; report extent/R_eff and NEVER certify loss.
"""
from __future__ import annotations
import copy
import hashlib
import json
import math
from datetime import datetime, timezone
from pathlib import Path
from typing import Callable
import numpy as np

ENGINE_VERSION="0.8.1"
CONFORMAL_SOURCE="https://www.comsol.com/paper/study-of-bending-losses-in-optical-fibers-using-comsol-multiphysics-61111"
COMSOL_V3_SOURCE="https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.step_index_fiber_bend/step_index_fiber_bend.html"

def _finite(name,value,positive=False):
    if isinstance(value,bool):
        raise ValueError(f"{name}: boolean not allowed")
    try: v=float(value)
    except (TypeError,ValueError):
        raise ValueError(f"{name}: finite number required") from None
    if not math.isfinite(v) or (positive and v<=0):
        raise ValueError(f"{name}: finite positive value required" if positive else f"{name}: finite value required")
    return v


def _outer_extent_um(design):
    radius=_finite("domain_radius_um",design["domain_radius_um"],positive=True)
    bc=design["boundary"]
    thickness=_finite("boundary.thickness_um",bc.get("thickness_um",0)) if bc["type"]=="pml" else 0.
    return radius+thickness


def configure_bend(design,radius_mm,direction_deg=0.0,
                   reff_multiplier=None,stress_source_url=None,stress_source_note=None):
    """Make transient input copy; do not change source design JSON schema v1.0.0."""
    r=_finite("bend radius [mm]",radius_mm,positive=True)
    angle=_finite("bend angle [deg]",direction_deg)%360.
    if reff_multiplier is None:
        if stress_source_url is not None or stress_source_note is not None:
            raise ValueError("stress source supplied without R_eff multiplier")
        mult=1.
        correction={"mode":"none","multiplier":1.0,"source":None}
    else:
        mult=_finite("R_eff/R multiplier",reff_multiplier,positive=True)
        if not isinstance(stress_source_url,str) or not stress_source_url.startswith(("https://","http://")):
            raise ValueError("stress correction requires a real literature source URL")
        if not stress_source_note or not isinstance(stress_source_note,str) or not stress_source_note.strip():
            raise ValueError("stress correction requires applicability note (fiber, wavelength, conditions)")
        correction={"mode":"user_literature_coefficient","multiplier":mult,
                    "source":stress_source_url,"applicability":stress_source_note}
    reff_um=r*1000.*mult
    extent=_outer_extent_um(design)
    if reff_um<=extent:
        raise ValueError(f"R_eff={reff_um:g} um <= computational outer radius={extent:g} um; invalid conformal index")
    result=copy.deepcopy(design)
    result["_bend"]={"model":"2d_conformal","radius_mm":r,
                     "direction_deg":angle,"R_eff_mm":r*mult,
                     "stress_correction":correction,
                     "outer_extent_over_R_eff":extent/reff_um,
                     "source":CONFORMAL_SOURCE,
                     "status":"EXPERIMENTAL_NOT_BENCHMARKED"}
    return result


def conformal_factor(x_um,y_um,R_eff_mm,direction_deg):
    """Return 1+(x cosθ+y sinθ)/(1000 R_eff_mm), broadcasting arrays."""
    r=_finite("R_eff [mm]",R_eff_mm,positive=True)
    ang=math.radians(_finite("direction [deg]",direction_deg))
    x=np.asarray(x_um,dtype=float)
    y=np.asarray(y_um,dtype=float)
    if not np.isfinite(x).all() or not np.isfinite(y).all():
        raise ValueError("finite transverse coordinates required")
    factor=1.+(x*math.cos(ang)+y*math.sin(ang))/(1000*r)
    if not np.isfinite(factor).all() or np.any(factor<=0):
        raise ValueError("invalid nonpositive conformal factor")
    return factor


def bend_per_turn(loss_dB_m,radius_mm):
    """Attenuation for ONE circumference using physical R, not R_eff. Signed."""
    a=_finite("loss[dB/m]",loss_dB_m)
    r=_finite("physical R[mm]",radius_mm,positive=True)
    return a*(2*math.pi*r*1e-3)


def _normalized_design_hash(design):
    """Canonical UTF-8 JSON of the complete source input; excludes generated _bend."""
    src={k:v for k,v in design.items() if k!="_bend"}
    value=json.dumps(src,sort_keys=True,separators=(",",":"),ensure_ascii=False,allow_nan=False)
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


def _complex_mode_quadrature(mode):
    q=mode.get("field_quadrature")
    if not isinstance(q,dict):
        raise ValueError("full physical field_quadrature required for bend mode matching")
    values={}
    for part in ("Ex","Ey"):
        xy=np.asarray(q[part],dtype=float)
        if xy.ndim!=2 or xy.shape[1]!=2:
            raise ValueError("complex field quadrature must be [Re,Im] pairs")
        values[part]=xy[:,0]+1j*xy[:,1]
    return {"points_um":q["points_um"],"weights_um2":q["weights_um2"],**values}


def _solver_fem(design,mesh,n_modes,target_neff):
    try:
        from .vector_fem import solve
    except ImportError:
        from vector_fem import solve
    return solve(design,mesh,n_modes=n_modes,target_neff=target_neff,
                 export_full_quadrature=True,max_sample_points=40)


def bend_radius_direction_sweep(design,radii_mm,directions_deg,
                                reff_multiplier=None,stress_source_url=None,
                                stress_source_note=None,n_modes=2,
                                target_neff=None,solve_fn=None,mesh=None,
                                minimum_overlap=.75,output_dir=None):
    """One physical mesh for straight + all bends; actual FEM solve_fn may be injected for tests.

    No silent mode-index sorting. Only the first straight reference mode is
    tracked; this is a *candidate* and requires verified HE11 physical ID.
    If full overlap fails, record TODO, not a made-up loss value.
    """
    try:
        from .schema import validate_design
        from .geometry import build_mesh
        from .physical_metrics import match_modes_by_overlap
    except ImportError:
        from schema import validate_design
        from geometry import build_mesh
        from physical_metrics import match_modes_by_overlap
    validate_design(design)
    if design.get("_bend") is not None:
        raise ValueError("source design must be unbent; do not nest a bend")
    if design["boundary"]["type"]!="pml":
        raise ValueError("bend loss requires configured PML for candidate outgoing loss (not PEC/SBC)")
    if not isinstance(radii_mm,(tuple,list)) or not 1<=len(radii_mm)<=100:
        raise ValueError("1..100 bend radii needed")
    if not isinstance(directions_deg,(tuple,list)) or not 1<=len(directions_deg)<=36:
        raise ValueError("1..36 bend directions needed")
    if not isinstance(n_modes,int) or isinstance(n_modes,bool) or not 1<=n_modes<=8:
        raise ValueError("mode count 1..8")
    threshold=_finite("minimum overlap",minimum_overlap,positive=True)
    if threshold>1:raise ValueError("overlap must <=1")
    radii=[_finite("radius [mm]",r,positive=True) for r in radii_mm]
    dirs=[_finite("direction [deg]",a)%360. for a in directions_deg]
    if len(set(radii))!=len(radii) or len(set(dirs))!=len(dirs):
        raise ValueError("duplicate radii/directions not supported")
    # Solve large R first, continue to smaller radius to reduce accidental mode switches.
    radii.sort(reverse=True)
    if mesh is None:mesh=build_mesh(design,Path(output_dir)/"common_mesh" if output_dir else None)
    solver=solve_fn or _solver_fem
    straight=solver(design,mesh,n_modes,target_neff)
    if not straight.get("modes"):
        raise RuntimeError("straight fiber has no candidate reference modes")
    straight_fm=straight["modes"][0]
    root={"schema":"LSCNS_step6a_bend_sweep_v081","engine_version":ENGINE_VERSION,
          "created_utc":datetime.now(timezone.utc).isoformat(),
          "input_sha256":_normalized_design_hash(design),
          "source_design":design,
          "physical_model":{"conformal":"n_eq=n(x,y)*(1+x_out/R_eff)", "R_eff":"f_source*R or R",
                            "direction":"x_out=x*cos(theta)+y*sin(theta)",
                            "source":CONFORMAL_SOURCE},
          "stress_correction_requested":reff_multiplier is not None,
          "reference":{"mode_selector":"straight_modes[0], NOT certified HE11",
                       "alpha_total_dB_m":straight_fm["alpha_total_dB_m"],
                       "neff_real":straight_fm["neff_real"],
                       "classification":straight_fm.get("status")},
          "directions_deg":dirs,"radii_mm_descending":radii,
          "cases":[],"validation":{"V3":"EVIDENCE_REQUIRED",
                                   "V4_PML":"NOT_RUN","V5_MESH":"NOT_RUN",
                                   "loss_certified":False,
                                   "notes":"COMSOL reference must match geometry/materials/PML/bend formulation; solver may contain spurious modes"}}
    for direction in dirs:
        prev=straight["modes"]
        tracked_idx=0
        for radius in radii:
            case={"radius_mm":radius,"direction_deg":direction,
                  "alpha_total_dB_m":None,"delta_total_dB_m":None,
                  "loss_per_turn_total_dB":None,
                  "neff_real":None,"overlap_power":None,
                  "status":"UNCOMPUTED"}
            try:
                curved=configure_bend(design,radius,direction,reff_multiplier,
                                       stress_source_url,stress_source_note)
                case["R_eff_mm"]=curved["_bend"]["R_eff_mm"]
                case["outer_extent_over_R_eff"]=curved["_bend"]["outer_extent_over_R_eff"]
                case["stress_correction"]=curved["_bend"]["stress_correction"]
                result=solver(curved,mesh,n_modes,target_neff)
                if not result.get("modes"):
                    raise ValueError("no bent candidate modes")
                r=match_modes_by_overlap([_complex_mode_quadrature(m) for m in prev],
                                         [_complex_mode_quadrature(m) for m in result["modes"]],
                                         minimum_overlap=threshold)
                pick=next((m for m in r["assignments"] if m["previous_index"]==tracked_idx),None)
                if pick is None or not pick["tracked"]:
                    case["status"]="MODE_TRACKING_FAILED"
                    case["evidence_required"]="confirm common quadrature and HE11 mode identity, reduce R step"
                else:
                    candidate=result["modes"][pick["current_index"]]
                    case["overlap_power"]=pick["overlap_power"]
                    case["mode_classification"]=candidate["status"]
                    if candidate["status"]!="candidate_unverified":
                        case["status"]="SPURIOUS_OR_UNSTABLE_CANDIDATE"
                    else:
                        loss=_finite("total loss",candidate["alpha_total_dB_m"])
                        case.update({"status":"COMPUTED_UNVERIFIED",
                                     "neff_real":candidate["neff_real"],
                                     "neff_imag":candidate["neff_imag"],
                                     "alpha_total_dB_m":loss,
                                     "delta_total_dB_m":loss-straight_fm["alpha_total_dB_m"],
                                     "loss_per_turn_total_dB":bend_per_turn(loss,radius)})
                    # Advance only when the physical label is provisionally consistent.
                    if case["status"]=="COMPUTED_UNVERIFIED":
                        prev=result["modes"]
                        tracked_idx=pick["current_index"]
                if case["status"]!="COMPUTED_UNVERIFIED":
                    case.setdefault("evidence_required","valid physical mode, PML sign, overlap and mesh convergence")
            except (RuntimeError,ValueError,KeyError,IndexError) as err:
                case["status"]="SOLVE_FAILED"
                case["error"]=str(err)
            root["cases"].append(case)
            if output_dir:
                p=Path(output_dir)
                p.mkdir(parents=True,exist_ok=True)
                (p/"bend_sweep_partial.json").write_text(
                    json.dumps(root,indent=2,ensure_ascii=False,allow_nan=False)+"\n",encoding="utf-8")
    return root
