"""Reproducible independent STEP 3 V1/V4/V5 benchmarks, not product certification."""
import argparse
import copy
import json
import platform
import importlib.metadata
from pathlib import Path
from datetime import datetime,timezone
import numpy as np
from scipy.special import jnp_zeros
try:
    from .schema import load_presets,save_design
    from .geometry import build_mesh
    from .vector_fem import solve
    from .step3_reference import he11_exact,leaky_circular_exact
    from .modal_study import run_cases,match_subspace,summarize,input_digest
    from .modal_diagnostics import disk_probes
except ImportError:
    from schema import load_presets,save_design
    from geometry import build_mesh
    from vector_fem import solve
    from step3_reference import he11_exact,leaky_circular_exact
    from modal_study import run_cases,match_subspace,summarize,input_digest
    from modal_diagnostics import disk_probes


def benchmark_design(kind="guided"):
    d=load_presets()["index_guiding__step"]
    d["name"]="STEP3 synthetic "+kind+" independent benchmark"
    d["metadata"]={"status":"SYNTHETIC_NUMERICAL_BENCHMARK_NOT_A_PRODUCT",
                   "material_basis":"explicit test indices, not inferred silica/air product constants"}
    d["domain_radius_um"]=6.
    d["parameters"]={"core_radius_um":1.3,"delta_n":.06}
    d["materials"]={"core":{"model":"constant","n":1.5},"background":{"model":"constant","n":1.44}}
    d["mesh"].update(elements_per_wavelength=4,curvature_segments=192,bulk_max_um=.8)
    d["solver"]={"element_order":2,"max_dofs":120000,"backend":"beta_squared","tolerance":1e-11}
    d["boundary"].update(type="pec",thickness_um=0)
    if kind=="radiative":
        d["parameters"].update(core_radius_um=3.,delta_n=.45)
        d["materials"]={"core":{"model":"constant","n":1.0},"background":{"model":"constant","n":1.45}}
        d["boundary"].update(type="pml",thickness_um=2.,strength=4.,order=3)
    elif kind=="uniform_absorbing":
        d["domain_radius_um"]=3.
        d["parameters"].update(core_radius_um=1.,delta_n=.001)
        d["materials"]={key:{"model":"constant","n":1.45,"kappa":1e-4} for key in ("core","background")}
    return d


def run_validation(output_dir):
    out=Path(output_dir);out.mkdir(parents=True,exist_ok=True)
    guided=benchmark_design();guided_ref=he11_exact(1.5,1.44,1.3,1.55)
    guided_cases=[]
    for count,curvature in ((4,144),(6,192),(8,240)):
        d=copy.deepcopy(guided);d["mesh"].update(elements_per_wavelength=count,curvature_segments=curvature)
        guided_cases.append(("guided_mesh_"+str(count),d))
    for radius in (6.,7.2,8.):
        d=copy.deepcopy(guided);d["domain_radius_um"]=radius;d["mesh"]["curvature_segments"]=240
        guided_cases.append(("guided_window_"+str(radius),d))
    gr,example=run_cases(guided,guided_cases,out/"guided",target_neff=guided_ref["neff_real"])
    complete_g=all(r["state"]=="COMPUTED" and r["tracking"]["status"]=="PASS_TRACKING" for r in gr)
    error=max((max(abs(n-guided_ref["neff_real"]) for n in r["neff_real"]) for r in gr if r["state"]=="COMPUTED"),default=float("inf"))
    mesh_rows=[r for r in gr if r["name"].startswith("guided_mesh") and r["state"]=="COMPUTED"]
    last=float(max(abs(a-b) for a,b in zip(sorted(mesh_rows[-1]["neff_real"]),sorted(mesh_rows[-2]["neff_real"])))) if len(mesh_rows)>=2 else float("inf")
    V1={"status":"PASS_SCOPED" if complete_g and error<=1e-5 and last<=1e-5 else "FAIL",
        "reference":guided_ref,"maximum_absolute_Re_error":error,"last_mesh_Re_change":last,
        "absolute_Re_tolerance":1e-5,"cases":gr,"scope":"synthetic circular guided HE11 only"}

    # Uniform circular PEC waveguide: exact complex TE11 beta, no FEM reference fit.
    uniform=benchmark_design("uniform_absorbing");k0=2*np.pi/1.55
    exact_uniform=np.sqrt((1.45+1e-4j)**2-(jnp_zeros(1,1)[0]/(k0*3))**2)
    ur=solve(uniform,n_modes=2,target_neff=exact_uniform.real,max_sample_points=16)
    uerror=max(abs(complex(m["neff_real"],m["neff_imag"])-exact_uniform) for m in ur["modes"])
    # Same matrix, different pencil; tests the linearization instead of reusing it as an oracle.
    d=copy.deepcopy(guided);d["solver"]["element_order"]=1;d["mesh"]["curvature_segments"]=96
    mesh=build_mesh(d);p,w=disk_probes(5.5,1.3)
    br=solve(d,mesh,n_modes=2,target_neff=guided_ref["neff_real"],max_sample_points=8,probe_points=p,probe_weights=w)
    d["solver"]["backend"]="companion"
    cr=solve(d,mesh,n_modes=2,target_neff=guided_ref["neff_real"],max_sample_points=8,probe_points=p,probe_weights=w)
    pencil_match=match_subspace(br,cr)
    pencil_error=max(abs(a-b) for a,b in zip(sorted(m["neff_real"] for m in br["modes"]),sorted(m["neff_real"] for m in cr["modes"])))
    integrity=gr+[dict(state="COMPUTED",**summarize(ur,range(len(ur["modes"]))))]
    good_integrity=all(r["state"]=="COMPUTED" and r["all_finite"] and r["max_qep_residual"]<=1e-7 and r["max_pencil_residual"]<=1e-7 and r["max_gauss_residual"]<=1e-7 and all(x>0 for x in r["power_W"]) for r in integrity)
    V4={"status":"PASS_SCOPED" if good_integrity and len(ur["modes"])==2 and uerror<=1e-5 and pencil_error<=1e-9 and pencil_match["status"]=="PASS_TRACKING" else "FAIL",
        "uniform_TE11_reference":{"neff_real":exact_uniform.real,"neff_imag":exact_uniform.imag,"maximum_complex_error":uerror},
        "beta_squared_companion_Re_difference":pencil_error,"pencil_subspace":pencil_match,
        "scope":"finite E/H, forward power, residuals, physical probe rank; not proof against every spurious mode"}

    radiative=benchmark_design("radiative")
    leak_ref=leaky_circular_exact(1.,1.45,3.,1.55,.981+.005j)
    cases=[("radiative_baseline",copy.deepcopy(radiative))]
    for count,curvature in ((4,144),(6,192),(8,240)):
        d=copy.deepcopy(radiative);d["mesh"].update(elements_per_wavelength=count,curvature_segments=curvature);cases.append(("radiative_mesh_"+str(count),d))
    for radius in (6.,6.6,7.2):
        d=copy.deepcopy(radiative);d["domain_radius_um"]=radius;cases.append(("radiative_window_"+str(radius),d))
    for field,values in (("thickness_um",(1.5,2.,2.5)),("strength",(3.,4.,5.)),("order",(2,3,4))):
        for value in values:
            d=copy.deepcopy(radiative);d["boundary"][field]=value;cases.append(("pml_"+field+"_"+str(value),d))
    lr,leaky_example=run_cases(radiative,cases,out/"radiative",target_neff=leak_ref["neff_real"])
    complete_l=all(r["state"]=="COMPUTED" and r["tracking"]["status"]=="PASS_TRACKING" for r in lr)
    im=np.asarray([value for r in lr if r["state"]=="COMPUTED" for value in r["neff_imag"]])
    radiative_integrity=all(r["state"]=="COMPUTED" and r["max_qep_residual"]<=1e-7 and r["max_pencil_residual"]<=1e-7 and r["max_gauss_residual"]<=1e-7 and r["all_finite"] and all(p>0 for p in r["power_W"]) for r in lr)
    V4["radiative_integrity_passed"]=radiative_integrity
    if not radiative_integrity:V4["status"]="FAIL"
    relative_error=float(np.max(np.abs(im-leak_ref["neff_imag"]))/leak_ref["neff_imag"]) if len(im) else float("inf")
    spread=float(np.ptp(im)/leak_ref["neff_imag"]) if len(im) else float("inf")
    re_error=max((max(abs(n-leak_ref["neff_real"]) for n in r["neff_real"]) for r in lr if r["state"]=="COMPUTED"),default=float("inf"))
    V5={"status":"PASS_SCOPED" if complete_l and np.min(im)>0 and relative_error<=.05 and spread<=.05 and re_error<=1e-4 else "FAIL",
        "reference":leak_ref,"maximum_relative_Im_error":relative_error,"relative_Im_spread":spread,
        "maximum_absolute_Re_error":re_error,"Im_relative_tolerance":.05,"Re_absolute_tolerance":1e-4,
        "cases":lr,"scope":"synthetic strongly radiating circular air-like core in n=1.45 infinite exterior; NOT tiny-loss HCF certification"}
    report={"schema":"LSCNS_step3_validation_v085","engine_version":"0.8.5","timestamp_utc":datetime.now(timezone.utc).isoformat(),
            "environment":{"python":platform.python_version(),**{k:importlib.metadata.version(k) for k in ("numpy","scipy","scikit-fem","gmsh")}},
            "V1":V1,"V4":V4,"V5":V5,"all_scoped_tests_passed":all(x["status"]=="PASS_SCOPED" for x in (V1,V4,V5)),
            "product_loss_certified":False,"next_step_started":False}
    (out/"step3_validation.json").write_text(json.dumps(report,indent=2,allow_nan=False)+"\n")
    for label,result in (("guided",example),("radiative",leaky_example)):
        if result is None:continue
        for m in result["modes"]:m.pop("tracking_fields",None)
        result["benchmark_validation"]={"scope":label+" synthetic input only","input_digest":input_digest(result["input"]),
                                        "checks":{"V1":V1["status"] if label=="guided" else "NOT_APPLICABLE","V4":V4["status"],"V5":V5["status"] if label=="radiative" else "NOT_APPLICABLE"},
                                        "product_loss_certified":False}
        (out/(label+"_mode_result.json")).write_text(json.dumps(result,indent=2,allow_nan=False)+"\n")
        save_design(result["input"],out/(label+"_input.json"))
    return report


def main():
    p=argparse.ArgumentParser(description=__doc__);p.add_argument("--out",required=True);a=p.parse_args()
    report=run_validation(a.out)
    print({k:report[k]["status"] for k in ("V1","V4","V5")},flush=True)
    if not report["all_scoped_tests_passed"]:raise SystemExit(1)

if __name__=="__main__":main()
