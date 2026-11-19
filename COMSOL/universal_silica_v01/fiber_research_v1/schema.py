"""Versioned design IO and STEP 1 v0.8.4 numeric/geometric guards. stdlib-only."""
import json
import math
from pathlib import Path

VERSION = "1.0.0"
VARIANTS = {"index_guiding":["step","graded","trench_assisted","w_profile","multilayer"],"solid_pcf":["hex","square","lma","high_nonlinearity","elliptic_birefringent"],"pbgf":["7_cell","19_cell"],"kagome":["hypocycloid"],"anti_resonant":["single_ring","nanf","dnanf","conjoined_tube","negative_curvature"],"bragg":["concentric"],"mcf":["hex","ring","trench_assisted"],"polarization_maintaining":["panda","bow_tie","elliptic_core"]}
ROOT = Path(__file__).resolve().parent

def load_presets():
    return json.loads((ROOT / "presets.json").read_text(encoding="utf-8"))

def number(obj, key, minimum=None, maximum=None, integer=False):
    value = obj.get(key)
    if isinstance(value, bool) or not isinstance(value, (int,float)) or not math.isfinite(value):
        raise ValueError(f"{key}: finite number required")
    if minimum is not None and value < minimum or maximum is not None and value > maximum:
        raise ValueError(f"{key}: value outside allowed range")
    if integer and int(value) != value:
        raise ValueError(f"{key}: integer required")
    return value

def validate_design(d):
    if not isinstance(d, dict) or d.get("schema_version") != VERSION:
        raise ValueError("unsupported schema_version; expected 1.0.0")
    f = d.get("family")
    v = d.get("variant")
    if not isinstance(f,str) or not isinstance(v,str):
        raise ValueError("family and variant must be strings")
    if v not in VARIANTS.get(f, []):
        raise ValueError(f"unsupported fiber family/variant: {f}/{v}")
    if not isinstance(d.get("name"), str) or not d["name"].strip():
        raise ValueError("name required")
    if len(d["name"])>120: raise ValueError("name exceeds 120 characters")
    if "target_neff" in d: number(d,"target_neff",.000001,3.5)
    number(d, "wavelength_nm", 250, 3500)
    radius = number(d, "domain_radius_um", 1.01, 1000)
    p = d.get("parameters")
    if not isinstance(p, dict) or len(p) > 50:
        raise ValueError("parameters object required")
    for key, val in p.items():
        number(p, key, -1e6, 1e6)
    for key in ("ellipse_ratio", "nested_radius_ratio", "graded_exponent", "trench_width_um", "layer_pitch_um"):
        if key in p: number(p, key, 1e-6, 1e6)
    for key in ("gap_um", "core_surround_um"):
        if key in p: number(p, key, 0, 1e6)
    req = {"index_guiding":("core_radius_um","delta_n"),"solid_pcf":("pitch_um","d_over_pitch","rings","defect_cells"),"pbgf":("pitch_um","d_over_pitch","rings","defect_cells","core_surround_um"),"kagome":("pitch_um","rings","core_radius_um","wall_thickness_um"),"anti_resonant":("core_radius_um","tube_count","tube_radius_um","wall_thickness_um"),"bragg":("core_radius_um","layer_count","layer_pitch_um"),"mcf":("core_radius_um","core_count","pitch_um"),"polarization_maintaining":("core_radius_um","delta_n","rod_radius_um","rod_offset_um")}
    for k in req[f]:
        number(p,k,0 if k=="core_surround_um" else 0.000001)
    for k in ("rings","defect_cells","tube_count","layer_count","core_count"):
        if k in p:
            number(p,k,1,30,integer=True)
    if f in ("solid_pcf","pbgf") and not (0 < p["d_over_pitch"] < .95):
        raise ValueError("d_over_pitch must be in (0,0.95)")
    if f=="pbgf" and p["defect_cells"] not in (7,19):
        raise ValueError("PBGF defect cells must be 7 or 19")
    if f=="anti_resonant" and (not 4<=p["tube_count"]<=12 or p["wall_thickness_um"]>=p["tube_radius_um"]/2):
        raise ValueError("invalid anti-resonant tube count/wall")
    if f=="mcf" and v=="trench_assisted": number(p,"trench_width_um",.000001)
    if f=="mcf" and (p["core_count"]>19 or p["pitch_um"]<=2*p["core_radius_um"]):
        raise ValueError("MCF cores overlap or count > 19")
    if f=="polarization_maintaining" and v!="elliptic_core" and p["rod_offset_um"]<=p["rod_radius_um"]+p["core_radius_um"]:
        raise ValueError("stress rods overlap optical core")
    mats=d.get("materials")
    if not isinstance(mats,dict) or "background" not in mats:
        raise ValueError("materials.background required")
    # STEP6b: Panda/Bow-tie rods are separate mechanical physical groups, never merged with silica cladding.
    if f=="polarization_maintaining" and v in ("panda","bow_tie") and "stress_rod" not in mats:
        raise ValueError("stress_rod optical material required for separate mechanical domain; cannot guess its composition")
    for mid,m in mats.items():
        if not isinstance(m,dict):
            raise ValueError(f"{mid}: material object required")
        model=m.get("model")
        if model in ("ge_doped_sellmeier","f_doped_sellmeier"):
            raise NotImplementedError(f"{mid}: TODO verified composition-specific Sellmeier source")
        if model not in ("pure_silica","air","constant"):
            raise ValueError(f"{mid}: unknown material")
        if model=="constant":
            number(m,"n",.5,3.5)
        number({"kappa":m.get("kappa",0)},"kappa",0,.099999)
    mesh=d.get("mesh")
    if not isinstance(mesh,dict):
        raise ValueError("mesh required")
    number(mesh,"bulk_max_um",.001,20)
    for key,lo,hi in (("through_wall_elements",3,30),("elements_per_wavelength",4,50),("curvature_segments",24,256)):
        number(mesh,key,lo,hi,integer=True)
    if mesh.get("resolution_policy", "boundary_refined") not in ("boundary_refined", "wavelength"):
        raise ValueError("mesh.resolution_policy must be boundary_refined or wavelength")
    if "wall_probe_angles" in mesh: number(mesh,"wall_probe_angles",8,128,integer=True)
    if "strict_checks" in mesh and not isinstance(mesh["strict_checks"],bool):
        raise ValueError("mesh.strict_checks must be boolean")
    if mesh.get("symmetry") not in ("full","c6v_experimental"):
        raise ValueError("mesh.symmetry invalid")
    if mesh["symmetry"] != "full":
        raise NotImplementedError("C6v sector boundary formulation not yet validated; use full")
    b=d.get("boundary")
    if not isinstance(b,dict) or b.get("type") not in ("pec","pml","sbc"):
        raise ValueError("boundary.type invalid")
    t=number(b,"thickness_um",0,100)
    number(b,"strength",.00001,30)
    number(b,"order",1,6,integer=True)
    if b["type"]=="pml" and not (0<t<.8*radius):
        raise ValueError("PML thickness invalid")
    try:
        from .design_checks import check_geometry
    except ImportError:
        from design_checks import check_geometry
    check_geometry(d)
    return d

def load_design(path):
    return validate_design(json.loads(Path(path).read_text(encoding="utf-8")))

def save_design(d,path):
    validate_design(d)
    dst=Path(path); dst.parent.mkdir(parents=True,exist_ok=True)
    dst.write_text(json.dumps(d,ensure_ascii=False,indent=2,allow_nan=False)+"\n",encoding="utf-8")
    return dst

