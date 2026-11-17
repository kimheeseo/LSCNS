"""STEP 5 v0.8.3: scalar 2D fixed-beta Bloch--Floquet unit-cell eigensolver.

TIME/SPACE: exp(i*beta*z-i*omega*t), r=A*(u,v), n is real and nondispersive.
Scalar proxy -Laplacian_t(Ez)+beta^2 Ez = (omega/c)^2 n(r)^2 Ez.
Boundary: E(q+e_i)=exp(i*k_Bloch.dot(a_i))*E(q).
Central finite differences in oblique fractional coordinates include mixed
derivatives. This is NOT the full-vector 3D Maxwell eigenproblem or an FEM.
A sampled gap is a CANDIDATE, never a certified complete 2D photonic gap.
"""
from __future__ import annotations
import copy
import hashlib
import json
import math
from pathlib import Path
import numpy as np
from scipy.sparse import coo_matrix, diags
from scipy.sparse.linalg import eigsh

SCHEMA_INPUT = "LSCNS_STEP5_BLOCH_INPUT_V083"
SCHEMA_OUTPUT = "LSCNS_STEP5_BLOCH_RESULT_V083"
LATTICES = ("triangular", "hexagonal", "square", "kagome")
MOTIFS = {
    "triangular": ((0., 0.),),
    "hexagonal": ((1./3, 1./3), (2./3, 2./3)),  # two-site honeycomb basis
    "square": ((0., 0.),),
    "kagome": ((.5, 0.), (0., .5), (.5, .5)),    # idealized three-site Kagome
}

def _finite(x, name, low=None, high=None):
    if isinstance(x, bool) or not isinstance(x, (int, float)) or not math.isfinite(x):
        raise ValueError(name + " must be a finite number")
    if low is not None and x < low or high is not None and x > high:
        raise ValueError(name + " outside allowed range")
    return float(x)

def primitive(lattice, pitch_um):
    a = float(pitch_um)
    if lattice == "square":
        return np.array([[a, 0.], [0., a]], float)
    if lattice in ("triangular", "hexagonal", "kagome"):
        return np.array([[a, a/2], [0., math.sqrt(3)*a/2]], float)
    raise ValueError("unknown lattice")

def smallest_separation(lattice, pitch_um):
    A = primitive(lattice, pitch_um)
    sites = MOTIFS[lattice]
    distances = []
    for i, site in enumerate(sites):
        for j, other in enumerate(sites):
            for du in (-1, 0, 1):
                for dv in (-1, 0, 1):
                    if i == j and du == 0 and dv == 0:
                        continue
                    diff = A @ (np.asarray(site)-np.asarray(other)+np.array([du, dv]))
                    distances.append(float(np.linalg.norm(diff)))
    return min(distances)

def validate_config(cfg):
    if not isinstance(cfg, dict) or cfg.get("schema") != SCHEMA_INPUT:
        raise ValueError("STEP5 input schema required: " + SCHEMA_INPUT)
    lattice = cfg.get("lattice")
    if lattice not in LATTICES:
        raise ValueError("lattice must be triangular/hexagonal/square/kagome")
    pitch = _finite(cfg.get("pitch_um"), "pitch_um", 0.1, 1000)
    hole = _finite(cfg.get("hole_diameter_over_pitch"), "hole_diameter_over_pitch", .01, .95)
    _finite(cfg.get("n_host"), "n_host", .9, 5.)
    _finite(cfg.get("n_air"), "n_air", .9, 5.)
    d_nearest = smallest_separation(lattice, pitch)
    if hole*pitch >= .98*d_nearest:
        raise ValueError("overlapping holes in this multi-site unit cell")
    for field, lo, hi in (("grid_n", 10, 96), ("bands", 2, 16),
                          ("path_points_per_segment", 2, 30), ("kmesh_n", 2, 12)):
        v = cfg.get(field)
        if isinstance(v, bool) or not isinstance(v, int) or not lo <= v <= hi:
            raise ValueError(field + " must be an integer in allowed range")
    if cfg["bands"] >= cfg["grid_n"]**2 - 1:
        raise ValueError("bands must be smaller than grid_n^2-1")
    betas = cfg.get("beta_norms")
    if not isinstance(betas, list) or not 1 <= len(betas) <= 10:
        raise ValueError("beta_norms must contain 1..10 normalized beta values")
    for b in betas:
        _finite(b, "beta_norms entry", .001, 5.)
    if len(betas)*cfg["kmesh_n"]**2+3*cfg["path_points_per_segment"] > 260:
        raise ValueError("too many eigensolves; reduce kmesh/path sampling")
    return cfg

def hole_mask(cfg):
    validate_config(cfg)
    n = cfg["grid_n"]
    A = primitive(cfg["lattice"], cfg["pitch_um"])
    q = (np.arange(n)+.5)/n
    U, V = np.meshgrid(q, q, indexing="ij")
    d2 = np.full((n,n), np.inf)
    for u0,v0 in MOTIFS[cfg["lattice"]]:
        for dx in (-1,0,1):
            for dy in (-1,0,1):
                du, dv = U-u0-dx, V-v0-dy
                xx = A[0,0]*du+A[0,1]*dv
                yy = A[1,0]*du+A[1,1]*dv
                d2 = np.minimum(d2, xx*xx+yy*yy)
    return d2 < (cfg["hole_diameter_over_pitch"]*cfg["pitch_um"]/2)**2

def dielectric_grid(cfg):
    air = hole_mask(cfg)
    return np.where(air, cfg["n_air"]**2, cfg["n_host"]**2).ravel(), air

def _phase_wrap(index, delta, n, reciprocal_fraction):
    total = index+delta
    return total % n, int(math.floor(total/n)), reciprocal_fraction

def assemble_operator(cfg, k_frac, beta_norm):
    """Sparse Hermitian generalized eigen K phi=lambda M phi; coordinates in um."""
    validate_config(cfg)
    N = cfg["grid_n"]
    lattice = primitive(cfg["lattice"], cfg["pitch_um"])
    inv = np.linalg.inv(lattice)
    g = inv @ inv.T
    step = 1./N
    g11, g22, g12 = g[0,0], g[1,1], g[0,1]
    b_um = 2*math.pi*beta_norm/cfg["pitch_um"]
    eps, _ = dielectric_grid(cfg)
    kx, ky = map(float, k_frac)
    if not all(math.isfinite(v) for v in (kx,ky)):
        raise ValueError("nonfinite k_frac")
    diagonal = 2*(g11+g22)/(step*step)+b_um*b_um
    # -d_uv coefficient: -2*g12 * (u+v+ - u+v- - u-v+ + u-v-) / (4*h^2).
    stencils = ((0,0,diagonal),
                (1,0,-g11/step**2),(-1,0,-g11/step**2),
                (0,1,-g22/step**2),(0,-1,-g22/step**2),
                (1,1,-g12/(2*step**2)),(-1,-1,-g12/(2*step**2)),
                (1,-1,g12/(2*step**2)),(-1,1,g12/(2*step**2)))
    rows=[]; cols=[]; vals=[]
    for i in range(N):
        for j in range(N):
            r=i*N+j
            for di,dj,c in stencils:
                if abs(c)<1e-14:
                    continue
                ti,wi,_=_phase_wrap(i,di,N,kx)
                tj,wj,_=_phase_wrap(j,dj,N,ky)
                phase=np.exp(2j*math.pi*(kx*wi+ky*wj))
                rows.append(r);cols.append(ti*N+tj);vals.append(c*phase)
    K=coo_matrix((vals,(rows,cols)),shape=(N*N,N*N),dtype=complex).tocsr()
    M=diags(eps,format="csr").astype(complex)
    return K,M

def eigenfrequencies(cfg, k_frac, beta_norm):
    K,M = assemble_operator(cfg,k_frac,beta_norm)
    vals = eigsh(K, k=cfg["bands"], M=M, sigma=-1e-7, which="LM",
                 tol=2e-9, maxiter=3000, return_eigenvectors=False)
    vals = np.sort(vals.real)
    if np.min(vals)<-1e-8 or not np.all(np.isfinite(vals)):
        raise RuntimeError("nonphysical generalized eigenvalue: mesh/phase issue")
    normalized = np.sqrt(np.maximum(vals,0))*cfg["pitch_um"]/(2*math.pi)
    return normalized.tolist()

def high_symmetry_path(lattice, per_segment):
    if lattice == "square":
        labels=("Γ","X","M","Γ")
        vertices=((0.,0.),(.5,0.),(.5,.5),(0.,0.))
    else:
        labels=("Γ","M","K","Γ")
        vertices=((0.,0.),(.5,.5),(2./3,1./3),(0.,0.))
    pts=[]; names=[]
    for i in range(3):
        for j in range(per_segment):
            t=j/per_segment
            pts.append([(1-t)*vertices[i][q]+t*vertices[i+1][q] for q in (0,1)])
            names.append(labels[i] if j==0 else "")
    pts.append(list(vertices[-1]));names.append(labels[-1])
    return pts,names

def _kmesh_points(n,lattice):
    # Fundamental reciprocal parallelogram tiles the full BZ; finite sampling only.
    pts=[[i/n-.5,j/n-.5] for i in range(n) for j in range(n)]
    path,_=high_symmetry_path(lattice,2)
    for p in path:
        if not any(np.linalg.norm(np.asarray(p)-q)<1e-12 for q in pts):
            pts.append(p)
    return pts

def analyze(cfg):
    validate_config(cfg)
    eps,holes=dielectric_grid(cfg)
    lattice=primitive(cfg["lattice"],cfg["pitch_um"])
    B=2*math.pi*np.linalg.inv(lattice).T
    ks,names=high_symmetry_path(cfg["lattice"],cfg["path_points_per_segment"])
    distance=[0.]
    for i in range(1,len(ks)):
        distance.append(distance[-1]+float(np.linalg.norm((B/(2*math.pi))@(np.asarray(ks[i])-np.asarray(ks[i-1])))*cfg["pitch_um"]))
    bands=[eigenfrequencies(cfg,k,cfg["beta_norms"][0]) for k in ks]
    mesh_pts=_kmesh_points(cfg["kmesh_n"],cfg["lattice"])
    gaps=[]; beta_summaries=[]
    for beta in cfg["beta_norms"]:
        spectrum=np.asarray([eigenfrequencies(cfg,k,beta) for k in mesh_pts],dtype=float)
        lo=spectrum.max(axis=0)
        hi=spectrum.min(axis=0)
        found=[]
        for j in range(cfg["bands"]-1):
            left=float(lo[j]);right=float(hi[j+1])
            if right>left+1e-8:
                gap={"band_below":j+1,"band_above":j+2,"f_low":left,"f_high":right,
                     "width":right-left,"relative_width":(right-left)/((right+left)/2),
                     "beta_norm":beta,"beta_over_k0_at_f_low":beta/left,
                     "beta_over_k0_at_f_high":beta/right,
                     "status":"SAMPLED_GAP_CANDIDATE"}
                found.append(gap);gaps.append(gap)
        beta_summaries.append({"beta_norm":beta,"sampled_k_count":len(mesh_pts),
                               "gap_candidates":len(found)})
    payload=json.dumps(cfg,sort_keys=True,separators=(",",":"),allow_nan=False)
    return {"schema":SCHEMA_OUTPUT,"engine_version":"0.8.3",
            "status":"SCALAR_SAMPLED_UNVERIFIED","input":copy.deepcopy(cfg),
            "input_sha256":hashlib.sha256(payload.encode()).hexdigest(),
            "model":"2D scalar fixed-beta -Delta_t Ez + beta^2 Ez = k0^2 n^2 Ez",
            "k_phase_convention":"E(r+a_i)=exp(i*kB·a_i)E(r)",
            "unit_convention":"pitch_um, beta_norm=beta*pitch/(2*pi), f=pitch/lambda=omega*pitch/(2*pi*c)",
            "materials":"real constant indices; wavelength dispersion not modeled",
            "lattice_vectors_um":lattice.tolist(),"reciprocal_vectors_um_inv":B.tolist(),
            "basis_sites_fractional":[list(p) for p in MOTIFS[cfg["lattice"]]],
            "air_fill_fraction_discrete":float(holes.mean()),
            "unit_cell_air_mask":holes.astype(int).tolist(),
            "path":{"k_frac":ks,"labels":names,"distance_norm":distance,
                    "bands_f":bands,"beta_norm":cfg["beta_norms"][0]},
            "gap_map":{"candidates":gaps,"beta_summaries":beta_summaries,
                       "sampled_k_points":mesh_pts,"light_lines_beta_over_k0":
                       [{"label":"air","n":cfg["n_air"]},{"label":"host","n":cfg["n_host"]}],
                       "status":"SAMPLED_ONLY_NOT_COMPLETE_BZ_PROOF"},
            "checks":{"hermitian_operator":"tested synthetically; input-specific unverified",
                      "mesh_convergence":"NOT_RUN","vector_maxwell":"NOT_IMPLEMENTED",
                      "complete_bandgap_certification":"NOT_VERIFIED",
                      "material_dispersion":"NOT_MODELED"}}

def defect_handoff(result, candidate_index, rings=4, defect_cells=7, core_surround_um=0.15):
    """Generate only triangular-compatible design for existing finite vector FEM.
    Requires an actual sampled gap candidate, not a guessed spectral interval.
    """
    if result.get("schema")!=SCHEMA_OUTPUT:
        raise ValueError("STEP5 computed result schema required")
    inp=result["input"]
    if inp["lattice"]!="triangular":
        raise NotImplementedError("Full-section FEM handoff supports triangular holes only; other lattices require CAD integration")
    gaps=result["gap_map"]["candidates"]
    if isinstance(candidate_index,bool) or not isinstance(candidate_index,int) or not 0<=candidate_index<len(gaps):
        raise ValueError("select an existing gap candidate index")
    if defect_cells not in (7,19) or not 2<=int(rings)<=12 or int(rings)!=rings:
        raise ValueError("defect_cells must be 7/19, rings 2..12")
    if not 0<=core_surround_um<=inp["pitch_um"]:
        raise ValueError("invalid core_surround_um")
    # Existing finite-section pbgf geometry is triangular hole lattice with missing 7/19 cells.
    # Import as normal design JSON, but modal confinement still needs independent V2/V4/V5.
    from schema import load_presets, validate_design
    gap=gaps[candidate_index]
    fmid=(gap["f_low"]+gap["f_high"])/2
    wavelength_nm=1000*inp["pitch_um"]/fmid
    if not 250 <= wavelength_nm <= 3500:
        raise ValueError("gap-center wavelength lies outside base design schema; choose pitch")
    base=copy.deepcopy(load_presets()["pbgf__"+str(defect_cells)+"_cell"])
    base["name"]="STEP5 proposed "+str(defect_cells)+"-cell triangular PBGF"
    base["wavelength_nm"]=wavelength_nm
    p=base["parameters"]
    p.update(pitch_um=inp["pitch_um"],
             d_over_pitch=inp["hole_diameter_over_pitch"],
             rings=int(rings), defect_cells=defect_cells,
             core_surround_um=float(core_surround_um))
    base["domain_radius_um"]=max(base["domain_radius_um"], (int(rings)+2)*inp["pitch_um"])
    base["materials"]["background"]={"model":"constant","n":inp["n_host"]}
    base["materials"]["silica"]={"model":"constant","n":inp["n_host"]}
    base["materials"]["air"]={"model":"constant","n":inp["n_air"]}
    base["boundary"]={"type":"pml","thickness_um":max(3.,inp["pitch_um"]*1.5),
                      "strength":2.,"order":3}
    base["metadata"]={"status":"proposal_not_validated",
                      "provenance":"STEP5 sampled scalar gap -> existing triangular finite pbgf approximation",
                      "source_input_sha256":result["input_sha256"],
                      "selected_gap":gap,
                      "fixed_beta_um_inverse":2*math.pi*gap["beta_norm"]/inp["pitch_um"],
                      "limitations":"no direct proof of true confinement, vector gap or matching full-section neff"}
    validate_design(base)
    return base

def example_config():
    return {"schema":SCHEMA_INPUT,"lattice":"triangular","pitch_um":3.0,
            "hole_diameter_over_pitch":0.42,"n_host":1.45,"n_air":1.0,
            "grid_n":20,"bands":5,"path_points_per_segment":4,"kmesh_n":3,
            "beta_norms":[.2,.4,.6],
            "input_basis":"illustrative constant indices, not validated silica Sellmeier or commercial fiber"}
