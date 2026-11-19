"""STEP 1 v0.8.4: geometric preflight and honest capability diagnostics.

No material coefficients or optical cutoff are inferred here. Lengths are um.
Numeric bounds are software limits, not manufacturer design rules.
"""
import math

ENGINE_VERSION = "0.8.4"


def lattice_centers(rings, pitch, defect_cells=1, square=False):
    omit = {1: 0, 7: 1, 19: 2}[defect_cells]
    if square:
        if defect_cells != 1:
            raise ValueError("square PCF currently supports a single missing cell only")
        return [(q*pitch, r*pitch) for q in range(-rings, rings+1)
                for r in range(-rings, rings+1) if q or r]
    return [(pitch*(q+r/2), pitch*math.sqrt(3)*r/2)
            for q in range(-rings, rings+1) for r in range(-rings, rings+1)
            if omit < max(abs(q), abs(r), abs(q+r)) <= rings]


def core_centers(n, pitch, kind):
    if kind == "ring":
        return [(pitch*math.cos(2*math.pi*i/n), pitch*math.sin(2*math.pi*i/n))
                for i in range(n)]
    result = [(0., 0.)]
    ring = 1
    while len(result) < n:
        points = [(pitch*(q+r/2), pitch*math.sqrt(3)*r/2)
                  for q in range(-ring, ring+1) for r in range(-ring, ring+1)
                  if max(abs(q), abs(r), abs(q+r)) == ring]
        result.extend(sorted(points, key=lambda xy: math.atan2(xy[1], xy[0])))
        ring += 1
    return result[:n]


def check_geometry(d):
    """Reject truncated/overlapping inputs before OCC; no silent clipping."""
    p = d["parameters"]
    f, v, R = d["family"], d["variant"], d["domain_radius_um"]
    needed = {"background"}

    def inside(x, y, radius):
        # Keep the legacy disk() clearance, now checked before meshing.
        if math.hypot(x, y)+radius >= .99*R:
            raise ValueError("geometry inclusion outside physical domain (1% clearance required)")

    if f == "index_guiding":
        needed.add("core")
        a = p["core_radius_um"]
        inside(0, 0, a)
        if v in ("trench_assisted", "w_profile"):
            needed.add("trench")
            lo, hi = p.get("trench_inner_um", 1.3*a), p.get("trench_outer_um", 1.8*a)
            if not a < lo < hi:
                raise ValueError("trench requires core_radius < trench_inner < trench_outer")
            inside(0, 0, hi)
        if v in ("graded", "multilayer"):
            needed.add("silica")
    elif f in ("solid_pcf", "pbgf"):
        needed.add("air")
        defect = int(p["defect_cells"])
        if defect not in (1, 7, 19):
            raise ValueError("defect_cells must be 1, 7 or 19")
        if f == "pbgf" and defect != {"7_cell": 7, "19_cell": 19}[v]:
            raise ValueError("PBGF variant and defect_cells disagree")
        pitch, rr = p["pitch_um"], p["pitch_um"]*p["d_over_pitch"]/2
        ey = p.get("ellipse_ratio", 1)
        pts = lattice_centers(int(p["rings"]), pitch, defect, v == "square")
        if not pts:
            raise ValueError("no cladding holes remain: increase rings beyond defect")
        for x, y in pts:
            inside(x, y, max(rr, rr*ey))
        # The same axis-aligned ellipses share their aspect ratio. Test lattice neighbours
        # in scaled coordinates instead of a circular bounding-box approximation.
        span=2*int(p["rings"])
        vectors=[(q*pitch,r*pitch) if v=="square" else
                 (pitch*(q+r/2),pitch*math.sqrt(3)*r/2)
                 for q in range(-span,span+1) for r in range(-span,span+1) if q or r]
        if any((x/(2*rr))**2+(y/(2*rr*ey))**2 <= 1 for x, y in vectors):
            raise ValueError("PCF air holes touch or overlap")
        if f == "pbgf":
            needed.add("silica")
            core = (1 if defect == 7 else 2)*pitch*.87
            surround = core+p["core_surround_um"]
            inside(0, 0, surround)
            if any(math.hypot(x, y)-rr <= surround for x, y in pts):
                raise ValueError("PBGF core surround overlaps cladding holes")
    elif f in ("anti_resonant", "kagome"):
        needed.add("silica")
        core = p["core_radius_um"]
        n = int(p.get("tube_count", 6) if f == "anti_resonant" else 8)
        tube = p.get("tube_radius_um", p.get("pitch_um", 3)*1.8)
        wall = p["wall_thickness_um"]
        gap = p.get("gap_um", .6)
        distance = core+tube+gap
        inside(distance, 0, tube)
        if not 0 < wall < tube/2:
            raise ValueError("tube wall must lie in (0, tube radius / 2)")
        if 2*distance*math.sin(math.pi/n) <= 2*tube:
            raise ValueError("adjacent tubes touch or overlap; conjoined CAD is not implemented")
        if v in ("nanf", "dnanf"):
            nested = tube*p.get("nested_radius_ratio", .48)
            if nested >= tube-wall or .7*wall >= nested:
                raise ValueError("nested tube wall touches parent or closes its air interior")
            if v == "dnanf" and (nested*.48 >= nested-.7*wall or .5*wall >= nested*.48):
                raise ValueError("second nested tube wall touches parent or closes its interior")
    elif f == "bragg":
        needed.update(("air", "high", "low"))
        inside(0, 0, p["core_radius_um"]+p["layer_count"]*p["layer_pitch_um"])
    elif f == "mcf":
        needed.add("core")
        radius = p["core_radius_um"]
        if v == "trench_assisted":
            needed.add("trench")
            radius += p["trench_width_um"]
        pts = core_centers(int(p["core_count"]), p["pitch_um"], v)
        for x, y in pts:
            inside(x, y, radius)
        if any(math.dist(a, b) <= 2*radius for i, a in enumerate(pts) for b in pts[i+1:]):
            raise ValueError("MCF cores/trenches touch or overlap (ring pitch is a radius)")
    elif f == "polarization_maintaining":
        needed.add("core")
        inside(0, 0, p["core_radius_um"]*(max(1, p.get("ellipse_ratio", 1)) if v == "elliptic_core" else 1))
        if v != "elliptic_core":
            needed.add("stress_rod")
            inside(p["rod_offset_um"], 0, p["rod_radius_um"])
    missing = needed-set(d["materials"])
    if missing:
        raise ValueError("missing physical-domain materials: "+", ".join(sorted(missing)))


def design_diagnostics(d):
    """Describe modeled geometry and assumptions, without claiming optical validation."""
    f, v = d["family"], d["variant"]
    approximations = []
    if v in ("graded", "multilayer") and f == "index_guiding":
        approximations.append("legacy alternating constant-index rings; not a continuous graded-index profile or user-specified arbitrary stack")
    if f == "pbgf":
        approximations.append("7/19 missing lattice cells with circular core/surround surrogate; no bandgap certificate")
    if f == "kagome" or v in ("conjoined_tube", "negative_curvature", "bow_tie"):
        approximations.append("legacy circular/elliptical surrogate; exact named CAD topology TODO")
    if v in ("nanf", "dnanf"):
        approximations.append("concentric nested tubes; realistic tangent nested geometry TODO")
    inert = sorted(set(d["parameters"]) & {"delta_n", "trench_delta_n", "stress_delta_n", "graded_exponent", "curvature"})
    return {"engine_version": ENGINE_VERSION, "status": "DESIGN_ONLY_NOT_OPTICALLY_CERTIFIED",
            "geometry_fidelity": "approximate" if approximations else "parametric",
            "approximations": approximations,
            "parameters_not_used_as_material_coefficients_or_exact_shape": inert,
            "material_basis": "materials.* defines optical index; parameter delta_n does not override it",
            "assumptions": ["air model uses constant n=1.00027; not pressure/temperature calibrated",
                            "no implicit Ge/F Sellmeier or stress-rod composition"],
            "mesh_resolution_policy": d["mesh"].get("resolution_policy", "boundary_refined"),
            "optical_certification": False}
