"""STEP 2 v0.8.4: measured mesh diagnostics, not optical convergence.

Wall checks sample straight radial normals of circular annuli. They do not certify
every point on a wall or an arbitrary curved/touching-wall topology.
"""
import math
import numpy as np

try:
    from .materials import index
except ImportError:
    from materials import index


def material_size_targets(design, names):
    wavelength = design["wavelength_nm"]*.001
    count = design["mesh"]["elements_per_wavelength"]
    targets = {}
    for name in names:
        mat = design["materials"]["background" if name == "pml" else name]
        n = index(mat, design["wavelength_nm"]).real
        targets[name] = {"n_real": float(n), "wavelength_in_material_um": wavelength/n,
                         "requested_edge_um": min(design["mesh"]["bulk_max_um"], wavelength/(n*count))}
    return targets


def segment_intervals(points, triangles, start, end):
    """Exact barycentric clipping of a segment against linear triangles; t in [0,1]."""
    xy = points[triangles]
    e1, e2 = xy[:, 1]-xy[:, 0], xy[:, 2]-xy[:, 0]
    determinant = e1[:, 0]*e2[:, 1]-e1[:, 1]*e2[:, 0]
    if np.any(np.abs(determinant) < 1e-25):
        raise ValueError("degenerate triangle in wall measurement")

    def bary(p):
        z = np.asarray(p)-xy[:, 0]
        u = (z[:, 0]*e2[:, 1]-z[:, 1]*e2[:, 0])/determinant
        v = (e1[:, 0]*z[:, 1]-e1[:, 1]*z[:, 0])/determinant
        return np.column_stack((1-u-v, u, v))

    b0 = bary(start)
    db = bary(end)-b0
    low, high = np.zeros(len(xy)), np.ones(len(xy))
    keep = np.ones(len(xy), dtype=bool)
    for k in range(3):
        positive, negative = db[:, k] > 1e-12, db[:, k] < -1e-12
        keep &= (positive | negative | (b0[:, k] >= -1e-10))
        bound = np.divide(-b0[:, k], db[:, k], out=np.zeros(len(xy)), where=positive | negative)
        low[positive] = np.maximum(low[positive], bound[positive])
        high[negative] = np.minimum(high[negative], bound[negative])
    keep &= high-low > 1e-9
    return sorted((float(a), float(b)) for a, b in zip(low[keep], high[keep]))


def _probe_summary(intervals):
    unique = []
    for a, b in intervals:
        if not unique or max(abs(a-unique[-1][0]), abs(b-unique[-1][1])) > 1e-8:
            unique.append((a, b))
    covered, last = 0., 0.
    for a, b in unique:
        covered += max(0., b-max(last, a))
        last = max(last, b)
    return len(unique), covered, max((b-a for a, b in unique), default=1.)


def measure_walls(points, triangles, material_ids, walls, requested, angles=24):
    xy = points[triangles]
    centres = xy.mean(axis=1)
    edges = np.max(np.linalg.norm(xy-np.roll(xy, 1, axis=1), axis=2), axis=1)
    ids = np.asarray(material_ids)
    reports = []
    for wall in walls:
        centre = np.asarray(wall["centre_um"])
        outer, thickness = wall["outer_radius_um"], wall["thickness_um"]
        radius = np.linalg.norm(centres-centre, axis=1)
        selected = (ids == wall["material"]) & (radius+edges >= outer-thickness) & (radius-edges <= outer)
        subset = triangles[selected]
        probes = []
        for i in range(angles):
            theta = 2*math.pi*(i+.371)/angles  # avoid most shared-edge radial lines
            normal = np.array([math.cos(theta), math.sin(theta)])
            start, end = centre+(outer-thickness)*normal, centre+outer*normal
            count, coverage, longest = _probe_summary(segment_intervals(points, subset, start, end))
            probes.append({"angle_deg": theta*180/math.pi, "crossed_triangles": count,
                           "covered_fraction": coverage, "longest_segment_um": longest*thickness})
        minimum = min(x["crossed_triangles"] for x in probes)
        coverage = min(x["covered_fraction"] for x in probes)
        reports.append(dict(wall, samples=angles, requested_elements=requested,
                            min_crossed_triangles=minimum, min_covered_fraction=coverage,
                            max_normal_segment_um=max(x["longest_segment_um"] for x in probes),
                            status="PASS_SAMPLED" if minimum >= requested and coverage >= .99 else "FAIL_SAMPLED",
                            probes=probes))
    return {"basis": "circular-wall radial segment / linear triangle intersection; sampled only",
            "coverage_minimum_policy": .99, "walls": reports,
            "status": ("NOT_APPLICABLE" if not reports else
                       "PASS_SAMPLED" if all(x["status"] == "PASS_SAMPLED" for x in reports) else "FAIL_SAMPLED")}


def measured_resolution(points, triangles, material_ids, targets):
    xy = points[triangles]
    h = np.max(np.linalg.norm(xy-np.roll(xy, 1, axis=1), axis=2), axis=1)
    ids = np.asarray(material_ids)
    result = {}
    # Gmsh target length is not a hard maximum. Expose both requested and actual;
    # 1.35 is an explicit engineering acceptance policy, not a physical constant.
    allowance = 1.35
    for name, target in targets.items():
        sizes = h[ids == name]
        limit = target["requested_edge_um"]
        result[name] = dict(target, measured_edge_max_um=float(sizes.max()),
                            max_target_ratio=float(sizes.max()/limit),
                            triangles_over_target=int(np.count_nonzero(sizes > limit)),
                            triangles_over_allowance=int(np.count_nonzero(sizes > allowance*limit)),
                            status="PASS_POLICY" if sizes.max() <= allowance*limit else "FAIL_POLICY")
    return {"max_edge_allowance": allowance, "materials": result,
            "status": "PASS_POLICY" if all(x["status"] == "PASS_POLICY" for x in result.values()) else "FAIL_POLICY",
            "scope": "phase wavelength resolution only; PML stretching and absorption skin depth not certified"}
