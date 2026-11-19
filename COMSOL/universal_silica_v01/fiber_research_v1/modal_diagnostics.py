"""Field-based independence and subspace tracking; no neff-only deduplication."""
import numpy as np


def weighted_columns(fields, weights):
    """fields shape (points, components, modes); quadrature weights in um²."""
    f=np.asarray(fields,dtype=complex);w=np.asarray(weights,dtype=float)
    if f.ndim!=3 or w.shape!=(f.shape[0],) or not np.isfinite(f).all() or not np.isfinite(w).all() or np.any(w<=0):
        raise ValueError("finite field tensor and positive matching weights required")
    return (f*np.sqrt(w)[:,None,None]).reshape(-1,f.shape[-1])


def independent_indices(columns, tolerance=1e-7):
    """Rank-revealing reorthogonalized Gram-Schmidt; exact degeneracy is retained."""
    x=np.asarray(columns,dtype=complex)
    if x.ndim!=2 or not np.isfinite(x).all():raise ValueError("finite matrix required")
    chosen=[];basis=[]
    for i in range(x.shape[1]):
        z=x[:,i].copy();norm=np.linalg.norm(z)
        if norm==0:continue
        z/=norm
        for _ in range(2):
            for b in basis:z-=b*np.vdot(b,z)
        length=np.linalg.norm(z)
        if length>tolerance:
            chosen.append(i);basis.append(z/length)
    return chosen


def principal_correlations(reference, candidate, rank_tolerance=1e-8):
    """Whiten both field subspaces before computing principal-angle cosines."""
    def orth(x):
        x=np.asarray(x,dtype=complex)
        if x.ndim!=2 or not np.isfinite(x).all():raise ValueError("finite subspace required")
        u,s,_=np.linalg.svd(x,full_matrices=False)
        return u[:,s>rank_tolerance*s[0]] if len(s) and s[0]>0 else u[:,:0]
    a,b=orth(reference),orth(candidate)
    if a.shape[0]!=b.shape[0] or a.shape[1]!=b.shape[1] or not a.shape[1]:return []
    return np.clip(np.linalg.svd(a.conj().T@b,compute_uv=False),0,1).tolist()


def disk_probes(radius_um, core_radius_um=None, radial_order=12, angular_points=24):
    """Polar Gauss-Legendre points, identical across independently remeshed solves."""
    edges=[0.]
    if core_radius_um is not None and 0<core_radius_um<radius_um:edges.append(float(core_radius_um))
    edges.append(float(radius_um))
    q,w=np.polynomial.legendre.leggauss(radial_order)
    points=[];weights=[]
    for lo,hi in zip(edges[:-1],edges[1:]):
        for r,wr in zip(lo+(q+1)*(hi-lo)/2,w*(hi-lo)/2):
            for j in range(angular_points):
                theta=2*np.pi*(j+.271)/angular_points
                points.append([r*np.cos(theta),r*np.sin(theta)]);weights.append(wr*r*2*np.pi/angular_points)
    return np.asarray(points),np.asarray(weights)
