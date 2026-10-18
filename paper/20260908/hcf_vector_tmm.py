"""Unfitted full-vector *concentric* Maxwell leaky-mode diagnostic.

Not a model of the actual 5-tube DNANF cross section and not total attenuation.
Convention exp(i beta z + i m phi - i omega t); only Im(beta)>0 is accepted.
Fields matched at each interface: Ez, Z0*Hz, Ephi, Z0*Hphi.
The geometry/container and Sellmeier inputs are retained from the user's TMM.
Reference: Bird, Opt. Express 25, 23215 (2017), Eqs. (4)-(8), (25).
Our complex exp(i*m*phi) basis is equivalent to Bird's sine/cosine basis.
Exact Bessel functions are used, not the asymptotic Hankel approximation.
"""
from __future__ import annotations

from dataclasses import dataclass, asdict
import numpy as np
from scipy import optimize, special

from hcf_concentric_tmm_20260908 import ConcentricLeakyModeModel, U01


def full_chord_model(wavelength_nm, core_radius_um=14.75,
                     large_um=31.05, middle_um=23.75, small_um=7.70,
                     wall_um=0.50):
    """Full radial chord of three internally tangent tubes (9 regions).

    The old 7-region surrogate stops after the small tube's *near* wall.
    This geometry-derived alternative retains its air bore and the three
    contiguous far walls (3*t). The internally tangent, equal-thickness
    assumption is explicit; the real 5-tube azimuthal geometry is still absent.
    Parameters and layer selection must never be tuned to attenuation data.
    """
    from hcf_concentric_tmm_20260908 import silica_index_sellmeier
    n=silica_index_sellmeier(wavelength_nm*1e-9)
    g1=large_um-2*wall_um-middle_um
    g2=middle_um-2*wall_um-small_um
    inner_bore=small_um-2*wall_um
    if min(g1,g2,inner_bore)<=0:
        raise ValueError("Nonpositive gap/bore in tangent geometry")
    return ConcentricLeakyModeModel("9-layer full-chord concentric diagnostic",
        core_radius_um,(n,1.,n,1.,n,1.,n),
        (wall_um,g1,wall_um,g2,wall_um,inner_bore,3*wall_um))


@dataclass(frozen=True)
class VectorSolution:
    wavelength_nm: float
    layers: int
    u_real: float
    u_imag: float
    loss_db_km: float
    neff_real: float
    neff_imag: float
    matching_smin_ratio: float
    outgoing_smin_ratio: float
    determinant_abs: float
    root_count: int
    converged: bool

    def to_dict(self):
        return asdict(self)


def _parameters(u, wavelength_nm, model, n):
    # Units in this solver are micrometres.  Compute q without cancellation.
    k = 2*np.pi/(wavelength_nm*1e-3)
    qc = u/model.core_radius_um
    q = np.sqrt(complex((n*n-model.n_core**2)*k*k + qc*qc))
    if q.real < 0:
        q = -q
    beta = np.sqrt(complex(model.n_core**2*k*k-qc*qc))
    if beta.real < 0:
        beta = -beta
    return k, q, beta


def _basis(u, wavelength_nm, model, n, radius_um, kind="both", m=1):
    """Maxwell basis; longitudinal E/H columns, no scalar approximation.

    Ephi = -beta*m/(q^2*r)*Ez - i*k0/q^2*d(Z0 Hz)/dr
    Z0 Hphi = i*k0*n^2/q^2*dEz/dr - beta*m/(q^2*r)*(Z0 Hz).
    This is the exact boundary-condition formulation, not Bird's asymptotic
    Hankel expansion or Bache's single-FEM-scale empirical extension.
    """
    k,q,beta = _parameters(u,wavelength_nm,model,n)
    x=q*radius_um
    if kind == "core":
        f=np.array([special.jv(m,x)]); fp=np.array([special.jvp(m,x)])
    elif kind == "outer":
        f=np.array([special.hankel1(m,x)]); fp=np.array([special.h1vp(m,x)])
    else:
        f=np.array([special.jv(m,x),special.yv(m,x)])
        fp=np.array([special.jvp(m,x),special.yvp(m,x)])
    count=len(f)
    B=np.zeros((4,2*count),complex)
    B[0,:count]=f; B[1,count:]=f
    cross=-beta*m/(q*q*radius_um)
    B[2,:count]=cross*f; B[2,count:]=-1j*k/q*fp
    B[3,:count]=1j*k*n*n/q*fp; B[3,count:]=cross*f
    return B


def _normalized(A):
    # Column rescaling changes no physical boundary condition.
    return A / np.maximum(np.linalg.norm(A,axis=0),1e-300)


def reduced_matching(u,wavelength_nm,model):
    """Propagate the two outgoing polarizations inward as a 4x2 subspace."""
    n=model.region_indices
    radii=model.boundary_radii_m*1e6
    out=_basis(u,wavelength_nm,model,n[-1],radii[-1],"outer")
    out=_normalized(out)
    for j in range(len(n)-2,0,-1):
        coeff=np.linalg.solve(_basis(u,wavelength_nm,model,n[j],radii[j]),out)
        out=_basis(u,wavelength_nm,model,n[j],radii[j-1]) @ coeff
        out=np.linalg.qr(out,mode="reduced")[0]
    core=_basis(u,wavelength_nm,model,n[0],radii[0],"core")
    return np.column_stack((_normalized(core),-out))


def outgoing_matching(u,wavelength_nm,model):
    """Independent outward transfer direction for cross-checking the root."""
    n=model.region_indices; r=model.boundary_radii_m*1e6
    core=_normalized(_basis(u,wavelength_nm,model,n[0],r[0],"core"))
    for j in range(1,len(n)-1):
        c=np.linalg.solve(_basis(u,wavelength_nm,model,n[j],r[j-1]),core)
        core=_basis(u,wavelength_nm,model,n[j],r[j])@c
        core=np.linalg.qr(core,mode="reduced")[0]
    out=_normalized(_basis(u,wavelength_nm,model,n[-1],r[-1],"outer"))
    return np.column_stack((core,-out))


def _smin(A):
    s=np.linalg.svd(A,compute_uv=False)
    return float(s[-1]/s[0])


def solve_vector_mode(wavelength_nm,model,seed=None,multistart=True):
    """Track the HE11-like root near u01, never minimize benchmark error."""
    def fun(x):
        try:
            d=np.linalg.det(reduced_matching(complex(*x),wavelength_nm,model))
            return [d.real,d.imag]
        except (np.linalg.LinAlgError,ValueError,FloatingPointError):
            return [1e3,1e3]
    seeds=[]
    if seed is not None:
        seeds.append(complex(seed))
    if multistart or seed is None:
        seeds += [complex(ur,ui) for ur in (U01-.06,U01,U01+.06)
                  for ui in (-.003,-.0001,-1e-6,-1e-8)]
    candidates=[]
    for guess in seeds:
        opt=optimize.root(fun,[guess.real,guess.imag],tol=1e-10,
                          options={"maxfev":600})
        u=complex(*opt.x)
        if not (1.8<u.real<3.0 and -0.3<u.imag<0):
            continue
        A=reduced_matching(u,wavelength_nm,model)
        s=_smin(A)
        # Do not turn an amplifying/failed root into a positive loss with abs.
        beta=_parameters(u,wavelength_nm,model,model.n_core)[2]
        if s<1e-9 and beta.imag>0 and not any(abs(u-v[0])<1e-7 for v in candidates):
            candidates.append((u,s,abs(np.linalg.det(A))))
    if not candidates:
        if seed is not None and not multistart:
            return solve_vector_mode(wavelength_nm,model,seed,True)
        raise RuntimeError(f"No validated HE11-like root at {wavelength_nm} nm, {model.name}")
    target=complex(seed) if seed is not None else U01
    u,s,d=min(candidates,key=lambda v:abs(v[0]-target))
    sout=_smin(outgoing_matching(u,wavelength_nm,model))
    k,_,b=_parameters(u,wavelength_nm,model,model.n_core)
    return VectorSolution(float(wavelength_nm),len(model.region_indices),
        float(u.real),float(u.imag),float(2*b.imag*10/np.log(10)*1e9),
        float((b/k).real),float((b/k).imag),s,sout,float(d),len(candidates),
        bool(s<1e-9 and sout<1e-6))


def high_precision_check(wavelength_nm,model,seed,dps=40):
    """Arbitrary-precision outward boundary determinant (no QR/rescaling).

    A separate implementation of the same Maxwell equations.  Numerical
    consistency only; this is NOT independent validation against DNANF FEM.
    """
    import mpmath as mp
    with mp.workdps(dps):
        k=2*mp.pi/mp.mpf(str(wavelength_nm*1e-3))
        a=mp.mpf(str(model.core_radius_um))
        ns=[mp.mpf(str(x)) for x in model.region_indices]
        rr=[mp.mpf(str(x)) for x in model.boundary_radii_m*1e6]
        def basis(u,n,r,kind):
            q=mp.sqrt((n*n-ns[0]**2)*k*k+(u/a)**2)
            beta=mp.sqrt(ns[0]**2*k*k-(u/a)**2)
            z=q*r
            fs=([lambda z:mp.besselj(1,z)] if kind=="core" else
                [lambda z:mp.hankel1(1,z)] if kind=="outer" else
                [lambda z:mp.besselj(1,z),lambda z:mp.bessely(1,z)])
            B=mp.matrix(4,2*len(fs))
            for j,f in enumerate(fs):
                v=f(z); vp=mp.diff(f,z); cross=-beta/(q*q*r)
                B[0,j]=v; B[1,j+len(fs)]=v
                B[2,j]=cross*v; B[2,j+len(fs)]=-1j*k/q*vp
                B[3,j]=1j*k*n*n/q*vp; B[3,j+len(fs)]=cross*v
            return B
        def det(u):
            C=basis(u,ns[0],rr[0],"core")
            for j in range(1,len(ns)-1):
                C=basis(u,ns[j],rr[j],"both")*(basis(u,ns[j],rr[j-1],"both")**-1)*C
            O=basis(u,ns[-1],rr[-1],"outer")
            A=mp.matrix(4,4)
            for i in range(4):
                for j in range(2):
                    A[i,j]=C[i,j]; A[i,j+2]=-O[i,j]
            return mp.det(A)
        z=mp.mpc(seed.real,seed.imag)
        root=mp.findroot(det,(z,z+mp.mpc("1e-7","1e-9")),
                         tol=mp.mpf(10)**(-dps+10),maxsteps=30)
        b=mp.sqrt(ns[0]**2*k*k-(root/a)**2)
        return {"dps":dps,"u_real":float(root.real),"u_imag":float(root.imag),
                "loss_db_km":float(2*b.imag*10/mp.log(10)*mp.mpf("1e9")),
                "raw_determinant_abs":float(abs(det(root)))}
