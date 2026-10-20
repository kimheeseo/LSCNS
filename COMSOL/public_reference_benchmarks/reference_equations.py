"""Reference equations only. Never substitutes computed FEM results.
Circular vector characteristic follows continuity of longitudinal E/H at r=a.
All lengths: um. Infinite cladding. Real, isotropic, lossless material indices.
"""
import numpy as np
from scipy.optimize import brentq
from scipy.special import j0, j1, k0, k1, jv, jvp, kv, kvp, jn_zeros


def scalar_circular_fundamental(n_core, n_clad, radius_um, wavelength_um):
    V=2*np.pi*radius_um/wavelength_um*np.sqrt(n_core*n_core-n_clad*n_clad)
    # LP01 fundamental occupies the first J0 branch even when V exceeds cutoff.
    hi=min(V*(1-1e-10),jn_zeros(0,1)[0]*(1-1e-10))
    def fun(u):
        w=np.sqrt(V*V-u*u)
        return u*j1(u)/j0(u)-w*k1(w)/k0(w)
    u=brentq(fun,1e-8,hi,xtol=1e-13)
    neff=np.sqrt(n_core*n_core-(u*wavelength_um/(2*np.pi*radius_um))**2)
    return dict(neff=float(neff),u=float(u),V=float(V))


def vector_he11(n_core, n_clad, radius_um, wavelength_um):
    """HE11 first branch of the exact hybrid-mode characteristic equation.

    (A+B)(nc^2*A+ns^2*B)=neff^2*(1/u^2+1/w^2)^2,
    A=J1'(u)/(u*J1(u)), B=K1'(w)/(w*K1(w)).
    This independent analytical reduction is not a new vector FEM implementation.
    """
    V=2*np.pi*radius_um/wavelength_um*np.sqrt(n_core*n_core-n_clad*n_clad)
    def ne(u):return np.sqrt(n_core*n_core-(u*wavelength_um/(2*np.pi*radius_um))**2)
    def determinant(u):
        w=np.sqrt(V*V-u*u)
        A=jvp(1,u)/(u*jv(1,u));B=kvp(1,w)/(w*kv(1,w))
        return (A+B)*(n_core*n_core*A+n_clad*n_clad*B)-ne(u)**2*(1/u**2+1/w**2)**2
    grid=np.linspace(.01,min(V*(1-1e-7),jn_zeros(0,1)[0]),1200)
    brackets=[(a,b) for a,b in zip(grid[:-1],grid[1:]) if determinant(a)*determinant(b)<0]
    if len(brackets)!=1:raise RuntimeError(f'HE11 expected one first-branch root, got {len(brackets)}')
    u=brentq(determinant,*brackets[0],xtol=1e-13)
    # Scale-free residual avoids reporting cancellation in large factors as accuracy.
    w=np.sqrt(V*V-u*u);A=jvp(1,u)/(u*jv(1,u));B=kvp(1,w)/(w*kv(1,w))
    lhs=(A+B)*(n_core*n_core*A+n_clad*n_clad*B)
    rhs=ne(u)**2*(1/u**2+1/w**2)**2
    return dict(neff=float(ne(u)),u=float(u),V=float(V),characteristic_relative_residual=float(abs(lhs-rhs)/max(abs(lhs),abs(rhs))))
