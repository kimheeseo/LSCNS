"""Independent circular-interface Maxwell references (Bessel matching).

exp(i beta z-i omega t), passive Im(neff)>0. Nonmagnetic media.
No fitted parameters. Inputs are explicit refractive indices and lengths in um.
The exact HE11 characteristic is also used in the pre-existing full_vector asset;
this implementation takes arbitrary explicit core/cladding indices.
"""
import numpy as np
from scipy.optimize import brentq, root
from scipy.special import jv, jvp, kv, kvp, hankel1, h1vp


def he11_exact(n_core,n_clad,radius_um,wavelength_um):
    if not 0<n_clad<n_core or radius_um<=0 or wavelength_um<=0:raise ValueError("invalid guided step-index reference")
    k=2*np.pi/wavelength_um;V=k*radius_um*np.sqrt(n_core*n_core-n_clad*n_clad)
    if not 0<V<2.4048255577:raise ValueError("HE11 reference restricted to single-spatial-mode V<2.405")
    def equation(u):
        w=np.sqrt(V*V-u*u)
        a=jvp(1,u)/(u*jv(1,u));b=kvp(1,w)/(w*kv(1,w))
        ne2=n_core*n_core-(u/(k*radius_um))**2
        return (a+b)*(n_core*n_core*a+n_clad*n_clad*b)-ne2*(1/u**2+1/w**2)**2
    grid=np.linspace(V*1e-5,V*(1-1e-7),2000);values=[equation(u) for u in grid]
    roots=[brentq(equation,a,b,xtol=1e-13) for a,b,fa,fb in zip(grid[:-1],grid[1:],values[:-1],values[1:])
           if np.isfinite(fa) and np.isfinite(fb) and fa*fb<0]
    if not roots:raise RuntimeError("HE11 root not bracketed")
    u=min(roots)
    return {"neff_real":float(np.sqrt(n_core*n_core-(u/(k*radius_um))**2)),
            "V":float(V),"u":float(u),"characteristic_residual":float(equation(u)),
            "method":"exact vector HE11 J1/K1 tangential-interface characteristic; infinite cladding"}


def circular_outgoing_characteristic(neff,n_core,n_clad,radius_um,wavelength_um,azimuthal_order=1):
    """Four tangential E/H conditions reduced to J/Hankel logarithmic derivatives.

    Outgoing Hankel-1 is continued on the open radiation sheet Re(q_clad)>0.
    This function is intended for an anti-guiding core n_core<n_clad, not a
    guided K-Bessel exterior or an arbitrary multilayer capillary.
    """
    k=2*np.pi/wavelength_um;m=azimuthal_order
    u=k*radius_um*np.sqrt(n_core*n_core-neff*neff+0j)
    v=k*radius_um*np.sqrt(n_clad*n_clad-neff*neff+0j)
    a=jvp(m,u)/(u*jv(m,u));b=h1vp(m,v)/(v*hankel1(m,v))
    return (a-b)*(n_core*n_core*a-n_clad*n_clad*b)-neff*neff*m*m*(1/u**2-1/v**2)**2


def leaky_circular_exact(n_core,n_clad,radius_um,wavelength_um,guess):
    if not 0<n_core<n_clad:raise ValueError("outgoing reference expects anti-guiding n_core<n_clad")
    def fun(z):
        value=circular_outgoing_characteristic(complex(*z),n_core,n_clad,radius_um,wavelength_um)
        return [value.real,value.imag]
    sol=root(fun,[complex(guess).real,complex(guess).imag],tol=1e-11)
    value=complex(*sol.x);res=float(np.linalg.norm(fun(sol.x)))
    if not sol.success or res>1e-8 or not (0<value.real<n_core and value.imag>0):
        raise RuntimeError("outgoing circular reference failed: "+str(sol.message))
    return {"neff_real":value.real,"neff_imag":value.imag,"characteristic_residual":res,
            "method":"exact vector m=1 circular core, regular J1 interior / outgoing Hankel-1 infinite exterior"}
