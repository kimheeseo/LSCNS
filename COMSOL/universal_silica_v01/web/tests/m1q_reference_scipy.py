"""M1 v0.5.1 independent step-index LP01 roots (development-only).

Run from web/: python tests/m1q_reference_scipy.py
Requires SciPy; browser app has no Python/Colab/server dependency.
This checks stored *analytic* reference values, not COMSOL vector-mode neff.
"""
from __future__ import annotations
import math
from scipy.special import j0,j1,k0,k1
from scipy.optimize import brentq

CASES=(
    ("official_comsol_145",1550.,8.,1.4378,.0079,1.4444018444092133),
    ("requested_41",1550.,4.1,1.4440,.005,1.446081132401304),
    ("low_v",1550.,4.1,1.4440,.0032,1.444863128995567),
    ("different_wavelength",1310.,4.1,1.4440,.005,1.446608620045317),
)
def lp01_neff(wavelength_nm,a_um,n_clad,delta_n):
    n_core=n_clad+delta_n
    k0=2*math.pi/(wavelength_nm/1000.)
    v=k0*a_um*math.sqrt(n_core*n_core-n_clad*n_clad)
    def characteristic(u):
        w=math.sqrt(v*v-u*u)
        return u*j1(u)/j0(u)-w*k1(w)/k0(w)
    # Independent SciPy Brent bracketing; root tolerance 1e-13 in u.
    upper=min(v,2.404825557695773)*(1-1e-9)
    u=brentq(characteristic,1e-8,upper,xtol=1e-13,rtol=1e-14,maxiter=100)
    neff=math.sqrt(n_core*n_core-(u/(k0*a_um))**2)
    return neff,v,characteristic(u)

if __name__=="__main__":
    for name,w,a,ncl,dn,reference in CASES:
        actual,v,res=lp01_neff(w,a,ncl,dn)
        difference=abs(actual-reference)
        assert difference<=1e-10,(name,difference)
        assert abs(res)<=1e-9,(name,res)
        print(f"{name:22s}  V={v:.9f}  LP01_neff={actual:.13f}  diff={difference:.3e} PASS")
    print("PASS 4/4 independent SciPy root references.")
