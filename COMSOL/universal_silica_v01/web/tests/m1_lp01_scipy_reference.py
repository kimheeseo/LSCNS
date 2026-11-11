"""Independent cross-check of the browser M1 LP01 Bessel root.

Local use only: python tests/m1_lp01_scipy_reference.py
Requires numpy and scipy in the developer's validation environment.
No Python/Colab dependency is added to the deployed GitHub Pages site.
"""
from __future__ import annotations
import math
from scipy.optimize import brentq
from scipy.special import j0,j1,k0,k1

CASES = (
    ("comsol_step", 1550.0, 8.0, 1.4457, 1.4378, 1.4444018444092133),
    ("single_mode", 1550.0, 4.1, 1.4472, 1.4440, 1.4448631289955670),
)
COMSOL_PUBLIC_ROUNDED=1.4444

def lp01(wavelength_nm:float, a_um:float, n_core:float, n_clad:float):
    k0=2*math.pi/(wavelength_nm/1000)
    v=k0*a_um*math.sqrt(n_core*n_core-n_clad*n_clad)
    limit=min(v,2.404825557695773)*(1-1e-12)
    def equation(u):
        w=math.sqrt(v*v-u*u)
        return u*j1(u)/j0(u)-w*k1(w)/k0(w)
    u=brentq(equation,1e-10,limit,xtol=1e-13,rtol=1e-14)
    n_eff=math.sqrt(n_core*n_core-(u/(k0*a_um))**2)
    return {"V":v,"u":u,"w":math.sqrt(v*v-u*u),"n_eff":n_eff,
            "characteristic_residual":equation(u)}

if __name__=="__main__":
    for name,lam,a,n1,n2,expected in CASES:
        actual=lp01(lam,a,n1,n2)
        delta=abs(actual["n_eff"]-expected)
        assert delta<=1e-10,(name,delta)
        print(f"{name}: neff={actual['n_eff']:.13f}, V={actual['V']:.10f}, |Δ|={delta:.3e} [PASS]")
        if name=="comsol_step":
            error=abs(actual["n_eff"]-COMSOL_PUBLIC_ROUNDED)
            assert error<=5e-5
            print(f"COMSOL rounded comparison: |Δ|={error:.3e} [within 4-decimal rounding interval]")
