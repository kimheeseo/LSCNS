"""Numerical regression tests with independent analytic/physical references."""
from pathlib import Path
import sys
import numpy as np
import pytest
ROOT=Path(__file__).resolve().parents[2]
for stage in (5,6,7,8):sys.path.insert(0,str(ROOT/f'waveoptics_m{stage}'/'src'))
from fiber_modes import Fiber, solve_modes, scalar_metrics, lp01_analytic, radial_lp01_aeff, dispersion_lp01
from radial_fem import solve_radial, radial_dispersion
from mcf_hcf import capillary_neff, capillary_dispersion
from m6_inverse_design import reflectance
from m5_spectral import spectrum

@pytest.mark.parametrize('key,value',[(k,v) for k in ('wavelength_um','radius_um','delta_n','domain_halfwidth_um','n_cladding') for v in (np.nan,np.inf,0,-1)])
def test_invalid_fiber(key,value):
    with pytest.raises(ValueError):Fiber(**{key:value})

@pytest.mark.parametrize('cores',[[],[(np.nan,0)],[(18,0)],[(0,)],[(0,0,0)]])
def test_invalid_cores(cores):
    with pytest.raises(ValueError):solve_modes(Fiber(),n=15,cores=cores)

@pytest.mark.parametrize('n',[14,16,15.5,True])
def test_invalid_mesh(n):
    with pytest.raises(ValueError):solve_modes(Fiber(),n=n)

def test_coating_default_is_nominal():
    n,d=1.2,.32
    expected=np.mean([np.mean(spectrum(np.linspace(1.3,1.65,81),[(n,d)],polarization=p,angle_deg=a)['R']) for p,a in [('TE',0),('TM',0),('TE',20),('TM',20)]])
    assert reflectance(n,d)==pytest.approx(expected,abs=1e-15)
    assert reflectance(n,d)<.001

@pytest.mark.parametrize('radius',[10.,15.,30.])
def test_capillary_exact_dispersion(radius):
    w,h=1.55,.01
    ne=lambda v:capillary_neff(v,radius_um=radius)
    d2=(-ne(w+2*h)+16*ne(w+h)-30*ne(w)+16*ne(w-h)-ne(w-2*h))/(12*h*h)
    assert capillary_dispersion(w,radius)==pytest.approx(-w/299792458*1e12*d2,rel=2e-6)

@pytest.mark.parametrize('args',[(np.nan,15,1.00027),(1.55,np.nan,1.00027),(1.55,15,np.nan)])
def test_invalid_capillary(args):
    with pytest.raises(ValueError):capillary_neff(args[0],radius_um=args[1],n_air=args[2])

@pytest.mark.parametrize('w,a,dn',[(1.31,3.5,.004),(1.55,4.1,.005),(1.625,5.,.003),(1.55,5.5,.002)])
def test_radial_independent_bessel(w,a,dn):
    f=Fiber(wavelength_um=w,radius_um=a,delta_n=dn,domain_halfwidth_um=50.)
    sol=solve_radial(f);ref=lp01_analytic(f);metrics=radial_lp01_aeff(f)
    assert sol['guided'] and sol['residual']<1e-9
    assert abs(sol['n_eff']-ref['n_eff'])<2e-7
    for name in ('Aeff_um2','MFD_2nd_um'):
        assert sol[name]==pytest.approx(metrics[name],rel=5e-4)

def test_quadrature_rectangle_constant_field():
    # Exact integral on a 2x1 rectangle: Aeff=2, centered <r^2>=(4+1)/12.
    xy=np.array([[0,0],[2,0],[2,1],[0,1.]])
    s=dict(xy=xy,triangles=np.array([[0,1,2],[0,2,3]]),fields=np.ones((4,1)),neff=[1.],residuals=[0.])
    m=scalar_metrics(s)
    assert m['Aeff_um2']==pytest.approx(2.,abs=1e-12)
    assert m['MFD_2nd_um']==pytest.approx(2*np.sqrt(2*5/12),abs=1e-12)
    assert m['centroid_um']==pytest.approx([1.,.5],abs=1e-12)

def test_quadrature_linear_field():
    # On unit right triangle psi=x: integral x^2=1/12, integral x^4=1/30.
    s=dict(xy=np.array([[0,0],[1,0],[0,1.]]),triangles=np.array([[0,1,2]]),fields=np.array([[0],[1.],[0]]),neff=[1.],residuals=[0.])
    assert scalar_metrics(s)['Aeff_um2']==pytest.approx(5/24,abs=1e-12)

def test_radial_dispersion_and_mesh_step_convergence():
    f=Fiber(domain_halfwidth_um=40.)
    vals=[radial_dispersion(f,step_um=h)['D_ps_nm_km'] for h in (.01,.005,.0025)]
    assert max(vals)-min(vals)<.002
    assert vals[-1]==pytest.approx(dispersion_lp01(),abs=.02)
    fine=radial_dispersion(f,core_elements=320,clad_elements=1280)['D_ps_nm_km']
    assert abs(fine-vals[1])<.02
