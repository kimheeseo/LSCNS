"""Acceptance criteria frozen BEFORE the numerical solver was written.

No golden values originate from a numerical FEM solution. All targets use
closed-form special cases or infinite-cladding analytical dispersion roots.
"""
import numpy as np
import pytest
from .analytic_slab import exact_modes, l2_field_error


def solve(**kwargs):
    # Delayed import allows oracle tests to run before solver code exists.
    from waveoptics.slab import solve_slab
    return solve_slab(**kwargs)


BASE = dict(n_core=1.50, n_clad=1.45, width_um=3.0, wavelength_um=1.55,
            padding_um=20.0, max_modes=6)
STRONG = dict(n_core=3.45, n_clad=1.44, width_um=0.4, wavelength_um=1.55,
              padding_um=4.0, max_modes=6)


@pytest.mark.parametrize('pol', ['TE', 'TM'])
def test_analytical_oracle_closed_form(pol):
    # Construct u=w=pi/4 for TE, or w=u/rho for TM.
    nc, ns, wavelength = 1.5, 1.45, 1.55
    rho = 1 if pol == 'TE' else nc**2/ns**2
    u = np.pi/4
    w = u/rho
    a = np.hypot(u,w)/(2*np.pi/wavelength*np.sqrt(nc**2-ns**2))
    mode = exact_modes(nc, ns, 2*a, wavelength, pol)[0]
    expected = np.sqrt(nc**2-(u/a/(2*np.pi/wavelength))**2)
    assert abs(mode.u-u) < 2e-13
    assert abs(mode.w-w) < 2e-13
    assert abs(mode.neff-expected) < 2e-13


@pytest.mark.parametrize('pol', ['TE', 'TM'])
def test_neff_modes_fields_residual(pol):
    sol = solve(**BASE, polarization=pol, h_um=0.0375)
    refs = exact_modes(BASE['n_core'], BASE['n_clad'], BASE['width_um'], BASE['wavelength_um'], pol)
    assert len(sol.modes) == len(refs) == 2
    for mode, ref in zip(sol.modes, refs):
        assert abs(mode.neff-ref.neff) < 1e-5
        assert l2_field_error(sol, mode, ref) < 1e-3
        assert mode.relative_residual < 1e-9
        assert np.isrealobj(mode.field)
        assert mode.field[0] == mode.field[-1] == 0
        # True integral of the interpolated P1 scalar field, not nodal sum.
        f = mode.field
        norm2 = np.sum(np.diff(sol.x_um)*(f[:-1]**2+f[:-1]*f[1:]+f[1:]**2)/3)
        assert abs(norm2-1) < 1e-11
    x = sol.x_um
    for index, mode in enumerate(sol.modes):
        mirror = np.interp(-x, x, mode.field)
        assert np.max(np.abs(mode.field-(-1)**index*mirror)) < 1e-8


@pytest.mark.parametrize('pol', ['TE', 'TM'])
def test_three_level_convergence(pol):
    hs = [0.15, 0.075, 0.0375]
    refs = exact_modes(1.5, 1.45, 3.0, 1.55, pol)
    sols = [solve(**BASE, polarization=pol, h_um=h) for h in hs]
    actual_h = np.array([s.h_max_um for s in sols])
    for j, ref in enumerate(refs):
        for errors in (
            np.array([abs(s.modes[j].neff-ref.neff) for s in sols]),
            np.array([l2_field_error(s, s.modes[j], ref) for s in sols]),
        ):
            assert np.all(np.diff(errors) < 0)
            slopes = np.log(errors[:-1]/errors[1:])/np.log(actual_h[:-1]/actual_h[1:])
            assert np.all((slopes > 1.8) & (slopes < 2.2))


@pytest.mark.parametrize('pol', ['TE', 'TM'])
def test_high_contrast_slab(pol):
    sol = solve(**STRONG, polarization=pol, h_um=0.0025)
    refs = exact_modes(3.45, 1.44, 0.4, 1.55, pol)
    assert len(sol.modes) == len(refs) == 2
    for mode, ref in zip(sol.modes, refs):
        assert abs(mode.neff-ref.neff) < 2e-4
        assert l2_field_error(sol, mode, ref) < 1e-3
        assert mode.relative_residual < 1e-9


@pytest.mark.parametrize('pol', ['TE', 'TM'])
def test_boundary_truncation_separately(pol):
    # Fixed exact h=0.05: only padding changes, no mesh-accuracy confound.
    short = solve(**{**BASE, 'padding_um': 15.0}, polarization=pol, h_um=0.05)
    long = solve(**{**BASE, 'padding_um': 20.0}, polarization=pol, h_um=0.05)
    assert len(short.modes) == len(long.modes) == 2
    assert max(abs(a.neff-b.neff) for a,b in zip(short.modes,long.modes)) < 1e-10


@pytest.mark.parametrize('pol', ['TE', 'TM'])
def test_interface_flux(pol):
    # P1 one-sided derivatives approximate continuous physical flux to O(h).
    errors = []
    for h in (0.01, 0.005, 0.0025):
        sol = solve(**STRONG, polarization=pol, h_um=h)
        j = np.argmin(abs(sol.x_um-0.2))
        assert abs(sol.x_um[j]-0.2) < 1e-13
        f = sol.modes[0].field
        left = (f[j]-f[j-1])/(sol.x_um[j]-sol.x_um[j-1])
        right = (f[j+1]-f[j])/(sol.x_um[j+1]-sol.x_um[j])
        if pol == 'TM':
            left /= 3.45**2
            right /= 1.44**2
        errors.append(abs(left-right)/max(abs(left),abs(right)))
    assert errors[-1] < 0.03
    slopes = np.log(np.array(errors[:-1])/errors[1:])/np.log(2)
    assert np.all((slopes > 0.8) & (slopes < 1.2))


@pytest.mark.parametrize('pol', ['TE', 'TM'])
def test_single_mode_slab(pol):
    sol = solve(**{**BASE, 'width_um': 1.0}, polarization=pol, h_um=0.025)
    refs = exact_modes(1.5,1.45,1.0,1.55,pol)
    assert len(sol.modes) == len(refs) == 1
    assert abs(sol.modes[0].neff-refs[0].neff) < 1e-5


def test_te_tm_are_distinct():
    te = solve(**STRONG, polarization='TE', h_um=0.0025)
    tm = solve(**STRONG, polarization='TM', h_um=0.0025)
    assert te.modes[0].neff > tm.modes[0].neff
    assert te.modes[0].neff-tm.modes[0].neff > 0.1


def test_backend_equivalence():
    # gmsh must actually generate the aligned 1D mesh; native is cross-check.
    gm = solve(**BASE, polarization='TM', h_um=0.075, mesh_backend='gmsh')
    native = solve(**BASE, polarization='TM', h_um=0.075, mesh_backend='native')
    assert gm.mesh_backend == 'gmsh'
    assert np.allclose(gm.x_um,native.x_um,rtol=0,atol=2e-10)
    assert np.allclose([m.neff for m in gm.modes], [m.neff for m in native.modes], rtol=0,atol=1e-10)


@pytest.mark.parametrize('change', [
    {'n_core': 1.4}, {'n_clad': 0}, {'n_core': 1.5+0.01j},
    {'wavelength_um': 0}, {'width_um': -1}, {'padding_um': 0},
    {'h_um': 0}, {'h_um': float('nan')}, {'polarization':'TEM'},
    {'mesh_backend':'unknown'}, {'max_modes':0}, {'max_modes':1.5},
])
def test_invalid_input(change):
    with pytest.raises(ValueError):
        solve(**{**BASE, 'polarization':'TE', 'h_um':0.075, **change})
