"""M2 acceptance tests written before any M2 solver code.

Frozen thresholds: rectangle |delta neff|<3e-4; fiber <2e-4;
area relative error <2%; field relative L2 <3%; pencil/Gauss residual <1e-8.
No tolerance may be increased to make an implementation pass.
"""
import numpy as np
import pytest
from .analytic_vector import rectangle_spectrum, FiberReference, relative_field_error


def rectangle(h=.03, n=1.5, candidates=10, window=(.85,1.5)):
    from waveoptics.vector import rectangle_mesh, solve_vector
    mesh = rectangle_mesh(1.2,.9,h)
    return solve_vector(mesh,n,wavelength_um=1.55,target_neff=1.4,
                        neff_window=window,candidates=candidates)


def fiber(h=.015, radius=3., candidates=8):
    from waveoptics.vector import fiber_mesh, solve_vector
    mesh = fiber_mesh(.3,radius,h,outer_factor=5.)
    n = np.where(mesh.region == 1,1.5,1.0)
    return solve_vector(mesh,n,wavelength_um=1.,target_neff=1.25,
                        neff_window=(1.0001,1.4999),candidates=candidates)


def test_reference_rectangle_closed_form():
    refs = rectangle_spectrum()
    assert [p[0] for p in refs] == ['TE10','TE01','TE11','TM11']
    assert abs(refs[0][1]**2-(1.5**2-(1.55/2.4)**2)) < 1e-14


def test_reference_fiber_interface_and_area():
    ref = FiberReference()
    neff = ref.neff
    assert abs(ref.characteristic(neff)) < 1e-10
    t = np.arange(64)*2*np.pi/64
    inner = ref.electric(np.array([(.3-1e-9)*np.cos(t),(.3-1e-9)*np.sin(t)]),neff)
    outer = ref.electric(np.array([(.3+1e-9)*np.cos(t),(.3+1e-9)*np.sin(t)]),neff)
    tangent = np.array([-np.sin(t),np.cos(t)])
    normal = np.array([np.cos(t),np.sin(t)])
    assert np.max(abs(np.sum((inner[:2]-outer[:2])*tangent,axis=0))) < 2e-7
    assert np.max(abs(inner[2]-outer[2])) < 2e-7
    assert np.max(abs(np.sum((1.5**2*inner[:2]-outer[:2])*normal,axis=0))) < 2e-7
    assert .1 < ref.area() < 2.


def test_rectangle_spectrum_no_extra_modes():
    sol = rectangle()
    refs = rectangle_spectrum()
    assert len(sol.modes) == len(refs) == 4
    for mode,(_,exact) in zip(sol.modes,refs):
        assert abs(mode.neff-exact) < 3e-4
        assert mode.pencil_residual < 1e-8
        assert mode.gauss_residual < 1e-8
        assert abs(mode.neff.imag) < 1e-10
        assert np.all(mode.edge_coefficients[sol.boundary_edge_dofs] == 0)
        assert np.all(mode.scalar_coefficients[sol.boundary_scalar_dofs] == 0)


def test_rectangle_fields_area_magnetic():
    sol = rectangle()
    m = sol.modes[0]
    xy, weights, e, hbar = sol.quadrature_fields(m)
    x = xy[0]
    expected = np.array([np.zeros_like(x),np.sin(np.pi*x/1.2),np.zeros_like(x)])
    assert relative_field_error(e,expected,weights) < .03
    assert abs(m.effective_area_um2/(2*1.2*.9/3)-1) < .02
    assert abs(np.sum(np.sum(abs(e)**2,axis=0)*weights)-1) < 1e-10
    # TE10: Z0 Hx=-neff Ey; Z0 Hz=-i/k0 d(Ey)/dx.
    assert np.sqrt(np.sum(abs(hbar[0]+m.neff*e[1])**2*weights)) < 1e-8
    assert np.sqrt(np.sum(abs(e[2])**2*weights)) < 1e-6


def test_rectangle_three_mesh_levels():
    sols = [rectangle(h) for h in (.12,.06,.03)]
    exact = rectangle_spectrum()[0][1]
    neff_errors = np.array([abs(s.modes[0].neff-exact) for s in sols])
    field_errors = []
    for s in sols:
        xy,w,e,_ = s.quadrature_fields(s.modes[0])
        f = np.array([np.zeros_like(xy[0]),np.sin(np.pi*xy[0]/1.2),np.zeros_like(xy[0])])
        field_errors.append(relative_field_error(e,f,w))
    for errors,limits in [(neff_errors,(1.8,2.2)),(np.array(field_errors),(.8,1.2))]:
        assert np.all(np.diff(errors) < 0)
        p = np.log(errors[:-1]/errors[1:])/np.log(2)
        assert np.all((p > limits[0]) & (p < limits[1]))


def test_complex_index_loss_against_exact():
    n = 1.5+.001j
    sol = rectangle(n=n)
    refs = rectangle_spectrum(n=n)
    assert len(sol.modes) == 4
    for mode,(_,exact) in zip(sol.modes,refs):
        assert abs(mode.neff-exact) < 3e-4
        assert abs(mode.neff.imag/exact.imag-1) < .002
        loss = (20/np.log(10))*(2*np.pi/1.55)*exact.imag*1e6
        assert abs(mode.loss_db_per_m/loss-1) < .002
        assert mode.loss_db_per_m > 0


def test_fiber_exact_vector_pair_and_area():
    sol = fiber()
    ref = FiberReference()
    assert len(sol.modes) == 2
    for mode in sol.modes:
        assert abs(mode.neff-ref.neff) < 2e-4
        assert abs(mode.effective_area_um2/ref.area()-1) < .02
        assert mode.pencil_residual < 1e-8
        assert mode.gauss_residual < 1e-8


def test_fiber_vector_field_subspace():
    # The HE11 pair is degenerate: compare to the complete analytic polarization
    # subspace rather than arbitrarily choosing one numerical eigenvector angle.
    sol = fiber()
    ref = FiberReference()
    for mode in sol.modes:
        xy,w,e,_ = sol.quadrature_fields(mode)
        f = np.stack([ref.electric(xy,ref.neff,angle) for angle in (0,np.pi/2)])
        mat = (f*np.sqrt(w)[None,None,...]).reshape(2,-1).T
        target = (e*np.sqrt(w)).ravel()
        coeff = np.linalg.lstsq(mat,target,rcond=None)[0]
        assert np.linalg.norm(target-mat@coeff)/np.linalg.norm(target) < .03


def test_fiber_three_mesh_levels():
    sols = [fiber(h) for h in (.06,.03,.015)]
    ref = FiberReference()
    errors = np.array([max(abs(m.neff-ref.neff) for m in s.modes) for s in sols])
    assert np.all(np.diff(errors) < 0)
    slopes = np.log(errors[:-1]/errors[1:])/np.log(2)
    assert np.all((slopes > 1.5) & (slopes < 2.5))


def test_fiber_outer_boundary_convergence():
    sols = [fiber(h=.02,radius=r) for r in (2.,2.5,3.)]
    ref = FiberReference()
    for s in sols:
        assert len(s.modes) == 2
        assert max(abs(m.neff-ref.neff) for m in s.modes) < 4e-4
    # Compare errors to the same infinite-cladding reference; remeshing changes
    # element shapes, so this is a practical bound, not a pure truncation oracle.
    means = [np.mean([m.neff.real for m in s.modes]) for s in sols]
    assert max(abs(np.diff(means))) < 5e-5


def test_arbitrary_index_callable_and_array_agree():
    from waveoptics.vector import rectangle_mesh, solve_vector
    mesh = rectangle_mesh(1.2,.9,.06)
    kwargs = dict(wavelength_um=1.55,target_neff=1.4,neff_window=(.85,1.5),candidates=8)
    a = solve_vector(mesh,lambda x: 1.5+0*x[0],**kwargs)
    b = solve_vector(mesh,np.full(mesh.mesh.nelements,1.5),**kwargs)
    assert np.max(abs(np.array([m.neff for m in a.modes])-
                         np.array([m.neff for m in b.modes]))) < 1e-10


@pytest.mark.parametrize('changes',[
    {'wavelength_um':0}, {'target_neff':0}, {'neff_window':(1.5,.8)},
    {'candidates':0}, {'candidates':2.5}, {'n':1.5-.01j},
    {'n':0}, {'n':float('nan')}, {'n':np.ones(3)},
])
def test_invalid_inputs(changes):
    from waveoptics.vector import rectangle_mesh, solve_vector
    mesh = rectangle_mesh(1.2,.9,.1)
    args = dict(n=1.5,wavelength_um=1.55,target_neff=1.4,
                neff_window=(.85,1.5),candidates=8)
    args.update(changes)
    with pytest.raises(ValueError):
        solve_vector(mesh,**args)
