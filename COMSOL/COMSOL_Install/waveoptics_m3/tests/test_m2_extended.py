"""Additional independent checks beyond the frozen initial acceptance suite."""
import numpy as np
from scipy.optimize import root, brentq
from scipy.special import jv, jvp, kv, kvp, iv, ivp
from skfem import asm, BilinearForm
from skfem.helpers import dot, curl, grad
from .analytic_vector import FiberReference


def absorbing_fiber_reference(nc=1.5+.001j,ns=1.+.0002j):
    a,k0 = .3,2*np.pi
    def characteristic(z):
        u = a*k0*np.sqrt(nc*nc-z*z)
        w = a*k0*np.sqrt(z*z-ns*ns)
        p,q = jvp(1,u)/(u*jv(1,u)),kvp(1,w)/(w*kv(1,w))
        return (p+q)*(nc*nc*p+ns*ns*q)-z*z*(1/u**2+1/w**2)**2
    def fun(x):
        value = characteristic(x[0]+1j*x[1])
        return [value.real,value.imag]
    result = root(fun,[FiberReference().neff,.0007],tol=1e-11)
    assert result.success
    exact = result.x[0]+1j*result.x[1]
    assert abs(characteristic(exact)) < 1e-10
    return exact


def finite_pec_fiber_reference(radius):
    """Exact circular finite-cladding Maxwell solution; polygon FEM is separate.

    Ez uses K1-cD*I1 (Dirichlet at R), Hz uses K1-cN*I1 (Neumann at R).
    The 2x2 tangential interface determinant supplies beta. No FEM targets.
    """
    a,k0,nc,ns = .3,2*np.pi,1.5,1.
    def det(neff):
        beta = k0*neff
        qc2,qs2 = k0**2*nc**2-beta**2,k0**2*ns**2-beta**2
        qc,kap = np.sqrt(qc2),np.sqrt(-qs2)
        u,w = a*qc,a*kap
        dc = qc*jvp(1,u)/jv(1,u)
        cD = kv(1,kap*radius)/iv(1,kap*radius)
        cN = kvp(1,kap*radius)/ivp(1,kap*radius)
        dD = kap*(kvp(1,w)-cD*ivp(1,w))/(kv(1,w)-cD*iv(1,w))
        dN = kap*(kvp(1,w)-cN*ivp(1,w))/(kv(1,w)-cN*iv(1,w))
        coupling = beta/a*(1/qc2-1/qs2)
        return coupling**2-k0**2*(dc/qc2-dN/qs2)*(nc**2*dc/qc2-ns**2*dD/qs2)
    exact = brentq(det,1.1,1.3,xtol=5e-15)
    assert abs(det(exact)) < 1e-9
    return exact


def test_finite_boundary_analytic_oracle():
    infinite = FiberReference().neff
    errors = np.array([abs(finite_pec_fiber_reference(r)-infinite) for r in (1.,1.5,2.)])
    assert np.all(np.diff(errors) < 0)
    assert errors[-1] < 1e-7


def test_absorbing_step_index_fiber_exact_complex_root():
    from waveoptics.vector import fiber_mesh,solve_vector
    mesh = fiber_mesh(.3,3.,.015)
    nc,ns = 1.5+.001j,1.+.0002j
    n = np.where(mesh.region == 1,nc,ns)
    sol = solve_vector(mesh,n,wavelength_um=1.,target_neff=1.25,
                       neff_window=(1.0001,1.4999),candidates=8)
    exact = absorbing_fiber_reference(nc,ns)
    assert len(sol.modes) == 2
    for mode in sol.modes:
        assert abs(mode.neff-exact) < 2e-4
        assert abs(mode.neff.imag/exact.imag-1) < .002
        assert mode.pencil_residual < 1e-8
        assert mode.gauss_residual < 1e-8


def test_compatible_gradient_subspace_curl_free():
    from waveoptics.vector import rectangle_mesh,solve_vector
    sol = solve_vector(rectangle_mesh(1.2,.9,.12),1.5,wavelength_um=1.55,
                       target_neff=1.4,neff_window=(.85,1.5),candidates=8)
    bt,bz = sol.edge_basis,sol.scalar_basis
    @BilinearForm
    def edge_mass(u,v,w):
        return dot(u,v)
    @BilinearForm
    def cross(u,v,w):
        return dot(grad(u),v)
    from scipy.sparse.linalg import spsolve
    scalar = np.random.default_rng(51).standard_normal(bz.N)
    scalar[sol.boundary_scalar_dofs] = 0
    projected = spsolve(asm(edge_mass,bt),asm(cross,bz,bt)@scalar)
    field = bt.interpolate(projected)
    direct = bz.interpolate(scalar).grad
    error = np.sqrt(np.sum(np.sum(abs(field-direct)**2,axis=0)*bt.dx))
    assert error < 1e-11
    assert np.sqrt(np.sum(field.curl**2*bt.dx)) < 1e-10


def test_point_sampling_and_raw_export(tmp_path):
    from waveoptics.vector import rectangle_mesh,solve_vector
    sol = solve_vector(rectangle_mesh(1.2,.9,.06),1.5,wavelength_um=1.55,
                       target_neff=1.4,neff_window=(.85,1.5),candidates=8)
    xy,w,e,_ = sol.quadrature_fields(sol.modes[0])
    chosen = np.arange(0,xy.shape[1],127)
    sampled = sol.sample_electric(sol.modes[0],xy[:,chosen])
    assert np.max(abs(sampled-e[:,chosen])) < 1e-10
    filename = tmp_path/'modes.npz'
    sol.save_npz(filename)
    raw = np.load(filename)
    assert np.array_equal(raw['triangles'],sol.geometry.mesh.t)
    assert np.max(abs(raw['neff']-[m.neff for m in sol.modes])) == 0
    assert np.max(abs(raw['edge_coefficients'][0]-sol.modes[0].edge_coefficients)) == 0
