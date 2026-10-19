"""Interface-conforming radial P1 FEM for circular scalar LP0m modes.

Independent discretization of (r psi')' + k0^2 n^2 r psi = beta^2 r psi.
Natural regularity at r=0, psi(R)=0. All lengths in micrometers.
Only circular step-index fibers; not a full-vector Maxwell solver.
"""
from dataclasses import replace
import numpy as np
from numpy.polynomial.legendre import leggauss
from scipy.sparse import coo_matrix
from scipy.sparse.linalg import eigsh
from fiber_modes import Fiber, C0


def solve_radial(fiber=Fiber(), core_elements=160, clad_elements=640, outer_radius_um=None):
    for count in (core_elements, clad_elements):
        if not isinstance(count, (int, np.integer)) or isinstance(count, bool) or count < 4:
            raise ValueError('element counts must be integers >=4')
    outer = fiber.domain_halfwidth_um if outer_radius_um is None else float(outer_radius_um)
    if not np.isfinite(outer) or outer <= fiber.radius_um:
        raise ValueError('outer radius must be finite and exceed core radius')
    a = fiber.radius_um
    r = np.r_[np.linspace(0, a, core_elements + 1),
              np.linspace(a, outer, clad_elements + 1)[1:]]
    x, w = leggauss(4)
    t = (x + 1) / 2
    phi = np.stack((1-t, t), axis=1)
    h = np.diff(r)
    rq = r[:-1, None] + h[:, None]*t
    weights = h[:, None]*w/2
    mass = np.einsum('eq,qi,qj->eij', weights*rq, phi, phi)
    stiff = ((r[1:]**2-r[:-1]**2)/(2*h**2))[:, None, None]*np.array([[1., -1.], [-1., 1.]])
    indices = np.r_[np.full(core_elements, fiber.core), np.full(clad_elements, fiber.clad)]
    k0 = 2*np.pi/fiber.wavelength_um
    local = k0**2*indices[:, None, None]**2*mass-stiff
    edges = np.stack((np.arange(len(h)), np.arange(len(h))+1), axis=1)
    rows = np.broadcast_to(edges[:, :, None], local.shape).ravel()
    cols = np.broadcast_to(edges[:, None, :], local.shape).ravel()
    def matrix(v):
        return coo_matrix((v.ravel(), (rows, cols)), shape=(len(r), len(r))).tocsr()[:-1, :-1]
    A, M = matrix(local), matrix(mass)
    val, vec = eigsh(A, k=1, M=M, sigma=(k0*fiber.core)**2, which='LM',
                     v0=np.ones(len(r)-1), tol=1e-11)
    beta2 = float(val[0]); q = vec[:, 0]
    residual = np.linalg.norm(A@q-beta2*(M@q))/np.linalg.norm(beta2*(M@q))
    field = np.r_[q, 0.]
    if field[0] < 0:
        field = -field
    fieldq = field[edges]@phi.T
    intensity = fieldq**2
    wt = 2*np.pi*weights*rq
    power = np.sum(wt*intensity)
    aeff = power**2/np.sum(wt*intensity**2)
    moment = np.sum(wt*intensity*rq**2)/power
    neff = np.sqrt(beta2)/k0 if beta2 > 0 else float('nan')
    tail = float(np.sum((wt*intensity)[r[:-1] >= .8*outer])/power)
    return dict(n_eff=float(neff), Aeff_um2=float(aeff), MFD_2nd_um=float(2*np.sqrt(2*moment)),
                residual=float(residual), guided=bool(fiber.clad < neff < fiber.core),
                outer_radius_um=outer, core_elements=core_elements, clad_elements=clad_elements,
                tail_power_fraction=tail, radius_um=r, field=field/np.sqrt(power))


def radial_dispersion(fiber=Fiber(), step_um=.005, core_elements=160, clad_elements=640):
    """Five-point spectral derivative on the SAME interface-conforming mesh.

    Requires step and mesh/domain convergence before interpreting a result.
    Constant delta_n is a model assumption, not a doped-silica Sellmeier law.
    """
    if not np.isfinite(step_um) or step_um <= 0:
        raise ValueError('step_um must be positive and finite')
    offsets = np.arange(-2, 3)
    values = np.array([solve_radial(replace(fiber, wavelength_um=fiber.wavelength_um+j*step_um),
                         core_elements, clad_elements)['n_eff'] for j in offsets])
    d2 = np.dot(np.array([-1., 16., -30., 16., -1.]), values)/(12*step_um**2)
    D = -fiber.wavelength_um/C0*1e12*d2
    return dict(D_ps_nm_km=float(D), step_um=step_um, neff=values.tolist())
