"""Independent infinite-cladding analytical oracle; no solver imports.

BYU ECE 360, section 7.3, equations (7.77), (7.78), (7.82),
(7.91), (7.92), with surrounding epsilon replaced by epsilon_s.
Here a is the HALF thickness and mode numbering begins at zero.
"""
from dataclasses import dataclass
import numpy as np
from scipy.optimize import brentq


@dataclass(frozen=True)
class ExactMode:
    neff: float
    beta_per_um: float
    u: float
    w: float
    order: int
    a_um: float

    def field(self, x):
        x = np.asarray(x)
        a, u, w = self.a_um, self.u, self.w
        if self.order % 2 == 0:
            inside = np.cos(u * x / a)
            outside = np.cos(u) * np.exp(-w * np.maximum(np.abs(x) - a, 0) / a)
            norm2 = a * (1 + np.sin(2*u)/(2*u)) + a*np.cos(u)**2/w
        else:
            inside = np.sin(u * x / a)
            outside = np.sign(x) * np.sin(u) * np.exp(-w*np.maximum(np.abs(x)-a, 0)/a)
            norm2 = a * (1 - np.sin(2*u)/(2*u)) + a*np.sin(u)**2/w
        return np.where(np.abs(x) <= a, inside, outside) / np.sqrt(norm2)


def exact_modes(n_core, n_clad, width_um, wavelength_um, polarization):
    a = width_um / 2
    k0 = 2*np.pi/wavelength_um
    V = a*k0*np.sqrt(n_core**2-n_clad**2)
    rho = 1.0 if polarization == 'TE' else n_core**2/n_clad**2
    modes = []
    for order in range(int(np.ceil(2*V/np.pi))):
        lo = order*np.pi/2
        if lo >= V:
            break
        hi = min((order+1)*np.pi/2, V)
        # atan formulation is equivalent to tan/cot equations without poles.
        def phase(u):
            return u-lo-np.arctan2(rho*np.sqrt(max(V*V-u*u, 0)), u)
        u = brentq(phase, lo, hi, xtol=5e-15, rtol=1e-14)
        w = np.sqrt(V*V-u*u)
        beta = np.sqrt((k0*n_core)**2-(u/a)**2)
        modes.append(ExactMode(beta/k0, beta, u, w, order, a))
    return modes


def l2_field_error(solution, mode, exact):
    """8-point element quadrature of actual P1 field, phase/sign aligned."""
    x = solution.x_um
    nodes, weights = np.polynomial.legendre.leggauss(8)
    dx = np.diff(x)
    points = (x[:-1,None]+x[1:,None])/2 + dx[:,None]*nodes[None,:]/2
    values = mode.field[:-1,None]*(1-nodes[None,:])/2 + mode.field[1:,None]*(1+nodes[None,:])/2
    ref = exact.field(points)
    measure = dx[:,None]*weights[None,:]/2
    sign = 1 if np.sum(measure*values*ref) >= 0 else -1
    return float(np.sqrt(np.sum(measure*(sign*values-ref)**2) / np.sum(measure*ref**2)))
