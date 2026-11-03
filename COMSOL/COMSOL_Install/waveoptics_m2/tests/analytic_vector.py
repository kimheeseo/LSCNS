"""Independent Maxwell references; never import the FEM implementation.

e^(i beta z-i omega t); lengths in micrometres.  The fiber reference is
the exact m=1 hybrid characteristic equation, NOT a scalar LP approximation.
"""
from dataclasses import dataclass
import numpy as np
from scipy.optimize import brentq
from scipy.special import jv, jvp, kv, kvp
from scipy.integrate import quad


def rectangle_spectrum(n=1.5, width=1.2, height=.9, wavelength=1.55,
                       window=(.85, 1.5)):
    out = []
    for m in range(8):
        for l in range(8):
            if m == l == 0:
                continue
            neff = np.sqrt(complex(n*n-(wavelength/2)**2*((m/width)**2+(l/height)**2)))
            if window[0] < neff.real < window[1]:
                out.append((f'TE{m}{l}', neff))
                if m and l:
                    out.append((f'TM{m}{l}', neff))
    return sorted(out, key=lambda p: -p[1].real)


@dataclass
class FiberReference:
    nc: float = 1.5
    ns: float = 1.0
    a: float = .3
    wavelength: float = 1.0

    @property
    def k0(self):
        return 2*np.pi/self.wavelength

    def characteristic(self, neff):
        u = self.a*self.k0*np.sqrt(self.nc**2-neff**2)
        w = self.a*self.k0*np.sqrt(neff**2-self.ns**2)
        p = jvp(1, u)/(u*jv(1, u))
        q = kvp(1, w)/(w*kv(1, w))
        return (p+q)*(self.nc**2*p+self.ns**2*q)-neff**2*(1/u**2+1/w**2)**2

    @property
    def neff(self):
        # V=2.107 < first J0 zero: a single HE11 polarization doublet.
        grid = np.linspace(self.ns+1e-5, self.nc-1e-5, 4001)
        roots = []
        for left, right in zip(grid[:-1], grid[1:]):
            if self.characteristic(left)*self.characteristic(right) < 0:
                roots.append(brentq(self.characteristic, left, right, xtol=5e-15))
        assert len(roots) == 1, 'Reference is deliberately limited to this single-mode case.'
        return roots[0]

    def electric(self, xy, neff=None, angle=0.):
        """One real linear polarization; arbitrary overall complex amplitude.

        Ez=f(r)cos(theta), Z0 Hz=B f(r)sin(theta).  Transverse fields follow
        E_t=i[beta grad(Ez)-k0 z_cross grad(Z0 Hz)]/(k0^2 n^2-beta^2).
        B is fixed by tangential E continuity at r=a.
        """
        if neff is None:
            neff = self.neff
        beta = self.k0*neff
        qc2 = self.k0**2*self.nc**2-beta**2
        qs2 = self.k0**2*self.ns**2-beta**2
        qc, kap = np.sqrt(qc2), np.sqrt(-qs2)
        u, w = qc*self.a, kap*self.a
        dc = qc*jvp(1,u)/jv(1,u)
        ds = kap*kvp(1,w)/kv(1,w)
        b = -beta/(self.a*self.k0)*(1/qc2-1/qs2)/(dc/qc2-ds/qs2)
        r = np.maximum(np.hypot(xy[0],xy[1]), 1e-14)
        theta = np.arctan2(xy[1],xy[0])-angle
        core = r < self.a
        f = np.empty_like(r)
        df = np.empty_like(r)
        f[core] = jv(1,qc*r[core])/jv(1,u)
        df[core] = qc*jvp(1,qc*r[core])/jv(1,u)
        f[~core] = kv(1,kap*r[~core])/kv(1,w)
        df[~core] = kap*kvp(1,kap*r[~core])/kv(1,w)
        q2 = np.where(core,qc2,qs2)
        er = 1j*(beta*df+self.k0*b*f/r)*np.cos(theta)/q2
        et = -1j*(beta*f/r+self.k0*b*df)*np.sin(theta)/q2
        # cylindrical axes rotate with the actual spatial angle, not polarization.
        actual = theta+angle
        return np.array([er*np.cos(actual)-et*np.sin(actual),
                         er*np.sin(actual)+et*np.cos(actual), f*np.cos(theta)])

    def area(self):
        neff = self.neff
        theta = np.arange(256)*2*np.pi/256
        def radial(r, power):
            e = self.electric(np.array([r*np.cos(theta),r*np.sin(theta)]), neff)
            intensity = np.sum(abs(e)**2,axis=0)
            return float(2*np.pi*r*np.mean(intensity**power))
        moments = [sum(quad(lambda r: radial(r,power),l,r,epsabs=1e-9,
                            epsrel=2e-10,limit=150)[0]
                       for l,r in [(0,self.a),(self.a,np.inf)]) for power in (1,2)]
        return moments[0]**2/moments[1]


def relative_field_error(actual, reference, weights):
    # Least-squares complex phase/amplitude alignment; references stay independent.
    factor = np.sum(np.conj(reference)*actual*weights)/np.sum(abs(reference)**2*weights)
    return float(np.sqrt(np.sum(abs(actual-factor*reference)**2*weights)/
                         np.sum(abs(actual)**2*weights)))
