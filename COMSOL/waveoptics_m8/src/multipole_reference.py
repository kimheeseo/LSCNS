"""Independent scalar two-circle boundary matching in an infinite cladding.

Own-core exterior = sum b_m K_m(w*r)/K_m(w*a) cos(m*theta).
Interior logarithmic radial derivative = q*J_m'(q*a)/J_m(q*a).
Fourier-project continuity of the normal derivative on one circle, after
eliminating the interior amplitude using continuity of the field.
Reflection: b_right,m = parity*(-1)**m*b_left,m. Only y-even modes.
No FEM mesh, finite box, vector polarization, bending or random coupling.
"""
import numpy as np
from scipy.special import jv,jvp,kv,kvp
from scipy.optimize import brentq


def validate(fiber, pitch_um):
    if not np.isfinite(pitch_um) or pitch_um <= 2*fiber.radius_um:
        raise ValueError('pitch must exceed two core radii')
    if not 0 < fiber.v < 2.404825557695773:
        raise ValueError('benchmark restricted to isolated-core V<2.405')


def boundary_matrix(neff, fiber, pitch_um=16., order=6, parity=1, samples=256):
    validate(fiber,pitch_um)
    if not isinstance(order,(int,np.integer)) or isinstance(order,bool) or not 0 <= order <= 20:
        raise ValueError('multipole order must be integer 0..20')
    if parity not in (-1,1) or samples < 8*(order+1):
        raise ValueError('invalid parity or angular sampling')
    if not fiber.clad < neff < fiber.core:
        raise ValueError('reference root must be guided')
    a=fiber.radius_um; k=2*np.pi/fiber.wavelength_um
    q=k*np.sqrt(fiber.core**2-neff**2); w=k*np.sqrt(neff**2-fiber.clad**2)
    m=np.arange(order+1); theta=2*np.pi*np.arange(samples)/samples
    ct,st=np.cos(theta),np.sin(theta)
    # Left circle centered at 0; right circle at (+pitch, 0).
    dx=a*ct-pitch_um; dy=a*st; rr=np.hypot(dx,dy); angle=np.arctan2(dy,dx)
    dr=(dx*ct+dy*st)/rr
    dt=(-dy*ct+dx*st)/rr**2
    cos=np.cos(theta[:,None]*m)
    other=kv(m,w*rr[:,None])/kv(m,w*a)*np.cos(angle[:,None]*m)
    other_derivative=(w*kvp(m,w*rr[:,None])/kv(m,w*a)*dr[:,None]*np.cos(angle[:,None]*m)
                      -kv(m,w*rr[:,None])/kv(m,w*a)*m*dt[:,None]*np.sin(angle[:,None]*m))
    projection=cos.T/samples
    projection[1:]*=2
    log_derivative=q*jvp(m,q*a)/jv(m,q*a)
    diagonal=w*kvp(m,w*a)/kv(m,w*a)-log_derivative
    matrix=np.diag(diagonal)+(projection@other_derivative-log_derivative[:,None]*(projection@other))*(parity*(-1.)**m)
    # Row scaling improves determinant conditioning and leaves its roots intact.
    return matrix/(1+np.abs(log_derivative[:,None]))


def solve_reference(fiber, pitch_um=16., order=6, samples=256):
    validate(fiber,pitch_um)
    roots=[]; singular=[]
    for parity in (1,-1):
        def fun(n):return float(np.linalg.det(boundary_matrix(n,fiber,pitch_um,order,parity,samples)))
        grid=np.linspace(fiber.clad+fiber.delta_n*1e-5, fiber.core-fiber.delta_n*1e-5,120)
        values=[fun(n) for n in grid]
        brackets=[(x,y) for x,y,u,v in zip(grid[:-1],grid[1:],values[:-1],values[1:]) if u*v<0]
        if len(brackets)!=1:
            raise RuntimeError(f'expected one y-even root for parity {parity}, found {len(brackets)}')
        n=brentq(fun,*brackets[0],xtol=5e-15,rtol=1e-15)
        roots.append(float(n))
        sv=np.linalg.svd(boundary_matrix(n,fiber,pitch_um,order,parity,samples),compute_uv=False)
        singular.append(float(sv[-1]/sv[0]))
    split=roots[0]-roots[1]
    if split<=0:raise RuntimeError('unresolved even/odd splitting')
    return dict(n_even=roots[0],n_odd=roots[1],index_split=split,
                coupling_length_mm=fiber.wavelength_um/(2*split)/1000,
                order=order,angular_samples=samples,boundary_relative_singular=singular,
                reference='independent scalar Fourier-Bessel matching; infinite cladding')
