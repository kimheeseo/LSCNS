"""P1 FEM generalized eigenproblem; NO analytical dispersion implementation.

Coordinates and wavelength use micrometres, beta uses inverse micrometres.
TE solves E_y; TM solves H_y. Real positive isotropic n, mu_r=1 only.
"""
from dataclasses import dataclass
import warnings
import numpy as np
from scipy.linalg import eigh
from scipy.sparse.linalg import eigsh
from skfem import Basis, BilinearForm, ElementLineP1, MeshLine, asm
from .mesh import aligned_mesh


@dataclass(frozen=True)
class Mode:
    order: int
    polarization: str
    neff: float
    beta_per_um: float
    field: np.ndarray
    relative_residual: float
    boundary_decay_estimate: float


@dataclass(frozen=True)
class SlabSolution:
    x_um: np.ndarray
    modes: tuple[Mode, ...]
    h_max_um: float
    elements: int
    dofs: int
    mesh_backend: str
    mode_limit_reached: bool
    parameters: dict


def _positive_real(value, name):
    if isinstance(value, (bool, complex)) or np.iscomplexobj(value):
        raise ValueError(f'{name} must be a positive real scalar; complex materials are unsupported in M1.')
    try:
        result = float(value)
    except (TypeError, ValueError):
        raise ValueError(f'{name} must be a positive real scalar.') from None
    if not np.isfinite(result) or result <= 0:
        raise ValueError(f'{name} must be finite and positive.')
    return result


def solve_slab(*, n_core=1.50, n_clad=1.45, width_um=3.0,
               wavelength_um=1.55, padding_um=20.0, h_um=0.0375,
               polarization='TE', max_modes=6, mesh_backend='gmsh'):
    """Return guided modes sorted by descending n_eff, with L2 scalar fields.

    max_modes is the number of highest eigenpairs requested, NOT a guarantee
    of finding all modes. Increase it if mode_limit_reached is True. Finite
    Dirichlet cladding must also be verified for every new geometry.
    """
    values = {name:_positive_real(value,name) for name,value in
              [('n_core',n_core),('n_clad',n_clad),('width_um',width_um),
               ('wavelength_um',wavelength_um),('padding_um',padding_um),('h_um',h_um)]}
    n_core,n_clad,width_um,wavelength_um,padding_um,h_um = values.values()
    if n_core <= n_clad:
        raise ValueError('M1 requires n_core > n_clad > 0.')
    if polarization not in ('TE','TM'):
        raise ValueError("polarization must be 'TE' or 'TM'.")
    if mesh_backend not in ('gmsh','native'):
        raise ValueError("mesh_backend must be 'gmsh' or 'native'.")
    if isinstance(max_modes,bool) or not isinstance(max_modes,(int,np.integer)) or max_modes < 1:
        raise ValueError('max_modes must be a positive integer.')
    x,t = aligned_mesh(width_um,padding_um,h_um,mesh_backend)
    if len(x) < 4:
        raise ValueError('At least two interior degrees of freedom are required; decrease h_um.')
    basis = Basis(MeshLine(x[None,:],t),ElementLineP1(),intorder=4)
    mids = (x[t[0]]+x[t[1]])/2
    # Elementwise coefficients: never interpolate n across a jump.
    n2 = np.where(abs(mids) < width_um/2,n_core**2,n_clad**2)[:,None]
    if polarization == 'TE':
        p,q,r = np.ones_like(n2),n2,np.ones_like(n2)
    else:
        p,q,r = 1/n2,np.ones_like(n2),1/n2
    k0 = 2*np.pi/wavelength_um

    @BilinearForm
    def stiffness(u,v,w):
        return w.p*u.grad[0]*v.grad[0]

    @BilinearForm
    def weighted_mass(u,v,w):
        return w.coefficient*u*v

    K = asm(stiffness,basis,p=p)
    Q = asm(weighted_mass,basis,coefficient=q)
    R = asm(weighted_mass,basis,coefficient=r)
    L2 = asm(weighted_mass,basis,coefficient=np.ones_like(n2))
    interior = np.arange(1,len(x)-1)
    A = (-K+k0*k0*Q)[interior][:,interior].tocsr()
    B = R[interior][:,interior].tocsr()
    count = min(int(max_modes),len(interior))
    if count == len(interior):
        eigenvalues,eigenvectors = eigh(A.toarray(),B.toarray())
    else:
        # Shift is ABOVE the upper bound beta² <= k0² max(n²).
        shift = (k0*n_core)**2*(1+1e-6)
        eigenvalues,eigenvectors = eigsh(A,k=count,M=B,sigma=shift,which='LM',
            tol=1e-12,v0=np.random.default_rng(734).normal(size=len(interior)),maxiter=10000)
    modes = []
    for j in np.argsort(eigenvalues)[::-1]:
        beta2 = float(eigenvalues[j])
        if beta2 <= 0:
            continue
        beta = np.sqrt(beta2)
        neff = beta/k0
        # Discrete box radiation modes below cladding index are excluded.
        if not n_clad+1e-10 < neff < n_core:
            continue
        vector = eigenvectors[:,j]
        av,bv = A@vector,B@vector
        residual = np.linalg.norm(av-beta2*bv)/(np.linalg.norm(av)+abs(beta2)*np.linalg.norm(bv))
        field = np.zeros(len(x))
        field[interior] = vector
        field /= np.sqrt(field @ (L2 @ field))
        # A deterministic sign, not a physical phase or power normalization.
        if field[np.argmax(abs(field))] < 0:
            field = -field
        alpha = np.sqrt(beta2-(k0*n_clad)**2)
        estimate = float(np.exp(-alpha*padding_um))
        if estimate > 1e-6:
            warnings.warn(f'{polarization}{len(modes)}: estimated relative cladding decay '
                          f'{estimate:.3e}; increase padding and check domain convergence.',RuntimeWarning)
        modes.append(Mode(len(modes),polarization,float(neff),float(beta),field,float(residual),estimate))
    limit = len(modes) == count and count < len(interior)
    if limit:
        warnings.warn('All requested eigenpairs are guided; increase max_modes to check completeness.',RuntimeWarning)
    return SlabSolution(x,tuple(modes),float(np.max(np.diff(x))),len(x)-1,len(interior),
                        mesh_backend,limit,{**values,'polarization':polarization,'max_modes':int(max_modes)})
