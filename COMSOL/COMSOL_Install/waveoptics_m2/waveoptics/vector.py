"""Full-vector 2D waveguide modes with a compatible Nedelec/P1 pair.

Convention exp(i*beta*z-i*omega*t); E=(e_t,-i*beta*phi).  Lengths are um.
PEC exterior; no PML and no radiation-loss claim.  Material n may be a scalar,
one value per triangle, or callable n(x) at common integration points.
"""
from dataclasses import dataclass, field
import math
import warnings
import numpy as np
from scipy.sparse import bmat, csr_matrix
from scipy.sparse.linalg import LinearOperator, eigs, splu
from skfem import MeshTri, Basis, ElementTriN1, ElementTriP1, BilinearForm, asm
from skfem.helpers import dot, grad, curl


@dataclass
class VectorMesh:
    mesh: MeshTri
    region: np.ndarray
    metadata: dict = field(default_factory=dict)


def _positive(value, name):
    if not np.isscalar(value) or np.iscomplexobj(value) or not np.isfinite(value) or value <= 0:
        raise ValueError(f'{name} must be a finite positive real number.')
    return float(value)


def _gmsh_start(name):
    import gmsh
    if gmsh.isInitialized():
        raise RuntimeError('gmsh is already initialized; run meshing in a separate process.')
    gmsh.initialize([name,'-nopopup'],readConfigFiles=False)
    gmsh.option.setNumber('General.Terminal',0)
    gmsh.option.setNumber('General.NumThreads',1)
    gmsh.model.add(name)
    return gmsh


def _read_gmsh(gmsh, regions, metadata):
    tags, xyz, _ = gmsh.model.mesh.getNodes()
    p = xyz.reshape(-1,3)[:,:2].T
    mapping = {int(t):i for i,t in enumerate(tags)}
    cells, labels = [], []
    for surface, label in regions:
        types, _, connectivity = gmsh.model.mesh.getElements(2,surface)
        for typ, conn in zip(types,connectivity):
            if typ != 2:
                raise RuntimeError('Only first-order triangular meshes are supported.')
            for tri in conn.reshape(-1,3):
                cells.append([mapping[int(t)] for t in tri])
                labels.append(label)
    used = np.unique(cells)
    old_to_new = np.full(len(tags),-1,dtype=int)
    old_to_new[used] = np.arange(len(used))
    mesh = MeshTri(p[:,used],old_to_new[np.array(cells).T])
    return VectorMesh(mesh,np.asarray(labels,dtype=int),metadata)


def rectangle_mesh(width_um, height_um, h_um):
    """gmsh triangular rectangle, origin (0,0); h is a diameter upper bound.

    Coordinate spacing <=h/2 gives diagonal length <=h/sqrt(2), with margin
    for floating-point CAD. The actual maximum diameter is always reported.
    """
    w = _positive(width_um,'width_um')
    h = _positive(height_um,'height_um')
    size = _positive(h_um,'h_um')
    gm = _gmsh_start('vector-rectangle')
    try:
        g = gm.model.geo
        points = [g.addPoint(x,y,0) for x,y in [(0,0),(w,0),(w,h),(0,h)]]
        curves = [g.addLine(points[i],points[(i+1)%4]) for i in range(4)]
        surface = g.addPlaneSurface([g.addCurveLoop(curves)])
        g.synchronize()
        for i,c in enumerate(curves):
            gm.model.mesh.setTransfiniteCurve(c,max(1,math.ceil(2*(w if i%2==0 else h)/size-1e-12))+1)
        gm.model.mesh.setTransfiniteSurface(surface,'Left',points)
        gm.model.mesh.generate(2)
        return _read_gmsh(gm,[(surface,0)],dict(kind='rectangle',width_um=w,height_um=h,
                          h_nominal_um=size,coordinate_spacing_bound_um=size/2))
    finally:
        gm.finalize()


def fiber_mesh(core_radius_um, outer_radius_um, h_um, outer_factor=5.):
    """gmsh disk/annulus, conforming interface.  Linear circle approximation.

    Core and interface size ~h, cladding grows to outer_factor*h.  Cell labels
    1=core, 0=cladding are assigned from CAD surfaces, not centroid thresholds.
    """
    a = _positive(core_radius_um,'core_radius_um')
    radius = _positive(outer_radius_um,'outer_radius_um')
    size = _positive(h_um,'h_um')
    factor = _positive(outer_factor,'outer_factor')
    if radius <= a or factor < 1:
        raise ValueError('outer_radius_um must exceed core radius; outer_factor must be >=1.')
    gm = _gmsh_start('vector-fiber')
    try:
        g = gm.model.geo
        center = g.addPoint(0,0,0,size)
        def circle(r, local_size):
            pts = [g.addPoint(r*np.cos(t),r*np.sin(t),0,local_size) for t in np.arange(4)*np.pi/2]
            return [g.addCircleArc(pts[i],center,pts[(i+1)%4]) for i in range(4)]
        inner = circle(a,size)
        outer = circle(radius,size*factor)
        loop = g.addCurveLoop(inner)
        core = g.addPlaneSurface([loop])
        clad = g.addPlaneSurface([g.addCurveLoop(outer),loop])
        g.synchronize()
        for c in inner:
            gm.model.mesh.setTransfiniteCurve(c,max(2,math.ceil(np.pi*a/(2*size)-1e-12))+1)
        distance = gm.model.mesh.field.add('Distance')
        gm.model.mesh.field.setNumbers(distance,'CurvesList',inner)
        gm.model.mesh.field.setNumber(distance,'Sampling',200)
        threshold = gm.model.mesh.field.add('Threshold')
        for key,val in [('InField',distance),('SizeMin',size),('SizeMax',factor*size),
                        ('DistMin',a*.5),('DistMax',a*3)]:
            gm.model.mesh.field.setNumber(threshold,key,val)
        gm.model.mesh.field.setAsBackgroundMesh(threshold)
        gm.option.setNumber('Mesh.MeshSizeExtendFromBoundary',0)
        gm.option.setNumber('Mesh.MeshSizeFromPoints',0)
        gm.option.setNumber('Mesh.MeshSizeFromCurvature',0)
        gm.option.setNumber('Mesh.Algorithm',6)
        gm.model.mesh.generate(2)
        return _read_gmsh(gm,[(core,1),(clad,0)],dict(kind='fiber',core_radius_um=a,
                    outer_radius_um=radius,h_nominal_um=size,outer_factor=factor))
    finally:
        gm.finalize()


@dataclass
class VectorMode:
    neff: complex
    beta_per_um: complex
    effective_area_um2: float
    loss_db_per_m: float
    pencil_residual: float
    gauss_residual: float
    edge_coefficients: np.ndarray
    scalar_coefficients: np.ndarray


@dataclass
class VectorSolution:
    geometry: VectorMesh
    edge_basis: Basis
    scalar_basis: Basis
    modes: list
    boundary_edge_dofs: np.ndarray
    boundary_scalar_dofs: np.ndarray
    wavelength_um: float
    diagnostics: dict

    def quadrature_fields(self, mode):
        """x(2,N), weights(N), E(3,N), Z0*H(3,N) at identical quadrature points.

        Electric field normalization: integral |E|^2 dA=1 (NOT one watt).
        Z0 H follows directly from curl_beta(E)/(i*k0).
        """
        et = self.edge_basis.interpolate(mode.edge_coefficients)
        phi = self.scalar_basis.interpolate(mode.scalar_coefficients)
        beta = mode.beta_per_um
        e = np.concatenate([np.asarray(et),(-1j*beta*np.asarray(phi))[None,...]],axis=0)
        combined = np.asarray(et)+phi.grad
        hbar = np.array([-beta*combined[1],beta*combined[0],-1j*et.curl])/(2*np.pi/self.wavelength_um)
        xy = np.asarray(self.edge_basis.global_coordinates()).reshape(2,-1)
        return xy,self.edge_basis.dx.ravel(),e.reshape(3,-1),hbar.reshape(3,-1)

    def sample_electric(self, mode, xy):
        """Pointwise E; points must be inside. Normal jumps remain discontinuous."""
        xy = np.asarray(xy,dtype=float)
        if xy.ndim != 2 or xy.shape[0] != 2:
            raise ValueError('xy must have shape (2,N).')
        et = self.edge_basis.interpolator(mode.edge_coefficients)(xy)
        phi = self.scalar_basis.interpolator(mode.scalar_coefficients)(xy)
        return np.vstack([et,-1j*mode.beta_per_um*phi])

    def save_npz(self, path):
        """Lossless raw FEM data; no nodal smoothing of transverse fields."""
        np.savez_compressed(path,points_um=self.geometry.mesh.p,triangles=self.geometry.mesh.t,
            region=self.geometry.region,wavelength_um=self.wavelength_um,
            neff=np.asarray([m.neff for m in self.modes]),
            edge_coefficients=np.asarray([m.edge_coefficients for m in self.modes]),
            scalar_coefficients=np.asarray([m.scalar_coefficients for m in self.modes]),
            effective_area_um2=np.asarray([m.effective_area_um2 for m in self.modes]),
            loss_db_per_m=np.asarray([m.loss_db_per_m for m in self.modes]))


def solve_vector(geometry, n, *, wavelength_um, target_neff, neff_window,
                 candidates=10, eig_tol=1e-10, maxiter=3000):
    """Solve A u=-beta^2 B u using LU shift-invert and nonsymmetric ARPACK.

    A has a zero longitudinal block; B is indefinite. No positive-definite
    generalized eigensolver is used. Output is only a target-near spectral
    subset; neither a window filter nor candidate count certifies completeness.
    """
    wavelength = _positive(wavelength_um,'wavelength_um')
    target = _positive(target_neff,'target_neff')
    _positive(eig_tol,'eig_tol')
    if (not isinstance(candidates,(int,np.integer)) or isinstance(candidates,bool) or candidates < 1):
        raise ValueError('candidates must be a positive integer.')
    if not isinstance(maxiter,(int,np.integer)) or maxiter < 1:
        raise ValueError('maxiter must be a positive integer.')
    window = np.asarray(neff_window)
    if (window.shape != (2,) or np.iscomplexobj(window) or not np.all(np.isfinite(window))
            or not 0 < window[0] < window[1]):
        raise ValueError('neff_window must be a finite ordered positive interval.')
    if isinstance(geometry,MeshTri):
        geometry = VectorMesh(geometry,np.zeros(geometry.nelements,dtype=int),dict(kind='custom'))
    if not isinstance(geometry,VectorMesh):
        raise ValueError('geometry must be a VectorMesh or MeshTri.')
    mesh = geometry.mesh
    bt = Basis(mesh,ElementTriN1(),intorder=6)
    bz = Basis(mesh,ElementTriP1(),quadrature=bt.quadrature)
    if callable(n):
        material = np.asarray(n(np.asarray(bt.global_coordinates())),dtype=complex)
        try:
            material = np.broadcast_to(material,bt.dx.shape)
        except ValueError as exc:
            raise ValueError('n(x) must broadcast to the integration-point shape.') from exc
    else:
        material = np.asarray(n,dtype=complex)
        if material.ndim == 0:
            material = np.full(bt.dx.shape,material)
        elif material.shape == (mesh.nelements,):
            material = np.broadcast_to(material[:,None],bt.dx.shape)
        else:
            raise ValueError('n must be a scalar, callable, or one value per triangle.')
    if not np.all(np.isfinite(material)) or np.any(material.real <= 0) or np.any(material.imag < 0):
        raise ValueError('n must be finite with Re(n)>0 and Im(n)>=0 (passive convention).')
    eps = material**2
    dtype = complex if np.any(material.imag != 0) else float
    if dtype is float:
        eps = eps.real
    k0 = 2*np.pi/wavelength

    @BilinearForm(dtype=dtype)
    def cc(u,v,w):
        return curl(u)*curl(v)
    @BilinearForm(dtype=dtype)
    def mm(u,v,w):
        return dot(u,v)
    @BilinearForm(dtype=dtype)
    def me(u,v,w):
        return w.eps*dot(u,v)
    @BilinearForm(dtype=dtype)
    def coupling(u,v,w):
        return dot(grad(u),v)
    @BilinearForm(dtype=dtype)
    def stiffness(u,v,w):
        return dot(grad(u),grad(v))
    @BilinearForm(dtype=dtype)
    def mass_eps(u,v,w):
        return w.eps*u*v
    @BilinearForm(dtype=dtype)
    def eps_coupling(u,v,w):
        return w.eps*dot(grad(u),v)
    kc = asm(cc,bt)
    mt = asm(mm,bt)
    met = asm(me,bt,eps=eps)
    g = asm(coupling,bz,bt)
    s = asm(stiffness,bz)
    q = asm(mass_eps,bz,eps=eps)
    ge = asm(eps_coupling,bz,bt,eps=eps)
    at = kc-k0**2*met
    zz = s-k0**2*q
    a_full = bmat([[at,None],[None,csr_matrix((bz.N,bz.N))]],format='csr')
    b_full = bmat([[mt,g],[g.T,zz]],format='csr')
    boundary_t = bt.get_dofs().all()
    boundary_z = bz.get_dofs().all()
    free_t = np.setdiff1d(np.arange(bt.N),boundary_t)
    free_z = np.setdiff1d(np.arange(bz.N),boundary_z)
    free = np.r_[free_t,bt.N+free_z]
    a = a_full[free][:,free].tocsc()
    b = b_full[free][:,free].tocsc()
    if candidates >= len(free)-1:
        raise ValueError('candidates must be smaller than the free system size minus one.')
    sigma = -(k0*target)**2
    lu = splu(a-sigma*b)
    op = LinearOperator(a.shape,matvec=lambda v: lu.solve(b@v),dtype=dtype)
    rng = np.random.default_rng(173)
    v0 = rng.standard_normal(len(free)).astype(dtype)
    theta,vectors = eigs(op,k=candidates,which='LM',tol=eig_tol,maxiter=maxiter,
                         v0=v0,ncv=min(len(free),max(2*candidates+1,30)))
    lambdas = sigma+1/theta
    result = VectorSolution(geometry,bt,bz,[],boundary_t,boundary_z,wavelength,{})
    rejected = []
    for j, lam in enumerate(lambdas):
        beta = np.sqrt(-complex(lam))
        if beta.real < 0:
            beta = -beta
        neff = beta/k0
        if not window[0] < neff.real < window[1]:
            rejected.append(dict(neff=[neff.real,neff.imag],reason='outside_window'))
            continue
        if dtype is float and abs(neff.imag) > 1e-8:
            rejected.append(dict(neff=[neff.real,neff.imag],reason='nonreal_lossless'))
            continue
        if neff.imag < -1e-8:
            rejected.append(dict(neff=[neff.real,neff.imag],reason='negative_attenuation'))
            continue
        u = vectors[:,j]
        au, bu = a@u, b@u
        residual = np.linalg.norm(au-lam*bu)/(np.linalg.norm(au)+abs(lam)*np.linalg.norm(bu))
        all_u = np.zeros(bt.N+bz.N,dtype=complex)
        all_u[free] = u
        et, phi = all_u[:bt.N],all_u[bt.N:]
        lhs = (ge.T@et)[free_z]
        rhs = (beta**2*q@phi)[free_z]
        # For TE modes both Gauss terms vanish. A scale independent of Ez avoids
        # treating floating-point cancellation as physical divergence error.
        gauss = np.linalg.norm(lhs-rhs)/(k0**2*np.linalg.norm(met@et)+
                     abs(beta)**2*np.linalg.norm(q@phi)+np.finfo(float).tiny)
        if residual > 1e-8 or gauss > 1e-8:
            rejected.append(dict(neff=[neff.real,neff.imag],reason='residual',
                                 pencil_residual=float(residual),gauss_residual=float(gauss)))
            continue
        mode = VectorMode(neff,beta,0.,float(20/np.log(10)*beta.imag*1e6),
                          float(residual),float(gauss),et,phi)
        _,weights,e,_ = result.quadrature_fields(mode)
        norm = np.sqrt(np.sum(np.sum(abs(e)**2,axis=0)*weights))
        phase = np.exp(-1j*np.angle(et[np.argmax(abs(et))]))
        mode.edge_coefficients *= phase/norm
        mode.scalar_coefficients *= phase/norm
        _,weights,e,_ = result.quadrature_fields(mode)
        intensity = np.sum(abs(e)**2,axis=0)
        mode.effective_area_um2 = float(np.sum(intensity*weights)**2/np.sum(intensity**2*weights))
        result.modes.append(mode)
    result.modes.sort(key=lambda m:-m.neff.real)
    edge_lengths = np.linalg.norm(mesh.p[:,mesh.facets[0]]-mesh.p[:,mesh.facets[1]],axis=0)
    core_cells = np.where(geometry.region == 1)[0]
    core_h = None
    if len(core_cells):
        core_h = float(np.max(edge_lengths[mesh.t2f[:,core_cells]]))
    result.diagnostics = dict(triangles=mesh.nelements,nodes=mesh.nvertices,
        edge_dofs=bt.N,scalar_dofs=bz.N,free_dofs=len(free),h_max_um=float(max(edge_lengths)),
        h_core_max_um=core_h,candidates=candidates,returned=len(result.modes),
        target_neff=target,window=window.tolist(),rejected=rejected,
        completeness_certified=False,element_pair='ElementTriN1 + ElementTriP1',
        normalization='integral |E|^2 dA = 1',boundary='PEC')
    if len(result.modes) == candidates:
        warnings.warn('All requested candidates lie in the window; request more to check mode count.',RuntimeWarning)
    return result
