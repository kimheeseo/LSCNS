"""M3 2D scalar Maxwell scattering (TE=E_z, TM=H_z).

Lengths are micrometres, time exp(-i omega t). P1 Galerkin with Cartesian
complex-stretch PML and a complete discrete transverse modal DtN map.
Only real, lossless port materials; passive complex interior indices allowed.
PEC transverse walls: TE Dirichlet, TM natural Neumann. No 3D vector ports.
"""
from dataclasses import dataclass, field
import math
import numpy as np
from scipy.linalg import eigh
from scipy.sparse import coo_matrix
from scipy.sparse.linalg import splu
from scipy.special import hankel1
from skfem import Basis, ElementTriP1, BilinearForm, LinearForm, asm
from .vector import VectorMesh, _gmsh_start, _read_gmsh, _positive

Z0=376.730313668


@dataclass(frozen=True)
class PML:
    x_physical: tuple
    y_physical: tuple
    left: float=0.
    right: float=0.
    bottom: float=0.
    top: float=0.
    strength: float=4.
    power: int=2

    def __post_init__(self):
        for bounds in [self.x_physical,self.y_physical]:
            if len(bounds)!=2 or not np.all(np.isfinite(bounds)) or bounds[1]<=bounds[0]:
                raise ValueError('Physical intervals must be finite and increasing.')
        for value in [self.left,self.right,self.bottom,self.top,self.strength]:
            if not np.isreal(value) or not np.isfinite(value) or value<0:
                raise ValueError('PML widths and strength must be finite nonnegative real values.')
        if not isinstance(self.power,int) or self.power<1: raise ValueError('PML power must be an integer >=1.')

    @property
    def outer(self):
        return (self.x_physical[0]-self.left,self.x_physical[1]+self.right,
                self.y_physical[0]-self.bottom,self.y_physical[1]+self.top)

    def stretch(self,x):
        def one(coord,bounds,minus,plus):
            sigma=np.zeros_like(coord,dtype=float)
            if minus: sigma+=self.strength*(np.maximum(bounds[0]-coord,0)/minus)**self.power
            if plus: sigma+=self.strength*(np.maximum(coord-bounds[1],0)/plus)**self.power
            return 1+1j*sigma
        return one(x[0],self.x_physical,self.left,self.right),one(x[1],self.y_physical,self.bottom,self.top)


def guide_mesh(x_breaks,y_breaks,h_um):
    """Conforming gmsh rectangular subdomains. Coordinate spacing <=h/3.

    Breaks must include discontinuous material and PML interfaces explicitly.
    Region IDs enumerate each CAD rectangle; n is supplied to the solver.
    """
    h=_positive(h_um,'h_um');xs=np.asarray(x_breaks,dtype=float);ys=np.asarray(y_breaks,dtype=float)
    if xs.ndim!=1 or ys.ndim!=1 or min(len(xs),len(ys))<2 or not np.all(np.isfinite(np.r_[xs,ys])) or np.any(np.diff(xs)<=0) or np.any(np.diff(ys)<=0):
        raise ValueError('Mesh breaks must be finite strictly increasing vectors.')
    gm=_gmsh_start('scalar-guide')
    try:
        g=gm.model.geo; points={}; lines={}; surfaces=[]
        for i,x in enumerate(xs):
            for j,y in enumerate(ys): points[i,j]=g.addPoint(float(x),float(y),0)
        def line(a,b):
            if (a,b) in lines:return lines[a,b]
            if (b,a) in lines:return -lines[b,a]
            lines[a,b]=g.addLine(points[a],points[b]);return lines[a,b]
        for i in range(len(xs)-1):
            for j in range(len(ys)-1):
                corners=[(i,j),(i+1,j),(i+1,j+1),(i,j+1)]
                curves=[line(corners[t],corners[(t+1)%4]) for t in range(4)]
                s=g.addPlaneSurface([g.addCurveLoop(curves)])
                surfaces.append((s,[points[a] for a in corners]))
        g.synchronize()
        for (a,b),tag in lines.items():
            distance=np.hypot(xs[a[0]]-xs[b[0]],ys[a[1]]-ys[b[1]])
            gm.model.mesh.setTransfiniteCurve(tag,max(1,math.ceil(3*distance/h-1e-12))+1)
        for s,corners in surfaces:gm.model.mesh.setTransfiniteSurface(s,'Left',corners)
        gm.model.mesh.generate(2)
        return _read_gmsh(gm,[(s,i) for i,(s,_) in enumerate(surfaces)],dict(kind='guide',bounds=(xs[0],xs[-1],ys[0],ys[-1]),h_nominal_um=h,x_breaks=xs.tolist(),y_breaks=ys.tolist()))
    finally:gm.finalize()


def cylinder_mesh(radius_um,pml,h_um):
    """Conforming gmsh circular inclusion and nine-block rectangular PML.

    First-order straight geometry approximates the circular interface; actual
    maximum element diameter is reported, not equated to the gmsh size input.
    """
    a=_positive(radius_um,'radius_um'); h=_positive(h_um,'h_um')
    if not isinstance(pml,PML):raise ValueError('A PML instance is required.')
    if min(pml.left,pml.right,pml.bottom,pml.top)<=0:raise ValueError('Cylinder needs PML on all four sides.')
    if not (pml.x_physical[0]<-a<a<pml.x_physical[1] and pml.y_physical[0]<-a<a<pml.y_physical[1]):
        raise ValueError('Cylinder must be inside the physical box.')
    gm=_gmsh_start('scalar-cylinder')
    try:
        g=gm.model.geo;xl,xr,yb,yt=pml.outer
        xs=[xl,*pml.x_physical,xr];ys=[yb,*pml.y_physical,yt];points={};lines={}
        for i,x in enumerate(xs):
            for j,y in enumerate(ys):points[i,j]=g.addPoint(x,y,0,h)
        center=g.addPoint(0,0,0,h)
        circ=[g.addPoint(a*np.cos(t),a*np.sin(t),0,h) for t in np.arange(4)*np.pi/2]
        arcs=[g.addCircleArc(circ[i],center,circ[(i+1)%4]) for i in range(4)]
        cloop=g.addCurveLoop(arcs);core=g.addPlaneSurface([cloop]);surfaces=[(core,1)]
        def line(a,b):
            if (a,b) in lines:return lines[a,b]
            if (b,a) in lines:return -lines[b,a]
            lines[a,b]=g.addLine(points[a],points[b]);return lines[a,b]
        for i in range(3):
            for j in range(3):
                c=[(i,j),(i+1,j),(i+1,j+1),(i,j+1)]
                loop=g.addCurveLoop([line(c[t],c[(t+1)%4]) for t in range(4)])
                s=g.addPlaneSurface([loop,cloop] if i==j==1 else [loop]);surfaces.append((s,0))
        g.synchronize()
        gm.option.setNumber('Mesh.MeshSizeMin',h)
        gm.option.setNumber('Mesh.MeshSizeMax',h)
        gm.option.setNumber('Mesh.Algorithm',6)
        for arc in arcs:gm.model.mesh.setTransfiniteCurve(arc,max(2,math.ceil(np.pi*a/(2*h)))+1)
        gm.model.mesh.generate(2)
        return _read_gmsh(gm,surfaces,dict(kind='cylinder',radius_um=a,bounds=pml.outer,h_nominal_um=h))
    finally:gm.finalize()


def _pol(polarization):
    if polarization not in ('TE','TM'):raise ValueError("polarization must be 'TE' (Ez) or 'TM' (Hz).")
    return polarization


def _index(n,x):
    values=np.asarray(n(x) if callable(n) else n,dtype=complex)
    try: values=np.broadcast_to(values,x.shape[1:])
    except ValueError as exc:raise ValueError('n must be scalar or callable returning the coordinate trailing shape.') from exc
    if not np.all(np.isfinite(values)) or np.any(values.real<=0) or np.any(values.imag<0):
        raise ValueError('Only finite passive indices Re(n)>0, Im(n)>=0 are supported.')
    return values


def _pq(n,pol):return (np.ones_like(n),n*n) if pol=='TE' else (1/(n*n),np.ones_like(n))


@BilinearForm(dtype=complex)
def _helmholtz(u,v,w):
    return w.dxcoef*u.grad[0]*v.grad[0]+w.dycoef*u.grad[1]*v.grad[1]-w.masscoef*u*v


def _volume(geometry,n,wl,pol,pml):
    basis=Basis(geometry.mesh,ElementTriP1(),intorder=6);x=np.asarray(basis.global_coordinates())
    nv=_index(n,x);p,q=_pq(nv,pol)
    if pml:
        sx,sy=pml.stretch(x)
        bounds=np.array([geometry.mesh.p[0].min(),geometry.mesh.p[0].max(),geometry.mesh.p[1].min(),geometry.mesh.p[1].max()])
        if not np.allclose(bounds,pml.outer,rtol=0,atol=1e-10):raise ValueError('Mesh bounds must equal PML outer bounds.')
    else:sx=sy=1.
    k=2*np.pi/wl
    matrix=asm(_helmholtz,basis,dxcoef=p*sy/sx,dycoef=p*sx/sy,masscoef=k*k*q*sx*sy).tocsc()
    return basis,matrix,nv,p,q


@dataclass
class PortModes:
    nodes: np.ndarray
    y_um: np.ndarray
    beta: np.ndarray
    vectors: np.ndarray
    weighted_mass: np.ndarray
    propagating: np.ndarray
    dtn: np.ndarray


def _port_modes(geometry,n,wl,pol,side):
    mesh=geometry.mesh;x0=mesh.p[0].min() if side=='left' else mesh.p[0].max()
    nodes=np.flatnonzero(np.isclose(mesh.p[0],x0,rtol=0,atol=1e-10))
    nodes=nodes[np.argsort(mesh.p[1,nodes])];y=mesh.p[1,nodes];N=len(y)
    if N<3:raise ValueError('Port requires >=3 transverse nodes.')
    K=np.zeros((N,N));Mp=K.copy();Mq=K.copy()
    z,w=np.polynomial.legendre.leggauss(6);shape=np.array([(1-z)/2,(1+z)/2])
    eps_x=1e-10*(mesh.p[0].max()-mesh.p[0].min());inside_x=x0+eps_x if side=='left' else x0-eps_x
    for j,dy in enumerate(np.diff(y)):
        yc=(y[j]+y[j+1])/2+dy*z/2;nv=_index(n,np.array([np.full_like(yc,inside_x),yc]))
        if np.any(abs(nv.imag)>1e-14):raise ValueError('Port materials must be real and lossless.')
        p,q=_pq(nv.real,pol);weights=w*dy/2
        ind=np.ix_([j,j+1],[j,j+1])
        K[ind]+=np.array([[1,-1],[-1,1]])*np.sum(weights*p)/dy**2
        Mp[ind]+=(shape*(weights*p))@shape.T
        Mq[ind]+=(shape*(weights*q))@shape.T
    active=np.arange(1,N-1) if pol=='TE' else np.arange(N)
    val,vec=eigh((K-(2*np.pi/wl)**2*Mq)[np.ix_(active,active)],Mp[np.ix_(active,active)])
    beta=np.sqrt(-val+0j)
    if np.any(abs(beta)<1e-7*(2*np.pi/wl)):raise ValueError('Port mode at cutoff: zero power normalization is undefined.')
    full=np.zeros((N,len(val)));full[active]=vec
    for j in range(full.shape[1]):
        # deterministic sign: first significantly nonzero nodal coefficient
        first=np.flatnonzero(abs(full[:,j])>1e-7*np.max(abs(full[:,j])))[0]
        if full[first,j]<0:full[:,j]*=-1
    prop=np.flatnonzero((beta.real>1e-8)&(abs(beta.imag)<1e-10))
    trace=Mp@full
    dtn=(trace*beta)@trace.T
    return PortModes(nodes,y,beta,full,Mp,prop,dtn)


def _diagnostics(geometry,basis,free,A,u,b):
    p=geometry.mesh.p;t=geometry.mesh.t
    edges=np.concatenate([np.linalg.norm(p[:,t[i]]-p[:,t[j]],axis=0) for i,j in [(0,1),(1,2),(2,0)]])
    rhsnorm=np.linalg.norm(b[free]);res=np.linalg.norm((A@u-b)[free])
    return dict(h_nominal_um=geometry.metadata.get('h_nominal_um'),h_max_um=float(edges.max()),triangles=int(t.shape[1]),dofs=int(basis.N),free_dofs=len(free),linear_residual=float(res/max(rhsnorm,1e-300)),element='P1',quadrature_order=6)


def _solve(A,b,dirichlet):
    free=np.setdiff1d(np.arange(A.shape[0]),dirichlet);u=np.zeros(A.shape[0],complex)
    if np.any(b):u[free]=splu(A[free][:,free].tocsc()).solve(b[free])
    return u,free


@dataclass
class ScatteringSolution:
    geometry: VectorMesh
    basis: Basis
    field: np.ndarray
    wavelength_um: float
    polarization: str
    index: object
    diagnostics: dict
    s_parameters: dict=field(default_factory=dict)
    port_modes: dict=field(default_factory=dict)
    incident_port: str|None=None
    incident_mode: int|None=None
    background_index: float|None=None
    incidence_angle_rad: float=0.
    quadrature_index: np.ndarray|None=None
    _point_finder: object=field(default=None,init=False,repr=False)

    def sample(self,xy,total=False):
        xy=np.asarray(xy,dtype=float)
        if xy.ndim!=2 or xy.shape[0]!=2:raise ValueError('xy must have shape (2,N).')
        if not np.all(np.isfinite(xy)):raise ValueError('Point coordinates must be finite.')
        if self._point_finder is None:
            from matplotlib.tri import Triangulation
            mesh=self.geometry.mesh
            self._point_finder=Triangulation(mesh.p[0],mesh.p[1],mesh.t.T).get_trifinder()
        cells=self._point_finder(xy[0],xy[1])
        if np.any(cells<0):raise ValueError('Sampling points must lie inside the mesh.')
        # Explicit P1 barycentric evaluation avoids element_finder's dense
        # all-cell fallback for large grids near irregular material interfaces.
        vertices=self.geometry.mesh.t[:,cells];p=self.geometry.mesh.p
        v0=p[:,vertices[0]];d1=p[:,vertices[1]]-v0;d2=p[:,vertices[2]]-v0;delta=xy-v0
        determinant=d1[0]*d2[1]-d1[1]*d2[0]
        l1=(delta[0]*d2[1]-delta[1]*d2[0])/determinant
        l2=(d1[0]*delta[1]-d1[1]*delta[0])/determinant
        value=self.field[vertices[0]]*(1-l1-l2)+self.field[vertices[1]]*l1+self.field[vertices[2]]*l2
        if total:
            if self.background_index is None:raise ValueError('Port solution already contains the total field; use total=False.')
            k=2*np.pi*self.background_index/self.wavelength_um
            value+=np.exp(1j*k*(np.cos(self.incidence_angle_rad)*xy[0]+np.sin(self.incidence_angle_rad)*xy[1]))
        return value

    def quadrature_fields(self,total=False):
        """Return xy, weights, E(3,N), H(3,N); physical-domain interpretation.

        TE scalar has E_z units; TM scalar has H_z units. PML fields are
        computational stretched fields, not ordinary physical E and H.
        """
        x=np.asarray(self.basis.global_coordinates());fi=self.basis.interpolate(self.field)
        u=np.asarray(fi).copy();g=fi.grad.copy()
        if total:
            if self.background_index is None:raise ValueError('Port solution is already total.')
            k=2*np.pi*self.background_index/self.wavelength_um;d=np.array([np.cos(self.incidence_angle_rad),np.sin(self.incidence_angle_rad)])
            ui=np.exp(1j*k*(d[0]*x[0]+d[1]*x[1]));u+=ui;g+=1j*k*d[:,None,None]*ui
        k0=2*np.pi/self.wavelength_um;E=np.zeros((3,*u.shape),complex);H=E.copy()
        if self.polarization=='TE':E[2]=u;H[:2]=np.array([g[1],-g[0]])/(1j*k0*Z0)
        else:H[2]=u;E[:2]=1j*Z0*np.array([g[1],-g[0]])/(k0*(self.quadrature_index if self.quadrature_index is not None else _index(self.index,x))**2)
        return x.reshape(2,-1),self.basis.dx.ravel(),E.reshape(3,-1),H.reshape(3,-1)

    def scattering_width(self,radius_um,orders=8,samples=1440):
        """Outgoing Hankel fit on a physical exterior circle, units um.

        All retained Fourier orders must be radiating in homogeneous background.
        This method is for plane-wave scattering, not port solutions.
        """
        if self.background_index is None:raise ValueError('Plane-wave scattered field required.')
        radius=_positive(radius_um,'radius_um')
        if not isinstance(orders,int) or orders<0 or not isinstance(samples,int) or samples<=2*orders:raise ValueError('Need samples>2*orders and nonnegative integer orders.')
        theta=np.arange(samples)*2*np.pi/samples;xy=radius*np.array([np.cos(theta),np.sin(theta)])
        nv=_index(self.index,xy)
        if not np.allclose(nv,self.background_index,rtol=0,atol=1e-12):raise ValueError('Extraction circle must be in homogeneous background.')
        bounds=self.diagnostics.get('physical_bounds')
        if bounds and not (xy[0].min()>bounds[0] and xy[0].max()<bounds[1] and xy[1].min()>bounds[2] and xy[1].max()<bounds[3]):raise ValueError('Extraction circle must be inside physical region.')
        f=self.sample(xy);m=np.arange(-orders,orders+1);k=2*np.pi*self.background_index/self.wavelength_um
        c=np.mean(f[None,:]*np.exp(-1j*m[:,None]*theta),axis=1)/hankel1(m,k*radius)
        return float(4/k*np.sum(abs(c)**2))

    def save_npz(self,path):
        data=dict(points_um=self.geometry.mesh.p,triangles=self.geometry.mesh.t,region=self.geometry.region,
                  field=self.field,wavelength_um=self.wavelength_um,polarization=self.polarization,
                  field_kind="scattered" if self.background_index is not None else "total",
                  index_at_quadrature=self.quadrature_index,
                  quadrature_xy_um=np.asarray(self.basis.global_coordinates()),quadrature_weights_um2=self.basis.dx,
                  background_index=self.background_index if self.background_index is not None else np.nan,
                  incidence_angle_rad=self.incidence_angle_rad)
        for name,s in self.s_parameters.items():data['S_'+name]=s
        for name,port in self.port_modes.items():
            data['port_'+name+'_beta_per_um']=port.beta;data['port_'+name+'_nodes']=port.nodes;data['port_'+name+'_vectors']=port.vectors
        np.savez_compressed(path,**data)


def solve_ports(geometry,n,wavelength_um,polarization,incident_port='left',incident_mode=0,pml=None,ports=('left','right')):
    """One modal incident wave. Returned S arrays contain propagating modes.

    Mode numbers start at 0, decreasing Re(beta). All discrete evanescent modes
    are retained in DtN. The full S matrix follows by exciting each port/mode.
    Geometry needs straight vertical ports and horizontal PEC transverse walls.
    """
    pol=_pol(polarization);wl=_positive(wavelength_um,'wavelength_um')
    if not ports or len(set(ports))!=len(ports) or any(s not in ('left','right') for s in ports):raise ValueError('Ports must be a nonempty distinct subset of left/right.')
    if incident_port not in ports:raise ValueError('incident_port must be an enabled port.')
    if not isinstance(incident_mode,int) or incident_mode<0:raise ValueError('incident_mode must be a nonnegative integer.')
    if pml and (pml.bottom or pml.top or (pml.left and 'left' in ports) or (pml.right and 'right' in ports)):
        raise ValueError('PML may terminate x only opposite a port; transverse PML port modes are not supported.')
    basis,A,nv,p,q=_volume(geometry,n,wl,pol,pml);b=np.zeros(basis.N,complex);modes={}
    for side in ports:
        pm=_port_modes(geometry,n,wl,pol,side);modes[side]=pm
        ii,jj=np.meshgrid(pm.nodes,pm.nodes,indexing='ij')
        A=A+coo_matrix((-1j*pm.dtn.ravel(),(ii.ravel(),jj.ravel())),shape=A.shape).tocsc()
    pm=modes[incident_port]
    if incident_mode>=len(pm.propagating):raise ValueError('Requested incident propagating mode does not exist (possibly cutoff).')
    j=pm.propagating[incident_mode];b[pm.nodes]=-2j*pm.beta[j]*(pm.weighted_mass@pm.vectors[:,j])
    mesh=geometry.mesh;boundary=mesh.boundary_nodes();bounds=(mesh.p[0].min(),mesh.p[0].max(),mesh.p[1].min(),mesh.p[1].max())
    dirichlet=[]
    if pol=='TE':dirichlet.extend(boundary[np.isclose(mesh.p[1,boundary],bounds[2])|np.isclose(mesh.p[1,boundary],bounds[3])])
    for side,pos in [('left',bounds[0]),('right',bounds[1])]:
        if side not in ports:dirichlet.extend(boundary[np.isclose(mesh.p[0,boundary],pos)])
    u,free=_solve(A,b,np.unique(dirichlet));S={};beta_in=pm.beta[j].real
    for side,port in modes.items():
        c=port.vectors.T@(port.weighted_mass@u[port.nodes])
        if side==incident_port:c[j]-=1
        S[side]=c[port.propagating]*np.sqrt(port.beta[port.propagating].real/beta_in)
    diag=_diagnostics(geometry,basis,free,A,u,b)
    diag.update(ports={side:dict(trace_modes=len(m.beta),propagating_modes=len(m.propagating),evanescent_modes=len(m.beta)-len(m.propagating)) for side,m in modes.items()},power_out=float(sum(np.sum(abs(v)**2) for v in S.values())),port_reference_planes_um={s:bounds[0 if s=='left' else 1] for s in ports})
    return ScatteringSolution(geometry,basis,u,wl,pol,n,diag,S,modes,incident_port,incident_mode,quadrature_index=nv)


def solve_plane_wave(geometry,n_inside,n_background,wavelength_um,polarization,pml,angle_rad=0.):
    """Compact-contrast scattered field with Cartesian PML.

    n_inside scalar selects region 1 on a cylinder mesh; callable supplies a
    general index distribution on any conforming mesh with rectangular exterior.
    Material contrast cannot touch the PML. Background is real and lossless.
    """
    pol=_pol(polarization);wl=_positive(wavelength_um,'wavelength_um');nb=_positive(n_background,'n_background')
    if not isinstance(pml,PML):raise ValueError('Plane-wave scattering requires PML.')
    if not np.isfinite(angle_rad):raise ValueError('angle_rad must be finite.')
    if callable(n_inside):index=n_inside
    else:
        if geometry.metadata.get('kind')!='cylinder':raise ValueError('Use a callable index for a non-cylinder geometry.')
        a=geometry.metadata['radius_um'];ni=complex(n_inside)
        # Material quadrature labels follow CAD, including the polygonal circle.
        index=lambda x:np.where(np.hypot(x[0],x[1])<a,ni,nb)
    basis=Basis(geometry.mesh,ElementTriP1(),intorder=6)
    if not callable(n_inside):
        # Do not reclassify triangles by circle radius at material quadrature.
        nv=np.broadcast_to(np.where(geometry.region[:,None]==1,complex(n_inside),nb),basis.dx.shape)
        _index(nv,np.asarray(basis.global_coordinates()))
        # _volume receives a callable preserving cell labels for these quadrature points.
        assemble_index=lambda x:nv
        basis,A,nv,p,q=_volume(geometry,assemble_index,wl,pol,pml)
    else:basis,A,nv,p,q=_volume(geometry,index,wl,pol,pml)
    x=np.asarray(basis.global_coordinates());sx,sy=pml.stretch(x)
    outside=(abs(sx-1)>1e-14)|(abs(sy-1)>1e-14)
    if np.any(abs(nv[outside]-nb)>1e-12):raise ValueError('Material contrast must be zero in PML.')
    k0=2*np.pi/wl;kb=k0*nb;direction=np.array([np.cos(angle_rad),np.sin(angle_rad)])
    ui=np.exp(1j*kb*(direction[0]*x[0]+direction[1]*x[1]));pb,qb=_pq(np.array(nb),pol)
    @LinearForm(dtype=complex)
    def source(v,w):
        return -(p-pb)*1j*kb*ui*(direction[0]*v.grad[0]+direction[1]*v.grad[1])+k0*k0*(q-qb)*ui*v
    b=asm(source,basis);u,free=_solve(A,b,geometry.mesh.boundary_nodes());diag=_diagnostics(geometry,basis,free,A,u,b)
    diag.update(physical_bounds=(*pml.x_physical,*pml.y_physical),background_index=nb,angle_rad=float(angle_rad))
    return ScatteringSolution(geometry,basis,u,wl,pol,index,diag,background_index=nb,incidence_angle_rad=float(angle_rad),quadrature_index=nv)


def solve_s_matrix(geometry,n,wavelength_um,polarization):
    """Complete propagating two-port S matrix and channel order.

    Each column is a separate unit modal incidence. Current implementation
    rebuilds/factorizes each column; this is intentionally a small serial tool.
    """
    first=solve_ports(geometry,n,wavelength_um,polarization)
    channels=[(side,i) for side in ('left','right') for i in range(len(first.s_parameters[side]))]
    columns=[];diagnostics=[]
    for side,i in channels:
        sol=first if (side,i)==('left',0) else solve_ports(geometry,n,wavelength_um,polarization,side,i)
        columns.append(np.concatenate([sol.s_parameters[s] for s in ('left','right')]))
        diagnostics.append(sol.diagnostics)
    return np.array(columns).T,channels,diagnostics
