"""Curved P2 scalar FEM for identical, straight, lossless two-core fibers.
Quarter-domain parity constraints avoid mesh asymmetry contaminating splitting.
Core interfaces are Gmsh arcs; P2 edge midpoints are projected onto circles.
"""
import numpy as np
from scipy.sparse.linalg import eigsh
from skfem import MeshTri,MeshTri2,Basis,ElementTriP2,BilinearForm,asm
from skfem.helpers import dot,grad
from multipole_reference import validate


def build_mesh(fiber,pitch_um=16.,halfwidth_um=32.,mesh_size_um=.6):
    import gmsh
    validate(fiber,pitch_um)
    if not np.isfinite(halfwidth_um) or halfwidth_um <= pitch_um/2+2*fiber.radius_um:
        raise ValueError('halfwidth must leave at least one core radius of cladding')
    if not np.isfinite(mesh_size_um) or mesh_size_um<=0 or mesh_size_um>fiber.radius_um:
        raise ValueError('mesh size must be positive and <= core radius')
    if gmsh.isInitialized():raise RuntimeError('Gmsh already initialized: finish the existing model first')
    gmsh.initialize()
    try:
        gmsh.option.setNumber('General.Terminal',0)
        gmsh.option.setNumber('General.NumThreads',1)
        gmsh.model.add('symmetric_two_core')
        geo=gmsh.model.geo; a=fiber.radius_um;c=pitch_um/2;L=halfwidth_um;h=mesh_size_um
        def p(x,y):return geo.addPoint(x,y,0,h)
        bl,br,tr,tl=p(0,0),p(L,0),p(L,L),p(0,L)
        left,right,top,center=p(c-a,0),p(c+a,0),p(c,a),p(c,0)
        arc1=geo.addCircleArc(right,center,top);arc2=geo.addCircleArc(top,center,left)
        diameter=geo.addLine(left,right)
        core=geo.addPlaneSurface([geo.addCurveLoop([diameter,arc1,arc2])])
        lines=[geo.addLine(bl,left),geo.addLine(right,br),geo.addLine(br,tr),geo.addLine(tr,tl),geo.addLine(tl,bl)]
        clad=geo.addPlaneSurface([geo.addCurveLoop([lines[0],-arc2,-arc1,*lines[1:]])])
        geo.synchronize()
        gmsh.option.setNumber('Mesh.MeshSizeMin',h)
        gmsh.option.setNumber('Mesh.MeshSizeMax',h)
        gmsh.option.setNumber('Mesh.Algorithm',6)
        gmsh.model.mesh.generate(2)
        tags,coords,_=gmsh.model.mesh.getNodes();coords=np.asarray(coords).reshape(-1,3)[:,:2]
        lookup={int(t):i for i,t in enumerate(tags)}
        tri=[];is_core=[]
        for surface,inside in ((core,True),(clad,False)):
            types,_,nodes=gmsh.model.mesh.getElements(2,surface)
            for typ,nod in zip(types,nodes):
                if typ!=2:raise RuntimeError('expected linear triangles')
                t=np.array([lookup[int(x)] for x in nod]).reshape(-1,3)
                tri.extend(t);is_core.extend([inside]*len(t))
        # Drop unused geometric center node.
        tri=np.asarray(tri); used=np.unique(tri); remap=np.full(len(coords),-1);remap[used]=np.arange(len(used))
        linear=MeshTri(np.ascontiguousarray(coords[used].T),np.ascontiguousarray(remap[tri].T))
        mesh=MeshTri2.from_mesh(linear)
        # P2 internal interface edges are not exterior boundary facets.
        ids=mesh.facets
        r=np.hypot(linear.p[0]-c,linear.p[1])
        curved=np.flatnonzero(np.all(np.abs(r[ids]-a)<1e-8,axis=0))
        mid=Basis(mesh,ElementTriP2()).dofs.facet_dofs[0,curved]
        vector=mesh.doflocs[:,mid]-np.array([[c],[0.]])
        mesh.doflocs[:,mid]=np.array([[c],[0.]])+a*vector/np.linalg.norm(vector,axis=0)
        return mesh,np.array(is_core,dtype=bool)
    finally:gmsh.finalize()


@BilinearForm
def stiffness(u,v,w):return dot(grad(u),grad(v))
@BilinearForm
def mass(u,v,w):return u*v
@BilinearForm
def weighted_mass(u,v,w):return w.n2*u*v


def solve_conforming(fiber,pitch_um=16.,halfwidth_um=32.,mesh_size_um=.6,return_fields=False):
    mesh,inside=build_mesh(fiber,pitch_um,halfwidth_um,mesh_size_um)
    basis=Basis(mesh,ElementTriP2(),intorder=8)
    k=2*np.pi/fiber.wavelength_um
    K=asm(stiffness,basis);M=asm(mass,basis)
    n2=np.where(inside,fiber.core**2,fiber.clad**2)[:,None]*np.ones_like(basis.dx)
    N=asm(weighted_mass,basis,n2=n2);A=k*k*N-K
    x,y=basis.doflocs;L=halfwidth_um
    outer=np.isclose(x,L,atol=1e-9)|np.isclose(y,L,atol=1e-9)
    roots=[];residuals=[];fields=[]
    for parity in (1,-1):
        fixed=outer | ((np.isclose(x,0,atol=1e-9)) if parity==-1 else False)
        free=np.flatnonzero(~fixed);AA=A[free][:,free];MM=M[free][:,free]
        val,vec=eigsh(AA,k=1,M=MM,sigma=(k*fiber.core)**2,which='LM',tol=2e-12,v0=np.ones(len(free)))
        q=vec[:,0];beta2=float(val[0]);n=np.sqrt(beta2)/k
        roots.append(float(n));residuals.append(float(np.linalg.norm(AA@q-beta2*(MM@q))/np.linalg.norm(beta2*(MM@q))))
        full=np.zeros(basis.N);full[free]=q
        if np.sum(full)<0:full=-full
        fields.append(full)
    if not fiber.clad<roots[1]<roots[0]<fiber.core:raise RuntimeError('two guided parity modes not resolved')
    split=roots[0]-roots[1]
    result=dict(n_even=roots[0],n_odd=roots[1],index_split=split,
                coupling_length_mm=fiber.wavelength_um/(2*split)/1000,
                kappa_per_m=np.pi*split/(fiber.wavelength_um*1e-6),
                residuals=residuals,quarter_dofs=int(basis.N),quarter_elements=int(mesh.nelements),
                mesh_size_um=mesh_size_um,halfwidth_um=halfwidth_um,pitch_um=pitch_um,
                method='curved P2 scalar FEM; x/y parity quarter domain')
    if return_fields:result.update(mesh=mesh,basis=basis,fields=np.array(fields))
    return result


def coherent_transfer(length_m,coupling_length_mm):
    """Lossless identical two-mode model: P2=sin^2(pi*z/(2*Lc)).

    This is deterministic coherent transfer, not measured/random cable XT.
    """
    z=np.asarray(length_m,float)
    if np.any(~np.isfinite(z)) or np.any(z<0) or not np.isfinite(coupling_length_mm) or coupling_length_mm<=0:
        raise ValueError('finite nonnegative length and positive coupling length required')
    p2=np.sin(np.pi*z/(2*coupling_length_mm*1e-3))**2
    return dict(P1=1-p2,P2=p2)
