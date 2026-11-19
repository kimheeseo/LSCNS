"""Mixed vector Maxwell assembly retained and extended for STEP 3 v0.8.5.

E=(Et,Ez) exp(i beta z-i omega t), Nedelec edge / H1 elements of order 1 or 2.
K0+beta K1+beta² K2 is solved by vector_spectrum.py; output diagnostics and
mode selection live in vector_solver_v085.py. Synthetic circular benchmarks
are independently checked; general precision fiber leakage is not certified.
"""
import math
import cmath

import numpy as np
from scipy.sparse import bmat, csc_matrix
from scipy.sparse.linalg import splu

ENGINE_VERSION="0.8.6"

try:
    from .materials import index
except ImportError:
    from materials import index

Z0 = 376.730313668  # Preserved CODATA 2018 vacuum impedance, ohm.
# https://tsapps.nist.gov/publication/get_pdf.cfm?pub_id=932441
# CODATA 2022 is 376.730313412(59): see STEP3_V085_ALGORITHM_20261119.md.

def _tensor(x, eps, design, photoelastic_delta=None):
    """Radial complex coordinate stretch, tensor ε=detJ J^-1 ε J^-T, μ^-1=J J^T/detJ."""
    xx, yy = x[0], x[1]
    bend=design.get("_bend")
    if bend is not None:
        # STEP6a: equivalent straight waveguide. Apply n_eq^2 BEFORE PML.
        # x,y in um; R_eff in mm. PML keeps this varying background index.
        try:
            from .bending import conformal_factor
        except ImportError:
            from bending import conformal_factor
        if bend.get("model")!="2d_conformal":
            raise ValueError("only sourced 2d_conformal bending is implemented")
        f=conformal_factor(xx,yy,bend["R_eff_mm"],bend["direction_deg"])
        eps=eps*f*f
    r=np.maximum(np.sqrt(xx*xx+yy*yy),1e-12)
    radius=float(design["domain_radius_um"])
    bc=design["boundary"]
    sr=np.ones_like(r,dtype=complex)
    st=np.ones_like(r,dtype=complex)
    if bc["type"]=="pml":
        th=bc["thickness_um"]
        q=np.clip((r-radius)/th,0.,1.)
        strength=bc["strength"]
        order=int(bc["order"])
        sr=1+1j*strength*q**order
        # r complex coordinate = r + i integral_R^r sigma(r') dr'
        st=1+1j*strength*th*q**(order+1)/((order+1)*r)
    ux,uy=xx/r,yy/r
    mu_r=sr/st;mu_t=st/sr
    mxx=mu_t+(mu_r-mu_t)*ux*ux
    myy=mu_t+(mu_r-mu_t)*uy*uy
    mxy=(mu_r-mu_t)*ux*uy
    er=eps/mu_r; et=eps/mu_t
    exx=et+(er-et)*ux*ux
    eyy=et+(er-et)*uy*uy
    exy=(er-et)*ux*uy
    ezz=eps*sr*st
    if photoelastic_delta is not None:
        # STEP6b: material-averaged real strain-optic epsilon offsets relative to n0^2.
        # PML radial coordinate stretch maps tensor from r/theta to global xy:
        # eps'_rr = eps_rr*st/sr, eps'_tt = eps_tt*sr/st, eps'_rt = eps_rt,
        # eps'_zz = eps_zz*sr*st. No invented xz/yz polarization components.
        pxx,pxy,pyy,pzz=photoelastic_delta
        drr=pxx*ux*ux+2*pxy*ux*uy+pyy*uy*uy
        dtt=pxx*uy*uy-2*pxy*ux*uy+pyy*ux*ux
        drt=(pyy-pxx)*ux*uy+pxy*(ux*ux-uy*uy)
        dr=drr*(st/sr)
        dt=dtt*(sr/st)
        exx=exx+dr*ux*ux+dt*uy*uy-2*drt*ux*uy
        eyy=eyy+dr*uy*uy+dt*ux*ux+2*drt*ux*uy
        exy=exy+(dr-dt)*ux*uy+drt*(ux*ux-uy*uy)
        ezz=ezz+pzz*sr*st
    return exx,exy,eyy,ezz,myy,-mxy,mxx,1/(sr*st)


def _assemble(design, raw, sbc_target_neff=None):
    try:
        from skfem import MeshTri, Basis, FacetBasis, BilinearForm, asm
        from skfem.element import ElementTriN1, ElementTriP1, ElementTriP0, ElementTriN2, ElementTriP2
        from skfem.helpers import curl, grad
    except ImportError as exc:
        raise RuntimeError("Install scikit-fem: pip install -r requirements.txt") from exc

    coords=np.asarray(raw["points_um"],dtype=float)
    triangles=np.asarray(raw["triangles"],dtype=int)
    if coords.ndim!=2 or coords.shape[1]!=2 or not np.isfinite(coords).all():raise ValueError("mesh points must be finite [N,2]")
    if triangles.ndim!=2 or triangles.shape[1]!=3 or not len(triangles) or triangles.min()<0 or triangles.max()>=len(coords):raise ValueError("mesh triangles must be valid [M,3] indices")
    if len(raw["material_ids"])!=len(triangles):raise ValueError("one material label per triangle required")
    mesh=MeshTri(coords.T, triangles.T)
    order=int(design.get("solver",{}).get("element_order",1))
    if order not in (1,2): raise ValueError("solver.element_order must be 1 or 2")
    ebe=ElementTriN1() if order==1 else ElementTriN2()
    eze=ElementTriP1() if order==1 else ElementTriP2()
    quadrature_order=4 if order==1 else 8  # Et^4 requires degree 8 for quadratic edge fields.
    be=Basis(mesh,ebe,intorder=quadrature_order)
    bz=Basis(mesh,eze,intorder=quadrature_order)
    b0=Basis(mesh,ElementTriP0(),intorder=quadrature_order)
    estimated=(be.N+bz.N if design["boundary"]["type"]=="sbc" else
               be.N+bz.N-len(be.get_dofs().all())-len(bz.get_dofs().all()))
    limit=design.get("solver",{}).get("max_dofs",80000)
    if estimated>limit:raise MemoryError(f"mixed DOFs {estimated} exceed solver.max_dofs={limit} before matrix assembly")
    materials=design["materials"]
    element_material=np.asarray(raw["material_ids"])
    eps=np.asarray([index(materials["background"] if name=="pml" else materials[name],
                          design["wavelength_nm"])**2
                    for name in element_material],dtype=complex)
    epsfield=b0.interpolate(eps)
    pm_fields={}
    if "pm_eps_delta" in raw:
        delta=np.asarray(raw["pm_eps_delta"],dtype=float)
        if delta.shape!=(len(triangles),4) or not np.isfinite(delta).all():
            raise ValueError("STEP6b optical epsilon tensor delta must be finite [Ntri,4]")
        pm_fields={k:b0.interpolate(delta[:,j]) for j,k in
                   enumerate(("pmxx","pmxy","pmyy","pmzz"))}
    k0=2*math.pi/(design["wavelength_nm"]*.001)

    def coefficients(w):
        return _tensor(
            w.x,w.eps,design,
            tuple(w[k] for k in ("pmxx","pmxy","pmyy","pmzz"))
            if pm_fields else None)
    assembly_kwargs={"eps":epsfield,**pm_fields}

    @BilinearForm(dtype=np.complex128)
    def ktt(u,v,w):
        ex,exy,ey,ez,gx,gxy,gy,mz=coefficients(w)
        return mz*curl(u)*curl(v)-k0*k0*(ex*u[0]*v[0]+exy*(u[0]*v[1]+u[1]*v[0])+ey*u[1]*v[1])

    @BilinearForm(dtype=np.complex128)
    def kzz(u,v,w):
        ex,exy,ey,ez,gx,gxy,gy,mz=coefficients(w)
        du,dv=grad(u),grad(v)
        return (gx*du[0]*dv[0]+gxy*(du[0]*dv[1]+du[1]*dv[0])+gy*du[1]*dv[1]
                -k0*k0*ez*u*v)

    @BilinearForm(dtype=np.complex128)
    def kez(u,v,w):
        ex,exy,ey,ez,gx,gxy,gy,mz=coefficients(w)
        du=grad(u)
        return 1j*(gx*du[0]*v[0]+gxy*(du[0]*v[1]+du[1]*v[0])+gy*du[1]*v[1])

    @BilinearForm(dtype=np.complex128)
    def kze(u,v,w):
        ex,exy,ey,ez,gx,gxy,gy,mz=coefficients(w)
        dv=grad(v)
        return -1j*(gx*u[0]*dv[0]+gxy*(u[0]*dv[1]+u[1]*dv[0])+gy*u[1]*dv[1])

    @BilinearForm(dtype=np.complex128)
    def mass_tt(u,v,w):
        ex,exy,ey,ez,gx,gxy,gy,mz=coefficients(w)
        return gx*u[0]*v[0]+gxy*(u[0]*v[1]+u[1]*v[0])+gy*u[1]*v[1]

    @BilinearForm(dtype=np.complex128)
    def gauss_e(u,v,w):
        ex,exy,ey,ez,gx,gxy,gy,mz=coefficients(w)
        dv=grad(v)
        return -(ex*u[0]*dv[0]+exy*(u[0]*dv[1]+u[1]*dv[0])+ey*u[1]*dv[1])

    @BilinearForm(dtype=np.complex128)
    def gauss_z(u,v,w):
        ex,exy,ey,ez,gx,gxy,gy,mz=coefficients(w)
        return ez*u*v

    @BilinearForm
    def scalar_mass(u,v,w): return u*v

    scalar_test_rows=np.setdiff1d(np.arange(bz.N),bz.get_dofs().all())
    mass=asm(scalar_mass,bz).tocsc()
    gauss_mass_factor=splu(mass[scalar_test_rows,:][:,scalar_test_rows].astype(complex).tocsc())
    Kee=asm(ktt,be,**assembly_kwargs).astype(complex)
    Kzz=asm(kzz,bz,**assembly_kwargs).astype(complex)
    Kez=asm(kez,bz,be,**assembly_kwargs).astype(complex) # edge row, axial col
    Kze=asm(kze,be,bz,**assembly_kwargs).astype(complex) # axial row, edge col
    Mee=asm(mass_tt,be,**assembly_kwargs).astype(complex)
    Ge=asm(gauss_e,be,bz,**assembly_kwargs).astype(complex)
    Gz=asm(gauss_z,bz,**assembly_kwargs).astype(complex)

    if design["boundary"]["type"]=="sbc":
        # Only a first-order fixed-target radiation approximation; not a DtN boundary.
        target=complex(sbc_target_neff if sbc_target_neff is not None else design.get("target_neff",index(materials["background"],design["wavelength_nm"]).real))
        clad=index(materials["background"],design["wavelength_nm"])
        kt=k0*cmath.sqrt(clad*clad-target*target+0j)
        @BilinearForm(dtype=np.complex128)
        def robin_e(u,v,w):
            ut=-w.n[1]*u[0]+w.n[0]*u[1]
            vt=-w.n[1]*v[0]+w.n[0]*v[1]
            return -1j*kt*ut*vt
        @BilinearForm(dtype=np.complex128)
        def robin_z(u,v,w):
            return -1j*kt*u*v
        Kee+=asm(robin_e,FacetBasis(mesh,ebe))
        Kzz+=asm(robin_z,FacetBasis(mesh,eze))

    ne,nz=be.N,bz.N
    ze=csc_matrix((ne,nz),dtype=complex)
    zz=csc_matrix((nz,nz),dtype=complex)
    K0=bmat([[Kee,ze],[ze.T,Kzz]],format="csc")
    K1=bmat([[None,Kez],[Kze,None]],format="csc")
    K2=bmat([[Mee,None],[None,zz]],format="csc")
    if design["boundary"]["type"]=="sbc":
        keep=np.arange(ne+nz) # radiation boundary: no PEC restriction
    else:
        bound_e=be.get_dofs().all()
        bound_z=bz.get_dofs().all()
        mask=np.ones(ne+nz,dtype=bool)
        mask[bound_e]=False
        mask[ne+bound_z]=False
        keep=np.flatnonzero(mask)
    if len(keep)<20:
        raise RuntimeError("too few interior mixed DOFs")
    return {"K0":K0[keep,:][:,keep],"K1":K1[keep,:][:,keep],
            "K2":K2[keep,:][:,keep],"keep":keep,"full_size":ne+nz,
            "ne":ne,"mesh":mesh,"be":be,"bz":bz,"eps":eps,"element_material":element_material,
            "pm_anisotropic_optics":bool(pm_fields),"element_order":order,
            "gauss_test_rows":scalar_test_rows,"gauss_mass_factor":gauss_mass_factor,
            "Ge":Ge,"Gz":Gz,"k0":k0}


def _field_samples(system,vector,beta,design,max_points,export_full_quadrature=False):
    be,bz=system["be"],system["bz"]
    ne=system["ne"]
    e=vector[:ne];z=vector[ne:]
    ef=be.interpolate(e);zf=bz.interpolate(z)
    et=np.asarray(ef);ez=np.asarray(zf);gz=zf.grad;cz=ef.curl
    curlx=gz[1]-1j*beta*et[1];curly=1j*beta*et[0]-gz[0]
    coefficients=_tensor(be.global_coordinates(),np.ones_like(ez),design)
    gxx,gxy,gyy,gzz=coefficients[4:]
    hx=(gxx*curlx+gxy*curly)/(1j*system["k0"]*Z0)
    hy=(gxy*curlx+gyy*curly)/(1j*system["k0"]*Z0)
    hz=gzz*cz/(1j*system["k0"]*Z0)
    six={"Ex":et[0],"Ey":et[1],"Ez":ez,"Hx":hx,"Hy":hy,"Hz":hz}
    dx=be.dx
    qpoints=np.moveaxis(np.asarray(be.global_coordinates()),0,-1)
    qphysical=np.broadcast_to(system["element_material"][:,None]!="pml",dx.shape)
    poynting=.5*np.real(et[0]*np.conj(hy)-et[1]*np.conj(hx))
    input_power=float(np.sum(poynting[qphysical]*dx[qphysical])*1e-12)
    if not np.isfinite(input_power) or input_power<=0:raise RuntimeError("nonfinite/non-forward physical optical power")
    ep=np.concatenate([np.asarray(six[key])[qphysical] for key in ("Ex","Ey","Ez")])
    phase=np.angle(ep[np.argmax(np.abs(ep))])
    scale=np.exp(-1j*phase)/math.sqrt(input_power)
    vector*=scale
    six={key:val*scale for key,val in six.items()}
    if not all(np.isfinite(v).all() for v in six.values()):raise RuntimeError("nonfinite E/H field")
    normalized_power=float(np.sum((.5*np.real(six["Ex"]*np.conj(six["Hy"])-six["Ey"]*np.conj(six["Hx"])))[qphysical]*dx[qphysical])*1e-12)
    eq=sum(np.abs(six[k])**2 for k in ("Ex","Ey","Ez"))
    hq=sum(np.abs(six[k])**2 for k in ("Hx","Hy","Hz"))
    ez_fraction=float(np.sum((np.abs(six["Ez"])**2*dx)[qphysical])/np.sum((eq*dx)[qphysical]))
    hz_fraction=float(np.sum((np.abs(six["Hz"])**2*dx)[qphysical])/np.sum((hq*dx)[qphysical]))
    stokes0=float(np.sum(((np.abs(six["Ex"])**2+np.abs(six["Ey"])**2)*dx)[qphysical]))
    stokes1=float(np.sum(((np.abs(six["Ex"])**2-np.abs(six["Ey"])**2)*dx)[qphysical]))
    cross=np.sum((six["Ex"]*np.conj(six["Ey"])*dx)[qphysical])
    stokes2,stokes3=float(2*cross.real),float(-2*cross.imag)
    # Quadrature cell averages, not nodal values or exact centroid evaluation.
    fields={key:np.sum(val*dx,axis=-1)/np.sum(dx,axis=-1) for key,val in six.items()}
    points=system["mesh"].p.T
    tris=system["mesh"].t.T
    centroids=points[tris].mean(axis=1)
    a,b,c=points[tris[:,0]],points[tris[:,1]],points[tris[:,2]]
    area=np.abs((b[:,0]-a[:,0])*(c[:,1]-a[:,1])-(b[:,1]-a[:,1])*(c[:,0]-a[:,0]))*.5
    intensity=np.abs(fields["Ex"])**2+np.abs(fields["Ey"])**2+np.abs(fields["Ez"])**2
    total=float(np.sum(intensity*area))
    if not math.isfinite(total) or total<=0: raise RuntimeError("nonfinite/zero optical field")
    frac={key:float(np.sum(np.abs(fields[key])**2*area)/total) for key in ("Ex","Ey","Ez")}
    r=np.linalg.norm(centroids,axis=1)
    in_pml=(system["element_material"]=="pml")
    pml_fraction=float(np.sum(intensity[in_pml]*area[in_pml])/total)
    physical=~in_pml
    # STEP 4 v0.8.0: output Aeff/MFD from transverse Et ONLY, not Ez or total E.
    transverse=np.abs(fields["Ex"])**2+np.abs(fields["Ey"])**2
    ip=transverse[physical];ap=area[physical]
    power=float(np.sum(ip*ap));a2=float(np.sum(ip*ip*ap))
    aeff=power*power/a2 if a2>0 else float("nan")
    pts=centroids[physical]
    cx=float(np.sum(pts[:,0]*ip*ap)/power)
    cy=float(np.sum(pts[:,1]*ip*ap)/power)
    moment=float(np.sum(((pts[:,0]-cx)**2+(pts[:,1]-cy)**2)*ip*ap)/power)
    mfd=2*math.sqrt(2*moment)
    chosen=np.linspace(0,len(centroids)-1,min(max_points,len(centroids)),dtype=int)
    samples={"coordinates_um":np.round(centroids[chosen],5).tolist(),
             "sampling":"quadrature cell average plotted at triangle centroid; not a nodal field",
             "components":{name:[[round(float(v.real),10),round(float(v.imag),10)]
                         for v in data[chosen]] for name,data in fields.items()}}
    quadrature=None
    if export_full_quadrature:
        # This can be a very large JSON: use only on Colab/local for mode-following.
        def serial(values):
            return [[float(v.real),float(v.imag)] for v in values[qphysical]]
        quadrature={"points_um":qpoints[qphysical].tolist(),
                    "weights_um2":dx[qphysical].tolist(),
                    **{key:serial(value) for key,value in six.items()},
                    "material_labels":np.broadcast_to(system["element_material"][:,None],dx.shape)[qphysical].tolist(),
                    "sampling":"actual FEM physical-domain Gauss quadrature; six components, material labels",
                    "quadrature_schema":"LSCNS_FE_PHYSICAL_QUADRATURE_V086",
                    "scope":"all physical-region FE Gauss quadrature points; PML excluded",
                    "warning":"field quadrature is not automatically remeshed; use identical probes for tracking"}
    return samples,{"Aeff_um2":aeff,"MFD_2nd_um":mfd,
                    "Aeff_basis":"Et=Ex,Ey only; FEM cell-average diagnostic; PML excluded",
                    "power_W":normalized_power,"normalization_error":abs(normalized_power-1.),"normalization":"physical-domain longitudinal Poynting flux = 1 W; global phase fixed",
                    "finite_six_components":True,
                    "longitudinal_E_fraction":ez_fraction,"longitudinal_H_fraction":hz_fraction,
                    "field_content_class":"TE-like" if ez_fraction<1e-6 and hz_fraction>=1e-6 else
                                          "TM-like" if hz_fraction<1e-6 and ez_fraction>=1e-6 else "hybrid_or_quasi_TEM",
                    "field_content_basis":"longitudinal energy fractions with heuristic 1e-6 threshold, not a symmetry irrep",
                    "integrated_transverse_Stokes":{"S1_over_S0":stokes1/stokes0,"S2_over_S0":stokes2/stokes0,"S3_over_S0":stokes3/stokes0,
                         "axis_angle_deg":float(.5*np.degrees(np.arctan2(stokes2,stokes1))),"S3_definition":"2 Im(conj(Ex) Ey)"},
                    "MFD_PetermannII_status":"TODO: full radial field or verified far-field transform required",
                    "polarization_fractions":frac,
                    "PML_intensity_fraction":pml_fraction,
                    "polarization_class":"quasi_x" if frac["Ex"]>.85 else
                      "quasi_y" if frac["Ey"]>.85 else "mixed",
                    "symmetry_class":"unclassified (requires field overlap / group action)"},fields,quadrature


def solve(design,mesh=None,n_modes=2,target_neff=None,max_sample_points=800,export_full_quadrature=False,probe_points=None,probe_weights=None):
    """Stable public entry; FEM assembly retained, v0.8.5 diagnostics delegated."""
    try: from .vector_solver_v085 import solve_modes
    except ImportError: from vector_solver_v085 import solve_modes
    return solve_modes(design,mesh,n_modes,target_neff,max_sample_points,export_full_quadrature,probe_points,probe_weights)
