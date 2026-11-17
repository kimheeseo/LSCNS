"""Experimental edge-element mixed vector optical mode eigenproblem.

E=(E_t,E_z) exp(i beta z-i omega t). HCurl(Nedelec-1) for E_t,
H1(P1) for E_z. Quadratic polynomial in beta assembled with weak Maxwell:
curl(mu^-1 curl E)-k0^2 eps E=0. Sparse companion shift-invert via scipy.
Not validated for precision PML leakage; inspect convergence and weak Gauss residual.
"""
import json
import math
import cmath
from pathlib import Path

import numpy as np
from scipy.sparse import bmat, csc_matrix, eye, diags
from scipy.sparse.linalg import eigs, splu, LinearOperator

try:
    from .schema import validate_design
    from .geometry import build_mesh
    from .materials import index, provenance
    from .physical_metrics import confinement_loss
except ImportError:
    from schema import validate_design
    from geometry import build_mesh
    from materials import index, provenance
    from physical_metrics import confinement_loss

Z0 = 376.730313668

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


def _assemble(design, raw):
    try:
        from skfem import MeshTri, Basis, FacetBasis, BilinearForm, asm
        from skfem.element import ElementTriN1, ElementTriP1, ElementTriP0
        from skfem.helpers import curl, grad
    except ImportError as exc:
        raise RuntimeError("Install scikit-fem: pip install -r requirements.txt") from exc

    coords=np.asarray(raw["points_um"],dtype=float)
    triangles=np.asarray(raw["triangles"],dtype=int)
    mesh=MeshTri(coords.T, triangles.T)
    ebe=ElementTriN1(); eze=ElementTriP1()
    be=Basis(mesh,ebe,intorder=4)
    bz=Basis(mesh,eze,intorder=4)
    b0=Basis(mesh,ElementTriP0(),intorder=4)
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

    Kee=asm(ktt,be,**assembly_kwargs).astype(complex)
    Kzz=asm(kzz,bz,**assembly_kwargs).astype(complex)
    Kez=asm(kez,bz,be,**assembly_kwargs).astype(complex) # edge row, axial col
    Kze=asm(kze,be,bz,**assembly_kwargs).astype(complex) # axial row, edge col
    Mee=asm(mass_tt,be,**assembly_kwargs).astype(complex)
    Ge=asm(gauss_e,be,bz,**assembly_kwargs).astype(complex)
    Gz=asm(gauss_z,bz,**assembly_kwargs).astype(complex)

    if design["boundary"]["type"]=="sbc":
        # Only a first-order fixed-target radiation approximation; not a DtN boundary.
        target=complex(design.get("target_neff",index(materials["background"],design["wavelength_nm"]).real))
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
            "ne":ne,"mesh":mesh,"be":be,"bz":bz,"eps":eps,
            "pm_anisotropic_optics":bool(pm_fields),
            "Ge":Ge,"Gz":Gz,"k0":k0}


def _field_samples(system,vector,beta,design,max_points,export_full_quadrature=False):
    be,bz=system["be"],system["bz"]
    ne=system["ne"]
    e=vector[:ne];z=vector[ne:]
    ef=be.interpolate(e);zf=bz.interpolate(z)
    et=ef.value;ez=zf.value;gz=zf.grad;cz=ef.curl
    hx=(gz[1]-1j*beta*et[1])/(1j*system["k0"]*Z0)
    hy=(1j*beta*et[0]-gz[0])/(1j*system["k0"]*Z0)
    hz=cz/(1j*system["k0"]*Z0)
    six={"Ex":et[0],"Ey":et[1],"Ez":ez,"Hx":hx,"Hy":hy,"Hz":hz}
    # Quadrature-point values; do not mislabel as nodal DOFs.
    fields={key:np.mean(val,axis=-1) for key,val in six.items()}
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
    in_pml=(r>float(design["domain_radius_um"]))
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
             "sampling":"triangle centroid, mean of finite-element quadrature values",
             "components":{name:[[round(float(v.real),10),round(float(v.imag),10)]
                         for v in data[chosen]] for name,data in fields.items()}}
    quadrature=None
    if export_full_quadrature:
        # This can be a very large JSON: use only on Colab/local for mode-following.
        def serial(values):
            return [[float(z.real),float(z.imag)] for z in values[physical]]
        quadrature={"points_um":centroids[physical].tolist(),
                    "weights_um2":area[physical].tolist(),
                    "Ex":serial(fields["Ex"]),"Ey":serial(fields["Ey"]),
                    "scope":"full physical-region triangle centroids, approximate FE quadrature",
                    "warning":"test field overlap + h refinement; no automatic remeshing interpolation"}
    return samples,{"Aeff_um2":aeff,"MFD_2nd_um":mfd,
                    "Aeff_basis":"Et=Ex,Ey only; FEM triangle-centroid quadrature; PML excluded",
                    "MFD_PetermannII_status":"TODO: full radial field or verified far-field transform required",
                    "polarization_fractions":frac,
                    "PML_intensity_fraction":pml_fraction,
                    "polarization_class":"quasi_x" if frac["Ex"]>.85 else
                      "quasi_y" if frac["Ey"]>.85 else "mixed",
                    "symmetry_class":"unclassified (requires field overlap / group action)"},fields,quadrature


def solve(design,mesh=None,n_modes=2,target_neff=None,max_sample_points=800,export_full_quadrature=False):
    validate_design(design)
    if mesh is None: mesh=build_mesh(design)
    s=_assemble(design,mesh)
    K0,K1,K2=(s["K"+str(k)] for k in range(3))
    n=K0.shape[0]
    if n>22000:
        raise MemoryError("sparse companion too large (>44,000 DOFs); use smaller mesh or HPC")
    if not 1<=n_modes<=8: raise ValueError("n_modes 1..8")
    k0=s["k0"]
    target=target_neff if target_neff is not None else design.get("target_neff")
    if target is None: target=index(design["materials"].get("core",design["materials"]["background"]),design["wavelength_nm"]).real-.001
    if not 0<float(target)<3.5: raise ValueError("target neff invalid")
    sigma=k0*float(target)
    I=eye(n,format="csc",dtype=complex)
    Z=csc_matrix((n,n),dtype=complex)
    A=bmat([[Z,I],[-K0,-K1]],format="csc")
    B=bmat([[I,Z],[Z,K2]],format="csc")
    # OP = (A - sigma B)^-1 B, mu = (beta - sigma)^-1.
    factor=splu(A-sigma*B)
    op=LinearOperator((2*n,2*n),matvec=lambda x:factor.solve(B@x),
                      dtype=complex)
    mu,vectors=eigs(op,k=min(2*n_modes+8,2*n-2),which="LM",tol=1e-8,maxiter=2500)
    found=[]
    for j,value in enumerate(mu):
        if abs(value)<1e-14 or not np.isfinite(value): continue
        beta=sigma+1/value
        neff=beta/k0
        if not (0<neff.real<3.5) or abs(neff.imag)>.05: continue
        e_reduced=vectors[:n,j]
        quad=K0@e_reduced+beta*K1@e_reduced+beta*beta*K2@e_reduced
        denom=np.linalg.norm(K0@e_reduced)+abs(beta)*np.linalg.norm(K1@e_reduced)+abs(beta)**2*np.linalg.norm(K2@e_reduced)
        residual=float(np.linalg.norm(quad)/max(denom,1e-30))
        if not math.isfinite(residual) or residual>1e-5: continue
        full=np.zeros(s["full_size"],dtype=complex)
        full[s["keep"]]=e_reduced
        e=full[:s["ne"]];z=full[s["ne"]:]
        g_e=s["Ge"]@e
        g_z=1j*beta*(s["Gz"]@z)
        div_res=float(np.linalg.norm(g_e+g_z)/(np.linalg.norm(g_e)+np.linalg.norm(g_z)+1e-30))
        try: samples,info,fields,quadrature=_field_samples(s,full,beta,design,max_sample_points,export_full_quadrature=export_full_quadrature)
        except RuntimeError: continue
        status="candidate_unverified"
        if neff.imag < -1e-9 or info["PML_intensity_fraction"]>.4 or div_res>5e-2:
            status="spurious_or_unstable_candidate"
        signed_loss=confinement_loss(float(neff.imag),design["wavelength_nm"])
        record=dict(neff_real=float(neff.real),neff_imag=float(neff.imag),
                    alpha_total_dB_m=signed_loss["alpha_total_dB_m"],
                    alpha_total_dB_km=signed_loss["alpha_total_dB_km"],
                    alpha_basis="total Im(neff): radiative + material; PML/convergence unverified",
                    residual_relative=residual,weak_gauss_relative=div_res,
                    classification=info,field_samples=samples,status=status)
        if quadrature is not None: record["field_quadrature"]=quadrature
        found.append(record)
    if not found: raise RuntimeError("no converged finite candidate modes; inspect mesh/PML, target, divergence")
    found.sort(key=lambda m:(m["status"]!="candidate_unverified",
                              abs(m["neff_real"]-float(target)),m["residual_relative"]))
    unique=[]
    for mode in found:
        if any(abs(mode["neff_real"]-z["neff_real"])<1e-10 and abs(mode["neff_imag"]-z["neff_imag"])<1e-10 for z in unique):
            continue
        unique.append(mode)
        if len(unique)>=n_modes: break
    return {"schema":"LSCNS_fiber_research_vector_v1","schema_version":"1.0.0","engine_version":"0.8.2" if s["pm_anisotropic_optics"] else
            ("0.8.1" if design.get("_bend") is not None else "0.8.0"),
            "method":"2D HCurl(Nedelec1)+H1(P1) transverse/longitudinal E mixed quadratic beta FEM, scipy sparse shift-invert",
            "convention":"exp(i beta z - i omega t); Im(neff)>0 is passive loss",
            "validation":{"level":"EXPERIMENTAL_NOT_BENCHMARKED",
                          "mesh_converged":False,"pml_converged":False,
                          "loss_dB_km_certified":False,
                          "warnings":["SBC uses approximate fixed-target Robin impedance" if design["boundary"]["type"]=="sbc" else
                                      "PML loss requires thickness, strength, mesh and physical domain convergence",
                                      "Spurious vector modes require independent HE11 and divergence/overlap benchmarks",
                                      "No computed Bloch bandgap or stress-optic birefringence"] +
                                     (["2D conformal macro-bend is weak-curvature approximation, NOT COMSOL axisymmetric bend V3",
                                       "R_eff stress multiplier must be user-supplied with applicable citation; bent loss and turn loss NOT certified"]
                                      if design.get("_bend") is not None else [])},
            "input":design,"mesh_quality":mesh["quality"],"modes":unique,
            "photoelastic_coupling":{"enabled":s["pm_anisotropic_optics"],
                                    "status":"EXPERIMENTAL_ANISOTROPIC_MAXWELL_NOT_BENCHMARKED"
                                    if s["pm_anisotropic_optics"] else "disabled",
                                    "source_sha256":mesh.get("pm_source_sha256"),
                                    "basis":"physical epsilon tensor from thermomechanical strain-optic, per-element, PML tensor transformed"},
            "materials_provenance":{key:provenance(mat) for key,mat in design["materials"].items()}}
