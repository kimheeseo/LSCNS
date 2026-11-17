"""STEP 6b / LSCNS v0.8.2 — generalized-plane-strain thermal FEM and strain optic.

SI material input ONLY (all E_Pa, nu, alpha_1_per_K, p11, p12 user-supplied
with independent source_url). No guessed Panda rod properties.
Coordinates in mesh are um; numerical elasticity assembly uses x/L0 and E/Eref
for conditioning; displacement output um; stresses Pa; strain unitless.

Thermal generalized plane strain (free axial resultant, no imposed shear):
eps = [[ux,x, (ux,y+uy,x)/2, 0],
       [(ux,y+uy,x)/2, uy,y, 0],
       [0, 0, eps_zz_const]]
lambda = E nu/((1+nu)(1-2nu)); mu=E/(2(1+nu)).
sigma = 2 mu(eps-alpha*DeltaT I)+lambda(tr(eps)-3alpha*DeltaT) I.
Weak in-plane system K u + C eps_zz = f_T.
Free axial resultant C^T u + G eps_zz = h_T, i.e. integral sigma_zz dA=0.
Outer surface traction-free, with only 3 in-plane gauge pins removing
rigid translations + rotation. PML excluded entirely from mechanics.

Photoelastic (strain-optic, NOT stress-optic p coefficients):
delta Q_ii = p11 eps_ii + p12 sum_{j!=i} eps_jj,
delta Q_xy = (p11-p12) eps_xy   (p44=(p11-p12)/2 Voigt).
Q=eps_r^-1 = I/n0^2+delta Q; epsilon_r = Q^-1; 2x2 transverse
principal n from Q eigenvalues. Spatial core average of Q is a
MATERIAL birefringence proxy, never a modal effective-index splitting.

References:
COMSOL generalized-plane-strain: 
https://doc.comsol.com/6.4/doc/com.comsol.help.models.sme.stress_optical_generalized/stress_optical_generalized.html
COMSOL photoelastic waveguide:
https://doc.comsol.com/6.4/doc/com.comsol.help.models.woptics.stress_optical/stress_optical.html
Isotropic p tensor and conventions:
https://pmc.ncbi.nlm.nih.gov/articles/PMC12203310/
"""
from __future__ import annotations
from datetime import datetime, timezone
from pathlib import Path
import hashlib
import json
import math
import numpy as np
from scipy.sparse import coo_matrix, bmat, csc_matrix
from scipy.sparse.linalg import spsolve
try:
    from .materials import index
except ImportError:
    from materials import index

ENGINE_VERSION="0.8.2"
MODEL="generalized_plane_strain_uniform_axial_thermal"
SOURCE_COMSOL="https://doc.comsol.com/6.4/doc/com.comsol.help.models.sme.stress_optical_generalized/stress_optical_generalized.html"


def _num(obj,key,lo=None,hi=None):
    v=obj.get(key)
    if isinstance(v,bool) or not isinstance(v,(float,int)) or not math.isfinite(v):
        raise ValueError(f"{key}: finite numeric value required; TODO cite source")
    if (lo is not None and v<=lo) or (hi is not None and v>=hi):
        raise ValueError(f"{key}: outside supported interval")
    return float(v)


def _url(obj,key):
    val=obj.get(key)
    if not isinstance(val,str) or not val.startswith(("https://","http://")) or len(val)>1024:
        raise ValueError(f"{key}: explicit source URL required; TODO material provenance")
    return val


def _signed_area(p):
    a,b,c=p[:,0],p[:,1],p[:,2]
    return .5*((b[:,0]-a[:,0])*(c[:,1]-a[:,1])-(b[:,1]-a[:,1])*(c[:,0]-a[:,0]))


def validate_pm_config(config,mesh,design):
    if config.get("schema")!="LSCNS_step6b_materials_v1":
        raise ValueError("expected LSCNS_step6b_materials_v1")
    dT=_num(config,"delta_T_K")
    if abs(dT)>2000:
        raise ValueError("delta_T_K outside safety guard; supply calibrated model")
    if design.get("family")!="polarization_maintaining":
        raise ValueError("STEP6b presently supports PM design family only")
    if design.get("variant") not in ("panda","bow_tie","elliptic_core"):
        raise ValueError("unsupported PM geometry")
    raw_ids=np.asarray(mesh["material_ids"],dtype=str)
    required=sorted(set(raw_ids)-{"pml"})
    if not required or "core" not in required or "background" not in required:
        raise ValueError("core/background mechanical domains required")
    if design["variant"] in ("panda","bow_tie") and "stress_rod" not in required:
        raise ValueError("stress_rod is missing in gmsh physical groups; old meshes unsupported")
    if not isinstance(config.get("materials"),dict):
        raise ValueError("materials mapping required")
    props={}
    photo_complete=True
    for name in required:
        v=config["materials"].get(name)
        if not isinstance(v,dict):
            raise ValueError(f"{name}: missing material parameters, TODO")
        E=_num(v,"E_Pa",lo=0)
        nu=_num(v,"poisson_ratio",lo=-.99,hi=.4999)
        alpha=_num(v,"alpha_1_per_K")
        _url(v,"mechanical_source_url")
        p11=v.get("p11");p12=v.get("p12")
        if p11 is None or p12 is None:
            photo_complete=False
            p11=p12=None
        else:
            p11=_num(v,"p11");p12=_num(v,"p12")
            if abs(p11)>10 or abs(p12)>10:
                raise ValueError(f"{name}: photoelastic coefficient outside sanity range")
            _url(v,"photoelastic_source_url")
        if name not in design["materials"]:
            raise ValueError(f"optical material {name} absent in design; no guessed refractive index")
        # n0 is a reference index not necessarily thermo-optically updated.
        n0=index(design["materials"][name],design["wavelength_nm"])
        if abs(n0.imag)>1e-14:
            raise NotImplementedError("complex absorptive n0 cannot be used for real stress-optic Q; TODO")
        props[name]={"E_Pa":E,"nu":nu,"alpha":alpha,"p11":p11,"p12":p12,"n0":float(n0.real)}
    for k in config["materials"]:
        if k not in required and k!="pml":
            raise ValueError(f"unexpected material {k} in mechanics inputs")
    return dT,props,photo_complete


def isotropic_photoelastic(n0,p11,p12,strain_xx,strain_yy,strain_zz,strain_xy):
    """Return Q 3x3 and epsilon_r 3x3, principal transverse refractive indices."""
    if not all(math.isfinite(float(v)) for v in
               (n0,p11,p12,strain_xx,strain_yy,strain_zz,strain_xy)) or n0<=0:
        raise ValueError("finite tensor parameters and positive n0 required")
    e=np.array([strain_xx,strain_yy,strain_zz],dtype=float)
    q0=1/(n0*n0)
    d=(p11-p12)*e+p12*np.sum(e)
    shear=(p11-p12)*strain_xy
    Q=np.diag(q0+d)
    Q[0,1]=Q[1,0]=shear
    eigen=np.linalg.eigvalsh(Q)
    if np.min(eigen)<=0:raise ValueError("nonphysical negative dielectric impermeability eigenvalue")
    eps=np.linalg.inv(Q)
    nt=1/np.sqrt(np.linalg.eigvalsh(Q[:2,:2]))
    return Q,eps,tuple(sorted(map(float,nt),reverse=True))


def _assemble_triangle(points,tri,mat):
    """CST generalized plane strain with nondimensional geometry and E/Eref."""
    xyz=points[tri]
    x,y=xyz[:,0],xyz[:,1]
    signed=.5*((x[1]-x[0])*(y[2]-y[0])-(y[1]-y[0])*(x[2]-x[0]))
    A=abs(signed)
    if A<=1e-15:raise ValueError("zero-area elastic element")
    det=2*signed
    dNdx=np.array([y[1]-y[2],y[2]-y[0],y[0]-y[1]])/det
    dNdy=np.array([x[2]-x[1],x[0]-x[2],x[1]-x[0]])/det
    B=np.zeros((3,6))
    for k in range(3):
        B[0,2*k]=dNdx[k]
        B[1,2*k+1]=dNdy[k]
        B[2,2*k]=dNdy[k]   # gamma_xy engineering
        B[2,2*k+1]=dNdx[k]
    lam=mat["lambda"];mu=mat["mu"];a=mat["alpha_dT"]
    D=np.array([[lam+2*mu,lam,0.],[lam,lam+2*mu,0.],[0.,0.,mu]])
    K=A*B.T@D@B
    C=A*B.T@np.array([lam,lam,0.])
    thermal=(3*lam+2*mu)*a
    F=A*B.T@np.array([thermal,thermal,0.])
    G=A*(lam+2*mu)
    H=A*thermal
    return A,B,K,C,F,G,H


def solve_thermal_stress(design,mesh,config,max_samples=1200):
    """Area-weighted thermal FEM + photoelastic tensor; modal B remains TODO.

    PML nodes removed; no numerical material constants inserted.
    Returns (JSON-compatible summary, per-full-element optical epsilon delta).
    """
    dT,props,photo_complete=validate_pm_config(config,mesh,design)
    pts_um=np.asarray(mesh["points_um"],dtype=float)
    tri_all=np.asarray(mesh["triangles"],dtype=int)
    labels=np.asarray(mesh["material_ids"],dtype=str)
    if pts_um.ndim!=2 or pts_um.shape[1]!=2 or tri_all.ndim!=2 or tri_all.shape[1]!=3:
        raise ValueError("mesh points Nx2 / triangles Mx3 required")
    if len(labels)!=len(tri_all) or len(tri_all)<2 or not np.isfinite(pts_um).all():
        raise ValueError("invalid physical mesh")
    mask=labels!="pml"
    physical_elements=np.flatnonzero(mask)
    physical_tris=tri_all[mask]
    phy_labels=labels[mask]
    vertices=np.unique(physical_tris)
    reverse=np.full(len(pts_um),-1,dtype=int)
    reverse[vertices]=np.arange(len(vertices))
    tri=reverse[physical_tris]
    xy_um=pts_um[vertices]
    L0=max(float(np.max(np.abs(xy_um))),1e-12)
    xy=xy_um/L0
    E_ref=max(p["E_Pa"] for p in props.values())
    norm={}
    for name,p in props.items():
        en=p["E_Pa"]/E_ref
        nu=p["nu"]
        norm[name]={"lambda":en*nu/((1+nu)*(1-2*nu)),
                    "mu":en/(2*(1+nu)),"alpha_dT":p["alpha"]*dT}
    n=len(vertices); ndof=2*n
    rr=[];cc=[];vv=[]; rhs=np.zeros(ndof)
    Cvec=np.zeros(ndof); G=0.; H=0.; areas=np.empty(len(tri))
    Bs=[]
    for k,(t,lab) in enumerate(zip(tri,phy_labels)):
        A,B,K,C,F,g,h=_assemble_triangle(xy,t,norm[lab])
        areas[k]=A
        Bs.append(B)
        dof=np.array([q for i in t for q in (2*i,2*i+1)],dtype=int)
        rr.extend(np.repeat(dof,6));cc.extend(np.tile(dof,6));vv.extend(K.ravel())
        rhs[dof]+=F;Cvec[dof]+=C
        G+=g;H+=h
    K=coo_matrix((vv,(rr,cc)),shape=(ndof,ndof)).tocsc()
    A=bmat([[K,csc_matrix(Cvec.reshape(-1,1))],
            [csc_matrix(Cvec.reshape(1,-1)),csc_matrix([[G]])]],format="csc")
    rhsfull=np.concatenate((rhs,[H]))
    # Remove 2 translations and 1 in-plane rotation (gauge only).
    anchor=int(np.argmin(np.sum(xy*xy,axis=1)))
    dif=xy[:,0]-xy[anchor,0]
    other=int(np.argmax(np.abs(dif)))
    if abs(dif[other])<1e-9:raise ValueError("cannot fix rotational rigid-body mode")
    fixed={2*anchor,2*anchor+1,2*other+1}
    free=np.asarray([i for i in range(ndof+1) if i not in fixed])
    solution=np.zeros(ndof+1)
    solution[free]=spsolve(A[free,:][:,free],rhsfull[free])
    if not np.isfinite(solution).all():
        raise RuntimeError("thermal mechanics linear solver failed")
    u=solution[:ndof]
    epszz=float(solution[-1])
    strain=np.zeros((len(tri),4))
    stress=np.zeros((len(tri),4))
    pts=xy_um[tri].mean(axis=1)
    delta_eps=np.zeros((len(tri_all),4),dtype=float)
    qcore=np.zeros((3,3),float)
    acore=0.
    samples=[]
    maxstress=0.
    axial_resultant=0.
    for k,(t,lab) in enumerate(zip(tri,phy_labels)):
        dof=np.array([q for i in t for q in (2*i,2*i+1)])
        exx,eyy,gxy=Bs[k]@u[dof]
        exy=.5*gxy
        p=props[lab]
        lam=norm[lab]["lambda"];mu=norm[lab]["mu"];a=norm[lab]["alpha_dT"]
        tstrain=exx+eyy+epszz
        sxx=(2*mu*(exx-a)+lam*(tstrain-3*a))*E_ref
        syy=(2*mu*(eyy-a)+lam*(tstrain-3*a))*E_ref
        sxy=2*mu*exy*E_ref
        szz=(2*mu*(epszz-a)+lam*(tstrain-3*a))*E_ref
        strain[k]=[exx,eyy,epszz,exy]
        stress[k]=[sxx,syy,szz,sxy]
        maxstress=max(maxstress,abs(sxx),abs(syy),abs(szz),abs(sxy))
        axial_resultant+=szz*areas[k]
        nloc=None
        if photo_complete:
            Q,eps,nloc=isotropic_photoelastic(p["n0"],p["p11"],p["p12"],
                                               exx,eyy,epszz,exy)
            epsref=p["n0"]**2
            delta_eps[physical_elements[k]]=[
                eps[0,0]-epsref,eps[0,1],eps[1,1]-epsref,eps[2,2]-epsref]
            if lab=="core":
                qcore+=Q*areas[k]
                acore+=areas[k]
        samples.append({"point_um":[float(pts[k,0]),float(pts[k,1])],
                        "material":str(lab),"stress_Pa":[float(sxx),float(syy),float(szz),float(sxy)],
                        "strain":[float(exx),float(eyy),float(epszz),float(exy)],
                        "local_principal_n":list(nloc) if nloc is not None else None})
    coreB={"status":"TODO_PHOTOELASTIC","required":"all p11, p12 and their source URLs for each mechanical material"}
    if photo_complete:
        if acore<=0:raise RuntimeError("no core triangles")
        Qbar=qcore/acore
        nbars=1/np.sqrt(np.linalg.eigvalsh(Qbar[:2,:2]))
        nfast=float(min(nbars)); nslow=float(max(nbars))
        bv=nslow-nfast
        coreB={"status":"MATERIAL_PROXY_ONLY","B_core_area_weighted":bv,
               "n_principal_slow":nslow,"n_principal_fast":nfast,
               "beat_length_proxy_m":design["wavelength_nm"]*1e-9/bv if bv>0 else None,
               "Q_core_area_weighted":Qbar.tolist(),
               "note":"n tensor core-area average, NOT full-vector modal neff_x/neff_y; geometric birefringence not included"}
    settings={"design":design,"materials":config}
    canon=json.dumps(settings,ensure_ascii=False,sort_keys=True,separators=(",",":"),allow_nan=False)
    fingerprint=hashlib.sha256(canon.encode("utf-8")).hexdigest()
    stride=np.linspace(0,len(samples)-1,min(max_samples,len(samples)),dtype=int)
    result={"schema":"LSCNS_step6b_stress_result_v082","engine_version":ENGINE_VERSION,
            "timestamp_utc":datetime.now(timezone.utc).isoformat(),"input_sha256":fingerprint,
            "source_design":design,"material_config":config,
            "method":MODEL,
            "mechanics":{"delta_T_K":dT,"axial_strain_uniform":epszz,
                         "max_abs_stress_Pa":maxstress,
                         "normalized_axial_resultant_Pa_area_fraction":axial_resultant/max(sum(areas),1e-30),
                         "displacement_um_max":float(np.max(np.linalg.norm(u.reshape(-1,2)*L0,axis=1))),
                         "physical_triangles":len(tri),"gauge":"2 translations + in-plane rotation fixed; exterior traction free",
                         "source":SOURCE_COMSOL},
            "photoelastic":{"all_coefficients_present":photo_complete,
                            "core_material_birefringence":coreB},
            "field_samples":[samples[i] for i in stride],
            "validation":{"mechanics_unit_tests":"SEE_CI","optical_modal_birefringence":"NOT_CERTIFIED",
                          "COMSOL_comparison":"EVIDENCE_REQUIRED",
                          "V4_PML":"NOT_RUN","V5_MESH":"NOT_RUN",
                          "notes":["linear isotropic elasticity generalized plane strain, no z-warping",
                                   "uniform axial epsilon_zz from net axial force zero",
                                   "isotropic strain-optic p11/p12 only, no dn/dT material baseline",
                                   "core-area Q mean is NOT true modal birefringence",
                                   "no thermo-viscoelastic cooling/anneal/fictive temperature"]}}
    return result,delta_eps if photo_complete else None
