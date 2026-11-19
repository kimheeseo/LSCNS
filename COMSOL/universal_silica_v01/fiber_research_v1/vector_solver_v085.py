"""v0.8.5 diagnostics and field-based mode selection around existing FEM assembly."""
import math
import numpy as np
try:
    from .schema import validate_design
    from .geometry import build_mesh
    from .materials import index,provenance
    from .physical_metrics import confinement_loss
    from .vector_spectrum import spectrum
    from .modal_diagnostics import disk_probes,weighted_columns,independent_indices
except ImportError:
    from schema import validate_design
    from geometry import build_mesh
    from materials import index,provenance
    from physical_metrics import confinement_loss
    from vector_spectrum import spectrum
    from modal_diagnostics import disk_probes,weighted_columns,independent_indices

ENGINE_VERSION="0.8.6"


def probe_e(system,full,points):
    xy=np.asarray(points,dtype=float).T
    et=np.asarray(system["be"].probes(xy)@full[:system["ne"]]).reshape(2,-1)
    ez=np.asarray(system["bz"].probes(xy)@full[system["ne"]:])
    result=np.column_stack((et.T,ez))
    if not np.isfinite(result).all():raise RuntimeError("nonfinite FE probe field")
    return result


def _solve_once(design,mesh=None,n_modes=2,target_neff=None,max_sample_points=800,
                export_full_quadrature=False,probe_points=None,probe_weights=None,sbc_target=None):
    try:
        from .vector_fem import _assemble,_field_samples
    except ImportError:
        from vector_fem import _assemble,_field_samples
    validate_design(design)
    if mesh is None:mesh=build_mesh(design)
    options=design.get("solver",{})
    if not isinstance(n_modes,int) or isinstance(n_modes,bool) or not 1<=n_modes<=64:raise ValueError("n_modes 1..64 (large spectra remain subject to solver resource limits)")
    target=target_neff if target_neff is not None else design.get("target_neff")
    if target is None:target=index(design["materials"].get("core",design["materials"]["background"]),design["wavelength_nm"]).real-.001
    if not np.isfinite(target) or not 0<float(target)<3.5:raise ValueError("target neff invalid")
    s=_assemble(design,mesh,sbc_target_neff=target if sbc_target is None else sbc_target)
    K0,K1,K2=(s["K"+str(k)] for k in range(3))
    candidates,metadata=spectrum(s,float(target),n_modes,options)
    custom_probes=probe_points is not None
    if probe_points is None:
        probe_points,probe_weights=disk_probes(.95*design["domain_radius_um"],design["parameters"].get("core_radius_um"))
    else:
        probe_points=np.asarray(probe_points,dtype=float);probe_weights=np.asarray(probe_weights,dtype=float)
        if probe_points.ndim!=2 or probe_points.shape[1]!=2 or len(probe_points)>10000 or not np.isfinite(probe_points).all():raise ValueError("finite [N,2] probes required, N<=10000")
        if np.max(np.linalg.norm(probe_points,axis=1))>=design["domain_radius_um"]:raise ValueError("tracking probes must lie inside physical domain")
    if np.shape(probe_weights)!=(len(probe_points),) or not np.isfinite(probe_weights).all() or np.any(probe_weights<=0):raise ValueError("matching positive probe weights required")
    found=[];diagnostic_fields=[]
    rows=s["gauss_test_rows"]
    for beta,e_reduced,linear_residual in candidates:
        neff=beta/s["k0"]
        if not (0<neff.real<3.5) or abs(neff.imag)>.5:continue
        q=K0@e_reduced+beta*K1@e_reduced+beta*beta*K2@e_reduced
        denom=np.linalg.norm(K0@e_reduced)+abs(beta)*np.linalg.norm(K1@e_reduced)+abs(beta)**2*np.linalg.norm(K2@e_reduced)
        residual=float(np.linalg.norm(q)/max(denom,1e-30))
        if not math.isfinite(residual) or max(residual,linear_residual)>options.get("residual_tolerance",1e-7):continue
        full=np.zeros(s["full_size"],dtype=complex);full[s["keep"]]=e_reduced
        e,z=full[:s["ne"]],full[s["ne"]:]
        ge=s["Ge"]@e;gz=1j*beta*(s["Gz"]@z);g=(ge+gz)[rows]
        ef=np.asarray(s["be"].interpolate(e));zf=np.asarray(s["bz"].interpolate(z))
        dnorm=float(np.sqrt(np.sum(np.abs(s["eps"][:,None])**2*(np.abs(ef[0])**2+np.abs(ef[1])**2+np.abs(zf)**2)*s["be"].dx)))
        dual=float(np.sqrt(max(0.,np.vdot(g,s["gauss_mass_factor"].solve(g)).real)))
        div_res=dual/max(s["k0"]*dnorm,1e-30)
        cancellation=float(np.linalg.norm(g)/(np.linalg.norm(ge[rows])+np.linalg.norm(gz[rows])+1e-30))
        try:
            samples,info,fields,quadrature=_field_samples(s,full,beta,design,max_sample_points,export_full_quadrature)
            probe=probe_e(s,full,probe_points)
        except RuntimeError:continue
        parities={}
        norm=float(np.sum(np.abs(probe)**2*probe_weights[:,None]))
        for axis,label in ((0,"x"),(1,"y")):
            mirrored=np.array(probe_points,copy=True);mirrored[:,axis]*=-1
            reflected=probe_e(s,full,mirrored);reflected[:,axis]*=-1
            corr=np.sum(np.conj(probe)*reflected*probe_weights[:,None])/norm
            parities[label]={"correlation_real":float(corr.real),"correlation_imag":float(corr.imag),
                             "parity":"even" if corr.real>.98 else "odd" if corr.real<-.98 else "mixed_or_asymmetric"}
        info["mirror_symmetry_diagnostic"]=parities
        info["symmetry_class"]="mirror x="+parities["x"]["parity"]+", y="+parities["y"]["parity"]+"; not an automatic HE/TE/TM label"
        status="candidate_unverified"
        if neff.imag < -options.get("imaginary_noise_floor",1e-10) or info["PML_intensity_fraction"]>.4 or div_res>1e-5:
            status="spurious_or_unstable_candidate"
        signed_loss=confinement_loss(float(neff.imag),design["wavelength_nm"])
        record={"neff_real":float(neff.real),"neff_imag":float(neff.imag),
                "alpha_total_dB_m":signed_loss["alpha_total_dB_m"],"alpha_total_dB_km":signed_loss["alpha_total_dB_km"],
                "alpha_basis":"signed total Im(neff): radiative + material; independent input-matched study required",
                "residual_relative":residual,"linearized_residual_relative":linear_residual,
                "weak_gauss_relative":div_res,"weak_gauss_cancellation_relative":cancellation,
                "gauss_basis":"interior H1 test functions; scalar-mass dual norm / (k0 * L2 norm of untransformed epsilon E)",
                "classification":info,"field_samples":samples,"status":status}
        if quadrature is not None:record["field_quadrature"]=quadrature
        if custom_probes:
            record["tracking_fields"]={"points_um":np.asarray(probe_points).tolist(),"weights_um2":np.asarray(probe_weights).tolist(),
                                       "E":[[[float(v.real),float(v.imag)] for v in row] for row in probe],
                                       "basis":"actual mixed finite-element interpolation on fixed physical-domain probes"}
        found.append(record);diagnostic_fields.append(probe)
    if not found:raise RuntimeError("no finite residual-qualified forward modes; inspect target, mesh/PML and resource limits")
    order=sorted(range(len(found)),key=lambda j:(found[j]["status"]!="candidate_unverified",abs(found[j]["neff_real"]-float(target)),found[j]["residual_relative"]))
    columns=weighted_columns(np.stack([diagnostic_fields[j] for j in order],axis=-1),probe_weights)
    keep=independent_indices(columns,options.get("independence_tolerance",1e-7))
    unique=[found[order[j]] for j in keep[:n_modes]]
    group=0
    for j,mode in enumerate(unique):
        if j and abs(complex(mode["neff_real"],mode["neff_imag"])-complex(unique[j-1]["neff_real"],unique[j-1]["neff_imag"]))>options.get("degeneracy_tolerance",1e-5):group+=1
        mode["degenerate_cluster"]=group
    warnings=["SBC is fixed-target first-order Robin, not cylindrical DtN" if design["boundary"]["type"]=="sbc" else "PML loss needs independent mesh, window, thickness, strength and order validation",
              "field independence and weak Gauss diagnostics do not prove all spurious modes absent",
              "named fiber geometry and material assumptions remain those of the design"]
    if design.get("_bend") is not None:warnings.append("2D conformal bend remains an approximation, not COMSOL axisymmetric bend")
    return {"schema":"LSCNS_fiber_research_vector_v1","schema_version":"1.0.0","engine_version":ENGINE_VERSION,
            "method":f"2D HCurl(Nedelec{s['element_order']})+H1(P{s['element_order']}) mixed quadratic beta Maxwell FEM; {metadata['backend']} sparse shift-invert",
            "convention":"exp(i beta z - i omega t); Im(neff)>0 is passive loss",
            "solver_diagnostics":dict(metadata,returned_modes=len(unique),requested_modes=n_modes,
                                      duplicate_vectors_removed=len(found)-len(keep),all_returned_fields_finite=True),
            "validation":{"level":"DIAGNOSTICS_ONLY_INDEPENDENT_STUDY_REQUIRED","mesh_converged":False,"pml_converged":False,"loss_dB_km_certified":False,"warnings":warnings},
            "input":design,"mesh_quality":mesh["quality"],"modes":unique,
            "photoelastic_coupling":{"enabled":s["pm_anisotropic_optics"],"status":"EXPERIMENTAL_ANISOTROPIC_MAXWELL_NOT_BENCHMARKED" if s["pm_anisotropic_optics"] else "disabled",
                                    "source_sha256":mesh.get("pm_source_sha256"),"basis":"physical xy-block and zz epsilon offsets; general xz/yz tensor unsupported"},
            "materials_provenance":{key:provenance(mat) for key,mat in design["materials"].items()}}


def solve_modes(design,mesh=None,n_modes=2,target_neff=None,max_sample_points=800,
                export_full_quadrature=False,probe_points=None,probe_weights=None):
    iterations=int(design.get("solver",{}).get("sbc_iterations",1)) if design.get("boundary",{}).get("type")=="sbc" else 1
    if iterations==1:return _solve_once(design,mesh,n_modes,target_neff,max_sample_points,export_full_quadrature,probe_points,probe_weights)
    validate_design(design)
    if mesh is None:mesh=build_mesh(design)
    search_target=target_neff if target_neff is not None else design.get("target_neff")
    if search_target is None:search_target=index(design["materials"].get("core",design["materials"]["background"]),design["wavelength_nm"]).real-.001
    boundary_target=complex(search_target);history=[]
    for i in range(iterations):
        result=_solve_once(design,mesh,n_modes,search_target,max_sample_points,export_full_quadrature,probe_points,probe_weights,boundary_target)
        chosen=result["modes"][:min(2,n_modes)]
        updated=sum(complex(m["neff_real"],m["neff_imag"]) for m in chosen)/len(chosen)
        delta=abs(updated-boundary_target)
        history.append({"iteration":i+1,"boundary_target_real":boundary_target.real,"boundary_target_imag":boundary_target.imag,
                        "solution_real":updated.real,"solution_imag":updated.imag,"change":float(delta)})
        if delta<1e-8:break
        boundary_target=updated;search_target=updated.real
    result["solver_diagnostics"]["sbc_iteration"]={"history":history,"self_consistent":bool(delta<1e-8),
                                                 "basis":"iterated-target first-order Robin; NOT an exact cylindrical DtN boundary"}
    return result
