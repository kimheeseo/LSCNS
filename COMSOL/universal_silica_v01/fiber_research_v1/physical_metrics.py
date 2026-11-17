"""LSCNS STEP 4 — physical formulas (engine v0.8.0; lengths and units explicit).

Conventions: E,H ~ exp(i*beta*z-i*omega*t); lambda in metres unless *_nm.
k0 [rad/m] = 2*pi/lambda_m; c=299792458 [m/s].
alpha_total [dB/m] = 20/ln(10) * k0 * Im(neff).  No abs(Im).
alpha_total [dB/km] = 1000 * alpha_total[dB/m].
I_t=|Ex|^2+|Ey|^2; dA in um^2 on transverse plane.
Aeff [um^2] = (integral I_t dA)^2 / integral I_t^2 dA.
MFD_2nd [um] = 2*sqrt(2 * integral |r-rbar|^2 I_t dA/integral I_t dA).
MFD_PetermannII(radial) [um] = 2*sqrt(2 * int |E(r)|^2 r dr/int |dE/dr|^2 r dr).
For non-axisymmetric vector fields, the gradient analogue is tagged PII-LIKE, not
a standardized Petermann-II MFD (IEC 60793-1-45).
D = -(lambda_m/c) * d^2 Re(neff)/d lambda_m^2 [s/m^2].
D[ps/(nm km)] = D[s/m^2]*1e6; S[ps/(nm^2 km)] = dD/dlambda_nm.
gamma [1/(W m)] = 2*pi*n2_eff[m^2/W]/(lambda_m*Aeff[m^2]).
HCF n2_eff proxy uses intensity^2 weighted fraction of supplied n2 per material;
this scalar weak-guidance approximation is NOT a rigorous tensor/gas nonlinear solver.
B = abs(nx-ny), beat length Lb [m] = lambda_m/B (B=0 => null).
HOMER = min(alpha_HOM[dB/m])/max(alpha_FM[dB/m]) for explicitly identified modes.
ARROW resonance lambda_m = 2*t*sqrt(n_glass^2-n_air^2)/m [same length unit].
Mode tracking uses weighted normalized complex transverse overlap and a 1:1
global Hungarian assignment; equal meshes and verified field samples REQUIRED.

References for methods:
IEC 60793-1-45:2024 (Petermann-II); Malitson 1965 (pure silica n);
Belardi et al / Poletti style antiresonant analytical approximation;
https://doi.org/10.1364/OE.27.003824 (HOMER).
"""
from __future__ import annotations
from dataclasses import dataclass
from typing import Sequence
import math
import numpy as np
from scipy.optimize import linear_sum_assignment

ENGINE_VERSION="0.8.0"
C_M_S=299792458.0
DECIBEL_PER_NEPER_POWER=20.0/math.log(10.0)
FORMULA_SCOPE="analytical post-processing; not a FEM/PML validation"


def _finite_float(x,name,positive=False,allow_zero=False):
    if isinstance(x,bool):
        raise ValueError(f"{name}: bool is not numerical")
    try: f=float(x)
    except (ValueError,TypeError): raise ValueError(f"{name}: finite number required") from None
    if not math.isfinite(f) or (positive and (f<0 if allow_zero else f<=0)):
        raise ValueError(f"{name}: invalid finite/positive numeric value")
    return f


def confinement_loss(neff_imag:float, wavelength_nm:float):
    """Return signed total loss [dB/m], [dB/km]. Loss IS NOT necessarily radiation-only."""
    imag=_finite_float(neff_imag,"Im(neff)")
    lam=_finite_float(wavelength_nm,"wavelength_nm",positive=True)*1e-9
    loss=DECIBEL_PER_NEPER_POWER*(2*math.pi/lam)*imag
    return {"alpha_total_dB_m":loss,"alpha_total_dB_km":1000*loss,
            "is_passive_sign":imag>=0.0,
            "scope":"TOTAL loss if Im(neff) includes material absorption; NOT certified confinement loss"}


def _axis(name,arr):
    x=np.asarray(arr,dtype=float)
    if x.ndim!=1 or len(x)<5 or not np.isfinite(x).all() or not np.all(np.diff(x)>0):
        raise ValueError(f"{name}: 1D strictly increasing finite array length >=5 required")
    return x


def _grid(x_um,y_um,ex,ey,mask=None):
    x=_axis("x_um",x_um); y=_axis("y_um",y_um)
    dx=np.diff(x);dy=np.diff(y)
    if not (np.allclose(dx,dx[0],rtol=1e-10,atol=1e-12) and
            np.allclose(dy,dy[0],rtol=1e-10,atol=1e-12)):
        raise ValueError("field integrals require a uniform Cartesian grid")
    ex=np.asarray(ex,dtype=complex);ey=np.asarray(ey,dtype=complex)
    if ex.shape!=(len(y),len(x)) or ey.shape!=ex.shape:
        raise ValueError("Ex/Ey must be [ny,nx] uniform-grid arrays")
    if not (np.isfinite(ex).all() and np.isfinite(ey).all()):
        raise ValueError("non-finite electric field")
    if mask is None: mask=np.ones(ex.shape,dtype=bool)
    else: mask=np.asarray(mask,dtype=bool)
    if mask.shape!=ex.shape or np.count_nonzero(mask)<5:
        raise ValueError("physical region mask invalid (exclude PML)")
    return x,y,ex,ey,mask,float(dx[0]*dy[0])


def field_area_mfd(x_um,y_um,ex,ey,physical_mask=None):
    """Physical Et-only Aeff and second-moment MFD; gradient MFD is PII-like only."""
    x,y,ex,ey,valid,dA=_grid(x_um,y_um,ex,ey,physical_mask)
    I=np.abs(ex)**2+np.abs(ey)**2
    w=I*valid
    P=float(np.sum(w)*dA)
    if P<=0: raise ValueError("zero transverse electric-field intensity")
    Q=float(np.sum(I*I*valid)*dA)
    XX,YY=np.meshgrid(x,y)
    cx=float(np.sum(XX*w)*dA/P)
    cy=float(np.sum(YY*w)*dA/P)
    rr=(XX-cx)**2+(YY-cy)**2
    second=2*math.sqrt(2*float(np.sum(rr*w)*dA/P))
    amp=np.sqrt(I)
    ddY,ddX=np.gradient(amp,y,x,edge_order=2)
    # Not a rigorously window-insensitive PII measurement: boundary truncation must converge.
    gradient=float(np.sum((ddX**2+ddY**2)*valid)*dA)
    pii_like=2*math.sqrt(2*P/gradient) if gradient>0 else None
    return {"Aeff_um2":P*P/Q,"MFD_2nd_um":second,
            "MFD_PII_like_um":pii_like,"MFD_PII_status":"2D transverse-amplitude-gradient proxy, NOT standardized Petermann-II for arbitrary vectors",
            "centroid_um":[cx,cy],"physical_region_integral_um2":float(np.count_nonzero(valid)*dA),
            "field_integral_basis":"Et only, PML excluded by physical_mask"}


def petermann_ii_radial(r_um,complex_radial_amplitude):
    """True near-field derivative Petermann-II for axisymmetric scalar mode amplitude."""
    r=_axis("r_um",r_um)
    e=np.asarray(complex_radial_amplitude,dtype=complex)
    if e.shape!=r.shape or not np.isfinite(e).all() or r[0]<0:
        raise ValueError("radial amplitude/axis invalid")
    de=np.gradient(e,r,edge_order=2)
    p=float(np.trapezoid(np.abs(e)**2*r,r))
    g=float(np.trapezoid(np.abs(de)**2*r,r))
    if p<=0 or g<=0: raise ValueError("invalid radial field/derivative")
    return {"MFD_PetermannII_um":2*math.sqrt(2*p/g),
            "method":"axisymmetric near-field derivative integral, verify outer-radius convergence"}


def birefringence(neff_x,neff_y,wavelength_nm):
    a=_finite_float(complex(neff_x).real,"neff_x")
    b=_finite_float(complex(neff_y).real,"neff_y")
    lam=_finite_float(wavelength_nm,"wavelength_nm",positive=True)*1e-9
    B=abs(a-b)
    return {"B":B,"beat_length_m":(lam/B if B else None),
            "requires":"x/y polarization identity and mode overlap checked independently"}


def mode_neff_spacing(modes):
    arr=[_finite_float(complex(v).real,"Re(neff)") for v in modes]
    if len(arr)<2: raise ValueError("2 or more modes required")
    return [{"i":i,"j":j,"delta_Re_neff":abs(arr[i]-arr[j])}
            for i in range(len(arr)) for j in range(i+1,len(arr))]


def nonlinear_gamma(n2_m2_W,wavelength_nm,Aeff_um2):
    n2=_finite_float(n2_m2_W,"n2[m^2/W]",positive=True,allow_zero=True)
    lam=_finite_float(wavelength_nm,"wavelength_nm",positive=True)*1e-9
    area=_finite_float(Aeff_um2,"Aeff[um^2]",positive=True)*1e-12
    return {"gamma_W_inv_m_inv":(2*math.pi*n2/(lam*area)),
            "method":"scalar weak-guidance; n2 explicitly user-supplied; Et-only Aeff"}


def hcf_effective_n2(x_um,y_um,ex,ey,glass_mask,n2_air_m2_W=None,n2_glass_m2_W=None,physical_mask=None):
    """Quartic |Et|^4 weighted n2 proxy. Does NOT use linear power fractions as n2 weights."""
    if n2_air_m2_W is None or n2_glass_m2_W is None:
        return {"status":"TODO_MATERIAL_N2","needed":"measured/cited n2_air and n2_glass at wavelength, pressure and gas composition"}
    air=_finite_float(n2_air_m2_W,"n2_air",positive=True,allow_zero=True)
    glass=_finite_float(n2_glass_m2_W,"n2_glass",positive=True,allow_zero=True)
    x,y,ex,ey,valid,dA=_grid(x_um,y_um,ex,ey,physical_mask)
    gm=np.asarray(glass_mask,dtype=bool)
    if gm.shape!=valid.shape or not np.any(gm&valid) or not np.any((~gm)&valid):
        raise ValueError("glass and air masks both required in physical domain")
    I=np.abs(ex)**2+np.abs(ey)**2
    w=I*I*valid
    wtotal=float(np.sum(w)*dA)
    if wtotal<=0:raise ValueError("nonzero field required")
    wglass=float(np.sum(w*gm)*dA/wtotal)
    return {"n2_eff_m2_W":wglass*glass+(1-wglass)*air,
            "quartic_glass_weight":wglass,"quartic_air_weight":1-wglass,
            "method":"Et^4 scalar overlap approximation; nonlinear material coefficients supplied by user",
            "not_equivalent_to":"glass optical power fraction (Poynting) or tensor Kerr perturbation"}


def glass_power_fraction(x_um,y_um,ex,ey,hx,hy,glass_mask,physical_mask=None):
    """Real axial power density Sz=0.5 Re(Ex Hy* - Ey Hx*)."""
    x,y,ex,ey,valid,dA=_grid(x_um,y_um,ex,ey,physical_mask)
    hx=np.asarray(hx,dtype=complex);hy=np.asarray(hy,dtype=complex)
    gm=np.asarray(glass_mask,dtype=bool)
    if hx.shape!=ex.shape or hy.shape!=ex.shape or gm.shape!=valid.shape:
        raise ValueError("Hx/Hy and glass mask must match Et grid")
    if not np.isfinite(hx).all() or not np.isfinite(hy).all():
        raise ValueError("nonfinite H fields")
    sz=0.5*np.real(ex*np.conj(hy)-ey*np.conj(hx))
    p=float(np.sum(sz*valid)*dA)
    if not math.isfinite(p) or p<=0:
        raise ValueError("net forward power <=0; cannot calculate physical power fractions")
    pg=float(np.sum(sz*valid*gm)*dA)
    return {"P_glass_fraction":pg/p,"P_air_fraction":1-pg/p,
            "net_power_proxy":p,"unit_of_power_proxy":"E-field * H-field * um^2; only fractions dimensionless",
            "warning":"PML excluded; local backward/negative Sz and leaky-mode normalization need review"}


def homer(fundamental_losses_dB_m, higher_order_losses_dB_m, mode_ids_verified=False):
    """Lowest HOM loss / highest FM loss. Returns dimensionless, NOT dB."""
    if not mode_ids_verified:
        return {"status":"TODO_MODE_CLASSIFICATION","needed":"verified FM/HOM labels from physical field profiles and mode overlap"}
    fm=[_finite_float(x,"FM loss",positive=True) for x in fundamental_losses_dB_m]
    hom=[_finite_float(x,"HOM loss",positive=True) for x in higher_order_losses_dB_m]
    if not fm or not hom:raise ValueError("both FM and HOM mode loss sets required")
    return {"HOMER_ratio":min(hom)/max(fm),"numerator_min_HOM_dB_m":min(hom),
            "denominator_max_FM_dB_m":max(fm),
            "validation_status":"requires individually converged FEM/PML losses and FM/HOM labels"}


def arrow_resonance_lines(wall_thickness_um,n_glass,n_air,orders=range(1,5)):
    """Constant-index ARROW resonances; dispersion must be solved self-consistently later."""
    t=_finite_float(wall_thickness_um,"wall_thickness_um",positive=True)
    ng=_finite_float(n_glass,"n_glass",positive=True)
    na=_finite_float(n_air,"n_air",positive=True)
    if ng<=na:raise ValueError("glass must have larger refractive index")
    orders=list(orders)
    if not orders or any(isinstance(m,bool) or not isinstance(m,int) or not 1<=m<=50 for m in orders):
        raise ValueError("orders must be positive integers 1..50")
    return {"resonance_lines":[{"m":m,"lambda_nm":2000*t*math.sqrt(ng*ng-na*na)/m} for m in orders],
            "method":"constant-index ARROW wall approximation; resonance NOT an FEM eigensolution",
            "needed_for_dispersive_glass":"n_glass(lambda), n_air(lambda) and self-consistent root at each m"}


def overlay_arrow_with_fem(arrow,tracked_fem=None):
    """Overlay data only: never synthesize optical FEM curves."""
    data={"arrow_lines":arrow["resonance_lines"],"fem_curves":[]}
    if tracked_fem is not None:
        for line in tracked_fem:
            if not line.get("overlap_verified",False):
                raise ValueError("FEM overlay requires overlap-verified mode tracking")
            lam=_axis("wavelength_nm",line["wavelength_nm"])
            n=np.asarray(line["neff_real"],dtype=float)
            if n.shape!=lam.shape or not np.isfinite(n).all():
                raise ValueError("invalid FEM mode curve")
            data["fem_curves"].append({"mode_id":line.get("mode_id"),"wavelength_nm":lam.tolist(),"neff_real":n.tolist()})
    return data


def _normalize_mode(mode):
    pts=np.asarray(mode["points_um"],dtype=float)
    w=np.asarray(mode["weights_um2"],dtype=float)
    ex=np.asarray(mode["Ex"],dtype=complex);ey=np.asarray(mode["Ey"],dtype=complex)
    if pts.ndim!=2 or pts.shape[1]!=2 or len(pts)<3 or w.shape!=(len(pts),) or ex.shape!=w.shape or ey.shape!=w.shape:
        raise ValueError("identical 2D quadrature points and Ex/Ey required")
    if not all(np.isfinite(a).all() for a in (pts,w,ex,ey)) or np.any(w<=0):
        raise ValueError("finite fields/positive quadrature weights required")
    q=np.sum(w*(np.abs(ex)**2+np.abs(ey)**2))
    if q<=0:raise ValueError("zero Et norm")
    return pts,w,ex,ey,float(q)


def match_modes_by_overlap(previous_modes,current_modes,minimum_overlap=0.75):
    """Strict same-quadrature mesh matching; reject sample-only data and remeshed coordinates."""
    threshold=_finite_float(minimum_overlap,"minimum_overlap",positive=True)
    if not 0<threshold<=1:raise ValueError("overlap threshold (0,1]")
    if not previous_modes or not current_modes:raise ValueError("both mode sets needed")
    prev=[_normalize_mode(m) for m in previous_modes]
    curr=[_normalize_mode(m) for m in current_modes]
    ref_pts,ref_weights=prev[0][0],prev[0][1]
    for group in (prev,curr):
        for pts,w,*_ in group:
            if pts.shape!=ref_pts.shape or not np.allclose(pts,ref_pts,rtol=0,atol=1e-10):
                raise ValueError("mesh points differ; require verified interpolation to common quadrature")
            if not np.allclose(w,ref_weights,rtol=1e-10,atol=1e-12):
                raise ValueError("quadrature weights differ; cannot compare overlap")
    score=np.zeros((len(prev),len(curr)),float)
    phases=np.zeros_like(score)
    for i,(_,w,ex1,ey1,n1) in enumerate(prev):
        for j,(_,_,ex2,ey2,n2) in enumerate(curr):
            inner=np.sum(w*(np.conj(ex1)*ex2+np.conj(ey1)*ey2))
            score[i,j]=min(1,float(abs(inner)**2/(n1*n2)))
            phases[i,j]=-float(np.angle(inner))
    rr,cc=linear_sum_assignment(-score)
    assignments=[{"previous_index":int(i),"current_index":int(j),
                  "overlap_power":float(score[i,j]),"phase_alignment_rad":float(phases[i,j]),
                  "tracked":bool(score[i,j]>=threshold)}
                 for i,j in zip(rr,cc)]
    return {"assignments":assignments,"all_tracked":bool(len(rr)==len(prev)==len(curr) and all(a["tracked"] for a in assignments)),
            "minimum_overlap":threshold,
            "method":"Hungarian 1-to-1 assignment of normalized complex vector Et overlap with matching area weights",
            "limitations":"near degeneracy may require subspace tracking; mesh change must first be interpolated and verified"}


def dispersion_from_tracked_sweep(wavelength_nm,neff_real,adjacent_overlap_power,minimum_overlap=.75):
    """Third-degree local 5-point fit. D/S and lambda0 require pre-verified mode continuity."""
    lam=_axis("wavelength_nm",wavelength_nm)
    nn=np.asarray(neff_real,dtype=float)
    overlaps=np.asarray(adjacent_overlap_power,dtype=float)
    if len(lam)<5 or nn.shape!=lam.shape or overlaps.shape!=(len(lam)-1,) or not np.isfinite(nn).all() or not np.isfinite(overlaps).all():
        raise ValueError("at least five finite samples and N-1 overlap scores required")
    if any(o<minimum_overlap or o>1 for o in overlaps):
        raise ValueError("ambiguous or switched mode: dispersion calculation BLOCKED")
    second=[]
    for k in range(len(lam)):
        start=min(max(0,k-2),len(lam)-5)
        sl=slice(start,start+5)
        xx=lam[sl]-lam[k]
        coefs=np.polyfit(xx,nn[sl],3)
        second.append(2*coefs[-3])  # per nm^2
    second=np.asarray(second)
    Dsi=-(lam*1e-9/C_M_S)*(second*1e18) # s/m^2
    dps=Dsi*1e6
    slope=np.gradient(dps,lam,edge_order=2) # ps/(nm^2 km)
    zero=[]
    for i in range(len(dps)-1):
        if dps[i]==0: zero.append(float(lam[i]))
        elif dps[i]*dps[i+1]<0:
            zero.append(float(lam[i]-dps[i]*(lam[i+1]-lam[i])/(dps[i+1]-dps[i])))
    if dps[-1]==0: zero.append(float(lam[-1]))
    return {"wavelength_nm":lam.tolist(),"D_ps_nm_km":dps.tolist(),
            "S_ps_nm2_km":slope.tolist(),
            "zero_dispersion_lambda_nm":zero,
            "derivative_method":"local cubic fit (five points), slopes from D gradient; confirm step-size convergence",
            "mode_tracking":"verified by adjacent complex-field overlap at >= minimum_overlap; regridding requires independent check"}


def surface_anticrossing_candidates(tracked_pair):
    """Find avoided-crossing candidates from local gap minimum and exchanged glass fraction.

    Requires input arrays for already-overlap-tracked core and glass mode branches.
    No numerical threshold for 'surface mode' is invented; the caller must label
    the glass-localized branch from validated field integrals.
    """
    if not tracked_pair.get("overlap_verified",False) or not tracked_pair.get("glass_mode_identified",False):
        return {"status":"TODO_MODE_IDENTITY",
                "needed":"two overlap-tracked FEM branches, physically glass-localized mode fields and converged neff"}
    lam=_axis("wavelength_nm",tracked_pair["wavelength_nm"])
    a=np.asarray(tracked_pair["neff_core"],dtype=float)
    b=np.asarray(tracked_pair["neff_glass"],dtype=float)
    fa=np.asarray(tracked_pair["glass_fraction_core"],dtype=float)
    fb=np.asarray(tracked_pair["glass_fraction_glass"],dtype=float)
    if any(v.shape!=lam.shape or not np.isfinite(v).all() for v in (a,b,fa,fb)):
        raise ValueError("valid pair curves required")
    if any(np.any(v<0) or np.any(v>1) for v in (fa,fb)):
        raise ValueError("glass power fractions [0,1]")
    delta=np.abs(a-b)
    candidates=[]
    for i in range(1,len(lam)-1):
        if delta[i]<delta[i-1] and delta[i]<delta[i+1]:
            exchange=(fa[i-1]-fb[i-1])*(fa[i+1]-fb[i+1])<0
            if exchange:
                candidates.append({"lambda_nm":float(lam[i]),"min_gap_Re_neff":float(delta[i]),
                                   "status":"anticrossing_candidate_only",
                                   "note":"finite step resolution; confirm field patterns, refinement and polarization"})
    return {"status":"candidate_screening_only","candidates":candidates}
