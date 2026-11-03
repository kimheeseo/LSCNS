"""M5 dispersive spectral S-parameter toolkit, normal-incidence TE/TM, exp(-i wt).
Independent 1D Maxwell transfer-matrix oracle; optional M3 FEM cross-check.
Units: wavelengths and thicknesses micrometres; not a vector-3D FEM implementation.
"""
from __future__ import annotations
from pathlib import Path
import numpy as np
from scipy.constants import c
# Fused-silica Malitson Sellmeier: λ in µm, valid primarily near visible–IR.
B=np.array([.6961663,.4079426,.8974794])
C=np.array([.0684043**2,.1162414**2,9.896161**2])
def silica_index(wavelength_um):
    w=np.asarray(wavelength_um,dtype=float)
    if np.any(~np.isfinite(w)) or np.any((w<.21)|(w>3.7)):raise ValueError("Sellmeier restricted to 0.21..3.7 um")
    return np.sqrt(1.+np.sum(B*w[...,None]**2/(w[...,None]**2-C),axis=-1))
def silica_group_index(wavelength_um):
    w=np.asarray(wavelength_um,dtype=float);n=silica_index(w)
    dndw=-np.sum(B*w[...,None]*C/(w[...,None]**2-C)**2,axis=-1)/n
    return n-w*dndw
def _admittance(n,ct,pol):
    if pol=="TE":return n*ct
    if pol=="TM":return n/ct
    raise ValueError("polarization must be TE or TM")
def spectrum(wavelength_um,layers=(),substrate="silica",incident=1.,polarization="TE",angle_deg=0.):
    w=np.atleast_1d(np.asarray(wavelength_um,dtype=float))
    if len(w)==0 or np.any(~np.isfinite(w)) or np.any(w<=0):raise ValueError("positive finite wavelengths")
    ns=silica_index(w) if substrate=="silica" else np.broadcast_to(np.asarray(substrate,dtype=float),w.shape)
    if np.any(ns<=0) or incident<=0:raise ValueError("positive indices required")
    if not 0<=angle_deg<85:raise ValueError("angle_deg must be in [0,85)")
    s0=np.sin(np.deg2rad(angle_deg));cos0=np.sqrt(1.-s0*s0)
    Ys=_admittance(ns,np.sqrt(1-(incident*s0/ns)**2+0j),polarization)
    Y0=_admittance(incident,cos0,polarization)
    a=np.ones(w.shape,complex);b=np.zeros(w.shape,complex);cc=b.copy();d=a.copy()
    for nraw,depth in layers:
        n=np.broadcast_to(np.asarray(nraw,dtype=float),w.shape)
        if np.any(n<=0) or depth<0 or not np.isfinite(depth):raise ValueError("passive positive n and nonnegative thickness")
        ct=np.sqrt(1-(incident*s0/n)**2+0j)
        Y=_admittance(n,ct,polarization);phase=2*np.pi*n*ct*depth/w
        cs=np.cos(phase);ss=np.sin(phase)
        A=cs;Bv=-1j*ss/Y;Cv=-1j*Y*ss;D=cs
        a,b,cc,d=a*A+b*Cv,a*Bv+b*D,cc*A+d*Cv,cc*Bv+d*D
    denom=Y0*a+Y0*Ys*b+cc+Ys*d
    refl=(Y0*a+Y0*Ys*b-cc-Ys*d)/denom
    trans=2*Y0/denom
    R=np.abs(refl)**2;T=np.real(Ys)/np.real(Y0)*np.abs(trans)**2
    return {"wavelength_um":w,"S11":refl,"S21":trans,"R":R,"T":T,"A":1.-R-T,"substrate_n":ns}
def group_delay_fs(wavelength_um,phase):
    lam=np.asarray(wavelength_um,float);ph=np.unwrap(np.angle(phase))
    omega=2*np.pi*c/(lam*1e-6)
    return np.gradient(ph,omega)*1e15
def run(outdir):
    import json
    import matplotlib;matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    out=Path(outdir);out.mkdir(parents=True,exist_ok=True)
    wl=np.linspace(1.3,1.65,101);bare=spectrum(wl);n0=np.sqrt(float(silica_index(1.55)));depth=1.55/(4*n0)
    ar=spectrum(wl,layers=[(n0,depth)])
    tm=spectrum(wl,layers=[(n0,depth)],polarization="TM",angle_deg=20)
    te=spectrum(wl,layers=[(n0,depth)],polarization="TE",angle_deg=20)
    gd=group_delay_fs(wl,ar["S21"])
    stats={"lambda_range_um":[float(wl[0]),float(wl[-1])],"sellmeier_n_1p55":float(silica_index(1.55)),
        "group_index_1p55":float(silica_group_index(1.55)),"quarterwave_n":n0,"quarterwave_depth_um":depth,
        "bare_reflectance_1p55":float(spectrum([1.55])["R"][0]),"AR_reflectance_1p55":float(spectrum([1.55],layers=[(n0,depth)])["R"][0]),
        "max_energy_error":float(np.max(np.abs(ar["R"]+ar["T"]-1))),"mean_group_delay_fs":float(np.mean(gd))}
    (out/"m5_summary.json").write_text(json.dumps(stats,indent=2))
    np.savetxt(out/"m5_spectrum.csv",np.column_stack([wl,bare["R"],ar["R"],te["R"],tm["R"],gd]),delimiter=",",header="wavelength_um,R_bare,R_AR_normal,R_AR_TE20,R_AR_TM20,group_delay_fs",comments="")
    fig,axes=plt.subplots(1,2,figsize=(11,4),constrained_layout=True)
    axes[0].plot(wl,bare["R"]*100,label="bare silica");axes[0].plot(wl,ar["R"]*100,label="quarter-wave coating")
    axes[0].plot(wl,te["R"]*100,"--",label="TE 20 deg");axes[0].plot(wl,tm["R"]*100,"--",label="TM 20 deg")
    axes[0].legend(fontsize=8);axes[0].set(xlabel="Wavelength (um)",ylabel="Reflected power (%)",title="M5 complex S-parameter spectrum")
    axes[1].plot(wl,silica_index(wl),label="phase n");axes[1].plot(wl,silica_group_index(wl),label="group index")
    axes[1].set(xlabel="Wavelength (um)",ylabel="Index",title="Fused silica dispersion");axes[1].legend()
    fig.savefig(out/"m5_spectrum.png",dpi=160);plt.close(fig)
    return stats
if __name__=="__main__":
    import argparse
    p=argparse.ArgumentParser();p.add_argument("--out",default="results")
    a=p.parse_args();print(run(a.out))
