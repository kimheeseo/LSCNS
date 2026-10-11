"""M6 inverse design of a single-layer broadband antireflection film.
Forward model: independently verified M5 1D Maxwell S-matrix (not 3D FEM).
Objective: mean TE/TM reflectance over 1.30–1.65 um at 0/20deg, including
+/- 3% fabrication thickness error. Deterministic multi-start bounded search.
"""
from pathlib import Path
import json,sys
import numpy as np
from scipy.optimize import minimize
sys.path.insert(0,str(Path(__file__).resolve().parents[2]/"waveoptics_m5"/"src"))
from m5_spectral import spectrum,silica_index

W=np.linspace(1.30,1.65,81)
CASES=[("TE",0.),("TM",0.),("TE",20.),("TM",20.)]
def reflectance(n,d,wavelengths=W,errors=(0.,)):
    all_r=[]
    for scale in errors:
        for pol,angle in CASES:
            r=spectrum(wavelengths,layers=[(float(n),float(d)*scale)],polarization=pol,angle_deg=angle)["R"]
            all_r.append(np.mean(r))
    return float(np.mean(all_r))
def optimize():
    n_q=np.sqrt(float(silica_index(1.55)));d_q=1.55/(4*n_q)
    bounds=[(1.05,1.42),(.15,.44)];errors=(.97,1.,1.03)
    def objective(x):return reflectance(x[0],x[1],errors=errors)
    candidates=[]
    for start in [[n_q,d_q],[1.15,.33],[1.3,.28],[1.20,.38]]:
        fit=minimize(objective,start,method="L-BFGS-B",bounds=bounds,options={"maxiter":300,"ftol":1e-13})
        candidates.append(fit)
    sol=min(candidates,key=lambda f:f.fun)
    if not sol.success and np.linalg.norm(sol.jac)>1e-4:raise RuntimeError("Optimization did not converge")
    n,d=map(float,sol.x)
    bare=float(np.mean([np.mean(spectrum(W,polarization=p,angle_deg=a)["R"]) for p,a in CASES]))
    q=reflectance(n_q,d_q,errors=errors)
    return dict(best_index=n,best_thickness_um=d,objective_R=float(sol.fun),bare_R=bare,quarterwave_R=q,
                improvement_vs_bare_pct=float(100*(1-sol.fun/bare)),improvement_vs_quarterwave_pct=float(100*(1-sol.fun/q)),
                bound_index=list(bounds[0]),bound_thickness_um=list(bounds[1]),nominal_n=n_q,nominal_thickness_um=d_q,
                wavelength_range_um=[float(W[0]),float(W[-1])],num_wavelengths=len(W),cases=CASES,thickness_error_factors=list(errors),
                success=bool(sol.success),method="multi-start bounded L-BFGS-B")
def plot(sol,outdir):
    import matplotlib;matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    out=Path(outdir);out.mkdir(parents=True,exist_ok=True)
    n,d=sol["best_index"],sol["best_thickness_um"]
    qn,qd=sol["nominal_n"],sol["nominal_thickness_um"]
    r0=spectrum(W)["R"];rq=spectrum(W,[(qn,qd)])["R"];ro=spectrum(W,[(n,d)])["R"]
    errlo=spectrum(W,[(n,d*.97)])["R"];errhi=spectrum(W,[(n,d*1.03)])["R"]
    fig,axes=plt.subplots(1,2,figsize=(11,4),constrained_layout=True)
    axes[0].plot(W,100*r0,label="bare");axes[0].plot(W,100*rq,label="quarter-wave");axes[0].plot(W,100*ro,label="optimized")
    axes[0].fill_between(W,100*np.minimum(errlo,errhi),100*np.maximum(errlo,errhi),alpha=.2,label="thickness ±3%")
    axes[0].set(xlabel="Wavelength (um)",ylabel="Normal-incidence reflectance (%)",title="M6 coating inverse design");axes[0].legend(fontsize=8)
    N=np.linspace(1.08,1.40,43);D=np.linspace(.20,.42,43)
    Z=np.array([[reflectance(nn,dd,wavelengths=W[::4],errors=(1.,)) for nn in N] for dd in D])
    mesh=axes[1].contourf(N,D,Z*100,levels=24,cmap="viridis")
    axes[1].plot([n],[d],"ro",label="robust optimum");axes[1].set(xlabel="Film index",ylabel="Thickness (um)",title="Coating design landscape")
    fig.colorbar(mesh,ax=axes[1],label="Avg reflectance (%)");axes[1].legend(fontsize=8)
    fig.savefig(out/"m6_optimization.png",dpi=155);plt.close(fig)
    np.savetxt(out/"m6_response.csv",np.column_stack([W,r0,rq,ro,errlo,errhi]),delimiter=",",header="wavelength_um,bare_R,quarterwave_R,optimized_R,opt_minus3percent_R,opt_plus3percent_R",comments="")
def run(outdir):
    out=Path(outdir);out.mkdir(parents=True,exist_ok=True);sol=optimize()
    plot(sol,out);(out/"m6_design.json").write_text(json.dumps(sol,indent=2));return sol
if __name__=="__main__":
    import argparse
    a=argparse.ArgumentParser();a.add_argument("--out",default="results")
    args=a.parse_args();print(json.dumps(run(args.out),indent=2))
