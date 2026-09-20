"""Reproduce Carena et al. JLT 2012 Fig. 5 with the user's numerical GN engine
and compare a representative subset against EGN_adaptive.py.

Paper Fig. 5 convention:
- 9 channels, 32 GBd
- 100-km spans, EDFA NF=5 dB and exact span-loss compensation
- 4th-order super-Gaussian Tx filter, optimized Bopt ~= channel spacing
- target BER=1e-3; linear XI/ISI penalty obtained from back-to-back Fig. 3
- Figs. 4-7 use INCOHERENT NLI accumulation (paper Eq. 17)

The GN run below approximates the NRZ/SG optical PSD by sinc^2(NRZ)*4th-order
super-Gaussian and feeds that PSD to gn_integral_general.py. No fitted NLI
scale factor is used.

EGN_adaptive's precision path accepts rectangular spectra and coherent
multi-span accumulation only. For the requested Fig.-5 comparison we therefore
compute its FULL one-span EGN NLI coefficient, then apply the paper's
incoherent N-span scaling. This is explicitly a nearest-scope comparison, not
a claim that Carena-2012 Fig. 5 is an EGN benchmark.
"""
from __future__ import annotations
import json, math, sys
from pathlib import Path
from functools import lru_cache
import numpy as np
import pandas as pd
from scipy.integrate import quad

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
EGN_DIR = ROOT / "EGN_model"
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(EGN_DIR))

import gn_integral_general as gn
import gn_integral_general_modulation as perf
import EGN_adaptive as egn

OUT = HERE / "results"
OUT.mkdir(exist_ok=True)

RS_GBD = 32.0
BN_GHZ = 12.5  # 0.1 nm convention used by the paper
SPAN_KM = 100.0
NF_DB = 5.0
CUT = 4
NCH = 9


@lru_cache(maxsize=None)
def _sg_norm(rs_THz: float, bandwidth_THz: float) -> float:
    # PSD before optical filtering: NRZ sinc^2. Power transfer of 4th-order
    # super-Gaussian: exp[-ln2*(2f/B)^8], where B is bilateral -3 dB BW.
    B = float(bandwidth_THz)
    Rs = float(rs_THz)
    def raw(f):
        x = f / Rs
        sinc = 1.0 if x == 0 else math.sin(math.pi*x)/(math.pi*x)
        sg = math.exp(-math.log(2.0)*(2.0*f/B)**8)
        return sinc*sinc*sg
    return quad(raw, -B, B, epsabs=1e-13, epsrel=1e-10, limit=300)[0]


def paper_like_psd(spacing_GHz: float):
    Rs = RS_GBD/1000.0
    B = float(spacing_GHz)/1000.0
    norm = _sg_norm(Rs, B)
    def shape(offset_THz):
        x = np.asarray(offset_THz, dtype=float)
        sinc2 = np.sinc(x/Rs)**2
        sg = np.exp(-math.log(2.0)*(2.0*x/B)**8)
        return sinc2*sg/norm
    return shape, B


def make_paper_system(spacing_GHz: float, power_dBm: float=0.0):
    spacing = float(spacing_GHz)/1000.0
    centers = (np.arange(NCH)-(NCH-1)/2.0)*spacing
    p = gn.dbm_to_w(power_dBm)
    shape, support = paper_like_psd(spacing_GHz)
    channels = tuple(
        gn.Channel(float(fc), p, RS_GBD, pulse_shape="custom",
                   custom_psd=shape, custom_support_half_width_THz=support,
                   label=f"ch{k}")
        for k,fc in enumerate(centers)
    )
    return gn.WDMSystem(channels)


def ase_per_span_W(span) -> float:
    return perf.ase_noise_power_edfa([span], RS_GBD*1e9, 1550.0)


def lmax_from_eta(eta_W_inv2: float, ase1_W: float, required_osnr_db: float):
    # Fig.3 OSNR -> matched-filter SNR_ASE via paper Eq.(8):
    # SNR = OSNR * B_N/R_s.
    snr_req_db = float(required_osnr_db) + 10.0*math.log10(BN_GHZ/RS_GBD)
    snr_req = 10.0**(snr_req_db/10.0)
    eta = float(eta_W_inv2)
    A = float(ase1_W)
    if eta <= 0 or A <= 0:
        raise ValueError("eta and ASE must be positive")
    p_opt = (A/(2.0*eta))**(1.0/3.0)
    snr_one_span = p_opt/(A + eta*p_opt**3)
    nmax = max(0, int(math.floor(snr_one_span/snr_req + 1e-12)))
    return {
        "required_snr_db": snr_req_db,
        "popt_dBm": gn.w_to_dbm(p_opt),
        "snr_one_span_db": 10*math.log10(snr_one_span),
        "nspans": nmax,
        "Lmax_km": nmax*SPAN_KM,
    }


def svg_fig5(df: pd.DataFrame, path: Path):
    import matplotlib.pyplot as plt
    mods = ["BPSK","QPSK","8QAM","16QAM"]
    for fiber in ["PSCF","SMF","NZDSF"]:
        d = df[df.fiber==fiber]
        fig, ax = plt.subplots(figsize=(7.4,5.0))
        for mod in mods:
            m=d[d.modulation==mod].sort_values("net_SE_bit_s_Hz")
            if m.empty: continue
            ax.plot(m.net_SE_bit_s_Hz, m.paper_Lmax_km, marker="o", linestyle="--", label=f"Paper {mod}")
            ax.plot(m.net_SE_bit_s_Hz, m.gn_Lmax_km, marker="x", linestyle="-", label=f"GN code {mod}")
        ax.set_yscale("log")
        ax.set_ylim(180,22000)
        ax.set_xlim(.9,6.0)
        ax.set_xlabel("Net spectral efficiency [bit/s/Hz]")
        ax.set_ylabel("Maximum reach [km]")
        ax.set_title(f"Carena 2012 Fig. 5 reproduction — {fiber}")
        ax.grid(True, which="both", alpha=.25)
        ax.legend(ncol=2, fontsize=7)
        fig.tight_layout()
        fig.savefig(path.with_name(path.stem+f"_{fiber}"+path.suffix))
        plt.close(fig)


def svg_parity(df: pd.DataFrame, path: Path):
    import matplotlib.pyplot as plt
    fig, ax = plt.subplots(figsize=(6.2,6.0))
    lo=max(150, float(df.paper_Lmax_km.min())*.75)
    hi=float(df.paper_Lmax_km.max())*1.25
    ax.loglog([lo,hi],[lo,hi],linestyle="--",label="ideal agreement")
    ax.scatter(df.paper_Lmax_km, df.gn_Lmax_km, marker="o", label="GN")
    ax.scatter(df.paper_Lmax_km, df.egn_Lmax_km, marker="x", label="EGN adaptive*")
    ax.set_xlabel("Paper Fig.5 Lmax [km]")
    ax.set_ylabel("Code-predicted Lmax [km]")
    ax.set_title("Paper vs GN vs EGN (50/38.4-GHz subset)")
    ax.grid(True, which="both", alpha=.25)
    ax.legend()
    fig.tight_layout()
    fig.savefig(path)
    plt.close(fig)


def main():
    ref = pd.read_csv(HERE/"paper_fig5_digitized.csv")

    # ---------- GN Fig. 5 reproduction ----------
    gn_cache = {}
    gn_opt = gn.GNIntegralOptions(sobol_power=14, seed=17, accumulation="incoherent",
                                  z_quadrature_order=64)
    for (fiber, alpha, D, gamma, spacing), _ in ref.groupby(
        ["fiber","alpha_dB_km","D_ps_nm_km","gamma_W_inv_km","spacing_GHz"]):
        span = gn.Span(SPAN_KM, alpha, gamma, D_ps_nm_km=D,
                       noise_figure_db=NF_DB)
        system = make_paper_system(spacing, 0.0)
        p_nli_1 = gn.integrate_nli_over_channel(system,[span],CUT,gn_opt,receiver_points=5)
        eta = p_nli_1/(1e-3**3)
        gn_cache[(fiber,float(spacing))]=(eta,ase_per_span_W(span),p_nli_1)

    grows=[]
    for _, r in ref.iterrows():
        eta,A,pn = gn_cache[(r.fiber,float(r.spacing_GHz))]
        reach=lmax_from_eta(eta,A,r.paper_required_OSNR_dB_0p1nm)
        d=r.to_dict()
        d.update({
            "gn_eta_W_inv2":eta,
            "gn_one_span_nli_W_at_0dBm":pn,
            "ase_one_span_W":A,
            "gn_Lmax_km":reach["Lmax_km"],
            "gn_popt_dBm":reach["popt_dBm"],
            "gn_error_pct":abs(reach["Lmax_km"]-r.paper_Lmax_km)/r.paper_Lmax_km*100.0,
            "gn_signed_error_pct":(reach["Lmax_km"]-r.paper_Lmax_km)/r.paper_Lmax_km*100.0,
            "required_snr_db":reach["required_snr_db"],
        })
        grows.append(d)
    gdf=pd.DataFrame(grows)
    gdf.to_csv(OUT/"figure5_reproduction.csv",index=False)

    # ---------- full EGN representative comparison ----------
    # Use the 50-GHz Fig.-5 points (12 combinations = 3 fibers x 4 formats).
    # This keeps the full-EGN comparison computationally bounded while covering
    # every fiber and modulation family in the paper.
    subset = gdf[np.isclose(gdf.spacing_GHz,50.0)].copy()
    erows=[]
    eopt=egn.EGNFullOptions(receiver_points=5,max_receiver_points=5,panel_order=8,
                            quadrature_rtol=7e-4,verify_convergence=False,
                            strict_convergence=False)

    def egn_reach_multispan(row):
        n0=max(1,int(round(float(row.paper_Lmax_km)/SPAN_KM)))
        system=egn.WDMSystem.equispaced(NCH,float(row.spacing_GHz),RS_GBD,0.0)
        span=egn.Span(SPAN_KM,float(row.alpha_dB_km),float(row.gamma_W_inv_km),
                      D_ps_nm_km=float(row.D_ps_nm_km),noise_figure_db=NF_DB)
        A=float(row.ase_one_span_W)
        snr_req=10.0**(float(row.required_snr_db)/10.0)

        # Explicit numerical guard built into EGN_adaptive.py.
        max_supported=max(1,int(math.floor(600.0/(2.0*span.alpha_field_per_km*SPAN_KM))))
        # Evaluate a compact, non-fitted span grid. Include the paper span count
        # when it lies inside the EGN implementation's validated numerical range.
        anchors=[1,2,5,10,20,40,60,80,100,120,max_supported]
        around=[n0-4,n0-2,n0,n0+2,n0+4]
        counts=sorted(set(int(n) for n in anchors+around if 1<=int(n)<=max_supported))
        brs=egn.egn_span_sweep(system,span,counts,CUT,str(row.modulation),full_options=eopt)

        ns=[]; snrs=[]; popts=[]; etas=[]; ratios=[]; ratios_db=[]
        for n in counts:
            br=brs[n]
            eta=float(br.total_egn_W)/(1e-3**3)
            p_opt=(n*A/(2.0*eta))**(1.0/3.0)
            snr=p_opt/(n*A+eta*p_opt**3)
            ns.append(float(n)); snrs.append(float(snr)); popts.append(float(p_opt))
            etas.append(float(eta)); ratios.append(float(br.total_egn_W/br.gn_total_W))
            ratios_db.append(float(br.ratio_to_gn_db))

        ns=np.asarray(ns); snrs=np.asarray(snrs); popts=np.asarray(popts)
        etas=np.asarray(etas); ratios=np.asarray(ratios); ratios_db=np.asarray(ratios_db)
        order=np.argsort(ns); ns,snrs,popts,etas,ratios,ratios_db=[x[order] for x in (ns,snrs,popts,etas,ratios,ratios_db)]

        # Find an SNR crossing; interpolate in log(SNR) vs span count.
        above=snrs>=snr_req
        if np.all(above):
            return {"status":"right_censored_span_limit","max_supported_spans":max_supported,
                    "nspans":float(ns[-1]),"Lmax_km":math.nan,
                    "popt_dBm":egn.w_to_dbm(popts[-1]),"eta_at_Lmax_W_inv2":etas[-1],
                    "egn_to_gn_nli_ratio_at_Lmax":ratios[-1],
                    "egn_to_gn_nli_ratio_at_Lmax_db":ratios_db[-1],
                    "snr_at_Lmax_db":10*math.log10(snrs[-1])}
        if not np.any(above):
            return {"status":"below_threshold_at_one_span","max_supported_spans":max_supported,
                    "nspans":1.0,"Lmax_km":0.0,"popt_dBm":egn.w_to_dbm(popts[0]),
                    "eta_at_Lmax_W_inv2":etas[0],
                    "egn_to_gn_nli_ratio_at_Lmax":ratios[0],
                    "egn_to_gn_nli_ratio_at_Lmax_db":ratios_db[0],
                    "snr_at_Lmax_db":10*math.log10(snrs[0])}

        last_pass=np.where(above)[0][-1]
        if last_pass==len(ns)-1:
            return {"status":"right_censored_span_limit","max_supported_spans":max_supported,
                    "nspans":float(ns[-1]),"Lmax_km":math.nan,
                    "popt_dBm":egn.w_to_dbm(popts[-1]),"eta_at_Lmax_W_inv2":etas[-1],
                    "egn_to_gn_nli_ratio_at_Lmax":ratios[-1],
                    "egn_to_gn_nli_ratio_at_Lmax_db":ratios_db[-1],
                    "snr_at_Lmax_db":10*math.log10(snrs[-1])}

        i0,i1=last_pass,last_pass+1
        y0,y1=np.log(snrs[i0]),np.log(snrs[i1])
        yt=math.log(snr_req)
        frac=(yt-y0)/(y1-y0) if y1!=y0 else 0.0
        n_cross=float(ns[i0]+frac*(ns[i1]-ns[i0]))
        # Interpolate secondary observables at the same crossing only for reporting.
        popt=float(popts[i0]+frac*(popts[i1]-popts[i0]))
        eta=float(np.exp(np.log(etas[i0])+frac*(np.log(etas[i1])-np.log(etas[i0]))))
        ratio=float(ratios[i0]+frac*(ratios[i1]-ratios[i0]))
        ratio_db=float(ratios_db[i0]+frac*(ratios_db[i1]-ratios_db[i0]))
        return {"status":"interpolated_full_egn","max_supported_spans":max_supported,
                "nspans":n_cross,"Lmax_km":n_cross*SPAN_KM,
                "popt_dBm":egn.w_to_dbm(popt),"eta_at_Lmax_W_inv2":eta,
                "egn_to_gn_nli_ratio_at_Lmax":ratio,
                "egn_to_gn_nli_ratio_at_Lmax_db":ratio_db,
                "snr_at_Lmax_db":10*math.log10(snr_req)}

    for _, r in subset.iterrows():
        reach=egn_reach_multispan(r)
        d=r.to_dict()
        d.update({
            "egn_eta_W_inv2":reach["eta_at_Lmax_W_inv2"],
            "egn_to_gn_nli_ratio":reach["egn_to_gn_nli_ratio_at_Lmax"],
            "egn_to_gn_nli_ratio_db":reach["egn_to_gn_nli_ratio_at_Lmax_db"],
            "egn_Lmax_km":reach["Lmax_km"],
            "egn_popt_dBm":reach["popt_dBm"],
            "egn_error_pct":(abs(reach["Lmax_km"]-r.paper_Lmax_km)/r.paper_Lmax_km*100.0
                             if math.isfinite(reach["Lmax_km"]) else math.nan),
            "egn_signed_error_pct":((reach["Lmax_km"]-r.paper_Lmax_km)/r.paper_Lmax_km*100.0
                                    if math.isfinite(reach["Lmax_km"]) else math.nan),
            "egn_status":reach["status"],
            "egn_max_supported_spans":reach["max_supported_spans"],
            "egn_span_count":reach["nspans"],
            "egn_snr_at_Lmax_db":reach["snr_at_Lmax_db"],
        })
        erows.append(d)
        print("EGN_ROW_JSON="+json.dumps({k:(v.item() if hasattr(v,"item") else v) for k,v in d.items()},default=float,separators=(",",":")))
    edf=pd.DataFrame(erows)
    edf.to_csv(OUT/"paper_gn_egn_comparison.csv",index=False)

    # Error summaries. MAPE is accompanied by digitization uncertainty because
    # the source paper supplies curves/markers, not the raw Fig.5 numeric table.
    summary={
        "figure5_points":int(len(gdf)),
        "gn_mape_pct":float(gdf.gn_error_pct.mean()),
        "gn_median_ape_pct":float(gdf.gn_error_pct.median()),
        "gn_max_ape_pct":float(gdf.gn_error_pct.max()),
        "egn_subset_points":int(len(edf)),
        "egn_subset_mape_pct":float(edf.egn_error_pct.mean()),
        "egn_subset_median_ape_pct":float(edf.egn_error_pct.median()),
        "egn_subset_max_ape_pct":float(edf.egn_error_pct.max()),
        "gn_same_subset_mape_pct":float(edf.gn_error_pct.mean()),
        "paper_digitization_uncertainty_pct_range":[
            float(ref.digitization_uncertainty_pct.min()),
            float(ref.digitization_uncertainty_pct.max())],
        "important_scope_notes":[
            "Paper Fig.5 GN uses incoherent span accumulation; the GN reproduction matches that convention.",
            "GN Tx PSD uses NRZ sinc^2 multiplied by a fourth-order super-Gaussian with Bopt=spacing, normalized to channel power.",
            "EGN_adaptive is evaluated with its native full coherent multi-span physics on the 50/38.4-GHz subset; its rectangular-spectrum requirement differs from the paper's optimized super-Gaussian Tx spectrum.",
            "No fitted NLI scale factor is used.",
            "Paper Fig.5 and Fig.3 values are digitized from the supplied PDF; error metrics therefore include digitization uncertainty."
        ]
    }
    (OUT/"summary.json").write_text(json.dumps(summary,indent=2),encoding="utf-8")
    svg_fig5(gdf,OUT/"figure5_reproduction.svg")
    svg_parity(edf,OUT/"paper_gn_egn_parity.svg")
    print("SUMMARY_JSON="+json.dumps(summary,separators=(",",":")))
    print("\nGN worst 10 points:")
    print(gdf.sort_values("gn_error_pct",ascending=False)[
        ["fiber","modulation","spacing_GHz","paper_Lmax_km","gn_Lmax_km","gn_error_pct"]
    ].head(10).to_string(index=False))
    print("\nGN/EGN comparison:")
    print(edf[["fiber","modulation","spacing_GHz","paper_Lmax_km","gn_Lmax_km","egn_Lmax_km","gn_error_pct","egn_error_pct"]].to_string(index=False))


if __name__=="__main__":
    main()
