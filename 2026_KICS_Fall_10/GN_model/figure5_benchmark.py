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
    """Single requested comparison graph: paper vs GN vs EGN on the same x-axis."""
    import matplotlib.pyplot as plt
    d=df.copy()
    order={"PSCF":0,"SMF":1,"NZDSF":2}
    d["_order"]=d.fiber.map(order)
    d=d.sort_values("_order")
    x=np.arange(len(d))
    labels=[f"{r.fiber}\n{r.modulation} {r.spacing_GHz:g} GHz" for _,r in d.iterrows()]
    fig, ax = plt.subplots(figsize=(8.0,5.2))
    ax.plot(x,d.paper_Lmax_km,marker="o",label="Paper Fig. 5")
    ax.plot(x,d.gn_Lmax_km,marker="s",label="GN model")
    ax.plot(x,d.egn_Lmax_km,marker="^",label="EGN adaptive")
    ax.set_yscale("log")
    ax.set_xticks(x,labels)
    ax.set_ylabel("Maximum reach [km]")
    ax.set_title("Paper vs GN vs EGN — representative QPSK, 50 GHz")
    ax.grid(True,which="both",alpha=.25)
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

    # ---------- EGN representative comparison ----------
    # Carena-2012 Fig. 5 is a GN benchmark and its plotted model uses the paper's
    # incoherent span accumulation (Eq. 17). EGN_adaptive's native full path is
    # coherent and rectangular-spectrum. Therefore a direct "EGN should match
    # Fig. 5" accuracy claim would be invalid.
    #
    # To satisfy the requested three-way comparison without hiding this mismatch,
    # use a clearly-labelled paper-convention proxy on the representative 50-GHz
    # QPSK points (PSCF/SMF/NZDSF): calculate the FULL one-span EGN NLI from
    # EGN_adaptive.py, then apply the same incoherent N-span scaling used by the
    # 2012 Fig. 5 plotting convention. No fitted scale factor is used.
    subset = gdf[np.isclose(gdf.spacing_GHz,50.0) & (gdf.modulation=="QPSK")].copy()
    erows=[]
    eopt=egn.EGNFullOptions(receiver_points=5,max_receiver_points=5,panel_order=8,
                            quadrature_rtol=5e-4,verify_convergence=False,
                            strict_convergence=False)

    for _, r in subset.iterrows():
        system=egn.WDMSystem.equispaced(NCH,float(r.spacing_GHz),RS_GBD,0.0)
        span=egn.Span(SPAN_KM,float(r.alpha_dB_km),float(r.gamma_W_inv_km),
                      D_ps_nm_km=float(r.D_ps_nm_km),noise_figure_db=NF_DB)
        br=egn.full_egn_nli_power(
            system,[span],CUT,str(r.modulation),
            gn_options=egn.GNIntegralOptions(accumulation="coherent"),
            full_options=eopt)
        eta_egn=float(br.total_egn_W)/(1e-3**3)
        reach=lmax_from_eta(eta_egn,float(r.ase_one_span_W),float(r.paper_required_OSNR_dB_0p1nm))
        d=r.to_dict()
        d.update({
            "egn_eta_W_inv2":eta_egn,
            "egn_one_span_nli_W_at_0dBm":float(br.total_egn_W),
            "egn_to_gn_nli_ratio":float(br.total_egn_W/br.gn_total_W),
            "egn_to_gn_nli_ratio_db":float(br.ratio_to_gn_db),
            "egn_Lmax_km":float(reach["Lmax_km"]),
            "egn_popt_dBm":float(reach["popt_dBm"]),
            "egn_error_pct":abs(float(reach["Lmax_km"])-float(r.paper_Lmax_km))/float(r.paper_Lmax_km)*100.0,
            "egn_signed_error_pct":(float(reach["Lmax_km"])-float(r.paper_Lmax_km))/float(r.paper_Lmax_km)*100.0,
            "egn_status":"paper_incoherent_scaling_of_one_span_full_egn",
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
            "The EGN three-way comparison is a paper-convention proxy: full one-span EGN is calculated by EGN_adaptive, then scaled incoherently across spans like Carena-2012 Fig.5. It is not a native multi-span EGN accuracy benchmark.",
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
