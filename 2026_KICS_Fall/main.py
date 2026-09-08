"""G.654.E GN report: one main entry point; all launch powers are total DP power.
Source model: kimheeseo/LSCNS dnanf_fig2_hcf_colab.ipynb (retrieved 2026-09-08).
Change only USER_INPUTS. Fiber and source-system parameters are immutable.
"""

# ===================== 이 부분만 수정하세요 =====================
USER_INPUTS = {
    "task": "all",                   # all | snr | compare | point
    "model": "gn_finite",            # gn_finite | legacy | calibrated
    "spans": list(range(1, 51)),      # 계산할 span 수, 예: [10, 20, 30]
    "launch_min_dbm": -10.0,
    "launch_max_dbm": 10.0,
    "launch_step_db": 0.05,
    "point_launch_dbm": 4.0,          # 단일 운용점의 채널당 DP 전력
    "point_spans": 30,
    "reference_spans_assumed": 30,    # 첨부 그림에 없는 조건: 비교를 위한 가정
    "bandwidth_policy": "provisional", # provisional: 조건부 계산 | strict: 불일치 시 중지
    "include_historical_diagnostic": False, # 이전 4.8 THz ×0.5 진단선 추가
    "save_outputs": True,
    "show_outputs": True,
    "output_dir": "G654E_results",
}
# ================================================================

from dataclasses import dataclass, asdict
from pathlib import Path
import json
import platform
import zipfile
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from scipy.optimize import least_squares
import scipy

try:
    from IPython import get_ipython
    from IPython.display import display, Markdown, FileLink, HTML
    if get_ipython() is not None:
        get_ipython().run_line_magic("matplotlib", "inline")
except ImportError:
    display = Markdown = FileLink = None


@dataclass(frozen=True)
class FiberFixed:
    wavelength_nm: float = 1550.0
    attenuation_db_km: float = 0.166
    effective_area_um2: float = 125.0
    dispersion_ps_nm_km: float = 21.0
    n2_m2_w: float = 2.2e-20


@dataclass(frozen=True)
class SystemFixed:
    channels: int = 90
    symbol_rate_gbd: float = 95.0
    spacing_ghz: float = 95.0
    span_length_km: float = 80.0
    noise_figure_db: float = 5.0
    transceiver_snr_db: float = 18.0
    stated_gain_bandwidth_thz: float = 4.8
    qam_order_record_only: int = 64
    shannon_gap_db_record_only: float = 3.0
    polarizations: int = 2


FIBER = FiberFixed()
SYSTEM = SystemFixed()
C0 = 299_792_458.0
H_PLANCK = 6.626_070_15e-34
DB_PER_NEPER_POWER = 10.0 / np.log(10.0)
SOURCE_NOTEBOOK_URL = "https://github.com/kimheeseo/LSCNS/blob/main/paper/hcf-optimum-launch-power-reproduction/dnanf_fig2_hcf_colab.ipynb"
SOURCE_NOTEBOOK_SHA256 = "d99239f33a9082f6d38f29a9c16596a9284c4bb3ff7818518b3f909e5c21d8c8"
REFERENCE_IMAGE_SHA256 = "d23fa4e0c7c44246e25754f7be3315d8265243095a4ac7a94506bd0923e4d220"

# 첨부된 파란 곡선을 픽셀 좌표로 읽은 값. 실측 원시 데이터가 아닙니다.
# x: pixel 88 -> -10 dBm, 653 -> +10 dBm
# y: pixel 80 -> 16 dB, 331 -> 2 dB; 파란 선의 열별 중앙값을 선형 보간.
# 격자 0.1 dB는 이미지보다 촘촘하므로 201개 점은 독립 측정 201회가 아닙니다.
REFERENCE_X_DBM = np.linspace(-10.0, 10.0, 201)
REFERENCE_Y_DB = np.array([
    6.043824701195, 6.155378486056, 6.219521912351, 6.335956175299, 6.442629482072, 6.500498007968, 6.683864541833, 6.734760956175,
    6.824701195219, 6.948107569721, 7.026892430279, 7.135657370518, 7.265338645418, 7.326693227092, 7.441035856574, 7.532370517928,
    7.605577689243, 7.691334661355, 7.824501992032, 7.894223107570, 7.996015936255, 8.107569721116, 8.167529880478, 8.329282868526,
    8.380876494024, 8.470119521912, 8.566334661355, 8.645119521912, 8.804780876494, 8.858466135458, 8.937250996016, 9.043924302789,
    9.111553784861, 9.201494023904, 9.308167330677, 9.362549800797, 9.485258964143, 9.587051792829, 9.669322709163, 9.757868525896,
    9.836653386454, 9.920318725100, 9.994223107570, 10.086254980080, 10.171314741036, 10.258466135458, 10.363745019920, 10.443924302789,
    10.505976095618, 10.601494023904, 10.680278884462, 10.761155378486, 10.862948207171, 10.944521912351, 11.007968127490, 11.074203187251,
    11.180876494024, 11.231772908367, 11.310557768924, 11.408167330677, 11.509960159363, 11.546912350598, 11.625697211155, 11.731673306773,
    11.777689243028, 11.844621513944, 11.953386454183, 12.011952191235, 12.073306772908, 12.149302788845, 12.228087649402, 12.290836653386,
    12.357768924303, 12.408665338645, 12.541832669323, 12.594123505976, 12.664541832669, 12.723804780876, 12.812350597610, 12.881374501992,
    12.904382470120, 12.955278884462, 13.089840637450, 13.127490039841, 13.163745019920, 13.242529880478, 13.321314741036, 13.393824701195,
    13.423107569721, 13.490039840637, 13.566733067729, 13.629482071713, 13.679681274900, 13.705478087649, 13.784262948207, 13.880478087649,
    13.908366533865, 13.936952191235, 13.987848605578, 14.047808764940, 14.131474103586, 14.140537848606, 14.187250996016, 14.243027888446,
    14.298804780876, 14.333665338645, 14.395019920319, 14.438247011952, 14.468924302789, 14.521912350598, 14.549800796813, 14.577689243028,
    14.644621513944, 14.689243027888, 14.745019920319, 14.745019920319, 14.800796812749, 14.815438247012, 14.828685258964, 14.889342629482,
    14.940239043825, 14.963247011952, 14.996015936255, 14.996015936255, 15.023904382470, 15.023904382470, 15.050398406375, 15.079681274900,
    15.079681274900, 15.135458167331, 15.135458167331, 15.137549800797, 15.188446215139, 15.191235059761, 15.191235059761, 15.191235059761,
    15.191235059761, 15.191235059761, 15.191235059761, 15.191235059761, 15.219123505976, 15.191235059761, 15.191235059761, 15.191235059761,
    15.141035856574, 15.107569721116, 15.095019920319, 15.079681274900, 15.079681274900, 15.053884462151, 15.030876494024, 15.023904382470,
    14.996015936255, 14.961852589641, 14.856573705179, 14.828685258964, 14.809163346614, 14.800796812749, 14.745019920319, 14.689243027888,
    14.689243027888, 14.582569721116, 14.531673306773, 14.466135458167, 14.393625498008, 14.298804780876, 14.244422310757, 14.187250996016,
    14.075697211155, 14.047808764940, 13.929282868526, 13.876294820717, 13.774501992032, 13.637151394422, 13.570916334661, 13.462151394422,
    13.367330677291, 13.265537848606, 13.159561752988, 13.089840637450, 12.932270916335, 12.849302788845, 12.724501992032, 12.571115537849,
    12.424701195219, 12.329183266932, 12.166733067729, 12.067729083665, 11.922709163347, 11.790936254980, 11.684262948207, 11.517629482072,
    11.364940239044, 11.224800796813, 11.059561752988, 10.903386454183, 10.765338645418, 10.660059760956, 10.458565737052, 10.324003984064,
    10.171314741036,
], dtype=float)
REFERENCE_X_DBM.setflags(write=False)
REFERENCE_Y_DB.setflags(write=False)
_ALLOWED_INPUTS = frozenset({"task", "model", "spans", "launch_min_dbm", "launch_max_dbm",
    "launch_step_db", "point_launch_dbm", "point_spans", "reference_spans_assumed",
    "bandwidth_policy", "include_historical_diagnostic", "save_outputs", "show_outputs", "output_dir"})

MODEL_NAMES = {
    "legacy": "Previous GN (infinite effective length)",
    "gn_finite": "GN with finite effective length",
    "calibrated": "Reference-calibrated GN (empirical)",
    "historical_diagnostic": "Historical 4.8 THz x0.5 diagnostic",
}
MODEL_COLORS = {"legacy": "#D55E00", "gn_finite": "#6B7280",
                "calibrated": "#009E73", "historical_diagnostic": "#CC79A7"}


def _validate(user_inputs):
    allowed = _ALLOWED_INPUTS
    unknown = set(user_inputs) - allowed
    if unknown:
        raise ValueError(f"지원하지 않는 입력 또는 고정값 변경: {sorted(unknown)}")
    cfg = {**USER_INPUTS, **user_inputs}
    if asdict(FIBER) != asdict(FiberFixed()) or asdict(SYSTEM) != asdict(SystemFixed()):
        raise ValueError("고정 조건이 변경되었습니다. 원래 G.654.E/시스템 조건을 복구하세요.")
    if cfg["task"] not in {"all", "snr", "compare", "point"}:
        raise ValueError("task: all, snr, compare, point 중 하나를 입력하세요.")
    if cfg["model"] not in {"legacy", "gn_finite", "calibrated"}:
        raise ValueError("model: legacy, gn_finite, calibrated 중 하나를 입력하세요.")
    if cfg["bandwidth_policy"] not in {"provisional", "strict"}:
        raise ValueError("bandwidth_policy: provisional 또는 strict")
    vals = list(cfg["spans"])
    for name, value in [("spans", v) for v in vals] + [
        ("point_spans", cfg["point_spans"]),
        ("reference_spans_assumed", cfg["reference_spans_assumed"])]:
        if isinstance(value, bool) or not isinstance(value, (int, np.integer)) or not 1 <= value <= 50:
            raise ValueError(f"{name}: 원 조건 범위인 1~50의 정수만 입력하세요.")
    if not vals:
        raise ValueError("spans를 하나 이상 입력하세요.")
    cfg["spans"] = sorted(set(int(v) for v in vals))
    for key in ["launch_min_dbm", "launch_max_dbm", "launch_step_db", "point_launch_dbm"]:
        if isinstance(cfg[key], bool) or not np.isfinite(cfg[key]):
            raise ValueError(f"{key}: 유한한 숫자가 필요합니다.")
    lo, hi, step = [float(cfg[k]) for k in ["launch_min_dbm", "launch_max_dbm", "launch_step_db"]]
    if hi <= lo or step <= 0 or (hi-lo)/step > 20000:
        raise ValueError("전력 범위는 min < max, step > 0, 최대 20,001점으로 지정하세요.")
    if min(lo, cfg["point_launch_dbm"]) < -100 or max(hi, cfg["point_launch_dbm"]) > 40:
        raise ValueError("이 보고서의 수치 탐색 범위는 -100~+40 dBm/channel입니다.")
    return cfg


def derive_parameters():
    """Unit-safe conversion. Alpha here is POWER attenuation in km^-1."""
    lam = FIBER.wavelength_nm * 1e-9
    rate = SYSTEM.symbol_rate_gbd * 1e9
    alpha = FIBER.attenuation_db_km / DB_PER_NEPER_POWER
    beta2 = -(lam**2 / (2*np.pi*C0)) * (FIBER.dispersion_ps_nm_km*1e-6) * 1000
    gamma = 2*np.pi*FIBER.n2_m2_w / (lam*FIBER.effective_area_um2*1e-12) * 1000
    leff = -np.expm1(-alpha*SYSTEM.span_length_km) / alpha
    linf = 1/alpha
    gain = 10**(FIBER.attenuation_db_km*SYSTEM.span_length_km/10)
    ase = H_PLANCK*(C0/lam)*10**(SYSTEM.noise_figure_db/10)*rate*(gain-1)
    return {"wavelength_m": lam, "frequency_hz": C0/lam, "rate_hz": rate,
            "alpha_power_per_km": alpha, "alpha_field_per_km": alpha/2,
            "beta2_s2_per_km": beta2, "beta2_ps2_per_km": beta2*1e24,
            "gamma_per_w_km": gamma, "effective_length_km": leff,
            "asymptotic_length_km": linf,
            "span_gain_db": FIBER.attenuation_db_km*SYSTEM.span_length_km,
            "span_gain_linear": gain, "ase_per_span_w": ase,
            "wdm_occupied_bandwidth_thz": SYSTEM.channels*SYSTEM.spacing_ghz/1000,
            "channels_that_fit_4p8thz": int(SYSTEM.stated_gain_bandwidth_thz*1000//SYSTEM.spacing_ghz)}


def gn_eta_per_span(params, finite=True, bandwidth_thz=None):
    """Equal-channel central-channel GN approximation; eta excludes span count.
    Power P is total two-polarization channel power. No extra factor 0.5.
    Finite Leff improves the numerator; this remains a closed-form approximation.
    bandwidth_thz is used ONLY for the separately labelled historical diagnostic.
    """
    b, r = abs(params["beta2_s2_per_km"]), params["rate_hz"]
    linf = params["asymptotic_length_km"]
    leff = params["effective_length_km"] if finite else linf
    nch = SYSTEM.channels if bandwidth_thz is None else bandwidth_thz*1e12/(SYSTEM.spacing_ghz*1e9)
    argument = (np.pi**2/2)*b*linf*r**2*nch**(2*r/(SYSTEM.spacing_ghz*1e9))
    return float((8/27)*params["gamma_per_w_km"]**2*leff**2/(np.pi*b*linf*r**2)*np.arcsinh(argument))


def simulate(launch_dbm, spans, eta_per_span, params):
    launch = np.atleast_1d(np.asarray(launch_dbm, dtype=float))
    power = 1e-3*10**(launch/10)
    ase = np.full_like(power, spans*params["ase_per_span_w"])
    nli = spans*eta_per_span*power**3
    trx = power/10**(SYSTEM.transceiver_snr_db/10)
    snr = power/(ase+nli+trx)
    return pd.DataFrame({"launch_dbm": launch, "spans": spans,
        "distance_km": spans*SYSTEM.span_length_km, "signal_w": power,
        "ase_w": ase, "nli_w": nli, "trx_equivalent_noise_w": trx,
        "snr_linear": snr, "snr_db": 10*np.log10(snr)})


def optimum(eta, params):
    # d(ASE/P + eta*P^2 + 1/SNRtrx)/dP = 0; incoherent identical-span case.
    p = (params["ase_per_span_w"]/(2*eta))**(1/3)
    return float(10*np.log10(p/1e-3))


def fit_reference(params, eta_finite, assumed_spans):
    """Only one empirical k is fitted. Fiber, ASE, TRX and integer Ns stay fixed."""
    def residual(log_k):
        return simulate(REFERENCE_X_DBM, assumed_spans, eta_finite*np.exp(log_k[0]), params)["snr_db"].to_numpy()-REFERENCE_Y_DB
    fit = least_squares(residual, np.log([0.5]), bounds=(np.log([1e-4]), np.log([100.])),
                        ftol=1e-12, xtol=1e-12, gtol=1e-12)
    if not fit.success:
        raise RuntimeError("참조 곡선 보정계수 계산 실패: "+fit.message)
    return float(np.exp(fit.x[0]))


def error_metrics(predicted_db):
    delta = np.asarray(predicted_db)-REFERENCE_Y_DB
    # dB percentage is not meaningful: percentages are computed in linear SNR.
    return {"mae_db": float(np.mean(np.abs(delta))),
            "rmse_db": float(np.sqrt(np.mean(delta**2))),
            "max_abs_error_db": float(np.max(np.abs(delta))),
            "linear_snr_mape_pct": float(100*np.mean(np.abs(10**(delta/10)-1)))}


def _show_table(title, frame, enabled):
    if not enabled:
        return
    if display is not None:
        display(Markdown("**"+title+"**"))
        with pd.option_context("display.max_rows", 60, "display.max_columns", 20):
            display(HTML(frame.to_html(index=False, float_format=lambda v: f"{v:.7g}")))
    else:
        print(title)
        print(frame.to_string(index=False))


def _plot_snr(curves, cfg, params, eta):
    fig, ax = plt.subplots(figsize=(11, 6.5), layout="constrained")
    cmap = plt.get_cmap("viridis")
    selected = set(cfg["spans"]) if len(cfg["spans"]) <= 8 else {1,10,20,30,40,50,cfg["spans"][0],cfg["spans"][-1]}
    for ns, frame in curves.groupby("spans"):
        strong = ns in selected
        ax.plot(frame.launch_dbm, frame.snr_db, color=cmap((ns-1)/49),
                lw=2.2 if strong else 0.6, alpha=1 if strong else .20,
                label=f"{ns} {'span' if ns==1 else 'spans'} / {ns*80:,} km" if strong else None)
    p_opt = optimum(eta, params)
    if cfg["launch_min_dbm"] <= p_opt <= cfg["launch_max_dbm"]:
        ax.axvline(p_opt, color="#374151", ls="--", lw=1.3, label=f"Analytical optimum: {p_opt:+.2f} dBm/ch")
    ax.axhline(18, color="#94A3B8", ls=":", lw=1)
    ax.set(xlabel="Launch power per channel, total DP (dBm)", ylabel="End-to-end SNR (dB)",
           title="G.654.E | "+MODEL_NAMES[cfg["model"]],
           xlim=(cfg["launch_min_dbm"],cfg["launch_max_dbm"]))
    ax.grid(color="#E2E8F0", lw=.8)
    ax.spines[["top","right"]].set_visible(False)
    ax.legend(loc="lower center", ncol=2, fontsize=8.5, framealpha=.97)
    fig.get_layout_engine().set(rect=(0,.065,1,.90))
    fig.text(.5,.025,"90 x 95 GHz = 8.55 THz; supplied 4.8 THz gain bandwidth is insufficient. Conditional model only.",
             ha="center", fontsize=8.8, color="#92400E")
    return fig


def _plot_comparison(comparison, cfg):
    fig, (ax, err) = plt.subplots(2,1,figsize=(11,8),sharex=True,
                                 gridspec_kw={"height_ratios":[3,1]},layout="constrained")
    x = comparison.launch_dbm
    ax.plot(x, comparison.reference_snr_db, color="#0072B2", lw=2.8, label="Supplied blue curve (digitized)")
    models = ["legacy","gn_finite","calibrated"]
    if cfg["include_historical_diagnostic"]:
        models.append("historical_diagnostic")
    for model in models:
        y = comparison[model+"_snr_db"]
        style = "-" if model=="calibrated" else "--"
        ax.plot(x,y,lw=2,ls=style,color=MODEL_COLORS[model],label=MODEL_NAMES[model])
        err.plot(x,y-comparison.reference_snr_db,lw=1.6,ls=style,color=MODEL_COLORS[model])
    ax.set(title=f"G.654.E comparison | {cfg['reference_spans_assumed']} spans ASSUMED for the reference",
           ylabel="End-to-end SNR (dB)")
    ax.legend(loc="lower center",fontsize=8.6)
    err.axhline(0,color="#374151",lw=.8)
    err.set(xlabel="Launch power per channel, total DP (dBm)",ylabel="Model - blue\n(dB)",xlim=(-10,10))
    for a in [ax,err]:
        a.grid(color="#E2E8F0",lw=.8)
        a.spines[["top","right"]].set_visible(False)
    fig.get_layout_engine().set(rect=(0,.085,1,.88))
    fig.text(.5,.048,"90-channel results assume flat gain across 8.55 THz; the stated 4.8 THz is insufficient.",
             ha="center",fontsize=9,color="#92400E")
    fig.text(.5,.020,"Green uses k fitted to this same image; this is reference agreement, not independent validation.",
             ha="center",fontsize=9,color="#374151")
    return fig


def main(user_inputs=None):
    """Return tables, metadata and figures; change USER_INPUTS to select results."""
    cfg = _validate({} if user_inputs is None else user_inputs)
    p = derive_parameters()
    conflict = p["wdm_occupied_bandwidth_thz"] > SYSTEM.stated_gain_bandwidth_thz
    message = ("입력 정합성: 90채널 × 95 GHz = 8.55 THz이며, 제시된 증폭 대역폭 4.8 THz를 초과합니다. "
               "원 입력은 변경하지 않았습니다. 계산에는 90채널을 사용하며, 전 대역 평탄 이득을 가정한 조건부 결과입니다.")
    if conflict and cfg["bandwidth_policy"] == "strict":
        raise ValueError(message+" strict 정책으로 계산을 중지합니다.")
    if cfg["show_outputs"]:
        print(message)
        print("비교 그림의 span 수는 미표기: reference_spans_assumed는 비교용 가정입니다.")
        if cfg["model"] == "calibrated":
            print("calibrated 선택: 동일 그림에서 추정한 k를 사용합니다. 다른 span에의 전이는 미검증입니다.")
    eta_legacy = gn_eta_per_span(p, finite=False)
    eta_finite = gn_eta_per_span(p, finite=True)
    k = fit_reference(p, eta_finite, cfg["reference_spans_assumed"])
    # 이전 답변의 숫자를 재현하기 위해서만 4.343 반올림을 보존합니다.
    historical_params = dict(p)
    historical_a = FIBER.attenuation_db_km/4.343
    historical_params.update(effective_length_km=-np.expm1(-historical_a*80)/historical_a,
                             asymptotic_length_km=1/historical_a)
    eta_4p8_inf = gn_eta_per_span(p, finite=False, bandwidth_thz=4.8)
    eta_4p8_fin = gn_eta_per_span(historical_params, finite=True, bandwidth_thz=4.8)
    etas = {"legacy": eta_legacy, "gn_finite": eta_finite, "calibrated": k*eta_finite,
            "historical_diagnostic": .5*eta_4p8_fin}
    selected_eta = etas[cfg["model"]]
    grid = np.arange(cfg["launch_min_dbm"],cfg["launch_max_dbm"]+cfg["launch_step_db"]*1e-7,cfg["launch_step_db"])
    if grid[-1] < cfg["launch_max_dbm"]-1e-8:
        grid = np.append(grid,cfg["launch_max_dbm"])
    curves = pd.concat([simulate(grid,ns,selected_eta,p) for ns in cfg["spans"]],ignore_index=True)
    curves["model"] = cfg["model"]
    curves["conditional_bandwidth_conflict"] = conflict
    p_opt = optimum(selected_eta,p)
    opt_rows = []
    for ns, frame in curves.groupby("spans"):
        opt = simulate([p_opt],int(ns),selected_eta,p).iloc[0]
        sample = frame.loc[frame.snr_db.idxmax()]
        opt_rows.append({"spans": int(ns), "distance_km": int(ns)*80,
            "analytic_optimum_launch_dbm": p_opt, "analytic_max_snr_db": opt.snr_db,
            "grid_best_launch_dbm": sample.launch_dbm, "grid_best_snr_db": sample.snr_db,
            "optimum_inside_scan": bool(grid[0]<=p_opt<=grid[-1]), "model": cfg["model"]})
    optimum_table = pd.DataFrame(opt_rows)
    point = simulate([cfg["point_launch_dbm"]],cfg["point_spans"],selected_eta,p)
    point["model"] = cfg["model"]
    comparison = pd.DataFrame({"launch_dbm":REFERENCE_X_DBM,"reference_snr_db":REFERENCE_Y_DB})
    metric_rows = []
    for model, eta in etas.items():
        pred = simulate(REFERENCE_X_DBM,cfg["reference_spans_assumed"],eta,p).snr_db.to_numpy()
        comparison[model+"_snr_db"] = pred
        comparison[model+"_error_db"] = pred-REFERENCE_Y_DB
        best_p = optimum(eta,p)
        metric_rows.append({"model":model,"spans_assumed":cfg["reference_spans_assumed"],
            "eta_per_span_w_inv2":eta, "optimum_launch_dbm":best_p,
            "maximum_snr_db":float(simulate([best_p],cfg["reference_spans_assumed"],eta,p).snr_db.iloc[0]),
            **error_metrics(pred)})
    metrics = pd.DataFrame(metric_rows)
    stages = pd.DataFrame([
        ["Previous infinite-length, 90ch",eta_legacy,"provisional 8.55 THz"],
        ["Finite effective length, 90ch",eta_finite,"main GN baseline; k=1"],
        ["Calibrated finite GN, 90ch",k*eta_finite,f"empirical k={k:.8f}; Ns assumed"],
        ["Historical: infinite-length, 4.8 THz",eta_4p8_inf,"incompatible with 90 Nyquist channels"],
        ["Historical: finite length, 4.8 THz",eta_4p8_fin,"diagnostic only"],
        ["Historical: finite, 4.8 THz, x0.5",.5*eta_4p8_fin,"0.5 has no verified polarization justification"],
    ],columns=["stage","eta_per_span_w_inv2","meaning"])
    metadata = {"report_version":"2026-09-08.1", "fiber_fixed":asdict(FIBER),"system_fixed":asdict(SYSTEM),
        "user_inputs":cfg,"derived":p,"eta_models_per_span_w_inv2":etas,"empirical_k":k,
        "bandwidth_conflict":conflict,"reference_spans_known":False,
        "reference_kind":"user-supplied raster graph; experimental status unknown",
        "reference_pixel_scale_db":14/251,"reference_uncertainty_note":"about +/-0.1 dB; heuristic, not a confidence interval",
        "reference_peak_launch_dbm":float(REFERENCE_X_DBM[np.argmax(REFERENCE_Y_DB)]),
        "reference_peak_snr_db":float(np.max(REFERENCE_Y_DB)),
        "source_notebook_url":SOURCE_NOTEBOOK_URL,"source_notebook_sha256":SOURCE_NOTEBOOK_SHA256,
        "reference_image_sha256":REFERENCE_IMAGE_SHA256,
        "versions":{"python":platform.python_version(),"numpy":np.__version__,"pandas":pd.__version__,"scipy":scipy.__version__},
        "assumptions":["identical 80 km spans; one loss-compensating amplifier per span",
                       "constant NF; no saturation; incoherent NLI accumulation; central-channel approximation",
                       "signal power is total DP power; no additional polarization factor",
                       "no inter-channel Raman scattering, gain tilt, PMD/PDL or filter penalties",
                       "64QAM and Shannon gap are recorded but not used in physical SNR",
                       "span count is not a wet-repeater count; receiver-end amplifier is included"]}
    tables = {"derived_parameters":pd.DataFrame(p.items(),columns=["parameter","value"]),
        "fiber_fixed":pd.DataFrame(asdict(FIBER).items(),columns=["parameter","value"]),
        "system_fixed":pd.DataFrame(asdict(SYSTEM).items(),columns=["parameter","value"]),
        "snr_curves":curves,"optimum_summary":optimum_table,"operating_point":point,
        "reference_comparison":comparison,"reference_error_metrics":metrics,"eta_stages":stages}
    figures = {}
    plt.rcParams.update({"font.family":"DejaVu Sans","font.size":10,"axes.titlesize":13,"savefig.facecolor":"white"})
    if cfg["task"] in {"all","snr"}:
        figures["G654E_snr_vs_launch"] = _plot_snr(curves,cfg,p,selected_eta)
    if cfg["task"] in {"all","compare"}:
        figures["G654E_reference_comparison"] = _plot_comparison(comparison,cfg)
    _show_table("고정 G.654.E 물성",tables["fiber_fixed"],cfg["show_outputs"])
    _show_table("계산된 주요 파라미터",tables["derived_parameters"],cfg["show_outputs"])
    if cfg["task"] in {"all","snr"}:
        excerpt = optimum_table if len(optimum_table)<=10 else optimum_table[optimum_table.spans.isin([1,10,20,30,40,50])]
        _show_table("최적점 요약 (전체 span 결과는 CSV)",excerpt,cfg["show_outputs"])
    if cfg["task"] in {"all","compare"}:
        _show_table("첨부 파란 곡선과의 일치도: 동일 그림 보정 포함",metrics,cfg["show_outputs"])
        _show_table("η 변화 단계: 과거 진단 조건까지 공개",stages,cfg["show_outputs"])
    if cfg["task"] in {"all","point"}:
        _show_table("입력 운용점의 SNR 및 잡음 성분",point,cfg["show_outputs"])
    out = Path(cfg["output_dir"])
    saved = []
    if cfg["save_outputs"]:
        out.mkdir(parents=True,exist_ok=True)
        for name,frame in tables.items():
            path=out/(name+".csv")
            frame.to_csv(path,index=False,encoding="utf-8-sig")
            saved.append(path)
        meta_path=out/"parameters_and_provenance.json"
        meta_path.write_text(json.dumps(metadata,ensure_ascii=False,indent=2),encoding="utf-8")
        saved.append(meta_path)
        for name,fig in figures.items():
            path=out/(name+".png")
            fig.savefig(path,dpi=180)
            saved.append(path)
        zip_path=out/"G654E_results.zip"
        with zipfile.ZipFile(zip_path,"w",zipfile.ZIP_DEFLATED) as z:
            for path in saved:
                z.write(path,path.name)
    if cfg["show_outputs"]:
        plt.show()
        if cfg["save_outputs"]:
            print("CSV/PNG/설정 ZIP:", str(out/"G654E_results.zip"))
            if display is not None:
                display(FileLink(str(out/"G654E_results.zip")))
    else:
        for fig in figures.values():
            plt.close(fig)
    return {"config":cfg,"parameters":p,"eta":etas,"empirical_k":k,
            "tables":tables,"figures":figures,"metadata":metadata}


if __name__ == "__main__":
    RESULTS = main(USER_INPUTS)
