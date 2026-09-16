import numpy as np

from gn_integral_general import (
    WDMSystem,
    Span,
    GNIntegralOptions,
    gn_nli_psd_qmc,
    integrate_nli_over_channel,
)

from gn_integral_general_modulation import (
    evaluate_performance,
    launch_power_vs_gsnr,
)


# =========================================================
# 1. WDM 입력값
# =========================================================

N_CHANNELS = 50
SPACING_GHz = 95.0
BAUD_GBd = 95.0
LAUNCH_POWER_dBm = 0.0


system = WDMSystem.equispaced(
    n_channels=N_CHANNELS,
    spacing_GHz=SPACING_GHz,
    baud_GBd=BAUD_GBd,
    power_dBm=LAUNCH_POWER_dBm,
    pulse_shape="rect",
    rolloff=0.0,
)


# =========================================================
# 2. Fiber / Span 입력값
# G.654.E 예
# =========================================================

N_SPANS = 30

span = Span(
    length_km=80.0,
    alpha_db_per_km=0.166,
    gamma_W_inv_km=0.713,
    D_ps_nm_km=21.0,
    wavelength_nm=1550.0,
    noise_figure_db=5.0,
)

spans = [span for _ in range(N_SPANS)]


# =========================================================
# 3. GN numerical integration 설정
# =========================================================

options = GNIntegralOptions(
    sobol_power=14,          # 처음 테스트용: 2^14 samples
    seed=1,
    accumulation="coherent",
    z_quadrature_order=64,
)


# =========================================================
# 4. CUT 선택
# =========================================================

cut_index = 25
cut = system.channels[cut_index]

print("CUT index =", cut_index)
print("CUT center frequency =", cut.center_THz, "THz")


# =========================================================
# 5. GN NLI PSD 계산
# =========================================================

result_psd = gn_nli_psd_qmc(
    system=system,
    spans=spans,
    f_THz=cut.center_THz,
    options=options,
)

print("\n===== GN NLI PSD =====")
print(result_psd)


# =========================================================
# 6. CUT 전체 NLI power 계산
# =========================================================

p_nli = integrate_nli_over_channel(
    system=system,
    spans=spans,
    cut_index=cut_index,
    options=options,
    receiver_points=3,
)

print("\n===== NLI Power =====")
print("P_NLI =", p_nli, "W")


# =========================================================
# 7. ASE + NLI + TRX → GSNR / BER / Capacity
# =========================================================

performance = evaluate_performance(
    system=system,
    spans=spans,
    cut_index=cut_index,
    modulation="64QAM",
    coding_rate=1.0,
    trx_snr_db=18.0,
    shannon_gap_db=3.0,
    gn_options=options,
    receiver_points=3,
    nli_model="gn",
)

print("\n===== System Performance =====")
print(performance)


# =========================================================
# 8. Launch Power sweep
# =========================================================

powers = np.arange(-10.0, 11.0, 1.0)

sweep = launch_power_vs_gsnr(
    base_system=system,
    spans=spans,
    cut_index=cut_index,
    launch_power_dBm=powers,
    modulation="64QAM",
    coding_rate=1.0,
    trx_snr_db=18.0,
    shannon_gap_db=3.0,
    gn_options=options,
    receiver_points=3,
    nli_model="gn",
)

print("\n===== Launch Power vs GSNR =====")

for p, gsnr, snr_nli in zip(
    sweep["launch_power_dBm"],
    sweep["gsnr_db"],
    sweep["snr_nli_db"],
):
    print(
        f"Power = {p:5.1f} dBm, "
        f"GSNR = {gsnr:7.3f} dB, "
        f"SNR_NLI = {snr_nli:7.3f} dB"
    )

# 입력값
# n_channels		WDM 채널 수
# spacing_GHz		채널 중심 주파수 간격
# baud_GBd		    각 채널 symbol rate
# power_dBm    		채널당 launch power
# pulse_shape	    Nyquist rectangular PSD -> "rect"
# length_km		    span당 길이
# alpha_db_per_km	G.654.E attenuation
# gamma_W_inv_km	nonlinear coefficient
# D_ps_nm_km		dispersion
# noise_figure_db	EDFA NF
# N_SPANS		    30 × 80 = 2400 km
# accumulation     span NLI complex amplitude coherent 합 -> "coherent"
# modulation	    BER/rate 계산용 modulation
# trx_snr_db	    transceiver SNR
# shannon_gap_db	Shannon limit 대비 gap

# 결과값
# 1. GNResult
# - f_THz: NLI를 계산한 CUT의 상대 중심 주파수
# - g_nli_W_per_THz: CUT 중심에서의 NLI Power Spectral Density
# - integration_bounds_THz: 실제 occupied WDM bandwidth

# 2. P_NLI
# 3. PerformanceResult
# 4. GSNR
# 5.SNR_ASE
# 6. SNR_NLI
# 7. BER
# 8. P_ASE
# 9. Shannon capacity
