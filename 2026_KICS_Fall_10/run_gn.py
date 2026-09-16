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
# n_channels         WDM 채널 수
# spacing_GHz        채널 중심 주파수 간격 [GHz]
# baud_GBd           각 채널 symbol rate [GBd]
# power_dBm          채널당 launch power [dBm]
# pulse_shape        Tx PSD shape, Nyquist rectangular -> "rect"

# length_km          span당 길이 [km]
# alpha_db_per_km    광섬유 attenuation [dB/km]
# gamma_W_inv_km     nonlinear coefficient [1/(W·km)]
# D_ps_nm_km         dispersion [ps/(nm·km)]
# noise_figure_db    EDFA noise figure [dB]

# N_SPANS            span 수 (예: 30 × 80 km = 2400 km)
# accumulation       span별 NLI accumulation 방식
#                    -> "coherent": complex amplitude 합 후 power 계산
#                    -> "incoherent": span별 NLI power 합
# modulation         BER / data rate 계산용 modulation
# trx_snr_db         transceiver SNR [dB]
# shannon_gap_db     Shannon limit 대비 implementation gap [dB]


# 결과값
# 1. GNResult
# - f_THz
#   NLI PSD를 계산한 주파수, 일반적으로 CUT 중심의 상대 주파수 [THz]
# - g_nli_W_per_THz
#   해당 주파수에서의 NLI Power Spectral Density [W/THz]
# - qmc_samples
#   GN 2-D 적분에 사용한 Sobol QMC sample 수
# - integration_bounds_THz
#   GN 적분에 사용한 전체 WDM occupied spectrum의
#   최소/최대 주파수 경계 [THz]
# - accumulation
#   사용된 NLI accumulation 방식

# 2. P_NLI
# - CUT receiver bandwidth 내에서 G_NLI(f)를 적분한 총 NLI power [W]

# 3. PerformanceResult
# - launch_power_dBm
#   CUT 채널 launch power [dBm]
# - cut_output_power_W
#   링크 출력에서의 CUT signal power [W]
# - p_ase_W
#   CUT bandwidth 내 ASE noise power [W]
# - p_nli_W
#   CUT bandwidth 내 nonlinear interference power [W]
# - snr_ase_db
#   ASE만 고려한 SNR [dB]
# - snr_nli_db
#   NLI만 고려한 SNR [dB]
# - gsnr_db
#   ASE + NLI + optional TRX noise를 고려한 GSNR [dB]
# - ber
#   modulation에 따른 approximate uncoded BER
# - gross_rate_gbps
#   CUT 1채널의 DP gross data rate [Gb/s]
# - net_rate_gbps
#   coding_rate를 적용한 CUT 1채널 net data rate [Gb/s]
#
# - shannon_gap_capacity_gbps
#   GSNR과 Shannon gap을 이용한 CUT 1채널 capacity [Gb/s]
