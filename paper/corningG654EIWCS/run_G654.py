import math
import numpy as np
import matplotlib.pyplot as plt

from EGN_adaptive import (
    WDMSystem,
    Span,
    GNIntegralOptions,
    launch_power_vs_gsnr,
)


# ============================================================
# 1. G.654.E fiber parameters
# ============================================================

alpha_db_per_km = 0.166
Aeff_um2 = 125.0
D_ps_nm_km = 21.0
n2 = 2.2e-20
wavelength_nm = 1550.0

# gamma = 2*pi*n2 / (lambda*Aeff)
Aeff_m2 = Aeff_um2 * 1e-12
wavelength_m = wavelength_nm * 1e-9

gamma_W_inv_km = (
    2.0 * math.pi * n2
    / (wavelength_m * Aeff_m2)
    * 1e3
)

print("gamma =", gamma_W_inv_km, "1/(W km)")


# ============================================================
# 2. WDM parameters
# ============================================================

N_CHANNELS = 90
SPACING_GHz = 95.0
BAUD_GBd = 95.0

# Reference system is made at 0 dBm/ch.
# launch_power_vs_gsnr() then sweeps -10 ~ +10 dBm.
system = WDMSystem.equispaced(
    n_channels=N_CHANNELS,
    spacing_GHz=SPACING_GHz,
    baud_GBd=BAUD_GBd,
    power_dBm=0.0,
    pulse_shape="rect",
    rolloff=0.0,
)

# 90 channels -> no exact 0-THz center channel.
# index 44 = -47.5 GHz
# index 45 = +47.5 GHz
cut_index = 45


# ============================================================
# 3. Link
# ============================================================

span = Span(
    length_km=80.0,
    alpha_db_per_km=alpha_db_per_km,
    gamma_W_inv_km=gamma_W_inv_km,
    D_ps_nm_km=D_ps_nm_km,
    wavelength_nm=wavelength_nm,
    noise_figure_db=5.0,
)

N_SPANS = 30

spans = [span] * N_SPANS


# ============================================================
# 4. Numerical GN / EGN options
# ============================================================

gn_options = GNIntegralOptions(
    sobol_power=14,
    seed=1,
    accumulation="coherent",
    z_quadrature_order=64,
    egn_frequency_order=64,
)


# ============================================================
# 5. Launch power sweep
# ============================================================

launch_power_dBm = np.arange(-10.0, 11.0, 1.0)

result = launch_power_vs_gsnr(
    base_system=system,
    spans=spans,
    cut_index=cut_index,
    launch_power_dBm=launch_power_dBm,

    modulation="64QAM",

    # Figure condition
    trx_snr_db=18.0,
    shannon_gap_db=3.0,

    gn_options=gn_options,

    # Receiver integration
    receiver_points=3,

    # IMPORTANT:
    # 90 channels cannot use the strict precision egn_full path.
    nli_model="egn_sci",

    # First-order NLI follows P^3 scaling.
    use_cubic_scaling=True,
)


# ============================================================
# 6. Print all SNR_tot results
# ============================================================

print("\nLaunch Power [dBm]     SNR_tot [dB]")
print("----------------------------------")

for p, snr in zip(
    result["launch_power_dBm"],
    result["gsnr_db"],
):
    print(f"{p:8.1f}              {snr:8.4f}")


# ============================================================
# 7. SNR at exactly +4 dBm
# ============================================================

idx_4 = np.where(
    np.isclose(result["launch_power_dBm"], 4.0)
)[0][0]

snr_at_4 = result["gsnr_db"][idx_4]

print("\n================================")
print("Launch power = +4 dBm")
print(f"SNR_tot      = {snr_at_4:.6f} dB")
print("================================")


# ============================================================
# 8. Find optimum launch power
# ============================================================

idx_best = np.argmax(result["gsnr_db"])

best_power = result["launch_power_dBm"][idx_best]
best_snr = result["gsnr_db"][idx_best]

print("\nOptimum launch power")
print(f"Power   = {best_power:.1f} dBm")
print(f"SNR_tot = {best_snr:.6f} dB")


# ============================================================
# 9. Plot
# ============================================================

plt.figure(figsize=(8, 5))

plt.plot(
    result["launch_power_dBm"],
    result["gsnr_db"],
    marker="o",
    label="EGN_adaptive.py"
)

plt.scatter(
    [4.0],
    [snr_at_4],
    s=80,
    label=f"+4 dBm = {snr_at_4:.2f} dB"
)

plt.xlabel("Launch Power [dBm]")
plt.ylabel("SNR_tot [dB]")
plt.title("G.654.E - Launch Power vs SNR_tot")

plt.grid(True)
plt.legend()
plt.tight_layout()
plt.show()
