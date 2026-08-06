import numpy as np

from gn_model_gsnr import SNR_GD, optimized_capacity_extended


COMMON = dict(
    lamb=1550.0,
    alpha=0.15,
    NF=4.5,
    N=75,
    gamma=0.81,
    Ltot=6000.0,
    B=4.3e12,
    D=-21e-12,
    Rs=100e9,
    gamma_XT=0.0,
    l_FIFO=0.0,
    delta_f=112.5e9,
)


def test_snr_curve_has_finite_interior_maximum():
    powers = np.geomspace(1e-3, 0.063, 200)
    snr = SNR_GD(powers, **COMMON)
    idx = int(np.argmax(snr))
    assert np.all(np.isfinite(snr))
    assert 0 < idx < len(powers) - 1
    assert snr[idx] > snr[0]
    assert snr[idx] > snr[-1]


def test_xt_db_and_linear_units_match():
    powers = np.array([0.01])
    db_value = -80.0
    linear_value = 10 ** (db_value / 10)
    snr_db = SNR_GD(
        powers, **{**COMMON, "gamma_XT": db_value}, gamma_xt_unit="dB_per_km"
    )
    snr_linear = SNR_GD(
        powers,
        **{**COMMON, "gamma_XT": linear_value},
        gamma_xt_unit="linear_per_km",
    )
    np.testing.assert_allclose(snr_db, snr_linear, rtol=1e-12)


def test_optimizer_returns_capacity_point_snr():
    spans = np.array([60.0, 80.0, 100.0])
    powers = np.geomspace(1e-3, 0.063, 120)
    eta = np.full((len(spans), len(powers)), 0.025)
    result = optimized_capacity_extended(
        spans,
        6000.0,
        4.5,
        18e3,
        1.0,
        eta,
        0.10,
        12,
        0.81,
        0.0,
        0.15,
        0.0,
        -21e-12,
        100e9,
        112.5e9,
        powers,
        1550.0,
        4.3e12,
        50.0,
        return_details=True,
    )
    best_idx = int(np.where(spans == result.optimal_span_length_km)[0][0])
    assert result.maximum_capacity_bps > 0
    assert result.gsnr_at_capacity_optimum_linear == result.gsnr_by_span_linear[best_idx]
    assert result.launch_power_by_span_w[best_idx] <= 10 ** ((18 - 30) / 10)
