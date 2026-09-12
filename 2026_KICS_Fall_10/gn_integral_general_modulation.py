"""Modulation/system-performance layer for ``gn_integral_general.py``.

Adds:
- modulation metadata (BPSK/QPSK/8QAM/16QAM/32QAM/64QAM/256QAM)
- EDFA ASE accumulation
- NLI power in a selected CUT
- SNR_ASE, SNR_NLI, optional transceiver SNR, and GSNR
- approximate AWGN BER by modulation
- launch-power-vs-GSNR sweep
- gross modulation rate and Shannon-gap capacity estimates

Important
---------
MODULATION_PHI now means the signed Carena (2014) Eq. (6) coefficient
Phi=mu4-2, computed directly from explicitly defined equiprobable constellations.
It is NOT Channel.modulation_phi (a legacy SCI multiplier, default 1).
ModulationFormat.phi retains that legacy multiplier meaning; its default is 1.

Default nli_model='gn' preserves the original pure-GN behavior. Select
nli_model='egn_sci' to pass both mu4 and mu6 to the engine, which computes the
link-dependent SCI ratio using Carena Eqs. (5)-(9), Appendix C. Rectangular
spectra/coherent accumulation only. XCI/MCI remain GN, not full WDM EGN.
"""
from __future__ import annotations

from dataclasses import dataclass, replace
from typing import Iterable, Literal, Optional, Sequence
import math
import numpy as np
from scipy.special import erfc

from gn_integral_general import (
    Channel, WDMSystem, Span, GNIntegralOptions,
    integrate_nli_over_channel, dbm_to_w, w_to_dbm,
)

H_PLANCK = 6.62607015e-34
C_M_PER_S = 299792458.0


@dataclass(frozen=True)
class ModulationFormat:
    name: str
    M: int
    coding_rate: float = 1.0
    # Legacy engine SCI multiplier, NOT the MODULATION_PHI table coefficient.
    phi: float = 1.0
    mu4: float = 2.0
    mu6: float = 6.0

    @property
    def egn_phi(self) -> float:
        return self.mu4 - 2.0

    @property
    def egn_psi(self) -> float:
        return self.mu6 - 9.0 * self.mu4 + 12.0

    @property
    def bits_per_symbol(self) -> float:
        return math.log2(self.M)

    def gross_dp_rate_gbps(self, baud_GBd: float) -> float:
        """Dual-polarization gross line rate, before FEC."""
        return 2.0 * float(baud_GBd) * self.bits_per_symbol

    def net_dp_rate_gbps(self, baud_GBd: float) -> float:
        return self.gross_dp_rate_gbps(baud_GBd) * float(self.coding_rate)


MODULATIONS = {
    "BPSK": 2,
    "QPSK": 4,
    "8QAM": 8,
    "16QAM": 16,
    "32QAM": 32,
    "64QAM": 64,
    "256QAM": 256,
}


def _modulation_key(name: str) -> str:
    key = str(name).upper().replace("-", "")
    aliases = {"PMQPSK": "QPSK", "PM16QAM": "16QAM", "PM64QAM": "64QAM"}
    key = aliases.get(key, key)
    if key not in MODULATIONS:
        raise ValueError(f"unsupported modulation {name!r}; choose {list(MODULATIONS)}")
    return key


def constellation_symbols(name: str) -> np.ndarray:
    """Explicit equal-probability coordinates, scaled to E[|X|**2]=1.

    BPSK: +/-1. QPSK/square QAM: I,Q in the equally spaced odd-integer grid.
    8QAM: chosen cross constellation, {+/-1+/-j} plus four cardinal points
    at radius 1+sqrt(3). 32QAM: 6x6 odd-integer grid with four corners removed.
    These 8/32QAM geometries are explicit modeling choices, not universal
    definitions or claims that these constellations were used in either paper.
    """
    key = _modulation_key(name)
    if key == "BPSK":
        symbols = np.array([-1, 1], dtype=complex)
    elif key == "8QAM":
        radius = 1.0 + math.sqrt(3.0)
        symbols = np.array([1+1j, 1-1j, -1+1j, -1-1j,
                            radius, -radius, 1j*radius, -1j*radius], dtype=complex)
    elif key == "32QAM":
        levels = (-5, -3, -1, 1, 3, 5)
        symbols = np.array([i + 1j*q for i in levels for q in levels
                            if not (abs(i) == 5 and abs(q) == 5)], dtype=complex)
    else:
        side = math.isqrt(MODULATIONS[key])
        levels = np.arange(1 - side, side, 2)
        symbols = (levels[:, None] + 1j * levels[None, :]).ravel()
    return symbols / math.sqrt(float(np.mean(np.abs(symbols)**2)))


@dataclass(frozen=True)
class ConstellationMoments:
    mu4: float
    mu6: float
    egn_phi: float
    egn_psi: float
    normalized_mean_power: float
    pseudo_variance_abs: float


def constellation_moments(
    symbols: Sequence[complex],
    probabilities: Optional[Sequence[float]] = None,
) -> ConstellationMoments:
    """Compute standardized radial moments from actual symbols, not a table.

    Optional probabilities permit an explicitly supplied shaped constellation.
    The built-in formats all use uniform probabilities. Radial moments alone
    are not a general model of noncircular or correlated-polarization signals;
    pseudo_variance_abs reports |E[X**2]| for awareness (BPSK gives 1).
    """
    x = np.asarray(symbols, dtype=complex)
    if x.ndim != 1 or x.size == 0 or not np.all(np.isfinite(x)):
        raise ValueError("symbols must be a finite nonempty one-dimensional constellation")
    if probabilities is None:
        p = np.full(x.size, 1.0 / x.size)
    else:
        p = np.asarray(probabilities, dtype=float)
        if p.shape != x.shape or not np.all(np.isfinite(p)) or np.any(p < 0) or p.sum() <= 0:
            raise ValueError("probabilities must be finite, nonnegative and match symbols")
        p = p / p.sum()
    power = float(np.sum(p * np.abs(x)**2))
    if power <= 0 or not math.isfinite(power):
        raise ValueError("constellation must have finite positive average power")
    xn = x / math.sqrt(power)
    e2 = float(np.sum(p * np.abs(xn)**2))
    mu4 = float(np.sum(p * np.abs(xn)**4) / e2**2)
    mu6 = float(np.sum(p * np.abs(xn)**6) / e2**3)
    return ConstellationMoments(mu4, mu6, mu4 - 2.0,
                                mu6 - 9.0 * mu4 + 12.0, e2,
                                float(abs(np.sum(p * xn**2))))


MODULATION_MOMENTS = {
    name: constellation_moments(constellation_symbols(name)) for name in MODULATIONS
}

# DIRECTLY CALCULATED constellation 4th moment, using constellation_symbols()
# and constellation_moments(): Phi = E[|X|^4]/E[|X|^2]^2 - 2.
# Definition: Carena, Bosco, Curri, Jiang, Poggiolini, Forghieri,
# "EGN model of non-linear fiber propagation", Opt. Express 22 (2014),
# DOI 10.1364/OE.22.016335, Eq. (6); BPSK/QPSK/16QAM/64QAM cross-check Table 1.
# Dar et al., "Properties of nonlinear noise in long, dispersion-uncompensated
# fiber links", Opt. Express 21 (2013), DOI 10.1364/OE.21.025685, Eq. (25):
# the same normalized correction enters as (mu4-2)*chi2, NOT an SCI multiplier.
# BPSK/QPSK: +/-1 / square coordinates; 16/64/256QAM: odd-integer square grids.
# 8QAM and 32QAM: the explicit cross geometries documented above. Their moments
# must be recalculated if a different geometry or symbol probability is used.
# This is a deliberate semantic correction of the old, ungrounded dictionary:
# negative values here MUST NOT be assigned to Channel.modulation_phi.
# No approximate or fitted numeric multipliers remain in this table.
MODULATION_PHI = {name: moments.egn_phi for name, moments in MODULATION_MOMENTS.items()}


def get_modulation(name: str, coding_rate: float = 1.0) -> ModulationFormat:
    key = _modulation_key(name)
    moments = MODULATION_MOMENTS[key]
    return ModulationFormat(key, MODULATIONS[key], float(coding_rate), phi=1.0,
                            mu4=moments.mu4, mu6=moments.mu6)


def _channel_for_nli(ch: Channel, mod: ModulationFormat, nli_model: str) -> Channel:
    """Input adapter only; all NLI integration and moment correction is in the engine."""
    if nli_model == "gn":
        return replace(ch, modulation_phi=mod.phi, egn_mu4=None, egn_mu6=None)
    if nli_model == "egn_sci":
        return replace(ch, modulation_phi=mod.phi, egn_mu4=mod.mu4, egn_mu6=mod.mu6)
    raise ValueError("nli_model must be 'gn' or 'egn_sci'")


def ber_awgn_approx(gsnr_linear: float, modulation: str) -> float:
    """Approximate uncoded BER vs GSNR.

    - QPSK reduces to 0.5*erfc(sqrt(SNR/2)), matching the paper's PM-QPSK
      convention.
    - Square M-QAM uses the standard Gray-coded AWGN approximation.
    - 8QAM/32QAM are treated by the same generic M-QAM approximation and
      should be considered approximate because constellation geometry varies.
    """
    s = max(float(gsnr_linear), 0.0)
    mod = get_modulation(modulation)
    M = mod.M
    if M == 2:
        return float(0.5 * erfc(math.sqrt(s)))
    k = math.log2(M)
    q = 0.5 * erfc(math.sqrt(3.0 * s / (2.0 * (M - 1.0))))
    ber = (4.0 / k) * (1.0 - 1.0 / math.sqrt(M)) * q
    return float(min(max(ber, 0.0), 0.5))


def _span_output_power_transfer(span: Span) -> float:
    return float(span.total_field_transfer) ** 2


def link_output_signal_power(system: WDMSystem, spans: Sequence[Span], cut_index: int) -> float:
    p = float(system.channels[int(cut_index)].power_W)
    for sp in spans:
        p *= _span_output_power_transfer(sp)
    return p


def ase_noise_power_edfa(
    spans: Sequence[Span],
    receiver_bandwidth_Hz: float,
    wavelength_nm: float = 1550.0,
) -> float:
    """Accumulate lumped-EDFA ASE at the link output.

    Uses the source-paper convention G_ASE=(G-1) F h nu for dual-pol
    unilateral ASE PSD, then propagates each amplifier's ASE through all
    downstream spans. Distributed-Raman ASE is not included by this helper.
    """
    spans = tuple(spans)
    nu = C_M_PER_S / (float(wavelength_nm) * 1e-9)
    total_psd_W_per_Hz = 0.0
    downstream = 1.0
    # Work backwards. Noise added at EDFA n sees downstream spans only.
    for n in range(len(spans) - 1, -1, -1):
        sp = spans[n]
        G = float(sp.lumped_power_gain)
        if G > 1.0 + 1e-15:
            F = 10.0 ** (float(sp.noise_figure_db) / 10.0)
            g_ase = (G - 1.0) * F * H_PLANCK * nu
            total_psd_W_per_Hz += g_ase * downstream
        downstream *= _span_output_power_transfer(sp)
    return float(total_psd_W_per_Hz * float(receiver_bandwidth_Hz))


def _db(x: float) -> float:
    return 10.0 * math.log10(float(x)) if x > 0 else -math.inf


@dataclass(frozen=True)
class PerformanceResult:
    launch_power_dBm: float
    cut_output_power_W: float
    p_ase_W: float
    p_nli_W: float
    snr_ase_db: float
    snr_nli_db: float
    gsnr_db: float
    ber: float
    gross_rate_gbps: float
    net_rate_gbps: float
    shannon_gap_capacity_gbps: float


def evaluate_performance(
    system: WDMSystem,
    spans: Sequence[Span],
    cut_index: int,
    modulation: str = "QPSK",
    coding_rate: float = 1.0,
    trx_snr_db: Optional[float] = None,
    shannon_gap_db: float = 0.0,
    gn_options: GNIntegralOptions = GNIntegralOptions(),
    receiver_points: int = 7,
    nli_model: Literal["gn", "egn_sci"] = "gn",
) -> PerformanceResult:
    """Compute NLI/ASE/GSNR/BER. Opt in to rectangular SCI EGN with nli_model='egn_sci'."""
    ch = system.channels[int(cut_index)]
    p_sig = link_output_signal_power(system, spans, cut_index)
    mod = get_modulation(modulation, coding_rate)
    ch = _channel_for_nli(ch, mod, nli_model)
    channels = list(system.channels)
    channels[int(cut_index)] = ch
    sys_for_nli = replace(system, channels=tuple(channels))
    p_nli = integrate_nli_over_channel(sys_for_nli, spans, cut_index, gn_options, receiver_points)
    p_ase = ase_noise_power_edfa(spans, ch.baud_GBd * 1e9, ch_center_wavelength_nm(spans, ch))

    noise = p_ase + p_nli
    if trx_snr_db is not None:
        noise += p_sig / (10.0 ** (float(trx_snr_db) / 10.0))

    snr_ase = math.inf if p_ase == 0 else p_sig / p_ase
    snr_nli = math.inf if p_nli == 0 else p_sig / p_nli
    gsnr = math.inf if noise == 0 else p_sig / noise

    ber = 0.0 if math.isinf(gsnr) else ber_awgn_approx(gsnr, mod.name)
    gross = mod.gross_dp_rate_gbps(ch.baud_GBd)
    net = mod.net_dp_rate_gbps(ch.baud_GBd)
    gap = 10.0 ** (float(shannon_gap_db) / 10.0)
    shannon = math.inf if math.isinf(gsnr) else 2.0 * ch.baud_GBd * math.log2(1.0 + gsnr / gap)

    return PerformanceResult(
        launch_power_dBm=w_to_dbm(ch.power_W),
        cut_output_power_W=p_sig,
        p_ase_W=p_ase,
        p_nli_W=p_nli,
        snr_ase_db=math.inf if math.isinf(snr_ase) else _db(snr_ase),
        snr_nli_db=math.inf if math.isinf(snr_nli) else _db(snr_nli),
        gsnr_db=math.inf if math.isinf(gsnr) else _db(gsnr),
        ber=ber,
        gross_rate_gbps=gross,
        net_rate_gbps=net,
        shannon_gap_capacity_gbps=shannon,
    )


def ch_center_wavelength_nm(spans: Sequence[Span], ch: Channel) -> float:
    """Use the fiber reference wavelength for ASE; relative WDM offsets are small."""
    return float(spans[0].wavelength_nm) if spans else 1550.0


def launch_power_vs_gsnr(
    base_system: WDMSystem,
    spans: Sequence[Span],
    cut_index: int,
    launch_power_dBm: Iterable[float],
    modulation: str = "QPSK",
    coding_rate: float = 1.0,
    trx_snr_db: Optional[float] = None,
    shannon_gap_db: float = 0.0,
    gn_options: GNIntegralOptions = GNIntegralOptions(),
    receiver_points: int = 7,
    use_cubic_scaling: bool = True,
    nli_model: Literal["gn", "egn_sci"] = "gn",
) -> dict:
    """Return launch-power-vs-GSNR (plus ASE/NLI/BER/rate arrays).

    By default one expensive GN integral is evaluated at ``base_system`` and
    all channel powers are then shifted together. The exact GN-model cubic
    law P_NLI -> a^3 P_NLI is used, making dense power sweeps fast.

    Set use_cubic_scaling=False to recompute the full GN integral at every
    power point (useful as a numerical cross-check, but much slower).
    nli_model='egn_sci' selects link-dependent moment correction of SCI only.
    Its ratio is power-independent, so the existing cubic scaling still holds.
    """
    pgrid = np.asarray(list(launch_power_dBm), dtype=float)
    if pgrid.size == 0:
        raise ValueError("launch_power_dBm is empty")
    cut = base_system.channels[int(cut_index)]
    p0 = float(cut.power_W)
    p0_dbm = w_to_dbm(p0)

    # Reference output signal and noises.
    p_sig0 = link_output_signal_power(base_system, spans, cut_index)
    p_ase = ase_noise_power_edfa(spans, cut.baud_GBd * 1e9, ch_center_wavelength_nm(spans, cut))
    mod = get_modulation(modulation, coding_rate)
    channels = list(base_system.channels)
    channels[int(cut_index)] = _channel_for_nli(cut, mod, nli_model)
    sys_for_nli = replace(base_system, channels=tuple(channels))
    p_nli0 = integrate_nli_over_channel(sys_for_nli, spans, cut_index, gn_options, receiver_points)

    gap = 10.0 ** (float(shannon_gap_db) / 10.0)

    out = {
        "launch_power_dBm": [], "gsnr_db": [], "snr_ase_db": [], "snr_nli_db": [],
        "ber": [], "p_nli_W": [], "p_ase_W": [], "cut_output_power_W": [],
        "gross_rate_gbps": [], "net_rate_gbps": [], "shannon_gap_capacity_gbps": [],
    }

    for p_dbm in pgrid:
        scale = 10.0 ** ((float(p_dbm) - p0_dbm) / 10.0)
        if use_cubic_scaling:
            p_sig = p_sig0 * scale
            p_nli = p_nli0 * scale**3
        else:
            sys_i = sys_for_nli.with_common_power_offset_db(float(p_dbm) - p0_dbm)
            p_sig = link_output_signal_power(sys_i, spans, cut_index)
            p_nli = integrate_nli_over_channel(sys_i, spans, cut_index, gn_options, receiver_points)

        noise = p_ase + p_nli
        if trx_snr_db is not None:
            noise += p_sig / (10.0 ** (float(trx_snr_db) / 10.0))
        gsnr = math.inf if noise == 0 else p_sig / noise
        snr_ase = math.inf if p_ase == 0 else p_sig / p_ase
        snr_nli = math.inf if p_nli == 0 else p_sig / p_nli
        ber = 0.0 if math.isinf(gsnr) else ber_awgn_approx(gsnr, mod.name)
        shannon = math.inf if math.isinf(gsnr) else 2.0 * cut.baud_GBd * math.log2(1.0 + gsnr / gap)

        out["launch_power_dBm"].append(float(p_dbm))
        out["gsnr_db"].append(math.inf if math.isinf(gsnr) else _db(gsnr))
        out["snr_ase_db"].append(math.inf if math.isinf(snr_ase) else _db(snr_ase))
        out["snr_nli_db"].append(math.inf if math.isinf(snr_nli) else _db(snr_nli))
        out["ber"].append(float(ber))
        out["p_nli_W"].append(float(p_nli))
        out["p_ase_W"].append(float(p_ase))
        out["cut_output_power_W"].append(float(p_sig))
        out["gross_rate_gbps"].append(mod.gross_dp_rate_gbps(cut.baud_GBd))
        out["net_rate_gbps"].append(mod.net_dp_rate_gbps(cut.baud_GBd))
        out["shannon_gap_capacity_gbps"].append(float(shannon))

    return {k: np.asarray(v, dtype=float) for k, v in out.items()}


# Backward-compatible descriptive alias used in earlier project notes.
def simulate_snr_curve(*args, **kwargs):
    return launch_power_vs_gsnr(*args, **kwargs)


def constellation_moment_self_test() -> dict:
    """Check coordinates/normalization and Carena Eq. (6), Table 1 (not fitted)."""
    table1_phi = {"BPSK": -1.0, "QPSK": -1.0, "16QAM": -17.0/25.0,
                  "64QAM": -13.0/21.0}
    rows = {}
    for name, count in MODULATIONS.items():
        x = constellation_symbols(name)
        assert len(x) == count and len(np.unique(x)) == count
        moments = constellation_moments(x)
        assert abs(moments.normalized_mean_power - 1.0) < 1e-12
        assert abs(np.mean(x)) < 1e-12
        scaled = constellation_moments(x * (3.0 + 2.0j))
        assert np.isclose(moments.mu4, scaled.mu4, rtol=1e-12, atol=0.0)
        assert np.isclose(moments.mu6, scaled.mu6, rtol=1e-12, atol=0.0)
        if name in table1_phi:
            assert abs(moments.egn_phi - table1_phi[name]) < 1e-12
        rows[name] = {"mu4": moments.mu4, "mu6": moments.mu6,
                      "Phi": moments.egn_phi, "Psi": moments.egn_psi}
    # Paper prints 1161/646 for 64QAM Psi; direct exact grid moment is
    # 5548/3087. Difference is ~5e-7; keep the exact calculation, not rounding.
    psi64 = MODULATION_MOMENTS["64QAM"].egn_psi
    assert abs(psi64 - 5548.0/3087.0) < 1e-12
    assert abs(psi64 - 1161.0/646.0) < 1e-6
    return {"passed": True, "source": "Carena 2014 Eq. (6), Table 1", "rows": rows}


__all__ = [
    "ModulationFormat", "MODULATIONS", "MODULATION_PHI", "get_modulation", "ber_awgn_approx",
    "ase_noise_power_edfa", "link_output_signal_power", "PerformanceResult",
    "evaluate_performance", "launch_power_vs_gsnr", "simulate_snr_curve",
    "ConstellationMoments", "MODULATION_MOMENTS", "constellation_symbols",
    "constellation_moments", "constellation_moment_self_test",
]
