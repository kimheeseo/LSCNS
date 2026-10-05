"""EGN_adaptive.py: phase-aware, convergence-checked Carena (2014) EGN.

한글: final_EGN(1).py의 GN/변조 자료형을 유지하고 Full-EGN 적분 경로를 교체했습니다.
English: preserves GN/modulation data types and replaces the Full-EGN integrator.

The precision Full-EGN path supports rectangular equal-baud spectra, identical
loss-compensated EDFA spans, beta2 dispersion, and the central symmetric WDM
comb required by Appendix B. SCI and XCI also work for one/two channels.
Beta3, distributed gain and heterogeneous spans remain available through the
legacy GN API; the new precision EGN path rejects those unvalidated profiles.

Changes: exact linear-frequency inner primitives (E1/Ei), exact receiver
integration for the GN term, support-boundary splitting, stationary-point and
phase-based splitting for fixed-sum terms, adaptive outer quadrature, explicit
positive/negative MCI regions, and separate frequency/receiver convergence checks.

No scale factors or literature target values enter the physics calculations.
A numerical convergence pass is NOT a guarantee of <3% error against paper
curves or SSFM. See the accompanying validation report for measured errors.
Dependencies: Python >=3.10, numpy, scipy. This file is self-contained.
Reference: Carena et al., Opt. Express 22, 16335-16362 (2014),
https://doi.org/10.1364/OE.22.016335, Eqs. 5-12, 18, Appendices A-C.
"""
from __future__ import annotations

from dataclasses import dataclass, field, replace
from functools import lru_cache
from typing import Callable, Iterable, Literal, Optional, Sequence
import math
import numpy as np
from scipy.stats import qmc
from scipy.special import roots_legendre, erfc

C_NM_PER_PS = 299792.458
DP_GN_COEFF = 16.0 / 27.0

PulseShape = Literal["rect", "raised_cosine", "rrc", "custom"]
AccumulationMode = Literal["coherent", "incoherent"]
AmplificationMode = Literal["lumped_edfa", "ideal_distributed", "backward_raman", "custom"]


def dbm_to_w(dbm: float) -> float:
    return 1e-3 * 10.0 ** (float(dbm) / 10.0)


def w_to_dbm(w: float) -> float:
    return 10.0 * math.log10(float(w) / 1e-3)


def alpha_field_from_db(alpha_db_per_km: float) -> float:
    """Power attenuation [dB/km] -> field attenuation [1/km]."""
    return float(alpha_db_per_km) * math.log(10.0) / 20.0


def beta2_from_D(D_ps_nm_km: float, wavelength_nm: float = 1550.0) -> float:
    """Convert D [ps/(nm km)] to beta2 [ps^2/km]."""
    return -float(D_ps_nm_km) * float(wavelength_nm) ** 2 / (2.0 * math.pi * C_NM_PER_PS)


def beta_of_f_offset(f_THz, beta2_ps2_km: float, beta3_ps3_km: float = 0.0):
    """Propagation-constant offset beta(f), Appendix-G Eq. (G.1).

    Constant and first-order terms are omitted because they cancel in the FWM
    phase mismatch. f_THz is numerically 1/ps.
    """
    f = np.asarray(f_THz, dtype=float)
    return 2.0 * np.pi**2 * beta2_ps2_km * f**2 + (4.0 / 3.0) * np.pi**3 * beta3_ps3_km * f**3


def phase_mismatch_beta23(f1_THz, f2_THz, f_THz, beta2_ps2_km: float, beta3_ps3_km: float = 0.0):
    """Return q [1/km] using the Eq.-(G.2) sign convention.

    q = 4*pi^2*(f1-f)*(f2-f) * [beta2 + pi*beta3*(f1+f2)]

    This reduces exactly to the beta2-only convention used by the original
    gn_integral.py when beta3=0.
    """
    f1 = np.asarray(f1_THz, dtype=float)
    f2 = np.asarray(f2_THz, dtype=float)
    f = np.asarray(f_THz, dtype=float)
    return 4.0 * np.pi**2 * (f1 - f) * (f2 - f) * (
        float(beta2_ps2_km) + np.pi * float(beta3_ps3_km) * (f1 + f2)
    )


def _safe_complex_span_integral(q, alpha_field_per_km: float, span_km: float):
    """Integral int_0^L exp(-2 alpha z) exp(j q z) dz."""
    q = np.asarray(q, dtype=float)
    a = float(alpha_field_per_km)
    L = float(span_km)
    den = 2.0 * a - 1j * q
    num = 1.0 - np.exp((-2.0 * a + 1j * q) * L)
    small = np.abs(den) < 1e-14
    if np.ndim(q) == 0:
        return complex(L if bool(small) else num / den)
    out = np.empty(q.shape, dtype=complex)
    out[small] = L
    out[~small] = num[~small] / den[~small]
    return out


def _normalised_rc_psd(offset_THz, symbol_rate_THz: float, rolloff: float):
    """Unit-area raised-cosine *power* spectrum [1/THz].

    The power spectrum of an RRC pulse is raised cosine, so pulse_shape='rrc'
    intentionally uses this same PSD.
    """
    x = np.abs(np.asarray(offset_THz, dtype=float))
    Rs = float(symbol_rate_THz)
    r = float(rolloff)
    if not (0.0 <= r <= 1.0):
        raise ValueError("rolloff must be in [0, 1]")
    if r == 0.0:
        return np.where(x <= Rs / 2.0, 1.0 / Rs, 0.0)
    a = (1.0 - r) * Rs / 2.0
    b = (1.0 + r) * Rs / 2.0
    out = np.zeros_like(x, dtype=float)
    out[x <= a] = 1.0 / Rs
    mask = (x > a) & (x <= b)
    out[mask] = 0.5 / Rs * (1.0 + np.cos(np.pi * (x[mask] - a) / (r * Rs)))
    return out


@dataclass(frozen=True)
class Channel:
    """One WDM channel.

    Parameters
    ----------
    center_THz:
        Center frequency relative to an arbitrary optical reference.
    power_W:
        Total dual-polarization launch power.
    baud_GBd:
        Symbol rate.
    pulse_shape:
        'rect', 'raised_cosine', 'rrc', or 'custom'.
    rolloff:
        Used by raised-cosine/RRC power PSD.
    custom_psd:
        Callable receiving offset frequency [THz] and returning a *unit-area*
        PSD shape [1/THz].
    custom_support_half_width_THz:
        Required for custom PSDs so QMC bounds remain finite.
    modulation_phi:
        Legacy SCI-only multiplier; 1.0 preserves pure GN. NOT the signed
        excess kurtosis Phi=mu4-2 used in Carena (2014), Eq. (6).
    egn_mu4, egn_mu6:
        Optional standardized symbol moments. Both must be provided to use
        link-dependent EGN SCI, with modulation_phi left at 1.0. Rectangular
        spectra and coherent accumulation only; other channels stay GN.
    """
    center_THz: float
    power_W: float
    baud_GBd: float
    pulse_shape: PulseShape = "rect"
    rolloff: float = 0.0
    label: str = ""
    custom_psd: Optional[Callable[[np.ndarray], np.ndarray]] = field(default=None, compare=False, repr=False)
    custom_support_half_width_THz: Optional[float] = None
    # Legacy SCI multiplier, Phi=1이면 GN과 동일; NOT EGN excess kurtosis.
    modulation_phi: float = 1.0
    egn_mu4: Optional[float] = None
    egn_mu6: Optional[float] = None

    @property
    def symbol_rate_THz(self) -> float:
        return float(self.baud_GBd) / 1000.0

    @property
    def support_half_width_THz(self) -> float:
        if self.pulse_shape == "rect":
            return self.symbol_rate_THz / 2.0
        if self.pulse_shape in ("raised_cosine", "rrc"):
            return (1.0 + float(self.rolloff)) * self.symbol_rate_THz / 2.0
        if self.pulse_shape == "custom":
            if self.custom_support_half_width_THz is None:
                raise ValueError("custom PSD requires custom_support_half_width_THz")
            return float(self.custom_support_half_width_THz)
        raise ValueError(f"unsupported pulse_shape={self.pulse_shape!r}")

    def psd(self, f_THz):
        off = np.asarray(f_THz, dtype=float) - float(self.center_THz)
        if self.pulse_shape == "rect":
            shape = np.where(np.abs(off) <= self.symbol_rate_THz / 2.0, 1.0 / self.symbol_rate_THz, 0.0)
        elif self.pulse_shape in ("raised_cosine", "rrc"):
            shape = _normalised_rc_psd(off, self.symbol_rate_THz, self.rolloff)
        elif self.pulse_shape == "custom":
            if self.custom_psd is None:
                raise ValueError("custom pulse_shape requires custom_psd callable")
            shape = np.asarray(self.custom_psd(off), dtype=float)
            shape = np.where(np.abs(off) <= self.support_half_width_THz, shape, 0.0)
        else:
            raise ValueError(f"unsupported pulse_shape={self.pulse_shape!r}")
        return float(self.power_W) * shape


@dataclass(frozen=True)
class WDMSystem:
    channels: tuple[Channel, ...]

    def __post_init__(self):
        if len(self.channels) == 0:
            raise ValueError("at least one channel is required")

    @classmethod
    def equispaced(
        cls,
        n_channels: int,
        spacing_GHz: float,
        baud_GBd: float,
        power_dBm: float,
        pulse_shape: PulseShape = "rect",
        rolloff: float = 0.0,
    ) -> "WDMSystem":
        n = int(n_channels)
        centers_GHz = (np.arange(n) - (n - 1) / 2.0) * float(spacing_GHz)
        p = dbm_to_w(power_dBm)
        chs = tuple(
            Channel(
                center_THz=float(fc / 1000.0), power_W=p, baud_GBd=float(baud_GBd),
                pulse_shape=pulse_shape, rolloff=float(rolloff), label=f"ch{k}"
            )
            for k, fc in enumerate(centers_GHz)
        )
        return cls(chs)

    def psd(self, f_THz):
        f = np.asarray(f_THz, dtype=float)
        out = np.zeros_like(f, dtype=float)
        for ch in self.channels:
            out += ch.psd(f)
        return out

    def support_bounds_THz(self) -> tuple[float, float]:
        lo = min(ch.center_THz - ch.support_half_width_THz for ch in self.channels)
        hi = max(ch.center_THz + ch.support_half_width_THz for ch in self.channels)
        return float(lo), float(hi)

    def with_common_power_offset_db(self, offset_db: float) -> "WDMSystem":
        s = 10.0 ** (float(offset_db) / 10.0)
        return WDMSystem(tuple(replace(ch, power_W=ch.power_W * s) for ch in self.channels))


@dataclass(frozen=True)
class Span:
    """One fiber span.

    ``gain_db`` is the lumped *power* gain at the span end.  If None and the
    mode is 'lumped_edfa', exact span-loss compensation is assumed.

    For backward Raman, Eq. (105) is represented by
        2*g_hat(z) = C_R * P_p0 * exp(2*alpha_p*z)
    where alpha_p is the *field* pump attenuation coefficient.
    """
    length_km: float
    alpha_db_per_km: float
    gamma_W_inv_km: float
    D_ps_nm_km: Optional[float] = None
    wavelength_nm: float = 1550.0
    beta2_ps2_km: Optional[float] = None
    beta3_ps3_km: float = 0.0
    gain_db: Optional[float] = None
    noise_figure_db: float = 5.0
    dcu_ps2: float = 0.0
    amplification: AmplificationMode = "lumped_edfa"
    raman_C_R_per_W_km: float = 0.0
    raman_pump_W: float = 0.0
    raman_pump_alpha_field_per_km: float = 0.0
    custom_field_gain_coeff: Optional[Callable[[np.ndarray], np.ndarray]] = field(default=None, compare=False, repr=False)

    @property
    def alpha_field_per_km(self) -> float:
        return alpha_field_from_db(self.alpha_db_per_km)

    @property
    def beta2(self) -> float:
        if self.beta2_ps2_km is not None:
            return float(self.beta2_ps2_km)
        if self.D_ps_nm_km is None:
            raise ValueError("provide either beta2_ps2_km or D_ps_nm_km")
        return beta2_from_D(self.D_ps_nm_km, self.wavelength_nm)

    @property
    def lumped_power_gain(self) -> float:
        if self.gain_db is not None:
            return 10.0 ** (float(self.gain_db) / 10.0)
        if self.amplification == "lumped_edfa":
            # Exact compensation of conventional power loss alpha_dB*L.
            return 10.0 ** (float(self.alpha_db_per_km) * float(self.length_km) / 10.0)
        return 1.0

    def field_gain_coeff(self, z_km):
        z = np.asarray(z_km, dtype=float)
        if self.amplification == "lumped_edfa":
            return np.zeros_like(z)
        if self.amplification == "ideal_distributed":
            return np.full_like(z, self.alpha_field_per_km)
        if self.amplification == "backward_raman":
            ap = float(self.raman_pump_alpha_field_per_km)
            return 0.5 * float(self.raman_C_R_per_W_km) * float(self.raman_pump_W) * np.exp(2.0 * ap * z)
        if self.amplification == "custom":
            if self.custom_field_gain_coeff is None:
                raise ValueError("custom amplification requires custom_field_gain_coeff")
            return np.asarray(self.custom_field_gain_coeff(z), dtype=float)
        raise ValueError(f"unsupported amplification={self.amplification!r}")

    def distributed_field_log_gain(self, z_km):
        """Return int_0^z g_hat(xi) dxi for common profiles."""
        z = np.asarray(z_km, dtype=float)
        if self.amplification == "lumped_edfa":
            return np.zeros_like(z)
        if self.amplification == "ideal_distributed":
            return self.alpha_field_per_km * z
        if self.amplification == "backward_raman":
            ap = float(self.raman_pump_alpha_field_per_km)
            A = 0.5 * float(self.raman_C_R_per_W_km) * float(self.raman_pump_W)
            if abs(ap) < 1e-15:
                return A * z
            return A * (np.exp(2.0 * ap * z) - 1.0) / (2.0 * ap)
        # Custom: numerical cumulative integral, scalar/vector friendly.
        flat = np.atleast_1d(z)
        vals = []
        for zz in flat:
            if zz <= 0:
                vals.append(0.0)
                continue
            grid = np.linspace(0.0, float(zz), 257)
            vals.append(float(np.trapz(self.field_gain_coeff(grid), grid)))
        arr = np.asarray(vals).reshape(np.shape(z) if np.ndim(z) else (1,))
        return float(arr[0]) if np.ndim(z) == 0 else arr

    @property
    def total_field_transfer(self) -> float:
        """Linear field transfer through fiber + distributed gain + lumped gain."""
        L = float(self.length_km)
        log_g = float(self.distributed_field_log_gain(L))
        return math.exp(-self.alpha_field_per_km * L + log_g) * math.sqrt(self.lumped_power_gain)


@dataclass(frozen=True)
class GNIntegralOptions:
    sobol_power: int = 18
    seed: int = 1
    accumulation: AccumulationMode = "coherent"
    z_quadrature_order: int = 64
    # Per-segment nested quadrature for optional EGN SCI; does not affect GN.
    egn_frequency_order: int = 512


@dataclass(frozen=True)
class GNResult:
    f_THz: float
    g_nli_W_per_THz: float
    qmc_samples: int
    integration_bounds_THz: tuple[float, float]
    accumulation: str


def _local_span_source_integral(span: Span, q, z_order: int = 64):
    """Complex single-span source integral.

    Lumped-EDFA case is analytic. Distributed cases evaluate the Eq.-(103)
    profile integral numerically:
        int exp(2*int_0^z g_hat dxi) exp(-2 alpha z) exp(j q z) dz.
    """
    q = np.asarray(q, dtype=float)
    if span.amplification == "lumped_edfa":
        return _safe_complex_span_integral(q, span.alpha_field_per_km, span.length_km)

    # Gauss-Legendre z integration; vectorised in q chunks by broadcasting.
    x, w = np.polynomial.legendre.leggauss(int(z_order))
    L = float(span.length_km)
    z = 0.5 * (x + 1.0) * L
    wz = 0.5 * L * w
    logg = span.distributed_field_log_gain(z)
    envelope = np.exp(2.0 * logg - 2.0 * span.alpha_field_per_km * z)
    if np.ndim(q) == 0:
        return complex(np.sum(wz * envelope * np.exp(1j * float(q) * z)))
    # (Nq, Nz) can be large, so chunk.
    flat = q.ravel()
    out = np.empty(flat.shape, dtype=complex)
    chunk = 32768
    for i in range(0, len(flat), chunk):
        qq = flat[i:i+chunk, None]
        out[i:i+chunk] = np.sum((wz * envelope)[None, :] * np.exp(1j * qq * z[None, :]), axis=1)
    return out.reshape(q.shape)


def _link_amplitudes(f1, f2, f, spans: Sequence[Span], z_order: int = 64):
    """Per-span complex FWM amplitudes following the structure of Eq. (100)."""
    spans = tuple(spans)
    N = len(spans)
    if N == 0:
        raise ValueError("at least one span is required")

    T = np.asarray([sp.total_field_transfer for sp in spans], dtype=float)
    pre = np.ones(N, dtype=float)
    for n in range(1, N):
        pre[n] = pre[n-1] * T[n-1]
    post = np.ones(N, dtype=float)
    running = 1.0
    for n in range(N-1, -1, -1):
        running *= T[n]
        post[n] = running

    q_list = [phase_mismatch_beta23(f1, f2, f, sp.beta2, sp.beta3_ps3_km) for sp in spans]

    # Prefix phase from all previous spans (fiber + optional lumped DCU).
    prefix_phase = np.zeros_like(np.asarray(f1, dtype=float), dtype=float)
    amps = []
    for n, sp in enumerate(spans):
        qn = q_list[n]
        local = _local_span_source_integral(sp, qn, z_order=z_order)
        scale = float(sp.gamma_W_inv_km) * pre[n]**3 * post[n]
        amps.append(scale * np.exp(1j * prefix_phase) * local)
        q_dcu = 4.0 * np.pi**2 * (np.asarray(f1)-np.asarray(f)) * (np.asarray(f2)-np.asarray(f)) * float(sp.dcu_ps2)
        prefix_phase = prefix_phase + qn * float(sp.length_km) + q_dcu
    return amps


def link_kernel(f1, f2, f, spans: Sequence[Span], accumulation: AccumulationMode = "coherent", z_order: int = 64):
    """Return |sum span amplitudes|^2 (coherent) or sum |.|^2 (incoherent)."""
    amps = _link_amplitudes(f1, f2, f, spans, z_order=z_order)
    if accumulation == "coherent":
        return np.abs(np.sum(np.stack(amps, axis=0), axis=0))**2
    if accumulation == "incoherent":
        return np.sum(np.stack([np.abs(a)**2 for a in amps], axis=0), axis=0)
    raise ValueError("accumulation must be 'coherent' or 'incoherent'")


@dataclass(frozen=True)
class EGNSCICoefficients:
    """Carena (2014) SCI coefficients before multiplication by P_CUT**3.

    kappa2a/kappa2b are the 80/81 and 16/81 terms in Eq. (8), evaluated
    using Appendix C Eqs. (43)/(44). kappa3 is Eq. (45).
    Units: 1/(W**2 THz) for PSD, or 1/W**2 after receiver integration.
    """
    kappa1: float
    kappa2a: float
    kappa2b: float
    kappa3: float

    @property
    def kappa2(self) -> float:
        return self.kappa2a + self.kappa2b

    def corrected(self, mu4: float, mu6: float) -> float:
        """Eq. (5): kappa1 + (mu4-2)*kappa2 + (mu6-9*mu4+12)*kappa3."""
        _validate_egn_moments(mu4, mu6)
        return (self.kappa1 + (float(mu4) - 2.0) * self.kappa2
                + (float(mu6) - 9.0 * float(mu4) + 12.0) * self.kappa3)

    def sci_multiplier(self, mu4: float, mu6: float) -> float:
        """Link-dependent SCI/GN ratio; never substitute mu4 or mu4/2 here."""
        value = self.corrected(mu4, mu6)
        if not math.isfinite(value) or value < 0.0:
            raise ValueError("negative/nonfinite EGN SCI: check moments and quadrature convergence")
        if self.kappa1 == 0.0:
            if value != 0.0:
                raise ValueError("nonzero SCI correction with zero GN SCI")
            return 1.0
        return value / self.kappa1


def _validate_egn_moments(mu4: float, mu6: float):
    m4, m6 = float(mu4), float(mu6)
    if not (math.isfinite(m4) and math.isfinite(m6)):
        raise ValueError("EGN moments must be finite")
    # E[Y]=1, Y=|X|**2: E[Y**2]>=1, E[Y**3]>=E[Y**2]**2.
    if m4 < 1.0 - 1e-12 or m6 < m4**2 - 1e-12:
        raise ValueError("inconsistent standardized fourth/sixth moments")


@lru_cache(maxsize=16)
def _egn_legendre_rule(order: int):
    if int(order) != order or order < 16:
        raise ValueError("egn_frequency_order must be an integer >= 16")
    x, w = roots_legendre(int(order))
    x.setflags(write=False)
    w.setflags(write=False)
    return x, w


def _egn_split_rule(order: int, split: float):
    x, w = _egn_legendre_rule(order)
    edges = [-0.5] + ([float(split)] if -0.5 < split < 0.5 else []) + [0.5]
    nodes, weights = [], []
    for lo, hi in zip(edges[:-1], edges[1:]):
        nodes.append(lo + (x + 1.0) * (hi - lo) / 2.0)
        weights.append(w * (hi - lo) / 2.0)
    return np.concatenate(nodes), np.concatenate(weights)


def _egn_link_function(f1, f2, f, spans: Sequence[Span], z_order: int):
    """Complex link amplitude, retaining phase required by EGN corrections.

    Fast homogeneous EDFA path is Carena Eqs. (10)-(12). Other links reuse
    the existing span amplitudes, following the generalization on p. 16341.
    This helper is used only by the optional EGN path.
    """
    spans = tuple(spans)
    if not spans:
        raise ValueError("at least one span is required")
    first = spans[0]
    if (first.amplification == "lumped_edfa" and first.gain_db is None
            and first.dcu_ps2 == 0.0 and all(sp == first for sp in spans)):
        q = phase_mismatch_beta23(f1, f2, f, first.beta2, first.beta3_ps3_km)
        single = first.gamma_W_inv_km * _safe_complex_span_integral(
            q, first.alpha_field_per_km, first.length_km)
        theta = np.remainder(q * first.length_km + np.pi, 2.0 * np.pi) - np.pi
        n = len(spans)
        phasor = (n * np.sinc(n * theta / (2.0 * np.pi))
                  / np.sinc(theta / (2.0 * np.pi))
                  * np.exp(0.5j * (n - 1) * theta))
        return single * phasor
    return np.sum(_link_amplitudes(f1, f2, f, spans, z_order=z_order), axis=0)


def egn_sci_coefficients(
    cut_channel: Channel,
    spans: Sequence[Span],
    f_THz: float,
    options: GNIntegralOptions = GNIntegralOptions(),
) -> EGNSCICoefficients:
    """Compute rectangular-spectrum SCI, Carena Eqs. (7), (43)-(45).

    For normalized offsets u=(f-center)/Rs, the pulse spectrum is 1/Rs.
    Each coefficient therefore has a common 1/Rs prefactor. Inner integration
    limits enforce the exact three-frequency SCI domain; the outer interval
    is split at its boundary kink. No fit or literature target enters here.

    Only the paper's rectangular, zero spectral-phase pulse is supported.
    Nonzero roll-off/custom spectral phase needs extra EGN terms (Sect. 2).
    """
    if options.accumulation != "coherent":
        raise ValueError("EGN SCI requires coherent accumulation")
    rectangular = (cut_channel.pulse_shape == "rect" or
                   (cut_channel.pulse_shape in ("rrc", "raised_cosine")
                    and cut_channel.rolloff == 0.0))
    if not rectangular:
        raise ValueError("EGN SCI currently requires a rectangular spectrum (rolloff=0)")
    rs = cut_channel.symbol_rate_THz
    if not math.isfinite(rs) or rs <= 0.0:
        raise ValueError("baud_GBd must be positive and finite")
    center = float(cut_channel.center_THz)
    t = (float(f_THz) - center) / rs
    if not math.isfinite(t) or abs(t) > 0.5 + 1e-12:
        raise ValueError("EGN SCI evaluation frequency must lie inside the CUT receiver band")
    spans = tuple(spans)
    n = options.egn_frequency_order
    v, wv = _egn_legendre_rule(n)
    outer, wo = _egn_split_rule(n, t)
    lo = np.maximum(-0.5, t - outer - 0.5)
    hi = np.minimum(0.5, t - outer + 0.5)
    inner = lo[:, None] + (v[None, :] + 1.0) * (hi - lo)[:, None] / 2.0
    wi = wv[None, :] * (hi - lo)[:, None] / 2.0
    amp = _egn_link_function(center + rs * outer[:, None], center + rs * inner,
                             f_THz, spans, options.z_quadrature_order)
    inner_amplitude = np.sum(wi * amp, axis=1)
    k1 = 16.0 / (27.0 * rs) * np.sum(wo[:, None] * wi * np.abs(amp)**2)
    k2a = 80.0 / (81.0 * rs) * np.sum(wo * np.abs(inner_amplitude)**2)
    k3 = 16.0 / (81.0 * rs) * abs(np.sum(wo * inner_amplitude))**2

    # Eq. (44): outer variable is f3; f1 = f3 - f2 + f.
    outer, wo = _egn_split_rule(n, -t)
    lo = np.maximum(-0.5, outer + t - 0.5)
    hi = np.minimum(0.5, outer + t + 0.5)
    inner = lo[:, None] + (v[None, :] + 1.0) * (hi - lo)[:, None] / 2.0
    wi = wv[None, :] * (hi - lo)[:, None] / 2.0
    amp = _egn_link_function(center + rs * (outer[:, None] - inner + t),
                             center + rs * inner, f_THz, spans, options.z_quadrature_order)
    inner_amplitude = np.sum(wi * amp, axis=1)
    k2b = 16.0 / (81.0 * rs) * np.sum(wo * np.abs(inner_amplitude)**2)
    return EGNSCICoefficients(float(k1), float(k2a), float(k2b), float(k3))


def integrate_egn_sci_coefficients(
    cut_channel: Channel,
    spans: Sequence[Span],
    options: GNIntegralOptions = GNIntegralOptions(),
    receiver_points: int = 31,
) -> EGNSCICoefficients:
    """Integrate SCI coefficients over +/- Rs/2, Carena Eq. (13); units 1/W**2.

    The ratio of integrated coefficients is a band-effective SCI multiplier,
    not a universal per-modulation constant. Uses the same coefficient
    implementation as the modulation-aware QMC path.
    """
    x, w = roots_legendre(int(receiver_points))
    half = cut_channel.symbol_rate_THz / 2.0
    total = np.zeros(4)
    spans = tuple(spans)
    for xx, ww in zip(x, w):
        c = egn_sci_coefficients(cut_channel, spans,
                                 cut_channel.center_THz + half * xx, options)
        total += half * ww * np.array([c.kappa1, c.kappa2a, c.kappa2b, c.kappa3])
    return EGNSCICoefficients(*(float(v) for v in total))


def gn_nli_psd_qmc(
    system: WDMSystem,
    spans: Sequence[Span],
    f_THz: float,
    options: GNIntegralOptions = GNIntegralOptions(),
    cut_channel: Optional[Channel] = None,
) -> GNResult:
    """Evaluate the full WDM GN double integral at one output frequency.

    Integration covers the total occupied WDM support. If cut_channel is
    supplied, only samples with all three frequencies inside its fixed spectral
    support receive modulation_phi. The mask is centered on the CUT center,
    not the output quadrature frequency. No CUT (or Phi=1) uses legacy GN.
    For overlapping channel supports this is a geometric mask, not an exact
    decomposition by channel origin. Optional egn_mu4/egn_mu6 compute a
    frequency- and link-dependent SCI ratio using EGN integrals, preserving
    XCI/MCI as GN. This is not full WDM EGN.
    """
    if cut_channel is not None:
        phi = float(cut_channel.modulation_phi)
        if not math.isfinite(phi) or phi < 0.0:
            raise ValueError("modulation_phi must be finite and nonnegative")
        use_egn = cut_channel.egn_mu4 is not None or cut_channel.egn_mu6 is not None
        if use_egn:
            if cut_channel.egn_mu4 is None or cut_channel.egn_mu6 is None:
                raise ValueError("provide both egn_mu4 and egn_mu6")
            if phi != 1.0:
                raise ValueError("do not combine manual modulation_phi with EGN moments")
            _validate_egn_moments(cut_channel.egn_mu4, cut_channel.egn_mu6)
            # The geometric SCI mask is unambiguous only for nonoverlapping CUTs.
            matching = [ch for ch in system.channels
                        if ch.center_THz == cut_channel.center_THz]
            if len(matching) != 1:
                raise ValueError("EGN requires one unique CUT center in system.channels")
            for other in system.channels:
                if other.center_THz == cut_channel.center_THz:
                    continue
                if abs(other.center_THz - cut_channel.center_THz) < (
                        other.support_half_width_THz + cut_channel.support_half_width_THz - 1e-15):
                    raise ValueError("EGN SCI requires CUT support not to overlap other channels")
            if (cut_channel.egn_mu4, cut_channel.egn_mu6) != (2.0, 6.0):
                phi = egn_sci_coefficients(cut_channel, spans, f_THz, options).sci_multiplier(
                    cut_channel.egn_mu4, cut_channel.egn_mu6)
    lo, hi = system.support_bounds_THz()
    width = hi - lo
    sob = qmc.Sobol(d=2, scramble=True, seed=int(options.seed))
    u = sob.random_base2(int(options.sobol_power))
    f1 = lo + width * u[:, 0]
    f2 = lo + width * u[:, 1]
    f3 = f1 + f2 - float(f_THz)

    G1 = system.psd(f1)
    G2 = system.psd(f2)
    G3 = system.psd(f3)
    active = (G1 > 0.0) & (G2 > 0.0) & (G3 > 0.0)
    if not np.any(active):
        return GNResult(float(f_THz), 0.0, 2**int(options.sobol_power), (lo, hi), options.accumulation)

    K = np.zeros_like(f1, dtype=float)
    K[active] = link_kernel(
        f1[active], f2[active], float(f_THz), spans,
        accumulation=options.accumulation, z_order=options.z_quadrature_order,
    )
    integrand = DP_GN_COEFF * G1 * G2 * G3 * K
    if cut_channel is not None and phi != 1.0:
        half_bw = cut_channel.support_half_width_THz
        center = float(cut_channel.center_THz)
        sci_mask = ((np.abs(f1 - center) <= half_bw)
                    & (np.abs(f2 - center) <= half_bw)
                    & (np.abs(f3 - center) <= half_bw))
        phi_correction = np.where(sci_mask, phi, 1.0)
        integrand = integrand * phi_correction
    val = float(np.mean(integrand) * width**2)
    return GNResult(float(f_THz), val, 2**int(options.sobol_power), (lo, hi), options.accumulation)


def gn_nli_psd_multi_seed(
    system: WDMSystem,
    spans: Sequence[Span],
    f_THz: float,
    sobol_power: int = 18,
    seeds: Iterable[int] = (1, 2, 3, 4),
    accumulation: AccumulationMode = "coherent",
    z_quadrature_order: int = 64,
    cut_channel: Optional[Channel] = None,
    egn_frequency_order: int = 512,
) -> dict:
    """Repeated scrambled-Sobol estimate with numerical uncertainty summary."""
    vals = []
    for s in seeds:
        res = gn_nli_psd_qmc(
            system, spans, f_THz,
            GNIntegralOptions(sobol_power=sobol_power, seed=int(s), accumulation=accumulation,
                              z_quadrature_order=z_quadrature_order,
                              egn_frequency_order=egn_frequency_order),
            cut_channel=cut_channel,
        )
        vals.append(res.g_nli_W_per_THz)
    a = np.asarray(vals, dtype=float)
    return {
        "mean_W_per_THz": float(np.mean(a)),
        "std_W_per_THz": float(np.std(a, ddof=1)) if len(a) > 1 else 0.0,
        "relative_std_pct": float(np.std(a, ddof=1) / np.mean(a) * 100.0) if len(a) > 1 and np.mean(a) != 0 else 0.0,
        "values_W_per_THz": a,
    }


def integrate_nli_over_channel(
    system: WDMSystem,
    spans: Sequence[Span],
    cut_index: int,
    options: GNIntegralOptions = GNIntegralOptions(),
    receiver_points: int = 7,
) -> float:
    """Integrate G_NLI over the CUT receiver band; return NLI power [W].

    A simple rectangular receiver noise band of width Rs is used, matching the
    common GN-model channel-noise integration convention. For detailed receiver
    matched-filter studies, replace this helper with an explicit |H_Rx(f)|^2.
    """
    ch = system.channels[int(cut_index)]
    # Gauss-Legendre over +/- Rs/2 (not the RRC excess band).
    x, w = np.polynomial.legendre.leggauss(int(receiver_points))
    half = ch.symbol_rate_THz / 2.0
    fs = ch.center_THz + half * x
    wf = half * w
    vals = []
    for k, ff in enumerate(fs):
        # Change seed across output-frequency nodes to reduce correlated QMC artefacts.
        opt = replace(options, seed=int(options.seed) + 1009 * k)
        vals.append(gn_nli_psd_qmc(system, spans, float(ff), opt, cut_channel=ch).g_nli_W_per_THz)
    return float(np.sum(wf * np.asarray(vals)))


def common_power_scaling_self_test() -> dict:
    """Small P^3 scaling test for the full WDM engine."""
    system = WDMSystem.equispaced(3, 50.0, 32.0, -3.0)
    spans = [Span(80.0, 0.2, 1.3, D_ps_nm_km=17.0) for _ in range(2)]
    opt = GNIntegralOptions(sobol_power=12, seed=7)
    g1 = gn_nli_psd_qmc(system, spans, 0.0, opt).g_nli_W_per_THz
    g2 = gn_nli_psd_qmc(system.with_common_power_offset_db(3.01029995664), spans, 0.0, opt).g_nli_W_per_THz
    ratio = g2 / g1
    return {"ratio_for_2x_power": ratio, "target": 8.0, "error_pct": abs(ratio - 8.0) / 8.0 * 100.0}


def egn_backward_compat_self_test() -> dict:
    """Check default/explicit Phi=1 against the uncorrected GN path."""
    system = WDMSystem.equispaced(3, 50.0, 32.0, -3.0)
    spans = [Span(80.0, 0.2, 1.3, D_ps_nm_km=17.0)]
    opt = GNIntegralOptions(sobol_power=12, seed=7)
    ch = replace(system.channels[1], modulation_phi=1.0)
    errors = []
    for f in (ch.center_THz, ch.center_THz + 0.3 * ch.symbol_rate_THz):
        old = gn_nli_psd_qmc(system, spans, f, opt).g_nli_W_per_THz
        new = gn_nli_psd_qmc(system, spans, f, opt, cut_channel=ch).g_nli_W_per_THz
        errors.append(abs(new - old) / max(abs(old), np.finfo(float).tiny))
    # Independently reconstruct legacy receiver integration: no CUT argument.
    x, w = np.polynomial.legendre.leggauss(3)
    half = ch.symbol_rate_THz / 2.0
    old_power = float(np.sum(half * w * np.asarray([
        gn_nli_psd_qmc(system, spans, float(ch.center_THz + half * xx),
                       replace(opt, seed=opt.seed + 1009 * k)).g_nli_W_per_THz
        for k, xx in enumerate(x)
    ])))
    new_power = integrate_nli_over_channel(system, spans, 1, opt, 3)
    errors.append(abs(new_power - old_power) / max(abs(old_power), np.finfo(float).tiny))
    old_multi = gn_nli_psd_multi_seed(system, spans, ch.center_THz, 12, (7, 8))
    new_multi = gn_nli_psd_multi_seed(
        system, spans, ch.center_THz, 12, (7, 8), cut_channel=ch)
    errors.extend(np.abs(
        new_multi["values_W_per_THz"] - old_multi["values_W_per_THz"]
    ) / np.maximum(np.abs(old_multi["values_W_per_THz"]), np.finfo(float).tiny))
    error = float(max(errors))
    assert error <= 1e-10, f"Phi=1 regression: {error}"
    return {"passed": True, "max_relative_error": error, "tolerance": 1e-10}


def modulation_phi_self_test() -> dict:
    """Legacy multiplier algebra only: SCI scales, XCI/MCI stays unchanged.

    The deliberately chosen 1.2 is a plumbing test, not a physical EGN value.
    See egn_literature_self_test for the external model comparison.
    """
    spans = [Span(80.0, 0.2, 1.3, D_ps_nm_km=17.0)]
    opt = GNIntegralOptions(sobol_power=13, seed=7)
    ratios = {}
    for n in (1, 3):
        system = WDMSystem.equispaced(n, 50.0, 32.0, -3.0)
        index = n // 2
        def power(phi):
            channels = list(system.channels)
            channels[index] = replace(channels[index], modulation_phi=phi)
            return integrate_nli_over_channel(
                replace(system, channels=tuple(channels)), spans, index, opt, 3)
        p0, p1, p12 = (power(phi) for phi in (0.0, 1.0, 1.2))
        assert p12 > p1 > 0.0
        assert np.isclose(p12 - p0, 1.2 * (p1 - p0), rtol=1e-10, atol=0.0)
        if n == 1:
            assert p0 == 0.0
            assert np.isclose(p12 / p1, 1.2, rtol=1e-10, atol=0.0)
        else:
            assert p0 > 0.0
        ratios[f"{n}_channel_ratio"] = p12 / p1
    return {"passed": True, **ratios}


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




def ch_center_wavelength_nm(spans, ch):
    return float(spans[0].wavelength_nm)



import math
import numpy as np
from scipy.special import exp1, expi, roots_legendre
from functools import lru_cache

@lru_cache(maxsize=32)
def _panel_rule(n):
    return roots_legendre(n)

def _adaptive_integral(fun, edges, omega=0., rtol=2e-5, order=12, phase_step=24*np.pi, max_panels=8192):
    edges=np.unique(edges); parts=[]
    for a,b in zip(edges[:-1],edges[1:]):
        nn=max(1,int(np.ceil((b-a)*omega/phase_step)))
        ee=np.linspace(a,b,nn+1)
        parts.extend(zip(ee[:-1],ee[1:]))
    if len(parts)>max_panels: raise RuntimeError('phase panels exceed budget')
    parts=np.array(parts)
    def evaluate(parts,n):
        x,w=_panel_rule(n);mid=parts.mean(axis=1);half=(parts[:,1]-parts[:,0])/2
        xx=mid[:,None]+half[:,None]*x
        v=fun(xx.ravel()).reshape((len(parts),n,-1))
        return np.sum(v*w[None,:,None],axis=1)*half[:,None]
    total=None;error=None;count=0
    for lev in range(12):
        vl=evaluate(parts,order);vh=evaluate(parts,2*order)
        err=abs(vh-vl);count+=len(parts)
        # Conservative mixed norm, used on individual physical components.
        estimate=vh.sum(axis=0) if total is None else total+vh.sum(axis=0)
        scale=np.maximum(abs(estimate),max(1e-12,1e-4*max(abs(estimate))))
        crit=rtol*scale*((parts[:,1]-parts[:,0])/(edges[-1]-edges[0]))[:,None]
        ok=np.all(err<=crit,axis=1)
        if total is None: total=np.zeros_like(vh[0]);error=np.zeros_like(err[0])
        total+=vh[ok].sum(axis=0);error+=err[ok].sum(axis=0)
        if np.all(ok):return total,error,count
        bad=parts[~ok];mid=bad.mean(axis=1)
        parts=np.concatenate([np.column_stack([bad[:,0],mid]),np.column_stack([mid,bad[:,1]])])
        if count+len(parts)>max_panels:raise RuntimeError(f'adaptive budget exceeded: {count}, {len(parts)}')
    raise RuntimeError('adaptive depth exceeded')

class _EGNQuadrature:
    def __init__(self,span,counts,rs,rtol=1e-4,phase_step=32*np.pi,order=12):
        self.span=span;self.ns=np.array(counts,int);self.N=int(max(counts));self.a=2*span.alpha_field_per_km;self.L=span.length_km
        self.rs=rs;self.c=4*np.pi**2*span.beta2*rs**2;self.gamma=span.gamma_W_inv_km
        self.r=np.exp(-self.a*self.L);self.b=self.L*np.arange(1,self.N+1)
        self.rtol=rtol;self.phase_step=phase_step;self.order=order;self.stats=[];self.cache={}
        if self.a*self.L*self.N>600:
            raise ValueError("This exact primitive requires N * power_alpha * L <= 600; split-profile extensions are not validated here.")

    def mu(self,q):
        th=np.remainder(q[:,None]*self.L+np.pi,2*np.pi)-np.pi
        single=-np.expm1((-self.a+1j*q[:,None])*self.L)/(self.a-1j*q[:,None])
        return self.gamma*single*self.ns*np.sinc(self.ns*th/(2*np.pi))/np.sinc(th/(2*np.pi))*np.exp(.5j*(self.ns-1)*th)

    def primitive(self,q,need_gn=True):
        """Exact antiderivatives of mu(q), |mu(q)|^2 and q|mu(q)|^2.

        mu(q)/gamma = [1-r*exp(i*q*L)]/(a-i*q) * sum exp(i*q*n*L).
        The exponential integral E1 evaluates the linear-frequency integral.
        |mu|^2 has cosine coefficients c0=N*(1-r)^2+2*r,
        cm=2*(N-m)*(1-r)^2 (0<m<N), cN=-2*r.
        This is algebra from Carena Eqs. (10)-(12), not an approximation/fitting.
        Positive a keeps E1/Ei arguments away from the negative-real-axis cut.
        """
        z=(self.a-1j*q[:,None])*self.b
        v=np.exp(self.a*self.b)*exp1(z)
        j=-1j*v
        cs=np.cumsum(j,axis=1)-j
        amp=1j*np.log(self.a-1j*q[:,None])+(1-self.r)*cs[:,self.ns-1]-self.r*j[:,self.ns-1]
        if not need_gn:return amp,None,None
        ww=np.exp(-self.a*self.b)*expi(np.conj(z))
        integ=(v+ww).imag/(2*self.a)
        cc=np.cumsum(integ,axis=1)-integ
        cm=np.cumsum(integ*np.arange(1,self.N+1),axis=1)-integ*np.arange(1,self.N+1)
        gn=(self.ns*(1-self.r)**2+2*self.r)*np.arctan(q[:,None]/self.a)/self.a
        gn+=2*(1-self.r)**2*(self.ns*cc[:,self.ns-1]-cm[:,self.ns-1])-2*self.r*integ[:,self.ns-1]
        ii=(ww-v).real/2
        cc=np.cumsum(ii,axis=1)-ii
        cm=np.cumsum(ii*np.arange(1,self.N+1),axis=1)-ii*np.arange(1,self.N+1)
        gn1=(self.ns*(1-self.r)**2+2*self.r)*np.log(self.a**2+q[:,None]**2)/2
        gn1+=2*(1-self.r)**2*(self.ns*cc[:,self.ns-1]-cm[:,self.ns-1])-2*self.r*ii[:,self.ns-1]
        return amp,gn,gn1

    def linear_inner(self,x,lo,hi,t,need_gn=True):
        k=self.c*(x-t);ql=k*(lo-t);qh=k*(hi-t)
        amp=np.zeros((len(x),len(self.ns)),complex);gn=np.zeros_like(amp.real)
        small=abs(k)*self.N*self.L*np.maximum.reduce([abs(lo-t),abs(hi-t),hi-lo])<.05
        for sel in [~small]:
            if not np.any(sel):continue
            al,glo,_=self.primitive(ql[sel],need_gn);ah,ghi,_=self.primitive(qh[sel],need_gn)
            amp[sel]=self.gamma*(ah-al)/k[sel,None]
            if need_gn:gn[sel]=self.gamma**2*(ghi-glo)/k[sel,None]
        if np.any(small):
            u,w=_panel_rule(12);yy=(lo[small,None]+hi[small,None])/2+(hi[small,None]-lo[small,None])*u/2
            vals=self.mu((k[small,None]*(yy-t)).ravel()).reshape((sum(small),12,-1))
            amp[small]=np.sum(w[None,:,None]*vals,axis=1)*(hi[small,None]-lo[small,None])/2
            if need_gn:gn[small]=np.sum(w[None,:,None]*abs(vals)**2,axis=1)*(hi[small,None]-lo[small,None])/2
        return amp,gn

    def outer(self,A,B,C,t):
        key=('o',A,B,C,round(t,14))
        if key in self.cache:return self.cache[key]
        left=max(A[0],t+C[0]-B[1]);right=min(A[1],t+C[1]-B[0]);n=len(self.ns)
        if right<=left:return np.zeros((3,n))
        edges=[left,right]
        edges.extend(z for z in [t,t+C[0]-B[0],t+C[1]-B[1]] if left<z<right)
        omega=abs(self.c)*self.N*self.L*max(abs(B[0]-t),abs(B[1]-t),abs(C[0]),abs(C[1]))*2
        def fun(x):
            lo=np.maximum(B[0],t+C[0]-x);hi=np.minimum(B[1],t+C[1]-x)
            amp,gn=self.linear_inner(x,lo,hi,t,need_gn=False)
            return np.column_stack([abs(amp)**2,amp.real,amp.imag])
        value,err,count=_adaptive_integral(fun,edges,omega,self.rtol,self.order,self.phase_step)
        self.stats.append((key,count,float(max(err/np.maximum(abs(value),1e-15)))))
        v=value.reshape((3,n));out=np.array([np.zeros(n),v[0],v[1]**2+v[2]**2]);self.cache[key]=out
        return out

    def gn_receiver_integrated(self,A,B,C):
        # x=f1-f; at fixed x the overlap in receiver f is a trapezoid in y=f2-f.
        # Both zeroth and first y moments of |mu(c*x*y)|^2 are integrated exactly.
        D0=(-.5,.5)
        left=max(A[0]-D0[1],C[0]-B[1]);right=min(A[1]-D0[0],C[1]-B[0])
        if right<=left:return np.zeros(len(self.ns))
        key=('g',A,B,C)
        if key in self.cache:return self.cache[key]
        edges=[left,right]+[z for z in [0,A[0]-D0[0],A[1]-D0[1],C[0]-B[0],C[1]-B[1]] if left<z<right]
        def fun(x):
            dl=np.maximum(D0[0],A[0]-x);dh=np.minimum(D0[1],A[1]-x)
            el=np.maximum(B[0],C[0]-x);eh=np.minimum(B[1],C[1]-x)
            knots=np.sort(np.column_stack([el-dh,el-dl,eh-dh,eh-dl]),axis=1)
            total=np.zeros((len(x),len(self.ns)))
            k=self.c*x
            for i in range(3):
                lo=knots[:,i];hi=knots[:,i+1];mid=(lo+hi)/2
                top=np.minimum(dh,eh-mid);bot=np.maximum(dl,el-mid)
                slope=-(eh-mid<dh).astype(float)+(el-mid>dl).astype(float)
                intercept=top-bot-slope*mid
                small=abs(k)*self.N*self.L*np.maximum.reduce([abs(lo),abs(hi),hi-lo])<.08
                use=(~small)&(hi>lo)
                if np.any(use):
                    _,vl,ml=self.primitive(k[use]*lo[use]);_,vh,mh=self.primitive(k[use]*hi[use])
                    total[use]+=self.gamma**2*(intercept[use,None]*(vh-vl)/k[use,None]+slope[use,None]*(mh-ml)/k[use,None]**2)
                if np.any(small):
                    gx,gw=_panel_rule(12);y=mid[small,None]+(hi[small,None]-lo[small,None])*gx/2
                    v=abs(self.mu((k[small,None]*y).ravel()))**2
                    v=v.reshape((sum(small),12,-1))
                    wg=gw*(hi[small,None]-lo[small,None])/2*(intercept[small,None]+slope[small,None]*y)
                    total[small]+=np.sum(wg[:,:,None]*v,axis=1)
            return total
        omega=4*abs(self.c)*self.N*self.L*max(1,abs(B[0])+.5,abs(B[1])+.5)
        v,e,cnt=_adaptive_integral(fun,edges,omega,self.rtol,self.order,self.phase_step)
        self.stats.append((key,cnt,float(max(e/np.maximum(abs(v),1e-15)))))
        self.cache[key]=v;return v


    def sum(self,A,B,C,t):
        key=('s',A,B,C,round(t,14))
        if key in self.cache:return self.cache[key]
        left=max(A[0]+B[0],t+C[0]);right=min(A[1]+B[1],t+C[1]);n=len(self.ns)
        if right<=left:return np.zeros(n)
        edges=[left,right]+[z for z in [A[0]+B[1],A[1]+B[0],2*t] if left<z<right]
        # Inner quadrature splits at the quadratic stationary point q=u1+u2.
        def fun(ss):
            out=[]
            for s in ss:
                lo=max(A[0],s-B[1]);hi=min(A[1],s-B[0]);mid=s/2
                ee=[lo]+([mid] if lo<mid<hi else [])+[hi]
                total=np.zeros(n,complex)
                for a,b in zip(ee[:-1],ee[1:]):
                    q0=self.c*(a-t)*(s-a-t);q1=self.c*(b-t)*(s-b-t)
                    panels=max(1,int(np.ceil(abs(q1-q0)*self.N*self.L/(4*np.pi))))
                    # Uniform u partition is safe using max phase derivative.
                    om=abs(self.c)*self.N*self.L*max(abs(s-2*a),abs(s-2*b))
                    panels=max(panels,int(np.ceil(om*(b-a)/(4*np.pi))))
                    bounds=np.linspace(a,b,panels+1);gx,gw=_panel_rule(self.order)
                    u=(bounds[:-1,None]+bounds[1:,None])/2+(bounds[1:,None]-bounds[:-1,None])*gx/2
                    w=(bounds[1:,None]-bounds[:-1,None])*gw/2
                    vv=self.mu((self.c*(u-t)*(s-u-t)).ravel()).reshape((-1,n))
                    total+=np.sum(w.ravel()[:,None]*vv,axis=0)
                out.append(abs(total)**2)
            return np.array(out)
        om=2*abs(self.c)*self.N*self.L*max(abs(A[0]-t),abs(A[1]-t),abs(B[0]-t),abs(B[1]-t))
        v,e,cnt=_adaptive_integral(fun,edges,om,self.rtol,self.order,self.phase_step)
        self.stats.append((key,cnt,float(max(e/np.maximum(abs(v),1e-15)))))
        self.cache[key]=v;return v


# PUBLIC PRECISION API
# The definitions below deliberately replace only the Full-EGN computation.
@dataclass(frozen=True)
class EGNFullOptions:
    """Numerical controls; none are physical calibration parameters.

    ``quadrature_rtol`` controls local integration estimates. ``convergence_rtol``
    controls the change in SCI, XCI, XMCI and total EGN after refinement.
    Receiver order and phase/panel resolution are checked separately.
    Set verify_convergence=False only for exploratory runs; the result will
    explicitly have diagnostics.converged=False.
    The old frequency_order control is replaced by panel_order/phase_step_rad.
    """
    receiver_points: int = 7
    max_receiver_points: int = 31
    panel_order: int = 12
    phase_step_rad: float = 32.0 * math.pi
    quadrature_rtol: float = 1e-4
    convergence_rtol: float = 3e-3
    verify_convergence: bool = True
    strict_convergence: bool = True
    symmetry_rtol: float = 2e-8
    power_rtol: float = 2e-8


@dataclass(frozen=True)
class EGNConvergenceReport:
    converged: bool
    frequency_converged: bool
    receiver_converged: bool
    frequency_relative_change: float
    receiver_relative_change: float
    convergence_rtol: float
    receiver_points: int
    refined_panel_order: int
    refined_phase_step_rad: float
    checked_span_counts: tuple[int, ...]
    note: str


class EGNConvergenceError(RuntimeError):
    """Requested numerical tolerance was not demonstrated; never a silent pass."""


@dataclass(frozen=True)
class EGNBreakdown:
    modulation: str
    gn_total_W: float
    sci_gn_W: float
    xci_gn_W: float
    mci_gn_W: float
    sci_correction_W: float
    xci_correction_W: float
    mci_correction_W: float
    total_egn_W: float
    assumptions: tuple[str, ...]
    diagnostics: EGNConvergenceReport

    @property
    def egn_sci_W(self): return self.sci_gn_W + self.sci_correction_W
    @property
    def egn_xci_W(self): return self.xci_gn_W + self.xci_correction_W
    @property
    def egn_mci_W(self): return self.mci_gn_W + self.mci_correction_W
    @property
    def egn_xmci_W(self): return self.egn_xci_W + self.egn_mci_W
    @property
    def correction_W(self): return self.sci_correction_W + self.xci_correction_W + self.mci_correction_W
    @property
    def ratio_to_gn(self): return self.total_egn_W / self.gn_total_W if self.gn_total_W else math.nan
    @property
    def ratio_to_gn_db(self): return 10*math.log10(self.ratio_to_gn) if self.ratio_to_gn>0 else math.nan


def _precision_scope(system, span, counts, cut_index, mods, options):
    ci=int(cut_index)
    if not 0<=ci<len(system.channels): raise IndexError("CUT index out of range")
    if (not counts or any(int(n)!=n or n<1 for n in counts)
            or len(set(counts))!=len(counts)):
        raise ValueError("span counts must be distinct positive integers")
    if span.amplification!='lumped_edfa' or span.gain_db is not None or span.dcu_ps2!=0 or span.beta3_ps3_km!=0:
        raise ValueError("Precision EGN requires beta2-only, loss-compensated lumped EDFA spans (gain_db=None, dcu=0). Use the GN API for other profiles.")
    if not all(math.isfinite(v) for v in [span.length_km,span.alpha_db_per_km,span.gamma_W_inv_km,span.beta2]):
        raise ValueError("Fiber parameters must be finite")
    if span.length_km<=0 or span.alpha_db_per_km<=0 or span.gamma_W_inv_km<0:
        raise ValueError("Require L>0, attenuation>0 and gamma>=0")
    cut=system.channels[ci];rs=cut.symbol_rate_THz
    if not math.isfinite(rs) or rs<=0: raise ValueError("Symbol rate must be positive")
    for ch,m in zip(system.channels,mods):
        if not ((ch.pulse_shape=='rect') or (ch.pulse_shape in ('rrc','raised_cosine') and ch.rolloff==0)):
            raise ValueError("Precision EGN requires rectangular, zero-phase spectra; rolloff must be zero")
        if not math.isclose(ch.symbol_rate_THz,rs,rel_tol=1e-12):
            raise ValueError("Precision EGN requires equal symbol rates")
        if not math.isfinite(ch.power_W) or ch.power_W<=0 or not math.isfinite(ch.center_THz):
            raise ValueError("Require positive finite channel powers and finite centers")
        if ch.modulation_phi!=1 or ch.egn_mu4 is not None or ch.egn_mu6 is not None:
            raise ValueError("Supply moments through modulation/channel_modulations, not pre-corrected Channel fields")
        _validate_egn_moments(m.mu4,m.mu6)
    indices=sorted(range(len(system.channels)),key=lambda j:system.channels[j].center_THz)
    centers=np.array([system.channels[j].center_THz for j in indices])
    if len(centers)>1 and np.any(np.diff(centers)<rs-1e-14):
        raise ValueError("Channel spectra must not overlap")
    grid=None
    if len(centers)>=3:
        H=(len(centers)-1)//2
        if len(centers)%2!=1 or indices[H]!=ci:
            raise ValueError("Appendix B requires an odd symmetric comb and the central CUT")
        diffs=np.diff(centers)
        if not np.allclose(diffs,diffs.mean(),rtol=options.symmetry_rtol,atol=1e-14):
            raise ValueError("Appendix B requires equal channel spacing")
        powers=np.array([ch.power_W for ch in system.channels])
        if not np.allclose(powers,powers[ci],rtol=options.power_rtol,atol=1e-20):
            raise ValueError("Appendix B requires equal channel powers")
        if any(not math.isclose(m.mu4,mods[ci].mu4,rel_tol=1e-12) or not math.isclose(m.mu6,mods[ci].mu6,rel_tol=1e-12) for m in mods):
            raise ValueError("This Appendix-B path requires a common modulation distribution")
        grid={i:indices[i+H] for i in range(-H,H+1)}
    if (options.receiver_points<3 or options.max_receiver_points<options.receiver_points
            or options.panel_order<8 or options.phase_step_rad<=0
            or not 0<options.quadrature_rtol<1 or not 0<options.convergence_rtol<1):
        raise ValueError("Invalid precision quadrature controls")
    return rs,grid


class _EGNSystemIntegrator:
    """Carries spectrum domains and power weights separately from quadrature."""
    def __init__(self,system,span,counts,ci,mods,options):
        self.system=system;self.ci=ci;self.mods=mods;self.options=options
        rs,self.grid=_precision_scope(system,span,counts,ci,mods,options)
        self.kernel=_EGNQuadrature(span,counts,rs,options.quadrature_rtol,options.phase_step_rad,options.panel_order)
        c=system.channels[ci].center_THz
        self.intervals=[((ch.center_THz-c)/rs-.5,(ch.center_THz-c)/rs+.5) for ch in system.channels]
        self.powers=np.array([ch.power_W for ch in system.channels])
        # Equal-rate endpoints are rounded only to merge floating-point duplicates.
        self.intervals=[tuple(round(v,13) for v in r) for r in self.intervals]

    def gn(self):
        n=len(self.intervals);out=np.zeros((3,len(self.kernel.ns)))
        for a in range(n):
          for b in range(a,n):
           for c in range(n):
            A,B,C=(self.intervals[j] for j in (a,b,c))
            if max(A[0]-.5,C[0]-B[1])>=min(A[1]+.5,C[1]-B[0]):continue
            used={a,b,c};typ=0 if used=={self.ci} else (1 if len(used-{self.ci})==1 else 2)
            # Exchange f1/f2 and global frequency reflection preserve beta2 GN.
            candidates=[]
            for aa,bb in [(A,B),(B,A)]:
                candidates.append((aa,bb,C))
                candidates.append(((-aa[1],-aa[0]),(-bb[1],-bb[0]),(-C[1],-C[0])))
            aa,bb,cc=min(candidates)
            v=self.kernel.gn_receiver_integrated(aa,bb,cc)
            out[typ]+=DP_GN_COEFF*(1 if a==b else 2)*self.powers[a]*self.powers[b]*self.powers[c]*v
        return out

    def correction_psd_normalized(self,t):
        """Returns Rs * correction PSD, with actual DP power weights."""
        k=self.kernel;R=self.intervals;ci=self.ci;p=self.powers;mods=self.mods
        out=np.zeros((3,len(k.ns)));m=mods[ci];C=R[ci]
        if m.egn_phi or m.egn_psi:
            o=k.outer(C,C,C,t);s=k.sum(C,C,C,t) if m.egn_phi else 0.
            out[0]=p[ci]**3*(m.egn_phi*(80/81*o[1]+16/81*s)+m.egn_psi*16/81*o[2])
        for j,I in enumerate(R):
            if j==ci:continue
            b=mods[j]
            # Eq. (18): powers are Pcut*Pint^2, Pcut^2*Pint and Pint^3.
            if b.egn_phi:
                out[1]+=p[ci]*p[j]**2*b.egn_phi*80/81*k.outer(C,I,I,t)[1]
            if m.egn_phi:
                out[1]+=p[ci]**2*p[j]*m.egn_phi*(80/81*k.outer(I,C,C,t)[1]+16/81*k.sum(C,C,I,t))
            if b.egn_phi or b.egn_psi:
                o=k.outer(I,I,I,t)
                out[1]+=p[j]**3*(b.egn_phi*(80/81*o[1]+16/81*k.sum(I,I,I,t))+b.egn_psi*16/81*o[2])
        if self.grid is not None and m.egn_phi:
            H=(len(R)-1)//2
            def rr(i):return R[self.grid[i]]
            # Eq. (39)-(41): evaluate both mirrored domains at the actual f.
            # Replacing them by 2*one_domain is only generally valid after a
            # symmetric receiver integration, not pointwise at f != 0.
            for sign in [-1,1]:
                for n in range(1,H+1):
                    out[2]+=m.egn_phi*p[ci]**3*80/81*k.outer(rr(-sign),rr(sign*n),rr(sign*n),t)[1]
                for n in range(2,H+1):
                    out[2]+=m.egn_phi*p[ci]**3*80/81*k.outer(rr(sign),rr(sign*n),rr(sign*n),t)[1]
                    ms=[n//2] if n%2==0 else [(n-1)//2,(n+1)//2]
                    for mm in ms:out[2]+=m.egn_phi*p[ci]**3*16/81*k.sum(rr(sign*mm),rr(sign*mm),rr(sign*n),t)
        return out

    def corrections(self,receiver_points):
        xs,ws=roots_legendre(int(receiver_points));out=np.zeros((3,len(self.kernel.ns)))
        symmetric=(len(self.intervals)==1 or self.grid is not None)
        for x,w in zip(xs,ws):
            if symmetric and x< -1e-14:continue
            fac=w if symmetric and x>1e-14 else w/2
            out+=fac*self.correction_psd_normalized(float(x)/2)
        return out


def _observable_change(gn0,corr0,gn1,corr1):
    """Check final observables, since GN/correction cancellation amplifies errors."""
    def observ(g,c):
        v=g+c
        return np.array([v[0],v[1],v[1]+v[2],v.sum(axis=0)])
    a=observ(gn0,corr0);b=observ(gn1,corr1)
    floor=np.maximum(abs(b[-1])*1e-10,1e-30)
    mask=np.maximum(abs(a),abs(b))>floor
    rel=np.where(mask,abs(a-b)/np.maximum(abs(b),floor),0.)
    return float(np.max(rel))


def egn_span_sweep(system: WDMSystem, span: Span, span_counts: Sequence[int], cut_index: int,
                   modulation: str | ModulationFormat='QPSK',
                   channel_modulations: Optional[Sequence[str | ModulationFormat]]=None,
                   full_options: EGNFullOptions=EGNFullOptions()) -> dict[int, EGNBreakdown]:
    """Compute many identical-span lengths using the same exact primitives.

    The convergence status is a numerical result for these inputs; it makes
    no claim about agreement with a paper, experiment or SSFM.
    """
    counts=tuple(span_counts);ci=int(cut_index)
    def mod(x):return x if isinstance(x,ModulationFormat) else get_modulation(x)
    mods=[mod(modulation)]*len(system.channels) if channel_modulations is None else [mod(x) for x in channel_modulations]
    if len(mods)!=len(system.channels):raise ValueError('One modulation per channel is required')
    o=full_options;eng=_EGNSystemIntegrator(system,span,counts,ci,mods,o)
    gn=eng.gn();nr=int(o.receiver_points);corr=eng.corrections(nr)
    freq_delta=rx_delta=math.nan;fc=rc=False;final_o=o
    if o.verify_convergence:
        final_o=replace(o,quadrature_rtol=o.quadrature_rtol/3,panel_order=o.panel_order+4,phase_step_rad=o.phase_step_rad/2)
        fine=_EGNSystemIntegrator(system,span,counts,ci,mods,final_o)
        g2=fine.gn();c2=fine.corrections(nr)
        freq_delta=_observable_change(gn,corr,g2,c2);fc=freq_delta<=o.convergence_rtol
        gn,corr,eng=g2,c2,fine
        while nr<o.max_receiver_points:
            next_nr=min(2*nr+1,int(o.max_receiver_points));cnext=eng.corrections(next_nr)
            rx_delta=_observable_change(gn,corr,gn,cnext);rc=rx_delta<=o.convergence_rtol
            nr,corr=next_nr,cnext
            if rc:break
    passed=fc and rc
    report=EGNConvergenceReport(passed,fc,rc,freq_delta,rx_delta,o.convergence_rtol,nr,final_o.panel_order,final_o.phase_step_rad,tuple(int(n) for n in counts),
        'Independent phase/panel and receiver refinements; empirical numerical estimates, not rigorous error bounds or paper/SSFM accuracy guarantees.' if o.verify_convergence else 'Exploratory calculation: convergence was not checked.')
    if o.verify_convergence and o.strict_convergence and not passed:
        raise EGNConvergenceError(f'EGN convergence target {o.convergence_rtol:g} not demonstrated: {report}')
    result={}
    for i,n in enumerate(counts):
        g=gn[:,i];c=corr[:,i];total=float(sum(g+c))
        if not np.all(np.isfinite(g+c)) or np.any(g+c < -1e-12*max(float(sum(g)),1e-30)) or total<0:
            raise EGNConvergenceError('Non-finite/negative NLI component; no clipping or replacement was applied')
        result[int(n)]=EGNBreakdown(mods[ci].name,float(sum(g)),*map(float,g),*map(float,c),total,
            ('total dual-polarization channel power','rectangular equal-baud spectra, beta2-only identical loss-compensated EDFA spans',
             'Appendix-B symmetric equal-power/common-modulation MCI scope' if len(mods)>=3 else 'SCI/XCI scope',
             'same GN physics; exact receiver/inner integration replaces stochastic GN integration in this precision path','no fitted constants'),report)
    return result


def full_egn_nli_power(system: WDMSystem, spans: Sequence[Span], cut_index: int,
                       modulation: str | ModulationFormat='QPSK',
                       channel_modulations: Optional[Sequence[str | ModulationFormat]]=None,
                       gn_options: GNIntegralOptions=GNIntegralOptions(),
                       full_options: EGNFullOptions=EGNFullOptions()) -> EGNBreakdown:
    """Precision replacement for final_EGN.py's full_egn_nli_power API.

    GNIntegralOptions is retained for call compatibility; its Sobol/order
    controls are not used here. Set precision controls with EGNFullOptions.
    """
    spans=tuple(spans)
    if not spans or any(s!=spans[0] for s in spans):
        raise ValueError('Precision Full-EGN requires one or more identical spans')
    if gn_options.accumulation!='coherent':raise ValueError('EGN requires coherent accumulation')
    return egn_span_sweep(system,spans[0],[len(spans)],cut_index,modulation,channel_modulations,full_options)[len(spans)]


@dataclass(frozen=True)
class FullPerformanceResult:
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
    nli_breakdown: EGNBreakdown


def evaluate_performance(system,spans,cut_index,modulation='QPSK',coding_rate=1.,
                         trx_snr_db=None,shannon_gap_db=0.,gn_options=GNIntegralOptions(),
                         receiver_points=7,nli_model='egn_full',channel_modulations=None,
                         full_egn_options=None):
    """GSNR/BER/rate adapter. The default now uses the precision Full-EGN path."""
    spans=tuple(spans);ch=system.channels[int(cut_index)];mod=get_modulation(modulation,coding_rate)
    br=None
    if nli_model=='egn_full':
        fo=full_egn_options or EGNFullOptions(receiver_points=receiver_points)
        br=full_egn_nli_power(system,spans,cut_index,modulation,channel_modulations,gn_options,fo)
        pn=br.total_egn_W
    elif nli_model in ('gn','egn_sci'):
        # Compatibility path: egn_sci uses the original fixed-order SCI routine.
        cs=list(system.channels);cs[int(cut_index)]=_channel_for_nli(ch,mod,nli_model)
        pn=integrate_nli_over_channel(replace(system,channels=tuple(cs)),spans,cut_index,gn_options,receiver_points)
    else:raise ValueError("nli_model must be gn, egn_sci or egn_full")
    ps=link_output_signal_power(system,spans,cut_index)
    pa=ase_noise_power_edfa(spans,ch.baud_GBd*1e9,ch_center_wavelength_nm(spans,ch))
    noise=pa+pn+(0. if trx_snr_db is None else ps/10**(float(trx_snr_db)/10))
    gs=ps/noise if noise else math.inf
    vals=(w_to_dbm(ch.power_W),ps,pa,pn,_db(ps/pa) if pa else math.inf,
          _db(ps/pn) if pn else math.inf,_db(gs),ber_awgn_approx(gs,modulation),
          mod.gross_dp_rate_gbps(ch.baud_GBd),mod.net_dp_rate_gbps(ch.baud_GBd),
          2*ch.baud_GBd*math.log2(1+gs/10**(float(shannon_gap_db)/10)))
    return FullPerformanceResult(*vals,br) if br is not None else PerformanceResult(*vals)


def launch_power_vs_gsnr(base_system,spans,cut_index,launch_power_dBm,modulation='QPSK',
                         coding_rate=1.,trx_snr_db=None,shannon_gap_db=0.,
                         gn_options=GNIntegralOptions(),receiver_points=7,
                         nli_model='egn_full',channel_modulations=None,full_egn_options=None,
                         use_cubic_scaling=True):
    """Common WDM power sweep; exact first-order P^3 scaling avoids repeated integrals."""
    grid=np.asarray(list(launch_power_dBm),float)
    if not len(grid) or not np.all(np.isfinite(grid)):raise ValueError('Power grid must be finite and nonempty')
    spans=tuple(spans);ci=int(cut_index);p0=base_system.channels[ci].power_W;pdb=w_to_dbm(p0)
    args=dict(modulation=modulation,coding_rate=coding_rate,trx_snr_db=trx_snr_db,shannon_gap_db=shannon_gap_db,
              gn_options=gn_options,receiver_points=receiver_points,nli_model=nli_model,
              channel_modulations=channel_modulations,full_egn_options=full_egn_options)
    ref=evaluate_performance(base_system,spans,ci,**args)
    out={k:[] for k in ['launch_power_dBm','gsnr_db','snr_ase_db','snr_nli_db','ber','p_nli_W','p_ase_W','cut_output_power_W','gross_rate_gbps','net_rate_gbps','shannon_gap_capacity_gbps']}
    ch=base_system.channels[ci];m=get_modulation(modulation,coding_rate);gap=10**(shannon_gap_db/10)
    for p in grid:
        scale=10**((p-pdb)/10)
        if not use_cubic_scaling:
            v=evaluate_performance(base_system.with_common_power_offset_db(float(p-pdb)),spans,ci,**args)
            for key in out:out[key].append(float(getattr(v,key)))
            continue
        ps=ref.cut_output_power_W*scale;pn=ref.p_nli_W*scale**3;pa=ref.p_ase_W
        noise=pa+pn+(0 if trx_snr_db is None else ps/10**(trx_snr_db/10));gs=ps/noise if noise else math.inf
        vals=(p,_db(gs),_db(ps/pa) if pa else math.inf,_db(ps/pn) if pn else math.inf,
              ber_awgn_approx(gs,modulation),pn,pa,ps,m.gross_dp_rate_gbps(ch.baud_GBd),m.net_dp_rate_gbps(ch.baud_GBd),2*ch.baud_GBd*math.log2(1+gs/gap))
        for key,val in zip(out,vals):out[key].append(float(val))
    return {k:np.asarray(v) for k,v in out.items()}


simulate_snr_curve=launch_power_vs_gsnr
__all__=['Channel','WDMSystem','Span','GNIntegralOptions','ModulationFormat','EGNFullOptions',
         'EGNConvergenceReport','EGNConvergenceError','EGNBreakdown','egn_span_sweep',
         'full_egn_nli_power','evaluate_performance','launch_power_vs_gsnr','simulate_snr_curve',
         'get_modulation','constellation_symbols','constellation_moments','MODULATION_PHI',
         'gn_nli_psd_qmc','integrate_nli_over_channel','dbm_to_w','w_to_dbm','beta2_from_D']


if __name__=='__main__':
    import argparse,json
    from dataclasses import asdict
    parser=argparse.ArgumentParser(description='Convergence-checked Carena EGN demonstration')
    parser.add_argument('--demo',action='store_true')
    parser.add_argument('--channels',type=int,default=3)
    parser.add_argument('--spans',type=int,default=5)
    parser.add_argument('--modulation',default='QPSK')
    args=parser.parse_args()
    if not args.demo:parser.print_help()
    else:
        sys0=WDMSystem.equispaced(args.channels,33.6,32.,0.)
        span0=Span(100.,.22,1.3,D_ps_nm_km=16.7)
        r=full_egn_nli_power(sys0,[span0]*args.spans,args.channels//2,args.modulation)
        payload=asdict(r);payload['eta_xmci_W_inv2']=r.egn_xmci_W/sys0.channels[args.channels//2].power_W**3
        print(json.dumps(payload,ensure_ascii=False,indent=2))
