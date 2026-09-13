"""final_EGN.py — GN + modulation-aware Enhanced GN (EGN) research engine.

Foundation
----------
* GN baseline: Poggiolini et al., ``A Detailed Analytical Derivation of the
  GN Model of Non-Linear Interference in Coherent Optical Transmission
  Systems`` (arXiv:1209.0394 v13), including the dual-polarization 16/27
  GNRF, beta2/beta3, coherent non-identical-span link accumulation, optional
  DCU, and numerical distributed-gain/Raman profiles.
* EGN: Carena et al., Optics Express 22, 16335-16362 (2014). SCI uses
  Eqs. (5)-(9) with Appendix-C reductions (43)-(45). Full-WDM XCI correction
  uses Eq. (18) and Appendix A, Eqs. (30),(32),(34),(35), evaluated through
  factorized two-dimensional quadratures. MCI correction uses Appendix B
  Eqs. (36)-(41), likewise reduced to nested one-dimensional quadratures.

Design rule
-----------
The original GN numerical path is preserved. ``egn_full`` is computed as
GN baseline + SCI non-Gaussian correction + XCI non-Gaussian correction +
MCI non-Gaussian correction. No fit, empirical offset, or paper-target tuning
is used.

Supported full-EGN scope
------------------------
* dual-polarization Manakov convention; channel powers are total DP powers;
* coherent accumulation; rectangular zero-phase spectra; equal symbol rates;
* XCI: arbitrary non-overlapping INT locations/powers and per-channel
  modulation moments;
* MCI Appendix-B path: odd, symmetric, equally spaced WDM comb with equal
  channel powers and a common modulation format, exactly matching Eq. (37);
* BPSK/QPSK/8QAM/16QAM/32QAM/64QAM/256QAM constellation moments are
  calculated from explicit symbol coordinates (8/32QAM geometries are the
  documented modeling choices in this file).

Limits
------
* Carena's printed full-EGN equations are for rectangular spectra. RRC with
  non-zero roll-off and custom spectral phase stay available for the GN
  baseline but are rejected by ``egn_full`` rather than silently approximated.
* Mid-link channel add/drop is not modeled, consistently with the limitation
  discussed around Eq. (100) in the GN derivation.
* Heterogeneous/distributed-Raman GN propagation is numerically supported,
  but full-EGN validation for such links is less established than the
  homogeneous EDFA cases used by Carena; validate against SSFM before design
  sign-off.
* Near-full dispersion compensation can violate the underlying GN/EGN
  assumptions; this code does not turn such a result into a guaranteed model.

The module is self-contained and does not require the original three .py
files at runtime.
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


def egn_literature_self_test(
    frequency_order: int = 512,
    receiver_points: int = 31,
    tolerance_db: float = 0.2,
    strict: bool = False,
) -> dict:
    """External SCI benchmark against Carena (2014), Fig. 1, 50 spans.

    References are extracted from the supplied publisher PDF, p. 8/16342,
    not generated by this engine. Green/blue/red curves are EGN/GN/SSFM.
    The source vector endpoints and axis calibration are retained below so
    the extraction is reproducible. Stroke/axis precision is about 0.1 dB.
    Source SHA256: 2bec32a57d69cb092b50854706fb78d85ac656514580696a45e60f0ad773aec8
    DOI: 10.1364/OE.22.016335. Parameters: Sect. 3, p. 16343.

    Theory uses rectangular spectra, as in the paper's stated derivation;
    SSFM used RC rolloff=0.05. Wavelength for D->beta2 is assumed 1550 nm.
    Comparisons to EGN theory and SSFM are reported separately. strict=True
    raises if either exceeds tolerance; a failed SSFM comparison is not hidden.
    This does not validate XCI/MCI or all modulation formats.

    Dar (2013) Fig. 5(b)'s ~6.5 dB is NOT an SCI reference: single polarization,
    ideal distributed gain, 500 km, SCI back-propagated away. It is not fitted.
    """
    # (D, gamma, y_top, y_bottom, eta_top, eta_bottom, y_GN, y_EGN, y_SSFM,
    #  text-reported approximate GN-to-simulation gap on p. 16343)
    source = {
        "SMF": (16.7, 1.3, 77.52001953125, 217.32000732421875, 45.0, 15.0,
                87.84280395507812, 94.2622299194336, 92.10208892822266, 1.1),
        "NZDSF": (3.8, 1.5, 250.32000732421875, 390.1199951171875, 50.0, 15.0,
                  255.4815216064453, 263.2221984863281, 263.64166259765625, 2.1),
        "LS": (-1.8, 2.2, 423.1199951171875, 562.8599853515625, 55.0, 20.0,
               424.8609313964844, 434.2813720703125, 435.9611511230469, 2.8),
    }
    ch = Channel(0.0, 1e-3, 32.0)
    opt = GNIntegralOptions(egn_frequency_order=frequency_order)
    rows = []
    for name, data in source.items():
        d, gamma, yt, yb, et, eb, ygn, yegn, yssfm, text_gap = data
        reference = [et + (y - yt) * (eb - et) / (yb - yt)
                     for y in (ygn, yegn, yssfm)]
        spans = [Span(100.0, 0.22, gamma, D_ps_nm_km=d, wavelength_nm=1550.0)] * 50
        c = integrate_egn_sci_coefficients(ch, spans, opt, receiver_points)
        actual_gn = 10.0 * math.log10(c.kappa1)
        actual_egn = 10.0 * math.log10(c.corrected(1.0, 1.0))
        actual_gap = actual_gn - actual_egn
        rows.append({
            "fiber": name, "calculated_gn_eta_db": actual_gn,
            "calculated_egn_eta_db": actual_egn,
            "reference_gn_eta_db": reference[0], "reference_egn_eta_db": reference[1],
            "reference_ssfm_eta_db": reference[2],
            "sci_multiplier": c.sci_multiplier(1.0, 1.0),
            "calculated_gn_to_egn_gap_db": actual_gap,
            "reference_gn_to_egn_gap_db": reference[0] - reference[1],
            "egn_eta_error_db": abs(actual_egn - reference[1]),
            "gn_eta_error_db": abs(actual_gn - reference[0]),
            "nli_ratio_error_db": abs(actual_gap - (reference[0] - reference[1])),
            "ssfm_eta_error_db": abs(actual_egn - reference[2]),
            "text_reported_gap_db": text_gap,
            "text_gap_error_db": abs(actual_gap - text_gap),
        })
    max_equation_error = max(max(r["egn_eta_error_db"], r["gn_eta_error_db"],
                                 r["nli_ratio_error_db"]) for r in rows)
    max_ssfm_error = max(r["ssfm_eta_error_db"] for r in rows)
    equation_pass = max_equation_error <= tolerance_db
    ssfm_pass = max_ssfm_error <= tolerance_db
    result = {"passed": equation_pass and ssfm_pass,
              "equation_reproduction_passed": equation_pass,
              "ssfm_comparison_passed": ssfm_pass,
              "tolerance_db": float(tolerance_db),
              "max_equation_error_db": max_equation_error,
              "max_ssfm_error_db": max_ssfm_error,
              "frequency_order": frequency_order, "receiver_points": receiver_points,
              "rows": rows}
    if strict and not result["passed"]:
        raise AssertionError(f"Literature comparison exceeds {tolerance_db} dB: {result}")
    return result


__all__ = [
    "Channel", "WDMSystem", "Span", "GNIntegralOptions", "GNResult",
    "dbm_to_w", "w_to_dbm", "alpha_field_from_db", "beta2_from_D",
    "beta_of_f_offset", "phase_mismatch_beta23", "link_kernel",
    "gn_nli_psd_qmc", "gn_nli_psd_multi_seed", "integrate_nli_over_channel",
    "common_power_scaling_self_test", "egn_backward_compat_self_test",
    "modulation_phi_self_test", "egn_literature_self_test", "DP_GN_COEFF",
    "EGNSCICoefficients", "egn_sci_coefficients", "integrate_egn_sci_coefficients",
]


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






# =============================================================================
# Full-WDM EGN extension (Carena 2014 Eq. 18, Appendices A-B)
# =============================================================================

FullNLIModel = Literal["gn", "egn_sci", "egn_full"]

@dataclass(frozen=True)
class EGNFullOptions:
    """Options specific to the full-WDM EGN correction path.

    ``frequency_order`` overrides ``GNIntegralOptions.egn_frequency_order``
    when not None.  Values 48-128 are useful for exploratory runs; convergence
    studies should increase the order until the requested observable is stable.
    """
    frequency_order: Optional[int] = None
    receiver_points: int = 7
    symmetry_rtol: float = 2e-8
    power_rtol: float = 2e-8
    strict_mci: bool = True


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

    @property
    def egn_sci_W(self) -> float:
        return self.sci_gn_W + self.sci_correction_W

    @property
    def egn_xci_W(self) -> float:
        return self.xci_gn_W + self.xci_correction_W

    @property
    def egn_mci_W(self) -> float:
        return self.mci_gn_W + self.mci_correction_W

    @property
    def correction_W(self) -> float:
        return self.sci_correction_W + self.xci_correction_W + self.mci_correction_W

    @property
    def ratio_to_gn(self) -> float:
        return self.total_egn_W / self.gn_total_W if self.gn_total_W > 0 else math.inf

    @property
    def ratio_to_gn_db(self) -> float:
        r = self.ratio_to_gn
        return 10.0 * math.log10(r) if r > 0 and math.isfinite(r) else (-math.inf if r == 0 else math.inf)


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


def _as_modulation_list(system: WDMSystem, modulation: str | ModulationFormat,
                        channel_modulations: Optional[Sequence[str | ModulationFormat]]) -> list[ModulationFormat]:
    if channel_modulations is None:
        m = modulation if isinstance(modulation, ModulationFormat) else get_modulation(modulation)
        return [m for _ in system.channels]
    if len(channel_modulations) != len(system.channels):
        raise ValueError("channel_modulations must have one entry per WDM channel")
    return [x if isinstance(x, ModulationFormat) else get_modulation(str(x)) for x in channel_modulations]


def _validate_full_egn_spectral_scope(system: WDMSystem, spans: Sequence[Span], cut_index: int,
                                      options: GNIntegralOptions) -> float:
    if options.accumulation != "coherent":
        raise ValueError("Carena full EGN requires coherent accumulation")
    if not spans:
        raise ValueError("at least one span is required")
    rs = float(system.channels[int(cut_index)].symbol_rate_THz)
    if rs <= 0 or not math.isfinite(rs):
        raise ValueError("CUT symbol rate must be finite and positive")
    for ch in system.channels:
        if not (ch.pulse_shape == "rect" or
                (ch.pulse_shape in ("rrc", "raised_cosine") and float(ch.rolloff) == 0.0)):
            raise ValueError("egn_full implements the Carena rectangular-spectrum equations only")
        if not math.isclose(float(ch.symbol_rate_THz), rs, rel_tol=1e-12, abs_tol=1e-15):
            raise ValueError("egn_full currently requires equal symbol rates, as in Carena Eqs. (18),(37)-(41)")
    ordered = sorted(system.channels, key=lambda c: c.center_THz)
    for a, b in zip(ordered[:-1], ordered[1:]):
        if float(b.center_THz) - float(a.center_THz) < a.support_half_width_THz + b.support_half_width_THz - 1e-14:
            raise ValueError("egn_full requires non-overlapping rectangular channel supports")
    return rs


def _gl_interval(lo: float, hi: float, order: int):
    if hi <= lo:
        return np.empty(0), np.empty(0)
    x, w = roots_legendre(int(order))
    return lo + (x + 1.0) * (hi - lo) / 2.0, w * (hi - lo) / 2.0


def _intersect(a: tuple[float, float], b: tuple[float, float]) -> tuple[float, float]:
    return max(a[0], b[0]), min(a[1], b[1])


def _norm_interval(ch: Channel, cut: Channel, rs: float) -> tuple[float, float]:
    d = (float(ch.center_THz) - float(cut.center_THz)) / rs
    return d - 0.5, d + 0.5


def _full_order(options: GNIntegralOptions, full_options: EGNFullOptions) -> int:
    n = int(full_options.frequency_order if full_options.frequency_order is not None
            else options.egn_frequency_order)
    if n < 16:
        raise ValueError("full-EGN frequency_order must be >= 16")
    return n


def _reduced_fixed_outer(A, B, C, t, cut_center, rs, spans, options, order):
    """Appendix-C factorization: integral_A |integral_{B,C} mu du2|^2 du1."""
    u1, w1 = _gl_interval(A[0], A[1], order)
    if u1.size == 0:
        return 0.0
    total = 0.0
    for x, wx in zip(u1, w1):
        lo, hi = _intersect(B, (t + C[0] - x, t + C[1] - x))
        u2, w2 = _gl_interval(lo, hi, order)
        if u2.size == 0:
            continue
        amp = _egn_link_function(cut_center + rs*x, cut_center + rs*u2,
                                 cut_center + rs*t, spans, options.z_quadrature_order)
        inner = np.sum(w2 * amp)
        total += wx * abs(inner)**2
    return float(total)


def _reduced_fixed_sum(A, B, C, t, cut_center, rs, spans, options, order):
    """Appendix-C change of variables: integral_q |integral mu(u,q-u) du|^2 dq."""
    qlo, qhi = _intersect((A[0]+B[0], A[1]+B[1]), (t+C[0], t+C[1]))
    q, wq = _gl_interval(qlo, qhi, order)
    if q.size == 0:
        return 0.0
    total = 0.0
    for qq, ww in zip(q, wq):
        lo, hi = _intersect(A, (qq-B[1], qq-B[0]))
        u1, w1 = _gl_interval(lo, hi, order)
        if u1.size == 0:
            continue
        u2 = qq-u1
        amp = _egn_link_function(cut_center + rs*u1, cut_center + rs*u2,
                                 cut_center + rs*t, spans, options.z_quadrature_order)
        inner = np.sum(w1 * amp)
        total += ww * abs(inner)**2
    return float(total)


def _reduced_pair_square(A, B, C, t, cut_center, rs, spans, options, order):
    """Factor a four-fold term into |double integral mu|^2 (Appendix-C logic)."""
    u1, w1 = _gl_interval(A[0], A[1], order)
    z = 0.0j
    for x, wx in zip(u1, w1):
        lo, hi = _intersect(B, (t+C[0]-x, t+C[1]-x))
        u2, w2 = _gl_interval(lo, hi, order)
        if u2.size == 0:
            continue
        amp = _egn_link_function(cut_center + rs*x, cut_center + rs*u2,
                                 cut_center + rs*t, spans, options.z_quadrature_order)
        z += wx * np.sum(w2 * amp)
    return float(abs(z)**2)


def _xci_correction_psd(system: WDMSystem, spans: Sequence[Span], cut_index: int,
                        f_THz: float, mods: Sequence[ModulationFormat],
                        options: GNIntegralOptions, full_options: EGNFullOptions) -> float:
    """Carena Eq. (18) non-Gaussian XCI correction, Appendix-A reduced form.

    The GN-like kappa_m1 pieces are already present in the preserved full-WDM
    GN baseline.  This routine adds only Phi/Psi correction pieces k12, k22,
    k32, k42 and k43, preventing double counting.
    """
    ci = int(cut_index)
    cut = system.channels[ci]
    rs = float(cut.symbol_rate_THz)
    center = float(cut.center_THz)
    t = (float(f_THz)-center)/rs
    if abs(t) > 0.5 + 1e-12:
        return 0.0
    order = _full_order(options, full_options)
    A0 = (-0.5, 0.5)
    pc = float(cut.power_W)
    phi_a = float(mods[ci].egn_phi)
    corr = 0.0
    for j, inte in enumerate(system.channels):
        if j == ci:
            continue
        I = _norm_interval(inte, cut, rs)
        k12 = 80.0/(81.0*rs) * _reduced_fixed_outer(A0, I, I, t, center, rs, spans, options, order)
        k22 = 80.0/(81.0*rs) * _reduced_fixed_outer(I, A0, A0, t, center, rs, spans, options, order)
        k32 = 16.0/(81.0*rs) * _reduced_fixed_sum(A0, A0, I, t, center, rs, spans, options, order)
        k42 = (80.0/(81.0*rs) * _reduced_fixed_outer(I, I, I, t, center, rs, spans, options, order)
               + 16.0/(81.0*rs) * _reduced_fixed_sum(I, I, I, t, center, rs, spans, options, order))
        k43 = 16.0/(81.0*rs) * _reduced_pair_square(I, I, I, t, center, rs, spans, options, order)
        pi = float(inte.power_W)
        phi_b = float(mods[j].egn_phi)
        psi_b = float(mods[j].egn_psi)
        corr += (pc*pi*pi*phi_b*k12
                 + pc*pc*pi*phi_a*(k22+k32)
                 + pi**3*(phi_b*k42 + psi_b*k43))
    return float(corr)


def _symmetric_grid_info(system: WDMSystem, cut_index: int, mods: Sequence[ModulationFormat],
                         rtol: float, power_rtol: float):
    ci = int(cut_index)
    cut = system.channels[ci]
    n = len(system.channels)
    if n % 2 != 1:
        raise ValueError("Appendix-B MCI Eq. (37) requires an odd channel count")
    centers = np.asarray([float(ch.center_THz) for ch in system.channels])
    order = np.argsort(centers)
    pos = int(np.where(order == ci)[0][0])
    H = (n-1)//2
    if pos != H:
        raise ValueError("Appendix-B MCI requires CUT at the center of a symmetric comb")
    sorted_centers = centers[order]
    diffs = np.diff(sorted_centers)
    spacing = float(np.mean(diffs)) if diffs.size else 0.0
    if diffs.size and not np.allclose(diffs, spacing, rtol=rtol, atol=max(abs(spacing)*rtol,1e-14)):
        raise ValueError("Appendix-B MCI requires equally spaced channels")
    rel = sorted_centers - float(cut.center_THz)
    target = np.arange(-H,H+1)*spacing
    if not np.allclose(rel,target,rtol=rtol,atol=max(abs(spacing)*rtol,1e-14)):
        raise ValueError("Appendix-B MCI requires spectral symmetry about the CUT")
    powers = np.asarray([float(system.channels[k].power_W) for k in order])
    if not np.allclose(powers,powers[H],rtol=power_rtol,atol=max(abs(powers[H])*power_rtol,1e-18)):
        raise ValueError("Appendix-B MCI Eq. (37) requires equal per-channel power")
    mu4 = np.asarray([mods[k].mu4 for k in order],float)
    mu6 = np.asarray([mods[k].mu6 for k in order],float)
    if not (np.allclose(mu4,mu4[H],rtol=1e-12,atol=1e-14) and np.allclose(mu6,mu6[H],rtol=1e-12,atol=1e-14)):
        raise ValueError("Appendix-B MCI Eq. (38) path currently requires a common modulation distribution")
    by_index = {i:int(order[i+H]) for i in range(-H,H+1)}
    return H, spacing, by_index, float(powers[H]), float(mods[ci].egn_phi)


def _mci_correction_psd(system: WDMSystem, spans: Sequence[Span], cut_index: int,
                        f_THz: float, mods: Sequence[ModulationFormat],
                        options: GNIntegralOptions, full_options: EGNFullOptions) -> float:
    """Carena Appendix-B Eq. (38)-(41) MCI correction.

    This is deliberately strict about Eq. (37): odd/symmetric/equal-spacing,
    equal launch power and common modulation.  The factor 2 in Eqs. (39)-(41)
    accounts for the mirrored negative-frequency regions.
    """
    if len(system.channels) < 3:
        return 0.0
    H, spacing, byidx, pch, phi_b = _symmetric_grid_info(
        system, cut_index, mods, full_options.symmetry_rtol, full_options.power_rtol)
    if H < 1:
        return 0.0
    cut = system.channels[int(cut_index)]
    rs=float(cut.symbol_rate_THz); center=float(cut.center_THz); t=(float(f_THz)-center)/rs
    order=_full_order(options,full_options)
    def R(idx): return _norm_interval(system.channels[byidx[int(idx)]],cut,rs)
    km1=0.0
    for n in range(1,H+1):
        km1 += 2.0*80.0/(81.0*rs)*_reduced_fixed_outer(R(-1),R(n),R(n),t,center,rs,spans,options,order)
    km2=0.0
    for n in range(2,H+1):
        km2 += 2.0*80.0/(81.0*rs)*_reduced_fixed_outer(R(1),R(n),R(n),t,center,rs,spans,options,order)
    km3=0.0
    for n in range(2,H+1):
        ms=[n//2] if n%2==0 else [(n-1)//2,(n+1)//2]
        for m in ms:
            km3 += 2.0*16.0/(81.0*rs)*_reduced_fixed_sum(R(m),R(m),R(n),t,center,rs,spans,options,order)
    return float(phi_b * pch**3 * (km1+km2+km3))


def _gn_breakdown_power_qmc(system: WDMSystem, spans: Sequence[Span], cut_index: int,
                            options: GNIntegralOptions, receiver_points: int = 7) -> tuple[float,float,float]:
    """GN SCI/XCI/MCI geometric breakdown using common Sobol samples.

    SCI: all three mixing frequencies in CUT. XCI: frequencies originate from
    CUT plus exactly one interferer. MCI: every remaining active GN triplet.
    Their sum is a numerical diagnostic; the authoritative GN total remains
    ``integrate_nli_over_channel`` for backward compatibility.
    """
    ci=int(cut_index); cut=system.channels[ci]
    x,w=roots_legendre(int(receiver_points)); half=cut.symbol_rate_THz/2.0
    out=np.zeros(3,float)
    lo,hi=system.support_bounds_THz(); width=hi-lo
    for k,(xx,ww) in enumerate(zip(x,w)):
        f=float(cut.center_THz+half*xx)
        sob=qmc.Sobol(d=2,scramble=True,seed=int(options.seed)+1009*k)
        u=sob.random_base2(int(options.sobol_power)); f1=lo+width*u[:,0]; f2=lo+width*u[:,1]; f3=f1+f2-f
        G1=system.psd(f1);G2=system.psd(f2);G3=system.psd(f3);active=(G1>0)&(G2>0)&(G3>0)
        if not np.any(active): continue
        K=np.zeros_like(f1);K[active]=link_kernel(f1[active],f2[active],f,spans,options.accumulation,options.z_quadrature_order)
        base=DP_GN_COEFF*G1*G2*G3*K
        def idx_of(v):
            z=np.full(v.shape,-1,int)
            for jj,ch in enumerate(system.channels):
                h=ch.support_half_width_THz
                z[(v>=float(ch.center_THz)-h)&(v<=float(ch.center_THz)+h)]=jj
            return z
        i1=idx_of(f1);i2=idx_of(f2);i3=idx_of(f3)
        sci=(i1==ci)&(i2==ci)&(i3==ci)
        xci=np.zeros_like(sci)
        for jj in range(len(system.channels)):
            if jj==ci: continue
            mask=np.isin(i1,[ci,jj])&np.isin(i2,[ci,jj])&np.isin(i3,[ci,jj]) & ((i1==jj)|(i2==jj)|(i3==jj))
            xci |= mask
        mci=active & ~sci & ~xci
        fac=half*ww*width**2
        out += fac*np.array([np.mean(base*sci),np.mean(base*xci),np.mean(base*mci)])
    return tuple(float(v) for v in out)


def full_egn_nli_power(system: WDMSystem, spans: Sequence[Span], cut_index: int,
                       modulation: str | ModulationFormat = "QPSK",
                       channel_modulations: Optional[Sequence[str | ModulationFormat]] = None,
                       gn_options: GNIntegralOptions = GNIntegralOptions(),
                       full_options: EGNFullOptions = EGNFullOptions()) -> EGNBreakdown:
    """Return GN + SCI/XCI/MCI full-EGN channel NLI power.

    This routine is the main research API. It never modifies the GN baseline.
    Non-Gaussian corrections are added separately, exactly matching the EGN
    decomposition ``G_EGN = G_GN + G_corr``.  No fitted constants are used.
    """
    spans=tuple(spans);ci=int(cut_index)
    rs=_validate_full_egn_spectral_scope(system,spans,ci,gn_options)
    mods=_as_modulation_list(system,modulation,channel_modulations)
    gn_total=float(integrate_nli_over_channel(system,spans,ci,gn_options,full_options.receiver_points))
    # SCI Eq. (5): add only non-Gaussian part to GN baseline.
    sci_c=integrate_egn_sci_coefficients(system.channels[ci],spans,gn_options,full_options.receiver_points)
    mc=mods[ci]
    sci_corr=float(system.channels[ci].power_W**3*(mc.egn_phi*sci_c.kappa2+mc.egn_psi*sci_c.kappa3))
    # Receiver integration of XCI/MCI correction PSD.
    xr,wr=roots_legendre(int(full_options.receiver_points));half=rs/2.0
    xci_corr=0.0;mci_corr=0.0
    mci_supported=True;mci_error=None
    for xx,ww in zip(xr,wr):
        ff=float(system.channels[ci].center_THz+half*xx)
        xci_corr += half*ww*_xci_correction_psd(system,spans,ci,ff,mods,gn_options,full_options)
        if len(system.channels)>=3:
            try:
                mci_corr += half*ww*_mci_correction_psd(system,spans,ci,ff,mods,gn_options,full_options)
            except ValueError as exc:
                mci_supported=False;mci_error=str(exc);break
    if not mci_supported:
        if full_options.strict_mci:
            raise ValueError("Full MCI correction outside Appendix-B Eq.(37) scope: "+str(mci_error))
        mci_corr=0.0
    sci_gn,xci_gn,mci_gn=_gn_breakdown_power_qmc(system,spans,ci,gn_options,full_options.receiver_points)
    total=float(gn_total+sci_corr+xci_corr+mci_corr)
    assumptions=("Poggiolini-1209.0394 GN baseline preserved",
                 "Carena-2014 rectangular-spectrum SCI/XCI corrections",
                 "Carena Appendix-B Eq.(37) MCI symmetry/equal-power scope" if len(system.channels)>=3 else "single/two-channel: no MCI",
                 "no fitting or empirical NLI offset")
    return EGNBreakdown(mods[ci].name,gn_total,sci_gn,xci_gn,mci_gn,
                        float(sci_corr),float(xci_corr),float(mci_corr),total,assumptions)


def evaluate_performance(system: WDMSystem, spans: Sequence[Span], cut_index: int,
                         modulation: str = "QPSK", coding_rate: float = 1.0,
                         trx_snr_db: Optional[float] = None, shannon_gap_db: float = 0.0,
                         gn_options: GNIntegralOptions = GNIntegralOptions(), receiver_points: int = 7,
                         nli_model: FullNLIModel = "gn",
                         channel_modulations: Optional[Sequence[str | ModulationFormat]] = None,
                         full_egn_options: Optional[EGNFullOptions] = None):
    """Backward-compatible GN/SCI-EGN API plus ``nli_model='egn_full'``."""
    if nli_model in ("gn","egn_sci"):
        ch=system.channels[int(cut_index)];p_sig=link_output_signal_power(system,spans,cut_index)
        mod=get_modulation(modulation,coding_rate)
        ch2=_channel_for_nli(ch,mod,nli_model);chs=list(system.channels);chs[int(cut_index)]=ch2
        sys2=replace(system,channels=tuple(chs))
        p_nli=integrate_nli_over_channel(sys2,spans,cut_index,gn_options,receiver_points)
        p_ase=ase_noise_power_edfa(spans,ch.baud_GBd*1e9,ch_center_wavelength_nm(spans,ch))
        noise=p_ase+p_nli
        if trx_snr_db is not None: noise += p_sig/(10**(float(trx_snr_db)/10.0))
        gsnr=p_sig/noise if noise>0 else math.inf
        snra=p_sig/p_ase if p_ase>0 else math.inf;snrn=p_sig/p_nli if p_nli>0 else math.inf
        gap=10**(float(shannon_gap_db)/10.0);cap=2.0*ch.baud_GBd*math.log2(1.0+gsnr/gap)
        return PerformanceResult(w_to_dbm(ch.power_W),p_sig,p_ase,p_nli,_db(snra),_db(snrn),_db(gsnr),
                                 ber_awgn_approx(gsnr,modulation),mod.gross_dp_rate_gbps(ch.baud_GBd),
                                 mod.net_dp_rate_gbps(ch.baud_GBd),cap)
    if nli_model!="egn_full": raise ValueError("nli_model must be 'gn', 'egn_sci', or 'egn_full'")
    fo=full_egn_options or EGNFullOptions(receiver_points=receiver_points)
    br=full_egn_nli_power(system,spans,cut_index,modulation,channel_modulations,gn_options,fo)
    ch=system.channels[int(cut_index)];mod=get_modulation(modulation,coding_rate)
    p_sig=link_output_signal_power(system,spans,cut_index)
    p_ase=ase_noise_power_edfa(spans,ch.baud_GBd*1e9,ch_center_wavelength_nm(spans,ch))
    noise=p_ase+br.total_egn_W
    if trx_snr_db is not None: noise += p_sig/(10**(float(trx_snr_db)/10.0))
    gsnr=p_sig/noise if noise>0 else math.inf
    gap=10**(float(shannon_gap_db)/10.0);cap=2.0*ch.baud_GBd*math.log2(1.0+gsnr/gap)
    return FullPerformanceResult(w_to_dbm(ch.power_W),p_sig,p_ase,br.total_egn_W,
                                 _db(p_sig/p_ase) if p_ase>0 else math.inf,
                                 _db(p_sig/br.total_egn_W) if br.total_egn_W>0 else math.inf,
                                 _db(gsnr),ber_awgn_approx(gsnr,modulation),
                                 mod.gross_dp_rate_gbps(ch.baud_GBd),mod.net_dp_rate_gbps(ch.baud_GBd),cap,br)


def full_egn_self_test(fast: bool = True) -> dict:
    """Physics/regression smoke tests; no literature fitting is performed."""
    order=20 if fast else 64
    sob=10 if fast else 14
    opt=GNIntegralOptions(sobol_power=sob,seed=7,egn_frequency_order=order)
    fo=EGNFullOptions(frequency_order=order,receiver_points=3 if fast else 7)
    # GN regression and P^3 invariance use the preserved original tests.
    reg=egn_backward_compat_self_test();p3=common_power_scaling_self_test();mom=constellation_moment_self_test()
    # Single-channel full EGN must equal the existing SCI-EGN construction.
    sys1=WDMSystem.equispaced(1,50.0,32.0,-3.0)
    spans=[Span(80.0,0.2,1.3,D_ps_nm_km=17.0)]
    br1=full_egn_nli_power(sys1,spans,0,"QPSK",gn_options=opt,full_options=fo)
    c=integrate_egn_sci_coefficients(sys1.channels[0],spans,opt,fo.receiver_points)
    q=get_modulation("QPSK");ref=sys1.channels[0].power_W**3*c.corrected(q.mu4,q.mu6)
    single_rel=abs(br1.total_egn_W-ref)/max(abs(ref),np.finfo(float).tiny)
    # Three-channel XCI/MCI and nine-channel general MCI structural checks.
    sys3=WDMSystem.equispaced(3,33.6,32.0,-6.0)
    br3=full_egn_nli_power(sys3,spans,1,"QPSK",gn_options=opt,full_options=fo)
    sys9=WDMSystem.equispaced(9,33.6,32.0,-8.0)
    br9=full_egn_nli_power(sys9,spans,4,"QPSK",gn_options=opt,full_options=fo)
    finite=all(np.isfinite([br3.total_egn_W,br9.total_egn_W,br3.xci_correction_W,br9.mci_correction_W]))
    return {"passed":bool(reg["passed"] and finite and single_rel < (2e-2 if fast else 5e-3)),
            "gn_backward_compat":reg,"p3":p3,"constellation_moments":mom,
            "single_channel_full_vs_sci_rel_error":single_rel,
            "three_channel":br3,"nine_channel":br9,
            "note":"Fast mode is a smoke test. Increase Sobol/quadrature orders for publication-quality convergence."}


def literature_sci_check(*args, **kwargs):
    """Expose the pre-existing Carena Fig.1 50-span benchmark without fitting."""
    return egn_literature_self_test(*args, **kwargs)


# Public names intentionally include both the original engine API and the new full-EGN API.
__all__ = [name for name in globals() if not name.startswith("_")]

# Override the legacy sweep so the final public API also accepts egn_full.
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
    nli_model: FullNLIModel = "gn",
    channel_modulations: Optional[Sequence[str | ModulationFormat]] = None,
    full_egn_options: Optional[EGNFullOptions] = None,
) -> dict:
    """Common launch-power sweep for GN, SCI-EGN, or full EGN.

    When all WDM powers are shifted by the same dB offset, first-order GN/EGN
    NLI obeys the exact cubic law, so a single reference integration can be
    reused. Set ``use_cubic_scaling=False`` to recompute each point.
    """
    pgrid=np.asarray(list(launch_power_dBm),float)
    if pgrid.size==0: raise ValueError("launch_power_dBm is empty")
    cut=base_system.channels[int(cut_index)]; p0=float(cut.power_W); p0db=w_to_dbm(p0)
    p_sig0=link_output_signal_power(base_system,spans,cut_index)
    p_ase=ase_noise_power_edfa(spans,cut.baud_GBd*1e9,ch_center_wavelength_nm(spans,cut))
    fo=full_egn_options or EGNFullOptions(receiver_points=receiver_points)
    if nli_model=="egn_full":
        ref=full_egn_nli_power(base_system,spans,cut_index,modulation,channel_modulations,gn_options,fo)
        p_nli0=ref.total_egn_W
    else:
        mod=get_modulation(modulation,coding_rate);chs=list(base_system.channels)
        chs[int(cut_index)]=_channel_for_nli(cut,mod,nli_model)
        p_nli0=integrate_nli_over_channel(replace(base_system,channels=tuple(chs)),spans,cut_index,gn_options,receiver_points)
    mod=get_modulation(modulation,coding_rate);gap=10**(float(shannon_gap_db)/10.0)
    out={k:[] for k in ("launch_power_dBm","gsnr_db","snr_ase_db","snr_nli_db","ber","p_nli_W","p_ase_W","cut_output_power_W","gross_rate_gbps","net_rate_gbps","shannon_gap_capacity_gbps")}
    for pp in pgrid:
        scale=10**((float(pp)-p0db)/10.0)
        if use_cubic_scaling:
            ps=p_sig0*scale; pn=p_nli0*scale**3
        else:
            sysi=base_system.with_common_power_offset_db(float(pp)-p0db)
            ps=link_output_signal_power(sysi,spans,cut_index)
            if nli_model=="egn_full": pn=full_egn_nli_power(sysi,spans,cut_index,modulation,channel_modulations,gn_options,fo).total_egn_W
            else:
                mi=get_modulation(modulation,coding_rate);cc=list(sysi.channels);cc[int(cut_index)]=_channel_for_nli(cc[int(cut_index)],mi,nli_model)
                pn=integrate_nli_over_channel(replace(sysi,channels=tuple(cc)),spans,cut_index,gn_options,receiver_points)
        noise=p_ase+pn
        if trx_snr_db is not None: noise+=ps/(10**(float(trx_snr_db)/10.0))
        gs=math.inf if noise==0 else ps/noise; sa=math.inf if p_ase==0 else ps/p_ase; sn=math.inf if pn==0 else ps/pn
        sh=math.inf if math.isinf(gs) else 2*cut.baud_GBd*math.log2(1+gs/gap)
        vals=(float(pp),_db(gs) if not math.isinf(gs) else math.inf,_db(sa) if not math.isinf(sa) else math.inf,_db(sn) if not math.isinf(sn) else math.inf,0.0 if math.isinf(gs) else ber_awgn_approx(gs,mod.name),pn,p_ase,ps,mod.gross_dp_rate_gbps(cut.baud_GBd),mod.net_dp_rate_gbps(cut.baud_GBd),sh)
        for k,v in zip(out,vals): out[k].append(float(v))
    return {k:np.asarray(v,float) for k,v in out.items()}


def simulate_snr_curve(*args, **kwargs):
    return launch_power_vs_gsnr(*args, **kwargs)


VALIDATION_SNAPSHOT = {
    "carena_fig1_frequency_order": 256,
    "carena_fig1_receiver_points": 21,
    "max_equation_reproduction_error_db": 0.114866926208812,
    "max_egn_vs_digitized_ssfm_error_db": 0.43681519892827936,
    "interpretation": (
        "SCI equations reproduce the digitized GN/EGN curves within about 0.115 dB in this check. "
        "The larger SSFM difference is not fitted away; the 2014 paper itself reports residual analytical-vs-simulation gaps, especially in challenging cases."
    ),
    "full_xci_mci_status": (
        "Eq.(18)/Appendix-A and Eq.(36)-(41)/Appendix-B correction paths are implemented with Appendix-C-style factorization. "
        "A publication-grade independent Fig.3/6/8 digitized-curve regression is still recommended before using full EGN as a sign-off reference."
    ),
}

__all__ = [name for name in globals() if not name.startswith("_")]
