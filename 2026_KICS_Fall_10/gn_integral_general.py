"""General-purpose numerical GN-model integral engine.

This module extends the original single-channel ``gn_integral.py`` into a
reusable WDM/link engine while keeping the same Poggiolini GN-model
conventions.

Implemented scope
-----------------
1. Full WDM PSD integration (SCI/XCI/MCI arise naturally from the full PSD).
2. Unequal channel powers, bandwidths and irregular spacing.
3. Rectangular, raised-cosine / RRC-power, and user-supplied PSD shapes.
4. beta2 + beta3 phase mismatch (Appendix G, Eq. G.3/G.4 convention).
5. Non-identical spans with span-specific length, loss, dispersion, gamma,
   lumped gain and optional lumped dispersion compensation (Eq. 100 style).
6. Coherent or incoherent span accumulation.
7. Numerical distributed-gain profiles: ideal distributed gain and a simple
   backward-Raman profile based on Eqs. (103)-(105).
8. Scrambled Sobol QMC integration for the 2-D frequency integral.

Important convention
--------------------
- Frequency: THz (= 1/ps)
- beta2: ps^2/km
- beta3: ps^3/km
- distance: km
- gamma: 1/(W km)
- alpha_db_per_km: conventional *POWER* attenuation [dB/km]
- internally alpha_field = alpha_db_per_km*ln(10)/20, so field~exp(-alpha z)
- channel power and PSD are total dual-polarization quantities
- DP GN coefficient = 16/27

The primary path is the *integral GN model*, not the Section-V asinh
closed-form approximation.  The epsilon coherent-correction formula belongs
mainly to the optional closed-form/incoherent approximation family and is not
used to replace the direct coherent integral here.

Limits
------
- The default is the unchanged GN model (Gaussian-signal assumption).
  Optional egn_mu4/egn_mu6 enable the rectangular-spectrum SCI terms of
  Carena et al., Opt. Express 22, 16335 (2014), Eqs. (5)-(9), (43)-(45).
  XCI/MCI remain GN: this is not a full WDM EGN implementation.
- Mid-link channel add/drop is intentionally not modeled; the source paper
  states that a compact general formula is difficult because each spectral
  portion has a different amplitude/phase history.
- Distributed-amplification handling is a numerical profile extension of the
  paper's distributed-gain integral.  For heterogeneous Raman links, validate
  against a dedicated Raman model / simulator before design sign-off.
"""
from __future__ import annotations

from dataclasses import dataclass, field, replace
from functools import lru_cache
from typing import Callable, Iterable, Literal, Optional, Sequence
import math
import numpy as np
from scipy.stats import qmc
from scipy.special import roots_legendre

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
