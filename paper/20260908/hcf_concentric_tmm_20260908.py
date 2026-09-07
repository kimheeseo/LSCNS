"""Scalar concentric leaky-mode solver for hollow-core antiresonant surrogates.

This is a research/diagnostic model, not a full-vector DNANF FEM replacement.
It solves the scalar radial Helmholtz equation in arbitrary concentric annuli,
with a regular J_m core field and a purely outgoing H_m^(1) exterior field.

Time/axial convention: exp(i*beta*z - i*omega*t).  A decaying mode therefore
has Im(beta) > 0 and power attenuation 2*Im(beta) [Np/m].
"""

from __future__ import annotations

from dataclasses import asdict, dataclass
from typing import Iterable, Sequence

import numpy as np
from scipy import optimize, special


U01 = float(special.jn_zeros(0, 1)[0])
DB_PER_NEPER_POWER = 10.0 / np.log(10.0)


def silica_index_sellmeier(wavelength_m: float) -> float:
    """Malitson fused-silica Sellmeier index; wavelength is in metres."""
    lam_um = float(wavelength_m) * 1e6
    b = np.array([0.6961663, 0.4079426, 0.8974794])
    c_um = np.array([0.0684043, 0.1162414, 9.896161])
    n_sq = 1.0 + np.sum(b * lam_um**2 / (lam_um**2 - c_um**2))
    return float(np.sqrt(n_sq))


def _radiation_sqrt(z: complex) -> complex:
    """Choose transverse-wave-number branch with Re(q)>=0."""
    q = complex(np.sqrt(complex(z)))
    if q.real < 0.0 or (abs(q.real) < 1e-15 and q.imag > 0.0):
        q = -q
    return q


@dataclass(frozen=True)
class ConcentricLeakyModeModel:
    """Core + finite annuli + infinite exterior.

    `annulus_indices` and `annulus_thickness_um` describe only the finite
    annuli between the core and the infinite exterior.
    """

    name: str
    core_radius_um: float
    annulus_indices: tuple[float, ...]
    annulus_thickness_um: tuple[float, ...]
    n_core: float = 1.0
    n_outer: float = 1.0
    azimuthal_order: int = 0

    def __post_init__(self) -> None:
        if self.core_radius_um <= 0:
            raise ValueError("core_radius_um must be positive")
        if len(self.annulus_indices) != len(self.annulus_thickness_um):
            raise ValueError("annulus_indices and annulus_thickness_um must have equal length")
        if any(n <= 0 for n in self.annulus_indices):
            raise ValueError("all refractive indices must be positive")
        if any(t <= 0 for t in self.annulus_thickness_um):
            raise ValueError("all annulus thicknesses must be positive")

    @property
    def core_radius_m(self) -> float:
        return self.core_radius_um * 1e-6

    @property
    def annulus_thickness_m(self) -> tuple[float, ...]:
        return tuple(t * 1e-6 for t in self.annulus_thickness_um)

    @property
    def boundary_radii_m(self) -> np.ndarray:
        return np.cumsum((self.core_radius_m, *self.annulus_thickness_m))

    @property
    def region_indices(self) -> tuple[float, ...]:
        return (self.n_core, *self.annulus_indices, self.n_outer)


@dataclass(frozen=True)
class ModeSolution:
    model_name: str
    wavelength_nm: float
    converged: bool
    u_real: float
    u_imag: float
    beta_real_per_m: float
    beta_imag_per_m: float
    neff_real: float
    neff_imag: float
    loss_db_km: float
    characteristic_abs: float
    direct_transfer_residual_abs: float
    relative_smallest_singular_value: float
    root_message: str
    roots_found: int

    def to_dict(self) -> dict:
        return asdict(self)


def beta_from_u(u: complex, wavelength_m: float, core_radius_m: float, n_core: float = 1.0) -> complex:
    """Convert dimensionless core transverse parameter u=q_core*a to beta."""
    k0 = 2.0 * np.pi / wavelength_m
    beta = complex(np.sqrt(complex((n_core * k0) ** 2 - (u / core_radius_m) ** 2)))
    if beta.real < 0.0:
        beta = -beta
    # exp(i beta z-i omega t): choose the axially decaying resonance.
    if beta.imag < 0.0:
        beta = beta.conjugate()
    return beta


def _region_basis(
    region_number: int,
    region_count: int,
    n: float,
    beta: complex,
    k0: float,
    r: float,
    m: int,
) -> np.ndarray:
    """Return basis columns [field; (1/k0)*d(field)/dr]."""
    q = _radiation_sqrt((n * k0) ** 2 - beta**2)
    x = q * r
    derivative_factor = q / k0

    if region_number == 0:  # regular at r=0: J_m only
        return np.array(
            [[special.jv(m, x)], [derivative_factor * special.jvp(m, x, 1)]],
            dtype=complex,
        )
    if region_number == region_count - 1:  # outgoing at infinity: H_m^(1) only
        return np.array(
            [[special.hankel1(m, x)], [derivative_factor * special.h1vp(m, x, 1)]],
            dtype=complex,
        )
    return np.array(
        [
            [special.jv(m, x), special.yv(m, x)],
            [
                derivative_factor * special.jvp(m, x, 1),
                derivative_factor * special.yvp(m, x, 1),
            ],
        ],
        dtype=complex,
    )


def _jy_basis(n: float, beta: complex, k0: float, r: float, m: int) -> np.ndarray:
    """Two-column J_m/Y_m basis used by the explicit coefficient TMM."""
    q = _radiation_sqrt((n * k0) ** 2 - beta**2)
    x = q * r
    return np.array(
        [
            [special.jv(m, x), special.yv(m, x)],
            [(q / k0) * special.jvp(m, x, 1), (q / k0) * special.yvp(m, x, 1)],
        ],
        dtype=complex,
    )


def direct_transfer_outgoing_residual(
    u: complex, wavelength_m: float, model: ConcentricLeakyModeModel
) -> complex:
    """Explicit transfer-matrix cross-check.

    At interface r_i, c_(i+1)=F_(i+1)^(-1)F_i c_i.  The regular core has
    c_core=[1,0].  In the outer J/Y basis an outgoing H^(1)=J+iY wave obeys
    B_out-i*A_out=0.  Repeated transfer is less stable than the global block
    determinant, so it is used as a diagnostic rather than the root objective.
    """
    k0 = 2.0 * np.pi / wavelength_m
    beta = beta_from_u(u, wavelength_m, model.core_radius_m, model.n_core)
    coefficients = np.array([1.0 + 0.0j, 0.0 + 0.0j])
    for interface, radius in enumerate(model.boundary_radii_m):
        left = _jy_basis(
            model.region_indices[interface], beta, k0, radius, model.azimuthal_order
        )
        right = _jy_basis(
            model.region_indices[interface + 1], beta, k0, radius, model.azimuthal_order
        )
        coefficients = np.linalg.solve(right, left @ coefficients)
        coefficients /= max(float(np.max(np.abs(coefficients))), 1e-300)
    return complex(coefficients[1] - 1j * coefficients[0])


def matching_matrix(u: complex, wavelength_m: float, model: ConcentricLeakyModeModel) -> np.ndarray:
    """Block matrix enforcing field and derivative continuity at every boundary.

    Each pair of block columns is equivalent to the usual transfer relation
    c_(i+1) = F_(i+1)^(-1) F_i c_i.  The global block form avoids repeatedly
    multiplying poorly conditioned transfer matrices.
    """
    k0 = 2.0 * np.pi / wavelength_m
    beta = beta_from_u(u, wavelength_m, model.core_radius_m, model.n_core)
    indices = model.region_indices
    boundaries = model.boundary_radii_m
    region_count = len(indices)
    dimensions = [1] + [2] * len(model.annulus_indices) + [1]
    starts = np.cumsum([0] + dimensions[:-1])
    matrix = np.zeros((2 * (region_count - 1), sum(dimensions)), dtype=complex)

    for interface, radius in enumerate(boundaries):
        left = _region_basis(
            interface,
            region_count,
            indices[interface],
            beta,
            k0,
            radius,
            model.azimuthal_order,
        )
        right = _region_basis(
            interface + 1,
            region_count,
            indices[interface + 1],
            beta,
            k0,
            radius,
            model.azimuthal_order,
        )
        rows = slice(2 * interface, 2 * interface + 2)
        matrix[rows, starts[interface] : starts[interface] + dimensions[interface]] = left
        matrix[
            rows,
            starts[interface + 1] : starts[interface + 1] + dimensions[interface + 1],
        ] = -right
    return matrix


def characteristic(u: complex, wavelength_m: float, model: ConcentricLeakyModeModel) -> complex:
    """Column-scaled characteristic determinant; its zeros are the leaky modes."""
    matrix = matching_matrix(u, wavelength_m, model)
    scales = np.max(np.abs(matrix), axis=0)
    if np.any(~np.isfinite(scales)) or np.any(scales == 0.0):
        return complex(np.nan, np.nan)
    return complex(np.linalg.det(matrix / scales))


def beta_loss_db_km(beta: complex) -> float:
    """Power attenuation: 2*Im(beta)*(10/ln(10))*1000."""
    return float(2.0 * abs(beta.imag) * DB_PER_NEPER_POWER * 1000.0)


def _root_residual(x: Sequence[float], wavelength_m: float, model: ConcentricLeakyModeModel) -> np.ndarray:
    value = characteristic(complex(float(x[0]), float(x[1])), wavelength_m, model)
    if not np.isfinite(value.real) or not np.isfinite(value.imag):
        return np.array([1e6, 1e6])
    return np.array([value.real, value.imag])


def solve_fundamental_mode(
    wavelength_nm: float,
    model: ConcentricLeakyModeModel,
    real_seeds: Iterable[float] | None = None,
    imag_seeds: Iterable[float] | None = None,
    maxfev: int = 5000,
    determinant_tolerance: float = 1e-8,
    singular_value_tolerance: float = 1e-10,
) -> ModeSolution:
    """Multi-start complex root search near the scalar LP01/HE11 surrogate.

    A candidate is accepted only when both the characteristic determinant and
    the relative smallest singular value independently confirm singularity.
    No number is returned as a valid loss if those checks fail.
    """
    wavelength_m = wavelength_nm * 1e-9
    if real_seeds is None:
        real_seeds = (U01 - 0.08, U01, U01 + 0.08)
    if imag_seeds is None:
        imag_seeds = (-1e-2, -1e-3, -1e-4, -1e-5, -1e-6, -1e-7, -1e-8)

    candidates: list[tuple[complex, object, float, float]] = []
    messages: list[str] = []
    best_diagnostic: tuple[float, complex, float, float, str] | None = None
    for ur in real_seeds:
        for ui in imag_seeds:
            result = optimize.root(
                _root_residual,
                np.array([ur, ui], dtype=float),
                args=(wavelength_m, model),
                method="hybr",
                options={"xtol": 1e-11, "maxfev": maxfev},
            )
            messages.append(str(result.message))
            u = complex(float(result.x[0]), float(result.x[1]))
            # Restrict to the low-radial-order resonance branch. Extremely
            # negative Im(u) roots are mathematically valid but are not the
            # core-like fundamental sought here.
            if not (1.5 < u.real < 3.5 and -1.0 < u.imag < 0.0):
                continue
            det_abs = abs(characteristic(u, wavelength_m, model))
            singular_values = np.linalg.svd(matching_matrix(u, wavelength_m, model), compute_uv=False)
            relative_smin = float(singular_values[-1] / singular_values[0])
            score = max(
                det_abs / determinant_tolerance,
                relative_smin / singular_value_tolerance,
            )
            if best_diagnostic is None or score < best_diagnostic[0]:
                best_diagnostic = (score, u, det_abs, relative_smin, str(result.message))
            if det_abs <= determinant_tolerance and relative_smin <= singular_value_tolerance:
                if not any(abs(u - old[0]) < 1e-6 for old in candidates):
                    candidates.append((u, result, det_abs, relative_smin))

    if not candidates:
        if best_diagnostic is None:
            detail = "No optimizer result stayed on 1.5<Re(u)<3.5 and -1<Im(u)<0."
        else:
            _, u_best, det_best, smin_best, optimizer_message = best_diagnostic
            detail = (
                f"Best low-order candidate u={u_best.real:.9g}{u_best.imag:+.3g}j, "
                f"|det|={det_best:.3e} (limit {determinant_tolerance:.1e}), "
                f"relative smin={smin_best:.3e} (limit {singular_value_tolerance:.1e}); "
                f"optimizer: {optimizer_message}"
            )
        return ModeSolution(
            model_name=model.name,
            wavelength_nm=float(wavelength_nm),
            converged=False,
            u_real=np.nan,
            u_imag=np.nan,
            beta_real_per_m=np.nan,
            beta_imag_per_m=np.nan,
            neff_real=np.nan,
            neff_imag=np.nan,
            loss_db_km=np.nan,
            characteristic_abs=np.nan,
            direct_transfer_residual_abs=np.nan,
            relative_smallest_singular_value=np.nan,
            root_message="No independently verified fundamental-like root. " + detail,
            roots_found=0,
        )

    # The scalar fundamental is the accepted root whose real part is nearest u01.
    u, result, det_abs, relative_smin = min(candidates, key=lambda item: abs(item[0].real - U01))
    beta = beta_from_u(u, wavelength_m, model.core_radius_m, model.n_core)
    neff = beta / (2.0 * np.pi / wavelength_m)
    return ModeSolution(
        model_name=model.name,
        wavelength_nm=float(wavelength_nm),
        converged=True,
        u_real=float(u.real),
        u_imag=float(u.imag),
        beta_real_per_m=float(beta.real),
        beta_imag_per_m=float(beta.imag),
        neff_real=float(neff.real),
        neff_imag=float(neff.imag),
        loss_db_km=beta_loss_db_km(beta),
        characteristic_abs=float(det_abs),
        direct_transfer_residual_abs=float(
            abs(direct_transfer_outgoing_residual(u, wavelength_m, model))
        ),
        relative_smallest_singular_value=float(relative_smin),
        root_message=str(result.message),
        roots_found=len(candidates),
    )


def bouncing_ray_te_loss_db_km(
    wavelength_nm: float,
    core_radius_um: float = 14.75,
    wall_thickness_um: float = 0.50,
    silica_index: float | None = None,
) -> float:
    """TE-like single-wall thin-wall bouncing-ray/ARROW estimate.

    This is the appropriate closed-form comparator for the unweighted scalar
    field/derivative continuity used by this notebook.  A hybrid HE11 estimate
    generally combines TE- and TM-like contributions and need not equal it.
    """
    wavelength_m = wavelength_nm * 1e-9
    a = core_radius_um * 1e-6
    t = wall_thickness_um * 1e-6
    k0 = 2.0 * np.pi / wavelength_m
    n = silica_index_sellmeier(wavelength_m) if silica_index is None else float(silica_index)
    kappa = U01 / a
    sigma = k0 * np.sqrt(n**2 - 1.0)
    phase = sigma * t
    denominator = 4.0 * np.cos(phase) ** 2 + (
        kappa / sigma + sigma / kappa
    ) ** 2 * np.sin(phase) ** 2
    alpha_per_m = 2.0 * U01 / (a**2 * k0 * denominator)
    return float(alpha_per_m * DB_PER_NEPER_POWER * 1000.0)


def bache_bouncing_ray_loss_db_km(
    wavelength_nm: float,
    polarization: str = "hybrid",
    core_radius_um: float = 14.75,
    wall_thickness_um: float = 0.50,
    silica_index: float | None = None,
) -> float:
    """Bache et al. Eqs. (15)-(17), without an FEM fitting factor.

    ``polarization`` may be ``TE``, ``TM`` or ``hybrid``.  The hybrid HE11-like
    result is the arithmetic mean of the TE and TM power-loss coefficients,
    exactly as stated in Eq. (17) of the paper.  This function deliberately
    does *not* apply the paper's design-specific ``f_FEM`` multiplier.
    """
    wavelength_m = wavelength_nm * 1e-9
    a = core_radius_um * 1e-6
    t = wall_thickness_um * 1e-6
    k0 = 2.0 * np.pi / wavelength_m
    n = silica_index_sellmeier(wavelength_m) if silica_index is None else float(silica_index)
    kappa = U01 / a
    sigma = k0 * np.sqrt(n**2 - 1.0)
    phase = sigma * t
    prefactor = 2.0 * U01 / (a**2 * k0)

    denominator_te = 4.0 * np.cos(phase) ** 2 + (
        kappa / sigma + sigma / kappa
    ) ** 2 * np.sin(phase) ** 2
    denominator_tm = 4.0 * np.cos(phase) ** 2 + (
        n**2 * kappa / sigma + sigma / (n**2 * kappa)
    ) ** 2 * np.sin(phase) ** 2
    te = prefactor / denominator_te * DB_PER_NEPER_POWER * 1000.0
    tm = prefactor / denominator_tm * DB_PER_NEPER_POWER * 1000.0

    key = polarization.strip().lower()
    if key == "te":
        return float(te)
    if key == "tm":
        return float(tm)
    if key in {"hybrid", "he11"}:
        return float((te + tm) / 2.0)
    raise ValueError("polarization must be 'TE', 'TM', or 'hybrid'")


def multilayer_bouncing_ray_loss_db_km(
    wavelength_nm: float,
    layer_indices: Sequence[float],
    layer_thickness_um: Sequence[float],
    polarization: str = "hybrid",
    core_radius_um: float = 14.75,
) -> float:
    """Planar multilayer extension of Bache's bouncing-ray loss.

    The incidence angle is fixed by the MS transverse wave number ``u01/a``.
    A standard lossless characteristic matrix returns the power reflectance of
    the finite stack.  The per-bounce transmission is then converted with
    Bache Eq. (6).  The result is geometry-only: no Petrovich loss value, FEM
    result, fitted scale factor, or regression coefficient is used.

    This is a useful independent cross-check of the concentric TMM, but it is
    still a one-dimensional radial surrogate and omits azimuthal confinement.
    """
    if len(layer_indices) != len(layer_thickness_um):
        raise ValueError("layer_indices and layer_thickness_um must have equal length")
    if not layer_indices:
        raise ValueError("at least one finite layer is required")

    wavelength_m = wavelength_nm * 1e-9
    a = core_radius_um * 1e-6
    k0 = 2.0 * np.pi / wavelength_m
    beta = np.sqrt(k0**2 - (U01 / a) ** 2 + 0j)
    indices = [1.0, *map(float, layer_indices), 1.0]
    q = [_radiation_sqrt((n * k0) ** 2 - beta**2) for n in indices]

    def one_pol(pol: str) -> float:
        if pol == "te":
            admittance = [qi / k0 for qi in q]
        elif pol == "tm":
            admittance = [n**2 * k0 / qi for n, qi in zip(indices, q)]
        else:
            raise ValueError("internal polarization must be te or tm")

        matrix = np.eye(2, dtype=complex)
        for region, thickness_um in enumerate(layer_thickness_um, start=1):
            phase = q[region] * float(thickness_um) * 1e-6
            y = admittance[region]
            layer_matrix = np.array(
                [
                    [np.cos(phase), 1j * np.sin(phase) / y],
                    [1j * y * np.sin(phase), np.cos(phase)],
                ],
                dtype=complex,
            )
            matrix = matrix @ layer_matrix

        a11, a12, a21, a22 = matrix.ravel()
        y_in = (a21 + a22 * admittance[-1]) / (
            a11 + a12 * admittance[-1]
        )
        reflection = (admittance[0] - y_in) / (admittance[0] + y_in)
        transmission_per_bounce = max(0.0, 1.0 - float(abs(reflection) ** 2))
        bounce_rate = U01 / (2.0 * a**2 * k0)
        alpha_per_m = transmission_per_bounce * bounce_rate
        return float(alpha_per_m * DB_PER_NEPER_POWER * 1000.0)

    key = polarization.strip().lower()
    if key == "te":
        return one_pol("te")
    if key == "tm":
        return one_pol("tm")
    if key in {"hybrid", "he11"}:
        return float((one_pol("te") + one_pol("tm")) / 2.0)
    raise ValueError("polarization must be 'TE', 'TM', or 'hybrid'")


def build_models(
    wavelength_nm: float,
    gap1_um: float,
    gap2_um: float | None = None,
    core_radius_um: float = 14.75,
    wall_thickness_um: float = 0.50,
    fixed_silica_index: float | None = None,
) -> dict[str, ConcentricLeakyModeModel]:
    """Create the requested 3/5-layer models and optional 7-layer diagnostic."""
    n_si = (
        silica_index_sellmeier(wavelength_nm * 1e-9)
        if fixed_silica_index is None
        else float(fixed_silica_index)
    )
    models = {
        "3-layer": ConcentricLeakyModeModel(
            "3-layer: air | silica | outgoing air",
            core_radius_um,
            (n_si,),
            (wall_thickness_um,),
        ),
        "5-layer": ConcentricLeakyModeModel(
            "5-layer: air | silica | air gap | silica | outgoing air",
            core_radius_um,
            (n_si, 1.0, n_si),
            (wall_thickness_um, gap1_um, wall_thickness_um),
        ),
    }
    if gap2_um is not None:
        models["7-layer"] = ConcentricLeakyModeModel(
            "7-layer diagnostic: air | silica | gap | silica | gap | silica | outgoing air",
            core_radius_um,
            (n_si, 1.0, n_si, 1.0, n_si),
            (wall_thickness_um, gap1_um, wall_thickness_um, gap2_um, wall_thickness_um),
        )
    return models


def solve_models(
    wavelengths_nm: Iterable[float],
    gap1_um: float,
    gap2_um: float | None = None,
    **model_kwargs,
) -> list[ModeSolution]:
    solutions: list[ModeSolution] = []
    for wavelength_nm in wavelengths_nm:
        for model in build_models(wavelength_nm, gap1_um, gap2_um, **model_kwargs).values():
            solutions.append(solve_fundamental_mode(wavelength_nm, model))
    return solutions


if __name__ == "__main__":
    # Public HCF2 midpoint diameters under an internally tangent radial mapping:
    # gap1=(31.05-2*0.50)-23.75=6.30 um; gap2=(23.75-2*0.50)-7.70=15.05 um.
    for solution in solve_models((1310.0, 1550.0), gap1_um=6.30, gap2_um=15.05):
        print(solution.to_dict())
