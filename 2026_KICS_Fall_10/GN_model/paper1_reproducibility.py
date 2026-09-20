"""Reproducibility checks against Poggiolini et al., A Detailed Analytical Derivation
of the GN Model of Non-Linear Interference in Coherent Optical Transmission Systems.

This is a physics/implementation regression test, not a fit to digitized paper curves.
"""
from __future__ import annotations
import json, math, sys
from pathlib import Path
import numpy as np

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(ROOT))

import gn_integral_general as gn

OUT = HERE / "results"
OUT.mkdir(exist_ok=True)


def relerr(a, b):
    return abs(float(a)-float(b))/max(abs(float(b)), np.finfo(float).tiny)


def main():
    rows = []

    # Paper convention: alpha is FIELD attenuation; input data are commonly POWER dB/km.
    alpha_db = 0.22
    expected_alpha_field = alpha_db * math.log(10.0) / 20.0
    actual_alpha_field = gn.alpha_field_from_db(alpha_db)
    rows.append(dict(test="dB/km -> field alpha", expected=expected_alpha_field,
                     actual=actual_alpha_field, relative_error=relerr(actual_alpha_field, expected_alpha_field)))

    # beta2-only phase mismatch used in the paper's identical-span EDFA expressions.
    f1, f2, f = 0.041, -0.073, 0.006  # THz
    beta2 = -21.27  # ps^2/km
    expected_q = 4*math.pi**2*(f1-f)*(f2-f)*beta2
    actual_q = float(gn.phase_mismatch_beta23(f1, f2, f, beta2, 0.0))
    rows.append(dict(test="beta2 phase mismatch", expected=expected_q, actual=actual_q,
                     relative_error=relerr(actual_q, expected_q)))

    # One compensated EDFA span: code kernel must reduce to gamma^2 |rho|^2.
    sp = gn.Span(100.0, 0.22, 1.3, D_ps_nm_km=16.7)
    q = float(gn.phase_mismatch_beta23(f1, f2, f, sp.beta2, 0.0))
    a = sp.alpha_field_per_km
    rho = (1.0 - np.exp((-2*a + 1j*q)*sp.length_km))/(2*a - 1j*q)
    expected_1 = abs(sp.gamma_W_inv_km*rho)**2
    actual_1 = float(gn.link_kernel(f1, f2, f, [sp], accumulation="coherent"))
    rows.append(dict(test="single-span link function", expected=expected_1, actual=actual_1,
                     relative_error=relerr(actual_1, expected_1)))

    # Eq.-17-style incoherent accumulation: identical span powers add linearly.
    nsp = 7
    expected_inc = nsp * expected_1
    actual_inc = float(gn.link_kernel(f1, f2, f, [sp]*nsp, accumulation="incoherent"))
    rows.append(dict(test="identical-span incoherent accumulation", expected=expected_inc, actual=actual_inc,
                     relative_error=relerr(actual_inc, expected_inc)))

    # Eq.-18-style coherent accumulation: geometric phase series.
    geom = sum(np.exp(1j*q*sp.length_km*n) for n in range(nsp))
    expected_coh = abs(sp.gamma_W_inv_km*rho*geom)**2
    actual_coh = float(gn.link_kernel(f1, f2, f, [sp]*nsp, accumulation="coherent"))
    rows.append(dict(test="identical-span coherent accumulation", expected=expected_coh, actual=actual_coh,
                     relative_error=relerr(actual_coh, expected_coh)))

    # Paper-1 generalized non-identical-span structure (Eq.-100 family):
    # each local source is phase-shifted by all previous spans.
    s1 = gn.Span(73.0, 0.19, 1.05, D_ps_nm_km=18.2)
    s2 = gn.Span(91.0, 0.23, 1.42, D_ps_nm_km=4.5)
    def local(s):
        qq = float(gn.phase_mismatch_beta23(f1, f2, f, s.beta2, 0.0))
        aa = s.alpha_field_per_km
        rr = (1.0 - np.exp((-2*aa + 1j*qq)*s.length_km))/(2*aa - 1j*qq)
        return qq, s.gamma_W_inv_km*rr
    q1, A1 = local(s1)
    _, A2local = local(s2)
    expected_het = abs(A1 + np.exp(1j*q1*s1.length_km)*A2local)**2
    actual_het = float(gn.link_kernel(f1, f2, f, [s1, s2], accumulation="coherent"))
    rows.append(dict(test="non-identical-span coherent phase history", expected=expected_het, actual=actual_het,
                     relative_error=relerr(actual_het, expected_het)))

    # Exact cubic NLI-power scaling is a defining first-order GN-model invariant.
    p3 = gn.common_power_scaling_self_test()
    rows.append(dict(test="common launch-power cubic scaling", expected=p3["target"],
                     actual=p3["ratio_for_2x_power"], relative_error=p3["error_pct"]/100.0))

    # Numerical reproducibility of the full 2-D WDM integral, reported separately
    # from the exact equation-level checks above.
    system = gn.WDMSystem.equispaced(3, 50.0, 32.0, -3.0)
    spans = [gn.Span(80.0, 0.2, 1.3, D_ps_nm_km=17.0)] * 2
    qmc = []
    for power in (11, 13, 15):
        r = gn.gn_nli_psd_multi_seed(system, spans, 0.0, sobol_power=power,
                                     seeds=(3, 11, 29, 47), accumulation="coherent")
        qmc.append(dict(sobol_power=power, samples=2**power,
                        mean_W_per_THz=r["mean_W_per_THz"],
                        relative_std_pct=r["relative_std_pct"]))
    qmc[-1]["relative_change_vs_previous_pct"] = abs(qmc[-1]["mean_W_per_THz"]-qmc[-2]["mean_W_per_THz"]) / qmc[-1]["mean_W_per_THz"] * 100
    qmc[-2]["relative_change_vs_previous_pct"] = abs(qmc[-2]["mean_W_per_THz"]-qmc[-3]["mean_W_per_THz"]) / qmc[-2]["mean_W_per_THz"] * 100

    max_exact = max(r["relative_error"] for r in rows)
    result = {
        "paper": "Poggiolini et al., A Detailed Analytical Derivation of the GN Model...",
        "equation_level_tests": rows,
        "max_equation_relative_error": max_exact,
        "max_equation_error_pct": 100*max_exact,
        "qmc_convergence": qmc,
        "assessment": (
            "PASS" if max_exact < 1e-10 else "CHECK"
        ),
        "scope_note": "Equation-level reproduction tests the derivation implemented by the engine; QMC dispersion is a numerical integration uncertainty, not a physics-model error."
    }
    (OUT/"paper1_reproducibility.json").write_text(json.dumps(result, indent=2), encoding="utf-8")
    print("PAPER1_RESULT_JSON=" + json.dumps(result, separators=(",",":")))
    print("\nEquation-level checks")
    for r in rows:
        print(f'{r["test"]}: rel.err={100*r["relative_error"]:.3e}%')
    print("\nQMC convergence")
    for r in qmc:
        print(r)


if __name__ == "__main__":
    main()
