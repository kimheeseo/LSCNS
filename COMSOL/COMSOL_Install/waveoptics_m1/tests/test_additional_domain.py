"""Additional acceptance checks prompted by reported cladding-decay warnings.

Original 28 tests and tolerances remain untouched. This file was added after
the initial implementation passed, to resolve the high-contrast TM1 warning.
"""
import numpy as np
import pytest
from .test_m1 import solve, STRONG
from .analytic_slab import exact_modes, l2_field_error


@pytest.mark.parametrize('pol',['TE','TM'])
def test_high_contrast_three_level_convergence(pol):
    hs = [0.01,0.005,0.0025]
    # Larger domain controls the slowly decaying TM1 tail.
    sols = [solve(**{**STRONG,'padding_um':8.0},polarization=pol,h_um=h) for h in hs]
    refs = exact_modes(3.45,1.44,0.4,1.55,pol)
    for j,ref in enumerate(refs):
        for errors in (
            [abs(s.modes[j].neff-ref.neff) for s in sols],
            [l2_field_error(s,s.modes[j],ref) for s in sols],
        ):
            errors = np.array(errors)
            assert np.all(np.diff(errors) < 0)
            slope = np.log(errors[:-1]/errors[1:])/np.log(2)
            assert np.all((slope > 1.8) & (slope < 2.2))


@pytest.mark.parametrize('pol',['TE','TM'])
def test_high_contrast_domain_convergence(pol):
    sols = [solve(**{**STRONG,'padding_um':p},polarization=pol,h_um=0.0025) for p in [4.0,6.0,8.0]]
    assert all(len(s.modes)==2 for s in sols)
    delta46 = max(abs(a.neff-b.neff) for a,b in zip(sols[0].modes,sols[1].modes))
    delta68 = max(abs(a.neff-b.neff) for a,b in zip(sols[1].modes,sols[2].modes))
    assert delta46 < 1e-8
    assert delta68 < 1e-10
