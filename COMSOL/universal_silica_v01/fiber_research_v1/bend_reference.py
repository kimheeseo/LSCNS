"""Explicit COMSOL bend normalization; no inferred radius or rounded reference.

COMSOL 6.4 step_index_fiber_bend: neff_power = neff_geometry*rAverage/r0.
Raw eigenvalue is lambda=-i*beta*r0 in the documented convention.
All lengths below are SI. imag_sign converts the raw propagation convention to
exp(i*beta*z-i*omega*t); negative passive loss is retained, never abs()-corrected.
"""
import math


def normalize_bend(value, *, representation, wavelength_m, r_average_m,
                   r_power_m, imag_sign):
    z = complex(value)
    numbers = (z.real, z.imag, wavelength_m, r_average_m, r_power_m)
    if not all(math.isfinite(v) for v in numbers):
        raise ValueError('finite complex value and radii required')
    if min(wavelength_m, r_average_m, r_power_m) <= 0 or imag_sign not in (-1, 1):
        raise ValueError('positive SI lengths and explicit imag_sign +/-1 required')
    k0 = 2*math.pi/wavelength_m
    if representation == 'neff_geometry':
        angular = z*k0*r_average_m
    elif representation == 'angular_propagation':
        angular = z
    elif representation == 'solver_eigenvalue_minus_i_angular':
        angular = 1j*z
    else:
        raise ValueError('explicit supported representation required')
    angular = complex(angular.real, imag_sign*angular.imag)
    geometry = angular/(k0*r_average_m)
    power = angular/(k0*r_power_m)
    return {'neff_geometry_real': geometry.real, 'neff_geometry_imag': geometry.imag,
            'neff_power_real': power.real, 'neff_power_imag': power.imag,
            'alpha_geometry_dB_m': 20/math.log(10)*angular.imag/r_average_m,
            'alpha_power_dB_m': 20/math.log(10)*angular.imag/r_power_m,
            'physical_cross_validation_certified': False}
