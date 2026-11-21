import math
import unittest
from bend_reference import normalize_bend


class BendReference(unittest.TestCase):
    def test_geometry_and_power_radius_are_distinct(self):
        r = normalize_bend(1.445-2e-8j, representation='neff_geometry',
                           wavelength_m=1.55e-6, r_average_m=.003,
                           r_power_m=.003004, imag_sign=-1)
        self.assertAlmostEqual(r['neff_power_real'], 1.445*.003/.003004)
        self.assertGreater(r['alpha_power_dB_m'], 0)
        self.assertFalse(r['physical_cross_validation_certified'])

    def test_raw_eigenvalue_conversion(self):
        k0=2*math.pi/1.55e-6
        m=(1.445-2e-8j)*k0*.003
        a=normalize_bend(m, representation='angular_propagation', wavelength_m=1.55e-6,
                         r_average_m=.003,r_power_m=.003004,imag_sign=-1)
        b=normalize_bend(-1j*m, representation='solver_eigenvalue_minus_i_angular',
                         wavelength_m=1.55e-6,r_average_m=.003,r_power_m=.003004,imag_sign=-1)
        self.assertEqual(a,b)

    def test_missing_normalization_is_rejected_and_sign_is_preserved(self):
        for radius in (0,float('nan'),-1):
            with self.assertRaises(ValueError):
                normalize_bend(1.4,representation='neff_geometry',wavelength_m=1e-6,
                               r_average_m=radius,r_power_m=.003,imag_sign=1)
        r=normalize_bend(1.4-1e-8j,representation='neff_geometry',wavelength_m=1e-6,
                         r_average_m=.003,r_power_m=.003,imag_sign=1)
        self.assertLess(r['alpha_power_dB_m'],0)


if __name__ == '__main__': unittest.main()
