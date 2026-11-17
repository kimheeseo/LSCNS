"""STEP5 synthetic math verification; does NOT certify PCF products or complete gap."""
import copy
import math
import unittest
import numpy as np
from step5_bloch import (
    SCHEMA_OUTPUT, analyze, assemble_operator, defect_handoff,
    eigenfrequencies, example_config, hole_mask, primitive,
    smallest_separation, validate_config
)

class Step5BlochTests(unittest.TestCase):
    def setUp(self):
        self.cfg=example_config()
        self.cfg.update(lattice="square",grid_n=12,bands=3,
                        kmesh_n=2,path_points_per_segment=2,beta_norms=[.4])

    def test_unit_cell_primitive_area(self):
        for lattice in ("triangular","hexagonal","kagome"):
            self.assertAlmostEqual(np.linalg.det(primitive(lattice,3)),
                                   9*math.sqrt(3)/2)
        self.assertAlmostEqual(np.linalg.det(primitive("square",3)),9)

    def test_motif_separation(self):
        p=3.
        self.assertAlmostEqual(smallest_separation("square",p),p)
        self.assertAlmostEqual(smallest_separation("triangular",p),p)
        self.assertAlmostEqual(smallest_separation("hexagonal",p),p/math.sqrt(3))
        self.assertAlmostEqual(smallest_separation("kagome",p),p/2)

    def test_overlapping_honeycomb_rejected(self):
        self.cfg.update(lattice="hexagonal",hole_diameter_over_pitch=.7)
        with self.assertRaises(ValueError):
            validate_config(self.cfg)

    def test_mask_4_lattices(self):
        for lattice in ("triangular","hexagonal","kagome","square"):
            cfg=copy.deepcopy(self.cfg)
            cfg.update(lattice=lattice,hole_diameter_over_pitch=.2)
            mask=hole_mask(cfg)
            self.assertEqual(mask.shape,(12,12))
            self.assertGreater(mask.sum(),0)
            self.assertLess(mask.sum(),144)

    def test_hermitian_periodic_oblique(self):
        self.cfg.update(lattice="triangular",hole_diameter_over_pitch=.2)
        K,M=assemble_operator(self.cfg,(.31,-.21),.4)
        err=K-K.getH()
        self.assertLess(max(abs(err.data),default=0),1e-10)
        self.assertTrue(np.all(M.diagonal().real>0))

    def test_homogeneous_gamma_analytic(self):
        self.cfg.update(n_air=1.45,n_host=1.45)
        frequencies=eigenfrequencies(self.cfg,(0,0),.4)
        self.assertAlmostEqual(frequencies[0],.4/1.45,delta=2e-6)

    def test_homogeneous_bloch_phase_analytic(self):
        self.cfg.update(n_air=1.45,n_host=1.45,grid_n=22)
        f=eigenfrequencies(self.cfg,(.25,0),.4)[0]
        analytic=math.sqrt(.25**2+.4**2)/1.45
        self.assertAlmostEqual(f,analytic,delta=.0006)

    def test_beta_normalization_and_light_lines(self):
        self.cfg.update(n_air=1.0,n_host=1.45)
        out=analyze(self.cfg)
        self.assertEqual(out["schema"],SCHEMA_OUTPUT)
        self.assertEqual(out["checks"]["complete_bandgap_certification"],"NOT_VERIFIED")
        self.assertEqual(out["gap_map"]["status"],"SAMPLED_ONLY_NOT_COMPLETE_BZ_PROOF")
        self.assertEqual([x["n"] for x in out["gap_map"]["light_lines_beta_over_k0"]],[1.0,1.45])
        self.assertEqual(len(out["path"]["bands_f"]),7)
        self.assertEqual(len(out["path"]["bands_f"][0]),3)
        for gap in out["gap_map"]["candidates"]:
            self.assertGreater(gap["f_high"],gap["f_low"])
            self.assertAlmostEqual(gap["beta_over_k0_at_f_low"],gap["beta_norm"]/gap["f_low"])

    def test_handoff_requires_computed_candidate_and_triangular(self):
        synthetic={"schema":SCHEMA_OUTPUT,
                   "input":copy.deepcopy(self.cfg),
                   "input_sha256":"unit-test-only",
                   "gap_map":{"candidates":[{"f_low":1.0,"f_high":1.1,"beta_norm":.4,
                                             "status":"SAMPLED_GAP_CANDIDATE"}]}}
        with self.assertRaises(NotImplementedError):
            defect_handoff(synthetic,0)
        synthetic["input"]["lattice"]="triangular"
        with self.assertRaises(ValueError):
            defect_handoff(synthetic,1)
        try:
            result=defect_handoff(synthetic,0,rings=4,defect_cells=7)
        except ImportError:
            self.skipTest("schema module package resolution missing")
        self.assertEqual(result["family"],"pbgf")
        self.assertEqual(result["parameters"]["defect_cells"],7)
        self.assertEqual(result["metadata"]["status"],"proposal_not_validated")
        self.assertEqual(result["boundary"]["type"],"pml")
        self.assertAlmostEqual(result["wavelength_nm"],3000/1.05)

if __name__=="__main__":
    unittest.main()
