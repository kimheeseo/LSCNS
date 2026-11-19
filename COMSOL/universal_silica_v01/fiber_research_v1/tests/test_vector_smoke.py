"""Dependency-gated numerical smoke checks. Not physical accuracy validation."""
import copy
import json
import sys
import unittest
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from schema import load_presets

try:
    import gmsh
    import skfem
    import scipy
    HAVE_DEPS=True
except ImportError:
    HAVE_DEPS=False

@unittest.skipUnless(HAVE_DEPS,"requires installed gmsh/scikit-fem/scipy")
class NumericalSmoke(unittest.TestCase):
    def test_small_step_mesh(self):
        from geometry import build_mesh
        d=copy.deepcopy(load_presets()["index_guiding__step"])
        d["domain_radius_um"]=9
        d["parameters"]["core_radius_um"]=2.2
        d["mesh"]["bulk_max_um"]=1.0
        d["mesh"]["elements_per_wavelength"]=4
        m=build_mesh(d)
        self.assertGreater(m["quality"]["triangles"],100)
        self.assertGreater(len(m["quality"]["materials"]),1)
        self.assertGreater(m["quality"]["area_min_um2"],0)
        self.assertEqual(len(m["triangles"]),len(m["material_ids"]))

    def test_small_step_vector_candidate(self):
        from geometry import build_mesh
        from vector_fem import solve
        d=copy.deepcopy(load_presets()["index_guiding__step"])
        d["domain_radius_um"]=9
        d["parameters"]["core_radius_um"]=2.2
        d["mesh"]["bulk_max_um"]=1.0
        d["mesh"]["elements_per_wavelength"]=4
        m=build_mesh(d)
        s=solve(d,m,n_modes=1,target_neff=1.446,max_sample_points=10,export_full_quadrature=True)
        self.assertEqual(s["schema"],"LSCNS_fiber_research_vector_v1")
        self.assertTrue(s["modes"])
        self.assertEqual(sorted(s["modes"][0]["field_samples"]["components"]),
                         sorted(("Ex","Ey","Ez","Hx","Hy","Hz")))
        self.assertFalse(s["validation"]["loss_dB_km_certified"])
        self.assertEqual(s["engine_version"],"0.8.6")
        self.assertIn("Et=Ex,Ey",s["modes"][0]["classification"]["Aeff_basis"])
        self.assertAlmostEqual(s["modes"][0]["alpha_total_dB_km"],1000*s["modes"][0]["alpha_total_dB_m"])
        q=s["modes"][0]["field_quadrature"]
        self.assertEqual(len(q["points_um"]),len(q["weights_um2"]))
        self.assertEqual(len(q["Ex"]),len(q["Ey"]))
        self.assertGreater(len(q["points_um"]),10)

        # The complex first-order QEP couplings MUST not disappear via implicit float casting.
        from vector_fem import _assemble
        a=_assemble(d,m)
        self.assertGreater(a["K1"].nnz,0)
        self.assertGreater(float(abs(a["K1"].data.imag).max()),0.0)


    def test_step6a_bent_vector_candidate_smoke_only(self):
        """Actual gmsh + skfem path, NEVER a COMSOL V3 or leakage loss PASS."""
        from geometry import build_mesh
        from vector_fem import solve
        from bending import configure_bend
        d=copy.deepcopy(load_presets()["index_guiding__step"])
        d["domain_radius_um"]=9
        d["parameters"]["core_radius_um"]=2.2
        d["mesh"]["bulk_max_um"]=1.0
        d["mesh"]["elements_per_wavelength"]=4
        d["boundary"]["type"]="pml"
        d["boundary"]["thickness_um"]=2.0
        curved=configure_bend(d,50.0,45.0)
        mesh=build_mesh(d)
        out=solve(curved,mesh,n_modes=1,target_neff=1.446,max_sample_points=10)
        self.assertEqual(out["engine_version"],"0.8.6")
        self.assertFalse(out["validation"]["loss_dB_km_certified"])
        self.assertEqual(out["input"]["_bend"]["direction_deg"],45.0)
        self.assertTrue(out["modes"])
        self.assertIn("alpha_total_dB_m",out["modes"][0])


if __name__=="__main__": unittest.main()

