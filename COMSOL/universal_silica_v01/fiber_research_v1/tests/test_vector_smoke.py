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
        s=solve(d,m,n_modes=1,target_neff=1.446,max_sample_points=10)
        self.assertEqual(s["schema"],"LSCNS_fiber_research_vector_v1")
        self.assertTrue(s["modes"])
        self.assertEqual(sorted(s["modes"][0]["field_samples"]["components"]),
                         sorted(("Ex","Ey","Ez","Hx","Hy","Hz")))
        self.assertFalse(s["validation"]["loss_dB_km_certified"])

if __name__=="__main__": unittest.main()
