"""Stdlib smoke tests. No gmsh/skfem necessary; physics acceptance is separate."""
import sys
import json
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from schema import load_presets, validate_design, save_design, load_design
from materials import silica_index
from geometry import hex_centers, mcf_centers

class DesignTests(unittest.TestCase):
    def test_all_25_presets(self):
        presets=load_presets()
        self.assertEqual(25,len(presets))
        for name,d in presets.items():
            with self.subTest(name=name):
                self.assertIs(validate_design(d),d)

    def test_round_trip(self):
        d=load_presets()["anti_resonant__nanf"]
        with TemporaryDirectory() as tmp:
            dst=save_design(d,Path(tmp)/"nanf.json")
            self.assertEqual(d,load_design(dst))
            self.assertEqual("1.0.0",json.loads(dst.read_text())["schema_version"])

    def test_bad_version_and_nonfinite(self):
        d=json.loads(json.dumps(load_presets()["index_guiding__step"]))
        d["schema_version"]="2.0.0"
        with self.assertRaises(ValueError): validate_design(d)
        d["schema_version"]="1.0.0";d["parameters"]["core_radius_um"]=float("nan")
        with self.assertRaises(ValueError): validate_design(d)

    def test_unknown_sellmeier_guard(self):
        d=json.loads(json.dumps(load_presets()["index_guiding__step"]))
        d["materials"]["core"]={"model":"ge_doped_sellmeier"}
        with self.assertRaises(NotImplementedError): validate_design(d)

    def test_pml_and_mcf_overlap(self):
        d=json.loads(json.dumps(load_presets()["mcf__hex"]))
        d["parameters"]["pitch_um"]=2
        with self.assertRaises(ValueError): validate_design(d)
        d=load_presets()["anti_resonant__nanf"]
        d=json.loads(json.dumps(d))
        d["boundary"]["thickness_um"]=999
        with self.assertRaises(ValueError): validate_design(d)

    def test_silica_ref(self):
        self.assertAlmostEqual(1.4440236217,silica_index(1550),places=7)

    def test_ring_layout(self):
        self.assertEqual(6,len(mcf_centers(6,18,"ring")))
        self.assertEqual(18,len(hex_centers(2,1,omit=0))) # all except center

if __name__=="__main__": unittest.main()
