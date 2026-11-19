"""Independent topology/area/normal-intersection checks for STEP 1–2 (no new optics)."""
import copy
import json
import math
from pathlib import Path
import sys
import unittest
from tempfile import TemporaryDirectory
import numpy as np
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from schema import load_presets, validate_design, save_design, load_design
from design_checks import design_diagnostics
from mesh_checks import segment_intervals, measure_walls, material_size_targets


class DesignContract(unittest.TestCase):
    def test_all_families_roundtrip_and_json_schema(self):
        from jsonschema import Draft202012Validator
        schema=json.loads((Path(__file__).resolve().parents[1]/"fiber_design.schema.json").read_text())
        Draft202012Validator.check_schema(schema)
        validator=Draft202012Validator(schema)
        with TemporaryDirectory() as tmp:
            for name,d in load_presets().items():
                with self.subTest(name=name):
                    validator.validate(d)
                    self.assertEqual(load_design(save_design(d, Path(tmp)/(name+".json"))),d)

    def test_rejects_string_boolean_and_bad_optional_numbers(self):
        for key,value in [("ellipse_ratio",0),("rings",True),("rings","3"),
                          ("ellipse_ratio",float("inf")),("ellipse_ratio",-1)]:
            d=load_presets()["solid_pcf__hex"];d["parameters"][key]=value
            with self.subTest(key=key,value=value),self.assertRaises(ValueError):validate_design(d)

    def test_cannot_clip_cladding_holes(self):
        d=load_presets()["solid_pcf__hex"];d["domain_radius_um"]=8
        with self.assertRaisesRegex(ValueError,"outside physical"):validate_design(d)

    def test_defect_and_ellipse_constraints(self):
        d=load_presets()["pbgf__7_cell"];d["parameters"]["defect_cells"]=19
        with self.assertRaisesRegex(ValueError,"disagree"):validate_design(d)
        d=load_presets()["solid_pcf__elliptic_birefringent"];d["parameters"]["ellipse_ratio"]=5
        with self.assertRaisesRegex(ValueError,"overlap"):validate_design(d)
        d=load_presets()["solid_pcf__hex"];d["parameters"].update(rings=1,defect_cells=7)
        with self.assertRaisesRegex(ValueError,"no cladding"):validate_design(d)

    def test_ring_mcf_uses_adjacent_chord_not_ring_radius(self):
        d=load_presets()["mcf__ring"]
        d["parameters"].update(core_count=19,pitch_um=12)
        # 12 > 2*3.5, but adjacent chord is about 3.95 < 7.
        with self.assertRaisesRegex(ValueError,"MCF.*overlap"):validate_design(d)

    def test_trench_overlap_and_pm_outer_extent(self):
        d=load_presets()["mcf__trench_assisted"];d["parameters"]["trench_width_um"]=7
        with self.assertRaisesRegex(ValueError,"MCF.*overlap"):validate_design(d)
        d=load_presets()["polarization_maintaining__panda"];d["parameters"]["rod_offset_um"]=28
        with self.assertRaisesRegex(ValueError,"outside physical"):validate_design(d)

    def test_nested_tube_and_missing_material(self):
        d=load_presets()["anti_resonant__nanf"];d["parameters"]["nested_radius_ratio"]=.99
        with self.assertRaisesRegex(ValueError,"nested tube"):validate_design(d)
        d=load_presets()["index_guiding__step"];del d["materials"]["core"]
        with self.assertRaisesRegex(ValueError,"missing.*materials"):validate_design(d)

    def test_approximations_are_not_certificates(self):
        for name in ("index_guiding__graded","kagome__hypocycloid","anti_resonant__conjoined_tube"):
            r=design_diagnostics(load_presets()[name])
            self.assertEqual(r["geometry_fidelity"],"approximate")
            self.assertFalse(r["optical_certification"])

    def test_schema_rejects_incomplete_nested_objects(self):
        from jsonschema import Draft202012Validator
        v=Draft202012Validator(json.loads((Path(__file__).resolve().parents[1]/"fiber_design.schema.json").read_text()))
        for parent,key in [("mesh","bulk_max_um"),("boundary","order")]:
            d=load_presets()["index_guiding__step"];del d[parent][key]
            self.assertTrue(list(v.iter_errors(d)))


class GeometricMeasurements(unittest.TestCase):
    def test_segment_clipping_independent_square(self):
        pts=np.array([[0,0],[1,0],[1,1],[0,1]],float)
        tris=np.array([[0,1,2],[0,2,3]])
        intervals=segment_intervals(pts,tris,[-1,.25],[1,.25])
        self.assertAlmostEqual(sum(b-a for a,b in intervals),.5)
        self.assertEqual(len(segment_intervals(pts,tris,[-1,2],[1,2])),0)

    def test_wall_measurement_detects_coarse_annulus(self):
        def annulus(layers):
            count=64
            pts=np.array([[r*math.cos(a),r*math.sin(a)] for r in np.linspace(1,1.3,layers+1)
                          for a in np.arange(count)*2*math.pi/count])
            tri=[]
            for j in range(layers):
                for i in range(count):
                    a=j*count+i;b=j*count+(i+1)%count;c=a+count;dd=b+count
                    tri.extend([[a,b,c],[b,dd,c]])
            t=np.array(tri)
            return measure_walls(pts,t,np.repeat('silica',len(t)),
                                 [{"centre_um":[0,0],"outer_radius_um":1.3,"thickness_um":.3,"material":"silica"}],3,8)
        self.assertEqual(annulus(1)["status"],"FAIL_SAMPLED")
        self.assertEqual(annulus(4)["status"],"PASS_SAMPLED")

    def test_high_index_target_comes_from_material(self):
        d=load_presets()["index_guiding__step"];d["materials"]["core"]={"model":"constant","n":2.5}
        target=material_size_targets(d,["core"])["core"]
        self.assertAlmostEqual(target["requested_edge_um"],1.55/(2.5*8))


try:
    import gmsh
    HAVE_GMSH=True
except (ImportError,OSError):
    HAVE_GMSH=False


@unittest.skipUnless(HAVE_GMSH,"Gmsh runtime required")
class ActualMesh(unittest.TestCase):
    def design(self, bragg=False):
        d=load_presets()["bragg__concentric" if bragg else "index_guiding__step"]
        d["domain_radius_um"]=4
        d["parameters"]["core_radius_um"]=1
        if bragg:d["parameters"].update(layer_count=2,layer_pitch_um=.4)
        d["wavelength_nm"]=3500
        d["mesh"].update(elements_per_wavelength=4,through_wall_elements=3,wall_probe_angles=8,resolution_policy="wavelength")
        d["boundary"].update(type="pml",thickness_um=1)
        return d

    def test_ellipse_y_major_axis_and_area(self):
        from geometry import build_mesh
        d=load_presets()["polarization_maintaining__elliptic_core"]
        d["domain_radius_um"]=4;d["parameters"]["core_radius_um"]=1
        d["boundary"].update(type="pml",thickness_um=1)
        d["wavelength_nm"]=3500;d["mesh"]["elements_per_wavelength"]=4
        m=build_mesh(d);q=m["quality"]
        expected=math.pi*d["parameters"]["ellipse_ratio"]
        self.assertAlmostEqual(q["cad_material_area_um2"]["core"],expected,places=8)
        edges=m["boundary_edges"]["INTERFACE_background__core"]
        boundary=m["points_um"][np.unique(edges)]
        self.assertAlmostEqual(float(abs(boundary[:,0]).max()),1,places=2)
        self.assertAlmostEqual(float(abs(boundary[:,1]).max()),1.7,places=2)

    def test_pml_area_labels_and_whole_domain_resolution(self):
        from geometry import build_mesh
        m=build_mesh(self.design());q=m["quality"]
        self.assertIn("PML_INTERFACE",m["boundary_edges"])
        self.assertIn("OUTER_PML",m["boundary_edges"])
        self.assertEqual(q["wavelength_resolution"]["status"],"PASS_POLICY")
        self.assertAlmostEqual(q["cad_material_area_um2"]["pml"],math.pi*(25-16),places=8)
        self.assertLess(abs(q["material_area_um2"]["core"]-math.pi)/math.pi,.02)
        for name,edges in m["boundary_edges"].items():
            self.assertGreater(len(edges),0)
            self.assertLess(int(edges.max()),len(m["points_um"]))

    def test_bragg_actual_through_wall_resolution_and_analytic_area(self):
        from geometry import build_mesh
        q=build_mesh(self.design(True))["quality"]
        self.assertEqual(q["wall_resolution"]["status"],"PASS_SAMPLED")
        self.assertEqual(len(q["wall_resolution"]["walls"]),2)
        expected=math.pi*(1.4**2-1)
        self.assertLess(abs(q["material_area_um2"]["high"]-expected)/expected,.02)

    def test_three_actual_distinct_meshes_without_optical_claim(self):
        from mesh_study import run_mesh_study
        with TemporaryDirectory() as tmp:
            r=run_mesh_study(self.design(),tmp)
            self.assertEqual(r["status"],"MESH_REFINEMENT_COMPUTED")
            self.assertEqual(len(r["cases"]),3)
            self.assertFalse(r["optical_convergence"])
            errors=[c["area_relative_errors"]["core"] for c in r["cases"]]
            self.assertLess(errors[-1],errors[0])


if __name__=="__main__":unittest.main()
