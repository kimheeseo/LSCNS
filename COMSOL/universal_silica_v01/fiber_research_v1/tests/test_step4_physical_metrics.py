"""STEP4 v0.8.0 analytical unit tests. Do NOT mark numerical FEM verification as passed."""
import math
import sys
import unittest
from pathlib import Path
import numpy as np
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from physical_metrics import (C_M_S,confinement_loss,field_area_mfd,
    petermann_ii_radial,birefringence,mode_neff_spacing,nonlinear_gamma,
    hcf_effective_n2,glass_power_fraction,homer,arrow_resonance_lines,
    overlay_arrow_with_fem,match_modes_by_overlap,dispersion_from_tracked_sweep,
    surface_anticrossing_candidates,ENGINE_VERSION)


def grid(w=4.0):
    x=np.linspace(-18,18,361)
    xx,yy=np.meshgrid(x,x)
    return x, np.exp(-(xx*xx+yy*yy)/w**2).astype(complex)


class Step4FormulaTests(unittest.TestCase):
    def test_version(self):
        self.assertEqual(ENGINE_VERSION,"0.8.0")

    def test_dB_m_km_and_negative_sign(self):
        lam=1550.;imag=1e-9
        value=confinement_loss(imag,lam)
        expected=(20/math.log(10))*2*math.pi/(lam*1e-9)*imag
        self.assertAlmostEqual(value["alpha_total_dB_m"],expected,places=12)
        self.assertAlmostEqual(value["alpha_total_dB_km"],1000*expected,places=12)
        self.assertAlmostEqual(confinement_loss(-imag,lam)["alpha_total_dB_m"],-expected,places=12)
        self.assertFalse(confinement_loss(-imag,lam)["is_passive_sign"])
        self.assertAlmostEqual(confinement_loss(0,lam)["alpha_total_dB_m"],0)

    def test_aeff_and_mfd_gaussian(self):
        w=4.0;x,E=grid(w)
        z=np.zeros_like(E)
        r=field_area_mfd(x,x,E,z)
        self.assertAlmostEqual(r["Aeff_um2"],math.pi*w*w,delta=.15)
        self.assertAlmostEqual(r["MFD_2nd_um"],2*w,delta=.06)
        self.assertAlmostEqual(r["MFD_PII_like_um"],2*w,delta=.07)
        self.assertIn("NOT standardized",r["MFD_PII_status"])

    def test_petermann_radial(self):
        r=np.linspace(0,24,1601)
        f=np.exp(-r*r/4.**2)
        result=petermann_ii_radial(r,f)
        self.assertAlmostEqual(result["MFD_PetermannII_um"],8.,delta=.025)

    def test_mask_does_not_count_pml(self):
        x,E=grid()
        mask=np.ones_like(E,dtype=bool)
        mask[:5]=mask[-5:]=False
        a=field_area_mfd(x,x,E,np.zeros_like(E),mask)
        b=field_area_mfd(x,x,E,np.zeros_like(E))
        self.assertAlmostEqual(a["Aeff_um2"],b["Aeff_um2"],delta=.0001)

    def test_gamma_units_and_missing_n2(self):
        n2=2.2e-20;lam=1550.;aeff=125.
        g=nonlinear_gamma(n2,lam,aeff)
        self.assertAlmostEqual(g["gamma_W_inv_m_inv"],2*math.pi*n2/(lam*1e-9*aeff*1e-12))
        with self.assertRaises(ValueError): nonlinear_gamma(None,lam,aeff)

    def test_glass_n2_quartic_proxy_with_user_values(self):
        x=np.linspace(-1,1,5)
        ex=np.ones((5,5),complex);ey=np.zeros_like(ex)
        mask=np.zeros((5,5),bool);mask[:2]=True
        self.assertEqual(hcf_effective_n2(x,x,ex,ey,mask)["status"],"TODO_MATERIAL_N2")
        r=hcf_effective_n2(x,x,ex,ey,mask,2e-23,2e-20)
        expected=(2e-23)*.6+(2e-20)*.4
        self.assertAlmostEqual(r["n2_eff_m2_W"],expected)
        self.assertAlmostEqual(r["quartic_glass_weight"],.4)

    def test_poynting_fraction(self):
        x=np.linspace(-1,1,5)
        ex=np.ones((5,5),complex);zeros=np.zeros_like(ex)
        hy=np.ones_like(ex)/(376.73)
        glass=np.zeros((5,5),bool);glass[:2]=True
        p=glass_power_fraction(x,x,ex,zeros,zeros,hy,glass)
        self.assertAlmostEqual(p["P_glass_fraction"],.4,places=12)
        with self.assertRaises(ValueError):
            glass_power_fraction(x,x,ex,zeros,zeros,-hy,glass)

    def test_birefringence_beats_and_spacing(self):
        b=birefringence(1.45,1.44999,1550)
        self.assertAlmostEqual(b["B"],1e-5,places=10)
        self.assertAlmostEqual(b["beat_length_m"],.155,places=6)
        self.assertIsNone(birefringence(1.45,1.45,1550)["beat_length_m"])
        self.assertEqual(len(mode_neff_spacing([1.45,1.44,1.43])),3)

    def test_homer_verified_and_unverified(self):
        self.assertEqual(homer([.1],[.4])["status"],"TODO_MODE_CLASSIFICATION")
        h=homer([.10,.12],[.6,.8],mode_ids_verified=True)
        self.assertAlmostEqual(h["HOMER_ratio"],5.)
        with self.assertRaises(ValueError):
            homer([0.],[.5],mode_ids_verified=True)

    def test_arrow_units_and_overlay_no_synthetic_fem(self):
        val=arrow_resonance_lines(.45,1.45,1.00027,[1,2])
        expected=2000*.45*math.sqrt(1.45**2-1.00027**2)
        self.assertAlmostEqual(val["resonance_lines"][0]["lambda_nm"],expected)
        self.assertAlmostEqual(val["resonance_lines"][1]["lambda_nm"],expected/2)
        self.assertEqual(overlay_arrow_with_fem(val)["fem_curves"],[])
        with self.assertRaises(ValueError):
            overlay_arrow_with_fem(val,[{"overlap_verified":False}])

    def test_dispersion_quadratic_units_slope(self):
        lam=np.arange(1450.,1660.,10.)
        q=-1.7e-9  # analytic test polynomial coefficient, not a material constant
        neff=1.45+q*(lam-1550)**2
        result=dispersion_from_tracked_sweep(lam,neff,np.ones(len(lam)-1))
        idx=int(np.where(lam==1550)[0][0])
        D=-(1550*1e-9/C_M_S)*(2*q*1e18)*1e6
        S=-(1e-9/C_M_S)*(2*q*1e18)*1e6
        self.assertAlmostEqual(result["D_ps_nm_km"][idx],D,places=6)
        self.assertAlmostEqual(result["S_ps_nm2_km"][idx],S,places=7)
        with self.assertRaises(ValueError):
            dispersion_from_tracked_sweep(lam,neff,[.5]*(len(lam)-1))

    def test_zero_dispersion_in_sweep(self):
        lam=np.arange(1480.,1630.,10.)
        x=lam-1550.
        neff=1.45+2e-12*x**3  # analytic unit-test polynomial, not fiber dispersion
        result=dispersion_from_tracked_sweep(lam,neff,np.ones(len(lam)-1))
        self.assertEqual(len(result["zero_dispersion_lambda_nm"]),1)
        self.assertAlmostEqual(result["zero_dispersion_lambda_nm"][0],1550.,places=5)

    def test_overlap_swapped_modes_and_orthogonality(self):
        pts=np.array([[0,0],[1,0],[1,1],[0,1]],float)
        wt=np.ones(4)
        xmode={"points_um":pts,"weights_um2":wt,
               "Ex":np.array([1,1,1,1],complex),"Ey":np.zeros(4,complex)}
        ymode={"points_um":pts,"weights_um2":wt,
               "Ex":np.zeros(4,complex),"Ey":np.array([1,1,1,1],complex)}
        response=match_modes_by_overlap([xmode,ymode],[ymode,xmode])
        self.assertTrue(response["all_tracked"])
        self.assertEqual([(x["previous_index"],x["current_index"]) for x in response["assignments"]],[(0,1),(1,0)])
        moved=dict(xmode,points_um=np.array(pts)+.01)
        with self.assertRaises(ValueError):
            match_modes_by_overlap([xmode],[moved])

    def test_anticrossing_candidate_requires_physics(self):
        self.assertEqual(surface_anticrossing_candidates({})["status"],"TODO_MODE_IDENTITY")
        lam=[1530,1540,1550,1560,1570]
        p={"overlap_verified":True,"glass_mode_identified":True,
           "wavelength_nm":lam,
           "neff_core":[1.1,1.08,1.072,1.06,1.04],
           "neff_glass":[1.01,1.05,1.071,1.09,1.12],
           "glass_fraction_core":[.1,.15,.45,.8,.9],
           "glass_fraction_glass":[.9,.85,.55,.2,.1]}
        r=surface_anticrossing_candidates(p)
        self.assertEqual(r["status"],"candidate_screening_only")
        self.assertEqual(len(r["candidates"]),1)
        self.assertEqual(r["candidates"][0]["lambda_nm"],1550)

    def test_reject_nan_and_bad_field(self):
        x,E=grid()
        with self.assertRaises(ValueError):
            field_area_mfd(x,x,E,np.zeros((4,4),complex))
        with self.assertRaises(ValueError):
            arrow_resonance_lines(.45,1.,1.01)
        with self.assertRaises(ValueError):
            confinement_loss(float("nan"),1550)

if __name__=="__main__":unittest.main()
