import sys
from pathlib import Path
import math
import numpy as np
import pytest
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from hcf_complex_fem import (HCFDesign,geometry,is_silica,mesh_geometry,assemble,solve,
                             silica_index,ideal_neff,peaks_from_fem_sweep,design_from_dict)

def test_default_geometry_matches_browser_reference():
    d=HCFDesign();g=geometry(d)
    assert len(g['tubes'])==12
    assert g['outer_tube_diameter_um']==pytest.approx(24.96,abs=1e-10)
    assert g['neighbor_gap_um']==pytest.approx(3.77,abs=1e-10)
    assert g['radius_extent_um']==pytest.approx(41.21,abs=1e-10)
    assert g['core_radius_um']==16

def test_three_nested_tubes_and_numerical_mesh():
    d=HCFDesign(nesting=3,thirdRatio=.55)
    g=geometry(d)
    assert len(g['tubes'])==18
    xy,tri,boundary,window,half=mesh_geometry(d,g)
    assert len(xy)>3000 and len(tri)>5000
    assert boundary.any() and np.max(np.abs(xy))<=half+1e-7

def test_tube_wall_index_annulus():
    d=HCFDesign(nesting=1)
    t=geometry(d)['tubes'][0]
    assert bool(is_silica(np.array([t['x']+t['r']-t['wall']/2]),np.array([t['y']]),[t])[0])
    assert not bool(is_silica(np.array([t['x']]),np.array([t['y']]),[t])[0])
    assert not bool(is_silica(np.array([0]),np.array([0]),geometry(d)['tubes'])[0])

def test_silica_sellmeier_1550():
    assert silica_index(1550)==pytest.approx(1.4440236,rel=1e-5)

def test_pml_complex_assembly_and_modes():
    d=HCFDesign(background_h_um=2.5,wall_edge_h_um=.90,n_modes=4)
    A,M,xy,lump,core,diag=assemble(d)
    assert A.shape[0]==M.shape[0] and A.nnz>4*A.shape[0]
    assert np.any(np.abs(M.data.imag)>1e-15)
    assert np.any(lump>0) and core.any()
    r=solve(d)
    assert 0.998<r['neff_real']<1.002
    assert r['ideal_J0_overlap']>.7
    assert r['converged_and_validated'] is False
    assert r['full_vector'] is False
    assert r['mode_identification_pass']

def test_complex_neff_loss_formula():
    d=HCFDesign()
    r=solve(d)
    if r['neff_imag']>0:
        expected=40*math.pi*1000*r['neff_imag']/(d.wavelengthNm*1e-9*math.log(10))
        assert r['confinement_loss_dB_km_unverified']==pytest.approx(expected)
    else:
        assert r['confinement_loss_dB_km_unverified'] is None

def test_mode_field_is_from_fem():
    r=solve(HCFDesign(background_h_um=2.5,wall_edge_h_um=.9),include_field=True)
    assert len(r['mode_points'])>100
    assert max(z['intensity'] for z in r['mode_points'])>0.9
    assert min(z['intensity'] for z in r['mode_points'])>=0

def test_spectrum_extrema_are_marked_unverified():
    template=solve(HCFDesign(background_h_um=3.5,wall_edge_h_um=1.2,n_modes=3))
    arr=[]
    for x,y in [(1300,2.),(1400,1.),(1500,3.),(1600,2.)]:
        z={**template,'input':{**template['input'],'wavelengthNm':x},
          'confinement_loss_dB_km_unverified':y,'mode_identification_pass':True}
        arr.append(z)
    feat=peaks_from_fem_sweep(arr)
    assert [(x['kind'],x['wavelength_nm']) for x in feat]==[
        ('antiresonance_loss_dip_candidate',1400),('resonance_loss_peak_candidate',1500)]
    arr[1]['mode_identification_pass']=False
    assert not peaks_from_fem_sweep(arr)

def test_design_validation_rejects_overlap_and_out_of_range():
    with pytest.raises(ValueError):geometry(HCFDesign(autoFit=False,tubeCount=12))
    with pytest.raises(ValueError):design_from_dict({'wall1Um':-2})
    with pytest.raises(ValueError):HCFDesign(wavelengthNm=9999).validate()


def test_mode_rank_data_has_all_modes():
    d=HCFDesign(n_modes=3,wall_edge_h_um=1.)
    r=solve(d)
    assert len(r['ranked_modes'])==3
    assert r['ranked_modes'][0]['J0_overlap']>=r['ranked_modes'][-1]['J0_overlap']
