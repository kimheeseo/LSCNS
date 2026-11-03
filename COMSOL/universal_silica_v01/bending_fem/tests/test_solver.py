"""Numerical tests for an independently computed scalar 2D FEM/PML eigenvalue."""
from dataclasses import replace
from pathlib import Path
import sys
import numpy as np
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from solver import Fiber, silica_index, index_field, assemble, solve

def model():
    return Fiber(core_radius_um=4.1, delta_n_core=.005,
                 bend_diameter_mm=10., window_um=22.,
                 pml_thickness_um=10.85, grid_h_um=.7, n_modes=10)

def test_silica_sellmeier_and_matrix():
    assert abs(silica_index(1550.)-1.4440236217)<1e-7
    A,M,coords,nx,h=assemble(model())
    assert A.shape==M.shape
    assert A.shape[0]==(nx-2)**2
    assert np.iscomplexobj(A.data)
    assert h>0

def test_complex_mode_and_physical_sign():
    y=solve(model())
    assert y["neff_imag"]>0
    assert y["bend_loss_dB_per_turn"]>0
    assert y["core_fraction"]>.7
    assert y["physical_fraction"]>.8
    assert y["is_fully_validated"] is False
    # Stable regression only on this *illustrative* mesh; not a paper reference.
    assert abs(y["bend_loss_dB_per_turn"]-67.5562)<1

def test_straight_limit():
    y=solve(replace(model(),bend_diameter_mm=0))
    assert abs(y["neff_imag"])<1e-7
    assert y["bend_loss_dB_per_turn"] is None
    assert y["core_fraction"]>.9

def test_mesh_refinement_not_automatically_certified():
    x=solve(replace(model(),grid_h_um=.6))
    y=solve(replace(model(),grid_h_um=.5))
    assert abs(x["neff_real"]-y["neff_real"])<3e-4
    relative=100*abs(x["bend_loss_dB_per_turn"]-y["bend_loss_dB_per_turn"])/y["bend_loss_dB_per_turn"]
    assert relative<2
    assert x["is_fully_validated"] is False

def test_pml_thickness_sensitivity():
    a=solve(replace(model(),pml_thickness_um=6.2))
    b=solve(replace(model(),pml_thickness_um=10.85))
    assert abs(a["bend_loss_dB_per_turn"]-b["bend_loss_dB_per_turn"])>1
    assert b["is_fully_validated"] is False

def test_mode_intensity_from_actual_eigenvector():
    x=solve(replace(model(),grid_h_um=.9),include_field=True)
    field=np.asarray(x["mode_intensity_image"])
    assert field.ndim==2 and field.size>1000
    assert abs(field.max()-1)<1e-12
    assert field.min()>=0

def test_invalid_geometry():
    import pytest
    with pytest.raises(ValueError):
        solve(replace(model(),bend_diameter_mm=-1))
    with pytest.raises(ValueError):
        solve(replace(model(),window_um=4))
    with pytest.raises(ValueError):
        solve(replace(model(),pml_thickness_um=0))
