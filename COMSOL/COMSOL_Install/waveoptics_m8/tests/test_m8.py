"""Tests of physical references, independent formulations and error handling."""
from pathlib import Path
import sys
import numpy as np
import pytest
ROOT=Path(__file__).resolve().parents[1]
sys.path[:0]=[str(ROOT/'src'),str(ROOT.parent/'waveoptics_m7'/'src')]
from fiber_modes import Fiber,lp01_analytic
from multipole_reference import solve_reference,boundary_matrix
from capillary_fem import solve_capillary
from conforming_mcf import solve_conforming,coherent_transfer,build_mesh
from mcf_hcf import capillary_neff,capillary_dispersion


def test_independent_infinite_domain_reference():
    f=Fiber();a=solve_reference(f,16,4);b=solve_reference(f,16,8,samples=512)
    assert b['index_split']==pytest.approx(a['index_split'],rel=1e-8)
    assert max(b['boundary_relative_singular'])<1e-10
    # Test MORE boundary harmonics than were used in finding the root.
    for n,p in [(b['n_even'],1),(b['n_odd'],-1)]:
        matrix=boundary_matrix(n,f,16,8,p,512)
        _,_,v=np.linalg.svd(matrix)
        padded=np.r_[v[-1],np.zeros(6)]
        full=boundary_matrix(n,f,16,14,p,1024)
        assert np.linalg.norm(full@padded)/np.linalg.norm(full)<1e-7


def test_large_pitch_limit_matches_single_core_bessel():
    f=Fiber();r=solve_reference(f,60,6);single=lp01_analytic(f)['n_eff']
    assert r['n_even']==pytest.approx(single,abs=3e-11)
    assert r['n_odd']==pytest.approx(single,abs=3e-11)


def test_curved_mesh_matches_material_interfaces():
    f=Fiber();mesh,inside=build_mesh(f,16,28,1.)
    c=8.;a=f.radius_um
    r=np.hypot(mesh.p[0]-c,mesh.p[1])
    assert np.count_nonzero(inside)>10 and np.count_nonzero(~inside)>10
    # Every facet joining core/clad triangles has all P2 nodes on the circle.
    from skfem import Basis,ElementTriP2
    shared=np.flatnonzero((mesh.f2t[1]>=0)&(inside[mesh.f2t[0]]!=inside[np.maximum(mesh.f2t[1],0)]))
    assert len(shared)>5
    b=Basis(mesh,ElementTriP2())
    ids=np.r_[mesh.facets[:,shared].ravel(),b.dofs.facet_dofs[:,shared].ravel()]
    rr=np.hypot(b.doflocs[0,ids]-c,b.doflocs[1,ids])
    assert np.max(abs(rr-a))<1e-10
    assert np.all(b.dx>0)


def test_two_core_fem_against_mesh_free_reference():
    f=Fiber();r=solve_reference(f);s=solve_conforming(f,16,32,.7)
    assert max(abs(s[k]-r[k]) for k in ['n_even','n_odd'])<3e-8
    assert abs(s['coupling_length_mm']/r['coupling_length_mm']-1)<3e-5
    assert max(s['residuals'])<1e-9


@pytest.mark.parametrize('radius,wavelength',[(10,1.3),(15,1.55),(30,1.7)])
def test_capillary_independent_fem(radius,wavelength):
    a=solve_capillary(wavelength,radius,elements=80);b=solve_capillary(wavelength,radius,elements=640)
    n=float(capillary_neff(wavelength,radius));D=float(capillary_dispersion(wavelength,radius))
    assert abs(b['n_eff']-n)<1e-9
    assert abs(b['n_eff']-n)<abs(a['n_eff']-n)/10
    assert b['D_ps_nm_km']==pytest.approx(D,rel=1e-6)


def test_transfer_power_and_units():
    length=np.array([0.,.005,.01,.02]);t=coherent_transfer(length,10)
    assert t['P2']==pytest.approx([0,.5,1,0],abs=1e-14)
    assert t['P1']+t['P2']==pytest.approx(np.ones(4),abs=1e-14)


@pytest.mark.parametrize('args',[dict(elements=True),dict(elements=3),dict(elements=4.5),dict(radius_um=np.nan),dict(wavelength_um=-1),dict(n_air=.9)])
def test_invalid_capillary_fem(args):
    with pytest.raises(ValueError):solve_capillary(**args)


@pytest.mark.parametrize('pitch',[0,8.2,np.nan,np.inf])
def test_invalid_mcf_pitch(pitch):
    with pytest.raises(ValueError):solve_reference(Fiber(),pitch)


@pytest.mark.parametrize('length,Lc',[(-1,10),(np.nan,10),(1,0),(1,np.inf)])
def test_invalid_transfer(length,Lc):
    with pytest.raises(ValueError):coherent_transfer(length,Lc)
