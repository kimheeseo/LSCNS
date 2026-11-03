import sys
from pathlib import Path
import numpy as np
import pytest
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from vector_modes import Configuration,solve,he11_exact_step,integrate_intensity,export_result,make_index


def test_gaussian_intensity_integrals():
    x=np.linspace(-20,20,201);X,Y=np.meshgrid(x,x);w=5
    I=np.exp(-2*(X*X+Y*Y)/w**2)
    r=integrate_intensity(I,x)
    assert abs(r['MFD_2nd_um']-2*w)<.01
    assert abs(r['MFD_PII_um']-2*w)<.02
    assert abs(r['Aeff_um2']-np.pi*w*w)<.01


def test_integral_scaling():
    x=np.linspace(-20,20,121);X,Y=np.meshgrid(x,x);I=np.exp(-((X-1)**2+(Y+2)**2)/15)
    a,b=integrate_intensity(I,x),integrate_intensity(17.1*I,x)
    for k in ('MFD_2nd_um','MFD_PII_um','Aeff_um2'):
        assert abs(a[k]-b[k])<1e-9


def test_exact_vector_he11_consistency():
    c=Configuration(grid_n=69)
    o=solve(c,modes=4);r=he11_exact_step(c)
    assert len(o['modes'])>=2
    assert abs(o['modes'][0]['neff']-r)<2e-5
    assert abs(o['modes'][0]['neff']-o['modes'][1]['neff'])<1e-8
    assert o['modes'][0]['residual']<1e-8
    assert o['modes'][0]['edge_power_fraction']<1e-3
    assert abs(sum(o['modes'][0][x+'_intensity_fraction'] for x in ('Ex','Ey','Ez'))-1)<1e-9
    assert 0<o['modes'][0]['Ez_intensity_fraction']<0.02


def test_grid_comparison_error_bounded():
    c=Configuration(grid_n=45)
    val=solve(c,modes=3)['modes'][0]['neff'];ref=he11_exact_step(c)
    assert abs(val-ref)<1.2e-4


@pytest.mark.parametrize('profile,other',[
    ('graded',{'core_radius_um':5.0,'delta_n':0.006}),
    ('trench',{'core_radius_um':4.1,'delta_n':0.005,'trench_start_um':6,'trench_end_um':9}),
    ('multilayer',{'shells':((4.,.006),(5.,.002),(7.,0.),(9.,-.0008))}),
])
def test_other_silica_profiles(profile,other):
    c=Configuration(profile=profile,grid_n=57,**other)
    s=solve(c,modes=3)
    assert len(s['modes'])>=2
    r=s['modes'][0]
    assert 1.44<r['neff']<1.5
    assert 0<r['Aeff_um2']<500
    assert 0<r['MFD_2nd_um']<40
    assert r['residual']<1e-8


def test_export_and_map(tmp_path):
    c=Configuration(grid_n=45)
    r=solve(c,modes=3)
    out=export_result(r,tmp_path/'out.json',field_stride=2)
    assert out['schema']=='LSCNS_vector_mode_v1'
    m=out['modes'][0]
    assert m['Aeff_um2']>0
    f=m['field_map'];n=f['grid_size']
    assert n>=20
    for key in ('Ex','Ey','Ez','Itotal'):
        mat=np.asarray(f['intensity_normalized_to_max_total'][key])
        assert mat.shape==(n,n)
        assert np.isfinite(mat).all()


def test_reject_invalid_geometry():
    with pytest.raises(ValueError): Configuration(grid_n=32).validate()
    with pytest.raises(ValueError): Configuration(profile='hcf').validate()
    with pytest.raises(ValueError): Configuration(profile='trench',trench_start_um=1).validate()
