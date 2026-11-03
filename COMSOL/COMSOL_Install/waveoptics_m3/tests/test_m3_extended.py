"""Additional safety, mode-conversion and oblique-incidence checks."""
import numpy as np
import pytest
from waveoptics.scattering import guide_mesh,cylinder_mesh,solve_ports,solve_plane_wave,solve_s_matrix,PML,Z0
from analytic_scattering import cylinder_field


@pytest.mark.parametrize('pol',['TE','TM'])
def test_mode_conversion_s_matrix(pol):
    mesh=guide_mesh([0,.3,.55,.9],[0,.45,.9],.05)
    def n(x):return np.where((x[0]>.3)&(x[0]<.55)&(x[1]>.45),1.6,1.2)
    S,channels,diag=solve_s_matrix(mesh,n,1.,pol)
    assert np.linalg.norm(S-S.T)<1e-9
    assert np.linalg.norm(S.conj().T@S-np.eye(len(channels)))<1e-9
    assert abs(S[1,0])>1e-3
    assert max(d['linear_residual'] for d in diag)<1e-10


@pytest.mark.parametrize('pol',['TE','TM'])
def test_oblique_cylinder_and_fields(pol,tmp_path):
    p=PML((-.8,.8),(-.8,.8),left=.6,right=.6,top=.6,bottom=.6)
    s=solve_plane_wave(cylinder_mesh(.2,p,.025),1.5,1.,1.,pol,p,angle_rad=np.pi/4)
    th=np.arange(720)*2*np.pi/720;xy=.55*np.array([np.cos(th),np.sin(th)])
    rotation=np.array([[np.cos(np.pi/4),np.sin(np.pi/4)],[-np.sin(np.pi/4),np.cos(np.pi/4)]])
    ref=cylinder_field(rotation@xy,.2,1.5,1.,1.,pol)
    assert np.linalg.norm(s.sample(xy)-ref)/np.linalg.norm(ref)<.015
    inc=np.exp(1j*2*np.pi*(xy[0]+xy[1])/np.sqrt(2))
    assert np.max(abs(s.sample(xy,total=True)-s.sample(xy)-inc))<1e-13
    coords,weights,E,H=s.quadrature_fields(total=True)
    assert E.shape==H.shape==(3,len(weights));assert np.all(np.isfinite(E));assert np.all(np.isfinite(H))
    path=tmp_path/'fields.npz';s.save_npz(path)
    with np.load(path,allow_pickle=False) as data:
        assert data['field_kind']=='scattered'
        assert np.array_equal(data['field'],s.field)
        assert data['index_at_quadrature'].shape==s.basis.dx.shape


@pytest.mark.parametrize('pol',['TE','TM'])
def test_background_field_maxwell_units(pol):
    p=PML((-.6,.6),(-.6,.6),left=.4,right=.4,top=.4,bottom=.4)
    s=solve_plane_wave(cylinder_mesh(.2,p,.1),1.,1.,1.,pol,p)
    x,_,E,H=s.quadrature_fields(total=True);inc=np.exp(1j*2*np.pi*x[0])
    if pol=='TE':
        assert np.max(abs(E[2]-inc))<1e-14
        assert np.max(abs(H[1]+inc/Z0))<1e-14
    else:
        assert np.max(abs(H[2]-inc))<1e-14
        assert np.max(abs(E[1]-Z0*inc))<1e-12


def test_general_callable_index_and_pml_validation():
    p=PML((-.6,.6),(-.6,.6),left=.4,right=.4,top=.4,bottom=.4)
    mesh=guide_mesh([-1,-.6,-.2,.2,.6,1],[-1,-.6,-.15,.15,.6,1],.08)
    def n(x):return np.where((abs(x[0])<.2)&(abs(x[1])<.15),1.4,1.)
    s=solve_plane_wave(mesh,n,1.,1.,'TE',p)
    assert np.linalg.norm(s.field)>0
    assert s.diagnostics['linear_residual']<1e-10
    with pytest.raises(ValueError):solve_plane_wave(mesh,lambda x:np.full(x.shape[1:],1.4),1.,1.,'TE',p)
    with pytest.raises(ValueError):s.scattering_width(.1)
    with pytest.raises(ValueError):s.scattering_width(.7)
    with pytest.raises(ValueError):solve_ports(mesh,1.,1.,'TE',pml=p)


@pytest.mark.parametrize('pol',['TE','TM'])
def test_absorbing_cylinder_independent_reference(pol):
    p=PML((-.8,.8),(-.8,.8),left=.6,right=.6,top=.6,bottom=.6)
    s=solve_plane_wave(cylinder_mesh(.2,p,.025),1.5+.03j,1.,1.,pol,p)
    th=np.arange(720)*2*np.pi/720;xy=.55*np.array([np.cos(th),np.sin(th)])
    ref=cylinder_field(xy,.2,1.5+.03j,1.,1.,pol)
    assert np.linalg.norm(s.sample(xy)-ref)/np.linalg.norm(ref)<.015


def test_dense_point_sampler_affine_exactness():
    p=PML((-.6,.6),(-.6,.6),left=.4,right=.4,top=.4,bottom=.4)
    s=solve_plane_wave(cylinder_mesh(.2,p,.08),1.,1.,1.,'TE',p)
    x=s.geometry.mesh.p;s.field=(2+1j)*x[0]-(1+2j)*x[1]+.7
    lin=np.linspace(-.97,.97,121);xx,yy=np.meshgrid(lin,lin);xy=np.array([xx.ravel(),yy.ravel()])
    expected=(2+1j)*xy[0]-(1+2j)*xy[1]+.7
    assert np.max(abs(s.sample(xy)-expected))<1e-13
    with pytest.raises(ValueError):s.sample(np.array([[2.],[2.]]))


def test_diagnostics_json_serialization():
    import json
    s=solve_ports(guide_mesh([0,.4],[0,.6],.1),1.2,1.,'TE')
    decoded=json.loads(json.dumps(s.diagnostics))
    assert decoded['dofs']==s.basis.N
