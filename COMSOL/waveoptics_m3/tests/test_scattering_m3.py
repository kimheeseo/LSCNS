"""M3 acceptance contract, written before scattering solver implementation."""
import numpy as np
import pytest
from analytic_scattering import guide_beta, layer_s, cylinder_field, cylinder_width


def api():
    from waveoptics.scattering import guide_mesh, cylinder_mesh, solve_ports, solve_plane_wave, PML
    return guide_mesh,cylinder_mesh,solve_ports,solve_plane_wave,PML


@pytest.mark.parametrize('polarization',['TE','TM'])
def test_independent_oracles(polarization):
    n,L,W,wl=1.3,.9,.7,1.
    m=1 if polarization=='TE' else 0
    r,t=layer_s([n],[L],wl,W,m,polarization)
    assert abs(r)<1e-14
    assert abs(t-np.exp(1j*guide_beta(n,wl,W,m)*L))<1e-14
    assert cylinder_width(.2,1.,1.,1.,polarization)<1e-28
    assert abs(cylinder_width(.2,1.5,1.,1.,polarization,20)/cylinder_width(.2,1.5,1.,1.,polarization,24)-1)<1e-12


@pytest.mark.parametrize('polarization',['TE','TM'])
def test_straight_guide_three_meshes(polarization):
    gm,_,solve,_,_=api(); n,L,W,wl=1.3,.9,.7,1.
    m=1 if polarization=='TE' else 0
    beta=guide_beta(n,wl,W,m).real; errors=[]; diameters=[]
    for h in [.08,.04,.02]:
        sol=solve(gm([0,L],[0,W],h),n,wl,polarization,incident_port='left',incident_mode=0)
        errors.append(abs(sol.s_parameters['right'][0]-np.exp(1j*beta*L)))
        diameters.append(sol.diagnostics['h_max_um'])
        assert sol.diagnostics['linear_residual']<1e-10
        assert abs(sol.s_parameters['left'][0])<.02
    slopes=np.log(np.array(errors[:-1])/errors[1:])/np.log(np.array(diameters[:-1])/diameters[1:])
    assert errors[-1]<5e-3
    assert np.all((slopes>1.8)&(slopes<2.2)),(errors,slopes)


@pytest.mark.parametrize('polarization',['TE','TM'])
def test_dielectric_layer_s(polarization):
    gm,_,solve,_,_=api(); Ls=[.3,.25,.35]; ns=[1.2,1.6,1.2]; W=.6; wl=1.
    def index(x): return np.where((x[0]>.3)&(x[0]<.55),1.6,1.2)
    sol=solve(gm([0,.3,.55,.9],[0,W],.02),index,wl,polarization)
    r,t=layer_s(ns,Ls,wl,W,1 if polarization=='TE' else 0,polarization)
    assert abs(sol.s_parameters['left'][0]-r)<5e-3
    assert abs(sol.s_parameters['right'][0]-t)<5e-3
    power=sum(np.sum(np.abs(s)**2) for s in sol.s_parameters.values())
    assert abs(power-1)<1e-9


@pytest.mark.parametrize('polarization',['TE','TM'])
def test_multimode_reciprocity_and_unitarity(polarization):
    gm,_,solve,_,_=api(); mesh=gm([0,.3,.55,.9],[0,.6],.035)
    def n(x): return np.where((x[0]>.3)&(x[0]<.55),1.6,1.2)
    first=solve(mesh,n,1.,polarization)
    count=len(first.s_parameters['left']); cols=[]
    for port in ['left','right']:
        for mode in range(count):
            s=solve(mesh,n,1.,polarization,incident_port=port,incident_mode=mode)
            cols.append(np.concatenate([s.s_parameters['left'],s.s_parameters['right']]))
    S=np.array(cols).T
    assert np.linalg.norm(S-S.T)<1e-9
    assert np.linalg.norm(S.conj().T@S-np.eye(2*count))<1e-9


@pytest.mark.parametrize('polarization',['TE','TM'])
def test_asymmetric_interface_power_normalization(polarization):
    gm,_,solve,_,_=api()
    def n(x): return np.where(x[0]<.4,1.2,1.6)
    s=solve(gm([0,.4,.9],[0,.6],.02),n,1.,polarization)
    r,t=layer_s([1.2,1.6],[.4,.5],1.,.6,1 if polarization=='TE' else 0,polarization)
    assert abs(s.s_parameters['left'][0]-r)<5e-3
    assert abs(s.s_parameters['right'][0]-t)<5e-3
    assert abs(sum(np.sum(abs(v)**2) for v in s.s_parameters.values())-1)<1e-9


@pytest.mark.parametrize('polarization',['TE','TM'])
def test_pml_truncation_matches_analytic_reflection(polarization):
    gm,_,solve,_,PML=api(); L=1.5; start=1.; sigma=1.2; thickness=.5
    pml=PML((0,start),(0,.7),right=thickness,strength=sigma,power=2)
    s=solve(gm([0,start,L],[0,.7],.015),1.3,1.,polarization,pml=pml,ports=('left',))
    beta=guide_beta(1.3,1.,.7,1 if polarization=='TE' else 0).real
    exact=-np.exp(2j*beta*L-2*beta*sigma*thickness/3)
    assert abs(s.s_parameters['left'][0]-exact)<5e-3


@pytest.mark.parametrize('polarization',['TE','TM'])
def test_cylinder_three_meshes(polarization):
    _,cm,_,solve,PML=api(); pml=PML((-.8,.8),(-.8,.8),left=.6,right=.6,bottom=.6,top=.6,strength=4.)
    errors=[]; sizes=[]; widths=[]
    theta=np.arange(720)*2*np.pi/720; xy=.55*np.array([np.cos(theta),np.sin(theta)])
    reference=cylinder_field(xy,.2,1.5,1.,1.,polarization)
    exact_width=cylinder_width(.2,1.5,1.,1.,polarization)
    for h in [.08,.04,.02]:
        sol=solve(cm(.2,pml,h),1.5,1.,1.,polarization,pml=pml)
        field=sol.sample(xy)
        errors.append(float(np.linalg.norm(field-reference)/np.linalg.norm(reference)))
        sizes.append(sol.diagnostics['h_max_um'])
        widths.append(sol.scattering_width(.55,orders=8))
        assert sol.diagnostics['linear_residual']<1e-10
    slopes=np.log(np.array(errors[:-1])/errors[1:])/np.log(np.array(sizes[:-1])/sizes[1:])
    assert errors[-1]<.015,(errors,slopes)
    assert abs(widths[-1]/exact_width-1)<.015
    assert np.all((slopes>1.65)&(slopes<2.4)),(errors,slopes)


@pytest.mark.parametrize('polarization',['TE','TM'])
def test_pml_background_no_scattering(polarization):
    _,cm,_,solve,PML=api(); p=PML((-.6,.6),(-.6,.6),left=.4,right=.4,top=.4,bottom=.4)
    s=solve(cm(.2,p,.07),1.,1.,1.,polarization,pml=p)
    assert np.linalg.norm(s.field)<1e-13


@pytest.mark.parametrize('polarization',['TE','TM'])
def test_absorbing_layer_passivity(polarization):
    gm,_,solve,_,_=api()
    def n(x): return np.where((x[0]>.3)&(x[0]<.55),1.6+.03j,1.2)
    s=solve(gm([0,.3,.55,.9],[0,.6],.02),n,1.,polarization)
    r,t=layer_s([1.2,1.6+.03j,1.2],[.3,.25,.35],1.,.6,1 if polarization=='TE' else 0,polarization)
    assert abs(s.s_parameters['left'][0]-r)<5e-3
    assert abs(s.s_parameters['right'][0]-t)<5e-3
    assert sum(np.sum(abs(v)**2) for v in s.s_parameters.values())<1


def test_slab_port_from_independent_dispersion():
    gm,_,solve,_,_=api()
    from analytic_slab import exact_modes
    def n(x): return np.where(abs(x[1])<.25,1.5,1.)
    mesh=gm([0,.6],[-2.,-.25,.25,2.],.025)
    for pol in ['TE','TM']:
        s=solve(mesh,n,1.,pol)
        beta=exact_modes(1.5,1.,.5,1.,pol)[0].beta_per_um
        assert abs(s.port_modes['left'].beta[0].real-beta)<.003
        assert abs(s.s_parameters['left'][0])<.003
        assert abs(s.s_parameters['right'][0]-np.exp(1j*beta*.6))<.004


def test_invalid_inputs_and_cutoff():
    gm,_,solve,_,PML=api(); mesh=gm([0,.4],[0,.5],.1)
    for kw in [dict(polarization='bad'),dict(wavelength_um=0),dict(incident_mode=99),dict(incident_port='top')]:
        args=dict(wavelength_um=1.,polarization='TE');args.update(kw)
        with pytest.raises(ValueError): solve(mesh,1.2,**args)
    with pytest.raises(ValueError): solve(mesh,1.2+.01j,1.,'TE')
    with pytest.raises(ValueError): PML((0,1),(0,1),right=-1)
    with pytest.raises(ValueError): solve(mesh,-1.,1.,'TE')
    with pytest.raises(ValueError): solve(mesh,1.,1.,'TE') # exactly at continuum cutoff
