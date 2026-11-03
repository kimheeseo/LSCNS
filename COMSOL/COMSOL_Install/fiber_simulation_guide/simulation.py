"""Shared input/output protocol for two ideal step-index fiber examples."""
from dataclasses import replace
import numpy as np
import pandas as pd
from scipy.optimize import brentq
from fiber_model import Fiber, silica_index, lp01_analytic, radial_lp01_aeff, solve_radial, solve_modes, scalar_metrics, dispersion_lp01

PRESETS={
 'G.654.E':dict(analysis_wavelength_nm=1550.,reference_wavelength_nm=1550.,radius_um=5.5,target_mfd_um=12.5,manual_delta_n=.0034785793846155,n2_m2_W=2.2e-20),
 'G.657.A1':dict(analysis_wavelength_nm=1310.,reference_wavelength_nm=1310.,radius_um=4.1,target_mfd_um=9.2,manual_delta_n=.0046215295453259,n2_m2_W=2.6e-20)}


def positive(value,name):
    if not np.isfinite(value) or value<=0:raise ValueError(name+' must be positive and finite')
    return float(value)


def parse_counts(text,odd=False):
    values=[int(x.strip()) for x in text.split(',')]
    if not values or any(x<15 if odd else x<4 for x in values):raise ValueError('invalid mesh list')
    if odd and any(x%2!=1 for x in values):raise ValueError('2D grids must be odd integers >=15')
    return values


def prepare(config,fit_mfd=True,outer_um=40.):
    c=dict(config)
    for key in ('analysis_wavelength_nm','reference_wavelength_nm','radius_um','target_mfd_um','manual_delta_n','n2_m2_W'):
        c[key]=positive(c[key],key)
    if not np.isfinite(outer_um) or outer_um<=2*c['radius_um']:raise ValueError('outer radius must exceed twice core radius')
    refwl=c['reference_wavelength_nm']/1000
    nc=silica_index(refwl)
    def candidate(dn):return Fiber(wavelength_um=refwl,radius_um=c['radius_um'],delta_n=float(dn),domain_halfwidth_um=outer_um)
    def deviation(dn):return radial_lp01_aeff(candidate(dn))['MFD_petermann_ii_um']-c['target_mfd_um']
    if fit_mfd:
        v=np.linspace(.85,2.4046,100)
        dns=np.sqrt(nc**2+(v*refwl/(2*np.pi*c['radius_um']))**2)-nc
        vals=[deviation(dn) for dn in dns]
        pairs=[(a,b) for a,b,ya,yb in zip(dns[:-1],dns[1:],vals[:-1],vals[1:]) if ya*yb<=0]
        if not pairs:raise ValueError('MFD calibration has no supported V<2.405 solution; change assumed radius/target or use manual delta_n')
        dn=brentq(deviation,*pairs[0],xtol=1e-13)
    else:dn=c['manual_delta_n']
    f=replace(candidate(dn),wavelength_um=c['analysis_wavelength_nm']/1000)
    lp01_analytic(f) # raises on unsupported V; do not silently extrapolate
    ref=replace(f,wavelength_um=refwl);lp01_analytic(ref)
    return f,c,fit_mfd


def summarize(config,fit_mfd=True,outer_um=40.):
    f,c,fit=prepare(config,fit_mfd,outer_um)
    a=lp01_analytic(f);m=radial_lp01_aeff(f)
    atref=radial_lp01_aeff(replace(f,wavelength_um=c['reference_wavelength_nm']/1000))
    D=dispersion_lp01(f)
    S=(dispersion_lp01(replace(f,wavelength_um=f.wavelength_um+.005))-dispersion_lp01(replace(f,wavelength_um=f.wavelength_um-.005)))/10
    gamma=2*np.pi*c['n2_m2_W']/(f.wavelength_um*1e-6*m['Aeff_um2']*1e-12)*1000
    zero_nm=None;S0=None
    try:
        w0=brentq(lambda w:dispersion_lp01(replace(f,wavelength_um=w)),1.28,1.36,xtol=1e-9)
        zero_nm=w0*1000
        S0=(dispersion_lp01(replace(f,wavelength_um=w0+.005))-dispersion_lp01(replace(f,wavelength_um=w0-.005)))/10
    except ValueError:
        pass # unsupported wavelength or no sign change: do not invent a lambda0
    row=dict(lambda0_nm=zero_nm,S0_ps_nm2_km=S0,lambda0_status='FOUND_IN_1280_1360' if zero_nm is not None else 'NOT_RESOLVED_IN_SUPPORTED_INTERVAL',analysis_wavelength_nm=c['analysis_wavelength_nm'],MFD_reference_wavelength_nm=c['reference_wavelength_nm'],
      assumed_radius_um=c['radius_um'],delta_n=f.delta_n,delta_n_source='MFD_CALIBRATION' if fit else 'MANUAL',
      MFD_target_um=c['target_mfd_um'] if fit else None,MFD_PII_at_reference_um=atref['MFD_petermann_ii_um'],
      n_clad=f.clad,n_core=f.core,V=f.v,neff_Bessel=a['n_eff'],Aeff_um2=m['Aeff_um2'],
      MFD_PII_um=m['MFD_petermann_ii_um'],MFD_2nd_um=m['MFD_2nd_um'],D_ps_nm_km=D,S_ps_nm2_km=S,
      n2_assumed_m2_W=c['n2_m2_W'],gamma_W_inv_km_inv=gamma)
    return f,row


def mesh_study(f,core_counts=(40,80,160,320),clad_ratio=4,grid_counts=(61,81,121,161),box_um=24.,run_2d=True):
    if not isinstance(clad_ratio,int) or clad_ratio<1:raise ValueError('clad_ratio must be a positive integer')
    ref=lp01_analytic(f);m=radial_lp01_aeff(f);rows=[]
    for n in core_counts:
        s=solve_radial(f,core_elements=n,clad_elements=clad_ratio*n)
        rows.append(dict(method='RADIAL',core_elements=n,clad_elements=clad_ratio*n,total_elements=(1+clad_ratio)*n,
           outer_radius_um=f.domain_halfwidth_um,neff=s['n_eff'],neff_abs_error=abs(s['n_eff']-ref['n_eff']),
           Aeff_um2=s['Aeff_um2'],Aeff_error_pct=100*(s['Aeff_um2']/m['Aeff_um2']-1),
           MFD_2nd_um=s['MFD_2nd_um'],MFD_error_pct=100*(s['MFD_2nd_um']/m['MFD_2nd_um']-1),residual=s['residual']))
    radial=pd.DataFrame(rows);rows=[];last=None
    if run_2d:
        ff=replace(f,domain_halfwidth_um=box_um)
        for n in grid_counts:
            last=solve_modes(ff,n=n,modes=1);s=scalar_metrics(last)
            rows.append(dict(method='2D',grid_n=n,triangles=len(last['triangles']),box_halfwidth_um=box_um,spacing_um=2*box_um/(n-1),
              neff=s['n_eff'],neff_abs_error=abs(s['n_eff']-ref['n_eff']),Aeff_um2=s['Aeff_um2'],
              Aeff_error_pct=100*(s['Aeff_um2']/m['Aeff_um2']-1),MFD_2nd_um=s['MFD_2nd_um'],
              MFD_error_pct=100*(s['MFD_2nd_um']/m['MFD_2nd_um']-1),residual=s['residual']))
    return radial,pd.DataFrame(rows),last


def domain_study(f,radii=(30.,40.,50.),core_elements=320,clad_spacing_um=.04):
    positive(clad_spacing_um,'clad_spacing_um');rows=[]
    for r in radii:
        if r<=2*f.radius_um:raise ValueError('domain too small for this assumed core')
        count=max(4,int(np.ceil((r-f.radius_um)/clad_spacing_um)))
        s=solve_radial(f,core_elements=core_elements,clad_elements=count,outer_radius_um=r)
        rows.append(dict(outer_radius_um=r,core_elements=core_elements,clad_elements=count,actual_clad_spacing_um=(r-f.radius_um)/count,
           neff=s['n_eff'],Aeff_um2=s['Aeff_um2'],MFD_2nd_um=s['MFD_2nd_um'],residual=s['residual']))
    return pd.DataFrame(rows)
