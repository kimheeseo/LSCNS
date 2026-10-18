"""Unfitted predictions and separately evaluated HCF2 benchmark comparisons.

Prediction functions take geometry/wavelength only. No measured loss is passed
to a solver, a root-selection criterion, or a geometry selection procedure.
The benchmark was known previously; this is not a blind validation.
"""
from pathlib import Path
from itertools import product
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from hcf_concentric_tmm_20260908 import (
    ConcentricLeakyModeModel, U01, build_models, solve_fundamental_mode,
    bache_bouncing_ray_loss_db_km, silica_index_sellmeier)
from hcf_vector_tmm import solve_vector_mode, full_chord_model, high_precision_check

# Fixed before this spectrum evaluation; retained published nominal core and
# midpoint diameters. Ranges are not optimized against attenuation.
GEOMETRY = dict(core_radius_um=14.75, wall_um=0.50,
                large_um=31.05, middle_um=23.75, small_um=7.70)
WAVELENGTHS_NM = np.arange(1200., 1651., 10.)
PRIMARY_MODEL = 'new_vector7'
MODELS = ['old_scalar5', 'old_scalar7', PRIMARY_MODEL, 'fullchord_vector9']
LABELS = {'old_scalar5': 'Original scalar, 5 regions',
          'old_scalar7': 'Earlier scalar, 7 regions',
          'new_vector7': 'Revised vector, 7 regions (primary)',
          'fullchord_vector9': 'Full chord, vector 9 regions (assumption)'}
COLORS = dict(zip(MODELS, ['#cf6b39', '#b69230', '#773bb2', '#148b8b']))


def geometry_models(w, geometry=None):
    g = GEOMETRY if geometry is None else geometry
    gap1 = g['large_um'] - 2*g['wall_um'] - g['middle_um']
    gap2 = g['middle_um'] - 2*g['wall_um'] - g['small_um']
    m = build_models(w, gap1, gap2, core_radius_um=g['core_radius_um'],
                     wall_thickness_um=g['wall_um'])
    return m, full_chord_model(w, **g)


def predict_spectrum(wavelengths=WAVELENGTHS_NM):
    """Execute retained scalar code and modified vector code at every point."""
    rows = []
    previous = {}
    for i, w in enumerate(wavelengths):
        models, chord = geometry_models(w)
        for key, model in [('old_scalar5', models['5-layer']),
                           ('old_scalar7', models['7-layer']),
                           ('new_vector7', models['7-layer']),
                           ('fullchord_vector9', chord)]:
            if key.startswith('old_'):
                s = solve_fundamental_mode(w, model)
            else:
                # Continuation is by eigenvalue proximity, never attenuation.
                s = solve_vector_mode(w, model, seed=previous.get(key),
                                      multistart=True)
                previous[key] = complex(s.u_real, s.u_imag)
            if not s.converged:
                raise RuntimeError(f'Unconverged result: {key}, {w} nm')
            row = s.to_dict()
            row.update(model=key, quantity='confinement/leakage loss only')
            rows.append(row)
        if i % 10 == 0 or i == len(wavelengths)-1:
            print(f'Computed {i+1}/{len(wavelengths)} wavelengths; all four variants retained.')
    return pd.DataFrame(rows)


def bird_benchmark():
    """Bird Table 1 HE11 exact values, rc/lambda=15, epsilon=2.25.

    Eq.(3) fixes all layer widths at antiresonance; no Table 2 optimization.
    Table values are rounded to 0.001; difference below 0.0005 is rounding
    compatible. N counts finite layers, not our total number of regions.
    """
    rows=[]
    for N, ref in [(1,1.157),(2,.788),(3,.556),(4,.409)]:
        ns=tuple(1.5 if j%2==0 else 1. for j in range(N))
        ts=tuple(1/(4*np.sqrt(1.25)) if j%2==0 else np.pi*15/(2*U01)
                 for j in range(N))
        model=ConcentricLeakyModeModel('Bird Table 1',15,ns,ts,
                                       n_outer=1 if N%2 else 1.5)
        s=solve_vector_mode(1000,model)
        scaled=s.loss_db_km/1e9*15**(N+3)
        rows.append(dict(N_finite_layers=N, regions=N+2,
                         paper_scaled_loss=ref, code_scaled_loss=scaled,
                         relative_error_pct=100*(scaled/ref-1),
                         within_table_rounding=abs(scaled-ref)<=.0005,
                         converged=s.converged))
    return pd.DataFrame(rows)


def numerical_checks(predictions):
    rows=[]
    for w in [1310.,1550.]:
        models,chord=geometry_models(w)
        for key,m in [('new_vector7',models['7-layer']),('fullchord_vector9',chord)]:
            row=predictions.query('wavelength_nm == @w and model == @key').iloc[0]
            hp=high_precision_check(w,m,complex(row.u_real,row.u_imag),dps=40)
            rows.append(dict(wavelength_nm=w,model=key,
                             float64_loss_db_km=row.loss_db_km,
                             mp40_loss_db_km=hp['loss_db_km'],
                             difference_pct=100*(row.loss_db_km/hp['loss_db_km']-1),
                             mp40_determinant_abs=hp['raw_determinant_abs']))
    return pd.DataFrame(rows)


def single_wall_diagnostic():
    rows=[]
    for w in [1310.,1550.]:
        models,_=geometry_models(w)
        scalar=solve_fundamental_mode(w,models['3-layer']).loss_db_km
        vector=solve_vector_mode(w,models['3-layer']).loss_db_km
        te=bache_bouncing_ray_loss_db_km(w,'TE')
        hybrid=bache_bouncing_ray_loss_db_km(w,'hybrid')
        rows.append(dict(wavelength_nm=w,scalar3_db_km=scalar,
                         vector3_db_km=vector,bache_TE_db_km=te,
                         bache_hybrid_db_km=hybrid,
                         scalar_vs_TE_pct=100*(scalar/te-1),
                         vector_vs_hybrid_pct=100*(vector/hybrid-1)))
    return pd.DataFrame(rows)


def geometry_sensitivity():
    """All 27 diameter triples, two wavelengths; not a statistical CI.

    These are rectangular combinations of marginal published ranges, not
    measured joint cross sections. No minimum-error combination is selected.
    Wall thickness remains approximate; its uncertainty is not quantified.
    """
    rows=[]
    for d1,d2,d3 in product([30.4,31.05,31.7],[22.7,23.75,24.8],[7.,7.7,8.4]):
        geom=dict(GEOMETRY,large_um=d1,middle_um=d2,small_um=d3)
        for w in [1310.,1550.]:
            models,chord=geometry_models(w,geom)
            # Evaluate the primary model only; full chord nominal is reported
            # as a separate topology assumption, not tuned in this range scan.
            s=solve_vector_mode(w,models['7-layer'])
            rows.append(dict(wavelength_nm=w,**geom,loss_db_km=s.loss_db_km,
                             converged=s.converged))
    return pd.DataFrame(rows)


def compare(predictions,reference,reference_kind):
    # Reference values enter for the first time here, after prediction.
    merged=predictions.merge(reference[['wavelength_nm','paper_loss_db_km']],
                              on='wavelength_nm',how='inner',validate='many_to_one')
    merged['relative_error_pct']=100*(merged.loss_db_km-merged.paper_loss_db_km)/merged.paper_loss_db_km
    merged['absolute_percentage_error']=merged.relative_error_pct.abs()
    merged['exceeds_10pct']=merged.absolute_percentage_error>10
    merged['within_15pct']=merged.absolute_percentage_error<=15
    merged['reference_kind']=reference_kind
    merged['interpretation']='leakage-to-measured-total discrepancy; not validated total-loss accuracy'
    return merged


def metrics(comparison):
    return comparison.groupby('model',sort=False).agg(
        n=('wavelength_nm','size'),
        min_wavelength_nm=('wavelength_nm','min'),
        max_wavelength_nm=('wavelength_nm','max'),
        MAPE_pct=('absolute_percentage_error','mean'),
        max_APE_pct=('absolute_percentage_error','max'),
        n_exceeds_10pct=('exceeds_10pct','sum'),
        n_within_15pct=('within_15pct','sum')).reset_index()


def primary_comparison_table(comparison):
    cols=['wavelength_nm','paper_loss_db_km','loss_db_km',
          'relative_error_pct','absolute_percentage_error','exceeds_10pct']
    tab=comparison[comparison.model==PRIMARY_MODEL][cols].copy()
    tab=tab.rename(columns={'loss_db_km':'revised_vector7_db_km'})
    old=comparison[comparison.model=='old_scalar5'][['wavelength_nm','loss_db_km']]
    tab=tab.merge(old.rename(columns={'loss_db_km':'original_scalar5_db_km'}),on='wavelength_nm')
    return tab[['wavelength_nm','paper_loss_db_km','original_scalar5_db_km',
                'revised_vector7_db_km','relative_error_pct',
                'absolute_percentage_error','exceeds_10pct']]


def make_plots(predictions,raw,comparison,nominal,sensitivity,outdir):
    outdir=Path(outdir);outdir.mkdir(parents=True,exist_ok=True)
    plt.rcParams.update({'font.family':'DejaVu Sans','font.size':11,
                        'axes.spines.top':False,'axes.spines.right':False,
                        'figure.facecolor':'white','savefig.facecolor':'white'})
    paths=[]
    fig,ax=plt.subplots(figsize=(12,6.5),layout='constrained')
    data=raw[raw.wavelength_nm.between(1200,1650)]
    ax.semilogy(data.wavelength_nm,data.paper_loss_db_km,color='#164b88',lw=2.3,label='Petrovich HCF2 measured TOTAL (source data)')
    for key in MODELS:
        d=predictions[predictions.model==key]
        ax.semilogy(d.wavelength_nm,d.loss_db_km,color=COLORS[key],lw=1.8,
                    ls='--' if key!='new_vector7' else '-',label=LABELS[key]+' | leakage')
    ax.set(xlabel='Wavelength (nm)',ylabel='Loss (dB/km)',xlim=(1200,1650),
           title='20260908 | Unfitted code calculations vs HCF2 measurements')
    ax.grid(True,which='both',alpha=.16);ax.legend(fontsize=9,loc='upper right')
    p=outdir/'20260908_loss.png';fig.savefig(p,dpi=180);plt.close(fig);paths.append(p)
    fig,ax=plt.subplots(figsize=(12,6.5),layout='constrained')
    for key in MODELS:
        d=comparison[comparison.model==key]
        ax.semilogy(d.wavelength_nm,d.absolute_percentage_error,color=COLORS[key],
                    label=LABELS[key],lw=1.8,marker='.' if key==PRIMARY_MODEL else None)
    ax.axhline(10,color='#bb2535',lw=1.8,ls='--',label='10% threshold')
    ax.axhspan(0.1,10,color='#478957',alpha=.07)
    ax.set(xlabel='Wavelength (nm)',ylabel='Absolute percentage discrepancy (%)',
           title='Leakage-to-total discrepancy | 46 exact source-data wavelengths',
           xlim=(1200,1650),ylim=(.1,None))
    ax.grid(True,which='both',alpha=.16);ax.legend(fontsize=9)
    p=outdir/'20260908_error.png';fig.savefig(p,dpi=180);plt.close(fig);paths.append(p)
    tab=primary_comparison_table(nominal)
    fig,ax=plt.subplots(figsize=(12,3.5),layout='constrained');ax.axis('off')
    ax.set_title('20260908 | Actual executed results — primary revised vector model',loc='left',pad=22,fontweight='bold')
    cells=[]
    for r in tab.itertuples():
        cells.append([f'{r.wavelength_nm:.0f}',f'{r.paper_loss_db_km:.3f}',
                      f'{r.original_scalar5_db_km:.3f}',f'{r.revised_vector7_db_km:.6f}',
                      f'{r.relative_error_pct:+,.2f}%', 'YES' if r.exceeds_10pct else 'NO'])
    t=ax.table(cellText=cells,colLabels=['nm','Paper total','Original scalar 5','Revised vector 7','Signed error','Above 10%'],
               loc='center',cellLoc='center',colWidths=[.07,.14,.19,.19,.24,.13])
    t.auto_set_font_size(False);t.set_fontsize(11);t.scale(1,2)
    for (r,c),cell in t.get_celld().items():
        cell.set_edgecolor('#dce2ed')
        if r==0:cell.set_facecolor('#edf2f8');cell.set_text_props(weight='bold')
    ax.text(0,.08,'Loss units: dB/km. Calculated values are leakage only. 10–15% target: NOT achieved.',transform=ax.transAxes,color='#a22437')
    p=outdir/'20260908_results.png';fig.savefig(p,dpi=180);plt.close(fig);paths.append(p)
    fig,ax=plt.subplots(figsize=(9,5),layout='constrained')
    for i,w in enumerate([1310.,1550.]):
        d=sensitivity[sensitivity.wavelength_nm==w].loss_db_km.to_numpy()
        ax.scatter(i+np.linspace(-.14,.14,len(d)),d,color=COLORS[PRIMARY_MODEL],alpha=.6,s=28)
        v=nominal[(nominal.model==PRIMARY_MODEL)&(nominal.wavelength_nm==w)].iloc[0]
        ax.scatter(i,v.loss_db_km,marker='D',s=70,color='#101827',zorder=4)
        ax.hlines(v.paper_loss_db_km,i-.24,i+.24,color='#164b88',lw=3)
    ax.set(yscale='log',xticks=[0,1],xticklabels=['1310 nm','1550 nm'],ylabel='Loss (dB/km)',
           title='Published diameter-range combinations | no best-fit selection')
    ax.text(.02,.98,'Purple: all 27 geometry combinations\nBlack diamond: nominal; blue line: measured total',va='top',transform=ax.transAxes,fontsize=10)
    ax.grid(True,axis='y',which='both',alpha=.15)
    p=outdir/'20260908_geometry_sensitivity.png';fig.savefig(p,dpi=180);plt.close(fig);paths.append(p)
    return paths
