"""Carena (2012) Fig. 5 검증: 현재 GN 엔진을 실행하며 NLI fitting을 하지 않는다."""
from __future__ import annotations
import hashlib, json, math, sys, time
from pathlib import Path
from functools import lru_cache
import numpy as np
import pandas as pd
from scipy.integrate import quad
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D
from matplotlib.ticker import FixedLocator, ScalarFormatter

HERE=Path(__file__).resolve().parent
ROOT=HERE.parent
sys.path.insert(0,str(ROOT))
import gn_integral_general as gn
import gn_integral_general_modulation as perf
OUT=HERE/'result'
OUT.mkdir(exist_ok=True)
RS=32.; BN=12.5; SPAN=100.; NF=5.; CUT=4; NCH=9

@lru_cache(None)
def paper_like_psd(spacing):
    rs=RS/1000; bw=float(spacing)/1000
    def raw(f):
        return np.sinc(np.asarray(f)/rs)**2*np.exp(-math.log(2)*(2*np.asarray(f)/bw)**8)
    norm=quad(lambda f: float(raw(f)), -bw,bw,epsabs=1e-13,epsrel=1e-10)[0]
    return lambda f: raw(f)/norm, bw

def system_for_spacing(spacing):
    shape,support=paper_like_psd(float(spacing))
    return gn.WDMSystem(tuple(gn.Channel(float((k-4)*spacing/1000),1e-3,RS,
        pulse_shape='custom',custom_psd=shape,custom_support_half_width_THz=support)
        for k in range(NCH)))

def reach(eta,ase,osnr):
    required=10**(float(osnr)/10)*BN/RS
    p=(ase/(2*eta))**(1/3)
    continuous=p/(ase+eta*p**3)/required
    n=max(0,math.floor(continuous+1e-12))
    # NLI와 ASE를 비코히어런트 방식으로 누적한다.
    return n*SPAN,gn.w_to_dbm(p),continuous*SPAN

def plots(df):
    mods=['BPSK','QPSK','8QAM','16QAM']
    colors={'BPSK':'#b322b5','QPSK':'#df3232','8QAM':'#292929','16QAM':'#2758da'}
    ticks=[100,200,500,1000,2000,5000,10000,20000]
    def draw(ax,fiber):
        for mod in mods:
            d=df[(df.fiber==fiber)&(df.modulation==mod)].sort_values('net_SE_bit_s_Hz')
            ax.plot(d.net_SE_bit_s_Hz,d.code_Lmax_km,color=colors[mod],lw=1.7,marker='x',ms=4)
            ax.plot(d.net_SE_bit_s_Hz,d.paper_Lmax_km,color=colors[mod],lw=1.2,ls='--',marker='o',ms=4,mfc='white')
        ax.set(yscale='log',ylim=(90,24000),xlim=(.9,6.1),title=fiber,
            xlabel='Net spectral efficiency [bit/s/Hz]',ylabel='Maximum reach [km]')
        ax.yaxis.set_major_locator(FixedLocator(ticks)); ax.yaxis.set_major_formatter(ScalarFormatter())
        ax.tick_params(which='minor',labelleft=False)
        ax.grid(True,which='major',alpha=.25)
        handles=[Line2D([],[],color='black',marker='x',label='code'),
            Line2D([],[],color='black',ls='--',marker='o',mfc='white',label='paper')]
        ax.legend(handles=handles,loc='upper right',fontsize=9,framealpha=.9)
    fig,axs=plt.subplots(1,3,figsize=(15,5.3),sharey=True)
    for ax,fiber in zip(axs,['PSCF','SMF','NZDSF']): draw(ax,fiber)
    fig.legend(handles=[Line2D([],[],color=colors[m],lw=2,label='PM-'+m) for m in mods],
        loc='lower center',ncol=4,frameon=False)
    fig.suptitle('Carena 2012 Fig. 5: current GN code vs paper simulation markers',fontsize=13)
    fig.tight_layout(rect=(0,.07,1,.94))
    fig.savefig(OUT/'figure5_code_vs_paper.png',dpi=180)
    fig.savefig(OUT/'figure5_code_vs_paper.svg')
    plt.close(fig)
    for fiber in ['PSCF','SMF','NZDSF']:
        fig,ax=plt.subplots(figsize=(7.2,5.5));draw(ax,fiber)
        fig.legend(handles=[Line2D([],[],color=colors[m],lw=2,label='PM-'+m) for m in mods],loc='lower center',ncol=4,fontsize=9,frameon=False)
        fig.tight_layout(rect=(0,.07,1,1));fig.savefig(OUT/f'figure5_{fiber}.png',dpi=180);plt.close(fig)

def main():
    start=time.perf_counter();ref=pd.read_csv(HERE/'paper_fig5_digitized.csv')
    assert len(ref)==63 and not ref.duplicated(['fiber','modulation','spacing_GHz']).any()
    assert np.allclose(ref.net_SE_bit_s_Hz,ref.modulation.map({'BPSK':2,'QPSK':4,'8QAM':6,'16QAM':8})*25/ref.spacing_GHz,atol=6e-5)
    cache={};conv=[]
    # 같은 조건의 두 해상도/시드로 독립 재계산한다.
    for (fiber,alpha,d,gamma,spacing),_ in ref.groupby(['fiber','alpha_dB_km','D_ps_nm_km','gamma_W_inv_km','spacing_GHz']):
        span=gn.Span(SPAN,alpha,gamma,D_ps_nm_km=d,noise_figure_db=NF)
        sys0=system_for_spacing(spacing)
        vals=[]
        for power,seed,rp in [(17,1,7),(18,2,11)]:
            opt=gn.GNIntegralOptions(sobol_power=power,seed=seed,accumulation='incoherent')
            val=gn.integrate_nli_over_channel(sys0,[span],CUT,opt,receiver_points=rp)
            vals.append(val/(1e-3**3))
        ase=perf.ase_noise_power_edfa([span],RS*1e9,1550.)
        cache[(fiber,float(spacing))]=(vals[1],ase,vals[0])
        conv.append(dict(fiber=fiber,spacing_GHz=spacing,eta_17_seed1_rx7=vals[0],eta_18_seed2_rx11=vals[1],relative_change_pct=abs(vals[1]/vals[0]-1)*100))
        print(f'{fiber} {spacing:g} GHz: eta={vals[1]:.6g}, change={conv[-1]["relative_change_pct"]:.3f}%',flush=True)
    rows=[]
    for _,r in ref.iterrows():
        eta,ase,eta_lo=cache[(r.fiber,float(r.spacing_GHz))]
        L,p,lc=reach(eta,ase,r.paper_required_OSNR_dB_0p1nm)
        Ll,_,_=reach(eta_lo,ase,r.paper_required_OSNR_dB_0p1nm)
        row=r.to_dict();row.update(code_eta_W_inv2=eta,ase_per_span_W=ase,code_Lmax_km=L,
            code_continuous_reach_km=lc,code_popt_dBm=p,code_low_resolution_reach_km=Ll,
            absolute_error_pct=abs(L/r.paper_Lmax_km-1)*100,
            signed_error_pct=(L/r.paper_Lmax_km-1)*100,
            required_snr_db=r.paper_required_OSNR_dB_0p1nm+10*math.log10(BN/RS))
        rows.append(row)
    df=pd.DataFrame(rows);df.to_csv(OUT/'figure5_comparison.csv',index=False)
    pd.DataFrame(conv).to_csv(OUT/'figure5_convergence.csv',index=False)
    byfiber=df.groupby('fiber').absolute_error_pct.agg(['count','mean','max'])
    bymod=df.groupby('modulation').absolute_error_pct.agg(['count','mean','max'])
    byfiber.to_csv(OUT/'figure5_error_by_fiber.csv');bymod.to_csv(OUT/'figure5_error_by_modulation.csv')
    summary=dict(points=len(df),mape_pct=float(df.absolute_error_pct.mean()),median_ape_pct=float(df.absolute_error_pct.median()),
        max_ape_pct=float(df.absolute_error_pct.max()),long_reach_mape_pct=float(df.loc[df.paper_Lmax_km>=1000,'absolute_error_pct'].mean()),
        by_fiber=json.loads(byfiber.to_json(orient='index')),by_modulation=json.loads(bymod.to_json(orient='index')),
        max_eta_resolution_change_pct=max(c['relative_change_pct'] for c in conv),
        max_reach_resolution_change_km=float(abs(df.code_Lmax_km-df.code_low_resolution_reach_km).max()),
        settings=dict(sobol_power=18,seed=2,receiver_points=11,accumulation='incoherent',channels=9,baud_GBd=32,span_km=100,NF_dB=5,trx_snr_db=None),
        source_sha256={f:hashlib.sha256((ROOT/f).read_bytes()).hexdigest() for f in ['gn_integral_general.py','gn_integral_general_modulation.py']},
        elapsed_seconds=time.perf_counter()-start,
        reference='Figure 5 simulation markers digitized from the user-supplied publisher PDF; not author raw data.',
        limitations=['Approximate NRZ sinc-squared PSD and 4th-order SG filter; rectangular receiver integration.',
            'Figure 3 back-to-back OSNR is external calibration for linear XI/ISI, not a GN NLI fit.',
            'Reported 4-8% digitization estimates are inherited approximate estimates, not statistical confidence intervals.',
            'Two-setting sensitivity checks change seed, Sobol samples and receiver points together; not an absolute error bound.'])
    (OUT/'figure5_summary.json').write_text(json.dumps(summary,indent=2,ensure_ascii=False)+'\n')
    plots(df)
    print(json.dumps(summary,indent=2,ensure_ascii=False));return df,summary

if __name__=='__main__': main()
