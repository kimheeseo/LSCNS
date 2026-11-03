"""Adapters from the workbench to the existing independently verified FEM solvers.
All lengths at this API are micrometers except wavelength_nm and length_m.
No HTTP/network access, proprietary COMSOL API or simulated UI-only solver.
"""
from pathlib import Path
import sys,io,base64,time
from dataclasses import replace
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.collections import LineCollection
ROOT=Path(__file__).resolve().parents[1]
sys.path[:0]=[str(ROOT/'waveoptics_m7'/'src'),str(ROOT/'waveoptics_m8'/'src')]
from fiber_modes import Fiber,lp01_analytic,radial_lp01_aeff,C0
from radial_fem import solve_radial,radial_dispersion
from capillary_fem import solve_capillary
from multipole_reference import solve_reference
from scipy.special import j0,jn_zeros
from scipy.integrate import quad

DEFAULTS=dict(model='single',wavelength_nm=1550.,radius_um=4.1,delta_n=.005,
              outer_um=40.,pitch_um=16.,core_elements=160,clad_elements=640,
              mesh_size_um=.7,hcf_radius_um=15.,n_air=1.00027,hcf_elements=320,
              derivative_step_nm=5.,sweep_start_nm=1530.,sweep_end_nm=1625.,sweep_points=7)
SCOPES=dict(single='원형 스텝형 · 약도파 스칼라 LP01 · 외벽 Dirichlet',
            mcf='동일 2코어 · 약도파 스칼라 · 곡선 P2 · 1/4 영역 대칭',
            hcf='이상적 J0/Dirichlet 모세관 · 일정 기체 굴절률 · 방사형 P1')
PRESETS=[dict(id='single',name='단일 코어 · 검증 기본값',values=DEFAULTS),
 dict(id='g654',name='G.654.E 등가 예시 · 실제 단면 아님',values={**DEFAULTS,'radius_um':5.5,'delta_n':.0034785793846155}),
 dict(id='g657',name='G.657.A1 등가 예시 · 실제 단면 아님',values={**DEFAULTS,'wavelength_nm':1310.,'delta_n':.004621529545326,'sweep_start_nm':1290.,'sweep_end_nm':1550.}),
 dict(id='scuba125',name='SCUBA 125 · 단순 등가/예비 모델',values={**DEFAULTS,'radius_um':5.5,'delta_n':.0036378189433495},
      note='반경 5.5 µm 가정, 공칭 MFD 12.3 µm로 Δn 보정. Aeff 공칭 125 µm²와 약 5.15% 차이. 제조사 단면·손실 재현 아님.'),
 dict(id='mcf',name='2코어 MCF · 곡선 P2',values={**DEFAULTS,'model':'mcf','outer_um':32.}),
 dict(id='hcf',name='HCF · 이상적 모세관 기준',values={**DEFAULTS,'model':'hcf'})]


def normalize(raw):
    if not isinstance(raw,dict):raise ValueError('입력은 JSON 객체여야 합니다.')
    unknown=set(raw)-set(DEFAULTS)
    if unknown:raise ValueError('알 수 없는 입력: '+', '.join(sorted(unknown)))
    c={**DEFAULTS,**raw}
    if c['model'] not in SCOPES:raise ValueError('지원 모델: single, mcf, hcf')
    limits={'wavelength_nm':(800,2200),'radius_um':(1,10),'delta_n':(.0001,.02),'outer_um':(15,60),
            'pitch_um':(3,30),'core_elements':(40,640),'clad_elements':(80,2560),'mesh_size_um':(.3,1.5),
            'hcf_radius_um':(3,100),'n_air':(1,1.1),'hcf_elements':(40,1280),'derivative_step_nm':(.5,10),
            'sweep_start_nm':(800,2200),'sweep_end_nm':(800,2200),'sweep_points':(3,11)}
    integers={'core_elements','clad_elements','hcf_elements','sweep_points'}
    for key,(lo,hi) in limits.items():
        value=c[key]
        if isinstance(value,bool) or not isinstance(value,(int,float)) or not np.isfinite(value) or not lo<=value<=hi:
            raise ValueError(f'{key}: {lo}~{hi} 범위의 유한한 숫자가 필요합니다.')
        if key in integers and int(value)!=value:raise ValueError(key+': 정수를 입력하십시오.')
        c[key]=int(value) if key in integers else float(value)
    if c['sweep_start_nm']>=c['sweep_end_nm']:raise ValueError('파장 스윕의 시작값은 끝값보다 작아야 합니다.')
    if c['model']!='hcf':
        f=make_fiber(c);lp01_analytic(f)  # supported isolated-core fundamental benchmark only
        if c['outer_um']<=2*c['radius_um']:raise ValueError('해석영역 반경/반폭은 코어 반경의 두 배보다 커야 합니다.')
        if c['model']=='mcf':
            if c['pitch_um']<=2*c['radius_um']:raise ValueError('중심 간격은 코어 지름보다 커야 합니다.')
            if c['outer_um']<=c['pitch_um']/2+2*c['radius_um']:raise ValueError('코어와 외곽 경계 사이의 여유가 부족합니다.')
            if c['outer_um']>45:raise ValueError('MCF UI의 영역 반폭은 최대 45 µm입니다.')
            if c['mesh_size_um']>c['radius_um']:raise ValueError('MCF 메시 크기는 코어 반경 이하여야 합니다.')
            if c['outer_um']/c['mesh_size_um']>150:raise ValueError('MCF UI의 L/h 제한(150)을 초과했습니다. 메시/영역을 조정하십시오.')
    return c


def make_fiber(c):return Fiber(wavelength_um=c['wavelength_nm']/1000,radius_um=c['radius_um'],delta_n=c['delta_n'],domain_halfwidth_um=c['outer_um'])
def pct(v,r):return 100*abs(v-r)/abs(r)
def metric(key,label,value,unit='',reference=None,error_kind='relative_pct'):
    d=dict(key=key,label=label,value=float(value),unit=unit)
    if reference is not None:
        d.update(reference=float(reference),error=float(abs(value-reference) if error_kind=='absolute' else pct(value,reference)),error_kind=error_kind)
    return d


def figure(fig,caption):
    b=io.BytesIO();fig.savefig(b,format='png',dpi=130,bbox_inches='tight',facecolor='white');plt.close(fig)
    return dict(image='data:image/png;base64,'+base64.b64encode(b.getvalue()).decode(),caption=caption)


def radial_plots(r,q,model,radius):
    R=min(float(r[-1]),radius*3 if model=='single' else radius)
    t=np.linspace(-R,R,181);X,Y=np.meshgrid(t,t);amp=np.interp(np.hypot(X,Y),r,q,right=0.)
    fig,ax=plt.subplots(figsize=(7.5,4.7),constrained_layout=True)
    im=ax.pcolormesh(X,Y,amp**2/np.max(q**2),shading='auto',cmap='inferno',vmin=0,vmax=1)
    ax.add_patch(plt.Circle((0,0),radius,fill=False,color='#77d9ee',lw=.9))
    ax.set(xlabel='x (um)',ylabel='y (um)',title='FEM intensity |psi|^2',aspect='equal');fig.colorbar(im,ax=ax,label='Relative scalar intensity')
    mode=figure(fig,'방사형 FEM 장의 회전 대칭 재구성입니다. 벡터 전기장 성분이 아닙니다.')
    fig,ax=plt.subplots(figsize=(7.5,4.5),constrained_layout=True)
    ax.plot(r,q/np.max(abs(q)),label='FEM scalar amplitude');ax.axvline(radius,color='#75859b',ls='--',label='Core boundary')
    ax.set(xlabel='Radius (um)',ylabel='Normalized amplitude',title='Numerical radial field');ax.grid(alpha=.2);ax.legend()
    return dict(mode=mode,profile=figure(fig,'고유값 문제에서 계산한 실제 FEM 장입니다.'))


def radial_mesh(r,radius):
    fig,ax=plt.subplots(figsize=(8,3.5),constrained_layout=True)
    ax.vlines(r,0,1,color='#446b9b',lw=.5);ax.axvline(radius,color='#cf6534',lw=2,label='Material/wall boundary')
    ax.set(xlabel='Radius (um)',yticks=[],title=f'Actual radial P1 mesh: {len(r)-1} elements');ax.legend()
    return figure(fig,'실제 1차원 방사형 요소입니다. 화면에 임의의 2D 삼각형 메시를 대신 표시하지 않습니다.')


def mcf_mesh_plot(mesh):
    from skfem import Basis,ElementTriP2
    b=Basis(mesh,ElementTriP2());fac=mesh.facets;mid=b.dofs.facet_dofs[0]
    ids=np.stack([fac[0],mid,fac[1]],1);lines=b.doflocs[:,ids].transpose(1,2,0)
    fig,ax=plt.subplots(figsize=(7.5,5),constrained_layout=True)
    ax.add_collection(LineCollection(lines,colors='#547394',linewidths=.22));ax.autoscale();ax.set(aspect='equal',xlabel='x (um)',ylabel='y (um)',title=f'Actual curved P2 mesh: quarter domain, {mesh.nelements} triangles')
    return figure(fig,'실제 계산 영역은 1/4 단면입니다. 원호의 P2 중간 노드를 포함해 표시합니다.')


def mcf_field_plot(s,c):
    ed=s['basis'].element_dofs
    tri=np.concatenate([ed[[0,3,5]].T,ed[[3,1,4]].T,ed[[5,4,2]].T,ed[[3,4,5]].T])
    xy=s['basis'].doflocs;center=xy[:,tri].mean(axis=2)
    tri=tri[(center[0]<c['pitch_um']/2+12)&(center[1]<13)]
    fig,axes=plt.subplots(1,2,figsize=(10,4),constrained_layout=True)
    for i,p in enumerate([1,-1]):
        q=s['fields'][i]/np.max(abs(s['fields'][i]))
        for sx,sy in [(1,1),(-1,1),(1,-1),(-1,-1)]:
            im=axes[i].tripcolor(sx*xy[0],sy*xy[1],tri,q*(p if sx<0 else 1),shading='gouraud',cmap='RdBu_r',vmin=-1,vmax=1,rasterized=True)
        for x in [-c['pitch_um']/2,c['pitch_um']/2]:axes[i].add_patch(plt.Circle((x,0),c['radius_um'],fill=False,color='#333',lw=.7))
        axes[i].set(aspect='equal',xlabel='x (um)',ylabel='y (um)',title=['Even supermode','Odd supermode'][i],xlim=(-c['pitch_um']/2-12,c['pitch_um']/2+12),ylim=(-13,13));fig.colorbar(im,ax=axes[i],shrink=.7)
    return figure(fig,'곡선 P2 FEM 슈퍼모드입니다. 실제 P2 노드값을 표시하며, 대칭 조건으로 전체 단면을 복원했습니다.')


def cap_ref(c):
    a=c['hcf_radius_um'];wl=c['wavelength_nm']/1000;n=c['n_air'];u=jn_zeros(0,1)[0];q=(u/(2*np.pi*a))**2
    ne=np.sqrt(n*n-q*wl*wl);D=wl/C0*1e12*q*n*n/ne**3
    p=quad(lambda t:j0(u*t)**2*t,0,1)[0];p4=quad(lambda t:j0(u*t)**4*t,0,1)[0]
    return dict(n_eff=float(ne),D=float(D),Aeff=float(2*np.pi*a*a*p*p/p4))


def compute_one(c,plots=True,dispersion=True):
    model=c['model'];extra={};outplots={}
    if model=='single':
        f=make_fiber(c);s=solve_radial(f,c['core_elements'],c['clad_elements']);a=lp01_analytic(f);m=radial_lp01_aeff(f)
        rows=[metric('neff','유효굴절률',s['n_eff'],reference=a['n_eff'],error_kind='absolute'),metric('aeff','유효면적',s['Aeff_um2'],'µm²',m['Aeff_um2']),metric('mfd','2차 모멘트 MFD',s['MFD_2nd_um'],'µm',m['MFD_2nd_um'])]
        r,q=s['radius_um'],s['field'];deriv=np.diff(q)/np.diff(r)
        denom=2*np.pi*np.sum(deriv**2*(r[1:]**2-r[:-1]**2)/2)
        rows.append(metric('mfd_pii','Petermann-II MFD',2*np.sqrt(2/denom),'µm'))
        if dispersion:
            d=radial_dispersion(f,c['derivative_step_nm']/1000,c['core_elements'],c['clad_elements'])['D_ps_nm_km']
            rows.append(metric('D','분산 D',d,'ps/(nm·km)'))
        extra=dict(V=f.v,n_clad=f.clad,n_core=f.core,elements=c['core_elements']+c['clad_elements'],residual=s['residual'],tail_power_fraction=s['tail_power_fraction'])
        if plots:outplots={**radial_plots(r,q,model,c['radius_um']),'mesh':radial_mesh(r,c['radius_um'])}
    elif model=='hcf':
        s=solve_capillary(c['wavelength_nm']/1000,c['hcf_radius_um'],c['n_air'],c['hcf_elements'],return_field=True);a=cap_ref(c)
        rows=[metric('neff','유효굴절률',s['n_eff'],reference=a['n_eff'],error_kind='absolute'),metric('aeff','유효면적',s['Aeff_um2'],'µm²',a['Aeff']),metric('D','분산 D',s['D_ps_nm_km'],'ps/(nm·km)',a['D']),metric('kt2','횡방향 고유값 kt²',s['transverse_kt2_per_um2'],'1/µm²')]
        extra=dict(elements=c['hcf_elements'],residual=s['residual'])
        if plots:outplots={**radial_plots(s['radius'],s['field'],model,c['hcf_radius_um']),'mesh':radial_mesh(s['radius'],c['hcf_radius_um'])}
    else:
        from conforming_mcf import solve_conforming,coherent_transfer
        s=solve_conforming(make_fiber(c),c['pitch_um'],c['outer_um'],c['mesh_size_um'],return_fields=plots);a=solve_reference(make_fiber(c),c['pitch_um'],order=8)
        rows=[metric('neff_even','짝모드 neff',s['n_even'],reference=a['n_even'],error_kind='absolute'),metric('neff_odd','홀모드 neff',s['n_odd'],reference=a['n_odd'],error_kind='absolute'),metric('split','Δneff',s['index_split'],reference=a['index_split']),metric('Lc','결합 길이 Lc',s['coupling_length_mm'],'mm',a['coupling_length_mm']),metric('kappa','결합 계수 κ',s['kappa_per_m'],'1/m')]
        extra=dict(elements=s['quarter_elements'],dofs=s['quarter_dofs'],residual=max(s['residuals']),V=make_fiber(c).v)
        if plots:
            z=np.linspace(0,2*s['coupling_length_mm']*1e-3,301);t=coherent_transfer(z,s['coupling_length_mm'])
            fig,ax=plt.subplots(figsize=(7.5,4.5),constrained_layout=True)
            ax.plot(z*1000,t['P1'],label='Core 1');ax.plot(z*1000,t['P2'],label='Core 2');ax.set(xlabel='Length (mm)',ylabel='Power fraction',title='Coherent transfer, identical two-core model');ax.legend();ax.grid(alpha=.2)
            outplots=dict(mode=mcf_field_plot(s,c),mesh=mcf_mesh_plot(s['mesh']),profile=figure(fig,'동일 무손실 두 모드의 전력 전달입니다. 실제 케이블의 랜덤 누화가 아닙니다.'))
    return dict(metrics=rows,details=extra,plots=outplots)


def run_job(raw,action='solve'):
    start=time.time();c=normalize(raw)
    if action not in ['solve','mesh','convergence','domain','sweep']:raise ValueError('지원하지 않는 해석 작업입니다.')
    model=c['model'];result=dict(inputs=c,action=action,scope=SCOPES[model],reference_scope='동일 이상화 스칼라 모델의 독립 기준. 제품·COMSOL 실측 대비 오차 아님.',warnings=[],plots={},rows=[])
    if model=='hcf':result['warnings'].append('유리벽·반공진·PML·복소 neff·구속 손실·굽힘 손실은 계산하지 않습니다.')
    else:result['warnings'].append('벡터 편광·PML·임의 형상·실제 제품 굴절률 분포·실측 정확도는 검증 범위 밖입니다.')
    if model=='mcf':result['warnings'].append('동일 직선 2코어의 스칼라 결합이며, 제조 편차·굽힘·랜덤 케이블 누화를 포함하지 않습니다.')
    if action=='mesh':
        if model=='mcf':
            from conforming_mcf import build_mesh
            mesh,_=build_mesh(make_fiber(c),c['pitch_um'],c['outer_um'],c['mesh_size_um']);result['plots']['mesh']=mcf_mesh_plot(mesh);result['details']=dict(elements=mesh.nelements)
        else:
            if model=='single':r=np.r_[np.linspace(0,c['radius_um'],c['core_elements']+1),np.linspace(c['radius_um'],c['outer_um'],c['clad_elements']+1)[1:]];a=c['radius_um']
            else:r=np.linspace(0,c['hcf_radius_um'],c['hcf_elements']+1);a=c['hcf_radius_um']
            result['plots']['mesh']=radial_mesh(r,a);result['details']=dict(elements=len(r)-1)
        result['metrics']=[]
    elif action=='solve':result.update(compute_one(c))
    else:
        configs=[];name='';unit=''
        if action=='convergence':
            name='Mesh resolution';unit=''
            if model=='single':
                configs=[{**c,'core_elements':max(10,int(c['core_elements']*t)),'clad_elements':max(20,int(c['clad_elements']*t))} for t in [.25,.5,1.]];x=[z['core_elements']+z['clad_elements'] for z in configs]
            elif model=='hcf':configs=[{**c,'hcf_elements':max(10,int(c['hcf_elements']*t))} for t in [.25,.5,1.]];x=[z['hcf_elements'] for z in configs]
            else:
                configs=[{**c,'mesh_size_um':min(c['radius_um'],c['mesh_size_um']*t)} for t in [1.8,1.35,1.]];x=[z['mesh_size_um'] for z in configs];name='Target mesh h';unit='um'
        elif action=='domain':
            if model=='hcf':raise ValueError('HCF 모세관의 외벽은 실제 모델 경계입니다. 반경 변경을 영역 수렴으로 해석할 수 없습니다. 메시 수렴을 사용하십시오.')
            name='Outer radius/halfwidth';unit='um'
            maximum=45. if model=='mcf' else 80.
            base=c['outer_um'];x=[base,base+2.,base+4.]
            if x[-1]>maximum:raise ValueError('영역 수렴을 위해 입력 영역을 조금 줄여 주십시오.')
            configs=[{**c,'outer_um':v} for v in x];result['warnings'].append('영역별로 재메시하므로 잔여 공간 이산화 오차도 포함됩니다.')
        else:
            name='Wavelength';unit='nm';x=np.linspace(c['sweep_start_nm'],c['sweep_end_nm'],c['sweep_points']).tolist();configs=[{**c,'wavelength_nm':w} for w in x]
            for z in configs:
                if model!='hcf':lp01_analytic(make_fiber(z))
        records=[]
        for xx,z in zip(x,configs):
            d=compute_one(z,plots=False,dispersion=False)
            row=dict(x=float(xx),**{m['key']:m['value'] for m in d['metrics']})
            for m in d['metrics']:
                if 'error' in m:row[m['key']+'_error']=m['error']
            records.append(row)
        keys=['neff','aeff'] if model!='mcf' else ['split','Lc']
        if model=='hcf' and action=='sweep':keys=['neff','D']
        fig,axes=plt.subplots(1,2,figsize=(10,4),constrained_layout=True)
        for ax,key in zip(axes,keys):
            y=[r[key] for r in records];ax.plot(x,y,'o-',color='#2867ad');ax.set(xlabel=f'{name} ({unit})' if unit else name,ylabel=key,title=key);ax.ticklabel_format(axis='y',useOffset=False);ax.grid(alpha=.2)
        result.update(last_condition=configs[-1],rows=records,metrics=d['metrics'],details=dict(points=len(records)),plots={action:figure(fig,'실제 해석기를 각 조건에서 다시 실행한 결과입니다. 보간된 저장값이 아닙니다.')})
        if action=='sweep' and model=='single':result['warnings'].append('단일 코어 파장 스윕은 neff·Aeff·MFD를 계산합니다. D는 개별 모드 해석에서 동일 메시의 5점 미분으로 계산합니다.')
    result['runtime_seconds']=time.time()-start
    # Strict JSON: never return a plot containing invalid numeric data.
    return result
