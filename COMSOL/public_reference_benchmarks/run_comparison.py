"""Execute existing FEM against published cases; COMSOL itself is not executed."""
from pathlib import Path
import sys,os,json,csv,time,hashlib,platform,subprocess,importlib.util
from datetime import datetime,timezone
import importlib.metadata as md
import numpy as np
from scipy.special import j0,jn_zeros
from scipy.integrate import quad
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
ROOT=Path(__file__).resolve().parent;COMSOL=ROOT.parent;OUT=ROOT/'results'
OUT.mkdir(exist_ok=True)
sys.path[:0]=[str(COMSOL/'waveoptics_m1'),str(COMSOL/'waveoptics_m7/src'),str(COMSOL/'waveoptics_m8/src')]
from waveoptics.slab import solve_slab
from fiber_modes import Fiber
from radial_fem import solve_radial
from conforming_mcf import solve_conforming
from multipole_reference import solve_reference
from capillary_fem import solve_capillary
from reference_equations import scalar_circular_fundamental,vector_he11
spec=importlib.util.spec_from_file_location('independent_slab',COMSOL/'waveoptics_m1/tests/analytic_slab.py')
analytic=importlib.util.module_from_spec(spec);sys.modules[spec.name]=analytic;spec.loader.exec_module(analytic)
CHECKS=[]
def check(name,passed,detail=None):
    CHECKS.append(dict(name=name,passed=bool(passed),detail=detail));print('CHECK '+json.dumps(CHECKS[-1]),flush=True)
def timed(fun,*a,**kw):
    t=time.perf_counter();r=fun(*a,**kw);return r,time.perf_counter()-t
def pct(x,r):return float(100*abs(x-r)/abs(r))
def write_csv(name,rows):
    with (OUT/name).open('w',encoding='utf-8-sig',newline='') as h:
        w=csv.DictWriter(h,fieldnames=list(rows[0]));w.writeheader();w.writerows(rows)
def plot_xy(name,x,series,xlabel,ylabel,title):
    fig,ax=plt.subplots(figsize=(7,4))
    for label,y in series.items():ax.plot(x,y,'o-',label=label)
    ax.set(xlabel=xlabel,ylabel=ylabel,title=title);ax.grid(alpha=.2);ax.legend()
    fig.tight_layout();fig.savefig(OUT/name);plt.close(fig)

def slab():
    p=dict(n_core=1.5,n_clad=1.,width_um=1.,wavelength_um=1.55,padding_um=3.,polarization='TE',max_modes=6)
    refs=analytic.exact_modes(1.5,1.,1.,1.55,'TE');rows=[]
    for h in [.04,.02,.01,.005]:
        s,sec=timed(solve_slab,**p,h_um=h)
        check('M1 mode count h='+str(h),len(s.modes)==len(refs),dict(fem=len(s.modes),analytic=len(refs)))
        for m,r in zip(s.modes,refs):
            rows.append(dict(mode='TE'+str(m.order),h_um=s.h_max_um,elements=s.elements,dofs=s.dofs,neff=m.neff,
                reference_neff=r.neff,abs_error=abs(m.neff-r.neff),relative_error_pct=pct(m.neff,r.neff),
                field_relative_l2_error=analytic.l2_field_error(s,m,r),runtime_seconds=sec))
    domains=[]
    for pad in [3.,5.,7.]:
        d,sec=timed(solve_slab,**{**p,'padding_um':pad},h_um=.005)
        domains.append(dict(padding_um=pad,neff_TE0=d.modes[0].neff,neff_TE1=d.modes[1].neff,runtime_seconds=sec))
    t=[r for r in rows if r['mode']=='TE0']
    check('M1 TE0 finest absolute error <1e-5',t[-1]['abs_error']<1e-5,t[-1])
    check('M1 TE0 field relative L2 <2e-4',t[-1]['field_relative_l2_error']<2e-4)
    check('M1 TE0 mesh error decreases',all(a['abs_error']>b['abs_error'] for a,b in zip(t[:-1],t[1:])))
    check('M1 TE0 padding5->7 change <1e-9',abs(domains[-1]['neff_TE0']-domains[-2]['neff_TE0'])<1e-9)
    fig,ax=plt.subplots(figsize=(7,4))
    for m,r in zip(s.modes,refs):
        x=np.linspace(-3.5,3.5,2001);v=np.interp(x,s.x_um,m.field);ex=r.field(x)
        if np.dot(v,ex)<0:v=-v
        ax.plot(x,v,label='FEM TE'+str(m.order));ax.plot(x,ex,'--',label='Analytic TE'+str(m.order))
    ax.set(xlabel='Transverse coordinate (um)',ylabel='L2-normalized scalar field',title='Published slab parameters; transverse eigenproblem')
    ax.grid(alpha=.2);ax.legend();fig.tight_layout();fig.savefig(OUT/'m1_fields.svg');plt.close(fig)
    write_csv('m1_mesh.csv',rows);write_csv('m1_domain.csv',domains)
    return dict(parameters=p,mesh=rows,domain=domains,reference_kind='published analytical equations; not raw COMSOL output',
        analytic_modes=[dict(mode='TE'+str(r.order),neff=r.neff) for r in refs],port_propagation_comparison=False,
        documentation_discrepancy='With full width 1 um, V_half=2.266>pi/2, both TE0/TE1 exist. Official prose says TE0 only; do not use it as mode count oracle.')

def circular():
    f=Fiber(wavelength_um=1.55,radius_um=8.,delta_n=1.4457-1.4378,n_cladding=1.4378,domain_halfwidth_um=40.)
    sc=scalar_circular_fundamental(f.core,f.clad,8.,1.55);ve=vector_he11(f.core,f.clad,8.,1.55);rows=[]
    for nc,ns in [(80,320),(160,640),(320,1280),(640,2560)]:
        s,sec=timed(solve_radial,f,nc,ns)
        rows.append(dict(core_elements=nc,clad_elements=ns,neff=s['n_eff'],scalar_reference=sc['neff'],
            scalar_abs_error=abs(s['n_eff']-sc['neff']),scalar_relative_error_pct=pct(s['n_eff'],sc['neff']),
            Aeff_um2=s['Aeff_um2'],MFD_2nd_um=s['MFD_2nd_um'],residual=s['residual'],runtime_seconds=sec))
    domain=[]
    for R in [32.,40.,48.]:
        ff=Fiber(wavelength_um=1.55,radius_um=8.,delta_n=f.delta_n,n_cladding=f.clad,domain_halfwidth_um=R)
        ns=int(round((R-8)/.0125));d,sec=timed(solve_radial,ff,640,ns)
        domain.append(dict(outer_radius_um=R,clad_elements=ns,neff=d['n_eff'],runtime_seconds=sec))
    weak=vector_he11(1.4501,1.45,20.,1.55);weaksc=scalar_circular_fundamental(1.4501,1.45,20.,1.55)
    check('M7 weak contrast vector/scalar limit <1e-7',abs(weak['neff']-weaksc['neff'])<1e-7,dict(vector=weak,scalar=weaksc))
    check('M7 finest scalar absolute error <1e-8',rows[-1]['scalar_abs_error']<1e-8,rows[-1])
    check('M7 mesh error decreases',all(a['scalar_abs_error']>b['scalar_abs_error'] for a,b in zip(rows[:-1],rows[1:])))
    check('M7 vector matches printed rounding interval (assumed nearest)',abs(ve['neff']-1.4444)<=.00005,ve)
    check('M7 vector characteristic residual <1e-10',ve['characteristic_relative_residual']<1e-10)
    check('M7 R40->48 neff change <1e-9',abs(domain[-1]['neff']-domain[-2]['neff'])<1e-9)
    plot_xy('m7_convergence.svg',[r['core_elements'] for r in rows],{'Scalar absolute neff error':[r['scalar_abs_error'] for r in rows]},'Core elements (cladding=4x)','Absolute neff error','Published circular fiber; existing radial P1 FEM')
    write_csv('m7_mesh.csv',rows);write_csv('m7_domain.csv',domain)
    return dict(parameters=dict(wavelength_um=1.55,n_core=f.core,n_clad=f.clad,radius_um=8,outer_radius_um=40,V=f.v),
        scalar_reference=sc,vector_characteristic_reference=ve,published_comsol_neff_rounded=1.4444,
        published_comsol_precision='four decimal places; no precise raw COMSOL export retrieved',
        finest_fem_to_printed_neff_abs_difference=abs(rows[-1]['neff']-1.4444),
        scalar_vector_abs_model_difference=abs(sc['neff']-ve['neff']),
        scalar_vector_relative_model_difference_pct=pct(sc['neff'],ve['neff']),mesh=rows,domain=domain,production_ui_range_changed=False)

def mcf():
    nc=1.444/np.sqrt(1-2*.0025)
    def fiber(ncore=nc):return Fiber(wavelength_um=1.55,radius_um=3.5,delta_n=ncore-1.444,n_cladding=1.444,domain_halfwidth_um=40.)
    f=fiber();pitch=9.;refs=[]
    for order in [6,8,10]:
        r,sec=timed(solve_reference,f,pitch,order=order,samples=512)
        r.update(kappa_per_m=np.pi*r['index_split']/1.55e-6,runtime_seconds=sec);refs.append(r)
    ref=refs[-1];mesh=[]
    for h in [1.,.7,.5,.35]:
        s,sec=timed(solve_conforming,f,pitch,40.,h)
        mesh.append(dict(h_um=h,elements=s['quarter_elements'],dofs=s['quarter_dofs'],n_even=s['n_even'],n_odd=s['n_odd'],
            split=s['index_split'],kappa_per_m=s['kappa_per_m'],Lc_mm=s['coupling_length_mm'],
            kappa_reference_per_m=ref['kappa_per_m'],kappa_error_pct=pct(s['kappa_per_m'],ref['kappa_per_m']),
            Lc_error_pct=pct(s['coupling_length_mm'],ref['coupling_length_mm']),runtime_seconds=sec))
    domains=[]
    for L in [32.,36.,40.]:
        s,sec=timed(solve_conforming,f,pitch,L,.5)
        domains.append(dict(halfwidth_um=L,kappa_per_m=s['kappa_per_m'],Lc_mm=s['coupling_length_mm'],runtime_seconds=sec))
    extended=[]
    for L in [48.,56.,64.]:
        z,sec=timed(solve_conforming,f,pitch,L,.35)
        extended.append(dict(halfwidth_um=L,h_um=.35,elements=z['quarter_elements'],dofs=z['quarter_dofs'],
            n_even=z['n_even'],n_odd=z['n_odd'],kappa_per_m=z['kappa_per_m'],Lc_mm=z['coupling_length_mm'],
            kappa_error_pct=pct(z['kappa_per_m'],ref['kappa_per_m']),
            Lc_error_pct=pct(z['coupling_length_mm'],ref['coupling_length_mm']),runtime_seconds=sec))
    check('MCF expanded L56->64 kappa change <0.001%',pct(extended[-2]['kappa_per_m'],extended[-1]['kappa_per_m'])<.001,extended)
    check('MCF expanded L64 kappa error <0.001%',extended[-1]['kappa_error_pct']<.001,extended[-1])
    gaps=[]
    for gap in [1.,2.,4.,6.]:
        rr=ref if gap==2 else solve_reference(f,7+gap,order=10,samples=512)
        s,sec=timed(solve_conforming,f,7+gap,40.,.5);kr=np.pi*rr['index_split']/1.55e-6
        gaps.append(dict(gap_um=gap,pitch_um=7+gap,kappa_per_m=s['kappa_per_m'],reference_kappa_per_m=kr,
            Lc_mm=s['coupling_length_mm'],kappa_error_pct=pct(s['kappa_per_m'],kr),runtime_seconds=sec))
    alternatives=[]
    for label,ncore in [('squared_index_primary',nc),('linear_relative_to_cladding',1.444*(1+.0025)),('linear_relative_to_core',1.444/(1-.0025))]:
        rr=solve_reference(fiber(ncore),pitch,order=10,samples=512);kr=np.pi*rr['index_split']/1.55e-6
        alternatives.append(dict(convention=label,n_core=ncore,delta_n=ncore-1.444,kappa_per_m=kr,Lc_mm=rr['coupling_length_mm'],
            kappa_difference_from_primary_pct=pct(kr,ref['kappa_per_m'])))
    check('MCF angular order8->10 kappa change <0.01%',pct(refs[-2]['kappa_per_m'],ref['kappa_per_m'])<.01)
    check('MCF finest scalar kappa error <0.1%',mesh[-1]['kappa_error_pct']<.1,mesh[-1])
    check('MCF L36->40 kappa change <0.1%',pct(domains[-2]['kappa_per_m'],domains[-1]['kappa_per_m'])<.1,domains)
    check('MCF published trend kappa decreases with edge gap',all(a['kappa_per_m']>b['kappa_per_m'] for a,b in zip(gaps[:-1],gaps[1:])),gaps)
    plot_xy('mcf_gap.svg',[r['gap_um'] for r in gaps],{'Existing curved P2 FEM':[r['kappa_per_m'] for r in gaps],'Independent scalar multipole':[r['reference_kappa_per_m'] for r in gaps]},'Edge gap dg (um)','Kappa (1/m)','Zhao et al. parameters; scalar approximation')
    for name,rows in [('mcf_mesh.csv',mesh),('mcf_domain.csv',domains),('mcf_gap.csv',gaps),('mcf_delta_sensitivity.csv',alternatives),('mcf_extended_domain.csv',extended)]:write_csv(name,rows)
    return dict(parameters=dict(n_clad=f.clad,n_core=f.core,delta_n=f.delta_n,relative_delta=.0025,radius_um=3.5,
        gap_um=2,pitch_um=9,wavelength_um=1.55,V=f.v),independent_reference=refs,mesh=mesh,domain=domains,gap_sweep=gaps,
        delta_sensitivity=alternatives,extended_domain=extended,published_comsol_comparison='parameters/trend only; graph not digitized, no raw beta/kappa',
        numerical_comsol_error_pct=None,polarization_comparison=False)

def hcf():
    wl=1.55;R=15.;ng=1.00027;u=jn_zeros(0,1)[0]
    def exact(radius):
        kt2=(u/radius)**2;q=kt2/(2*np.pi)**2;ne=np.sqrt(ng*ng-q*wl*wl)
        p=quad(lambda x:j0(u*x)**2*x,0,1,epsabs=1e-13)[0];p4=quad(lambda x:j0(u*x)**4*x,0,1,epsabs=1e-13)[0]
        return dict(neff=float(ne),kt2=float(kt2),Aeff_um2=float(2*np.pi*radius**2*p*p/p4),
            D_ps_nm_km=float(wl/299792458.*1e12*q*ng*ng/ne**3))
    ref=exact(R);rows=[]
    for N in [80,160,320,640]:
        s,sec=timed(solve_capillary,wl,R,ng,N)
        rows.append(dict(elements=N,neff=s['n_eff'],neff_abs_error=abs(s['n_eff']-ref['neff']),
            Aeff_um2=s['Aeff_um2'],Aeff_error_pct=pct(s['Aeff_um2'],ref['Aeff_um2']),
            D_ps_nm_km=s['D_ps_nm_km'],D_error_pct=pct(s['D_ps_nm_km'],ref['D_ps_nm_km']),
            kt2_error_pct=pct(s['transverse_kt2_per_um2'],ref['kt2']),runtime_seconds=sec))
    radius=[]
    for r in [10.,15.,20.,25.]:
        s=solve_capillary(wl,r,ng,640);rr=exact(r)
        radius.append(dict(radius_um=r,neff=s['n_eff'],reference_neff=rr['neff'],Aeff_um2=s['Aeff_um2'],D_ps_nm_km=s['D_ps_nm_km']))
    check('HCF N640 neff absolute error <1e-9',rows[-1]['neff_abs_error']<1e-9,rows[-1])
    check('HCF N640 Aeff error <0.0001%',rows[-1]['Aeff_error_pct']<.0001)
    check('HCF N640 D error <0.0001%',rows[-1]['D_error_pct']<.0001)
    check('HCF mesh neff error decreases',all(a['neff_abs_error']>b['neff_abs_error'] for a,b in zip(rows[:-1],rows[1:])))
    plot_xy('hcf_radius.svg',[r['radius_um'] for r in radius],{'D (ps/(nm km))':[r['D_ps_nm_km'] for r in radius]},'Physical radius (um)','D (ps/(nm km))','Ideal capillary; no glass-wall leakage')
    write_csv('hcf_mesh.csv',rows);write_csv('hcf_radius.csv',radius)
    return dict(parameters=dict(wavelength_um=wl,radius_um=R,n_gas_constant=ng),parameter_origin='chosen baseline; not the full Kelly et al. fiber geometry',
        reference=ref,reference_kind='published capillary approximation Eq.(1), exact for current ideal Dirichlet disk',
        mesh=rows,radius_sweep=radius,numerical_comsol_error_pct=None,confinement_loss_comparison=False,
        rejected_comparison='Official microstructured fiber is silica core with vector modes/PML, not the same HCF geometry.')

def table(headers,rows):
    return '\n'.join(['| '+' | '.join(headers)+' |','| '+' | '.join(['---']*len(headers))+' |']+['| '+' | '.join(str(x) for x in r)+' |' for r in rows])

def main():
    data=dict(M1=slab(),M7=circular(),MCF=mcf(),HCF=hcf())
    try:commit=subprocess.check_output(['git','rev-parse','HEAD'],cwd=COMSOL,text=True).strip()
    except Exception:commit='unknown'
    paths=[Path(__file__),ROOT/'reference_equations.py',ROOT/'references.json',
        COMSOL/'waveoptics_m1/waveoptics/slab.py',COMSOL/'waveoptics_m1/waveoptics/mesh.py',
        COMSOL/'waveoptics_m1/tests/analytic_slab.py',COMSOL/'waveoptics_m7/src/fiber_modes.py',
        COMSOL/'waveoptics_m7/src/radial_fem.py',COMSOL/'waveoptics_m8/src/conforming_mcf.py',
        COMSOL/'waveoptics_m8/src/multipole_reference.py',COMSOL/'waveoptics_m8/src/capillary_fem.py']
    data['metadata']=dict(executed_utc=datetime.now(timezone.utc).isoformat(),source_commit=commit,
        run_url=('https://github.com/'+os.environ['GITHUB_REPOSITORY']+'/actions/runs/'+os.environ['GITHUB_RUN_ID']) if 'GITHUB_RUN_ID' in os.environ else None,
        python=platform.python_version(),platform=platform.platform(),
        packages={p:md.version(p) for p in ['numpy','scipy','matplotlib','scikit-fem','gmsh']},
        source_sha256={str(p.relative_to(COMSOL)):hashlib.sha256(p.read_bytes()).hexdigest() for p in paths},
        comsol_executed=False,comsol_runtime_speedup=None)
    data['checks']=CHECKS
    (OUT/'comparison.json').write_text(json.dumps(data,ensure_ascii=False,indent=2,allow_nan=False)+'\n',encoding='utf-8')
    m1,m7,mc,hc=[data[k] for k in ['M1','M7','MCF','HCF']]
    m1fine=[r for r in m1['mesh'] if r['mode']=='TE0'][-1]
    summary=[
        dict(case='M1 TE0',value=m1fine['neff'],metric='neff',reference_kind='published analytical equation; not COMSOL export',error=m1fine['abs_error'],error_unit='absolute neff'),
        dict(case='M7 LP01',value=m7['mesh'][-1]['neff'],metric='neff',reference_kind='independent scalar circular equation',error=m7['mesh'][-1]['scalar_abs_error'],error_unit='absolute neff'),
        dict(case='MCF',value=mc['extended_domain'][-1]['kappa_per_m'],metric='kappa 1/m',reference_kind='independent scalar multipole; not published COMSOL',error=mc['extended_domain'][-1]['kappa_error_pct'],error_unit='relative percent'),
        dict(case='HCF',value=hc['mesh'][-1]['neff'],metric='neff',reference_kind='published capillary approximation; not full COMSOL',error=hc['mesh'][-1]['neff_abs_error'],error_unit='absolute neff')]
    write_csv('summary.csv',summary)
    graph=None
    if (OUT/'mcf_graph_estimates.json').exists():
        graph=json.loads((OUT/'mcf_graph_estimates.json').read_text())
    report=['# 공개 레퍼런스 비교 결과: M1 · M7 · MCF · HCF','',
        '이 실행의 Python FEM 결과로 자동 생성했습니다. **COMSOL 소프트웨어를 실행하지 않았습니다.** 해석식·독립 기준 오차를 COMSOL 실측 오차로 표시하지 않습니다.','',
        '- 실행 UTC: '+data['metadata']['executed_utc'],'- 소스 커밋: '+commit,
        '- 실행 기록: '+str(data['metadata']['run_url']),
        '- 검사 결과: '+str(sum(x['passed'] for x in CHECKS))+'/'+str(len(CHECKS))+' 통과','',
        '## 핵심 수치','',
        table(['대상','계산값','기준','오차'],[[r['case'],f"{r['value']:.12g}",r['reference_kind'],f"{r['error']:.6g} ({r['error_unit']})"] for r in summary]),'',
        '## M1 슬랩','',
        '공식 C1 조건: n_core=1.5, n_clad=1, 전두께 1 µm, λ=1.55 µm, 전체 횡단면 7 µm(양쪽 padding=3 µm). 기존 M1 TE 해석기를 실행했습니다. 비교는 횡단면 고유문제이며 공식 예제의 포트·전파 해석은 대상이 아닙니다.','',
        table(['모드','h (µm)','FEM neff','독립 해석 neff','절대 오차','장 상대 L2 오차'],[[r['mode'],f"{r['h_um']:.6g}",f"{r['neff']:.12f}",f"{r['reference_neff']:.12f}",f"{r['abs_error']:.5g}",f"{r['field_relative_l2_error']:.5g}"] for r in m1['mesh']]),'',
        '**문서 불일치:** 공식 설명의 TE0만 전파한다는 문장은 주어진 전두께·굴절률의 일반적 모드 개수와 맞지 않습니다. 반두께 기준 V≈2.266>π/2이므로 TE0·TE1이 존재하며 FEM과 독립 분산식으로 확인했습니다. TE0 입사와 다른 유도 모드가 존재하지 않는 것은 구분해야 합니다. 이 문구를 맞추기 위해 모드를 삭제하지 않았습니다. TE1은 같은 padding에서 더 약하게 감쇠해 경계 영향이 클 수 있습니다.','',
        '![M1 FEM과 해석장](m1_fields.svg)','',
        '## M7 원형 광섬유','',
        '공식 C2의 코어 반경 8 µm, 외곽 40 µm, n_core=1.4457, n_clad=1.4378, λ=1550 nm를 기존 Fiber API와 radial FEM에 전달했습니다. LP01 별도 스칼라 기준식을 첫 J0 분기에서 풀어 V>2.405 기본 모드를 비교했습니다. 기존 UI/오라클 제한은 바꾸지 않았습니다.','',
        f"- 실제 V: {m7['parameters']['V']:.9f}",
        f"- 스칼라 LP01 기준 neff: {m7['scalar_reference']['neff']:.12f}",
        f"- 벡터 HE11 특성방정식 neff: {m7['vector_characteristic_reference']['neff']:.12f}",
        f"- 스칼라/벡터 모델 차이: {m7['scalar_vector_abs_model_difference']:.6g} (절대), {m7['scalar_vector_relative_model_difference_pct']:.6g}%",
        '- 공식 공개값은 **1.4444**로 소수점 네 자리입니다. 원시 수치가 없어 정밀 COMSOL 오차율은 산출하지 않았습니다. 반올림 구간 검사는 최근접 반올림을 가정한 참고 검사입니다.',
        '- 벡터 값은 독립 원형 Maxwell 특성방정식 계산입니다. 새 벡터 FEM이나 COMSOL 실행 결과가 아닙니다.','',
        table(['코어/클래딩 요소','FEM neff','스칼라 절대 오차'],[[f"{r['core_elements']}/{r['clad_elements']}",f"{r['neff']:.12f}",f"{r['scalar_abs_error']:.6g}"] for r in m7['mesh']]),'',
        '![M7 수렴](m7_convergence.svg)','',
        '## MCF 2코어','',
        'P1의 a=3.5 µm, n_clad=1.444, λ=1550 nm, 상대 Δ=0.25%, 가장자리 간격 dg=2 µm를 사용했습니다. 중심 간격 p=2a+dg=9 µm. 기본 변환은 Δ=(n_core²−n_clad²)/(2n_core²)를 가정하며, 원문 변환식이 명확하지 않아 대안의 민감도를 함께 계산했습니다.','',
        f"- 변환 n_core={mc['parameters']['n_core']:.12f}, 절대 Δn={mc['parameters']['delta_n']:.12f}",
        '- 곡선 P2 FEM을 독립 무한 클래딩 Fourier–Bessel 경계 정합과 비교했습니다. 논문 벡터/편광·경계·메시와 완전히 동일하다는 주장은 하지 않습니다.',
        '- Figure 2의 간격 증가 → κ 감소 경향을 재현했습니다. 그래프를 정밀 디지타이즈하지 않았고 원시 β/κ 표가 없어 **논문 COMSOL 대비 정밀 오차율은 미산출**입니다.','',
        table(['h (µm)','κ (1/m)','Lc (mm)','독립 스칼라 κ 오차 (%)'],[[r['h_um'],f"{r['kappa_per_m']:.9f}",f"{r['Lc_mm']:.9f}",f"{r['kappa_error_pct']:.6g}"] for r in mc['mesh']]),'',
        table(['Δ 가정','n_core','κ (1/m)','기본 가정 대비 κ 차이 (%)'],[[r['convention'],f"{r['n_core']:.12f}",f"{r['kappa_per_m']:.9f}",f"{r['kappa_difference_from_primary_pct']:.6g}"] for r in mc['delta_sensitivity']]),'',
        '**경계 영향 추가 확인:** L=40 µm에서 h를 줄여도 오차가 약 0.0086%에 머물렀습니다. 같은 h=0.35 µm로 L을 48/56/64 µm까지 늘려 독립 기준에 접근하는지 확인했습니다. 아래는 메시 세분화와 별도로 수행한 계산영역 확대 결과입니다. 기존 UI의 영역 제한은 변경하지 않았습니다.','',
        table(['L (µm)','h (µm)','κ (1/m)','Lc (mm)','독립 스칼라 κ 오차 (%)'],[[r['halfwidth_um'],r['h_um'],f"{r['kappa_per_m']:.9f}",f"{r['Lc_mm']:.9f}",f"{r['kappa_error_pct']:.6g}"] for r in mc['extended_domain']]),'',
        '![MCF 간격 변화](mcf_gap.svg)','',
        '## HCF','',
        'H1 Eq.(1)의 모세관 근사 neff=sqrt(n_gas²−(u01·λ/(2πR))²)를 기준으로 기존 radial P1 FEM을 검증했습니다. R=15 µm·λ=1550 nm·일정 n_gas=1.00027은 선택한 검증 입력이며 논문의 실제 AR-HCF 단면·압력을 재현한 것이 아닙니다. Aeff는 J0 적분, D는 같은 식을 일정 n_gas로 미분한 기준입니다.','',
        table(['요소','FEM neff','neff 절대 오차','Aeff (µm²)','Aeff 오차 (%)','D','D 오차 (%)'],[[r['elements'],f"{r['neff']:.12f}",f"{r['neff_abs_error']:.6g}",f"{r['Aeff_um2']:.8f}",f"{r['Aeff_error_pct']:.6g}",f"{r['D_ps_nm_km']:.9f}",f"{r['D_error_pct']:.6g}"] for r in hc['mesh']]),'',
        '**전체 HCF COMSOL 비교 미지원:** 현재 유리벽·반공진·PML·복소 neff가 없어 구속 손실이나 실제 NANF/ARF 오차율을 계산할 수 없습니다. C3의 유리 코어 미세구조 예제도 구조가 달라 비교에서 제외했습니다. 반경 변경은 물리 구조 변경이며 영역 수렴이 아닙니다.','',
        '![HCF 반경 변화](hcf_radius.svg)','',
        '## 오차 원인과 사용 범위','',
        '- 메시 의존성은 이산화와 외곽 경계 절단의 수치 오차입니다. 각 CSV에 메시/영역 변화를 분리했습니다. M7 영역 비교는 클래딩 요소 간격을 일정하게 유지합니다.',
        '- M7 스칼라/벡터 차이는 Maxwell 계면 조건을 약도파 근사로 줄인 모델 차이이며 메시만 늘려 제거할 수 없습니다.',
        '- MCF는 상대 Δ 정의·벡터/스칼라·원시 수치 부재가 남습니다. 독립 스칼라 오차를 논문 오차로 부르지 않습니다.',
        '- HCF 모세관 검증은 유리벽 반공진 구조의 검증이 아닙니다.',
        '- 실행 시간은 이 러너의 Python 시간만 JSON/CSV에 기록했습니다. 동일 하드웨어·정확도 COMSOL 자료가 없어 속도/메모리 우열은 판단하지 않습니다.',
        '- 이 조건의 검증은 G.654.E/G.657.A1 제품 인증·굽힘·감쇠·PMD·케이블 누화 정확도 보증이 아닙니다.','',
        '## 출처·재현·검사','',
        '[출처 전체](../references.json) · [실행법](../README.md) · [수치/환경/해시](comparison.json) · [CSV 요약](summary.csv)','',
        table(['검사','통과'],[[x['name'],str(x['passed'])] for x in CHECKS])]
    if graph:
        report+=['','## 논문 MCF 그림 수치 추출 검토','',
            '원시 COMSOL 수치표가 없어 Figure 2(a)의 녹색 FEM 곡선을 이미지 좌표로 추출했습니다. 축 보정과 원본 PDF 해시는 mcf_graph_estimates.json에 기록합니다. 그림 오차는 근사값이며 모델/Δ 정의/편광 차이를 포함한 엄밀한 수치 오차가 아닙니다.','',
            '추출 상태: '+graph.get('status','unknown')]
        if graph.get('status')=='estimated':
            report+=['',table(['가장자리 간격 (µm)','그림 κ 근사 (1/m)','이미지 좌표 불확실도 (±1/m)'],[[r['gap_um'],f"{r['kappa_estimate_per_m']:.3f}",f"{r['raster_uncertainty_per_m']:.3f}"] for r in graph['samples']]),
                '','![논문 Figure 2a 확인용 발췌](mcf_published_figure2a.png)','Zhao, Ren, Yin & Wang (2021), DOI 10.4236/opj.2021.118029, Figure 2(a), CC BY 4.0. 확인을 위한 그림 발췌·크롭이며 수치 추정은 이번 코드에서 수행했습니다.']
            p=next(r for r in graph['samples'] if r['gap_um']==2)
            diff=pct(mc['extended_domain'][-1]['kappa_per_m'],p['kappa_estimate_per_m'])
            report+=['',f"dg=2 µm: 현재 스칼라 κ={mc['extended_domain'][-1]['kappa_per_m']:.6f} /m, 그림 κ≈{p['kappa_estimate_per_m']:.3f}±{p['raster_uncertainty_per_m']:.3f} /m. 명목 상대 차이는 약 {diff:.3f}%입니다.",
                '**이 값은 그림 추출·축 가정이 맞는지 확인한 후 읽어야 합니다. 큰 차이가 남으면 재현 일치로 판정하지 않으며, Δ 정의·기하/편광·원문 수치 확보가 필요합니다. 독립 스칼라 FEM 오차가 작다는 이유로 논문과의 불일치를 덮지 않습니다.**']
        else:report+=['','자동 그림 추출을 완료하지 못했습니다. 원인: '+graph.get('error','unknown')+'; 정밀 논문 오차율은 미산출입니다.']
    (OUT/'COMPARISON_REPORT.md').write_text('\n'.join(report)+'\n',encoding='utf-8')
    print('COMPARISON_SUMMARY '+json.dumps(dict(metadata=data['metadata'],summary=summary,M7_vector=m7['vector_characteristic_reference'],MCF_delta_sensitivity=mc['delta_sensitivity'],checks=CHECKS),ensure_ascii=False),flush=True)
    if not all(c['passed'] for c in CHECKS):raise SystemExit('Declared acceptance checks failed; outputs retained.')
if __name__=='__main__':main()
