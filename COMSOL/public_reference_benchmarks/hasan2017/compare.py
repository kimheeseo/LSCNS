"""Hasan arXiv:1708.06879v2 Fig.2 vs existing ideal-capillary FEM.
This is a MODEL DISCREPANCY comparison, not validation of an ARF solver.
pip install numpy scipy pillow pymupdf
python compare.py [--pdf /path/to/paper.pdf] [--solver /path/to/capillary_fem.py]
"""
from pathlib import Path
import argparse, hashlib, importlib.util, io, json, math, csv, platform
from datetime import datetime, timezone
from urllib.request import urlopen
import numpy as np
from PIL import Image
import pymupdf

ROOT = Path(__file__).resolve().parent
PIN = 'a65fc034fdd9c658ce57bfebe6ee47497d5690e3'
PDF_URL = 'https://arxiv.org/pdf/1708.06879v2'
SOLVER_URL = f'https://raw.githubusercontent.com/kimheeseo/LSCNS/{PIN}/COMSOL/waveoptics_m8/src/capillary_fem.py'
C = 299792458.

def empirical(wl, f2):
    # Eq.(3)/(4), using Fig.2 fitted parameters, wavelengths/lengths in um.
    b = f2 / (12. * .2)
    a = 2.405**2 / (8 * math.pi**2 * 1.1039**2 * 12.**2)
    ne = 1 - a * wl**2 / (1-b*wl**2)**2
    npp = -2*a*(1+8*b*wl**2+3*b*b*wl**4)/(1-b*wl**2)**4
    beta2_cm = wl**3*npp*1e22/(2*math.pi*C*C)
    return ne, beta2_cm

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--pdf', type=Path)
    ap.add_argument('--solver', type=Path)
    args = ap.parse_args()
    pdf = args.pdf.read_bytes() if args.pdf else urlopen(PDF_URL, timeout=60).read()
    solver = args.solver.read_bytes() if args.solver else urlopen(SOLVER_URL, timeout=60).read()
    if hashlib.sha256(pdf).hexdigest() != '6b45b8848232612d95092290af5a29a3b3c2d969464503ee271e6a09b535c3c6':
        raise ValueError('PDF differs from inspected v2; recalibrate figure axes before use')
    if hashlib.sha256(solver).hexdigest() != 'a501e437a6c093caedcd063e1b8dba93d382ce0cafefb80b71523c0b0c058d33':
        raise ValueError('Solver differs from pinned source')
    # Execute the pinned existing solver, not a fitted replacement.
    namespace = {'__name__': 'pinned_capillary'}
    exec(compile(solver, 'pinned_capillary_fem.py', 'exec'), namespace)
    solve = namespace['solve_capillary']
    doc = pymupdf.open(stream=pdf, filetype='pdf')
    images = [doc.extract_image(x[0]) for x in doc[3].get_images()]
    figure = next(x['image'] for x in images if (x['width'],x['height']) == (2512,465))
    rgb = np.array(Image.open(io.BytesIO(figure)).convert('RGB')).astype(float)
    green = (rgb[:,:,1] > 70) & (rgb[:,:,1] > 1.3*rgb[:,:,0]) & (rgb[:,:,1] > 1.3*rgb[:,:,2])
    # Native embedded Fig.2 pixels, manually inspected against printed ticks.
    # x ranges .5..1.5 um; y 0..-100 fs^2/cm.
    axes = [(6,173.,921.,.0480), (8,968.,1716.,.0608), (10,1762.,2510.,.0729)]
    top,bottom = 16.5,362.
    rows=[]
    for n,left,right,f2 in axes:
        for wl in [.6,.7,.8,.9,1.,1.1,1.2,1.3]:
            x = round(left+(wl-.5)*(right-left))
            en,eb = empirical(wl,f2)
            predicted_y = top-eb/100*(bottom-top)
            # The empirical curve only locates the window. Reference values
            # are measured from green FEM pixels, never filled from the fit.
            y0=max(20,int(predicted_y-20)); y1=min(300,int(predicted_y+21))
            yy,xx=np.nonzero(green[y0:y1,x-10:x+11])
            if len(yy)<2:
                raise RuntimeError(f'No usable green FEM pixels: n={n}, wl={wl}')
            yy=yy+y0
            ref=-100*(float(np.median(yy))-top)/(bottom-top)
            # Conservative raster reading interval: observed spread plus
            # four native pixels for manual axis placement/curve ambiguity.
            unc=100*((float(yy.max()-yy.min())/2)+4)/(bottom-top)
            s=solve(wl,12.,1.,640)
            current=-s['D_ps_nm_km']*1e-6*(wl*1e-6)**2/(2*math.pi*C)*1e28
            rows.append(dict(tubes=n,wavelength_um=wl,
                published_FEM_GVD_estimate_fs2_cm=ref,raster_reading_halfwidth_fs2_cm=unc,
                current_capillary_FEM_GVD_fs2_cm=current,
                signed_difference_fs2_cm=current-ref,
                nominal_absolute_difference_pct=100*abs(current-ref)/abs(ref),
                paper_empirical_GVD_recomputed_fs2_cm=eb,
                paper_empirical_neff_recomputed=en,
                current_capillary_neff=s['n_eff'],pixel_x=x,
                green_pixel_y_min=int(yy.min()),green_pixel_y_max=int(yy.max())))
    # Verify differentiating Eq.(4) and capillary mesh sensitivity separately.
    derivative=[]
    for f2 in [.0480,.0608,.0729]:
        a=2.405**2/(8*math.pi**2*1.1039**2*12**2); b=f2/2.4
        f=lambda z: 1-a*z*z/(1-b*z*z)**2
        exact=empirical(1.,f2)[1]
        for h in [.01,.005,.0025]:
            npp=(-f(1+2*h)+16*f(1+h)-30*f(1)+16*f(1-h)-f(1-2*h))/(12*h*h)
            val=npp*1e22/(2*math.pi*C*C)
            err=abs(val/exact-1)
            if err>1e-5: raise AssertionError('Eq4 derivative cross-check')
            derivative.append(dict(f2=f2,step_um=h,relative_difference=err))
    mesh=[]
    for N in [320,640,1280]:
        s=solve(1.,12.,1.,N)
        mesh.append(dict(elements=N,neff=s['n_eff'],D_ps_nm_km=s['D_ps_nm_km']))
    metadata=dict(executed_utc=datetime.now(timezone.utc).isoformat(),python=platform.python_version(),
        reference_url=PDF_URL,pdf_sha256=hashlib.sha256(pdf).hexdigest(),
        solver_source_commit=PIN,solver_url=SOLVER_URL,solver_sha256=hashlib.sha256(solver).hexdigest(),
        figure_embedded_size=[2512,465],figure_sha256=hashlib.sha256(figure).hexdigest(),
        axes_pixels=dict(y_zero=top,y_minus100=bottom,x_ranges=axes,sample_x_halfwidth=10),
        parameters=dict(radius_um=12,wall_um=.2,gap_um=1,n_air_constant=1.,elements=640),
        COMSOL_used_in_this_work=False,COMSOL_use_in_arxiv_text='not specified; FEM and PML specified',
        interpretation='different physical models; discrepancies are NOT ARF solver error or product accuracy',
        uncertainty='heuristic raster reading interval, not statistical or total physical uncertainty')
    result=dict(metadata=metadata,rows=rows,capillary_mesh=mesh,empirical_derivative_checks=derivative)
    (ROOT/'comparison.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n')
    with (ROOT/'comparison_24.csv').open('w',newline='',encoding='utf-8-sig') as h:
        w=csv.DictWriter(h,fieldnames=list(rows[0]));w.writeheader();w.writerows(rows)
    report=['# Hasan 2017: 공개 FEM 그래프와 현재 HCF 모델 비교','',
        '**현재 모세관 모델의 적용 한계를 확인하는 비교입니다. ARF 솔버 검증 통과를 뜻하지 않습니다.**',
        '', '출처: [arXiv:1708.06879v2](https://arxiv.org/abs/1708.06879v2), Figure 2. 본문에는 FEM·PML이 명시되어 있으나 COMSOL 명칭은 확인되지 않았습니다.',
        '', '## 비교 방법','',
        '- 기준값: PDF 4쪽 Figure 2의 **초록색 FEM 곡선**을 원본 내장 이미지에서 추출했습니다. 빨간 경험식으로 기준값을 대체하지 않았습니다.',
        '- 6/8/10개 튜브 × 파장 0.6–1.3 µm의 8개 지점 = 24개 GVD 표본입니다. 독립적인 24개 모델 검증을 뜻하지 않습니다.',
        '- 논문: R=12 µm, t=0.2 µm, g=1 µm인 반공진 튜브 구조. 현재 FEM: R=12 µm, 일정 n_air=1인 J0/Dirichlet 원판. 현재 솔버에는 튜브·유리벽·PML이 없습니다.',
        '- 단위는 그림과 동일한 **fs²/cm**이며 통신 분산 D[ps/(nm·km)]와 구분합니다.',
        '- 각 파장 위치의 좌우 10픽셀(약 ±0.0134 µm)에서 초록 곡선을 읽습니다. ±범위는 이 창의 곡선 픽셀 분포와 축 보정 여유 4픽셀로 추정한 읽기 범위입니다. 통계적 신뢰구간이나 전체 오차 한계가 아닙니다. 작은 차이에 대한 일치 판정에는 원시 FEM 수치가 필요합니다.',
        '', '## 공개 FEM 추정값 vs 현재 계산값','',
        '| 튜브 수 | 파장 (µm) | 논문 FEM 그래프 (fs²/cm) | 현재 모세관 FEM (fs²/cm) | 명목 차이 (%) |',
        '|---|---|---|---|---|']
    for r in rows:
        report.append(f"| {r['tubes']} | {r['wavelength_um']:.1f} | {r['published_FEM_GVD_estimate_fs2_cm']:.1f} ± {r['raster_reading_halfwidth_fs2_cm']:.1f} | {r['current_capillary_FEM_GVD_fs2_cm']:.2f} | {r['nominal_absolute_difference_pct']:.1f} |")
    report += ['', '## 판정과 필요한 개선','',
        '현재 모델은 같은 반경/파장에서 튜브 수와 관계없이 같은 값을 냅니다. 논문 FEM 곡선은 튜브 수에 따라 달라집니다. 따라서 현재 이상 모세관은 이 ARF의 구조별 분산을 독립적으로 예측하는 도구로 검증되지 않았습니다. 이 차이는 메시 증가만으로 해결되지 않습니다.',
        '', '논문의 Eq.(3)/(4)를 별도 경험식으로 구현해 CSV/JSON에 재계산 값을 기록했습니다. 이는 저자의 FEM에 맞춘 모델을 재현한 것으로, 새로운 독립 FEM 검증이 아닙니다. 공진에서 떨어진 조건에 한정하며, Eq.(2)의 공진항은 이번 비교에서 구현하지 않았습니다.',
        '', '실제 ARF 계산을 검증하려면 튜브 단면·유리 분산·외곽 영역·PML·벡터 모드 및 모드 추적을 구현하고, 얇은 벽 메시·영역/PML·파장 미분 간격을 각각 수렴시킨 뒤 같은 그래프와 비교해야 합니다. 논문의 미공개 재료/PML 세부는 가정을 기록해야 합니다.',
        '', 'arXiv v2의 5쪽 본문에서는 유효면적 결과/공식이나 손실 곡선이 확인되지 않았습니다. 제목/초록의 유효면적까지 검증했다고 주장하지 않습니다. 정식 출판본(JLT 36, 4060–4065, 2018; DOI 10.1109/JLT.2018.2854722)은 별도 판본입니다.',
        '', '## 재현','', '```bash', 'pip install numpy scipy pillow pymupdf', 'python compare.py', '```',
        '', '논문 PDF와 기존 솔버 소스를 내려받습니다. 로컬 파일은 `--pdf`, `--solver`로 지정할 수 있습니다. [수치·해시·설정](comparison.json) · [24개 표본 CSV](comparison_24.csv).',
        '', '논문 경험식의 해석적 2차 미분을 세 파장 간격의 5점 차분으로 확인했고 상대 차이는 모두 1e-5 미만이었습니다. 모세관 메시 320/640/1280 결과도 JSON에 기록했습니다. 이 확인은 ARF 모델 검증과 별개입니다.']
    (ROOT/'README.md').write_text('\n'.join(report)+'\n',encoding='utf-8')
    print(json.dumps(dict(rows_at_1um=[r for r in rows if r['wavelength_um']==1],mesh=mesh,metadata=metadata),indent=2))

if __name__=='__main__': main()
