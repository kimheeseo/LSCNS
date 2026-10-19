"""Build a self-contained Colab report; embed the exact source files and hashes."""
from pathlib import Path
import ast,hashlib
import nbformat as nbf
ROOT=Path(__file__).resolve().parents[1]
SOURCES=[ROOT.parent/'waveoptics_m7'/'src'/'fiber_modes.py',ROOT/'src'/'multipole_reference.py',ROOT/'src'/'capillary_fem.py',ROOT/'src'/'conforming_mcf.py',ROOT/'src'/'mcf_hcf.py',ROOT/'validation'/'run_validation.py']


def embedded_source():
    parts=[];hashes={}
    for p in SOURCES:
        text=p.read_text();tree=ast.parse(text);key=str(p.relative_to(ROOT.parent))
        hashes[key]=hashlib.sha256(text.encode()).hexdigest()
        lines=[]
        for node in tree.body:
            if p.name=='mcf_hcf.py':
                if isinstance(node,ast.FunctionDef) and node.name in ['solve_two_core','capillary_neff','capillary_dispersion']:
                    lines.append(ast.get_source_segment(text,node))
                continue
            if isinstance(node,ast.ImportFrom) and node.module in ['fiber_modes','multipole_reference','mcf_hcf','capillary_fem','conforming_mcf']:continue
            if isinstance(node,ast.Assign) and any(isinstance(t,ast.Name) and t.id=='ROOT' for t in node.targets):continue
            if isinstance(node,ast.Expr) and isinstance(node.value,ast.Call) and ast.unparse(node.value.func)=='sys.path.insert':continue
            if isinstance(node,ast.Assign) and any(isinstance(t,ast.Subscript) and ast.unparse(t.value)=='sys.path' for t in node.targets):continue
            if isinstance(node,ast.If) and '__name__' in ast.unparse(node.test):continue
            if isinstance(node,(ast.FunctionDef,ast.ClassDef)) and node.decorator_list:
                lines.append('\n'.join(text.splitlines()[node.decorator_list[0].lineno-1:node.end_lineno]))
            else:lines.append(ast.get_source_segment(text,node))
        if p.name=='mcf_hcf.py':lines.insert(0,'U01=2.404825557695773')
        parts.append('# SOURCE: '+key+'\n'+'\n\n'.join(lines))
    return '\n\n'.join(parts),hashes


def build():
    engine,hashes=embedded_source()
    ns={'__name__':'__main__'};exec(compile(engine,'<embedded-engine>','exec'),ns)
    assert ns['Fiber'](wavelength_um=1.55).radius_um==4.1
    assert hasattr(ns['stiffness'],'assemble')
    nb=nbf.v4.new_notebook()
    md=nbf.v4.new_markdown_cell;code=nbf.v4.new_code_cell
    nb.cells=[md('''# M8: HCF 모세관 · 2코어 MCF 검증 보고서

**런타임 → 모두 실행**으로 재계산합니다. 저장된 결과도 바로 볼 수 있습니다.

HCF는 이상적 모세관 J0/Dirichlet 모델입니다. 실제 ARF/NANF의 누설 손실 검증은 포함하지 않습니다. MCF는 동일 원형 코어 2개의 약도파 스칼라 모델입니다. 오차율은 동일 모델의 독립 기준 해 대비 값이며, 실물 제품 대비 오차가 아닙니다.

[HCF 상세 보고서](https://github.com/kimheeseo/LSCNS/blob/main/COMSOL/waveoptics_m8/HCF_VALIDATION_REPORT.md) · [MCF 상세 보고서](https://github.com/kimheeseo/LSCNS/blob/main/COMSOL/waveoptics_m8/MCF_2CORE_VALIDATION_REPORT.md)'''),
    code('''# 설치: Colab/Linux에서 Gmsh 공유 라이브러리가 필요합니다.
import sys,subprocess,importlib.util
try:
    import gmsh
except (ImportError,OSError):
    subprocess.check_call(['apt-get','update','-qq'])
    subprocess.check_call(['apt-get','install','-y','-qq','libglu1-mesa','libxft2','libxrender1','libxinerama1','libxcursor1'])
required=['numpy==2.3.5','scipy==1.17.0','matplotlib==3.10.8','scikit-fem==12.0.2','gmsh==4.15.2']
# 이미 검증된 버전이면 설치를 생략합니다.
from importlib.metadata import version,PackageNotFoundError
needs_install=False
for requirement in required:
    name,wanted=requirement.split('==')
    try:needs_install |= version(name)!=wanted
    except PackageNotFoundError:needs_install=True
if needs_install:
    subprocess.check_call([sys.executable,'-m','pip','install','-q',*required])
    print('설치 완료. 기존에 다른 버전을 import했다면 런타임을 재시작한 뒤 모두 실행하십시오.')
else: print('검증된 의존성 버전 확인 완료')'''),
    md('''## 입력값

파장은 µm, 코어 반경·간격·메시·해석영역은 µm입니다. 간격은 코어 중심 사이 거리입니다. Δn는 core−cladding입니다. 공기 굴절률은 파장에 대해 일정하다고 가정합니다.

MCF는 단일 코어 V<2.405, 동일 코어, 비중첩 구조를 사용합니다. 아래 설정은 보고서 검증 예입니다. 새 조건에서 판정 실패가 나오면 실패 수치와 수렴 결과를 함께 확인하십시오.'''),
    code('''#@title 기본 설계 입력
wavelength_um = 1.55 #@param {type:"number"}
core_radius_um = 4.1 #@param {type:"number"}
delta_n = 0.005 #@param {type:"number"}
pitch_um = 16.0 #@param {type:"number"}
hcf_radius_um = 15.0 #@param {type:"number"}
n_air = 1.00027 #@param {type:"number"}
CONFIG=dict(wavelength_um=wavelength_um,core_radius_um=core_radius_um,delta_n=delta_n,
            pitch_um=pitch_um,hcf_radius_um=hcf_radius_um,n_air=n_air,
            mesh_sizes_um=[1.0,0.7,0.45,0.3],halfwidth_um=40.0,
            domain_halfwidths_um=[28.0,32.0,40.0],pitches_um=[12.0,16.0,20.0,24.0],
            hcf_elements=[40,80,160,320,640])
print(CONFIG)'''),
    md('''## 내장 해석 엔진

아래 셀은 저장소의 실제 소스를 내장합니다. HCF 방사형 FEM과 Bessel 기준, MCF 곡선 P2 FEM과 독립 무한영역 원통함수 정합을 별도로 계산합니다. 외부 저장소를 런타임에 내려받지 않습니다.'''),
    code(engine),
    code('''from pathlib import Path
import tempfile
OUTPUT=Path(tempfile.mkdtemp(prefix='m8_hcf_mcf_'))
result=run_validation(OUTPUT,CONFIG,include_legacy=True)
print('계산값과 CSV/PNG 저장:',OUTPUT)'''),
    md('''## 계산값과 오차율

오차율(%) = 100×|계산−기준|/|기준|입니다. 작은 neff 오차와 별도로 MCF의 Δneff 및 결합 길이 오차를 확인합니다. 통과 개수는 오차율이 아닙니다.'''),
    code('''from IPython.display import display,Markdown,Image
h=result['hcf_finest'];m=result['mcf_finest']
text=f"""|항목|계산값|기준 대비 오차|
|---|---:|---:|
|HCF neff|{h['n_eff']:.13f}|절대 {h['neff_abs_error']:.6g} / 상대 {h['neff_error_pct']:.6g}%|
|HCF D [ps/(nm km)]|{h['D_ps_nm_km']:.10f}|{h['D_error_pct']:.6g}%|
|HCF Aeff [um²]|{h['Aeff_um2']:.8f}|{h['Aeff_error_pct']:.6g}%|
|MCF even neff|{m['n_even']:.13f}|절대 {m['neff_even_abs_error']:.6g}|
|MCF odd neff|{m['n_odd']:.13f}|절대 {m['neff_odd_abs_error']:.6g}|
|MCF Δneff|{m['index_split']:.12g}|{m['split_error_pct']:.6g}%|
|MCF Lc [mm]|{m['coupling_length_mm']:.10f}|{m['Lc_error_pct']:.6g}%|"""
display(Markdown(text))
print('수치 판정:',result['passed'],'/',result['total'])
for check in result['checks']:
    print('PASS' if check['passed'] else 'FAIL',check['name'],check['value'],'<=',check['limit'])'''),
    code('''for name in ['01_convergence.png','02_mcf_supermodes.png','03_coupling.png']:
    display(Image(filename=str(OUTPUT/name)))'''),
    md('''## 한계와 오차 원인

HCF 메시 오차는 방사형 P1 이산화에서 발생합니다. 유리벽·반공진·PML·구속 손실은 모델에 없습니다. 실물 HCF 대비 오차율은 미확인입니다.

MCF 잔여 오차는 FEM 공간/기하 근사와 외곽 경계에서 발생합니다. 기존 Cartesian P1은 재료 경계 분류도 변합니다. 원통함수 기준에도 전개 차수 절단·근 찾기 오차가 있으므로 차수와 각도 샘플 수를 검사합니다. 메시·영역·기준 해 수렴을 함께 봐야 합니다.

전력 전달은 동일 코어의 두 모드에 대한 무손실 모델입니다. z=Lc에서 P2=1이며 비팅 주기는 2Lc입니다. 실제 광섬유의 랜덤 누화(dB/km), 굽힘, 편광, 제조 편차를 포함하지 않습니다.

문헌: [Marcatili & Schmeltzer (1964)](https://doi.org/10.1002/j.1538-7305.1964.tb04108.x), [White et al. (2002)](https://doi.org/10.1364/JOSAB.19.002322). 이번 스칼라 코드는 이 논문들의 전체 벡터 모델을 구현하거나 실험 표를 재현한 것은 아닙니다.'''),
    code('''import zipfile
PACKAGE=OUTPUT/'M8_HCF_MCF_Results.zip'
with zipfile.ZipFile(PACKAGE,'w',zipfile.ZIP_DEFLATED) as z:
    for path in sorted(OUTPUT.iterdir()):
        if path.is_file() and path!=PACKAGE:z.write(path,path.name)
print('결과 ZIP:',PACKAGE)
try:
    from google.colab import files
    files.download(str(PACKAGE))
except ImportError:
    print('로컬 실행에서는 출력 디렉터리에서 ZIP을 확인하십시오.')''')]
    nb.metadata.update(kernelspec=dict(display_name='Python 3',language='python',name='python3'),language_info=dict(name='python',version='3.12'),
                       colab=dict(name='M8_HCF_MCF_2Core_Validation.ipynb'),embedded_source_sha256=hashes,
                       execution_verification='Execute with validation/replay_notebook.py; Colab browser UI not manually tested.')
    nb.cells[3].metadata['cellView']='form'
    path=ROOT/'M8_HCF_MCF_2Core_Validation.ipynb';nbf.validate(nb);nbf.write(nb,path);print(path)
    return path


if __name__=='__main__':build()
