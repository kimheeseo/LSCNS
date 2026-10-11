"""Generate verified M4/M5/M6 numerical results, figures and markdown reports.
No result is labeled validated unless its assertions pass. M4 is a limited rebuilt
PEC-cavity benchmark, not a restoration of missing prior PML/port implementations.
"""
from pathlib import Path
import json,sys
ROOT=Path(__file__).resolve().parent
for k in (4,5,6):sys.path.insert(0,str(ROOT/f'waveoptics_m{k}'/'src'))
from m4_cavity import run as m4run,solve as m4solve,analytic_te101
from m5_spectral import run as m5run,spectrum,silica_index,silica_group_index
from m6_inverse_design import run as m6run,reflectance
import numpy as np
def check(condition,message):
    if not condition:raise AssertionError(message)
def main():
    m4dir=ROOT/'waveoptics_m4';m5dir=ROOT/'waveoptics_m5';m6dir=ROOT/'waveoptics_m6'
    for p in (m4dir,m5dir,m6dir):(p/'screenshots').mkdir(parents=True,exist_ok=True)
    m4=m4run(m4dir/'screenshots')
    check(len(m4)==3,'three mesh convergence runs')
    check(all(x['relative_residual']<1e-5 for x in m4),'eigenvector residual')
    check(m4[-1]['relative_error']<.20,'3D cavity TE101 analytic reference error <20%')
    check(m4[-1]['relative_error']<m4[0]['relative_error'],'refinement trend')
    case,_=m4solve(index=1.5,shape=(3,2,4))
    check(abs(case['frequency_hz']*1.5/m4[0]['frequency_hz']-1)<.025,'n=1.5 frequency scaling')
    m4report=f"""# M4 검증 보고서 — 3D Nédélec PEC cavity 수치 재구현 (제한적 복구)

**재검증 범위:** 실제 3D tetrahedral first-order Nédélec FEM, 등방성 비자성 유전체, PEC 벽, 전자기 cavity eigenmode. 기존 M4의 3D PML / mode port / complex S 행렬 전체 소스는 현재 GitHub에서 찾지 못하여 **복원·재검증했다고 주장하지 않습니다.**

## 방법
좌표 (x,y,z) 단위 μm. Maxwell 고유치 ∇×μᵣ⁻¹∇×E=k₀²εᵣE. Edge Nédélec Nᵢⱼ=λᵢ∇λⱼ−λⱼ∇λᵢ. Curl exact + tetra 4-point degree-2 mass quadrature, sparse generalized eigsh shift-invert, tangential E=0 PEC edges.

독립 직육면체 공동 TE101 기준: f=c/(2n)√[(1/Lx)²+(1/Lz)²], Lx=2, Ly=1, Lz=3 μm; 이산화에 따른 오차/잔차.

| 분할(nx,ny,nz) | tetra | DOF(자유) | FEM frequency (THz) | analytic (THz) | 상대오차 (%) | 선형 고유잔차 |
|---|---:|---:|---:|---:|---:|---:|
"""
    for r in m4:m4report+=f"| {r['shape']} | {r['tetrahedra']} | {r['free_dofs']} | {r['frequency_hz']/1e12:.6f} | {r['analytic_te101_hz']/1e12:.6f} | {100*r['relative_error']:.3f} | {r['relative_residual']:.2e} |\n"
    m4report+=f"""
## 추가 검증
n=1.5의 고유주파수는 n=1.0 대비 1/n 스케일법칙 **오차 {100*abs(case['frequency_hz']*1.5/m4[0]['frequency_hz']-1):.3f}%**로 일치.

![3D field](screenshots/m4_cavity_field.png)
![convergence](screenshots/m4_convergence.png)

**통과:** 3개 메시, PEC 고유주파수 독립 기준, 고유잔차, 메시 개선 추세, n 스케일링. **미검증:** 3D mode ports, full-vector PML, S parameters, 실제 COMSOL 비교, 물리적 개방 광섬유 모드.
"""
    (m4dir/'M4_VALIDATION_REPORT.md').write_text(m4report,encoding='utf8')
    m5=m5run(m5dir/'screenshots')
    check(m5['max_energy_error']<1e-9,'unitary lossless multilayer')
    check(m5['AR_reflectance_1p55']<m5['bare_reflectance_1p55']*.01,'quarter-wave AR near zero')
    check(m5['group_index_1p55']>m5['sellmeier_n_1p55'],'positive material group contribution')
    same=spectrum([1.2,1.5,1.65],[(1.0,.25)],substrate=1.0)
    check(np.max(same['R'])<1e-20,'uniform background must not reflect')
    (m5dir/'M5_VALIDATION_REPORT.md').write_text(f"""# M5 광 재료 분산 및 파장별 S-행렬 검증 보고서

**M5 범위:** 파장별 복소 S11/S21, 실리카 Sellmeier 물질 분산, TE/TM/입사각 비교, 위상 기반 group delay. 1D Maxwell transfer-matrix 별도 해석 오라클로, M3 2D FEM 또는 M4 3D FEM에서 직접 sweep한 결과는 아닙니다.

- λ: 1.30–1.65 μm, 101점; 기준 λ₀=1.55 μm.
- fused silica n(1.55)={m5['sellmeier_n_1p55']:.9f}, group index={m5['group_index_1p55']:.9f}.
- 이상적인 단층 코팅 n={m5['quarterwave_n']:.7f}, 두께={m5['quarterwave_depth_um']:.7f} μm.
- uncoated R(1.55)={m5['bare_reflectance_1p55']:.8g}, coated R(1.55)={m5['AR_reflectance_1p55']:.8g}.
- R+T=1의 최대 오차 {m5['max_energy_error']:.3g}; 균일 매질 R≈0 추가 검증.
- 파장별 복소 전송 위상으로 평균 지연 약 {m5['mean_group_delay_fs']:.5f} fs (코팅 전파 구간 기준).
- 독립 재료 모델이며 실제 상용 코팅/고굴절률 광섬유 사양으로 검증한 값은 아님.

![M5 spectral response](screenshots/m5_spectrum.png)

**후속 보완:** M4 3D ports/PML 복구 이후 수치 FEM의 S(λ)와 독립 정식화 교차검증.
""",encoding='utf8')
    m6=m6run(m6dir/'screenshots')
    check(m6['success'],'bounded multi-start solver exit success')
    check(1.05<=m6['best_index']<=1.42 and .15<=m6['best_thickness_um']<=.44,'manufacturing design constraints')
    check(m6['objective_R']<m6['bare_R']*.1,'effective reduction')
    check(m6['objective_R']<=m6['quarterwave_R']+1e-8,'not worse than quarter-wave initial feasible design')
    (m6dir/'M6_VALIDATION_REPORT.md').write_text(f"""# M6 광학 단층 박막 역설계 · 제조 공차 민감도 검증 보고서

**M6 범위:** M5의 Maxwell 전송행렬을 forward model로 사용하는 경계 제약 최적화. TE/TM, 0°/20°, λ=1.30–1.65 μm 81점 및 두께 ±3% 공차를 포함한 평균 반사율 최소화.

최적화: 다중 초기점 + SciPy L-BFGS-B, n ∈ [1.05,1.42], 두께 d ∈ [0.15,0.44] μm. **실제 사용 가능한 박막 재료 또는 공정 호환성을 검증한 것이 아님**.

| 출력 항목 | 계산 |
|---|---:|
| 최적 박막 굴절률 | {m6['best_index']:.8f} |
| 최적 두께 (μm) | {m6['best_thickness_um']:.8f} |
| 무코팅 평균 R | {m6['bare_R']:.8g} |
| 통상 quarter-wave 강건 평균 R | {m6['quarterwave_R']:.8g} |
| 최적 강건 평균 R | {m6['objective_R']:.8g} |
| 무코팅 대비 반사율 감소 | {m6['improvement_vs_bare_pct']:.3f}% |
| quarter-wave 대비 감소 | {m6['improvement_vs_quarterwave_pct']:.3f}% |

검증: 다중 시작점 최적화 성공, 설계 bounds 만족, 무코팅 대비 개선, 기준 quarter-wave 해 대비 개선. 재료 분산은 실리카 기판에만 반영하였고 코팅 박막은 상수 n 이상화 가정.

![Inverse design](screenshots/m6_optimization.png)

**후속 보완:** 다층막, 실제 박막 Sellmeier 데이터, 생산 제약(흡수, 응력, 두께 분포), 2D/3D FEM spotcheck.
""",encoding='utf8')
    summary={"stage":"M4-M6","m4":{"num_meshes":3,"finest_frequency_THz":m4[-1]['frequency_hz']/1e12,"finest_relative_error":m4[-1]['relative_error']},
    "m5":m5,"m6":m6,"validated_assertions":12,"scope":"M4 cavity FEM subset; M5/M6 1D Maxwell transfer matrix; not COMSOL equivalent"}
    (ROOT/'WAVEOPTICS_M4_M6_EXECUTION_SUMMARY.json').write_text(json.dumps(summary,indent=2),encoding='utf8')
    print(json.dumps(summary,indent=2))
if __name__=="__main__":main()
