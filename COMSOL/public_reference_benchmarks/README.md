# 공개 레퍼런스 비교: M1 · M7 · MCF · HCF

COMSOL 공식 문서와 공개 논문에 나온 조건을 기존 Python FEM으로 재계산합니다. **COMSOL 프로그램을 직접 실행한 결과가 아닙니다.** 원문에 없는 정밀 COMSOL 수치나 실행 시간을 만들어 넣지 않습니다.

- [비교 결과 보고서](results/COMPARISON_REPORT.md) — 실제 실행 후 생성
- [수치 원본](results/comparison.json) · [비교 표 CSV](results/summary.csv)
- [출처·조건·변환·비교 가능 여부](references.json)
- [자동 실행](https://github.com/kimheeseo/LSCNS/actions/workflows/public-optics-reference-comparison.yml)

## 비교 구성

| 대상 | 레퍼런스 | 수행하는 비교 | 구분 |
|---|---|---|---|
| M1 | COMSOL Dielectric Slab Waveguide | 같은 단면의 TE 고유값·장 분포·메시·경계 수렴 | 공개 방정식 대비 수치 검증. 포트/전파 계산 제외 |
| M7 | COMSOL Step-Index Fiber | a=8 µm, n_core=1.4457, n_clad=1.4378, 1550 nm | 공개 neff=1.4444의 반올림 수준 비교 + scalar/vector 모델 차이 |
| MCF | Zhao et al., 2021, DOI 10.4236/opj.2021.118029 | 같은 2코어 조건의 κ·Lc·간격 경향, 독립 원통함수 기준 | 논문 그래프와 정성 비교. 원시 COMSOL 수치 부재로 정밀 오차율 미산출 |
| HCF | Kelly et al., 2021, DOI 10.1364/OPTICA.424224, Eq.(1) | 공개 모세관 근사식 대비 neff·D·Aeff 메시 수렴 | 모세관 근사 검증. ARF/NANF COMSOL 손실 비교 아님 |

M7 비교는 기존 `radial_fem.solve_radial`의 상수 굴절률 입력을 별도 스크립트에서 사용합니다. UI의 V<2.405 제한이나 기존 검증 기준을 변경하지 않습니다. V>2.405의 기본 모드 비교를 다른 고차 모드 전체의 검증으로 해석하지 않습니다.

MCF 논문의 Δ=0.25%는 앱의 절대 Δn=0.0025와 다릅니다. 변환 가정과 대안의 민감도를 함께 계산합니다. 논문 간격 dg는 가장자리 간격이며 앱의 중심 간격 p=2a+dg로 바꿉니다.

## 실행

저장소 최상위에서 Python 3.12와 필요한 Linux 공유 라이브러리를 준비합니다. 기존 COMSOL/research/requirements.txt의 고정 버전을 사용합니다.

```bash
python -m pip install -r COMSOL/research/requirements.txt
python COMSOL/public_reference_benchmarks/run_comparison.py
```

GitHub Actions는 같은 명령으로 실행하고 생성된 JSON/CSV/Markdown/SVG를 이 폴더의 results에 올립니다. 출력에는 소스 커밋·패키지 버전·파일 SHA256을 기록합니다.

## 오차·시간 해석

neff는 절대 차이와 상대 차이를 함께 기록하며, κ/Lc/Aeff/D는 상대 오차(%)를 사용합니다. 기준이 해석식인지, 독립 수치 방법인지, 공개된 COMSOL 반올림값인지 구분합니다. 메시 수렴 허용치는 이번 벤치마크를 위한 채택 기준이며 제품 정확도 보증이 아닙니다.

실행 시간은 해당 GitHub 러너에서 Python 계산에 걸린 시간만 기록합니다. COMSOL의 동일 하드웨어·동일 정확도 실행 자료가 없어 속도 배수나 메모리 우열은 계산하지 않습니다.

공식 미세구조 광섬유 예제는 유리 코어·공기 구멍·벡터/PML 구조입니다. 현재 HCF의 중공 Dirichlet 원판과 다르므로 두 neff를 빼서 HCF 오차율을 계산하지 않습니다.
