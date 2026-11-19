# STEP 3 v0.8.5 — 검증표 (2026-11-19)

STEP 1·2 도구 적용은 선행 commit `4343ba96ad5d57bfbda830c2e73664913afde4d4`에 기록했다. STEP 3는 기존 조립/스칼라/Yee/Colab 모듈 위에 기능을 추가했다. 이 보고서는 **아래 명시한 합성 원형 시험의 수치 검증**이며 COMSOL 동등성이나 임의 섬유의 제품 성능 인증이 아니다.

## 검증 정책과 결과

V번호의 원래 개별 사양이 제공되지 않았으므로 [설계서](STEP3_V085_DESIGN_20261119.md)에 정의한 공개 기준을 사용했다. 허용오차는 공인 규격이 아닌 이번 수용 정책이다. 계산 후 기준을 완화하지 않았다.

| 항목 | 독립 기준 / 실제 계산 | 허용 기준 | 결과 |
| --- | --- | --- | --- |
| V1 | 원형 step-index 정확한 벡터 HE11; 최대 Re 절대오차 6.29662878e-06, 마지막 메시 차이 1.26109825e-06 | 각각 ≤1e-5; 3개 메시와 3개 창, 두 독립 편광 | PASS_SCOPED |
| V4 | 복소 균질 원형 PEC TE11 해 최대 복소오차 1.4144407e-06; β²/companion Re 차이 0; 부분공간 상관 1 | 복소오차 ≤1e-5, pencil 비교 ≤1e-9; 전력 양수, 유한한 6성분, 잔차 ≤1e-7 | PASS_SCOPED |
| V5 | J1/Hankel-1 방사 해; 각 모드 Im 최대 상대오차 0.499925%, 전체 Im spread 0.770479%; Re 절대오차 2.43649929e-05 | Im 오차/spread 각각 ≤5%, Re ≤1e-4, 16개 메시/창/PML 조합에서 동일 부분공간 추적 | PASS_SCOPED |

원형 HE11/방사 22개 case 전체 최대 QEP 잔차 `2.14753e-14`, pencil 잔차 `2.16984e-14`, Gauss 잔차 `7.11633e-13`, 최소 모드 부분공간 상관 `0.999636633`. 각 case에서 두 모드의 실제 정규화 전력은 양수이며 약 1 W이다. 수치는 [원시 JSON](../universal_silica_v01/fiber_research_v1/benchmarks/step3_v085/step3_validation.json)에 보존했다. V5는 편광 쌍의 평균으로 오차를 줄이지 않고 **각 개별 Im 값**으로 판정한다.

## 시험 입력

- HE11: λ=1.55 µm, n_core=1.50, n_clad=1.44, a=1.3 µm, V=2.21330269530326, 정확한 neff=1.468523976355074. 계산창 6/7.2/8 µm, PEC 외곽, N2/P2. 메시 4/6/8 요소 파장 입력 및 곡률 144/192/240 동시 개선; 혼합 DOF 17,097/26,128/35,985. 무한 클래딩 해와 유한창 효과도 함께 검사한다.
- 복소 TE11: 원형 PEC R=3 µm, 균질 n=1.45+i1e-4. 정확한 neff=sqrt(n²−(j'11/(k0 R))²). 방사 손실이 없는 물질 흡수 시험이다.
- 방사 m=1: λ=1.55 µm, n_core=1.00, n_outside=1.45, a=3 µm. 정규 J1와 outgoing Hankel-1 접선 E/H 연속조건의 정확한 근 neff=0.981264795001076+i0.00468339885554407. 실제 공기의 상태 방정식/상용 HCF 재료 모델을 뜻하지 않는다. 메시 혼합 DOF 35,526/60,188/85,424. 창 6/6.6/7.2 µm, PML 두께 1.5/2/2.5 µm, 강도 3/4/5, 다항 차수 2/3/4를 검사했다.
- 위 굴절률·길이·흡수 수치는 **명시적 합성 시험 조건**이며 문헌 재료값으로 추정한 숫자가 아니다. 고손실 방사 시험을 극소 누설 HCF 검증으로 확장하지 않는다.

## 회귀 및 도구 검증

- Python: 기존 STEP 1·2 및 보존된 STEP 4/5/6 자산 테스트에 혼합 FEM/퇴화/복소 pencil/SBC/유한성/독립 기준 검사를 추가했다. 최종 실행 결과는 `step3_v085/python_tests.txt` 및 CI 기록에 남긴다.
- JS: 6성분 유한성/배열 일치, 모든 모드 검사, 입력 일치 보호, 기존 메시 입력 비교 테스트.
- 웹 실제 index: STEP 1·2/STEP 3 상단 바로가기, HCF/vector 표시모드에서도 접근, 실제 계산 두 예시, 모드 선택, 불일치 결과 거절, 입력 변경 후 이전 결과 제거, 390 px 화면 검사를 CI에 추가했다. 브라우저 결과는 CI 실행 후 아래 증거에 기록한다.
- 개별 `solve`는 항상 `DIAGNOSTICS_ONLY_INDEPENDENT_STUDY_REQUIRED`와 인증 false를 반환한다. 동봉 benchmark의 입력/검증 메타데이터를 일반 사용자 설계에 전이하지 않는다.

## 재현

```bash
cd COMSOL/universal_silica_v01/fiber_research_v1
python -m pip install -r requirements.txt
OPENBLAS_NUM_THREADS=1 OMP_NUM_THREADS=1 python -m unittest discover -s tests -v
OPENBLAS_NUM_THREADS=1 OMP_NUM_THREADS=1 python validate_step3.py --out out_step3
python cli.py solve benchmarks/step3_v085/guided_input.json --modes 2 --target-neff 1.468524 --out out_guided
python cli.py solve benchmarks/step3_v085/radiative_input.json --modes 2 --target-neff 0.981265 --out out_radiative
python cli.py study benchmarks/step3_v085/radiative_input.json --modes 2 --target-neff 0.981265 --out out_sensitivity
```

Colab에서도 같은 Python 명령을 실행할 수 있다. Linux Gmsh에 `libglu1-mesa`가 필요할 수 있다. 기존 Full-vector Colab notebook은 보존한다. 웹 상단 **STEP 3 벡터 FEM** → **실계산 예시 적용** 또는 동일 입력 적용 후 `mode_result.json 가져오기`. 웹에서 벡터 고유계는 계산하지 않는다.

## CI 증거

코드 업로드 후 실제 GitHub Actions 실행 URL·결과를 추가한다. 업로드 전 로컬 결과와 CI 실행 결과를 혼동하지 않는다.

다음 제안: 승인 후 고체 코어 PCF의 독립 벤치마크(V2/V6 정의·문헌 수치 포함). 이번 작업에서 해당 단계를 시작하지 않았다.
