# Workbench에서 바로 불러오는 입력 예제

각 링크에서 **Raw / Download raw file**로 JSON을 저장한 뒤, 로컬 Workbench의 **모델 → 열기**로 불러오고 **모드 해석**을 누릅니다. 모델 JSON은 입력만 복원하며, 저장된 결과를 불러오는 파일은 아닙니다.

| 모델 파일 | 주요 조건 | 비교할 실제 실행 기록 |
|---|---|---|
| [single_model.json](single_model.json) | 1550 nm, a=4.1 µm, Δn=0.005, R=40 µm, 요소 160/640 | [단일 코어](../sample_result.json) |
| [mcf_model.json](mcf_model.json) | 1550 nm, a=4.1 µm, Δn=0.005, p=16 µm, **L=28 µm**, h=0.7 µm | [2코어 MCF](../evidence/mcf_smoke.json) |
| [hcf_model.json](hcf_model.json) | 1550 nm, 중공 반경 15 µm, n_air=1.00027, 요소 320 | [이상적 모세관](../evidence/hcf_smoke.json) |

MCF 예제는 화면 프리셋의 기본 L=32 µm와 다릅니다. HCF 파일의 `outer_um` 등 숨겨진 다른 모델용 입력은 모델 JSON 형식 호환성을 위해 들어 있으며, HCF 해석은 `hcf_radius_um`을 사용합니다.

[설치·사용·결과 해석 가이드](../USER_GUIDE.md#examples)에서 출력값과 해석 범위를 확인하십시오. 예제는 현재 검증된 v0.1 설치 ZIP에서 사용할 수 있습니다. 이 JSON은 입력 형식 버전 1이며, 계산 엔진 버전과 별도입니다.
