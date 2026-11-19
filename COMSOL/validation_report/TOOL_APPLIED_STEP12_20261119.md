# STEP 1·2 웹 도구 적용 — 2026-11-19

사용자 요청에 따라 이미 main에 있던 v0.8.4 기능을 실제 도구에서 찾을 수 있도록 적용 경로를 보완했다.

- 상단 바로가기에 STEP 1·2 설계·메시 진입 링크를 추가했다.
- #research-fiber-v1 진입 시 HCF/vector 선택 화면에서도 연구 패널이 숨겨지지 않도록 했다.
- 앱 헤더를 v0.8.4로 표시하고, 기존 scalar 계산 코어가 v0.7.1임을 패널에서 별도로 명시했다.
- 기존 scalar 수식, CAD, JSON schema 및 Colab 결과 경로는 변경하지 않았다.
- 실제 Python mesh_preview.json 가져오기, 설계 편집 시 결과 무효화 및 모바일 패널은 기존 CI 통과 결과를 유지한다. 이번 진입 경로는 브라우저 회귀 검사에 추가한다.

[도구 진입](https://kimheeseo.github.io/LSCNS/COMSOL/universal_silica_v01/web/?v=20261119-v084-applied#research-fiber-v1)

이후 STEP 3는 별도 v0.8.5 작업으로 진행한다. 이 문서는 모드/손실 정확도 인증을 뜻하지 않는다.

## 후속 실제 화면 확인 — STEP 3 v0.8.5

실제 전체 index 회귀에서 기존 HCF/vector `:not(...)` 숨김 규칙의 높은 CSS 우선순위를 확인했고 commit `5a72113fa4631f954e4a4f2dfb85805a8af8f7b9`에서 연구 패널을 해당 규칙에서 제외했다. 연구 링크 활성 시 표시하고 다른 HCF/vector 화면에서는 기존 전용 레이아웃을 유지한다. [최종 종합 CI](https://github.com/kimheeseo/LSCNS/actions/runs/44455761996)와 [실제 화면 회귀](https://github.com/kimheeseo/LSCNS/actions/runs/44455761971)가 모두 success다. 기존 모델 마법사는 정상 Cancel 버튼으로 닫아 검사한다. STEP 3의 코드·검증표·한계는 [검증 보고서](STEP3_V085_VERIFICATION_20261119.md)에 정리했다.
