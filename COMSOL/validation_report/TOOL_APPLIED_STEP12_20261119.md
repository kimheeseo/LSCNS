# STEP 1·2 웹 도구 적용 — 2026-11-19

사용자 요청에 따라 이미 main에 있던 v0.8.4 기능을 실제 도구에서 찾을 수 있도록 적용 경로를 보완했다.

- 상단 바로가기에 STEP 1·2 설계·메시 진입 링크를 추가했다.
- #research-fiber-v1 진입 시 HCF/vector 선택 화면에서도 연구 패널이 숨겨지지 않도록 했다.
- 앱 헤더를 v0.8.4로 표시하고, 기존 scalar 계산 코어가 v0.7.1임을 패널에서 별도로 명시했다.
- 기존 scalar 수식, CAD, JSON schema 및 Colab 결과 경로는 변경하지 않았다.
- 실제 Python mesh_preview.json 가져오기, 설계 편집 시 결과 무효화 및 모바일 패널은 기존 CI 통과 결과를 유지한다. 이번 진입 경로는 브라우저 회귀 검사에 추가한다.

[도구 진입](https://kimheeseo.github.io/LSCNS/COMSOL/universal_silica_v01/web/?v=20261119-v084-applied#research-fiber-v1)

이후 STEP 3는 별도 v0.8.5 작업으로 진행한다. 이 문서는 모드/손실 정확도 인증을 뜻하지 않는다.
