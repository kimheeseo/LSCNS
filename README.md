# LS Cable Digital Twin and Optical Tools

## 바로 실행

| 도구 | 링크 | 버전 및 상태 |
|---|---|---|
| **데이터센터 3D — Three.js 디지털 트윈** | [최신 3D 시뮬레이터 열기](https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html?engine=three) | **v4.8.0-rc.1** · WebGL2 / Three.js r186.1 · 릴리스 후보 |
| 기존 3D 렌더 | [Legacy 3D 열기](https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html?engine=legacy) | 기존 v4.7.0 계산·운영 코어 유지 |
| 원본 2D Gold Pixel Tour | [2D 탐방](https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_Campus.html) | 원본 보호 |
| 데이터센터 BOM 및 제품 툴 | [DataCenter 메인](https://kimheeseo.github.io/LSCNS/DCI/DataCenter/) | 기존 BOM 도구 보존 |

## 사용 설명서 · 개발 이력

- **[3D 시뮬레이터 사용 방법](DCI/DataCenter/README.md)** — 시작, PC/모바일 조작, 랙 서버 인출/문/카세트, 광·전력·냉각, MMR/Co-location, Troubleshooting.
- **[v4.7.0 → v4.8.0-rc.1 버전업 상세 이력 TXT](DCI/DataCenter/LS3D_VERSION_UPGRADE_HISTORY.txt)** — 1~20단계의 변경 내용, 모듈 캐시 정책, 버전 승격 및 롤백 방법.
- **[20단계 최종 릴리스 보고서](DCI/DataCenter/architecture/STEP20_RELEASE_REPORT.md)** — 범위, 보호된 기존 파일, 테스트 근거, 실제 브라우저 확인 필요 항목.
- [20단계 개발 로드맵](DCI/DataCenter/LS3D_V4_20_TASK_ROADMAP.md) · [개발/검증 진행 현황](DCI/DataCenter/LS3D_PROGRESS_STATUS.md)
- [STEP 19 GitHub Actions 브라우저 QA](https://github.com/kimheeseo/LSCNS/actions/workflows/ls3d-step19-quality.yml)

### v4.8.0-rc.1 릴리스 후보의 검증 상태

20단계 구현 및 문서 정리를 GitHub main에 반영했습니다. Three.js/랙/MMR/고객별 배선/모바일 관련 JavaScript 모의 회귀 테스트 22/22는 19단계 재검증에서 PASS였습니다. **최신 빌드의 독립 Chromium/WebGL2 실화면·모바일·FPS 테스트와 GitHub Pages 배포 결과는 아직 확인되지 않았습니다.** 따라서 v4.8.0 정식 릴리스가 아니라 rc.1로 표기합니다. GitHub Actions 검증 결과를 확인한 후 정식 승격 여부를 판단하세요.

### 이전 Phase 7 v4.7.0 기능

기존 시뮬레이터는 공유 설계 패킷 lsdc-twin-bom/1.0의 정보를 요약하고 구성 호환성 및 근거 부족을 표시하며, 회사 서버실을 모델링된 공간 범위에 한정하는 기능을 포함합니다. 이번 Three.js 업데이트는 기존 데이터·계산/BOM 호환성을 유지하면서 시각 탐방 레이어를 추가한 것입니다.
