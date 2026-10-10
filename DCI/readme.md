# AI Data Center DCI

AI 데이터센터의 **Compute / Network / Optical / Rack / Power / Cooling / BOM**을 설계·검토하기 위한 웹 도구 모음입니다.

## 바로 실행

아래 페이지는 GitHub 저장소 화면이 아니라 **GitHub Pages 배포 사이트**에서 실행합니다. 별도 설치 없이 데스크톱 또는 모바일 브라우저에서 열 수 있습니다.

| 도구 | 실행 페이지 | 주요 내용 |
|---|---|---|
| DataCenter Tool | [▶ 데이터센터 설계·BOM 도구](https://kimheeseo.github.io/LSCNS/DCI/DataCenter/) | Compute, Network, Optical, Rack, Power, Cooling 설계 및 BOM |
| 3D Digital Twin | [▶ LS Datacenter Campus 3D](https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html) | 캠퍼스 7개 구역, 설비 선택, 랙 점검, 작업자 이동·업무 시점 |
| 2D Gold Pixel Tour | [▶ LS Datacenter Campus 2D](https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_Campus.html) | 픽셀아트 캠퍼스 탐험 및 설비 조사 |
| CPO Supply Chain | [▶ CPO 공급망 도구](https://kimheeseo.github.io/cpo-supply-chain/) | CPO 공급망 및 업체 정보 |

## 3D Digital Twin 사용 방법

자세한 화면 조작, 작업자 시점, 장애 시나리오, 그래프·로그 및 설정 방법은 [3D Digital Twin 상세 사용 안내(TXT)](./DataCenter/LS_Datacenter_3D_사용법.txt)를 참고하세요. 계산식·입력값·장애 대응·ToR/광 링크·JSON/BOM 연동까지 포함한 문서는 [3D Digital Twin 운영·계산 가이드북(Word)](./DataCenter/LS_Datacenter_3D_Guidebook_v4.5.3.docx)에서 확인할 수 있습니다.

1. [3D Digital Twin 페이지](https://kimheeseo.github.io/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html)를 엽니다.
2. 왼쪽 메뉴에서 Campus, Utility Yard, Cooling Plant, Gray Zone, Data Hall, Network / MMR, NOC / Safety 중 구역을 선택합니다.
3. 장비를 클릭하면 오른쪽 설비 검사 패널에서 상세 정보를 확인할 수 있습니다.
4. 상단의 작업자 카드 중 한 명을 선택하면 해당 직원의 근무 장면을 따라갑니다. 화면 안의 시점 버튼으로 근접 눈높이 보기와 넓게 따라보기를 전환할 수 있습니다.
5. 마우스 드래그로 회전, Shift+드래그로 이동, 휠로 확대·축소합니다. 모바일에서는 화면 터치 조작을 사용합니다.

랙을 선택하면 전면·후면, 도어, 서버 인출, 광 카세트와 케이블 연결을 확인할 수 있습니다. 모델은 교육·설계 검토용 개념 모델이며 실측 CAD/BIM이나 실시간 시설 데이터가 아닙니다.

## GitHub Pages에서 실행·배포

- 배포 주소: [https://kimheeseo.github.io/LSCNS/DCI/DataCenter/](https://kimheeseo.github.io/LSCNS/DCI/DataCenter/)
- 공개 소스: [DCI/DataCenter](https://github.com/kimheeseo/LSCNS/tree/main/DCI/DataCenter)
- 배포 상태: [GitHub Actions](https://github.com/kimheeseo/LSCNS/actions)

변경 사항은 `main` 브랜치에 반영한 뒤 GitHub Actions의 **pages build and deployment** 작업이 성공하면 GitHub Pages에 게시됩니다. 소스 HTML 파일을 GitHub에서 직접 실행하는 방식이 아니라, 배포된 Pages URL을 브라우저로 여는 방식입니다.

## 로컬에서 미리보기

저장소를 내려받은 뒤 저장소 최상위 폴더에서 정적 웹 서버를 실행합니다.

```bash
python -m http.server 8000
```

브라우저에서 다음 주소를 엽니다.

- DataCenter Tool: http://localhost:8000/DCI/DataCenter/
- 3D Digital Twin: http://localhost:8000/DCI/DataCenter/LS_Datacenter_3D.html
- 2D Gold Pixel Tour: http://localhost:8000/DCI/DataCenter/LS_Datacenter_Campus.html

## 주요 기능

- GPU / Rack / Leaf-Spine-Core 네트워크 설계
- Optical / Fiber / Connector / Transceiver BOM 산정
- Power / Cooling sizing 및 2D/3D 시각화
- Supply Chain / Product Mapping
- Excel / CSV / Design JSON export
- 3D 디지털 트윈에서 7개 구역 탐색, 랙 내부 점검, 작업자 이동 확인

## Repository

- `DataCenter/` — 공개 UI, 시각화, 정적 runtime assets
- `current_version/` — 현재 개발 버전
- `updates/` — 날짜별 업데이트 이력
- CPO Supply Chain — 별도 repository: [kimheeseo/cpo-supply-chain](https://github.com/kimheeseo/cpo-supply-chain)

Validation / reference case는 private `kimheeseo/others/DCI/`에서 관리합니다.

- [현재 개발 버전](./current_version/README.md)
- [업데이트 내역](./updates/README.md)

## DataCenter Tool backend

DataCenter 화면은 LSCNS GitHub Pages에서 실행됩니다. 핵심 계산·BOM 제품 매칭은 private backend에서 처리되므로, backend 연결 상태에 따라 해당 기능의 사용 가능 여부가 달라질 수 있습니다. 2D Gold Pixel Tour와 3D Digital Twin은 공개 정적 페이지로 실행됩니다.

## Source protection

- Private source: `kimheeseo/others/DCI/DataCenter/`
- Production calculation API: Railway private service
- Public Pages에는 핵심 `dc-engine.min.js`와 `catalog-design-match.js`를 두지 않습니다.
- 브라우저는 입력 JSON을 backend로 보내고 결과 JSON만 받아 표시합니다.
- 제품 카탈로그 자체는 제조사 공개 자료이므로 public reference data로 유지할 수 있습니다.

> 주의: 과거 public Git commit에는 이전 소스가 남아 있을 수 있습니다. 완전한 과거 이력 제거는 별도의 history rewrite 또는 runtime-only public repository 재생성이 필요합니다.
