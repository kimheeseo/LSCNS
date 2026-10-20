# LS Wave Optics Workbench v0.1

**[처음 사용자용 상세 가이드: 알 수 있는 것 → 설치 → 사용법 → 결과 해석](USER_GUIDE.md)** · [불러올 수 있는 예제 모델 3개](examples/README.md) · [Windows/Linux 최종 검증](FINAL_VERIFICATION.md)

회사 PC에서 실행하는 **Python FEM + 브라우저 UI**입니다. 모델 트리 → 기하·재료 → 실제 메시 → 연구 → 결과·오차 확인 순서로 사용합니다. COMSOL 소프트웨어나 COMSOL API를 사용하지 않습니다.

[설치용 소스 ZIP](https://github.com/kimheeseo/LSCNS/raw/refs/heads/main/COMSOL/workbench/LS_WaveOptics_Workbench_v0.1.zip) · [화면 미리보기](https://kimheeseo.github.io/LSCNS/COMSOL/workbench/) · [전체 소스](https://github.com/kimheeseo/LSCNS/tree/main/COMSOL/workbench)

GitHub Pages에서는 Python을 실행할 수 없으므로 **저장된 실제 실행 결과와 단면 미리보기**를 제공합니다. 입력값을 바꾸면 이전 결과를 현재 결과로 표시하지 않고 재계산 필요 상태로 전환합니다. 새 계산은 아래 로컬 Python 앱에서 실행하십시오.

## Windows에서 시작

1. Python 3.12를 설치합니다.
2. 위 **설치용 소스 ZIP**을 다운로드·압축 해제합니다. 필요한 M7/M8 해석기가 포함되어 있습니다. `workbench` 폴더만 단독으로 받으면 기존 해석기 경로가 없습니다.
3. `COMSOL/workbench/START_WINDOWS.bat`를 실행합니다. 첫 실행은 가상환경 생성과 공개 Python 의존성 설치를 수행합니다.
4. 브라우저가 열리면 `Python 엔진 연결됨`을 확인합니다. 기본 주소는 `http://127.0.0.1:8765/`입니다.
5. 오른쪽에서 입력값을 바꾸고 **모드 해석**을 누릅니다. 종료는 Python 실행 창에서 Ctrl+C입니다.

수동 설치 및 실행:

```powershell
cd LS_WaveOptics_Workbench_v0.1
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r COMSOL/workbench/requirements.txt
.\.venv\Scripts\python.exe COMSOL/workbench/launch.py
```

첫 설치에는 인터넷이 필요합니다. 의존성 설치 후에는 실제 계산과 데이터 처리가 이 PC 안에서 수행됩니다. 회사 PC의 Python·패키지 설치 정책에 따라 설치 가능 여부를 확인하십시오.

## macOS / Linux

```bash
python3.12 -m venv .venv
.venv/bin/python -m pip install -r COMSOL/workbench/requirements.txt
# Ubuntu의 Gmsh 공유 라이브러리:
sudo apt-get install -y libglu1-mesa libxft2 libxrender1 libxinerama1 libxcursor1
.venv/bin/python COMSOL/workbench/launch.py
```

`--port 8766`으로 포트를 바꾸거나 `--no-browser`로 자동 브라우저 열기를 생략할 수 있습니다. 계산 서버는 127.0.0.1에만 연결됩니다. 한 번에 한 작업을 실행합니다.

## 지원하는 해석

| 모델 | 실제 엔진 | 출력 |
|---|---|---|
| 원형 단일 코어 | M7 방사형 P1 FEM | neff, Aeff, 2차 모멘트 MFD, FEM Petermann-II MFD, 동일 메시 5점 미분 D, 실제 장 |
| 동일 2코어 MCF | M8 곡선 P2 FEM · 대칭 1/4 영역 | 짝/홀 neff, Δneff, κ, 결합 길이, 실제 슈퍼모드, 이상적 전력 전달 |
| 이상적 HCF 모세관 | M8 방사형 P1 FEM | neff, Aeff, kt², 일정 기체 굴절률의 D, 실제 장 |

- 단일 코어 및 MCF는 검증된 독립 기준의 적용 범위인 **단일 코어 V<2.405**를 사용합니다.
- MCF 결과는 제조 편차·굽힘·랜덤 누화가 없는 동일 직선 코어입니다.
- HCF는 J0/Dirichlet 모세관입니다. 실제 ARF/NANF의 유리벽·반공진·PML·구속 손실은 미구현입니다.
- 스칼라 장 ψ를 표시합니다. Ex/Ey/Ez 전기장이나 편광을 계산했다고 표시하지 않습니다.

## 화면과 조작

- **모델 탐색기:** 전역 정의·기하·재료·물리·메시·연구·결과로 이동합니다.
- **단면:** 입력값의 기하 미리보기입니다. 단일 코어 R은 계산 경계이며 실제 클래딩 반경과 다릅니다.
- **메시 생성:** 실제 해석에 사용하는 방사형 P1 또는 곡선 P2 메시를 표시합니다.
- **모드 해석:** 새 FEM 계산을 실행합니다. 웹 보간이나 저장값 조회로 대체하지 않습니다.
- **메시 수렴:** 현재 설정의 1/4·1/2·전체 요소 수(방사형), 또는 서로 다른 목표 h(MCF)를 비교합니다.
- **영역 수렴:** R/L, R/L+2 µm, R/L+4 µm를 비교합니다. MCF는 각 영역을 재메시합니다. HCF 외벽은 물리 경계이므로 영역 수렴 버튼을 비활성화합니다.
- **파장 스윕:** 각 파장에서 해석기를 다시 실행합니다. 단일 코어는 neff·Aeff·MFD를 계산하며, D는 개별 모드 해석에서 계산합니다.
- **확대/축소/맞춤:** 표시 크기를 조절합니다. 계산 메시를 변경하지 않습니다.
- **모델 저장/열기:** 버전이 있는 JSON 형식으로 입력값을 저장·복원합니다.
- **결과 내보내기:** JSON은 입력·결과·기준·오차·그래프를 함께, CSV는 수치 표, PNG는 현재 그래프를 저장합니다.

연구 결과 카드에는 **마지막 연구 조건**의 값이 표시됩니다. 전체 조건은 연구 표와 JSON을 확인하십시오. 실패는 로그에 표시하며, 이전 성공 결과를 새 결과처럼 바꾸지 않습니다.

## 오차와 COMSOL 비교

표시하는 오차는 동일 이상화 모델의 독립 Bessel/원통함수 기준 대비 **수치 오차**입니다. 제조사 제품·실측·COMSOL 결과와의 오차를 뜻하지 않습니다. 단일 코어의 FEM Petermann-II MFD 및 분산은 직접 계산값이며 UI에서 독립 오차를 표시하지 않습니다. 별도 파장 미분/메시 수렴을 확인하십시오.

COMSOL Wave Optics는 벡터 모드·전자기파·포트·PML 등의 기능을 제공합니다. 이번 UI가 이런 물리 기능을 새로 구현한 것은 아닙니다. 같은 방정식·재료·경계조건·메시 수렴 조건으로 맞춘 원형 스칼라 벤치마크부터 비교해야 합니다. **실제 COMSOL과의 교차 검증은 아직 수행하지 않았습니다.**

- [COMSOL 공식 Wave Optics 개요](https://doc.comsol.com/6.3/doc/com.comsol.help.woptics/woptics_ug_intro.4.03.html)
- [M7/M8 연구용 검증](https://github.com/kimheeseo/LSCNS/blob/main/COMSOL/research/README.md)
- [2코어 MCF 검증 보고서](https://github.com/kimheeseo/LSCNS/blob/main/COMSOL/waveoptics_m8/MCF_2CORE_VALIDATION_REPORT.md)
- [HCF 검증 보고서](https://github.com/kimheeseo/LSCNS/blob/main/COMSOL/waveoptics_m8/HCF_VALIDATION_REPORT.md)

SCUBA 125 예시는 반경 5.5 µm를 가정하고 MFD에 보정한 단순 등가 모델입니다. 실제 단면을 의미하지 않습니다. [Lightera 기술자료 ET07179 v3](https://datasheet.lightera.com/versao_et_pdf?id=a0ATT000002m1nm2AA&lang=en-us)의 MFD 공칭 12.3 µm, Aeff 공칭 125 µm² 중 MFD만 보정했고, 예측 Aeff는 약 118.56 µm²로 공칭값보다 약 5.15% 작습니다. 감쇠와 PMD는 이 모델로 예측하지 않습니다.

## 검증 재실행

```bash
python -m pip install pytest==9.1.1
python -m pytest COMSOL/workbench/qa/test_workbench.py -q
```

GitHub Actions는 수치/HTTP 테스트와 Chromium에서 실제 계산·파라미터 변경·모델 저장·CSV 다운로드·MCF 메시·HCF 수렴·파장 스윕·정적 미리보기·모바일 레이아웃을 확인합니다. 실행 결과와 스크린샷은 해당 Actions 실행의 `wave-optics-workbench-qa` 아티팩트로 저장합니다.


[변경 기록](CHANGELOG.md) · [UI 검증 보고서 및 실제 화면](VALIDATION_REPORT.md)

## 배포본과 최신 문서

설치 ZIP은 Windows/Linux에서 검증한 v0.1 파일 그대로 유지합니다. `SOURCE_MANIFEST.json`은 그 배포 시점의 파일 해시 기록이며, 이후 GitHub에 보완한 README·상세 가이드·예제 모델은 ZIP에 들어 있지 않습니다. 최신 설치·사용법과 재현 예제는 [USER_GUIDE.md](USER_GUIDE.md)를 확인하십시오. 이번 문서 보완은 계산 엔진과 UI 코드를 변경하지 않습니다.
