# Simple Preform Data Analyzer

`alls.xlsx`에 저장된 광섬유 프리폼 품질 이력 데이터를 **그룹별로 분류하고, 품질 Report·오류 목록·월별 추세 그래프를 생성하는 Windows용 분석 도구**입니다.

## 준비

- Python 3.x
- 입력 파일: `alls.xlsx`
- Python 패키지

```bash
pip install openpyxl matplotlib pyinstaller
```

## 사용 방법

```bash
python app.py
```

1. **alls.xlsx 파일 선택**  
   입력 파일을 선택합니다.
2. **1단계: 그룹별 xlsx 생성**  
   - E열 값의 앞 두 글자 기준으로 데이터를 분류합니다.
   - 각 그룹의 데이터를 다시 D열 값별 시트로 나누어 `AL.xlsx`, `LE.xlsx` 등 그룹 파일을 생성합니다.
3. **2단계: 선택 파일 Report 생성**  
   - 분석할 그룹을 선택합니다.
   - 시트별 평균, 단선 횟수와 주요 품질 항목을 정리한 `[그룹]_report.xlsx`를 생성합니다.
   - 기준을 벗어난 행은 Report에서 빨간색으로 표시하고 `error.txt`에 기록합니다.
4. **3단계: 월별 결과·그래프 생성**  
   - `hdate`를 기준으로 월별 시트를 추가합니다.
   - 1550 nm 감쇠 I/E·O/E, 1310 nm MFD I/E·O/E의 월 평균 그래프 4개를 PNG로 저장합니다.

## 입력 데이터 전제

이 프로그램은 `alls.xlsx`의 첫 번째 시트를 원본 데이터로 사용하며, 다음 열 위치와 형식을 전제로 합니다.

| 열 | 용도 |
|---|---|
| A열 | hdate (예: `2025.05`) |
| D열 | 그룹 내 세부 시트 분류값 |
| E열 | 그룹 분류값 및 spool 번호 |
| F열 이후 | 광학·기하·분산·굽힘 등 품질 측정값 |

E열의 끝 문자가 숫자이고 0이 아니면 Report 생성 시 해당 행을 제외합니다. spool 번호의 뒤에서 5·6·7번째 문자에 포함된 A/B/C… 표기를 사용해 단선 횟수를 계산합니다.

## 생성 파일

| 파일 | 내용 |
|---|---|
| `[그룹].xlsx` | E열 앞 두 글자로 나눈 그룹 파일. D열 값별 시트 포함 |
| `[그룹]_report.xlsx` | 시트별 평균, 단선 횟수, 품질 이상 표시, 월별 시트 |
| `error.txt` | 품질 기준을 벗어난 spool 목록과 발생 시각 |
| `[그룹]_*.png` | 월별 감쇠·MFD 평균 추세 그래프 4개 |
| `run_log.txt` | 실행 진행 상황과 처리 로그 |

## 자동 품질 판정 기준

| 항목 | 허용 범위 / 조건 |
|---|---|
| MFD 1310 nm I/E, O/E | 8.8 ~ 9.2 |
| Cutoff 2 m − 22 m | 음수이면 이상 |
| Clad Diameter I/E, O/E | 124.3 ~ 125.7 |
| Dispersion slope at ZDW | 0.073 ~ 0.090 |

판정 기준이나 Report 항목은 `app.py`의 `mapping` 및 `check_range()` 부분에서 수정할 수 있습니다.

## 주의사항

- **1단계는 입력 `alls.xlsx`에 그룹 시트를 추가하고 0 값을 빈칸으로 바꿉니다. 실행 전 원본을 반드시 복사해 두십시오.**
- 데이터 열 순서가 바뀌면 Report 값과 품질 판정이 잘못될 수 있습니다.
- 결과 파일과 로그는 입력 파일과 같은 폴더에 저장됩니다.

## 실행 파일 만들기

Windows 실행 파일은 다음 명령으로 생성합니다.

```bash
pyinstaller --onefile --windowed --name lsexcelpreform app.py
```

생성 파일: `dist/lsexcelpreform.exe`

## 참고

- [사용 시작 영상](https://youtu.be/Gsl1WYRC1nI?si=XFRbgvzFufa7gd1c)
- [배포 파일 다운로드](https://drive.google.com/drive/folders/1G7vG1GKtajdI2Q21o0q0UUIVP4btDATo?usp=drive_link)
