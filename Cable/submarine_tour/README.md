# 한국–일본 데이터센터 광신호 투어

[투어 실행](https://kimheeseo.github.io/LSCNS/Cable/submarine_tour/)

한국 데이터센터의 서버에서 출발해 육상 광망·부산 CLS·해안 BMH·해저케이블·후쿠오카 CLS를 거쳐 일본 데이터센터 서버에 도착하는 한국어 교육용 3D 웹 앱입니다. 기존 Drawing Tower 앱과 독립적으로 실행됩니다. Three.js r128을 cdnjs에서 불러오는 단일 `index.html`이며, 장비·표지·단면은 기본 지오메트리와 Canvas로 생성하고, 동물의 곡면은 Blender에서 제작해 HTML 안에 내장합니다. 별도의 모델·이미지 파일을 다운로드하지 않습니다.

## 포설선 저장 탱크·한국 공장 외형 v2.7 · 2026-11-27

- 첫 번째 참고 사진처럼 포설선 갑판에 **원형 케이블 저장 탱크, 여러 겹의 케이블, 중앙 허브, 황색 난간, 아치형 이송 구조물**을 추가했습니다. 선폭과 갑판도 넓혀 탱크가 선체 안에 배치되도록 했습니다. 기존 보조 드럼·선미 시브·ROV 연결은 유지합니다.
- **장비 그림 → 저장 탱크 → 3D로 보기**로 위에서 확대 관찰합니다. 탱크를 직접 클릭해도 안내를 열며, **장면 재생** 시 저장부가 천천히 회전하고 일시정지하면 멈춥니다. 적층·회전은 시각화이며 실제 케이블 용량·장력·권취 제어를 계산하지 않습니다.
- 두 번째 사진의 외형을 참고해 한국 데이터센터 뒤편에 **LS전선 동해공장 참고 모델**을 배치했습니다. 청색 지붕의 공장동, 저층 사무동, 세로 리브와 비대칭 상부가 있는 높은 수직 타워, LS 표지, 부지 도로와 드럼을 표현했습니다. 실제 공장 치수·배치 또는 동해공장과 부산 데이터센터의 지리적 인접 관계를 재현하지 않습니다.
- **장비 그림 → 동해공장 → 3D로 보기**로 확대하거나 공장 모델을 직접 클릭합니다. 데이터센터에서 출발하는 기존 위치·속도·급전 계산은 유지하고, 한국 육상 구간에서 공장을 잠시 바라보는 시선만 추가했습니다. 공장은 통신 광신호의 경유 장비가 아닙니다.
- **48개 검사 통과**: 저장 탱크의 정면 선택 광선, 재생·일시정지와 진행률·전기 수치 보존, 한국 측 공장과 데이터센터의 별도 배치 및 관찰 중 경로 보존을 포함합니다. 브라우저 호환 화면에서 실제 적층 케이블·황색 구조물과 공장 지붕·수직 타워를 확인했습니다. 큰 갑판·바닥 면이 세부 모델을 덮지 않도록 호환 화면의 그리기 순서를 보정했습니다.

## 데이터센터 간 광신호·해안 원형 BMH v2.6 · 2026-11-27

- 전체 투어를 **한국 데이터센터 → 한국 육상 광망 → 부산 CLS → 부산 BMH → 해저케이블 → 후쿠오카 BMH → 후쿠오카 CLS → 일본 육상 광망 → 일본 데이터센터**로 확장했습니다. 두 데이터센터에는 서버 랙과 엣지 라우터를 배치합니다. 마지막은 일본 서버의 수신 완료 화면입니다.
- 처음 구성도에 양쪽 데이터센터와 육상망을 추가했습니다. 기본 **데이터센터 광신호 투어**는 비급전 구성이고, **광신호 + PFE 급전 체험**은 해저 급전 설명을 함께 표시합니다. 상단 **광신호 1인칭 투어**는 현재 선택한 구성을 유지합니다.
- BMH는 제공된 해안 사진의 형태를 참고해 **모래밭의 원형 콘크리트 입구·열린 녹슨 금속 덮개·바위·마른 풀**로 재구성했습니다. 입구 아래에는 사다리와 지하 접속 공간을 둡니다. 사진 파일 자체는 앱에 삽입하지 않습니다. 실제 BMH는 형태가 다양하며 이 모델은 치수·설계·작업 절차를 재현하지 않습니다.
- 부산에서는 원형 입구를 통해 아래로 내려가며, 일본 측에서는 지하 접속 공간에서 입구로 상승한 뒤 CLS로 이동합니다. **장비 그림 → 해안 BMH / 데이터센터 → 부산 측 / 후쿠오카 측 → 3D로 보기**로 개별 관찰할 수 있습니다. 기존 CLS 문 출입과 SLTE·ODF·PFE, 해양 생물·작업 장면은 유지합니다.
- 전체 카메라 진행률 `journey`(기본 140초)와 기존 해저 회로 진행률 `u`를 분리했습니다. `journey=0.18…0.82`만 `u=0…1`로 변환합니다. 양쪽 육상 구간을 해저 설계 길이에 합산하지 않으며, PFE의 금색 도체·전류는 데이터센터까지 연장하지 않습니다. 청록색 육상 광경로는 별도 생성합니다. 시간·축척과 점은 교육용 표현으로 전파 지연·실제 패킷 시뮬레이션이 아닙니다.
- 46개 검사 통과: 데이터센터 출발·도착, 육상/해저 진행률 분리, 새 구간의 문 통과, 원형 입구의 개방 광선, 부산 하강·일본 상승, 선택·복귀 시 전체 진행률 보존을 포함합니다. 브라우저 호환 화면에서 서버 출발·수신과 원형 BMH 외형을 확인했습니다. 호환 렌더러는 원거리 관찰 시 지하 면 대신 어두운 입구 프록시를 사용하고 입구 접근 시 내부를 표시합니다.

## CLS 실내 SLTE·ODF·PFE v2.5 · 2026-11-27

- 참고 그림처럼 부산·후쿠오카 양쪽 CLS의 실내를 **SLTE / ODF / PFE / 전원반** 순서로 나란히 배치했습니다. 비급전 모드에서는 기존과 같이 PFE와 DC 급전 경로를 숨깁니다.
- ODF에는 광접속 어댑터·패치 패널·여장 루프를 표현했습니다. 구성도와 장비 그림에 **ODF · 광분배** 버튼을 추가했고, 3D의 ODF 랙을 직접 눌러도 역할 설명을 엽니다. 부산/후쿠오카를 선택해 각 장비를 확대할 수 있습니다.
- 청록색 **육상망–SLTE–ODF–광섬유** 접속과 금색 **PFE–금속 도체** 급전을 별도로 연결합니다. 광신호의 양방향 점과 DC 방향 표시는 재생·일시정지에 연동합니다. 장비 연결은 배선 역할을 보여주는 개념 표현이며, 실제 포트 수·접속 손실·장비 내부 회로는 모델링하지 않습니다.
- CLS 출입 체험의 실내 관찰 구간은 SLTE·ODF·PFE 순서로 시선을 이동합니다. 출입구와 통로를 유지하면서 랙을 뒤쪽에 배치해 카메라가 장비 본체를 관통하지 않도록 했습니다.
- 41개 검사 통과: 양쪽 CLS의 세 장비 정면 선택 광선, 지역 선택, ODF 관찰 시 모드·전류·진행률 보존, 광접속과 급전의 연결 분리, 새 배치에서의 출입 동선 등을 포함합니다. 브라우저 호환 화면에서 세 랙의 배치와 ODF 직접 클릭 및 후쿠오카 확대 화면을 확인했습니다.

## CLS 문 출입 v2.4 · 2026-11-27

- 자동 1인칭 투어는 부산 CLS 실내에서 문을 열고 퇴실한 뒤 기존 케이블 경로로 이어지며, 후쿠오카에서는 출입구에 접근해 문을 열고 실내 장비 앞으로 들어갑니다. 카메라가 문을 통과할 때 문은 완전히 열린 상태이며 통과 후 닫힙니다.
- **장비 그림 → 육양국 CLS → 부산 CLS / 후쿠오카 CLS → 문 열고 출입 체험**으로 선택한 육양국의 입장·장비 관찰·퇴실을 별도로 체험합니다. 문 자체를 클릭해도 해당 CLS 출입 체험을 시작합니다. 기본 속도에서 약 14초이며, 재생·일시정지·속도·이전/다음 구간·출입 진행 슬라이더를 사용할 수 있습니다.
- 출입 체험 중 케이블 진행 위치와 전기 계산값은 유지합니다. **투어로 돌아가기**로 원래 경로에 복귀합니다. 두 CLS는 장비 관찰을 위해 앞면과 지붕 일부를 개방한 교육용 모형입니다.
- 실제 r128 지오메트리로 닫힌 문과 열린 통로의 광선 교차, 자동 투어의 통과 시 문 열림, 두 CLS의 입장·실내 관찰·퇴실, 일시정지·재시작·복귀 및 탐색 중 전기 진행률 보존을 검사했습니다. 브라우저 호환 화면에서 닫힌 문·열린 출입구·실내 시점을 확인했습니다.

## SLTE·PFE 분류 v2.3 · 2026-11-27

| 분류 | 전달 경로 | 역할 | 표시 |
| --- | --- | --- | --- |
| SLTE · 광단국 장비 | 광섬유 | 통신 광신호 송수신·파장 신호 처리 | 청록색 |
| PFE · 급전 장비 | 금속 도체 | 해저 리피터용 정전류 DC 전원 공급 | 금색 |
| CLS · 육양국 | 육상 시설 | 이 예시에서 SLTE·PFE 등 장비를 배치하는 시설 | 시설 그림 |

- **장비 그림**과 구성도 카드에 **SLTE · 광신호 / PFE · DC급전**을 별도 항목으로 표시합니다. PFE 전용 그림과 두 장비의 역할 구분표를 추가했습니다.
- PFE 랙이 CLS 설명으로 연결되던 클릭 분류를 수정했습니다. 부산·후쿠오카의 SLTE 2개와 PFE 2개는 각각 독립 장비로 등록합니다. 3D 확대 시 건물 간판에 가리지 않도록 랙 정면 시점을 사용합니다.
- 비급전 모드에서 PFE 그림을 여는 것만으로 전류나 구성이 바뀌지 않습니다. **급전형으로 전환해 3D 보기**를 누를 때 급전형 예시로 전환합니다. 비급전으로 되돌리면 PFE 모델과 급전 경로를 숨깁니다.
- 브라우저에서 PFE·SLTE 랙을 직접 클릭해 각 장비의 설명이 열리는 것을 확인했습니다. 검사에는 등록 분류, 비급전 그림 열기 시 전류 보존, 명시적인 3D 전환 및 복원, 확대 시점의 중앙 선택 광선이 PFE 랙에 도달하는 조건을 포함합니다.

## 모바일 닫기·잠수부 v2.2 · 2026-11-27

- **초기 구성도 닫기**를 Three.js와 분리한 작은 네이티브 UI 스크립트가 먼저 실행합니다. CDN 다운로드와 3D 시작을 비동기로 처리하며, 3D가 아직 준비되지 않았거나 다운로드가 실패해도 닫기·구성도 재열기가 동작합니다. 로딩 중 사용자가 닫은 상태는 3D 초기화 후에도 유지합니다.
- 모바일 닫기는 줄바꿈 없는 72×44 CSS px 이상의 영역을 확보했습니다. 구성도 머리말을 스크롤 상단에 고정하고 닫힌 오버레이의 숨김 스타일을 명시했습니다.
- 부산 연안에 **잠수부의 금속 보호 프레임 용접 보수** 장면을 추가했습니다. 잠수복·헬멧·마스크·탱크·팔·토치와 토치의 빛·불꽃·기포를 표현합니다. 광섬유 케이블 자체의 수중 용접이나 실제 시공 절차를 뜻하지 않습니다.
- **장비 그림 → 잠수부 → 3D로 보기 → 장면 재생**으로 작업을 확대 관찰할 수 있습니다. 이때 투어 진행 위치와 전기 수치를 유지합니다. 1인칭 투어에서도 연안 잠수부 구간을 지납니다.
- 브라우저의 390×780 CSS px 프레임에서 정상 3D와 3D 로더 미실행 상태 모두 닫기/재열기를 확인했고, 버튼 크기 72×44 px와 한 줄 표시를 확인했습니다. 실제 Android 앱 내 브라우저에서 직접 검증한 것은 아닙니다.

## 장비 안내 v2.1 · 2026-11-27

- **장비 그림** 또는 구성도의 그림 카드를 눌러 **육양국(CLS) / 광단국 장비(SLTE) / 포설선 / ROV**를 선택합니다. 글을 길게 나열하는 대신 네 가지 벡터 안내 그림과 짧은 역할 설명을 제공합니다.
- **3D로 보기**는 선택한 장비로 카메라를 이동합니다. 3D의 건물·랙·선박·ROV를 클릭해도 해당 그림과 설명이 열립니다. 드래그와 두 손가락 조작은 클릭 선택으로 처리하지 않습니다.
- 부산·후쿠오카 건물을 CLS로 표시합니다. SLTE는 통신 광신호, PFE는 DC 급전, CLS는 이들을 배치하는 육상 시설로 구분합니다. 그림은 급전형 장비 배치의 개념도이며 실제 건축·배치도가 아닙니다.
- 포설선에는 선교·케이블 드럼·선미 시브·이송 장치·A 프레임을, ROV에는 부력재·추진기·카메라·조명·작업팔을 표현했습니다. 투어에 포설선·ROV를 관찰하는 구간을 추가했습니다.
- 보라색 **엄빌리컬**은 선박과 ROV 사이의 전원·제어 연결입니다. PFE의 해저 통신 케이블 급전 회로에 접속하지 않습니다. 선박의 드럼 회전, 작은 상하 동요와 ROV의 추진기·작업팔 동작은 재생/일시정지에 연동됩니다. 선박은 정점 유지 상태의 점검 장면으로 표시하며 실제 포설·매설·장력 계산은 하지 않습니다.
- 장비 그림을 열면 투어가 일시정지합니다. 장비 확대 관찰 중 전기적 수치와 진행 거리는 기존 투어 위치를 유지하고, 수심은 관찰 장비의 위치를 표시합니다. **1인칭**이나 진행 슬라이더로 투어 시점에 돌아갈 수 있습니다.

## 그래픽 v2.0 · 2026-11-27

- Blender 4.3.2의 Subdivision Surface로 상어 몸체·꼬리·등지느러미·가슴지느러미와 거북이 등껍질·물갈퀴를 제작했습니다. 미리 삼각분할한 3,172개 정점과 법선·UV를 HTML에 내장합니다.
- 금속의 거칠기·반사, 모래 입자·잔물결, 거북이 등껍질·피부 무늬를 Canvas 텍스처로 표현합니다. PFE/SLTE 표지와 랙 상태등, 받침대·실내 바닥도 개선했습니다.
- WebGL에서는 ACES 톤매핑, 환경 반사, PCF 부드러운 그림자, 수면 파도·법선 변화, 해저 집광 무늬와 수중 광선을 사용합니다. 깊은 바다의 색감과 거리 안개, 부드러운 화면 가장자리 효과를 적용했습니다.
- 상단 **그래픽 선명 / 그래픽 가벼움** 버튼으로 화질을 전환합니다. 선명은 최대 DPR 2와 2048 그림자, 가벼움은 최대 DPR 1로 그림자·광선을 줄입니다. 동선과 전기 계산은 동일합니다.
- WebGL 미지원 시 사용하는 Canvas 호환 화면도 뒷면 제거, 거리 안개, 표지의 원근 표현, 텍스처 샘플링, 면 경계 완화로 개선했습니다. 정지한 화면은 재계산하지 않습니다.

## 사용 방법

1. 첫 구성도에서 한국 데이터센터부터 일본 데이터센터까지의 전체 연결을 확인합니다.
2. **데이터센터 광신호 투어**로 시작합니다. 해저 급전도 보려면 **광신호 + PFE 급전 체험**을 선택합니다. 상단 **광신호 1인칭 투어**는 현재 구성을 유지합니다.
3. 재생·일시정지·처음으로·이전/다음 구간·전체 진행 슬라이더·속도(0.25–4배)를 사용합니다. 전체 진행률 옆의 해저 km는 별도 표시하며 육상 거리는 포함하지 않습니다.
4. **장비 그림**에서 해안 BMH·데이터센터·CLS·SLTE·ODF·PFE 등을 선택하고 양쪽 지역을 고른 뒤 **3D로 보기**로 관찰합니다. CLS는 **문 열고 출입 체험**도 제공합니다.
5. 장면 드래그는 시선 이동, 전체 보기에서는 회전입니다. 휠·두 손가락은 확대, **광신호 표시**는 이동하는 광신호 표시를 전환합니다. **설계 예시**는 기존 해저 DC 회로의 가정값을 조절합니다.

## 전기 모델

급전형은 양단 PFE가 정전류를 유지하는 단일 금속 도체 직렬 회로와 이상적인 해수·접지 귀로를 가정합니다. 후쿠오카 PFE도 급전원이며, 수전 장치로 취급하지 않습니다. 각 PFE는 전체 요구 전압을 절반씩 분담합니다.

- 도체 전압강하: `Vc = I × Rdc × L`
- 리피터 전압강하 합계: `Vr = N × Vrep`
- 총 요구 전압: `Vtotal = Vc + Vr`
- 대지 기준 단자 전위: 부산 `+Vtotal/2`, 후쿠오카 `−Vtotal/2`
- 경로 u 지점의 도체 전위: `V(u) = Vtotal/2 − I × Rdc × L × u − Npassed(u) × Vrep`
- 도체 손실: `Pcable = I² × Rdc × L`
- 리피터 소비전력: `Prep = I × N × Vrep`
- 전체 전력: `Ptotal = I × Vtotal = Pcable + Prep`

기본 예시는 `L=220 km`, `I=1 A`, `Rdc=1 Ω/km`, `N=2`, `Vrep=50 V`입니다. 요구 전압은 320 V, 양단 분담은 +160/−160 V, 도체 손실은 220 W, 리피터 소비전력은 100 W입니다. 이는 계산 흐름을 보여주는 가정이며 실제 설비 정격이나 실용 설계 권고가 아닙니다. 비급전·무중계 모드는 해저 리피터용 급전을 사용하지 않으므로 전류·전압·급전 전력이 0입니다.

### 범위와 한계

- 실제 부산–후쿠오카 해저케이블의 항로·설비·리피터 배치·해저 지형을 재현하지 않습니다. 통신 가능 거리나 광 링크 예산을 검증하지 않습니다.
- 수심, 경로 길이, 동물 크기, 케이블 지름 및 이동 시간은 교육용으로 축척·과장했습니다. 표시 길이는 경로 진행률에 설계 길이를 곱한 값입니다. 동물의 출현은 실제 생태 분포를 의미하지 않습니다.
- 전력 흐름은 정상상태 설명입니다. 점의 속도와 카메라 이동은 전자 드리프트 속도·신호 전파 속도·전원 인가 과도응답을 뜻하지 않습니다. 광신호는 광섬유를, 전류는 금속 도체를 따라갑니다.
- 접지·귀로 저항, 전원 효율, 전압 여유, 정격 제한, 절연, 고장·절체, 단단 급전은 생략했습니다. 절반씩 나누는 양단 전압은 선택한 예시 조건입니다.

## 구성 및 검증

- `index.html`: UI, DC 계산, 경로, 장비·환경·상어·거북 생성 함수, 카메라, 터치 조작, 애니메이션을 포함한 단일 앱
- `tests/verify.cjs`: 실제 Three.js r128 수학/지오메트리와 최소 DOM·렌더러 대역으로 수행하는 회귀 검사

WebGL이 가능하면 WebGLRenderer를 사용합니다. 불가능한 환경에서는 Canvas CPU 미리보기로 전환하며 최대 약 12 fps로 제한합니다. 호환 화면은 기본 재질 무늬·안개·상징적인 광선을 지원하지만, GPU의 실시간 그림자·반사·집광 효과는 생략합니다. 화면에는 “호환 그래픽”으로 표시합니다. 로딩에는 cdnjs 접속이 필요합니다. 정적 호스팅 또는 GitHub Pages에서 실행할 수 있습니다.

```sh
curl -L https://cdnjs.cloudflare.com/ajax/libs/three.js/r128/three.min.js -o /tmp/three-r128.js
node Cable/submarine_tour/tests/verify.cjs /tmp/three-r128.js
```

48개 검사 통과: 모드 전환, 양단 전압 및 전력 균형, 500개 설계 조합, 리피터 전압강하, 재생·속도·정지·탐색, 챕터 순서, 동물 동작 정지, 설계 반영, 도착 및 지오메트리 수 고정(1462). Blender 메시의 좌표·인덱스·법선, 화질 전환 시 계산·위치 보존, 실제 r128 셰이더 소스에 파도·집광 코드가 삽입되는 것도 검사합니다. 브라우저 Canvas 호환 화면에서 출발·상어·거북·후쿠오카 도착과 무늬·표지·화질 버튼을 확인했습니다. 장비 그림 4종 선택, 3D 시점 전환과 전압·진행 위치 보존, 선박/ROV 동작 정지, 챕터 순서 및 패널 상호 전환을 검사했습니다. 브라우저에서 네 그림과 포설선·ROV 3D 모델, ROV 직접 클릭으로 설명이 열리는 것을 확인했습니다. GPU 셰이더의 실제 드라이버 컴파일이나 그림자 품질을 검사한 것은 아닙니다. 실제 GPU 성능과 모바일 기기의 터치 성능은 별도 실기 검증이 필요합니다.

### Blender 모델 재생성

Blender 4.x가 설치된 환경에서 저장소 루트를 기준으로 실행합니다. 생성된 데이터는 HTML의 `BLENDER_MARINE_BEGIN/END` 구간만 갱신하며, 웹 실행 중 Blender를 호출하지 않습니다.

```sh
blender -b --python Cable/submarine_tour/tools/build_marine_models.py -- /absolute/path/Cable/submarine_tour/index.html
```

## 기술 참고

- [NEC: Power Feeding Equipment for Optical Submarine Cable Systems](https://www.nec.com/en/global/techrep/journal/g10/n01/pdf/100107.pdf) — 정전류 급전, 양단 전압 분담, 접지
- [ITU-T G Supplement 42 (2018)](https://www.itu.int/rec/dologin_pub.asp?id=T-REC-G.Sup42-201810-I%21%21PDF-E&lang=e&type=items) — 무중계 시스템 구성과 PFE 필요 여부

### 장비 안내의 기술 참고

- [FCC 24-119 Appendix B](https://docs.fcc.gov/public/attachments/FCC-24-119A1.pdf) — BMH·CLS·육상 백홀·데이터센터의 연결 역할만 참조했습니다. 이 앱의 부산·후쿠오카 경로와 BMH 외형은 실제 설비 정보가 아닙니다.

- [CommScope ODF 제품 안내](https://www.commscope.com/network-type/central-officeheadend/optical-distribution-frames-odf/) — 광접속·패칭·접속 트레이의 역할을 참조. 외형과 포트 배치는 직접 만든 개념 모형입니다.

- [FCC 기술 설명: CLS·SLTE와 육상망 연결](https://www.govinfo.gov/content/pkg/FR-2025-03-13/html/2025-03718.htm) — 시스템 구성과 장비 역할의 기술 설명을 참조
- [SubCom Reliance Class 선박 개요](https://www.subcom.com/documents/RelianceClass-VesselSpecs-Letter-2025-web.pdf) — 케이블 드럼·이송 장치·ROV 운용 역할
- [SubCom ST200 ROV 개요](https://www.subcom.com/documents/submersibles/SubCom-ST200-ROVSpecs-Letter-Web.pdf) — 추진기·카메라·조명·작업팔의 역할
- [NOAA ROV 설명](https://oceanexplorer.noaa.gov/technology/subs-rovs/) — 선박과 ROV의 테더 전력·제어·영상 전달

위 자료는 역할과 구성의 확인에만 사용했습니다. 그림·3D 모델은 직접 만든 개념 표현이며 특정 선박·ROV의 외형이나 정격을 복제하지 않습니다.

### 보수 장면의 표현 범위

수중 금속 구조물 보수의 일반적인 시각 요소는 [TWI 설명](https://www.twi-global.com/technical-knowledge/faqs/what-is-underwater-welding)을 참고했습니다. 광섬유 해저케이블의 접속·수리는 [KCS 기술 소개](https://k-kcs.co.jp/english/solution/introduction.html)의 선상 회수·접속 등 별도 과정과 구분했습니다. 실제 작업 깊이·전압·용접 조건·잠수 작업 절차는 계산하거나 안내하지 않습니다.

## 향후 확장

1. 단단 급전, PFE 절체, 케이블 지락 위치에 따른 전압 분포 비교
2. 실제 승인된 항로·수심 데이터와 케이블 구조별 단면 프리셋
3. SLTE·EDFA·광신호·OSNR·전력 예산을 연결하는 시스템 설계 학습 모드


## v2.8 — shore-end graphics and campus transition (2026-11-27)
- Monotone cubic camera interpolation across the campus; existing CLS door timing, design parameters and route fractions are retained. Added paved service approach.
- Both BMHs have rectangular concrete cutaways, open inspection hatch, marine sleeve, joint housing, optical/DC/earth ducts; power and earth branches are hidden in unpowered mode. Illustrative dimensions, not an as-built layout.
- 18 temporary shore-end floats with surface movement (separate from the operating signal/power route); yellow spiral identification bands on the two armored shore sections. Not a claim that all installed submarine cables use yellow bands.
- Added structural and camera seam regression assertions; actual GPU/browser rendering needs separate verification.


## v2.9 — factory-first exterior observation (2026-11-27)
- Automatic sequence: Donghae factory reference overview → Korean data center (server, router, passive ODF) → Korean CLS cutaway (SLTE, ODF, PFE in powered mode) → BMH and existing wet route → Japanese CLS cutaway → Japanese data center. Factory is not part of the actual optical circuit.
- Automatic camera never walks into DC/CLS buildings; optional manual CLS door visit remains independent. Cable progress and power equations unchanged.
- Revised regression assertions for factory-first order, outdoor camera, closed automatic doors, and independent manual door visit. JavaScript syntax checked; browser/GPU and full regression execution not performed in this session.


## v3.0 — 2026-11-27 (partial graphics and performance upgrade)
- Enhanced existing carousel, striped coils, payout animation and factory campus; no duplicate facility.
- Closed DC and CLS roofs, translucent front facade, HVAC, rack components, DC component click descriptions, CTE/NMS and separately colored illustrative earth path.
- Shared primitive geometry, instanced repeated details and coils in WebGL, preserved CPU fallback. Pointer move coalesced into animation frame; paused actor updates skipped. Five quality modes and Auto hysteresis; rolling frame timing/renderer statistics API.
- Existing power equations, length, repeater logic and dry camera path retained. Existing manual CLS door visit remains available independently.
- Validation: JavaScript syntax compilation and exact-source checks on four protected functions only. Regression test updated/added but not executed; real browser screenshots, FPS, GPU/CPU/mobile measurements, full feature regression, LOD/texture compression, damping and full tour ship-stop sequencing remain open. No performance target claimed. New assets are procedural Three.js, not Blender exports.

Performance API: `SubmarineTour.performanceStats()`, `SubmarineTour.resetPerformance()`, `SubmarineTour.setQuality('auto')`. Sample a comparable scene and quality for at least 240 frames; values are rolling browser frame intervals, not GPU benchmark results. The CPU preview does not report actual WebGL draw/triangle/FPS metrics. The factory is a photo-inspired educational model.

## v3.0.1 — 2026-11-27
- Separate automatic ship/carousel surface view and underwater ROV view using continuous camera blends; no route or power equation change.
- Frame-based pointer damping; distance-based repeated-detail visibility, low-quality water geometry; CTE/NMS/power cabinet descriptions.
- Corrected stale v2.9 UI-label assertion (baseline 49/50, one pre-existing mismatch). Actual Three.js r128 plus original test harness executed in a JavaScript isolate with Node APIs and renderer substituted: 54/54 checks pass. Baseline unique scene geometry objects 1,732; v3.0 test 553 (not GPU measurements). Browser/GPU, screenshots, real mobile and 60/30 FPS targets remain unverified. Texture atlas/compression and full environment map/postprocessing quality tiers are not implemented.

## v3.1.0 — 2026-11-27
- Black jacket and regular yellow stripes on the entire wet route, retaining the two shore armor meshes. Illustrative color scheme, not a universal cable specification.
- Blue depth-dependent haze and lighting, visible overhead water, stronger shallow caustics, crossed light shafts, 20 stylized coral patches and six animated fish schools. Procedural merged meshes, not photographs or a surveyed marine habitat. Geometry count increases with detail; browser/GPU performance and appearance are unverified.
- Slow bounded vessel translation (±3 scene units, maximum 0.075 scene units/s at 1×), heave, tether interpolation and tracking automatic camera. Not a real ship speed or laying-rate model. Pausing freezes movement. Protected optical/power calculations and route preserved.
- 56/56 regression checks with actual Three.js r128 and substituted DOM/renderer/Node APIs; syntax checks passed. No real browser screenshot or FPS measurement.
