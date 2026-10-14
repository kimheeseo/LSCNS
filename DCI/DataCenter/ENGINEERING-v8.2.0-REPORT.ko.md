# DC BOM Design Tool v8.2.0 — 계산 로직 개선 및 검산 보고서

작성일: 2026-10-14 (KST). 계획 설계용이며 모든 기본 계수는 가정입니다. PASS는 입력 산식과 포트·링크 예산의 통과를 뜻하며 제조사 인증, 실측 성능, 시공 승인을 뜻하지 않습니다.

## 1. 변경 요약

| 요청 | 구현 내용 |
|---|---|
| P1 | 서버–Leaf 연결 Rail/Single/Dual과 2/3-tier를 분리. Rail은 독립 fabric, Dual은 기존 NIC를 A/B로 분배. R/2 하향, ceil((R/2)/oversub) 상향. 독립 domain별 R²/2, R³/4 용량 확인 및 자동 3-tier 승급. 작은 fabric은 병렬 링크를 사용한 compact Clos로 스위치 수를 줄임. 실제 포트 예약 경로를 구성도·Excel에 연결. |
| P2 | Server–Leaf/Leaf–Spine/Spine–Core별 표준 길이, 광섬유 감쇠, 접속점 손실, 접속 여유 마진, splice를 검증. 테이블의 reach/loss를 수정 가능. Structured 최소 4쌍/P2P 최소 2쌍. 실패 시 다른 optics 후보 및 연결 방식 변경 안내. MPO Type-B, APC/UPC 충돌 경고. |
| P3 | IT=배치된 모든 장비 kW 합계, 시설=IT×PUE. UPS=IT/PF/효율×여유율; 발전기·변압기에 PUE 기반 냉각·비IT 부하 포함. 모듈 이용률과 N/N+1/2N 반영. 액냉/잔여 공랭 및 물 유량, CDU duty/설치 수 분리. 공랭·rear-door 한계와 냉각 부족 경고. |
| P4 | GPU→ceil(GPU/GPU-per-unit), 랙당 min(RU/전력/냉각) 및 결정 제약 표시. 시스템 계수 테이블 편집. 광링크 2 모듈, DAC/AOC 1 cable. 2×400G는 스위치 종단별로만 공유. 예비율은 케이블·광모듈만, 모두 올림. 트레이 길이+수직상승+양단여유+포설여유→표준 길이. 서로 다른 포트 경로 간 trunk 임의 pooling 방지. |
| P5 | GPU 입력/IT MW 입력/NVL72 랙 입력. IT 모드에서 예약 여유 적용 후 랙·서버·GPU 역산. 인프라 부하를 추가한 총 IT가 예산을 초과하면 경고. |
| P6 | GB200/GB300 NVL72 별도 rack-scale 경로. 72 GPU/rack, 전력 120/140kW 계획 가정. Air/Auto/rear-door 비활성 및 DLC 안내. NVLink 구리 backplane은 OEM kit/RFQ로 분리; 포함 tray는 이중 발주 금지 표기. |
| P7 | 용량과 GPU당 read GB/s 요구를 모두 사용. 노드 수=max(용량 기준, 대역폭 기준). EC 8+2=0.8/3중복제=1/3. NVMe-oF/RoCE 추가, Fibre Channel=legacy SAN. |
| P8 | OOB/BMC/1G 스위치/콘솔, A/B PDU, 3상 분산 계획, C19/C20 계획 전원선, busway/RPP branch, MPO through-adapter cassette/panel, ground/label 직접 수량 산정. 트레이 경로는 측정값 입력 시만 수량 표시, 충전율은 보수적인 전체 경로 공유 상한으로 경고. containment=옵션 allowance, 소화설비=RFQ. |

기존 입력/결과/광배선/전력/냉각/BOM/제품 후보/Excel 흐름을 유지했습니다. 기존 콜로케이션과 별도 '용량' 도구는 기존 경로를 유지합니다. AI·GPU 기본 설계와 민감도 비교는 새 브라우저 계산 모듈을 사용합니다. 정확한 SKU·가격·호스트 호환이 확인되지 않은 품목은 RFQ입니다.

## 2. 계산 규칙과 토폴로지 해석

- 'Host'는 독립 fabric의 NIC 종단입니다. 예: 8-NIC HGX 서버 128대의 Rail에서는 Rail마다 128 host이며, 8개 Rail 합계는 1,024 링크입니다. 8개 Rail의 1,024 GPU 전체를 단일 Rail의 용량 한계에 대입하지 않습니다.
- R=64에서는 하향 d=32, 1:1이면 상향 u=32, 2:1이면 u=16입니다. 요청 비율이 정수 포트와 맞지 않으면 실제 예약 비율 d/u를 표시합니다. 2-tier 최대 host=2,048/domain, 3-tier=65,536/domain은 균일 radix 및 동속 링크의 이상적인 상한입니다. ToR의 랙 지역성 때문에 그보다 일찍 계층이 필요할 수 있습니다.
- 2-tier compact Spine 수 S는 `u의 약수` 중 `Leaf 수×u/S≤R`을 만족하는 최소값. 모든 Leaf는 모든 Spine에 u/S개 링크를 연결합니다. 미사용 포트가 있는 경우에도 스위치 발주 수량에 예비율을 더하지 않습니다.
- 3-tier는 최대 d Leaf/pod. pod Spine 수 S는 u의 약수 중 `pod 최대 Leaf 수×u/S≤R/2`의 최소값. Leaf–Spine 병렬 수=u/S. Spine의 상향 예약은 최대로 채워진 pod의 하향량과 동일하게 둡니다. Core는 같은 Spine-plane별 독립 group으로 연결하며 radix를 넘지 않는 병렬 분배를 계산합니다. 작은 마지막 pod는 상향 여유가 생길 수 있습니다.
- 스위치 수는 토폴로지·병렬 연결 정책에 따라 달라지므로 유일한 이론값이 아닙니다. 검산은 이 명시된 compact Clos 구성의 링크/대역폭/포트 보존을 확인합니다. 실제 routing, ECMP/LAG, InfiniBand subnet manager, firmware와 ASIC 지원은 RFQ 검증 사항입니다.
- Dual ToR는 기존 NIC를 A/B로 분할합니다. 동일 트래픽을 양쪽에서 각각 full bandwidth로 처리하는 2N fabric 요구이면 NIC·fabric을 별도로 증설해야 합니다.

## 3. 검증 공통 가정

R=64, switch=2kW/2RU, 48RU rack 중 4RU 예약, 랙 전력·냉각 각각 100kW, 예약 headroom=20%. H200=8 GPU/8RU/10.2kW/8×400G, B200=8 GPU/10RU/14.3kW/8×400G. C는 랙 한도를 각각 120kW로 설정. NVL72=72 GPU/120kW/72×400G/rack, 4 Rail 계획. 실제 OEM 구성은 필수 확인.

트레이 입력 Server–Leaf=30m, 상위 구간=100m, 수직상승=3m, 양단여유=0m, 포설여유=15%, 예비=10%. 따라서 서버 `(30+3)×1.15=37.95m→50m`, 상위 `(100+3)×1.15=118.45m→120m custom RFQ`입니다. 100m를 초과하는 표준 목록 밖의 길이는 10m 단위 사용자정의/RFQ로 표시합니다.

PUE=1.2, PF=0.9, UPS η=0.96, 시설 headroom=20%, UPS=1,000kVA/80%, generator=2,500kW/80%, transformer=3,000kVA/80%, N+1. Frontend/OOB 포함, storage/운영 appliance 미선택. 외부 시설 부하는 PUE 기반 일괄 가정이며 실제 설비별 동시율 모델은 아닙니다.

## 4. 시나리오 결과표

| 항목 | A: B200 1,024 GPU | B: H200 4,096 GPU | C: GB200 NVL72 4랙 |
|---|---:|---:|---:|
| 시스템 수 | 128 | 512 | 4 |
| GPU | 1024 | 4096 | 288 |
| Compute 랙 | 32 | 103 | 4 |
| 전체 IT 랙 | 37 | 118 | 7 |
| Leaf | 32 | 128 | 12 |
| Spine | 16 | 64 | 8 |
| Super-spine | 0 | 64 | 0 |
| Server–Leaf 링크 | 1024 | 4096 | 288 |
| Leaf–Spine 링크 | 1024 | 2048 | 384 |
| Spine–Core 링크 | 0 | 2048 | 0 |
| Fabric 설치 광모듈 (Frontend allowance 별도) | 4096 | 16384 | 1344 |
| Compute kW | 1830.4 | 5222.4 | 480 |
| 총 IT kW | 1937.4 | 5778.9 | 527 |
| 시설 kW | 2324.88 | 6934.68 | 632.4 |
| UPS 설계 kVA / 설치 모듈 | 2690.833 / 5 | 8026.25 / 12 | 731.944 / 2 |
| 발전기 kW / 모듈 | 2789.856 / 3 | 8321.616 / 6 | 758.88 / 2 |
| 변압기 kVA / 모듈 | 3099.84 / 3 | 9246.24 / 5 | 843.2 / 2 |
| 액냉 열부하 kW | 0 | 0 | 360 |
| 잔여 공랭 kW (다른 IT 포함) | 1937.4 | 5778.9 | 167 |
| 유량 L/min | 0 | 0 | 516.006 |
| CDU duty / 설치 | 0 / 0 | 0 / 0 | 1 / 2 |

### A 단계별 검산

1. `ceil(1024/8)=128` 서버. RU=`floor((48−4)/10)=4`, 전력/냉각=`floor(100×0.8/14.3)=5`. 결정 제약=RU. `ceil(128/4)=32` compute 랙.
2. 각 Rail 128 NIC, Leaf=`ceil(128/32)=4`, u=32. Spine S=2이면 Leaf당 16 병렬 링크, Spine당 `4×16=64`포트. 8 Rail 합계 Leaf 32/Spine 16.
3. Server 링크=128×8=1,024, uplink=32×32=1,024. 전체 fabric 링크=2,048, 일반 광모듈=4,096. 구간별 10% 예비를 각각 올려 구매하면 서버 2,253+uplink 2,253=4,506 모듈.
4. 서버 구간 SR4 계획손실=`3×0.050+4×0.35=1.550dB ≤1.9dB`. 상위 DR4=`0.35×0.120+4×0.35=1.442dB ≤3dB`. 두 reach 조건도 통과. 수치는 범용 계획 테이블 가정이며 특정 NVIDIA SR4 SKU는 50m 등 더 짧은 reach를 가질 수 있으므로 제품 검증이 필요합니다.
5. Compute=`128×14.3=1830.4kW`. Fabric 48 switches×2=96kW. Frontend 13×0.5=6.5kW, OOB 9×0.5=4.5kW. IT=1937.4kW, 시설=×1.2=2324.88kW.
6. UPS=`1937.4/0.9/0.96×1.2=2690.833kVA`, usable module=800kVA. `ceil(2690.833/800)+1=5`. 발전기 `ceil(2789.856/2000)+1=3`, 변압기 `ceil(3099.84/2400)+1=3`.
7. Air 기본 가정에서 실제 랙 57.2kW가 30kW 계획 공랭 한계를 넘으므로 경고. 수량 검산 통과와 냉각 적합성 확정은 별개입니다.

### B 단계별 검산

1. `ceil(4096/8)=512` 서버. RU=`floor(44/8)=5`, 전력/냉각=`floor(80/10.2)=7`. RU 제약. `ceil(512/5)=103` compute 랙.
2. Rail당 512 host→16 Leaf. u=`ceil(32/2)=16`. 3-tier에서 pod Spine 8이면 각 Leaf–Spine 2개, Spine 하향 16×2=32/상향32. 각 Spine plane Core 1개. Rail당 Leaf16/Spine8/Core8, 합계128/64/64.
3. Server=4,096, Leaf–Spine=128×16=2,048, Core=64×32=2,048. 총8,192링크/16,384일반 광모듈. 구간별 예비올림 구매=9,012+4,506+4,506=18,024.
4. P2P 서버 SR4 손실=`3×0.050+2×0.35=0.850dB`. 상위 DR4=`0.35×0.120+2×0.35=0.742dB`. 일반 계획표 기준 통과.
5. Compute=512×10.2=5222.4kW. Fabric 256×2=512kW, Frontend49×0.5=24.5kW, OOB40×0.5=20kW. IT=5778.9kW/시설=6934.68kW.
6. UPS=`5778.9/0.9/0.96×1.2=8026.25kVA→12모듈(N+1)`. 발전기8321.616kW→6, 변압기9246.24kVA→5.
7. 각 Rail은 1 pod에 들어가므로 3-tier가 용량상 필수는 아닙니다. 사용자가 지정한 3-tier 구성으로 계산합니다. 실제 랙 51kW는 기본 Air 한계를 초과하여 경고합니다.

### C 단계별 검산

1. 서버당 8GPU 경로를 사용하지 않습니다. 4 rack×72=288GPU, 4×120=480kW compute.
2. 액냉 75%=360kW, compute 잔여 공랭=120kW. 랙당 액냉90/공랭30kW. 네트워크 등 다른 IT 47kW를 공랭에 포함하면 총 공랭=167kW.
3. 물의 `ρ=1000kg/m³, cp=4186J/(kg·K), ΔT=10K`: `360000/(1000×4186×10)=0.0086001m³/s=516.006L/min`. 즉 1kW당1.433L/min.
4. CDU600kW: duty=`ceil(360/600)=1`, N+1 설치=2. 유량은 duty 회로 총 요구량이며 설치 2대에 단순 반분한 운전 유량이 아닙니다.
5. NVLink copper backplane OEM kit=4, 포함 NVLink switch trays=36(중복 발주 금지). 외부 scale-out과 내부 NVLink는 별도 품목입니다.
6. IT=480+fabric20×2+Frontend7×0.5+OOB7×0.5=527kW. 시설632.4kW. UPS731.944kVA→2모듈, 발전기758.88kW→2, 변압기843.2kVA→2(N+1).

## 5. 검증 결과 및 제한

자동 검증 19개 통과: A/B/C 독립 수식, 전력 체인, 6조합, 2-tier 한계/승급, radix8 full fat-tree R³/4, DAC/AOC, twin-port, 예비재고, 강제 광손실 실패, low-loss, 정수 oversub 양자화, storage capacity/bandwidth, IT power 역산, NVL72 air 교정, N/N+1/2N, 비정상 입력, route별 trunk 올림.

실행: `node DCI/DataCenter/tests/engineering-v82.test.cjs`.

미검증: 실제 OEM SKU/firmware/FEC/광극성 및 APC/UPC, routing 처리량, 실측 소비전력, 액체 압력/관경/수질, 개별 냉각 설비 부하·동시율, 전기 보호협조·전원선 정격, 시공 좌표/패널 배치·트레이 공유 경로, 소화설비. '이론 일치'는 명시한 planning model의 보존식과 수량 검산 범위입니다.

## 6. 신규·변경 가정 목록

| 영역 | 기본값/규칙 |
|---|---|
| Topology | radix64; 각 tier 독립 축; d=R/2/u=ceil(d/oversub); 기본1:1; compact parallel links; Rail 독립, Dual NIC 분배 |
| 링크 | P2P2쌍/Structured4쌍; MPO0.35/low-loss0.25dB, LC0.25dB; OM4 3/OS2 0.35dB/km; 마진0dB(편집 가능), splice0개/0.1dB |
| 길이 | vertical3m/slack15%/양단0m; 1/2/3/5/10/20/30/50/100m 올림; 초과custom RFQ |
| optics | 400G DAC3m/800G DAC2m; AOC30m; SR100m/1.9dB; DR500m/3dB; FR2km/4dB; LR4-6 6km/6.3dB. 800G FR/LR는2×400G 계획군; 실제800G native SKU로 간주하지 않음 |
| Sharing | switch-side 2×400G 옵션 기본OFF. server-side NIC 모듈은 공유하지 않음. 선택시 전체 cage/광커넥터/FEC 검증 요구 |
| UPS | 효율0.96/PF0.9/1000kVA 모듈/80%사용; headroom20%; N+1기본 |
| 냉각 | DLC75%/Immersion100%; waterρ1000/cp4186/ΔT10; CDU600kW/N+1; Air30kW/rear-door50kW 경고 한계 |
| 시스템 | H20010.2kW/8RU/8×400G, B20014.3kW/10RU/8×400G, B30015kW/10RU/8×800G. 모두8GPU. NVL72 72GPU/120 또는140kW/72외부포트, 4Rail 계획 |
| 스토리지 | EC0.8/replica1/3; raw1000TB/node/read50GB/s/node; GPU당read0(default대역폭 미요구); 10kW/4RU/node |
| 부속 | console32port; panel24channel/cassette12channel; tray300×50mm/cableOD5mm/충전율40%; tray measuredroute0→RFQ |

모든 유효 입력 및 시스템/optics 테이블은 결과의 `assumptions` 객체에 보관되며 Excel Engineering Audit에 포함됩니다. 변경하면 자동 재계산하며, 모순 입력은 경고와 결과를 함께 제공합니다.

## 7. 코드 및 근거 자료

- `engineering-v82.js`: 독립 계산 모듈(CommonJS/browser 공용).
- `engineering-v82-ui.js`, `engineering-v82.css`: 기존 입력 모듈에 추가한 가정·검증 UI.
- `app-v82.js`: 최신 재계산만 렌더하는 revision guard.
- `planning-upgrade-v82.js`, `workbench-ui-v82.js`, `procurement-upgrade-v82.js`, `bom-completeness-v82.js`: 기존 탭/보고서와 신규 산식 연결.
- `engineering-v82-validation.json`: 재현 가능한 입력과 결과.

검토한 제조사 자료(이 문서의 계획 테이블 전체를 인증하는 출처는 아님):

- NVIDIA DGX GB Rack Hardware: https://docs.nvidia.com/dgx/dgxgb200-user-guide/hardware.html (72GPU/18compute trays/9NVLink trays 등 rack 구성)
- NVIDIA Transceivers and Fiber Details: https://docs.nvidia.com/networking/display/400g100gpam4ovdev/Transceivers-and-Fiber-Details-100G-PAM4 (2×400G와800G 명칭/커넥터 주의)
- NVIDIA 400G SR4 specifications: https://networking-docs.nvidia.com/mms1z00ns400sr4 (SKU별50m/APC 등 generic planning과 다름)
- Cisco400G modules: https://www.cisco.com/c/en/us/products/collateral/interfaces-modules/transceiver-modules/datasheet-c78-743172.html (FR4/LR4-6 budget/reach 및 SKU 차이)
