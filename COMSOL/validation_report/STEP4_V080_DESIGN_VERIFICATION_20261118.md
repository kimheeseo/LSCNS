# STEP 4 · 계산 물리량 — 설계 / 구현 / 검증표 (v0.8.0)

일자: 2026-11-18 KST. 범위: STEP 4 ONLY. STEP 5 Bloch–Floquet/Photonic Crystal bandgap 기능은 미착수.

## 1. 설계 및 원칙
- 단위·부호 계약은 Python 소스 physical_metrics.py 최상단에 고정. e^(iβz−iωt), Im(neff)>0 ⇒ 수동계 감쇠.
- 복소 neff 허수부의 부호는 절댓값으로 바꾸지 않는다. 재료 흡수가 있으면 confinement-only가 아니라 **total loss**.
- 모든 Aeff와 second-moment MFD는 Et=(Ex,Ey) 및 물리 영역 적분을 사용한다. Ez를 포함하던 과거 실험 결과와 정의가 다르므로 basis metadata를 저장.
- 파장/설계 스윕은 공통 면적 가중 complex transverse field overlap의 Hungarian 일대일 배정. mesh 좌표가 다르거나 인접 overlap이 기준 미달이면 D/S/λ0 계산을 거절.
- 알 수 없는 n₂, 응력광학 재료계수, HOM 정체, PML 손실 인증은 TODO 및 필요 자료를 명시한다.
- 기존 디자인 JSON schema_version 1.0.0은 하위 호환 보존하고 **엔진 버전 v0.8.0**을 별도 표시.

## 2. 공식 및 단위 계약

| 계산량 | 공식 | 단위 / 동작 |
| --- | --- | --- |
| total α | α=(20/ln10)(2π/λ)Im(neff) | λ[m], dB/m. ×1000=dB/km. 부호 보존 |
| Aeff | (∫It dA)²/∫It²dA, It=|Ex|²+|Ey|² | µm², PML 제외 |
| 2nd MFD | 2√(2⟨|r−r̄|²⟩) | µm |
| Petermann-II | 2√(2 ∫|E(r)|² r dr / ∫|dE/dr|²rdr) | 축대칭 radial scalar만 Petermann-II. 비축대칭 2D gradient 방식은 PII-like 표시 |
| D, S, λ0 | D=−(λ/c)d² Re(neff)/dλ², S=dD/dλ, λ0:D=0 | ps/(nm km), ps/(nm² km), nm; 겹침 추적된 5+점 필요 |
| γ | 2πn₂/(λAeff) | n₂[m²/W], λ[m], Aeff[m²], γ[W⁻¹m⁻¹] |
| HCF effective n₂ | ∫ n₂(x,y)It²dA / ∫It²dA | 가스/유리 조성별 n₂ 필수, scalar weak-guidance 근사 |
| 유리 power 비율 | Pglass/Ptotal, Sz=0.5 Re(ExHy*−EyHx*) | 무차원, glass/air mask 필수 |
| HOMER | min αHOM / max αFM | 무차원, 양의 수렴 손실 및 FM/HOM 모드 식별 필수 |
| 복굴절 / beat | B=|nx−ny|; LB=λ/B | 무차원, m |
| 모드 간격 | |Re(neff,i)−Re(neff,j)| | 무차원 |
| ARROW 공진 | λm=2t√(ng²−na²)/m | t[µm]→nm, **상수 굴절률 근사** |
| anticrossing | 두 추적 모드의 neff gap 국소 최소 + glass fraction 교환 | 후보만 표시, 실제 surface mode 인증 아님 |
| overlap | |∫Et,i*·Et,jdA|²/(∫|Et,i|²dA ∫|Et,j|²dA) | 면적 가중 무차원, 공통 메시 필수 |

## 3. 실제 구현 경로
- Python physical_metrics.py: 위 수식·단위·오류 및 TODO 처리.
- Python vector_fem.py: Et-only Aeff/MFD, dB/m & dB/km, 선택적 full quadrature E 복소장 출력.
- Python step4_cli.py: 단일 FEM 결과 후처리 / 공통 메시 파장 FEM 스윕 / 추적된 분산 결과.
- 웹 research_step4_v080.mjs: 정적 결과 가져오기, Re(neff) vs 파장에 ARROW 공진 수직선, D(λ) 표시. 실제 FEM 데이터 없으면 임의 곡선을 만들지 않음.
- tests/test_step4_physical_metrics.py 및 tests/test_step4_cli.py: analytic Gaussian/다항식/overlap/HCF guard 테스트.

## 4. 검증표

증거: GitHub Actions https://github.com/kimheeseo/LSCNS/actions/runs/44138171783 (Python 3.11 및 Node JS 검사). 이 보고서의 PASS는 **공식 계산/구현 단위 검증**이지 물리적으로 수렴한 HCF 예측을 뜻하지 않는다.

| 항목 | 시험 내용 | 판정 |
| --- | --- | --- |
| α dB/m ↔ dB/km | 1000배, Im 부호, zero 입력 | PASS |
| Gaussian Aeff/MFD | Aeff≈πw², MFD≈2w | PASS |
| 축대칭 radial Petermann-II | Gaussian 2w | PASS |
| n₂/γ/HCF quartic | SI 환산, air/glass 값 누락 시 TODO | PASS |
| 유리 Poynting | 균일 정방/역방향 광력 샘플 | PASS |
| HOMER/B/LB/Δneff | 정식 및 모드 미인증 시 차단 | PASS |
| D/S/λ0 | analytic quadratic/cubic neff, overlap 불충분 차단 | PASS |
| overlap tracking | x/y 모드 순서 교환 정정, mesh 불일치 차단 | PASS |
| ARROW, anticrossing | 파장 단위/차수, 근접 교환 후보 검사 | PASS (후보 분류만) |
| Python/JavaScript | compileall, Node --check, JSON Schema 25종 | PASS |
| STEP 3 V1/V4/V5 | 독립 HE11/벡터/메시·PML 수렴 및 분리 수용 | NOT PASS / 미검증 |
| HCF 정량 손실 | 실제 모드와 벽 메시, PML 독립 변화 | NOT PASS / 미검증 |

CI는 총 29개의 회귀/단위/기존 smoke 테스트 통과 기록을 갖는다. 실제 FEM 광학 검증은 별도로 진행해야 한다.

## 5. 실행법
설치 폴더: COMSOL/universal_silica_v01/fiber_research_v1/

    python step4_cli.py result out_step/mode_result.json --out step4_metrics.json
    python step4_cli.py result out_step/mode_result.json --n2-m2-W 2.2e-20 --out step4_metrics.json
    python step4_cli.py solve-sweep presets/index_guiding__step.json --start-nm 1500 --stop-nm 1600 --points 7 --out out_sweep --modes 2

두 번째 줄의 n₂ 값은 **입력 방법 예시**일 뿐 실제 해당 광섬유의 공식 계수가 아니다. 실행 시 직접 확인한 공인 값으로 대체해야 한다. 7회 FEM 스윕은 로컬/Colab에서 실행하며, 메모리/시간 사용량이 클 수 있다.

## 6. 다음 단계 제안 — 승인 전 미착수
사용자 지정 순서에 맞춰 먼저 STEP 3의 V1/V4/V5 독립 검증 게이트를 통과시킨 후, 고체 PCF(V2/V6), STEP 5 Bloch-bandgap 순으로 진행할 것을 제안한다. V1/V4/V5의 정량 정의가 기존 자료에서 명확하지 않으면 추정 없이 원자료를 확인한다.

## 출처
- IEC 60793-1-45:2024 (Petermann-II): https://www.iec.ch/
- Petermann I/II 구분: https://optiwave.com/optifiber/optifiber-tutorials-optifiber/optical-fiber-mode-field-diameter-and-mode-area-definitions/
- HOMER: https://doi.org/10.1364/OE.27.003824
- ARROW resonance: https://www.mdpi.com/2079-6439/6/4/89
