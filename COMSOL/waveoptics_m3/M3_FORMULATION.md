# M3 정식화와 선작성 검증 계약

시간 e^{-iωt}; 평면 x,y; 불변 방향 z. TE=E_z, TM=H_z.
∇·(p∇u)+k₀²qu=0. TE (p,q)=(1,n²), TM=(1/n²,1).
약형 a(u,v)=∫ D∇u·∇v̅−k₀²Quv̅−iΣΓ〈Tu,v〉.
PML sx=1+iσx, sy=1+iσy; D=p diag(sy/sx,sx/sy), Q=q sx sy.
σj=σmax(dj/tj)^r; dj는 물리 영역 바깥 거리. 코너는 두 신장을 모두 적용.
포트 횡방향 Kφ−k₀²Mqφ=−β²Mpφ, φ*Mpφ=1.
벽 PEC: TE u=0, TM p∂nu=0. 슬랩 포트는 유한 횡방향 창으로 근사.
포트 외향 플럭스 p∂nu=iT(u−2uinc). β의 제곱근 가지는 Reβ≥0, Imβ≥0.
입사 부하 a(u,v)=−2i〈Tuinc,v〉. T의 이산 행렬은 MpΦ diag(β)Φ*Mp.
유한 포트 trace 공간의 모든 고유벡터를 유지하여 소멸 모드도 DtN에 포함.
좌표 단위 μm, 파장 μm. 삼각형 P1 Galerkin, 직사각형 좌표 간격 ≤h/3, 6차 체적 적분, 직접 희소 LU.
무손실 실수 포트만 지원. 포트 진폭 c=Φ*Mp uΓ, b=c−ainc.
전파 모드 Smi=bm √(βm/βi), 전력 |S|². β≈0 cutoff에서는 명시적으로 거부.
일반 재료 경계 u와 p∂nu 연속. 내부 재료의 수동 복소 n 지원.

평면파 산란: u=ui+us, ui=exp(ikb x), 물리 contrast가 PML과 겹치지 않아야 함.
aPML(us,v)=−∫(p−pb)∇ui·∇v̅+k₀²∫(q−qb)ui v̅.
이 부하는 배경 항의 차이로 유도되며 재료 contrast 영역에만 작용.
PML 외곽 us=0. 출력 총장은 물리 영역에서 ui+us로 재구성.
전자기 필드 복원: TE Z₀H=(∂yu,−∂xu,0)/(ik₀), E=(0,0,u).
TM E=iZ₀(∂yu,−∂xu,0)/(k₀ n²), H=(0,0,u).
코드 TM 기본 u는 H_z; 모든 필드는 임의 입사 진폭 기준이며 S는 단위 무관.

독립 해석 기준: 직사각형 β²=(k₀n)²−(mπ/W)²;
층 전달행렬 state=(u,p∂xu), Mj=[[cosβL,sinβL/(pβ)],[-pβ sinβL,cosβL]].
원통 exterior us=Σm i^m am Hm^(1)(kb r)e^{imθ};
am=(pi ki Ji' Jb−pb kb Jb' Ji)/(pb kb Hb' Ji−pi ki Ji' Hb).
산란폭 Csca=(4/kb)Σ|am|² [μm]; 3D 단면적 아님.
수렴 차수 slope=log(eh/eh2)/log(hmax/hmax2); 실제 최대 삼각형 변 길이 보고.
pytest 임계값은 최초 작성 시 고정하며 테스트 실패 후 늘리지 않음.

검증 범위: PEC 유한 폭 포트, 유한 창 슬랩 포트, 평면파 원통 PML.
개방 광도파관의 PML 횡단 모드 포트, M2 벡터 모드의 3D 포트 결합은 이번 범위 아님.
