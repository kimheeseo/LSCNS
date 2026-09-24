# CPO Component Explorer

GitHub 기반 **CPO(Co-Packaged Optics) Component Explorer** 프로젝트입니다.

- 저장 위치: `kimheeseo/LSCNS/DCI/CPO/`
- 기존 참고 사이트: `https://cpo-component-explorer.harrykim9463.chatgpt.site`
- 운영 목표: Data Center BOM Tool과 유사하게 GitHub에서 사이트 소스와 변경 이력을 관리
- 기본 언어: 한국어
- 향후 다국어: 영어 / 중국어 등 선택형 UI

---


## v1.7.0 — 2026-09-25

### Vendor expansion, Korean news titles, and live stock charts

- Renamed component `07 Data fiber` to `07 Optical fiber`.
- Added Optical fiber companies:
  - LS전선
  - YOFC
  - ZTT
- Added `OpenLight` to the Optical Engine / PIC company list.
- Added `GlobalFoundries` and `UMC` to the Common Package company list.
- Renamed `01 외부 광원` to `01 레이저` across the component selector and diagram UI.
- Converted the stored CPO / AI optical-interconnect news headlines to Korean display titles while preserving the original source links.
- Removed the bottom explanatory notice about static GitHub Pages / JSON news maintenance.
- Added a current-market stock-chart panel to company detail cards:
  - Public companies load a current daily chart through TradingView.
  - Private/unlisted companies show a clear "direct stock chart unavailable" message.
  - LS전선 is unlisted; its card clearly labels the LS Corp. chart as a parent-company reference.
- Added/validated market mappings for major public companies including Broadcom, NVIDIA, Marvell, Coherent, Lumentum, Corning, Intel, Cisco, TSMC, ASE, Amkor, GlobalFoundries, UMC, Fujikura, Sumitomo Electric, Furukawa Electric, Hengtong, YOFC, ZTT, and LS Corp. reference.
- Updated the site/package version to `v1.7.0`.

---

## v1.6.0 — 2026-09-25

### Component-only 10-second highlight interaction

- Removed the top `LSCNS / DCI / CPO` breadcrumb from the homepage header.
- Changed diagram interaction so only the selected component and its matching callout are highlighted.
- Added a repeating sparkle/pulse effect that runs for approximately 10 seconds after each component click.
- After 10 seconds, only the visual highlight is cleared; the selected component, company list, product information, and news detail remain selected.
- Optical Engine selection highlights the optical-engine modules only.
- Switch ASIC selection highlights the central ASIC only.
- ELS, PM fiber, common package, FAU, data fiber, and front connector selections use the same component-specific behavior.
- Clicking another component immediately cancels the previous highlight and starts a new 10-second highlight for the new selection.
- Initial page load and search filtering do not trigger the 10-second animation automatically.
- Updated the site/package version to `v1.6.0`.

---

## v1.5.0 — 2026-09-24

### More realistic 3D CPO package visual

- Refined the homepage CPO package to look closer to a physical hardware render while keeping it browser-rendered and sharp at different resolutions.
- Added stronger perspective, board thickness, multilayer package depth, metallic ELS housing, package traces, gold ASIC pin detail, raised optical-engine modules, metallic FAU structures, blue front connectors, board screws, and passive SMD details.
- Improved blue data-fiber and orange PM-fiber depth using shadows and optical glow effects.
- Preserved interactive component highlighting:
  - Optical Engine selection highlights all eight optical-engine modules in blue.
  - Switch ASIC selection highlights the center switch chip in orange.
  - ELS, PM fiber, common package, FAU, data fiber, and front connectors retain selection effects.
- Kept the existing related-company, representative-product, official-link, and recent-news panels.
- Updated the site/package version to `v1.5.0`.
- Render deployment settings are unchanged.

---

## v1.4.0 — 2026-09-24

### Crisp interactive 3D CPO model + component highlight actions

- Removed the low-resolution raster CPO artwork from the main UI to prevent blur/pixelation when the browser scales the diagram.
- Rebuilt the CPO package as an original HTML/CSS 3D model rather than copying the supplied reference layout.
- The new model stays sharp at different browser sizes because the board, ASIC, optical engines, ELS, FAUs, fibers, and connectors are rendered as browser elements rather than a stretched bitmap.
- Added persistent selection highlighting on the actual hardware:
  - Optical Engine selection highlights all eight optical-engine modules in blue.
  - Switch ASIC selection highlights the central ASIC with an orange outline/glow.
  - ELS, PM fiber, package, FAU, data-fiber routes, and front connectors also receive component-specific highlighting.
- Added a short pulse animation when a component is selected.
- Diagram clicks now keep the diagram in view instead of immediately scrolling away, so the highlight action can be seen.
- Related-company, product/role, official-link, and recent-news panels continue to update from the selected component.
- Updated the site/package version to `v1.4.0`.
- Render deployment settings remain unchanged: `npm install` → `npm start`, health check `/api/health`.

---

## v1.3.0 — 2026-09-24

### 3D CPO package image + interactive hotspot update

- Replaced the previous schematic SVG illustration with a rendered 3D/isometric CPO package image.
- Added the generated image as `assets/cpo-package-3d.webp`.
- Preserved all 8 interactive component selections using transparent responsive hotspot buttons over the image:
  - 01 External Laser Source
  - 02 PM Fiber
  - 03 Optical Engine / PIC + EIC
  - 04 Switch ASIC
  - 05 Common Package
  - 06 FAU
  - 07 Data Fiber
  - 08 Front Connector
- Clicking a hotspot continues to update the related-company list and company detail/news panel.
- Optical Engine remains the default selection, with Broadcom shown first in the detail panel.
- Updated the site/package version to `v1.3.0`.
- Render deployment remains compatible with the existing `npm install` / `npm start` setup.

---

## v1.2.1 — 2026-09-24

### Render Web Service deployment support

- Added `package.json` so Render can successfully run `npm install`.
- Added `server.js` using the built-in Node.js HTTP server; no external runtime dependency is required.
- Added `npm start` script for Render Web Service deployment.
- Added `/api/health` endpoint for Render health checks.
- Server binds to `0.0.0.0` and uses Render's `PORT` environment variable.
- Static files in `DCI/CPO/` are served directly, with `index.html` as the default page.
- Updated the site version badge to `v1.2.1`.

### Render settings

- Service Type: `Web Service`
- Repository: `kimheeseo/LSCNS`
- Branch: `main`
- Root Directory: `DCI/CPO`
- Runtime: `Node`
- Build Command: `npm install`
- Start Command: `npm start`
- Health Check Path: `/api/health`
- Auto Deploy: `Yes`

---

## v1.2.0 — 2026-09-24

### Interactive CPO package diagram + vendor/news explorer

- Rebuilt the main HTML around an interactive CPO package / physical-layout schematic inspired by the supplied reference image.
- Added 8 clickable CPO component regions:
  - External Laser Source (ELS)
  - PM Fiber
  - Optical Engine / PIC + EIC
  - Switch ASIC
  - Common Package / Substrate
  - FAU
  - Data Fiber
  - Front Connector / VSFF
- Clicking a component now lists relevant companies for that component.
- Clicking a company now shows:
  - company role in the CPO stack
  - representative product / technology focus
  - official company/product link
  - curated recent CPO / AI optical-interconnect article links
  - a live "latest news search" link for additional current coverage
- Added curated 2026 CPO-related news for major ecosystem companies including Coherent, Lumentum, Ayar Labs, Broadcom, NVIDIA, Marvell, Lightmatter, Molex, Corning, SENKO, US Conec, Fujikura, Sumitomo Electric, Furukawa Electric, Hengtong, and Intel.
- Added search, component selector, component quick cards, and responsive mobile layout.
- Updated page version badge to v1.2.0.
- Static-site note: the page does not call a live news API; curated links are stored in the HTML and can later be moved into JSON data files for easier maintenance.

---

## v1.1.0 — 2026-09-24

### Interactive HTML prototype added

- Added `DCI/CPO/index.html` as a self-contained interactive CPO Component Explorer.
- Added clickable CPO architecture flow.
- Added component/category search and filtering.
- Added representative company chips by component.
- Added selected-component detail view.
- Added roadmap for product/news/simulation/multilingual expansion.
- Current HTML is an initial structure/UI prototype; product/news data will be validated and expanded in later versions.

---

## v1.0.0 — 2026-09-24

### GitHub 프로젝트 운영 구조 정의

CPO Component Explorer를 GitHub 기반 정적 사이트로 이전·운영하기 위한 기본 구조와 업데이트 원칙을 정의했습니다.

### 핵심 기능

1. **CPO 구성도 기반 탐색**
   - CPO 시스템/모듈 그림에서 부품 클릭
   - 클릭한 부품의 관련 업체 및 제품 확인

2. **부품별 업체 Ecosystem**
   - Laser / Light Source
   - Modulator
   - PIC / Optical Engine
   - ASIC / xPU
   - FAU
   - MPO / VSFF / SN / MDC 등 Connector
   - PM Fiber / SMF / 관련 Fiber
   - DSP / Driver / TIA
   - 기타 CPO 관련 부품

3. **업체 상세 정보**
   - 업체명
   - 공식 홈페이지
   - 대표 CPO 관련 제품
   - 간단한 기술/제품 설명
   - 해당 기업의 최근 CPO 관련 뉴스 제목 약 5건

4. **제품 및 업체 데이터 분리**
   사이트 코드와 데이터를 분리하여 유지보수성을 높이는 방향을 기본으로 합니다.

   ```text
   DCI/CPO/
   ├── index.html
   ├── styles.css
   ├── app.js
   ├── data/
   │   ├── components.json
   │   ├── companies.json
   │   ├── products.json
   │   └── news.json
   ├── assets/
   │   └── images/
   └── readme.md
   ```

5. **다국어 지원**
   - 한국어를 기본 표시 언어로 사용
   - 향후 영어 / 중국어 등 언어 선택 시 설명 및 UI 텍스트 전환

---

## 향후 업데이트 원칙

앞으로 CPO Component Explorer를 수정할 때는 다음 순서로 관리합니다.

```text
사이트 기능/데이터 수정
        ↓
GitHub DCI/CPO 코드 업데이트
        ↓
버전 증가
        ↓
readme.md에 변경사항 기록
```

예:

```text
v1.1.0 — PM Fiber ecosystem update
v1.2.0 — Connector / VSFF product update
v1.3.0 — Vendor product + official URL + news update
```

### README 변경 이력 작성 항목

- 버전
- 수정 날짜
- 추가/변경된 기능
- 신규 업체/제품
- 데이터 구조 변경
- UI/UX 변경
- 오류 수정
- 향후 추가 예정 기능

---

## 계획된 사이트 동작 예

예를 들어 **PM Fiber**를 클릭하면:

```text
PM Fiber
  ├─ Hengtong
  │   ├─ 대표 제품
  │   ├─ 공식 홈페이지
  │   └─ 최근 CPO 관련 뉴스
  ├─ Fujikura
  ├─ Sumitomo Electric
  └─ 기타 관련 업체
```

Connector를 클릭하면:

```text
Connector
  ├─ MPO
  ├─ SN
  ├─ MDC
  └─ VSFF
```

각 항목에서 관련 업체, 대표 제품, 공식 URL, 최근 CPO 관련 정보를 확인할 수 있도록 구성합니다.

---

## 개발 방향

단순 업체 목록이 아니라 다음 형태의 **CPO Supply-Chain / Component Intelligence Tool**로 발전시키는 것을 목표로 합니다.

```text
CPO Architecture
      ↓
Component
      ↓
Company
      ↓
Representative Product
      ↓
Official Product / Company URL
      ↓
Recent CPO News
```

추후에는 제품 비교, 검색/필터, 기술 분류, CPO 시뮬레이션 결과 연계 등의 기능도 추가할 수 있습니다.
