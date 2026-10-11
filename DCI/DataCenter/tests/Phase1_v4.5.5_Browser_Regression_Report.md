# LS Datacenter 3D v4.5.5 — Phase 1 browser regression report

- Run date: 2026-10-11 (UTC)
- Tested commit: `ce54294670458fdbbb90a9c6b1913b0fc1cf3c2a`
- CI run: https://github.com/kimheeseo/LSCNS/actions/runs/38104463902
- Download full JSON, Markdown report and desktop/mobile screenshots: https://github.com/kimheeseo/LSCNS/actions/runs/38104463902/artifacts/11688619098
- Browser: Playwright 1.56.1 / Chromium Headless Shell 141.0.7390.37 (runner Ubuntu Linux; SwiftShader WebGL).
- Tested layouts: desktop 1440×900; simulated mobile 390×844, DPR 2.

## Overall result

**46 PASS / 0 FAIL / 0 page errors / 0 console errors.**

### Interactions covered

- On first visit, onboarding tour appears and its Skip action restores interactions.
- FOV ruler mounts with 12 log ticks; values for 39.7 cm, 74.9 cm, 1.21 m agree with FOV calculation; toggle is clickable by pointer and touch.
- Desktop wheel zoom works (30.00 m → 19.91 m).
- Mobile pinch events reached the canvas (2 touchstart events, 5 touchmove events; 40.00 m → 14.46 m). No sideways overflow at 390 px; HUD stays inside viewport.
- Rack front inspector: door opening, cassette extraction and server tray extraction pass.
- Camera/near-far and LOD stages: server tray, GPU board, GPU package and conceptual micron/nanometre illustration pass.
- Power, cooling, core network, optical-fiber-cut and normal-state transitions pass.
- Walk-mode activation and exit pass. Worker follow, first-person camera button, and exit pass.
- N, N+1, 2N comparison renders 3 cards and highlights selected 2N policy.
- Existing 2D tour and Word guidebook links remain present; CSV download succeeds.

## FPS (GitHub Actions software WebGL only)

| Scene | Mean FPS | p95 frame time |
|---|---:|---:|
| Desktop campus, medium quality | 8.0 | 150.0 ms |
| Desktop campus, low quality | 11.7 | 116.7 ms |
| Desktop GPU card LOD | 27.7 | 50.1 ms |
| Desktop conceptual die LOD | 39.4 | 50.0 ms |
| Simulated mobile campus | 11.2 | 100.1 ms |

Low quality provided around **46% higher observed FPS** than medium quality in this single CI run. These figures are **not** hardware-accelerated PC or physical phone FPS; the software renderer and shared CI resources introduce substantial uncertainty. A second run produced somewhat different numbers. Do not use these figures as production sizing criteria.

## Regression issue discovered and fixed

After selecting Compare, the telemetry/events drawer expanded at fixed `z-index:20`. It overlapped the worker POV HUD, which used a lower stacking level, so its first-person button could not receive actual pointer clicks. The fix automatically collapses the telemetry drawer upon entry into worker follow view and restores the prior expanded state on normal exit. The corresponding browser hit-target and real-click tests now pass.

Earlier false-negative test issues were also isolated:
- The first-use 5-step tour had to be dismissed before testing buttons behind it.
- The simulated phone needed to scroll the canvas into its visible viewport before emitting two-finger CDP events.

## Limitations and follow-up

- Actual Windows desktop GPU, Mac GPU, Android device and iPhone Safari were **not** tested. Mobile results use Chromium touch emulation.
- Real-site network/CDN latency, browsers other than Chromium, WebGL context loss, and long-duration stress tests were **not** measured.
- The campus view is considerably slower than near-field LOD on the CI renderer. Next performance optimization candidates: scene-wide visibility culling, geometry batching, reducing distant rack/NPC detail, and avoiding expensive canvas graph refreshes during camera interaction.
- The FOV display at micron/nanometre scale is a semantic illustration; **not measured CAD/BIM data**. Original caveats remain in the webpage.

The automated suite can be repeated through `.github/workflows/ls-datacenter-3d-phase1-browser.yml` or by editing the tested source files.
