// STEP19: run ONLY integrated browser acceptance, not legacy draft suites.
// npm install --no-save @playwright/test@^1.56.0 three@0.186.1
// npx playwright install chromium
// npx playwright test --config=DCI/DataCenter/ls3d-three/playwright.step19.config.mjs
import {defineConfig,devices} from '@playwright/test';
import {fileURLToPath} from 'node:url';
export default defineConfig({
  testDir:'.',
  testMatch:'step19.integration.spec.mjs',
  fullyParallel:false,
  retries:process.env.CI?1:0,
  workers:1,
  timeout:90000,
  expect:{timeout:20000},
  reporter:[
    ['list'],
    ['html',{open:'never',outputFolder:'DCI/DataCenter/ls3d-three/step19-playwright-report'}],
    ['json',{outputFile:'DCI/DataCenter/ls3d-three/step19-results.json'}]
  ],
  outputDir:'DCI/DataCenter/ls3d-three/step19-results',
  use:{
    ...devices['Desktop Chrome'],
    browserName:'chromium',
    headless:true,
    viewport:{width:1440,height:900},
    screenshot:'on',
    trace:'retain-on-failure',
    video:'retain-on-failure',
    launchOptions:{
      args:[
        '--enable-webgl',
        '--enable-unsafe-swiftshader',
        '--use-angle=swiftshader',
        '--disable-dev-shm-usage',
        '--ignore-gpu-blocklist'
      ]
    }
  },
  webServer:{
    command:'python3 -m http.server 8765 --bind 127.0.0.1',
    // The config lives below DCI/DataCenter/ls3d-three, but the site URL is
    // repository-relative. Serve the repository root, not the config folder.
    cwd:fileURLToPath(new URL('../../../',import.meta.url)),
    url:'http://127.0.0.1:8765/DCI/DataCenter/LS_Datacenter_3D.html',
    timeout:30000,
    reuseExistingServer:!process.env.CI
  }
});
