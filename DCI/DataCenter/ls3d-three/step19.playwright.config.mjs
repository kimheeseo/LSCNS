import {defineConfig,devices} from '@playwright/test';

export default defineConfig({
  testDir:'.',
  testMatch:'step19.integration.playwright.spec.mjs',
  timeout:60000,
  expect:{timeout:12000},
  fullyParallel:false,
  workers:1,
  retries:0,
  reporter:[
    ['line'],
    ['html',{outputFolder:'step19-playwright-report',open:'never'}],
    ['json',{outputFile:'step19-test-results.json'}]
  ],
  outputDir:'step19-test-results',
  use:{
    ...devices['Desktop Chrome'],
    viewport:{width:1440,height:900},
    baseURL:process.env.LS3D_BASE_URL||
      'http://127.0.0.1:8765/DCI/DataCenter/LS_Datacenter_3D.html',
    launchOptions:{
      args:['--enable-webgl','--use-gl=angle','--use-angle=swiftshader',
        '--enable-unsafe-swiftshader','--disable-gpu-sandbox']
    },
    screenshot:'only-on-failure',
    trace:'retain-on-failure',
    video:'retain-on-failure'
  }
});
