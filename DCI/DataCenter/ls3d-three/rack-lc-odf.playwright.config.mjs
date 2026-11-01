import {defineConfig,devices} from '@playwright/test';
import {fileURLToPath} from 'node:url';
export default defineConfig({
 testDir:'.',testMatch:'rack-lc-odf.browser.spec.mjs',workers:1,retries:0,timeout:90000,
 expect:{timeout:20000},
 reporter:[['list'],['json',{outputFile:'rack-lc-odf-browser-results.json'}]],
 outputDir:'rack-lc-odf-browser-results',
 use:{...devices['Desktop Chrome'],browserName:'chromium',headless:true,
  viewport:{width:1440,height:900},screenshot:'on',trace:'retain-on-failure',
  launchOptions:{args:['--enable-webgl','--enable-unsafe-swiftshader',
   '--use-angle=swiftshader','--disable-dev-shm-usage','--ignore-gpu-blocklist']}},
 webServer:{command:'python3 -m http.server 8765 --bind 127.0.0.1',
  cwd:fileURLToPath(new URL('../../../',import.meta.url)),
  url:'http://127.0.0.1:8765/DCI/DataCenter/LS_Datacenter_3D.html',timeout:30000,
  reuseExistingServer:!process.env.CI}
});
