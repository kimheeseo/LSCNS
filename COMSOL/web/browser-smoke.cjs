/* GitHub Actions Playwright smoke for standalone GitHub Pages Wave Optics Lab */
'use strict';
const {chromium}=require('playwright'),http=require('node:http'),fs=require('node:fs'),path=require('node:path');
const root=path.resolve(__dirname,'..'),results=path.join(__dirname,'test-results');
fs.mkdirSync(results,{recursive:true});
const server=http.createServer((req,res)=>{
 const p=decodeURIComponent(new URL(req.url,'http://localhost').pathname);
 const relative=p==='/'?'index.html':p.replace(/^\/+/,''),target=path.resolve(root,relative);
 if(!target.startsWith(root+path.sep)){res.writeHead(403).end();return}
 fs.readFile(target,(err,buf)=>{
  if(err){res.writeHead(404).end('not found');return}
  res.writeHead(200,{'Content-Type':({'.js':'text/javascript','.css':'text/css','.html':'text/html'}[path.extname(target)]||'application/octet-stream')+';charset=utf-8'}).end(buf);
 });
});
const rec={checks:[],errors:[]};
const check=(name,passed,info)=>{rec.checks.push({name,passed:!!passed,info});console.log((passed?'PASS ':'FAIL ')+name+' '+(info||''))};
(async()=>{
 let browser;
 try{
  await new Promise(resolve=>server.listen(8588,'127.0.0.1',resolve));
  browser=await chromium.launch({headless:true,args:['--no-sandbox','--disable-dev-shm-usage']});
  const desktop=await browser.newPage({viewport:{width:1400,height:920}});
  desktop.on('pageerror',e=>rec.errors.push('desktop: '+e.message));
  await desktop.goto('http://127.0.0.1:8588/',{waitUntil:'networkidle'});
  await desktop.waitForSelector('#metrics .metric',{timeout:12000});
  let v=await desktop.evaluate(()=>({metrics:document.querySelector('#metrics').innerText,canvas:document.querySelector('#chart-a').width,map:document.querySelector('#mode-map').width}));
  check('desktop-LP01-plot',v.metrics.includes('1.4461')&&v.canvas>0&&v.map>0,JSON.stringify(v).slice(0,280));
  await desktop.locator('#radius').fill('4.4');
  await desktop.locator('#calculate').click();
  v=await desktop.locator('#metrics').innerText();check('desktop-input-recompute',!v.includes('1.4461047'),v.slice(0,200));
  await desktop.locator('[data-mode="mcf"]').click();
  let m=await desktop.locator('#metrics').innerText();
  check('desktop-MCF',m.includes('10.43')&&m.includes('0.000074'),m.slice(0,230));
  await desktop.locator('[data-mode="hcf"]').click();
  let h=await desktop.locator('#metrics').innerText();
  check('desktop-HCF',h.includes('0.9994878')&&h.includes('3.37'),h.slice(0,220));
  await desktop.screenshot({path:path.join(results,'desktop-wave-optics.png'),fullPage:true});
  const mobile=await browser.newPage({viewport:{width:390,height:844},deviceScaleFactor:2,isMobile:true,hasTouch:true});
  mobile.on('pageerror',e=>rec.errors.push('mobile: '+e.message));
  await mobile.goto('http://127.0.0.1:8588/',{waitUntil:'networkidle'});
  await mobile.waitForSelector('#metrics .metric',{timeout:12000});
  const geo=await mobile.evaluate(()=>({width:innerWidth,scroll:document.documentElement.scrollWidth,plot:document.querySelector('#chart-a').getBoundingClientRect().toJSON(),metricCount:document.querySelectorAll('.metric').length}));
  check('mobile-layout-no-horizontal-scroll',geo.scroll<=geo.width+4,JSON.stringify(geo).slice(0,350));
  await mobile.locator('[data-mode="mcf"]').click();
  check('mobile-touch-tab',await mobile.locator('#metrics').innerText().then(v=>v.includes('10.43')));
  await mobile.screenshot({path:path.join(results,'mobile-wave-optics.png'),fullPage:true});
  await desktop.close();await mobile.close();
 }catch(e){rec.errors.push(String(e.stack||e));console.error(e)}
 finally{
  if(browser)await browser.close().catch(()=>{});
  await new Promise(resolve=>server.close(resolve));
  rec.passed=rec.checks.filter(x=>x.passed).length;rec.failed=rec.checks.filter(x=>!x.passed).length;
  fs.writeFileSync(path.join(results,'report.json'),JSON.stringify(rec,null,2));
  console.log('REPORT '+JSON.stringify({passed:rec.passed,failed:rec.failed,errors:rec.errors.length}));
  if(rec.failed||rec.errors.length)process.exitCode=1;
 }
})();
