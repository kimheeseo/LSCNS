/* LS 3D Phase 3 A/B browser benchmark: baseline v4.6.0 vs optimized v4.6.1.
   Same Chromium/host/viewport. GPU uses headless SwiftShader; measured, NOT hardware FPS. */
'use strict';
const {chromium}=require('playwright'),fs=require('node:fs'),path=require('node:path'),http=require('node:http'),os=require('node:os');
const root=path.resolve(__dirname,'..'),baseline=path.resolve(__dirname,'test-results','phase3','baseline'),out=path.resolve(__dirname,'test-results','phase3');
fs.mkdirSync(out,{recursive:true});const results={host:{platform:os.platform(),cpu:os.cpus()[0]?.model||'',arch:os.arch(),node:process.version},timestamp:new Date().toISOString(),samples:[],errors:[],warnings:[]};
const sv=http.createServer((req,res)=>{
 const url=new URL(req.url,'http://localhost'),p=decodeURIComponent(url.pathname);
 if(!/^\/(current|baseline)\//.test(p)){res.writeHead(403).end();return}
 const base=p.startsWith('/baseline/')?baseline:root;
 const rel=p.replace(/^\/(current|baseline)\//,'');
 const name=path.resolve(base,rel);
 if(!name.startsWith(base+path.sep)){res.writeHead(403).end();return}
 fs.readFile(name,(err,buf)=>{
 if(err){res.writeHead(404).end('missing resource '+rel);return}
 res.setHeader('Content-Type',({'.js':'application/javascript','.css':'text/css','.html':'text/html'}[path.extname(name)]||'application/octet-stream')+';charset=utf-8');
 res.setHeader('Cache-Control','no-store');res.writeHead(200).end(buf);
 });
});
async function measure(page,stage,version){
 await page.goto('http://127.0.0.1:9875/'+version+'/LS_Datacenter_3D.html?perf=3',{waitUntil:'domcontentloaded',timeout:35000});
 await page.waitForFunction(()=>window.__LS3D_TEST__?.camera&&window.LS3D_PHASE1&&window.LS3D_CONFIG,{timeout:24000});
 await page.evaluate(q=>{const x=window.__LS3D_TEST__;window.LS3D_CONFIG.quality=q;x.cameraDesired.distance=164;x.cameraDesired.target=[0,2,0];},stage);
 const skip=page.locator('#tourSkip');if(await skip.count())try{if(await page.locator('#guide.open').count())await skip.click({timeout:2500})}catch(_){}
 await page.waitForTimeout(1200);
 const fps=await page.evaluate(async()=>{
   const stamps=[],start=performance.now();await new Promise(resolve=>{function frame(t){stamps.push(t);if(t-start<4200)requestAnimationFrame(frame);else resolve()}requestAnimationFrame(frame)});
   const dt=stamps.slice(1).map((v,i)=>v-stamps[i]).filter(x=>x>0),secs=(stamps.at(-1)-stamps[0])/1000;
   dt.sort((a,b)=>a-b);const q=(n)=>dt[Math.min(dt.length-1,Math.floor((dt.length-1)*n))];
   return{meanFps:Math.round((dt.length/secs)*10)/10,medianFrameMs:Math.round(q(.5)*10)/10,p95FrameMs:Math.round(q(.95)*10)/10,frames:dt.length,seconds:Math.round(secs*100)/100,stats:window.LS3D_RENDER_METRICS?{...window.LS3D_RENDER_METRICS}:null,gl:!!window.__LS3D_TEST__.gl};
 });
 results.samples.push({version,quality:stage,...fps});
 console.log('SAMPLE '+version+' '+stage+' '+JSON.stringify(fps));
 if(!fps.gl||fps.frames<3)throw Error('WebGL or frames missing: '+version+' '+stage);
 await page.screenshot({path:path.join(out,version+'-'+stage+'.png')});
}
(async()=>{
 await new Promise(resolve=>sv.listen(9875,'127.0.0.1',resolve));let browser;
 try{
 browser=await chromium.launch({headless:true,args:['--no-sandbox','--disable-dev-shm-usage','--enable-webgl','--use-gl=angle','--use-angle=swiftshader','--disable-background-timer-throttling']});
 const ctx=await browser.newContext({viewport:{width:1440,height:900},deviceScaleFactor:1});
 const page=await ctx.newPage();
 page.on('pageerror',e=>results.errors.push(String(e)));
 for(const [version,quality]of [['baseline','medium'],['current','medium'],['baseline','low'],['current','low']])await measure(page,quality,version);
 await ctx.close();
 }catch(e){results.errors.push(String(e?.stack||e));console.error(e);}
 finally{
 if(browser)await browser.close().catch(()=>{});
 await new Promise(resolve=>sv.close(resolve));
 const by=(v,q)=>results.samples.find(s=>s.version===v&&s.quality===q);
 results.comparison=['medium','low'].map(q=>{const b=by('baseline',q),n=by('current',q);return{quality:q,baselineFPS:b?.meanFps||0,optimizedFPS:n?.meanFps||0,ratio:b&&n?Math.round(n.meanFps/b.meanFps*1000)/1000:null,optimizedVertices:n?.stats?.triVertices||null,culledAssets:n?.stats?.culledAssets||0,liteRacks:n?.stats?.liteRacks||0,mediumRacks:n?.stats?.mediumRacks||0,fullRacks:n?.stats?.fullRacks||0}});
 fs.writeFileSync(path.join(out,'benchmark.json'),JSON.stringify(results,null,2));
 const rows=['# Phase 3 same-run software WebGL A/B benchmark','','**Environment:** Chromium Headless Shell / Linux Ubuntu / SwiftShader · not a physical-device FPS result','','| Quality | v4.6.0 FPS | v4.6.1 FPS | Ratio | Visible vertices (new) | Culled | Lite racks | Medium racks |','|---|---:|---:|---:|---:|---:|---:|---:|'];
 for(const x of results.comparison)rows.push('| '+x.quality+' | '+x.baselineFPS+' | '+x.optimizedFPS+' | '+x.ratio+'x | '+x.optimizedVertices+' | '+x.culledAssets+' | '+x.liteRacks+' | '+x.mediumRacks+' |');
 if(results.errors.length)rows.push('','## Errors',...results.errors);
 fs.writeFileSync(path.join(out,'benchmark.md'),rows.join('\n'));
 console.log('COMPARISON '+JSON.stringify(results.comparison));console.log('ERRORS '+results.errors.length);
 }
 process.exitCode=results.errors.length?1:0;
})();
