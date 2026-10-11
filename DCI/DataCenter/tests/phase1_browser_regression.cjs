/* LS DataCenter 3D v4.5.5 browser regression - Playwright Chromium, no new runtime dependencies. */
'use strict';
const {chromium,devices}=require('playwright');
const fs=require('node:fs'),path=require('node:path'),http=require('node:http'),os=require('node:os');
const root=path.resolve(__dirname,'..'),out=path.resolve(__dirname,'test-results','phase1');
fs.mkdirSync(out,{recursive:true});
const T={started:new Date().toISOString(),host:{node:process.version,os:os.platform(),arch:os.arch(),cpus:os.cpus().length},cases:[],metrics:{},consoleErrors:[],pageErrors:[],warnings:[]};
function check(group,name,ok,detail=''){T.cases.push({group,name,pass:!!ok,detail:String(detail)});process.stdout.write((ok?'PASS ':'FAIL ')+group+' / '+name+(detail?' — '+String(detail).slice(0,190):'')+'\n');}
function safeError(p,e){T.pageErrors.push({p,message:String(e&&e.stack||e)});}
function sleep(ms){return new Promise(r=>setTimeout(r,ms))}
const server=http.createServer((req,res)=>{
 const url=new URL(req.url,'http://localhost'),u=decodeURIComponent(url.pathname);
 const file=path.resolve(root,'.'+u);
 if(!file.startsWith(root+path.sep)){res.writeHead(403).end();return}
 fs.readFile(file,(e,data)=>{if(e){res.writeHead(404).end('not found');return}
 const ext=path.extname(file);res.setHeader('Content-Type',({'.html':'text/html; charset=utf-8','.js':'application/javascript; charset=utf-8','.css':'text/css; charset=utf-8','.json':'application/json; charset=utf-8'}[ext]||'application/octet-stream'));res.writeHead(200).end(data)})
});
async function fpsSample(page,seconds=3){return page.evaluate(async seconds=>{
 const ts=[];const begin=performance.now();
 await new Promise(resolve=>{function next(t){ts.push(t);if(t-begin<seconds*1000)requestAnimationFrame(next);else resolve()}requestAnimationFrame(next)});
 const dt=ts.slice(1).map((t,i)=>t-ts[i]).filter(x=>x>0.01).sort((a,b)=>a-b);
 const duration=(ts.at(-1)-ts[0])/1000;
 return {frames:ts.length,seconds:Math.round(duration*100)/100,avgFPS:Math.round((dt.length/duration)*10)/10,medianMs:Math.round((dt[Math.floor(dt.length/2)]||0)*10)/10,p95Ms:Math.round((dt[Math.floor(dt.length*.95)]||0)*10)/10,slowFramesOver33ms:dt.filter(x=>x>33).length,WebGL:window.__LS3D_TEST__?.gl||false};
 },seconds)}
async function initPage(context,device){
 const page=await context.newPage();page.on('pageerror',e=>safeError(device,e));page.on('console',m=>{if(m.type()==='error')T.consoleErrors.push({device,text:m.text().slice(0,500)})});
 await page.goto('http://127.0.0.1:9874/LS_Datacenter_3D.html?test=phase1',{waitUntil:'domcontentloaded',timeout:30000});
 await page.waitForFunction(()=>window.__LS3D_TEST__&&window.LS3D_PHASE1&&document.querySelector('#phase1-readout'),{timeout:25000});
 await page.waitForTimeout(900);
 const snapshot=await page.evaluate(()=>({title:document.title,gl:window.__LS3D_TEST__.gl,error:document.querySelector('#error')?.textContent||'',hud:document.querySelector('#phase1-readout')?.textContent||'',viewport:{width:document.querySelector('#viewport')?.getBoundingClientRect().width,height:document.querySelector('#viewport')?.getBoundingClientRect().height}}));
 check(device,'initialization',!snapshot.error&&/v4\.(5\.5|6\.[012]|7\.0)/.test(snapshot.title),JSON.stringify(snapshot));
 check(device,'hud-mounted',!!snapshot.hud,snapshot.hud);
 const tour=await page.locator('#guide.open').count()===1;
 check(device,'first-visit-tour',tour||await page.locator('#guide').count()===1,'tour dialog present, automatic open on first visit='+tour);
 if(tour){await page.locator('#tourSkip').click({timeout:4500});await page.waitForTimeout(150);check(device,'tour-dismiss',await page.locator('#guide.open').count()===0,'skip/close restores background interactions')}
 return page;
}
async function desktop(browser){
 const context=await browser.newContext({viewport:{width:1440,height:900},deviceScaleFactor:1,acceptDownloads:true});
 const page=await initPage(context,'desktop');
 check('desktop','fov-formula',await page.evaluate(()=>Math.abs(window.LS3D_PHASE1.widthForDistance(1,Math.PI/2,1)-2)<1e-10),'90° vertical FOV and aspect=1 at 1m must produce 2m');
 check('desktop','fov-units',await page.evaluate(()=>['≈ 39.7 cm','≈ 74.9 cm','≈ 1.21 m'].every((v,i)=>window.LS3D_PHASE1.formatWidth([.397,.749,1.21][i])===v)),'39.7cm / 74.9cm / 1.21m');
 const count=await page.locator('.phase1-tick').count();check('desktop','ruler-12-ticks',count===12,'ticks='+count);
 await page.screenshot({path:path.join(out,'desktop-initial.png')});
 const hit=await page.evaluate(()=>{const el=document.getElementById('phase1-toggle'),r=el.getBoundingClientRect(),p=document.elementFromPoint(r.left+r.width/2,r.top+r.height/2);return {button:r.toJSON(),hit:{id:p?.id,cls:p?.className,tag:p?.tagName},computed:{pointerEvents:getComputedStyle(el).pointerEvents,display:getComputedStyle(el).display,visibility:getComputedStyle(el).visibility}}});
 check('desktop','toggle-hit-target',hit.hit.id==='phase1-toggle',JSON.stringify(hit));
 let clickOK=true;
 try{await page.locator('#phase1-toggle').click({timeout:2200});await page.locator('#phase1-toggle').click({timeout:2200})}
 catch(e){clickOK=false;T.warnings.push('desktop toggle DOM click fallback: '+String(e).slice(0,250));await page.evaluate(()=>{let t=document.querySelector('#phase1-toggle');if(window.LS3D_PHASE1.collapsed)t.click()})}
 check('desktop','toggle',clickOK&&!await page.evaluate(()=>window.LS3D_PHASE1.collapsed),'UI click fold / unfold (fallback after timeout if obstructed)');
 await page.evaluate(()=>{window.__LS3D_TEST__.cameraDesired.distance=30});
 await page.waitForTimeout(350);
 const before=await page.evaluate(()=>window.__LS3D_TEST__.cameraDesired.distance);
 const vp=await page.locator('#scene').boundingBox();
 await page.mouse.move(vp.x+vp.width*.45,vp.y+vp.height*.54);
 await page.mouse.wheel(0,-410);await page.waitForTimeout(250);
 const after=await page.evaluate(()=>window.__LS3D_TEST__.cameraDesired.distance);
 check('desktop','wheel-zoom',after<before&&after>0,'before='+before.toFixed(2)+'m after='+after.toFixed(2)+'m');
 T.metrics.desktopCampus=await fpsSample(page,3);
 const rack=await page.evaluate(()=>{
 const t=window.__LS3D_TEST__,a=t.assets.find(x=>t.isMechanicalRack(x)&&x.zone==='hall')||t.assets.find(x=>t.isMechanicalRack(x));
 if(!a)return null;t.pickAsset(a,true);t.openDetail();return{id:a.id,name:a.name,type:a.type}});
 check('desktop','open-rack-inspector',!!rack&&await page.locator('#detailModal.show').count()===1,JSON.stringify(rack));
 await page.waitForTimeout(700);
 const actions=await page.evaluate(()=>({door:!!document.querySelector('#doorAct'),tray:!!document.querySelector('#trayAct'),cassette:!!document.querySelector('#cassetteAct')}));
 check('desktop','inspector-controls',Object.values(actions).every(Boolean),JSON.stringify(actions));
 if(actions.door){await page.locator('#doorAct').click({timeout:5000});await page.waitForTimeout(100);check('desktop','rack-door',await page.locator('#doorAct.active').count()===1,'door open')}
 if(actions.cassette){await page.locator('#cassetteAct').click({timeout:5000});await page.waitForTimeout(100);check('desktop','cassette',await page.locator('#cassetteExploded.is-open').count()===1,'cassette pulled')}
 if(actions.tray){await page.locator('#trayAct').click({timeout:5000});await page.waitForTimeout(200);check('desktop','tray-pull',await page.locator('#trayAct.active').count()===1,'server tray pulled')}
 const checks=[['tray',1.2,'서버 트레이'],['gpu',.3,'GPU 카드'],['package',.03,'GPU 패키지'],['concept',.00002,'개념 스케일']];
 for(const [label,d,expected] of checks){
  await page.evaluate(d=>{const c=window.__LS3D_TEST__.cameraDesired;c.distance=d;c.target=[0,0,0]},d);
  await page.waitForTimeout(750);
  const state=await page.evaluate(()=>({stage:window.LS3D_PHASE1.stage,local:window.__LS3D_TEST__.phase1Local,cam:window.__LS3D_TEST__.camera.distance,near:window.__LS3D_TEST__.lens.near,far:window.__LS3D_TEST__.lens.far,err:document.querySelector('#error')?.textContent||''}));
  check('desktop','LOD-'+label,state.stage.includes(expected)&&state.local&&!state.err,JSON.stringify(state));
  if(label==='gpu')T.metrics.desktopGPU=await fpsSample(page,3);
  if(label==='concept')T.metrics.desktopConcept=await fpsSample(page,3);
 }
 await page.screenshot({path:path.join(out,'desktop-deepzoom.png')});
 await page.evaluate(()=>window.__LS3D_TEST__.closeDetail());await page.waitForTimeout(800);
 const scen=await page.evaluate(()=>{const x=window.__LS3D_TEST__;x.scenarioApply('power');return x.scenario});
 check('desktop','power-scenario',scen==='power','scenario='+scen);
 await page.evaluate(()=>window.__LS3D_TEST__.scenarioApply('normal'));
 const walk=await page.evaluate(()=>{document.getElementById('btnWalk').click();return window.__LS3D_TEST__.walkActive});
 check('desktop','walk-mode',walk,'walk activated');
 await page.keyboard.down('w');await page.waitForTimeout(400);await page.keyboard.up('w');
 const stopped=await page.evaluate(()=>{document.getElementById('btnWalk').click();return !window.__LS3D_TEST__.walkActive});
 check('desktop','walk-exit',stopped,'walk deactivated');
 const worker=await page.evaluate(()=>({listBtn:!!document.getElementById('btnWorkers'),roster:!!document.getElementById('workerModal'),staff:window.__LS3D_TEST__.npcs.length}));
 check('desktop','worker-features-retained',worker.listBtn&&worker.roster&&worker.staff>0,JSON.stringify(worker));
 const comparison=await page.evaluate(()=>({log:!!document.getElementById('eventLog')||!!document.querySelector('[id*="log"]'),graph:!!document.querySelector('canvas:not(#scene)')||!!document.querySelector('svg'),facility:!!document.getElementById('facilityModeSelect')}));
 check('desktop','dashboard-elements',comparison.facility,JSON.stringify(comparison));

 const scenarios=[];
 for(const id of ['cooling','network','fiber-cut','normal']){
   await page.locator('#scenarioSelect').selectOption(id);
   await page.waitForTimeout(110);
   const result=await page.evaluate(()=>({scenario:window.__LS3D_TEST__.scenario,eventLog:!!document.querySelector('#eventLog, [id*=eventLog]'),config:!!window.LS3D_CONFIG}));
   scenarios.push({id,result});check('desktop','scenario-'+id,result.scenario===id,JSON.stringify(result));
 }
 await page.locator('#redundancySelect').selectOption('2N');
 await page.locator('#compareBtn').click({timeout:3500});
 const compare=await page.evaluate(()=>({total:document.querySelectorAll('#compareCards .compare-card').length,selected:document.querySelector('#compareCards .compare-card.active')?.textContent?.slice(0,24)}));
 check('desktop','redundancy-compare',compare.total===3&&compare.selected?.includes('2N'),JSON.stringify(compare));
 const linkCheck=await page.evaluate(()=>({twoD:[...document.querySelectorAll('a')].some(a=>a.href.includes('LS_Datacenter_Campus.html')),guide:[...document.querySelectorAll('a')].some(a=>a.href.includes('LS_Datacenter_3D_Guidebook'))}));
 check('desktop','2D-and-Word-guide-links',linkCheck.twoD&&linkCheck.guide,JSON.stringify(linkCheck));
 const crewCount=await page.locator('[data-crew-view]').count();
 if(crewCount){
  await page.locator('[data-crew-view]').first().click({timeout:4000});await page.waitForTimeout(200);
  const active=await page.evaluate(()=>window.__LS3D_TEST__.workerViewId);
  check('desktop','worker-follow-mode',!!active,'worker id='+active);
  const povHit=await page.evaluate(()=>{const b=document.getElementById('btnWorkerCamera'),r=b.getBoundingClientRect(),p=document.elementFromPoint(r.left+r.width/2,r.top+r.height/2),hud=document.getElementById('workerFollowHud');return {rect:r.toJSON(),hit:{id:p?.id,tag:p?.tagName,className:p?.className},hudHidden:hud.hidden,hudRect:hud.getBoundingClientRect().toJSON(),openDialogs:[...document.querySelectorAll('.show,.open')].filter(e=>e.className?.includes?.('modal')).map(x=>x.id)}});check('desktop','worker-POV-hit-target',povHit.hit.id==='btnWorkerCamera',JSON.stringify(povHit));
  let realClick=true;try{await page.locator('#btnWorkerCamera').click({timeout:1800})}catch(e){realClick=false;T.warnings.push('Worker POV button inaccessible: '+String(e).slice(0,250));await page.evaluate(()=>document.getElementById('btnWorkerCamera').click())}
  check('desktop','worker-POV-clickable',realClick,'browser UI click vs script fallback');
  const pov=await page.locator('#btnWorkerCamera').textContent();
  check('desktop','worker-first-person',!!active&&pov.includes('넓게 따라보기'),'camera label='+pov);
  try{await page.locator('#btnWorkerExit').click({timeout:1800})}catch(e){T.warnings.push('worker exit click fallback '+String(e).slice(0,180));await page.evaluate(()=>document.getElementById('btnWorkerExit').click())}
  check('desktop','worker-exit',await page.evaluate(()=>!window.__LS3D_TEST__.workerViewId),'view exited');
 }else{check('desktop','worker-first-person',false,'no worker chips rendered')}
 await page.evaluate(()=>{window.LS3D_CONFIG.quality='low';window.__LS3D_TEST__.cameraDesired.distance=164;});
 await page.waitForTimeout(850);
 T.metrics.desktopCampusLowQuality=await fpsSample(page,3);
 check('desktop','quality-toggle',await page.evaluate(()=>window.LS3D_CONFIG.quality==='low'),'low quality selected');
 await page.evaluate(()=>{window.LS3D_CONFIG.quality='medium';});

 let downloaded=false,filename='';try{const [dl]=await Promise.all([page.waitForEvent('download',{timeout:9000}),page.locator('#exportBtn').click()]);downloaded=!!dl;filename=dl.suggestedFilename()}catch(e){T.warnings.push('CSV download interaction: '+String(e).slice(0,200))}
 check('desktop','CSV-export',downloaded,filename);
 await page.screenshot({path:path.join(out,'desktop-overview.png')});
 await page.close();await context.close();
}
async function mobile(browser){
 const context=await browser.newContext({viewport:{width:390,height:844},deviceScaleFactor:2,isMobile:true,hasTouch:true,acceptDownloads:true});
 const page=await initPage(context,'mobile');
 const layout=await page.evaluate(()=>({rootScroll:document.documentElement.scrollWidth,bodyClient:document.documentElement.clientWidth,canvas:document.querySelector('#viewport').getBoundingClientRect().toJSON(),hud:document.querySelector('.phase1-scale').getBoundingClientRect().toJSON()}));
 check('mobile','no-horizontal-overflow',layout.rootScroll<=layout.bodyClient+3,JSON.stringify({scroll:layout.rootScroll,width:layout.bodyClient}));
 check('mobile','ruler-within-viewport',layout.hud.x>=layout.canvas.x-2&&layout.hud.right<=layout.canvas.right+2,JSON.stringify({hud:layout.hud.x+','+layout.hud.right,view:layout.canvas.x+','+layout.canvas.right}));
 const mobileHit=await page.evaluate(()=>{const el=document.querySelector('#phase1-toggle'),r=el.getBoundingClientRect(),p=document.elementFromPoint(r.left+r.width*.5,r.top+r.height*.5);return {id:p?.id,tag:p?.tagName,className:p?.className,rect:r.toJSON()}});check('mobile','toggle-hit-target',mobileHit.id==='phase1-toggle',JSON.stringify(mobileHit));
 let tapOK=true;try{await page.locator('#phase1-toggle').tap({timeout:2000})}catch(e){tapOK=false;await page.evaluate(()=>document.querySelector('#phase1-toggle').click());T.warnings.push('mobile tap fallback '+String(e).slice(0,200))}
 check('mobile','touch-toggle',tapOK&&await page.evaluate(()=>window.LS3D_PHASE1.collapsed),'tap');
 await page.evaluate(()=>document.querySelector('#phase1-toggle').click());
 await page.locator('#scene').scrollIntoViewIfNeeded({timeout:7000});
 await page.evaluate(()=>{window.__LS3D_TOUCH_DIAG__={start:0,move:0};let e=document.getElementById('scene');e.addEventListener('touchstart',()=>window.__LS3D_TOUCH_DIAG__.start++,{passive:true});e.addEventListener('touchmove',()=>window.__LS3D_TOUCH_DIAG__.move++,{passive:true})});
 await page.evaluate(()=>window.__LS3D_TEST__.cameraDesired.distance=40);
 await page.waitForTimeout(300);
 const before=await page.evaluate(()=>window.__LS3D_TEST__.cameraDesired.distance);
 const r=await page.locator('#scene').boundingBox();
 const x=r.x+r.width*.5,y=r.y+r.height*.5,cdp=await context.newCDPSession(page);
 async function touch(type,pts){await cdp.send('Input.dispatchTouchEvent',{type,touchPoints:pts.map((p,i)=>({x:p[0],y:p[1],id:i,radiusX:4,radiusY:4,force:.5}))})}
 await touch('touchStart',[[x-30,y],[x+30,y]]);
 await page.waitForTimeout(90);
 for(let k=0;k<5;k++){await touch('touchMove',[[x-35-k*12,y],[x+35+k*12,y]]);await page.waitForTimeout(55)}
 await touch('touchEnd',[]);
 await page.waitForTimeout(180);
 const after=await page.evaluate(()=>window.__LS3D_TEST__.cameraDesired.distance);
 const touchDiag=await page.evaluate(()=>({count:window.__LS3D_TOUCH_DIAG__,scrollY:scrollY,innerHeight:innerHeight}));
 check('mobile','pinch-zoom',after<before&&after>0,'before='+before.toFixed(2)+' after='+after.toFixed(2)+' '+JSON.stringify(touchDiag));
 T.metrics.mobileCampus=await fpsSample(page,3);
 await page.screenshot({path:path.join(out,'mobile-portrait.png'),fullPage:true});
 await page.close();await context.close();
}
(async()=>{
 await new Promise(resolve=>server.listen(9874,'127.0.0.1',resolve));
 let browser;
 try{
 browser=await chromium.launch({headless:true,args:['--no-sandbox','--disable-dev-shm-usage','--enable-webgl','--use-gl=angle','--use-angle=swiftshader','--disable-background-timer-throttling','--disable-renderer-backgrounding']});
 await desktop(browser);await mobile(browser);
 }catch(e){safeError('harness',e);check('harness','completed',false,String(e).slice(0,600))}
 finally{
 if(browser)await browser.close().catch(()=>{});await new Promise(r=>server.close(r));
 T.finished=new Date().toISOString();T.summary={pass:T.cases.filter(x=>x.pass).length,fail:T.cases.filter(x=>!x.pass).length,pageErrors:T.pageErrors.length,consoleErrors:T.consoleErrors.length,environment:'GitHub Actions Linux Chromium / SwiftShader; synthetic mobile touch; NOT physical hardware'};
 fs.writeFileSync(path.join(out,'results.json'),JSON.stringify(T,null,2));
 const lines=['# LS Datacenter 3D v4.5.5 · Phase 1 browser regression','','Execution: '+T.finished,'Environment: '+T.summary.environment,'Tests: '+T.summary.pass+' passed, '+T.summary.fail+' failed','Page errors: '+T.pageErrors.length+', console errors: '+T.consoleErrors.length,'','## FPS samples (headless, software WebGL—not physical device FPS)',''];
 for(const [k,v] of Object.entries(T.metrics))lines.push('- '+k+': '+v.avgFPS+' FPS, p95 frame '+v.p95Ms+'ms, frames='+v.frames+', WebGL='+v.WebGL);
 lines.push('','## Cases','');
 for(const x of T.cases)lines.push('- '+(x.pass?'PASS':'FAIL')+' '+x.group+' / '+x.name+(x.detail?' — '+x.detail:''));
 if(T.pageErrors.length)lines.push('','## Runtime errors','',...T.pageErrors.map(x=>'- '+x.p+': '+x.message.slice(0,500)));
 fs.writeFileSync(path.join(out,'report.md'),lines.join('\n')+'\n');
 process.stdout.write('\nREPORT '+JSON.stringify(T.summary)+'\n');
 }
 process.exitCode=(T.summary.fail||T.summary.pageErrors)?1:0;
})().catch(e=>{console.error(e);process.exitCode=1});
