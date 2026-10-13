// Run with NODE_PATH pointing at a Playwright installation; the app itself uses no library.
// CAMPUS_URL=https://.../LS_Datacenter_Campus.html also checks a deployed build.
const {chromium}=require('playwright'),assert=require('node:assert/strict'),path=require('node:path');
const url=process.env.CAMPUS_URL||'file://'+path.resolve(__dirname,'../LS_Datacenter_Campus.html');
(async()=>{const browser=await chromium.launch();try{
 const page=await browser.newPage({viewport:{width:1440,height:1000}}),errors=[],assets=[];
 page.on('pageerror',e=>errors.push(e.message));page.on('request',r=>{if(/^https?:/.test(r.url())&&r.resourceType()!=='document')assets.push(r.url())});
 await page.goto(url);await page.evaluate(()=>document.fonts.ready);
 assert.equal(await page.locator('body').getAttribute('data-training'),'junior');
 assert.equal(await page.locator('#count').innerText(),'0 / 12');
 assert.deepEqual(await page.evaluate(()=>[campusTrainingDiagnostics().juniorTotal,campusTrainingDiagnostics().seniorTotal]),[12,73]);
 assert.equal(await page.locator('#trainingSelector').evaluate(el=>el.parentElement.classList.contains('console')),true);
 assert.equal(await page.locator('main > *').count(),2,'selector must not create a third grid item');
 const sequence=['sign','security','transformer','gen','plantchiller','ups','cooling','server','tor','storage','mmr','noc'];
 // Location guidance is reversible and never marks a discovery; only the A interaction does.
 await page.locator('#juniorNextBtn').click();assert.equal(await page.evaluate(()=>target().id),'sign');assert.equal(await page.locator('#count').innerText(),'0 / 12');
 for(let i=0;i<sequence.length;i++){
  assert.equal(await page.evaluate(()=>target().id),sequence[i]);
  await page.locator('#action').click();assert.equal(await page.locator('#dialogBox').isVisible(),true);
  assert.equal(await page.locator('#count').innerText(),`${i+1} / 12`);
  assert.match(await page.locator('#pixelFieldExtra').innerText(),/무엇인가요.*왜 필요한가요/s);
  const ratio=await page.locator('[data-junior-picture]').evaluate(c=>c.clientWidth/c.clientHeight);assert.ok(Math.abs(ratio-3)<.05);
  if(i<sequence.length-1){await page.locator('.junior-next-action').click();assert.equal(await page.locator('#dialogBox').isVisible(),false);assert.equal(await page.locator('#count').innerText(),`${i+1} / 12`)}
 }
 await page.locator('#closeInvestigation').click();assert.equal(await page.locator('#juniorNextBtn').isDisabled(),true);
 await page.reload();assert.equal(await page.locator('#count').innerText(),'12 / 12');
 await page.locator('#book').click();assert.equal(await page.locator('#bookList section').count(),12);await page.locator('#bookList button').first().click();assert.equal(await page.locator('#dialogBox').isVisible(),true);await page.locator('#closeInvestigation').click();
 // Destination labels match the actual warp definitions in every junior zone.
 for(const z of ['campus','utility','coolingplant','gray','white','networkroom','ops']){
  await page.locator(`[data-zone="${z}"]`).click();
  const expected=await page.evaluate(()=>[...new Set(Object.values(maps[zone].doors).map(v=>v[0]))].sort());
  const actual=await page.locator('[data-door-to]').evaluateAll(a=>[...new Set(a.map(el=>el.dataset.doorTo))].sort());assert.deepEqual(actual,expected,z);
 }
 // Walk through a labeled door with the original keyboard control, then show its destination.
 await page.locator('[data-zone="gray"]').click();await page.evaluate(()=>enter('gray',11,1,'up'));await page.locator('#game').focus();assert.match(await page.locator('#nearTitle').innerText(),/출입문.*서버실/);
 await page.keyboard.down('ArrowUp');await page.waitForTimeout(80);await page.keyboard.up('ArrowUp');await page.waitForFunction(()=>zone==='white'&&!move);
 // Mobile explanations stay within the map; their final content and close action remain accessible.
 for(const width of [768,390,320]){
  await page.setViewportSize({width,height:844});
  for(const z of ['campus','utility','coolingplant','gray','white','networkroom','ops']){
   await page.locator(`[data-zone="${z}"]`).click();
   const overlap=await page.locator('.map-labels').evaluate(el=>{const a=[...el.children].map(e=>({door:e.dataset.doorTo,r:e.getBoundingClientRect()}));let n=0;for(let i=0;i<a.length;i++)for(let j=i+1;j<a.length;j++){const l=a[i],r=a[j];if((l.door||r.door)&&l.r.left<r.r.right&&l.r.right>r.r.left&&l.r.top<r.r.bottom&&l.r.bottom>r.r.top)n++}return n});
   assert.equal(overlap,0,`${width}px door label overlap in ${z}`);
  }await page.locator('[data-ui-mode="auto"]').click();
  await page.locator('#book').click();await page.locator('#bookList button').first().click();
  const bounds=await page.locator('#dialogBox').evaluate(el=>{const d=el.getBoundingClientRect(),v=document.querySelector('.viewport').getBoundingClientRect();return {contained:d.top>=v.top&&d.bottom<=v.bottom,width:d.width,height:d.height,scroll:el.scrollHeight>el.clientHeight}});
  assert.equal(bounds.contained,true,'dialog outside map at '+width);assert.ok(bounds.width>200&&bounds.height>100);if(width<=390)assert.equal(bounds.scroll,true);
  assert.equal(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth+1),true,'page overflow at '+width);
  await page.locator('#dialogBox').evaluate(el=>{el.scrollTop=el.scrollHeight});await page.locator('#closeInvestigation').click();
 }
 // Senior discoveries are separate, and switching/reloading preserves both sets.
 await page.setViewportSize({width:1440,height:1000});await page.locator('#seniorModeBtn').click();assert.equal(await page.locator('#count').innerText(),'0 / 73');assert.equal(await page.locator('#juniorGuide').isVisible(),false);assert.equal(await page.locator('[data-door-to]').count(),0);
 await page.locator('[data-mission-location="0"]').click();await page.locator('#action').click();assert.equal(await page.locator('#count').innerText(),'1 / 73');await page.locator('#closeInvestigation').click();
 await page.locator('#juniorModeBtn').click();assert.equal(await page.locator('#count').innerText(),'12 / 12');await page.locator('#seniorModeBtn').click();await page.reload();assert.equal(await page.locator('#count').innerText(),'1 / 73');
 await page.locator('#tourTheme').click();assert.equal(await page.locator('body').getAttribute('data-theme'),'night');await page.reload();assert.equal(await page.locator('body').getAttribute('data-theme'),'night');
 await page.setViewportSize({width:390,height:844});await page.locator('[data-mission-location="0"]').click();await page.locator('#action').click();assert.equal(await page.locator('#dialogBox').evaluate(el=>{const d=el.getBoundingClientRect(),v=document.querySelector('.viewport').getBoundingClientRect();return d.top>=v.top&&d.bottom<=v.bottom}),true);await page.locator('#closeInvestigation').click();
 assert.equal(await page.locator('script[src]').count(),0);assert.equal(await page.locator('link[rel="stylesheet"]').count(),0);assert.deepEqual(assets,[]);assert.deepEqual(errors,[]);
 console.log('PASS: 12 A discoveries; next guide does not discover; persistence and separate 73-entry senior notes; 7 door destinations and keyboard transit; 1440/768/390/320px dialog containment/scroll; night mode; single HTML; zero script errors.');
}finally{await browser.close()}})().catch(e=>{console.error(e);process.exitCode=1});
