// Run with Playwright installed: NODE_PATH=/path/to/node_modules node this-file.cjs
// Starts a local read-only HTTP server; optional CAMPUS_TEST_URL checks a deployed build.
const {chromium}=require('playwright');
const assert=require('node:assert/strict'),path=require('node:path'),fs=require('node:fs');
const root=path.resolve(__dirname,'../../..');
const server=process.env.CAMPUS_TEST_URL?null:require('node:child_process').spawn('python',['-m','http.server','8790','--bind','127.0.0.1'],{cwd:root,stdio:'ignore'});
(async()=>{
 await new Promise(r=>setTimeout(r,400));const browser=await chromium.launch({headless:true});
 try{
 const page=await browser.newPage({viewport:{width:1440,height:1100}}),errors=[];
 page.on('pageerror',e=>errors.push(e.message));await page.route('**/youtube.com/**',r=>r.abort());
 await page.goto(process.env.CAMPUS_TEST_URL||'http://127.0.0.1:8790/DCI/DataCenter/LS_Datacenter_Campus.html');await page.locator('[data-zone="white"]').click();
 await page.locator('#rackRowViews').waitFor({state:'visible'});
 const cases=[['supercluster','gpu-row-a-supercluster-pixel.svg','Supercluster'],['cabled_rear','gpu-row-b-cabled-rear-pixel.svg','후면 배선'],['dense_front','compute-b-dense-front-pixel.svg','고밀도 정면']];
 for(const [id,file,title] of cases){
  await page.locator(`[data-rack-view="${id}"]`).click();assert.match(await page.locator('#rackxTitle').innerText(),new RegExp(title));assert.ok((await page.locator('#rackxSpecialImage').getAttribute('src')).endsWith(file));
  await page.waitForFunction(()=>document.getElementById('rackxSpecialImage').complete&&document.getElementById('rackxSpecialImage').naturalWidth>0);
  assert.equal(await page.locator('#rackxSpecialDrawerBtn').isVisible(),false);
  await page.locator('#pixelDetailBtn').click();assert.equal(await page.locator('#pixelDetailBtn').getAttribute('aria-pressed'),'true');
  const size=await page.locator('.tor-image-frame').evaluate(e=>({outer:e.clientWidth,inner:e.scrollWidth}));assert.ok(size.inner>size.outer);
  await page.locator('#pixelOverviewBtn').click();assert.equal(await page.locator('#pixelOverviewBtn').getAttribute('aria-pressed'),'true');
  if(process.env.CAMPUS_QA_DIR){await page.evaluate(()=>document.fonts.ready);fs.mkdirSync(process.env.CAMPUS_QA_DIR,{recursive:true});await page.screenshot({path:path.join(process.env.CAMPUS_QA_DIR,id+'.png')});}
  await page.locator('#rackxClose').click();assert.equal(await page.locator('#rackExplorerModal').evaluate(e=>e.open),false);
 }
 // Reach the same objects using actual movement keys and the investigation button.
 for(const [id] of cases){
  await page.locator('[data-zone="white"]').click();
  const route=await page.evaluate(id=>{
   const o=maps.white.objects.find(o=>o.id===id),queue=[{x:p.x,y:p.y,path:[]}],seen=new Set([p.x+','+p.y]);
   while(queue.length){const cur=queue.shift();for(const [dir,[dx,dy]] of Object.entries(dirs)){
    const nx=cur.x+dx,ny=cur.y+dy;if(nx>=o.x&&nx<o.x+o.w&&ny>=o.y&&ny<o.y+o.h)return {steps:cur.path,face:dir};
    const tile=maps.white.grid[ny]?.[nx],key=nx+','+ny;
    if(!tile||['block','tree','wall','exit','gate','door'].includes(tile)||seen.has(key))continue;
    seen.add(key);queue.push({x:nx,y:ny,path:[...cur.path,{dir,x:nx,y:ny}]});
   }}throw Error('No reachable rack');
  },id);
  for(const step of route.steps){const key='Arrow'+step.dir[0].toUpperCase()+step.dir.slice(1);await page.keyboard.down(key);await page.waitForTimeout(90);await page.keyboard.up(key);await page.waitForFunction(s=>p.x===s.x&&p.y===s.y&&!move,step);}
  const key='Arrow'+route.face[0].toUpperCase()+route.face.slice(1);await page.keyboard.down(key);await page.waitForTimeout(60);await page.keyboard.up(key);
  await page.locator('#action').click();assert.equal(await page.locator('#rackxTitle').innerText(),await page.evaluate(id=>dedicatedViews[id].title,id));await page.locator('#rackxClose').click();
 }
 // Discovered entries reopen their own views from the notebook.
 await page.locator('#book').click();
 const section=page.locator('#bookList section').filter({hasText:'① GPU ROW A · Supercluster'});await section.getByRole('button').click();assert.match(await page.locator('#rackxTitle').innerText(),/Supercluster/);await page.locator('#rackxClose').click();
 await page.setViewportSize({width:390,height:844});await page.locator('[data-zone="white"]').click();
 for(const [id] of cases){await page.locator(`[data-rack-view="${id}"]`).click();await page.waitForFunction(()=>document.getElementById('rackxSpecialImage').complete&&document.getElementById('rackxSpecialImage').naturalWidth>0);assert.ok(await page.locator('#rackxSpecialImage').isVisible());await page.locator('#pixelDetailBtn').click();const sizes=await page.locator('.tor-image-frame').evaluate(e=>({outer:e.clientWidth,inner:e.scrollWidth}));assert.ok(sizes.inner>=960&&sizes.inner>sizes.outer);await page.locator('#pixelOverviewBtn').click();await page.locator('#rackxClose').click();}
 assert.equal(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth+1),true);
 await page.locator('[data-zone="gray"]').click();assert.equal(await page.locator('#rackRowViews').isVisible(),false);
 assert.deepEqual(errors,[]);console.log('PASS: distinct views, assets, zoom/reset, physical keyboard inspection, notebook reopen, mobile and zone visibility; zero page errors.');
 }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exitCode=1}).finally(()=>server?.kill());
