const {chromium}=require('playwright');
const {execFile}=require('node:child_process');
const {promisify}=require('node:util');
const assert=require('node:assert/strict');
(async()=>{const b=await chromium.launch({executablePath:process.env.LS3D_CHROMIUM_EXECUTABLE||'/usr/bin/chromium',args:['--enable-webgl','--use-gl=angle','--use-angle=swiftshader','--enable-unsafe-swiftshader','--disable-gpu-sandbox']});try{const p=await b.newPage();const errors=[];p.on('pageerror',e=>errors.push(e.message));
// Fetch public CDN assets through curl's existing system CA trust; TLS verification remains enabled.
await p.route('https://cdn.jsdelivr.net/**',async r=>{const {stdout}=await promisify(execFile)('curl',['--fail','--silent','--show-error','--location','--max-time','30',r.request().url()],{encoding:'buffer',maxBuffer:20*1024*1024});await r.fulfill({status:200,headers:{'access-control-allow-origin':'*'},contentType:r.request().resourceType()==='stylesheet'?'text/css':r.request().resourceType()==='font'?'font/woff2':'text/javascript',body:stdout});});

await p.addInitScript(()=>{let host;Object.defineProperty(window,'__LS3D_TEST__',{configurable:true,get(){return host},set(v){host=v;v.simRunning=false}});localStorage.setItem('ls3d-first-guide-v481','done');localStorage.setItem('ls3d-display-mode-v1','pc')});
await p.setViewportSize({width:1440,height:900});
if(process.argv[2]==='before')for(const file of ['LS_Datacenter_3D.html','ls3d-three/bootstrap.mjs','ls3d-three/scene-adapter.mjs'])await p.route('**/'+file+'*',async r=>{const {stdout}=await promisify(execFile)('git',['-C',require('node:path').resolve(__dirname,'../../../..'),'show','d34bfd0ec626313a1550bea5da9b114e21de93c6:DCI/DataCenter/'+file]);await r.fulfill({body:stdout,contentType:file.endsWith('.html')?'text/html':'text/javascript'});});
await p.goto('http://127.0.0.1:8765/DCI/DataCenter/LS_Datacenter_3D.html?engine=three');
await p.waitForFunction(()=>window.LS3D_THREE?.status==='ready');
await p.evaluate(()=>{document.getElementById('guide')?.classList.remove('open');window.__LS3D_TEST__.simRunning=false;window.LS3D_CONFIG.quality='high'});
const states=['campus','hall','rack','open','cable','selected'];
const fs=require('fs'),cameraFile=__dirname+'/before-camera.json',cameras=process.argv[2]==='after'?JSON.parse(fs.readFileSync(cameraFile,'utf8')):{};
for(const state of states){await p.evaluate(state=>{const T=window.__LS3D_TEST__,a=window.LS3D_THREE.adapter;T.closeRackScene();T.setZone(state==='campus'?'campus':'hall');if(!['campus','hall'].includes(state)){const rack=T.visibleAssets().find(x=>x.zone==='hall'&&T.isMechanicalRack(x));T.pickAsset(rack,true);if(state!=='selected'){document.getElementById('detailBtn').click();a.focusRackSide(rack.id,'front');}if(['open','cable'].includes(state))T.setMechanical(rack,'door');if(state==='cable')a.runRackAction('cable-patch');}},state);await p.waitForTimeout(2500);if(process.argv[2]==='before')cameras[state]=await p.evaluate(()=>JSON.parse(JSON.stringify(window.__LS3D_TEST__.camera)));else{await p.evaluate(c=>{const T=window.__LS3D_TEST__;Object.assign(T.cameraDesired,structuredClone(c));Object.assign(T.camera,structuredClone(c));},cameras[state]);await p.waitForTimeout(500);}await p.screenshot({path:__dirname+'/'+process.argv[2]+'-'+state+'.png'});}
if(process.argv[2]==='before')fs.writeFileSync(cameraFile,JSON.stringify(cameras,null,2));console.log('Captured',process.argv[2]);
}finally{await b.close()}})().catch(e=>{console.error(e);process.exitCode=1});
