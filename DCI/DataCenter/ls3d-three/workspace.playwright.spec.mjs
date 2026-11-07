import {test,expect} from './browser-fixtures.mjs';
const path='/DCI/DataCenter/LS_Datacenter_3D.html';
async function ready(page,engine='three'){
 await page.addInitScript(()=>{localStorage.setItem('ls3d-first-guide-v481','done');localStorage.setItem('ls3d-display-mode-v1',innerWidth<791?'mobile':'pc')});
 await page.goto('http://127.0.0.1:8765'+path+'?engine='+engine);
 await page.waitForFunction(e=>window.LS3D_THREE?.status===(e==='three'?'ready':'legacy'),engine);
 await page.evaluate(()=>{document.getElementById('guide')?.classList.remove('open');window.__LS3D_TEST__.simRunning=false;});
}
for(const size of [{width:1440,height:900},{width:1920,height:1080},{width:390,height:844}])test(`panels, quality and navigation at ${size.width}x${size.height}`,async({page},info)=>{
 await page.setViewportSize(size);await ready(page);
 const errors=[];page.on('pageerror',e=>errors.push(e.message));
 const explorer=page.getByRole('button',{name:'탐색',exact:true}),inspector=page.getByRole('button',{name:'정보',exact:true});
 if(size.width>790){const width=await page.locator('#sceneThree').evaluate(e=>e.getBoundingClientRect().width);await explorer.click();await inspector.click();await expect.poll(()=>page.locator('#sceneThree').evaluate(e=>e.getBoundingClientRect().width)).toBeGreaterThan(width+450);await explorer.click();await inspector.click();}
 else{await explorer.click();await expect(page.locator('#workspaceExplorer')).toBeVisible();await explorer.click();await expect(page.locator('#workspaceExplorer')).toBeHidden();}
 await page.screenshot({path:info.outputPath(`workspace-${size.width}.png`),fullPage:true});
 await page.evaluate(()=>window.__LS3D_TEST__.setZone('hall'));
 await expect(page.locator('.workspace-tree button').first()).toBeAttached();
 await page.evaluate(()=>{window.LS3D_CONFIG.quality='low'});await expect.poll(()=>page.evaluate(()=>window.LS3D_THREE.diagnostics.shadowSize)).toBe(0);
 await page.evaluate(()=>{window.LS3D_CONFIG.quality='high'});await expect.poll(()=>page.evaluate(()=>window.LS3D_THREE.diagnostics.shadowSize)).toBe(2048);
 expect(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth)).toBe(true);expect(errors).toEqual([]);
});
test('xray, doors, camera controls and topology preserve identity',async({page})=>{
 await ready(page);await page.evaluate(()=>{const T=window.__LS3D_TEST__;T.setZone('hall');T.pickAsset(T.visibleAssets().find(a=>T.isMechanicalRack(a)),true);document.getElementById('detailBtn').click();});
 await expect.poll(()=>page.evaluate(()=>window.LS3D_THREE.diagnostics.isolatedRack)).toBeTruthy();
 const before=await page.evaluate(()=>{const T=window.__LS3D_TEST__;return {id:T.selected.id,door:T.mechanical(T.selected).targetDoor,paths:JSON.stringify(T.flowPaths),cables:JSON.stringify(T.cableItems)}});
 await page.getByRole('button',{name:'투시 꺼짐',exact:true}).click();await expect.poll(()=>page.evaluate(()=>window.LS3D_THREE.adapter.rackXray)).toBe(true);
 expect(await page.evaluate(()=>window.__LS3D_TEST__.mechanical(window.__LS3D_TEST__.selected).targetDoor)).toBe(before.door);
 await page.getByRole('button',{name:'투시 켜짐',exact:true}).click();
 await page.evaluate(()=>{const a=window.LS3D_THREE.adapter,id=window.__LS3D_TEST__.selected.id;a.toggleRackFrontDoor(id);a.toggleRearRackDoor(id)});
 await expect.poll(()=>page.evaluate(()=>window.__LS3D_TEST__.mechanical(window.__LS3D_TEST__.selected).targetDoor)).not.toBe(before.door);
 const d=await page.evaluate(()=>window.__LS3D_TEST__.cameraDesired.distance);await page.getByRole('button',{name:'확대 +',exact:true}).click();expect(await page.evaluate(()=>window.__LS3D_TEST__.cameraDesired.distance)).toBeLessThan(d);
 const after=await page.evaluate(()=>{const T=window.__LS3D_TEST__;return {id:T.selected.id,paths:JSON.stringify(T.flowPaths),cables:JSON.stringify(T.cableItems)}});expect(after).toEqual({id:before.id,paths:before.paths,cables:before.cables});
 await page.evaluate(()=>window.__LS3D_TEST__.closeRackScene());await expect.poll(()=>page.evaluate(()=>window.LS3D_THREE.adapter.rackXray)).toBe(false);
});
test('explicit legacy and automatic fallback retain original canvas',async({page})=>{
 await ready(page,'legacy');await expect(page.locator('#scene')).toBeVisible();
 await ready(page);await page.evaluate(()=>window.LS3D_THREE.fallback(new Error('test fallback')));await expect(page.locator('#sceneThree')).toBeHidden();await expect(page.locator('#scene')).toBeVisible();expect(await page.evaluate(()=>window.LS3D_THREE.mode)).toBe('legacy');
});

test('Pages subdirectory loads relative modules and styles',async({page})=>{
 await page.addInitScript(()=>{localStorage.setItem('ls3d-display-mode-v1','pc');localStorage.setItem('ls3d-first-guide-v481','done')});
 const missing=[];page.on('response',r=>{if(r.url().includes('127.0.0.1')&&r.status()>=400)missing.push(r.url())});
 await page.goto('http://127.0.0.1:8766/LSCNS/DCI/DataCenter/LS_Datacenter_3D.html?engine=three');
 await page.waitForFunction(()=>window.LS3D_THREE?.status==='ready');await expect(page.locator('.workspace-tools')).toBeVisible();expect(missing).toEqual([]);
});
