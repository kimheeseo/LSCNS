// Optional verified transport for public CDN assets in managed browsers.
// curl preserves system CA verification; normal CI uses native browser HTTPS.
import {test as base,expect} from '@playwright/test';
import {execFile} from 'node:child_process';
import {promisify} from 'node:util';
const cache=new Map();
async function verifiedCdn(context){
  await context.route('https://cdn.jsdelivr.net/**',async route=>{
    const url=route.request().url();
    if(!cache.has(url))cache.set(url,promisify(execFile)('curl',['--fail','--silent','--show-error','--location','--max-time','30',url],{encoding:'buffer',maxBuffer:20*1024*1024}).then(r=>r.stdout));
    const type=route.request().resourceType();
    await route.fulfill({status:200,headers:{'access-control-allow-origin':'*'},contentType:type==='stylesheet'?'text/css':type==='font'?'font/woff2':'text/javascript',body:await cache.get(url)});
  });
}
export const test=base.extend({
  browser:[async({browser},use)=>{
    const original=browser.newContext;
    if(process.env.LS3D_VERIFIED_CDN==='1')browser.newContext=async function(...args){const context=await original.apply(this,args);await verifiedCdn(context);return context;};
    try{await use(browser);}finally{browser.newContext=original;}
  },{scope:'worker'}]
});
export {expect};
