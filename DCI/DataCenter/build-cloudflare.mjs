// Cloudflare Pages: root DCI/DataCenter, build node build-cloudflare.mjs, output .cloudflare-dist
import {readdir,mkdir,copyFile,writeFile,rm,stat} from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
const root=path.dirname(fileURLToPath(import.meta.url));
const out=path.join(root,'.cloudflare-dist');
const allowed=/\.(?:html|css|js|mjs|json|webmanifest|png|jpg|jpeg|webp|gif|svg|ico|woff2?|ttf|otf|glb|gltf|bin|mp4|webm|wasm|csv|txt)$/i;
await rm(out,{recursive:true,force:true});await mkdir(out,{recursive:true});
let count=0;
async function copy(dir,relative=''){
 for(const entry of await readdir(dir,{withFileTypes:true})){
  if(entry.name.startsWith('.')||['node_modules','__pycache__','test-results','playwright-report'].includes(entry.name))continue;
  const rel=path.join(relative,entry.name),src=path.join(dir,entry.name),dest=path.join(out,rel);
  if(entry.isDirectory()){await copy(src,rel);continue;}
  if(!entry.isFile()||!allowed.test(entry.name)||/\.(?:test|spec)\.m?js$/i.test(entry.name)||entry.name==='build-cloudflare.mjs')continue;
  if((await stat(src)).size>25*1024*1024)throw new Error('Cloudflare file exceeds 25 MiB: '+rel);
  await mkdir(path.dirname(dest),{recursive:true});await copyFile(src,dest);count++;
 }
}
await copy(root);
for(const required of ['LS_Datacenter_3D.html','ls3d-three/bootstrap.mjs','ls3d-three/scene-adapter.mjs'])await stat(path.join(out,required));
await writeFile(path.join(out,'index.html'),`<!doctype html><html lang="ko"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta http-equiv="refresh" content="0;url=./LS_Datacenter_3D.html?engine=three"><title>3D 데이터센터 투어</title><p><a href="./LS_Datacenter_3D.html?engine=three">3D 데이터센터 투어 시작</a></p><script>const u=new URL('./LS_Datacenter_3D.html',location.href);u.search=location.search;u.searchParams.set('engine','three');u.hash=location.hash;location.replace(u.href);<\/script></html>`);
await writeFile(path.join(out,'_headers'),`/*
  X-Content-Type-Options: nosniff
/*.html
  Cache-Control: no-cache
`);
console.log('Cloudflare static site ready: '+count+' files in '+out);
