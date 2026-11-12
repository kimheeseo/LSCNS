import fs from 'node:fs';
import os from 'node:os';
import {createHash} from 'node:crypto';
import assert from 'node:assert/strict';
import {solveYee,portableYee} from '../yee_solver.mjs';
import {config,buildYeeGrid,lossDbKm} from '../vector_geometry.mjs';
import {scalarReferences,compareReference,twoCoreMetrics,rowsCSV,migrateProject,jobsFor} from '../vector_study.mjs';
const out=new URL('../results/yee_validation/',import.meta.url);fs.mkdirSync(out,{recursive:true});
const report={generatedAt:new Date().toISOString(),runtime:{node:process.version,cpu:os.cpus()[0].model,platform:os.platform(),cores:os.cpus().length},cases:[]};
const base={wavelength_nm:1550,core_radius_um:4.1,n_core:1.449023621703261,n_clad:1.444023621703261,window_halfwidth_um:18,grid_n:48,modes:2};
const compact=(key,value)=>key==='labels'&&Array.isArray(value)?{length:value.length,sha256:createHash('sha256').update(JSON.stringify(value)).digest('hex'),generation:'synthetic circle a=4.1um halfwidth18um; see benchmark source'}:value;
const strip=r=>({...r,modes:r.modes.map(({field_maps,field_components,...m})=>m),index_map:undefined});
function run(name,cfg,validate){
 const memoryBefore=process.memoryUsage(),start=performance.now();
 try{
  const r=solveYee({...base,...cfg}),refs=scalarReferences(r.config),p=portableYee(r);
  validate?.(r,p);const c={name,config:r.config,wallMs:r.wallMs,processWallMs:performance.now()-start,memoryBefore,memoryAfter:process.memoryUsage(),processHighWaterRSSBytes:process.resourceUsage().maxRSS*1024,reference:refs,result:strip(r)};
  report.cases.push(c);console.log(name,Math.round(r.wallMs)+'ms',r.modes.map(m=>[m.neff,m.neff_imag,m.residual]));fs.writeFileSync(new URL('benchmark.json',out),JSON.stringify(report,compact,2));return {r,p,c};
 }catch(e){report.cases.push({name,config:cfg,error:e.message,wallMs:performance.now()-start});fs.writeFileSync(new URL('benchmark.json',out),JSON.stringify(report,compact,2));console.log('FAILED',name,e.message);return null;}
}
for(const n of [24,48,72,96])run('step-grid-'+n,{grid_n:n},r=>{assert.equal(r.modes.length,2);assert.ok(r.modes.every(m=>m.residual<1e-8&&m.neff>base.n_clad&&m.neff<base.n_core));assert.ok(r.modes[0].Ez_intensity_fraction>0);});
const fixture=JSON.parse(fs.readFileSync(new URL('../../full_vector/results/vector_mode_result.json',import.meta.url)));fixture.input.n_clad=base.n_clad;
for(const n of [48,68,96]){
 const x=run('Colab-physical-input-grid-'+n,{core_radius_um:5.5,n_core:1.4475022010878765,window_halfwidth_um:22,grid_n:n});
 if(x){x.c.comparison=compareReference(x.p,fixture);fs.writeFileSync(new URL('colab_comparison.json',out),JSON.stringify(x.c.comparison,null,2));}
}
const circle=run('circle-region',{grid_n:48,geometry:{kind:'regions',regions:[{kind:'circle',x:0,y:0,radius:4.1,material:'core'}]}});
const analytic=report.cases.find(x=>x.name==='step-grid-48');if(circle)assert.ok(Math.abs(circle.r.modes[0].neff-analytic.result.modes[0].neff)<1e-12,'same geometry path');
for(const size of [128,256,512]){
 const L=18,labels=Array.from({length:size*size},(_,q)=>{const x=-L+(q%size+.5)*2*L/size,y=L-(Math.floor(q/size)+.5)*2*L/size;return Math.hypot(x,y)<4.1?1:0;});
 run('circle-mask-'+size,{grid_n:48,geometry:{kind:'mask',width:size,height:size,labels,palette:{0:'clad',1:'core'},halfwidth_um:L}});
}
for(const subpixel of [false,true])run('subpixel-'+subpixel,{grid_n:48,subpixel});
for(const thickness of [3,4,5])for(const strength of [1,3,5])run('PML-t'+thickness+'-s'+strength,{grid_n:48,pml:{enabled:true,thickness_um:thickness,strength,order:3}});
for(const R of [0,4000,2000,1000])run('bend-R-'+R,{grid_n:48,bend_radius_um:R,pml:{enabled:true,thickness_um:4,strength:3,order:3}});
run('uniform-material-absorption',{grid_n:32,k_core:1e-8,k_clad:1e-8},r=>{assert.ok(r.modes[0].neff_imag>0);assert.ok(Math.abs(r.modes[0].neff_imag/1e-8-1)<.01);});
for(const n of [48,72,96]){const x=run('MCF-two-core-grid-'+n,{grid_n:n,window_halfwidth_um:30,modes:4,geometry:{kind:'mcf',count:2,pitch_um:16}});if(x)x.c.mcf=twoCoreMetrics(x.p,{length_mm:1});}
run('MCF-three-core',{grid_n:48,window_halfwidth_um:34,modes:6,geometry:{kind:'mcf',count:3,pitch_um:16}});
for(const n of [48,72])run('HCF-underresolved-grid-'+n,{grid_n:n,window_halfwidth_um:54,core_radius_um:16,geometry:{kind:'hcf'},pml:{enabled:true,thickness_um:8,strength:3,order:3}},r=>assert.ok(r.warnings.some(s=>s.includes('fewer than 3'))));
assert.throws(()=>buildYeeGrid({grid_n:400}),/grid/);assert.throws(()=>config({geometry:{kind:'hcf'}}),/requires M4/);
assert.throws(()=>jobsFor(base,[{parameter:'a',start:3,end:4,count:18},{parameter:'lambda',start:1500,end:1600,count:18}]),/36/);
const migration=migrateProject({version:'0.1',model:{profile:'step',wavelength_nm:1550,core_radius_um:4.1,delta_n_core:.005,clad_index_override:base.n_clad}});assert.ok(migration.warnings.length>0);
assert.ok(lossDbKm(-1e-8,1550)<0,'negative imaginary index never folded to positive');
assert.ok(rowsCSV([{a:'<raw>,"',neff:1.4}]).includes('"<raw>,"""'));
report.checks='geometry identity, absorption sign/scale, limits, migration, sweep cap, CSV quoting passed';
fs.writeFileSync(new URL('benchmark.json',out),JSON.stringify(report,compact,2));
console.log('DONE',report.cases.length,'cases');
