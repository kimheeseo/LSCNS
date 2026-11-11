/**
 * Resource benchmark for the user's actual Node runtime (NOT browser memory).
 * Node >=20; optionally run with --expose-gc:
 * node --expose-gc tests/benchmark_m3_v07.mjs
 * M3_GRIDS=29,41,51,69,81 node --expose-gc tests/benchmark_m3_v07.mjs
 * Reports process heap/RSS/external/ArrayBuffer before-after and peak RSS
 * as MEASURED host numbers. This cannot infer browser Worker RSS.
 */
import {performance} from "node:perf_hooks";
import {resourceUsage} from "node:process";
import {silicaIndex} from "../solver.mjs";
import {solveVectorFDFD} from "../vector_fdfd_browser.mjs";
import {estimatedTypedArrayMiB} from "../m3_study_metrics_v07.mjs";
const grids=(process.env.M3_GRIDS||"29,41,51,69").split(",").map(Number);
const input={profile:"step",wavelength_nm:1550,core_radius_um:5.5,window_halfwidth_um:22,
 n_clad:silicaIndex(1550),n_core:silicaIndex(1550)+.0034785793846155};
const mb=bytes=>+(bytes/1048576).toFixed(3);
const result=[];
for(const grid_n of grids){
 global.gc?.();
 const before=process.memoryUsage(),rssBefore=resourceUsage().maxRSS;
 const start=performance.now(),v=solveVectorFDFD({...input,grid_n},{modes:2}),
  ms=performance.now()-start,after=process.memoryUsage(),rssAfter=resourceUsage().maxRSS;
 const item={grid_n,measuredWallMs:+ms.toFixed(1),
  measuredProcessBeforeMiB:{rss:mb(before.rss),heap:mb(before.heapUsed),external:mb(before.external),arrayBuffers:mb(before.arrayBuffers)},
  measuredProcessAfterMiB:{rss:mb(after.rss),heap:mb(after.heapUsed),external:mb(after.external),arrayBuffers:mb(after.arrayBuffers)},
  measuredPeakRSSMiB:process.platform==="linux"?+(rssAfter/1024).toFixed(3):null,
  typedArrayModelEstimateMiB:+estimatedTypedArrayMiB(grid_n).toFixed(3),
  effectiveIndex:v.modes[0].neff,eigenResidual:v.modes[0].residual};
 result.push(item);
 console.log(JSON.stringify(item));
}
console.log("M3 v0.7 resource benchmark finished; host OS measurements are NOT browser heap measurements.");
