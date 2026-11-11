/** M3 worker: keep expensive FDFD eigen-iteration off the UI thread. */
import {solveVectorFDFD,vectorResultJSON} from "./vector_fdfd_browser.mjs?v=20261112-m3-residual1e7-v2";
self.onmessage=event=>{
 const {id,input,modes=2}=event.data||{};
 try{
  const result=solveVectorFDFD(input,{modes,onProgress:detail=>self.postMessage({id,type:"progress",detail})});
  const json=vectorResultJSON(result,{stride:1});
  self.postMessage({id,type:"result",result:json});
 }catch(error){self.postMessage({id,type:"error",message:String(error?.message||error)});}
};
