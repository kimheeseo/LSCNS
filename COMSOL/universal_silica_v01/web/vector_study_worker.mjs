import {runStudy} from './vector_study.mjs?v=20261112-m3m6-v2';
self.onmessage=e=>{
 const {id,input,axes,reference,jobs}=e.data??{};
 try{const report=runStudy(input,{axes,reference,jobs,onProgress:p=>self.postMessage({id,type:'progress',...p})});self.postMessage({id,type:'result',report});}
 catch(error){self.postMessage({id,type:'error',message:error.message});}
};
