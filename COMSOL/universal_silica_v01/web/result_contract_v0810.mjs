/* Data integrity only. A hash is not a scientific validation certificate. */
const stable=x=>x&&typeof x==='object'?Array.isArray(x)?x.map(stable):Object.fromEntries(Object.keys(x).sort().map(k=>[k,stable(x[k])])):x;
export async function verifyProvenance(result){
 if(!result.provenance?.canonical_input_json)return {status:'LEGACY_MISSING_PROVENANCE',certified:false};
 const text=result.provenance.canonical_input_json;const source=JSON.parse(text),input=result.provenance.hash_scope==='input_manifest'?result.input_manifest:result.input||result.source_design;
 if(result.provenance.hash_scope==='input_manifest'&&JSON.stringify(stable(input?.design))!==JSON.stringify(stable(result.input)))throw Error('manifest design과 result input 불일치');
 if(JSON.stringify(stable(source))!==JSON.stringify(stable(input)))throw Error('canonical input과 결과 input 불일치');
 const bytes=new TextEncoder().encode(text),digest=await crypto.subtle.digest('SHA-256',bytes),hex=[...new Uint8Array(digest)].map(b=>b.toString(16).padStart(2,'0')).join('');
 if(hex!==result.input_sha256)throw Error('입력 hash 불일치');if(!result.solver_version||(!Number.isFinite(Date.parse(result.timestamp_utc))||!/(Z|\+00:00)$/.test(result.timestamp_utc)))throw Error('solver version / UTC 실행 시각 누락');
 return {status:'INPUT_HASH_VERIFIED_ONLY_NOT_PHYSICS',input_sha256:hex,solver_version:result.solver_version,timestamp_utc:result.timestamp_utc,certified:false};
}
