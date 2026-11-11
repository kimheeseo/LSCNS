/**
 * M1 v0.5.1 independent report writers; design JSON and existing export
 * formats are untouched.
 */
function csvCell(value){
 const raw=value==null?"":String(value);
 // Escape user-supplied text and prevent spreadsheet formula injection.
 const safe=/^[=+@\-]/.test(raw.trimStart())?"'"+raw:raw;
 return '"'+safe.replace(/"/g,'""')+'"';
}
export function m1qReportCSV(report){
 if(report?.kind!=="LSWaveM1QuantitativeVerification"||!Array.isArray(report.cases)||!Array.isArray(report.regression))
  throw Error("Invalid M1 report");
 const columns=["row_type","case_id","source","wavelength_nm","core_radius_um","delta_n_core",
  "n_clad","n_core","V","lp01_analytic_neff","fem_neff","abs_error_neff","relative_error_percent",
  "tolerance_abs_neff","comsol_direct_status","metric","preupdate","postupdate",
  "regression_abs_error","regression_tolerance","status"];
 const rows=[];
 for(const c of report.cases){
  rows.push(["case",c.id,c.sourceType,c.inputs.wavelength_nm,c.inputs.core_radius_um,
   c.inputs.delta_n_core,c.inputs.n_clad,c.inputs.n_core,c.V,
   c.analyticNeff,c.femNeff,c.absError,c.relativePercent,c.toleranceAbsNeff,
   c.comsolStatus,"","","","","",c.status]);
 }
 for(const r of report.regression){
  for(const key of ["neff","mfd","aeff","D"])
   rows.push(["regression",r.id,"v0.5.0 before/after","","","","","","","","","","","","",
    key,r.before[key],r.after[key],r.absError[key],r.tolerances[key],
    r.absError[key]<=r.tolerances[key]?"PASS":"FAIL"]);
 }
 return "\uFEFF"+[columns,...rows].map(row=>row.map(csvCell).join(",")).join("\r\n")+"\r\n";
}
export function m1qReportJSON(report){
 if(report?.kind!=="LSWaveM1QuantitativeVerification")throw Error("Invalid M1 report");
 return JSON.stringify(report,null,2)+"\n";
}
