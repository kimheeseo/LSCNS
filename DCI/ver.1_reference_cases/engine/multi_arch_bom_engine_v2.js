/* v2 validation adapter: calculation remains reference-blind; validation adds input-independence audit. */
const v1 = require('./multi_arch_bom_engine.js');
const BLOCKED_OUTPUT_TOKENS = /(leaf|spine|core|tor|switch|rack|cable|optic|transceiver|ocs)[_-]?(count|qty|quantity|number|total)/i;
function walk(value, prefix='', out=[]) {
  if (Array.isArray(value)) value.forEach((v,i)=>walk(v, prefix+'['+i+']', out));
  else if (value && typeof value === 'object') Object.entries(value).forEach(([k,v])=>walk(v, prefix ? prefix+'.'+k : k, out));
  else out.push([prefix, value]);
  return out;
}
function auditDesignInput(input) {
  const fields = walk(input);
  const blocked = fields.filter(([k])=>BLOCKED_OUTPUT_TOKENS.test(k)).map(([k])=>k);
  const structural = fields.filter(([k])=>/^(topology|network_plan|fabric_design|rack_system|scale_unit|clos_unit|hpc_system)/.test(k)).map(([k])=>k);
  return {
    calculation_path: 'design_input + v1 shared formulas only; reference.json is read after calculation',
    blocked_output_fields: blocked,
    structural_policy_fields: structural,
    input_independent: blocked.length === 0,
    note: blocked.length ? 'Structural count-like fields were provided as a design policy; the case cannot claim fully autonomous A validation until product/constraint solvers derive them.' : 'No direct output-count field found in design input.'
  };
}
function runV2(input, reference) {
  const metrics = v1.deriveArchitecture(input);
  const numerical = v1.compareReference(metrics, reference.metrics, 10);
  const audit = auditDesignInput(input);
  const compared = Object.keys(numerical.metrics || {}).length;
  let level = 'C';
  if (numerical.status === 'PASS' && numerical.coverage_pct === 100 && audit.input_independent && compared >= 5) level = 'A';
  else if (numerical.status === 'PASS' && numerical.coverage_pct === 100 && audit.blocked_output_fields.length <= 2 && compared >= 3) level = 'A-';
  else if (numerical.status === 'PASS') level = 'B';
  return { engine: 'multi_arch_bom_engine_v2.js', calculated_metrics: metrics, numerical_validation: numerical, input_independence: audit, validation_level: level };
}
module.exports = { runV2, auditDesignInput };
