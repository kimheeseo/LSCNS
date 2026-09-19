/*
 * Node.js reference-case runner.
 * Usage:
 *   node validate_reference_case.js <case-folder>
 *
 * It loads design_input.json + reference.json and writes calculated.json
 * and validation.json using the shared engine. No case-specific formulas.
 */
const fs = require("fs");
const path = require("path");
const engine = require("./multi_arch_bom_engine.js");

function readJson(p){ return JSON.parse(fs.readFileSync(p, "utf8")); }

function runCase(caseDir){
  const input = readJson(path.join(caseDir, "design_input.json"));
  const reference = readJson(path.join(caseDir, "reference.json"));
  const calculatedMetrics = engine.deriveArchitecture(input);
  const validation = engine.compareReference(calculatedMetrics, reference.metrics, 10);

  const calculated = {
    case_id: reference.case_id,
    engine: "multi_arch_bom_engine.js",
    design_input: input,
    metrics: calculatedMetrics
  };
  const validationOut = Object.assign({
    case_id: reference.case_id,
    engine: "multi_arch_bom_engine.js"
  }, validation);

  fs.writeFileSync(path.join(caseDir, "calculated.json"), JSON.stringify(calculated, null, 2) + "\n");
  fs.writeFileSync(path.join(caseDir, "validation.json"), JSON.stringify(validationOut, null, 2) + "\n");
  return {calculated, validation: validationOut};
}

if (require.main === module) {
  const caseDir = process.argv[2];
  if (!caseDir) throw new Error("case folder is required");
  const out = runCase(caseDir);
  console.log(JSON.stringify(out.validation, null, 2));
}

module.exports = { runCase };
