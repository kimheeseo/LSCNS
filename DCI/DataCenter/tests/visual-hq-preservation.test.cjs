const test=require('node:test'),assert=require('node:assert/strict');
const fs=require('node:fs'),path=require('node:path'),crypto=require('node:crypto');
const baseline=require('./assets/visual-hq-preservation-20261126.json');
const engine=require('../engineering-v82.js');
const digest=value=>crypto.createHash('sha256').update(value).digest('hex');
test('visual upgrade leaves all six protected calculation/model files byte-identical',()=>{
  for(const [file,sha256] of Object.entries(baseline.protectedFiles))
    assert.equal(digest(fs.readFileSync(path.resolve(__dirname,'..',file))),sha256,file);
});
baseline.scenarios.forEach((scenario,index)=>{
  test('complete pre-upgrade calculation output matches scenario '+(index+1),()=>{
    assert.equal(digest(JSON.stringify(engine.calculate(scenario.input))),scenario.sha256);
  });
});
