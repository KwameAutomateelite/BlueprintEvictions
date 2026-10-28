const test = require('node:test');
const assert = require('node:assert/strict');
const patch = require('../integrations/n8n/notice-input-correction-oct27.json');
const params = name => patch.changes.find(c => c.node === name).parameters;
const quiet = { log() {} };
function run(overrides = {}) {
  const fields = {
    id:'recTEST', _noticePolicy:{enabled:true},
    "Tenant's Property Address":'123 EXAMPLE AVENUE, SAN JOSE, CA,95128',
    "Tenant Name(s)":'Alex Example', 'First Name':'Sam', 'Last Name':'Example',
    'Signer Title':'Property Manager', "Property Owner's Name":'Example Owner LLC',
    'Landlord Address':'456 Example Blvd, Cupertino, CA 95014',
    "Landlord's Preferred Email":'owner@example.invalid',
    'Landlords Preferred Phone':'4085550100',
    'Notice Type':'3-Day Notice to Pay Rent or Quit', 'Rent':1000,
    'Rent Due Day':'1st', 'Itemized Rent Due':'September 2026 $1000; October 2026 $1000',
    ...overrides
  };
  const ctx = new Function('$input','console','module',params('build context').jsCode)(
    {first:()=>({json:fields})},quiet,{exports:{}})[0].json;
  const get = name => ({item:{json:name==='Load Notice Runtime Policy'
    ? {id:'recTEST',_noticePolicy:{enabled:true,reviewer_email:'reviewer@example.invalid'}} : ctx}});
  const draft = {DAY_COUNT:3,CASE_NAME:'Example Owner v Example',NOTICE_TYPE:'residential'};
  const payload = new Function('$input','$',params('Build Notice Payload').jsCode)(
    {first:()=>({json:draft})},get)[0].json;
  const evaluate = (expr) => new Function('$json','$','return (' + expr.replace(/^={{\s*/,'').replace(/\s*}}$/,'') + ');')(payload,get);
  const ready = evaluate(params('Configured: Property Complete').conditions.conditions[0].leftValue);
  const review = JSON.parse(evaluate(params('Configured: Property Needs Attention').jsonBody));
  return {ctx,payload,ready,review};
}
test('complete intake reaches the fixed form with one correctly split location', () => {
  const {payload,ready} = run();
  assert.equal(ready,true);
  assert.equal(payload.property_address_street,'123 EXAMPLE AVENUE');
  assert.equal(payload.property_address_city,'SAN JOSE, CA 95128');
  assert.equal(payload.property_address,'123 EXAMPLE AVENUE, SAN JOSE, CA 95128');
  assert.equal(payload.signer_name,'Sam Example');
  assert.equal(payload.landlord_name,'Example Owner LLC');
  assert.equal(payload.total_amount_due,'$2,000.00');
});
test('a total without unpaid periods asks for periods and does not blame a complete address', () => {
  const {ready,review} = run({'Itemized Rent Due':'$2000.00'});
  assert.equal(ready,false);
  assert.match(review.message.body.content,/unpaid rent periods/);
  assert.doesNotMatch(review.message.body.content,/confirm the complete premises/);
});
test('an incomplete premises stays held without borrowing the landlord city', () => {
  const {ctx,ready,review} = run({"Tenant's Property Address":'123 Example St'});
  assert.equal(ctx.property.city,'');
  assert.equal(ready,false);
  assert.match(review.message.body.content,/complete premises address/);
});
test('missing ZIP is reported as an input issue, not filled by inference', () => {
  const {ready,ctx} = run({"Tenant's Property Address":'123 Example St, Sunnyvale, CA'});
  assert.equal(ctx.property.zip,'');
  assert.equal(ready,false);
});
